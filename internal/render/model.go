// Package render draws a sheet as an image (for agents that can see) or
// HTML (for people): cell values as formatted in the file, fonts, fills,
// borders, alignment, wrapping, merged cells, banding, and charts.
package render

import (
	"bytes"
	"fmt"
	"image/color"
	"math"
	"os"
	"regexp"
	"strconv"
	"strings"

	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/xlsx"
	"github.com/xuri/excelize/v2"
)

// Options choose what to draw.
type Options struct {
	Sheet   string      // "" for the first visible sheet
	Range   *grid.Range // nil for the used range and its charts
	Scale   float64     // pixel density; 2 by default
	Headers bool        // draw column letters and row numbers
	MaxRows int         // cap on rows drawn (default 60)
	MaxCols int         // cap on columns drawn (default 26)
}

// Result describes what was drawn.
type Result struct {
	Sheet   string `json:"sheet"`
	Range   string `json:"range"`
	Of      string `json:"of,omitempty"` // the full range, when only part was drawn
	Charts  int    `json:"charts"`
	Width   int    `json:"width"`
	Height  int    `json:"height"`
	Missing string `json:"not_drawn,omitempty"` // what the drawing leaves out
}

type border struct {
	width int // in pixels at scale 1; 0 for none
	color color.RGBA
}

type look struct {
	bold, italic bool
	underline    bool // single or double: drawn single
	strike       bool
	size         float64 // points
	ink          color.RGBA
	fill         *color.RGBA
	borders      [4]border // left, top, right, bottom
	halign       string    // left, center, right, or "" (general)
	valign       string    // top, center, bottom ("" means bottom)
	wrap         bool
}

type cell struct {
	text    string
	number  bool // right-aligned by default; shown as ### when too wide
	look    *look
	span    grid.Range // merged area, for its top-left cell
	covered bool       // inside a merged area, not its top-left cell
}

type placedChart struct {
	x, y, w, h int // pixels at scale 1, relative to the grid's origin
	chart      grid.Chart
}

// sheet is everything needed to draw a range.
type sheet struct {
	name      string
	rg        grid.Range
	of        grid.Range
	colX      []int // left edge of each column, plus the right edge of the last
	rowY      []int
	cells     map[[2]int]*cell
	gridlines bool
	charts    []placedChart
	buttons   map[[2]int]bool // cells showing a filter button
	wb        *grid.Workbook
	missing   []string
}

var defaultLook = &look{size: 11, ink: color.RGBA{0, 0, 0, 255}}

// load builds the drawing model of a sheet.
func load(path string, opts Options) (*sheet, error) {
	data, err := os.ReadFile(path)
	if err != nil {
		return nil, err
	}
	wb, err := xlsx.ReadBytes(data)
	if err != nil {
		return nil, fmt.Errorf("%s: %w", path, err)
	}
	f, err := excelize.OpenReader(bytes.NewReader(data))
	if err != nil {
		return nil, err
	}
	defer f.Close()

	var gs *grid.Sheet
	if opts.Sheet == "" {
		gs = wb.FirstVisible()
	} else if gs = wb.Sheet(opts.Sheet); gs == nil {
		return nil, fmt.Errorf("sheet %q not found (sheets: %s)", opts.Sheet, strings.Join(wb.SheetNames(), ", "))
	}
	if gs == nil {
		return nil, fmt.Errorf("workbook has no sheets")
	}
	if opts.MaxRows <= 0 {
		opts.MaxRows = 60
	}
	if opts.MaxCols <= 0 {
		opts.MaxCols = 26
	}
	sh := &sheet{name: gs.Name, cells: map[[2]int]*cell{}, wb: wb, gridlines: true}
	if v, err := f.GetSheetView(gs.Name, 0); err == nil && v.ShowGridLines != nil {
		sh.gridlines = *v.ShowGridLines
	}

	// What to draw: the range asked for, or the used range from A1 with
	// the charts on the sheet.
	if opts.Range != nil {
		sh.rg = *opts.Range
	} else {
		used := gs.Used()
		sh.rg = grid.Range{R1: 1, C1: 1, R2: max(used.R2, 1), C2: max(used.C2, 1)}
		for _, c := range gs.Charts {
			r, col, ok := grid.ParseCell(c.Anchor)
			if !ok {
				continue
			}
			endR, endC := r+18, col+8 // a 640x360 chart covers about this much
			if tr, tc, ok := grid.ParseCell(c.To); ok {
				endR, endC = tr, tc
			} else if c.Width > 0 {
				endC = col + c.Width/64
				endR = r + c.Height/20
			}
			sh.rg = sh.rg.Union(grid.Range{R1: r, C1: col, R2: endR, C2: endC})
		}
	}
	sh.of = sh.rg
	if sh.rg.Rows() > opts.MaxRows {
		sh.rg.R2 = sh.rg.R1 + opts.MaxRows - 1
	}
	if sh.rg.Cols() > opts.MaxCols {
		sh.rg.C2 = sh.rg.C1 + opts.MaxCols - 1
	}

	// Geometry. Excel widths count characters of the default font
	// (7 pixels each, plus 5 of padding); heights are points.
	sh.colX = []int{0}
	for c := sh.rg.C1; c <= sh.rg.C2; c++ {
		w, _ := f.GetColWidth(gs.Name, grid.ColName(c))
		px := int(math.Round(w*7 + 5))
		if w == 0 || gs.ColHidden(c) {
			px = 0
		}
		sh.colX = append(sh.colX, sh.colX[len(sh.colX)-1]+px)
	}
	// Rows the file gives no height of their own fit their text, as Excel
	// sizes them on open; the height of a row past the data is the default.
	heights := make([]int, sh.rg.Rows())
	auto := make([]bool, sh.rg.Rows())
	def, _ := f.GetRowHeight(gs.Name, grid.MaxRows)
	for r := sh.rg.R1; r <= sh.rg.R2; r++ {
		h, _ := f.GetRowHeight(gs.Name, r)
		heights[r-sh.rg.R1] = int(math.Round(h * 96 / 72))
		auto[r-sh.rg.R1] = h == def
		if gs.RowHidden(r) {
			heights[r-sh.rg.R1], auto[r-sh.rg.R1] = 0, false
		}
	}

	// Merged areas.
	merges, _ := f.GetMergeCells(gs.Name)
	for _, m := range merges {
		rg, err := grid.ParseRange(m.GetStartAxis() + ":" + m.GetEndAxis())
		if err != nil {
			continue
		}
		for r := rg.R1; r <= rg.R2; r++ {
			for c := rg.C1; c <= rg.C2; c++ {
				sh.cell(r, c).covered = r != rg.R1 || c != rg.C1
			}
		}
		sh.cell(rg.R1, rg.C1).span = rg
	}

	// Values and looks.
	looks := map[int]*look{}
	for r := sh.rg.R1; r <= sh.rg.R2; r++ {
		for c := sh.rg.C1; c <= sh.rg.C2; c++ {
			name := grid.CellName(r, c)
			id, _ := f.GetCellStyle(gs.Name, name)
			lk, ok := looks[id]
			if !ok {
				lk = lookOf(f, id)
				looks[id] = lk
			}
			raw := gs.At(r, c)
			if lk == defaultLook && (raw == nil || raw.IsEmpty()) {
				continue
			}
			ce := sh.cell(r, c)
			ce.look = lk
			if raw == nil || raw.IsEmpty() {
				continue
			}
			text, _ := f.GetCellValue(gs.Name, name)
			if raw.Formula != "" && (raw.Pending || text == "") {
				if v, err := f.CalcCellValue(gs.Name, name); err == nil {
					text = v
				} else if strings.HasPrefix(err.Error(), "#") {
					text = err.Error()
				} else {
					text = "?"
				}
			}
			v := wb.Value(gs, r, c)
			ce.text = text
			ce.number = v != nil && v.Kind == grid.Number
		}
	}

	// Filter buttons: on the first row of the sheet's filter and the
	// header rows of Excel tables that show them.
	sh.buttons = map[[2]int]bool{}
	addButtons := func(rg grid.Range) {
		for c := rg.C1; c <= rg.C2; c++ {
			if sh.rg.Contains(rg.R1, c) {
				sh.buttons[[2]int{rg.R1, c}] = true
			}
		}
	}
	if gs.Filter.R1 > 0 {
		addButtons(gs.Filter)
	}
	for _, t := range gs.Tables {
		if t.Filter {
			addButtons(t.Range)
		}
	}

	// Banding and other conditional fills of the form sheets writes.
	sh.banding(f, data)
	for _, t := range gs.Tables {
		if !t.Range.Intersect(sh.rg).Empty() {
			sh.missing = append(sh.missing, fmt.Sprintf("the table style of Excel table %q", t.Name))
		}
	}

	for r := sh.rg.R1; r <= sh.rg.R2; r++ {
		if !auto[r-sh.rg.R1] {
			continue
		}
		need := heights[r-sh.rg.R1]
		for c := sh.rg.C1; c <= sh.rg.C2; c++ {
			ce := sh.cells[[2]int{r, c}]
			if ce == nil || ce.covered || ce.look == nil || ce.text == "" || ce.span.Rows() > 1 {
				continue
			}
			lines := 1
			if ce.look.wrap {
				w := sh.edgeX(c+max(ce.span.Cols(), 1)) - sh.colX[c-sh.rg.C1] - 2*padX
				lines = len(wrapText(ce.text, ce.look, w))
			}
			need = max(need, rowFit(ce.look, lines))
		}
		heights[r-sh.rg.R1] = need
	}
	sh.rowY = []int{0}
	for _, h := range heights {
		sh.rowY = append(sh.rowY, sh.rowY[len(sh.rowY)-1]+h)
	}

	// Charts, placed by their anchor cells.
	for _, c := range gs.Charts {
		r, col, ok := grid.ParseCell(c.Anchor)
		if !ok || !sh.rg.Contains(r, col) {
			continue
		}
		x, y := sh.colX[col-sh.rg.C1], sh.rowY[r-sh.rg.R1]
		w, h := c.Width, c.Height
		if tr, tc, ok := grid.ParseCell(c.To); ok {
			w = sh.edgeX(tc) - x
			h = sh.edgeY(tr) - y
		}
		if w <= 0 || h <= 0 {
			w, h = 480, 288
		}
		sh.charts = append(sh.charts, placedChart{x: x, y: y, w: w, h: h, chart: c})
	}
	return sh, nil
}

func (sh *sheet) cell(r, c int) *cell {
	k := [2]int{r, c}
	if sh.cells[k] == nil {
		sh.cells[k] = &cell{}
	}
	return sh.cells[k]
}

// edgeX is the left edge of column c, extrapolating past the drawn range.
// continuous returns the last column a centre-across-selection cell
// spans: the empty cells after it that are set the same way.
func (sh *sheet) continuous(r, c int) int {
	if ce := sh.cells[[2]int{r, c}]; ce == nil || ce.span.Cols() > 1 {
		return c
	}
	for c+1 <= sh.rg.C2 {
		n := sh.cells[[2]int{r, c + 1}]
		if n == nil || n.covered || n.text != "" || n.look == nil || n.look.halign != "centerContinuous" {
			break
		}
		c++
	}
	return c
}

func (sh *sheet) edgeX(c int) int {
	i := c - sh.rg.C1
	if i < len(sh.colX) {
		return sh.colX[max(i, 0)]
	}
	return sh.colX[len(sh.colX)-1] + (i-len(sh.colX)+1)*64
}

func (sh *sheet) edgeY(r int) int {
	i := r - sh.rg.R1
	if i < len(sh.rowY) {
		return sh.rowY[max(i, 0)]
	}
	return sh.rowY[len(sh.rowY)-1] + (i-len(sh.rowY)+1)*20
}

var bandRule = regexp.MustCompile(`^MOD\(ROW\(\)(?:-(\d+))?,2\)=([01])$`)

// banding applies conditional fills that alternate rows, which is how
// sheets bands tables. Other conditional formats are listed as not drawn.
func (sh *sheet) banding(f *excelize.File, data []byte) {
	rules, err := xlsx.CondRules(data, sh.name)
	if err != nil {
		return
	}
	other := false
	for _, rule := range rules {
		var m []string
		if rule.Type == "expression" && len(rule.Formulas) == 1 {
			m = bandRule.FindStringSubmatch(strings.ReplaceAll(rule.Formulas[0], " ", ""))
		}
		if m == nil || rule.Dxf < 0 {
			other = true
			continue
		}
		st, err := f.GetConditionalStyle(rule.Dxf)
		if err != nil || len(st.Fill.Color) == 0 {
			continue
		}
		fill, ok := hexColor(st.Fill.Color[0])
		if !ok {
			continue
		}
		k, _ := strconv.Atoi(m[1])
		want, _ := strconv.Atoi(m[2])
		for _, part := range strings.Fields(rule.Range) {
			rg, err := grid.ParseRange(part)
			if err != nil {
				continue
			}
			for r := max(rg.R1, sh.rg.R1); r <= min(rg.R2, sh.rg.R2); r++ {
				if ((r-k)%2+2)%2 != want {
					continue
				}
				for c := max(rg.C1, sh.rg.C1); c <= min(rg.C2, sh.rg.C2); c++ {
					ce := sh.cell(r, c)
					lk := *defaultLook
					if ce.look != nil {
						lk = *ce.look
					}
					fc := fill
					lk.fill = &fc
					ce.look = &lk
				}
			}
		}
	}
	if other {
		sh.missing = append(sh.missing, "conditional formats other than banding")
	}
}

// lookOf turns an excelize style into what drawing needs.
func lookOf(f *excelize.File, id int) *look {
	if id == 0 {
		return defaultLook
	}
	st, err := f.GetStyle(id)
	if err != nil || st == nil {
		return defaultLook
	}
	lk := &look{size: 11, ink: color.RGBA{0, 0, 0, 255}}
	if st.Font != nil {
		lk.bold, lk.italic = st.Font.Bold, st.Font.Italic
		lk.underline, lk.strike = st.Font.Underline != "" && st.Font.Underline != "none", st.Font.Strike
		if st.Font.Size > 0 {
			lk.size = st.Font.Size
		}
		if c, ok := hexColor(st.Font.Color); ok {
			lk.ink = c
		}
	}
	if st.Fill.Type == "pattern" && st.Fill.Pattern == 1 && len(st.Fill.Color) > 0 {
		if c, ok := hexColor(st.Fill.Color[0]); ok {
			lk.fill = &c
		}
	}
	for _, b := range st.Border {
		i := map[string]int{"left": 0, "top": 1, "right": 2, "bottom": 3}[b.Type]
		if b.Type != "left" && b.Type != "top" && b.Type != "right" && b.Type != "bottom" {
			continue
		}
		w := 1
		if b.Style == 2 || b.Style == 5 || b.Style == 8 {
			w = 2
		}
		if b.Style == 6 {
			w = 3 // double
		}
		c, ok := hexColor(b.Color)
		if !ok {
			c = color.RGBA{0, 0, 0, 255}
		}
		if b.Style > 0 {
			lk.borders[i] = border{width: w, color: c}
		}
	}
	if st.Alignment != nil {
		lk.halign = st.Alignment.Horizontal
		lk.valign = st.Alignment.Vertical
		lk.wrap = st.Alignment.WrapText
	}
	return lk
}

func hexColor(s string) (color.RGBA, bool) {
	s = strings.TrimPrefix(s, "#")
	if len(s) == 8 {
		s = s[2:]
	}
	if len(s) != 6 {
		return color.RGBA{}, false
	}
	v, err := strconv.ParseUint(s, 16, 32)
	if err != nil {
		return color.RGBA{}, false
	}
	return color.RGBA{uint8(v >> 16), uint8(v >> 8), uint8(v), 255}, true
}

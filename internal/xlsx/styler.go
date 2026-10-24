package xlsx

import (
	"fmt"
	"math"
	"strings"

	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/style"
	"github.com/xuri/excelize/v2"
)

// A styler lays looks over the cells of a workbook, keeping what each cell
// has that the look does not set.
type styler struct {
	f     *excelize.File
	cache map[styleKey]int
	notes map[string]bool
	order []string
}

type styleKey struct {
	cur   int
	look  *Look
	edges uint8
}

// Edges of the range a cell is on, for outline and inner borders.
const (
	edgeTop uint8 = 1 << iota
	edgeBottom
	edgeLeft
	edgeRight
)

func newStyler(f *excelize.File) *styler {
	return &styler{f: f, cache: map[styleKey]int{}, notes: map[string]bool{}}
}

func (st *styler) note(s string) {
	if !st.notes[s] {
		st.notes[s] = true
		st.order = append(st.order, s)
	}
}

// Notes lists what the styler decided on its own, like a readable text
// colour on a new fill.
func (st *styler) Notes() []string { return st.order }

// apply lays l over r on sheet. only, when set, limits cell styles to the
// cells it accepts; widths, heights and merges then do not change.
func (st *styler) apply(sheet string, r grid.Range, l *Look, only func(row, col int) bool) error {
	f := st.f
	if r.Empty() {
		return nil
	}
	if only == nil && l.Merge != nil && !*l.Merge {
		if err := st.unmerge(sheet, r); err != nil {
			return err
		}
	}
	if l.CellStyle() {
		for row := r.R1; row <= r.R2; row++ {
			runFrom, runID := 0, -1
			flush := func(to int) error {
				if runID < 0 {
					return nil
				}
				return f.SetCellStyle(sheet, grid.CellName(row, runFrom), grid.CellName(row, to), runID)
			}
			for col := r.C1; col <= r.C2; col++ {
				id := -1
				if only == nil || only(row, col) {
					cur, err := f.GetCellStyle(sheet, grid.CellName(row, col))
					if err != nil {
						return err
					}
					var edges uint8
					if row == r.R1 {
						edges |= edgeTop
					}
					if row == r.R2 {
						edges |= edgeBottom
					}
					if col == r.C1 {
						edges |= edgeLeft
					}
					if col == r.C2 {
						edges |= edgeRight
					}
					if id, err = st.restyle(cur, l, edges); err != nil {
						return err
					}
					if id == cur {
						id = -1
					}
				}
				if id != runID {
					if err := flush(col - 1); err != nil {
						return err
					}
					runFrom, runID = col, id
				}
			}
			if err := flush(r.C2); err != nil {
				return err
			}
		}
	}
	if only != nil {
		return nil
	}
	// A number format that makes values wider than their column would show
	// #####: widen the column, as Excel does when a number is entered.
	if l.NumFmt != "" && l.Width == 0 {
		for col := r.C1; col <= r.C2; col++ {
			name := grid.ColName(col)
			cur, err := f.GetColWidth(sheet, name)
			if err != nil {
				return err
			}
			if need := st.fitWidth(sheet, r, col); need > cur+0.5 {
				if err := f.SetColWidth(sheet, name, name, need); err != nil {
					return err
				}
				st.note(fmt.Sprintf("widened column %s to %s so the numbers show (not #####)", name, grid.FormatNumber(need)))
			}
		}
	}
	if l.Width > 0 {
		if err := f.SetColWidth(sheet, grid.ColName(r.C1), grid.ColName(r.C2), l.Width); err != nil {
			return err
		}
	} else if l.Width < 0 {
		for col := r.C1; col <= r.C2; col++ {
			if err := f.SetColWidth(sheet, grid.ColName(col), grid.ColName(col), st.fitWidth(sheet, r, col)); err != nil {
				return err
			}
		}
	}
	if l.Height > 0 {
		for row := r.R1; row <= r.R2; row++ {
			if err := f.SetRowHeight(sheet, row, l.Height); err != nil {
				return err
			}
		}
	}
	if l.Merge != nil && *l.Merge && (r.Rows() > 1 || r.Cols() > 1) {
		if err := st.merge(sheet, r); err != nil {
			return err
		}
	}
	return nil
}

// fitWidth is a column width that shows the widest value in the range as
// Excel formats it.
func (st *styler) fitWidth(sheet string, r grid.Range, col int) float64 {
	w := 0.0
	for row := r.R1; row <= r.R2; row++ {
		v, err := st.f.GetCellValue(sheet, grid.CellName(row, col))
		if err == nil && v != "" {
			w = math.Max(w, style.TextWidth(v))
		}
	}
	p := style.Default
	return math.Round(min(max(w+2, p.MinWidth), p.MaxWidth)*100) / 100
}

func (st *styler) merge(sheet string, r grid.Range) error {
	var hidden []string
	for row := r.R1; row <= r.R2; row++ {
		for col := r.C1; col <= r.C2; col++ {
			if row == r.R1 && col == r.C1 {
				continue
			}
			cell := grid.CellName(row, col)
			v, _ := st.f.GetCellValue(sheet, cell, excelize.Options{RawCellValue: true})
			fm, _ := st.f.GetCellFormula(sheet, cell)
			if v != "" || fm != "" {
				hidden = append(hidden, cell)
			}
		}
	}
	if len(hidden) > 0 {
		more := ""
		if len(hidden) > 3 {
			more = fmt.Sprintf(" and %d more", len(hidden)-3)
			hidden = hidden[:3]
		}
		return fmt.Errorf("merging %s would hide %s%s, which hold values; clear them first, or use align=center-across, which centres without merging", r, strings.Join(hidden, ", "), more)
	}
	if err := st.unmerge(sheet, r); err != nil {
		return err
	}
	return st.f.MergeCell(sheet, grid.CellName(r.R1, r.C1), grid.CellName(r.R2, r.C2))
}

// unmerge undoes every merge that overlaps r.
func (st *styler) unmerge(sheet string, r grid.Range) error {
	merges, err := st.f.GetMergeCells(sheet)
	if err != nil {
		return err
	}
	for _, m := range merges {
		mr, err := grid.ParseRange(m.GetStartAxis() + ":" + m.GetEndAxis())
		if err != nil || mr.Intersect(r).Empty() {
			continue
		}
		if err := st.f.UnmergeCell(sheet, m.GetStartAxis(), m.GetEndAxis()); err != nil {
			return err
		}
	}
	return nil
}

// restyle returns the style of a cell with style cur once l is laid over
// it.
func (st *styler) restyle(cur int, l *Look, edges uint8) (int, error) {
	key := styleKey{cur, l, edges}
	if id, ok := st.cache[key]; ok {
		return id, nil
	}
	s, err := st.f.GetStyle(cur)
	if err != nil {
		return 0, err
	}
	if l.Reset {
		s = &excelize.Style{NumFmt: s.NumFmt, DecimalPlaces: s.DecimalPlaces, CustomNumFmt: s.CustomNumFmt}
	}
	normalizeFill(s)
	font := func() *excelize.Font {
		if s.Font == nil {
			s.Font = &excelize.Font{}
		}
		return s.Font
	}
	if l.Bold != nil {
		font().Bold = *l.Bold
	}
	if l.Italic != nil {
		font().Italic = *l.Italic
	}
	if l.Underline != nil {
		font().Underline = map[bool]string{true: "single", false: ""}[*l.Underline]
	}
	if l.Strike != nil {
		font().Strike = *l.Strike
	}
	if l.Font != "" {
		font().Family = l.Font
	}
	if l.Size > 0 {
		font().Size = l.Size
	}
	if l.Color != "" {
		font().Color = l.Color
		font().ColorTheme, font().ColorIndexed = nil, 0
	}
	switch l.Fill {
	case "":
	case "none":
		s.Fill = excelize.Fill{}
	default:
		s.Fill = excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{l.Fill}}
		// Keep the text readable on the new fill unless a colour was
		// asked for.
		if l.Color == "" {
			text := "000000"
			if s.Font != nil && s.Font.ColorTheme == nil {
				if c := strings.TrimPrefix(s.Font.Color, "#"); len(c) >= 6 {
					text = strings.ToUpper(c[len(c)-6:])
				}
			}
			if contrast(text, l.Fill) < 4.5 {
				better := readableOn(l.Fill)
				font().Color, font().ColorTheme = better, nil
				st.note(fmt.Sprintf("text set to %s, readable on the %s fill (give color= to choose)", map[string]string{"000000": "black", "FFFFFF": "white"}[better], "#"+l.Fill))
			}
		}
	}
	switch l.NumFmt {
	case "":
	case "General":
		s.NumFmt, s.DecimalPlaces, s.CustomNumFmt = 0, nil, nil
	default:
		code := l.NumFmt
		s.NumFmt, s.DecimalPlaces, s.CustomNumFmt = 0, nil, &code
	}
	align := func() *excelize.Alignment {
		if s.Alignment == nil {
			s.Alignment = &excelize.Alignment{}
		}
		return s.Alignment
	}
	if l.Align != "" {
		align().Horizontal = l.Align
		if l.Align == "general" {
			align().Horizontal = ""
		}
	}
	if l.VAlign != "" {
		align().Vertical = l.VAlign
	}
	if l.Wrap != nil {
		align().WrapText = *l.Wrap
	}
	if l.Indent != nil {
		align().Indent = *l.Indent
		if h := align().Horizontal; *l.Indent > 0 && h != "left" && h != "right" && h != "distributed" {
			align().Horizontal = "left"
		}
	}
	if l.Rotate != nil {
		align().TextRotation = *l.Rotate
	}
	for _, b := range l.Borders {
		for _, side := range b.sidesAt(edges) {
			kept := s.Border[:0]
			for _, x := range s.Border {
				if x.Type != side {
					kept = append(kept, x)
				}
			}
			s.Border = kept
			if b.Style > 0 {
				s.Border = append(s.Border, excelize.Border{Type: side, Color: b.Color, Style: b.Style})
			}
		}
	}
	id, err := st.f.NewStyle(s)
	if err != nil {
		return 0, err
	}
	st.cache[key] = id
	return id, nil
}

// sidesAt lists the sides a border draws on a cell at the given edges of
// its range.
func (b Border) sidesAt(edges uint8) []string {
	all := []string{"top", "bottom", "left", "right"}
	bits := []uint8{edgeTop, edgeBottom, edgeLeft, edgeRight}
	var out []string
	switch b.Sides {
	case "all":
		return all
	case "outline", "inner":
		for i, side := range all {
			if (edges&bits[i] != 0) == (b.Sides == "outline") {
				out = append(out, side)
			}
		}
		return out
	}
	return []string{b.Sides}
}

// normalizeFill repairs fills excelize reads but then refuses to write:
// a pattern without a type (openpyxl writes <patternFill/>, which means
// none) comes back as pattern -1, and a solid fill can carry a second
// colour.
func normalizeFill(s *excelize.Style) {
	if s.Fill.Type == "pattern" && (s.Fill.Pattern < 0 || s.Fill.Pattern > 18) {
		s.Fill = excelize.Fill{}
	}
	if len(s.Fill.Color) > 1 && s.Fill.Type == "pattern" {
		s.Fill.Color = s.Fill.Color[:1]
	}
}

package xlsx

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"math"
	"os"
	"strings"

	"github.com/eemax/sheets/internal/analyze"
	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/style"
	"github.com/xuri/excelize/v2"
)

// FormatOptions control fmt.
type FormatOptions struct {
	Preset    style.Preset
	Sheets    []string   // only these sheets; empty means all
	Overrides []Override // forced column types
	DryRun    bool       // work out the result without writing
}

// Override forces the number format kind of a column.
type Override struct {
	Sheet  string // empty matches every sheet
	Column string // header text or column letters
	Format grid.Format
}

// FormatResult describes what fmt did, so agents can check the inference.
type FormatResult struct {
	Sheets []FormattedSheet `json:"sheets"`
}

// FormattedSheet is the outcome for one sheet.
type FormattedSheet struct {
	Name   string           `json:"name"`
	Tables []FormattedTable `json:"tables,omitempty"`
	Blocks int              `json:"blocks,omitempty"` // regions without a header
	Notes  []string         `json:"notes,omitempty"`
}

// FormattedTable lists the types fmt chose for a table's columns.
type FormattedTable struct {
	Range   string            `json:"range"`
	Total   bool              `json:"total_row,omitempty"`
	Columns []FormattedColumn `json:"columns"`
}

// FormattedColumn is one column's chosen format.
type FormattedColumn struct {
	Col    string `json:"col"`
	Header string `json:"header"`
	Type   string `json:"type"`
}

// Format restyles a workbook in place.
func Format(path string, opts FormatOptions) (*FormatResult, error) {
	data, err := os.ReadFile(path)
	if err != nil {
		return nil, err
	}
	fm, res, err := formatBytes(data, opts)
	if err != nil {
		return nil, fmt.Errorf("%s: %w", path, err)
	}
	defer fm.f.Close()
	if opts.DryRun {
		return res, nil
	}
	return res, save(fm.f, path, fm.patches)
}

// Check reports whether Format would change nothing.
func Check(path string, opts FormatOptions) (bool, *FormatResult, error) {
	data, err := os.ReadFile(path)
	if err != nil {
		return false, nil, err
	}
	base, err := excelize.OpenReader(bytes.NewReader(data))
	if err != nil {
		return false, nil, fmt.Errorf("%s: %w", path, err)
	}
	defer base.Close()
	before, err := serialize(base, nil)
	if err != nil {
		return false, nil, err
	}
	fm, res, err := formatBytes(data, opts)
	if err != nil {
		return false, nil, fmt.Errorf("%s: %w", path, err)
	}
	defer fm.f.Close()
	after, err := serialize(fm.f, fm.patches)
	if err != nil {
		return false, nil, err
	}
	same, err := samePackage(before, after)
	return same, res, err
}

// samePackage compares two xlsx packages part by part.
func samePackage(a, b []byte) (bool, error) {
	pa, err := parts(a)
	if err != nil {
		return false, err
	}
	pb, err := parts(b)
	if err != nil {
		return false, err
	}
	if len(pa) != len(pb) {
		return false, nil
	}
	for name, da := range pa {
		if db, ok := pb[name]; !ok || !bytes.Equal(da, db) {
			return false, nil
		}
	}
	return true, nil
}

func parts(data []byte) (map[string][]byte, error) {
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, err
	}
	out := map[string][]byte{}
	for _, zf := range zr.File {
		r, err := zf.Open()
		if err != nil {
			return nil, err
		}
		b, err := io.ReadAll(r)
		r.Close()
		if err != nil {
			return nil, err
		}
		out[zf.Name] = b
	}
	return out, nil
}

func formatBytes(data []byte, opts FormatOptions) (*formatter, *FormatResult, error) {
	wb, layouts, err := parse(data)
	if err != nil {
		return nil, nil, err
	}
	for _, name := range opts.Sheets {
		if wb.Sheet(name) == nil {
			return nil, nil, fmt.Errorf("sheet %q not found (sheets: %s)", name, strings.Join(wb.SheetNames(), ", "))
		}
	}
	f, err := excelize.OpenReader(bytes.NewReader(data))
	if err != nil {
		return nil, nil, err
	}
	fm := &formatter{
		f: f, wb: wb, p: opts.Preset, layouts: layouts, patches: newPatches(),
		styles: map[cellStyle]int{}, overrides: opts.Overrides, matched: make([]bool, len(opts.Overrides)),
	}
	fail := func(err error) (*formatter, *FormatResult, error) {
		f.Close()
		return nil, nil, err
	}
	res := &FormatResult{}
	if err := f.SetDefaultFont(fm.p.Font); err != nil {
		return fail(err)
	}
	for _, s := range wb.Sheets {
		if len(opts.Sheets) > 0 && !nameIn(s.Name, opts.Sheets) {
			continue
		}
		fs, err := fm.sheet(s)
		if err != nil {
			return fail(fmt.Errorf("sheet %q: %w", s.Name, err))
		}
		res.Sheets = append(res.Sheets, fs)
	}
	for i, ok := range fm.matched {
		if !ok {
			return fail(fmt.Errorf("--as %s: no table column with that header or letter", opts.Overrides[i].Column))
		}
	}
	return fm, res, nil
}

func nameIn(name string, list []string) bool {
	for _, n := range list {
		if strings.EqualFold(n, name) {
			return true
		}
	}
	return false
}

type formatter struct {
	f         *excelize.File
	wb        *grid.Workbook
	p         style.Preset
	layouts   map[string]*layout
	patches   *patches
	styles    map[cellStyle]int
	bandDxf   *int
	overrides []Override
	matched   []bool
}

// cellStyle is everything fmt sets on a cell; it doubles as the style cache
// key.
type cellStyle struct {
	numFmt  string
	bold    bool
	size    float64
	color   string
	fill    string
	align   string
	valign  string
	wrap    bool
	border  bool
	topRule bool
}

func (fm *formatter) styleID(cs cellStyle) (int, error) {
	if id, ok := fm.styles[cs]; ok {
		return id, nil
	}
	st := &excelize.Style{Font: &excelize.Font{Family: fm.p.Font, Size: cs.size, Bold: cs.bold, Color: cs.color}}
	if cs.fill != "" {
		st.Fill = excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{cs.fill}}
	}
	if cs.border {
		top := excelize.Border{Type: "top", Color: fm.p.Border, Style: 1}
		if cs.topRule {
			// Medium weight so it wins over the thin border of the row above.
			top.Color, top.Style = fm.p.TotalBorder, 2
		}
		st.Border = []excelize.Border{
			{Type: "left", Color: fm.p.Border, Style: 1},
			top,
			{Type: "bottom", Color: fm.p.Border, Style: 1},
			{Type: "right", Color: fm.p.Border, Style: 1},
		}
	}
	if cs.align != "" || cs.valign != "" || cs.wrap {
		st.Alignment = &excelize.Alignment{Horizontal: cs.align, Vertical: cs.valign, WrapText: cs.wrap}
	}
	if cs.numFmt != "" {
		code := cs.numFmt
		st.CustomNumFmt = &code
	}
	id, err := fm.f.NewStyle(st)
	if err != nil {
		return 0, err
	}
	fm.styles[cs] = id
	return id, nil
}

// bandStyle returns the conditional format used for banding, reusing an
// identical one already in the file so repeated runs do not pile them up.
func (fm *formatter) bandStyle() (int, error) {
	if fm.bandDxf != nil {
		return *fm.bandDxf, nil
	}
	want := strings.ToUpper(fm.p.BandFill)
	for i := 0; ; i++ {
		st, err := fm.f.GetConditionalStyle(i)
		if err != nil || st == nil {
			break
		}
		if st.Font == nil && len(st.Border) == 0 && st.CustomNumFmt == nil && st.NumFmt == 0 &&
			len(st.Fill.Color) == 1 && strings.HasSuffix(strings.ToUpper(st.Fill.Color[0]), want) {
			fm.bandDxf = &i
			return i, nil
		}
	}
	id, err := fm.f.NewConditionalStyle(&excelize.Style{Fill: excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{fm.p.BandFill}}})
	if err != nil {
		return 0, err
	}
	fm.bandDxf = &id
	return id, nil
}

// colFormat is the decision for one column.
type colFormat struct {
	fmt     grid.Format // Kind General: decide per cell
	plain   bool        // no thousands separator (ids, years)
	numeric bool        // right-align the header
	forced  bool        // from --as: applies to every number in the column
	wrap    bool        // text too long for the widest column: wrap it
}

func (fm *formatter) decide(s *grid.Sheet, reg analyze.Region, col analyze.Column) colFormat {
	cf := colFormat{}
	kind, sym := col.Kind, col.Symbol
	switch col.Type {
	case "integer", "decimal", "currency", "percent", "date", "datetime", "time":
	default:
		kind = grid.General
	}
	for i, o := range fm.overrides {
		if o.Sheet != "" && !strings.EqualFold(o.Sheet, s.Name) {
			continue
		}
		if strings.EqualFold(strings.TrimSpace(o.Column), strings.TrimSpace(col.Header)) && col.Header != "" ||
			strings.EqualFold(o.Column, grid.ColName(col.Col)) {
			fm.matched[i] = true
			kind, cf.forced = o.Format.Kind, true
			if o.Format.Symbol != "" {
				sym = o.Format.Symbol
			}
		}
	}
	f := grid.Format{Kind: kind, Symbol: sym}
	switch kind {
	case grid.Integer:
		cf.plain = analyze.IDLike(col.Header) || col.Min >= 1900 && col.Max <= 2100 && col.Filled > 0 && !cf.forced
	case grid.Decimal:
		f.Decimals = decimalPlaces(col.Decimals, math.Max(math.Abs(col.Min), math.Abs(col.Max)))
	case grid.Currency:
		if col.Decimals > 0 {
			f.Decimals = 2
		}
	case grid.Percent:
		f.Decimals = min(col.Decimals, 2)
	case grid.Time:
		f.Decimals = timeSeconds(fm.wb, s, reg.Body, col.Col)
	}
	cf.fmt = f
	cf.numeric = kind != grid.General && kind != grid.TextFmt
	return cf
}

// decimalPlaces shows two decimals, or up to four for small magnitudes
// (rates, ratios) where two would hide the value.
func decimalPlaces(needed int, magnitude float64) int {
	if magnitude < 10 {
		return min(max(needed, 2), 4)
	}
	return 2
}

// propagate gives formula columns the type of what they compute from:
// =Units*Price is money when Price is. A division is left alone (a ratio of
// two amounts is not an amount), and date arithmetic stays a date only
// when one date is shifted (=Due+30), not when dates are subtracted.
func (fm *formatter) propagate(s *grid.Sheet, reg analyze.Region, cols []analyze.Column, cfs []colFormat) {
	index := map[int]int{}
	for i, c := range cols {
		index[c.Col] = i
	}
	for pass := 0; pass < 3; pass++ {
		changed := false
		for i, c := range cols {
			if cfs[i].forced || len(c.Formulas) == 0 || c.FormulaGroupsMore > 0 {
				continue
			}
			kind := cfs[i].fmt.Kind
			if kind != grid.Integer && kind != grid.Decimal {
				continue
			}
			var currency, dates int
			symbol := ""
			ok := true
			for _, g := range c.Formulas {
				cell := s.Get(g.Range.R1, c.Col)
				if formula.Divides(cell.Formula) {
					ok = false
					break
				}
				for _, rc := range formula.SameRowColumns(cell.Formula, g.Range.R1) {
					j, in := index[rc]
					if !in || rc == c.Col {
						continue
					}
					switch cfs[j].fmt.Kind {
					case grid.Currency:
						currency++
						if symbol == "" {
							symbol = cfs[j].fmt.Symbol
						}
					case grid.Date, grid.DateTime:
						dates++
					}
				}
			}
			if !ok {
				continue
			}
			first := s.Get(c.Formulas[0].Range.R1, c.Col).Formula
			switch {
			case currency > 0 && dates == 0 && callsOnly(first, keepsAmount):
				cfs[i].fmt = grid.Format{Kind: grid.Currency, Symbol: symbol}
				if c.Decimals > 0 {
					cfs[i].fmt.Decimals = 2
				}
				cfs[i].plain = false
				changed = true
			case dates == 1 && currency == 0 && !formula.Multiplies(first) && callsOnly(first, keepsDate):
				cfs[i].fmt = grid.Format{Kind: grid.Date}
				if c.Min != math.Trunc(c.Min) || c.Max != math.Trunc(c.Max) {
					cfs[i].fmt.Kind = grid.DateTime
				}
				cfs[i].plain = false
				changed = true
			}
		}
		if !changed {
			return
		}
	}
}

// Functions whose result has the type of their arguments.
var (
	keepsAmount = setOf("SUM ROUND ROUNDUP ROUNDDOWN MROUND CEILING FLOOR CEILING.MATH FLOOR.MATH MAX MIN ABS IF IFS IFERROR IFNA AVERAGE SUMIF SUMIFS AVERAGEIF AVERAGEIFS")
	keepsDate   = setOf("EDATE EOMONTH WORKDAY WORKDAY.INTL MAX MIN IF IFS IFERROR IFNA")
)

func setOf(s string) map[string]bool {
	m := map[string]bool{}
	for _, f := range strings.Fields(s) {
		m[f] = true
	}
	return m
}

func callsOnly(f string, allowed map[string]bool) bool {
	for _, fn := range formula.Calls(f) {
		if !allowed[fn] {
			return false
		}
	}
	return true
}

func timeSeconds(wb *grid.Workbook, s *grid.Sheet, body grid.Range, col int) int {
	secs := 0
	s.Each(grid.Range{R1: body.R1, C1: col, R2: body.R2, C2: col}, func(_, _ int, c *grid.Cell) {
		if c.Kind == grid.Number && int(math.Round(c.Num*86400))%60 != 0 {
			secs = 1
		}
	})
	return secs
}

// numFormat picks the format for one numeric cell.
func (fm *formatter) numFormat(c *grid.Cell, cf colFormat) (grid.Format, bool) {
	own := fm.wb.FormatOf(c)
	f := cf.fmt
	switch {
	case cf.forced:
	case f.Kind == grid.General:
		f = own
		if f.Kind == grid.General || f.Kind == grid.TextFmt {
			f = grid.Format{Kind: grid.Integer}
			if c.Num != math.Trunc(c.Num) {
				f = grid.Format{Kind: grid.Decimal, Decimals: decimalPlaces(analyze.DecimalsNeeded(c.Num), math.Abs(c.Num))}
			}
		}
		return f, false
	case own.Kind != grid.General && own.Kind != f.Kind && !(isPlainNumber(own.Kind) && isPlainNumber(f.Kind)):
		// A minority cell with its own explicit kind keeps it.
		return own, false
	}
	return f, cf.plain
}

func isPlainNumber(k grid.FormatKind) bool { return k == grid.Integer || k == grid.Decimal }

// fill tracks the widest content of each column.
type widths map[int]float64

func (w widths) see(col int, v float64) {
	if v > w[col] {
		w[col] = v
	}
}

func (fm *formatter) sheet(s *grid.Sheet) (FormattedSheet, error) {
	out := FormattedSheet{Name: s.Name}
	used := s.Used()
	if used.Empty() {
		return out, nil
	}
	name := s.Name
	f := fm.f

	// Strip: cell, row and column styles, and conditional formats.
	lay := fm.layouts[name]
	if err := stripStyles(f, name, lay); err != nil {
		return out, err
	}
	if lay != nil {
		fm.patches.plainCols[name] = true
	}
	cfs, err := f.GetConditionalFormats(name)
	if err != nil {
		return out, err
	}
	for ref := range cfs {
		if err := f.UnsetConditionalFormat(name, ref); err != nil {
			return out, err
		}
	}

	// Excel tables carry their own style and filter; a sheet filter
	// overlapping one makes the file invalid. Their regions only get number
	// formats and widths.
	inListObject := func(r grid.Range) string {
		if t := s.TableAt(r); t != nil {
			return t.Name
		}
		return ""
	}

	w := widths{}
	var tables []analyze.Region
	for _, reg := range analyze.Regions(fm.wb, s) {
		lo := inListObject(reg.Range)
		switch {
		case lo != "":
			out.Notes = append(out.Notes, fmt.Sprintf("%s is Excel table %q: kept its table style, applied number formats and widths only", reg.Range, lo))
		case reg.IsTable():
			tables = append(tables, reg)
		default:
			out.Blocks++
		}
		ft, err := fm.region(s, reg, w, lo != "")
		if err != nil {
			return out, err
		}
		if reg.IsTable() || lo != "" {
			out.Tables = append(out.Tables, ft)
		}
	}

	for col := used.C1; col <= used.C2; col++ {
		v, ok := w[col]
		if !ok {
			continue
		}
		width := math.Round(min(max(v+2, fm.p.MinWidth), fm.p.MaxWidth)*100) / 100
		letter := grid.ColName(col)
		if err := f.SetColWidth(name, letter, letter, width); err != nil {
			return out, err
		}
	}

	// One table: filter its header and freeze the rows above its data.
	panes := &excelize.Panes{}
	if len(tables) == 1 {
		t := tables[0]
		ref := grid.Range{R1: t.HeaderRow, C1: t.Range.C1, R2: t.Body.R2, C2: t.Range.C2}
		// AutoFilter replaces all sheet properties (tab colour, fit to
		// page, ...); put them back.
		props, err := f.GetSheetProps(name)
		if err != nil {
			return out, err
		}
		if err := f.AutoFilter(name, ref.String(), nil); err != nil {
			return out, err
		}
		if err := f.SetSheetProps(name, &props); err != nil {
			return out, err
		}
		if t.HeaderRow <= 10 {
			panes = &excelize.Panes{Freeze: true, YSplit: t.HeaderRow, TopLeftCell: grid.CellName(t.HeaderRow+1, 1), ActivePane: "bottomLeft"}
		}
	} else if lay != nil {
		// No single table to filter: remove any filter left from before.
		fm.patches.noFilter[name] = true
		if err := f.DeleteDefinedName(&excelize.DefinedName{Name: "_xlnm._FilterDatabase", Scope: name}); err != nil && !strings.Contains(err.Error(), "no defined name") {
			return out, err
		}
	}
	if err := f.SetPanes(name, panes); err != nil {
		return out, err
	}
	if err := f.SetSheetDimension(name, grid.Range{R1: 1, C1: 1, R2: used.R2, C2: used.C2}.String()); err != nil {
		return out, err
	}
	show := fm.p.Gridlines
	return out, f.SetSheetView(name, 0, &excelize.ViewOptions{ShowGridLines: &show})
}

func (fm *formatter) region(s *grid.Sheet, reg analyze.Region, w widths, listObject bool) (FormattedTable, error) {
	p := fm.p
	name := s.Name
	table := reg.IsTable() && !listObject
	cols := analyze.Profile(fm.wb, s, reg)
	cfs := make([]colFormat, len(cols))
	ft := FormattedTable{Range: reg.Range.String(), Total: reg.TotalRow > 0}
	for i, c := range cols {
		cfs[i] = fm.decide(s, reg, c)
	}
	fm.propagate(s, reg, cols, cfs)
	wraps := false
	for i, c := range cols {
		if style.TextWidth(strings.Repeat("n", c.MaxLen))+2 > p.MaxWidth && c.Type != "empty" {
			cfs[i].wrap, wraps = true, true
		}
	}
	for i, c := range cols {
		if reg.IsTable() {
			t := c.Type
			if cfs[i].fmt.Kind != grid.General {
				t = cfs[i].fmt.Kind.String()
				if cfs[i].fmt.Kind == grid.Currency {
					t += " " + firstNonEmpty(cfs[i].fmt.Symbol, "$")
				}
			}
			ft.Columns = append(ft.Columns, FormattedColumn{Col: grid.ColName(c.Col), Header: c.Header, Type: t})
		}
	}
	set := func(row, col int, cs cellStyle) error {
		id, err := fm.styleID(cs)
		if err != nil {
			return err
		}
		cell := grid.CellName(row, col)
		return fm.f.SetCellStyle(name, cell, cell, id)
	}

	// Title rows, or a title standing on its own above a table.
	titleRows := reg.TitleRows
	if reg.IsTitle {
		titleRows = 1
	}
	for r := reg.Range.R1; r < reg.Range.R1+titleRows; r++ {
		var err error
		s.Each(grid.Range{R1: r, C1: reg.Range.C1, R2: r, C2: reg.Range.C2}, func(row, col int, _ *grid.Cell) {
			if err == nil {
				err = set(row, col, cellStyle{bold: true, size: p.TitleSize, color: p.TitleColor})
			}
		})
		if err != nil {
			return ft, err
		}
	}

	// Header row.
	if reg.HeaderRow > 0 && !listObject {
		filter := 0.0
		if table {
			filter = 3 // room for the filter button
		}
		for i, c := range cols {
			align := "left"
			if cfs[i].numeric {
				align = "right"
			}
			if err := set(reg.HeaderRow, c.Col, cellStyle{bold: true, size: p.Size, color: p.HeaderFont, fill: p.HeaderFill, align: align, valign: "center", border: true}); err != nil {
				return ft, err
			}
			w.see(c.Col, style.TextWidth(c.Header)*1.1+filter)
		}
	}

	if reg.IsTitle {
		return ft, nil
	}

	// Body and total rows, one column at a time in runs of equal style.
	last := reg.Body.R2
	if reg.TotalRow > 0 {
		last = reg.TotalRow
	}
	for i, colInfo := range cols {
		col := colInfo.Col
		runStart, runID := 0, -1
		flush := func(end int) error {
			if runID < 0 {
				return nil
			}
			return fm.f.SetCellStyle(name, grid.CellName(runStart, col), grid.CellName(end, col), runID)
		}
		for row := reg.Body.R1; row <= last; row++ {
			cs, styled, width := fm.bodyCell(s, row, col, cfs[i], table, row == reg.TotalRow)
			if styled && wraps {
				cs.valign = "top"
				cs.wrap = cfs[i].wrap
			}
			if width > 0 && (table || s.RowLen(row) > 0 && rowCells(s, row, reg.Range) > 1) {
				w.see(col, width)
			}
			id := -1
			if styled {
				var err error
				if id, err = fm.styleID(cs); err != nil {
					return ft, err
				}
			}
			if id != runID {
				if err := flush(row - 1); err != nil {
					return ft, err
				}
				runStart, runID = row, id
			}
		}
		if err := flush(last); err != nil {
			return ft, err
		}
	}

	// Banding on data rows, by conditional format so it survives sorting.
	if table && p.BandFill != "" && reg.Body.Rows() >= 2 {
		dxf, err := fm.bandStyle()
		if err != nil {
			return ft, err
		}
		rule := []excelize.ConditionalFormatOptions{{Type: "formula", Criteria: fmt.Sprintf("MOD(ROW()-%d,2)=1", reg.Body.R1), Format: &dxf}}
		if err := fm.f.SetConditionalFormat(name, reg.Body.String(), rule); err != nil {
			return ft, err
		}
	}
	return ft, nil
}

func rowCells(s *grid.Sheet, row int, r grid.Range) int {
	n := 0
	s.Each(grid.Range{R1: row, C1: r.C1, R2: row, C2: r.C2}, func(int, int, *grid.Cell) { n++ })
	return n
}

// bodyCell computes the style of a data or total cell and the width of its
// content.
func (fm *formatter) bodyCell(s *grid.Sheet, row, col int, cf colFormat, table, total bool) (cellStyle, bool, float64) {
	p := fm.p
	cs := cellStyle{size: p.Size, border: table, bold: total, topRule: total}
	c := fm.wb.Value(s, row, col)
	if c == nil || c.IsEmpty() {
		return cs, table, 0
	}
	width := 0.0
	switch c.Kind {
	case grid.Number:
		f, plain := fm.numFormat(c, cf)
		cs.numFmt = style.NumFmt(f, plain)
		width = float64(len([]rune(style.Render(c.Num, f, plain, fm.wb.Date1904))))
	case grid.Text, grid.Error:
		width = style.TextWidth(c.Text)
	case grid.Bool:
		width = 5
	}
	if total {
		width *= 1.1
	}
	if p.InputColor != "" {
		switch {
		case c.Formula != "" && formula.CrossSheet(c.Formula):
			cs.color = p.LinkColor
		case c.Formula != "":
			cs.color = p.FormulaColor
		case c.Kind == grid.Number:
			cs.color = p.InputColor
		}
	}
	return cs, true, width
}

func firstNonEmpty(a, b string) string {
	if a != "" {
		return a
	}
	return b
}

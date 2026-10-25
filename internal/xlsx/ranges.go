package xlsx

import (
	"errors"
	"fmt"
	"sort"
	"strings"

	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/style"
	"github.com/eemax/sheets/internal/words"
	"github.com/xuri/excelize/v2"
)

// Range operations: clear, copy, move and sort. They work on a workbook
// as opened (each is a command of its own), reading cells from the parsed
// copy and styles, notes and links from excelize.

// taken is a cell's content taken up to be put down elsewhere.
type taken struct {
	cell    grid.Cell // its value; for a formula, its result
	formula string    // "" for a constant
	dynamic bool      // a dynamic array formula (it spills)
	style   int
	note    *excelize.Comment
	link    *grid.Link
}

func (h *taken) empty() bool {
	return h.cell.IsEmpty() && h.style == 0 && h.note == nil && h.link == nil
}

// block is what a range holds, by offset from its top-left cell.
type block struct {
	sheet  string
	r      grid.Range
	cells  map[[2]int]*taken
	merges []grid.Range // merges inside r
}

// errNotFresh guards the range operations, which read the workbook as
// opened.
var errNotFresh = errors.New("internal error: a range operation needs a freshly opened workbook")

func (e *Editor) fresh() bool { return len(e.changed) == 0 && !e.structural && !e.sheetsChanged }

// stretch limits a range given as whole rows or columns to the part of the
// sheet in use, styled cells included.
func (e *Editor) stretch(sheet string, r grid.Range) grid.Range {
	if r.R2 < grid.MaxRows && r.C2 < grid.MaxCols {
		return r
	}
	ext := e.used(sheet)
	if lay := e.layouts[sheet]; lay != nil {
		for _, c := range lay.cells {
			ext = ext.Union(c)
		}
		if s := e.wb.Sheet(sheet); s != nil {
			for _, n := range s.Notes {
				ext = ext.Union(grid.Range{R1: n.Row, C1: n.Col, R2: n.Row, C2: n.Col})
			}
		}
	}
	if ext.Empty() {
		ext = grid.Range{R1: 1, C1: 1, R2: 1, C2: 1}
	}
	if r.R2 == grid.MaxRows {
		r.R1, r.R2 = max(r.R1, 1), max(ext.R2, r.R1)
	}
	if r.C2 == grid.MaxCols {
		r.C1, r.C2 = max(r.C1, 1), max(ext.C2, r.C1)
	}
	return r
}

// spillsOn lists a sheet's dynamic arrays: anchor → the range it fills.
func (e *Editor) spillsOn(sheet string) map[[2]int]grid.Range {
	if lay := e.layouts[sheet]; lay != nil {
		return lay.spills
	}
	return nil
}

// ghost tells whether a cell shows part of a dynamic array's result: a
// value the file stores but that is not the cell's own.
func ghost(spills map[[2]int]grid.Range, row, col int) bool {
	for a, sp := range spills {
		if sp.Contains(row, col) && (a[0] != row || a[1] != col) {
			return true
		}
	}
	return false
}

// checkSpills refuses a range that holds part of a dynamic array's result
// without the formula, or the formula without all of its result.
func (e *Editor) checkSpills(sheet string, r grid.Range, verb string) error {
	for a, sp := range e.spillsOn(sheet) {
		if sp.Intersect(r).Empty() || sp.Rows()*sp.Cols() == 1 {
			continue
		}
		anchor := grid.CellName(a[0], a[1])
		if !r.Contains(a[0], a[1]) {
			return fmt.Errorf("%s holds part of the result of the dynamic array formula in %s (it fills %s); %s the formula's cell instead", grid.FormatRef(sheet, r), anchor, sp, verb)
		}
		if !within([]grid.Range{sp}, r) {
			return fmt.Errorf("%s holds the dynamic array formula in %s but not all of its result (%s); include all of it", grid.FormatRef(sheet, r), anchor, sp)
		}
	}
	return nil
}

// checkMerges refuses a range that cuts through merged cells, as Excel
// does.
func (e *Editor) checkMerges(sheet string, r grid.Range) ([]grid.Range, error) {
	ms, err := e.f.GetMergeCells(sheet)
	if err != nil {
		return nil, err
	}
	var inside []grid.Range
	for _, m := range ms {
		mr, err := grid.ParseRange(m.GetStartAxis() + ":" + m.GetEndAxis())
		if err != nil || mr.Intersect(r).Empty() {
			continue
		}
		if !within([]grid.Range{mr}, r) {
			return nil, fmt.Errorf("%s cuts through the merged cells %s; include all of them, or unmerge them first (sheets style %s unmerge)", grid.FormatRef(sheet, r), mr, cliRange(sheet, mr))
		}
		inside = append(inside, mr)
	}
	return inside, nil
}

// cliRange writes a range for a suggested command, quoted for the shell.
func cliRange(sheet string, r grid.Range) string { return words.Quote(sheet + "!" + r.String()) }

// tableRows says which rows of r are a header or total row of an Excel
// table, which cannot be emptied or overwritten.
func (e *Editor) tableEdge(sheet string, r grid.Range) (tableNow, int, bool) {
	for _, t := range e.tablesOn(sheet) {
		if t.ref.Intersect(r).Empty() {
			continue
		}
		if t.headers > 0 && r.R1 <= t.ref.R1 && t.ref.R1 <= r.R2 {
			return t, t.ref.R1, true
		}
		if t.totals > 0 && r.R1 <= t.ref.R2 && t.ref.R2 <= r.R2 {
			return t, t.ref.R2, true
		}
	}
	return tableNow{}, 0, false
}

// take reads what a range holds: values and formulas (parts of a dynamic
// array's result are not the cells' own and are left out), styles, notes
// and links.
func (e *Editor) take(sheet string, r grid.Range) (*block, error) {
	b := &block{sheet: sheet, r: r, cells: map[[2]int]*taken{}}
	at := func(row, col int) *taken {
		k := [2]int{row - r.R1, col - r.C1}
		h := b.cells[k]
		if h == nil {
			h = &taken{}
			b.cells[k] = h
		}
		return h
	}
	spills := e.spillsOn(sheet)
	if s := e.wb.Sheet(sheet); s != nil {
		s.Each(r, func(row, col int, c *grid.Cell) {
			if c.Formula == "" && ghost(spills, row, col) {
				return
			}
			h := at(row, col)
			h.cell = *c
			h.formula = c.Formula
			if _, ok := spills[[2]int{row, col}]; ok && c.Formula != "" {
				h.dynamic = true
			}
		})
		for _, n := range s.Links {
			if r.Contains(n.Row, n.Col) {
				l := n
				at(n.Row, n.Col).link = &l
			}
		}
	}
	// Styles: of the cells with content, and of styled empty cells.
	styled := map[[2]int]bool{}
	for k := range b.cells {
		styled[k] = true
	}
	if lay := e.layouts[sheet]; lay != nil {
		for _, run := range lay.cells {
			x := run.Intersect(r)
			for row := x.R1; row <= x.R2 && !x.Empty(); row++ {
				for col := x.C1; col <= x.C2; col++ {
					styled[[2]int{row - r.R1, col - r.C1}] = true
				}
			}
		}
	}
	for k := range styled {
		id, err := e.f.GetCellStyle(sheet, grid.CellName(r.R1+k[0], r.C1+k[1]))
		if err != nil {
			return nil, err
		}
		if id != 0 {
			at(r.R1+k[0], r.C1+k[1]).style = id
		}
	}
	cs, err := e.f.GetComments(sheet)
	if err != nil {
		return nil, err
	}
	for _, c := range cs {
		row, col, ok := grid.ParseCell(c.Cell)
		if ok && r.Contains(row, col) {
			c := c
			at(row, col).note = &c
		}
	}
	if b.merges, err = e.checkMerges(sheet, r); err != nil {
		return nil, err
	}
	for k, h := range b.cells {
		if h.empty() {
			delete(b.cells, k)
		}
	}
	return b, nil
}

// threadsIn tells whether a range holds a comment thread (Excel 365),
// which sheets cannot move or copy.
func (e *Editor) threadIn(sheet string, r grid.Range) (string, bool) {
	if s := e.wb.Sheet(sheet); s != nil {
		for _, n := range s.Notes {
			if n.Thread && r.Contains(n.Row, n.Col) {
				return grid.CellName(n.Row, n.Col), true
			}
		}
	}
	return "", false
}

// put writes a held cell. text is the formula to write (moved or
// shifted), when the cell holds one.
func (e *Editor) put(sheet string, row, col int, h *taken, text string, what pasteWhat) error {
	cell := grid.CellName(row, col)
	if what.contents {
		e.noteFresh(sheet, row, col)
		if err := e.clearSpill(sheet, row, col); err != nil {
			return err
		}
		delete(e.patches.errorCells[sheet], cell)
		var err error
		switch {
		case text != "":
			err = e.setFormulaAs(sheet, cell, text, h.dynamic)
		case h.cell.Kind == grid.Text:
			err = setText(e.f, sheet, cell, h.cell.Text)
		case h.cell.Kind == grid.Number:
			err = e.f.SetCellFloat(sheet, cell, h.cell.Num, -1, 64)
		case h.cell.Kind == grid.Bool:
			err = e.f.SetCellBool(sheet, cell, h.cell.Bool)
		case h.cell.Kind == grid.Error:
			if err = e.f.SetCellDefault(sheet, cell, h.cell.Text); err == nil {
				if e.patches.errorCells[sheet] == nil {
					e.patches.errorCells[sheet] = map[string]bool{}
				}
				e.patches.errorCells[sheet][cell] = true
			}
		default:
			err = e.f.SetCellDefault(sheet, cell, "")
		}
		if err != nil {
			return fmt.Errorf("%s: %w", grid.FormatRef(sheet, grid.Range{R1: row, C1: col, R2: row, C2: col}), err)
		}
		e.touch(sheet, row, col)
		if !h.cell.IsEmpty() {
			e.grow(sheet, grid.Range{R1: row, C1: col, R2: row, C2: col})
		}
		if f := e.wb.FormatOf(&h.cell); what.numFmt && h.cell.Kind == grid.Number && f.Kind != grid.General {
			// A cell with a number format of its own keeps it, as with
			// Excel's Paste Values.
			general, err := e.generalNumFmt(sheet, cell)
			if err != nil {
				return err
			}
			if general {
				code := f.Code
				if code == "" || strings.EqualFold(code, "General") {
					code = style.NumFmt(f, false)
				}
				if err := e.applyNumFmt(sheet, cell, code); err != nil {
					return err
				}
			}
		}
	}
	if what.formats {
		if err := e.f.SetCellStyle(sheet, cell, cell, h.style); err != nil {
			return err
		}
	}
	if what.notes {
		if h.note != nil {
			c := *h.note
			c.Cell = cell
			if err := e.addHeldNote(sheet, heldNote{row, col, c}); err != nil {
				return err
			}
		} else if err := e.deleteNote(sheet, cell); err != nil {
			return err
		}
	}
	if what.links {
		if err := e.f.SetCellHyperLink(sheet, cell, "", "None"); err != nil {
			return err
		}
		if l := h.link; l != nil {
			var opts []excelize.HyperlinkOpts
			if l.Tooltip != "" {
				tip := l.Tooltip
				opts = append(opts, excelize.HyperlinkOpts{Tooltip: &tip})
			}
			var err error
			if l.URL != "" {
				err = e.f.SetCellHyperLink(sheet, cell, l.URL, "External", opts...)
			} else {
				err = e.f.SetCellHyperLink(sheet, cell, l.Location, "Location", opts...)
			}
			if err != nil {
				return err
			}
		}
	}
	return nil
}

// setFormulaAs writes a formula, keeping it a dynamic array formula when
// it was one, though its text alone may not show it (=A1:A5).
func (e *Editor) setFormulaAs(sheet, cell, text string, dynamic bool) error {
	if !dynamic || dynamicFormula(formula.ToFile(text)) {
		return e.setFormula(sheet, cell, text)
	}
	if err := e.f.SetCellDefault(sheet, cell, ""); err != nil {
		return err
	}
	t, ref := excelize.STCellFormulaTypeArray, cell
	if err := e.f.SetCellFormula(sheet, cell, formula.ToFile(text), excelize.FormulaOpts{Type: &t, Ref: &ref}); err != nil {
		return err
	}
	e.patches.markCell(sheet, cell, true)
	return nil
}

// pasteWhat is what a paste writes. numFmt gives pasted values their
// number format when formats are not pasted (Excel's "values and number
// formats"): a date pasted as a value stays a date.
type pasteWhat struct {
	contents, formats, notes, links, numFmt bool
}

var pasteAll = pasteWhat{contents: true, formats: true, notes: true, links: true}

// wipe empties the cells of r that a paste of what would otherwise leave
// behind: the cells a block does not cover.
func (e *Editor) wipe(sheet string, r grid.Range, keep map[[2]int]*taken, what pasteWhat) error {
	empty := &taken{}
	s := e.wb.Sheet(sheet)
	clear := func(row, col int) error {
		if keep[[2]int{row - r.R1, col - r.C1}] != nil {
			return nil
		}
		return e.put(sheet, row, col, empty, "", what)
	}
	seen := map[[2]int]bool{}
	visit := func(row, col int) error {
		if !r.Contains(row, col) || seen[[2]int{row, col}] {
			return nil
		}
		seen[[2]int{row, col}] = true
		return clear(row, col)
	}
	var err error
	if s != nil && what.contents {
		s.Each(r, func(row, col int, _ *grid.Cell) {
			if err == nil {
				err = visit(row, col)
			}
		})
	}
	if err != nil {
		return err
	}
	if lay := e.layouts[sheet]; lay != nil && what.formats {
		for _, run := range lay.cells {
			x := run.Intersect(r)
			for row := x.R1; row <= x.R2 && !x.Empty(); row++ {
				for col := x.C1; col <= x.C2; col++ {
					if err := visit(row, col); err != nil {
						return err
					}
				}
			}
		}
	}
	if s != nil && what.notes {
		for _, n := range s.Notes {
			if !n.Thread {
				if err := visit(n.Row, n.Col); err != nil {
					return err
				}
			}
		}
	}
	if s != nil && what.links {
		for _, l := range s.Links {
			if err := visit(l.Row, l.Col); err != nil {
				return err
			}
		}
	}
	return nil
}

// ClearWhat is what clear takes away.
type ClearWhat struct {
	Contents bool // values and formulas
	Formats  bool // cell styles, merges, conditional formats within, recorded styling
	Notes    bool
	Links    bool
}

// RangeResult reports a range operation.
type RangeResult struct {
	From  string   `json:"from,omitempty"`
	To    string   `json:"to"`
	Cells int      `json:"cells"` // cells changed
	Notes []string `json:"notes,omitempty"`
}

// ClearRange empties cells: their contents, formats, notes or links, as
// asked. Data validation stays, as in Excel.
func (e *Editor) ClearRange(sheet string, r grid.Range, w ClearWhat) (RangeResult, error) {
	if !e.fresh() {
		return RangeResult{}, errNotFresh
	}
	sheet, err := e.sheetName(sheet)
	if err != nil {
		return RangeResult{}, err
	}
	r = e.stretch(sheet, r)
	res := RangeResult{To: grid.FormatRef(sheet, r)}
	if w.Contents {
		if t, row, ok := e.tableEdge(sheet, r); ok && row == t.ref.R1 {
			return res, fmt.Errorf("row %d is the header row of Excel table %s, which needs its column names; clear the rows below it", row, t)
		}
		if err := e.checkSpills(sheet, r, "clear"); err != nil {
			return res, err
		}
	}
	if w.Notes {
		if cell, ok := e.threadIn(sheet, r); ok {
			return res, fmt.Errorf("%s!%s holds a comment thread (Excel 365); sheets cannot remove those: delete it in Excel, or leave out --notes", grid.QuoteSheet(sheet), cell)
		}
	}
	changed := map[[2]int]bool{}
	what := pasteWhat{contents: w.Contents, formats: w.Formats, notes: w.Notes, links: w.Links}
	if w.Contents {
		if err := e.unshareSheet(sheet); err != nil {
			return res, err
		}
		if s := e.wb.Sheet(sheet); s != nil {
			spills := e.spillsOn(sheet)
			s.Each(r, func(row, col int, c *grid.Cell) {
				if c.Formula != "" || !ghost(spills, row, col) {
					changed[[2]int{row, col}] = true
				}
			})
		}
		// A dynamic array's result goes with its formula.
		for a := range e.spillsOn(sheet) {
			if r.Contains(a[0], a[1]) {
				if err := e.clearSpill(sheet, a[0], a[1]); err != nil {
					return res, err
				}
			}
		}
	}
	if w.Formats {
		if lay := e.layouts[sheet]; lay != nil {
			for _, run := range lay.cells {
				x := run.Intersect(r)
				for row := x.R1; row <= x.R2 && !x.Empty(); row++ {
					for col := x.C1; col <= x.C2; col++ {
						changed[[2]int{row, col}] = true
					}
				}
			}
		}
	}
	if s := e.wb.Sheet(sheet); s != nil {
		for _, n := range s.Notes {
			if w.Notes && r.Contains(n.Row, n.Col) {
				changed[[2]int{n.Row, n.Col}] = true
			}
		}
		for _, l := range s.Links {
			if w.Links && r.Contains(l.Row, l.Col) {
				changed[[2]int{l.Row, l.Col}] = true
			}
		}
	}
	if err := e.wipe(sheet, r, nil, what); err != nil {
		return res, err
	}
	res.Cells = len(changed)
	if w.Formats {
		n, err := e.clearFormats(sheet, r)
		if err != nil {
			return res, err
		}
		res.Notes = append(res.Notes, n...)
	}
	return res, nil
}

// clearFormats takes away the formats of r beyond cell styles: merges,
// conditional formats lying within it, and recorded styling over it.
func (e *Editor) clearFormats(sheet string, r grid.Range) ([]string, error) {
	var notes []string
	ms, err := e.f.GetMergeCells(sheet)
	if err != nil {
		return nil, err
	}
	for _, m := range ms {
		mr, err := grid.ParseRange(m.GetStartAxis() + ":" + m.GetEndAxis())
		if err == nil && !mr.Intersect(r).Empty() {
			if err := e.f.UnmergeCell(sheet, m.GetStartAxis(), m.GetEndAxis()); err != nil {
				return nil, err
			}
			e.dropMerges()
		}
	}
	cfs, err := e.CondFormats(sheet)
	if err != nil {
		return nil, err
	}
	kept := 0
	for _, cf := range cfs {
		rs := sqrefRanges(cf.Range)
		if overlaps(rs, []grid.Range{r}) && !within(rs, r) {
			kept++
		}
	}
	if _, err := e.ClearCond(sheet, &r); err != nil {
		return nil, err
	}
	if kept > 0 {
		notes = append(notes, fmt.Sprintf("kept %s that reach beyond %s (sheets cond --list)", countOf(kept, "conditional format"), r))
	}
	if err := e.unrecordStyle(sheet, r); err != nil {
		return nil, err
	}
	return notes, nil
}

// unrecordStyle takes r out of the recorded styling of a sheet, so fmt does
// not lay it there again.
func (e *Editor) unrecordStyle(sheet string, r grid.Range) error {
	for _, rec := range e.Records() {
		if rec.Kind != "style" || !strings.EqualFold(rec.Sheet, sheet) {
			continue
		}
		var areas []grid.Range
		cut := false
		for _, a := range rec.Areas {
			if a.Intersect(r).Empty() {
				areas = append(areas, a)
				continue
			}
			cut = true
			areas = append(areas, subtract(a, r)...)
		}
		if cut {
			if err := setAreas(e.f, rec, areas); err != nil {
				return err
			}
		}
	}
	return nil
}

// dropValidation is ClearValidation for a range that may hold none.
func (e *Editor) dropValidation(sheet string, r grid.Range) error {
	_, err := e.ClearValidation(sheet, &r)
	return err
}

// validationsIn lists the data validation rules over r: each rule's part
// within r, with its formulas as they read from that part's first cell.
func (e *Editor) validationsIn(sheet string, r grid.Range) ([]*excelize.DataValidation, error) {
	dvs, err := e.f.GetDataValidations(sheet)
	if err != nil {
		return nil, err
	}
	var out []*excelize.DataValidation
	for _, dv := range dvs {
		rs := sqrefRanges(dv.Sqref)
		if len(rs) == 0 {
			continue
		}
		anchor := rs[0]
		var parts []string
		var first *grid.Range
		for _, x := range rs {
			if p := x.Intersect(r); !p.Empty() {
				parts = append(parts, p.String())
				if first == nil {
					first = &p
				}
			}
		}
		if first == nil {
			continue
		}
		c := *dv
		c.Sqref = strings.Join(parts, " ")
		// Formulas are relative to the first cell of the rule's range.
		dr, dc := first.R1-anchor.R1, first.C1-anchor.C1
		c.Formula1 = escapeXML(formula.Shift(unescapeXML(dv.Formula1), dr, dc))
		c.Formula2 = escapeXML(formula.Shift(unescapeXML(dv.Formula2), dr, dc))
		out = append(out, &c)
	}
	return out, nil
}

// placeValidations adds rules read by validationsIn at another place: each
// range and formula goes through fix.
func (e *Editor) placeValidations(sheet string, dvs []*excelize.DataValidation, dr, dc int, fix func(string) string) error {
	for _, dv := range dvs {
		var parts []string
		for _, x := range sqrefRanges(dv.Sqref) {
			y := grid.Range{R1: x.R1 + dr, C1: x.C1 + dc, R2: x.R2 + dr, C2: x.C2 + dc}
			if y.R1 < 1 || y.C1 < 1 || y.R2 > grid.MaxRows || y.C2 > grid.MaxCols {
				continue
			}
			parts = append(parts, y.String())
		}
		if len(parts) == 0 {
			continue
		}
		c := *dv
		c.Sqref = strings.Join(parts, " ")
		c.Formula1 = escapeXML(fix(unescapeXML(dv.Formula1)))
		c.Formula2 = escapeXML(fix(unescapeXML(dv.Formula2)))
		if err := e.f.AddDataValidation(sheet, &c); err != nil {
			return err
		}
	}
	return nil
}

func escapeXML(s string) string {
	return strings.NewReplacer("&", "&amp;", "<", "&lt;", ">", "&gt;", `"`, "&quot;").Replace(s)
}

// recordsIn lists the recorded styling of a sheet over r: style entries
// with their areas cut to r, and cond entries lying within it.
func (e *Editor) recordsIn(sheet string, r grid.Range) []Record {
	var out []Record
	for _, rec := range e.Records() {
		if !strings.EqualFold(rec.Sheet, sheet) {
			continue
		}
		switch rec.Kind {
		case "style":
			var areas []grid.Range
			for _, a := range rec.Areas {
				if x := a.Intersect(r); !x.Empty() {
					areas = append(areas, x)
				}
			}
			if len(areas) > 0 {
				rec.Areas = areas
				out = append(out, rec)
			}
		case "cond":
			if within(rec.Areas, r) {
				out = append(out, rec)
			}
		}
	}
	return out
}

// placeRecords records styling taken from one place at another, moved by
// (dr, dc): style entries as they are, cond entries by applying their rule
// again there, its formula moved by fix.
func (e *Editor) placeRecords(sheet string, recs []Record, dr, dc int, fix func(string) string) error {
	for _, rec := range recs {
		var areas []grid.Range
		for _, a := range rec.Areas {
			areas = append(areas, grid.Range{R1: a.R1 + dr, C1: a.C1 + dc, R2: a.R2 + dr, C2: a.C2 + dc})
		}
		switch rec.Kind {
		case "style":
			if _, err := addRecord(e.f, Record{Kind: "style", Sheet: sheet, Areas: areas, Args: rec.Args}); err != nil {
				return err
			}
		case "cond":
			args := append([]string{}, rec.Args...)
			for i, a := range args {
				if strings.HasPrefix(a, "=") {
					args[i] = "=" + fix(a[1:])
				}
			}
			if _, err := e.Cond(sheet, areas[0], args); err != nil {
				return err
			}
		}
	}
	return nil
}

// dest works out where a copy of src lands given at: its top-left cell, or
// a range the copy fills (a whole multiple of src, as Excel tiles a paste).
func dest(src, at grid.Range) ([]grid.Range, error) {
	if at.Rows() == 1 && at.Cols() == 1 || at.Rows() == src.Rows() && at.Cols() == src.Cols() {
		return []grid.Range{{R1: at.R1, C1: at.C1, R2: at.R1 + src.Rows() - 1, C2: at.C1 + src.Cols() - 1}}, nil
	}
	if at.Rows()%src.Rows() != 0 || at.Cols()%src.Cols() != 0 {
		return nil, fmt.Errorf("%s is %d×%d, which %s (%d×%d) does not fill evenly; give its top-left cell", at, at.Rows(), at.Cols(), src, src.Rows(), src.Cols())
	}
	var out []grid.Range
	for r := at.R1; r <= at.R2; r += src.Rows() {
		for c := at.C1; c <= at.C2; c += src.Cols() {
			out = append(out, grid.Range{R1: r, C1: c, R2: r + src.Rows() - 1, C2: c + src.Cols() - 1})
		}
	}
	return out, nil
}

// CopyOptions say what a copy takes.
type CopyOptions struct {
	Values  bool // values only: a formula gives its result; the target keeps its formats
	Formats bool // formats only
}

// CopyRange copies cells as Excel's copy and paste does: values and
// formulas (relative references shift with them), formats, merges, notes,
// links, data validation, and the recorded styling over them. With
// Values, formulas paste their results.
func (e *Editor) CopyRange(from string, src grid.Range, to string, at grid.Range, o CopyOptions) (RangeResult, error) {
	if !e.fresh() {
		return RangeResult{}, errNotFresh
	}
	from, err := e.sheetName(from)
	if err != nil {
		return RangeResult{}, err
	}
	to, err = e.sheetName(to)
	if err != nil {
		return RangeResult{}, err
	}
	src = e.stretch(from, src)
	res := RangeResult{From: grid.FormatRef(from, src)}
	what := pasteAll
	switch {
	case o.Values:
		what = pasteWhat{contents: true, numFmt: true}
	case o.Formats:
		what = pasteWhat{formats: true}
	}
	if what.contents && !o.Values {
		if err := e.checkSpills(from, src, "copy"); err != nil {
			return res, err
		}
	}
	if cell, ok := e.threadIn(from, src); ok && what.notes {
		return res, fmt.Errorf("%s holds a comment thread (Excel 365) in %s, which sheets cannot copy; copy around it, or copy --values", grid.FormatRef(from, src), cell)
	}
	b, err := e.take(from, src)
	if err != nil {
		return res, err
	}
	if o.Values {
		if err := e.valuesOf(b); err != nil {
			return res, err
		}
	}
	targets, err := dest(src, at)
	if err != nil {
		return res, err
	}
	var dvs []*excelize.DataValidation
	var recs []Record
	if what.formats {
		if dvs, err = e.validationsIn(from, src); err != nil {
			return res, err
		}
		recs = e.recordsIn(from, src)
		if cfs, err := e.CondFormats(from); err == nil {
			unrecorded := 0
			for _, cf := range cfs {
				if !cf.Recorded && overlaps(sqrefRanges(cf.Range), []grid.Range{src}) {
					unrecorded++
				}
			}
			if unrecorded > 0 {
				res.Notes = append(res.Notes, fmt.Sprintf("%s not made by sheets cond stayed behind", countOf(unrecorded, "conditional format")))
			}
		}
	}
	for _, d := range targets {
		if d.R2 > grid.MaxRows || d.C2 > grid.MaxCols {
			return res, fmt.Errorf("copying %s to %s would run off the sheet", src, grid.CellName(d.R1, d.C1))
		}
		if what.contents {
			if t, row, ok := e.tableEdge(to, d); ok && row == t.ref.R1 {
				return res, fmt.Errorf("%s would overwrite row %d, the header row of Excel table %s, which names its columns; copy into its data rows", grid.FormatRef(to, d), row, t)
			}
			if err := e.checkSpills(to, d, "paste over"); err != nil {
				return res, err
			}
			if err := e.checkTableFormulas(b, from, to, d); err != nil {
				return res, err
			}
		}
		if _, err := e.checkMerges(to, d); err != nil {
			return res, err
		}
	}
	if err := e.unshareSheet(from); err != nil {
		return res, err
	}
	if err := e.unshareSheet(to); err != nil {
		return res, err
	}
	for _, d := range targets {
		dr, dc := d.R1-src.R1, d.C1-src.C1
		if what.formats {
			if err := e.unmergeIn(to, d); err != nil {
				return res, err
			}
		}
		if err := e.wipe(to, d, b.cells, what); err != nil {
			return res, err
		}
		for k, h := range b.cells {
			text := ""
			if h.formula != "" && !o.Values {
				text = formula.Shift(h.formula, dr, dc)
			}
			if err := e.put(to, d.R1+k[0], d.C1+k[1], h, text, what); err != nil {
				return res, err
			}
		}
		res.Cells += len(b.cells)
		if what.formats {
			for _, m := range b.merges {
				if err := e.f.MergeCell(to, grid.CellName(m.R1+dr, m.C1+dc), grid.CellName(m.R2+dr, m.C2+dc)); err != nil {
					return res, err
				}
				e.dropMerges()
			}
			if err := e.unrecordStyle(to, d); err != nil {
				return res, err
			}
			shift := func(f string) string { return formula.Shift(f, dr, dc) }
			if err := e.placeRecords(to, recs, dr, dc, shift); err != nil {
				return res, err
			}
			if len(dvs) > 0 {
				if err := e.dropValidation(to, d); err != nil {
					return res, err
				}
				if err := e.placeValidations(to, dvs, dr, dc, shift); err != nil {
					return res, err
				}
			}
		}
	}
	all := targets[0]
	for _, d := range targets[1:] {
		all = all.Union(d)
	}
	res.To = grid.FormatRef(to, all)
	return res, nil
}

// unmergeIn undoes the merges inside r.
func (e *Editor) unmergeIn(sheet string, r grid.Range) error {
	ms, err := e.f.GetMergeCells(sheet)
	if err != nil {
		return err
	}
	for _, m := range ms {
		mr, err := grid.ParseRange(m.GetStartAxis() + ":" + m.GetEndAxis())
		if err == nil && within([]grid.Range{mr}, r) {
			if err := e.f.UnmergeCell(sheet, m.GetStartAxis(), m.GetEndAxis()); err != nil {
				return err
			}
			e.dropMerges()
		}
	}
	return nil
}

// valuesOf turns a block's formulas into their results, as Paste Values
// does; a dynamic array's result is pasted whole.
func (e *Editor) valuesOf(b *block) error {
	s := e.wb.Sheet(b.sheet)
	if s == nil {
		return nil
	}
	// The results of dynamic arrays are values to paste too.
	for a, sp := range e.spillsOn(b.sheet) {
		if b.r.Intersect(sp).Empty() {
			continue
		}
		x := sp.Intersect(b.r)
		for row := x.R1; row <= x.R2; row++ {
			for col := x.C1; col <= x.C2; col++ {
				if row == a[0] && col == a[1] {
					continue
				}
				if c := s.At(row, col); c != nil && !c.IsEmpty() {
					k := [2]int{row - b.r.R1, col - b.r.C1}
					if b.cells[k] == nil {
						b.cells[k] = &taken{}
					}
					b.cells[k].cell = *c
				}
			}
		}
	}
	var unknown []string
	for k, h := range b.cells {
		if h.formula == "" {
			continue
		}
		row, col := b.r.R1+k[0], b.r.C1+k[1]
		c := e.wb.Value(s, row, col)
		if c.Pending {
			v, _, ok := e.Evaluate(b.sheet, row, col)
			if !ok {
				unknown = append(unknown, grid.CellName(row, col))
				continue
			}
			c = &v
		}
		h.cell = grid.Cell{Kind: c.Kind, Num: c.Num, Text: c.Text, Bool: c.Bool, Fmt: c.Fmt}
		h.formula, h.dynamic = "", false
	}
	if len(unknown) > 0 {
		sort.Strings(unknown)
		more := ""
		if len(unknown) > 5 {
			more = fmt.Sprintf(" and %d more", len(unknown)-5)
			unknown = unknown[:5]
		}
		return fmt.Errorf("the sheets calculator cannot compute %s%s, so their values are unknown until Excel computes them; copy the formulas instead", strings.Join(unknown, ", "), more)
	}
	return nil
}

// checkTableFormulas refuses pasting formulas that use table references
// (Sales[@Units]) outside their table, where they would mean something
// else.
func (e *Editor) checkTableFormulas(b *block, from, to string, d grid.Range) error {
	tables := e.tablesOn(to)
	for k, h := range b.cells {
		if h.formula == "" || !strings.Contains(h.formula, "[") || len(formula.DepsOf(h.formula).Tables) == 0 {
			continue
		}
		row, col := d.R1+k[0], d.C1+k[1]
		inside := false
		for _, t := range tables {
			if t.ref.Contains(row, col) {
				inside = true
			}
		}
		if !inside || !strings.EqualFold(from, to) {
			return fmt.Errorf("%s holds =%s, which uses table references and would land outside its table at %s; copy --values, or copy within the table",
				grid.FormatRef(from, grid.Range{R1: b.r.R1 + k[0], C1: b.r.C1 + k[1], R2: b.r.R1 + k[0], C2: b.r.C1 + k[1]}), h.formula, grid.CellName(row, col))
		}
	}
	return nil
}

// MoveRange moves cells as Excel's cut and paste does. Everything the
// cells hold goes with them: values, formulas (which keep reading the
// cells they read), formats, merges, notes, links, data validation and
// recorded styling. References anywhere in the workbook to the moved cells
// follow them; references to the cells pasted over become #REF!.
func (e *Editor) MoveRange(from string, src grid.Range, to string, at grid.Range) (RangeResult, error) {
	if !e.fresh() {
		return RangeResult{}, errNotFresh
	}
	from, err := e.sheetName(from)
	if err != nil {
		return RangeResult{}, err
	}
	to, err = e.sheetName(to)
	if err != nil {
		return RangeResult{}, err
	}
	src = e.stretch(from, src)
	res := RangeResult{From: grid.FormatRef(from, src)}
	if !(at.Rows() == 1 && at.Cols() == 1 || at.Rows() == src.Rows() && at.Cols() == src.Cols()) {
		return res, fmt.Errorf("%s is %d×%d but %s is %d×%d; give the top-left cell to move to", at, at.Rows(), at.Cols(), src, src.Rows(), src.Cols())
	}
	m := formula.Move{From: from, Src: src, To: to, DR: at.R1 - src.R1, DC: at.C1 - src.C1}
	dst := m.Dst()
	res.To = grid.FormatRef(to, dst)
	same := strings.EqualFold(from, to)
	if m.DR == 0 && m.DC == 0 && same {
		return res, nil
	}
	if dst.R2 > grid.MaxRows || dst.C2 > grid.MaxCols {
		return res, fmt.Errorf("moving %s to %s would run off the sheet", src, grid.CellName(dst.R1, dst.C1))
	}
	for _, x := range []struct {
		sheet string
		r     grid.Range
	}{{from, src}, {to, dst}} {
		for _, t := range e.tablesOn(x.sheet) {
			if !t.ref.Intersect(x.r).Empty() {
				return res, fmt.Errorf("%s reaches into Excel table %s; sheets does not move cells into, out of or within Excel tables (sort them with sheets sort)", grid.FormatRef(x.sheet, x.r), t)
			}
		}
	}
	if cell, ok := e.threadIn(from, src); ok {
		return res, fmt.Errorf("%s holds a comment thread (Excel 365) in %s, which sheets cannot move; move it in Excel", grid.FormatRef(from, src), cell)
	}
	if err := e.checkSpills(from, src, "move"); err != nil {
		return res, err
	}
	for a, sp := range e.spillsOn(to) {
		if same && src.Contains(a[0], a[1]) || sp.Intersect(dst).Empty() || sp.Rows()*sp.Cols() == 1 {
			continue
		}
		if !dst.Contains(a[0], a[1]) || !within([]grid.Range{sp}, dst) {
			return res, fmt.Errorf("%s would land on part of the result of the dynamic array formula in %s (it fills %s); move it elsewhere", grid.FormatRef(to, dst), grid.CellName(a[0], a[1]), sp)
		}
	}
	b, err := e.take(from, src)
	if err != nil {
		return res, err
	}
	ms, err := e.f.GetMergeCells(to)
	if err != nil {
		return res, err
	}
	for _, mc := range ms {
		mr, err := grid.ParseRange(mc.GetStartAxis() + ":" + mc.GetEndAxis())
		if err != nil || mr.Intersect(dst).Empty() || within([]grid.Range{mr}, dst) || same && within([]grid.Range{mr}, src) {
			continue
		}
		return res, fmt.Errorf("%s would land on part of the merged cells %s; unmerge them first", grid.FormatRef(to, dst), mr)
	}
	if err := e.unshareSheet(from); err != nil {
		return res, err
	}
	if err := e.unshareSheet(to); err != nil {
		return res, err
	}

	// What reads the cells: formulas, names, charts and conditional
	// formats follow them. Formulas reading a range the move took cells
	// out of keep reading the same range, as in Excel; say so.
	now, err := e.formulasNow()
	if err != nil {
		return res, err
	}
	var partial, broken []string
	rewritten := 0
	for _, f := range now {
		if strings.EqualFold(f.sheet, from) && src.Contains(f.row, f.col) {
			continue // moves with the block
		}
		if strings.EqualFold(f.sheet, to) && dst.Contains(f.row, f.col) {
			continue // pasted over
		}
		for _, u := range formula.RefUses(f.text) {
			sheet := u.Sheet
			if sheet == "" {
				sheet = f.sheet
			}
			if strings.EqualFold(sheet, from) && !u.Range.Intersect(src).Empty() && !within([]grid.Range{u.Range}, src) &&
				!(same && within([]grid.Range{dst}, u.Range)) {
				partial = append(partial, fmt.Sprintf("%s (=%s)", grid.FormatRef(f.sheet, grid.Range{R1: f.row, C1: f.col, R2: f.row, C2: f.col}), f.text))
				break
			}
		}
		text := formula.MoveRefs(f.text, f.sheet, f.sheet, m)
		if text == f.text {
			continue
		}
		if err := e.setFormula(f.sheet, grid.CellName(f.row, f.col), text); err != nil {
			return res, err
		}
		e.touch(f.sheet, f.row, f.col)
		if strings.Count(text, "#REF!") > strings.Count(f.text, "#REF!") {
			broken = append(broken, grid.FormatRef(f.sheet, grid.Range{R1: f.row, C1: f.col, R2: f.row, C2: f.col}))
		} else {
			rewritten++
		}
	}
	for _, n := range e.f.GetDefinedName() {
		if strings.HasPrefix(n.Name, "_xlnm.") || strings.HasPrefix(n.Name, recordPrefix) {
			continue
		}
		ref := formula.MoveRefs(strings.TrimPrefix(n.RefersTo, "="), "", "", m)
		if ref == strings.TrimPrefix(n.RefersTo, "=") {
			continue
		}
		if err := e.f.DeleteDefinedName(&excelize.DefinedName{Name: n.Name, Scope: n.Scope}); err != nil {
			return res, err
		}
		n.RefersTo = ref
		if err := e.f.SetDefinedName(&n); err != nil {
			return res, err
		}
	}
	e.patches.rangeRewrites = append(e.patches.rangeRewrites, func(f string) string { return formula.MoveRefs(f, "", "", m) })
	for home, sec := range e.patches.conds {
		sec.formulas(func(f string) string { return formula.MoveRefs(f, home, home, m) })
	}

	// Formats that belong to ranges: validation, recorded styling and
	// conditional formats.
	dvs, err := e.validationsIn(from, src)
	if err != nil {
		return res, err
	}
	recs := e.recordsIn(from, src)
	var conds []Record // cond entries, which follow on the same sheet
	var styles []Record
	for _, rec := range recs {
		if rec.Kind == "cond" {
			conds = append(conds, rec)
		} else {
			styles = append(styles, rec)
		}
	}
	cfs, err := e.CondFormats(from)
	if err != nil {
		return res, err
	}
	stayed := 0
	for _, cf := range cfs {
		rs := sqrefRanges(cf.Range)
		if !overlaps(rs, []grid.Range{src}) || cf.Banding {
			continue
		}
		if !within(rs, src) || !same && !cf.Recorded {
			stayed++
		}
	}
	if sec := e.patches.conds[from]; sec != nil && same {
		// Wholly within the block, they go with it.
		sec.each(func(el []byte) []byte {
			if bandRule.Match(el) {
				return el
			}
			return withSqref(el, func(sq string) string {
				rs := sqrefRanges(sq)
				if len(rs) == 0 || !within(rs, src) {
					return sq
				}
				var out []string
				for _, r := range rs {
					out = append(out, grid.Range{R1: r.R1 + m.DR, C1: r.C1 + m.DC, R2: r.R2 + m.DR, C2: r.C2 + m.DC}.String())
				}
				return strings.Join(out, " ")
			})
		})
	}
	if stayed > 0 {
		res.Notes = append(res.Notes, fmt.Sprintf("%s over %s stayed where %s (conditional formats move whole, within a sheet)",
			countOf(stayed, "conditional format"), src, map[bool]string{true: "it was", false: "they were"}[stayed == 1]))
	}

	// The cells: clear the dynamic arrays leaving, paste, then empty what
	// the block left.
	for a := range e.spillsOn(from) {
		if src.Contains(a[0], a[1]) {
			if err := e.clearSpill(from, a[0], a[1]); err != nil {
				return res, err
			}
		}
	}
	if err := e.unmergeIn(from, src); err != nil {
		return res, err
	}
	if err := e.unmergeIn(to, dst); err != nil {
		return res, err
	}
	if err := e.wipe(to, dst, b.cells, pasteAll); err != nil {
		return res, err
	}
	for k, h := range b.cells {
		text := ""
		if h.formula != "" {
			text = formula.MoveRefs(h.formula, from, to, m)
		}
		if err := e.put(to, dst.R1+k[0], dst.C1+k[1], h, text, pasteAll); err != nil {
			return res, err
		}
	}
	left := []grid.Range{src}
	if same {
		left = subtract(src, dst)
	}
	for _, r := range left {
		if err := e.wipe(from, r, nil, pasteAll); err != nil {
			return res, err
		}
	}
	for _, mg := range b.merges {
		if err := e.f.MergeCell(to, grid.CellName(mg.R1+m.DR, mg.C1+m.DC), grid.CellName(mg.R2+m.DR, mg.C2+m.DC)); err != nil {
			return res, err
		}
		e.dropMerges()
	}
	res.Cells = len(b.cells)

	fix := func(f string) string { return formula.MoveRefs(f, from, to, m) }
	if err := e.dropValidation(from, src); err != nil {
		return res, err
	}
	if len(dvs) > 0 {
		if err := e.dropValidation(to, dst); err != nil {
			return res, err
		}
		if err := e.placeValidations(to, dvs, m.DR, m.DC, fix); err != nil {
			return res, err
		}
	}
	if err := e.unrecordStyle(from, src); err != nil {
		return res, err
	}
	if err := e.unrecordStyle(to, dst); err != nil {
		return res, err
	}
	if err := e.placeRecords(to, styles, m.DR, m.DC, fix); err != nil {
		return res, err
	}
	if same {
		for _, rec := range conds {
			if err := setAreas(e.f, rec, []grid.Range{{R1: rec.Range().R1 + m.DR, C1: rec.Range().C1 + m.DC, R2: rec.Range().R2 + m.DR, C2: rec.Range().C2 + m.DC}}); err != nil {
				return res, err
			}
		}
		if err := e.rewriteCondRecords(func(f, home string) string { return formula.MoveRefs(f, home, home, m) }); err != nil {
			return res, err
		}
	} else {
		// Rules already recorded first: the ones re-made on the other
		// sheet take their formulas from fix.
		if err := e.rewriteCondRecords(func(f, home string) string { return formula.MoveRefs(f, home, home, m) }); err != nil {
			return res, err
		}
		if err := e.placeRecords(to, conds, m.DR, m.DC, fix); err != nil {
			return res, err
		}
		for _, rec := range conds {
			r := rec.Range()
			if _, err := e.ClearCond(from, &r); err != nil {
				return res, err
			}
		}
	}

	if rewritten > 0 {
		res.Notes = append(res.Notes, fmt.Sprintf("%s elsewhere now %s the cells where they went", countOf(rewritten, "formula"), map[bool]string{true: "reads", false: "read"}[rewritten == 1]))
	}
	if len(broken) > 0 {
		sort.Strings(broken)
		e.warnings = append(e.warnings, fmt.Sprintf("%s now %s #REF!: %s the cells the move pasted over, as in Excel: %s",
			countOf(len(broken), "formula"), map[bool]string{true: "contains", false: "contain"}[len(broken) == 1],
			map[bool]string{true: "it read", false: "they read"}[len(broken) == 1], listOf(broken, 10)))
	}
	if len(partial) > 0 {
		sort.Strings(partial)
		list := partial[:min(len(partial), 5)]
		if len(partial) > 5 {
			list = append(list, fmt.Sprintf("and %d more", len(partial)-5))
		}
		res.Notes = append(res.Notes, fmt.Sprintf("%s %s a range the move took cells out of, and still %s that range, as in Excel: %s",
			countOf(len(partial), "formula"), map[bool]string{true: "reads", false: "read"}[len(partial) == 1], map[bool]string{true: "reads", false: "read"}[len(partial) == 1], strings.Join(list, "; ")))
	}
	return res, nil
}

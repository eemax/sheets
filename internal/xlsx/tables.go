package xlsx

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"html"
	"regexp"
	"strconv"
	"strings"

	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
	"github.com/xuri/excelize/v2"
)

// Excel tables (Insert › Table) are parts of their own: a range whose
// first row names the columns, perhaps a totals row at the bottom. excelize
// does not keep them right through edits: it skips tables stored under an
// absolute part path (as openpyxl writes them) and resets totals rows. So
// sheets rebuilds every table part from the file's original XML when
// saving. Its ranges follow moved rows and columns through the same
// rewrites as charts. Its column names follow the header cells, as Excel
// requires.

// tableSource is a table as the file had it when opened.
type tableSource struct {
	sheet string
	tablePart
}

// tableNow is where a table stands after the edits so far.
type tableNow struct {
	name    string
	sheet   string
	ref     grid.Range
	headers int
	totals  int // 0 once its totals row is deleted
	columns []string
}

// columnNames reads the column names of a table part, in order.
func columnNames(data []byte) []string {
	var out []string
	for _, tag := range tableColumn.FindAll(data, -1) {
		if m := nameAttr.FindSubmatch(tag); m != nil {
			out = append(out, html.UnescapeString(string(m[2])))
		}
	}
	return out
}

// calculatedFormulas reads the formulas of a table part's calculated
// columns, one per column ("" for the others); nil when there are none.
func calculatedFormulas(data []byte) []string {
	var out []string
	found := false
	for _, m := range tableColumnElem.FindAllSubmatch(data, -1) {
		f := ""
		if c := calculatedFormula.FindSubmatch(m[1]); c != nil {
			f, found = strings.TrimPrefix(html.UnescapeString(string(c[1])), "="), true
		}
		out = append(out, f)
	}
	if !found {
		return nil
	}
	return out
}

// formulaTable is the table as structured references see it.
func (t tableNow) formulaTable() formula.Table {
	return formula.Table{Name: t.name, Sheet: t.sheet, Ref: t.ref, Headers: t.headers, Totals: t.totals, Columns: t.columns}
}

// locate follows a range of sheet through the range rewrites: the sheet it
// is now on, and where. ok is false when the range was deleted.
func locate(sheet string, rg grid.Range, rewrites []func(string) string) (string, grid.Range, bool) {
	f := ChartRef(sheet, rg)
	for _, fn := range rewrites {
		f = fn(f)
	}
	i := strings.LastIndex(f, "!")
	if i < 0 {
		return "", grid.Range{}, false
	}
	moved, err := grid.ParseRange(strings.ReplaceAll(f[i+1:], "$", ""))
	if err != nil {
		return "", grid.Range{}, false
	}
	name := f[:i]
	if strings.HasPrefix(name, "'") && strings.HasSuffix(name, "'") && len(name) >= 2 {
		name = strings.ReplaceAll(name[1:len(name)-1], "''", "'")
	}
	return name, moved, true
}

// now locates the table after the rewrites.
func (t tableSource) now(rewrites []func(string) string) (tableNow, bool) {
	sheet, ref, ok := locate(t.sheet, t.ref, rewrites)
	if !ok {
		return tableNow{}, false
	}
	tn := tableNow{name: t.name, sheet: sheet, ref: ref, headers: t.headers, columns: columnNames(t.xml)}
	if t.totals > 0 {
		totals := t.ref
		totals.R1 = totals.R2 - t.totals + 1
		if _, _, ok := locate(t.sheet, totals, rewrites); ok {
			tn.totals = t.totals
		}
	}
	return tn, true
}

// ExtendTable makes the charts and the Excel table of a table whose data
// is body take in n rows written right below it: their ranges within its
// columns that end at its last row grow by n. Call it after inserting the
// rows, when there was something below to move.
func (e *Editor) ExtendTable(sheet string, body grid.Range, n int) {
	e.patches.rangeRewrites = append(e.patches.rangeRewrites, func(f string) string {
		return formula.ExtendRangesIn(f, "", sheet, body.R2, body.R2+n, body.C1, body.C2)
	})
	// Formulas reading a column of the table whole take in the new rows,
	// as Excel's references to a table's columns do.
	if all, err := e.formulasNow(); err == nil {
		for _, f := range all {
			grown := formula.GrowRanges(f.text, f.sheet, sheet, body.R1, body.R2, body.R2+n, body.C1, body.C2)
			if grown == f.text {
				continue
			}
			if err := e.setFormula(f.sheet, grid.CellName(f.row, f.col), grown); err != nil {
				e.warnings = append(e.warnings, "could not extend "+grid.FormatRef(f.sheet, grid.Range{R1: f.row, C1: f.col, R2: f.row, C2: f.col})+": "+err.Error())
				continue
			}
			e.touch(f.sheet, f.row, f.col)
		}
	}
	if err := e.growRecords(sheet, body, n); err != nil {
		e.warnings = append(e.warnings, "could not extend recorded styling: "+err.Error())
	}
}

// tablesOn lists the Excel tables on a sheet where they stand now.
func (e *Editor) tablesOn(sheet string) []tableNow {
	var out []tableNow
	for _, t := range e.patches.tables {
		if tn, ok := t.now(e.patches.rangeRewrites); ok && strings.EqualFold(tn.sheet, sheet) {
			out = append(out, tn)
		}
	}
	return out
}

func (t tableNow) String() string { return fmt.Sprintf("%q (%s)", t.name, t.ref) }

// checkTableRows refuses deleting rows a table cannot do without: its
// header row, or all of its data rows.
func (e *Editor) checkTableRows(sheet string, from, to int) error {
	for _, t := range e.tablesOn(sheet) {
		if t.headers > 0 && from <= t.ref.R1 && to >= t.ref.R1 {
			return fmt.Errorf("row %d is the header row of Excel table %s; sheets cannot delete it (Excel would remove the table)", t.ref.R1, t)
		}
		first, last := t.ref.R1+t.headers, t.ref.R2-t.totals
		if first <= last && from <= first && to >= last {
			return fmt.Errorf("rows %d:%d are all the data rows of Excel table %s; a table keeps at least one", first, last, t)
		}
	}
	return nil
}

// checkTableCols refuses adding or removing columns inside a table: its
// columns are named parts of it, which sheets does not add or remove.
func (e *Editor) checkTableCols(sheet string, from, to int, insert bool) error {
	for _, t := range e.tablesOn(sheet) {
		if insert && from > t.ref.C1 && from <= t.ref.C2 {
			return fmt.Errorf("column %s is inside Excel table %s; sheets cannot add table columns (insert at %s or %s, outside it)",
				grid.ColName(from), t, grid.ColName(t.ref.C1), grid.ColName(t.ref.C2+1))
		}
		if !insert && from <= t.ref.C2 && to >= t.ref.C1 {
			return fmt.Errorf("columns %s:%s overlap Excel table %s; sheets cannot remove table columns", grid.ColName(from), grid.ColName(to), t)
		}
	}
	return nil
}

var (
	refAttr     = regexp.MustCompile(`(\sref=")([^"]*)(")`)
	totalsAttr  = regexp.MustCompile(`\stotalsRowCount="\d+"`)
	tableColumn = regexp.MustCompile(`<tableColumn\b[^>]*>`)
	// A whole tableColumn element, with what it holds.
	tableColumnElem   = regexp.MustCompile(`(?s)<tableColumn\b(?:[^>]*/>|[^>]*>(.*?)</tableColumn>)`)
	calculatedFormula = regexp.MustCompile(`(?s)<calculatedColumnFormula\b[^>]*>([^<]*)</calculatedColumnFormula>`)
	nameAttr          = regexp.MustCompile(`(\sname=")([^"]*)(")`)
	// Formulas a table part keeps: of calculated columns and totals.
	tableFormula = regexp.MustCompile(`(<(?:calculatedColumnFormula|totalsRowFormula)\b[^>]*>)([^<]*)(</[a-zA-Z]+>)`)
)

// renameColumnsIn applies the column renames to the cell formulas of a
// worksheet part.
func renameColumnsIn(b []byte, sheet string, p *patches) []byte {
	return rewriteCells(b, func(string) bool { return true }, func(m []byte, ref string) []byte {
		if !bytes.Contains(m, []byte("<f")) || !bytes.Contains(m, []byte("[")) {
			return m
		}
		r, c, _ := grid.ParseCell(ref)
		return rewriteFormulas(m, func(f string) string {
			return p.renameColumns(f, func(cr *columnRename) bool {
				return strings.EqualFold(cr.sheet, sheet) && cr.ref.Contains(r, c)
			})
		})
	})
}

// fixTables rebuilds the table parts for saving: ranges moved with the
// rows and columns, column names taken from the header cells. Header cells
// left empty or repeating a name get one, as Excel does.
func (e *Editor) fixTables() error {
	e.patches.tableXML = map[string][]byte{}
	e.patches.columnRenames = nil
	names := map[string]string{} // part → table name
	for _, t := range e.patches.tables {
		tn, ok := t.now(e.patches.rangeRewrites)
		if i, err := e.f.GetSheetIndex(tn.sheet); !ok || err != nil || i < 0 {
			continue // gone with its sheet; excelize has dropped it
		}
		data := t.xml
		failed := false
		data = refAttr.ReplaceAllFunc(data, func(m []byte) []byte {
			sub := refAttr.FindSubmatch(m)
			rg, err := grid.ParseRange(string(sub[2]))
			if err != nil {
				return m
			}
			_, moved, ok := locate(t.sheet, rg, e.patches.rangeRewrites)
			if !ok {
				failed = true
				return m
			}
			return []byte(string(sub[1]) + moved.String() + string(sub[3]))
		})
		if failed {
			continue
		}
		if tn.totals == 0 && t.totals > 0 {
			data = totalsAttr.ReplaceAll(data, nil)
		}
		if tn.headers > 0 {
			var err error
			if data, err = e.nameColumns(tn, data); err != nil {
				return err
			}
		}
		e.patches.tableXML[strings.ToLower(t.part)] = data
		names[strings.ToLower(t.part)] = t.name
	}

	// Formulas kept in table parts (calculated columns, totals) follow
	// renamed columns too; formulas in cells follow when saving.
	if len(e.patches.columnRenames) > 0 {
		for part, data := range e.patches.tableXML {
			e.patches.tableXML[part] = tableFormula.ReplaceAllFunc(data, func(m []byte) []byte {
				g := tableFormula.FindSubmatch(m)
				text := html.UnescapeString(string(g[2]))
				out := e.patches.renameColumns(text, func(cr *columnRename) bool { return strings.EqualFold(cr.table, names[part]) })
				if out == text {
					return m
				}
				return append(append(append([]byte{}, g[1]...), xmlEscape(out)...), g[3]...)
			})
		}
	}
	return nil
}

// columnRename is a table's columns renamed after their headers.
type columnRename struct {
	sheet, table string
	ref          grid.Range
	renames      map[string]string // lower-cased old name → new
	pairs        []string          // "Units" → "Qty", for the output
	used         bool              // a formula named one of them
}

// renameColumns applies the column renames to a formula; inside tells
// whether the formula sits inside a table (where [Units] means its column).
func (p *patches) renameColumns(f string, inside func(*columnRename) bool) string {
	for _, cr := range p.columnRenames {
		if g := formula.RenameTableColumns(f, cr.table, inside(cr), cr.renames); g != f {
			f, cr.used = g, true
		}
	}
	return f
}

// renameNotes describes the renamed columns, after saving has applied them.
func (p *patches) renameNotes() []string {
	var out []string
	for _, cr := range p.columnRenames {
		note := fmt.Sprintf("columns of Excel table %q renamed after their headers: %s", cr.table, strings.Join(cr.pairs, ", "))
		if cr.used {
			note += "; formulas naming them now use the new names, as Excel would"
		}
		out = append(out, note)
	}
	return out
}

// nameColumns sets each table column's name to its header cell's text.
func (e *Editor) nameColumns(t tableNow, data []byte) ([]byte, error) {
	cols := tableColumn.FindAllIndex(data, -1)
	if len(cols) != t.ref.Cols() {
		return data, nil // not the shape sheets knows; leave it be
	}
	seen := map[string]bool{}
	names := make([]string, len(cols))
	for i := range cols {
		cell := grid.CellName(t.ref.R1, t.ref.C1+i)
		text, err := e.f.GetCellValue(t.sheet, cell, excelize.Options{RawCellValue: true})
		if err != nil {
			return nil, err
		}
		if f, _ := e.f.GetCellFormula(t.sheet, cell); f != "" {
			text = "" // header cells hold text, not formulas
		}
		switch typ, _ := e.f.GetCellType(t.sheet, cell); {
		case text == "" || seen[strings.ToLower(text)]:
			// As Excel names them: by position, kept unique.
			n := i + 1
			for seen[strings.ToLower("Column"+strconv.Itoa(n))] {
				n++
			}
			was := text
			text = "Column" + strconv.Itoa(n)
			if err := setText(e.f, t.sheet, cell, text); err != nil {
				return nil, err
			}
			e.touch(t.sheet, t.ref.R1, t.ref.C1+i)
			why := "was empty"
			if was != "" {
				why = fmt.Sprintf("repeated %q", was)
			}
			e.warnings = append(e.warnings, fmt.Sprintf("header %s of Excel table %q %s; it is now %q, as Excel would name it", cell, t.name, why, text))
		case typ != excelize.CellTypeSharedString && typ != excelize.CellTypeInlineString:
			// A number or date as a header: Excel keeps header cells as text.
			if err := setText(e.f, t.sheet, cell, text); err != nil {
				return nil, err
			}
			e.touch(t.sheet, t.ref.R1, t.ref.C1+i)
		}
		seen[strings.ToLower(text)] = true
		names[i] = text
	}

	// Rebuild from the end so earlier offsets stay valid.
	out := append([]byte(nil), data...)
	cr := &columnRename{sheet: t.sheet, table: t.name, ref: t.ref, renames: map[string]string{}}
	for i := len(cols) - 1; i >= 0; i-- {
		tag := out[cols[i][0]:cols[i][1]]
		sub := nameAttr.FindSubmatchIndex(tag)
		if sub == nil {
			continue
		}
		old := html.UnescapeString(string(tag[sub[4]:sub[5]]))
		want := encodeColumnName(names[i])
		if old == want {
			continue
		}
		cr.renames[strings.ToLower(old)] = want
		cr.pairs = append([]string{fmt.Sprintf("%q → %q", old, want)}, cr.pairs...)
		newTag := string(tag[:sub[4]]) + attrEscape(want) + string(tag[sub[5]:])
		out = append(out[:cols[i][0]], append([]byte(newTag), out[cols[i][1]:]...)...)
	}
	if len(cr.renames) > 0 {
		e.patches.columnRenames = append(e.patches.columnRenames, cr)
	}
	return out, nil
}

// encodeColumnName writes line breaks and tabs the way Excel stores them
// in table column names.
func encodeColumnName(s string) string {
	return strings.NewReplacer("\r", "_x000d_", "\n", "_x000a_", "\t", "_x0009_").Replace(s)
}

func xmlEscape(s string) []byte {
	var b bytes.Buffer
	xml.EscapeText(&b, []byte(s))
	return b.Bytes()
}

func attrEscape(s string) string {
	return strings.NewReplacer("&", "&amp;", "<", "&lt;", ">", "&gt;", `"`, "&quot;").Replace(s)
}

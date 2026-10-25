package cli

import (
	"fmt"
	"math"
	"sort"
	"strconv"
	"strings"
	"time"

	"github.com/dop251/goja"
	"github.com/eemax/sheets/internal/analyze"
	"github.com/eemax/sheets/internal/book"
	"github.com/eemax/sheets/internal/engine"
	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/input"
)

// The ExcelScript objects: Workbook, Worksheet, Range, Table and their
// parts, as Office Scripts names them (getValues, setFormula, ...). Go
// methods become JS methods with a lower-case first letter. Positions in
// the API count from 0, as in Office Scripts.

// ---- Workbook ----

type jsWorkbook struct{ s *scriptSession }

func (w *jsWorkbook) sheetObj(name string) *jsWorksheet {
	return &jsWorksheet{w.s, w.s.ref(name)}
}

func (w *jsWorkbook) GetWorksheet(name string) goja.Value {
	for _, n := range w.s.sheetNames() {
		if strings.EqualFold(n, name) {
			return w.s.vm.ToValue(w.sheetObj(n))
		}
	}
	return goja.Undefined()
}

func (w *jsWorkbook) GetWorksheets() goja.Value {
	var out []any
	for _, n := range w.s.sheetNames() {
		out = append(out, w.sheetObj(n))
	}
	return w.s.array(out)
}

func (w *jsWorkbook) GetActiveWorksheet() *jsWorksheet { return w.GetFirstWorksheet(true) }

func (w *jsWorkbook) GetFirstWorksheet(visibleOnly bool) *jsWorksheet {
	wb := w.s.state()
	sh := wb.Sheets[0]
	if visibleOnly {
		sh = wb.FirstVisible()
	}
	return w.sheetObj(sh.Name)
}

func (w *jsWorkbook) GetLastWorksheet(visibleOnly bool) *jsWorksheet {
	wb := w.s.state()
	for i := len(wb.Sheets) - 1; i >= 0; i-- {
		if !visibleOnly || !wb.Sheets[i].Hidden {
			return w.sheetObj(wb.Sheets[i].Name)
		}
	}
	return w.sheetObj(wb.Sheets[0].Name)
}

func (w *jsWorkbook) AddWorksheet(name string) *jsWorksheet {
	names := w.s.sheetNames()
	taken := func(n string) bool {
		for _, m := range names {
			if strings.EqualFold(m, n) {
				return true
			}
		}
		return false
	}
	if name == "" {
		for n := len(names) + 1; ; n++ {
			if name = fmt.Sprintf("Sheet%d", n); !taken(name) {
				break
			}
		}
	}
	if taken(name) {
		w.s.fail("a worksheet named %q already exists", name)
	}
	w.s.must(w.s.editor().AddSheet(name, "", ""))
	w.s.done = append(w.s.done, "sheet "+name+" added")
	return w.sheetObj(name)
}

func (w *jsWorkbook) GetTable(name string) goja.Value {
	if t, ok := w.s.table(name); ok {
		return w.s.vm.ToValue(t)
	}
	return goja.Undefined()
}

func (w *jsWorkbook) GetTables() goja.Value {
	var out []any
	for _, t := range w.s.tables() {
		out = append(out, t)
	}
	return w.s.array(out)
}

func (w *jsWorkbook) GetNamedItem(name string) goja.Value {
	for _, n := range w.s.state().Names {
		if strings.EqualFold(n.Name, name) && !strings.HasPrefix(n.Name, "_") {
			return w.s.vm.ToValue(&jsNamedItem{w.s, n})
		}
	}
	return goja.Undefined()
}

func (w *jsWorkbook) GetNamedItems() goja.Value {
	var out []any
	for _, n := range w.s.state().Names {
		if !strings.HasPrefix(n.Name, "_") {
			out = append(out, &jsNamedItem{w.s, n})
		}
	}
	return w.s.array(out)
}

func (w *jsWorkbook) GetName() string { return baseName(w.s.file) }

func baseName(path string) string {
	if i := strings.LastIndexAny(path, `/\`); i >= 0 {
		return path[i+1:]
	}
	return path
}

// ---- Named items ----

type jsNamedItem struct {
	s *scriptSession
	n grid.DefinedName
}

func (n *jsNamedItem) GetName() string    { return n.n.Name }
func (n *jsNamedItem) GetFormula() string { return "=" + strings.TrimPrefix(n.n.Ref, "=") }
func (n *jsNamedItem) GetRange() goja.Value {
	sheet, rg, err := analyze.Resolve(n.s.state(), strings.TrimPrefix(n.n.Ref, "="))
	if err != nil || rg == nil {
		return goja.Undefined()
	}
	return n.s.vm.ToValue(&jsRange{n.s, n.s.ref(sheet), *rg})
}

// ---- Worksheet ----

type jsWorksheet struct {
	s   *scriptSession
	ref *sheetRef
}

func (w *jsWorksheet) ToJSON() string { return w.ref.name }

func (w *jsWorksheet) GetName() string { return w.s.sheetName(w.ref) }

func (w *jsWorksheet) SetName(name string) {
	old := w.s.sheetName(w.ref)
	if old == name {
		return
	}
	w.s.must(w.s.editor().RenameSheet(old, name))
	w.s.renamed(old, name)
	w.s.done = append(w.s.done, fmt.Sprintf("sheet %s renamed %s", old, name))
}

func (w *jsWorksheet) GetPosition() int {
	for i, sh := range w.s.state().Sheets {
		if sh == w.s.sheet(w.ref) {
			return i
		}
	}
	return -1
}

func (w *jsWorksheet) Delete() {
	name := w.s.sheetName(w.ref)
	w.s.must(w.s.editor().DeleteSheet(name))
	w.s.done = append(w.s.done, "sheet "+name+" deleted")
}

func (w *jsWorksheet) Activate() {}

func (w *jsWorksheet) GetVisibility() string {
	if w.s.sheet(w.ref).Hidden {
		return "Hidden"
	}
	return "Visible"
}

func (w *jsWorksheet) SetVisibility(v string) {
	name := w.s.sheetName(w.ref)
	w.s.must(w.s.editor().SetVisible(name, strings.EqualFold(v, "Visible")))
}

func (w *jsWorksheet) Copy(position string, relativeTo *jsWorksheet) *jsWorksheet {
	name := w.s.sheetName(w.ref)
	to := name + " (2)"
	for n := 3; w.s.hasSheet(to); n++ {
		to = fmt.Sprintf("%s (%d)", name, n)
	}
	notes, err := w.s.editor().CopySheet(name, to)
	w.s.must(err)
	w.s.warnings = append(w.s.warnings, notes...)
	w.s.done = append(w.s.done, fmt.Sprintf("sheet %s copied to %s", name, to))
	return &jsWorksheet{w.s, w.s.ref(to)}
}

// GetRange takes an address on this sheet: "A1:C9", "B:B", "3:5"; none
// is the whole sheet.
func (w *jsWorksheet) GetRange(address string) *jsRange {
	if address == "" {
		return &jsRange{w.s, w.ref, grid.Range{R1: 1, C1: 1, R2: grid.MaxRows, C2: grid.MaxCols}}
	}
	if i := strings.LastIndex(address, "!"); i >= 0 {
		if sheet := strings.ReplaceAll(strings.Trim(address[:i], "'"), "''", "'"); !strings.EqualFold(sheet, w.ref.name) {
			w.s.fail("getRange(%q): the address is on another sheet; use workbook.getWorksheet(%q).getRange(...)", address, sheet)
		}
		address = address[i+1:]
	}
	rg, err := grid.ParseRange(address)
	if err != nil {
		w.s.fail("getRange(%q): %v", address, err)
	}
	return &jsRange{w.s, w.ref, rg}
}

func (w *jsWorksheet) GetRangeByIndexes(row, col, rows, cols int) *jsRange {
	if row < 0 || col < 0 || rows < 1 || cols < 1 {
		w.s.fail("getRangeByIndexes(%d, %d, %d, %d): positions count from 0, sizes from 1", row, col, rows, cols)
	}
	return &jsRange{w.s, w.ref, grid.Range{R1: row + 1, C1: col + 1, R2: row + rows, C2: col + cols}}
}

func (w *jsWorksheet) GetCell(row, col int) *jsRange { return w.GetRangeByIndexes(row, col, 1, 1) }

func (w *jsWorksheet) GetUsedRange(valuesOnly bool) goja.Value {
	u := w.s.sheet(w.ref).Used()
	if u.Empty() {
		return goja.Undefined()
	}
	return w.s.vm.ToValue(&jsRange{w.s, w.ref, u})
}

func (w *jsWorksheet) GetTables() goja.Value {
	var out []any
	name := w.s.sheet(w.ref).Name
	for _, t := range w.s.tables() {
		if strings.EqualFold(t.sheet().Name, name) {
			out = append(out, t)
		}
	}
	return w.s.array(out)
}

func (w *jsWorksheet) GetFreezePanes() *jsFreeze { return &jsFreeze{w} }

// ---- Freeze panes ----

type jsFreeze struct{ w *jsWorksheet }

func (f *jsFreeze) freeze(row, col int) {
	name := f.w.s.sheetName(f.w.ref)
	at := grid.Range{R1: row + 1, C1: col + 1, R2: row + 1, C2: col + 1}
	f.w.s.must(f.w.s.editor().View(name, book.ViewOptions{Freeze: &at}))
}

func (f *jsFreeze) FreezeRows(n int)    { f.freeze(n, 0) }
func (f *jsFreeze) FreezeColumns(n int) { f.freeze(0, n) }
func (f *jsFreeze) Unfreeze()           { f.freeze(0, 0) }
func (f *jsFreeze) FreezeAt(r *jsRange) { f.freeze(r.r.R2, r.r.C2) }

// ---- Range ----

type jsRange struct {
	s     *scriptSession
	sheet *sheetRef
	r     grid.Range
}

func (r *jsRange) ToJSON() string     { return r.GetAddress() }
func (r *jsRange) GetAddress() string { return grid.FormatRef(r.sheet.name, r.r) }
func (r *jsRange) GetAddressLocal() string {
	return r.GetAddress()
}
func (r *jsRange) GetWorksheet() *jsWorksheet { return &jsWorksheet{r.s, r.sheet} }
func (r *jsRange) GetRowIndex() int           { return r.r.R1 - 1 }
func (r *jsRange) GetColumnIndex() int        { return r.r.C1 - 1 }
func (r *jsRange) GetRowCount() int           { return r.r.Rows() }
func (r *jsRange) GetColumnCount() int        { return r.r.Cols() }
func (r *jsRange) GetCellCount() int          { return r.r.Rows() * r.r.Cols() }

func (r *jsRange) sub(rg grid.Range) *jsRange {
	if rg.R1 < 1 || rg.C1 < 1 || rg.R2 > grid.MaxRows || rg.C2 > grid.MaxCols || rg.Empty() {
		r.s.fail("%s: the range would leave the sheet", r.GetAddress())
	}
	return &jsRange{r.s, r.sheet, rg}
}

func (r *jsRange) GetCell(row, col int) *jsRange {
	return r.sub(grid.Range{R1: r.r.R1 + row, C1: r.r.C1 + col, R2: r.r.R1 + row, C2: r.r.C1 + col})
}
func (r *jsRange) GetRow(i int) *jsRange {
	return r.sub(grid.Range{R1: r.r.R1 + i, C1: r.r.C1, R2: r.r.R1 + i, C2: r.r.C2})
}
func (r *jsRange) GetColumn(i int) *jsRange {
	return r.sub(grid.Range{R1: r.r.R1, C1: r.r.C1 + i, R2: r.r.R2, C2: r.r.C1 + i})
}
func (r *jsRange) GetOffsetRange(rows, cols int) *jsRange {
	return r.sub(grid.Range{R1: r.r.R1 + rows, C1: r.r.C1 + cols, R2: r.r.R2 + rows, C2: r.r.C2 + cols})
}
func (r *jsRange) GetResizedRange(rows, cols int) *jsRange {
	return r.sub(grid.Range{R1: r.r.R1, C1: r.r.C1, R2: r.r.R2 + rows, C2: r.r.C2 + cols})
}
func (r *jsRange) GetEntireRow() *jsRange {
	return r.sub(grid.Range{R1: r.r.R1, C1: 1, R2: r.r.R2, C2: grid.MaxCols})
}
func (r *jsRange) GetEntireColumn() *jsRange {
	return r.sub(grid.Range{R1: 1, C1: r.r.C1, R2: grid.MaxRows, C2: r.r.C2})
}
func (r *jsRange) GetLastCell() *jsRange {
	b := r.bounded()
	return r.sub(grid.Range{R1: b.R2, C1: b.C2, R2: b.R2, C2: b.C2})
}
func (r *jsRange) GetLastRow() *jsRange {
	b := r.bounded()
	return r.sub(grid.Range{R1: b.R2, C1: b.C1, R2: b.R2, C2: b.C2})
}
func (r *jsRange) GetLastColumn() *jsRange {
	b := r.bounded()
	return r.sub(grid.Range{R1: b.R1, C1: b.C2, R2: b.R2, C2: b.C2})
}

// GetUsedRange is the part of the range holding values.
func (r *jsRange) GetUsedRange(valuesOnly bool) goja.Value {
	u := r.s.sheet(r.sheet).Used().Intersect(r.r)
	if u.Empty() {
		return goja.Undefined()
	}
	return r.s.vm.ToValue(r.sub(u))
}

// GetSurroundingRegion is the block of cells around the range's first
// cell, as Ctrl+A finds it (Excel's current region).
func (r *jsRange) GetSurroundingRegion() *jsRange {
	sh := r.s.sheet(r.sheet)
	for _, reg := range analyze.Regions(r.s.state(), sh) {
		if reg.Range.Contains(r.r.R1, r.r.C1) {
			return r.sub(reg.Range)
		}
	}
	return r.sub(grid.Range{R1: r.r.R1, C1: r.r.C1, R2: r.r.R1, C2: r.r.C1})
}

// bounded cuts whole rows and columns down to the used range, for reading.
func (r *jsRange) bounded() grid.Range {
	b := r.r
	if b.R2 == grid.MaxRows || b.C2 == grid.MaxCols {
		u := r.s.sheet(r.sheet).Used()
		if u.Empty() {
			u = grid.Range{R1: 1, C1: 1, R2: 1, C2: 1}
		}
		if b.R2 == grid.MaxRows {
			b.R2 = max(u.R2, b.R1)
		}
		if b.C2 == grid.MaxCols {
			b.C2 = max(u.C2, b.C1)
		}
	}
	return b
}

// readLimit keeps a read of a whole sheet's worth of cells in bounds.
const readLimit = 5_000_000

func (r *jsRange) cells(fn func(c *grid.Cell, row, col int) any) goja.Value {
	b := r.bounded()
	if b.Rows()*b.Cols() > readLimit {
		r.s.fail("%s holds %d cells; read it in parts", r.GetAddress(), b.Rows()*b.Cols())
	}
	wb := r.s.state()
	sh := r.s.sheet(r.sheet)
	rows := make([]any, 0, b.Rows())
	for row := b.R1; row <= b.R2; row++ {
		line := make([]any, 0, b.Cols())
		for col := b.C1; col <= b.C2; col++ {
			line = append(line, fn(wb.Value(sh, row, col), row, col))
		}
		rows = append(rows, r.s.array(line))
	}
	return r.s.array(rows)
}

// value is a cell's value as Office Scripts gives it.
func (r *jsRange) value(c *grid.Cell, row, col int) any {
	switch {
	case c == nil:
		return ""
	case c.Pending:
		r.s.unknown[grid.FormatRef(r.sheet.name, grid.Range{R1: row, C1: col, R2: row, C2: col})] = true
		return nil
	}
	switch c.Kind {
	case grid.Number:
		return c.Num
	case grid.Bool:
		return c.Bool
	case grid.Text, grid.Error:
		return c.Text
	}
	return ""
}

func (r *jsRange) GetValues() goja.Value { return r.cells(r.value) }
func (r *jsRange) GetValue() any {
	return r.value(r.s.state().Value(r.s.sheet(r.sheet), r.r.R1, r.r.C1), r.r.R1, r.r.C1)
}

func (r *jsRange) formulaOf(c *grid.Cell, row, col int) any {
	if c != nil && c.Formula != "" {
		return "=" + c.Formula
	}
	return r.value(c, row, col)
}

func (r *jsRange) GetFormulas() goja.Value { return r.cells(r.formulaOf) }
func (r *jsRange) GetFormula() any {
	return r.formulaOf(r.s.state().Value(r.s.sheet(r.sheet), r.r.R1, r.r.C1), r.r.R1, r.r.C1)
}

// text is a cell as Excel shows it, by its number format.
func (r *jsRange) text(c *grid.Cell, row, col int) any {
	if c == nil {
		return ""
	}
	if c.Pending {
		r.value(c, row, col)
		return ""
	}
	wb := r.s.state()
	switch c.Kind {
	case grid.Number:
		f := wb.FormatOf(c)
		code := f.Code
		if code == "" {
			code = "General"
		}
		return engine.FormatNumber(c.Num, code, wb.Date1904)
	case grid.Bool:
		if c.Bool {
			return "TRUE"
		}
		return "FALSE"
	}
	return c.Text
}

func (r *jsRange) GetTexts() goja.Value { return r.cells(r.text) }
func (r *jsRange) GetText() any {
	return r.text(r.s.state().Value(r.s.sheet(r.sheet), r.r.R1, r.r.C1), r.r.R1, r.r.C1)
}

func (r *jsRange) GetValueTypes() goja.Value {
	return r.cells(func(c *grid.Cell, _, _ int) any {
		switch {
		case c == nil || c.IsEmpty():
			return "Empty"
		case c.Pending:
			return "Unknown"
		case c.Kind == grid.Number:
			return "Double"
		case c.Kind == grid.Bool:
			return "Boolean"
		case c.Kind == grid.Error:
			return "Error"
		}
		return "String"
	})
}
func (r *jsRange) GetValueType() string {
	v := r.GetValueTypes().Export().([]any)
	return v[0].([]any)[0].(string)
}

func (r *jsRange) GetNumberFormats() goja.Value {
	wb := r.s.state()
	return r.cells(func(c *grid.Cell, _, _ int) any {
		if c == nil {
			return "General"
		}
		if f := wb.FormatOf(c); f.Code != "" {
			return f.Code
		}
		return "General"
	})
}
func (r *jsRange) GetNumberFormat() string {
	return r.GetNumberFormats().Export().([]any)[0].([]any)[0].(string)
}

// ---- Range writes ----

// writable refuses writing whole rows or columns, which would fill a
// million cells.
func (r *jsRange) writable(what string) grid.Range {
	if r.r.R2 == grid.MaxRows || r.r.C2 == grid.MaxCols {
		r.s.fail("%s: %s whole rows or columns; give a bounded range", r.GetAddress(), what)
	}
	return r.r
}

// grid2 reads a 2-D array of the range's shape.
func (r *jsRange) grid2(v goja.Value, what string) [][]goja.Value {
	rg := r.writable(what)
	obj, ok := v.(*goja.Object)
	if !ok || obj.ClassName() != "Array" {
		r.s.fail("%s(%s): give an array of rows, each an array of values", what, r.GetAddress())
	}
	rows := r.list(obj)
	if len(rows) != rg.Rows() {
		r.s.fail("%s: %s has %d rows, the values %d", what, r.GetAddress(), rg.Rows(), len(rows))
	}
	out := make([][]goja.Value, len(rows))
	for i, row := range rows {
		ro, ok := row.(*goja.Object)
		if !ok || ro.ClassName() != "Array" {
			r.s.fail("%s: row %d is not an array", what, i)
		}
		out[i] = r.list(ro)
		if len(out[i]) != rg.Cols() {
			r.s.fail("%s: %s has %d columns, row %d of the values %d", what, r.GetAddress(), rg.Cols(), i, len(out[i]))
		}
	}
	return out
}

func (r *jsRange) list(o *goja.Object) []goja.Value {
	n := int(o.Get("length").ToInteger())
	out := make([]goja.Value, n)
	for i := 0; i < n; i++ {
		out[i] = o.Get(fmt.Sprint(i))
	}
	return out
}

// input reads a value for a cell; ok is false for null and undefined,
// which leave the cell as it is, as in Office Scripts.
func (s *scriptSession) input(v goja.Value) (input.Value, bool) {
	if v == nil || goja.IsUndefined(v) || goja.IsNull(v) {
		return input.Value{}, false
	}
	switch x := v.Export().(type) {
	case bool:
		return input.Value{Kind: grid.Bool, Bool: x}, true
	case int64:
		return input.Value{Kind: grid.Number, Num: float64(x)}, true
	case float64:
		if math.IsNaN(x) || math.IsInf(x, 0) {
			s.fail("%v is not a number a cell can hold", x)
		}
		return input.Value{Kind: grid.Number, Num: x}, true
	case string:
		return input.Parse(x, false), true
	case time.Time:
		// In UTC, as (serial - 25569) * 86400000 turns serials into Dates.
		serial := grid.TimeToSerial(x.UTC())
		kind := grid.Date
		if serial != math.Trunc(serial) {
			kind = grid.DateTime
		}
		return input.Value{Kind: grid.Number, Num: serial, Fmt: grid.Format{Kind: kind}}, true
	}
	s.fail("a cell takes a number, text, true or false, or a Date; got %s", v.String())
	return input.Value{}, false
}

// put writes one cell.
func (r *jsRange) put(ed book.Editor, row, col int, v input.Value) {
	if v.Formula != "" {
		v.Formula = formula.FillRow(v.Formula, row)
		f, err := sheetTableRefs(r.s.state0(), v.Formula)
		r.s.must(err)
		v.Formula = f
	}
	r.s.must(ed.Set(r.sheet.name, row, col, v))
	r.s.cells++
	if v.Formula != "" {
		r.s.written = append(r.s.written, cellRef{r.sheet.name, row, col, spills(v.Formula)})
		r.s.recent[cellKey{strings.ToLower(r.sheet.name), row, col}] = true
	}
}

func (r *jsRange) SetValues(v goja.Value) {
	rows := r.grid2(v, "setValues")
	r.s.sheetName(r.sheet)
	ed := r.s.editor()
	for i, row := range rows {
		for j, x := range row {
			if in, ok := r.s.input(x); ok {
				r.put(ed, r.r.R1+i, r.r.C1+j, in)
			}
		}
	}
}

// SetValue sets every cell of the range to one value.
func (r *jsRange) SetValue(v goja.Value) {
	rg := r.writable("setValue")
	in, ok := r.s.input(v)
	if !ok {
		return
	}
	r.s.sheetName(r.sheet)
	ed := r.s.editor()
	for row := rg.R1; row <= rg.R2; row++ {
		for col := rg.C1; col <= rg.C2; col++ {
			c := in
			if in.Formula != "" && !formula.HasRowPlaceholder(in.Formula) {
				c.Formula = formula.Shift(in.Formula, row-rg.R1, col-rg.C1)
			}
			r.put(ed, row, col, c)
		}
	}
}

// SetFormula fills the range with a formula as Excel fills it down and
// right (relative references shift), as setValue does.
func (r *jsRange) SetFormula(f string) {
	if !strings.HasPrefix(f, "=") && f != "" {
		f = "=" + f
	}
	r.SetValue(r.s.vm.ToValue(f))
}

func (r *jsRange) SetFormulas(v goja.Value) { r.SetValues(v) }

func (r *jsRange) SetNumberFormat(code string) { r.style("format=" + code) }

func (r *jsRange) SetNumberFormats(v goja.Value) {
	rows := r.grid2(v, "setNumberFormats")
	for i, row := range rows {
		for j, x := range row {
			if x == nil || goja.IsUndefined(x) || goja.IsNull(x) {
				continue
			}
			r.GetCell(i, j).SetNumberFormat(x.String())
		}
	}
}

// style applies sheets style properties to the range.
func (r *jsRange) style(words ...string) {
	look, err := book.ParseLook(words)
	r.s.must(err)
	rg := r.r
	if rg.R2 == grid.MaxRows && rg.C2 != grid.MaxCols || rg.C2 == grid.MaxCols && rg.R2 != grid.MaxRows {
		// Whole columns or rows take widths and heights; their cells, the
		// used part.
		if !strings.HasPrefix(words[0], "width=") && !strings.HasPrefix(words[0], "height=") {
			rg = r.bounded()
		}
	}
	name := r.s.sheetName(r.sheet)
	_, err = r.s.editor().Style(name, rg, &look)
	r.s.must(err)
	r.s.styled[grid.FormatRef(name, rg)] = true
}

func (r *jsRange) Clear(applyTo string) {
	w := book.ClearWhat{}
	switch strings.ToLower(applyTo) {
	case "", "all":
		w = book.ClearWhat{Contents: true, Formats: true, Notes: true, Links: true}
	case "contents", "resetcontents":
		w.Contents = true
	case "formats":
		w.Formats = true
	case "hyperlinks", "removehyperlinks":
		w.Links = true
	default:
		r.s.fail("clear(%q): give ExcelScript.ClearApplyTo.all, contents, formats or hyperlinks", applyTo)
	}
	name, area := r.s.sheetName(r.sheet), r.bounded()
	_, err := r.s.rangeEditor().ClearRange(name, area, w)
	r.s.must(err)
	r.s.done = append(r.s.done, grid.FormatRef(name, area)+" cleared")
}

// lines tells whether the range is whole rows or whole columns.
func (r *jsRange) lines() (rows, cols bool) {
	return r.r.C1 == 1 && r.r.C2 == grid.MaxCols, r.r.R1 == 1 && r.r.R2 == grid.MaxRows
}

func (r *jsRange) Delete(shift string) {
	rows, cols := r.lines()
	name := r.s.sheetName(r.sheet)
	switch {
	case rows || !cols && strings.EqualFold(shift, "Up") && r.spansUsedWidth():
		r.s.must(r.s.editor().DeleteRows(name, r.r.R1, r.r.R2))
		r.s.done = append(r.s.done, lineSpan("row", r.r.R1, r.r.R2, strconv.Itoa)+" deleted")
	case cols || strings.EqualFold(shift, "Left") && r.spansUsedHeight():
		r.s.must(r.s.editor().DeleteCols(name, r.r.C1, r.r.C2))
		r.s.done = append(r.s.done, lineSpan("column", r.r.C1, r.r.C2, grid.ColName)+" deleted")
	default:
		r.s.fail("%s.delete: sheets deletes whole rows or columns (getEntireRow().delete()); shifting part of a row is not supported", r.GetAddress())
	}
}

func (r *jsRange) Insert(shift string) *jsRange {
	rows, cols := r.lines()
	name := r.s.sheetName(r.sheet)
	switch {
	case rows || !cols && strings.EqualFold(shift, "Down") && r.spansUsedWidth():
		r.s.must(r.s.editor().InsertRows(name, r.r.R1, r.r.Rows()))
		r.s.done = append(r.s.done, plural(r.r.Rows(), "row")+" inserted")
	case cols || strings.EqualFold(shift, "Right") && r.spansUsedHeight():
		r.s.must(r.s.editor().InsertCols(name, r.r.C1, r.r.Cols()))
		r.s.done = append(r.s.done, plural(r.r.Cols(), "column")+" inserted")
	default:
		r.s.fail("%s.insert: sheets inserts whole rows or columns (getEntireRow().insert(...)); shifting part of a row is not supported", r.GetAddress())
	}
	return r
}

// spansUsedWidth tells whether the range covers the sheet's used columns,
// so that shifting it up or down is shifting whole rows.
func (r *jsRange) spansUsedWidth() bool {
	u := r.s.sheet(r.sheet).Used()
	return r.r.C1 <= u.C1 && r.r.C2 >= u.C2
}

func (r *jsRange) spansUsedHeight() bool {
	u := r.s.sheet(r.sheet).Used()
	return r.r.R1 <= u.R1 && r.r.R2 >= u.R2
}

// CopyFrom pastes another range here, as Excel's copy and paste does.
func (r *jsRange) CopyFrom(source goja.Value, copyType string, skipBlanks, transpose bool) {
	src := r.s.rangeArg(source, r)
	if skipBlanks || transpose {
		r.s.fail("copyFrom: skipBlanks and transpose are not supported")
	}
	o := book.CopyOptions{}
	switch strings.ToLower(copyType) {
	case "", "all", "formulas":
	case "values":
		o.Values = true
	case "formats":
		o.Formats = true
	default:
		r.s.fail("copyFrom: copy type %q is not supported (all, values, formats)", copyType)
	}
	from, to, area := r.s.sheetName(src.sheet), r.s.sheetName(r.sheet), src.bounded()
	res, err := r.s.rangeEditor().CopyRange(from, area, to, r.r, o)
	r.s.must(err)
	r.s.done = append(r.s.done, fmt.Sprintf("%s copied to %s", grid.FormatRef(from, area), res.To))
	r.s.warnings = append(r.s.warnings, res.Notes...)
}

func (r *jsRange) MoveTo(dest goja.Value) {
	to := r.s.rangeArg(dest, r)
	from, at, area := r.s.sheetName(r.sheet), r.s.sheetName(to.sheet), r.bounded()
	res, err := r.s.rangeEditor().MoveRange(from, area, at, to.r)
	r.s.must(err)
	r.s.done = append(r.s.done, fmt.Sprintf("%s moved to %s", grid.FormatRef(from, area), res.To))
	r.s.warnings = append(r.s.warnings, res.Notes...)
}

// rangeArg reads a range given as a Range or an address on near's sheet.
func (s *scriptSession) rangeArg(v goja.Value, near *jsRange) *jsRange {
	if o, ok := v.Export().(*jsRange); ok {
		return o
	}
	if text, ok := v.Export().(string); ok {
		sheet, address := near.sheet, text
		if i := strings.LastIndex(text, "!"); i >= 0 {
			sheet, address = s.ref(strings.Trim(text[:i], "'")), text[i+1:]
		}
		return (&jsWorksheet{s, sheet}).GetRange(address)
	}
	s.fail("give a range or an address such as \"Sheet1!A1:B9\"")
	return nil
}

func (r *jsRange) GetFormat() *jsFormat { return &jsFormat{r} }
func (r *jsRange) GetSort() *jsSort     { return &jsSort{rg: r} }

func (r *jsRange) SetHyperlink(v goja.Value) {
	o, ok := v.(*goja.Object)
	if !ok {
		r.s.fail("setHyperlink: give {address, textToDisplay, screenTip}")
	}
	get := func(k string) string {
		x := o.Get(k)
		if x == nil || goja.IsUndefined(x) || goja.IsNull(x) {
			return ""
		}
		return x.String()
	}
	l := &grid.Link{URL: get("address"), Tooltip: get("screenTip")}
	if sub := get("documentReference"); sub != "" && l.URL == "" {
		l.Location = sub
	}
	name := r.s.sheetName(r.sheet)
	r.s.must(r.s.editor().SetLink(name, r.r.R1, r.r.C1, l, get("textToDisplay")))
}

// ---- Formats ----

type jsFormat struct{ rg *jsRange }

func (f *jsFormat) GetFont() *jsFont { return &jsFont{f.rg} }
func (f *jsFormat) GetFill() *jsFill { return &jsFill{f.rg} }
func (f *jsFormat) SetHorizontalAlignment(a string) {
	f.rg.style("align=" + map[string]string{"general": "general", "left": "left", "center": "center", "right": "right", "fill": "fill",
		"justify": "justify", "centeracrossselection": "center-across", "distributed": "distributed"}[strings.ToLower(a)])
}
func (f *jsFormat) SetVerticalAlignment(a string) {
	f.rg.style("valign=" + map[string]string{"top": "top", "center": "middle", "bottom": "bottom", "justify": "justify", "distributed": "distributed"}[strings.ToLower(a)])
}
func (f *jsFormat) SetWrapText(b bool)       { f.rg.style(fmt.Sprintf("wrap=%v", b)) }
func (f *jsFormat) SetIndentLevel(n int)     { f.rg.style(fmt.Sprintf("indent=%d", n)) }
func (f *jsFormat) SetTextOrientation(n int) { f.rg.style(fmt.Sprintf("rotate=%d", n)) }

// SetColumnWidth takes points, as Office Scripts does; Excel's widths are
// characters of the default font, seven pixels each plus five.
func (f *jsFormat) SetColumnWidth(points float64) {
	chars := math.Max((points*4/3-5)/7, 0.5)
	f.rg.style(fmt.Sprintf("width=%.2f", chars))
}
func (f *jsFormat) SetRowHeight(points float64) { f.rg.style(fmt.Sprintf("height=%g", points)) }
func (f *jsFormat) AutofitColumns()             { f.rg.style("width=auto") }
func (f *jsFormat) AutofitRows()                {}

type jsFont struct{ rg *jsRange }

func (f *jsFont) SetBold(b bool)          { f.rg.style(fmt.Sprintf("bold=%v", b)) }
func (f *jsFont) SetItalic(b bool)        { f.rg.style(fmt.Sprintf("italic=%v", b)) }
func (f *jsFont) SetStrikethrough(b bool) { f.rg.style(fmt.Sprintf("strike=%v", b)) }
func (f *jsFont) SetUnderline(u string) {
	f.rg.style(fmt.Sprintf("underline=%v", !strings.EqualFold(u, "None")))
}
func (f *jsFont) SetColor(c string) { f.rg.style("color=" + c) }
func (f *jsFont) SetSize(n float64) { f.rg.style(fmt.Sprintf("size=%g", n)) }
func (f *jsFont) SetName(n string)  { f.rg.style("font=" + n) }

type jsFill struct{ rg *jsRange }

func (f *jsFill) SetColor(c string) { f.rg.style("fill=" + c) }
func (f *jsFill) Clear()            { f.rg.style("fill=none") }

// ---- Sorting ----

type jsSort struct {
	rg    *jsRange
	table *jsTable
}

// Apply sorts the range (or the table's data rows) by fields, each
// {key: column index within the range, ascending?: bool}.
func (s *jsSort) Apply(fields goja.Value, matchCase, hasHeaders bool, orientation string) {
	ss := s.rg.s
	if strings.EqualFold(orientation, "Columns") {
		ss.fail("sort: sheets sorts rows, not columns")
	}
	rg := s.rg
	if s.table != nil {
		rg = s.table.GetRangeBetweenHeaderAndTotal()
	}
	area := rg.bounded()
	if hasHeaders {
		area.R1++
	}
	obj, ok := fields.(*goja.Object)
	if !ok || obj.ClassName() != "Array" {
		ss.fail("sort.apply: give an array of {key, ascending}")
	}
	var keys []book.SortKey
	for _, f := range rg.list(obj) {
		fo, ok := f.(*goja.Object)
		if !ok {
			ss.fail("sort.apply: each field is {key, ascending}")
		}
		key := fo.Get("key")
		if key == nil || goja.IsUndefined(key) {
			ss.fail("sort.apply: a field needs its key, the column's index in the range")
		}
		col := area.C1 + int(key.ToInteger())
		if col < area.C1 || col > area.C2 {
			ss.fail("sort.apply: key %d is outside the range's %d columns", key.ToInteger(), area.Cols())
		}
		asc := fo.Get("ascending")
		keys = append(keys, book.SortKey{Col: col, Desc: asc != nil && !goja.IsUndefined(asc) && !asc.ToBoolean()})
	}
	name := ss.sheetName(rg.sheet)
	res, err := ss.rangeEditor().SortRange(name, area, keys)
	ss.must(err)
	ss.done = append(ss.done, res.Range+" sorted")
	ss.warnings = append(ss.warnings, res.Notes...)
}

// ---- Tables ----

// jsTable is an Excel table, or a table of a sheet, named after its sheet
// as sheets query names it (Sales, then Sales_2 for a second).
type jsTable struct {
	s     *scriptSession
	name  string
	excel bool
	on    *sheetRef // a sheet's table: its sheet, and which one, from the top
	index int
}

func (t *jsTable) ToJSON() string  { return t.name }
func (t *jsTable) GetName() string { return t.name }

// part is a part of the table where it is now: #All, #Headers, #Data or
// #Totals.
func (t *jsTable) part(spec string) *jsRange {
	if t.excel {
		sheet, rg, err := analyze.Resolve(t.s.state(), t.name+"["+spec+"]")
		if err != nil {
			t.s.fail("%v", err)
		}
		return &jsRange{t.s, t.s.ref(sheet), *rg}
	}
	reg, ok := t.region()
	if !ok {
		t.s.fail("the table %s no longer exists", t.name)
	}
	var rg grid.Range
	switch spec {
	case "#All":
		rg = grid.Range{R1: reg.HeaderRow, C1: reg.Range.C1, R2: max(reg.Body.R2, reg.TotalRow), C2: reg.Range.C2}
	case "#Headers":
		rg = grid.Range{R1: reg.HeaderRow, C1: reg.Range.C1, R2: reg.HeaderRow, C2: reg.Range.C2}
	case "#Data":
		rg = reg.Body
	case "#Totals":
		if reg.TotalRow == 0 {
			t.s.fail("the table %s has no total row", t.name)
		}
		rg = grid.Range{R1: reg.TotalRow, C1: reg.Range.C1, R2: reg.TotalRow, C2: reg.Range.C2}
	}
	return &jsRange{t.s, t.on, rg}
}

// region finds a sheet's table where it is now.
func (t *jsTable) region() (analyze.Region, bool) {
	sh := t.s.sheet(t.on)
	n := 0
	for _, reg := range analyze.Regions(t.s.state(), sh) {
		if reg.IsTable() && sh.TableAt(reg.Range) == nil {
			if n == t.index {
				return reg, true
			}
			n++
		}
	}
	return analyze.Region{}, false
}

func (t *jsTable) sheet() *grid.Sheet {
	r := t.part("#All")
	return t.s.sheet(r.sheet)
}

func (t *jsTable) GetWorksheet() *jsWorksheet { return &jsWorksheet{t.s, t.part("#All").sheet} }
func (t *jsTable) GetRange() *jsRange         { return t.part("#All") }
func (t *jsTable) GetHeaderRowRange() *jsRange {
	return t.part("#Headers")
}
func (t *jsTable) GetRangeBetweenHeaderAndTotal() *jsRange { return t.part("#Data") }
func (t *jsTable) GetTotalRowRange() goja.Value {
	if !t.GetShowTotals() {
		return goja.Undefined()
	}
	return t.s.vm.ToValue(t.part("#Totals"))
}
func (t *jsTable) GetShowTotals() bool {
	if !t.excel {
		reg, ok := t.region()
		return ok && reg.TotalRow > 0
	}
	_, _, err := analyze.Resolve(t.s.state(), t.name+"[#Totals]")
	return err == nil
}
func (t *jsTable) GetRowCount() int { return t.part("#Data").r.Rows() }

func (t *jsTable) columns() []*jsTableColumn {
	h := t.GetHeaderRowRange()
	sh := t.s.sheet(h.sheet)
	var out []*jsTableColumn
	for c := h.r.C1; c <= h.r.C2; c++ {
		out = append(out, &jsTableColumn{t, c - h.r.C1, strings.TrimSpace(analyze.Label(t.s.state(), sh, h.r.R1, c))})
	}
	return out
}

func (t *jsTable) GetColumns() goja.Value {
	var out []any
	for _, c := range t.columns() {
		out = append(out, c)
	}
	return t.s.array(out)
}

func (t *jsTable) GetColumnByName(name string) goja.Value {
	for _, c := range t.columns() {
		if strings.EqualFold(c.name, name) {
			return t.s.vm.ToValue(c)
		}
	}
	return goja.Undefined()
}

func (t *jsTable) GetColumn(key goja.Value) goja.Value {
	if n, ok := key.Export().(int64); ok {
		cols := t.columns()
		if n >= 0 && int(n) < len(cols) {
			return t.s.vm.ToValue(cols[n])
		}
		return goja.Undefined()
	}
	return t.GetColumnByName(key.String())
}

func (t *jsTable) GetSort() *jsSort { return &jsSort{rg: t.GetRange(), table: t} }

// AddRow appends a row; index, when given, must be the table's end.
func (t *jsTable) AddRow(index goja.Value, values goja.Value) {
	arr := t.s.vm.NewArray(values)
	t.AddRows(index, arr)
}

// AddRows appends rows at the table's end, as sheets write --append does:
// above its total row, filling its formula columns.
func (t *jsTable) AddRows(index goja.Value, values goja.Value) {
	if index != nil && !goja.IsUndefined(index) && !goja.IsNull(index) {
		if n := index.ToInteger(); n != -1 && int(n) != t.GetRowCount() {
			t.s.fail("addRows at %d: sheets adds rows at a table's end (index -1 or omitted)", n)
		}
	}
	obj, ok := values.(*goja.Object)
	if !ok || obj.ClassName() != "Array" {
		t.s.fail("addRows: give an array of rows, each an array of values")
	}
	sheet := t.sheet()
	cols := t.GetHeaderRowRange().r.Cols()
	var rows [][]input.Value
	for i, row := range t.part("#All").list(obj) {
		ro, ok := row.(*goja.Object)
		if !ok || ro.ClassName() != "Array" {
			t.s.fail("addRows: row %d is not an array", i)
		}
		cells := t.part("#All").list(ro)
		if len(cells) > cols {
			t.s.fail("addRows: row %d has %d values; the table has %d columns", i, len(cells), cols)
		}
		var line []input.Value
		for _, c := range cells {
			v, _ := t.s.input(c)
			line = append(line, v)
		}
		rows = append(rows, line)
	}
	if len(rows) == 0 {
		return
	}
	last := lastTable(t.s.state(), sheet)
	if all := t.part("#All").r; last == nil || last.Range.Intersect(all).Empty() {
		t.s.fail("addRows: sheets appends to the last table of a sheet; %s is not the last on %s", t.name, sheet.Name)
	}
	start := t.part("#All").r.C1
	ref := grid.QuoteSheet(sheet.Name) + "!" + grid.CellName(1, start)
	res, err := writeRows(t.s.editor(), t.s.file, ref, rows, false, true, false)
	t.s.must(err)
	t.s.cells += res.Rows * res.Cols
	t.s.written = append(t.s.written, res.written...)
	for _, c := range res.written {
		t.s.recent[cellKey{strings.ToLower(c.sheet), c.row, c.col}] = true
	}
	note := plural(res.Rows, "row") + " added to " + t.name
	if len(res.Filled) > 0 {
		note += ", filling " + strings.Join(res.Filled, ", ")
	}
	t.s.done = append(t.s.done, note)
}

// DeleteRowsAt deletes count rows of the table's data from index.
func (t *jsTable) DeleteRowsAt(index, count int) {
	if count < 1 {
		count = 1
	}
	body := t.part("#Data").r
	if index < 0 || index+count > body.Rows() {
		t.s.fail("deleteRowsAt(%d, %d): the table has %d rows", index, count, body.Rows())
	}
	name := t.sheet().Name
	t.s.must(t.s.editor().DeleteRows(name, body.R1+index, body.R1+index+count-1))
	t.s.done = append(t.s.done, fmt.Sprintf("%s deleted from %s", plural(count, "row"), t.name))
}

type jsTableColumn struct {
	t     *jsTable
	index int
	name  string
}

func (c *jsTableColumn) ToJSON() string  { return c.name }
func (c *jsTableColumn) GetName() string { return c.name }
func (c *jsTableColumn) GetIndex() int   { return c.index }
func (c *jsTableColumn) column(r *jsRange) *jsRange {
	return r.sub(grid.Range{R1: r.r.R1, C1: r.r.C1 + c.index, R2: r.r.R2, C2: r.r.C1 + c.index})
}
func (c *jsTableColumn) GetRange() *jsRange          { return c.column(c.t.GetRange()) }
func (c *jsTableColumn) GetHeaderRowRange() *jsRange { return c.column(c.t.GetHeaderRowRange()) }
func (c *jsTableColumn) GetRangeBetweenHeaderAndTotal() *jsRange {
	return c.column(c.t.GetRangeBetweenHeaderAndTotal())
}
func (c *jsTableColumn) GetTotalRowRange() goja.Value {
	if !c.t.GetShowTotals() {
		return goja.Undefined()
	}
	return c.t.s.vm.ToValue(c.column(c.t.part("#Totals")))
}

// table finds a table by name: an Excel table, or a sheet holding one.
func (s *scriptSession) table(name string) (*jsTable, bool) {
	for _, t := range s.tables() {
		if strings.EqualFold(t.name, name) {
			return t, true
		}
	}
	return nil, false
}

// tables lists the Excel tables, then the tables of each sheet that are
// not Excel tables.
func (s *scriptSession) tables() []*jsTable {
	wb := s.state()
	var out []*jsTable
	taken := map[string]bool{}
	for _, sh := range wb.Sheets {
		for _, t := range sh.Tables {
			out = append(out, &jsTable{s: s, name: t.Name, excel: true})
			taken[strings.ToLower(t.Name)] = true
		}
	}
	for _, sh := range wb.Sheets {
		n := 0
		for _, reg := range analyze.Regions(wb, sh) {
			if !reg.IsTable() || sh.TableAt(reg.Range) != nil {
				continue
			}
			name := sh.Name
			if n > 0 {
				name = fmt.Sprintf("%s_%d", sh.Name, n+1)
			}
			if !taken[strings.ToLower(name)] {
				taken[strings.ToLower(name)] = true
				out = append(out, &jsTable{s: s, name: name, on: s.ref(sh.Name), index: n})
			}
			n++
		}
	}
	return out
}

// lineSpan names rows or columns: "row 4", "rows 4:9", "columns B:C".
func lineSpan(word string, a, b int, name func(int) string) string {
	if a == b {
		return word + " " + name(a)
	}
	return word + "s " + name(a) + ":" + name(b)
}

// sheetName is a sheet's name now, without saving pending writes to read
// it.
func (s *scriptSession) sheetName(ref *sheetRef) string {
	for _, n := range s.sheetNames() {
		if strings.EqualFold(n, ref.name) {
			return n
		}
	}
	s.fail("the worksheet %q no longer exists", ref.name)
	return ""
}

// sheetNames lists the sheets as they are now.
func (s *scriptSession) sheetNames() []string {
	if s.ed != nil {
		return s.ed.SheetList()
	}
	return s.state().SheetNames()
}

func (s *scriptSession) hasSheet(name string) bool {
	for _, n := range s.sheetNames() {
		if strings.EqualFold(n, name) {
			return true
		}
	}
	return false
}

// array makes a JS array.
func (s *scriptSession) array(items []any) goja.Value {
	if items == nil {
		items = []any{}
	}
	return s.vm.NewArray(items...)
}

// state0 is the workbook as of the last read, for translating table
// references: the tables a formula names exist already.
func (s *scriptSession) state0() *grid.Workbook {
	if s.wb != nil {
		return s.wb
	}
	if s.ed != nil {
		return s.ed.Workbook()
	}
	return s.state()
}

// sortCells orders cell references by sheet, then row, then column.
func sortCells(cells []string) {
	key := func(ref string) (string, int, int) {
		sheet, cell := "", ref
		if i := strings.LastIndex(ref, "!"); i >= 0 {
			sheet, cell = ref[:i], ref[i+1:]
		}
		r, c, _ := grid.ParseCell(cell)
		return sheet, r, c
	}
	sort.Slice(cells, func(i, j int) bool {
		si, ri, ci := key(cells[i])
		sj, rj, cj := key(cells[j])
		if si != sj {
			return si < sj
		}
		if ri != rj {
			return ri < rj
		}
		return ci < cj
	})
}

// ---- Borders ----

// jsBorder is one edge of a range's borders; each change applies at once.
type jsBorder struct {
	rg                   *jsRange
	side                 string
	style, weight, color string
}

func (f *jsFormat) GetRangeBorder(index string) *jsBorder {
	side, ok := map[string]string{"edgetop": "top", "edgebottom": "bottom", "edgeleft": "left", "edgeright": "right",
		"insidehorizontal": "inner", "insidevertical": "inner"}[strings.ToLower(index)]
	if !ok {
		f.rg.s.fail("getRangeBorder(%q): sheets draws the edges (edgeTop, edgeBottom, edgeLeft, edgeRight) and inside lines, not diagonals", index)
	}
	return &jsBorder{rg: f.rg, side: side, style: "Continuous", weight: "Thin", color: "#000000"}
}

func (b *jsBorder) apply() {
	line := map[string]string{"hairline": "hair", "thin": "thin", "medium": "medium", "thick": "thick"}[strings.ToLower(b.weight)]
	switch strings.ToLower(b.style) {
	case "none":
		line = "none"
	case "double":
		line = "double"
	case "dash", "dashdot", "dashdotdot", "slantdashdot":
		line = "dashed"
	case "dot":
		line = "dotted"
	}
	if line == "" {
		line = "thin"
	}
	b.rg.style(fmt.Sprintf("border=%s:%s:%s", b.side, line, b.color))
}

func (b *jsBorder) SetStyle(s string)  { b.style = s; b.apply() }
func (b *jsBorder) SetWeight(w string) { b.weight = w; b.apply() }
func (b *jsBorder) SetColor(c string)  { b.color = c; b.apply() }
func (b *jsBorder) GetStyle() string   { return b.style }
func (b *jsBorder) GetWeight() string  { return b.weight }
func (b *jsBorder) GetColor() string   { return b.color }

// ---- Duplicates and finding ----

type jsRemoved struct{ removed, remaining int }

func (r *jsRemoved) GetRemoved() int         { return r.removed }
func (r *jsRemoved) GetUniqueRemaining() int { return r.remaining }

// RemoveDuplicates deletes the rows of the range repeating an earlier
// row's values in the given columns (indexes within the range), text
// compared without case, as Excel does. The range must span the sheet's
// used columns: the rows go whole.
func (r *jsRange) RemoveDuplicates(columns goja.Value, includesHeader bool) *jsRemoved {
	area := r.bounded()
	if !r.spansUsedWidth() {
		r.s.fail("removeDuplicates: sheets removes whole rows; give a range spanning the used columns (%s)", grid.Range{R1: area.R1, C1: r.s.sheet(r.sheet).Used().C1, R2: area.R2, C2: r.s.sheet(r.sheet).Used().C2})
	}
	var keys []int
	if o, ok := columns.(*goja.Object); ok && o.ClassName() == "Array" {
		for _, v := range r.list(o) {
			k := int(v.ToInteger())
			if k < 0 || k >= area.Cols() {
				r.s.fail("removeDuplicates: column %d is outside the range's %d columns", k, area.Cols())
			}
			keys = append(keys, area.C1+k)
		}
	}
	if len(keys) == 0 {
		for c := area.C1; c <= area.C2; c++ {
			keys = append(keys, c)
		}
	}
	first := area.R1
	if includesHeader {
		first++
	}
	wb, sh := r.s.state(), r.s.sheet(r.sheet)
	seen := map[string]bool{}
	var dups []int
	for row := first; row <= area.R2; row++ {
		var key strings.Builder
		for _, c := range keys {
			v := wb.Value(sh, row, c)
			if v != nil {
				key.WriteString(strings.ToLower(grid.Display(v, grid.Format{}, false)))
			}
			key.WriteByte(0)
		}
		if seen[key.String()] {
			dups = append(dups, row)
		}
		seen[key.String()] = true
	}
	name := r.s.sheetName(r.sheet)
	// From the bottom, a run of rows at a time.
	for i := len(dups) - 1; i >= 0; {
		j := i
		for j > 0 && dups[j-1] == dups[j]-1 {
			j--
		}
		r.s.must(r.s.editor().DeleteRows(name, dups[j], dups[i]))
		i = j - 1
	}
	if len(dups) > 0 {
		r.s.done = append(r.s.done, fmt.Sprintf("%s removed as duplicates", plural(len(dups), "row")))
	}
	return &jsRemoved{removed: len(dups), remaining: area.R2 - first + 1 - len(dups)}
}

// Find gives the first cell whose value shows text: a part of it, or all
// with completeMatch; without case unless matchCase.
func (r *jsRange) Find(text string, criteria goja.Value) goja.Value {
	var complete, matchCase, backwards bool
	if o, ok := criteria.(*goja.Object); ok {
		get := func(k string) goja.Value { return o.Get(k) }
		complete = get("completeMatch") != nil && get("completeMatch").ToBoolean()
		matchCase = get("matchCase") != nil && get("matchCase").ToBoolean()
		backwards = get("searchDirection") != nil && strings.EqualFold(get("searchDirection").String(), "Backwards")
	}
	want := text
	if !matchCase {
		want = strings.ToLower(text)
	}
	area := r.bounded()
	wb, sh := r.s.state(), r.s.sheet(r.sheet)
	match := func(row, col int) bool {
		c := wb.Value(sh, row, col)
		if c == nil || c.IsEmpty() {
			return false
		}
		got := grid.Display(c, wb.FormatOf(c), wb.Date1904)
		if !matchCase {
			got = strings.ToLower(got)
		}
		if complete {
			return got == want
		}
		return strings.Contains(got, want)
	}
	for i := 0; i < area.Rows()*area.Cols(); i++ {
		k := i
		if backwards {
			k = area.Rows()*area.Cols() - 1 - i
		}
		row, col := area.R1+k/area.Cols(), area.C1+k%area.Cols()
		if match(row, col) {
			return r.s.vm.ToValue(r.sub(grid.Range{R1: row, C1: col, R2: row, C2: col}))
		}
	}
	return goja.Undefined()
}

// ---- Filters ----

type jsAutoFilter struct{ w *jsWorksheet }

func (w *jsWorksheet) GetAutoFilter() *jsAutoFilter { return &jsAutoFilter{w} }

// Apply puts filter buttons on a range's header row; filtering rows by
// criteria is Excel's to do.
func (f *jsAutoFilter) Apply(rng goja.Value, column goja.Value, criteria goja.Value) {
	if criteria != nil && !goja.IsUndefined(criteria) && !goja.IsNull(criteria) {
		f.w.s.fail("autoFilter.apply: sheets puts filter buttons on a table; hiding rows by criteria is not supported (filter the values in the script instead)")
	}
	rg := f.w.s.rangeArg(rng, f.w.GetRange("A1"))
	area := rg.bounded()
	name := f.w.s.sheetName(f.w.ref)
	f.w.s.must(f.w.s.editor().View(name, book.ViewOptions{Filter: &area}))
}

func (f *jsAutoFilter) Remove() {
	name := f.w.s.sheetName(f.w.ref)
	f.w.s.must(f.w.s.editor().View(name, book.ViewOptions{NoFilter: true}))
}

// ---- What sheets leaves to its own commands ----

func (w *jsWorksheet) AddTable(address goja.Value, hasHeaders bool) {
	w.s.fail("addTable: sheets does not create Excel tables (Insert › Table). A block with a header row is a table for sheets already: workbook.getTable(%q) finds it", w.ref.name)
}

func (w *jsWorksheet) AddChart(kind string, source goja.Value) {
	w.s.fail("addChart: charts are made with sheets chart add (in the house style), after the script")
}

func (w *jsWorksheet) GetCharts() goja.Value {
	w.s.fail("getCharts: list charts with sheets chart list")
	return nil
}

func (r *jsRange) AddConditionalFormat(kind string) {
	r.s.fail("addConditionalFormat: conditional formats are set with sheets cond, after the script")
}

func (r *jsRange) GetDataValidation() {
	r.s.fail("getDataValidation: data validation is set with sheets validate, after the script")
}

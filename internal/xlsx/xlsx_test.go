package xlsx

import (
	"archive/zip"
	"bytes"
	"io"
	"os"
	"path/filepath"
	"slices"
	"strconv"
	"strings"
	"testing"
	"time"

	"github.com/eemax/sheets/internal/cache"
	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/input"
	"github.com/eemax/sheets/internal/style"
	"github.com/xuri/excelize/v2"
)

// writeBook writes rows to sheet Data of a new workbook, typing values the
// way the write command does ({r} in formulas is the row).
func writeBook(t *testing.T, rows [][]string) string {
	t.Helper()
	path := filepath.Join(t.TempDir(), "book.xlsx")
	e, err := OpenEditor(path, true)
	if err != nil {
		t.Fatal(err)
	}
	name, _, err := e.EnsureSheet("Data")
	if err != nil {
		t.Fatal(err)
	}
	for r, row := range rows {
		for c, v := range row {
			if err := e.Set(name, r+1, c+1, input.Parse(strings.ReplaceAll(v, "{r}", strconv.Itoa(r+1)), false)); err != nil {
				t.Fatal(err)
			}
		}
	}
	if err := e.Save(false); err != nil {
		t.Fatal(err)
	}
	return path
}

func TestRoundTrip(t *testing.T) {
	path := writeBook(t, [][]string{
		{"Date", "Region", "Revenue", "Share", "Total"},
		{"2026-07-01", "North", "$1,200.50", "12.5%", "=C2*2"},
		{"2026-07-02", "South", "$900", "7%", "=C3*2"},
	})
	wb, err := Read(path)
	if err != nil {
		t.Fatal(err)
	}
	if len(wb.Sheets) != 1 || wb.Sheets[0].Name != "Data" {
		t.Fatalf("sheets = %v", wb.SheetNames())
	}
	s := wb.Sheets[0]
	if got := s.Used().String(); got != "A1:E3" {
		t.Errorf("used = %s", got)
	}
	check := func(row, col int, want string) {
		t.Helper()
		c := wb.Value(s, row, col)
		if got := grid.Display(c, wb.FormatOf(c), wb.Date1904); got != want {
			t.Errorf("%s = %q, want %q", grid.CellName(row, col), got, want)
		}
	}
	check(1, 1, "Date")
	check(2, 1, "2026-07-01")
	check(2, 3, "1200.5")
	check(2, 4, "12.5%")
	check(2, 5, "2401")
	check(3, 5, "1800")
	if f := wb.FormatOf(s.At(2, 3)); f.Kind != grid.Currency || f.Symbol != "$" {
		t.Errorf("C2 format = %+v", f)
	}
	// The result is stored in the file, and trusted on reading it back.
	if c := s.At(2, 5); c.Formula != "C2*2" || c.Pending || c.Computed || c.Num != 2401 {
		t.Errorf("E2 = %+v", c)
	}
}

func TestFormatIdempotent(t *testing.T) {
	path := writeBook(t, [][]string{
		{"Q3 report"},
		{"Date", "Region", "Revenue", "Share", "Total", "Year"},
		{"2026-07-01", "North", "$1,200.50", "12.5%", "=C3*2", "2026"},
		{"2026-07-02", "South", "$900", "7%", "=C4*2", "2025"},
		{"2026-07-03", "East", "$300", "70%", "=C5*2", "2024"},
		{"Total", "", "=SUM(C3:C5)", "", "=SUM(E3:E5)", ""},
	})
	opts := FormatOptions{Preset: style.Default}
	ok, _, err := Check(path, opts)
	if err != nil {
		t.Fatal(err)
	}
	if ok {
		t.Fatal("unformatted file reported as formatted")
	}
	res, err := Format(path, opts)
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Sheets) != 1 || len(res.Sheets[0].Tables) != 1 {
		t.Fatalf("result = %+v", res)
	}
	tb := res.Sheets[0].Tables[0]
	if tb.Range != "A1:F6" || !tb.Total {
		t.Errorf("table = %+v", tb)
	}
	var types []string
	for _, c := range tb.Columns {
		types = append(types, c.Type)
	}
	// Total (=Revenue*2) takes the currency of Revenue.
	if got := strings.Join(types, ","); got != "date,text,currency $,percent,currency $,integer" {
		t.Errorf("types = %s", got)
	}
	ok, _, err = Check(path, opts)
	if err != nil {
		t.Fatal(err)
	}
	if !ok {
		t.Fatal("formatted file reported as unformatted")
	}
	// Formatting again must not change the file at all.
	before, _ := os.ReadFile(path)
	if _, err := Format(path, opts); err != nil {
		t.Fatal(err)
	}
	after, _ := os.ReadFile(path)
	if same, _ := samePackage(before, after); !same {
		t.Error("second format changed the file")
	}
}

// Overwriting the first cell of a filled-down (shared) formula must not
// erase the formulas below it, which excelize does on its own.
func TestOverwriteSharedFormulaAnchor(t *testing.T) {
	path := filepath.Join(t.TempDir(), "shared.xlsx")
	f := excelize.NewFile()
	for r := 1; r <= 4; r++ {
		f.SetSheetRow("Sheet1", grid.CellName(r, 1), &[]any{r, r * 10})
	}
	typ, ref := excelize.STCellFormulaTypeShared, "C1:C4"
	if err := f.SetCellFormula("Sheet1", "C1", "A1+B1", excelize.FormulaOpts{Type: &typ, Ref: &ref}); err != nil {
		t.Fatal(err)
	}
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	wb, err := Read(path)
	if err != nil {
		t.Fatal(err)
	}
	if c := wb.Sheets[0].At(3, 3); c.Formula != "A3+B3" || !c.Shared {
		t.Fatalf("C3 = %+v", c)
	}
	e, err := OpenEditor(path, false)
	if err != nil {
		t.Fatal(err)
	}
	if err := e.Set("Sheet1", 1, 3, input.Parse("=A1*B1", false)); err != nil {
		t.Fatal(err)
	}
	if err := e.Save(false); err != nil {
		t.Fatal(err)
	}
	wb, err = Read(path)
	if err != nil {
		t.Fatal(err)
	}
	s := wb.Sheets[0]
	for row, want := range map[int]string{1: "A1*B1", 2: "A2+B2", 4: "A4+B4"} {
		if got := s.Get(row, 3).Formula; got != want {
			t.Errorf("C%d = %q, want %q", row, got, want)
		}
	}
	if got := s.Used().String(); got != "A1:C4" {
		t.Errorf("used = %s", got)
	}
}

// A value far from the data must not make fmt sweep the rectangle between
// them (it used to touch every cell up to the stray one).
func TestFormatSparseSheet(t *testing.T) {
	path := writeBook(t, [][]string{{"Name", "Qty"}, {"a", "1"}, {"b", "2"}})
	e, err := OpenEditor(path, false)
	if err != nil {
		t.Fatal(err)
	}
	if err := e.Set("Data", 1000000, 52, input.Parse("stray", false)); err != nil {
		t.Fatal(err)
	}
	if err := e.Save(false); err != nil {
		t.Fatal(err)
	}
	done := make(chan error, 1)
	go func() { _, err := Format(path, FormatOptions{Preset: style.Default}); done <- err }()
	select {
	case err := <-done:
		if err != nil {
			t.Fatal(err)
		}
	case <-time.After(10 * time.Second):
		t.Fatal("fmt on a sparse sheet is too slow")
	}
}

// Formatting from other tools is removed wherever it lives: on cells
// (including empty ones), rows and whole columns.
func TestFormatStripsForeignStyles(t *testing.T) {
	path := filepath.Join(t.TempDir(), "styled.xlsx")
	f := excelize.NewFile()
	f.SetSheetRow("Sheet1", "A1", &[]any{"Item", "Cost"})
	f.SetSheetRow("Sheet1", "A2", &[]any{"a", 1})
	f.SetSheetRow("Sheet1", "A3", &[]any{"b", 2})
	loud, _ := f.NewStyle(&excelize.Style{Fill: excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{"FFFF00"}}, Font: &excelize.Font{Family: "Comic Sans MS"}})
	f.SetColStyle("Sheet1", "A:Z", loud)
	f.SetRowStyle("Sheet1", 8, 8, loud)
	f.SetCellStyle("Sheet1", "F6", "F6", loud)
	f.SetCellStyle("Sheet1", "A2", "B3", loud)
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	if _, err := Format(path, FormatOptions{Preset: style.Default}); err != nil {
		t.Fatal(err)
	}
	g, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	defer g.Close()
	for _, cell := range []string{"A2", "B3", "F6", "C8"} {
		id, _ := g.GetCellStyle("Sheet1", cell)
		st, _ := g.GetStyle(id)
		if len(st.Fill.Color) > 0 && st.Fill.Color[0] == "FFFF00" || st.Font != nil && st.Font.Family == "Comic Sans MS" {
			t.Errorf("%s keeps foreign style %+v", cell, st)
		}
	}
	if id, _ := g.GetColStyle("Sheet1", "M"); id != 0 {
		t.Errorf("column M keeps style %d", id)
	}
	if ok, _, err := Check(path, FormatOptions{Preset: style.Default}); err != nil || !ok {
		t.Errorf("not idempotent: %v %v", ok, err)
	}
}

// A real Excel table keeps its own style; a sheet filter over it would make
// the file invalid.
func TestFormatLeavesExcelTables(t *testing.T) {
	path := filepath.Join(t.TempDir(), "table.xlsx")
	f := excelize.NewFile()
	f.SetSheetRow("Sheet1", "A1", &[]any{"k", "v"})
	f.SetSheetRow("Sheet1", "A2", &[]any{"a", 1})
	f.SetSheetRow("Sheet1", "A3", &[]any{"b", 2})
	if err := f.AddTable("Sheet1", &excelize.Table{Range: "A1:B3", Name: "Things"}); err != nil {
		t.Fatal(err)
	}
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	res, err := Format(path, FormatOptions{Preset: style.Default})
	if err != nil {
		t.Fatal(err)
	}
	if notes := res.Sheets[0].Notes; len(notes) != 1 || !strings.Contains(notes[0], `"Things"`) {
		t.Errorf("notes = %v", notes)
	}
	xml := sheetXML(t, path, "xl/worksheets/sheet1.xml")
	if strings.Contains(xml, "<autoFilter") || strings.Contains(xml, "<conditionalFormatting") {
		t.Errorf("table got a sheet filter or banding: %s", xml)
	}
}

// When a sheet stops having exactly one table, the filter fmt set before
// goes away.
func TestFormatRemovesStaleFilter(t *testing.T) {
	path := writeBook(t, [][]string{{"Item", "Qty"}, {"a", "1"}, {"b", "2"}})
	opts := FormatOptions{Preset: style.Default}
	if _, err := Format(path, opts); err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(sheetXML(t, path, "xl/worksheets/sheet1.xml"), "<autoFilter") {
		t.Fatal("expected a filter on the only table")
	}
	e, _ := OpenEditor(path, false)
	for r, row := range [][]string{{"Other", "Cost"}, {"x", "5"}, {"y", "6"}} {
		for c, v := range row {
			e.Set("Data", r+1, c+5, input.Parse(v, false))
		}
	}
	if err := e.Save(false); err != nil {
		t.Fatal(err)
	}
	if _, err := Format(path, opts); err != nil {
		t.Fatal(err)
	}
	if xml := sheetXML(t, path, "xl/worksheets/sheet1.xml"); strings.Contains(xml, "<autoFilter") {
		t.Error("stale filter kept")
	}
	if wbx := sheetXML(t, path, "xl/workbook.xml"); strings.Contains(wbx, "_FilterDatabase") || strings.Contains(wbx, "<definedNames></definedNames>") {
		t.Errorf("filter name left behind: %s", wbx)
	}
	if ok, _, _ := Check(path, opts); !ok {
		t.Error("not idempotent")
	}
}

func TestReadRejectsEncrypted(t *testing.T) {
	_, err := ReadBytes(append([]byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1}, make([]byte, 512)...))
	if err == nil || !strings.Contains(err.Error(), "password-protected or in the legacy .xls format") {
		t.Errorf("err = %v", err)
	}
}

func TestUnescapeOOXML(t *testing.T) {
	for in, want := range map[string]string{
		"line_x000D_\nnext": "line\r\nnext",
		"a_x005F_x000D_b":   "a_x000D_b",
		"plain_text":        "plain_text",
		"_xZZZZ_":           "_xZZZZ_",
	} {
		if got := unescapeOOXML(in); got != want {
			t.Errorf("unescapeOOXML(%q) = %q, want %q", in, got, want)
		}
	}
}

// Workbooks using the 1904 date system get dates that read back the same.
func TestDate1904(t *testing.T) {
	path := filepath.Join(t.TempDir(), "mac.xlsx")
	f := excelize.NewFile()
	on := true
	if err := f.SetWorkbookProps(&excelize.WorkbookPropsOptions{Date1904: &on}); err != nil {
		t.Fatal(err)
	}
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	e, err := OpenEditor(path, false)
	if err != nil {
		t.Fatal(err)
	}
	if err := e.Set("Sheet1", 1, 1, input.Parse("2026-07-01", false)); err != nil {
		t.Fatal(err)
	}
	if err := e.Save(false); err != nil {
		t.Fatal(err)
	}
	wb, err := Read(path)
	if err != nil {
		t.Fatal(err)
	}
	c := wb.Sheets[0].At(1, 1)
	if got := grid.Display(c, wb.FormatOf(c), wb.Date1904); got != "2026-07-01" {
		t.Errorf("date = %s", got)
	}
}

func sheetXML(t *testing.T, path, part string) string {
	t.Helper()
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	p, err := parts(data)
	if err != nil {
		t.Fatal(err)
	}
	return string(p[part])
}

func TestModernFormulas(t *testing.T) {
	path := writeBook(t, [][]string{{"Name", "Score"}, {"Ada", "9"}, {"Bob", "4"}, {"Ada", "8"}})
	e, err := OpenEditor(path, false)
	if err != nil {
		t.Fatal(err)
	}
	for cell, f := range map[string]string{"D1": "=UNIQUE(A2:A4)", "E1": "=XLOOKUP(\"Bob\",A2:A4,B2:B4)", "F1": "=LET(x,B2,x*2)"} {
		r, c, _ := grid.ParseCell(cell)
		if err := e.Set("Data", r, c, input.Parse(f, false)); err != nil {
			t.Fatal(err)
		}
	}
	if v, _, ok := e.Evaluate("Data", 1, 5); !ok || v.Num != 4 {
		t.Errorf("XLOOKUP = %+v %v", v, ok)
	}
	if v, _, ok := e.Evaluate("Data", 1, 6); !ok || v.Num != 18 {
		t.Errorf("LET = %+v %v", v, ok)
	}
	if err := e.Save(false); err != nil {
		t.Fatal(err)
	}
	xml := sheetXML(t, path, "xl/worksheets/sheet1.xml")
	for _, want := range []string{
		`<c r="D1" cm="1" t="str"><f t="array" ref="D1:D2">_xlfn.UNIQUE(A2:A4)</f><v>Ada</v>`, // spilled, as Excel saves it
		`<f>_xlfn.XLOOKUP(&#34;Bob&#34;,A2:A4,B2:B4)</f><v>4</v>`,
		`<f>_xlfn.LET(_xlpm.x,B2,_xlpm.x*2)</f><v>18</v>`, // computed inlined
	} {
		if !strings.Contains(xml, want) {
			t.Errorf("missing %s in\n%s", want, xml)
		}
	}
	if md := sheetXML(t, path, "xl/metadata.xml"); !strings.Contains(md, `name="XLDAPR"`) {
		t.Errorf("metadata = %q", md)
	}
	wb, err := Read(path)
	if err != nil {
		t.Fatal(err)
	}
	if got := wb.Sheets[0].Get(1, 6).Formula; got != "LET(x,B2,x*2)" {
		t.Errorf("read back %q", got)
	}
}

// An Excel 365 file whose dynamic array has spilled: overwriting the
// anchor must clear the spilled values instead of leaving them as constants.
func TestOverwriteSpillAnchor(t *testing.T) {
	path := writeBook(t, [][]string{{"Name"}, {"Ada"}, {"Bob"}})
	e, _ := OpenEditor(path, false)
	e.Set("Data", 1, 3, input.Parse("=UNIQUE(A2:A3)", false))
	e.Set("Data", 2, 3, input.Parse("Bob", false)) // the spilled value, as Excel stores it
	if err := e.Save(false); err != nil {
		t.Fatal(err)
	}
	rewritePart(t, path, "xl/worksheets/sheet1.xml", func(s string) string {
		return strings.Replace(s, `<f t="array" ref="C1">`, `<f t="array" ref="C1:C2">`, 1)
	})
	e, err := OpenEditor(path, false)
	if err != nil {
		t.Fatal(err)
	}
	if err := e.Set("Data", 1, 3, input.Parse("done", false)); err != nil {
		t.Fatal(err)
	}
	if err := e.Save(false); err != nil {
		t.Fatal(err)
	}
	wb, _ := Read(path)
	if c := wb.Sheets[0].Get(2, 3); !c.IsEmpty() {
		t.Errorf("C2 kept spilled value %+v", c)
	}
	if strings.Contains(sheetXML(t, path, "xl/worksheets/sheet1.xml"), `cm="`) {
		t.Error("cm left on overwritten anchor")
	}
}

func rewritePart(t *testing.T, path, name string, fn func(string) string) {
	t.Helper()
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatal(err)
	}
	var out bytes.Buffer
	zw := zip.NewWriter(&out)
	for _, zf := range zr.File {
		r, _ := zf.Open()
		b, _ := io.ReadAll(r)
		r.Close()
		if zf.Name == name {
			b = []byte(fn(string(b)))
		}
		w, _ := zw.Create(zf.Name)
		w.Write(b)
	}
	zw.Close()
	if err := os.WriteFile(path, out.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
}

func TestFormatPropagatesTypes(t *testing.T) {
	path := writeBook(t, [][]string{
		{"Start", "Days", "Due", "Units", "Price", "Revenue", "Tax", "Per unit", "Year", "Count", "Span", "Notes"},
		{"2026-07-01", "30", "=A{r}+B{r}", "3", "$2.50", "=D{r}*E{r}", "=ROUND(F{r}*0.2,2)", "=F{r}/D{r}", "=YEAR(A{r})", "=COUNTIF(E:E,E{r})", "=C{r}-A{r}", strings.Repeat("long note ", 10)},
		{"2026-07-05", "10", "=A{r}+B{r}", "4", "$3.00", "=D{r}*E{r}", "=ROUND(F{r}*0.2,2)", "=F{r}/D{r}", "=YEAR(A{r})", "=COUNTIF(E:E,E{r})", "=C{r}-A{r}", "short"},
	})
	res, err := Format(path, FormatOptions{Preset: style.Default})
	if err != nil {
		t.Fatal(err)
	}
	got := map[string]string{}
	for _, c := range res.Sheets[0].Tables[0].Columns {
		got[c.Header] = c.Type
	}
	for col, want := range map[string]string{
		"Due": "date", "Revenue": "currency $", "Tax": "currency $", "Per unit": "decimal",
		"Year": "integer", "Count": "integer", "Span": "integer",
	} {
		if got[col] != want {
			t.Errorf("%s = %s, want %s", col, got[col], want)
		}
	}
	f, _ := excelize.OpenFile(path)
	defer f.Close()
	id, _ := f.GetCellStyle("Data", "L2")
	st, _ := f.GetStyle(id)
	if st.Alignment == nil || !st.Alignment.WrapText || st.Alignment.Vertical != "top" {
		t.Errorf("long notes not wrapped: %+v", st.Alignment)
	}
	// Headers are left-aligned, numbers below or not: the filter button
	// takes the right edge.
	id, _ = f.GetCellStyle("Data", "F1")
	if st, _ = f.GetStyle(id); st.Alignment == nil || st.Alignment.Horizontal != "left" {
		t.Errorf("Revenue header alignment: %+v", st.Alignment)
	}
	if ok, _, _ := Check(path, FormatOptions{Preset: style.Default}); !ok {
		t.Error("not idempotent")
	}
}

// Deleting a row that formulas point into gives #REF!, as Excel does;
// excelize alone would point them at the row above.
func TestDeleteRowsBreaksReferences(t *testing.T) {
	path := writeBook(t, [][]string{
		{"v", "", "", "", "=B5*2"},
		{"", "", "", "", "=SUM(B4:B6)"},
		{"", "", "", "", "=SUM(B5:B8)"},
		{"", "4"},
		{"", "5"},
		{"", "6"},
		{"", "7"},
		{"", "8"},
	})
	e, _ := OpenEditor(path, false)
	e.EnsureSheet("Other")
	e.Set("Other", 1, 1, input.Parse("=Data!B5*10", false))
	e.Set("Other", 2, 1, input.Parse("=Data!B6", false))
	e.Save(false)
	e, err := OpenEditor(path, false)
	if err != nil {
		t.Fatal(err)
	}
	if err := e.DeleteRows("Data", 5, 5); err != nil {
		t.Fatal(err)
	}
	if err := e.Save(false); err != nil {
		t.Fatal(err)
	}
	wb, _ := Read(path)
	data, other := wb.Sheet("Data"), wb.Sheet("Other")
	for _, c := range []struct {
		got, want string
	}{
		{data.Get(1, 5).Formula, "#REF!*2"},
		{data.Get(2, 5).Formula, "SUM(B4:B5)"},
		{data.Get(3, 5).Formula, "SUM(B5:B7)"},
		{other.Get(1, 1).Formula, "Data!#REF!*10"},
		{other.Get(2, 1).Formula, "Data!B5"},
	} {
		if c.got != c.want {
			t.Errorf("got %s, want %s", c.got, c.want)
		}
	}
	if w := strings.Join(e.Warnings(), " "); !strings.Contains(w, "#REF!") || !strings.Contains(w, "Other!A1") {
		t.Errorf("warnings = %q", w)
	}
}

// fixture copies a workbook from testdata (made by openpyxl, see
// scripts/fixtures.py) into a temporary directory.
func fixture(t *testing.T, name string) string {
	t.Helper()
	data, err := os.ReadFile(filepath.Join("testdata", name))
	if err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(t.TempDir(), name)
	if err := os.WriteFile(path, data, 0o644); err != nil {
		t.Fatal(err)
	}
	return path
}

// Moving rows rewrites the chart's drawing; excelize drops the namespace
// declaration Excel and openpyxl put on c:chart, which Excel then repairs.
// Renaming the sheet must reach the chart's series.
func TestStructureKeepsCharts(t *testing.T) {
	path := fixture(t, "rich.xlsx")
	e, err := OpenEditor(path, false)
	if err != nil {
		t.Fatal(err)
	}
	if err := e.InsertRows("Data", 2, 1); err != nil {
		t.Fatal(err)
	}
	if err := e.RenameSheet("Data", "Sales Data"); err != nil {
		t.Fatal(err)
	}
	if err := e.Save(false); err != nil {
		t.Fatal(err)
	}
	drawing := sheetXML(t, path, "xl/drawings/drawing1.xml")
	if strings.Contains(drawing, "<c:chart") && !strings.Contains(drawing, `xmlns:c="`) {
		t.Errorf("drawing uses c: without declaring it:\n%s", drawing)
	}
	// Renamed, and moved down by the row inserted above the data.
	if chart := sheetXML(t, path, "xl/charts/chart1.xml"); !strings.Contains(chart, "&#39;Sales Data&#39;!$B$3:$B$5") {
		t.Errorf("chart series not renamed and moved:\n%s", chart)
	}
	wb, err := Read(path)
	if err != nil {
		t.Fatal(err)
	}
	if wb.Sheet("Sales Data") == nil || wb.Sheet("T").Tables[0].Name != "Things" {
		t.Errorf("sheets = %v", wb.SheetNames())
	}
}

func TestDeclareNamespaces(t *testing.T) {
	in := `<?xml version="1.0"?>` + "\n" + `<xdr:wsDr xmlns:xdr="x"><c:chart r:id="rId1"/></xdr:wsDr>`
	out := string(declareNamespaces([]byte(in)))
	if !strings.Contains(out, `<xdr:wsDr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xdr="x">`) {
		t.Errorf("got %s", out)
	}
	if got := string(declareNamespaces([]byte(out))); got != out {
		t.Error("not idempotent")
	}
	// A declaration on an inner element does not cover other elements.
	in = `<xdr:wsDr xmlns:xdr="x"><c:chart r:id="1"/><c:chart xmlns:c="y" r:id="2"/></xdr:wsDr>`
	if out := string(declareNamespaces([]byte(in))); !strings.HasPrefix(out, `<xdr:wsDr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r=`) {
		t.Errorf("inner declaration taken for a root one: %s", out)
	}
}

// Formatting a workbook styled by another tool.
func TestFormatForeignWorkbook(t *testing.T) {
	path := fixture(t, "styled.xlsx")
	res, err := Format(path, FormatOptions{Preset: style.Default})
	if err != nil {
		t.Fatal(err)
	}
	var types []string
	for _, c := range res.Sheets[0].Tables[0].Columns {
		types = append(types, c.Type)
	}
	if got := strings.Join(types, ","); got != "text,date,currency €,currency €,percent" {
		t.Errorf("types = %s", got)
	}
	if ok, _, _ := Check(path, FormatOptions{Preset: style.Default}); !ok {
		t.Error("not idempotent")
	}
	f, _ := excelize.OpenFile(path)
	defer f.Close()
	id, _ := f.GetCellStyle("Budget 2026", "A1")
	st, _ := f.GetStyle(id)
	if st.Font == nil || st.Font.Family != "Calibri" || !st.Font.Bold || st.Font.Size != 14 {
		t.Errorf("title font = %+v", st.Font)
	}
	if props, _ := f.GetSheetProps("Budget 2026"); props.TabColorRGB == nil {
		t.Error("tab colour lost")
	}
}

// Group labels centre across the columns they group, with no borders
// inside the group.
func TestFormatGroupRow(t *testing.T) {
	path := writeBook(t, [][]string{
		{"Region", "Q1", "", "Q2", ""},
		{"Region", "Units", "Revenue", "Units", "Revenue"},
		{"North", "5", "$100", "6", "$120"},
		{"South", "7", "$140", "8", "$160"},
	})
	res, err := Format(path, FormatOptions{Preset: style.Default})
	if err != nil {
		t.Fatal(err)
	}
	if tb := res.Sheets[0].Tables[0]; tb.Columns[2].Header != "Q1 / Revenue" {
		t.Fatalf("columns = %+v", tb.Columns)
	}
	f, _ := excelize.OpenFile(path)
	defer f.Close()
	look := func(cell string) (string, string) {
		id, _ := f.GetCellStyle("Data", cell)
		st, _ := f.GetStyle(id)
		var sides []string
		for _, b := range st.Border {
			sides = append(sides, b.Type[:1])
		}
		slices.Sort(sides)
		return st.Alignment.Horizontal, strings.Join(sides, "")
	}
	for cell, want := range map[string][2]string{
		"A1": {"center", "blrt"},
		"B1": {"centerContinuous", "blt"},
		"C1": {"centerContinuous", "brt"},
		"D1": {"centerContinuous", "blt"},
		"E1": {"centerContinuous", "brt"},
	} {
		if a, s := look(cell); a != want[0] || s != want[1] {
			t.Errorf("%s: %s %s, want %s %s", cell, a, s, want[0], want[1])
		}
	}
	if ok, _, _ := Check(path, FormatOptions{Preset: style.Default}); !ok {
		t.Error("not idempotent")
	}
}

func TestReadHiddenRowsAndCols(t *testing.T) {
	f := excelize.NewFile()
	for r := 1; r <= 6; r++ {
		f.SetCellValue("Sheet1", "A"+strconv.Itoa(r), r)
	}
	f.SetRowVisible("Sheet1", 3, false)
	f.SetRowVisible("Sheet1", 4, false)
	f.SetRowVisible("Sheet1", 6, false)
	f.SetColVisible("Sheet1", "C:E", false)
	var buf bytes.Buffer
	if err := f.Write(&buf); err != nil {
		t.Fatal(err)
	}
	wb, err := ReadBytes(buf.Bytes())
	if err != nil {
		t.Fatal(err)
	}
	s := wb.Sheets[0]
	if want := []grid.Run{{From: 3, To: 4}, {From: 6, To: 6}}; !equalRuns(s.HiddenRows, want) {
		t.Errorf("hidden rows %v, want %v", s.HiddenRows, want)
	}
	if want := []grid.Run{{From: 3, To: 5}}; !equalRuns(s.HiddenCols, want) {
		t.Errorf("hidden cols %v, want %v", s.HiddenCols, want)
	}
	if !s.RowHidden(4) || s.RowHidden(5) || !s.ColHidden(5) || s.ColHidden(2) {
		t.Error("RowHidden/ColHidden disagree with the runs")
	}
}

func equalRuns(a, b []grid.Run) bool {
	if len(a) != len(b) {
		return false
	}
	for i := range a {
		if a[i] != b[i] {
			return false
		}
	}
	return true
}

// A large file is cached parsed: the second read decodes the same grid,
// and formulas without values still compute.
func TestReadCache(t *testing.T) {
	t.Setenv("SHEETS_CACHE_DIR", t.TempDir())
	f := excelize.NewFile()
	sw, err := f.NewStreamWriter("Sheet1")
	if err != nil {
		t.Fatal(err)
	}
	sw.SetRow("A1", []any{"Order", "Customer", "Units", "Price", "Revenue"})
	for r := 2; r <= 40000; r++ {
		sw.SetRow("A"+strconv.Itoa(r), []any{"ORD-" + strconv.Itoa(r), "Customer " + strconv.Itoa(r%997), r % 17, float64(r%100) / 4,
			excelize.Cell{Formula: "C" + strconv.Itoa(r) + "*D" + strconv.Itoa(r)}})
	}
	if err := sw.Flush(); err != nil {
		t.Fatal(err)
	}
	var buf bytes.Buffer
	if err := f.Write(&buf); err != nil {
		t.Fatal(err)
	}
	data := buf.Bytes()
	if len(data) < cache.MinSize {
		t.Fatalf("fixture is %d bytes, under the cache threshold", len(data))
	}
	first, err := ReadBytes(data)
	if err != nil {
		t.Fatal(err)
	}
	if _, ok := cache.Get(cache.Key("xlsx", data)); !ok {
		t.Fatal("not cached")
	}
	second, err := ReadBytes(data)
	if err != nil {
		t.Fatal(err)
	}
	a, b := first.Sheets[0], second.Sheets[0]
	if a.Used() != b.Used() || second.Resolver == nil {
		t.Fatalf("cached read differs: used %v vs %v, resolver %v", a.Used(), b.Used(), second.Resolver)
	}
	for r := 1; r <= 40000; r++ {
		for c := 1; c <= 5; c++ {
			if a.Get(r, c) != b.Get(r, c) {
				t.Fatalf("%s: %+v vs %+v", grid.CellName(r, c), a.Get(r, c), b.Get(r, c))
			}
		}
	}
	if v := second.Value(b, 3, 5); v.Kind != grid.Number || v.Num != 3*0.75 {
		t.Errorf("E3 = %+v, want 2.25", v)
	}
}

// openpyxl writes its default fill as <patternFill/>, which excelize reads
// back as an invalid pattern; typed values with a format still go in.
func TestNumFmtOnOpenpyxlStyles(t *testing.T) {
	path := fixture(t, "tables.xlsx")
	e, err := OpenEditor(path, false)
	if err != nil {
		t.Fatal(err)
	}
	if err := e.Set("Sales", 9, 1, input.Parse("$120.00", false)); err != nil {
		t.Fatal(err)
	}
	if err := e.Save(false); err != nil {
		t.Fatal(err)
	}
}

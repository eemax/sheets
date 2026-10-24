package xlsx

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"os"
	"regexp"
	"strconv"
	"strings"
	"testing"

	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/input"
	"github.com/eemax/sheets/internal/style"
	"github.com/xuri/excelize/v2"
)

// stored returns the result a cell stores in the file ("" when none), the
// way pandas or a preview reads it.
func stored(t *testing.T, path, part, cell string) string {
	t.Helper()
	x := sheetXML(t, path, part)
	m := regexp.MustCompile(`<c r="` + cell + `"[^>]*>(.*?)</c>`).FindStringSubmatch(x)
	if m == nil {
		t.Fatalf("no cell %s in %s", cell, part)
	}
	if v := regexp.MustCompile(`<v>(.*?)</v>`).FindStringSubmatch(m[1]); v != nil {
		return v[1]
	}
	return ""
}

const dataPart, summaryPart = "xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"

// Changing an input stores new results for everything that reads it,
// through totals, other sheets and names; what it cannot compute loses
// its result; what does not read it keeps its result as Excel stored it.
func TestStoredResultsFollowEdits(t *testing.T) {
	path := fixture(t, "cached.xlsx")
	e := edit(t, path, func(e *Editor) error { return e.Set("Data", 2, 2, input.Parse("100", false)) })
	for _, tt := range []struct{ part, cell, want string }{
		{dataPart, "B5", "110"},
		{summaryPart, "A1", "220"},  // Data!B5*Rate
		{summaryPart, "A2", "1000"}, // LET, computed inlined
		{summaryPart, "A3", "8"},    // reads B3 only: kept
		{summaryPart, "A4", "3"},
		{summaryPart, "A5", ""}, // XMATCH: the calculator cannot, Excel will
	} {
		if got := stored(t, path, tt.part, tt.cell); got != tt.want {
			t.Errorf("%s %s stores %q, want %q", tt.part, tt.cell, got, tt.want)
		}
	}
	if w := strings.Join(e.Warnings(), "\n"); !strings.Contains(w, "Summary!A5 (XMATCH is not supported") {
		t.Errorf("warnings do not name the result left to Excel: %s", w)
	}

	// Read back, the stored results are trusted; the missing one is
	// computed on demand.
	wb, err := Read(path)
	if err != nil {
		t.Fatal(err)
	}
	s := wb.Sheet("Summary")
	if c := s.At(1, 1); c.Pending || c.Num != 220 {
		t.Errorf("Summary!A1 read back as %+v", c)
	}
	if c := s.At(5, 1); !c.Pending {
		t.Errorf("Summary!A5 read back as %+v, want no stored result", c)
	}
}

// Moving rows changes what ranges and positions mean: those results are
// recomputed; a reference into deleted rows stores #REF!.
func TestStoredResultsFollowRows(t *testing.T) {
	path := fixture(t, "cached.xlsx")
	edit(t, path, func(e *Editor) error { return e.DeleteRows("Data", 3, 3) }) // South, 7
	for _, tt := range []struct{ part, cell, want string }{
		{dataPart, "B4", "8"}, // the total, moved up: =SUM(B2:B3)
		{summaryPart, "A1", "16"},
		{summaryPart, "A3", "#REF!"},
		{summaryPart, "A4", "2"}, // ROWS of a range that lost a row
	} {
		if got := stored(t, path, tt.part, tt.cell); got != tt.want {
			t.Errorf("%s %s stores %q, want %q", tt.part, tt.cell, got, tt.want)
		}
	}

	// Inserting rows inside the total's range: same total, recomputed.
	edit(t, path, func(e *Editor) error { return e.InsertRows("Data", 3, 2) })
	if got := stored(t, path, dataPart, "B6"); got != "8" {
		t.Errorf("total after inserting rows stores %q", got)
	}
}

// A file whose results are missing (openpyxl writes none) gets them on the
// next save, as far as the calculator can.
func TestStoredResultsFilledIn(t *testing.T) {
	f := excelize.NewFile()
	f.SetCellValue("Sheet1", "A1", 4)
	f.SetCellFormula("Sheet1", "B1", "A1*2")
	f.SetCellFormula("Sheet1", "C1", "B1+1")
	on := true
	f.SetCalcProps(&excelize.CalcPropsOptions{FullCalcOnLoad: &on})
	path := t.TempDir() + "/bare.xlsx"
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	edit(t, path, func(e *Editor) error { return e.Set("Sheet1", 5, 5, input.Parse("note", false)) })
	if b, c := stored(t, path, dataPart, "B1"), stored(t, path, dataPart, "C1"); b != "8" || c != "9" {
		t.Errorf("filled in B1=%q C1=%q", b, c)
	}
}

// The stamp vouches for results sheets stored; an edit by another tool
// that leaves the recalculation flag on voids it.
func TestStamp(t *testing.T) {
	path := fixture(t, "cached.xlsx")
	edit(t, path, func(e *Editor) error { return e.Set("Data", 2, 2, input.Parse("100", false)) })
	pending := func() bool {
		t.Helper()
		wb, err := Read(path)
		if err != nil {
			t.Fatal(err)
		}
		return wb.Sheet("Data").At(5, 2).Pending
	}
	if pending() {
		t.Fatal("results stored by sheets are not trusted")
	}

	// fmt restyles but changes no result: the stamp is renewed, and
	// fmt --check finds nothing to do afterwards.
	if _, err := Format(path, FormatOptions{Preset: style.Default}); err != nil {
		t.Fatal(err)
	}
	if pending() {
		t.Error("fmt voided the stamp")
	}
	if ok, _, err := Check(path, FormatOptions{Preset: style.Default}); err != nil || !ok {
		t.Errorf("fmt --check after fmt: %v %v", ok, err)
	}

	// Another tool changes an input and saves without new results.
	f, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	f.SetCellValue("Data", "B3", 1000)
	if err := f.Save(); err != nil {
		t.Fatal(err)
	}
	f.Close()
	if !pending() {
		t.Error("a stale result was trusted after another tool's edit")
	}
}

// Results the calculator may get wrong are shown but not stored.
func TestUncertainNotStored(t *testing.T) {
	path := writeBook(t, [][]string{{"Units", "Price", "Total"}, {"5", "2", "=A2*B2"}})
	edit(t, path, func(e *Editor) error { return e.Set("Data", 2, 1, input.Parse("1,200", true)) })
	if got := stored(t, path, dataPart, "C2"); got != "" {
		t.Errorf("C2 stores %q from text Excel may read as a number", got)
	}
}

// Out of time, results not computed are removed rather than left stale.
func TestRecalcBudget(t *testing.T) {
	old := recalcBudget
	recalcBudget = 0
	defer func() { recalcBudget = old }()
	path := fixture(t, "cached.xlsx")
	e := edit(t, path, func(e *Editor) error { return e.Set("Data", 2, 2, input.Parse("100", false)) })
	if got := stored(t, path, dataPart, "B5"); got != "" {
		t.Errorf("B5 keeps %q though it was not recomputed", got)
	}
	if got := stored(t, path, summaryPart, "A3"); got != "8" {
		t.Errorf("A3, untouched by the edit, lost its result: %q", got)
	}
	if w := strings.Join(e.Warnings(), "\n"); !strings.Contains(w, "stopped computing") {
		t.Errorf("warnings: %s", w)
	}
}

// Recomputing an array formula clears the rest of its old result.
func TestStoreResultsArray(t *testing.T) {
	in := []byte(`<row r="1"><c r="D1" cm="1" t="str"><f ref="D1:D3" t="array">UNIQUE(A1:A9)</f><v>Ada</v></c></row>` +
		`<row r="2"><c r="D2" s="4" t="str"><v>Bob</v></c><c r="E2"><v>5</v></c></row><row r="3"><c r="D3" t="s"><v>7</v></c></row>`)
	out := string(storeResults(in, map[string]grid.Cell{"D1": {}}))
	want := `<row r="1"><c r="D1" cm="1"><f ref="D1:D3" t="array">UNIQUE(A1:A9)</f></c></row>` +
		`<row r="2"><c r="D2" s="4"></c><c r="E2"><v>5</v></c></row><row r="3"><c r="D3"></c></row>`
	if out != want {
		t.Errorf("got\n%s\nwant\n%s", out, want)
	}
}

func TestStoreResultsXML(t *testing.T) {
	in := []byte(`<row r="1"><c r="A1" s="3" t="str"><f>B1*2</f><v>old</v></c><c r="B1"><v>2</v></c>` +
		`<c r="C1" t="e" cm="1"><f t="shared" si="0"/><v>#DIV/0!</v></c><c r="D1"/><c r="E1" t="str"><f>X</f></c></row>`)
	out := string(storeResults(in, map[string]grid.Cell{
		"A1": {Kind: grid.Number, Num: 4}, "B1": {Kind: grid.Number, Num: 9}, "C1": {Kind: grid.Text, Text: "a<b"},
		"D1": {Kind: grid.Number, Num: 1}, "E1": {},
	}))
	want := `<row r="1"><c r="A1" s="3"><f>B1*2</f><v>4</v></c><c r="B1"><v>2</v></c>` +
		`<c r="C1" cm="1" t="str"><f t="shared" si="0"/><v>a&lt;b</v></c><c r="D1"/><c r="E1"><f>X</f></c></row>`
	if out != want {
		t.Errorf("got\n%s\nwant\n%s", out, want)
	}
}

func TestFreshCaches(t *testing.T) {
	in := `<c:ser><c:tx><c:strRef><c:f>Data!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Units</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
		`<c:cat><c:strRef><c:f>Data!$A$2:$A$4</c:f><c:strCache><c:ptCount val="3"/></c:strCache></c:strRef></c:cat>` +
		`<c:val><c:numRef><c:f>Data!$B$2:$B$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/></c:numCache></c:numRef></c:val></c:ser>` +
		`<val><numRef><f>Other!$B$2:$B$4</f><numCache/></numRef></val>`
	out := string(freshCaches([]byte(in), nil, map[string][][2]int{"data": {{3, 2}}}))
	// The values took in the changed B3 and lose their cache; the series
	// name and categories keep theirs, as does the other sheet's series.
	want := `<c:ser><c:tx><c:strRef><c:f>Data!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Units</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
		`<c:cat><c:strRef><c:f>Data!$A$2:$A$4</c:f><c:strCache><c:ptCount val="3"/></c:strCache></c:strRef></c:cat>` +
		`<c:val><c:numRef><c:f>Data!$B$2:$B$4</c:f></c:numRef></c:val></c:ser>` +
		`<val><numRef><f>Other!$B$2:$B$4</f><numCache/></numRef></val>`
	if out != want {
		t.Errorf("got\n%s\nwant\n%s", out, want)
	}
}

// Results are stored with the type Excel gives them, as far as can be
// told; circular references and the calculator's #VALUE! are not stored.
func TestStoredTypes(t *testing.T) {
	path := writeBook(t, [][]string{
		{"Code", "Qty", "Text", "Flag", "Lookup", "Loop", "Value"},
		{"007", "5", `=TEXT(B2,"000")`, "=IF(B2>3,TRUE)", "=INDEX(A2:B2,1)", "=F2+1", `=VALUE(" 42 ")`},
		{"x", "6", `=B3&""`, "=B3>3", "=INDEX(B2:B3,2)", "=G3*2", "=1"},
	})
	e := edit(t, path, func(e *Editor) error { return e.Set("Data", 2, 2, input.Parse("4", false)) })
	for cell, want := range map[string]string{
		"C2": `t="str"><f>TEXT(B2,&#34;000&#34;)</f><v>004</v>`,
		"D2": `t="b"><f>IF(B2&gt;3,TRUE)</f><v>1</v>`,
		"E2": `<f>INDEX(A2:B2,1)</f></c>`, // "007" or 7: left to Excel
		"F2": `<f>F2+1</f></c>`,           // refers to itself
		"G2": `<f>VALUE(&#34; 42 &#34;)</f></c>`,
		"C3": `t="str"><f>B3&amp;&#34;&#34;</f><v>6</v>`,
		"E3": `<f>INDEX(B2:B3,2)</f><v>6</v>`,
	} {
		x := sheetXML(t, path, dataPart)
		m := regexp.MustCompile(`<c r="` + cell + `"[^>]*>.*?</c>`).FindString(x)
		if !strings.Contains(m, want) {
			t.Errorf("%s = %s, want it to hold %s", cell, m, want)
		}
	}
	if w := strings.Join(e.Warnings(), "\n"); !strings.Contains(w, "circular reference: Data!F2 refers to itself") {
		t.Errorf("warnings: %s", w)
	}
}

func TestCycles(t *testing.T) {
	path := writeBook(t, [][]string{
		{"=C1+1", "x", "=A1+1", "=A1*2", "=SUM(E2:E9)"}, // A1 and C1 loop; D1 reads the loop
		{"1", "", "", "", "=E1*2"},                      // E1 sums E2, which reads E1: a loop through a range
		{"2", "", "", "", "3"},
		{"=SUM(A2:A3)", "", "", "", ""}, // no loop
	})
	e := edit(t, path, func(e *Editor) error { return e.Set("Data", 3, 1, input.Parse("5", false)) })
	for cell, want := range map[string]string{"A1": "", "C1": "", "D1": "", "E1": "", "E2": "", "A4": "6"} {
		if got := stored(t, path, dataPart, cell); got != want {
			t.Errorf("%s stores %q, want %q", cell, got, want)
		}
	}
	if w := strings.Join(e.Warnings(), "\n"); !strings.Contains(w, "circular reference: 4 formulas refer to each other") ||
		!strings.Contains(w, "Data!A1, Data!C1, Data!E1, Data!E2") {
		t.Errorf("warnings: %s", w)
	}
}

// Reading results the file lacks follows the same rules as storing them.
func TestReadComputedTypes(t *testing.T) {
	f := excelize.NewFile()
	f.SetCellValue("Sheet1", "A1", 5)
	for cell, fm := range map[string]string{"B1": "C1+1", "C1": "B1+1", "D1": `TEXT(A1,"000")`, "E1": "IF(A1>3,TRUE)", "F1": "SUM(A1,#REF!)", "G1": "A1*2"} {
		f.SetCellFormula("Sheet1", cell, fm)
	}
	var buf bytes.Buffer
	if err := f.Write(&buf); err != nil {
		t.Fatal(err)
	}
	wb, err := ReadBytes(buf.Bytes())
	if err != nil {
		t.Fatal(err)
	}
	s := wb.Sheets[0]
	for col, want := range map[int]string{2: "?", 3: "?", 4: `"005"`, 5: "TRUE", 6: "?", 7: "10"} {
		v := wb.Value(s, 1, col)
		got := "?"
		switch {
		case v.Pending:
		case v.Kind == grid.Text:
			got = strconv.Quote(v.Text)
		default:
			got = grid.Display(v, wb.FormatOf(v), false)
		}
		if got != want {
			t.Errorf("%s = %s, want %s", grid.CellName(1, col), got, want)
		}
	}
}

// A result counts only when everything it reads does: a formula reading
// one the calculator cannot compute gets no result.
func TestTaintedResults(t *testing.T) {
	path := t.TempDir() + "/dep.xlsx"
	e, _ := OpenEditor(path, true)
	e.EnsureSheet("Data")
	for cell, f := range map[[2]int]string{{1, 1}: "=XMATCH(2,{1,2,3})", {1, 2}: "=A1+1", {1, 3}: "=SUM(A1:A2)*2", {1, 4}: "=B1*10", {2, 5}: "=5*2"} {
		if err := e.Set("Data", cell[0], cell[1], input.Parse(f, false)); err != nil {
			t.Fatal(err)
		}
	}
	if err := e.Save(false); err != nil {
		t.Fatal(err)
	}
	for cell, want := range map[string]string{"A1": "", "B1": "", "C1": "", "D1": "", "E2": "10"} {
		if got := stored(t, path, dataPart, cell); got != want {
			t.Errorf("%s stores %q, want %q", cell, got, want)
		}
	}
	if _, note, ok := e.Evaluate("Data", 1, 4); ok || !strings.Contains(note, "Data!A1") {
		t.Errorf("D1: %v %q", ok, note)
	}

	// Reading the file back says the same.
	wb, _ := Read(path)
	if v := wb.Value(wb.Sheet("Data"), 1, 3); !v.Pending {
		t.Errorf("C1 read back as %+v", v)
	}
}

// Where the calculator disagrees with a result Excel stored, nothing
// computed from it is trusted.
func TestCrossCheck(t *testing.T) {
	f := excelize.NewFile()
	f.SetCellValue("Sheet1", "A1", 5)
	f.SetCellValue("Sheet1", "C1", 1)
	f.SetCellFormula("Sheet1", "A2", "A1*2")
	f.SetCellFormula("Sheet1", "B2", "A2+C1")
	var buf bytes.Buffer
	f.Write(&buf)
	// As if Excel had stored 999 for A2 (the calculator says 10).
	data := patchZip(t, buf.Bytes(), dataPart, func(x string) string {
		x = regexp.MustCompile(`<c r="A2"[^>]*><f>A1\*2</f>`).ReplaceAllLiteralString(x, `<c r="A2"><f>A1*2</f><v>999</v>`)
		return regexp.MustCompile(`<c r="B2"[^>]*><f>A2\+C1</f>`).ReplaceAllLiteralString(x, `<c r="B2"><f>A2+C1</f><v>1000</v>`)
	})
	path := t.TempDir() + "/x.xlsx"
	os.WriteFile(path, data, 0o644)
	e := edit(t, path, func(e *Editor) error { return e.Set("Sheet1", 1, 3, input.Parse("2", false)) })
	if got := stored(t, path, dataPart, "B2"); got != "" {
		t.Errorf("B2 stores %q, computed from a value the calculator disputes", got)
	}
	if got := stored(t, path, dataPart, "A2"); got != "999" {
		t.Errorf("A2, not reached by the edit, changed to %q", got)
	}
	if _, note, _ := e.Evaluate("Sheet1", 2, 2); !strings.Contains(note, "where Excel stored 999") {
		t.Errorf("B2 note: %q", note)
	}
}

func patchZip(t *testing.T, data []byte, part string, fn func(string) string) []byte {
	t.Helper()
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
		if zf.Name == part {
			b = []byte(fn(string(b)))
		}
		w, _ := zw.Create(zf.Name)
		w.Write(b)
	}
	zw.Close()
	return out.Bytes()
}

// Table references are computed, by the cells they stand for.
func TestTableResults(t *testing.T) {
	path := fixture(t, "tables.xlsx") // SalesT: A1:C5, totals row 5; E2 =SUM(SalesT[Units])
	edit(t, path, func(e *Editor) error { return e.Set("Sales", 2, 2, input.Parse("10", false)) })
	if b5, e2 := stored(t, path, dataPart, "B5"), stored(t, path, dataPart, "E2"); b5 != "19" || e2 != "19" {
		t.Errorf("totals row B5 = %q, E2 = %q; want 19", b5, e2)
	}
}

// What the calculator cannot compute but Excel did stands at Excel's
// result for formulas reading it.
func TestFrozenResults(t *testing.T) {
	path := fixture(t, "cached.xlsx") // Summary!A5 =XMATCH(5,Data!B2:B4), stored 1
	edit(t, path, func(e *Editor) error { return e.Set("Summary", 5, 2, input.Parse("=A5*100", false)) })
	if got := stored(t, path, summaryPart, "B5"); got != "100" {
		t.Errorf("B5 stores %q, want 100", got)
	}
	if got := stored(t, path, summaryPart, "A5"); got != "1" {
		t.Errorf("A5 keeps %q, want Excel's 1", got)
	}
}

// Renamed and deleted sheets: results are computed with the formulas as
// they are written, which follow the sheets.
func TestResultsAfterSheetEdits(t *testing.T) {
	f := excelize.NewFile()
	f.SetSheetName("Sheet1", "Data")
	f.NewSheet("Summary")
	f.NewSheet("Old")
	f.SetCellValue("Data", "B2", 5)
	f.SetCellValue("Old", "A1", 1)
	f.SetCellFormula("Summary", "A1", "Data!B2*2")
	f.SetCellFormula("Summary", "A2", "SUM(Data!B2:B3)")
	f.SetCellFormula("Summary", "A3", "Old!A1+1")
	path := t.TempDir() + "/s.xlsx"
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	// No stored results: the save computes them, after the rename and the
	// deletion.
	edit(t, path, func(e *Editor) error {
		if err := e.RenameSheet("Data", "Sales 2026"); err != nil {
			return err
		}
		return e.DeleteSheet("Old")
	})
	for cell, want := range map[string]string{"A1": "10", "A2": "5", "A3": "#REF!"} {
		if got := stored(t, path, summaryPart, cell); got != want {
			t.Errorf("%s stores %q, want %q", cell, got, want)
		}
	}
}

// Array arithmetic over ranges is left to Excel: the calculator gets it
// wrong (an agent found SUMPRODUCT((A="East")*B*C) stored as nonsense).
func TestArrayArithmeticNotStored(t *testing.T) {
	path := writeBook(t, [][]string{
		{"Region", "Units", "Price", "", "=SUMPRODUCT((A2:A4=\"East\")*B2:B4*C2:C4)", "=SUMPRODUCT(B2:B4,C2:C4)"},
		{"East", "2", "3"}, {"West", "4", "5"}, {"East", "1", "10"},
	})
	if got := stored(t, path, dataPart, "E1"); got != "" {
		t.Errorf("E1 stores %q from array arithmetic", got)
	}
	if got := stored(t, path, dataPart, "F1"); got != "36" {
		t.Errorf("F1 stores %q, want 36", got)
	}
}

// Renaming a table column by its header renames it in formulas when saved;
// the results computed before that must use the new name too.
func TestResultsAfterColumnRename(t *testing.T) {
	path := writeBook(t, [][]string{
		{"Region", "Units", "Price"},
		{"East", "5", "30"},
		{"West", "2", "40"},
	})
	change(t, path, func(e *Editor) {
		must(t, e.f.AddTable("Data", &excelize.Table{Range: "A1:C3", Name: "Sales"}))
		must(t, e.Set("Data", 5, 1, input.Parse("=SUM(Sales[Units])", false)))
	})
	change(t, path, func(e *Editor) {
		must(t, e.Set("Data", 1, 2, input.Parse("Qty", false)))
	})
	wb, err := Read(path)
	must(t, err)
	s := wb.Sheet("Data")
	if c := s.At(5, 1); c.Formula != "SUM(Sales[Qty])" || c.Pending || c.Num != 7 {
		t.Errorf("A5 = %+v, want =SUM(Sales[Qty]) stored as 7", c)
	}
}

// The aggregates sheets computes itself follow Excel: errors pass COUNT
// by and count for COUNTA; text and logical values count for COUNTA only.
func TestNativeAggregates(t *testing.T) {
	cells := map[[2]int]grid.Cell{
		{1, 1}: {Kind: grid.Number, Num: 4},
		{2, 1}: {Kind: grid.Error, Text: "#N/A"},
		{3, 1}: {Kind: grid.Text, Text: "x"},
		{4, 1}: {Kind: grid.Bool, Bool: true},
		{5, 1}: {Kind: grid.Number, Num: -2},
		{1, 2}: {Kind: grid.Number, Num: 1},
		{2, 2}: {Kind: grid.Text, Text: "7"},
	}
	used := func(string) grid.Range { return grid.Range{R1: 1, C1: 1, R2: 6, C2: 2} }
	value := func(_ *judge, sheet string, r, c int) (grid.Cell, bool) {
		if !strings.EqualFold(sheet, "My Sheet") {
			t.Errorf("read sheet %q", sheet)
		}
		return cells[[2]int{r, c}], true
	}
	native := nativeAggregates(used, value)
	for f, want := range map[string]string{
		"COUNT('My Sheet'!A1:A6)":             "2",
		"COUNTA('My Sheet'!A1:A6)":            "5",
		"SUM('My Sheet'!A1:A6)":               "#N/A",
		"MAX('My Sheet'!A1:A6)":               "#N/A",
		"SUM('My Sheet'!B1:B6)":               "1",
		"AVERAGE('My Sheet'!B3:B6)":           "#DIV/0!",
		"MIN( 'My Sheet'!A3:A5 )":             "-2",
		"MAX('My Sheet'!A3,'My Sheet'!B1:B2)": "1",
	} {
		v, ok := native(nil, &formulaAt{sheet: "Other", text: f})
		got := "not native"
		if ok {
			got = grid.Display(&v, grid.Format{}, false)
		}
		if got != want {
			t.Errorf("%s = %s, want %s", f, got, want)
		}
	}
	if _, _, ok := aggregateOf("SUM(A1:A3)+1"); ok {
		t.Error("SUM(A1:A3)+1 is not an aggregate of references alone")
	}
}

// Running totals are scanned on from the one before, with the same results
// as scanning each range whole, errors and unknown values included.
func TestNativeAggregatesRunning(t *testing.T) {
	vals := map[int]grid.Cell{}
	for r := 1; r <= 40; r++ {
		vals[r] = grid.Cell{Kind: grid.Number, Num: float64(r)}
	}
	vals[7] = grid.Cell{Kind: grid.Text, Text: "x"}
	vals[20] = grid.Cell{Kind: grid.Error, Text: "#N/A"}
	unknown := 30
	used := func(string) grid.Range { return grid.Range{R1: 1, C1: 1, R2: 40, C2: 1} }
	native := nativeAggregates(used, func(_ *judge, _ string, r, _ int) (grid.Cell, bool) { return vals[r], r != unknown })
	for r := 1; r <= 40; r++ {
		for _, fn := range []string{"SUM", "COUNT", "COUNTA", "MAX"} {
			got, ok := native(nil, &formulaAt{sheet: "S", text: fmt.Sprintf("%s(A$1:A%d)", fn, r)})
			var want grid.Cell
			wantOK := r < unknown
			if wantOK {
				sum, n, filled, high := 0.0, 0, 0, 0.0
				var errCell *grid.Cell
				for i := 1; i <= r; i++ {
					switch v := vals[i]; v.Kind {
					case grid.Number:
						sum, n, filled = sum+v.Num, n+1, filled+1
						high = max(high, v.Num)
					case grid.Error:
						filled++
						if errCell == nil {
							errCell = &v
						}
					default:
						filled++
					}
				}
				switch {
				case errCell != nil && fn != "COUNT" && fn != "COUNTA":
					want = *errCell
				case fn == "SUM":
					want = grid.Cell{Kind: grid.Number, Num: sum}
				case fn == "COUNT":
					want = grid.Cell{Kind: grid.Number, Num: float64(n)}
				case fn == "COUNTA":
					want = grid.Cell{Kind: grid.Number, Num: float64(filled)}
				default:
					want = grid.Cell{Kind: grid.Number, Num: high}
				}
			}
			if ok != wantOK || ok && got != want {
				t.Fatalf("%s(A$1:A%d) = %+v %v, want %+v %v", fn, r, got, ok, want, wantOK)
			}
		}
	}
	// A range read before the unknown cell still answers.
	if v, ok := native(nil, &formulaAt{sheet: "S", text: "SUM(A$1:A5)"}); !ok || v.Num != 15 {
		t.Errorf("SUM(A$1:A5) after the rest = %+v %v", v, ok)
	}
}

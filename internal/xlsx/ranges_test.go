package xlsx

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"io"
	"os"
	"regexp"
	"strings"
	"testing"

	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/input"
	"github.com/xuri/excelize/v2"
)

// read opens a saved workbook as sheets reads it.
func read(t *testing.T, path string) *grid.Workbook {
	t.Helper()
	wb, err := Read(path)
	must(t, err)
	return wb
}

// show is a cell as sheets shows it, with its formula when it has one.
func cellShow(t *testing.T, wb *grid.Workbook, sheet, cell string) string {
	t.Helper()
	s := wb.Sheet(sheet)
	if s == nil {
		t.Fatalf("no sheet %s", sheet)
	}
	r, c, _ := grid.ParseCell(cell)
	x := s.At(r, c)
	if x == nil || x.IsEmpty() {
		return ""
	}
	v := wb.Value(s, r, c)
	t2 := grid.Display(v, wb.FormatOf(v), wb.Date1904)
	if v.Pending {
		t2 = "?"
	}
	if x.Formula != "" {
		return "=" + x.Formula + " → " + t2
	}
	return t2
}

func expect(t *testing.T, wb *grid.Workbook, sheet string, want map[string]string) {
	t.Helper()
	for cell, w := range want {
		if got := cellShow(t, wb, sheet, cell); got != w {
			t.Errorf("%s!%s = %q, want %q", sheet, cell, got, w)
		}
	}
}

// validPackage checks every XML part parses and no zip entry is misnamed:
// what Excel would refuse to open, or repair.
func validPackage(t *testing.T, path string) {
	t.Helper()
	data, err := os.ReadFile(path)
	must(t, err)
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	must(t, err)
	names := map[string]bool{}
	for _, f := range zr.File {
		names[f.Name] = true
		if strings.HasPrefix(f.Name, "/") {
			t.Errorf("zip entry %q starts with /", f.Name)
		}
		if !strings.HasSuffix(f.Name, ".xml") && !strings.HasSuffix(f.Name, ".vml") && !strings.HasSuffix(f.Name, ".rels") {
			continue
		}
		r, _ := f.Open()
		b, _ := io.ReadAll(r)
		r.Close()
		dec := xml.NewDecoder(bytes.NewReader(b))
		for {
			_, err := dec.Token()
			if err == io.EOF {
				break
			}
			if err != nil {
				t.Errorf("%s: %v", f.Name, err)
				break
			}
		}
	}
	ct := regexp.MustCompile(`PartName="/([^"]*)"`)
	for _, f := range zr.File {
		if f.Name != "[Content_Types].xml" {
			continue
		}
		r, _ := f.Open()
		b, _ := io.ReadAll(r)
		r.Close()
		for _, m := range ct.FindAllSubmatch(b, -1) {
			if !names[string(m[1])] {
				t.Errorf("content type for missing part %s", m[1])
			}
		}
	}
}

func TestCopyRange(t *testing.T) {
	path := writeBook(t, [][]string{
		{"Item", "Qty", "Price", "Total", "Day"},
		{"Pen", "2", "1.5", "=B{r}*C{r}", "2026-03-01"},
		{"Ink", "3", "4", "=B{r}*C{r}", "2026-03-02"},
	})
	change(t, path, func(e *Editor) {
		_, err := e.CopyRange("Data", rng("D3"), "Data", rng("D4:D6"), CopyOptions{})
		must(t, err)
	})
	wb := read(t, path)
	expect(t, wb, "Data", map[string]string{"D4": "=B4*C4 → 0", "D6": "=B6*C6 → 0"})

	// To another sheet: formulas shift, and keep reading their own sheet.
	change(t, path, func(e *Editor) {
		_, _, err := e.EnsureSheet("Copy")
		must(t, err)
		_, err = e.CopyRange("Data", rng("A1:E3"), "Copy", rng("B2"), CopyOptions{})
		must(t, err)
	})
	wb = read(t, path)
	expect(t, wb, "Copy", map[string]string{"B2": "Item", "E3": "=C3*D3 → 3", "F4": "2026-03-02"})

	// Values: results, with their number formats.
	change(t, path, func(e *Editor) {
		_, err := e.CopyRange("Data", rng("D2:E3"), "Data", rng("H2"), CopyOptions{Values: true})
		must(t, err)
	})
	wb = read(t, path)
	expect(t, wb, "Data", map[string]string{"H2": "3", "H3": "12", "I2": "2026-03-01"})

	// A paste that does not fill its target evenly is refused.
	e, err := OpenEditor(path, false)
	must(t, err)
	defer e.Close()
	if _, err := e.CopyRange("Data", rng("A1:B2"), "Data", rng("H10:J12"), CopyOptions{}); err == nil || !strings.Contains(err.Error(), "does not fill evenly") {
		t.Errorf("uneven paste: %v", err)
	}
}

func TestMoveRange(t *testing.T) {
	path := writeBook(t, [][]string{
		{"Item", "Qty", "Price", "Total"},
		{"Pen", "2", "1.5", "=B{r}*C{r}"},
		{"Ink", "3", "4", "=B{r}*C{r}"},
		{"Sum", "=SUM(B2:B3)", "", "=SUM(D2:D3)"},
	})
	change(t, path, func(e *Editor) {
		_, _, err := e.EnsureSheet("Calc")
		must(t, err)
		for cell, f := range map[string]string{"A1": "=Data!D4*2", "A2": "=SUM(Data!B2:B3)", "A3": "=Data!H11", "A4": "=SUM(Data!A1:A9)"} {
			r, c, _ := grid.ParseCell(cell)
			must(t, e.Set("Calc", r, c, input.Parse(f, false)))
		}
		must(t, e.f.SetDefinedName(&excelize.DefinedName{Name: "Total", RefersTo: "Data!$D$4"}))
	})
	var res RangeResult
	change(t, path, func(e *Editor) {
		var err error
		res, err = e.MoveRange("Data", rng("A1:D4"), "Data", rng("F8"))
		must(t, err)
	})
	wb := read(t, path)
	expect(t, wb, "Data", map[string]string{"A1": "", "F8": "Item", "I9": "=G9*H9 → 3", "G11": "=SUM(G9:G10) → 5"})
	expect(t, wb, "Calc", map[string]string{
		"A1": "=Data!I11*2 → 30", "A2": "=SUM(Data!G9:G10) → 5",
		"A3": "=Data!#REF! → #REF!",  // pasted over
		"A4": "=SUM(Data!A1:A9) → 0", // took cells out of the range: stays
	})
	if len(wb.Names) != 1 || wb.Names[0].Ref != "Data!$I$11" {
		t.Errorf("names = %+v", wb.Names)
	}
	if !strings.Contains(strings.Join(res.Notes, " "), "Calc!A4") {
		t.Errorf("notes = %q", res.Notes)
	}
	validPackage(t, path)

	// Across sheets: formulas moved keep reading their old sheet.
	change(t, path, func(e *Editor) {
		_, _, err := e.EnsureSheet("Archive")
		must(t, err)
		_, err = e.MoveRange("Data", rng("F8:I11"), "Archive", rng("A1"))
		must(t, err)
	})
	wb = read(t, path)
	expect(t, wb, "Archive", map[string]string{"D2": "=B2*C2 → 3", "B4": "=SUM(B2:B3) → 5"})
	expect(t, wb, "Calc", map[string]string{"A1": "=Archive!D4*2 → 30"})
}

func TestMoveFormats(t *testing.T) {
	path := writeBook(t, [][]string{
		{"Name", "Score", "Status"},
		{"Ada", "92", "Late"},
		{"Bob", "47", "On time"},
		{"Cy", "75", "Late"},
	})
	change(t, path, func(e *Editor) {
		_, err := e.Cond("Data", rng("A2:C4"), []string{`=$C2="Late"`, "fill=lightyellow"})
		must(t, err)
		_, err = e.Validate("Data", rng("B2:B4"), ValidationRule{Words: []string{"whole", "between", "0", "and", "100"}})
		must(t, err)
		_, err = e.Style("Data", rng("A1:C1"), look(t, "bold fill=lightblue"))
		must(t, err)
		_, err = e.Style("Data", rng("A1:C4"), look(t, "border=outline:thin"))
		must(t, err)
		must(t, e.SetNote("Data", 2, 1, "First", "Ann"))
		must(t, e.SetLink("Data", 3, 1, &grid.Link{URL: "https://example.com/bob"}, ""))
	})
	change(t, path, func(e *Editor) {
		_, err := e.MoveRange("Data", rng("A1:C4"), "Data", rng("E11"))
		must(t, err)
	})
	wb := read(t, path)
	s := wb.Sheet("Data")
	if len(s.Notes) != 1 || s.Notes[0].Row != 12 || s.Notes[0].Col != 5 || s.Notes[0].Text != "First" {
		t.Errorf("notes = %+v", s.Notes)
	}
	if len(s.Links) != 1 || s.Links[0].Row != 13 || s.Links[0].Col != 5 {
		t.Errorf("links = %+v", s.Links)
	}
	e, err := OpenEditor(path, false)
	must(t, err)
	cfs, err := e.CondFormats("Data")
	must(t, err)
	if len(cfs) != 1 || cfs[0].Range != "E12:G14" || !strings.Contains(cfs[0].Rule, `$G12="Late"`) || !cfs[0].Recorded {
		t.Errorf("conditional formats = %+v", cfs)
	}
	dvs, err := e.Validations("Data")
	must(t, err)
	if len(dvs) != 1 || dvs[0].Range != "F12:F14" {
		t.Errorf("validations = %+v", dvs)
	}
	var areas []string
	for _, rec := range e.Records() {
		areas = append(areas, rec.Kind+" "+rec.Ref)
	}
	e.Close()
	want := []string{"cond Data!E12:G14", "style Data!E11:G11", "style Data!E11:G14"}
	if strings.Join(areas, "; ") != strings.Join(want, "; ") {
		t.Errorf("records = %q, want %q", areas, want)
	}
	// fmt keeps all of it where it went.
	format(t, path)
	if h := cellLook(t, path, "Data", "E11"); styleFill(h) != "DDEBF7" {
		t.Errorf("moved header fill after fmt = %s", styleFill(h))
	}
	validPackage(t, path)
}

func TestMoveSplitsRecords(t *testing.T) {
	path := budget(t)
	change(t, path, func(e *Editor) {
		_, err := e.Style("Data", rng("A1:D1"), look(t, "fill=lightblue"))
		must(t, err)
	})
	change(t, path, func(e *Editor) {
		_, _, err := e.EnsureSheet("Other")
		must(t, err)
		_, err = e.MoveRange("Data", rng("C1:D2"), "Other", rng("A1"))
		must(t, err)
	})
	e, err := OpenEditor(path, false)
	must(t, err)
	defer e.Close()
	var got []string
	for _, rec := range e.Records() {
		got = append(got, rec.Ref+" "+rec.Text())
	}
	want := []string{"Data!A1:B1 style fill=lightblue", "Other!A1:B1 style fill=lightblue"}
	if strings.Join(got, "; ") != strings.Join(want, "; ") {
		t.Errorf("records = %q, want %q", got, want)
	}
}

func TestMoveRefuses(t *testing.T) {
	path := fixture(t, "tables.xlsx")
	e, err := OpenEditor(path, false)
	must(t, err)
	defer e.Close()
	s := e.wb.Sheets[0]
	tb := s.Tables[0].Range
	if _, err := e.MoveRange(s.Name, tb, s.Name, grid.Range{R1: tb.R1 + 20, C1: tb.C1, R2: tb.R1 + 20, C2: tb.C1}); err == nil || !strings.Contains(err.Error(), "Excel table") {
		t.Errorf("moving a table: %v", err)
	}
}

func TestSortRange(t *testing.T) {
	path := writeBook(t, [][]string{
		{"Key", "Value", "Running", "Note"},
		{"b", "3", "=SUM($B$2:B{r})", "x"},
		{"B", "1", "=SUM($B$2:B{r})", ""},
		{"", "2", "=SUM($B$2:B{r})", "blank key"},
		{"10", "5", "=SUM($B$2:B{r})", ""},
		{"TRUE", "4", "=SUM($B$2:B{r})", ""},
		{"a", "6", "=SUM($B$2:B{r})", ""},
	})
	change(t, path, func(e *Editor) {
		must(t, e.SetNote("Data", 2, 1, "first b", "Ann"))
		res, err := e.SortRange("Data", rng("A2:D7"), []SortKey{{Col: 1}})
		must(t, err)
		if res.Moved == 0 {
			t.Error("nothing moved")
		}
	})
	wb := read(t, path)
	// Numbers, then text ignoring case (b and B tie and keep their order),
	// then logical values; blanks last.
	var keys []string
	for r := 2; r <= 7; r++ {
		keys = append(keys, cellShow(t, wb, "Data", grid.CellName(r, 1)))
	}
	if got := strings.Join(keys, ","); got != "10,a,b,B,TRUE," {
		t.Errorf("order = %s", got)
	}
	expect(t, wb, "Data", map[string]string{"C2": "=SUM($B$2:B2) → 5", "C7": "=SUM($B$2:B7) → 21", "D4": "x", "D7": "blank key"})
	if n := wb.Sheet("Data").Notes; len(n) != 1 || n[0].Row != 4 || n[0].Text != "first b" {
		t.Errorf("notes = %+v", n)
	}

	// Descending: reversed, blanks still last.
	change(t, path, func(e *Editor) {
		_, err := e.SortRange("Data", rng("A2:D7"), []SortKey{{Col: 2, Desc: true}})
		must(t, err)
	})
	wb = read(t, path)
	var vals []string
	for r := 2; r <= 7; r++ {
		vals = append(vals, cellShow(t, wb, "Data", grid.CellName(r, 2)))
	}
	if got := strings.Join(vals, ","); got != "6,5,4,3,2,1" {
		t.Errorf("descending = %s", got)
	}
	validPackage(t, path)
}

func TestSortRefuses(t *testing.T) {
	path := writeBook(t, [][]string{
		{"Name", "Value", "Change"},
		{"a", "3", ""},
		{"b", "1", "=B3-B2"},
		{"c", "2", ""},
	})
	e, err := OpenEditor(path, false)
	must(t, err)
	_, err = e.SortRange("Data", rng("A2:C4"), []SortKey{{Col: 2}})
	if err == nil || !strings.Contains(err.Error(), "another row") {
		t.Errorf("a formula reading another row: %v", err)
	}
	e.Close()

	// A formula in its own row, reading a column left out of the sort.
	path = writeBook(t, [][]string{
		{"Name", "Value", "Twice"},
		{"a", "3", "=B{r}*2"},
		{"b", "1", "=B{r}*2"},
	})
	e, err = OpenEditor(path, false)
	must(t, err)
	_, err = e.SortRange("Data", rng("A2:A3"), []SortKey{{Col: 1, Desc: true}})
	if err != nil {
		t.Errorf("sorting a column no formula in it reads: %v", err)
	}
	e.Close()
	e, err = OpenEditor(path, false)
	must(t, err)
	_, err = e.SortRange("Data", rng("A2:C3"), []SortKey{{Col: 2}})
	if err != nil {
		t.Errorf("sorting whole rows: %v", err)
	}
	e.Close()
	e, err = OpenEditor(path, false)
	must(t, err)
	_, err = e.SortRange("Data", rng("C2:C3"), []SortKey{{Col: 3}})
	if err == nil || !strings.Contains(err.Error(), "outside the sorted columns") {
		t.Errorf("a formula reading a column left out: %v", err)
	}
	e.Close()

	// Hidden rows.
	change(t, path, func(e *Editor) {
		must(t, e.View("Data", ViewOptions{Hide: []grid.Range{{R1: 3, C1: 1, R2: 3, C2: grid.MaxCols}}}))
	})
	e, err = OpenEditor(path, false)
	must(t, err)
	defer e.Close()
	if _, err := e.SortRange("Data", rng("A2:C3"), []SortKey{{Col: 2}}); err == nil || !strings.Contains(err.Error(), "hidden") {
		t.Errorf("hidden rows: %v", err)
	}
}

func TestClearRange(t *testing.T) {
	path := writeBook(t, [][]string{
		{"Name", "Value"},
		{"a", "1"},
		{"b", "=B2*2"},
	})
	change(t, path, func(e *Editor) {
		_, err := e.Style("Data", rng("A1:B1"), look(t, "bold fill=lightblue"))
		must(t, err)
		must(t, e.SetNote("Data", 2, 1, "keep me", "sheets"))
	})
	change(t, path, func(e *Editor) {
		res, err := e.ClearRange("Data", rng("A2:B3"), ClearWhat{Contents: true})
		must(t, err)
		if res.Cells != 4 {
			t.Errorf("cleared %d cells", res.Cells)
		}
	})
	wb := read(t, path)
	expect(t, wb, "Data", map[string]string{"A2": "", "B3": "", "A1": "Name"})
	if len(wb.Sheet("Data").Notes) != 1 {
		t.Error("clearing contents took the note")
	}
	// Formats: the recorded styling goes too, so fmt does not bring it back.
	change(t, path, func(e *Editor) {
		_, err := e.ClearRange("Data", rng("A1"), ClearWhat{Formats: true})
		must(t, err)
		recs := e.Records()
		if len(recs) != 1 || recs[0].Ref != "Data!B1" {
			t.Errorf("records = %+v", recs)
		}
	})
	format(t, path)
	if s := cellLook(t, path, "Data", "B1"); styleFill(s) != "DDEBF7" {
		t.Errorf("B1 fill = %s", styleFill(s))
	}
	if s := cellLook(t, path, "Data", "A1"); styleFill(s) == "DDEBF7" {
		t.Error("A1 kept its recorded fill")
	}
	change(t, path, func(e *Editor) {
		_, err := e.ClearRange("Data", rng("A1:B3"), ClearWhat{Notes: true})
		must(t, err)
	})
	if n := read(t, path).Sheet("Data").Notes; len(n) != 0 {
		t.Errorf("notes left: %+v", n)
	}
	validPackage(t, path)
}

func TestReplace(t *testing.T) {
	path := writeBook(t, [][]string{
		{"Name", "Amount", "Twice"},
		{"Acme Inc.", "N/A", "=B{r}*2"},
		{"acme inc.", "1200", "=B{r}*2"},
	})
	change(t, path, func(e *Editor) {
		re := regexp.MustCompile(`(?i)acme inc\.`)
		got, err := e.Replace("", nil, func(s string) (string, bool) {
			return re.ReplaceAllString(s, "Acme Corp"), re.MatchString(s)
		}, false)
		must(t, err)
		if len(got) != 2 {
			t.Errorf("replaced %+v", got)
		}
	})
	change(t, path, func(e *Editor) {
		_, err := e.Replace("Data", nil, func(s string) (string, bool) { return "7", s == "N/A" }, false)
		must(t, err)
	})
	wb := read(t, path)
	expect(t, wb, "Data", map[string]string{"A2": "Acme Corp", "A3": "Acme Corp", "B2": "7", "C2": "=B2*2 → 14"})
	if c := wb.Sheet("Data").At(2, 2); c.Kind != grid.Number {
		t.Errorf("B2 is %s, want a number", c.Kind)
	}
}

func TestNotesMoveWithRows(t *testing.T) {
	// Made by openpyxl: its notes parts are named otherwise than excelize
	// expects, and must still move cleanly.
	path := fixture(t, "rich.xlsx")
	before := read(t, path).Sheet("Data").Notes
	if len(before) != 1 || before[0].Row != 2 || before[0].Col != 2 {
		t.Fatalf("fixture notes = %+v", before)
	}
	change(t, path, func(e *Editor) { must(t, e.InsertRows("Data", 1, 2)) })
	after := read(t, path).Sheet("Data").Notes
	if len(after) != 1 || after[0].Row != 4 || after[0].Col != 2 || after[0].Text != before[0].Text {
		t.Errorf("after inserting rows: %+v", after)
	}
	change(t, path, func(e *Editor) { must(t, e.SetNote("Data", 1, 1, "A new note", "sheets")) })
	notes := read(t, path).Sheet("Data").Notes
	if len(notes) != 2 {
		t.Errorf("notes = %+v", notes)
	}
	validPackage(t, path)
	// One note shape per note in the drawing.
	data, err := os.ReadFile(path)
	must(t, err)
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	must(t, err)
	shapes := 0
	for _, f := range zr.File {
		if strings.HasSuffix(f.Name, ".vml") {
			r, _ := f.Open()
			b, _ := io.ReadAll(r)
			r.Close()
			shapes += strings.Count(string(b), `ObjectType="Note"`)
		}
	}
	if shapes != 2 {
		t.Errorf("%d note shapes, want 2", shapes)
	}
}

func TestErrorValuesCopied(t *testing.T) {
	path := writeBook(t, [][]string{{"Code"}, {"x"}})
	rewritePart(t, path, "xl/worksheets/sheet1.xml", func(s string) string {
		return regexp.MustCompile(`<c r="A2"[^>]*>.*?</c>`).ReplaceAllString(s, `<c r="A2" t="e"><v>#N/A</v></c>`)
	})
	change(t, path, func(e *Editor) {
		_, err := e.CopyRange("Data", rng("A1:A2"), "Data", rng("C1"), CopyOptions{})
		must(t, err)
	})
	wb := read(t, path)
	if c := wb.Sheet("Data").At(2, 3); c == nil || c.Kind != grid.Error || c.Text != "#N/A" {
		t.Errorf("C2 = %+v, want the error #N/A", c)
	}
}

func TestThreads(t *testing.T) {
	path := fixture(t, "threads.xlsx")
	notes := read(t, path).Sheet("Data").Notes
	if len(notes) != 2 || !notes[1].Thread || notes[1].Author != "Ann" || notes[1].Text != "Is 47 right?\nBob: Yes, checked." {
		t.Fatalf("notes = %+v", notes)
	}
	// sheets cannot write threads.
	e, err := OpenEditor(path, false)
	must(t, err)
	if err := e.SetNote("Data", 3, 2, "x", "sheets"); err == nil || !strings.Contains(err.Error(), "comment thread") {
		t.Errorf("writing over a thread: %v", err)
	}
	e.Close()
	// A thread follows its row, whether rows are inserted or sorted.
	change(t, path, func(e *Editor) { must(t, e.InsertRows("Data", 2, 1)) })
	notes = read(t, path).Sheet("Data").Notes
	if len(notes) != 2 || notes[1].Row != 4 || !notes[1].Thread {
		t.Errorf("after inserting a row: %+v", notes)
	}
	change(t, path, func(e *Editor) {
		_, err := e.SortRange("Data", rng("A3:B5"), []SortKey{{Col: 2}})
		must(t, err)
	})
	wb := read(t, path)
	notes = wb.Sheet("Data").Notes
	// Bob (47) sorts first.
	if len(notes) != 2 || notes[1].Row != 3 || notes[1].Col != 2 || !notes[1].Thread || cellShow(t, wb, "Data", "A3") != "Bob" {
		t.Errorf("after sorting: %+v", notes)
	}
	validPackage(t, path)
	// Deleting the row deletes the thread.
	change(t, path, func(e *Editor) { must(t, e.DeleteRows("Data", 3, 3)) })
	if notes = read(t, path).Sheet("Data").Notes; len(notes) != 1 || notes[0].Thread {
		t.Errorf("after deleting its row: %+v", notes)
	}
	validPackage(t, path)
}

// A recorded rule's formula follows edits as the rule does, so a copy made
// from the record reads the right cells.
func TestCondRecordFollows(t *testing.T) {
	path := writeBook(t, [][]string{
		{"Name", "Score", "Status"},
		{"Ada", "92", "Late"},
		{"Bob", "47", "On time"},
	})
	change(t, path, func(e *Editor) {
		_, err := e.Cond("Data", rng("A2:C3"), []string{`=$C2="Late"`, "fill=lightyellow"})
		must(t, err)
	})
	change(t, path, func(e *Editor) { must(t, e.InsertRows("Data", 1, 2)) })
	change(t, path, func(e *Editor) {
		_, err := e.MoveRange("Data", rng("A3:C5"), "Data", rng("E10"))
		must(t, err)
	})
	change(t, path, func(e *Editor) { must(t, e.RenameSheet("Data", "Scores")) })
	change(t, path, func(e *Editor) {
		_, _, err := e.EnsureSheet("Copy")
		must(t, err)
		_, err = e.CopyRange("Scores", rng("E10:G12"), "Copy", rng("A1"), CopyOptions{})
		must(t, err)
	})
	e, err := OpenEditor(path, false)
	must(t, err)
	defer e.Close()
	var got []string
	for _, rec := range e.Records() {
		got = append(got, rec.Ref+" "+rec.Text())
	}
	want := []string{`Scores!E11:G12 cond '=$G11="Late"' fill=lightyellow`, `Copy!A2:C3 cond '=$G2="Late"' fill=lightyellow`}
	if strings.Join(got, "\n") != strings.Join(want, "\n") {
		t.Errorf("records:\n%s\nwant:\n%s", strings.Join(got, "\n"), strings.Join(want, "\n"))
	}
	for _, sheet := range []string{"Scores", "Copy"} {
		cfs, err := e.CondFormats(sheet)
		must(t, err)
		for _, cf := range cfs {
			if strings.Contains(cf.Rule, "#REF!") {
				t.Errorf("%s: %+v", sheet, cf)
			}
		}
	}
}

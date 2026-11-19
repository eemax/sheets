package cli

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"reflect"
	"regexp"
	"strings"
	"testing"
	"time"

	"github.com/eemax/sheets/internal/diff"
	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/xlsx"
)

func TestLostFeatures(t *testing.T) {
	f := func(kind, sheet, where, what string) xlsx.Feature {
		return xlsx.Feature{Kind: kind, Sheet: sheet, Where: where, What: what}
	}
	before := []xlsx.Feature{
		f(xlsx.FeatureValidation, "Data", "B2:B6", "whole validation"),
		f(xlsx.FeatureValidation, "Data", "C2:C6", "list validation"),
		f(xlsx.FeatureChart, "Data", "F2", "column chart"),
		f(xlsx.FeatureCond, "Old", "A1:A3", "formula rule"),
		f(xlsx.FeatureName, "", "Rate", "Data!$B$1"),
		f(xlsx.FeatureName, "", "Gone", "Data!$B$2"),
		f(xlsx.FeatureTable, "Data", "A1:D6", "Sales"),
		f(xlsx.FeatureShape, "Removed", "A1", "TextBox 1"),
	}
	after := []xlsx.Feature{
		f(xlsx.FeatureValidation, "Data", "B2:B7", "whole validation"), // grew with a row
		f(xlsx.FeatureChart, "Data", "G2", "column chart"),             // moved
		f(xlsx.FeatureCond, "New", "A1:A3", "formula rule"),            // its sheet renamed
		f(xlsx.FeatureName, "", "rate", "Data!$B$1"),
		f(xlsx.FeatureName, "", "Other", "Data!$B$2"), // a name is matched by name
		f(xlsx.FeatureTable, "Data", "A1:D7", "Sales"),
	}
	got := lostFeatures(before, after, map[string]string{"data": "Data", "old": "New", "removed": ""})
	want := []lostFeature{
		{Kind: xlsx.FeatureValidation, Where: "Data!C2:C6", What: "list validation"},
		{Kind: xlsx.FeatureName, Where: "Gone", What: "Data!$B$2"},
		{Kind: xlsx.FeatureShape, Where: "Removed!A1", What: "TextBox 1"},
	}
	if !reflect.DeepEqual(got, want) {
		t.Errorf("lost\n got %+v\nwant %+v", got, want)
	}
	if got := lostFeatures(nil, after, nil); got == nil || len(got) != 0 {
		t.Errorf("nothing before: %+v", got)
	}
	// Both validations grew with a row inserted, then the whole one was
	// cleared: the one lost is the whole one, not the decimal one, which is
	// now in another place.
	got = lostFeatures([]xlsx.Feature{f(xlsx.FeatureValidation, "Data", "B2:B6", "whole validation"), f(xlsx.FeatureValidation, "Data", "C2:C6", "decimal validation")},
		[]xlsx.Feature{f(xlsx.FeatureValidation, "Data", "C2:C7", "decimal validation")}, map[string]string{"data": "Data"})
	if want := []lostFeature{{Kind: xlsx.FeatureValidation, Where: "Data!B2:B6", What: "whole validation"}}; !reflect.DeepEqual(got, want) {
		t.Errorf("lost after a move\n got %+v\nwant %+v", got, want)
	}
	// A validation that changed its type in place is not lost.
	if got := lostFeatures([]xlsx.Feature{f(xlsx.FeatureValidation, "Data", "B2:B6", "whole validation")},
		[]xlsx.Feature{f(xlsx.FeatureValidation, "Data", "B2:B7", "decimal validation")}, nil); len(got) != 0 {
		t.Errorf("changed: %+v", got)
	}
}

// TestStaleWhy checks what a stale formula is said to miss: lines added or
// removed, a sheet renamed, or both.
func TestStaleWhy(t *testing.T) {
	renamed := &report{renames: [][2]string{{"data", "Facts"}}}
	renamed.setEdits(nil)
	if got := renamed.staleWhy("SUM(Data!B2:B6)", "Summary"); got != "a sheet it reads was renamed" {
		t.Errorf("renamed: %q", got)
	}
	lines := &report{}
	lines.setEdits([]lineEdit{{sheet: "Data", rowsAdded: [][2]int{{3, 3}}}})
	if got := lines.staleWhy("SUM(Data!B2:B6)", "Summary"); got != "rows or columns added or removed moved what it reads" {
		t.Errorf("lines: %q", got)
	}
	both := &report{renames: [][2]string{{"data", "Facts"}}}
	both.setEdits([]lineEdit{{sheet: "Facts", rowsAdded: [][2]int{{3, 3}}}})
	if got := both.staleWhy("SUM(Data!B2:B6)", "Summary"); got != "rows or columns added or removed moved what it reads, and a sheet it reads was renamed" {
		t.Errorf("both: %q", got)
	}
}

// TestRewritesRemoved checks what a formula reading a sheet removed may have
// become: sheets writes =SUM(#REF!), Excel =SUM(#REF!B2:B6).
func TestRewritesRemoved(t *testing.T) {
	r := &report{removedSheets: []string{"My Data"}}
	r.setEdits(nil)
	for _, now := range []string{"SUM(#REF!)+B1", "SUM(#REF!B2:B6)+B1"} {
		if !r.adjusted(folded("SUM('My Data'!B2:B6)+B1"), "Summary", folded(now), 2, 2) {
			t.Errorf("=%s is not what it may have become", now)
		}
	}
	if r.adjusted(folded("SUM('My Data'!B2:B6)+B1"), "Summary", folded("SUM(#REF!)+B2"), 2, 2) {
		t.Errorf("a formula changed besides is what it may have become")
	}
}

// A share of the total in every row and every tenth row deleted: hundreds
// of runs of rows removed, and every formula reads past all of them. What
// Excel writes for them is no formula changed, and the shares moved with
// the rows removed (the total), not with a formula blamed.
func TestCheckAgainstManyRunsShare(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("SHEETS_CACHE_DIR", filepath.Join(dir, ".cache"))
	build := func(path string, drop bool) {
		var keep []int
		for i := 2; i <= 5001; i++ {
			if !drop || i%10 != 0 {
				keep = append(keep, i)
			}
		}
		var b strings.Builder
		b.WriteString("Region\tUnits\tShare\n")
		for j, i := range keep {
			fmt.Fprintf(&b, "R%d\t%d\t=B%d/$B$%d\n", i, i%50+1, j+2, len(keep)+2)
		}
		fmt.Fprintf(&b, "Total\t=SUM(B2:B%d)\n", len(keep)+1)
		if out, code := runCLI(b.String(), "write", path, "Data!A1"); code != 0 {
			t.Fatalf("write: %s", out)
		}
	}
	before, after := filepath.Join(dir, "a.xlsx"), filepath.Join(dir, "b.xlsx")
	build(before, false)
	build(after, true)
	js, code := runCLI("", "--json", "check", after, "--against", before)
	if code != 0 {
		t.Fatalf("exit %d:\n%.2000s", code, js)
	}
	var r againstResult
	if err := json.Unmarshal([]byte(js), &r); err != nil {
		t.Fatal(err)
	}
	if r.Summary.Formulas != 0 || len(r.Unexplained) != 0 {
		t.Fatalf("formulas changed %d, unexplained %d: %+v", r.Summary.Formulas, len(r.Unexplained), r.Changed[:min(3, len(r.Changed))])
	}
	for _, m := range r.Moved {
		if !strings.HasPrefix(m.Cause, "rows removed: ") {
			t.Fatalf("moved: %+v", m)
		}
	}
}

// Long text cells, changed: the text output shows each cut, as sheets diff
// does, so it stays short; --json gives them whole.
func TestCheckAgainstWide(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("SHEETS_CACHE_DIR", filepath.Join(dir, ".cache"))
	build := func(path, ch string) {
		var b strings.Builder
		for r := 0; r < 12; r++ {
			fmt.Fprintf(&b, "%s%d\t=LEN(A%d)\n", strings.Repeat(ch, 20000), r, r+1)
		}
		if out, code := runCLI(b.String(), "write", path, "S!A1"); code != 0 {
			t.Fatalf("write: %s", out)
		}
	}
	before, after := filepath.Join(dir, "a.xlsx"), filepath.Join(dir, "b.xlsx")
	build(before, "a")
	build(after, "b")
	out, _ := runCLI("", "check", after, "--against", before)
	if len(out) > 8000 {
		t.Errorf("%d bytes of text:\n%.3000s", len(out), out)
	}
	if !strings.Contains(out, "…") {
		t.Errorf("nothing cut:\n%.3000s", out)
	}
	js, _ := runCLI("", "--json", "check", after, "--against", before)
	if !strings.Contains(js, strings.Repeat("b", 20000)) {
		t.Errorf("--json does not give the values whole")
	}
}

// A chart on a chart sheet: lost with its sheet, but not for a rename of the
// chart sheet, which sheets does not follow as it does a worksheet's.
func TestLostChartsheet(t *testing.T) {
	chart := func(sheet string) xlsx.Feature {
		return xlsx.Feature{Kind: xlsx.FeatureChart, Sheet: sheet, What: "bar chart", Chartsheet: true}
	}
	before := []xlsx.Feature{chart("Chart1")}
	got := lostFeatures(before, nil, map[string]string{"data": "Data"})
	want := []lostFeature{{Kind: xlsx.FeatureChart, Where: "Chart1", What: "bar chart on a chart sheet"}}
	if !reflect.DeepEqual(got, want) {
		t.Errorf("lost\n got %+v\nwant %+v", got, want)
	}
	if got := lostFeatures(before, before, nil); len(got) != 0 {
		t.Errorf("kept: %+v", got)
	}
	if got := lostFeatures(before, []xlsx.Feature{chart("Chart 1")}, nil); len(got) != 0 {
		t.Errorf("renamed: %+v", got)
	}
	// Two chart sheets, one lost: the one left keeps its own.
	two := []xlsx.Feature{chart("A"), chart("B")}
	if got := lostFeatures(two, []xlsx.Feature{chart("B")}, nil); len(got) != 1 || got[0].Where != "A" {
		t.Errorf("one of two: %+v", got)
	}
	// A chart on a worksheet is not matched by a chart sheet's.
	onSheet := xlsx.Feature{Kind: xlsx.FeatureChart, Sheet: "Data", Where: "F2", What: "bar chart"}
	if got := lostFeatures([]xlsx.Feature{onSheet}, before, map[string]string{"data": "Data"}); len(got) != 1 {
		t.Errorf("a chart on a sheet and one on a chart sheet: %+v", got)
	}
}

// TestSameValuesDates checks that dates compare as dates, whichever date
// system each workbook counts from (1904 is 1,462 days behind 1900), and
// that plain numbers do not move with it.
func TestSameValuesDates(t *testing.T) {
	mk := func(date1904 bool, serial float64, style uint16) valueOf {
		wb := &grid.Workbook{Date1904: date1904, Formats: []grid.Format{{}, {Kind: grid.Date}}}
		s := grid.NewSheet("S")
		s.Set(1, 1, grid.Cell{Kind: grid.Number, Num: serial, Fmt: style})
		wb.Sheets = append(wb.Sheets, s)
		return value(wb, s, 1, 1)
	}
	for _, c := range []struct {
		name string
		a, b valueOf
		same bool
	}{
		{"the same date", mk(true, 43860, 1), mk(false, 45322, 1), true},
		{"the same date, the other way", mk(false, 45322, 1), mk(true, 43860, 1), true},
		{"a serial kept, another date", mk(true, 43860, 1), mk(false, 43860, 1), false},
		{"one system", mk(false, 45322, 1), mk(false, 45322, 1), true},
		{"numbers", mk(true, 43860, 0), mk(false, 43860, 0), true},
		{"numbers, 1,462 apart", mk(true, 43860, 0), mk(false, 45322, 0), false},
	} {
		if got := sameInputs(c.a, c.b); got != c.same {
			t.Errorf("%s (inputs): %v, want %v", c.name, got, c.same)
		}
		if got := sameResults(c.a, c.b); got != c.same {
			t.Errorf("%s (results): %v, want %v", c.name, got, c.same)
		}
	}
}

// Typed values compare exactly; computed results to Excel's 15 digits.
func TestSameValuesExact(t *testing.T) {
	num := func(n float64) valueOf {
		return valueOf{cell: grid.Cell{Kind: grid.Number, Num: n}, num: n, known: true}
	}
	tenth, fifth := 0.1, 0.2
	for _, c := range []struct {
		name          string
		a, b          float64
		inputs, calcs bool
	}{
		{"a barcode", 7350012345678, 7350012349999, false, false},
		{"one cent on 12,345,678.90", 12345678.90, 12345678.91, false, false},
		{"an order number", 4500012345, 4500012346, false, false},
		{"the same", 7350012345678, 7350012345678, true, true},
		{"past the 15th digit", 0.1, 0.1000000000000001, false, true},
		{"rounding of a sum", tenth + fifth, 0.3, false, true},
		{"the 13th digit", 1, 1.0000000000001, false, false},
		{"what cancellation leaves", 0, 5.5e-17, false, true},
		{"small, and not zero", 1e-9, 2e-9, false, false},
	} {
		if got := sameInputs(num(c.a), num(c.b)); got != c.inputs {
			t.Errorf("%s (inputs): %v, want %v", c.name, got, c.inputs)
		}
		if got := sameResults(num(c.a), num(c.b)); got != c.calcs {
			t.Errorf("%s (results): %v, want %v", c.name, got, c.calcs)
		}
	}
	a, b := num(0.1), num(0.1000000000000001)
	a.text, b.text = "0.1", "0.1"
	if x, y := inputTexts(a, b); x == y {
		t.Errorf("two inputs that differ read alike: %q, %q", x, y)
	}
}

func TestBrokenRow(t *testing.T) {
	s := grid.NewSheet("S")
	for r := 2; r <= 5; r++ {
		s.Set(r, 3, grid.Cell{Formula: fmt.Sprintf("A%d*B%d", r, r)})
	}
	s.Set(3, 3, grid.Cell{Formula: "A3*B2"})
	old := func(r int) string {
		k, _ := formula.Pattern(folded(fmt.Sprintf("a%d*B%d", r, r)), r, 3)
		return k
	}
	now, _ := formula.Pattern(folded("A3*B2"), 3, 3)
	if got := brokenRow(s, 3, 3, old(3), now); !strings.Contains(got, "unlike the rows around it, which compute =A{r}*B{r}") {
		t.Errorf("one row broken: %q", got)
	}
	if got := brokenRow(s, 3, 3, old(3), ""); !strings.Contains(got, "a value where") {
		t.Errorf("a value: %q", got)
	}
	// The whole column changed alike: no flag.
	t2 := grid.NewSheet("S")
	for r := 2; r <= 4; r++ {
		t2.Set(r, 3, grid.Cell{Formula: fmt.Sprintf("A%d+B%d", r, r)})
	}
	now2, _ := formula.Pattern(folded("A3+B3"), 3, 3)
	if got := brokenRow(t2, 3, 3, old(3), now2); got != "" {
		t.Errorf("a column changed alike: %q", got)
	}
}

func TestFolded(t *testing.T) {
	if got := folded(`maxifs(a1:a2,'my sheet'!b1,"<>Mid")`); got != `MAXIFS(A1:A2,'MY SHEET'!B1,"<>Mid")` {
		t.Errorf("folded %s", got)
	}
	// LibreOffice writes TRUE() and FALSE() for TRUE and FALSE.
	for _, c := range []struct{ in, want string }{
		{`VLOOKUP("k2",A1:B1,2,FALSE())`, `VLOOKUP("k2",A1:B1,2,FALSE)`},
		{`IF(B1>0,TRUE(),false())`, `IF(B1>0,TRUE,FALSE)`},
		{`AND(true(),NOT(FALSE()))`, `AND(TRUE,NOT(FALSE))`},
		{`TRUE()`, `TRUE`},
		{`if(a1,"TRUE()","x")`, `IF(A1,"TRUE()","x")`},
		{`MYTRUE()+XFALSE()`, `MYTRUE()+XFALSE()`},
		{`Sheet.TRUE()`, `SHEET.TRUE()`},
	} {
		if got := folded(c.in); got != c.want {
			t.Errorf("folded %s = %s, want %s", c.in, got, c.want)
		}
	}
	if folded(`IF(B1>0,TRUE,FALSE)`) != folded(`IF(B1>0,TRUE(),FALSE())`) {
		t.Errorf("TRUE and TRUE() differ")
	}
}

func TestPairReplaced(t *testing.T) {
	p := func(o, n int) diff.Pair { return diff.Pair{Old: o, New: n} }
	got := pairReplaced([]diff.Pair{p(1, 1), p(2, 0), p(3, 0), p(0, 2), p(4, 3), p(0, 4), p(5, 5)})
	want := []diff.Pair{p(1, 1), p(2, 2), p(3, 0), p(4, 3), p(0, 4), p(5, 5)}
	if !reflect.DeepEqual(got, want) {
		t.Errorf("pairs %v", got)
	}
}

func TestRuns(t *testing.T) {
	if got := runs([]int{7, 3, 4, 5, 9}); !reflect.DeepEqual(got, [][2]int{{3, 5}, {7, 7}, {9, 9}}) {
		t.Errorf("runs %v", got)
	}
}

// TestSameRef checks that a name Excel rewrote for lines added and removed
// and sheets renamed is the same name, and one that points elsewhere is not.
func TestSameRef(t *testing.T) {
	r := &report{renames: [][2]string{{"data", "Facts"}}}
	r.setEdits([]lineEdit{{sheet: "Summary", rowsAdded: [][2]int{{1, 1}}}, {sheet: "Facts", rowsRemoved: [][2]int{{6, 6}}}})
	for _, c := range []struct {
		old, now string
		same     bool
	}{
		{"Summary!$B$1", "Summary!$B$1", true},
		{"Summary!$B$1", "Summary!$B$2", true},         // a row added above
		{"Summary!$B$1", "Summary!$B$3", false},        // moved elsewhere
		{"Data!$B$2:$B$6", "Facts!$B$2:$B$5", true},    // renamed, and its last row removed
		{"Data!$B$2:$B$6", "Facts!$B$2:$B$6", false},   // renamed, but not shrunk
		{"Data!$B$2:$B$6", "Summary!$B$2:$B$6", false}, // another sheet
	} {
		if got := r.sameRef(c.old, c.now); got != c.same {
			t.Errorf("%s → %s: %v, want %v", c.old, c.now, got, c.same)
		}
	}
	if (&report{}).sameRef("Summary!$B$1", "Summary!$B$2") {
		t.Errorf("nothing was added, yet the name moved")
	}
}

// TestReplay checks that what Excel writes for lines removed and added is
// worked out for many runs of them (deleting every fourth row of a table is
// hundreds), and that lines past what a formula reads touch nothing of it.
func TestReplay(t *testing.T) {
	var removed [][2]int
	for r := 5; r <= 1001; r += 4 {
		removed = append(removed, [2]int{r, r})
	}
	r := &report{}
	r.setEdits([]lineEdit{{sheet: "Data", rowsRemoved: removed}})
	f, ok := r.replay("SUM(B2:B1001)", "Data", false)
	if want := "SUM(B2:B751)"; !ok || f != want {
		t.Errorf("replay %q %v, want %q", f, ok, want)
	}
	// A formula on another sheet, and one that reads only rows above.
	f, ok = r.replay("SUM(Data!B2:B1001)+B3", "Summary", false)
	if want := "SUM(Data!B2:B751)+B3"; !ok || f != want {
		t.Errorf("replay %q %v, want %q", f, ok, want)
	}
	if f, _ := r.replay("B2*C2", "Data", false); f != "B2*C2" {
		t.Errorf("above the rows removed: %q", f)
	}
	// Rows added or removed within what a sort moved are left out when sure.
	r = &report{}
	r.setEdits([]lineEdit{{sheet: "Data", rowsAdded: [][2]int{{3, 3}}, movedNew: [2]int{2, 8}}})
	if f, _ := r.replay("SUM(B2:B6)", "Data", true); f != "SUM(B2:B6)" {
		t.Errorf("within a sort, sure: %q", f)
	}
	if f, _ := r.replay("SUM(B2:B6)", "Data", false); f != "SUM(B2:B7)" {
		t.Errorf("within a sort: %q", f)
	}
	if _, ok := (&report{}).replay("SUM(B2:B6)", "Data", false); ok {
		t.Errorf("nothing to replay, yet it did")
	}
}

// Deleting every fourth row of a table (a filter, a script's delete_rows) is
// hundreds of runs of rows removed: the total Excel shrank is no formula
// changed, and its value moved with the rows removed as its cause.
func TestCheckAgainstManyRuns(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("SHEETS_CACHE_DIR", filepath.Join(dir, ".cache"))
	build := func(path string, drop bool) {
		var b strings.Builder
		b.WriteString("Region\tUnits\tPrice\n")
		n := 0
		for i := 1; i <= 1000; i++ {
			if drop && i%4 == 0 {
				continue
			}
			n++
			fmt.Fprintf(&b, "R%d\t%d\t%d\n", i, i%50, 1+i%13)
		}
		fmt.Fprintf(&b, "Total\t=SUM(B2:B%d)\n", n+1)
		if out, code := runCLI(b.String(), "write", path, "Data!A1"); code != 0 {
			t.Fatalf("write: %s", out)
		}
	}
	before, after := filepath.Join(dir, "a.xlsx"), filepath.Join(dir, "b.xlsx")
	build(before, false)
	build(after, true)
	js, code := runCLI("", "--json", "check", after, "--against", before)
	if code != 0 {
		t.Fatalf("exit %d:\n%s", code, js)
	}
	var r againstResult
	if err := json.Unmarshal([]byte(js), &r); err != nil {
		t.Fatal(err)
	}
	if r.Summary.Formulas != 0 || len(r.Changed) != 0 {
		t.Errorf("formulas changed: %+v", r.Changed)
	}
	if len(r.Moved) != 1 || !strings.HasPrefix(r.Moved[0].Cause, "rows removed: ") {
		t.Errorf("moved: %+v", r.Moved)
	}
}

// runCLI runs sheets as a command, in-process.
func runCLI(stdin string, args ...string) (string, int) {
	var out, errs bytes.Buffer
	code := Main(args, strings.NewReader(stdin), &out, &errs)
	return out.String() + errs.String(), code
}

// rewriteSheet copies a workbook, its first sheet's XML rewritten.
func rewriteSheet(t *testing.T, src, dst string, rewrite func([]byte) []byte) {
	t.Helper()
	zr, err := zip.OpenReader(src)
	if err != nil {
		t.Fatal(err)
	}
	defer zr.Close()
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	for _, f := range zr.File {
		rc, err := f.Open()
		if err != nil {
			t.Fatal(err)
		}
		data, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			t.Fatal(err)
		}
		if f.Name == "xl/worksheets/sheet1.xml" {
			was := data
			if data = rewrite(data); bytes.Equal(data, was) {
				t.Fatal("nothing rewritten")
			}
		}
		w, err := zw.Create(f.Name)
		if err != nil {
			t.Fatal(err)
		}
		w.Write(data)
	}
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(dst, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
}

// TestCheckAgainstStoredResults checks the results a file stores that the
// sheets calculator disagrees with. An input changed with the results left
// as they were (a tool writing the XML, Excel calculating by hand): the
// formulas reading it hold what they held, which is unknown, with the input
// named, not clean. And a result Excel stored that sheets computes otherwise,
// in the old version, against a new one with no results (openpyxl's round
// trip): the difference is the calculator's, unknown, not unexplained.
func TestCheckAgainstStoredResults(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("SHEETS_CACHE_DIR", filepath.Join(dir, ".cache"))
	before := filepath.Join("..", "xlsx", "testdata", "against_before.xlsx")
	type result struct {
		Summary     againstSummary `json:"summary"`
		Unexplained []movedValue   `json:"unexplained"`
		Unknown     []unknownValue `json:"unknown"`
	}
	run := func(after, before string) (result, int) {
		t.Helper()
		out, code := runCLI("", "--json", "check", after, "--against", before)
		var r result
		if err := json.Unmarshal([]byte(out), &r); err != nil {
			t.Fatalf("%v: %s", err, out)
		}
		return r, code
	}

	stale := filepath.Join(dir, "stale.xlsx")
	rewriteSheet(t, before, stale, func(b []byte) []byte {
		return bytes.Replace(b, []byte(`<c r="B2"><v>5</v>`), []byte(`<c r="B2"><v>6</v>`), 1)
	})
	r, code := run(stale, before)
	if r.Summary.Inputs != 1 || r.Summary.Moved != 0 || r.Summary.Unexplained != 0 {
		t.Errorf("results left as they were: %+v", r.Summary)
	}
	held := map[string]bool{}
	for _, u := range r.Unknown {
		held[u.Cell] = true
		if !strings.HasPrefix(u.Why, "its result may be out of date") || !strings.Contains(u.Why, "(input Data!B2)") || u.Before != u.After {
			t.Errorf("results left as they were: %+v", u)
		}
	}
	for _, c := range []string{"Data!D2", "Data!B7", "Summary!B2", "Summary!B7"} {
		if !held[c] {
			t.Errorf("results left as they were: %s not unknown (%+v)", c, r.Unknown)
		}
	}
	if code != 0 {
		t.Errorf("results left as they were: exit %d", code)
	}
	// Against itself: nothing changed, so nothing is out of date.
	if r, _ := run(stale, stale); len(r.Unknown) != 0 {
		t.Errorf("against itself: %+v", r.Unknown)
	}

	excel := filepath.Join(dir, "excel.xlsx")
	rewriteSheet(t, before, excel, func(b []byte) []byte {
		return bytes.Replace(b, []byte("<f>B3*C3</f><v>140</v>"), []byte("<f>B3*C3</f><v>141</v>"), 1)
	})
	bare := filepath.Join(dir, "bare.xlsx")
	results := regexp.MustCompile(`(</f>)<v>[^<]*</v>`)
	rewriteSheet(t, excel, bare, func(b []byte) []byte { return results.ReplaceAll(b, []byte("$1")) })
	for _, c := range [][2]string{{bare, excel}, {excel, bare}} {
		r, _ := run(c[0], c[1])
		if len(r.Unexplained) != 0 || r.Summary.Moved != 0 {
			t.Errorf("%s against %s: unexplained %+v, %d moved", filepath.Base(c[0]), filepath.Base(c[1]), r.Unexplained, r.Summary.Moved)
		}
		if len(r.Unknown) != 1 || r.Unknown[0].Cell != "Data!D3" || !strings.HasPrefix(r.Unknown[0].Why, "the difference is the calculator's, not a change") {
			t.Errorf("%s against %s: unknown %+v", filepath.Base(c[0]), filepath.Base(c[1]), r.Unknown)
		}
	}
}

// A file with many findings (each formula that breaks its column's pattern
// is one) does not make the text output long: ten are listed, with the
// count of the rest; --json lists them all.
func TestCheckAgainstFindingsBounded(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("SHEETS_CACHE_DIR", filepath.Join(dir, ".cache"))
	build := func(path string, broken bool) {
		var b strings.Builder
		b.WriteString("Label\tVal\tCalc\n")
		for i := 2; i <= 401; i++ {
			if broken && i%4 == 0 {
				fmt.Fprintf(&b, "L%d\t%d\t=B%d*%d\n", i, i, i-1, i)
			} else {
				fmt.Fprintf(&b, "L%d\t%d\t=B%d*2\n", i, i, i)
			}
		}
		if out, code := runCLI(b.String(), "write", path, "Data!A1"); code != 0 {
			t.Fatalf("write: %s", out)
		}
	}
	before, after := filepath.Join(dir, "a.xlsx"), filepath.Join(dir, "b.xlsx")
	build(before, false)
	build(after, true)
	out, code := runCLI("", "check", after, "--against", before)
	if lines := strings.Split(strings.TrimSpace(out), "\n"); len(lines) > 60 {
		t.Errorf("%d lines:\n%s", len(lines), out)
	}
	if code != 0 && code != 1 {
		t.Errorf("exit %d", code)
	}
	if !strings.Contains(out, "  … ") || !strings.Contains(out, "more (--json lists them all)") {
		t.Errorf("no count of the rest:\n%s", out)
	}
	js, _ := runCLI("", "--json", "check", after, "--against", before)
	var r struct {
		Findings []json.RawMessage `json:"findings"`
	}
	if err := json.Unmarshal([]byte(js), &r); err != nil {
		t.Fatal(err)
	}
	if len(r.Findings) < 100 {
		t.Errorf("--json lists %d findings, want them all (100 broken rows)", len(r.Findings))
	}
}

// TestCheckAgainstLarge checks that check --against stays bounded, and
// records its time, on 150,000 rows of which every tenth changed. Run
// with SHEETS_LARGE=1.
func TestCheckAgainstLarge(t *testing.T) {
	if os.Getenv("SHEETS_LARGE") == "" {
		t.Skip("SHEETS_LARGE=1 runs it (150,000 rows)")
	}
	const rows = 150000
	dir := t.TempDir()
	t.Setenv("SHEETS_CACHE_DIR", filepath.Join(dir, ".cache"))
	run := func(stdin string, args ...string) (string, int) {
		var out, errs bytes.Buffer
		code := Main(args, strings.NewReader(stdin), &out, &errs)
		return out.String() + errs.String(), code
	}
	build := func(path string, changed bool) {
		var b strings.Builder
		b.WriteString("Region\tUnits\tPrice\tRevenue\n")
		for i := 1; i <= rows; i++ {
			units := i % 50
			if changed && i%10 == 0 {
				units += 1
			}
			fmt.Fprintf(&b, "R%d\t%d\t%d\t=B{r}*C{r}\n", i%7, units, 1+i%13)
		}
		if out, code := run(b.String(), "write", path, "Data!A1"); code != 0 {
			t.Fatalf("write: %s", out)
		}
		last := rows + 1
		if out, code := run("", "set", path, "Summary!A1", "Units", "Summary!B1", fmt.Sprintf("=SUM(Data!B2:B%d)", last),
			"Summary!A2", "Revenue", "Summary!B2", "=SUM(Data!D:D)"); code != 0 {
			t.Fatalf("set: %s", out)
		}
	}
	before, after := filepath.Join(dir, "before.xlsx"), filepath.Join(dir, "after.xlsx")
	start := time.Now()
	build(before, false)
	build(after, true)
	t.Logf("wrote two workbooks of %d rows in %s", rows, time.Since(start).Round(time.Millisecond))

	start = time.Now()
	out, code := run("", "check", after, "--against", before)
	took := time.Since(start)
	t.Logf("check --against on %d rows: %s", rows, took.Round(time.Millisecond))
	if code != 0 {
		t.Errorf("exit %d:\n%s", code, out)
	}
	lines := strings.Split(strings.TrimSpace(out), "\n")
	if len(lines) > 60 || len(out) > 6000 {
		t.Errorf("output not bounded: %d lines, %d bytes:\n%s", len(lines), len(out), out)
	}
	head := fmt.Sprintf("%s against %s: 15000 inputs changed, 0 formulas changed, 15002 computed values moved (each with its cause), 0 unexplained, 0 features lost, 0 unknown", after, before)
	if lines[0] != head {
		t.Errorf("summary:\n got %s\nwant %s", lines[0], head)
	}
	for _, want := range []string{"  … 14990 more (--json lists them all)", "  … 14992 more (--json lists them all)"} {
		if !strings.Contains(out, want) {
			t.Errorf("no %q in:\n%s", want, out)
		}
	}
	t.Logf("output, %d lines:\n%s", len(lines), out)

	// Every tenth formula broken, not an input changed: 15,000 findings and
	// 15,000 formulas changed, flagged; the text stays as short.
	broken := filepath.Join(dir, "broken.xlsx")
	{
		var b strings.Builder
		b.WriteString("Region\tUnits\tPrice\tRevenue\n")
		for i := 1; i <= rows; i++ {
			if i%10 == 0 {
				fmt.Fprintf(&b, "R%d\t%d\t%d\t=B{r-1}*C{r}\n", i%7, i%50, 1+i%13)
			} else {
				fmt.Fprintf(&b, "R%d\t%d\t%d\t=B{r}*C{r}\n", i%7, i%50, 1+i%13)
			}
		}
		if out, code := run(b.String(), "write", broken, "Data!A1"); code != 0 {
			t.Fatalf("write: %s", out)
		}
		if out, code := run("", "set", broken, "Summary!A1", "Units", "Summary!B1", fmt.Sprintf("=SUM(Data!B2:B%d)", rows+1),
			"Summary!A2", "Revenue", "Summary!B2", "=SUM(Data!D:D)"); code != 0 {
			t.Fatalf("set: %s", out)
		}
		start = time.Now()
		out, _ := run("", "check", broken, "--against", before)
		t.Logf("check --against on %d rows with every tenth formula broken: %s", rows, time.Since(start).Round(time.Millisecond))
		lines := strings.Split(strings.TrimSpace(out), "\n")
		if len(lines) > 60 || len(out) > 12000 {
			t.Errorf("output with findings not bounded: %d lines, %d bytes:\n%s", len(lines), len(out), out)
		}
		if !strings.Contains(out, "15000 formulas changed") || !strings.Contains(out, "more (--json lists them all)") {
			t.Errorf("broken formulas:\n%s", out)
		}
		t.Logf("output with findings, %d lines, %d bytes", len(lines), len(out))
	}

	start = time.Now()
	js, code := run("", "--json", "check", after, "--against", before)
	t.Logf("check --against --json on %d rows: %s, %d bytes", rows, time.Since(start).Round(time.Millisecond), len(js))
	var r againstResult
	if err := json.Unmarshal([]byte(js), &r); err != nil || code != 0 {
		t.Fatalf("json (exit %d): %v", code, err)
	}
	if len(r.Changed) != 15000 || len(r.Moved) != 15002 || len(r.Unexplained) != 0 {
		t.Errorf("json: %d changed, %d moved, %d unexplained", len(r.Changed), len(r.Moved), len(r.Unexplained))
	}
	if m := r.Moved[len(r.Moved)-1]; m.Cause == "" {
		t.Errorf("no cause: %+v", m)
	}
}

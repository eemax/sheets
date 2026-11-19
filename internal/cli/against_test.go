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
	// A validation or conditional format by its type and rule, a chart by its
	// type and the data it shows.
	rule := func(kind, sheet, where, what, typ, rule string) xlsx.Feature {
		return xlsx.Feature{Kind: kind, Sheet: sheet, Where: where, What: what, Type: typ, Rule: rule}
	}
	chart := func(sheet, where string, series ...string) xlsx.Feature {
		return xlsx.Feature{Kind: xlsx.FeatureChart, Sheet: sheet, Where: where, What: "column chart", Type: "column", Series: series}
	}
	whole := func(where string) xlsx.Feature {
		return rule(xlsx.FeatureValidation, "Data", where, "whole validation", "whole", "between|0|100")
	}
	decimal := func(where string) xlsx.Feature {
		return rule(xlsx.FeatureValidation, "Data", where, "decimal validation", "decimal", "greaterThanOrEqual|0")
	}
	before := []xlsx.Feature{
		whole("B2:B6"),
		rule(xlsx.FeatureValidation, "Data", "C2:C6", "list validation", "list", "Summary!$D$2:$D$6"),
		chart("Data", "F2", "Data!$D$2:$D$6"),
		rule(xlsx.FeatureCond, "Old", "A1:A3", "formula rule", "expression", "$A1>5"),
		f(xlsx.FeatureName, "", "Rate", "Data!$B$1"),
		f(xlsx.FeatureName, "", "Gone", "Data!$B$2"),
		f(xlsx.FeatureTable, "Data", "A1:D6", "Sales"),
		f(xlsx.FeatureShape, "Removed", "A1", "TextBox 1"),
	}
	after := []xlsx.Feature{
		whole("B2:B7"),                        // grew with a row
		chart("Data", "G2", "Data!$D$2:$D$6"), // moved
		rule(xlsx.FeatureCond, "New", "A1:A3", "formula rule", "expression", "$A1>5"), // its sheet renamed
		f(xlsx.FeatureName, "", "rate", "Data!$B$1"),
		f(xlsx.FeatureName, "", "Other", "Data!$B$2"), // a name is matched by name
		f(xlsx.FeatureTable, "Data", "A1:D7", "Sales"),
	}
	got := lostFeatures(before, after, map[string]string{"data": "Data", "old": "New", "removed": ""}, nil, nil)
	want := []lostFeature{
		{Kind: xlsx.FeatureValidation, Where: "Data!C2:C6", What: "list validation"},
		{Kind: xlsx.FeatureName, Where: "Gone", What: "Data!$B$2"},
		{Kind: xlsx.FeatureShape, Where: "Removed!A1", What: "TextBox 1"},
	}
	if !reflect.DeepEqual(got, want) {
		t.Errorf("lost\n got %+v\nwant %+v", got, want)
	}
	if got := lostFeatures(nil, after, nil, nil, nil); got == nil || len(got) != 0 {
		t.Errorf("nothing before: %+v", got)
	}
	// Both validations grew with a row inserted, then the whole one was
	// cleared: the one lost is the whole one, not the decimal one, which is
	// now in another place.
	got = lostFeatures([]xlsx.Feature{whole("B2:B6"), decimal("C2:C6")}, []xlsx.Feature{decimal("C2:C7")}, map[string]string{"data": "Data"}, nil, nil)
	if want := []lostFeature{{Kind: xlsx.FeatureValidation, Where: "Data!B2:B6", What: "whole validation"}}; !reflect.DeepEqual(got, want) {
		t.Errorf("lost after a move\n got %+v\nwant %+v", got, want)
	}
	// A column deleted took the whole validation: the decimal one now where it
	// was is not it.
	deleted := func(home, ref string) string {
		return map[string]string{"B2:B6": "", "C2:C6": "B2:B6"}[ref]
	}
	got = lostFeatures([]xlsx.Feature{whole("B2:B6"), decimal("C2:C6")}, []xlsx.Feature{decimal("B2:B6")}, map[string]string{"data": "Data"}, nil, deleted)
	if want := []lostFeature{{Kind: xlsx.FeatureValidation, Where: "Data!B2:B6", What: "whole validation"}}; !reflect.DeepEqual(got, want) {
		t.Errorf("lost with a column\n got %+v\nwant %+v", got, want)
	}
	// Three alike, the middle column deleted: the one lost is the one that was
	// there, not the last, which moved into its place.
	middle := func(home, ref string) string {
		return map[string]string{"B2:B6": "B2:B6", "C2:C6": "", "D2:D6": "C2:C6"}[ref]
	}
	got = lostFeatures([]xlsx.Feature{whole("B2:B6"), whole("C2:C6"), whole("D2:D6")}, []xlsx.Feature{whole("B2:B6"), whole("C2:C6")}, map[string]string{"data": "Data"}, nil, middle)
	if want := []lostFeature{{Kind: xlsx.FeatureValidation, Where: "Data!C2:C6", What: "whole validation"}}; !reflect.DeepEqual(got, want) {
		t.Errorf("lost of three alike\n got %+v\nwant %+v", got, want)
	}
	// Another of its kind is no reason to say nothing of the one that is gone:
	// a validation of another type on other cells is not the one it was. One
	// on the same cells whose formula, operator or type changed is it,
	// edited: changed, not lost.
	if got := lostFeatures([]xlsx.Feature{whole("B2:B6")}, []xlsx.Feature{decimal("B2:B7")}, nil, nil, nil); len(got) != 1 {
		t.Errorf("another type on other cells: %+v", got)
	}
	for name, now := range map[string]xlsx.Feature{
		"its type":     decimal("B2:B6"),
		"its formula":  rule(xlsx.FeatureValidation, "Data", "B2:B6", "whole validation", "whole", "between|0|200"),
		"its operator": rule(xlsx.FeatureValidation, "Data", "B2:B6", "whole validation", "whole", "notBetween|0|100"),
	} {
		lost, changed := compareFeatures([]xlsx.Feature{whole("B2:B6")}, []xlsx.Feature{now}, nil, nil, nil)
		if len(lost) != 0 || len(changed) != 1 || changed[0].Where != "Data!B2:B6" || changed[0].Before == changed[0].After {
			t.Errorf("changed %s: lost %+v, changed %+v", name, lost, changed)
		}
	}
	// A conditional format replaced by another kind (a banding rule) is lost,
	// whatever else the sheet holds.
	cond := rule(xlsx.FeatureCond, "Data", "D2:D6", "cell value rule", "cellIs", "greaterThan|140")
	band := rule(xlsx.FeatureCond, "Data", "A2:D6", "formula rule", "expression", "MOD(ROW(),2)=0")
	if got := lostFeatures([]xlsx.Feature{cond}, []xlsx.Feature{band}, nil, nil, nil); len(got) != 1 || got[0].Kind != xlsx.FeatureCond {
		t.Errorf("replaced conditional format: %+v", got)
	}
	// A chart added does not stand for the one dropped, though it is of its
	// type: it shows other data, and is somewhere else. One of the same data
	// is the chart, moved.
	old := chart("Data", "F2", "Data!$D$2:$D$6")
	if got := lostFeatures([]xlsx.Feature{old}, []xlsx.Feature{chart("Data", "H20", "Data!$B$2:$B$6")}, nil, nil, nil); len(got) != 1 || got[0].Kind != xlsx.FeatureChart {
		t.Errorf("chart replaced by another: %+v", got)
	}
	if got := lostFeatures([]xlsx.Feature{old}, []xlsx.Feature{chart("Data", "H20", "Data!$D$2:$D$6")}, nil, nil, nil); len(got) != 0 {
		t.Errorf("chart moved: %+v", got)
	}
	// The same place, other data: the chart's data was edited.
	if got := lostFeatures([]xlsx.Feature{old}, []xlsx.Feature{chart("Data", "F2", "Data!$D$2:$D$9")}, nil, nil, nil); len(got) != 0 {
		t.Errorf("chart edited in place: %+v", got)
	}
	// Rows inserted above it moved its data and its place: where the lines
	// added put them, it is where it was.
	shifted := func(home, ref string) string {
		return map[string]string{"F2": "F5", "Data!$D$2:$D$6": "Data!$D$5:$D$9"}[ref]
	}
	if got := lostFeatures([]xlsx.Feature{old}, []xlsx.Feature{chart("Data", "F5", "Data!$D$5:$D$9")}, nil, nil, shifted); len(got) != 0 {
		t.Errorf("chart after rows inserted: %+v", got)
	}
	// An Excel table renamed is the table, by its columns; with none renamed,
	// another table is no stand-in.
	tbl := func(name, where string) xlsx.Feature {
		return xlsx.Feature{Kind: xlsx.FeatureTable, Sheet: "Data", Where: where, What: name, Rule: "Region|Units"}
	}
	renamedTo := tableRenames([]xlsx.Feature{tbl("Sales", "A1:B3")}, []xlsx.Feature{tbl("Sales2", "A1:B3")}, nil)
	if !reflect.DeepEqual(renamedTo, map[string]string{"sales": "Sales2"}) {
		t.Errorf("table renames: %v", renamedTo)
	}
	if got := lostFeatures([]xlsx.Feature{tbl("Sales", "A1:B3")}, []xlsx.Feature{tbl("Sales2", "A1:B3")}, nil, renamedTo, nil); len(got) != 0 {
		t.Errorf("table renamed: %+v", got)
	}
	if got := lostFeatures([]xlsx.Feature{tbl("Sales", "A1:B3")}, []xlsx.Feature{tbl("Sales2", "A1:B3")}, nil, nil, nil); len(got) != 1 {
		t.Errorf("table by another name: %+v", got)
	}
	other := tbl("Other", "A1:B3")
	other.Rule = "Item|Price"
	if got := tableRenames([]xlsx.Feature{tbl("Sales", "A1:B3")}, []xlsx.Feature{other}, nil); len(got) != 0 {
		t.Errorf("a table of other columns is no rename: %v", got)
	}
	// A shape or picture is kept by its place, or by its name elsewhere; one
	// of another name in another place is not it.
	shape := f(xlsx.FeatureShape, "Data", "F2", "TextBox 2")
	if got := lostFeatures([]xlsx.Feature{shape}, []xlsx.Feature{f(xlsx.FeatureShape, "Data", "F2", "Shape 1")}, nil, nil, nil); len(got) != 0 {
		t.Errorf("shape renamed in place: %+v", got)
	}
	if got := lostFeatures([]xlsx.Feature{shape}, []xlsx.Feature{f(xlsx.FeatureShape, "Data", "H9", "TextBox 2")}, nil, nil, nil); len(got) != 0 {
		t.Errorf("shape moved: %+v", got)
	}
	if got := lostFeatures([]xlsx.Feature{shape}, []xlsx.Feature{f(xlsx.FeatureShape, "Data", "H9", "Other")}, nil, nil, nil); len(got) != 1 {
		t.Errorf("another shape: %+v", got)
	}
}

// ruleKey compares what a rule tests: the cells its formulas read are no part
// of it (rows and columns added or removed rewrite them), a sheet renamed no
// more; the text it matches and the numbers it compares with are.
func TestRuleKey(t *testing.T) {
	same := [][2]string{
		{"between|Summary!$D$2:$D$6", "between|Facts!$D$3:$D$7"},
		{"expression|$B2>$C2", "expression|$B5 > $C5"},
		{"greaterThan|140", "GREATERTHAN|140"},
		{"between|'Q 2026'!A1:A9", "between|Q!A1:A12"},
		{"between|Sales[Units]", "between|Sales2[Units]"},
		{"between|SUM(B2:B6)>5", "between|sum(C3:C9)>5"},
	}
	for _, c := range same {
		if ruleKey(c[0]) != ruleKey(c[1]) {
			t.Errorf("%q and %q differ: %q, %q", c[0], c[1], ruleKey(c[0]), ruleKey(c[1]))
		}
	}
	differ := [][2]string{
		{"greaterThan|140", "greaterThan|150"},
		{"between|0|100", "between|0|1000"},
		{`between|"Yes,No"`, `between|"Yes,Maybe"`},
		{`between|"A1,B2"`, `between|"A1,B3"`},
		{"expression|$B2>$C2", "expression|$B2<$C2"},
		{"expression|$B2>5", "expression|$B2>$C2"},
		{"between|LOG10(B2)", "between|LOG2(B2)"},
	}
	for _, c := range differ {
		if ruleKey(c[0]) == ruleKey(c[1]) {
			t.Errorf("%q and %q are alike: %q", c[0], c[1], ruleKey(c[0]))
		}
	}
}

func TestRenameTable(t *testing.T) {
	for _, c := range []struct{ in, want string }{
		{"SUM(Sales[Units])", "SUM(Sales2[Units])"},
		{"SUM(SALES[[#This Row],[Units]])+Sales[@Units]", "SUM(Sales2[[#This Row],[Units]])+Sales2[@Units]"},
		{`"Sales[x]"&Sales[Units]`, `"Sales[x]"&Sales2[Units]`},
		{"MySales[Units]", "MySales[Units]"},
		{"Data!Sales[Units]", "Data!Sales[Units]"},
		{"Sales", "Sales"},
	} {
		if got := renameTable(c.in, "sales", "Sales2"); got != c.want {
			t.Errorf("%s: %s, want %s", c.in, got, c.want)
		}
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
	got := lostFeatures(before, nil, map[string]string{"data": "Data"}, nil, nil)
	want := []lostFeature{{Kind: xlsx.FeatureChart, Where: "Chart1", What: "bar chart on a chart sheet"}}
	if !reflect.DeepEqual(got, want) {
		t.Errorf("lost\n got %+v\nwant %+v", got, want)
	}
	if got := lostFeatures(before, before, nil, nil, nil); len(got) != 0 {
		t.Errorf("kept: %+v", got)
	}
	if got := lostFeatures(before, []xlsx.Feature{chart("Chart 1")}, nil, nil, nil); len(got) != 0 {
		t.Errorf("renamed: %+v", got)
	}
	// Two chart sheets, one lost: the one left keeps its own.
	two := []xlsx.Feature{chart("A"), chart("B")}
	if got := lostFeatures(two, []xlsx.Feature{chart("B")}, nil, nil, nil); len(got) != 1 || got[0].Where != "A" {
		t.Errorf("one of two: %+v", got)
	}
	// A chart on a worksheet is not matched by a chart sheet's.
	onSheet := xlsx.Feature{Kind: xlsx.FeatureChart, Sheet: "Data", Where: "F2", What: "bar chart"}
	if got := lostFeatures([]xlsx.Feature{onSheet}, before, map[string]string{"data": "Data"}, nil, nil); len(got) != 1 {
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
	got := pairReplaced([]diff.Pair{p(1, 1), p(2, 0), p(3, 0), p(0, 2), p(4, 3), p(0, 4), p(5, 5)}, nil)
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
func mustRead(t *testing.T, path string) []byte {
	t.Helper()
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	return data
}

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

	// The receipt of it: bounded too, its pictures at most three ranges of
	// at most receiptRows rows, the rest counted.
	{
		receipt := filepath.Join(dir, "receipt.md")
		start = time.Now()
		js, code := run("", "--json", "check", after, "--against", before, "--receipt", receipt)
		t.Logf("check --against --receipt on %d rows: %s", rows, time.Since(start).Round(time.Millisecond))
		var r againstResult
		if err := json.Unmarshal([]byte(js), &r); err != nil || code != 0 || r.Receipt == nil {
			t.Fatalf("receipt (exit %d): %v", code, err)
		}
		md := string(mustRead(t, receipt))
		if len(md) > 30000 {
			t.Errorf("receipt not bounded: %d bytes", len(md))
		}
		if len(r.Receipt.Previews) != receiptPreviews || r.Receipt.Previews[1].After != "Summary!A1:C3" || r.Receipt.Unshown == 0 || !strings.Contains(md, "… and 14975 more: `sheets --json check") {
			t.Errorf("receipt: %d previews, %d not pictured:\n%s", len(r.Receipt.Previews), r.Receipt.Unshown, md)
		}
		for _, p := range r.Receipt.Previews {
			for _, img := range []string{p.BeforeImage, p.AfterImage} {
				if fi, err := os.Stat(img); err != nil || fi.Size() > 1<<20 {
					t.Errorf("picture %s: %v", img, err)
				}
			}
		}
		t.Logf("receipt, %d bytes", len(md))
	}

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

	// A block of cells beside the table cut and pasted below it: moved, not
	// changes, though the table's rows hold it and rows are added below.
	{
		side, cut := filepath.Join(dir, "side.xlsx"), filepath.Join(dir, "cut.xlsx")
		data, err := os.ReadFile(before)
		if err != nil {
			t.Fatal(err)
		}
		if err := os.WriteFile(side, data, 0o644); err != nil {
			t.Fatal(err)
		}
		if out, code := run("", "set", side, "Data!H1", "Notes", "Data!H2", "10", "Data!H3", "20", "Summary!C1", "=SUM(Data!H2:H3)"); code != 0 {
			t.Fatalf("set: %s", out)
		}
		if err := os.WriteFile(cut, mustRead(t, side), 0o644); err != nil {
			t.Fatal(err)
		}
		if out, code := run("", "move", cut, "Data!H1:H3", fmt.Sprintf("Data!H%d", rows+10)); code != 0 {
			t.Fatalf("move: %s", out)
		}
		start = time.Now()
		out, code := run("", "check", cut, "--against", side)
		t.Logf("check --against on %d rows with a block cut and pasted below: %s", rows, time.Since(start).Round(time.Millisecond))
		if code != 0 || !strings.Contains(out, ": 0 inputs changed, 0 formulas changed, 0 computed values moved, 0 unexplained, 0 features lost, 0 unknown, 3 rows moved") ||
			!strings.Contains(out, fmt.Sprintf("Data!H%d:H%d, were H1:H3", rows+10, rows+12)) {
			t.Errorf("block cut (exit %d):\n%s", code, out)
		}
		if n := len(strings.Split(strings.TrimSpace(out), "\n")); n > 12 {
			t.Errorf("block cut output not bounded: %d lines:\n%s", n, out)
		}
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

// Comparing features takes time in step with their number, not its square:
// one validation and one conditional format for each of n cells, four of
// each lost, the rest shifted by a row inserted above them. Four times the
// cells takes about four times as long (a scan of the new features for each
// old one would take sixteen).
func TestLostFeaturesMany(t *testing.T) {
	shift := func(_, ref string) string {
		rg, err := grid.ParseRange(ref)
		if err != nil {
			return ref
		}
		return grid.CellName(rg.R1+1, rg.C1)
	}
	run := func(n int) time.Duration {
		var before, after []xlsx.Feature
		for r := 1; r <= n; r++ {
			cell := grid.CellName(r, 2)
			before = append(before,
				xlsx.Feature{Kind: xlsx.FeatureValidation, Sheet: "Data", Where: cell, Type: "whole", Rule: "between|0|9"},
				xlsx.Feature{Kind: xlsx.FeatureCond, Sheet: "Data", Where: cell, Type: "cellIs", Rule: "greaterThan|5"})
			if r%(n/4) != 0 {
				after = append(after,
					xlsx.Feature{Kind: xlsx.FeatureValidation, Sheet: "Data", Where: grid.CellName(r+1, 2), Type: "whole", Rule: "between|0|9"},
					xlsx.Feature{Kind: xlsx.FeatureCond, Sheet: "Data", Where: grid.CellName(r+1, 2), Type: "cellIs", Rule: "greaterThan|5"})
			}
		}
		start := time.Now()
		got := lostFeatures(before, after, nil, nil, shift)
		took := time.Since(start)
		if len(got) != 8 {
			t.Fatalf("%d cells: %d lost, want 8: %v", n, len(got), got)
		}
		return took
	}
	small := run(5000)
	large := run(20000)
	t.Logf("5,000 cells: %s, 20,000: %s", small, large)
	if large > 10*small+50*time.Millisecond {
		t.Errorf("4 times the cells took %.1f times as long (%s, %s)", float64(large)/float64(small), small, large)
	}
}

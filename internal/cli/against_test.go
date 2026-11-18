package cli

import (
	"bytes"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
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

package cli

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"

	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/lint"
	"github.com/eemax/sheets/internal/xlsx"
)

func num(n float64) grid.Cell { return grid.Cell{Kind: grid.Number, Num: n} }

// A session follows cells through the parts of a command: a later part's
// rows inserted and deleted carry what an earlier one found, each value
// keeps its first result before and its last after, and a cell written over
// is a cell written, not a value moved.
func TestSessionCompose(t *testing.T) {
	s := newSession("book.xlsx")
	s.started = true
	s.effects(xlsx.Effects{
		Written: []xlsx.EffectCell{{Sheet: "Data", Row: 2, Col: 2}},
		Results: []xlsx.ResultChange{
			{Sheet: "Data", Row: 2, Col: 4, Before: num(6), Was: true, After: num(0), Now: true},
			{Sheet: "Data", Row: 9, Col: 4, Before: num(1), Was: true, After: num(2), Now: true},
			{Sheet: "Sum", Row: 1, Col: 1, Before: num(1), Was: true, After: grid.Cell{Kind: grid.Error, Text: "#DIV/0!"}, Now: true},
		},
	})
	// Two rows inserted above row 2, then row 9 (now 11) deleted; D2 (now
	// D4) moves back, and Sum!A1 is computed again.
	s.effects(xlsx.Effects{
		Ops: []xlsx.LineOp{{Sheet: "data", At: 2, N: 2}, {Sheet: "Data", At: 11, N: -1}},
		Results: []xlsx.ResultChange{
			{Sheet: "Data", Row: 4, Col: 4, Before: num(0), Was: true, After: num(6), Now: true},
			{Sheet: "Sum", Row: 1, Col: 1, Before: grid.Cell{Kind: grid.Error, Text: "#DIV/0!"}, Was: true, After: num(3), Now: true},
		},
	})
	r := s.report()
	if r.Cells != 1 {
		t.Errorf("cells %d, want 1", r.Cells)
	}
	if r.Moved != 1 || r.MovedList[0] != (changeValue{Cell: "Sum!A1", Before: "1", After: "3"}) {
		t.Errorf("moved %d %v, want Sum!A1 1 → 3 alone (D4 is back at 6, D9 was deleted)", r.Moved, r.MovedList)
	}
	if r.Errors != 0 || r.Unknown != 0 {
		t.Errorf("errors %d, unknown %d, want none", r.Errors, r.Unknown)
	}
	if s.written[cellPos{"data", 4, 2}] == "" {
		t.Errorf("the cell written is not where the rows put it: %v", s.written)
	}

	// Written over with a value: its error goes with the formula.
	s = newSession("book.xlsx")
	s.started = true
	s.effects(xlsx.Effects{Written: []xlsx.EffectCell{{Sheet: "A", Row: 1, Col: 1}},
		Results: []xlsx.ResultChange{{Sheet: "A", Row: 1, Col: 1, Written: true, After: grid.Cell{Kind: grid.Error, Text: "#N/A"}, Now: true}}})
	if r := s.report(); r.Errors != 1 || r.ErrorList[0].Before != "" {
		t.Errorf("a formula written that shows an error: %+v", r)
	}
	s.effects(xlsx.Effects{Written: []xlsx.EffectCell{{Sheet: "A", Row: 1, Col: 1}}})
	if r := s.report(); r.Errors != 0 || r.Cells != 1 {
		t.Errorf("written over with a value: %+v", r)
	}

	// A sheet renamed carries its cells; one deleted takes them.
	s = newSession("book.xlsx")
	s.started = true
	s.effects(xlsx.Effects{Results: []xlsx.ResultChange{
		{Sheet: "Old", Row: 1, Col: 1, Before: num(1), Was: true, After: num(2), Now: true},
		{Sheet: "Gone", Row: 1, Col: 1, Before: num(1), Was: true, After: num(2), Now: true},
	}})
	s.effects(xlsx.Effects{Ops: []xlsx.LineOp{{Sheet: "Old", Rename: "New"}, {Sheet: "gone", Gone: true}}})
	if r := s.report(); r.Moved != 1 || r.MovedList[0].Cell != "New!A1" {
		t.Errorf("renamed and deleted: %+v", r)
	}
	if got := s.renamed(); got["old"] != "New" || got["gone"] != "" || len(got) != 2 {
		t.Errorf("renamed %v", got)
	}
}

// Unknown values: those that had a result, or were written, and have none
// now, grouped by why; a result missing before and still missing is not
// new.
func TestSessionUnknown(t *testing.T) {
	s := newSession("book.xlsx")
	s.started = true
	var rs []xlsx.ResultChange
	for i := 1; i <= 7; i++ {
		rs = append(rs, xlsx.ResultChange{Sheet: "A", Row: i, Col: 1, Before: num(1), Was: true, Why: "circular reference: Excel cannot compute it"})
	}
	rs = append(rs, xlsx.ResultChange{Sheet: "A", Row: 9, Col: 1, Written: true, Why: "INFO is not supported by the sheets calculator"},
		xlsx.ResultChange{Sheet: "A", Row: 10, Col: 1, Why: "never known"})
	s.effects(xlsx.Effects{Results: rs})
	r := s.report()
	if r.Unknown != 8 || len(r.UnknownList) != 2 {
		t.Fatalf("unknown %d in %d groups, want 8 in 2: %+v", r.Unknown, len(r.UnknownList), r.UnknownList)
	}
	if u := r.UnknownList[0]; u.Count != 7 || len(u.Cells) != changeListed || !strings.HasPrefix(u.Why, "circular") {
		t.Errorf("first group %+v", u)
	}
}

// Where the session's lines put a reference of the file before.
func TestShiftRef(t *testing.T) {
	s := &session{ops: []xlsx.LineOp{
		{Sheet: "Data", At: 2, N: 1},              // a row before row 2
		{Sheet: "Data", Cols: true, At: 3, N: -2}, // columns C:D deleted
		{Sheet: "Data", Rename: "Facts"},
	}}
	for _, c := range []struct{ home, ref, want string }{
		{"Data", "B2:B9", "B3:B10"},
		{"Data", "E5", "C6"},
		{"Data", "C2:D9", ""},      // its columns deleted
		{"Data", "B2:E2", "B3:C3"}, // part of it
		{"Data", "A:A", "A:A"},
		{"Other", "B2", "B2"},
		{"x", "Data!$B$2:$B$9", "Facts!B3:B10"},
	} {
		if got := s.shiftRef(c.home, c.ref); got != c.want {
			t.Errorf("shiftRef(%q, %q) = %q, want %q", c.home, c.ref, got, c.want)
		}
	}
}

// The lint compared rule by rule, by cells: a finding that grows counts,
// one gone counts down, and a new one is shown.
func TestLintDelta(t *testing.T) {
	before := []lint.Finding{
		{Rule: "error-value", Severity: lint.Error, Sheet: "A", Cells: []string{"B2"}, Message: "1 cell shows #DIV/0!"},
		{Rule: "mixed-types", Severity: lint.Note, Sheet: "A", Cells: []string{"C2"}, Message: "a number column also holds text"},
	}
	after := []lint.Finding{
		{Rule: "error-value", Severity: lint.Error, Sheet: "A", Cells: []string{"B2", "B3"}, Message: "2 cells show #DIV/0!"},
		{Rule: "broken-reference", Severity: lint.Error, Sheet: "A", Cells: []string{"D1"}, Message: "1 formula refers to deleted cells (#REF!)"},
	}
	d := lintDelta(before, after)
	if !d.Compared || d.Errors != 2 || d.Warnings != 0 || d.Notes != -1 {
		t.Fatalf("delta %+v", d)
	}
	if len(d.Rules) != 3 || d.Rules[2].Rule != "mixed-types" || d.Rules[2].Message != "" {
		t.Errorf("rules %+v", d.Rules)
	}
	if d.Rules[0].Message == "" || !strings.HasPrefix(d.Rules[0].Message, "A!") {
		t.Errorf("a rule that rose shows a finding: %+v", d.Rules[0])
	}
	if d := lintDelta(before, before); len(d.Rules) != 0 || lintHead(d) != "as before" {
		t.Errorf("no change: %+v", d)
	}
}

// The report, printed, keeps to its lines however much there is to say.
func TestPrintChangeBounded(t *testing.T) {
	r := &changeReport{Cells: 150000, Moved: 9000, Errors: 70, Unknown: 400, FeaturesLost: 12}
	for i := 0; i < changeListed; i++ {
		r.MovedList = append(r.MovedList, changeValue{Cell: "A!B1", Before: "1", After: "2"})
		r.ErrorList = append(r.ErrorList, changeValue{Cell: "A!C1", After: "#REF!"})
		r.FeaturesLostList = append(r.FeaturesLostList, lostFeature{Kind: "chart", Where: "A!F2"})
	}
	for i := 0; i < 8; i++ {
		r.UnknownList = append(r.UnknownList, changeUnknown{Why: strings.Repeat("why ", i+1), Count: 50, Cells: []string{"A!D1"}})
	}
	r.Lint = &changeLint{Compared: true, Errors: 70}
	for i := 0; i < changeRules; i++ {
		r.Lint.Rules = append(r.Lint.Rules, changeRule{Rule: "error-value", Before: 0, After: 70, Message: strings.Repeat("x", 300)})
	}
	var out bytes.Buffer
	(&app{stdout: &out}).printChange(r)
	lines := strings.Split(strings.TrimRight(out.String(), "\n"), "\n")
	if len(lines) >= 15 {
		t.Errorf("%d lines:\n%s", len(lines), out.String())
	}
	if !strings.Contains(out.String(), "unknown: 250 more for other reasons") {
		t.Errorf("the reasons past three are counted:\n%s", out.String())
	}
}

// withChange adds the report to a command's JSON as a field, its text as
// written.
func TestWithChange(t *testing.T) {
	b, err := marshalPlain(withChange(struct {
		File string `json:"file"`
	}{"a<b>.xlsx"}, &changeReport{Cells: 2}))
	if err != nil {
		t.Fatal(err)
	}
	if got := string(b); !strings.HasPrefix(got, `{"file":"a<b>.xlsx","change":{"cells":2,`) {
		t.Errorf("got %s", got)
	}
	if got := withChange([]int{1}, &changeReport{}); got == nil {
		t.Error("nil")
	}
}

// reportLines are the change report's lines in a command's output: from
// its head line to the next-step hints.
func reportLines(out string) []string {
	var lines []string
	in := false
	for _, l := range strings.Split(strings.TrimRight(out, "\n"), "\n") {
		switch {
		case strings.HasPrefix(l, "change: "):
			in = true
		case strings.HasPrefix(l, "next: "):
			in = false
		}
		if in {
			lines = append(lines, l)
		}
	}
	return lines
}

// On 150,000 rows, in both editors, the report after set, rows insert and
// sort, and after a batch of 50 edits (streaming), stays under 15 lines.
func TestChangeLarge(t *testing.T) {
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
	var b strings.Builder
	b.WriteString("Region\tUnits\tPrice\tRevenue\n")
	for i := 1; i <= rows; i++ {
		fmt.Fprintf(&b, "R%d\t%d\t%d\t=B{r}*C{r}\n", i%7, i%50, 1+i%13)
	}
	base := filepath.Join(dir, "base.xlsx")
	if out, code := run(b.String(), "write", base, "Data!A1"); code != 0 {
		t.Fatalf("write: %s", out)
	}
	if out, code := run("", "set", base, "Summary!A1", "Revenue", "Summary!B1", "=SUM(Data!D:D)", "Summary!B2", "=1/Data!B2"); code != 0 {
		t.Fatalf("set: %s", out)
	}
	data, err := os.ReadFile(base)
	if err != nil {
		t.Fatal(err)
	}
	var batch []string
	for i := 0; i < 50; i++ {
		switch i % 5 {
		case 0:
			batch = append(batch, fmt.Sprintf("set Data!B%d 0", 2+i*997))
		case 1:
			batch = append(batch, fmt.Sprintf("set Data!E%d '=1/Data!B%d'", 2+i, 2+(i-1)*997))
		case 2:
			batch = append(batch, fmt.Sprintf("set Data!F%d '=F%d+1'", 2+i, 2+i))
		case 3:
			batch = append(batch, fmt.Sprintf("rows insert Data!%d", 10+i))
		default:
			batch = append(batch, fmt.Sprintf("rows delete Data!%d", 10+i))
		}
	}
	for _, editor := range []string{"always", "never"} {
		t.Setenv("SHEETS_STREAM", editor)
		for _, c := range []struct {
			name, stdin string
			args        []string
		}{
			{"set", "", []string{"set", "", "Data!B5", "7"}},
			{"rows insert", "", []string{"rows", "insert", "", "Data!10"}},
			{"sort", "", []string{"sort", "", "Data", "Units"}},
			{"do (50 edits)", strings.Join(batch, "\n"), []string{"do", ""}},
		} {
			if c.name == "do (50 edits)" && editor == "never" {
				continue // each step loads the whole workbook: four minutes
			}
			path := filepath.Join(dir, "book.xlsx")
			if err := os.WriteFile(path, data, 0o644); err != nil {
				t.Fatal(err)
			}
			args := append([]string{}, c.args...)
			for i, a := range args {
				if a == "" {
					args[i] = path
				}
			}
			start := time.Now()
			out, code := run(c.stdin, args...)
			took := time.Since(start)
			if code != 0 {
				t.Fatalf("%s (%s): exit %d\n%s", c.name, editor, code, out)
			}
			lines := reportLines(out)
			t.Logf("%s, SHEETS_STREAM=%s: %s, %d lines: %s", c.name, editor, took.Round(time.Millisecond), len(lines), strings.Join(lines, " | "))
			if len(lines) == 0 || len(lines) >= 15 {
				t.Errorf("%s (%s): a report of %d lines:\n%s", c.name, editor, len(lines), out)
			}
		}
	}
}

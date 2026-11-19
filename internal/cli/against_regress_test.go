package cli

import (
	"path/filepath"
	"strings"
	"testing"
)

// What report-causes' last audit found (f5.1 to f5.3): a row edited in place is
// no row removed, a structure change that moved a value is named beside the
// input nearest it, and a total weighted by position moves with a sort.

// regressBook writes a table into Data of a new workbook from tab-separated
// lines, then runs sheets commands on it (as edited does).
func regressBook(t *testing.T, dir, name, tsv string, cmds ...[]string) string {
	t.Helper()
	path := filepath.Join(dir, name)
	if out, code := runCLI(tsv, "write", path, "Data!A1"); code != 0 {
		t.Fatalf("write: exit %d: %s", code, out)
	}
	return edited(t, path, path, cmds...)
}

// A structure change that moved a value is named with the changed input
// nearest it, first (f5.2): a row inserted and typed in, and an input in the
// same column changed, are both the cause of the totals over the column and of
// a MAX only the row's value can make; a row holding less than the MAX is no
// cause of it. Rows removed and an input that reach a value at the same
// distance, and a column added and typed in, are named alike.
func TestAgainstRegressAddedWithInput(t *testing.T) {
	dir, before := causesFixture(t)
	both := "rows added: Data!4, input Data!B2"
	for _, c := range []struct {
		name  string
		value string // typed in the row added
		b2    string // Data!B2 set to (5 before)
		want  map[string]string
	}{
		{"a row holding the new MAX", "10", "6", map[string]string{"Summary!B6": both, "Summary!B2": both, "Data!B8": both,
			"Summary!B7": "input Data!B2", "Data!D2": "input Data!B2", "Summary!B3": "input Data!B2"}},
		{"a row holding less than the MAX", "1", "10", map[string]string{"Summary!B6": "input Data!B2", "Summary!B2": both, "Data!B8": both}},
	} {
		after := edited(t, before, filepath.Join(dir, "a.xlsx"), []string{"rows", "insert", "Data!4"},
			[]string{"set", "Data!A4", "Zed", "Data!B4", c.value}, []string{"set", "Data!B2", c.b2})
		r, _ := against(t, after, before) // the exit code counts the file's own warnings too
		m := causes(r)
		for cell, want := range c.want {
			if m[cell] != want {
				t.Errorf("%s: %s: cause %q, want %q", c.name, cell, m[cell], want)
			}
		}
		if len(r.Unexplained) != 0 || r.Summary.Inputs != 1 || r.Summary.Flagged != 0 {
			t.Errorf("%s: %+v, unexplained %+v", c.name, r.Summary, r.Unexplained)
		}
	}
	// A row deleted and the tax rate changed: the tax and the net name both.
	after := edited(t, before, filepath.Join(dir, "a.xlsx"), []string{"rows", "delete", "Data!3"}, []string{"set", "Summary!B1", "0.3"})
	r, code := against(t, after, before)
	m := causes(r)
	for _, cell := range []string{"Summary!B4", "Summary!B5"} {
		if !strings.HasPrefix(m[cell], "rows removed: Data!3 in ") || !strings.HasSuffix(m[cell], ", input Summary!B1") {
			t.Errorf("a row removed and the rate set: %s: cause %q", cell, m[cell])
		}
	}
	if code != 0 || len(r.Unexplained) != 0 {
		t.Errorf("a row removed and the rate set: exit %d, unexplained %+v", code, r.Unexplained)
	}
	// A column inserted inside what a total reads, a value typed in it, and an
	// input in the total's row changed.
	wide := edited(t, before, filepath.Join(dir, "w.xlsx"), []string{"set", "Summary!C1", "=SUM(Data!B2:D2)"}, []string{"check", "--fix"})
	after = edited(t, wide, filepath.Join(dir, "a.xlsx"), []string{"cols", "insert", "Data!C"}, []string{"set", "Data!C2", "100"}, []string{"set", "Data!B2", "6"})
	r, code = against(t, after, wide)
	if m := causes(r); m["Summary!C1"] != "columns added: Data!C, input Data!B2" || m["Summary!B2"] != "input Data!B2" {
		t.Errorf("a column added and an input set: Summary!C1 %q, Summary!B2 %q", m["Summary!C1"], m["Summary!B2"])
	}
	if code != 0 || len(r.Unexplained) != 0 {
		t.Errorf("a column added and an input set: exit %d, unexplained %+v", code, r.Unexplained)
	}
}

// A total that weighs the rows by position, with an array constant or a call
// beside the range in SUMPRODUCT, moves with a sort: the rows moved are its
// cause, not unexplained (f5.3). One over the range alone, or the range twice,
// did not move.
func TestAgainstRegressSortWeighted(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("SHEETS_CACHE_DIR", filepath.Join(dir, ".cache"))
	before := regressBook(t, dir, "b.xlsx", "Region\tUnits\nNorth\t5\nSouth\t7\nEast\t3\nWest\t9\nMid\t4\n",
		[]string{"set", "Summary!A1", "=SUMPRODUCT(Data!B2:B6,{1;2;3;4;5})", "Summary!A2", "=SUMPRODUCT({1;2;3;4;5},Data!B2:B6)",
			"Summary!A3", "=SUMPRODUCT(Data!B2:B6,ROW(Data!B2:B6))", "Summary!A4", "=SUMPRODUCT(Data!B2:B6,Data!B2:B6)",
			"Summary!A5", "=SUM(Data!B2:B6)"},
		[]string{"check", "--fix"})
	after := edited(t, before, filepath.Join(dir, "a.xlsx"), []string{"sort", "Data!A1:B6", "Units:desc"})
	r, code := against(t, after, before)
	m := causes(r)
	for _, c := range []string{"Summary!A1", "Summary!A2", "Summary!A3"} {
		if !strings.HasPrefix(m[c], "rows moved: Data!") {
			t.Errorf("%s: cause %q, want rows moved", c, m[c])
		}
	}
	for _, c := range []string{"Summary!A4", "Summary!A5"} {
		if cause, ok := m[c]; ok {
			t.Errorf("%s (a total over the rows as a set) moved: %q", c, cause)
		}
	}
	if code != 0 || len(r.Unexplained) != 0 {
		t.Errorf("exit %d, unexplained %+v", code, r.Unexplained)
	}
}

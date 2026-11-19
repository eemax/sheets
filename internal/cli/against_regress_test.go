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

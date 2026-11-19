package cli

import (
	"os"
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

// regressTable is a table that ends its sheet, with no formula column.
const regressTable = "Region\tUnits\tPrice\nNorth\t5\t30\nSouth\t7\t20\nEast\t3\t50\nWest\t9\t10\nMid\t4\t25\n"

// changedOf maps each cell changed to its value before and now.
func changedOf(r causesResult) map[string][2]string {
	out := map[string][2]string{}
	for _, c := range r.Changed {
		out[c.Cell] = [2]string{c.Before, c.After}
	}
	return out
}

// The last row of a table edited in place, two of its cells cleared and the
// third set, with no row deleted, is its three cells changed, the two cleared
// named, no structure change and nothing flagged (f5.1): the formulas reading
// it, as they were, show that no row was removed. The same in a CSV file, where
// nothing reads the row: values changed, as their values tell.
func TestAgainstRegressLastRowEdited(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("SHEETS_CACHE_DIR", filepath.Join(dir, ".cache"))
	before := regressBook(t, dir, "b.xlsx", regressTable,
		[]string{"set", "Summary!A1", "=SUM(Data!B2:B6)", "Summary!A2", "=COUNTA(Data!A2:A6)", "Summary!A3", "=Data!A6"})
	after := edited(t, before, filepath.Join(dir, "a.xlsx"), []string{"clear", "Data!A6:C6"}, []string{"set", "Data!B6", "99"})
	r, code := against(t, after, before)
	want := map[string][2]string{"Data!A6": {`"Mid"`, ""}, "Data!B6": {"4", "99"}, "Data!C6": {"25", ""}}
	if s := r.Summary; s.Inputs != 3 || s.Formulas != 0 || s.Flagged != 0 || len(r.Structure) != 0 || code != 0 {
		t.Errorf("xlsx: exit %d, %+v, structure %+v, changed %+v", code, s, r.Structure, r.Changed)
	}
	if got := changedOf(r); len(got) != 3 || got["Data!A6"] != want["Data!A6"] || got["Data!B6"] != want["Data!B6"] || got["Data!C6"] != want["Data!C6"] {
		t.Errorf("xlsx: changed %+v", r.Changed)
	}
	if m := causes(r); m["Summary!A1"] != "input Data!B6" || m["Summary!A3"] != "input Data!A6" {
		t.Errorf("xlsx: causes %q", m)
	}
	// A CSV file whose last row Mid,4,25 became ,99,.
	b := filepath.Join(dir, "b.csv")
	a := filepath.Join(dir, "a.csv")
	writeText(t, b, "Region,Units,Price\nNorth,5,30\nSouth,7,20\nEast,3,50\nWest,9,10\nMid,4,25\n")
	writeText(t, a, "Region,Units,Price\nNorth,5,30\nSouth,7,20\nEast,3,50\nWest,9,10\n,99,\n")
	r, code = against(t, a, b)
	if s := r.Summary; s.Inputs != 3 || len(r.Structure) != 0 || code != 0 {
		t.Errorf("csv: exit %d, %+v, structure %+v, changed %+v", code, s, r.Structure, r.Changed)
	}
	if got := changedOf(r); len(got) != 3 || got["a!A6"] != want["Data!A6"] || got["a!B6"] != want["Data!B6"] || got["a!C6"] != want["Data!C6"] {
		t.Errorf("csv: changed %+v", r.Changed)
	}
	// The fuzz pool's openpyxl file, whose last row a defined name, an Excel
	// table and a total read, as they were: the row cleared, its formulas too,
	// and a cell set. Each of its nine cells changed (a formula of a column
	// cleared is flagged as such), no row removed, no name not rewritten.
	pool := filepath.Join(dir, "pool.xlsx")
	if err := os.WriteFile(pool, mustRead(t, filepath.Join("..", "fuzz", "testdata", "pool", "openpyxl-02.xlsx")), 0o644); err != nil {
		t.Fatal(err)
	}
	after = edited(t, pool, filepath.Join(dir, "a.xlsx"), []string{"clear", "Data!A7:P7"}, []string{"set", "Data!B7", "77"})
	r, _ = against(t, after, pool) // the exit code counts the file's own errors too
	if s := r.Summary; s.Inputs != 6 || s.Formulas != 3 || len(r.Structure) != 0 || changedOf(r)["Data!B7"] != [2]string{"40", "77"} {
		t.Errorf("pool file: %+v, structure %+v, changed %+v", s, r.Structure, r.Changed)
	}
}

// writeText writes a file for a test.
func writeText(t *testing.T, path, text string) {
	t.Helper()
	if err := os.WriteFile(path, []byte(text), 0o644); err != nil {
		t.Fatal(err)
	}
}

// The last row (or two) of a table deleted and a cell set in the row now at that
// index is still one row removed (two) and the cell typed in from empty, nothing
// flagged (report-causes f3.1), when what reads the rows was rewritten for the
// removal: a total over them, a defined name, an Excel table (the fuzz pool's
// openpyxl file has all three). Where nothing reads them, the rows deleted and
// the cells set are the same file as the row's cells cleared and typed over, and
// the values changed are what it tells: in a sheet nothing reads, in a CSV file,
// and after openpyxl's delete_rows, which rewrites no formula.
func TestAgainstRegressDeleteSetLast(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("SHEETS_CACHE_DIR", filepath.Join(dir, ".cache"))
	plain := regressBook(t, dir, "plain.xlsx", regressTable)
	read := edited(t, plain, filepath.Join(dir, "b.xlsx"), []string{"set", "Summary!A1", "=SUM(Data!B2:B6)"})
	pool := filepath.Join(dir, "pool.xlsx")
	if err := os.WriteFile(pool, mustRead(t, filepath.Join("..", "fuzz", "testdata", "pool", "openpyxl-02.xlsx")), 0o644); err != nil {
		t.Fatal(err)
	}
	for _, c := range []struct {
		name, before, rows, cell string
		removed                  string // the structure told: "" for none
		inputs                   int
	}{
		{"a total reads it", read, "6", "Data!B6", "row removed Data!6", 1},
		{"a total reads them", read, "5:6", "Data!B5", "rows removed Data!5:6", 1},
		{"a name, a table and totals read it", pool, "7", "Data!C7", "row removed Data!7", 1},
		{"nothing reads it", plain, "6", "Data!B6", "", 3},
	} {
		after := edited(t, c.before, filepath.Join(dir, "a.xlsx"), []string{"rows", "delete", "Data!" + c.rows}, []string{"set", c.cell, "77"})
		r, _ := against(t, after, c.before) // the exit code counts the pool file's own errors too
		var lines []string
		for _, e := range r.Structure {
			lines = append(lines, e.Change+" "+strings.SplitN(e.Where, " in ", 2)[0])
		}
		if s := r.Summary; s.Inputs != c.inputs || s.Formulas != 0 || s.Flagged != 0 || s.Unexplained != 0 || s.FeaturesLost != 0 ||
			c.removed == "" && len(lines) != 0 || c.removed != "" && (len(lines) != 1 || lines[0] != c.removed) {
			t.Errorf("%s: %+v, structure %q, changed %+v", c.name, s, lines, r.Changed)
		}
		if got := changedOf(r)[c.cell]; got[1] != "77" || c.inputs == 1 && got[0] != "" {
			t.Errorf("%s: %s %q", c.name, c.cell, got)
		}
	}
	// A CSV file whose last row was deleted and a value typed in its place.
	b, a := filepath.Join(dir, "b.csv"), filepath.Join(dir, "a.csv")
	writeText(t, b, "Region,Units,Price\nNorth,5,30\nSouth,7,20\nEast,3,50\nWest,9,10\nMid,4,25\n")
	writeText(t, a, "Region,Units,Price\nNorth,5,30\nSouth,7,20\nEast,3,50\nWest,9,10\n,77,\n")
	if r, code := against(t, a, b); r.Summary.Inputs != 3 || len(r.Structure) != 0 || code != 0 {
		t.Errorf("csv: exit %d, %+v, structure %+v", code, r.Summary, r.Structure)
	}
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

// A structure change is named beside the changed input whichever way the input
// moved the value by more, and beside a structure change of the other kind: the
// report's sum adds up (f0.1 of the last audit). Rows removed that took 7 off a
// total that an input raised by 45 are named with it, as with an input raised
// by 1; rows inserted and typed in that added 3 to a total an input lowered by
// 5, likewise; a row deleted and another inserted and typed in name both for the
// total they both moved, each the ones it alone moved. A MAX is named for the
// rows only when they held what it is (or was).
func TestAgainstRegressStructureAndInput(t *testing.T) {
	dir, before := causesFixture(t)
	for _, b2 := range []string{"6", "50"} {
		after := edited(t, before, filepath.Join(dir, "a.xlsx"), []string{"rows", "delete", "Data!3"}, []string{"set", "Data!B2", b2})
		r, code := against(t, after, before)
		m := causes(r)
		want := "rows removed: Data!3 in " + before + ", input Data!B2"
		for _, cell := range []string{"Summary!B2", "Data!B6"} {
			if m[cell] != want {
				t.Errorf("a row removed, B2 set to %s: %s: cause %q, want %q", b2, cell, m[cell], want)
			}
		}
		if m["Summary!B6"] != "input Data!B2" && b2 == "50" {
			t.Errorf("a row removed, B2 set to 50: MAX: cause %q", m["Summary!B6"])
		}
		if code != 0 || len(r.Unexplained) != 0 {
			t.Errorf("a row removed, B2 set to %s: exit %d, unexplained %+v", b2, code, r.Unexplained)
		}
	}
	// A row inserted and typed in, an input lowered by more than the row added.
	after := edited(t, before, filepath.Join(dir, "a.xlsx"), []string{"rows", "insert", "Data!4"},
		[]string{"set", "Data!A4", "Zed", "Data!B4", "3"}, []string{"set", "Data!B2", "0"})
	r, code := against(t, after, before)
	m := causes(r)
	for _, cell := range []string{"Summary!B2", "Data!B8"} {
		if m[cell] != "rows added: Data!4, input Data!B2" {
			t.Errorf("a row added, B2 set to 0: %s: cause %q", cell, m[cell])
		}
	}
	if _, moved := m["Summary!B6"]; moved || code != 0 || len(r.Unexplained) != 0 {
		t.Errorf("a row added, B2 set to 0: the MAX (9 still) moved, cause %q, exit %d, unexplained %+v", m["Summary!B6"], code, r.Unexplained)
	}
	// A row deleted, another inserted and typed in, no input changed.
	after = edited(t, before, filepath.Join(dir, "a.xlsx"), []string{"rows", "delete", "Data!3"}, []string{"rows", "insert", "Data!5"},
		[]string{"set", "Data!A5", "Zed", "Data!B5", "10", "Data!C5", "1"})
	r, code = against(t, after, before)
	m = causes(r)
	both := "rows removed: Data!3 in " + before + ", rows added: Data!5"
	for _, cell := range []string{"Summary!B2", "Data!B7"} {
		if m[cell] != both {
			t.Errorf("a row removed and one added: %s: cause %q, want %q", cell, m[cell], both)
		}
	}
	if want := "rows removed: Data!3 in " + before; m["Summary!B3"] != want || m["Data!D7"] != want {
		t.Errorf("a row removed and one added: revenue: causes %q, %q, want %q", m["Summary!B3"], m["Data!D7"], want)
	}
	if m["Summary!B6"] != "rows added: Data!5" || code != 0 || len(r.Unexplained) != 0 {
		t.Errorf("a row removed and one added: MAX cause %q, exit %d, unexplained %+v", m["Summary!B6"], code, r.Unexplained)
	}
}

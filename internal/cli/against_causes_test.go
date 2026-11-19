package cli

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"testing"
	"time"

	"github.com/eemax/sheets/internal/xlsx"
)

// What check --against gives as the cause of a value that moved: the cause
// that can move it. Rows a sort moved are no cause of a total over all of
// them; rows removed are no cause of a MAX that rose; a value nothing listed
// can account for stays unexplained, and what follows it follows it.

// causesFixture copies against_before.xlsx into a test's folder: Data!A1:D7
// (Region, Units, Price, Revenue = Units × Price, a total row), Summary
// (B2 =SUM(Data!B2:B6), B3 =SUM(Data!D2:D6), B4 =B3*TaxRate, B5 =B3-B4,
// B6 =MAX(Data!B2:B6), B7 =Data!B2) and the name TaxRate = Summary!$B$1.
func causesFixture(t *testing.T) (dir, before string) {
	t.Helper()
	dir = t.TempDir()
	t.Setenv("SHEETS_CACHE_DIR", filepath.Join(dir, ".cache"))
	before = filepath.Join(dir, "b.xlsx")
	if err := os.WriteFile(before, mustRead(t, filepath.Join("..", "xlsx", "testdata", "against_before.xlsx")), 0o644); err != nil {
		t.Fatal(err)
	}
	return dir, before
}

// edited copies a workbook and runs sheets commands on the copy: each
// command's words, the file going in after the first word, or after the
// second for rows, cols, sheet and chart.
func edited(t *testing.T, src, dst string, cmds ...[]string) string {
	t.Helper()
	if err := os.WriteFile(dst, mustRead(t, src), 0o644); err != nil {
		t.Fatal(err)
	}
	for _, c := range cmds {
		n := 1
		switch c[0] {
		case "rows", "cols", "sheet", "chart":
			n = 2
		}
		args := append(append(append([]string{}, c[:n]...), dst), c[n:]...)
		if out, code := runCLI("", args...); code != 0 {
			t.Fatalf("%v: exit %d: %s", c, code, out)
		}
	}
	return dst
}

// causesResult is what the tests read of check --against's JSON.
type causesResult struct {
	Summary         againstSummary   `json:"summary"`
	Changed         []changedCell    `json:"changed"`
	Moved           []movedValue     `json:"moved"`
	Unexplained     []movedValue     `json:"unexplained"`
	FeaturesLost    []lostFeature    `json:"features_lost"`
	FeaturesChanged []changedFeature `json:"features_changed"`
	Unknown         []unknownValue   `json:"unknown"`
	Structure       []structureEntry `json:"structure"`
}

// against runs check --against and reads its JSON.
func against(t *testing.T, after, before string) (causesResult, int) {
	t.Helper()
	out, code := runCLI("", "--json", "check", after, "--against", before)
	var r causesResult
	if err := json.Unmarshal([]byte(out), &r); err != nil {
		t.Fatalf("%v: %s", err, out)
	}
	return r, code
}

// causes maps each value moved to its cause.
func causes(r causesResult) map[string]string {
	out := map[string]string{}
	for _, m := range r.Moved {
		out[m.Cell] = m.Cause
	}
	return out
}

// A sort and one typed fix: the totals over the sorted rows moved for the
// input, not for the sort (change-report f5.1); a cell read by position
// moved for both.
func TestAgainstCausesSortTotal(t *testing.T) {
	dir, before := causesFixture(t)
	after := edited(t, before, filepath.Join(dir, "a.xlsx"), []string{"sort", "Data!A1:D6", "Units:desc"}, []string{"set", "Data!B2", "11"})
	r, code := against(t, after, before)
	m := causes(r)
	for _, c := range []string{"Data!D7", "Data!B7", "Summary!B2", "Summary!B3", "Summary!B4", "Summary!B5", "Summary!B6"} {
		if m[c] != "input Data!B2" {
			t.Errorf("%s: cause %q, want input Data!B2", c, m[c])
		}
	}
	if want := "input Data!B2, rows moved: Data!2 (were Data!5)"; m["Summary!B7"] != want {
		t.Errorf("Summary!B7 (=Data!B2): cause %q, want %q", m["Summary!B7"], want)
	}
	if code != 0 || len(r.Unexplained) != 0 {
		t.Errorf("exit %d, unexplained %+v", code, r.Unexplained)
	}
}

// A sort alone: the totals over the sorted rows did not move; a cell read
// by position and a lookup did, for the sort.
func TestAgainstCausesSortLookup(t *testing.T) {
	dir, before := causesFixture(t)
	base := edited(t, before, filepath.Join(dir, "base.xlsx"), []string{"set", "Summary!C1", `=VLOOKUP("North",Data!A2:D6,2,FALSE)`,
		"Summary!C2", "=INDEX(Data!B2:B6,1)", "Summary!C3", "=SUMPRODUCT(Data!B2:B6,Data!C2:C6)", "Summary!C4", "=COUNTIF(Data!B2:B6,\">4\")",
		"Summary!C5", "=SUM(Data!B2:B3)", "Summary!C6", "=SUMIF(Data!A2:A6,\"North\",Data!B2:B6)"})
	after := edited(t, base, filepath.Join(dir, "a.xlsx"), []string{"sort", "Data!A1:D6", "Units:desc"})
	r, code := against(t, after, base)
	m := causes(r)
	for _, c := range []string{"Summary!B7", "Summary!C2", "Summary!C5"} {
		if !strings.HasPrefix(m[c], "rows moved: ") {
			t.Errorf("%s (read by position): cause %q", c, m[c])
		}
	}
	for _, c := range []string{"Summary!B2", "Summary!B3", "Summary!B6", "Summary!C1", "Summary!C3", "Summary!C4", "Summary!C6", "Data!B7", "Data!D7"} {
		if _, ok := m[c]; ok {
			t.Errorf("%s moved: %q", c, m[c])
		}
	}
	if code != 0 || len(r.Unexplained)+len(r.Unknown) != 0 {
		t.Errorf("exit %d, unexplained %+v, unknown %+v", code, r.Unexplained, r.Unknown)
	}
}

// An Excel table's column summed, its rows sorted and one changed: the
// total moved for the input.
func TestAgainstCausesSortTable(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("SHEETS_CACHE_DIR", filepath.Join(dir, ".cache"))
	before := filepath.Join(dir, "b.xlsx")
	if err := os.WriteFile(before, mustRead(t, filepath.Join("..", "xlsx", "testdata", "against_table.xlsx")), 0o644); err != nil {
		t.Fatal(err)
	}
	after := edited(t, before, filepath.Join(dir, "a.xlsx"), []string{"sort", "Data!A1:B3", "Units:desc"}, []string{"set", "Data!B3", "6"})
	r, code := against(t, after, before)
	if got := causes(r)["Sum!A1"]; got != "input Data!B3" {
		t.Errorf("Sum!A1 (=SUM(Sales[Units])): cause %q, want input Data!B3 (%+v)", got, r.Moved)
	}
	if code != 0 {
		t.Errorf("exit %d", code)
	}
}

// 300 rows sorted, one edited: the total moved for the edit.
func TestAgainstCausesSortMany(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("SHEETS_CACHE_DIR", filepath.Join(dir, ".cache"))
	var b strings.Builder
	b.WriteString("Item\tUnits\tPrice\tRevenue\n")
	for i := 1; i <= 300; i++ {
		fmt.Fprintf(&b, "I%d\t%d\t%d\t=B{r}*C{r}\n", i, (i*37)%101, 1+i%9)
	}
	before := filepath.Join(dir, "b.xlsx")
	if out, code := runCLI(b.String(), "write", before, "Data!A1"); code != 0 {
		t.Fatal(out)
	}
	if out, code := runCLI("", "set", before, "Data!B302", "=SUM(B2:B301)", "Data!D302", "=SUM(D2:D301)"); code != 0 {
		t.Fatal(out)
	}
	after := edited(t, before, filepath.Join(dir, "a.xlsx"), []string{"sort", "Data!A1:D301", "Units:desc"}, []string{"set", "Data!B2", "500"})
	r, code := against(t, after, before)
	m := causes(r)
	for _, c := range []string{"Data!B302", "Data!D302"} {
		if m[c] != "input Data!B2" {
			t.Errorf("%s: cause %q", c, m[c])
		}
	}
	if code != 0 || r.Summary.MovedRows == 0 {
		t.Errorf("exit %d, %+v", code, r.Summary)
	}
}

// storeResult rewrites the result Summary stores for a cell (results trusted:
// no full calculation on load).
func storeResult(t *testing.T, src, dst string, cells map[string]string) {
	t.Helper()
	rewriteParts(t, src, dst, func(name string, b []byte) []byte {
		switch name {
		case "xl/workbook.xml":
			return bytes.Replace(b, []byte(` fullCalcOnLoad="true"`), nil, 1)
		case "xl/worksheets/sheet2.xml":
			for cell, v := range cells {
				re := regexp.MustCompile(`(<c r="` + cell + `"[^>]*><f>[^<]*</f>)<v>[^<]*</v>`)
				was := b
				if b = re.ReplaceAll(b, []byte("${1}<v>"+v+"</v>")); bytes.Equal(b, was) {
					t.Fatalf("no result of %s", cell)
				}
			}
		}
		return b
	})
}

// A total whose stored result changed with nothing in it changed, the rows
// it sums sorted: unexplained, as without the sort (f5.2); the cells that
// read it and held are told with it, not unknown; one that followed it
// moved with it as its cause.
func TestAgainstCausesSortUnexplained(t *testing.T) {
	dir, before := causesFixture(t)
	sorted := edited(t, before, filepath.Join(dir, "s.xlsx"), []string{"sort", "Data!A1:D6", "Units:desc"})
	for _, src := range []string{before, sorted} {
		after := filepath.Join(dir, "a.xlsx")
		storeResult(t, src, after, map[string]string{"B3": "999"})
		r, code := against(t, after, before)
		if len(r.Unexplained) != 1 || r.Unexplained[0].Cell != "Summary!B3" || code != 1 {
			t.Fatalf("%s: exit %d, unexplained %+v, moved %+v", filepath.Base(src), code, r.Unexplained, r.Moved)
		}
		if got := strings.Join(r.Unexplained[0].Held, " "); got != "Summary!B4 Summary!B5" {
			t.Errorf("%s: held %q", filepath.Base(src), got)
		}
		for _, u := range r.Unknown {
			t.Errorf("%s: unknown %+v", filepath.Base(src), u)
		}
		out, _ := runCLI("", "check", after, "--against", before)
		if !strings.Contains(out, "Summary!B3  630 → 999  read by Summary!B4, Summary!B5, which hold what they held") {
			t.Errorf("%s: text:\n%s", filepath.Base(src), out)
		}

		// The tax followed the total: moved, for it.
		storeResult(t, src, after, map[string]string{"B3": "999", "B4": "249.75"})
		r, code = against(t, after, before)
		if len(r.Unexplained) != 1 || r.Unexplained[0].Cell != "Summary!B3" || code != 1 {
			t.Fatalf("%s, followed: exit %d, unexplained %+v", filepath.Base(src), code, r.Unexplained)
		}
		if got := causes(r)["Summary!B4"]; got != "Summary!B3, which moved with nothing upstream changed" {
			t.Errorf("%s, followed: Summary!B4's cause %q", filepath.Base(src), got)
		}
	}
}

// A row removed and an input changed in the same column: the MAX that rose,
// and the totals, moved for the input (f5.3); the cell that read the row
// removed moved for the removal.
func TestAgainstCausesRemovedRise(t *testing.T) {
	dir, before := causesFixture(t)
	after := edited(t, before, filepath.Join(dir, "a.xlsx"), []string{"rows", "delete", "Data!2"}, []string{"set", "Data!B3", "77"})
	r, _ := against(t, after, before)
	m := causes(r)
	for _, c := range []string{"Summary!B6", "Data!B6", "Data!D6", "Summary!B2", "Summary!B3", "Summary!B4", "Summary!B5"} {
		if m[c] != "input Data!B3" {
			t.Errorf("%s: cause %q, want input Data!B3", c, m[c])
		}
	}
	if !strings.HasPrefix(m["Summary!B7"], "rows removed: Data!2") {
		t.Errorf("Summary!B7 (=Data!B2, now #REF!): cause %q", m["Summary!B7"])
	}
	if len(r.Unexplained) != 0 {
		t.Errorf("unexplained %+v", r.Unexplained)
	}
}

// A row removed and an input lowered: the totals fell for both, and say so;
// the MAX fell for the input (removing 5 leaves 9 the most).
func TestAgainstCausesRemovedFall(t *testing.T) {
	dir, before := causesFixture(t)
	after := edited(t, before, filepath.Join(dir, "a.xlsx"), []string{"rows", "delete", "Data!2"}, []string{"set", "Data!B4", "1"})
	r, _ := against(t, after, before)
	m := causes(r)
	for _, c := range []string{"Data!B6", "Summary!B2"} {
		if m[c] != "rows removed: Data!2 in "+before+", input Data!B4" {
			t.Errorf("%s: cause %q", c, m[c])
		}
	}
	if m["Summary!B6"] != "input Data!B4" {
		t.Errorf("Summary!B6 (MAX 9 → 7): cause %q", m["Summary!B6"])
	}
	// The row alone: what it moved is its doing.
	only := edited(t, before, filepath.Join(dir, "o.xlsx"), []string{"rows", "delete", "Data!2"})
	r, code := against(t, only, before)
	for c, cause := range causes(r) {
		if !strings.HasPrefix(cause, "rows removed: Data!2") {
			t.Errorf("row removed alone: %s: cause %q", c, cause)
		}
	}
	if len(r.Unexplained) != 0 || code != 1 { // exit 1: Summary!B7 reads #REF!, an error
		t.Errorf("row removed alone: exit %d, unexplained %+v", code, r.Unexplained)
	}
}

// stored rewrites the result a workbook stores for a formula's cell on the
// second sheet (Summary), and leaves the workbook not to be calculated on
// load, as a file another tool wrote and Excel calculated.
func stored(t *testing.T, src, dst, cell, value string) {
	t.Helper()
	re := regexp.MustCompile(`(<c r="` + cell + `"[^>]*><f>[^<]*</f>)<v>[^<]*</v>`)
	rewriteParts(t, src, dst, func(name string, b []byte) []byte {
		switch name {
		case "xl/workbook.xml":
			return regexp.MustCompile(` fullCalcOnLoad="(true|1)"`).ReplaceAll(b, nil)
		case "xl/worksheets/sheet2.xml":
			was := b
			if b = re.ReplaceAll(b, []byte("${1}<v>"+value+"</v>")); bytes.Equal(b, was) {
				t.Fatalf("no result stored for %s in %s", cell, b)
			}
		}
		return b
	})
}

// A row removed whose removal cannot have moved a value the way it moved
// (a MAX that rose, a SUM that rose though the row held 4): the value stored
// by another tool is unexplained, with exit 1, whether the formula was
// rewritten for the row (Excel's shrinking of its range) or an input in its
// column changed too (f0.1). What the row can move, it moves.
func TestAgainstCausesRemovedWrongWay(t *testing.T) {
	dir, before := causesFixture(t)
	deleted := edited(t, before, filepath.Join(dir, "d.xlsx"), []string{"rows", "delete", "Data!6"})
	unrelated := edited(t, before, filepath.Join(dir, "u.xlsx"), []string{"rows", "delete", "Data!6"}, []string{"set", "Data!C2", "31"})
	for _, c := range []struct {
		name, from, cell, value string
	}{
		{"a MAX that rose", deleted, "B6", "50"},
		{"a SUM that rose", deleted, "B2", "999"},
		{"a SUM that rose, an unrelated input changed", unrelated, "B2", "999"},
	} {
		after := filepath.Join(dir, "a.xlsx")
		stored(t, c.from, after, c.cell, c.value)
		r, code := against(t, after, before)
		got := false
		for _, u := range r.Unexplained {
			got = got || u.Cell == "Summary!"+c.cell
		}
		if m := causes(r)["Summary!"+c.cell]; !got || m != "" || code != 1 {
			t.Errorf("%s: exit %d, unexplained %+v, cause %q", c.name, code, r.Unexplained, m)
		}
	}
	// The way the removal can move them: the row held 4, so the total fell.
	r, _ := against(t, deleted, before)
	if m := causes(r)["Summary!B2"]; !strings.HasPrefix(m, "rows removed: Data!6") {
		t.Errorf("the total that fell: cause %q", m)
	}
}

// A row added that holds nothing cannot move a total over it: a total stored
// as another number is unexplained (f0.3). Rows added that hold a number
// move it as they can, with the cell they were given named.
func TestAgainstCausesAddedEmptyRow(t *testing.T) {
	dir, before := causesFixture(t)
	empty := edited(t, before, filepath.Join(dir, "i.xlsx"), []string{"rows", "insert", "Data!4"})
	r, code := against(t, empty, before)
	if len(r.Moved)+len(r.Unexplained) != 0 || code != 0 {
		t.Errorf("an empty row alone: exit %d, moved %+v, unexplained %+v", code, r.Moved, r.Unexplained)
	}
	for _, c := range []struct{ cell, value string }{{"B2", "999"}, {"B6", "50"}} {
		after := filepath.Join(dir, "a.xlsx")
		stored(t, empty, after, c.cell, c.value)
		r, code := against(t, after, before)
		got := false
		for _, u := range r.Unexplained {
			got = got || u.Cell == "Summary!"+c.cell
		}
		if m := causes(r)["Summary!"+c.cell]; !got || m != "" || code != 1 {
			t.Errorf("%s stored as %s: exit %d, unexplained %+v, moved %+v", c.cell, c.value, code, r.Unexplained, r.Moved)
		}
	}
	// A number in the row moves a total up and a MAX that it tops: the row
	// added is their cause.
	filled := edited(t, before, filepath.Join(dir, "f.xlsx"), []string{"rows", "insert", "Data!4"}, []string{"set", "Data!B4", "11"})
	r, code = against(t, filled, before)
	for _, c := range []string{"Summary!B2", "Summary!B6"} {
		if m := causes(r)[c]; m != "rows added: Data!4" {
			t.Errorf("%s with a number added: cause %q (unexplained %+v)", c, m, r.Unexplained)
		}
	}
	if len(r.Unexplained) != 0 || code != 0 {
		t.Errorf("a number added: exit %d, unexplained %+v", code, r.Unexplained)
	}
	// One that is lower than the MAX cannot have moved it, nor can it lower
	// the total.
	low := edited(t, before, filepath.Join(dir, "l.xlsx"), []string{"rows", "insert", "Data!4"}, []string{"set", "Data!B4", "2"})
	lowStored := filepath.Join(dir, "ls.xlsx")
	stored(t, low, lowStored, "B6", "50")
	if r, code := against(t, lowStored, before); code != 1 || len(r.Unexplained) != 1 || r.Unexplained[0].Cell != "Summary!B6" {
		t.Errorf("a MAX that rose for a lower number: exit %d, unexplained %+v, moved %+v", code, r.Unexplained, r.Moved)
	}
}

// rewriteParts copies a workbook, its parts rewritten.
func rewriteParts(t *testing.T, src, dst string, rewrite func(name string, b []byte) []byte) {
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
		w, err := zw.Create(f.Name)
		if err != nil {
			t.Fatal(err)
		}
		w.Write(rewrite(f.Name, data))
	}
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(dst, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
}

// A row deleted and a cell of the rows below set, on openpyxl's pool file
// with two rows alike (f5.5): one row removed and one input changed, Excel's
// own rewrite of a range and of a name no change, and the chart still there
// not lost.
func TestAgainstCausesDeleteSetAligned(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("SHEETS_CACHE_DIR", filepath.Join(dir, ".cache"))
	before := filepath.Join(dir, "b.xlsx")
	if err := os.WriteFile(before, mustRead(t, filepath.Join("..", "fuzz", "testdata", "pool", "openpyxl-02.xlsx")), 0o644); err != nil {
		t.Fatal(err)
	}
	for _, c := range []string{"B3", "C3"} {
		after := edited(t, before, filepath.Join(dir, "a.xlsx"), []string{"rows", "delete", "Data!2"}, []string{"set", "Data!" + c, "7"})
		r, code := against(t, after, before)
		s := r.Summary
		if s.FeaturesLost != 0 || s.Inputs != 1 || s.Formulas != 0 || s.Unexplained != 0 || code != 0 {
			t.Errorf("set %s: exit %d, %+v, lost %+v, changed %+v", c, code, s, r.FeaturesLost, r.Changed)
		}
		if len(r.Structure) != 1 || r.Structure[0].Change != "row removed" || r.Structure[0].Where != "Data!2 in "+before {
			t.Errorf("set %s: structure %+v", c, r.Structure)
		}
		if len(r.Changed) == 1 && (r.Changed[0].Cell != "Data!"+c || r.Changed[0].Was != c[:1]+"4") {
			t.Errorf("set %s: changed %+v", c, r.Changed)
		}
	}
}

// A validation's bound raised on the same cells, a dropdown's items
// changed, a number's threshold: features changed, not lost, and no exit 1
// for them (f5.6), in check --against and in the command's own report
// (f4.3). A highlight's threshold another tool edited, the same.
func TestAgainstCausesFeaturesChanged(t *testing.T) {
	dir, before := causesFixture(t)
	lists := edited(t, before, filepath.Join(dir, "lists.xlsx"), []string{"validate", "Data!F2:F6", "list", "1,2"})
	for _, c := range []struct {
		name, before string
		cmd          []string
	}{
		{"bound", before, []string{"validate", "Data!B2:B6", "whole", "between", "0", "and", "5000"}},
		{"threshold", before, []string{"validate", "Data!C2:C6", "number", ">= 5"}},
		{"dropdown", lists, []string{"validate", "Data!F2:F6", "list", "1,2,3"}},
	} {
		after := filepath.Join(dir, c.name+".xlsx")
		if err := os.WriteFile(after, mustRead(t, c.before), 0o644); err != nil {
			t.Fatal(err)
		}
		out, code := runCLI("", append([]string{c.cmd[0], after}, c.cmd[1:]...)...)
		if code != 0 {
			t.Fatalf("%s: %s", c.name, out)
		}
		if !strings.Contains(out, ", 0 features lost (1 changed)") || !strings.Contains(out, "\nchanged: validation Data!") {
			t.Errorf("%s: the command's report:\n%s", c.name, out)
		}
		changedOK(t, c.name, after, c.before)
	}
	cond := filepath.Join(dir, "cond.xlsx")
	rewriteSheet(t, before, cond, func(b []byte) []byte {
		return bytes.Replace(b, []byte("<formula>140</formula>"), []byte("<formula>120</formula>"), 1)
	})
	changedOK(t, "highlight", cond, before)

	// On other cells it is another rule: the one that was is lost.
	moved := edited(t, before, filepath.Join(dir, "moved.xlsx"), []string{"validate", "Data!B2:B6", "--clear"}, []string{"validate", "Data!B3:B6", "whole", "between", "0", "and", "5000"})
	if r, code := against(t, moved, before); len(r.FeaturesLost) != 1 || code != 1 {
		t.Errorf("rule on other cells: exit %d, lost %+v, changed %+v", code, r.FeaturesLost, r.FeaturesChanged)
	}
}

// A rule another replaced on the same cells is the old one lost, not an edit
// of it (f0.2): the banding sheets fmt writes over a highlight of another
// tool, a list a move pasted a whole-number rule over, a highlight cleared
// and replaced by data bars. Each is lost, with exit 1, in check --against
// and in the command's own report; none is in the features changed.
func TestAgainstCausesFeatureReplaced(t *testing.T) {
	dir, fixture := causesFixture(t)
	lostOnly := func(name, after, before, kind, where, cmdOut string) {
		t.Helper()
		r, code := against(t, after, before)
		found := false
		for _, f := range r.FeaturesLost {
			found = found || f.Kind == kind && f.Where == where
		}
		if !found || len(r.FeaturesChanged) != 0 || r.Summary.FeaturesChanged != 0 || code != 1 {
			t.Errorf("%s: exit %d, lost %+v, changed %+v", name, code, r.FeaturesLost, r.FeaturesChanged)
		}
		if cmdOut != "" && (strings.Contains(head(t, cmdOut), "changed)") || strings.Contains(head(t, cmdOut), ", 0 features lost")) {
			t.Errorf("%s: the command's report:\n%s", name, cmdOut)
		}
	}
	// A highlight of the data rows by another tool, then sheets fmt's banding
	// over the same rows: both are formula rules.
	before := filepath.Join(dir, "highlight.xlsx")
	rewriteSheet(t, fixture, before, func(b []byte) []byte {
		return bytes.Replace(b, []byte("</conditionalFormatting>"), []byte(`</conditionalFormatting><conditionalFormatting sqref="A2:D6"><cfRule type="expression" dxfId="0" priority="2"><formula>$B2&gt;6</formula></cfRule></conditionalFormatting>`), 1)
	})
	after := filepath.Join(dir, "banded.xlsx")
	if err := os.WriteFile(after, mustRead(t, before), 0o644); err != nil {
		t.Fatal(err)
	}
	out, _ := runCLI("", "fmt", after)
	lostOnly("fmt's banding", after, before, "conditional_format", "Data!A2:D6", out)
	if !strings.Contains(out, "conditional format Data!A2:D6 (formula rule)") {
		t.Errorf("fmt's report:\n%s", out)
	}

	// A list on the cells a move pasted over.
	listed := edited(t, fixture, filepath.Join(dir, "listed.xlsx"), []string{"validate", "Data!H2:H3", "list", "a,b"})
	after = filepath.Join(dir, "pasted.xlsx")
	if err := os.WriteFile(after, mustRead(t, listed), 0o644); err != nil {
		t.Fatal(err)
	}
	out, code := runCLI("", "move", after, "Data!B2:C3", "Data!H2")
	if code != 0 || !strings.Contains(head(t, out), ", 1 feature lost") || strings.Contains(out, "changed: validation") {
		t.Errorf("move over a list: exit %d:\n%s", code, out)
	}
	lostOnly("move over a list", after, listed, "validation", "Data!H2:H3", "")

	// A highlight cleared, data bars where it was.
	bars := edited(t, fixture, filepath.Join(dir, "bars.xlsx"), []string{"cond", "Data!D2:D6", "--clear"}, []string{"cond", "Data!D2:D6", "bars"})
	lostOnly("data bars for a highlight", bars, fixture, "conditional_format", "Data!D2:D6", "")
}

// ruleEdited, on its own: the same type or a kindred one (a bound on a value),
// a formula of the same shape (what a threshold edit leaves), and a validation
// of no type, which allows any value and stands in for any. ruleText says "any
// value" of it.
func TestAgainstCausesRuleEdited(t *testing.T) {
	val := func(typ, rule string) xlsx.Feature {
		return xlsx.Feature{Kind: xlsx.FeatureValidation, Type: typ, Rule: rule}
	}
	cond := func(typ, rule string) xlsx.Feature {
		return xlsx.Feature{Kind: xlsx.FeatureCond, Type: typ, Rule: rule}
	}
	for _, c := range []struct {
		name         string
		f, g         xlsx.Feature
		same, anyway bool // edited when the same type is asked for, and when not
	}{
		{"bound", val("whole", "between|0|1000"), val("whole", "between|0|5000"), true, true},
		{"operator", val("decimal", "greaterThan|0"), val("decimal", "greaterThanOrEqual|5"), true, true},
		{"list items", val("list", `"a,b"`), val("list", `"a,b,c"`), true, true},
		{"list source", val("list", `Summary!$D$2:$D$6`), val("list", `"a,b"`), true, true},
		{"list to whole", val("list", `"a,b"`), val("whole", "between|0|1000"), false, false},
		{"whole to decimal", val("whole", "between|0|1000"), val("decimal", "greaterThan|0"), false, true},
		{"whole to custom", val("whole", "between|0|1000"), val("custom", "ISNUMBER(B2)"), false, false},
		{"whole to any", val("whole", "between|0|1000"), val("", "between"), false, true},
		{"any to list", val("", "between"), val("list", `"a,b"`), false, true},
		{"threshold", cond("expression", "$B2>6"), cond("expression", "$B2>7"), true, true},
		{"comparison", cond("expression", "$B2>6"), cond("expression", "$B2<=6.5"), true, true},
		{"text", cond("expression", `$A2="North"`), cond("expression", `$A2="South"`), true, true},
		{"highlight to banding", cond("expression", "$B2>6"), cond("expression", "MOD(ROW()-2,2)=1"), false, false},
		{"banding shifted", cond("expression", "MOD(ROW()-2,2)=1"), cond("expression", "MOD(ROW()-3,2)=1"), true, true},
		{"custom formula", val("custom", "ISNUMBER(B2)"), val("custom", "AND(ISNUMBER(B2),B2>0)"), false, false},
		{"cell value to bars", cond("cellIs", "greaterThan|140"), cond("dataBar", ""), false, false},
		{"cell value", cond("cellIs", "greaterThan|140"), cond("cellIs", "lessThan|3"), true, true},
		{"cond to none", cond("cellIs", "greaterThan|1"), cond("", ""), false, false},
	} {
		if got := ruleEdited(c.f, c.g, true); got != c.same {
			t.Errorf("%s: same type asked: %v, want %v", c.name, got, c.same)
		}
		if got := ruleEdited(c.f, c.g, false); got != c.anyway {
			t.Errorf("%s: any type: %v, want %v", c.name, got, c.anyway)
		}
	}
	if got := ruleText(val("", "between")); got != "any value" {
		t.Errorf("ruleText of a validation of no type: %q", got)
	}
	if got := ruleText(val("whole", "between|0|1000")); got != "whole between 0 1000" {
		t.Errorf("ruleText: %q", got)
	}
}

// A whole-number validation replaced by "any value" with a message: an edit,
// said as the feature is, not as "none between" (f0.6), in check --against,
// its text and the command's short report.
func TestAgainstCausesAnyValue(t *testing.T) {
	dir, before := causesFixture(t)
	after := filepath.Join(dir, "any.xlsx")
	if err := os.WriteFile(after, mustRead(t, before), 0o644); err != nil {
		t.Fatal(err)
	}
	cmd, code := runCLI("", "validate", after, "Data!B2:B6", "any", "--prompt", "units")
	if code != 0 || !strings.Contains(cmd, "changed: validation Data!B2:B6 (whole between 0 1000 → any value)") {
		t.Errorf("the command's report: exit %d:\n%s", code, cmd)
	}
	r, code := against(t, after, before)
	if len(r.FeaturesChanged) != 1 || r.FeaturesChanged[0].Before != "whole between 0 1000" || r.FeaturesChanged[0].After != "any value" || code != 0 {
		t.Errorf("exit %d, changed %+v, lost %+v", code, r.FeaturesChanged, r.FeaturesLost)
	}
	if text, _ := runCLI("", "check", after, "--against", before); strings.Contains(text, "none") || !strings.Contains(text, "any value") {
		t.Errorf("text:\n%s", text)
	}
}

// changedOK checks check --against finds one feature changed and none lost.
func changedOK(t *testing.T, name, after, before string) {
	t.Helper()
	r, code := against(t, after, before)
	if len(r.FeaturesLost) != 0 || len(r.FeaturesChanged) != 1 || r.Summary.FeaturesChanged != 1 || code != 0 {
		t.Errorf("%s: exit %d, lost %+v, changed %+v", name, code, r.FeaturesLost, r.FeaturesChanged)
		return
	}
	if f := r.FeaturesChanged[0]; f.Before == f.After || !strings.HasPrefix(f.Where, "Data!") {
		t.Errorf("%s: %+v", name, f)
	}
	text, _ := runCLI("", "check", after, "--against", before)
	if !strings.Contains(text, "0 features lost (1 changed)") || !strings.Contains(text, "\nfeatures changed: rules edited on the cells they cover (not lost):\n") {
		t.Errorf("%s: text:\n%s", name, text)
	}
}

// A row inserted above what a defined name refers to by a tool that leaves
// names as they were (openpyxl's insert_rows): the name is flagged, with
// what Excel writes, and the values it moved are its doing, not the row's
// (f5.4). Inserted by sheets, which rewrites it, nothing is said of it.
func TestAgainstCausesNameNotRewritten(t *testing.T) {
	dir, fixture := causesFixture(t)
	before := edited(t, fixture, filepath.Join(dir, "taxed.xlsx"), []string{"set", "Data!E2", "=D2*TaxRate"}, []string{"check", "--fix"})
	rewrote := edited(t, before, filepath.Join(dir, "sheets.xlsx"), []string{"rows", "insert", "Summary!1"})
	r, _ := against(t, rewrote, before)
	for _, e := range r.Structure {
		if strings.Contains(e.Where, "TaxRate") {
			t.Errorf("rewritten by sheets: %+v", e)
		}
	}
	if _, ok := causes(r)["Data!E2"]; ok {
		t.Errorf("rewritten by sheets: Data!E2 moved: %+v", r.Moved)
	}
	stale := filepath.Join(dir, "stale.xlsx")
	rewriteParts(t, rewrote, stale, func(name string, b []byte) []byte {
		if name == "xl/workbook.xml" {
			was := b
			if b = bytes.Replace(b, []byte("Summary!$B$2</definedName>"), []byte("Summary!$B$1</definedName>"), 1); bytes.Equal(b, was) {
				t.Fatalf("no name in %s", b)
			}
		}
		return b
	})
	r, _ = against(t, stale, before)
	found := false
	for _, e := range r.Structure {
		if e.Change == "name not rewritten" {
			found = true
			if !strings.HasPrefix(e.Where, "TaxRate = Summary!$B$1: its reference is as it was") || !strings.HasSuffix(e.Where, "Excel writes Summary!$B$2") {
				t.Errorf("structure: %+v", e)
			}
		}
	}
	if !found {
		t.Errorf("no name flagged: %+v", r.Structure)
	}
	if got := causes(r)["Data!E2"]; !strings.HasPrefix(got, "name TaxRate not rewritten (Excel writes Summary!$B$2)") {
		t.Errorf("Data!E2's cause %q (%+v, unexplained %+v)", got, r.Moved, r.Unexplained)
	}
}

// A used range that spans the grid (A1:B1 and XFD1048576): check --against
// compares the cells there are, not the rectangle, in a few seconds on
// either side and either way round (f5.7); so does sheets diff.
func TestAgainstCausesWholeGrid(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("SHEETS_CACHE_DIR", filepath.Join(dir, ".cache"))
	small := func(name string, a int, far bool) string {
		path := filepath.Join(dir, name)
		if out, code := runCLI("", "set", path, "S!A1", fmt.Sprint(a), "S!B1", "=A1*2"); code != 0 {
			t.Fatal(out)
		}
		if !far {
			return path
		}
		grid := filepath.Join(dir, "far-"+name)
		rewriteSheet(t, path, grid, func(b []byte) []byte {
			b = regexp.MustCompile(`<dimension ref="[^"]*"/>`).ReplaceAll(b, []byte(`<dimension ref="A1:XFD1048576"/>`))
			return bytes.Replace(b, []byte("</sheetData>"), []byte(`<row r="1048576"><c r="XFD1048576"><v>5</v></c></row></sheetData>`), 1)
		})
		return grid
	}
	b, a, plain := small("b.xlsx", 1, true), small("a.xlsx", 3, true), small("p.xlsx", 1, false)
	for _, c := range [][2]string{{a, b}, {b, a}, {a, plain}, {plain, a}} {
		start := time.Now()
		r, _ := against(t, c[0], c[1])
		took := time.Since(start)
		t.Logf("%s against %s: %s", filepath.Base(c[0]), filepath.Base(c[1]), took.Round(time.Millisecond))
		if took > 5*time.Second*raceSlow {
			t.Errorf("%s against %s: %s", filepath.Base(c[0]), filepath.Base(c[1]), took)
		}
		if r.Summary.Unexplained != 0 || r.Summary.FeaturesLost != 0 {
			t.Errorf("%s against %s: %+v", filepath.Base(c[0]), filepath.Base(c[1]), r.Summary)
		}
		if got := causes(r)["S!B1"]; c[0] != plain && c[1] != plain && got != "input S!A1" {
			t.Errorf("%s against %s: S!B1's cause %q (%+v)", filepath.Base(c[0]), filepath.Base(c[1]), got, r.Changed)
		}
	}
	start := time.Now()
	if out, code := runCLI("", "diff", b, a); code != 0 || !strings.Contains(out, "A1  1 → 3") {
		t.Errorf("diff: exit %d: %s", code, out)
	}
	if took := time.Since(start); took > 5*time.Second*raceSlow {
		t.Errorf("diff: %s", took)
	}
}

// A sort on a sheet whose used range spans the grid (2,000 rows and one
// value at XFD1048576): each run of rows moved is a source as wide as the
// sheet, and check --against follows it through the columns that hold
// something, not the 16,384 of each row (f0.5). A total over the rows is
// not moved by it; a cell reading one by position is.
func TestAgainstCausesWholeGridSort(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("SHEETS_CACHE_DIR", filepath.Join(dir, ".cache"))
	var csv strings.Builder
	csv.WriteString("Item,Units\n")
	for i := 1; i <= 2000; i++ {
		fmt.Fprintf(&csv, "I%d,%d\n", i, i*37%101)
	}
	src := filepath.Join(dir, "S.csv")
	if err := os.WriteFile(src, []byte(csv.String()), 0o644); err != nil {
		t.Fatal(err)
	}
	plain := filepath.Join(dir, "plain.xlsx")
	if out, code := runCLI("", "convert", src, plain); code != 0 {
		t.Fatal(out)
	}
	if out, code := runCLI("", "set", plain, "S!D1", "=SUM(B2:B2001)", "S!D2", "=B2"); code != 0 {
		t.Fatal(out)
	}
	before := filepath.Join(dir, "b.xlsx")
	rewriteSheet(t, plain, before, func(b []byte) []byte {
		b = regexp.MustCompile(`<dimension ref="[^"]*"/>`).ReplaceAll(b, []byte(`<dimension ref="A1:XFD1048576"/>`))
		return bytes.Replace(b, []byte("</sheetData>"), []byte(`<row r="1048576"><c r="XFD1048576"><v>5</v></c></row></sheetData>`), 1)
	})
	after := edited(t, before, filepath.Join(dir, "a.xlsx"), []string{"sort", "S!A1:B2001", "Units:desc"})
	start := time.Now()
	r, code := against(t, after, before)
	took := time.Since(start)
	t.Logf("sort of 2,000 rows on a grid-wide sheet: %s", took.Round(time.Millisecond))
	if took > 5*time.Second*raceSlow {
		t.Errorf("took %s", took)
	}
	m := causes(r)
	if _, ok := m["S!D1"]; ok {
		t.Errorf("S!D1 (a total over the rows sorted) moved: %q", m["S!D1"])
	}
	if got := m["S!D2"]; !strings.HasPrefix(got, "rows moved: S!A2:B2 (were S!A") {
		t.Errorf("S!D2 (=B2): cause %q (moved %+v)", got, r.Moved)
	}
	if code != 0 || len(r.Unexplained) != 0 || r.Summary.MovedRows < 1900 {
		t.Errorf("exit %d, unexplained %+v, %d rows moved", code, r.Unexplained, r.Summary.MovedRows)
	}
}

// LibreOffice's round trip writes a TRUE typed in as =TRUE() and drops the
// quotes a sheet's name needs not ('Data'!B2 as Data!B2): Excel's and its
// own way of writing the same, not changes (the f5.5 pool file, saved again
// by LibreOffice).
func TestAgainstCausesLibreOfficeSpelling(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("SHEETS_CACHE_DIR", filepath.Join(dir, ".cache"))
	before := filepath.Join(dir, "b.xlsx")
	if out, code := runCLI("", "set", before, "Data!A1", "TRUE", "Data!A2", "FALSE", "Data!C1", "2", "Data!C2", "3",
		"Data!B1", "=SUM('Data'!C1:C2)", "Data!B2", "='Data'!C1*2", "Data!B3", "=SUM('A1'!B1:B2)"); code != 0 {
		t.Fatal(out)
	}
	after := filepath.Join(dir, "a.xlsx")
	rewriteSheet(t, before, after, func(b []byte) []byte {
		b = regexp.MustCompile(`(<c r="A1"[^>]*t="b"[^>]*>)<v>1</v>`).ReplaceAll(b, []byte("${1}<f>TRUE()</f><v>1</v>"))
		b = regexp.MustCompile(`(<c r="A2"[^>]*t="b"[^>]*>)<v>0</v>`).ReplaceAll(b, []byte("${1}<f>FALSE()</f><v>0</v>"))
		return bytes.ReplaceAll(b, []byte("&apos;Data&apos;!"), []byte("Data!"))
	})
	r, code := against(t, after, before)
	if r.Summary.Inputs+r.Summary.Formulas+r.Summary.Moved != 0 || code != 0 {
		t.Errorf("exit %d, %+v, changed %+v", code, r.Summary, r.Changed)
	}
	// A sheet named like a cell keeps its quotes; a boolean that changed is
	// a change.
	if folded("'A1'!B2") != "'A1'!B2" || folded("'Data'!b2") != "DATA!B2" {
		t.Errorf("folded: %q, %q", folded("'A1'!B2"), folded("'Data'!b2"))
	}
	flipped := filepath.Join(dir, "f.xlsx")
	rewriteSheet(t, before, flipped, func(b []byte) []byte {
		return regexp.MustCompile(`(<c r="A1"[^>]*t="b"[^>]*>)<v>1</v>`).ReplaceAll(b, []byte("${1}<f>FALSE()</f><v>0</v>"))
	})
	if r, _ := against(t, flipped, before); r.Summary.Formulas != 1 {
		t.Errorf("TRUE to =FALSE(): %+v", r.Changed)
	}
}

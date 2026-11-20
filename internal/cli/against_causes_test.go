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

// A sort of some of a table's columns that moved some rows and left others
// that shifted (openpyxl's pool file, a table sorted whole by its second
// column; the formulas of columns beside it stay where they are): what reads
// a row displaced by the rows that moved, by position, moved for the sort,
// whichever rows moved (f3.5, f5.1, f5.2): nothing is unexplained.
func TestAgainstCausesSortDisplaced(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("SHEETS_CACHE_DIR", filepath.Join(dir, ".cache"))
	before := filepath.Join(dir, "b.xlsx")
	if err := os.WriteFile(before, mustRead(t, filepath.Join("..", "fuzz", "testdata", "pool", "openpyxl-04.xlsx")), 0o644); err != nil {
		t.Fatal(err)
	}
	after := edited(t, before, filepath.Join(dir, "a.xlsx"), []string{"sort", "Data!A1:F12", "B"})
	r, _ := against(t, after, before) // the exit code counts the file's own errors too
	if len(r.Unexplained) != 0 {
		t.Errorf("unexplained %+v", r.Unexplained)
	}
	m := causes(r)
	for _, c := range []string{"Data!L9", "Data!L10", "Data!L11", "Data!P11"} {
		if !strings.Contains(m[c], "rows moved: ") {
			t.Errorf("%s: cause %q, want the sort", c, m[c])
		}
	}
}

// A column of running totals (=SUM($B$2:Bn), each a SUM of its own range) after
// a row deleted or inserted inside it costs no more than after a cell set
// (f3.4): whether the lines could have moved each total is told without
// reading its whole range again for each (n²/2 cells), and an empty row
// inserted before any range is read. Counted in the cells read, not in time.
func TestAgainstCausesRunningTotalsFast(t *testing.T) {
	rows := 3000
	if raceSlow > 1 {
		rows = 800 // the race detector: what it finds, not how fast (n²/2 is still 20 times the limit)
	}
	dir := t.TempDir()
	t.Setenv("SHEETS_CACHE_DIR", filepath.Join(dir, ".cache"))
	var b strings.Builder
	b.WriteString("Item\tUnits\tRunning\n")
	for i := 1; i <= rows; i++ {
		fmt.Fprintf(&b, "I%d\t%d\t=SUM($B$2:B%d)\n", i, (i*37)%101+1, i+1)
	}
	before := filepath.Join(dir, "b.xlsx")
	if out, code := runCLI(b.String(), "write", before, "Data!A1"); code != 0 {
		t.Fatalf("write: %s", out)
	}
	for _, c := range []struct {
		name  string
		moved int // the totals that moved
		cmd   []string
	}{
		{"delete", rows - 10, []string{"rows", "delete", "Data!5"}},
		{"insert", 0, []string{"rows", "insert", "Data!5"}},
		{"insert two", 0, []string{"rows", "insert", "Data!5", "-n", "2"}},
	} {
		after := edited(t, before, filepath.Join(dir, "a.xlsx"), c.cmd)
		rangeReads = 0
		r, _ := against(t, after, before)
		if len(r.Unexplained) != 0 || r.Summary.Moved < c.moved {
			t.Errorf("%s: %+v", c.name, r.Summary)
		}
		if rangeReads > 20*rows {
			t.Errorf("%s: %d cells read one by one to tell what the totals give: not linear in the %d rows (n²/2 is %d)", c.name, rangeReads, rows, rows*rows/2)
		}
	}
}

// storeAt rewrites the results a workbook stores for formula cells, by sheet
// part (sheet1 is Data, sheet2 Summary of the fixture), results trusted: no
// full calculation on load.
func storeAt(t *testing.T, src, dst string, parts map[string]map[string]string) {
	t.Helper()
	rewriteParts(t, src, dst, func(name string, b []byte) []byte {
		if name == "xl/workbook.xml" {
			return regexp.MustCompile(` fullCalcOnLoad="(true|1)"`).ReplaceAll(b, nil)
		}
		for part, cells := range parts {
			if name != "xl/worksheets/"+part+".xml" {
				continue
			}
			for cell, v := range cells {
				re := regexp.MustCompile(`(<c r="` + cell + `"[^>]*><f>[^<]*</f>)<v>[^<]*</v>`)
				was := b
				if b = re.ReplaceAll(b, []byte("${1}<v>"+v+"</v>")); bytes.Equal(b, was) {
					t.Fatalf("no result of %s in %s", cell, part)
				}
			}
		}
		return b
	})
}

// A formula carried with its row by a sort, or moved by rows added or removed
// above it, reads what it read: a wrong result stored for it is no work of the
// sort or the lines, and the totals reading it follow it, not the move (f3.2,
// f5.1, f5.2). A formula that reads where it stands (a ROW) moved for the
// sort.
func TestAgainstCausesCarriedRow(t *testing.T) {
	dir, before := causesFixture(t)
	only := func(name string, r causesResult, code int, cell string, follows ...string) {
		t.Helper()
		if len(r.Unexplained) != 1 || !strings.HasPrefix(r.Unexplained[0].Cell, cell) || code != 1 {
			t.Fatalf("%s: exit %d, unexplained %+v, moved %+v", name, code, r.Unexplained, r.Moved)
		}
		m := causes(r)
		for _, f := range follows {
			if want := cell + ", which moved with nothing upstream changed"; m[f] != want {
				t.Errorf("%s: %s: cause %q, want %q", name, f, m[f], want)
			}
		}
	}

	// Sorted, a result stored for the cell that was D5 (9 x 10 = 90) and for
	// the totals reading it.
	sorted := edited(t, before, filepath.Join(dir, "s.xlsx"), []string{"sort", "Data!A1:D6", "Units:desc"})
	after := filepath.Join(dir, "a.xlsx")
	storeAt(t, sorted, after, map[string]map[string]string{"sheet1": {"D2": "999", "D7": "1539"}, "sheet2": {"B3": "1539"}})
	r, code := against(t, after, before)
	only("sorted", r, code, "Data!D2", "Data!D7", "Summary!B3")
	if r.Unexplained[0].Was != "D5" {
		t.Errorf("sorted: %+v", r.Unexplained[0])
	}

	// An empty row inserted above it: it moves nothing.
	inserted := edited(t, before, filepath.Join(dir, "i.xlsx"), []string{"rows", "insert", "Data!3"})
	storeAt(t, inserted, after, map[string]map[string]string{"sheet1": {"D6": "999"}})
	r, code = against(t, after, before)
	only("inserted", r, code, "Data!D6")

	// A row removed above it: the formula's own row is as it was.
	deleted := edited(t, before, filepath.Join(dir, "d.xlsx"), []string{"rows", "delete", "Data!2"})
	storeAt(t, deleted, after, map[string]map[string]string{"sheet1": {"D5": "999"}})
	r, code = against(t, after, before)
	only("deleted", r, code, "Data!D5")

	// The same sort and no wrong result: nothing unexplained, as before.
	r, code = against(t, sorted, before)
	if len(r.Unexplained) != 0 || code != 0 {
		t.Errorf("sorted alone: exit %d, unexplained %+v", code, r.Unexplained)
	}

	// A formula that reads where it stands moves with the rows: ROW() went
	// with its row, and gives another number.
	rowed := edited(t, before, filepath.Join(dir, "rowed.xlsx"), []string{"set", "Data!E2", "=ROW()", "Data!E3", "=ROW()", "Data!E4", "=ROW()", "Data!E5", "=ROW()", "Data!E6", "=ROW()"})
	after = edited(t, rowed, filepath.Join(dir, "rowed-s.xlsx"), []string{"sort", "Data!A1:E6", "Units:desc"})
	r, code = against(t, after, rowed)
	moved := 0
	for _, m := range r.Moved {
		if strings.HasPrefix(m.Cell, "Data!E") && strings.Contains(m.Cause, "rows moved") {
			moved++
		}
	}
	if moved == 0 || len(r.Unexplained) != 0 || code != 0 {
		t.Errorf("ROW() in sorted rows: exit %d, unexplained %+v, moved %+v", code, r.Unexplained, r.Moved)
	}
}

// A row removed and an input changed in the same column: the MAX that rose
// moved for the input (the row held less than it was, f5.3); the totals moved
// for both, the row removed first, whichever way they moved (the row took 5
// off a total that rose with the input, f0.1 of the last audit); the cell that
// read the row removed moved for the removal.
func TestAgainstCausesRemovedRise(t *testing.T) {
	dir, before := causesFixture(t)
	after := edited(t, before, filepath.Join(dir, "a.xlsx"), []string{"rows", "delete", "Data!2"}, []string{"set", "Data!B3", "77"})
	r, _ := against(t, after, before)
	m := causes(r)
	if m["Summary!B6"] != "input Data!B3" {
		t.Errorf("Summary!B6 (MAX): cause %q, want input Data!B3", m["Summary!B6"])
	}
	for _, c := range []string{"Data!B6", "Data!D6", "Summary!B2", "Summary!B3", "Summary!B4", "Summary!B5"} {
		if want := "rows removed: Data!2 in " + before + ", input Data!B3"; m[c] != want {
			t.Errorf("%s: cause %q, want %q", c, m[c], want)
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
	re := regexp.MustCompile(`(<c r="` + cell + `"[^>]*><f[^>]*>[^<]*</f>)<v>[^<]*</v>`)
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

// A validation or conditional format edited only in the cells it refers to is
// a feature changed, as one whose items or bound changed is (f3.3, f5.6): a
// dropdown re-pointed at another range (the guide's own range form), a formula
// highlight testing another column, a threshold given as a cell, in
// check --against and in the command's own report. The same rule where rows
// inserted above it moved what it reads, as Excel rewrites it, is the same
// rule; where a tool left it as it was, it is changed, with what Excel
// writes.
func TestAgainstCausesFeaturesRepointed(t *testing.T) {
	dir, fixture := causesFixture(t)
	listed := edited(t, fixture, filepath.Join(dir, "listed.xlsx"), []string{"validate", "Data!A2:A6", "list", "=Summary!$D$2:$D$4"})
	after := filepath.Join(dir, "repointed.xlsx")
	if err := os.WriteFile(after, mustRead(t, listed), 0o644); err != nil {
		t.Fatal(err)
	}
	out, code := runCLI("", "validate", after, "Data!A2:A6", "list", "=Summary!$D$2:$D$6")
	if code != 0 || !strings.Contains(head(t, out), ", 0 features lost (1 changed)") || !strings.Contains(out, "\nchanged: validation Data!A2:A6") {
		t.Errorf("re-pointed dropdown: the command's report (exit %d):\n%s", code, out)
	}
	changedOK(t, "re-pointed dropdown", after, listed)

	// A threshold that is a cell, another cell.
	bound := edited(t, fixture, filepath.Join(dir, "bound.xlsx"), []string{"validate", "Data!B2:B6", "whole", ">=", "=Summary!$B$1"})
	moved := edited(t, bound, filepath.Join(dir, "bound2.xlsx"), []string{"validate", "Data!B2:B6", "whole", ">=", "=Summary!$B$2"})
	changedOK(t, "threshold cell", moved, bound)

	// A highlight testing another column: formula rules written by another
	// tool, one cell apart.
	rule := func(f string) func([]byte) []byte {
		return func(b []byte) []byte {
			return bytes.Replace(b, []byte("</conditionalFormatting>"), []byte(`</conditionalFormatting><conditionalFormatting sqref="A2:C6"><cfRule type="expression" dxfId="0" priority="2"><formula>`+f+`</formula></cfRule></conditionalFormatting>`), 1)
		}
	}
	hi := filepath.Join(dir, "hi.xlsx")
	rewriteSheet(t, fixture, hi, rule("$B2&gt;6"))
	hi2 := filepath.Join(dir, "hi2.xlsx")
	rewriteSheet(t, fixture, hi2, rule("$C2&gt;6"))
	changedOK(t, "highlight testing another column", hi2, hi)

	// Rows inserted above the list, rewritten as Excel does (sheets does the
	// same): the dropdown is as it was.
	inserted := edited(t, listed, filepath.Join(dir, "inserted.xlsx"), []string{"rows", "insert", "Summary!2"})
	if r, code := against(t, inserted, listed); len(r.FeaturesLost) != 0 || len(r.FeaturesChanged) != 0 || code != 0 {
		t.Errorf("rows inserted above what the list reads: exit %d, lost %+v, changed %+v", code, r.FeaturesLost, r.FeaturesChanged)
	}
	// The same rows inserted by a tool that leaves the rule as it was: changed,
	// what Excel writes said.
	stale := filepath.Join(dir, "stale.xlsx")
	rewriteParts(t, inserted, stale, func(name string, b []byte) []byte {
		if name == "xl/worksheets/sheet1.xml" {
			return bytes.ReplaceAll(b, []byte("$D$3:$D$5"), []byte("$D$2:$D$4"))
		}
		return b
	})
	r, code := against(t, stale, listed)
	if len(r.FeaturesLost) != 0 || len(r.FeaturesChanged) != 1 || code != 0 || !strings.Contains(r.FeaturesChanged[0].After, "Excel writes") {
		t.Errorf("rule left as it was: exit %d, lost %+v, changed %+v", code, r.FeaturesLost, r.FeaturesChanged)
	}
}

// A rule is written for the first cell of its cells: rows or columns deleted
// or inserted at its first row or column, or before it, leave it the same rule
// (f3.3: the cells it reads are compared as Excel rewrites them, a relative
// reference by where the first cell left stands). A rule whose first cell
// was deleted is the one of the next cell. (One whose own column was deleted
// reads #REF! now, and is lost, as the validation on it is.)
func TestAgainstCausesRuleFirstCell(t *testing.T) {
	dir, fixture := causesFixture(t)
	before := filepath.Join(dir, "highlight.xlsx")
	rewriteSheet(t, fixture, before, func(b []byte) []byte {
		return bytes.Replace(b, []byte("</conditionalFormatting>"), []byte(`</conditionalFormatting><conditionalFormatting sqref="B2:D6"><cfRule type="expression" dxfId="0" priority="2"><formula>$B2&gt;6</formula></cfRule></conditionalFormatting>`), 1)
	})
	for _, c := range []struct {
		name string
		cmd  []string
	}{
		{"first row deleted", []string{"rows", "delete", "Data!2"}},
		{"row above deleted", []string{"rows", "delete", "Data!1"}},
		{"first two rows deleted", []string{"rows", "delete", "Data!2:3"}},
		{"row inserted at the first", []string{"rows", "insert", "Data!2"}},
		{"row inserted above", []string{"rows", "insert", "Data!1"}},
		{"column before deleted", []string{"cols", "delete", "Data!A"}},
		{"column before inserted", []string{"cols", "insert", "Data!A"}},
		{"column inserted at the first", []string{"cols", "insert", "Data!B"}},
	} {
		after := edited(t, before, filepath.Join(dir, "a.xlsx"), c.cmd)
		r, _ := against(t, after, before) // the exit code counts the file's own errors too
		lost := r.FeaturesLost
		if c.name == "column before deleted" && len(lost) == 1 && lost[0].Kind == "validation" && lost[0].Where == "Data!A2:A6" {
			// The fixture's list on Data!A2:A6 (in Excel 2010's extension
			// list) had all its cells in the column: gone with it, as Excel
			// drops it, and lost.
			lost = nil
		}
		if len(lost) != 0 || len(r.FeaturesChanged) != 0 {
			t.Errorf("%s: lost %+v, changed %+v", c.name, r.FeaturesLost, r.FeaturesChanged)
		}
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
	if !strings.Contains(out, "conditional format Data!A2:D6 (formula rule: expression $B2>6)") {
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
	if raceSlow > 1 {
		// A grid-wide sheet costs 10 s under the race detector for what is a
		// test of speed; the plain and streaming stages run it.
		t.Skip("takes 10 s under -race; run without it")
	}
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
		var r causesResult
		took := runWithin(5*time.Second*raceSlow, func() { r, _ = against(t, c[0], c[1]) })
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
	took := runWithin(5*time.Second*raceSlow, func() {
		if out, code := runCLI("", "diff", b, a); code != 0 || !strings.Contains(out, "A1  1 → 3") {
			t.Errorf("diff: exit %d: %s", code, out)
		}
	})
	if took > 5*time.Second*raceSlow {
		t.Errorf("diff: %s", took)
	}
}

// runWithin runs a step until one run is within limit (three runs at most: a
// machine busy with many test processes slows one run, not each of them, and
// a step that is slow for a reason takes long in each) and gives the last
// run's time.
func runWithin(limit time.Duration, run func()) time.Duration {
	var took time.Duration
	for try := 0; try < 3; try++ {
		start := time.Now()
		run()
		if took = time.Since(start); took <= limit {
			break
		}
	}
	return took
}

// A sort on a sheet whose used range spans the grid (2,000 rows and one
// value at XFD1048576): each run of rows moved is a source as wide as the
// sheet, and check --against follows it through the columns that hold
// something, not the 16,384 of each row (f0.5). A total over the rows is
// not moved by it; a cell reading one by position is.
func TestAgainstCausesWholeGridSort(t *testing.T) {
	if raceSlow > 1 {
		// Under the race detector following a sort through a grid-wide sheet
		// takes 25 s, whatever the rows (300 or 2,000): the stage's longest.
		// The plain and streaming stages run it; the race stage checks the
		// same code through the other sorts here.
		t.Skip("takes 25 s under -race; run without it")
	}
	rows := 2000
	dir := t.TempDir()
	t.Setenv("SHEETS_CACHE_DIR", filepath.Join(dir, ".cache"))
	var csv strings.Builder
	csv.WriteString("Item,Units\n")
	for i := 1; i <= rows; i++ {
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
	if out, code := runCLI("", "set", plain, "S!D1", fmt.Sprintf("=SUM(B2:B%d)", rows+1), "S!D2", "=B2"); code != 0 {
		t.Fatal(out)
	}
	before := filepath.Join(dir, "b.xlsx")
	rewriteSheet(t, plain, before, func(b []byte) []byte {
		b = regexp.MustCompile(`<dimension ref="[^"]*"/>`).ReplaceAll(b, []byte(`<dimension ref="A1:XFD1048576"/>`))
		return bytes.Replace(b, []byte("</sheetData>"), []byte(`<row r="1048576"><c r="XFD1048576"><v>5</v></c></row></sheetData>`), 1)
	})
	after := edited(t, before, filepath.Join(dir, "a.xlsx"), []string{"sort", fmt.Sprintf("S!A1:B%d", rows+1), "Units:desc"})
	var r causesResult
	var code int
	took := runWithin(5*time.Second*raceSlow, func() { r, code = against(t, after, before) })
	t.Logf("sort of %d rows on a grid-wide sheet: %s", rows, took.Round(time.Millisecond))
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
	if code != 0 || len(r.Unexplained) != 0 || r.Summary.MovedRows < rows*19/20 {
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

// A row deleted, then a cell set in the row that moved into its place, on
// the pool's files from other producers (f1.2): the row deleted is the one
// removed, whichever rows beside it hold formulas of the same pattern or
// nothing at all (a blank row below the last of the data), one input is
// changed, and Excel's own rewrite of a formula is no change and no mistake.
func TestAgainstCausesDeleteSetNeighbour(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("SHEETS_CACHE_DIR", filepath.Join(dir, ".cache"))
	for _, c := range []struct {
		file string
		row  int
	}{
		{"openpyxl-32", 3}, {"xlsxwriter-39", 3}, {"xlsxwriter-35", 5}, {"xlsxwriter-09", 6}, {"xlsxwriter-15", 6},
		{"xlsxwriter-31", 6}, {"openpyxl-14", 7}, {"xlsxwriter-29", 7},
	} {
		before := filepath.Join(dir, c.file+".xlsx")
		if err := os.WriteFile(before, mustRead(t, filepath.Join("..", "fuzz", "testdata", "pool", c.file+".xlsx")), 0o644); err != nil {
			t.Fatal(err)
		}
		for _, col := range []string{"B", "C", "D"} {
			cell := fmt.Sprintf("Data!%s%d", col, c.row)
			after := edited(t, before, filepath.Join(dir, "a.xlsx"), []string{"rows", "delete", fmt.Sprintf("Data!%d", c.row)}, []string{"set", cell, "77"})
			r, _ := against(t, after, before) // the exit code counts the file's own errors too
			s := r.Summary
			if s.Inputs != 1 || s.Formulas != 0 || s.Flagged != 0 || s.Unexplained != 0 || s.FeaturesLost != 0 {
				t.Errorf("%s: %s: %+v, changed %+v", c.file, cell, s, r.Changed)
			}
			var rows []string
			for _, e := range r.Structure {
				if strings.HasPrefix(e.Change, "row") {
					rows = append(rows, e.Change+" "+strings.SplitN(e.Where, " in ", 2)[0])
				}
			}
			if want := fmt.Sprintf("row removed Data!%d", c.row); len(rows) != 1 || rows[0] != want {
				t.Errorf("%s: %s: structure %q, want %q", c.file, cell, rows, want)
			}
			if len(r.Changed) == 1 && r.Changed[0].Cell != cell {
				t.Errorf("%s: %s: changed %+v", c.file, cell, r.Changed)
			}
		}
	}
}

// The last row of a table deleted and a cell set in the row now at that index,
// when nothing moves up into its place (f3.1, f5.5): one row removed and one
// input changed (typed in from empty), Excel's own rewrite of a range and of a
// name no change, nothing flagged, no feature lost. The same with the last two
// rows deleted and the first of them set: two rows removed. A plain table by
// XlsxWriter, and openpyxl's pool file with the names and a chart.
func TestAgainstCausesDeleteSetLast(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("SHEETS_CACHE_DIR", filepath.Join(dir, ".cache"))
	for _, c := range []struct {
		file string
		last int // the last row of Data's table
	}{
		{"openpyxl-02", 7}, {"xlsxwriter-09", 6}, {"xlsxwriter-15", 6}, {"xlsxwriter-31", 6},
	} {
		before := filepath.Join(dir, c.file+".xlsx")
		if err := os.WriteFile(before, mustRead(t, filepath.Join("..", "fuzz", "testdata", "pool", c.file+".xlsx")), 0o644); err != nil {
			t.Fatal(err)
		}
		for _, run := range []struct {
			rows  string
			first int
			want  string
		}{
			{fmt.Sprint(c.last), c.last, fmt.Sprintf("row removed Data!%d", c.last)},
			{fmt.Sprintf("%d:%d", c.last-1, c.last), c.last - 1, fmt.Sprintf("rows removed Data!%d:%d", c.last-1, c.last)},
		} {
			for _, col := range []string{"B", "C", "D"} {
				cell := fmt.Sprintf("Data!%s%d", col, run.first)
				name := c.file + ": delete Data!" + run.rows + ", set " + cell
				after := edited(t, before, filepath.Join(dir, "a.xlsx"), []string{"rows", "delete", "Data!" + run.rows}, []string{"set", cell, "77"})
				r, _ := against(t, after, before) // the exit code counts the file's own errors too
				s := r.Summary
				if s.Inputs != 1 || s.Formulas != 0 || s.Flagged != 0 || s.Unexplained != 0 || s.FeaturesLost != 0 {
					t.Errorf("%s: %+v, changed %+v", name, s, r.Changed)
				}
				var lines []string
				for _, e := range r.Structure {
					lines = append(lines, e.Change+" "+strings.SplitN(e.Where, " in ", 2)[0])
				}
				if len(lines) != 1 || lines[0] != run.want {
					t.Errorf("%s: structure %q, want [%q]", name, lines, run.want)
				}
				if len(r.Changed) == 1 && (r.Changed[0].Cell != cell || r.Changed[0].Before != "" || r.Changed[0].After != "77") {
					t.Errorf("%s: changed %+v", name, r.Changed)
				}
			}
		}
	}
}

// totalsFixture copies against_totals.xlsx (Data!B2:B6 holding 5, 7, 3, 9, 4,
// an Excel table Sales over A1:B6 and the name Units over B2:B6; Sum!A1
// =SUM(Sales[Units]), A2 =MAX(Sales[Units]), A3 =SUM(Units), A4 =MAX(Units),
// A5 =COUNT(Sales[Units])) with its results stored, as Excel stores them.
func totalsFixture(t *testing.T) (dir, before string) {
	t.Helper()
	dir = t.TempDir()
	t.Setenv("SHEETS_CACHE_DIR", filepath.Join(dir, ".cache"))
	before = filepath.Join(dir, "b.xlsx")
	if err := os.WriteFile(before, mustRead(t, filepath.Join("..", "xlsx", "testdata", "against_totals.xlsx")), 0o644); err != nil {
		t.Fatal(err)
	}
	if out, code := runCLI("", "check", "--fix", before); code != 0 {
		t.Fatalf("check --fix: exit %d: %s", code, out)
	}
	return dir, before
}

// unexplainedAt tells whether a cell is among the values unexplained.
func unexplainedAt(r causesResult, cell string) bool {
	for _, u := range r.Unexplained {
		if u.Cell == cell {
			return true
		}
	}
	return false
}

// A total over an Excel table's column or a defined name is a call over
// references like one over a range (f1.1): a row removed is no cause of a MAX
// that rose or a SUM that rose, whichever way the formula reads the numbers,
// and a row added that holds nothing, inside the table, none of a total of
// another number. What the row can move, it moves.
func TestAgainstCausesRemovedTable(t *testing.T) {
	dir, before := totalsFixture(t)
	deleted := edited(t, before, filepath.Join(dir, "d.xlsx"), []string{"rows", "delete", "Data!6"})
	for _, c := range []struct{ cell, value string }{
		{"A1", "999"}, {"A2", "50"}, {"A3", "999"}, {"A4", "50"}, {"A5", "7"},
	} {
		after := filepath.Join(dir, "a.xlsx")
		stored(t, deleted, after, c.cell, c.value)
		r, code := against(t, after, before)
		if m := causes(r)["Sum!"+c.cell]; !unexplainedAt(r, "Sum!"+c.cell) || m != "" || code != 1 {
			t.Errorf("%s stored as %s: exit %d, cause %q, unexplained %+v", c.cell, c.value, code, m, r.Unexplained)
		}
	}
	// The way the row can move them: the 4 it held was in the sums and in the
	// count, below the MAX.
	r, code := against(t, deleted, before)
	m := causes(r)
	for _, c := range []string{"Sum!A1", "Sum!A3", "Sum!A5"} {
		if !strings.HasPrefix(m[c], "rows removed: Data!6") {
			t.Errorf("%s: cause %q (moved %+v)", c, m[c], r.Moved)
		}
	}
	if len(r.Unexplained) != 0 || code != 0 {
		t.Errorf("the row removed: exit %d, unexplained %+v", code, r.Unexplained)
	}
	// An empty row inserted inside the table moves nothing; a total stored as
	// another number is unexplained.
	empty := edited(t, before, filepath.Join(dir, "i.xlsx"), []string{"rows", "insert", "Data!4"})
	if r, code := against(t, empty, before); len(r.Moved)+len(r.Unexplained) != 0 || code != 0 {
		t.Errorf("an empty row in the table: exit %d, moved %+v, unexplained %+v", code, r.Moved, r.Unexplained)
	}
	for _, c := range []struct{ cell, value string }{{"A1", "999"}, {"A2", "50"}, {"A3", "999"}, {"A4", "50"}} {
		after := filepath.Join(dir, "a.xlsx")
		stored(t, empty, after, c.cell, c.value)
		r, code := against(t, after, before)
		if m := causes(r)["Sum!"+c.cell]; !unexplainedAt(r, "Sum!"+c.cell) || m != "" || code != 1 {
			t.Errorf("an empty row in the table, %s stored as %s: exit %d, cause %q, unexplained %+v", c.cell, c.value, code, m, r.Unexplained)
		}
	}
}

// A SUM or a COUNT that moved further than the lines removed or added could
// move it is not their doing (f1.4): the row of 4 removed lowers the total by
// 4, not by 28; a number added raises it by that number.
func TestAgainstCausesRemovedSize(t *testing.T) {
	dir, before := causesFixture(t)
	deleted := edited(t, before, filepath.Join(dir, "d.xlsx"), []string{"rows", "delete", "Data!6"})
	after := filepath.Join(dir, "a.xlsx")
	stored(t, deleted, after, "B2", "0")
	r, code := against(t, after, before)
	if m := causes(r)["Summary!B2"]; !unexplainedAt(r, "Summary!B2") || m != "" || code != 1 {
		t.Errorf("a SUM that fell by 28: exit %d, cause %q, unexplained %+v", code, m, r.Unexplained)
	}
	// What the row held is what the total fell by: the value sheets computes.
	r, code = against(t, deleted, before)
	if m := causes(r)["Summary!B2"]; !strings.HasPrefix(m, "rows removed: Data!6") || len(r.Unexplained) != 0 || code != 0 {
		t.Errorf("a SUM that fell by 4: exit %d, cause %q, unexplained %+v", code, m, r.Unexplained)
	}
	// A number added: the total rose by it, not by more.
	added := edited(t, before, filepath.Join(dir, "n.xlsx"), []string{"rows", "insert", "Data!4"}, []string{"set", "Data!B4", "11"})
	stored(t, added, after, "B2", "100")
	r, code = against(t, after, before)
	if m := causes(r)["Summary!B2"]; !unexplainedAt(r, "Summary!B2") || m != "" || code != 1 {
		t.Errorf("a SUM that rose by 72 for a 11 added: exit %d, cause %q, unexplained %+v", code, m, r.Unexplained)
	}
	r, code = against(t, added, before)
	if m := causes(r)["Summary!B2"]; m != "rows added: Data!4" || len(r.Unexplained) != 0 || code != 0 {
		t.Errorf("a SUM that rose by 11: exit %d, cause %q, unexplained %+v", code, m, r.Unexplained)
	}
	// A COUNT that fell by more than the numbers the row held.
	tdir, tbefore := totalsFixture(t)
	tdeleted := edited(t, tbefore, filepath.Join(tdir, "d.xlsx"), []string{"rows", "delete", "Data!6"})
	tafter := filepath.Join(tdir, "a.xlsx")
	stored(t, tdeleted, tafter, "A5", "2")
	r, code = against(t, tafter, tbefore)
	if m := causes(r)["Sum!A5"]; !unexplainedAt(r, "Sum!A5") || m != "" || code != 1 {
		t.Errorf("a COUNT that fell by 3: exit %d, cause %q, unexplained %+v", code, m, r.Unexplained)
	}
}

// A blank row inserted every few rows of a long table, and one input changed
// (f1.3): the total over the table moved for the input, and the blank rows
// are looked through together, not one run at a time (each used to cost a
// whole pass over the sheet: minutes at 150,000 rows).
func TestAgainstCausesAddedRunsFast(t *testing.T) {
	rows, every, limit := 12000, 6, 3*time.Second
	if raceSlow > 1 {
		rows, every, limit = 3000, 3, 0 // the race detector: what it finds, not how fast
	}
	dir := t.TempDir()
	t.Setenv("SHEETS_CACHE_DIR", filepath.Join(dir, ".cache"))
	build := func(path string, gaps bool) {
		var b strings.Builder
		b.WriteString("Item\tUnits\tPrice\tRevenue\n")
		n := 1
		for i := 0; i < rows; i++ {
			if gaps && i > 0 && i%every == 0 {
				b.WriteString("\n")
				n++
			}
			units := i%7 + 1
			if gaps && i == 1 {
				units = 100
			}
			n++
			fmt.Fprintf(&b, "I%d\t%d\t10\t=B%d*C%d\n", i, units, n, n)
		}
		if out, code := runCLI(b.String(), "write", path, "Data!A1"); code != 0 {
			t.Fatalf("write: %s", out)
		}
		if out, code := runCLI("", "set", path, "Summary!A1", "Total", "Summary!B1", fmt.Sprintf("=SUM(Data!D2:D%d)", n)); code != 0 {
			t.Fatalf("set: %s", out)
		}
	}
	before, after := filepath.Join(dir, "b.xlsx"), filepath.Join(dir, "a.xlsx")
	build(before, false)
	build(after, true)
	start := time.Now()
	r, code := against(t, after, before)
	took := time.Since(start)
	if m := causes(r)["Summary!B1"]; m != "input Data!B3" || r.Summary.Inputs != 1 || len(r.Unexplained) != 0 || code != 0 {
		t.Errorf("exit %d, %+v, Summary!B1 cause %q, unexplained %+v", code, r.Summary, m, r.Unexplained)
	}
	if limit > 0 && took > limit {
		t.Errorf("check --against with %d runs of rows added took %s, over %s", rows/every-1, took.Round(time.Millisecond), limit)
	}
}

// A sort with rows inserted above the sorted rows (a title row) or deleted
// inside them (f2.1): the block the sort moved is the rows as numbered now,
// so a total over all of them is still no work of the sort, and a result
// nothing accounts for stays unexplained.
func TestAgainstCausesSortRowsShifted(t *testing.T) {
	dir, before := causesFixture(t)
	for _, c := range []struct {
		name string
		cmds [][]string
	}{
		{"a row inserted above", [][]string{{"rows", "insert", "Data!1"}, {"sort", "Data!A2:D7", "Units:desc"}}},
		{"a row deleted inside", [][]string{{"rows", "delete", "Data!3"}, {"sort", "Data!A1:D5", "Units:desc"}}},
	} {
		shifted := edited(t, before, filepath.Join(dir, "s.xlsx"), c.cmds...)
		r, code := against(t, shifted, before)
		for _, m := range r.Moved {
			if m.Cell != "Summary!B7" && strings.Contains(m.Cause, "rows moved") {
				t.Errorf("%s: %s moved for the sort: %q", c.name, m.Cell, m.Cause)
			}
		}
		if code != 0 || len(r.Unexplained) != 0 || r.Summary.MovedRows == 0 {
			t.Errorf("%s: exit %d, %+v, unexplained %+v", c.name, code, r.Summary, r.Unexplained)
		}
		if m := causes(r)["Summary!B7"]; !strings.HasPrefix(m, "rows moved: Data!") {
			t.Errorf("%s: Summary!B7 (=Data!B2): cause %q", c.name, m)
		}
		// The total stored wrong: nothing accounts for it, with the sort as
		// without it.
		after := filepath.Join(dir, "a.xlsx")
		stored(t, shifted, after, "B2", "999")
		r, code = against(t, after, before)
		if !unexplainedAt(r, "Summary!B2") || code != 1 {
			t.Errorf("%s: a total stored as 999: exit %d, unexplained %+v, moved %+v", c.name, code, r.Unexplained, r.Moved)
		}
	}
	// An input changed among the rows sorted: the totals over all of them
	// moved for it, and the sort is named only for the cell that reads a row.
	set := edited(t, before, filepath.Join(dir, "set.xlsx"), []string{"rows", "insert", "Data!1"}, []string{"sort", "Data!A2:D7", "Units:desc"}, []string{"set", "Data!B4", "11"})
	r, code := against(t, set, before)
	m := causes(r)
	for _, c := range []string{"Data!B8", "Data!D8", "Summary!B2", "Summary!B3", "Summary!B4", "Summary!B5", "Summary!B6"} {
		if m[c] != "input Data!B4" {
			t.Errorf("%s: cause %q, want input Data!B4", c, m[c])
		}
	}
	if code != 0 || len(r.Unexplained) != 0 {
		t.Errorf("exit %d, unexplained %+v", code, r.Unexplained)
	}
}

// Cells cut and pasted past a total row (f2.2): the total did not move and
// reads the cells that left, so it moved for the move (the rows moved name
// the span they passed over), not with nothing to say why.
func TestAgainstCausesMoveBelowTotal(t *testing.T) {
	dir, before := causesFixture(t)
	for _, mv := range [][]string{{"move", "Data!A3:D4", "Data!A20"}, {"move", "Data!B3:C4", "Data!B20"}} {
		after := edited(t, before, filepath.Join(dir, "a.xlsx"), mv)
		r, code := against(t, after, before)
		if code != 0 || len(r.Unexplained) != 0 {
			t.Errorf("%v: exit %d, unexplained %+v", mv, code, r.Unexplained)
		}
		m := causes(r)
		for _, c := range []string{"Data!B7", "Summary!B2"} {
			if !strings.HasPrefix(m[c], "rows moved: Data!") {
				t.Errorf("%v: %s: cause %q (moved %+v)", mv, c, m[c], r.Moved)
			}
		}
	}
}

// The calls over what a sort moved cost the sorted rows once each, not once
// for every run of rows it moved (f2.3): 20,000 rows sorted, 200 lookups over
// the table, each run of rows a source of its own over the same block.
//
// The bar is the time with the lookups over the time of the same sort
// without them, the best of up to three runs each: a machine busy with other tests
// (check full runs three suites at once) slows both alike, where a wall-clock
// bar failed on load alone. On an M4: 2.5 s against 0.44 s (5.7 times); with
// f2.3's fix taken out, 7 s (16 times).
func TestAgainstCausesSortRunsFast(t *testing.T) {
	rows, lookups, bar := 20000, 200, 10.0
	if raceSlow > 1 {
		rows, lookups, bar = 3000, 40, 0 // the race detector: what it finds, not how fast
	}
	dir := t.TempDir()
	t.Setenv("SHEETS_CACHE_DIR", filepath.Join(dir, ".cache"))
	var b strings.Builder
	b.WriteString("Item\tUnits\tPrice\tRevenue\n")
	for i := 0; i < rows; i++ {
		fmt.Fprintf(&b, "I%d\t%d\t10\t=B{r}*C{r}\n", i, i*37%101)
	}
	plain := filepath.Join(dir, "p.xlsx")
	if out, code := runCLI(b.String(), "write", plain, "Data!A1"); code != 0 {
		t.Fatalf("write: %s", out)
	}
	before := filepath.Join(dir, "b.xlsx")
	if err := os.WriteFile(before, mustRead(t, plain), 0o644); err != nil {
		t.Fatal(err)
	}
	if out, code := runCLI("", "set", plain, "Summary!A1", "none"); code != 0 {
		t.Fatalf("set: %s", out)
	}
	args := []string{"set", before}
	for k := 0; k < lookups; k++ {
		args = append(args, fmt.Sprintf("Summary!A%d", k+1), fmt.Sprintf("I%d", k*97%rows), fmt.Sprintf("Summary!B%d", k+1), fmt.Sprintf("=VLOOKUP(A%d,Data!$A:$D,2,FALSE)", k+1))
	}
	if out, code := runCLI("", args...); code != 0 {
		t.Fatalf("set: %s", out)
	}
	sort := []string{"sort", fmt.Sprintf("Data!A1:D%d", rows+1), "Units:desc"}
	after := edited(t, before, filepath.Join(dir, "a.xlsx"), sort)
	plainAfter := edited(t, plain, filepath.Join(dir, "pa.xlsx"), sort)
	var r causesResult
	var code int
	if bar == 0 {
		r, code = against(t, after, before)
	}
	// Each timed again, the best kept, only while over the bar: a run slowed
	// by the machine is rerun, a slow step stays slow.
	var took, base time.Duration
	ratio := 0.0
	for try := 0; bar > 0 && try < 3 && (try == 0 || ratio > bar); try++ {
		start := time.Now()
		r, code = against(t, after, before)
		if d := time.Since(start); try == 0 || d < took {
			took = d
		}
		start = time.Now()
		against(t, plainAfter, plain)
		if d := time.Since(start); try == 0 || d < base {
			base = d
		}
		ratio = float64(took) / float64(max(base, time.Millisecond))
	}
	if code != 0 || len(r.Unexplained) != 0 || r.Summary.MovedRows < rows*9/10 {
		t.Errorf("exit %d, %+v, unexplained %+v", code, r.Summary, r.Unexplained)
	}
	if bar == 0 {
		return
	}
	t.Logf("%d rows sorted: %d lookups over them %s, none %s: %.1f times", rows, lookups, took.Round(time.Millisecond), base.Round(time.Millisecond), ratio)
	if ratio > bar {
		t.Errorf("check --against after a sort of %d rows with %d lookups over them took %s, %.1f times the %s of the same sort without them (bar %.0f times)", rows, lookups, took.Round(time.Millisecond), ratio, base.Round(time.Millisecond), bar)
	}
}

// A MAX or a MIN over dates in a workbook that counts them from 1904 (a
// Mac's), in a cell that is no date: the result is a serial of that system,
// and the dates are as many days short of the 1900 system the values are
// compared in. The row removed that held the latest (earliest) date moved it;
// a result nothing accounts for does not become the row's work (f2.4).
func TestAgainstCausesDatesMaxMin1904(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("SHEETS_CACHE_DIR", filepath.Join(dir, ".cache"))
	plain := filepath.Join(dir, "p.xlsx")
	if out, code := runCLI("Item\tDue\nI1\t2026-01-05\nI2\t2026-01-17\nI3\t2026-01-03\nI4\t2026-01-29\nI5\t2026-01-11\n", "write", plain, "Data!A1"); code != 0 {
		t.Fatal(out)
	}
	if out, code := runCLI("", "set", plain, "Summary!A1", "Latest", "Summary!B1", "=MAX(Data!B2:B6)", "Summary!A2", "Earliest", "Summary!B2", "=MIN(Data!B2:B6)"); code != 0 {
		t.Fatal(out)
	}
	before := filepath.Join(dir, "b.xlsx")
	rewriteParts(t, plain, before, func(name string, b []byte) []byte {
		if name == "xl/workbook.xml" {
			b = bytes.Replace(b, []byte("<workbookPr "), []byte(`<workbookPr date1904="true" `), 1)
		}
		return b
	})
	for _, c := range []struct{ name, row, cell string }{
		{"the latest date's row", "Data!5", "Summary!B1"},   // 2026-01-29
		{"the earliest date's row", "Data!4", "Summary!B2"}, // 2026-01-03
	} {
		after := edited(t, before, filepath.Join(dir, "a.xlsx"), []string{"rows", "delete", c.row})
		r, code := against(t, after, before)
		if m := causes(r)[c.cell]; !strings.HasPrefix(m, "rows removed: "+c.row) || len(r.Unexplained) != 0 || code != 0 {
			t.Errorf("%s: %s: exit %d, cause %q, unexplained %+v, moved %+v", c.name, c.cell, code, m, r.Unexplained, r.Moved)
		}
	}
	// A row that held neither date removed, and the MAX stored as another
	// date: no work of the row.
	deleted := edited(t, before, filepath.Join(dir, "d.xlsx"), []string{"rows", "delete", "Data!2"})
	after := filepath.Join(dir, "w.xlsx")
	stored(t, deleted, after, "B1", "47000")
	if r, code := against(t, after, before); !unexplainedAt(r, "Summary!B1") || code != 1 {
		t.Errorf("a MAX stored as another date: exit %d, unexplained %+v, moved %+v", code, r.Unexplained, r.Moved)
	}
}

// An array formula in rows a sort moved (the fuzz pool's XlsxWriter file, whose
// side columns hold one over two rows): it reads another row besides its own,
// so what the sort moved moved its result (f2.5), with rows inserted above
// the table or none.
func TestAgainstCausesArraySorted(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("SHEETS_CACHE_DIR", filepath.Join(dir, ".cache"))
	before := filepath.Join(dir, "b.xlsx")
	if err := os.WriteFile(before, mustRead(t, filepath.Join("..", "fuzz", "testdata", "pool", "xlsxwriter-05.xlsx")), 0o644); err != nil {
		t.Fatal(err)
	}
	for _, c := range []struct {
		name string
		cmds [][]string
		cell string
	}{
		{"sorted", [][]string{{"sort", "Data!A1:F12", "B:desc"}}, "Data!K3"},
		{"a row inserted above", [][]string{{"rows", "insert", "Data!1"}, {"sort", "Data!A2:F13", "B:desc"}}, "Data!K4"},
	} {
		after := edited(t, before, filepath.Join(dir, "a.xlsx"), c.cmds...)
		r, code := against(t, after, before)
		// (Exit 1 is check's own findings on the file: an error value in it.)
		if m := causes(r)[c.cell]; !strings.HasPrefix(m, "rows moved: Data!") || len(r.Unexplained) != 0 || code != 1 {
			t.Errorf("%s: %s: exit %d, cause %q, unexplained %+v", c.name, c.cell, code, m, r.Unexplained)
		}
	}
}

// The span a sort moved takes in the rows between where each moved row was
// and where it went, which hold another row's now though they did not move,
// as numbered now: with rows inserted above or deleted inside as without (the
// fuzz pool's openpyxl file, whose side column reads each row's own cell).
func TestAgainstCausesSortSpanBetween(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("SHEETS_CACHE_DIR", filepath.Join(dir, ".cache"))
	before := filepath.Join(dir, "b.xlsx")
	if err := os.WriteFile(before, mustRead(t, filepath.Join("..", "fuzz", "testdata", "pool", "openpyxl-24.xlsx")), 0o644); err != nil {
		t.Fatal(err)
	}
	for _, c := range []struct {
		name string
		cmds [][]string
	}{
		{"sorted, one value set", [][]string{{"sort", "Data!A1:F12", "B:desc"}, {"set", "Data!B4", "55"}}},
		{"a row inserted above, sorted", [][]string{{"rows", "insert", "Data!1"}, {"sort", "Data!A2:F13", "B:desc"}}},
		{"a row deleted inside, sorted", [][]string{{"rows", "delete", "Data!5"}, {"sort", "Data!A1:F11", "B:desc"}}},
	} {
		after := edited(t, before, filepath.Join(dir, "a.xlsx"), c.cmds...)
		r, _ := against(t, after, before)
		if len(r.Unexplained) != 0 {
			t.Errorf("%s: unexplained %+v", c.name, r.Unexplained)
		}
	}
}

// Two tables on one sheet, each sorted apart (two sort commands, or a batch
// with two): each run of rows moved is a span of its own, so a total over the
// first table is a total over all of the rows it moved, and a result stored
// wrong for it stays unexplained as it does after one sort (f4.1); an input
// set after the sorts is named alone.
func TestAgainstCausesSortApart(t *testing.T) {
	dir, fixture := causesFixture(t)
	before := edited(t, fixture, filepath.Join(dir, "two.xlsx"), []string{"set", "Data!A9", "Item", "Data!B9", "Stock", "Data!A10", "Pens", "Data!B10", "12",
		"Data!A11", "Ink", "Data!B11", "2", "Data!A12", "Paper", "Data!B12", "30", "Data!A13", "Tape", "Data!B13", "6",
		"Summary!C1", "=SUM(Data!B10:B13)", "Summary!C2", "=MAX(Data!B10:B13)"})
	sorted := edited(t, before, filepath.Join(dir, "s.xlsx"), []string{"sort", "Data!A1:D6", "Units:desc"}, []string{"sort", "Data!A9:B13", "Stock:desc"})
	r, code := against(t, sorted, before)
	if r.Summary.MovedRows < 2 || code != 0 || len(r.Unexplained) != 0 {
		t.Errorf("both sorted: exit %d, %+v, unexplained %+v", code, r.Summary, r.Unexplained)
	}
	for _, c := range []string{"Summary!B2", "Summary!B6", "Summary!C1", "Summary!C2"} {
		if m, ok := causes(r)[c]; ok {
			t.Errorf("both sorted: %s (a total) moved: %q", c, m)
		}
	}
	// A total over either table stored wrong, with one sort or both: unexplained.
	one := edited(t, before, filepath.Join(dir, "one.xlsx"), []string{"sort", "Data!A1:D6", "Units:desc"})
	for _, src := range []string{one, sorted} {
		for _, cell := range []string{"B2", "B6", "C1", "C2"} {
			after := filepath.Join(dir, "a.xlsx")
			stored(t, src, after, cell, "999")
			r, code := against(t, after, before)
			if m := causes(r)["Summary!"+cell]; !unexplainedAt(r, "Summary!"+cell) || m != "" || code != 1 {
				t.Errorf("%s stored as 999 (%s): exit %d, cause %q, unexplained %+v", cell, filepath.Base(src), code, m, r.Unexplained)
			}
		}
	}
	// An input set in the first table after both sorts: the totals over it name
	// the input, not the sorts; those over the second are as they were.
	set := edited(t, before, filepath.Join(dir, "set.xlsx"), []string{"sort", "Data!A1:D6", "Units:desc"}, []string{"sort", "Data!A9:B13", "Stock:desc"}, []string{"set", "Data!B3", "11"})
	r, code = against(t, set, before)
	m := causes(r)
	for _, c := range []string{"Summary!B2", "Summary!B3", "Summary!B6"} {
		if m[c] != "input Data!B3" {
			t.Errorf("an input set after two sorts: %s: cause %q, want input Data!B3", c, m[c])
		}
	}
	for _, c := range []string{"Summary!C1", "Summary!C2"} {
		if _, ok := m[c]; ok {
			t.Errorf("an input set after two sorts: %s moved: %q", c, m[c])
		}
	}
	if code != 0 || len(r.Unexplained) != 0 {
		t.Errorf("an input set after two sorts: exit %d, unexplained %+v", code, r.Unexplained)
	}
	// The same two sorts in one batch.
	batch := filepath.Join(dir, "batch.xlsx")
	if err := os.WriteFile(batch, mustRead(t, before), 0o644); err != nil {
		t.Fatal(err)
	}
	if out, code := runCLI("sort Data!A1:D6 Units:desc\nsort Data!A9:B13 Stock:desc\n", "do", batch); code != 0 {
		t.Fatalf("do: %s", out)
	}
	after := filepath.Join(dir, "a.xlsx")
	stored(t, batch, after, "B2", "999")
	if r, code := against(t, after, before); !unexplainedAt(r, "Summary!B2") || code != 1 {
		t.Errorf("batch: exit %d, unexplained %+v, moved %+v", code, r.Unexplained, r.Moved)
	}
}

// A total over sorted rows written as arithmetic over ranges inside the
// function that adds them up (the common conditional sums: SUMPRODUCT((A2:A6=
// "North")*B2:B6), SUMPRODUCT(B2:B6*C2:C6), SUM(B2:B6*C2:C6)) is no work of the
// sort either: no reordering of whole rows moves it (f4.2). Stored wrong it is
// unexplained; with an input set among the rows, the input is its cause alone.
func TestAgainstCausesSortElementwise(t *testing.T) {
	dir, fixture := causesFixture(t)
	forms := map[string]string{
		"F1": `=SUMPRODUCT((Data!A2:A6="North")*Data!B2:B6)`,
		"F2": `=SUMPRODUCT(Data!B2:B6*Data!C2:C6)`,
		"F3": `=SUM(Data!B2:B6*Data!C2:C6)`,
		"F4": `=SUMPRODUCT(--(Data!B2:B6>6),Data!C2:C6)`,
		"F5": `=SUMPRODUCT((Data!B2:B6>6)*Data!C2:C6*Summary!$B$1)`,
		"F6": `=MAX(Data!B2:B6*2)`,
	}
	var cmd []string
	cmd = append(cmd, "set")
	for _, c := range []string{"F1", "F2", "F3", "F4", "F5", "F6"} {
		cmd = append(cmd, "Summary!"+c, forms[c])
	}
	before := edited(t, fixture, filepath.Join(dir, "forms.xlsx"), cmd)
	sorted := edited(t, before, filepath.Join(dir, "s.xlsx"), []string{"sort", "Data!A1:D6", "Units:desc"})
	r, code := against(t, sorted, before)
	for c := range forms {
		if m, ok := causes(r)["Summary!"+c]; ok {
			t.Errorf("sorted: %s moved: %q", c, m)
		}
	}
	if code != 0 || len(r.Unexplained) != 0 || r.Summary.MovedRows == 0 {
		t.Errorf("sorted: exit %d, %+v, unexplained %+v", code, r.Summary, r.Unexplained)
	}
	for c := range forms {
		after := filepath.Join(dir, "a.xlsx")
		stored(t, sorted, after, c, "999")
		r, code := against(t, after, before)
		if m := causes(r)["Summary!"+c]; !unexplainedAt(r, "Summary!"+c) || m != "" || code != 1 {
			t.Errorf("%s (%s) stored as 999: exit %d, cause %q, unexplained %+v", c, forms[c], code, m, r.Unexplained)
		}
	}
	set := edited(t, before, filepath.Join(dir, "set.xlsx"), []string{"sort", "Data!A1:D6", "Units:desc"}, []string{"set", "Data!B4", "11"})
	r, code = against(t, set, before)
	for c := range forms {
		if m := causes(r)["Summary!"+c]; m != "input Data!B4" {
			t.Errorf("%s (%s) with an input set: cause %q, want input Data!B4", c, forms[c], m)
		}
	}
	if code != 0 || len(r.Unexplained) != 0 {
		t.Errorf("an input set: exit %d, unexplained %+v", code, r.Unexplained)
	}
	// A range read by position among them is still the sort's: the first row of
	// a column multiplied, not the column as a set.
	pos := edited(t, before, filepath.Join(dir, "pos.xlsx"), []string{"set", "Summary!G1", "=SUMPRODUCT(Data!B2:B6*Data!C3:C7)", "Summary!G2", "=SUMPRODUCT(Data!B2:B6*Data!C2)"})
	posSorted := edited(t, pos, filepath.Join(dir, "ps.xlsx"), []string{"sort", "Data!A1:D6", "Units:desc"})
	r, _ = against(t, posSorted, pos)
	for _, c := range []string{"Summary!G1", "Summary!G2"} {
		if m := causes(r)[c]; !strings.HasPrefix(m, "rows moved: ") {
			t.Errorf("%s: cause %q, want the sort", c, m)
		}
	}
}

// A dropdown lost from one range and another added on other cells (the classic
// slip of a script that replaces the sheet's validations while adding one) is
// one lost and one added, not the old one edited (f4.4): a rule is changed
// only on the cells it covered. A highlight the same.
func TestAgainstCausesFeaturesLostElsewhere(t *testing.T) {
	dir, fixture := causesFixture(t)
	before := edited(t, fixture, filepath.Join(dir, "lists.xlsx"), []string{"set", "Data!H2", "North", "Data!H3", "South", "Data!J2", "Open", "Data!J3", "Closed"},
		[]string{"validate", "Data!A2:A6", "list", "=$H$2:$H$3"})
	cleared := edited(t, before, filepath.Join(dir, "cleared.xlsx"))
	out, code := runCLI("", "validate", cleared, "Data!A2:A6", "--clear")
	if code != 0 || !strings.Contains(head(t, out), ", 1 feature lost") {
		t.Fatalf("the clear alone: exit %d:\n%s", code, out)
	}
	after := edited(t, before, filepath.Join(dir, "a.xlsx"), []string{"validate", "Data!A2:A6", "--clear"}, []string{"validate", "Data!E2:E6", "list", "=$J$2:$J$3"})
	r, code := against(t, after, before)
	if len(r.FeaturesLost) != 1 || r.FeaturesLost[0].Where != "Data!A2:A6" || len(r.FeaturesChanged) != 0 || r.Summary.FeaturesLost != 1 || code != 1 {
		t.Errorf("a dropdown cleared, another added: exit %d, lost %+v, changed %+v", code, r.FeaturesLost, r.FeaturesChanged)
	}
	if text, _ := runCLI("", "check", after, "--against", before); !strings.Contains(text, "1 feature lost") || strings.Contains(text, "features changed") {
		t.Errorf("text:\n%s", text)
	}
	// A highlight cleared on some cells and the same kind added on others.
	hi := func(sqref, f string) func([]byte) []byte {
		return func(b []byte) []byte {
			return bytes.Replace(b, []byte("</conditionalFormatting>"), []byte(`</conditionalFormatting><conditionalFormatting sqref="`+sqref+`"><cfRule type="expression" dxfId="0" priority="2"><formula>`+f+`</formula></cfRule></conditionalFormatting>`), 1)
		}
	}
	had := filepath.Join(dir, "had.xlsx")
	rewriteSheet(t, fixture, had, hi("A2:C6", "$B2&gt;6"))
	now := filepath.Join(dir, "now.xlsx")
	rewriteSheet(t, fixture, now, hi("F2:H6", "$G2&gt;6"))
	r, code = against(t, now, had)
	if len(r.FeaturesLost) != 1 || r.FeaturesLost[0].Where != "Data!A2:C6" || len(r.FeaturesChanged) != 0 || code != 1 {
		t.Errorf("a highlight on other cells: exit %d, lost %+v, changed %+v", code, r.FeaturesLost, r.FeaturesChanged)
	}
	// An edit on the same cells is still an edit.
	same := filepath.Join(dir, "same.xlsx")
	rewriteSheet(t, fixture, same, hi("A2:C6", "$C2&gt;6"))
	changedOK(t, "highlight on the same cells", same, had)
}

// Several rules of a kind on the same cells, some cleared and one edited: the
// edit is paired with the rule it resembles, and the one lost is named by its
// operator and value (f4.5).
func TestAgainstCausesFeaturesClosestRule(t *testing.T) {
	dir, fixture := causesFixture(t)
	before := edited(t, fixture, filepath.Join(dir, "two.xlsx"), []string{"cond", "Data!B2:B6", "> 6"}, []string{"cond", "Data!B2:B6", "< 4", "fill=lightgreen"})
	after := edited(t, before, filepath.Join(dir, "a.xlsx"), []string{"cond", "Data!B2:B6", "--clear"}, []string{"cond", "Data!B2:B6", "< 5", "fill=lightgreen"})
	r, code := against(t, after, before)
	if len(r.FeaturesChanged) != 1 || r.FeaturesChanged[0].Before != "cellIs lessThan 4" || r.FeaturesChanged[0].After != "cellIs lessThan 5" {
		t.Errorf("changed %+v", r.FeaturesChanged)
	}
	if len(r.FeaturesLost) != 1 || !strings.Contains(r.FeaturesLost[0].What, "cellIs greaterThan 6") || code != 1 {
		t.Errorf("exit %d, lost %+v", code, r.FeaturesLost)
	}
	// The other way about: the greater-than edited, the lesser cleared.
	after = edited(t, before, filepath.Join(dir, "b.xlsx"), []string{"cond", "Data!B2:B6", "--clear"}, []string{"cond", "Data!B2:B6", "> 7"})
	r, code = against(t, after, before)
	if len(r.FeaturesChanged) != 1 || r.FeaturesChanged[0].Before != "cellIs greaterThan 6" || r.FeaturesChanged[0].After != "cellIs greaterThan 7" {
		t.Errorf("changed %+v", r.FeaturesChanged)
	}
	if len(r.FeaturesLost) != 1 || !strings.Contains(r.FeaturesLost[0].What, "cellIs lessThan 4") || code != 1 {
		t.Errorf("exit %d, lost %+v", code, r.FeaturesLost)
	}
	text, _ := runCLI("", "check", after, "--against", before)
	if !strings.Contains(text, "cell value rule: cellIs lessThan 4") {
		t.Errorf("text:\n%s", text)
	}
}

// A value reached by a structure change only through a formula the change did
// not move is no work of the change (f4.3): the tracing stops at a formula
// whose value held, and passes only the formulas the lines could move. An
// empty row inserted or removed moves no total, so the tax on the total,
// stored as another number, is unexplained, and what reads it follows it; so
// is a MAX that held after an input changed, and what reads it. A row with
// numbers in it still moves the total, and the tax with it.
func TestAgainstCausesChainHeld(t *testing.T) {
	dir, fixture := causesFixture(t)
	// An empty row inserted: Summary!B3 (=SUM(Data!D2:D7)) holds 630.
	inserted := edited(t, fixture, filepath.Join(dir, "i.xlsx"), []string{"rows", "insert", "Data!4"})
	after := filepath.Join(dir, "a.xlsx")
	stored(t, inserted, after, "B4", "999")
	r, code := against(t, after, fixture)
	if !unexplainedAt(r, "Summary!B4") || code != 1 || len(r.Moved) != 0 || len(r.Unknown) != 0 {
		t.Errorf("an empty row inserted, the tax stored as 999: exit %d, unexplained %+v, moved %+v, unknown %+v", code, r.Unexplained, r.Moved, r.Unknown)
	}
	if len(r.Unexplained) == 1 && strings.Join(r.Unexplained[0].Held, " ") != "Summary!B5" {
		t.Errorf("the cell that reads it and held: %+v", r.Unexplained[0])
	}
	text, _ := runCLI("", "check", after, "--against", fixture)
	if !strings.Contains(text, "Summary!B4  157.5 → 999  read by Summary!B5, which holds what it held") {
		t.Errorf("text:\n%s", text)
	}
	// The total itself stored wrong, and the tax following it: the total is
	// unexplained, and the tax follows it, not the row.
	mid := filepath.Join(dir, "m.xlsx")
	stored(t, inserted, mid, "B3", "999")
	stored(t, mid, after, "B4", "249.75")
	r, code = against(t, after, fixture)
	if len(r.Unexplained) != 1 || r.Unexplained[0].Cell != "Summary!B3" || code != 1 {
		t.Fatalf("the total stored as 999: exit %d, unexplained %+v, moved %+v", code, r.Unexplained, r.Moved)
	}
	if got := causes(r)["Summary!B4"]; got != "Summary!B3, which moved with nothing upstream changed" {
		t.Errorf("the tax that followed the total: cause %q", got)
	}
	for _, u := range r.Unknown {
		t.Errorf("unknown %+v", u)
	}
	// An empty row removed, as the file had one: the same.
	gap := edited(t, fixture, filepath.Join(dir, "gap.xlsx"), []string{"rows", "insert", "Data!4"})
	removed := edited(t, gap, filepath.Join(dir, "r.xlsx"), []string{"rows", "delete", "Data!4"})
	stored(t, removed, after, "B4", "999")
	if r, code := against(t, after, gap); !unexplainedAt(r, "Summary!B4") || code != 1 || len(r.Moved) != 0 || len(r.Unknown) != 0 {
		t.Errorf("an empty row removed, the tax stored as 999: exit %d, unexplained %+v, moved %+v, unknown %+v", code, r.Unexplained, r.Moved, r.Unknown)
	}
	stored(t, removed, mid, "B3", "999")
	stored(t, mid, after, "B4", "249.75")
	r, code = against(t, after, gap)
	if len(r.Unexplained) != 1 || r.Unexplained[0].Cell != "Summary!B3" || causes(r)["Summary!B4"] != "Summary!B3, which moved with nothing upstream changed" || code != 1 {
		t.Errorf("an empty row removed, the total stored as 999: exit %d, unexplained %+v, moved %+v", code, r.Unexplained, r.Moved)
	}
	// An input that leaves a MAX as it was: what reads the MAX did not move for it.
	base := edited(t, fixture, filepath.Join(dir, "base.xlsx"), []string{"set", "Summary!F1", "=B6*2"})
	set := edited(t, base, filepath.Join(dir, "set.xlsx"), []string{"set", "Data!B2", "6"})
	stored(t, set, after, "F1", "999")
	r, code = against(t, after, base)
	if !unexplainedAt(r, "Summary!F1") || code != 1 || causes(r)["Summary!F1"] != "" {
		t.Errorf("an input that left the MAX as it was, what reads it stored as 999: exit %d, unexplained %+v, moved %+v", code, r.Unexplained, r.Moved)
	}
	if got := causes(r)["Summary!B2"]; got != "input Data!B2" {
		t.Errorf("the total the input moved: cause %q", got)
	}
	if r, code := against(t, set, base); len(r.Unexplained) != 0 || code != 0 || len(causes(r)) == 0 {
		t.Errorf("the input alone: exit %d, unexplained %+v", code, r.Unexplained)
	}
	// A row with numbers moves the total, and the tax that reads it with it.
	filled := edited(t, fixture, filepath.Join(dir, "f.xlsx"), []string{"rows", "insert", "Data!4"}, []string{"set", "Data!B4", "11", "Data!C4", "10", "Data!D4", "=B4*C4"})
	r, code = against(t, filled, fixture)
	for _, c := range []string{"Summary!B3", "Summary!B4"} {
		if m := causes(r)[c]; !strings.HasPrefix(m, "rows added: Data!4") {
			t.Errorf("a row with numbers: %s cause %q", c, m)
		}
	}
	if len(r.Unexplained) != 0 || code != 0 {
		t.Errorf("a row with numbers: exit %d, unexplained %+v", code, r.Unexplained)
	}
}

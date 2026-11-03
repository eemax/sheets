package xlsx

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"testing"

	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/input"
	"github.com/eemax/sheets/internal/xlsxtest"
	"github.com/xuri/excelize/v2"
)

// bothEditors runs a test with the fix worked out by the streaming editor
// and by the full one.
func bothEditors(t *testing.T, fn func(t *testing.T)) {
	for _, mode := range []string{"always", "never"} {
		t.Run("stream="+mode, func(t *testing.T) {
			t.Setenv("SHEETS_STREAM", mode)
			fn(t)
		})
	}
}

func checkData(t *testing.T, name string, data []byte) *CheckResult {
	t.Helper()
	res, err := CheckResults(name, data, CheckOptions{})
	if err != nil {
		t.Fatal(err)
	}
	if res.FixError != "" {
		t.Fatal(res.FixError)
	}
	return res
}

// changesOf lists a check's changes as "kind cell", in order.
func changesOf(res *CheckResult) map[string]Change {
	out := map[string]Change{}
	for _, c := range res.Changes {
		out[c.Kind+" "+grid.CellName(c.Row, c.Col)] = c
	}
	return out
}

// fixedTwice checks that a fixed file reads, keeps the results Excel
// vouched for, and that fixing it again changes nothing.
func fixedTwice(t *testing.T, before []byte, res *CheckResult) {
	t.Helper()
	if res.Data == nil {
		return
	}
	b, err := ReadBytes(before)
	must(t, err)
	a, err := ReadBytes(res.Data)
	must(t, err)
	if err := xlsxtest.CheckOrder(res.Data); err != nil {
		t.Errorf("cells out of order: %v", err)
	}
	if v, o := promisesKept(b, a); len(v)+len(o) > 0 {
		t.Errorf("changed what it must not: %v %v", v, o)
	}
	if again := checkData(t, "again.xlsx", res.Data); again.Data != nil {
		t.Errorf("a second fix changes %+v", again.Changes)
	}
}

// Results Excel stored are kept, even one the calculator disputes (a
// difference to note), and nothing computed from one is stored; a
// missing result is stored.
func TestCheckKeepsExcelResults(t *testing.T) {
	f := excelize.NewFile()
	f.SetCellValue("Sheet1", "A1", 5)
	for cell, text := range map[string]string{"A2": "A1*2", "A3": "A1+1", "A4": "A2+1", "A5": "SUM(A1:A2)"} {
		f.SetCellFormula("Sheet1", cell, text)
	}
	var buf bytes.Buffer
	must(t, f.Write(&buf))
	data := patchZip(t, buf.Bytes(), dataPart, func(x string) string {
		for cell, v := range map[string]string{"A2": "999", "A5": "1004"} {
			x = regexp.MustCompile(`<c r="`+cell+`"[^>]*>(<f>[^<]*</f>)(<v>[^<]*</v>)?`).ReplaceAllString(x, `<c r="`+cell+`">${1}<v>`+v+`</v>`)
		}
		return x
	})
	bothEditors(t, func(t *testing.T) {
		res := checkData(t, "x.xlsx", data)
		if res.Formulas != 4 || res.Vouched != 2 || res.Missing != 2 || res.Stale != 0 {
			t.Errorf("counts: %+v", res)
		}
		ch := changesOf(res)
		if c, ok := ch["stored A3"]; !ok || c.After.Num != 6 || len(ch) != 1 {
			t.Errorf("changes: %v", ch)
		}
		if len(res.Unknown) != 1 || res.Unknown[0].Row != 4 || !strings.Contains(res.Unknown[0].Why, "where Excel stored 999") {
			t.Errorf("unknown: %+v", res.Unknown)
		}
		if len(res.Differ) != 1 || res.Differ[0].Row != 2 || res.Differ[0].Calc.Num != 10 || res.Differ[0].Excel.Num != 999 {
			t.Errorf("differences: %+v", res.Differ)
		}
		fixedTwice(t, data, res)
	})
}

// openpyxl's files: no results, no prefixes. XlsxWriter's: a 0 for every
// formula, out of date; what sheets cannot compute loses it.
func TestCheckOtherProducers(t *testing.T) {
	for _, tc := range []struct {
		file string
		want map[string]string // change → value after
	}{
		{"unprefixed.xlsx", map[string]string{"prefixed C2": "", "prefixed D2": "", "prefixed E2": "", "stored C2": `"PK1"`, "stored D2": `"PK1,XX2,PK3"`, "stored E2": "30", "stored F2": "60"}},
		{"stale.xlsx", map[string]string{"prefixed C3": "", "replaced C2": "15", "replaced C3": "5", "dropped C4": "nothing", "dropped C5": "nothing",
			"replaced C6": "#DIV/0!", "replaced C7": "83", "replaced C8": "30"}},
		{"calccol.xlsx", map[string]string{"stored F2": "10", "stored F3": "26", "stored F4": "36", "stored H2": "36"}},
		// Dynamic arrays spill over the cells below them, in either editor.
		{"dynamic_xlsxwriter.xlsx", map[string]string{"replaced D1": `"North"`, "replaced E1": "16", "replaced F1": "7",
			"filled D2": `"South"`, "filled D3": `"East"`, "filled F2": "5", "filled F3": "3", "filled F4": "1"}},
	} {
		data, err := os.ReadFile("testdata/" + tc.file)
		must(t, err)
		t.Run(tc.file, func(t *testing.T) {
			bothEditors(t, func(t *testing.T) {
				res := checkData(t, tc.file, data)
				ch := changesOf(res)
				for k, want := range tc.want {
					c, ok := ch[k]
					if !ok {
						t.Errorf("no change %s (changes: %v)", k, ch)
						continue
					}
					if want != "" && show1(c.After) != want {
						t.Errorf("%s: %s, want %s", k, show1(c.After), want)
					}
				}
				if len(ch) != len(tc.want) {
					t.Errorf("%d changes, want %d: %v", len(ch), len(tc.want), ch)
				}
				fixedTwice(t, data, res)
			})
		})
	}
}

// A file Excel saved needs nothing: no file comes of it.
func TestCheckNothingToFix(t *testing.T) {
	data, err := os.ReadFile("testdata/cached.xlsx")
	must(t, err)
	bothEditors(t, func(t *testing.T) {
		res := checkData(t, "cached.xlsx", data)
		if res.Data != nil || len(res.Changes) != 0 || res.Vouched != 6 || res.Compared == 0 || res.Agreed != res.Compared {
			t.Errorf("%+v", res)
		}
	})
}

// Adding a prefix changes the formula, which Excel computed as #NAME?:
// where a result Excel vouched for reads it, adding it would change that
// result, so the fix stores results only, and says why.
func TestCheckWithholdsPrefixes(t *testing.T) {
	f := excelize.NewFile()
	for r := 1; r <= 3; r++ {
		f.SetCellValue("Sheet1", grid.CellName(r, 1), r)
	}
	f.SetCellFormula("Sheet1", "D1", "AGGREGATE(9,6,A1:A3)")
	f.SetCellFormula("Sheet1", "D2", "IFERROR(D1,0)")
	f.SetCellFormula("Sheet1", "D3", "A1+A2")
	var buf bytes.Buffer
	must(t, f.Write(&buf))
	data := patchZip(t, buf.Bytes(), dataPart, func(x string) string {
		x = regexp.MustCompile(`<c r="D1"[^>]*><f>([^<]*)</f>(<v>[^<]*</v>)?`).ReplaceAllString(x, `<c r="D1" t="e"><f>$1</f><v>#NAME?</v>`)
		x = regexp.MustCompile(`<c r="D2"[^>]*><f>([^<]*)</f>(<v>[^<]*</v>)?`).ReplaceAllString(x, `<c r="D2"><f>$1</f><v>0</v>`)
		return regexp.MustCompile(`<c r="D3"[^>]*><f>([^<]*)</f>(<v>[^<]*</v>)?`).ReplaceAllString(x, `<c r="D3"><f>$1</f>`)
	})
	bothEditors(t, func(t *testing.T) {
		res := checkData(t, "x.xlsx", data)
		if !strings.Contains(res.Withheld, "Sheet1!D2") {
			t.Errorf("withheld: %q", res.Withheld)
		}
		ch := changesOf(res)
		if _, ok := ch["stored D3"]; !ok {
			t.Errorf("changes: %v", ch)
		}
		if _, ok := ch["prefixed D1"]; ok {
			t.Errorf("prefixed D1 anyway")
		}
		fixedTwice(t, data, res)
		if res.Data != nil {
			a, _ := ReadBytes(res.Data)
			if c := a.Sheet("Sheet1").Get(2, 4); c.Kind != grid.Number || c.Num != 0 {
				t.Errorf("D2, vouched for, now %+v", c)
			}
		}
	})
}

// The comparison of a fix with the file refuses changes it must not make.
func TestCheckRefusesBrokenPromises(t *testing.T) {
	wb := func(v float64, pending bool) *grid.Workbook {
		s := grid.NewSheet("S")
		s.Set(1, 1, grid.Cell{Kind: grid.Number, Num: 1})
		s.Set(1, 2, grid.Cell{Kind: grid.Number, Num: v, Formula: "A1*2", Pending: pending})
		return &grid.Workbook{Sheets: []*grid.Sheet{s}}
	}
	res := &CheckResult{}
	if err := compareFiles(wb(2, false), wb(3, false), nil, res); err == nil || !strings.Contains(err.Error(), "S!B1 result 2 → 3") {
		t.Errorf("a vouched result changed: %v", err)
	}
	other := wb(2, false)
	other.Sheets[0].Set(1, 1, grid.Cell{Kind: grid.Number, Num: 7})
	if err := compareFiles(wb(2, false), other, nil, res); err == nil {
		t.Errorf("a value changed, accepted")
	}
	res = &CheckResult{}
	if err := compareFiles(wb(0, true), wb(2, false), nil, res); err != nil || len(res.Changes) != 1 || res.Changes[0].Kind != "replaced" {
		t.Errorf("a stale result replaced: %v %+v", err, res.Changes)
	}
}

// The streaming editor computes a formula entered as an array formula as
// one (it stored #VALUE! for =SUM(B2:B4*B2:B4)).
func TestStreamArrayFormula(t *testing.T) {
	t.Setenv("SHEETS_STREAM", "always")
	path := fixture(t, "stale.xlsx")
	data, err := os.ReadFile(path)
	must(t, err)
	a, err := OpenAuto(path, data)
	must(t, err)
	must(t, a.Set("Data", 2, 2, input.Parse("6", false)))
	must(t, a.Save(false))
	if got := stored(t, path, dataPart, "C7"); got != "94" {
		t.Errorf("C7 stores %q, want 94", got)
	}
}

// A link to another workbook only Excel can update: a result of it already
// out of date is removed rather than vouched for, and what reads it stays
// unknown however often the fix runs.
func TestCheckLinkedResults(t *testing.T) {
	f := excelize.NewFile()
	f.SetCellValue("Sheet1", "A1", 2)
	f.SetCellFormula("Sheet1", "B1", "[1]Rates!A1*A1")
	f.SetCellFormula("Sheet1", "B2", "A1*3")
	f.SetCellFormula("Sheet1", "B3", "B1+1")
	on := true
	must(t, f.SetCalcProps(&excelize.CalcPropsOptions{FullCalcOnLoad: &on}))
	var buf bytes.Buffer
	must(t, f.Write(&buf))
	data := patchZip(t, buf.Bytes(), dataPart, func(x string) string {
		for cell, v := range map[string]string{"B1": "7", "B2": "0", "B3": "0"} {
			x = regexp.MustCompile(`<c r="`+cell+`"[^>]*>(<f>[^<]*</f>)(<v>[^<]*</v>)?`).ReplaceAllString(x, `<c r="`+cell+`">${1}<v>`+v+`</v>`)
		}
		return x
	})
	bothEditors(t, func(t *testing.T) {
		res := checkData(t, "x.xlsx", data)
		ch := changesOf(res)
		if c, ok := ch["replaced B2"]; !ok || c.After.Num != 6 || len(ch) != 3 {
			t.Errorf("changes: %v", ch)
		}
		if c, ok := ch["dropped B1"]; !ok || !strings.Contains(c.Why, "another workbook") {
			t.Errorf("B1: %+v", c)
		}
		if _, ok := ch["dropped B3"]; !ok {
			t.Errorf("B3 kept a result computed from a stale link: %v", ch)
		}
		if len(res.Unknown) != 2 {
			t.Errorf("unknown: %+v", res.Unknown)
		}
		fixedTwice(t, data, res)
		again := checkData(t, "x.xlsx", res.Data)
		if len(again.Unknown) != 2 || again.Missing != 2 {
			t.Errorf("second run: %d unknown, %d missing", len(again.Unknown), again.Missing)
		}
		a, _ := ReadBytes(res.Data)
		if c := a.Value(a.Sheet("Sheet1"), 1, 2); !c.Pending {
			t.Errorf("B1 reads as known: %+v", c)
		}
	})
}

// openFixture reads a fixture's bytes.
func openFixture(t *testing.T, name string) []byte {
	t.Helper()
	data, err := os.ReadFile("testdata/" + name)
	must(t, err)
	return data
}

// valuesAt reads a fixed file's values of cells (a range of one sheet).
func valuesAt(t *testing.T, data []byte, sheet string, rg grid.Range) []string {
	t.Helper()
	wb, err := ReadBytes(data)
	must(t, err)
	var out []string
	s := wb.Sheet(sheet)
	for r := rg.R1; r <= rg.R2; r++ {
		for c := rg.C1; c <= rg.C2; c++ {
			c := s.Get(r, c)
			if c.Pending {
				out = append(out, "?")
				continue
			}
			out = append(out, show1(valueOnly(c)))
		}
	}
	return out
}

// Array formulas entered with Ctrl+Shift+Enter over several cells, as
// openpyxl and XlsxWriter write them: every cell of the range gets its
// value, in both editors; a prefix missing from one is added, the array
// kept.
func TestCheckLegacyArrays(t *testing.T) {
	op, xw := openFixture(t, "arrays_openpyxl.xlsx"), openFixture(t, "arrays_xlsxwriter.xlsx")
	bothEditors(t, func(t *testing.T) {
		res := checkData(t, "op.xlsx", op)
		for rg, want := range map[string]string{"D1:D4": "[10 40 90 160]", "F1:G1": "[1 2]", "H1:H3": "[10 10 10]", "J1:J3": `["" "big" "big"]`, "K1:K3": `["b" "a" "a"]`, "L1:L3": `["x" "y" "y"]`} {
			r, _ := grid.ParseRange(rg)
			if got := fmt.Sprint(valuesAt(t, res.Data, "S", r)); got != want {
				t.Errorf("openpyxl %s: %s, want %s", rg, got, want)
			}
		}
		if len(res.Unknown) != 0 {
			t.Errorf("unknown: %+v", res.Unknown)
		}
		fixedTwice(t, op, res)

		res = checkData(t, "xw.xlsx", xw)
		r, _ := grid.ParseRange("E2:E4")
		if got := fmt.Sprint(valuesAt(t, res.Data, "S", r)); got != "[10 40 90]" {
			t.Errorf("XlsxWriter E2:E4: %s", got)
		}
		if _, ok := changesOf(res)["prefixed G2"]; !ok {
			t.Errorf("G2 not prefixed: %v", changesOf(res))
		}
		a, _ := ReadBytes(res.Data)
		if c := a.Sheet("S").Get(2, 7); c.Unprefixed || c.Pending || c.Num != 2 {
			t.Errorf("G2: %+v", c)
		}
		if len(a.Sheet("S").Arrays) != 2 {
			t.Errorf("arrays now %v", a.Sheet("S").Arrays)
		}
		fixedTwice(t, xw, res)
	})
}

// In a workbook counting dates from 1904, what turns serials into dates
// is unknown rather than computed in the 1900 system; plain arithmetic
// and TEXT, which knows the system, are computed.
func TestCheck1904(t *testing.T) {
	data := openFixture(t, "dates1904.xlsx") // A1 2020-01-01; B1 =YEAR(A1), C1 =DATE(2020,1,1), D1 =A1+1, E1 =TEXT(A1,"yyyy-mm-dd")
	bothEditors(t, func(t *testing.T) {
		res := checkData(t, "d.xlsx", data)
		if got := fmt.Sprint(valuesAt(t, res.Data, "S", grid.Range{R1: 1, C1: 2, R2: 1, C2: 5})); got != `[? ? 42370 "2020-01-01"]` {
			t.Errorf("B1:E1: %s", got)
		}
		if len(res.Unknown) != 2 || !strings.Contains(res.Unknown[0].Why, "1904") {
			t.Errorf("unknown: %+v", res.Unknown)
		}
		fixedTwice(t, data, res)
	})
}

// The editors agree on volatile formulas (TODAY), INDIRECT, and a ROWS
// over a range holding the formula itself, which is not circular.
func TestCheckRefsBothEditors(t *testing.T) {
	data := openFixture(t, "refs.xlsx") // A2 =ROWS($A$2:A2), A3 =ROWS($A$2:A3), B2 =COLUMNS(A2:B2), C2 TODAY, C3 NOW, C4 INDIRECT("A2")
	bothEditors(t, func(t *testing.T) {
		res := checkData(t, "r.xlsx", data)
		if len(res.Unknown) != 0 {
			t.Errorf("unknown: %+v", res.Unknown)
		}
		if got := fmt.Sprint(valuesAt(t, res.Data, "S", grid.Range{R1: 2, C1: 1, R2: 3, C2: 1})); got != "[1 2]" {
			t.Errorf("A2:A3: %s", got)
		}
		if got := fmt.Sprint(valuesAt(t, res.Data, "S", grid.Range{R1: 4, C1: 3, R2: 4, C2: 3})); got != "[1]" {
			t.Errorf("C4: %s", got)
		}
	})
}

// A formula's empty text result (=IF(A1>5,A1,"")) survives every editor's
// save; Excel's results stay as they were, so a fix is no fault.
func TestEmptyTextResultKept(t *testing.T) {
	f := excelize.NewFile()
	f.SetCellValue("Sheet1", "A1", 1)
	f.SetCellFormula("Sheet1", "B1", `IF(A1>5,A1,"")`)
	f.SetCellFormula("Sheet1", "C1", "A1*2")
	f.SetCellFormula("Sheet1", "D1", "CONCAT(\"a\",A1)")
	var buf bytes.Buffer
	must(t, f.Write(&buf))
	data := patchZip(t, buf.Bytes(), dataPart, func(x string) string {
		x = regexp.MustCompile(`<c r="B1"[^>]*>(<f>[^<]*</f>)(<v>[^<]*</v>)?`).ReplaceAllString(x, `<c r="B1" t="str">${1}<v></v>`)
		return regexp.MustCompile(`<c r="C1"[^>]*>(<f>[^<]*</f>)(<v>[^<]*</v>)?`).ReplaceAllString(x, `<c r="C1">${1}<v>2</v>`)
	})
	for _, mode := range []string{"always", "never"} {
		t.Run(mode, func(t *testing.T) {
			t.Setenv("SHEETS_STREAM", mode)
			path := filepath.Join(t.TempDir(), "e.xlsx")
			must(t, os.WriteFile(path, data, 0o644))
			a, err := OpenAuto(path, data)
			if mode == "never" {
				e, err2 := OpenEditorData(path, data)
				must(t, err2)
				must(t, e.Set("Sheet1", 9, 9, input.Parse("1", false)))
				must(t, e.Save(false))
			} else {
				must(t, err)
				must(t, a.Set("Sheet1", 9, 9, input.Parse("1", false)))
				must(t, a.Save(false))
			}
			wb := read(t, path)
			if c := wb.Sheet("Sheet1").Get(1, 2); c.Pending || c.Kind != grid.Text || c.Text != "" {
				t.Errorf("B1 after an edit: %+v", c)
			}
		})
	}
	// The prefix CONCAT lacks is added: nothing Excel vouched for changes.
	bothEditors(t, func(t *testing.T) {
		res := checkData(t, "e.xlsx", data)
		if res.Withheld != "" {
			t.Errorf("withheld: %s", res.Withheld)
		}
		if _, ok := changesOf(res)["prefixed D1"]; !ok {
			t.Errorf("changes: %v", changesOf(res))
		}
		fixedTwice(t, data, res)
	})
}

// Excel's result is compared only where what the formula reads is
// Excel's own: one reading a formula Excel left without a result is not
// blamed for a difference that is its input's.
func TestCompareOnlyExcelsInputs(t *testing.T) {
	f := excelize.NewFile()
	f.SetCellValue("Sheet1", "A1", 1)
	f.SetCellFormula("Sheet1", "H1", "A1+3")
	f.SetCellFormula("Sheet1", "J2", "ROWS(A1:A3)*H1")
	var buf bytes.Buffer
	must(t, f.Write(&buf))
	data := patchZip(t, buf.Bytes(), dataPart, func(x string) string {
		x = regexp.MustCompile(`<c r="H1"[^>]*>(<f>[^<]*</f>)(<v>[^<]*</v>)?`).ReplaceAllString(x, `<c r="H1">${1}`)
		return regexp.MustCompile(`<c r="J2"[^>]*>(<f>[^<]*</f>)(<v>[^<]*</v>)?`).ReplaceAllString(x, `<c r="J2">${1}<v>3</v>`)
	})
	res := checkData(t, "c.xlsx", data)
	if len(res.Differ) != 0 || res.ReadsUnstored != 1 {
		t.Errorf("differences %+v, %d not compared", res.Differ, res.Uncomputable)
	}
}

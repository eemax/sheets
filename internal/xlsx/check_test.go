package xlsx

import (
	"bytes"
	"os"
	"regexp"
	"strings"
	"testing"

	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/input"
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

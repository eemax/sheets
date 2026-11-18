package engine

import (
	"fmt"
	"testing"

	"github.com/eemax/sheets/internal/grid"
)

// The classes WP-01 found the engine computing unlike Excel, each with
// Excel's answer (or unknown, where only Excel's language setting
// decides).

// COLUMN over an IF in an array formula: an error element stays itself,
// a value is #VALUE!, so MATCH(3,ERROR.TYPE(...),0) finds the first cell
// the IF kept.
func TestErrorTypeOfColumnOfIf(t *testing.T) {
	f := `=IFNA(INDEX(A%[1]d:C%[1]d,MATCH(3,ERROR.TYPE(COLUMN(IF(A%[1]d:C%[1]d<>"~N/A",A%[1]d:C%[1]d))),0)),"")`
	cells := map[string]any{
		"A2": "Completed", "B2": "#N/A", "C2": "#N/A",
		"A3": "#N/A", "B3": 44635, "C3": "#N/A",
		"A4": "#N/A", "B4": "#N/A", "C4": "ENR",
		"A5": "#N/A", "B5": "#N/A", "C5": "#N/A",
	}
	for r := 2; r <= 5; r++ {
		cells[grid.CellName(r, 4)] = fmt.Sprintf(f, r)
	}
	b := newBook(cells)
	for r := 2; r <= 5; r++ {
		b.arrays["Sheet1!"+grid.CellName(r, 4)] = Area{r, 4, r, 4}
	}
	e := New(b)
	for r, want := range map[int]string{2: `"Completed"`, 3: "44635", 4: `"ENR"`, 5: `""`} {
		if got := show(e.Compute("Sheet1", r, 4)); got != want {
			t.Errorf("D%d: got %s, want %s", r, got, want)
		}
	}
	for formula, want := range map[string]string{
		"COLUMN({1,2})":         "#VALUE!",
		"ERROR.TYPE(COLUMN(1))": "3",
	} {
		if got := show(e.Eval("Sheet1", 1, 9, formula)); got != want {
			t.Errorf("=%s: got %s, want %s", formula, got, want)
		}
	}
}

// Criteria: a date written as text reads only as ISO, alike in every
// language of Excel; a range of criteria in a normal formula is its
// implicit intersection (#VALUE! where there is none, a criterion like
// any error); a month's name alone is text.
func TestCriteriaUnlikeExcel(t *testing.T) {
	b := newBook(map[string]any{
		"A1": 43831, "A2": 43832, "A3": 43833, // 2020-01-01 .. 03
		"B1": 1, "B2": 2, "B3": 4,
		"C1": "A", "C2": "B", "C3": "A",
		"D1": "March", "D2": "April", "D3": "March",
		"E1": "#N/A", "E2": 5, "E3": "#N/A",
		"F6": "A", "F7": "B",
		"G2": "B",
	})
	e := New(b)
	unknown := func(crit string) string {
		return `unknown: its criterion "` + crit + `" holds a date written as text, which Excel reads by its language setting (write it as a DATE, or as 2020-01-02)`
	}
	for _, c := range []struct{ formula, want string }{
		{`SUMIF(A1:A3,"<1/2/2020",B1:B3)`, unknown("<1/2/2020")},
		{`SUMIF(A1:A3,"<1/13/2020",B1:B3)`, unknown("<1/13/2020")},
		{`COUNTIF(A1:A3,"Jan 2, 2020")`, unknown("Jan 2, 2020")},
		{`IFERROR(COUNTIF(A1:A3,"2/1/2020"),0)`, unknown("2/1/2020")},
		{`SUMIF(A1:A3,"<2020-01-02",B1:B3)`, "1"},
		{`SUMIF(A1:A3,">="&DATE(2020,1,2),B1:B3)`, "6"},
		{`SUMIF(A1:A3,"<=2020-01-02 12:00",B1:B3)`, "3"},
		{`SUMIF(D1:D3,"March",B1:B3)`, "5"},
		{`COUNTIF(D1:D3,"april")`, "1"},
		{`COUNTIF(B1:B3,">1,000")`, "0"},
		{`COUNTIF(E1:E3,#N/A)`, "2"},
		{`SUMIF(E1:E3,NA(),B1:B3)`, "5"},
		// In row 1, F6:F7 has no cell of the formula's row: #VALUE!.
		{`SUMIFS(B1:B3,C1:C3,F6:F7)`, "0"},
		// G1:G3 has one in row 1: G1, blank, which matches no letter.
		{`SUMIFS(B1:B3,C1:C3,G1:G3)`, "0"},
	} {
		if got := show(e.Eval("Sheet1", 1, 9, c.formula)); got != c.want {
			t.Errorf("=%s: got %s, want %s", c.formula, got, c.want)
		}
	}
	// In row 2 the same criteria range gives G2, "B".
	if got := show(e.Eval("Sheet1", 2, 9, `SUMIFS(B1:B3,C1:C3,G1:G3)`)); got != "2" {
		t.Errorf("=SUMIFS(B1:B3,C1:C3,G1:G3) in row 2: got %s, want 2", got)
	}
	// In an array formula a range of criteria gives one result each.
	b.arrays["Sheet1!I1"] = Area{1, 9, 2, 9}
	b.forms["Sheet1!I1"] = `SUMIFS(B1:B3,C1:C3,F6:F7)`
	e = New(b)
	if got := show(e.Compute("Sheet1", 1, 9)); got != "5" {
		t.Errorf("{=SUMIFS(B1:B3,C1:C3,F6:F7)}: got %s, want 5", got)
	}
	if a, ok := e.ArrayOf("Sheet1", 1, 9); !ok || show(a.At(1, 0)) != "2" {
		t.Errorf("{=SUMIFS(B1:B3,C1:C3,F6:F7)}: want 5 then 2, got %v", a)
	}
}

// In the 1904 date system serials count from 1904-01-01 (2020-01-01 is
// 42369): the date functions read and give serials of it. Only a date
// written as text and read as a number (arithmetic, VALUE) is read in the
// 1900 system, and counts as such (EpochUses), for the caller to leave
// unknown.
func TestDate1904(t *testing.T) {
	b := newBook(map[string]any{
		"A1": 42369, "A2": 42370, "A3": 42371, // 2020-01-01 .. 03
		"B1": 1, "B2": 2, "B3": 4,
		"C1": 42400, // 2020-02-01
	})
	e := New(b)
	e.Date1904 = true
	for _, c := range []struct {
		formula, want string
		counts        bool
	}{
		{"YEAR(A1)&\"-\"&MONTH(A1)&\"-\"&DAY(A1)", `"2020-1-1"`, false},
		{"DATE(2020,1,1)", "42369", false},
		{"DATE(1904,1,1)", "0", false},
		{"DATE(1903,12,31)", "#NUM!", false},
		{"YEAR(0)", "1904", false},
		{"YEAR(-1)", "#NUM!", false},
		{"WEEKDAY(A1)", "4", false}, // a Wednesday
		{"WEEKDAY(0)", "6", false},  // 1904-01-01, a Friday
		{"WEEKNUM(A3)", "1", false},
		{"ISOWEEKNUM(A1)", "1", false},
		{"EDATE(A1,1)", "42400", false},
		{"EOMONTH(A1,0)", "42399", false},
		{"WORKDAY(A3,1)", "42374", false}, // Friday to Monday
		{"NETWORKDAYS(A1,C1)", "23", false},
		{"NETWORKDAYS(A1,A3,A2)", "2", false},
		{"DATEDIF(A1,C1,\"d\")", "31", false},
		{"DAYS(C1,A1)", "31", false},
		{"YEARFRAC(A1,C1)", "0.0833333333333", false},
		{"DATEVALUE(\"2020-01-01\")", "42369", false},
		{"YEAR(\"2020-01-01\")", "2020", false},
		{"TIMEVALUE(\"12:00\")", "0.5", false},
		{"TEXT(A1,\"yyyy-mm-dd\")", `"2020-01-01"`, false},
		{"A1+1", "42370", false},
		{"SUMIF(A1:A3,\"<2020-01-02\",B1:B3)", "1", false},
		{"SUMIF(A1:A3,\">=\"&DATE(2020,1,2),B1:B3)", "6", false},
		{"COUNTIF(A1:A3,\"<=2020-01-02 00:00\")", "2", false},
		{"\"2020-01-01\"+0", "43831", true},
		{"VALUE(\"2020-01-01\")", "43831", true},
	} {
		before := EpochUses()
		got := show(e.Eval("Sheet1", 1, 9, c.formula))
		counted := EpochUses() != before
		if got != c.want || counted != c.counts {
			t.Errorf("=%s: got %s (counted %v), want %s (counted %v)", c.formula, got, counted, c.want, c.counts)
		}
	}
}

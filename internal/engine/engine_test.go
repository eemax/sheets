package engine

import (
	"math"
	"strings"
	"testing"

	"github.com/eemax/sheets/internal/grid"
)

// mapBook is a workbook in memory: cells by "Sheet!A1", a value or a
// formula (starting with =).
type mapBook struct {
	cells  map[string]Value
	forms  map[string]string
	names  map[string]string
	arrays map[string]Area // anchor "Sheet!A1" -> fill
}

func newBook(cells map[string]any) *mapBook {
	b := &mapBook{cells: map[string]Value{}, forms: map[string]string{}, names: map[string]string{}, arrays: map[string]Area{}}
	for k, v := range cells {
		if !strings.Contains(k, "!") {
			k = "Sheet1!" + k
		}
		switch x := v.(type) {
		case float64:
			b.cells[k] = num(x)
		case int:
			b.cells[k] = num(float64(x))
		case bool:
			b.cells[k] = boolean(x)
		case string:
			if strings.HasPrefix(x, "=") {
				b.forms[k] = x[1:]
			} else if strings.HasPrefix(x, "#") {
				b.cells[k] = errv(x)
			} else {
				b.cells[k] = str(x)
			}
		}
	}
	return b
}

func (b *mapBook) key(sheet string, row, col int) string {
	return sheet + "!" + grid.CellName(row, col)
}

func (b *mapBook) Cell(sheet string, row, col int) (Value, string) {
	k := b.key(sheet, row, col)
	return b.cells[k], b.forms[k]
}

func (b *mapBook) Extent(sheet string) (int, int, bool) {
	if sheet != "Sheet1" && sheet != "Data" {
		return 0, 0, false
	}
	rows, cols := 0, 0
	for k := range b.cells {
		if s, cell, _ := strings.Cut(k, "!"); s == sheet {
			r, c, _ := grid.ParseCell(cell)
			rows, cols = max(rows, r), max(cols, c)
		}
	}
	for k := range b.forms {
		if s, cell, _ := strings.Cut(k, "!"); s == sheet {
			r, c, _ := grid.ParseCell(cell)
			rows, cols = max(rows, r), max(cols, c)
		}
	}
	return rows, cols, true
}

func (b *mapBook) Name(name, sheet string) (string, bool) {
	t, ok := b.names[strings.ToUpper(name)]
	return t, ok
}

func (b *mapBook) ArrayAt(sheet string, row, col int) (int, int, Area, bool) {
	for k, a := range b.arrays {
		s, _, _ := strings.Cut(k, "!")
		if s == sheet && row >= a.R1 && row <= a.R2 && col >= a.C1 && col <= a.C2 {
			return a.R1, a.C1, a, true
		}
	}
	return 0, 0, Area{}, false
}

// show writes a value for comparison: numbers to 12 significant digits.
func show(v Value) string {
	switch v.Kind {
	case Number:
		if v.Num == math.Trunc(v.Num) && math.Abs(v.Num) < 1e15 {
			return NumberText(v.Num)
		}
		return NumberText(roundSig(v.Num, 12))
	case Text:
		return `"` + v.Str + `"`
	case Bool:
		if v.Bool {
			return "TRUE"
		}
		return "FALSE"
	case Error:
		if v.Str == ErrUnsupported {
			return "unknown: " + v.Why()
		}
		return v.Str
	case Empty:
		return "blank"
	}
	return "?"
}

// Excel's answers, as Excel gives them, for formulas over a small sheet.
func TestExcelAnswers(t *testing.T) {
	b := newBook(map[string]any{
		"A1": -100, "A2": 30, "A3": 40, "A4": 4, "A5": 50,
		"B1": "apple", "B2": "Banana", "B3": "cherry", "B4": "banana", "B5": 1,
		"C1": 123456789, "C2": 911313316543722.0, "C3": 0.0104166666666667, // 00:15
		"D1": "x", "D2": "y", "D3": "x", "D4": "", "D5": "x",
		"E1": 10, "E2": 20, "E3": 30, "E4": 40, "E5": 50,
		"F1": "#REF!", "F2": "1,200", "F3": "TRUE",
		"G1": 1, "G2": 3, "G3": 5, "G4": 7, "G5": 9,
		"H1": "a", "H2": "b", "H3": "c", "H4": "d", "H5": "e",
	})
	cases := []struct{ formula, want string }{
		// excelize got these wrong.
		{"NPV(0.09,A1:A5)", "-0.275110106438"},
		{"AVEDEV(A1:A5)", "42.24"},
		{"AND(A1:A5)", "TRUE"},
		{"RIGHT(C1,4)", `"6789"`},
		{"ISNUMBER(C2)", "TRUE"},
		{"TEXT(C2,\"0\")", `"911313316543722"`},
		{"IFNA(INDEX(B1:B5,MATCH(\"zzz\",B1:B5,0)),\"NEW\")", `"NEW"`},
		{"ROUND(F1,-3)", "#REF!"},
		{"MOD(C3-TIME(0,30,0)+1,1)", "0.989583333333"},
		{"VLOOKUP(\"cherry\",B1:E5,4,FALSE)", "30"},
		{"SUMIF(D1:D5,\"\",E1:E5)", "40"},
		{"Z99", "0"},
		{"IF(TRUE,Z99)", "0"},
		{"Z99&\"\"", `""`},
		// Arithmetic and text.
		{"0.1+0.2-0.3", "0"},
		{"-2^2", "4"},
		{"2^3^2", "64"},
		{"\"a\"&1.5&TRUE", `"a1.5TRUE"`},
		// Text with decimal or thousands marks reads by Excel's language.
		{"F2*2", "unknown: it reads \"1,200\" as a number or date, which depends on Excel's language setting"},
		{"\"x\"*2", "#VALUE!"},
		{"1/0", "#DIV/0!"},
		{"ROUND(2.675,2)", "2.68"},
		{"ROUNDDOWN(-2.5,0)", "-2"},
		{"MROUND(10,3)", "9"},
		{"CEILING(-2.5,-1)", "-3"},
		{"FLOOR.MATH(-2.5)", "-3"},
		{"CEILING.MATH(-2.5)", "-2"},
		{"INT(-1.5)", "-2"},
		{"1/3&\"\"", `"0.333333333333333"`},
		{"10^20&\"\"", `"1E+20"`},
		// Comparison.
		{"\"a\"=\"A\"", "TRUE"},
		{"1<\"a\"", "TRUE"},
		{"\"z\"<TRUE", "TRUE"},
		{"Z99=0", "TRUE"},
		{"Z99=\"\"", "TRUE"},
		// Aggregates skip text in ranges, not typed values.
		{"SUM(B1:B5)", "1"},
		{"SUM(\"3\",2)", "5"},
		{"SUM(\"x\",2)", "#VALUE!"},
		{"COUNT(A1:B5,\"7\",TRUE)", "8"},
		{"COUNTA(D1:D5)", "5"},
		{"COUNTBLANK(D1:D6)", "2"},
		{"AVERAGE(B1:B4)", "#DIV/0!"},
		{"MAX(B1:B4)", "0"},
		{"SUMPRODUCT((D1:D5=\"x\")*E1:E5)", "90"},
		{"SUMPRODUCT(--(D1:D5=\"x\"))", "3"},
		{"COUNTIF(B1:B5,\"b*\")", "2"},
		{"COUNTIF(B1:B5,\"<>banana\")", "3"},
		{"COUNTIF(E1:E5,\">=30\")", "3"},
		{"COUNTIFS(D1:D5,\"x\",E1:E5,\">10\")", "2"},
		{"SUMIFS(E1:E5,D1:D5,\"x\",E1:E5,\"<50\")", "40"},
		{"AVERAGEIF(D1:D5,\"x\",E1:E5)", "30"},
		{"SUM(COUNTIF(D1:D5,{\"x\",\"y\"}))", "4"},
		{"LARGE(E1:E5,2)", "40"},
		{"SMALL(E1:E5,{1,2})", "10"},
		{"MEDIAN(E1:E5)", "30"},
		{"STDEV(E1:E5)", "15.8113883008"},
		{"RANK(30,E1:E5)", "3"},
		{"PERCENTILE(E1:E5,0.25)", "20"},
		// Lookups.
		{"MATCH(25,E1:E5)", "2"},
		{"MATCH(25,E1:E5,1)", "2"},
		{"MATCH(\"BANANA\",B1:B5,0)", "2"},
		{"MATCH(\"b?n*\",B1:B5,0)", "2"},
		{"MATCH(25,E1:E5,)", "#N/A"},
		{"VLOOKUP(25,E1:H5,4)", `"b"`},
		{"VLOOKUP(25,E1:H5,4,)", "#N/A"},
		{"INDEX(E1:H5,3,4)", `"c"`},
		{"INDEX(E1:E5,3)", "30"},
		{"SUM(INDEX(E1:H5,0,1))", "150"},
		{"SUM(E1:INDEX(E1:E5,3))", "60"},
		{"LOOKUP(2,1/(D1:D5=\"x\"),E1:E5)", "50"},
		{"XLOOKUP(\"y\",D1:D5,E1:E5)", "20"},
		{"XLOOKUP(\"q\",D1:D5,E1:E5,\"none\")", `"none"`},
		{"XLOOKUP(6,G1:G5,H1:H5,,-1)", `"c"`},
		{"INDEX(H1:H5,MATCH(1,INDEX((D1:D5=\"x\")*(E1:E5>20),0),0))", `"c"`},
		{"ROWS(A:A)", "1048576"},
		{"OFFSET(E1,2,0)", "30"},
		{"SUM(OFFSET(E1,1,0,3,1))", "90"},
		{"INDIRECT(\"E\"&2)", "20"},
		{"ADDRESS(2,3)", `"$C$2"`},
		// Implicit intersection in a normal formula (this cell is row 1).
		{"E1:E5*2", "20"},
		{"SUM(E1:E5*2)", "20"},
		{"ROW(A3:A5)", "3"},
		// Text.
		{"LEN(\"héllo\")", "5"},
		{"PROPER(\"hello wORLD\")", `"Hello World"`},
		{"TRIM(\"  a   b  \")", `"a b"`},
		{"SUBSTITUTE(\"a-b-c\",\"-\",\"+\",2)", `"a-b+c"`},
		{"FIND(\"b\",\"abcb\",3)", "4"},
		{"SEARCH(\"B?\",\"abcb\")", "2"},
		{"MID(\"abcdef\",2,3)", `"bcd"`},
		{"REPLACE(\"abcdef\",2,3,\"X\")", `"aXef"`},
		{"TEXTJOIN(\",\",TRUE,D1:D5)", `"x,y,x,x"`},
		{"TEXTJOIN(\",\",,D1:D5)", `"x,y,x,x"`},
		{"CONCAT(H1:H3)", `"abc"`},
		{"VALUE(\"$1,200.50\")", "unknown: it reads \"$1,200.50\" as a number or date, which depends on Excel's language setting"},
		{"VALUE(\"12%\")", "0.12"},
		{"VALUE(\"2024-07-01\")", "45474"},
		{"VALUE(\"x\")", "#VALUE!"},
		{"TEXT(1234.567,\"#,##0.00\")", `"1,234.57"`},
		{"TEXT(0.256,\"0.0%\")", `"25.6%"`},
		{"TEXT(-5,\"0;(0)\")", `"(5)"`},
		{"TEXT(45474,\"yyyy-mm-dd\")", `"2024-07-01"`},
		{"TEXT(45474,\"dddd, mmmm d\")", `"Monday, July 1"`},
		{"TEXT(0.0219444444444,\"h:mm:ss\")", `"0:31:36"`},
		{"TEXT(1.5,\"[h]:mm\")", `"36:00"`},
		{"TEXT(123456789,\"000-00-0000\")", `"123-45-6789"`},
		{"TEXT(1234567,\"#,##0,\")", `"1,235"`},
		{"TEXT(4,\"[<3]\"\"dog\"\"0;\")", `""`},
		{"FIXED(1234.5,1)", `"1,234.5"`},
		{"DOLLAR(-1234.5)", `"($1,234.50)"`},
		{"CHAR(65)&CODE(\"a\")", `"A97"`},
		// Dates.
		{"DATE(2024,2,30)", "45352"},
		{"DATE(2024,14,1)", "45689"},
		{"DATE(1900,3,0)", "60"},
		{"YEAR(45474)&\"-\"&MONTH(45474)&\"-\"&DAY(45474)", `"2024-7-1"`},
		{"EOMONTH(45474,1)", "45535"},
		{"EOMONTH(45474,-12)", "45138"},
		{"EDATE(\"2024-01-31\",1)", "45351"},
		{"WEEKDAY(45474)", "2"},
		{"WEEKDAY(45474,2)", "1"},
		{"WEEKNUM(45474)", "27"},
		{"WEEKNUM(45474,21)", "27"},
		{"DATEDIF(\"2020-02-29\",\"2024-02-28\",\"Y\")", "3"},
		// Excel's own quirk: "MD" can go negative.
		{"DATEDIF(\"2024-01-31\",\"2024-03-01\",\"MD\")", "-1"},
		{"YEARFRAC(\"2024-01-01\",\"2024-07-01\")", "0.5"},
		{"NETWORKDAYS(\"2024-07-01\",\"2024-07-12\")", "10"},
		{"WORKDAY(\"2024-07-05\",1)", "45481"},
		{"DATEVALUE(\"Jan 2, 2024\")", "unknown: it reads \"Jan 2, 2024\" as a number or date, which depends on Excel's language setting"},
		{"DATEVALUE(\"2-Jan-2024\")", "unknown: it reads \"2-Jan-2024\" as a number or date, which depends on Excel's language setting"},
		{"DATEVALUE(\"1Jan2024\")", "unknown: it reads \"1Jan2024\" as a number or date, which depends on Excel's language setting"},
		{"TIMEVALUE(\"1:30 PM\")", "0.5625"},
		{"DATEVALUE(\"21/11/2024\")", "unknown: it reads \"21/11/2024\" as a number or date, which depends on Excel's language setting"},
		{"DATEVALUE(\"11/05/2024\")", "unknown: it reads \"11/05/2024\" as a number or date, which depends on Excel's language setting"},
		{"DATEVALUE(\"12/31/2024\")", "unknown: it reads \"12/31/2024\" as a number or date, which depends on Excel's language setting"},
		{"DATEVALUE(\"2024-12-31\")", "45657"},
		{"VALUE(\" 42 \")+VALUE(\"-1e3\")", "-958"},
		// Financial.
		{"PMT(0.05/12,360,200000)", "-1073.64324602"},
		{"FV(0.05,10,-100)", "1257.78925355"},
		{"PV(0.05,10,-100)", "772.173492918"},
		{"NPER(0.01,-100,1000)", "10.5886444594"},
		{"RATE(10,-100,800)", "0.0427749780351"},
		{"IRR({-100,30,40,50})", "0.0889633946933"},
		// Logic and errors.
		{"IF(FALSE,1)", "FALSE"},
		{"IF(FALSE,1,)", "0"},
		{"IFERROR(1/0,\"none\")", `"none"`},
		{"IFERROR(F1,0)", "0"},
		{"ISERROR(F1)", "TRUE"},
		{"ISERR(NA())", "FALSE"},
		{"ERROR.TYPE(1/0)", "2"},
		{"CHOOSE(2,\"a\",\"b\",\"c\")", `"b"`},
		{"SWITCH(2,1,\"one\",2,\"two\",\"other\")", `"two"`},
		{"IFS(E1>10,\"big\",TRUE,\"small\")", `"small"`},
		{"LET(x,2,y,x*3,x+y)", "8"},
		{"XOR(TRUE,FALSE,TRUE)", "FALSE"},
		{"NOT(0)", "TRUE"},
		{"SUM({1,2;3,4}*2)", "20"},
		{"UNKNOWNFN(1)", "#NAME?"},
		{"IFERROR(UNKNOWNFN(1),0)", "0"},
		{"CELL(\"row\",A1)", "unknown: CELL is not supported by the sheets calculator"},
		{"IFERROR(CELL(\"row\",A1),0)", "unknown: CELL is not supported by the sheets calculator"},
		{"ISERROR(CELL(\"row\",A1))", "unknown: CELL is not supported by the sheets calculator"},
		{"[1]Sheet1!A1", "unknown: links to another workbook"},
		{"Missing!A1", "#REF!"},
		{"nosuchname", "#NAME?"},
	}
	e := New(b)
	for _, c := range cases {
		got := show(e.Eval("Sheet1", 1, 26, c.formula))
		if got != c.want {
			t.Errorf("=%s: got %s, want %s", c.formula, got, c.want)
		}
	}
}

// An array formula evaluates its ranges as arrays, and fills its range.
func TestArrayFormulas(t *testing.T) {
	b := newBook(map[string]any{
		"A1": "x", "A2": "y", "A3": "x", "A4": "z",
		"B1": 10, "B2": 20, "B3": 30, "B4": 40,
		"C1": "=SUM(IF(A1:A4=\"x\",B1:B4))",
		"C2": "=SUM(IF(A1:A4=\"x\",B1:B4))",
		"D1": "=INDEX(B1:B4,SMALL(IF(A1:A4=\"x\",ROW(A1:A4)),2))",
		"E1": "=TRANSPOSE(B1:B4)",
	})
	b.arrays["Sheet1!C1"] = Area{1, 3, 1, 3}
	b.arrays["Sheet1!D1"] = Area{1, 4, 1, 4}
	b.arrays["Sheet1!E1"] = Area{1, 5, 1, 8}
	e := New(b)
	for cell, want := range map[string]string{"C1": "40", "C2": "0", "D1": "30", "E1": "10", "F1": "20", "H1": "40"} {
		r, c, _ := grid.ParseCell(cell)
		if got := show(e.cellValue("Sheet1", r, c)); got != want {
			t.Errorf("%s: got %s, want %s", cell, got, want)
		}
	}
}

// Formulas reading each other compute once; a circular reference is
// marked, not looped.
func TestDependencies(t *testing.T) {
	b := newBook(map[string]any{
		"A1": 5, "A2": "=A1*2", "A3": "=A2+A1", "A4": "=SUM(A1:A3)",
		"B1": "=B2", "B2": "=B1",
		"Data!A1": 7, "C1": "=Data!A1+1",
	})
	e := New(b)
	if got := show(e.Compute("Sheet1", 4, 1)); got != "30" {
		t.Errorf("A4: %s", got)
	}
	if got := show(e.Compute("Sheet1", 1, 2)); got != "unknown: circular reference" && !strings.HasPrefix(got, "#CYCLE") {
		t.Errorf("B1: %s", got)
	}
	if !e.Cycle {
		t.Error("circular reference not noted")
	}
	if got := show(e.Compute("Sheet1", 1, 3)); got != "8" {
		t.Errorf("C1: %s", got)
	}
}

func TestArrayUse(t *testing.T) {
	for _, c := range []struct {
		f             string
		array, spills bool
	}{
		{"SUM(IF(A1:A3>1,A1:A3))", true, false},
		{"SUM(A1:A3*B1:B3)", true, false},
		{"MAX(A1:A3-B1:B3)", true, false},
		{"AND(A1:A3>1)", true, false},
		{"SUM(LEN(A1:A3))", true, false},
		{"TEXTJOIN(\",\",TRUE,IF(A1:A3>1,A1:A3,\"\"))", true, false},
		{"SUM(B:B*C:C)", true, false},
		{"A1:A3*2", true, true},
		{"A1:A3", true, true},
		{"IF(A1:A3>1,\"y\",\"n\")", true, true},
		{"B:B*2", false, false}, // the cell in the formula's row, as written for older Excel
		{"SUMPRODUCT((A1:A3>1)*B1:B3)", false, false},
		{"SUM(A1:A3)/COUNT(B1:B3)", false, false},
		{"INDEX(A1:A10,MATCH(5,B1:B10,0))", false, false},
		{"VLOOKUP(D1,A1:B10,2,FALSE)", false, false},
		{"SUMIFS(C:C,A:A,\"x\",B:B,\">1\")", false, false},
		{"IFERROR(A1/B1,0)", false, false},
		{"@A1:A3*2", false, false},
		{"A1*2", false, false},
		{"_xlfn.XLOOKUP(D1,A1:A9,B1:B9)", false, false},
		{"SUMPRODUCT(--(A2:A4=\"East\"),B2:B4)", false, false},
		{"SUMPRODUCT({1,2,3},A1:A3)", false, false},
		{"ROUND(SUM(Data!B2:B9)/COUNT(B2:B9),2)", false, false},
		{"TEXT(A1,\"0\")&\" of \"&ROWS(A1:A9)", false, false},
		{"(A1:A9)*2", true, true},
		{"ABS(A1:A9)", true, true},
		{"SUM((A1:A9))", false, false},
		{"NETWORKDAYS(A1,B1,H1:H9)", false, false},
		{"NETWORKDAYS(A1:A3,B1)", true, true},
		{"VLOOKUP(A2:A4,D:E,2,FALSE)", true, true},
		{"COUNTIF(A:A,B2:B4)", true, true},
		{"SUMIF(A:A,B2:B4,C:C)", true, true},
		{"SUMIFS(C:C,A:A,B2:B4)", true, true},
		{"COUNTIFS(A:A,\"x\",B:B,C2:C4)", true, true},
		{"SUM(COUNTIF(A:A,B2:B4))", true, false},
		{"INDEX(A:A,B2:B4)", true, true},
		{"D2#*2", true, true},
		{`SMALL(IF(INDEX($T$2:$T$600,,MATCH("x",$T$1:$V$1,0))=CB2,ROW($T$2:$T$600)),1)`, true, false},
		{"INDEX(A2:C9,3,2)*2", false, false},
		{"SUM(INDEX(A2:C9,0,2)*2)", true, false},
		{"SUM(OFFSET(A1,0,0,5,1)*2)", true, false},
		{"OFFSET(A1,1,1)*2", false, false},
		{"INDEX(A2:B2,1)*2", false, false},
		{"INDEX(A2:A9,0)*2", true, true},
		{"INDEX(A2:C9,2)*2", true, true},
		{"COUNTA(D2#)", false, false},
		{"SUM(D2#*2)", true, false},
	} {
		array, spills := ArrayUse(c.f)
		if array != c.array || spills != c.spills {
			t.Errorf("ArrayUse(%s) = %v, %v; want %v, %v", c.f, array, spills, c.array, c.spills)
		}
	}
}

package formula

import (
	"slices"
	"testing"

	"github.com/eemax/sheets/internal/grid"
)

func TestShift(t *testing.T) {
	tests := []struct {
		f      string
		dr, dc int
		want   string
	}{
		{"A2+B2", 1, 0, "A3+B3"},
		{"$A$2+B$2+$C2", 3, 1, "$A$2+C$2+$C5"},
		{"SUM(D2:D10)", 2, 0, "SUM(D4:D12)"},
		{"SUM(A:A)", 0, 2, "SUM(C:C)"},
		{"SUM(2:4)", 1, 0, "SUM(3:5)"},
		{"Sheet2!A1*2", 1, 0, "Sheet2!A2*2"},
		{"'My Sheet'!B3", 0, 1, "'My Sheet'!C3"},
		{"Q1!A1", 1, 0, "Q1!A2"}, // Q1 is a sheet name here
		{`IF(A1>0,"B2",A1)`, 1, 0, `IF(A2>0,"B2",A2)`},
		{"LOG10(A1)", 1, 0, "LOG10(A2)"},
		{"Table1[Amount]*A1", 1, 0, "Table1[Amount]*A2"},
		{"A1*100", 1, 0, "A2*100"},
		{"A1*1E5", 1, 0, "A2*1E5"},
		{"A1-1", -1, 0, "#REF!-1"},
		{"TaxRate*A1", 1, 0, "TaxRate*A2"},
	}
	for _, tt := range tests {
		if got := Shift(tt.f, tt.dr, tt.dc); got != tt.want {
			t.Errorf("Shift(%q, %d, %d) = %q, want %q", tt.f, tt.dr, tt.dc, got, tt.want)
		}
	}
}

func TestPattern(t *testing.T) {
	k1, d1 := Pattern("D2-E2", 2, 6)
	k2, _ := Pattern("D3-E3", 3, 6)
	if k1 != k2 {
		t.Errorf("keys differ: %q vs %q", k1, k2)
	}
	if d1 != "D{r}-E{r}" {
		t.Errorf("display = %q", d1)
	}
	if _, d := Pattern("SUM(B$2:B2)+A1", 5, 3); d != "SUM(B$2:B{r-3})+A{r-4}" {
		t.Errorf("display = %q", d)
	}
	k3, _ := Pattern("D3-E2", 3, 6)
	if k3 == k1 {
		t.Error("different relative formulas share a key")
	}
}

func TestCrossSheet(t *testing.T) {
	for f, want := range map[string]bool{
		"A1+B1":         false,
		"Rates!B1*A1":   true,
		`"a!b"&A1`:      false,
		"'My Sheet'!A1": true,
	} {
		if got := CrossSheet(f); got != want {
			t.Errorf("CrossSheet(%q) = %v", f, got)
		}
	}
}

func TestFillRow(t *testing.T) {
	for in, want := range map[string]string{
		"C{r}*D{r}":       "C7*D7",
		"SUM(B$2:B{r-1})": "SUM(B$2:B6)",
		"A{r+2}":          "A9",
		`"{r}"&A1`:        `"7"&A1`, // inside strings too; {r} is not Excel syntax
		"A1+B1":           "A1+B1",
	} {
		if got := FillRow(in, 7); got != want {
			t.Errorf("FillRow(%q) = %q, want %q", in, got, want)
		}
	}
	// A pattern from peek written back gives the original formula.
	_, p := Pattern("SUM(B$2:B6)+C7", 7, 4)
	if got := FillRow(p, 7); got != "SUM(B$2:B6)+C7" {
		t.Errorf("round trip = %q", got)
	}
}

func TestSameRowColumns(t *testing.T) {
	got := SameRowColumns("D7*E7+Rates!F7+$G$7+H6+SUM(I7:J7)", 7)
	want := []int{4, 5, 9, 10} // D, E and the I7:J7 range ends
	if len(got) != len(want) {
		t.Fatalf("got %v", got)
	}
	for i := range want {
		if got[i] != want[i] {
			t.Fatalf("got %v, want %v", got, want)
		}
	}
	if !Divides("A1/B1") || Divides(`"a/b"&A1`) {
		t.Error("Divides")
	}
}

func TestDeleteLines(t *testing.T) {
	tests := []struct {
		f, home string
		want    string
	}{
		{"B5*2", "Data", "#REF!*2"},
		{"SUM(B5:B5)", "Data", "SUM(#REF!)"},
		{"SUM(B4:B6)", "Data", "SUM(B4:B5)"},
		{"SUM(B5:B8)", "Data", "SUM(B5:B7)"},
		{"SUM(B2:B5)", "Data", "SUM(B2:B4)"},
		{"SUM(B1:B3)+B9+$B$9", "Data", "SUM(B1:B3)+B8+$B$8"},
		{"Data!B5*10", "Other", "Data!#REF!*10"},
		{"Data!B6", "Other", "Data!B5"},
		{"'Data'!$B$5+B5", "Other", "'Data'!#REF!+B5"},
		{"B5", "Other", "B5"},
		{"SUM(5:5)+SUM(7:9)", "Data", "SUM(#REF!)+SUM(6:8)"},
		{"SUM(A:A)", "Data", "SUM(A:A)"},
		{"Data!$B$5", "", "Data!#REF!"},
	}
	for _, tt := range tests {
		if got := DeleteLines(tt.f, tt.home, "Data", false, 5, 5); got != tt.want {
			t.Errorf("DeleteLines(%s on %s) = %s, want %s", tt.f, tt.home, got, tt.want)
		}
	}
	if got := DeleteLines("C1+SUM(C:D)+B1+E1", "Data", "Data", true, 3, 3); got != "#REF!+SUM(C:C)+B1+D1" {
		t.Errorf("columns: %s", got)
	}
}

func TestInsertLines(t *testing.T) {
	tests := []struct{ f, home, want string }{
		{"SUM(B2:B5)", "Data", "SUM(B2:B7)"},
		{"SUM(B3:B5)+B2", "Data", "SUM(B5:B7)+B2"},
		{"$B$9", "Data", "$B$11"},
		{"Data!B4+Other!B4", "Other", "Data!B6+Other!B4"},
		{"SUM(Data!B2:B4)", "Other", "SUM(Data!B2:B6)"},
		{"B4", "Other", "B4"},
		{"SUM(3:4)", "Data", "SUM(5:6)"},
		{"LET(x,B4,x*2)", "Data", "LET(x,B6,x*2)"},
	}
	for _, tt := range tests {
		if got := InsertLines(tt.f, tt.home, "Data", false, 3, 2); got != tt.want {
			t.Errorf("InsertLines(%s on %s) = %s, want %s", tt.f, tt.home, got, tt.want)
		}
	}
}

func TestExtendRanges(t *testing.T) {
	for in, want := range map[string]string{
		"SUM(B4:B9)":                "SUM(B4:B12)",
		"SUM(B$4:B$9)/COUNT(B4:B9)": "SUM(B$4:B$12)/COUNT(B4:B12)",
		"B9":                        "B9",
		"SUM(B4:B8)":                "SUM(B4:B8)",
		"SUM(Other!B4:B9)":          "SUM(Other!B4:B9)",
	} {
		if got := ExtendRanges(in, "Data", "Data", 9, 12); got != want {
			t.Errorf("ExtendRanges(%s) = %s, want %s", in, got, want)
		}
	}
}

func TestRefsAndNumbers(t *testing.T) {
	refs := Refs("SUM(C3:C7)+Other!B2*1.21+A:A")
	if len(refs) != 3 || refs[0].Range.String() != "C3:C7" || refs[1].Sheet != "Other" || refs[1].Range.String() != "B2" || refs[2].Range.C1 != 1 || refs[2].Range.R2 != grid.MaxRows {
		t.Errorf("Refs = %+v", refs)
	}
	nums := Numbers(`ROUND(B2*1.21,2)+SUM(3:5)+"7"+LOG10(A1)`)
	if len(nums) != 2 || nums[0] != 1.21 || nums[1] != 2 {
		t.Errorf("Numbers = %v", nums)
	}
}

func TestCheck(t *testing.T) {
	for f, ok := range map[string]bool{
		"=SUM(A1:A3)":                          true,
		"=IF(A1=\"x)\",1,0)":                   true,
		"='My Sheet'!A1+'It''s'!B2":            true,
		"=SUM(Sales[Units])":                   true,
		"=Sales[[#Totals],[Units]]":            true,
		"=SUM(T[Col'[1']])":                    true,
		"=T[It''s]":                            true,
		"=SUM({1,2;3,4})":                      true,
		"=A1%":                                 true,
		"=SUM(A1,)":                            true,
		"=-A1":                                 true,
		"=[1]Sheet1!A1":                        true,
		"=LET(x,1,x+1)":                        true,
		"=SUM(":                                false,
		"=SUM(A1))":                            false,
		"=A1+":                                 false,
		"=\"abc":                               false,
		"='My Sheet!A1":                        false,
		"=":                                    false,
		"=SUM(A1:A3":                           false,
		"={1,2":                                false,
		"=Sales[Units":                         false,
		"=A1:":                                 false,
		"=(A1+B1)*(C1":                         false,
		"=IF(A1>0,\"yes\",\"no\")":             true,
		"=TEXT(A1,\"#,##0\")&\" units\"":       true,
		"=SUMIFS(C:C,A:A,\"East\",B:B,\">2\")": true,
	} {
		if err := Check(f); (err == nil) != ok {
			t.Errorf("Check(%q) = %v, want ok=%v", f, err, ok)
		}
	}
}

func TestGrowRanges(t *testing.T) {
	for _, c := range []struct{ in, want string }{
		{"SUM(Sales!$B$2:$B$4)", "SUM(Sales!$B$2:$B$6)"},
		{"SUMIFS(Sales!B2:B4,Sales!A1:A4,\"x\")", "SUMIFS(Sales!B2:B6,Sales!A1:A6,\"x\")"},
		{"SUM(Sales!B3:B4)", "SUM(Sales!B3:B4)"}, // part of the column
		{"SUM(Sales!B2:B3)", "SUM(Sales!B2:B3)"}, // does not reach the end
		{"SUM(Sales!D2:D4)", "SUM(Sales!D2:D4)"}, // outside the table
		{"SUM(Other!B2:B4)", "SUM(Other!B2:B4)"}, // another sheet
		{"Sales!B2*2", "Sales!B2*2"},
		{"SUM(Sales!$B$2:B4)", "SUM(Sales!$B$2:B4)"}, // a running total
		{"SUM(Sales!B$2:$B4)", "SUM(Sales!B$2:$B4)"}, // a running total
		{"SUM(Sales!B2:$B$4)", "SUM(Sales!B2:$B$6)"}, // fixed at the end
	} {
		if got := GrowRanges(c.in, "Summary", "Sales", 2, 4, 6, 1, 3); got != c.want {
			t.Errorf("%s: got %s, want %s", c.in, got, c.want)
		}
	}
	// Unqualified ranges are the formula's own sheet's.
	if got := GrowRanges("SUM(B2:B4)", "Sales", "Sales", 2, 4, 6, 1, 3); got != "SUM(B2:B6)" {
		t.Errorf("own sheet: %s", got)
	}
}

// The spill operator is stored as Excel stores it.
func TestSpillRefs(t *testing.T) {
	for in, want := range map[string]string{
		"COUNTA(D2#)":         "COUNTA(_xlfn.ANCHORARRAY(D2))",
		"SUM(Sheet1!$D$2#)*2": "SUM(_xlfn.ANCHORARRAY(Sheet1!$D$2))*2",
		"SUM('My Sheet'!D2#)": "SUM(_xlfn.ANCHORARRAY('My Sheet'!D2))",
		`IF(A1="x#",D2#,0)`:   `IF(A1="x#",_xlfn.ANCHORARRAY(D2),0)`,
		"IFERROR(A1,#N/A)":    "IFERROR(A1,#N/A)",
	} {
		got := ToFile(in)
		if got != want {
			t.Errorf("ToFile(%s) = %s, want %s", in, got, want)
		}
		if back := FromFile(got); back != in {
			t.Errorf("FromFile(%s) = %s, want %s", got, back, in)
		}
	}
}

// Functions added after Excel 2007 stored without their prefix, as
// openpyxl writes them, are found; prefixed ones, names in strings,
// sheet names and structured references are not.
func TestUnprefixed(t *testing.T) {
	for _, tt := range []struct {
		stored string
		want   []string
	}{
		{`IFERROR(INDEX(A1:A4,AGGREGATE(15,6,ROW(A1:A4)/(A1:A4="x"),1)),"")`, []string{"AGGREGATE"}},
		{`textjoin(",",TRUE,A1:A4)&MAXIFS(B:B,A:A,"x")&TEXTJOIN("-",1,C1)`, []string{"TEXTJOIN", "MAXIFS"}},
		{`_xlfn.AGGREGATE(9,6,A1:A4)+_xlfn._xlws.FILTER(A1:A4,B1:B4)`, nil},
		{`SUM(A1:A4)&"TEXTJOIN(x)"`, nil},
		{`'IFS(1)'!A1+XLOOKUP!B2`, nil},
		{`SUM(Table1[[#This Row],[MAXIFS(]])`, nil},
		{`_xlfn.LET(_xlpm.x,1,_xlpm.x+IFS(TRUE,1))`, []string{"IFS"}},
		{`FILTER(A1:A4,B1:B4>1)`, []string{"FILTER"}},
	} {
		if got := Unprefixed(tt.stored); !slices.Equal(got, tt.want) {
			t.Errorf("Unprefixed(%s) = %v, want %v", tt.stored, got, tt.want)
		}
	}
}

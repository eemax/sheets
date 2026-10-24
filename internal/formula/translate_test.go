package formula

import (
	"testing"

	"github.com/eemax/sheets/internal/grid"
)

func TestExpandTables(t *testing.T) {
	// Sales: A1:D6 on sheet Data, header row 1, data rows 2-5, totals row 6.
	tables := []Table{{Name: "Sales", Sheet: "Data", Ref: grid.Range{R1: 1, C1: 1, R2: 6, C2: 4}, Headers: 1, Totals: 1,
		Columns: []string{"Region", "Units", "Unit price", "Total"}}}
	for _, tt := range []struct {
		f        string
		row, col int
		sheet    string
		want     string
		ok       bool
	}{
		{"SUM(Sales[Units])", 1, 9, "Summary", "SUM(Data!$B$2:$B$5)", true},
		{"SUBTOTAL(109,[Units])", 6, 2, "Data", "SUBTOTAL(109,Data!$B$2:$B$5)", true},
		{"[@Units]*[@[Unit price]]", 3, 4, "Data", "Data!$B$3*Data!$C$3", true},
		{"Sales[[#This Row],[Units]]*2", 4, 4, "Data", "Data!$B$4*2", true},
		{"SUM(Sales[[Units]:[Unit price]])", 1, 9, "Summary", "SUM(Data!$B$2:$C$5)", true},
		{"Sales[[#Totals],[Units]]", 1, 9, "Summary", "Data!$B$6", true},
		{"COUNTA(Sales[#All])", 1, 9, "Summary", "COUNTA(Data!$A$1:$D$6)", true},
		{"ROWS(Sales)", 1, 9, "Summary", "ROWS(Data!$A$2:$D$5)", true},
		{"VLOOKUP(\"North\",Sales[#Data],2,0)", 1, 9, "Summary", "VLOOKUP(\"North\",Data!$A$2:$D$5,2,0)", true},
		{"[@Units]", 3, 9, "Data", "", false},             // outside the table: no table to mean
		{"Sales[Units]*2", 1, 9, "Summary", "", false},    // a column in arithmetic: Excel spills it
		{"[@Units]", 1, 2, "Data", "", false},             // the header row has no this-row data
		{"SUM(Other[Units])", 1, 9, "Summary", "", false}, // no such table
		{"SUM(Sales[Nope])", 1, 9, "Summary", "", false},
		{`"Sales[Units]"&A1`, 1, 9, "Summary", `"Sales[Units]"&A1`, true},
	} {
		got, ok := ExpandTables(tt.f, tt.sheet, tt.row, tt.col, tables)
		if ok != tt.ok || ok && got != tt.want {
			t.Errorf("ExpandTables(%q) = %q %v, want %q %v", tt.f, got, ok, tt.want, tt.ok)
		}
	}
}

func TestInlineLet(t *testing.T) {
	for _, tt := range []struct {
		f, want string
		ok      bool
	}{
		{"LET(x,2,x*3)", "((2)*3)", true},
		{"LET(x,A1,y,x+1,x*y)", "((A1)*((A1)+1))", true},
		{"1+LET(a,B2,LET(b,a*2,b+a))", "1+((((B2)*2)+(B2)))", true},
		{"LET(x,1,LET(x,2,x))", "(((2)))", true}, // the inner x shadows the outer
		{"LET(r,RAND(),r-r)", "", false},
		{"LET(f,LAMBDA(n,n*2),f(3))", "", false},
		{"SUM(A1:A3)", "SUM(A1:A3)", true},
		{`LET(x,"x",x&Sheet1!x)`, `(("x")&Sheet1!x)`, true}, // a reference after ! is not the name
	} {
		got, ok := InlineLet(tt.f)
		if ok != tt.ok || ok && got != tt.want {
			t.Errorf("InlineLet(%q) = %q %v, want %q %v", tt.f, got, ok, tt.want, tt.ok)
		}
	}
}

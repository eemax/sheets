package formula

import "testing"

func TestRenameTableColumns(t *testing.T) {
	renames := map[string]string{"units": "Qty", "price": "Unit price", "a#b": "C[1]"}
	for _, tt := range []struct {
		in      string
		inTable bool
		want    string
	}{
		{"SUM(SalesT[Units])", false, "SUM(SalesT[Qty])"},
		{"SUM(salest[UNITS])", false, "SUM(salest[Qty])"},
		{"SalesT[[#This Row],[Units]]*SalesT[[#This Row],[Price]]", false, "SalesT[[#This Row],[Qty]]*SalesT[[#This Row],[Unit price]]"},
		{"SUM(SalesT[[Units]:[Price]])", false, "SUM(SalesT[[Qty]:[Unit price]])"},
		{"SUBTOTAL(109,SalesT[[#Totals],[Units]])", false, "SUBTOTAL(109,SalesT[[#Totals],[Qty]])"},
		{"COUNTA(SalesT[#Data])", false, "COUNTA(SalesT[#Data])"},
		{"SUM(OtherT[Units])", false, "SUM(OtherT[Units])"},                     // another table
		{"SUBTOTAL(109,[Units])", false, "SUBTOTAL(109,[Units])"},               // outside: not ours
		{"SUBTOTAL(109,[Units])", true, "SUBTOTAL(109,[Qty])"},                  // inside the table
		{"[@Units]*[@Price]", true, "[@Qty]*[@[Unit price]]"},                   // a space needs brackets
		{"[@[Units]]+[@Region]", true, "[@[Qty]]+[@Region]"},                    // others kept
		{"SalesT[a'#b]", false, "SalesT[C'[1']]"},                               // escapes both ways
		{`"SalesT[Units]"&SalesT[Units]`, false, `"SalesT[Units]"&SalesT[Qty]`}, // strings kept
		{"'My [Units] sheet'!A1+[1]Units!B2", true, "'My [Units] sheet'!A1+[1]Units!B2"},
		{"SUM([Units", true, "SUM([Units"}, // unbalanced: unchanged
	} {
		if got := RenameTableColumns(tt.in, "SalesT", tt.inTable, renames); got != tt.want {
			t.Errorf("RenameTableColumns(%q, %v) = %q, want %q", tt.in, tt.inTable, got, tt.want)
		}
	}
}

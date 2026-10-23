package formula

import "testing"

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

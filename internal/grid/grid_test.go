package grid

import "testing"

func TestColumns(t *testing.T) {
	for i, want := range map[int]string{1: "A", 26: "Z", 27: "AA", 702: "ZZ", 703: "AAA", 16384: "XFD"} {
		if got := ColName(i); got != want {
			t.Errorf("ColName(%d) = %s", i, got)
		}
		if got, ok := ColIndex(want); !ok || got != i {
			t.Errorf("ColIndex(%s) = %d", want, got)
		}
	}
	if _, ok := ColIndex("XFE"); ok {
		t.Error("XFE accepted")
	}
}

func TestParseRef(t *testing.T) {
	isSheet := func(s string) bool { return s == "Q1" || s == "My Sheet" }
	tests := []struct {
		in, sheet, rng string
	}{
		{"Sales!A1:C9", "Sales", "A1:C9"},
		{"'My Sheet'!B2", "My Sheet", "B2"},
		{"My Sheet!B2", "My Sheet", "B2"},
		{"A1:C9", "", "A1:C9"},
		{"C9:A1", "", "A1:C9"},
		{"Q1", "Q1", ""},
		{"B7", "", "B7"},
		{"Sales", "Sales", ""},
		{"Data!$A$1", "Data", "A1"},
	}
	for _, tt := range tests {
		r, err := ParseRef(tt.in, isSheet)
		if err != nil {
			t.Errorf("ParseRef(%q): %v", tt.in, err)
			continue
		}
		got := ""
		if r.Range != nil {
			got = r.Range.String()
		}
		if r.Sheet != tt.sheet || got != tt.rng {
			t.Errorf("ParseRef(%q) = %q %q", tt.in, r.Sheet, got)
		}
	}
	r, _ := ParseRef("Data!A:C", nil)
	if r.Range.C1 != 1 || r.Range.C2 != 3 || r.Range.R2 != MaxRows {
		t.Errorf("whole columns = %+v", r.Range)
	}
	if QuoteSheet("My Sheet") != "'My Sheet'" || QuoteSheet("Sales") != "Sales" || QuoteSheet("Q1") != "'Q1'" {
		t.Error("QuoteSheet")
	}
}

func TestClassifyFormat(t *testing.T) {
	tests := []struct {
		code string
		kind FormatKind
		dec  int
		sym  string
	}{
		{"General", General, 0, ""},
		{"0", Integer, 0, ""},
		{"#,##0.00_);(#,##0.00)", Decimal, 2, ""},
		{"0.0%", Percent, 1, ""},
		{`"$"#,##0.00_);("$"#,##0.00)`, Currency, 2, "$"},
		{`#,##0.00 "€"`, Currency, 2, "€"},
		{`[$€-407] #,##0`, Currency, 0, "€"},
		{`#,##0 "kr"`, Currency, 0, "kr"},
		{`#,##0.0 "kg"`, Decimal, 1, ""},
		{"yyyy-mm-dd", Date, 0, ""},
		{"mm/dd/yyyy", Date, 0, ""},
		{"d-mmm-yy", Date, 0, ""},
		{"mmm-yy", Date, 0, ""},
		{"h:mm AM/PM", Time, 0, ""},
		{"[h]:mm:ss", Time, 1, ""},
		{"yyyy-mm-dd hh:mm", DateTime, 0, ""},
		{"0.00E+00", Decimal, 2, ""},
		{"@", TextFmt, 0, ""},
		{`_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)`, Currency, 2, "$"},
		{"[Red]0.00", Decimal, 2, ""},
	}
	for _, tt := range tests {
		f := ClassifyFormat(tt.code)
		if f.Kind != tt.kind || f.Decimals != tt.dec || f.Symbol != tt.sym {
			t.Errorf("ClassifyFormat(%q) = %v %d %q", tt.code, f.Kind, f.Decimals, f.Symbol)
		}
	}
}

func TestDisplay(t *testing.T) {
	tests := []struct {
		c    Cell
		f    Format
		want string
	}{
		{Cell{Kind: Number, Num: 0.1 + 0.2}, Format{}, "0.3"},
		{Cell{Kind: Number, Num: 1234567.891}, Format{}, "1234567.891"},
		{Cell{Kind: Number, Num: 46204}, Format{Kind: Date}, "2026-07-01"},
		{Cell{Kind: Number, Num: 46204.5625}, Format{Kind: DateTime}, "2026-07-01 13:30"},
		{Cell{Kind: Number, Num: 0.125}, Format{Kind: Percent}, "12.5%"},
		{Cell{Kind: Bool, Bool: true}, Format{}, "TRUE"},
		{Cell{Kind: Number, Num: 1e-12}, Format{}, "1e-12"},
	}
	for _, tt := range tests {
		if got := Display(&tt.c, tt.f, false); got != tt.want {
			t.Errorf("Display(%+v) = %q, want %q", tt.c, got, tt.want)
		}
	}
}

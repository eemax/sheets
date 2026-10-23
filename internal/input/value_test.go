package input

import (
	"strings"
	"testing"

	"github.com/eemax/sheets/internal/grid"
)

func TestParse(t *testing.T) {
	tests := []struct {
		in   string
		kind grid.Kind
		num  float64
		fmt  grid.FormatKind
		text string
	}{
		{"42", grid.Number, 42, grid.General, ""},
		{"-3.5", grid.Number, -3.5, grid.General, ""},
		{"1,234.50", grid.Number, 1234.5, grid.Decimal, ""},
		{"1,234", grid.Number, 1234, grid.Integer, ""},
		{"(1,200)", grid.Number, -1200, grid.Integer, ""},
		{"12.5%", grid.Number, 0.125, grid.Percent, ""},
		{"$1,200.50", grid.Number, 1200.5, grid.Currency, ""},
		{"-$5", grid.Number, -5, grid.Currency, ""},
		{"€99", grid.Number, 99, grid.Currency, ""},
		{"1 200 kr", grid.Text, 0, grid.General, "1 200 kr"},
		{"250 kr", grid.Number, 250, grid.Currency, ""},
		{"2026-07-01", grid.Number, 46204, grid.Date, ""},
		{"2026-07-01 13:30", grid.Number, 46204 + 13.5/24, grid.DateTime, ""},
		{"13:30", grid.Number, 13.5 / 24, grid.Time, ""},
		{"007", grid.Text, 0, grid.General, "007"},
		{"12345678901234567890", grid.Text, 0, grid.General, "12345678901234567890"},
		{"'42", grid.Text, 0, grid.General, "42"},
		{"hello", grid.Text, 0, grid.General, "hello"},
		{"0.5", grid.Number, 0.5, grid.General, ""},
		{"1e5", grid.Text, 0, grid.General, "1e5"},
		{"2026-13-45", grid.Text, 0, grid.General, "2026-13-45"},
	}
	for _, tt := range tests {
		v := Parse(tt.in, false)
		if v.Kind != tt.kind || v.Fmt.Kind != tt.fmt || v.Text != tt.text || diff(v.Num, tt.num) > 1e-9 {
			t.Errorf("Parse(%q) = %+v", tt.in, v)
		}
	}
	if v := Parse("=SUM(A1:A3)", false); v.Formula != "SUM(A1:A3)" {
		t.Errorf("formula = %+v", v)
	}
	if v := Parse("TRUE", false); v.Kind != grid.Bool || !v.Bool {
		t.Errorf("bool = %+v", v)
	}
	if v := Parse("=A1", true); v.Kind != grid.Text || v.Text != "=A1" {
		t.Errorf("raw = %+v", v)
	}
}

func diff(a, b float64) float64 {
	if a > b {
		return a - b
	}
	return b - a
}

func TestTable(t *testing.T) {
	rows, err := Table(strings.NewReader("\xef\xbb\xbfa\tb\n1\tC:\\new\n"), "auto", false)
	if err != nil || len(rows) != 2 || rows[0][0].Text != "a" || rows[1][1].Text != `C:\new` {
		t.Fatalf("tsv = %+v, %v", rows, err)
	}
	rows, err = Table(strings.NewReader("a,b\n\"1,5\",2\n"), "auto", false)
	// "1,5" is not thousands grouping; guessing a decimal comma is unsafe.
	if err != nil || len(rows) != 2 || rows[1][0].Text != "1,5" || rows[1][1].Num != 2 {
		t.Fatalf("csv = %+v, %v", rows, err)
	}
	rows, err = Table(strings.NewReader(`[{"b":1,"a":"x"},{"a":"y","c":true}]`), "auto", false)
	if err != nil {
		t.Fatal(err)
	}
	var header []string
	for _, v := range rows[0] {
		header = append(header, v.Text)
	}
	if strings.Join(header, ",") != "b,a,c" || len(rows) != 3 || !rows[2][2].Bool {
		t.Fatalf("json objects = %+v", rows)
	}
	rows, err = Table(strings.NewReader(`[[1,"2026-01-02",null],[12345678901234567890]]`), "auto", false)
	if err != nil || rows[0][1].Fmt.Kind != grid.Date || !rows[0][2].IsEmpty() || rows[1][0].Kind != grid.Text {
		t.Fatalf("json arrays = %+v, %v", rows, err)
	}
}

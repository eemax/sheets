package diff

import (
	"reflect"
	"testing"
)

// A column inserted between the columns a formula column reads: Excel
// rewrites =B2*C2 to =B2*D2, so the formulas' relative reads change and the
// column is paired with its old self only with the option.
func TestColumnShifts(t *testing.T) {
	before := [][]any{
		{"Region", "Units", "Price", "Revenue"},
		{"North", 5, 30, "=B2*C2"},
		{"South", 7, 20, "=B3*C3"},
		{"East", 3, 50, "=B4*C4"},
		{"Total", "", "", "=SUM(D2:D4)"},
	}
	after := [][]any{
		{"Region", "Units", "Discount", "Price", "Revenue"},
		{"North", 5, 0, 30, "=B2*D2"},
		{"South", 7, 0, 20, "=B3*D3"},
		{"East", 3, 0, 50, "=B4*D4"},
		{"Total", "", "", "", "=SUM(E2:E4)"},
	}
	pairs := func(opts Options) []Pair {
		return Align(book("Data", before), book("Data", after), opts).Sheets[0].Cols
	}
	// Left off, as in sheets diff: Revenue is a column removed and another
	// added.
	if got := pairs(Options{}); !reflect.DeepEqual(got, []Pair{{1, 1}, {2, 2}, {0, 3}, {3, 4}, {4, 0}, {0, 5}}) {
		t.Errorf("without the option: %v", got)
	}
	want := []Pair{{1, 1}, {2, 2}, {0, 3}, {3, 4}, {4, 5}}
	if got := pairs(Options{ColumnShifts: true}); !reflect.DeepEqual(got, want) {
		t.Errorf("columns %v, want %v", got, want)
	}
	// Its rows still pair in place.
	al := Align(book("Data", before), book("Data", after), Options{ColumnShifts: true})
	for _, p := range al.Sheets[0].Rows {
		if p.Old != p.New {
			t.Errorf("rows %v", al.Sheets[0].Rows)
			break
		}
	}
	// A column whose formulas do another thing (+ for *) is not the old one
	// moved right.
	changed := [][]any{
		after[0],
		{"North", 5, 0, 30, "=B2+D2"},
		{"South", 7, 0, 20, "=B3+D3"},
		{"East", 3, 0, 50, "=B4+D4"},
		after[4],
	}
	if got := Align(book("Data", before), book("Data", changed), Options{ColumnShifts: true}).Sheets[0].Cols; reflect.DeepEqual(got, want) {
		t.Errorf("a column of another formula paired as the same: %v", got)
	}
}

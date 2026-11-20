package diff

import (
	"reflect"
	"testing"
)

// A column deleted that another column's formulas read: Excel rewrites them
// with #REF! and the column moves left.
func lostColumns() (before, after [][]any) {
	before = [][]any{
		{"Region", "Units", "Price", "Revenue"},
		{"North", 5, 30, "=B2*C2"},
		{"South", 7, 20, "=B3*C3"},
		{"East", 3, 50, "=B4*C4"},
		{"Total", "", "", "=SUM(D2:D4)"},
	}
	after = [][]any{
		{"Region", "Units", "Revenue"},
		{"North", 5, "=B2*#REF!"},
		{"South", 7, "=B3*#REF!"},
		{"East", 3, "=B4*#REF!"},
		{"Total", "", "=SUM(C2:C4)"},
	}
	return before, after
}

func TestLostRefsColumns(t *testing.T) {
	before, after := lostColumns()
	pairs := func(opts Options) []Pair {
		al := Align(book("Data", before), book("Data", after), opts)
		return al.Sheets[0].Cols
	}
	// Left off, as in sheets diff: the rewritten column is a column added,
	// and the old two removed.
	if got, want := pairs(Options{}), []Pair{{1, 1}, {2, 2}, {3, 0}, {4, 0}, {0, 3}}; !reflect.DeepEqual(got, want) {
		t.Errorf("without the option: %v, want %v", got, want)
	}
	want := []Pair{{1, 1}, {2, 2}, {3, 0}, {4, 3}}
	if got := pairs(Options{LostRefs: true}); !reflect.DeepEqual(got, want) {
		t.Errorf("columns %v, want %v", got, want)
	}
}

// A row whose formulas all lost a reference is the same row.
func TestLostRefsRows(t *testing.T) {
	before := [][]any{
		{"a", "b"},
		{"=A3*2", "=A3*3"},
		{"x", ""},
		{"y", ""},
	}
	after := [][]any{
		{"a", "b"},
		{"=#REF!*2", "=#REF!*3"},
		{"y", ""},
	}
	rows := func(opts Options) []Pair {
		return Align(book("S", before), book("S", after), opts).Sheets[0].Rows
	}
	want := []Pair{{1, 1}, {2, 2}, {3, 0}, {4, 3}}
	if got := rows(Options{LostRefs: true}); !reflect.DeepEqual(got, want) {
		t.Errorf("rows %v, want %v", got, want)
	}
}

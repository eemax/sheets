package engine

import (
	"strings"
	"testing"

	"github.com/eemax/sheets/internal/grid"
)

// IFS and SWITCH go through an array element by element, as Excel does in
// an array formula; INDEX over an array gives the values it refers to.
func TestLiftedInArrays(t *testing.T) {
	b := newBook(map[string]any{
		"A1": 1, "A2": 0, "A3": 3,
		"B1": 10, "B2": 20, "B3": 30,
		"D1": `=IFS(A1:A3>1,"a",TRUE,"b")`,
		"E1": `=SWITCH(A1:A3,1,"x","y")`,
		"F1": `=INDEX(B1:B3,A1:A3+1)`,
		"G1": `=SWITCH(A1:A3,1,"x")`,
	})
	for _, c := range []string{"D1", "E1", "F1", "G1"} {
		r, col, _ := grid.ParseCell(c)
		b.arrays["Sheet1!"+c] = Area{r, col, r + 2, col}
	}
	e := New(b)
	for cell, want := range map[string]string{
		"D1": `"b"`, "D2": `"b"`, "D3": `"a"`,
		"E1": `"x"`, "E2": `"y"`, "E3": `"y"`,
		"F1": "20", "F2": "10", "F3": "#REF!",
		"G1": `"x"`, "G2": "#N/A", "G3": "#N/A",
	} {
		r, c, _ := grid.ParseCell(cell)
		if got := show(e.cellValue("Sheet1", r, c)); got != want {
			t.Errorf("%s: got %s, want %s", cell, got, want)
		}
	}
}

// A function that takes one value of an array in an array formula, where
// Excel would go through it element by element, makes the result unknown
// rather than the first element's.
func TestCollapseUnknown(t *testing.T) {
	b := newBook(map[string]any{"A1": 1, "A2": 2})
	e := New(b)
	var failed *Value
	c := &ctx{e: e, sheet: "Sheet1", row: 5, col: 5, arr: true, failed: &failed}
	v := c.scalar(Value{Kind: Array, Arr: &Arr{Rows: 2, Cols: 1, V: []Value{num(1), num(2)}}})
	if v.Kind != Error || v.Str != ErrUnsupported || failed == nil || !strings.Contains(v.Why(), "element by element") {
		t.Errorf("got %s (%s)", show(v), v.Why())
	}
	// Outside an array formula, an array's first value, as before.
	c.arr, failed = false, nil
	if v := c.scalar(Value{Kind: Array, Arr: &Arr{Rows: 2, Cols: 1, V: []Value{num(1), num(2)}}}); show(v) != "1" {
		t.Errorf("outside: %s", show(v))
	}
}

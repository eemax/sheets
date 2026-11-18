package engine

import (
	"fmt"
	"testing"

	"github.com/eemax/sheets/internal/grid"
)

// The classes WP-01 found the engine computing unlike Excel, each with
// Excel's answer (or unknown, where only Excel's language setting
// decides).

// COLUMN over an IF in an array formula: an error element stays itself,
// a value is #VALUE!, so MATCH(3,ERROR.TYPE(...),0) finds the first cell
// the IF kept.
func TestErrorTypeOfColumnOfIf(t *testing.T) {
	f := `=IFNA(INDEX(A%[1]d:C%[1]d,MATCH(3,ERROR.TYPE(COLUMN(IF(A%[1]d:C%[1]d<>"~N/A",A%[1]d:C%[1]d))),0)),"")`
	cells := map[string]any{
		"A2": "Completed", "B2": "#N/A", "C2": "#N/A",
		"A3": "#N/A", "B3": 44635, "C3": "#N/A",
		"A4": "#N/A", "B4": "#N/A", "C4": "ENR",
		"A5": "#N/A", "B5": "#N/A", "C5": "#N/A",
	}
	for r := 2; r <= 5; r++ {
		cells[grid.CellName(r, 4)] = fmt.Sprintf(f, r)
	}
	b := newBook(cells)
	for r := 2; r <= 5; r++ {
		b.arrays["Sheet1!"+grid.CellName(r, 4)] = Area{r, 4, r, 4}
	}
	e := New(b)
	for r, want := range map[int]string{2: `"Completed"`, 3: "44635", 4: `"ENR"`, 5: `""`} {
		if got := show(e.Compute("Sheet1", r, 4)); got != want {
			t.Errorf("D%d: got %s, want %s", r, got, want)
		}
	}
	for formula, want := range map[string]string{
		"COLUMN({1,2})":         "#VALUE!",
		"ERROR.TYPE(COLUMN(1))": "3",
	} {
		if got := show(e.Eval("Sheet1", 1, 9, formula)); got != want {
			t.Errorf("=%s: got %s, want %s", formula, got, want)
		}
	}
}

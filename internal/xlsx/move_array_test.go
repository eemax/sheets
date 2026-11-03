package xlsx

import (
	"strings"
	"testing"

	"github.com/eemax/sheets/internal/grid"
)

// A moved array formula marked to calculate always (ca="1") leaves no
// copy behind (excelize kept it when the cell was emptied).
func TestMoveLeavesNoFormula(t *testing.T) {
	data := checkData(t, "x.xlsx", openFixture(t, "arrays_xlsxwriter.xlsx")).Data
	data = patchZip(t, data, dataPart, func(x string) string {
		if !strings.Contains(x, `<f t="array" ref="G2">`) {
			t.Fatalf("no array formula in G2: %s", x)
		}
		return strings.Replace(x, `<f t="array" ref="G2">`, `<f t="array" ref="G2" ca="1">`, 1)
	})
	e, err := OpenEditorData("x.xlsx", data)
	must(t, err)
	if _, err := e.MoveRange("S", grid.Range{R1: 2, C1: 7, R2: 2, C2: 7}, "S", grid.Range{R1: 9, C1: 8, R2: 9, C2: 8}); err != nil {
		t.Fatal(err)
	}
	out, err := e.Render()
	must(t, err)
	wb, err := ReadBytes(out)
	must(t, err)
	if c := wb.Sheet("S").Get(2, 7); c.Formula != "" {
		t.Errorf("G2 still holds =%s", c.Formula)
	}
	if c := wb.Sheet("S").Get(9, 8); c.Formula == "" {
		t.Errorf("H9 holds no formula")
	}
}

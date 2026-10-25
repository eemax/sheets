package xlsx

import (
	"fmt"
	"os"
	"strconv"
	"strings"
	"testing"

	"github.com/eemax/sheets/internal/engine"
	"github.com/eemax/sheets/internal/grid"
)

// TestEngineProbe evaluates expressions with the sheets calculator, as if in a
// cell (SHEETS_PROBE_CELL, default A1):
//
//	SHEETS_PROBE_FILE=book.xlsx SHEETS_PROBE_SHEET=Sheet1 SHEETS_PROBE_CELL=I3 \
//	SHEETS_PROBE='ROW(B2:B13)|SUM(A1:A3)' go test ./internal/xlsx -run TestEngineProbe -v
func TestEngineProbe(t *testing.T) {
	path := os.Getenv("SHEETS_PROBE_FILE")
	if path == "" {
		t.Skip("SHEETS_PROBE_FILE names the workbook")
	}
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	wb, _, _, err := parse(data)
	if err != nil {
		t.Fatal(err)
	}
	rs := wbResolver(wb)
	book := &gridBook{wb: wb, text: func(sheet string, row, col int, f string) string {
		return rs.calcText(&formulaAt{sheet: sheet, row: row, col: col, text: f})
	}}
	eng := engine.New(book)
	sheet := os.Getenv("SHEETS_PROBE_SHEET")
	if sheet == "" {
		sheet = wb.Sheets[0].Name
	}
	cell := os.Getenv("SHEETS_PROBE_CELL")
	if cell == "" {
		cell = "A1"
	}
	r, c, _ := grid.ParseCell(cell)
	for _, expr := range strings.Split(os.Getenv("SHEETS_PROBE"), "|") {
		v := eng.Eval(sheet, r, c, expr)
		t.Logf("%-50s → %s %q %v %s", expr, kindName(v.Kind), v.Str, v.Num, arrText(v))
	}
}

func kindName(k engine.Kind) string {
	return []string{"empty", "number", "text", "bool", "error", "array", "ref"}[k]
}

func arrText(v engine.Value) string {
	if v.Kind != engine.Array || v.Arr == nil {
		return ""
	}
	var parts []string
	for _, x := range v.Arr.V[:min(len(v.Arr.V), 20)] {
		switch x.Kind {
		case engine.Number:
			parts = append(parts, strconv.FormatFloat(x.Num, 'g', -1, 64))
		case engine.Empty:
			parts = append(parts, "∅")
		default:
			parts = append(parts, x.Str)
		}
	}
	return fmt.Sprintf("%dx%d [%s]", v.Arr.Rows, v.Arr.Cols, strings.Join(parts, " "))
}

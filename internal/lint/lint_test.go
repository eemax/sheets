package lint

import (
	"strings"
	"testing"

	"github.com/eemax/sheets/internal/grid"
)

// sheet builds a workbook from rows: strings starting with = are formulas
// (value given after |, e.g. "=A1/0|#DIV/0!"), other strings text, ints and
// floats numbers.
func book(rows ...[]any) *grid.Workbook {
	s := grid.NewSheet("S")
	for r, row := range rows {
		for c, v := range row {
			var cell grid.Cell
			switch x := v.(type) {
			case nil:
				continue
			case string:
				if strings.HasPrefix(x, "=") {
					f, val, _ := strings.Cut(x[1:], "|")
					cell = grid.Cell{Formula: f, Kind: grid.Number}
					if strings.HasPrefix(val, "#") {
						cell.Kind, cell.Text = grid.Error, val
					}
				} else {
					cell = grid.Cell{Kind: grid.Text, Text: x}
				}
			case int:
				cell = grid.Cell{Kind: grid.Number, Num: float64(x)}
			case float64:
				cell = grid.Cell{Kind: grid.Number, Num: x}
			}
			s.Set(r+1, c+1, cell)
		}
	}
	return &grid.Workbook{Sheets: []*grid.Sheet{s}, Formats: []grid.Format{{}}}
}

func rules(fs []Finding) map[string][]Finding {
	m := map[string][]Finding{}
	for _, f := range fs {
		m[f.Rule] = append(m[f.Rule], f)
	}
	return m
}

func TestFindings(t *testing.T) {
	wb := book(
		[]any{"Item", "Units", "Price", "Revenue", "Note"},
		[]any{"a", 1, 2.5, "=B2*C2", "ok"},
		[]any{"b", 2, 2.5, "=B3*C3", " padded "},
		[]any{"c", "1,200", 2.5, "=B4*C4", "x"},
		[]any{"d", 4, 2.5, "=B5*C5+1", "y"},
		[]any{"e", 5, 2.5, 99, "z"},
		[]any{"f", 6, 0, "=B7/C7|#DIV/0!", "w"},
		[]any{"g", 7, 2.5, "=B8*C8", "v"},
		[]any{"Total", "=SUM(B2:B6)", nil, "=SUM(D2:D8)", nil},
	)
	got := rules(Check(wb, Options{}))

	if f := got["error-value"]; len(f) != 1 || f[0].Cells[0] != "D7" || !strings.Contains(f[0].Message, "#DIV/0!") {
		t.Errorf("error-value = %+v", f)
	}
	inc := got["inconsistent-formula"]
	var cells []string
	for _, f := range inc {
		cells = append(cells, f.Cells[0])
	}
	if strings.Join(cells, ",") != "D5,D6,D7" {
		t.Errorf("inconsistent-formula cells = %v", cells)
	}
	if f := got["number-as-text"]; len(f) != 1 || f[0].Cells[0] != "B4" || !strings.Contains(f[0].Message, "1 value in Units is a number stored as text") {
		t.Errorf("number-as-text = %+v", f)
	}
	if f := got["total-range"]; len(f) != 1 || f[0].Cells[0] != "B9" || !strings.Contains(f[0].Message, "rows 7-8 left out") {
		t.Errorf("total-range = %+v", f)
	}
	if f := got["stray-whitespace"]; len(f) != 1 || f[0].Cells[0] != "E3" {
		t.Errorf("stray-whitespace = %+v", f)
	}
	// Errors sort first.
	if all := Check(wb, Options{}); all[0].Severity != Error {
		t.Errorf("first finding = %+v", all[0])
	}
}

func TestNoFalsePositives(t *testing.T) {
	wb := book(
		[]any{"Date", "Flow", "Balance", "Ref"},
		[]any{"x", 0, 1000, "=Table1[Flow]+Other!A1"},
		[]any{"y", 50, "=C2+B3", "=Other!A1"},
		[]any{"z", -20, "=C3+B4", "=Other!A1"},
		[]any{"w", 10, "=C4+B5", "=Other!A1"},
	)
	got := rules(Check(wb, Options{}))
	if f := got["inconsistent-formula"]; len(f) > 0 {
		t.Errorf("opening balance flagged: %+v", f)
	}
	if f := got["external-link"]; len(f) > 0 {
		t.Errorf("structured reference taken for an external link: %+v", f)
	}
}

func TestReferencesAndFunctions(t *testing.T) {
	wb := book(
		[]any{"=#REF!*2|#REF!", "=NOPE(1)|#NAME?", "=[1]Prices!A1", "=TODAY()", "=B9*1.21", "=C9*1.21", "=D9*1.21"},
	)
	wb.Names = []grid.DefinedName{{Name: "Old", Ref: "S!#REF!"}}
	got := rules(Check(wb, Options{}))
	if len(got["broken-reference"]) != 2 {
		t.Errorf("broken-reference = %+v", got["broken-reference"])
	}
	if f := got["unknown-function"]; len(f) != 1 || !strings.Contains(f[0].Message, "NOPE") {
		t.Errorf("unknown-function = %+v", f)
	}
	if len(got["error-value"]) != 0 {
		t.Errorf("errors explained by other rules reported twice: %+v", got["error-value"])
	}
	if len(got["external-link"]) != 1 || len(got["volatile-function"]) != 1 {
		t.Errorf("notes = %+v", got)
	}
	if f := got["repeated-constant"]; len(f) != 1 || !strings.HasPrefix(f[0].Message, "1.21 is typed into 3 formulas") {
		t.Errorf("repeated-constant = %+v", f)
	}
}

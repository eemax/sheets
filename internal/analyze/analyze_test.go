package analyze

import (
	"testing"

	"github.com/eemax/sheets/internal/grid"
)

func sheet(rows ...[]any) (*grid.Workbook, *grid.Sheet) {
	s := grid.NewSheet("S")
	for r, row := range rows {
		for c, v := range row {
			var cell grid.Cell
			switch x := v.(type) {
			case nil:
				continue
			case string:
				if len(x) > 1 && x[0] == '=' {
					cell = grid.Cell{Formula: x[1:], Kind: grid.Number}
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
	return &grid.Workbook{Sheets: []*grid.Sheet{s}, Formats: []grid.Format{{}}}, s
}

func TestRegions(t *testing.T) {
	wb, s := sheet(
		[]any{"Report"},
		nil,
		[]any{"Name", "Qty", "Price"},
		[]any{"a", 1, 2.5},
		[]any{"b", 3, 4.5},
		[]any{"Total", "=SUM(B4:B5)", nil},
		nil,
		[]any{"note"},
		nil,
		[]any{nil, nil, nil, nil, "x", "y"},
		[]any{nil, nil, nil, nil, "z", "w"},
	)
	regs := Regions(wb, s)
	if len(regs) != 4 {
		t.Fatalf("got %d regions: %+v", len(regs), regs)
	}
	if !regs[0].IsTitle {
		t.Errorf("A1 should caption the table: %+v", regs[0])
	}
	tb := regs[1]
	if !tb.IsTable() || tb.HeaderRow != 3 || tb.TotalRow != 6 || tb.Body.String() != "A4:C5" {
		t.Errorf("table = %+v", tb)
	}
	if regs[2].IsTitle || regs[2].Range.String() != "A8" {
		t.Errorf("note = %+v", regs[2])
	}
	if regs[3].IsTable() || regs[3].Range.String() != "E10:F11" {
		t.Errorf("text block = %+v", regs[3])
	}

	cols := Profile(wb, s, tb)
	if cols[0].Type != "text" || cols[1].Type != "integer" || cols[2].Type != "decimal" {
		t.Errorf("types = %s %s %s", cols[0].Type, cols[1].Type, cols[2].Type)
	}
	if cols[1].Sum != 4 || cols[2].Min != 2.5 || cols[2].Max != 4.5 {
		t.Errorf("stats = %+v %+v", cols[1], cols[2])
	}
}

func TestTitleRowsInsideRegion(t *testing.T) {
	wb, s := sheet(
		[]any{"Q3"},
		[]any{"Region", "Units"},
		[]any{"North", 5},
		[]any{"South", 7},
	)
	r := Regions(wb, s)[0]
	if r.TitleRows != 1 || r.HeaderRow != 2 || r.Body.String() != "A3:B4" || r.TotalRow != 0 {
		t.Errorf("region = %+v", r)
	}
}

func TestFormulaGroups(t *testing.T) {
	wb, s := sheet(
		[]any{"A", "B", "C"},
		[]any{1, 2, "=A2+B2"},
		[]any{3, 4, "=A3+B3"},
		[]any{5, 6, "=A4*B4"},
	)
	cols := Profile(wb, s, Regions(wb, s)[0])
	g := cols[2].Formulas
	if len(g) != 2 || g[0].Pattern != "A{r}+B{r}" || g[0].Range.String() != "C2:C3" || g[1].Range.String() != "C4" {
		t.Errorf("groups = %+v", g)
	}
}

func TestHeaderShapes(t *testing.T) {
	tests := []struct {
		name   string
		rows   [][]any
		header int
		title  int
	}{
		{"years as labels", [][]any{{"Region", 2024, 2025}, {"North", 5, 6}, {"South", 7, 8}}, 1, 0},
		{"duplicate labels", [][]any{{"Name", "Amount", "Amount"}, {"a", 1, 2}, {"b", 3, 4}}, 1, 0},
		{"crosstab corner", [][]any{{nil, "Total"}, {"North", 5}, {"South", 7}}, 1, 0},
		{"numbers are data", [][]any{{2024, 5}, {2025, 6}}, 0, 0},
		{"title over table", [][]any{{"Q3"}, {"Region", "Units"}, {"North", 5}}, 2, 1},
		{"all-text, distinct", [][]any{{"Name", "City"}, {"Ada", "London"}, {"Bob", "Paris"}}, 1, 0},
		{"all-text, too short", [][]any{{"x", "y"}, {"z", "w"}}, 0, 0},
	}
	for _, tt := range tests {
		wb, s := sheet(tt.rows...)
		r := Regions(wb, s)[0]
		if r.HeaderRow != tt.header || r.TitleRows != tt.title {
			t.Errorf("%s: header %d title %d, want %d %d", tt.name, r.HeaderRow, r.TitleRows, tt.header, tt.title)
		}
	}
}

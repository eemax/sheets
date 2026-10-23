package diff

import (
	"strings"
	"testing"

	"github.com/eemax/sheets/internal/grid"
)

func seq(s string) []uint64 {
	out := make([]uint64, len(s))
	for i, c := range s {
		out[i] = uint64(c)
	}
	return out
}

func TestAlign(t *testing.T) {
	m := align(seq("abcdef"), seq("abXcdeYf"))
	var got []string
	for _, p := range m {
		got = append(got, string("abcdef"[p[0]]))
	}
	if strings.Join(got, "") != "abcdef" {
		t.Errorf("matched %v", got)
	}
	steps := pairUp(3, 3, align(seq("aQc"), seq("aRc")), func(i, j int) bool { return true })
	if len(steps) != 3 || steps[1] != (step{1, 1}) {
		t.Errorf("changed middle line: %v", steps)
	}
	steps = pairUp(2, 3, align(seq("aQ"), seq("aNQ")), func(i, j int) bool { return i == 1 && j == 2 })
	if len(steps) != 3 || steps[1] != (step{-1, 1}) || steps[2] != (step{1, 2}) {
		t.Errorf("insert before a changed line: %v", steps)
	}
}

// book builds a workbook; sheets are given as name, rows.
func book(sheets ...any) *grid.Workbook {
	wb := &grid.Workbook{Formats: []grid.Format{{}, {Kind: grid.Currency, Symbol: "$", Decimals: 2}}}
	for i := 0; i < len(sheets); i += 2 {
		s := grid.NewSheet(sheets[i].(string))
		for r, row := range sheets[i+1].([][]any) {
			for c, v := range row {
				switch x := v.(type) {
				case string:
					if strings.HasPrefix(x, "=") {
						s.Set(r+1, c+1, grid.Cell{Formula: x[1:], Pending: true})
					} else {
						s.Set(r+1, c+1, grid.Cell{Kind: grid.Text, Text: x})
					}
				case int:
					s.Set(r+1, c+1, grid.Cell{Kind: grid.Number, Num: float64(x)})
				case float64: // currency
					s.Set(r+1, c+1, grid.Cell{Kind: grid.Number, Num: x, Fmt: 1})
				}
			}
		}
		wb.Sheets = append(wb.Sheets, s)
	}
	return wb
}

func TestWorkbooks(t *testing.T) {
	base := [][]any{
		{"Item", "Qty", "Total"},
		{"a", 1, "=B2*2"},
		{"b", 2, "=B3*2"},
		{"c", 3, "=B4*2"},
		{"Sum", "=SUM(B2:B4)", "=SUM(C2:C4)"},
	}
	old := book("Data", base, "Notes", [][]any{{"x"}})

	t.Run("same", func(t *testing.T) {
		if r := Workbooks(old, book("Data", base, "Notes", [][]any{{"x"}}), Options{}); !r.Same {
			t.Errorf("%+v", r)
		}
	})
	t.Run("value and formula", func(t *testing.T) {
		changed := [][]any{
			{"Item", "Qty", "Total"},
			{"a", 1, "=B2*2"},
			{"b", 20, "=B3*2"},
			{"c", 3, "=B4*3"},
			{"Sum", "=SUM(B2:B4)", "=SUM(C2:C4)"},
		}
		r := Workbooks(old, book("Data", changed, "Notes", [][]any{{"x"}}), Options{})
		if len(r.Sheets) != 1 {
			t.Fatalf("%+v", r)
		}
		var got []string
		for _, c := range r.Sheets[0].Cells {
			got = append(got, c.Cell+":"+c.Old+">"+c.New)
		}
		if strings.Join(got, " ") != "B3:2>20 C4:=B4*2>=B4*3" {
			t.Errorf("cells = %v", got)
		}
	})
	t.Run("rows inserted and deleted", func(t *testing.T) {
		moved := [][]any{
			{"Item", "Qty", "Total"},
			{"a", 1, "=B2*2"},
			{"new", 9, "=B3*2"},
			{"c", 3, "=B4*2"},
			{"Sum", "=SUM(B2:B4)", "=SUM(C2:C4)"},
		}
		r := Workbooks(old, book("Data", moved, "Notes", [][]any{{"x"}}), Options{})
		d := r.Sheets[0]
		if len(d.RowsAdded) != 1 || d.RowsAdded[0].From != 3 || len(d.RowsRemoved) != 1 || d.RowsRemoved[0].From != 3 {
			t.Errorf("rows: %+v / %+v", d.RowsAdded, d.RowsRemoved)
		}
		if len(d.Cells) != 0 {
			t.Errorf("cells: %+v", d.Cells)
		}
	})
	t.Run("column inserted, sheet renamed, order", func(t *testing.T) {
		wide := [][]any{
			{"Item", "Note", "Qty", "Total"},
			{"a", nil, 1, "=C2*2"},
			{"b", nil, 2, "=C3*2"},
			{"c", nil, 3, "=C4*2"},
			{"Sum", nil, "=SUM(C2:C4)", "=SUM(D2:D4)"},
		}
		r := Workbooks(old, book("Notes", [][]any{{"x"}}, "Stock", wide), Options{})
		var d *Sheet
		for i := range r.Sheets {
			if r.Sheets[i].Name == "Stock" {
				d = &r.Sheets[i]
			}
		}
		if d == nil || d.RenamedFrom != "Data" || len(d.ColsAdded) != 1 || d.ColsAdded[0].From != 2 || len(d.Cells) != 0 {
			t.Fatalf("%+v", r.Sheets)
		}
		if r.Order == nil {
			t.Error("order change missed")
		}
	})
	t.Run("formats", func(t *testing.T) {
		priced := [][]any{
			{"Item", "Qty", "Total"},
			{"a", 1.0, "=B2*2"},
			{"b", 2.0, "=B3*2"},
			{"c", 3, "=B4*2"},
			{"Sum", "=SUM(B2:B4)", "=SUM(C2:C4)"},
		}
		r := Workbooks(old, book("Data", priced, "Notes", [][]any{{"x"}}), Options{})
		d := r.Sheets[0]
		if len(d.Cells) != 0 || len(d.Formats) != 1 || d.Formats[0].Range != "B2:B3" || d.Formats[0].New != "currency $" {
			t.Errorf("%+v", d)
		}
	})
	t.Run("sheets and names", func(t *testing.T) {
		old.Names = []grid.DefinedName{{Name: "Rate", Ref: "Data!$B$2"}}
		nw := book("Data", base, "Extra", [][]any{{"y"}})
		nw.Names = []grid.DefinedName{{Name: "Rate", Ref: "Data!$B$3"}}
		r := Workbooks(old, nw, Options{})
		statuses := map[string]string{}
		for _, s := range r.Sheets {
			statuses[s.Name] = s.Status
		}
		if statuses["Extra"] != "added" || statuses["Notes"] != "removed" || len(r.Names) != 1 || r.Names[0].New != "Data!$B$3" {
			t.Errorf("%+v %+v", r.Sheets, r.Names)
		}
	})
}

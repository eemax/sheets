package diff

import (
	"testing"

	"github.com/eemax/sheets/internal/grid"
)

// withResult is a sheet's book with an array formula's result over rg.
func withResult(name string, rows [][]any, rg grid.Range) *grid.Workbook {
	wb := book(name, rows)
	wb.Sheets[0].Spills = []grid.Range{rg}
	wb.Sheets[0].Arrays = []grid.Range{rg}
	return wb
}

// filtered is a FILTER's result in A2 over two columns (its values in
// order), with a SUM of it in D2, and a note in row 10 when noted.
func filtered(values []any, noted bool) *grid.Workbook {
	rows := [][]any{{"Big"}, {"=FILTER(Data!A2:B6,Data!B2:B6>8)", values[1], nil, "=SUM(B2:B20)"}}
	for i := 2; i < len(values); i += 2 {
		rows = append(rows, []any{values[i], values[i+1]})
	}
	if noted {
		for len(rows) < 9 {
			rows = append(rows, []any{})
		}
		rows = append(rows, []any{nil, nil, nil, "Note"})
	}
	return withResult("Report", rows, grid.Range{R1: 2, C1: 1, R2: 1 + len(values)/2, C2: 2})
}

func inPlace(t *testing.T, name string, p SheetPair, rows int) {
	t.Helper()
	for r := 1; r <= rows; r++ {
		if got := pairOf(p.Rows, r); got != r {
			t.Errorf("%s: old row %d is %d (%+v)", name, r, got, p.Rows)
		}
	}
	if len(p.Moved) != 0 || len(p.Blocks) != 0 {
		t.Errorf("%s: moved %+v, blocks %+v", name, p.Moved, p.Blocks)
	}
}

// TestSpills checks that the rows of an array formula's result pair as its
// formula puts them, not by its values (Options.Spills): a result that lost
// a row, one sorted again (with a note below, past empty rows, or not), a
// SEQUENCE at the bottom of a sheet that shrank while its own row changed,
// and rows inserted through a SEQUENCE.
func TestSpills(t *testing.T) {
	big := []any{"North", 50, "South", 40, "East", 30, "West", 20, "Mid", 10}
	lost := []any{"North", 50, "South", 40, "West", 20, "Mid", 10}
	resorted := []any{"Up", 60, "North", 50, "South", 40, "East", 30, "West", 20}
	opts := Options{Moves: true, Spills: true}
	for _, c := range []struct {
		name  string
		after []any
		rows  int // the old rows that pair in place, with no note
	}{
		{"a row lost", lost, 5},
		{"sorted again", resorted, 6},
	} {
		a, b := filtered(big, false), filtered(c.after, false)
		if p := AlignSheet(a, a.Sheets[0], b, b.Sheets[0], Options{Moves: true}); pairOf(p.Rows, 4) == 4 && len(p.Moved) == 0 {
			t.Errorf("%s: in place without Spills: the case shows nothing", c.name)
		}
		inPlace(t, c.name, AlignSheet(a, a.Sheets[0], b, b.Sheets[0], opts), c.rows)
		a, b = filtered(big, true), filtered(c.after, true)
		inPlace(t, c.name+", a note below", AlignSheet(a, a.Sheets[0], b, b.Sheets[0], opts), 10)
	}

	seq := func(n int) *grid.Workbook {
		rows := [][]any{{n, "=SEQUENCE(A1)", nil, "=SUM(B1:B5)"}}
		for i := 2; i <= n; i++ {
			rows = append(rows, []any{nil, i})
		}
		return withResult("Data", rows, grid.Range{R1: 1, C1: 2, R2: n, C2: 2})
	}
	a, b := seq(5), seq(3)
	inPlace(t, "a SEQUENCE shrank", AlignSheet(a, a.Sheets[0], b, b.Sheets[0], opts), 3)

	// Two rows inserted at 3 through a SEQUENCE in B1:B4: the labels beside
	// it moved down, the result stayed where its formula puts it.
	a = withResult("Data", [][]any{{4, "=SEQUENCE(A1)", nil, "Item"}, {nil, 2, nil, "a"}, {nil, 3, nil, "b"},
		{nil, 4, nil, "c"}, {nil, nil, nil, "d"}, {nil, nil, nil, "e"}}, grid.Range{R1: 1, C1: 2, R2: 4, C2: 2})
	b = withResult("Data", [][]any{{4, "=SEQUENCE(A1)", nil, "Item"}, {nil, 2, nil, "a"}, {nil, 3}, {nil, 4},
		{nil, nil, nil, "b"}, {nil, nil, nil, "c"}, {nil, nil, nil, "d"}, {nil, nil, nil, "e"}}, grid.Range{R1: 1, C1: 2, R2: 4, C2: 2})
	if p := AlignSheet(a, a.Sheets[0], b, b.Sheets[0], Options{Moves: true}); pairOf(p.Rows, 3) != 3 {
		t.Errorf("rows inserted through a SEQUENCE: the case shows nothing without Spills: %+v", p.Rows)
	}
	p := AlignSheet(a, a.Sheets[0], b, b.Sheets[0], opts)
	for old, now := range map[int]int{1: 1, 2: 2, 3: 5, 4: 6, 5: 7, 6: 8} {
		if got := pairOf(p.Rows, old); got != now {
			t.Errorf("rows inserted through a SEQUENCE: old row %d is %d, not %d (%+v)", old, got, now, p.Rows)
		}
	}
	for _, q := range p.Rows {
		if q.Old == 0 && q.New != 3 && q.New != 4 {
			t.Errorf("rows inserted through a SEQUENCE: row %d added", q.New)
		}
	}
}

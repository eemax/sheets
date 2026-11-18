package diff

import (
	"reflect"
	"strconv"
	"testing"

	"github.com/eemax/sheets/internal/formula"
)

// sortedBeside is sorted() with cells beside the table that stay where they
// are: a note and a formula in E, past an empty column, or in D, next to it.
func sortedBeside(col int) (before, after [][]any) {
	before, after = sorted()
	side := []any{"Best", "=MAX(C2:C6)", "Note"}
	for r := range side {
		for _, rows := range []*[][]any{&before, &after} {
			row := append([]any{}, (*rows)[r]...)
			for len(row) < col-1 {
				row = append(row, nil)
			}
			(*rows)[r] = append(row, side[r])
		}
	}
	return before, after
}

func pairOf(ps []Pair, old int) int {
	for _, p := range ps {
		if p.Old == old {
			return p.New
		}
	}
	return 0
}

// TestBlocks checks that rows a sort moved in some columns, with cells
// beside them that stayed, pair as moved in those columns and in place
// elsewhere: no rows added or removed, nothing changed.
func TestBlocks(t *testing.T) {
	for _, col := range []int{5, 4} {
		before, after := sortedBeside(col)
		a, b := book("Data", before), book("Data", after)
		p := AlignSheet(a, a.Sheets[0], b, b.Sheets[0], Options{Moves: true})
		if len(p.Blocks) != 1 {
			t.Fatalf("beside in column %d: blocks %+v", col, p.Blocks)
		}
		bl := p.Blocks[0]
		if want := []Pair{{1, 1}, {2, 2}, {3, 3}}; !reflect.DeepEqual(bl.Cols, want) {
			t.Errorf("beside in column %d: block's columns %v", col, bl.Cols)
		}
		if want := []Moved{{From: 4, To: 2, Rows: 1}, {From: 6, To: 3, Rows: 1}, {From: 2, To: 4, Rows: 1}, {From: 3, To: 6, Rows: 1}}; !reflect.DeepEqual(bl.Moved, want) {
			t.Errorf("beside in column %d: moved %+v", col, bl.Moved)
		}
		for _, r := range []int{2, 4, 6} {
			if got := pairOf(bl.Rows, r); got != map[int]int{2: 4, 4: 2, 6: 3}[r] {
				t.Errorf("beside in column %d: old row %d in the block is %d", col, r, got)
			}
		}
		// The rest of the sheet in place.
		for r := 1; r <= 7; r++ {
			if got := pairOf(p.Rows, r); got != r {
				t.Errorf("beside in column %d: old row %d is %d", col, r, got)
			}
		}
		if len(p.Moved) != 0 {
			t.Errorf("beside in column %d: whole rows moved %+v", col, p.Moved)
		}
	}
}

// Whole rows sorted, whatever empty columns lie between, are whole rows
// moved; a changed column of repeated values (any value changed pairs with
// another row's) moved nothing.
func TestBlocksNot(t *testing.T) {
	before, after := sorted()
	for _, rows := range []*[][]any{&before, &after} {
		for r := range *rows {
			(*rows)[r] = append(append([]any{}, (*rows)[r]...), nil, (*rows)[r][0])
		}
	}
	a, b := book("Data", before), book("Data", after)
	p := AlignSheet(a, a.Sheets[0], b, b.Sheets[0], Options{Moves: true})
	if len(p.Blocks) != 0 || len(p.Moved) == 0 {
		t.Errorf("whole rows sorted: blocks %+v, moved %+v", p.Blocks, p.Moved)
	}
	var old, now [][]any
	for i := 1; i <= 60; i++ {
		v := i % 5
		old = append(old, []any{"R" + strconv.Itoa(i), v, "=B" + strconv.Itoa(i) + "*2"})
		if i%7 == 0 {
			v = (v + 1) % 5
		}
		now = append(now, []any{"R" + strconv.Itoa(i), v, "=B" + strconv.Itoa(i) + "*2"})
	}
	a, b = book("Data", old), book("Data", now)
	p = AlignSheet(a, a.Sheets[0], b, b.Sheets[0], Options{Moves: true})
	if len(p.Blocks) != 0 || len(p.Moved) != 0 {
		t.Errorf("repeated values changed: blocks %+v, moved %+v", p.Blocks, p.Moved)
	}
}

// TestRewrite checks that rows whose formulas read another sheet pair with
// their old selves when the old formulas are keyed as Excel rewrote them for
// a row added there: two rows alike (the same label and formula) read one
// row further, and so would look like their neighbours.
func TestRewrite(t *testing.T) {
	q := func(r1, r2 int) [][]any {
		f := "=SUM(Data!D" + strconv.Itoa(r1) + ":D" + strconv.Itoa(r2) + ")"
		return [][]any{{"Key", "Value"}, {"Lime", "=COUNTIF(Data!B" + strconv.Itoa(r1) + ":B" + strconv.Itoa(r2) + ",\">3\")"}, {"Pear", 76},
			{"Lime", f}, {"Lime", f}, {"R&D", "=COUNTIF(Data!B" + strconv.Itoa(r1) + ":B" + strconv.Itoa(r2) + ",\">3\")"}, {"Fig", f}}
	}
	a, b := book("Q", q(2, 7)), book("Q", q(3, 8))
	inPlace := func(p SheetPair) bool {
		for r := 1; r <= 7; r++ {
			if pairOf(p.Rows, r) != r {
				return false
			}
		}
		return true
	}
	if p := AlignSheet(a, a.Sheets[0], b, b.Sheets[0], Options{Moves: true}); inPlace(p) {
		t.Fatalf("the case no longer mis-pairs without Rewrite: %+v", p.Rows)
	}
	rewrite := func(sheet, f string) string { return formula.InsertLines(f, sheet, "Data", false, 2, 1) }
	if p := AlignSheet(a, a.Sheets[0], b, b.Sheets[0], Options{Moves: true, Rewrite: rewrite}); !inPlace(p) {
		t.Errorf("with Rewrite: %+v", p.Rows)
	}
}

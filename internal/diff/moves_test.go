package diff

import (
	"reflect"
	"testing"
)

// sorted is a sheet of five rows with a formula in each, sorted by Units:
// rows 2 to 6 come back as 4, 6, 2, 5, 3, their formulas following them.
func sorted() (before, after [][]any) {
	before = [][]any{
		{"Region", "Units", "Double"},
		{"North", 5, "=B2*2"},
		{"South", 9, "=B3*2"},
		{"East", 1, "=B4*2"},
		{"West", 7, "=B5*2"},
		{"Mid", 3, "=B6*2"},
		{"Total", "", "=SUM(C2:C6)"},
	}
	after = [][]any{
		{"Region", "Units", "Double"},
		{"East", 1, "=B2*2"},
		{"Mid", 3, "=B3*2"},
		{"North", 5, "=B4*2"},
		{"West", 7, "=B5*2"},
		{"South", 9, "=B6*2"},
		{"Total", "", "=SUM(C2:C6)"},
	}
	return before, after
}

func TestMovesOff(t *testing.T) {
	before, after := sorted()
	r := Workbooks(book("Data", before), book("Data", after), Options{})
	if len(r.Sheets) != 1 || r.Sheets[0].RowsMoved != nil {
		t.Fatalf("moves off: %+v", r.Sheets)
	}
	if len(r.Sheets[0].Cells) == 0 && len(r.Sheets[0].RowsAdded) == 0 {
		t.Errorf("a sort without move detection shows as changes: %+v", r.Sheets[0])
	}
}

func TestMovesSort(t *testing.T) {
	before, after := sorted()
	r := Workbooks(book("Data", before), book("Data", after), Options{Moves: true})
	if len(r.Sheets) != 1 {
		t.Fatalf("sheets: %+v", r.Sheets)
	}
	d := r.Sheets[0]
	if len(d.Cells)+len(d.RowsAdded)+len(d.RowsRemoved) != 0 {
		t.Errorf("a sort shows as changes: %+v", d)
	}
	// Every row whose number changed: West stays in row 5.
	n := movedCount(d.RowsMoved)
	if want := []Moved{{From: 4, To: 2, Rows: 1}, {From: 6, To: 3, Rows: 1}, {From: 2, To: 4, Rows: 1}, {From: 3, To: 6, Rows: 1}}; !reflect.DeepEqual(d.RowsMoved, want) {
		t.Errorf("moved rows %+v", d.RowsMoved)
	}
	if got := d.Summary(); got != plural(n, "row")+" moved" {
		t.Errorf("summary %q", got)
	}

	al := Align(book("Data", before), book("Data", after), Options{Moves: true})
	if len(al.Sheets) != 1 {
		t.Fatalf("align: %+v", al.Sheets)
	}
	// Each old row is paired with where it went.
	where := map[int]int{}
	for _, p := range al.Sheets[0].Rows {
		where[p.Old] = p.New
	}
	want := map[int]int{1: 1, 2: 4, 3: 6, 4: 2, 5: 5, 6: 3, 7: 7}
	if !reflect.DeepEqual(where, want) {
		t.Errorf("rows %v, want %v", where, want)
	}
}

// A changed row is still a change, and blank rows never move.
func TestMovesNotChanges(t *testing.T) {
	before := [][]any{{"Name", "N"}, {"a", 1}, {}, {"b", 2}, {"c", 3}}
	after := [][]any{{"Name", "N"}, {"c", 3}, {}, {"b", 20}, {"a", 1}}
	al := Align(book("S", before), book("S", after), Options{Moves: true})
	p := al.Sheets[0]
	var moved []Moved
	moved = append(moved, p.Moved...)
	for _, m := range moved {
		if m.From == 3 || m.To == 3 {
			t.Errorf("a blank row moved: %+v", moved)
		}
		if m.From == 4 {
			t.Errorf("a changed row moved: %+v", moved)
		}
	}
	if movedCount(moved) == 0 {
		t.Errorf("no rows moved: %+v", p)
	}
}

func TestMovedRuns(t *testing.T) {
	got := movedRuns([][2]int{{5, 1}, {6, 2}, {7, 3}, {1, 7}})
	want := []Moved{{From: 6, To: 2, Rows: 3}, {From: 2, To: 8, Rows: 1}}
	if !reflect.DeepEqual(got, want) {
		t.Errorf("runs %+v, want %+v", got, want)
	}
}

func TestAlignSheets(t *testing.T) {
	a := book("Keep", [][]any{{"x", 1}}, "Gone", [][]any{{"y", 2}})
	b := book("Keep", [][]any{{"x", 1}, {"z", 3}}, "New", [][]any{{"q", 9}, {"r", 8}})
	al := Align(a, b, Options{Moves: true})
	if len(al.Sheets) != 3 {
		t.Fatalf("sheets %+v", al.Sheets)
	}
	k, n, g := al.Sheets[0], al.Sheets[1], al.Sheets[2]
	if k.Old == nil || k.New == nil || k.New.Name != "Keep" {
		t.Errorf("kept %+v", k)
	}
	if !reflect.DeepEqual(k.Rows, []Pair{{1, 1}, {0, 2}}) || !reflect.DeepEqual(k.Cols, []Pair{{1, 1}, {2, 2}}) {
		t.Errorf("kept rows %v cols %v", k.Rows, k.Cols)
	}
	if n.Old != nil || n.New.Name != "New" || g.New != nil || g.Old.Name != "Gone" {
		t.Errorf("added %+v removed %+v", n, g)
	}
}

// Rows that repeat, a row inserted near the top and every tenth row
// changed: no row moved.
func TestMovesRepeatedRows(t *testing.T) {
	var before, after [][]any
	before = append(before, []any{"Region", "Units", "Double"})
	after = append(after, []any{"Region", "Units", "Double"})
	for i := 1; i <= 400; i++ {
		if i == 3 {
			after = append(after, []any{"New", 99, "=B4*2"})
		}
		u := i % 5
		before = append(before, []any{"R" + string(rune('a'+i%3)), u, ""})
		if i%10 == 0 {
			u += 7
		}
		after = append(after, []any{"R" + string(rune('a'+i%3)), u, ""})
	}
	al := Align(book("S", before), book("S", after), Options{Moves: true})
	if m := al.Sheets[0].Moved; len(m) != 0 {
		t.Errorf("rows moved: %+v", m)
	}
	r := Workbooks(book("S", before), book("S", after), Options{Moves: true})
	d := r.Sheets[0]
	if len(d.RowsAdded) != 1 || d.RowsAdded[0].From != 4 || len(d.Cells) != 40 {
		t.Errorf("added %+v, %d cells changed", d.RowsAdded, len(d.Cells))
	}
}

// A sort below a row inserted: the rows the sort renumbered moved, not
// those the insertion shifted.
func TestMovesSortShifted(t *testing.T) {
	before, _ := sorted()
	shifted := [][]any{
		{"Region", "Units", "Double"},
		{"Inserted", 0, ""},
		{"East", 1, "=B3*2"},
		{"Mid", 3, "=B4*2"},
		{"North", 5, "=B5*2"},
		{"West", 7, "=B6*2"},
		{"South", 9, "=B7*2"},
		{"Total", "", "=SUM(C3:C7)"},
	}
	al := Align(book("Data", before), book("Data", shifted), Options{Moves: true})
	var n int
	for _, m := range al.Sheets[0].Moved {
		n += m.Rows
		if m.To == 8 {
			t.Errorf("the total moved: %+v", al.Sheets[0].Moved)
		}
	}
	if n != 4 {
		t.Errorf("moved %+v", al.Sheets[0].Moved)
	}
}

// unitsTable is a table of regions in the order given, with Units as
// against_before.xlsx's Data has them and titles above it.
func unitsTable(titles int, order []string) [][]any {
	units := map[string]int{"North": 5, "South": 7, "East": 3, "West": 9, "Mid": 4, "Extra": 6}
	var rows [][]any
	for i := 0; i < titles; i++ {
		rows = append(rows, []any{"Title"})
	}
	rows = append(rows, []any{"Region", "Units"})
	for _, n := range order {
		rows = append(rows, []any{n, units[n]})
	}
	return append(rows, []any{"Total", "=SUM(B1:B2)"})
}

// A sort that comes with rows inserted or removed above or inside it moves
// every row it renumbered, whatever number each has now: a row's number
// says nothing once rows were added or removed, its rank among the others
// does.
func TestMovesSortWithRowEdits(t *testing.T) {
	was := []string{"North", "South", "East", "West", "Mid"}
	sortedNames := []string{"East", "Mid", "North", "South", "West"}
	for _, c := range []struct {
		name   string
		after  [][]any
		before [][]any
		moved  map[int]int // old row → new row
	}{
		{"two rows above", unitsTable(2, sortedNames), unitsTable(0, was),
			map[int]int{2: 6, 3: 7, 4: 4, 5: 8, 6: 5}},
		{"one row above", unitsTable(1, sortedNames), unitsTable(0, was),
			map[int]int{2: 5, 3: 6, 4: 3, 5: 7, 6: 4}},
		// South deleted, then sorted: West keeps number 5 but not its rank.
		{"a row deleted inside", unitsTable(0, []string{"East", "Mid", "North", "West"}), unitsTable(0, was),
			map[int]int{2: 4, 4: 2, 5: 5, 6: 3}},
		// A row added inside the sort, wherever it landed.
		{"a row added inside", unitsTable(0, []string{"East", "Mid", "North", "Extra", "South", "West"}), unitsTable(0, was),
			map[int]int{2: 4, 3: 6, 4: 2, 5: 7, 6: 3}},
	} {
		al := Align(book("Data", c.before), book("Data", c.after), Options{Moves: true})
		p := al.Sheets[0]
		got := map[int]int{}
		for _, m := range p.Moved {
			for k := 0; k < m.Rows; k++ {
				got[m.From+k] = m.To + k
			}
		}
		if !reflect.DeepEqual(got, c.moved) {
			t.Errorf("%s: moved %v (runs %+v), want %v", c.name, got, p.Moved, c.moved)
		}
	}
}

// Rows that kept their rank are not moved, though a sort moved the rows
// around them, nor are rows inserted above one that stayed.
func TestMovesKeptRank(t *testing.T) {
	before, after := sorted() // West is row 5 in both
	al := Align(book("Data", before), book("Data", after), Options{Moves: true})
	for _, m := range al.Sheets[0].Moved {
		if m.From == 5 || m.To == 5 {
			t.Errorf("West kept its place: %+v", al.Sheets[0].Moved)
		}
	}
	// With a row inserted above, West is row 6 now: it kept its rank.
	shifted := append([][]any{after[0], {"Inserted", 0, ""}}, after[1:]...)
	al = Align(book("Data", before), book("Data", shifted), Options{Moves: true})
	for _, m := range al.Sheets[0].Moved {
		if m.From == 5 || m.To == 6 {
			t.Errorf("West kept its rank: %+v", al.Sheets[0].Moved)
		}
	}
}

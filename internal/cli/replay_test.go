package cli

import (
	"fmt"
	"math/rand"
	"strings"
	"testing"

	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
)

// naiveShift is what shiftLines does, run by run: the runs removed bottom up,
// then those added top down.
func naiveShift(f, home, sheet string, cols bool, removed, added [][2]int) string {
	for i := len(removed) - 1; i >= 0; i-- {
		f = formula.DeleteLines(f, home, sheet, cols, removed[i][0], removed[i][1])
	}
	for _, a := range added {
		f = formula.InsertLines(f, home, sheet, cols, a[0], a[1]-a[0]+1)
	}
	return f
}

// randomRuns gives runs that do not touch, in increasing order, within
// 1..n.
func randomRuns(rng *rand.Rand, n int) [][2]int {
	var out [][2]int
	for at := 1 + rng.Intn(4); at <= n; {
		size := 1 + rng.Intn(3)
		if rng.Intn(3) == 0 {
			out = append(out, [2]int{at, min(at+size-1, n)})
		}
		at += size + 1 + rng.Intn(4)
	}
	return out
}

// TestShiftLines checks the replay of many runs of lines against replaying
// them one by one: references into the runs, ranges that start, end or lie
// in them, on the formula's sheet and on others, whole rows and columns.
func TestShiftLines(t *testing.T) {
	rng := rand.New(rand.NewSource(7))
	sheets := []string{"", "Data!", "'My Data'!", "Other!"}
	ref := func() string {
		r1, r2 := 1+rng.Intn(40), 1+rng.Intn(40)
		c1, c2 := 1+rng.Intn(12), 1+rng.Intn(12)
		q := sheets[rng.Intn(len(sheets))]
		switch rng.Intn(6) {
		case 0:
			return q + grid.CellName(r1, c1)
		case 1:
			return q + "$" + grid.ColName(c1) + "$" + fmt.Sprint(r1)
		case 2:
			return q + grid.ColName(c1) + ":" + grid.ColName(c2)
		case 3:
			return q + fmt.Sprint(r1) + ":" + fmt.Sprint(r2)
		}
		return q + grid.CellName(min(r1, r2), min(c1, c2)) + ":" + grid.CellName(max(r1, r2), max(c1, c2))
	}
	for i := 0; i < 3000; i++ {
		var parts []string
		for k := 0; k < 1+rng.Intn(4); k++ {
			parts = append(parts, ref())
		}
		f := "SUM(" + strings.Join(parts, ",") + ")"
		home := []string{"Data", "My Data", "Summary", ""}[rng.Intn(4)]
		sheet := []string{"Data", "My Data"}[rng.Intn(2)]
		cols := rng.Intn(3) == 0
		removed, added := randomRuns(rng, 40), randomRuns(rng, 50)
		if rng.Intn(4) == 0 {
			removed = nil
		}
		if rng.Intn(4) == 0 {
			added = nil
		}
		got := shiftLines(f, home, sheet, cols, removed, added)
		want := naiveShift(f, home, sheet, cols, removed, added)
		if got != want {
			t.Fatalf("%s on %q, lines of %s (cols %v) removed %v, added %v:\n got %s\nwant %s", f, home, sheet, cols, removed, added, got, want)
		}
	}
}

// Thousands of runs replay in time on thousands of formulas: a call for
// each line a reference ends on, not for each run.
func TestShiftLinesMany(t *testing.T) {
	var removed [][2]int
	for r := 10; r <= 150000; r += 10 {
		removed = append(removed, [2]int{r, r})
	}
	for row := 2; row <= 20000; row++ {
		f := fmt.Sprintf("B%d/$B$150002", row)
		got := shiftLines(f, "Data", "Data", false, removed, nil)
		if want := fmt.Sprintf("B%d/$B$135002", row-row/10); row%10 != 0 && got != want {
			t.Fatalf("%s: %s, want %s", f, got, want)
		}
	}
}

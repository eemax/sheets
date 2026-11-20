package cli

import (
	"sort"
	"strings"

	"github.com/eemax/sheets/internal/formula"
)

// Replaying lines added and removed on a formula, as Excel rewrites it: what
// formula.DeleteLines and formula.InsertLines do run by run, done with a call
// for each line a reference of the formula ends on rather than for each run.
// Deleting every tenth row of a table is hundreds of runs, and every formula
// reading below them (a share of the total) pays for each.

// shiftLines rewrites f, a formula on sheet home ("" for a name's
// reference), for the lines removed from sheet (runs numbered as before) and
// then those added (runs numbered as after), rows or, with cols set,
// columns. The runs come in increasing order and do not touch.
func shiftLines(f, home, sheet string, cols bool, removed, added [][2]int) string {
	if len(removed) == 0 && len(added) == 0 || !mentions(f, home, sheet) {
		return f
	}
	if len(removed) > 0 {
		f = deleteRuns(f, home, sheet, cols, removed)
	}
	if len(added) > 0 {
		f = insertRuns(f, home, sheet, cols, added)
	}
	return f
}

// mentions tells whether a formula may refer to a sheet: its own, or one it
// names.
func mentions(f, home, sheet string) bool {
	if strings.EqualFold(home, sheet) {
		return true
	}
	lf, ls := strings.ToLower(f), strings.ToLower(sheet)
	return strings.Contains(lf, ls) || strings.Contains(lf, strings.ReplaceAll(ls, "'", "''"))
}

// ends are the lines the references of f on sheet start and end at, in
// increasing order, once each.
func ends(f, home, sheet string, cols bool) []int {
	var out []int
	for _, ref := range formula.Refs(f) {
		s := ref.Sheet
		if s == "" {
			s = home
		}
		if !strings.EqualFold(s, sheet) {
			continue
		}
		if cols {
			out = append(out, ref.Range.C1, ref.Range.C2)
		} else {
			out = append(out, ref.Range.R1, ref.Range.R2)
		}
	}
	sort.Ints(out)
	k := 0
	for i, n := range out {
		if i == 0 || n != out[k-1] {
			out[k] = n
			k++
		}
	}
	return out[:k]
}

// runSizes gives, for each run, the lines in the runs before it (and, last,
// in all of them).
func runSizes(runs [][2]int) []int {
	pre := make([]int, len(runs)+1)
	for i, r := range runs {
		pre[i+1] = pre[i] + r[1] - r[0] + 1
	}
	return pre
}

// deleteRuns deletes the runs from f's references. A run holding a line a
// reference ends on is deleted as it is (the reference becomes #REF!, or a
// range shrinks); the runs between two such lines only move what is past
// them, so they are deleted as one run of their total size, right below the
// next. Bottom up, each keeps the numbers as before.
func deleteRuns(f, home, sheet string, cols bool, runs [][2]int) string {
	pts := ends(f, home, sheet, cols)
	if len(pts) == 0 {
		return f
	}
	pre := runSizes(runs)
	var ops [][2]int
	done := 0 // the runs below the last line looked at
	for _, p := range pts {
		k := sort.Search(len(runs), func(i int) bool { return runs[i][1] >= p })
		hit := k < len(runs) && runs[k][0] <= p
		if s := pre[k] - pre[done]; s > 0 {
			top := p
			if hit {
				top = runs[k][0]
			}
			ops = append(ops, [2]int{top - s, top - 1})
		}
		if hit {
			if n := len(ops); n == 0 || ops[n-1] != runs[k] {
				ops = append(ops, runs[k])
			}
			done = k + 1
		} else {
			done = k
		}
	}
	for i := len(ops) - 1; i >= 0; i-- {
		f = formula.DeleteLines(f, home, sheet, cols, ops[i][0], ops[i][1])
	}
	return f
}

// insertRuns inserts the runs (numbered as after) into f's references. A
// line kept is the u-th of those not added, so its place after is u plus
// the lines added before that; each line a reference ends on moves by its
// share, bottom up, with what is past it.
func insertRuns(f, home, sheet string, cols bool, runs [][2]int) string {
	pts := ends(f, home, sheet, cols)
	pre := runSizes(runs)
	moved := 0 // what the lines looked at have moved by
	for _, u := range pts {
		// The runs with fewer lines kept before them than u come before it.
		k := sort.Search(len(runs), func(i int) bool { return runs[i][0]-1-pre[i] >= u })
		if s := pre[k]; s > moved {
			f = formula.InsertLines(f, home, sheet, cols, u+moved, s-moved)
			moved = s
		}
	}
	return f
}

// outside keeps the runs that do not meet a span (zero for none).
func outside(runs [][2]int, span [2]int) [][2]int {
	if span[0] == 0 {
		return runs
	}
	var out [][2]int
	for _, run := range runs {
		if !within(run, span) {
			out = append(out, run)
		}
	}
	return out
}

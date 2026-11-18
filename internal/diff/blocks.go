package diff

import "github.com/eemax/sheets/internal/grid"

// Rows that moved on their own in some columns: a sort of a range with other
// cells beside it (a note, a lookup list, a formula reading the table), which
// stay where they are. Rows compared whole are then neither the same nor in
// order, so findMoves sees no row moved, and the rows of the range read as
// added, removed or changed. Columns apart from each other by a column empty
// in both versions are a block; each block's rows are aligned on their own,
// and a block whose rows moved is paired as its alignment has it, with the
// rest of the sheet paired in place around it, when that leaves fewer cells
// changed than pairing whole rows.

// blockAlign is a block of columns whose rows moved on their own: its
// columns (1-based pairs) and how its rows pair (0-based, as sheetAlign).
type blockAlign struct {
	cols          [][2]int
	rows          []step
	moved, placed [][2]int
}

// blocksOf splits the column pairs into blocks: runs of columns with
// something in them, apart where a column is empty in both versions.
func blocksOf(A, B side, pairs [][2]int) [][][2]int {
	used := func(x side) map[int]bool {
		m := map[int]bool{}
		x.s.Each(x.s.Used(), func(_, col int, _ *grid.Cell) { m[col] = true })
		return m
	}
	ua, ub := used(A), used(B)
	var out [][][2]int
	var cur [][2]int
	for _, p := range pairs {
		if !ua[p[0]] && !ub[p[1]] {
			if len(cur) > 0 {
				out = append(out, cur)
				cur = nil
			}
			continue
		}
		cur = append(cur, p)
	}
	if len(cur) > 0 {
		out = append(out, cur)
	}
	return out
}

// rowSigs describes each row of a side over some columns, for alignment,
// and tells which rows are blank there.
func rowSigs(x side, rows int, cols [][2]int, pick int) ([]uint64, []bool) {
	out := make([]uint64, rows)
	blank := make([]bool, rows)
	parts := make([]string, len(cols))
	for r := 1; r <= rows; r++ {
		empty := true
		for k, p := range cols {
			parts[k] = x.key(r, p[pick])
			empty = empty && parts[k] == ""
		}
		out[r-1], blank[r-1] = hash(parts...), empty
	}
	return out, blank
}

// similarOver tells whether two rows are mostly alike over some columns (a
// changed row rather than one removed and one added), as alignSheet does.
func similarOver(A, B side, cols [][2]int, lostRefs bool) func(i, j int) bool {
	return func(i, j int) bool {
		same, used := 0, 0
		for _, p := range cols {
			ka, kb := A.key(i+1, p[0]), B.key(j+1, p[1])
			if ka == "" && kb == "" {
				continue
			}
			used++
			if ka == kb || alike(A, i+1, p[0], B, j+1, p[1]) || lostRefs && lostRef(ka, kb) {
				same++
			}
		}
		return used > 0 && same*2 >= used
	}
}

// alignBlocks pairs, in sheets with several blocks of columns, the rows of
// each block that moved on its own, and the rest of the sheet in place
// around them, when that leaves fewer cells changed than full gives (whole
// rows paired). The blocks tried are the runs of columns, between empty
// ones, that are changed or not, both in place and with whole rows paired:
// a sort of some columns of a table changes those, in place, and not the
// others.
func alignBlocks(A, B side, opts Options, full sheetAlign) sheetAlign {
	pairs := matched(full.cols)
	fullRows := append(append(stepPairs(full.rows), full.placed...), full.moved...)
	ua, ub := A.s.Used(), B.s.Used()
	ra, rb := max(ua.R2, 0), max(ub.R2, 0)
	inPlace := make([][2]int, max(ra, rb))
	asIs := len(fullRows) == len(inPlace) // whole rows paired in place
	for r := range inPlace {
		inPlace[r] = [2]int{r, r}
		if r >= ra {
			inPlace[r][0] = -1
		}
		if r >= rb {
			inPlace[r][1] = -1
		}
		asIs = asIs && fullRows[r] == inPlace[r]
	}
	// runs of columns, by block apart from the rest, and whether each is
	// changed
	var parts [][][][2]int
	var changedRun [][]bool
	n, anyChanged := 0, false
	for _, g := range blocksOf(A, B, pairs) {
		var runs [][][2]int
		var flags []bool
		for k, p := range g {
			changed := differs(A, B, fullRows, p) && (asIs || differs(A, B, inPlace, p))
			if k == 0 || changed != flags[len(flags)-1] {
				runs, flags = append(runs, nil), append(flags, changed)
			}
			runs[len(runs)-1] = append(runs[len(runs)-1], p)
			anyChanged = anyChanged || changed
		}
		parts, changedRun = append(parts, runs), append(changedRun, flags)
		n += len(runs)
	}
	if n < 2 || !anyChanged {
		return full
	}
	var blocks []blockAlign
	var groups [][][2]int // all runs, those of moving blocks emptied
	moving := map[int]bool{}
	for i, runs := range parts {
		for k := 0; k < len(runs); k++ {
			cols := runs[k]
			if len(cols) == 0 || !changedRun[i][k] {
				continue // as it was, in place or with whole rows paired
			}
			sa, blankA := rowSigs(A, ra, cols, 0)
			sb, blankB := rowSigs(B, rb, cols, 1)
			inOrder, moved, placed := findMoves(sa, sb, fillGaps(sa, sb, align(sa, sb)), blankA, blankB)
			if !distinct(sa, sb, moved) {
				continue
			}
			skipA, skipB := map[int]bool{}, map[int]bool{}
			for _, m := range append(append([][2]int{}, moved...), placed...) {
				skipA[m[0]], skipB[m[1]] = true, true
			}
			rows := pairUpSkipping(ra, rb, inOrder, similarOver(A, B, cols, opts.LostRefs), skipA, skipB)
			// The columns beside it that its rows explain too (a formula
			// column reading the row, the same whether it moved or not) are
			// part of it.
			at := append(append(stepPairs(rows), placed...), moved...)
			fits := func(p [2]int) bool { return !differs(A, B, at, p) }
			if k > 0 {
				left := runs[k-1]
				for len(left) > 0 && fits(left[len(left)-1]) {
					cols = append([][2]int{left[len(left)-1]}, cols...)
					left = left[:len(left)-1]
				}
				runs[k-1] = left
			}
			if k+1 < len(runs) {
				right := runs[k+1]
				for len(right) > 0 && fits(right[0]) {
					cols = append(cols, right[0])
					right = right[1:]
				}
				runs[k+1] = right
			}
			runs[k] = cols
			blocks = append(blocks, blockAlign{cols: cols, rows: rows, moved: moved, placed: placed})
			moving[len(groups)+k] = true
		}
		groups = append(groups, runs...)
	}
	if len(blocks) == 0 {
		return full
	}
	// The rest of the sheet, in place: the columns of the blocks that stayed,
	// and each block that moved as one cell, filled or blank (its rows are
	// where they were, whatever they hold now).
	var still [][2]int
	var movingCols [][][2]int
	for g, cols := range groups {
		if moving[g] {
			movingCols = append(movingCols, cols)
		} else {
			still = append(still, cols...)
		}
	}
	if len(still) == 0 && len(movingCols) == 1 {
		return full // every column moved, together: whole rows
	}
	filled := func(x side, row int, cols [][2]int, pick int) bool {
		for _, p := range cols {
			if x.key(row, p[pick]) != "" {
				return true
			}
		}
		return false
	}
	sig := func(x side, rows, pick int) []uint64 {
		out := make([]uint64, rows)
		parts := make([]string, len(still)+len(movingCols))
		for r := 1; r <= rows; r++ {
			for k, p := range still {
				parts[k] = x.key(r, p[pick])
			}
			for k, cols := range movingCols {
				parts[len(still)+k] = map[bool]string{true: "filled"}[filled(x, r, cols, pick)]
			}
			out[r-1] = hash(parts...)
		}
		return out
	}
	sa, sb := sig(A, ra, 0), sig(B, rb, 1)
	similar := func(i, j int) bool {
		same, used := 0, 0
		for _, p := range still {
			ka, kb := A.key(i+1, p[0]), B.key(j+1, p[1])
			if ka == "" && kb == "" {
				continue
			}
			used++
			if ka == kb || alike(A, i+1, p[0], B, j+1, p[1]) || opts.LostRefs && lostRef(ka, kb) {
				same++
			}
		}
		for _, cols := range movingCols {
			fa, fb := filled(A, i+1, cols, 0), filled(B, j+1, cols, 1)
			if !fa && !fb {
				continue
			}
			used++
			if fa == fb {
				same++
			}
		}
		return used > 0 && same*2 >= used
	}
	skeleton := pairUp(ra, rb, fillGaps(sa, sb, align(sa, sb)), similar)
	// Fewer cells changed, or whole rows it is.
	cost := cellsChanged(A, B, stepPairs(skeleton), still)
	for _, bl := range blocks {
		cost += cellsChanged(A, B, append(append(stepPairs(bl.rows), bl.placed...), bl.moved...), bl.cols)
	}
	if cost >= cellsChanged(A, B, append(append(stepPairs(full.rows), full.placed...), full.moved...), pairs) {
		return full
	}
	return sheetAlign{cols: full.cols, rows: skeleton, blocks: blocks}
}

// distinct tells whether a row moved holds what only one row holds in each
// version: a sort, not a value changed to one that other rows hold too (in
// a column of repeated values, any change pairs with some other row).
func distinct(a, b []uint64, moved [][2]int) bool {
	if len(moved) == 0 {
		return false
	}
	count := map[uint64][2]int{}
	for _, h := range a {
		c := count[h]
		c[0]++
		count[h] = c
	}
	for _, h := range b {
		c := count[h]
		c[1]++
		count[h] = c
	}
	for _, m := range moved {
		if c := count[a[m[0]]]; c == [2]int{1, 1} {
			return true
		}
	}
	return false
}

// stepPairs gives an alignment's steps as pairs (0-based, -1 for a line
// only one side has).
func stepPairs(steps []step) [][2]int {
	out := make([][2]int, len(steps))
	for i, st := range steps {
		out[i] = [2]int{st.a, st.b}
	}
	return out
}

// differs tells whether a pairing of rows leaves a cell of a column changed.
func differs(A, B side, rows [][2]int, col [2]int) bool {
	for _, q := range rows {
		ka, kb := "", ""
		if q[0] >= 0 {
			ka = A.key(q[0]+1, col[0])
		}
		if q[1] >= 0 {
			kb = B.key(q[1]+1, col[1])
		}
		if ka != kb {
			return true
		}
	}
	return false
}

// cellsChanged counts the cells of some columns a pairing of rows leaves
// changed: those unlike their pair's, and those with something in them in
// rows only one side has.
func cellsChanged(A, B side, rows [][2]int, cols [][2]int) int {
	n := 0
	for _, q := range rows {
		for _, p := range cols {
			ka, kb := "", ""
			if q[0] >= 0 {
				ka = A.key(q[0]+1, p[0])
			}
			if q[1] >= 0 {
				kb = B.key(q[1]+1, p[1])
			}
			if ka != kb {
				n++
			}
		}
	}
	return n
}

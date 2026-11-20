package diff

import "github.com/eemax/sheets/internal/grid"

// Rows that moved on their own in some columns: a sort of a table with other
// cells past an empty column beside it (a note, a lookup list, a formula
// reading the table), which stay where they are. Rows compared whole are then
// neither the same nor in order, so findMoves sees no row moved, and the rows
// of the table read as added, removed or changed. Columns apart from each
// other by a column empty in both versions are a block; each block's rows are
// aligned on their own, every column of it together, and a block whose rows
// moved is paired as its alignment has it, with the rest of the sheet paired
// in place around it, when that leaves fewer cells changed than pairing whole
// rows. Columns of a block that moved without the rest of it are not rows
// moved: what they hold changed.

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
// rows paired). A block is tried whole, every column of it between empty
// ones: a row of it moved only when all its columns moved together, labels
// included. A sort of some columns of a table, or two values swapped in one
// column of it, splits its rows (North with East's numbers), which is cells
// changed, not rows moved (Excel's own sort warns of "data next to your
// selection"): the table's rows pair as the columns that stayed have them,
// so the cells moved are changed where they are.
func alignBlocks(A, B side, opts Options, full sheetAlign) sheetAlign {
	pairs := matched(full.cols)
	groups := blocksOf(A, B, pairs)
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
	// changed tells whether a column is changed, both in place and with
	// whole rows paired: a sort of some columns of a table changes those,
	// in place, and not the others.
	changed := func(p [2]int) bool {
		return differs(A, B, fullRows, p) && (asIs || differs(A, B, inPlace, p))
	}
	// movesIn aligns the rows of some columns on their own, and tells
	// whether rows moved there, one of a kind.
	movesIn := func(cols [][2]int) (blockAlign, bool) {
		sa, blankA := rowSigs(A, ra, cols, 0)
		sb, blankB := rowSigs(B, rb, cols, 1)
		inOrder, moved, placed := findMoves(sa, sb, fillGaps(sa, sb, align(sa, sb)), blankA, blankB)
		if !distinct(sa, sb, moved) {
			return blockAlign{}, false
		}
		skipA, skipB := map[int]bool{}, map[int]bool{}
		for _, m := range append(append([][2]int{}, moved...), placed...) {
			skipA[m[0]], skipB[m[1]] = true, true
		}
		rows := pairUpSkipping(ra, rb, inOrder, similarOver(A, B, cols, opts.LostRefs), skipA, skipB)
		return blockAlign{cols: cols, rows: rows, moved: moved, placed: placed}, true
	}
	var blocks []blockAlign
	var still [][2]int        // the columns the rest of the sheet pairs by
	var movingCols [][][2]int // those of each block that moved, and of each split
	split := false            // a table's rows split: some of its columns sorted
	for _, cols := range groups {
		var runs [][][2]int
		var flags []bool
		for k, p := range cols {
			c := changed(p)
			if k == 0 || c != flags[len(flags)-1] {
				runs, flags = append(runs, nil), append(flags, c)
			}
			runs[len(runs)-1] = append(runs[len(runs)-1], p)
		}
		if len(runs) == 1 && !flags[0] {
			still = append(still, cols...) // as it was
			continue
		}
		switch {
		case len(groups) > 1:
			if bl, ok := movesIn(cols); ok {
				blocks = append(blocks, bl)
				movingCols = append(movingCols, cols)
				continue
			}
		case len(full.moved) > 0:
			return full // the one block's whole rows moved
		}
		// Rows that moved in some of its columns, without the rest: those
		// columns changed, in the rows the rest pairs.
		var stay [][2]int
		var sorted [][][2]int
		for k, run := range runs {
			if flags[k] && len(run) < len(cols) {
				if _, ok := movesIn(run); ok {
					sorted = append(sorted, run)
					continue
				}
			}
			stay = append(stay, run...)
		}
		if len(sorted) == 0 || len(stay) == 0 {
			still = append(still, cols...)
			continue
		}
		still = append(still, stay...)
		movingCols = append(movingCols, sorted...)
		split = true
	}
	if len(blocks) == 0 && !split {
		return full
	}
	// The rest of the sheet, in place: the columns of the blocks that stayed,
	// and each block that moved (and each run of columns sorted without the
	// rest of its table) as one cell, filled or blank (its rows are where they
	// were, whatever they hold now).
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
	if split {
		// No rows moved there, whatever pairing whole rows makes of it.
		return sheetAlign{cols: full.cols, rows: skeleton, blocks: blocks}
	}
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

package diff

import "sort"

// Sequence alignment by patience diff (as git uses): lines unique in both
// sequences anchor the alignment, the longest increasing run of anchors is
// kept, and the gaps between anchors are aligned recursively. What stays
// unmatched is paired by similarity in pairUp.

// align returns matched index pairs (i into a, j into b), increasing in
// both.
func align(a, b []uint64) [][2]int {
	var out [][2]int
	alignRange(a, b, 0, len(a), 0, len(b), &out)
	return out
}

func alignRange(a, b []uint64, a1, a2, b1, b2 int, out *[][2]int) {
	// Common prefix and suffix.
	for a1 < a2 && b1 < b2 && a[a1] == b[b1] {
		*out = append(*out, [2]int{a1, b1})
		a1++
		b1++
	}
	var tail [][2]int
	for a1 < a2 && b1 < b2 && a[a2-1] == b[b2-1] {
		a2--
		b2--
		tail = append(tail, [2]int{a2, b2})
	}
	defer func() {
		for i := len(tail) - 1; i >= 0; i-- {
			*out = append(*out, tail[i])
		}
	}()
	if a1 == a2 || b1 == b2 {
		return
	}
	anchors := uniqueAnchors(a, b, a1, a2, b1, b2)
	if len(anchors) == 0 {
		lcsSmall(a, b, a1, a2, b1, b2, out)
		return
	}
	pi, pj := a1, b1
	for _, m := range anchors {
		alignRange(a, b, pi, m[0], pj, m[1], out)
		*out = append(*out, m)
		pi, pj = m[0]+1, m[1]+1
	}
	alignRange(a, b, pi, a2, pj, b2, out)
}

// uniqueAnchors finds values occurring exactly once in each range and
// keeps the longest run of them that is increasing in both.
func uniqueAnchors(a, b []uint64, a1, a2, b1, b2 int) [][2]int {
	type seen struct{ na, nb, ia, ib int }
	m := map[uint64]*seen{}
	for i := a1; i < a2; i++ {
		s := m[a[i]]
		if s == nil {
			s = &seen{}
			m[a[i]] = s
		}
		s.na++
		s.ia = i
	}
	for j := b1; j < b2; j++ {
		if s := m[b[j]]; s != nil {
			s.nb++
			s.ib = j
		}
	}
	var pairs [][2]int // in a order
	for i := a1; i < a2; i++ {
		if s := m[a[i]]; s.na == 1 && s.nb == 1 {
			pairs = append(pairs, [2]int{s.ia, s.ib})
		}
	}
	return longestIncreasing(pairs)
}

// longestIncreasing keeps the longest subsequence of pairs whose second
// index increases (patience sorting).
func longestIncreasing(pairs [][2]int) [][2]int {
	if len(pairs) == 0 {
		return nil
	}
	tails := []int{} // index into pairs of the smallest tail of each length
	prev := make([]int, len(pairs))
	for i, p := range pairs {
		lo, hi := 0, len(tails)
		for lo < hi {
			mid := (lo + hi) / 2
			if pairs[tails[mid]][1] < p[1] {
				lo = mid + 1
			} else {
				hi = mid
			}
		}
		if lo > 0 {
			prev[i] = tails[lo-1]
		} else {
			prev[i] = -1
		}
		if lo == len(tails) {
			tails = append(tails, i)
		} else {
			tails[lo] = i
		}
	}
	out := make([][2]int, len(tails))
	for k, i := len(tails)-1, tails[len(tails)-1]; k >= 0; k, i = k-1, prev[i] {
		out[k] = pairs[i]
	}
	return out
}

// lcsSmall aligns a gap without unique anchors by longest common
// subsequence, when it is small enough; larger gaps stay unmatched.
func lcsSmall(a, b []uint64, a1, a2, b1, b2 int, out *[][2]int) {
	n, m := a2-a1, b2-b1
	if n*m > 1_000_000 {
		return
	}
	dp := make([][]int32, n+1)
	for i := range dp {
		dp[i] = make([]int32, m+1)
	}
	for i := n - 1; i >= 0; i-- {
		for j := m - 1; j >= 0; j-- {
			if a[a1+i] == b[b1+j] {
				dp[i][j] = dp[i+1][j+1] + 1
			} else {
				dp[i][j] = max(dp[i+1][j], dp[i][j+1])
			}
		}
	}
	for i, j := 0, 0; i < n && j < m; {
		switch {
		case a[a1+i] == b[b1+j]:
			*out = append(*out, [2]int{a1 + i, b1 + j})
			i++
			j++
		case dp[i+1][j] >= dp[i][j+1]:
			i++
		default:
			j++
		}
	}
}

// step is one entry of an alignment walk.
type step struct {
	a, b int // -1 when the line is only on the other side
}

// pairUp walks an alignment and pairs unmatched lines that are similar
// enough (a changed row rather than a removed and an added one).
func pairUp(na, nb int, matches [][2]int, similar func(i, j int) bool) []step {
	return pairUpSkipping(na, nb, matches, similar, nil, nil)
}

// pairUpSkipping is pairUp leaving out the lines skipA and skipB hold
// (rows paired elsewhere, as moved).
func pairUpSkipping(na, nb int, matches [][2]int, similar func(i, j int) bool, skipA, skipB map[int]bool) []step {
	var out []step
	i, j := 0, 0
	gap := func(i2, j2 int) {
		for i < i2 || j < j2 {
			if i < i2 && skipA[i] {
				i++
				continue
			}
			if j < j2 && skipB[j] {
				j++
				continue
			}
			switch {
			case i < i2 && j < j2 && similar(i, j):
				out = append(out, step{i, j})
				i++
				j++
			case i < i2 && j+1 < j2 && similar(i, j+1):
				out = append(out, step{-1, j}) // a line was added before a changed one
				j++
			case j < j2 && i+1 < i2 && similar(i+1, j):
				out = append(out, step{i, -1}) // a line was removed before a changed one
				i++
			case i < i2 && (i2-i >= j2-j || j == j2):
				out = append(out, step{i, -1})
				i++
			default:
				out = append(out, step{-1, j})
				j++
			}
		}
	}
	for _, m := range matches {
		gap(m[0], m[1])
		out = append(out, step{m[0], m[1]})
		i, j = m[0]+1, m[1]+1
	}
	gap(na, nb)
	return out
}

// findMoves pairs the lines an alignment left unmatched that are the same
// on both sides (a row a sort took elsewhere), blank lines aside. Of all
// the pairs, a row moved when its rank among the paired rows changed: rows
// added or removed shift a row's number, not its rank, so a sort that comes
// with rows inserted or removed above or inside it moves every row it
// renumbered, whatever number each has now, and a row that kept its rank is
// in its place (placed: paired, but not in the others' order). Blank rows
// never move.
func findMoves(a, b []uint64, matches [][2]int, blankA, blankB []bool) (inOrder, moved, placed [][2]int) {
	inA, inB := make([]bool, len(a)), make([]bool, len(b))
	for _, m := range matches {
		inA[m[0]], inB[m[1]] = true, true
	}
	free := map[uint64][]int{}
	for i := range a {
		if !inA[i] && !blankA[i] {
			free[a[i]] = append(free[a[i]], i)
		}
	}
	all := append([][2]int{}, matches...)
	for j := range b {
		if inB[j] || blankB[j] {
			continue
		}
		if l := free[b[j]]; len(l) > 0 {
			all = append(all, [2]int{l[0], j})
			free[b[j]] = l[1:]
		}
	}
	if len(all) == len(matches) {
		return matches, nil, nil
	}
	sort.Slice(all, func(i, j int) bool { return all[i][0] < all[j][0] })
	// A pair's rank in the old version is its place in all; in the new one,
	// its place by new index.
	byNew := make([]int, len(all))
	for k := range byNew {
		byNew[k] = k
	}
	sort.Slice(byNew, func(x, y int) bool { return all[byNew[x]][1] < all[byNew[y]][1] })
	rankNew := make([]int, len(all))
	for r, k := range byNew {
		rankNew[k] = r
	}
	kept := make(map[[2]int]bool, len(matches))
	for _, p := range longestIncreasing(all) {
		kept[p] = true
	}
	for k, p := range all {
		switch {
		case blankA[p[0]] || blankB[p[1]]:
			// A blank row never moves: matched in order it stays, else it is
			// in its place.
			if kept[p] {
				inOrder = append(inOrder, p)
			} else {
				placed = append(placed, p)
			}
		case k != rankNew[k]:
			moved = append(moved, p)
		case kept[p]:
			inOrder = append(inOrder, p)
		default:
			placed = append(placed, p)
		}
	}
	if moved == nil {
		return inOrder, nil, placed
	}
	return inOrder, moved, placed
}

// lookAhead bounds how far fillGaps looks for a line's match.
const lookAhead = 64

// fillGaps matches, in order, the lines an alignment left unmatched
// between its matches that are the same on both sides, looking a few lines
// ahead past lines added, removed or changed. Patience diff anchors only on
// lines unique in both versions, so rows that repeat (and gaps too large to
// align line by line) are left to it; without it, move detection would pair
// repeated rows far apart.
func fillGaps(a, b []uint64, matches [][2]int) [][2]int {
	var out [][2]int
	pi, pj := 0, 0
	gap := func(i2, j2 int) {
		i, j := pi, pj
		for i < i2 && j < j2 {
			if a[i] == b[j] {
				out = append(out, [2]int{i, j})
				i++
				j++
				continue
			}
			found := false
			for k := 1; k <= lookAhead; k++ {
				if j+k < j2 && a[i] == b[j+k] {
					j += k // lines added
					found = true
					break
				}
				if i+k < i2 && a[i+k] == b[j] {
					i += k // lines removed
					found = true
					break
				}
			}
			if !found {
				i++ // a line changed
				j++
			}
		}
	}
	for _, m := range matches {
		gap(m[0], m[1])
		out = append(out, m)
		pi, pj = m[0]+1, m[1]+1
	}
	gap(len(a), len(b))
	return out
}

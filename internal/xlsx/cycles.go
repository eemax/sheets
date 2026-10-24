package xlsx

import (
	"sort"
	"strings"

	"github.com/eemax/sheets/internal/formula"
)

// Circular references. Excel reports them; the sheets calculator instead
// reads the cell being computed as its old stored result, and returns
// numbers that mean nothing. So formulas in a cycle, and formulas reading
// one, get no result from it.
//
// The graph has an edge from a formula to every formula it reads. A range
// can read thousands of formulas (a running total =SUM($E$2:E9) down a
// column of formulas), so a formula reaches a range through a segment tree
// over each column's formulas: O(log n) edges per range, not one per
// formula, and reachability is unchanged.

type cycleGraph struct {
	edges [][]int32 // node → successors; formulas first, then tree nodes
	n     int       // formulas
	fs    []*formulaAt
	id    map[*formulaAt]int
}

// columnTree is a segment tree over the formulas of one column, by row.
type columnTree struct {
	rows  []int // sorted rows of the column's formulas
	ids   []int // formula at each row
	base  int   // node number of the tree's root
	nodes int   // tree nodes, numbered base, base+1, ...
}

// newCycleGraph builds the graph of what formulas read: an edge from a
// formula to each formula it reads (through the column trees for ranges).
func newCycleGraph(fs []*formulaAt, deps func(f *formulaAt) ([]formula.Ref, bool)) *cycleGraph {
	g := &cycleGraph{n: len(fs), fs: fs, id: map[*formulaAt]int{}}
	g.edges = make([][]int32, len(fs))
	at := map[cellKey]int{}
	cols := map[cellKey]*columnTree{}
	for i, f := range fs {
		g.id[f] = i
		at[keyOf(f.sheet, f.row, f.col)] = i
		k := cellKey{strings.ToLower(f.sheet), 0, f.col}
		t := cols[k]
		if t == nil {
			t = &columnTree{}
			cols[k] = t
		}
		t.rows, t.ids = append(t.rows, f.row), append(t.ids, i)
	}
	sheetCols := map[string][]int{} // columns holding formulas, by sheet
	for k, t := range cols {
		sort.Sort(byRow{t})
		t.base = len(g.edges)
		t.nodes = 4 * len(t.rows)
		g.edges = append(g.edges, make([][]int32, t.nodes)...)
		g.build(t, 0, 0, len(t.rows)-1)
		sheetCols[k.sheet] = append(sheetCols[k.sheet], k.col)
	}
	for i, f := range fs {
		refs, ok := deps(f)
		if !ok {
			continue
		}
		for _, r := range refs {
			sheet := strings.ToLower(r.Sheet)
			rg := r.Range
			if rg.R1 == rg.R2 && rg.C1 == rg.C2 {
				if j, ok := at[cellKey{sheet, rg.R1, rg.C1}]; ok {
					g.edges[i] = append(g.edges[i], int32(j))
				}
				continue
			}
			for _, c := range sheetCols[sheet] {
				if c < rg.C1 || c > rg.C2 {
					continue
				}
				t := cols[cellKey{sheet, 0, c}]
				lo := sort.SearchInts(t.rows, rg.R1)
				hi := sort.SearchInts(t.rows, rg.R2+1) - 1
				if lo <= hi {
					g.cover(i, t, 0, 0, len(t.rows)-1, lo, hi)
				}
			}
		}
	}
	return g
}

// cycles finds the formulas in a circular reference, and those reading
// one.
func (g *cycleGraph) cycles() (reaching map[*formulaAt]bool, cycle []*formulaAt) {
	fs := g.fs
	// Tarjan's strongly connected components; each component comes out
	// after every component it reaches, so "reaches a cycle" is known for
	// its successors by then.
	n := len(g.edges)
	index := make([]int32, n)
	low := make([]int32, n)
	comp := make([]int32, n)
	for i := range index {
		index[i], comp[i] = -1, -1
	}
	var stack []int32
	onStack := make([]bool, n)
	var compReaches []bool
	next := int32(0)
	type frame struct {
		v    int32
		edge int
	}
	for root := 0; root < g.n; root++ {
		if index[root] >= 0 {
			continue
		}
		call := []frame{{int32(root), 0}}
		index[root], low[root] = next, next
		next++
		stack = append(stack, int32(root))
		onStack[root] = true
		for len(call) > 0 {
			fr := &call[len(call)-1]
			v := fr.v
			if fr.edge < len(g.edges[v]) {
				w := g.edges[v][fr.edge]
				fr.edge++
				if index[w] < 0 {
					index[w], low[w] = next, next
					next++
					stack = append(stack, w)
					onStack[w] = true
					call = append(call, frame{w, 0})
				} else if onStack[w] && index[w] < low[v] {
					low[v] = index[w]
				}
				continue
			}
			call = call[:len(call)-1]
			if len(call) > 0 {
				if u := call[len(call)-1].v; low[v] < low[u] {
					low[u] = low[v]
				}
			}
			if low[v] != index[v] {
				continue
			}
			// v roots a component: pop it.
			id := int32(len(compReaches))
			var members []int32
			for {
				w := stack[len(stack)-1]
				stack = stack[:len(stack)-1]
				onStack[w] = false
				comp[w] = id
				members = append(members, w)
				if w == v {
					break
				}
			}
			inCycle := len(members) > 1
			if !inCycle {
				for _, w := range g.edges[v] {
					inCycle = inCycle || w == v
				}
			}
			reaches := inCycle
			for _, m := range members {
				for _, w := range g.edges[m] {
					if c := comp[w]; c >= 0 && c != id && compReaches[c] {
						reaches = true
					}
				}
				if inCycle && int(m) < g.n {
					cycle = append(cycle, fs[m])
				}
			}
			compReaches = append(compReaches, reaches)
		}
	}
	reaching = map[*formulaAt]bool{}
	for i, f := range fs {
		if compReaches[comp[i]] {
			reaching[f] = true
		}
	}
	sort.Slice(cycle, func(i, j int) bool {
		a, b := cycle[i], cycle[j]
		if a.sheet != b.sheet {
			return a.sheet < b.sheet
		}
		if a.row != b.row {
			return a.row < b.row
		}
		return a.col < b.col
	})
	return reaching, cycle
}

// build links tree node k (covering formulas lo..hi of t) to its children,
// or at a leaf to the formula.
func (g *cycleGraph) build(t *columnTree, k, lo, hi int) {
	node := t.base + k
	if lo == hi {
		g.edges[node] = append(g.edges[node], int32(t.ids[lo]))
		return
	}
	mid := (lo + hi) / 2
	g.edges[node] = append(g.edges[node], int32(t.base+2*k+1), int32(t.base+2*k+2))
	g.build(t, 2*k+1, lo, mid)
	g.build(t, 2*k+2, mid+1, hi)
}

// cover links formula i to the tree nodes that cover formulas lo..hi.
func (g *cycleGraph) cover(i int, t *columnTree, k, nlo, nhi, lo, hi int) {
	if hi < nlo || nhi < lo {
		return
	}
	if lo <= nlo && nhi <= hi {
		g.edges[i] = append(g.edges[i], int32(t.base+k))
		return
	}
	mid := (nlo + nhi) / 2
	g.cover(i, t, 2*k+1, nlo, mid, lo, hi)
	g.cover(i, t, 2*k+2, mid+1, nhi, lo, hi)
}

type byRow struct{ t *columnTree }

func (s byRow) Len() int           { return len(s.t.rows) }
func (s byRow) Less(i, j int) bool { return s.t.rows[i] < s.t.rows[j] }
func (s byRow) Swap(i, j int) {
	s.t.rows[i], s.t.rows[j] = s.t.rows[j], s.t.rows[i]
	s.t.ids[i], s.t.ids[j] = s.t.ids[j], s.t.ids[i]
}

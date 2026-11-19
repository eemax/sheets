package xlsx

import (
	"sort"
	"strings"

	"github.com/eemax/sheets/internal/engine"
	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
)

// Tracing what reads what in a workbook, forward from cells, for telling
// why a value moved (sheets check --against): what each formula reads is
// read off the text the calculator computes (names resolved, table
// references translated into their cells), as when saving works out what
// an edit reaches. The cells of an array formula's result read its
// formula.

// At is a cell of a workbook.
type At struct {
	Sheet    string // as the workbook spells it
	Row, Col int
}

// Source is cells whose change the formulas reading them follow.
type Source struct {
	Sheet string
	Range grid.Range
	// Readers: only what reads the cells follows, not the formulas among
	// them (rows that moved, with their formulas, which read as before).
	Readers bool
	// Reordered is the block the cells moved within, among themselves (rows
	// a sort moved, cells cut and pasted): a formula reading them only as a
	// set (SUM, MAX, COUNT, an average, a table's column) over the whole of
	// the block's rows does not follow, as no reordering changes what it
	// computes. Across: they moved across columns too, so the set must take
	// in the block's columns as well. Zero for cells that changed.
	Reordered grid.Range
	Across    bool
	// Carried are the cells of the block that moved (as numbered now): a
	// formula among them moved with what it reads, rewritten as for a copy,
	// so it does not follow the source (what moved it, if anything, is the
	// row it now stands in).
	Carried []grid.Range
}

// Tracer follows a workbook's formulas forward from cells.
type Tracer struct {
	fs     []*formulaAt
	keys   []cellKey // each formula's cell, by index
	refs   [][]formula.Ref
	opaque map[cellKey]string   // why what a formula reads cannot be told
	at     map[cellKey]int      // formulas, by cell
	arrays map[cellKey]arrayAt  // array formulas over several cells, by anchor
	anchor map[cellKey]cellKey  // the cells of their results, but the anchor
	names  map[string][]cellKey // formulas by the names they use (lower-cased)
	sheets map[string]*grid.Sheet
	used   map[string]grid.Range // what each sheet uses, by lower-cased name
	idx    *depIndex             // who reads what, kept for SpreadAll
	cols   map[string][]int      // by lower-cased sheet: the columns worth looking at (see columnsOf)
	wb     *grid.Workbook
}

type arrayAt struct {
	sheet string
	rg    grid.Range
}

// NewTracer reads what each formula of a workbook reads.
func NewTracer(wb *grid.Workbook) *Tracer {
	t := &Tracer{opaque: map[cellKey]string{}, at: map[cellKey]int{}, arrays: map[cellKey]arrayAt{},
		anchor: map[cellKey]cellKey{}, names: map[string][]cellKey{}, sheets: map[string]*grid.Sheet{}, used: map[string]grid.Range{}, wb: wb}
	rs := wbResolver(wb)
	for _, s := range wb.Sheets {
		t.sheets[strings.ToLower(s.Name)] = s
		t.used[strings.ToLower(s.Name)] = s.Used()
		s.Each(s.Used(), func(r, c int, cell *grid.Cell) {
			if cell.Formula == "" {
				return
			}
			f := &formulaAt{sheet: s.Name, row: r, col: c, text: cell.Formula}
			f.calc = rs.calcText(f)
			k := keyOf(s.Name, r, c)
			t.at[k] = len(t.fs)
			t.fs = append(t.fs, f)
			t.keys = append(t.keys, k)
			d := formula.DepsOf(f.calc)
			for _, n := range formula.DepsOf(f.text).Names {
				ln := strings.ToLower(n)
				t.names[ln] = append(t.names[ln], k)
			}
			refs, ok := rs.ranges(s.Name, r, c, d, 0)
			if !ok && !d.Volatile && !d.External {
				refs, ok = rs.tableRanges(f, d)
			}
			if !ok || d.External {
				t.opaque[k] = opaqueWhy(f.calc, d)
				refs = nil
			} else if why := outsideWhy(f.calc); why != "" {
				// What it reads is known (a change there is its cause), but its
				// result comes from outside the workbook too.
				t.opaque[k] = why
			}
			t.refs = append(t.refs, refs)
		})
		for _, rg := range append(append([]grid.Range{}, s.Spills...), s.Arrays...) {
			if rg.Rows()*rg.Cols() < 2 {
				continue
			}
			a := keyOf(s.Name, rg.R1, rg.C1)
			t.arrays[a] = arrayAt{s.Name, rg}
			for r := rg.R1; r <= rg.R2; r++ {
				for c := rg.C1; c <= rg.C2; c++ {
					if r != rg.R1 || c != rg.C1 {
						t.anchor[keyOf(s.Name, r, c)] = a
					}
				}
			}
		}
	}
	return t
}

// opaqueWhy says why what a formula reads cannot be told.
func opaqueWhy(text string, d formula.Deps) string {
	switch {
	case d.External:
		return linkNote
	case d.Multi:
		return "it reads across several sheets (Jan:Dec!B2), which sheets cannot follow"
	case d.Volatile:
		var fns []string
		for _, fn := range formula.CalledFunctions(text) {
			switch fn {
			case "NOW", "TODAY", "RAND", "RANDBETWEEN", "RANDARRAY", "INDIRECT", "OFFSET", "CELL", "INFO", "GETPIVOTDATA":
				fns = append(fns, fn)
			}
		}
		if len(fns) == 0 {
			return "it calls a volatile function, whose result can change while what it reads does not"
		}
		return "it calls " + strings.Join(fns, ", ") + ", whose result can change while what it names does not"
	}
	return "what it reads cannot be told (a name or table reference sheets cannot follow)"
}

// outsideFns are the functions whose result comes from outside the
// workbook: a server (RTD), the web (WEBSERVICE), a data service
// (STOCKHISTORY, the CUBE functions, TRANSLATE, DETECTLANGUAGE, COPILOT),
// Python in Excel (PY). Their stored values change with nothing in the
// workbook changed.
var outsideFns = map[string]bool{"RTD": true, "WEBSERVICE": true, "STOCKHISTORY": true, "CUBEVALUE": true, "CUBEMEMBER": true,
	"CUBEMEMBERPROPERTY": true, "CUBERANKEDMEMBER": true, "CUBESET": true, "CUBESETCOUNT": true, "CUBEKPIMEMBER": true,
	"TRANSLATE": true, "DETECTLANGUAGE": true, "COPILOT": true, "PY": true}

// outsideWhy says which functions of a formula bring in a result from
// outside the workbook; "" when none does.
func outsideWhy(text string) string {
	var fns []string
	for _, fn := range formula.CalledFunctions(text) {
		if outsideFns[strings.TrimPrefix(fn, "_XLWS.")] {
			fns = append(fns, strings.TrimPrefix(fn, "_XLWS."))
		}
	}
	if len(fns) == 0 {
		return ""
	}
	return "it calls " + strings.Join(fns, ", ") + ", whose result comes from outside the workbook"
}

// Opaque says why what a formula reads cannot be told (a volatile
// function, a link to another workbook), or why its result comes from
// outside the workbook (RTD, WEBSERVICE); "" when neither.
func (t *Tracer) Opaque(sheet string, row, col int) string {
	return t.opaque[keyOf(sheet, row, col)]
}

// Formula tells whether a cell holds a formula.
func (t *Tracer) Formula(sheet string, row, col int) bool {
	_, ok := t.at[keyOf(sheet, row, col)]
	return ok
}

// Outer gives the call a formula cell is, alone (=SUM(Sales[Units]),
// =MAX(Units), =COUNT(B2:B9)), with what it is given resolved to references:
// the columns of Excel tables as the cells they are, defined names as what
// they refer to. It says false for a cell that holds no formula, for one that
// is no call over references alone, and for a name that is none.
func (t *Tracer) Outer(sheet string, row, col int) (engine.Call, bool) {
	id, ok := t.at[keyOf(sheet, row, col)]
	if !ok {
		return engine.Call{}, false
	}
	f := t.fs[id]
	return engine.OuterNamed(f.calc, func(on, n string) (engine.RefArg, bool) { return t.nameRef(f.sheet, on, n) })
}

// Anchor gives the cell of the array formula whose result a cell holds
// (not the formula's own cell).
func (t *Tracer) Anchor(sheet string, row, col int) (At, bool) {
	a, ok := t.anchor[keyOf(sheet, row, col)]
	if !ok {
		return At{}, false
	}
	return t.atOf(a), true
}

func (t *Tracer) atOf(k cellKey) At {
	if s := t.sheets[k.sheet]; s != nil {
		return At{s.Name, k.row, k.col}
	}
	return At{k.sheet, k.row, k.col}
}

// Using lists the formulas that use a defined name.
func (t *Tracer) Using(name string) []At {
	var out []At
	for _, k := range t.names[strings.ToLower(name)] {
		out = append(out, t.atOf(k))
	}
	return out
}

// Spread follows the formulas reading the sources, directly or through
// other formulas: each it reaches (and each cell of an array formula's
// result) gets the index of the source that reaches it first, breadth
// first, so the nearest cause. Formulas whose reads cannot be told
// (Opaque) are not reached, but a formula given as a source spreads. The
// cells of a source Reordered are not followed by a formula that reads
// them as a set over the whole block (see Source).
func (t *Tracer) Spread(sources []Source) map[At]int { return t.SpreadBut(sources, nil) }

// SpreadBut is Spread, the formulas in stop neither reached nor followed
// through: what moves them is their own (a formula changed).
func (t *Tracer) SpreadBut(sources []Source, stop map[At]bool) map[At]int {
	if stop == nil {
		return t.SpreadUnless(sources, nil)
	}
	return t.SpreadUnless(sources, func(at At) bool { return stop[at] })
}

// SpreadUnless is Spread, the formulas that blocked tells of (nil: none)
// neither reached nor followed through: what moves them is not what reaches
// them (a formula whose value held, a formula another change moved).
func (t *Tracer) SpreadUnless(sources []Source, blocked func(At) bool) map[At]int {
	x := &depIndex{cells: map[cellKey][]int{}, columns: map[cellKey]*rowTree{}, wide: map[string]map[grid.Range][]int{}}
	for id, refs := range t.refs {
		for _, r := range refs {
			x.add(id, r.Sheet, r.Range)
		}
	}
	out := map[At]int{}
	type item struct {
		k   cellKey
		src int
		// reordered: a cell of a source Reordered, whose readers are found
		// without taking them from the index (another way to the cell may
		// still be followed by all of them).
		reordered bool
		// wide: the source's cells are too many to look up one by one (rows
		// added across a sheet whose used range spans the grid): its readers
		// are found by what each formula reads.
		wide bool
	}
	var queue []item
	// queued: cells whose readers are looked for (queuedR: as cells of a
	// source Reordered); reached: formulas (and array results) that follow a
	// source.
	queued, queuedR, reached := map[cellKey]bool{}, map[cellKey]bool{}, map[cellKey]bool{}
	push := func(k cellKey, src int) {
		if !queued[k] {
			queued[k] = true
			queue = append(queue, item{k: k, src: src})
		}
	}
	var reach func(k cellKey, src int)
	reach = func(k cellKey, src int) {
		if reached[k] {
			return
		}
		reached[k] = true
		at := t.atOf(k)
		if blocked != nil && blocked(at) {
			return
		}
		out[at] = src
		push(k, src)
		if a, ok := t.arrays[k]; ok {
			for r := a.rg.R1; r <= a.rg.R2; r++ {
				for c := a.rg.C1; c <= a.rg.C2; c++ {
					if r != a.rg.R1 || c != a.rg.C1 {
						reach(keyOf(a.sheet, r, c), src)
					}
				}
			}
		}
	}
	for i, s := range sources {
		rg := s.Range
		if u, ok := t.used[strings.ToLower(s.Sheet)]; ok && !u.Empty() {
			rg = clipUsed(rg, u)
		}
		cols := t.columnsOf(s.Sheet, rg.C1, rg.C2)
		if !rg.Empty() && rg.Rows()*len(cols) > wideSource {
			queue = append(queue, item{src: i, wide: true})
			continue
		}
		for r := rg.R1; r <= rg.R2; r++ {
			for _, c := range cols {
				k := keyOf(s.Sheet, r, c)
				if _, ok := t.at[k]; ok && !s.Readers {
					reach(k, i) // a formula changed: what it computes follows
					continue
				}
				if s.Reordered.R2 > 0 {
					if !queued[k] && !queuedR[k] {
						queuedR[k] = true
						queue = append(queue, item{k: k, src: i, reordered: true})
					}
					continue
				}
				push(k, i)
			}
		}
	}
	// A source Reordered belongs to a group: those with the same sheet, block
	// and orientation (every run of rows a sort moved is a source of its own
	// over the same block), for orderFree to be told once of a formula and a
	// group, not once for each run.
	type blockKey struct {
		sheet  string
		block  grid.Range
		across bool
	}
	groups := map[blockKey]int{}
	group := make([]int, len(sources))
	sheetOf := make([]string, len(sources)) // lower-cased
	for i, s := range sources {
		if s.Reordered.R2 > 0 {
			sheetOf[i] = strings.ToLower(s.Sheet)
			k := blockKey{sheetOf[i], s.Reordered, s.Across}
			g, ok := groups[k]
			if !ok {
				g = len(groups)
				groups[k] = g
			}
			group[i] = g
		}
	}
	free := map[[2]int]bool{} // by formula and group
	// carried indexes each source's Carried by rows (the runs of one sort share
	// them: looked up for every formula a run's cells reach, so not scanned).
	type carriedKey struct {
		first *grid.Range
		n     int
	}
	type carriedRows struct {
		rs   []grid.Range // by first row
		last []int        // the largest last row among rs[:i+1]
	}
	carried := map[carriedKey]*carriedRows{}
	carriedAt := func(rgs []grid.Range, row, col int) bool {
		if len(rgs) <= 4 {
			for _, rg := range rgs {
				if rg.Contains(row, col) {
					return true
				}
			}
			return false
		}
		key := carriedKey{&rgs[0], len(rgs)}
		c := carried[key]
		if c == nil {
			c = &carriedRows{rs: append([]grid.Range{}, rgs...)}
			sort.SliceStable(c.rs, func(i, j int) bool { return c.rs[i].R1 < c.rs[j].R1 })
			hi := 0
			for _, rg := range c.rs {
				hi = max(hi, rg.R2)
				c.last = append(c.last, hi)
			}
			carried[key] = c
		}
		for i := sort.Search(len(c.rs), func(i int) bool { return c.rs[i].R1 > row }) - 1; i >= 0 && c.last[i] >= row; i-- {
			if c.rs[i].Contains(row, col) {
				return true
			}
		}
		return false
	}
	// follows tells whether a formula reading a source's cells follows it.
	follows := func(id, src int) bool {
		s := &sources[src]
		if s.Reordered.R2 == 0 {
			return true
		}
		// A formula that moved with its row was rewritten for the move, as for a
		// copy, and reads its own row: except an array's, whose result spans
		// rows and reads others besides (a sort leaves its formula where it
		// is, and what it reads moved).
		if f := t.fs[id]; t.keys[id].sheet == sheetOf[src] && carriedAt(s.Carried, f.row, f.col) && t.arrays[t.keys[id]].rg.Rows() < 2 {
			return false
		}
		key := [2]int{id, group[src]}
		isFree, ok := free[key]
		if !ok {
			isFree = t.orderFree(id, *s)
			free[key] = isFree
		}
		return !isFree
	}
	for len(queue) > 0 {
		it := queue[0]
		queue = queue[1:]
		if it.wide {
			s := sources[it.src]
			rg := s.Range
			if u, ok := t.used[strings.ToLower(s.Sheet)]; ok {
				rg = clipUsed(rg, u)
			}
			for id, f := range t.fs {
				k := keyOf(f.sheet, f.row, f.col)
				if !s.Readers && strings.EqualFold(f.sheet, s.Sheet) && rg.Contains(f.row, f.col) {
					reach(k, it.src)
					continue
				}
				for _, ref := range t.refs[id] {
					if strings.EqualFold(ref.Sheet, s.Sheet) && !ref.Range.Intersect(rg).Empty() && follows(id, it.src) {
						reach(k, it.src)
						break
					}
				}
			}
			continue
		}
		if !it.reordered {
			x.readers(it.k, func(id int) {
				f := t.fs[id]
				reach(keyOf(f.sheet, f.row, f.col), it.src)
			})
			continue
		}
		if queued[it.k] {
			continue // reached another way too: followed as that
		}
		x.readersOf(it.k, func(id int) {
			// One already reached is not looked at again (a lookup over the whole
			// table is a reader of every cell of it).
			if k := t.keys[id]; !reached[k] && follows(id, it.src) {
				reach(k, it.src)
			}
		})
	}
	return out
}

// wideSource is how many cells a source may span for Spread to look up the
// readers of each: past it, it reads each formula's references instead.
const wideSource = 1 << 20

// wholeRow is how many columns a source may span for Spread to look at each
// of them: past it, only the columns that can make a difference (columnsOf).
const wholeRow = 64

// clipUsed cuts a source that is whole rows or columns to the cells the sheet
// uses, which is what a formula reading them can find changed. A source that is
// cells is kept, though beyond what the sheet uses now: a cell cleared, the
// last of its row and column, is read by the formulas whose ranges reach it.
func clipUsed(rg, used grid.Range) grid.Range {
	if rg.R2 < grid.MaxRows && rg.C2 < grid.MaxCols {
		return rg
	}
	return rg.Intersect(grid.Range{R1: 1, C1: 1, R2: used.R2, C2: used.C2})
}

// columnsOf gives the columns c1..c2 of a source's rows that Spread looks
// at. A source spanning few columns gives them all. One spanning many (rows
// added or moved across a sheet, whose used range may reach the last column
// for one far cell) gives the first column and those where the sheet holds a
// cell or a formula reads from: of the rest no formula reads a cell and none
// is there, so looking at them finds nothing (a formula reading some of the
// source's cells reads the first column or starts in one of them).
func (t *Tracer) columnsOf(sheet string, c1, c2 int) []int {
	if c2 < c1 {
		return nil
	}
	if c2-c1 < wholeRow {
		cols := make([]int, 0, c2-c1+1)
		for c := c1; c <= c2; c++ {
			cols = append(cols, c)
		}
		return cols
	}
	if t.cols == nil {
		held := map[string]map[int]bool{}
		for _, s := range t.wb.Sheets {
			m := map[int]bool{}
			s.Each(s.Used(), func(_, col int, _ *grid.Cell) { m[col] = true })
			held[strings.ToLower(s.Name)] = m
		}
		for _, refs := range t.refs {
			for _, r := range refs {
				k := strings.ToLower(r.Sheet)
				if held[k] == nil {
					held[k] = map[int]bool{}
				}
				held[k][r.Range.C1] = true
			}
		}
		t.cols = map[string][]int{}
		for k, m := range held {
			cols := make([]int, 0, len(m))
			for c := range m {
				cols = append(cols, c)
			}
			sort.Ints(cols)
			t.cols[k] = cols
		}
	}
	all := t.cols[strings.ToLower(sheet)]
	cols := []int{c1}
	for i := sort.SearchInts(all, c1+1); i < len(all) && all[i] <= c2; i++ {
		cols = append(cols, all[i])
	}
	return cols
}

// orderFree tells whether a formula reads what it reads of a source's
// block (Source.Reordered) only as a set over all of its rows (and, Across,
// its columns): ranges given whole to SUM, MAX, COUNT and the like, those a
// SUMIF or SUMPRODUCT pairs spanning the same rows within the block.
func (t *Tracer) orderFree(id int, s Source) bool {
	f := t.fs[id]
	reads, ok := engine.OrderReads(f.calc, func(sheet, n string) (engine.RefArg, bool) { return t.nameRef(f.sheet, sheet, n) })
	if !ok {
		return false
	}
	p := s.Reordered
	type groupRef struct {
		on bool
		rg grid.Range
	}
	groups := map[int][]groupRef{}
	touched := map[int]bool{}
	read := false
	for _, r := range reads {
		sheet := r.Sheet
		if sheet == "" {
			sheet = f.sheet
		}
		rg := grid.Range{R1: r.Area.R1, C1: r.Area.C1, R2: r.Area.R2, C2: r.Area.C2}
		on := strings.EqualFold(sheet, s.Sheet)
		hit := on && !rg.Intersect(p).Empty()
		if r.Group > 0 {
			groups[r.Group] = append(groups[r.Group], groupRef{on, rg})
			touched[r.Group] = touched[r.Group] || hit
		}
		if !hit {
			continue
		}
		read = true
		if !r.Set || rg.R1 > p.R1 || rg.R2 < p.R2 || s.Across && (rg.C1 > p.C1 || rg.C2 < p.C2) {
			return false
		}
	}
	for g, refs := range groups {
		if !touched[g] {
			continue
		}
		for _, x := range refs {
			if s.Across || !x.on || x.rg.R1 != refs[0].rg.R1 || x.rg.R2 != refs[0].rg.R2 || x.rg.C1 < p.C1 || x.rg.C2 > p.C2 {
				return false
			}
		}
	}
	return read
}

// nameRef resolves a defined name a formula on home uses (qualified by
// sheet, or not) to the reference it stands for.
func (t *Tracer) nameRef(home, sheet, name string) (engine.RefArg, bool) {
	var found *grid.DefinedName
	for i, n := range t.wb.Names {
		if !strings.EqualFold(n.Name, name) {
			continue
		}
		scope := sheet
		if scope == "" {
			scope = home
		}
		switch {
		case strings.EqualFold(n.Scope, scope):
			found = &t.wb.Names[i]
		case n.Scope == "" && found == nil:
			found = &t.wb.Names[i]
		}
	}
	if found == nil {
		return engine.RefArg{}, false
	}
	r, ok := engine.Reference(strings.TrimPrefix(found.Ref, "="))
	if ok && r.Sheet == "" && found.Scope != "" {
		r.Sheet = found.Scope
	}
	return r, ok
}

// Reach is the sources that reach a cell: the indexes of the lowest few, in
// increasing order, and whether others do.
type Reach struct {
	Sources []int
	More    bool
}

// SpreadAll follows the formulas reading the sources like Spread, but tells
// of every source reaching each formula (and each cell of an array formula's
// result), up to limit of them, the lowest in index: a value that two inputs
// moved is the work of both. It stops where a cell has all it can hold, so
// the cost is bounded by limit and not by the sources.
func (t *Tracer) SpreadAll(sources []Source, limit int) map[At]Reach {
	return t.SpreadAllUnless(sources, limit, nil)
}

// SpreadAllUnless is SpreadAll, the formulas that blocked tells of (nil: none)
// neither reached nor followed through (see SpreadUnless).
func (t *Tracer) SpreadAllUnless(sources []Source, limit int, blocked func(At) bool) map[At]Reach {
	if t.idx == nil {
		t.idx = &depIndex{cells: map[cellKey][]int{}, columns: map[cellKey]*rowTree{}, wide: map[string]map[grid.Range][]int{}}
		for id, refs := range t.refs {
			for _, r := range refs {
				t.idx.add(id, r.Sheet, r.Range)
			}
		}
	}
	x := t.idx
	labels := map[cellKey][]int{} // up to limit+1 each: one over says there are more
	visited := map[cellKey]int{}  // by the source (plus one) that last reached it
	var queue []cellKey
	var reach func(k cellKey, src int)
	reach = func(k cellKey, src int) {
		if visited[k] == src+1 {
			return
		}
		visited[k] = src + 1
		if blocked != nil && blocked(t.atOf(k)) {
			return
		}
		if len(labels[k]) > limit {
			return // it has all it can hold, and so have those after it
		}
		labels[k] = append(labels[k], src)
		queue = append(queue, k)
		if a, ok := t.arrays[k]; ok {
			for r := a.rg.R1; r <= a.rg.R2; r++ {
				for c := a.rg.C1; c <= a.rg.C2; c++ {
					if r != a.rg.R1 || c != a.rg.C1 {
						reach(keyOf(a.sheet, r, c), src)
					}
				}
			}
		}
	}
	for i, s := range sources {
		rg := s.Range
		if u, ok := t.used[strings.ToLower(s.Sheet)]; ok && !u.Empty() {
			rg = clipUsed(rg, u)
		}
		queue = queue[:0]
		cols := t.columnsOf(s.Sheet, rg.C1, rg.C2)
		for r := rg.R1; r <= rg.R2; r++ {
			for _, c := range cols {
				k := keyOf(s.Sheet, r, c)
				if _, ok := t.at[k]; ok && !s.Readers {
					reach(k, i) // a formula changed: what it computes follows
					continue
				}
				visited[k] = i + 1
				queue = append(queue, k)
			}
		}
		for n := 0; n < len(queue); n++ {
			x.readersOf(queue[n], func(id int) {
				f := t.fs[id]
				reach(keyOf(f.sheet, f.row, f.col), i)
			})
		}
	}
	out := make(map[At]Reach, len(labels))
	for k, ls := range labels {
		r := Reach{Sources: ls}
		if len(ls) > limit {
			r.Sources, r.More = ls[:limit], true
		}
		out[t.atOf(k)] = r
	}
	return out
}

// Touching lists the formulas that read cells in lines of a sheet (rows, or
// columns with cols set), given as runs of line numbers in increasing order
// (the lines a later version removed, found by what this version reads).
// Each maps to the index of the first run it reads in.
func (t *Tracer) Touching(sheet string, runs [][2]int, cols bool) map[At]int {
	out := map[At]int{}
	if len(runs) == 0 {
		return out
	}
	for id, refs := range t.refs {
		f := t.fs[id]
		for _, r := range refs {
			if !strings.EqualFold(r.Sheet, sheet) {
				continue
			}
			lo, hi := r.Range.R1, r.Range.R2
			if cols {
				lo, hi = r.Range.C1, r.Range.C2
			}
			k := sort.Search(len(runs), func(i int) bool { return runs[i][1] >= lo })
			if k >= len(runs) || runs[k][0] > hi {
				continue
			}
			at := At{f.sheet, f.row, f.col}
			if prev, ok := out[at]; !ok || k < prev {
				out[at] = k
			}
		}
	}
	return out
}

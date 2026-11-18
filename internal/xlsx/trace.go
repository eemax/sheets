package xlsx

import (
	"sort"
	"strings"

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
}

// Tracer follows a workbook's formulas forward from cells.
type Tracer struct {
	fs     []*formulaAt
	refs   [][]formula.Ref
	opaque map[cellKey]string   // why what a formula reads cannot be told
	at     map[cellKey]int      // formulas, by cell
	arrays map[cellKey]arrayAt  // array formulas over several cells, by anchor
	anchor map[cellKey]cellKey  // the cells of their results, but the anchor
	names  map[string][]cellKey // formulas by the names they use (lower-cased)
	sheets map[string]*grid.Sheet
	used   map[string]grid.Range // what each sheet uses, by lower-cased name
}

type arrayAt struct {
	sheet string
	rg    grid.Range
}

// NewTracer reads what each formula of a workbook reads.
func NewTracer(wb *grid.Workbook) *Tracer {
	t := &Tracer{opaque: map[cellKey]string{}, at: map[cellKey]int{}, arrays: map[cellKey]arrayAt{},
		anchor: map[cellKey]cellKey{}, names: map[string][]cellKey{}, sheets: map[string]*grid.Sheet{}, used: map[string]grid.Range{}}
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
// (Opaque) are not reached, but a formula given as a source spreads.
func (t *Tracer) Spread(sources []Source) map[At]int {
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
	}
	var queue []item
	// queued: cells whose readers are looked for; reached: formulas (and
	// array results) that follow a source.
	queued, reached := map[cellKey]bool{}, map[cellKey]bool{}
	push := func(k cellKey, src int) {
		if !queued[k] {
			queued[k] = true
			queue = append(queue, item{k, src})
		}
	}
	var reach func(k cellKey, src int)
	reach = func(k cellKey, src int) {
		if reached[k] {
			return
		}
		reached[k] = true
		out[t.atOf(k)] = src
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
			// Whole rows or columns: the cells the sheet uses, which is
			// what a formula reading them can find changed.
			rg = rg.Intersect(grid.Range{R1: 1, C1: 1, R2: u.R2, C2: u.C2})
		}
		for r := rg.R1; r <= rg.R2; r++ {
			for c := rg.C1; c <= rg.C2; c++ {
				k := keyOf(s.Sheet, r, c)
				if _, ok := t.at[k]; ok && !s.Readers {
					reach(k, i) // a formula changed: what it computes follows
					continue
				}
				push(k, i)
			}
		}
	}
	for len(queue) > 0 {
		it := queue[0]
		queue = queue[1:]
		x.readers(it.k, func(id int) {
			f := t.fs[id]
			reach(keyOf(f.sheet, f.row, f.col), it.src)
		})
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

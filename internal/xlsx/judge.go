package xlsx

import (
	"math"
	"strconv"
	"strings"
	"time"

	"github.com/eemax/sheets/internal/engine"
	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
)

// The engine's answer for a formula holds only when it holds for
// everything the formula reads, and the judge checks that, through the
// graph of what reads what:
//
//   - the formula is computable at all (see computable) and the engine
//     computes it without meeting anything it cannot (a function it does
//     not have, a link to another workbook, a date written in a way only
//     Excel's language setting decides);
//   - it is not in, and does not read, a circular reference;
//   - where Excel stored a result the engine recomputes on the way, the
//     engine agrees with it;
//   - the same holds for everything it reads.

type judge struct {
	eng      *engine.Engine
	g        *cycleGraph
	cycle    map[*formulaAt]bool // in or reading a circular reference
	members  []*formulaAt        // in one
	rs       *resolver
	ty       *typer
	fresh    map[*formulaAt]bool // results being recomputed: none to check against
	frozen   map[*formulaAt]bool // taken at their stored results (see calcPlan)
	veto     map[*formulaAt]string
	native   func(j *judge, f *formulaAt) (grid.Cell, bool)
	stored   func(f *formulaAt) (grid.Cell, bool)
	deadline time.Time
	at       map[cellKey]*formulaAt

	evals map[*formulaAt]*judged
	state []int8   // per node: 0 not seen, 1 holds, 2 does not, 3 working
	why   []string // per node that does not hold: why
}

// judged is the engine's answer for a formula: a value, or err with a
// note saying why there is none.
type judged struct {
	v    grid.Cell
	note string
	err  error
}

// newJudge judges the engine's answers for the formulas of g. The engine
// takes formulas the plan freezes at their stored results, and computes
// aggregates of long ranges natively (see nativeAggregates).
func newJudge(eng *engine.Engine, g *cycleGraph, rs *resolver, ty *typer, fresh map[*formulaAt]bool, plan calcPlan,
	stored func(*formulaAt) (grid.Cell, bool), deadline time.Time) *judge {
	cycle, members := g.cycles()
	j := &judge{eng: eng, g: g, cycle: cycle, members: members, rs: rs, ty: ty, fresh: fresh, frozen: plan.frozen, veto: plan.veto, native: plan.native, stored: stored,
		deadline: deadline, at: map[cellKey]*formulaAt{}, evals: map[*formulaAt]*judged{}, state: make([]int8, len(g.edges)), why: make([]string, len(g.edges))}
	for _, f := range g.fs {
		j.at[keyOf(f.sheet, f.row, f.col)] = f
	}
	eng.Override = func(sheet string, row, col int) (engine.Value, bool) {
		f := j.at[keyOf(sheet, row, col)]
		if f == nil {
			return engine.Value{}, false
		}
		if j.frozen[f] {
			return cellToCalc(f.stored), true
		}
		if j.native != nil {
			if v, ok := j.native(j, f); ok {
				return cellToCalc(v), true
			}
		}
		return engine.Value{}, false
	}
	return j
}

// result is the engine's answer for f, computed once.
func (j *judge) result(f *formulaAt) *judged {
	if r, ok := j.evals[f]; ok {
		return r
	}
	var r *judged
	if j.frozen[f] {
		r = &judged{v: f.stored}
	} else {
		r = judgedOf(j.eng.Compute(f.sheet, f.row, f.col))
		if r.err == nil && r.v.Kind == grid.Error && r.v.Text == "#NAME?" {
			// Say which name Excel does not know.
			isName := func(n string) bool { return j.rs.lookup(n, f.sheet) != nil }
			if unknown := formula.UnknownFunctions(f.text, isName); len(unknown) > 0 {
				r.note = "unknown function " + strings.Join(unknown, ", ")
			}
		}
	}
	j.evals[f] = r
	return r
}

// judgedOf reads an engine value: unknown ones carry why.
func judgedOf(v engine.Value) *judged {
	if v.Kind == engine.Error {
		switch v.Str {
		case engine.ErrUnsupported:
			return &judged{note: v.Why(), err: errUnsupported}
		case engine.ErrCycle:
			return &judged{note: "it reads a circular reference", err: errUnsupported}
		}
	}
	return &judged{v: calcToCell(v)}
}

// holds reports whether the engine's answer for f can be trusted;
// reason says why not.
func (j *judge) holds(f *formulaAt) (bool, string) {
	i, ok := j.g.id[f]
	if !ok {
		return false, "not a formula sheets knows of"
	}
	if j.node(i) {
		return true, ""
	}
	return false, j.why[i]
}

func (j *judge) node(v int) bool {
	switch j.state[v] {
	case 1:
		return true
	case 2, 3:
		return false // 3: on the path, which cycles were marked to prevent
	}
	if v < j.g.n && j.frozen[j.g.fs[v]] {
		j.state[v] = 1 // its stored result stands, whatever it reads
		return true
	}
	j.state[v] = 3
	// What it reads first: a failure there is the cause, and the
	// engine's answer for this formula would be built on it.
	why := j.static(v)
	if why == "" {
		for _, w := range j.g.edges[v] {
			if !j.node(int(w)) {
				why = j.why[w]
				if int(w) < j.g.n && !strings.HasPrefix(why, "it reads ") {
					f := j.g.fs[w]
					why = "it reads " + grid.FormatRef(f.sheet, grid.Range{R1: f.row, C1: f.col, R2: f.row, C2: f.col}) + ", whose result is unknown here (" + why + ")"
				}
				break
			}
		}
	}
	if why == "" {
		why = j.check(v)
	}
	j.why[v] = why
	if why == "" {
		j.state[v] = 1
		return true
	}
	j.state[v] = 2
	return false
}

// static judges a formula by its text (nothing for a tree node); "" when
// the engine can compute it.
func (j *judge) static(v int) string {
	if v >= j.g.n {
		return ""
	}
	f := j.g.fs[v]
	if j.cycle[f] {
		return "circular reference: Excel cannot compute it"
	}
	if f.unprefixed {
		return unprefixedNote(f.text)
	}
	// What the calculator cannot do at all says more than why it is not
	// trusted here.
	if ok, why := computable(calcOf(f), j.rs, f.sheet); !ok {
		return why
	}
	return j.veto[f]
}

// unprefixedNote says why a formula stored without a function's prefix
// has no result.
func unprefixedNote(text string) string {
	fns := formula.Unprefixed(text)
	return strings.Join(fns, ", ") + map[bool]string{true: " is", false: " are"}[len(fns) == 1] +
		" stored without the _xlfn. prefix Excel needs, so Excel shows #NAME? until the formula is entered again (sheets set adds it)"
}

// calcOf is the text the engine computes for f.
func calcOf(f *formulaAt) string {
	if f.calc != "" {
		return f.calc
	}
	return f.text
}

// check judges the engine's answer for a formula whose inputs hold
// (nothing for a tree node); "" when it holds.
func (j *judge) check(v int) string {
	if v >= j.g.n {
		return ""
	}
	f := j.g.fs[v]
	if time.Now().After(j.deadline) {
		return "not computed in the time allowed"
	}
	r := j.result(f)
	if r.err != nil {
		return r.note
	}
	if j.fresh[f] || formula.DepsOf(f.text).Volatile {
		return ""
	}
	if want, ok := j.stored(f); ok && !agreesCell(r.v, want) {
		return "the sheets calculator gets " + show(r.v) + " where Excel stored " + show(want)
	}
	return ""
}

// agreesCell tells whether the engine's value matches a result Excel
// stored, allowing for the last digits of a float.
func agreesCell(got, want grid.Cell) bool {
	if got.Kind != want.Kind {
		return false
	}
	switch want.Kind {
	case grid.Number:
		a, b := got.Num, want.Num
		return a == b || math.Abs(a-b) <= 1e-9*math.Max(1, math.Max(math.Abs(a), math.Abs(b)))
	case grid.Text, grid.Error:
		return got.Text == want.Text
	case grid.Bool:
		return got.Bool == want.Bool
	}
	return true
}

func show(c grid.Cell) string {
	switch c.Kind {
	case grid.Text:
		return strconv.Quote(c.Text)
	case grid.Error:
		return c.Text
	case grid.Bool:
		return strings.ToUpper(strconv.FormatBool(c.Bool))
	case grid.Number:
		return grid.FormatNumber(c.Num)
	}
	return "nothing"
}

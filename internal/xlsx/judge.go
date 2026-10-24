package xlsx

import (
	"errors"
	"math"
	"strconv"
	"strings"
	"time"

	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
	"github.com/xuri/excelize/v2"
)

// The calculator answers for a formula by computing everything it reads,
// silently: a formula it cannot compute on the way reads as nothing, a
// circular reference as an old result, and where it disagrees with Excel
// the error carries on downstream. So its answer for a formula holds only
// when it holds for everything the formula reads, and the judge checks
// that, through the graph of what reads what:
//
//   - the formula is computable at all (see computable) and the calculator
//     computes it without failing;
//   - it is not in, and does not read, a circular reference;
//   - where Excel stored a result the calculator recomputes on the way,
//     the calculator agrees with it;
//   - the same holds for everything it reads.

type judge struct {
	f        *excelize.File
	names    func(string) bool
	g        *cycleGraph
	cycle    map[*formulaAt]bool // in or reading a circular reference
	members  []*formulaAt        // in one
	rs       *resolver
	ty       *typer
	fresh    map[*formulaAt]bool // results being recomputed: none to check against
	frozen   map[*formulaAt]bool // taken at their stored results (see calcPlan)
	stored   func(f *formulaAt) (grid.Cell, bool)
	deadline time.Time

	evals map[*formulaAt]*judged
	state []int8   // per node: 0 not seen, 1 holds, 2 does not, 3 working
	why   []string // per node that does not hold: why
}

// judged is the calculator's answer for a formula.
type judged struct {
	v    grid.Cell
	raw  string
	note string
	err  error
}

func newJudge(f *excelize.File, g *cycleGraph, rs *resolver, ty *typer, fresh, frozen map[*formulaAt]bool,
	stored func(*formulaAt) (grid.Cell, bool), deadline time.Time) *judge {
	cycle, members := g.cycles()
	return &judge{f: f, names: definedNames(f), g: g, cycle: cycle, members: members, rs: rs, ty: ty, fresh: fresh, frozen: frozen, stored: stored,
		deadline: deadline, evals: map[*formulaAt]*judged{}, state: make([]int8, len(g.edges)), why: make([]string, len(g.edges))}
}

// result is the calculator's answer for f, computed once.
func (j *judge) result(f *formulaAt) *judged {
	if r, ok := j.evals[f]; ok {
		return r
	}
	if j.frozen[f] {
		r := &judged{v: f.stored, raw: rawOf(f.stored)}
		j.evals[f] = r
		return r
	}
	v, raw, note, err := evaluateWith(j.f, j.names, f.sheet, grid.CellName(f.row, f.col))
	if errors.Is(err, errUncertain) || err == nil && v.Kind == grid.Error && v.Text == "#VALUE!" {
		if formula.InputKinds(calcOf(f), &sheetTyper{j.ty, f.sheet})&formula.KNumText == 0 {
			// #VALUE! from text no one reads as a number ("abc"*2):
			// Excel's answer too.
			v, err, note = grid.Cell{Kind: grid.Error, Text: "#VALUE!"}, nil, ""
		} else {
			err, note = errUncertain, "#VALUE! from the sheets calculator, which Excel may not give: it reads text as a number where it can"
		}
	}
	r := &judged{v: v, raw: raw, note: note, err: err}
	j.evals[f] = r
	return r
}

// holds reports whether the calculator's answer for f can be trusted;
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
	// calculator's answer for this formula would be built on it.
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
// the calculator can compute it.
func (j *judge) static(v int) string {
	if v >= j.g.n {
		return ""
	}
	f := j.g.fs[v]
	if j.cycle[f] {
		return "circular reference: Excel cannot compute it"
	}
	if ok, why := computable(calcOf(f), j.rs, f.sheet); !ok {
		return why
	}
	return ""
}

// calcOf is the text the calculator computes for f.
func calcOf(f *formulaAt) string {
	if f.calc != "" {
		return f.calc
	}
	return f.text
}

// rawOf writes a value as the calculator gives results, as text.
func rawOf(v grid.Cell) string {
	switch v.Kind {
	case grid.Number:
		return strconv.FormatFloat(v.Num, 'g', -1, 64)
	case grid.Text, grid.Error:
		return v.Text
	case grid.Bool:
		return strings.ToUpper(strconv.FormatBool(v.Bool))
	}
	return ""
}

// check judges the calculator's answer for a formula whose inputs hold
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
		return strings.TrimSuffix(strings.TrimSuffix(r.note, " (Excel computes it on open)"), "; Excel computes it on open")
	}
	if j.fresh[f] || formula.DepsOf(f.text).Volatile {
		return ""
	}
	if want, ok := j.stored(f); ok && !agrees(r, want) {
		return "the sheets calculator gets " + show(r.v) + " where Excel stored " + show(want)
	}
	return ""
}

// agrees tells whether the calculator's answer matches a result Excel
// stored, allowing for the last digits of a float.
func agrees(r *judged, want grid.Cell) bool {
	switch want.Kind {
	case grid.Number:
		n, err := strconv.ParseFloat(r.raw, 64)
		if r.v.Kind == grid.Bool || err != nil {
			return false
		}
		return n == want.Num || math.Abs(n-want.Num) <= 1e-9*math.Max(1, math.Max(math.Abs(n), math.Abs(want.Num)))
	case grid.Text:
		return r.v.Kind != grid.Error && r.raw == want.Text
	case grid.Bool:
		u := strings.ToUpper(r.raw)
		return (u == "TRUE" || r.raw == "1") == want.Bool && (u == "TRUE" || u == "FALSE" || r.raw == "1" || r.raw == "0")
	case grid.Error:
		return r.v.Kind == grid.Error && r.v.Text == want.Text
	}
	return false
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

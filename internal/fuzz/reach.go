package fuzz

import (
	"fmt"
	"math"
	"regexp"
	"sort"
	"strings"

	"github.com/eemax/sheets/internal/grid"
)

// The reach invariant: nothing outside an edit's reach changes. The model of
// each edit (model.command) says what it may change: the cells a set, write
// or clear names, the target of a copy, the source and target of a move, the
// range a sort takes, the cells a replace matches, the rows or columns an
// insert adds, a sheet the edit makes as a copy. Every other cell of the
// workbook is followed through the edit (rows and columns inserted or
// deleted move it, a sheet renamed takes it along), and must hold what it
// held before, where it stands now, and an empty cell must stay empty.
//
// Two changes outside the reach are expected:
//
//   - shifted references: an edit that moves cells (move, rows, cols, a
//     script's moveTo, insert and delete) or renames a sheet rewrites the
//     formulas that read them, so a formula may then differ from what it was
//     in its references only (each reference, or #REF!, is read as one
//     token); any other edit leaves every formula outside its reach as it
//     was, letter for letter;
//   - recomputed dependents: a formula's value may change when it reads a
//     cell that may (in the reach, or a formula whose value may change, to a
//     fixpoint), when its references were rewritten, or when its result was
//     unknown before or is unknown now (pending: sheets computes a result it
//     did not know, or drops one it cannot vouch for). The cells an array
//     fills (a dynamic array's spill, a legacy array's range) change with
//     their formula, and with an edit that reaches into them or cuts them.
//
// So a slip that writes a cell the edit does not name, moves one it should
// not, or loses or rewrites a formula elsewhere is a finding, as is a
// formula whose value changes with nothing it reads changed.
type reach struct {
	before *grid.Workbook
	sheets map[string]*grid.Sheet // the sheets before, by lower-case name
	// at are the cells outside the reach so far, non-empty before the edit,
	// by sheet (lower case) and where they stand now: where they stood.
	at map[string]map[[2]int]origin
	// rects is the reach, where it stands now, by sheet.
	rects map[string][]grid.Range
	// regions are the arrays of the workbook before (dynamic, legacy), where
	// they stand now; member the region each of their other cells belonged to.
	regions []*region
	member  map[origin]*region
	// rewrites: the edit may rewrite references (it moves cells, or renames
	// a sheet).
	rewrites bool
}

// origin is where a cell stood before the edit: its sheet (lower case) and
// cell.
type origin struct {
	sheet string
	r, c  int
}

// region is an array of the workbook before the edit, where it stands now.
type region struct {
	sheet     string
	rg        grid.Range // its range, with its formula's cell first
	disturbed bool       // the edit reached into it, or cut through it
	gone      bool       // its formula's row or column was deleted
}

func newReach(wb *grid.Workbook) *reach {
	rc := &reach{before: wb, sheets: map[string]*grid.Sheet{}, at: map[string]map[[2]int]origin{}, rects: map[string][]grid.Range{}, member: map[origin]*region{}}
	for _, s := range wb.Sheets {
		k := strings.ToLower(s.Name)
		rc.sheets[k] = s
		cells := map[[2]int]origin{}
		if u := s.Used(); u.R2 >= 1 {
			s.Each(u, func(r, c int, _ *grid.Cell) { cells[[2]int{r, c}] = origin{k, r, c} })
		}
		rc.at[k] = cells
		seen := map[[2]int]bool{}
		for _, a := range append(append([]grid.Range{}, s.Spills...), s.Arrays...) {
			if seen[[2]int{a.R1, a.C1}] {
				continue
			}
			seen[[2]int{a.R1, a.C1}] = true
			g := &region{sheet: k, rg: a}
			rc.regions = append(rc.regions, g)
			for r := a.R1; r <= a.R2; r++ {
				for c := a.C1; c <= a.C2; c++ {
					if r != a.R1 || c != a.C1 {
						rc.member[origin{k, r, c}] = g
					}
				}
			}
		}
	}
	return rc
}

// whole is every cell of a sheet.
var whole = grid.Range{R1: 1, C1: 1, R2: grid.MaxRows, C2: grid.MaxCols}

// touch puts a range of a sheet, as it stands now, in the edit's reach.
func (rc *reach) touch(k string, rg grid.Range) {
	if rc == nil || rg.Empty() {
		return
	}
	rc.rects[k] = append(rc.rects[k], rg)
	for p := range rc.at[k] {
		if rg.Contains(p[0], p[1]) {
			delete(rc.at[k], p)
		}
	}
	for _, g := range rc.regions {
		if g.sheet == k && !g.gone && !g.rg.Intersect(rg).Empty() {
			g.disturbed = true
		}
	}
}

// touchCells puts the cells now holding what stood at the origins match
// picks in the reach (a replace's matches).
func (rc *reach) touchCells(k string, match func(o origin) bool) {
	if rc == nil {
		return
	}
	for p, o := range rc.at[k] {
		if match(o) {
			rc.touch(k, grid.Range{R1: p[0], C1: p[1], R2: p[0], C2: p[1]})
		}
	}
}

// cell is a cell of the workbook before the edit.
func (rc *reach) cell(o origin) grid.Cell {
	if s := rc.sheets[o.sheet]; s != nil {
		return s.Get(o.r, o.c)
	}
	return grid.Cell{}
}

// shiftLine is where line p (a row, or a column) goes when n lines are
// inserted before at, or (n == 0) lines at..to are deleted; false when it is
// deleted.
func shiftLine(p, at, to, n int) (int, bool) {
	switch {
	case n > 0 && p >= at:
		return min(p+n, grid.MaxRows), true
	case n == 0 && p > to:
		return p - (to - at + 1), true
	case n == 0 && p >= at:
		return 0, false
	}
	return p, true
}

// shiftSpan is where the lines lo..hi go; false when all are deleted.
func shiftSpan(lo, hi, at, to, n int) (int, int, bool) {
	if n > 0 {
		if lo >= at {
			lo += n
		}
		if hi >= at {
			hi += n
		}
		return lo, hi, true
	}
	cnt := to - at + 1
	switch {
	case lo > to:
		lo -= cnt
	case lo >= at:
		lo = at
	}
	switch {
	case hi > to:
		hi -= cnt
	case hi >= at:
		hi = at - 1
	}
	return lo, hi, hi >= lo
}

// shift inserts n rows (columns) before at, or deletes at..to, on a sheet:
// the cells outside the reach, the reach and the arrays move with them, and
// the lines inserted are in the reach.
func (rc *reach) shift(k string, cols bool, at, to, n int) {
	if rc == nil {
		return
	}
	rc.rewrites = true
	i := 0
	if cols {
		i = 1
	}
	next := map[[2]int]origin{}
	for p, o := range rc.at[k] {
		q, ok := shiftLine(p[i], at, to, n)
		if !ok {
			continue
		}
		p[i] = q
		next[p] = o
	}
	rc.at[k] = next
	span := func(rg grid.Range) (grid.Range, bool) {
		lo, hi := rg.R1, rg.R2
		if cols {
			lo, hi = rg.C1, rg.C2
		}
		lo2, hi2, ok := shiftSpan(lo, hi, at, to, n)
		if cols {
			rg.C1, rg.C2 = lo2, min(hi2, grid.MaxCols)
		} else {
			rg.R1, rg.R2 = lo2, min(hi2, grid.MaxRows)
		}
		return rg, ok
	}
	var rects []grid.Range
	for _, rg := range rc.rects[k] {
		if g, ok := span(rg); ok {
			rects = append(rects, g)
		}
	}
	rc.rects[k] = rects
	for _, g := range rc.regions {
		if g.sheet != k || g.gone {
			continue
		}
		lo, hi := g.rg.R1, g.rg.R2
		if cols {
			lo, hi = g.rg.C1, g.rg.C2
		}
		// Lines inserted inside it, or deleted from it, cut it: sheets
		// computes it again (or refuses, for a legacy array).
		if n > 0 && lo < at && at <= hi || n == 0 && at <= hi && to >= lo {
			g.disturbed = true
		}
		if _, ok := shiftLine(lo, at, to, n); !ok {
			g.gone = true
			continue
		}
		g.rg, _ = span(g.rg)
	}
	if n > 0 {
		rg := grid.Range{R1: at, C1: 1, R2: at + n - 1, C2: grid.MaxCols}
		if cols {
			rg = grid.Range{R1: 1, C1: at, R2: grid.MaxRows, C2: at + n - 1}
		}
		rc.touch(k, rg)
	}
}

// rename follows a sheet to its new name (lower case).
func (rc *reach) rename(from, to string) {
	if rc == nil || from == to {
		return
	}
	rc.rewrites = true
	rc.at[to], rc.rects[to] = rc.at[from], rc.rects[from]
	delete(rc.at, from)
	delete(rc.rects, from)
	for _, g := range rc.regions {
		if g.sheet == from {
			g.sheet = to
		}
	}
}

// refToken is a reference in a formula's text, with its sheet, or one an
// edit broke (#REF!): what an edit that moves cells may rewrite.
var refToken = regexp.MustCompile(`(?:(?:'(?:[^']|'')+'|[A-Za-z0-9_]+)!)?(?:#REF!|\$?[A-Z]{1,3}\$?[0-9]+(?::\$?[A-Z]{1,3}\$?[0-9]+)?)`)

// masked is a formula's text with each reference one token.
func masked(f string) string { return refToken.ReplaceAllString(f, "§") }

// sameValue tells whether two cells hold the same value, both pending or
// neither.
func sameValue(a, b grid.Cell) bool {
	return a.Kind == b.Kind && a.Text == b.Text && a.Bool == b.Bool && a.Pending == b.Pending &&
		(a.Num == b.Num || math.IsNaN(a.Num) && math.IsNaN(b.Num))
}

// describe is a cell as a finding names it: its formula and its value.
func describe(c grid.Cell) string {
	if c.Formula != "" {
		return "=" + c.Formula + " (" + value(c) + ")"
	}
	return value(c)
}

// readRange is a range a formula reads: its sheet (lower case) and the range.
type readRange struct {
	sheet string
	rg    grid.Range
}

// inputs are the ranges a formula on a sheet (lower case) reads, by its
// references and the defined names it uses. ok is false when it reads
// something the model cannot place (then any value is expected of it).
func inputs(f, sheet string, names []usedName) (out []readRange, ok bool) {
	add := func(text, home string) bool {
		for _, m := range refRE.FindAllStringSubmatch(text, -1) {
			at := home
			if n := m[1] + m[2]; n != "" {
				at = strings.ToLower(strings.ReplaceAll(m[1], "''", "'") + m[2])
			}
			ref := m[3] + m[4]
			if m[5] != "" {
				ref += ":" + m[5] + m[6]
			}
			rg, err := grid.ParseRange(ref)
			if err != nil {
				return false
			}
			out = append(out, readRange{at, rg})
		}
		return true
	}
	text := stringsOut(f)
	ok = add(text, sheet)
	for _, n := range names {
		if n.re.MatchString(text) {
			// A name the edit defined again (rows deleted from its range) gives
			// another value with the same text.
			if n.moved || n.Scope != "" || !strings.Contains(n.Ref, "!") || !add(n.Ref, "") {
				ok = false
			}
		}
	}
	return out, ok
}

// usedName is a defined name, how a formula's text uses it, and whether the
// edit defined it again (moved: its range is another, or it is new).
type usedName struct {
	grid.DefinedName
	re    *regexp.Regexp
	moved bool
}

// usedNames are the names of the workbook after the edit; before are those
// it had before.
func usedNames(names, before []grid.DefinedName) []usedName {
	was := map[[2]string]string{}
	for _, n := range before {
		was[[2]string{strings.ToLower(n.Name), strings.ToLower(n.Scope)}] = n.Ref
	}
	var out []usedName
	for _, n := range names {
		ref, ok := was[[2]string{strings.ToLower(n.Name), strings.ToLower(n.Scope)}]
		out = append(out, usedName{n, regexp.MustCompile(`(?i)(^|[^A-Za-z0-9_.])` + regexp.QuoteMeta(n.Name) + `($|[^A-Za-z0-9_.(])`), !ok || ref != n.Ref})
	}
	return out
}

// stringsOut is a formula's text without its string literals ("Fig" is no
// reference, nor is N("p") a name).
func stringsOut(f string) string {
	var b strings.Builder
	in := false
	for _, r := range f {
		if r == '"' {
			in = !in
			continue
		}
		if !in {
			b.WriteRune(r)
		}
	}
	return b.String()
}

// check compares what the edit left outside its reach with what was there,
// and returns the first difference, "" when there is none. sheets are the
// sheets the model says the workbook holds (lower case).
func (rc *reach) check(after *grid.Workbook, sheets map[string]bool, order []string) string {
	now := map[string]*grid.Sheet{}
	for _, s := range after.Sheets {
		k := strings.ToLower(s.Name)
		now[k] = s
		if !sheets[k] {
			return fmt.Sprintf("the edit made a sheet, %s, it does not name", s.Name)
		}
	}
	get := func(k string, p [2]int) grid.Cell {
		if s := now[k]; s != nil {
			return s.Get(p[0], p[1])
		}
		return grid.Cell{}
	}
	inRects := func(k string, rg grid.Range) bool {
		for _, r := range rc.rects[k] {
			if !r.Intersect(rg).Empty() {
				return true
			}
		}
		return false
	}
	// may: the cells, where they stand now, whose value may change.
	may := map[string]map[[2]int]bool{}
	mark := func(k string, p [2]int) bool {
		if may[k] == nil {
			may[k] = map[[2]int]bool{}
		}
		if may[k][p] {
			return false
		}
		may[k][p] = true
		return true
	}
	// The arrays of the workbook now, by sheet.
	arrays := map[string][]grid.Range{}
	for k, s := range now {
		arrays[k] = append(append(arrays[k], s.Spills...), s.Arrays...)
	}
	for _, g := range rc.regions {
		if g.disturbed && !g.gone {
			mark(g.sheet, [2]int{g.rg.R1, g.rg.C1})
		}
	}
	// mayRead tells whether a range holds a cell whose value may change: in
	// the reach, marked, or filled by an array whose formula is marked.
	// anchored tells whether the formula of an array of the workbook now
	// (at its first cell) may give another result.
	anchored := func(k string, a [2]int) bool {
		return may[k][a] || inRects(k, grid.Range{R1: a[0], C1: a[1], R2: a[0], C2: a[1]})
	}
	mayRead := func(in readRange) bool {
		if inRects(in.sheet, in.rg) {
			return true
		}
		for p := range may[in.sheet] {
			if in.rg.Contains(p[0], p[1]) {
				return true
			}
		}
		for _, a := range arrays[in.sheet] {
			if !a.Intersect(in.rg).Empty() && anchored(in.sheet, [2]int{a.R1, a.C1}) {
				return true
			}
		}
		for _, g := range rc.regions {
			if g.sheet == in.sheet && !g.gone && !g.rg.Intersect(in.rg).Empty() && (g.disturbed || may[g.sheet][[2]int{g.rg.R1, g.rg.C1}]) {
				return true
			}
		}
		return false
	}

	// The cells outside the reach, in the sheets' order and by row.
	type tracked struct {
		k string
		p [2]int
		o origin
	}
	var cells []tracked
	for _, k := range order {
		var ps [][2]int
		for p := range rc.at[k] {
			ps = append(ps, p)
		}
		sort.Slice(ps, func(i, j int) bool {
			if ps[i][0] != ps[j][0] {
				return ps[i][0] < ps[j][0]
			}
			return ps[i][1] < ps[j][1]
		})
		for _, p := range ps {
			cells = append(cells, tracked{k, p, rc.at[k][p]})
		}
	}
	name := func(k string, p [2]int) string {
		if s := now[k]; s != nil {
			return s.Name + "!" + grid.CellName(p[0], p[1])
		}
		return k + "!" + grid.CellName(p[0], p[1])
	}

	// Formulas: rewritten in their references only, and only by an edit that
	// moves cells; a result unknown before or now may change.
	var formulas []tracked
	for _, t := range cells {
		b, a := rc.cell(t.o), get(t.k, t.p)
		if b.Formula == "" {
			continue
		}
		switch {
		case a.Formula == "":
			return fmt.Sprintf("%s: outside the edit's reach, it held %s, and holds %s now", name(t.k, t.p), describe(b), describe(a))
		case a.Formula != b.Formula && !rc.rewrites:
			return fmt.Sprintf("%s: outside the edit's reach, its formula =%s is rewritten as =%s by an edit that moves no cells", name(t.k, t.p), b.Formula, a.Formula)
		case a.Formula != b.Formula && masked(a.Formula) != masked(b.Formula):
			return fmt.Sprintf("%s: outside the edit's reach, its formula =%s is rewritten as =%s, other than in its references", name(t.k, t.p), b.Formula, a.Formula)
		case a.Formula != b.Formula || a.Pending || b.Pending:
			mark(t.k, t.p)
		}
		formulas = append(formulas, t)
	}
	// Recomputed dependents, to a fixpoint.
	names := usedNames(after.Names, rc.before.Names)
	for changed := true; changed; {
		changed = false
		for _, t := range formulas {
			if may[t.k][t.p] {
				continue
			}
			ins, ok := inputs(get(t.k, t.p).Formula, t.k, names)
			hit := !ok
			for _, in := range ins {
				hit = hit || mayRead(in)
			}
			if hit && mark(t.k, t.p) {
				changed = true
			}
		}
	}

	// Values.
	for _, t := range cells {
		b, a := rc.cell(t.o), get(t.k, t.p)
		if may[t.k][t.p] {
			continue
		}
		if g := rc.member[t.o]; g != nil && (g.disturbed || g.gone || may[g.sheet][[2]int{g.rg.R1, g.rg.C1}]) {
			continue // part of an array whose formula may give another result
		}
		if b.Formula == "" && a.Formula != "" || !sameValue(a, b) {
			why := ""
			if b.Formula != "" {
				why = ", and nothing it reads may have changed"
			}
			return fmt.Sprintf("%s: outside the edit's reach, it held %s, and holds %s now%s", name(t.k, t.p), describe(b), describe(a), why)
		}
	}
	// Cells empty before (or moved away) that hold something now.
	for _, k := range order {
		s := now[k]
		if s == nil {
			continue
		}
		u := s.Used()
		if u.R2 < 1 {
			continue
		}
		msg := ""
		s.Each(u, func(r, c int, cl *grid.Cell) {
			p := [2]int{r, c}
			if msg != "" {
				return
			}
			if _, ok := rc.at[k][p]; ok || inRects(k, grid.Range{R1: r, C1: c, R2: r, C2: c}) {
				return
			}
			for _, a := range arrays[k] {
				if a.Contains(r, c) && anchored(k, [2]int{a.R1, a.C1}) {
					return // filled by an array whose formula may change
				}
			}
			for _, g := range rc.regions {
				if g.sheet == k && !g.gone && g.rg.Contains(r, c) && (g.disturbed || may[k][[2]int{g.rg.R1, g.rg.C1}]) {
					return
				}
			}
			msg = fmt.Sprintf("%s: outside the edit's reach, it was empty, and holds %s now", name(k, p), describe(*cl))
		})
		if msg != "" {
			return msg
		}
	}
	return ""
}

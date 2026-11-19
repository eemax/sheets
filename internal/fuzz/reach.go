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
// held before, shown with the same number format, where it stands now, and
// an empty cell must stay empty.
//
// Two changes outside the reach are expected:
//
//   - shifted references: an edit that moves cells (move, rows, cols, a
//     script's moveTo, insert and delete) or renames a sheet rewrites the
//     formulas that read them, in their references only, each as the edit's
//     shift puts it (xform: what Excel does to a reference, edit by edit);
//     any other edit leaves every formula outside its reach as it was,
//     letter for letter; and the workbook's defined names the same way (no
//     edit names one: checkNames), none lost and none made, but those of a
//     sheet copied;
//   - recomputed dependents: a formula's value may change when it reads a
//     cell that may (in the reach, or a formula whose value may change, to a
//     fixpoint), when its references were rewritten, when it uses a defined
//     name the edit's shifts rewrote (rows deleted from its range), or when its
//     result was unknown before (pending: sheets computes a result it did not
//     know). A result known before and unknown after is a change: a formula
//     that reads nothing the edit touched keeps its result (the first version
//     excused it, which hid a save that drops results; fuzzer-long's audit,
//     f2.2). The cells
//     an array fills (a dynamic array's spill, a legacy array's range)
//     change with their formula, and with an edit that reaches into them or
//     cuts them.
//
// So a slip that writes a cell the edit does not name, moves one it should
// not, loses or rewrites a formula elsewhere (or shifts a reference other
// than the edit does), or shows a cell another way is a finding, as is a
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
	// they stand now; member the regions each of their other cells belonged
	// to (two arrays whose results meet share a cell: Excel shows #SPILL! at
	// one, and sheets lays them out in map order, so the cell is either's).
	regions []*region
	member  map[origin][]*region
	// xforms are what the edit does to references, in order: blocks moved,
	// lines inserted or deleted, sheets renamed (none: it rewrites none).
	xforms []xform
	// copies are the sheets the edit made as copies (lower case, as named
	// now): a sheet copied takes its names along.
	copies map[string]bool
}

// xform is one rewrite of references an edit makes, as Excel makes it:
//
//   - 'm', a block moved (move, a script's moveTo): a reference to a cell or a
//     range wholly inside src, on sheet, follows the block to sheet to, dr
//     rows and dc columns away; one wholly inside the cells it lands on
//     (which did not move themselves) becomes #REF!; the rest stay;
//   - 'l', lines (rows, or columns with cols) of sheet: n inserted before at,
//     each end of a reference at or past it moving n on; or (n == 0) at..end
//     deleted, a reference to a cell in them becoming #REF!, a range wholly
//     in them too, a range partly in them shrinking, and ends past them moving
//     back;
//   - 'r', sheet renamed to.
type xform struct {
	kind      byte
	sheet, to string // lower case
	src       grid.Range
	dr, dc    int
	cols      bool
	at, end   int
	n         int
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
	dynamic   bool       // a dynamic array's results (else a legacy array's range)
	was       grid.Range // its range before the edit
}

func newReach(wb *grid.Workbook) *reach {
	rc := &reach{before: wb, sheets: map[string]*grid.Sheet{}, at: map[string]map[[2]int]origin{}, rects: map[string][]grid.Range{}, member: map[origin][]*region{}, copies: map[string]bool{}}
	for _, s := range wb.Sheets {
		k := strings.ToLower(s.Name)
		rc.sheets[k] = s
		cells := map[[2]int]origin{}
		if u := s.Used(); u.R2 >= 1 {
			s.Each(u, func(r, c int, _ *grid.Cell) { cells[[2]int{r, c}] = origin{k, r, c} })
		}
		rc.at[k] = cells
		seen := map[[2]int]bool{}
		for i, a := range append(append([]grid.Range{}, s.Spills...), s.Arrays...) {
			if seen[[2]int{a.R1, a.C1}] {
				continue
			}
			seen[[2]int{a.R1, a.C1}] = true
			g := &region{sheet: k, rg: a, dynamic: i < len(s.Spills), was: a}
			rc.regions = append(rc.regions, g)
			for r := a.R1; r <= a.R2; r++ {
				for c := a.C1; c <= a.C2; c++ {
					if r != a.R1 || c != a.C1 {
						rc.member[origin{k, r, c}] = append(rc.member[origin{k, r, c}], g)
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

// resultsNote says that a cell was in the results of a dynamic array before
// the edit, other than its first cell, as spillNote does for the kinds and
// text invariants: ", in the results of the dynamic array at M2:M3" (a known
// bug's face reads it: spill-overwrites), "" when it was in none.
func (rc *reach) resultsNote(o origin) string {
	for _, g := range rc.member[o] {
		if g.dynamic {
			return ", in the results of the dynamic array at " + g.was.String()
		}
	}
	return ""
}

// shown are the texts a replace may read a value of the workbook before the
// edit as: as sheets shows it (its number format: a date as 2026-07-01, a
// percent as 12%), and a number as it is with no format too.
func (rc *reach) shown(c grid.Cell) []string {
	out := []string{grid.Display(&c, rc.before.FormatOf(&c), rc.before.Date1904)}
	if c.Kind == grid.Number {
		out = append(out, grid.ShowNumber(c.Num), grid.FormatNumber(c.Num))
	}
	return out
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
	rc.xforms = append(rc.xforms, xform{kind: 'l', sheet: k, cols: cols, at: at, end: to, n: n})
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
	rc.xforms = append(rc.xforms, xform{kind: 'r', sheet: from, to: to})
	rc.at[to], rc.rects[to] = rc.at[from], rc.rects[from]
	delete(rc.at, from)
	delete(rc.rects, from)
	if rc.copies[from] {
		rc.copies[to] = true
		delete(rc.copies, from)
	}
	for _, g := range rc.regions {
		if g.sheet == from {
			g.sheet = to
		}
	}
}

// copied says the edit made a sheet (lower case) as a copy of another: all of
// it is in the reach, and names scoped to it may be new.
func (rc *reach) copied(k string) {
	if rc == nil {
		return
	}
	rc.copies[k] = true
	rc.touch(k, whole)
}

// move follows a block of cells, src on sheet from (lower case), to sheet to,
// dr rows and dc columns away: what reads it follows it (see xform).
func (rc *reach) move(from, to string, src grid.Range, dr, dc int) {
	if rc == nil {
		return
	}
	rc.xforms = append(rc.xforms, xform{kind: 'm', sheet: from, to: to, src: src, dr: dr, dc: dc})
}

// refEnd is a cell a reference names, its row and column absolute or not.
type refEnd struct {
	r, c       int
	rAbs, cAbs bool
}

func (e refEnd) String() string {
	s := ""
	if e.cAbs {
		s = "$"
	}
	s += grid.ColName(e.c)
	if e.rAbs {
		s += "$"
	}
	return s + fmt.Sprint(e.r)
}

// fref is a reference in a formula's text: its sheet (lower case, "" for the
// formula's own), a cell or a range (a:b), or one an edit broke (#REF!); and
// where it stands in the text.
type fref struct {
	sheet      string
	a, b       refEnd
	rng, bad   bool
	start, end int
}

// wordByte tells whether a byte may be part of a name, a function or a sheet
// written without quotes (non-ASCII letters too).
func wordByte(c byte) bool {
	return c == '_' || c == '.' || c >= '0' && c <= '9' || c >= 'A' && c <= 'Z' || c >= 'a' && c <= 'z' || c >= 0x80
}

// cellAt reads a cell (B2, $B$2) at f[i:], and where it ends.
func cellAt(f string, i int) (refEnd, int, bool) {
	var e refEnd
	j := i
	if j < len(f) && f[j] == '$' {
		e.cAbs = true
		j++
	}
	k := j
	for k < len(f) && k-j < 4 && (f[k] >= 'A' && f[k] <= 'Z' || f[k] >= 'a' && f[k] <= 'z') {
		k++
	}
	if k == j || k-j > 3 {
		return e, i, false
	}
	col, ok := grid.ColIndex(strings.ToUpper(f[j:k]))
	if !ok {
		return e, i, false
	}
	if k < len(f) && f[k] == '$' {
		e.rAbs = true
		k++
	}
	d := k
	for d < len(f) && f[d] >= '0' && f[d] <= '9' {
		d++
	}
	if d == k {
		return e, i, false
	}
	row := 0
	for _, ch := range f[k:d] {
		row = row*10 + int(ch-'0')
		if row > grid.MaxRows {
			return e, i, false
		}
	}
	// A name or a function goes on (LOG10(, B2x): no cell.
	if row < 1 || d < len(f) && (wordByte(f[d]) || f[d] == '(' || f[d] == '!') {
		return e, i, false
	}
	e.r, e.c = row, col
	return e, d, true
}

// refsOf reads the references of a formula's text, outside its strings: a
// cell or a range, with the sheet it names, or #REF!.
func refsOf(f string) []fref {
	var out []fref
	for i := 0; i < len(f); {
		c := f[i]
		if c == '"' { // a string; "" inside it is a quote
			j := i + 1
			for j < len(f) {
				if f[j] == '"' {
					if j+1 < len(f) && f[j+1] == '"' {
						j += 2
						continue
					}
					break
				}
				j++
			}
			i = j + 1
			continue
		}
		if i > 0 && wordByte(f[i-1]) || !(c == '\'' || c == '$' || c == '#' || wordByte(c)) {
			i++
			continue
		}
		r := fref{start: i}
		j := i
		// The sheet: 'Bob''s Plan'! or Data!.
		if c == '\'' {
			k := i + 1
			for k < len(f) {
				if f[k] == '\'' {
					if k+1 < len(f) && f[k+1] == '\'' {
						k += 2
						continue
					}
					break
				}
				k++
			}
			if k+1 < len(f) && f[k+1] == '!' {
				r.sheet = strings.ToLower(strings.ReplaceAll(f[i+1:k], "''", "'"))
				j = k + 2
			}
		} else if wordByte(c) {
			k := i
			for k < len(f) && wordByte(f[k]) {
				k++
			}
			if k < len(f) && f[k] == '!' {
				r.sheet = strings.ToLower(f[i:k])
				j = k + 1
			}
		}
		if strings.HasPrefix(f[j:], "#REF!") {
			r.bad, r.end = true, j+5
			out = append(out, r)
			i = r.end
			continue
		}
		a, k, ok := cellAt(f, j)
		if !ok {
			// No reference here: past the word (a name, a function).
			for i++; i < len(f) && wordByte(f[i]); i++ {
			}
			continue
		}
		r.a, r.end = a, k
		if k < len(f) && f[k] == ':' {
			if b, k2, ok := cellAt(f, k+1); ok {
				r.b, r.rng, r.end = b, true, k2
			}
		}
		out = append(out, r)
		i = r.end
	}
	return out
}

// masked is a formula's text with each reference one token.
func masked(f string) string {
	var b strings.Builder
	at := 0
	for _, r := range refsOf(f) {
		b.WriteString(f[at:r.start])
		b.WriteString("§")
		at = r.end
	}
	b.WriteString(f[at:])
	return b.String()
}

// span is the range a reference names.
func (r fref) span() grid.Range {
	b := r.a
	if r.rng {
		b = r.b
	}
	return grid.Range{R1: min(r.a.r, b.r), C1: min(r.a.c, b.c), R2: max(r.a.r, b.r), C2: max(r.a.c, b.c)}
}

func within(r, outer grid.Range) bool {
	return r.R1 >= outer.R1 && r.R2 <= outer.R2 && r.C1 >= outer.C1 && r.C2 <= outer.C2
}

// apply is the reference after the rewrite; its sheet is never "".
func (x xform) apply(r fref) fref {
	if r.bad {
		return r
	}
	switch x.kind {
	case 'r':
		if r.sheet == x.sheet {
			r.sheet = x.to
		}
	case 'm':
		dst := grid.Range{R1: x.src.R1 + x.dr, C1: x.src.C1 + x.dc, R2: x.src.R2 + x.dr, C2: x.src.C2 + x.dc}
		switch rg := r.span(); {
		case r.sheet == x.sheet && within(rg, x.src):
			r.sheet = x.to
			r.a.r, r.a.c = r.a.r+x.dr, r.a.c+x.dc
			if r.rng {
				r.b.r, r.b.c = r.b.r+x.dr, r.b.c+x.dc
			}
		case r.sheet == x.to && within(rg, dst):
			r.bad = true
		}
	case 'l':
		if r.sheet != x.sheet {
			return r
		}
		line := func(e *refEnd) *int {
			if x.cols {
				return &e.c
			}
			return &e.r
		}
		ends := []*int{line(&r.a)}
		if r.rng {
			ends = append(ends, line(&r.b))
		}
		limit := grid.MaxRows
		if x.cols {
			limit = grid.MaxCols
		}
		if x.n > 0 {
			for _, p := range ends {
				if *p >= x.at {
					*p += x.n
				}
				if *p > limit {
					r.bad = true
				}
			}
			return r
		}
		in := func(p int) bool { return p >= x.at && p <= x.end }
		switch {
		case !r.rng && in(*ends[0]), r.rng && in(*ends[0]) && in(*ends[1]):
			r.bad = true
			return r
		case r.rng && in(*ends[0]):
			*ends[0] = x.end + 1
		case r.rng && in(*ends[1]):
			*ends[1] = x.at - 1
		}
		for _, p := range ends {
			if *p > x.end {
				*p -= x.end - x.at + 1
			}
		}
	}
	if r.a.r < 1 || r.a.c < 1 || r.rng && (r.b.r < 1 || r.b.c < 1) || r.a.r > grid.MaxRows || r.a.c > grid.MaxCols || r.b.r > grid.MaxRows || r.b.c > grid.MaxCols {
		r.bad = true
	}
	return r
}

// sameRef tells whether two references, their sheets resolved, name the same
// cells the same way (or both are #REF!).
func sameRef(a, b fref) bool {
	if a.bad || b.bad {
		return a.bad == b.bad
	}
	return a.sheet == b.sheet && a.rng == b.rng && a.a == b.a && (!a.rng || a.b == b.b)
}

// shifted checks a formula outside the reach against what the edit's
// rewrites (rc.xforms) make of it: its text, but for its references, as
// before; and each reference where the shifts put it (home and now are the
// formula's sheet before and now). It returns the formula the shifts give,
// and false when the formula saved is another.
func (rc *reach) shifted(before, after, home, now string, name func(k string) string) (string, bool) {
	if masked(before) != masked(after) {
		return "", false
	}
	bs, as := refsOf(before), refsOf(after)
	ok := len(bs) == len(as)
	var b strings.Builder
	at := 0
	for i, r := range bs {
		if r.sheet == "" {
			r.sheet = home
		}
		for _, x := range rc.xforms {
			r = x.apply(r)
		}
		b.WriteString(before[at:r.start])
		at = r.end
		if r.bad {
			b.WriteString("#REF!")
		} else {
			if r.sheet != now || bs[i].sheet != "" {
				b.WriteString(grid.QuoteSheet(name(r.sheet)) + "!")
			}
			b.WriteString(r.a.String())
			if r.rng {
				b.WriteString(":" + r.b.String())
			}
		}
		if ok {
			got := as[i]
			if got.sheet == "" {
				got.sheet = now
			}
			ok = sameRef(r, got)
		}
	}
	b.WriteString(before[at:])
	return b.String(), ok
}

// sameValue tells whether two cells hold the same value, both pending or
// neither.
func sameValue(a, b grid.Cell) bool {
	return a.Kind == b.Kind && a.Text == b.Text && a.Bool == b.Bool && a.Pending == b.Pending &&
		(a.Num == b.Num || math.IsNaN(a.Num) && math.IsNaN(b.Num))
}

// formatName is a number format as a finding names it.
func formatName(f grid.Format) string {
	if f.Code == "" || f.Code == "General" {
		return "General"
	}
	return f.Code
}

// sameFormat tells whether two number formats show a value the same way.
func sameFormat(a, b grid.Format) bool {
	return a.Kind == b.Kind && a.Decimals == b.Decimals && a.Symbol == b.Symbol && formatName(a) == formatName(b)
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

// checkNames judges the workbook's defined names, which no edit names: each
// name of the workbook before must be there after, under its scope (a sheet
// renamed takes it along), its references as the edit's shifts put them (as a
// formula's: rows deleted from Units shrink it), and letter for letter after
// an edit that moves no cells. No name may appear, but one scoped to a sheet
// the edit made as a copy (Excel copies a sheet's names with it). The first
// version took any name the edit rewrote as expected, and every formula using
// it as free to change: a set that rewrote Units, and the SUM(Units) it
// changed, were no finding (fuzzer-long's audit, f1.3). It returns the first
// difference, "" when there is none.
func (rc *reach) checkNames(after *grid.Workbook, display func(k string) string) string {
	key := func(n grid.DefinedName) [2]string {
		return [2]string{strings.ToLower(n.Name), strings.ToLower(n.Scope)}
	}
	// follow is where a scope (a sheet, lower case) stands after the edit.
	follow := func(k string) string {
		for _, x := range rc.xforms {
			if x.kind == 'r' && k != "" && k == x.sheet {
				k = x.to
			}
		}
		return k
	}
	now := map[[2]string]grid.DefinedName{}
	for _, n := range after.Names {
		now[key(n)] = n
	}
	scoped := func(n grid.DefinedName) string {
		if n.Scope == "" {
			return n.Name
		}
		return n.Name + " (of " + n.Scope + ")"
	}
	was := map[[2]string]bool{}
	for _, n := range rc.before.Names {
		k := key(n)
		k[1] = follow(k[1])
		was[k] = true
		a, ok := now[k]
		switch {
		case !ok:
			return fmt.Sprintf("the defined name %s, outside the edit's reach, is gone (it referred to =%s)", scoped(n), n.Ref)
		case a.Ref == n.Ref:
			if len(rc.xforms) == 0 {
				continue
			}
		case len(rc.xforms) == 0:
			return fmt.Sprintf("the defined name %s, outside the edit's reach, referred to =%s, and refers to =%s now, after an edit that moves no cells", scoped(n), n.Ref, a.Ref)
		}
		if want, ok := rc.shifted(n.Ref, a.Ref, strings.ToLower(n.Scope), k[1], display); !ok {
			if a.Ref == n.Ref {
				return fmt.Sprintf("the defined name %s, outside the edit's reach, refers to =%s as it did, and the edit's shift of references gives =%s", scoped(n), n.Ref, want)
			}
			return fmt.Sprintf("the defined name %s, outside the edit's reach, referred to =%s, and refers to =%s now, and the edit's shift of references gives =%s", scoped(n), n.Ref, a.Ref, want)
		}
	}
	for _, n := range after.Names {
		if k := key(n); !was[k] && !(k[1] != "" && rc.copies[k[1]]) {
			return fmt.Sprintf("the edit made a defined name, %s (=%s), it does not name", scoped(n), n.Ref)
		}
	}
	return ""
}

// usedName is a defined name, how a formula's text uses it, and whether the
// edit defined it again (moved: its range is another, or it is new; that
// it is the one the edit's shifts give, or a sheet copy's, checkNames says).
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
// reference, nor is N("p") a name). A sheet's name in single quotes is kept
// whole, double quotes and all ('Q1 "draft"'!B9 reads B9 of that sheet: its
// quotes once read as a string, which hid the reference).
func stringsOut(f string) string {
	var b strings.Builder
	str, name := false, false
	for _, r := range f {
		switch {
		case str:
			str = r != '"' // "" inside a string closes and opens it again: the same
			continue
		case name:
			name = r != '\'' // '' inside a name closes and opens it again: the same
		case r == '"':
			str = true
			continue
		case r == '\'':
			name = true
		}
		b.WriteRune(r)
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
	// anchored tells whether the formula of an array of the workbook now
	// (at its first cell) may give another result.
	anchored := func(k string, a [2]int) bool {
		return may[k][a] || inRects(k, grid.Range{R1: a[0], C1: a[1], R2: a[0], C2: a[1]})
	}
	// mayRead tells whether a range holds a cell whose value may change: in
	// the reach, marked, or filled by an array whose formula may change.
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

	// The name of a sheet (lower case) as it is now, or was.
	display := func(k string) string {
		if s := now[k]; s != nil {
			return s.Name
		}
		if s := rc.sheets[k]; s != nil {
			return s.Name
		}
		return k
	}

	// Formulas: rewritten only by an edit that moves cells, in their
	// references only, and each reference as the edit's shifts put it; a
	// result unknown before may change (sheets computes it), but not one
	// known before and unknown now (f2.2).
	var formulas []tracked
	for _, t := range cells {
		b, a := rc.cell(t.o), get(t.k, t.p)
		if b.Formula == "" {
			continue
		}
		switch {
		case a.Formula == "":
			return fmt.Sprintf("%s: outside the edit's reach, it held %s, and holds %s now", name(t.k, t.p), describe(b), describe(a))
		case a.Formula != b.Formula && len(rc.xforms) == 0:
			return fmt.Sprintf("%s: outside the edit's reach, its formula =%s is rewritten as =%s by an edit that moves no cells", name(t.k, t.p), b.Formula, a.Formula)
		case a.Formula != b.Formula && masked(a.Formula) != masked(b.Formula):
			return fmt.Sprintf("%s: outside the edit's reach, its formula =%s is rewritten as =%s, other than in its references", name(t.k, t.p), b.Formula, a.Formula)
		}
		if len(rc.xforms) > 0 {
			if want, ok := rc.shifted(b.Formula, a.Formula, t.o.sheet, t.k, display); !ok {
				if a.Formula == b.Formula {
					return fmt.Sprintf("%s: outside the edit's reach, its formula =%s is left as it was, and the edit's shift of references gives =%s", name(t.k, t.p), b.Formula, want)
				}
				return fmt.Sprintf("%s: outside the edit's reach, its formula =%s is rewritten as =%s, and the edit's shift of references gives =%s", name(t.k, t.p), b.Formula, a.Formula, want)
			}
		}
		if a.Formula != b.Formula || b.Pending {
			mark(t.k, t.p)
		}
		formulas = append(formulas, t)
	}
	// Defined names: kept, or rewritten as the edit's shifts put them.
	if msg := rc.checkNames(after, display); msg != "" {
		return msg
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

	// Values, and how each is shown (its number format: a date shown as
	// 46022 is a change a value does not show).
	for _, t := range cells {
		b, a := rc.cell(t.o), get(t.k, t.p)
		// Part of an array whose formula may give another result (of any
		// array it is in, when the results of two met).
		recomputed := may[t.k][t.p]
		for _, g := range rc.member[t.o] {
			recomputed = recomputed || g.disturbed || g.gone || may[g.sheet][[2]int{g.rg.R1, g.rg.C1}]
		}
		if !recomputed && (b.Formula == "" && a.Formula != "" || !sameValue(a, b)) {
			why := ""
			if b.Formula != "" {
				why = ", and nothing it reads may have changed"
			}
			return fmt.Sprintf("%s: outside the edit's reach, it held %s, and holds %s now%s%s", name(t.k, t.p), describe(b), describe(a), why, rc.resultsNote(t.o))
		}
		if fb, fa := rc.before.FormatOf(&b), after.FormatOf(&a); !sameFormat(fb, fa) {
			return fmt.Sprintf("%s: outside the edit's reach, it held %s shown as %s, and is shown as %s now%s", name(t.k, t.p), describe(b), formatName(fb), formatName(fa), rc.resultsNote(t.o))
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

package xlsx

import (
	"regexp"
	"sort"
	"strings"

	"github.com/eemax/sheets/internal/engine"
	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
)

// Results for a streaming edit are computed in a workbook of their own,
// holding only what the formulas being computed read: the values of the
// cells in their ranges, the formulas among those that are being computed
// too, and, for the rest, the results the file stores, which Excel
// computed. A sample of those (crossCheck per range) is computed again and
// checked against what the file stores, as the judge checks everything in
// a small workbook: a calculator that disagrees with Excel there is not
// trusted with the rest.

// crossCheck is how many stored results of a range are computed again.
const crossCheck = 16

// resolver turns names and tables into ranges, as they are now.
func (s *StreamEditor) resolver() *resolver {
	rs := &resolver{names: map[string][]nameDef{}}
	for _, n := range s.wb.Names {
		k := strings.ToLower(n.Name)
		rs.names[k] = append(rs.names[k], nameDef{strings.ToLower(n.Scope), s.rewrite(n.Ref, "")})
	}
	rs.tables = s.tablesNow("")
	return rs
}

func (s *StreamEditor) sheetOrder(name string) int {
	for i, sh := range s.wb.Sheets {
		if strings.EqualFold(sh.Name, name) {
			return i
		}
	}
	return len(s.wb.Sheets)
}

// kindsNow is what a cell written in this session holds.
func (s *StreamEditor) kindsNow(sheet string, row, col int) formula.Kinds {
	c := s.cellNow(sheet, row, col)
	if c.Kind == grid.Empty {
		return formula.KNumber
	}
	return formula.KindsOf(c)
}

func (s *StreamEditor) calcState() calcState {
	return calcState{formulas: s.formulasNow(), rs: s.resolver(), changed: s.changed, recompute: s.recompute,
		edited: len(s.ops) > 0, order: s.sheetOrder, calcFile: s.calcBook,
		typer: func(all map[cellKey]*formulaAt, computing map[*formulaAt]bool, rs *resolver) *typer {
			return newTyper(s.wb, s.changed, s.kindsNow, len(s.ops) > 0, all, computing, rs)
		}}
}

// unfollowed are functions whose inputs cannot be read off their
// arguments, or that answer for the workbook around them: the calculation
// workbook, which holds what formulas name, does not hold what they would
// look at.
var unfollowed = setOf(`INDIRECT OFFSET CELL INFO SHEET SHEETS`)

func setOf(s string) map[string]bool {
	m := map[string]bool{}
	for _, w := range strings.Fields(s) {
		m[w] = true
	}
	return m
}

// calcBook builds the calculation workbook for the formulas being
// computed (see above).
func (s *StreamEditor) calcBook(list []*formulaAt, rs *resolver, computing map[*formulaAt]bool) (engine.Book, calcPlan, error) {
	// The plan is made here: what is frozen (taken at its stored result)
	// is decided below, for what the formulas computed read.
	plan := calcPlan{frozen: map[*formulaAt]bool{}, veto: map[*formulaAt]string{}}
	at := map[cellKey]*formulaAt{}
	for _, f := range list {
		at[keyOf(f.sheet, f.row, f.col)] = f
	}
	type entry struct {
		v       grid.Cell
		formula string // to compute, as the calculator reads it
	}
	cells := map[cellKey]entry{}
	names := map[cellKey]string{}
	included := map[*formulaAt]bool{}
	// What the formulas included read: their ranges; the rows of each
	// column read so far (each is read once), and the rows holding an error
	// the calculator cannot hold.
	type sheetRange struct {
		sheet string // lower-cased
		rg    grid.Range
	}
	type colKey struct {
		sheet string // lower-cased
		col   int
	}
	reads := map[*formulaAt][]sheetRange{}
	covered := map[colKey]*rowSpans{}
	checked := map[colKey]int{}
	var queue []*formulaAt
	include := func(f *formulaAt) {
		if included[f] {
			return
		}
		included[f] = true
		delete(plan.frozen, f)
		k := keyOf(f.sheet, f.row, f.col)
		cells[k] = entry{formula: calcOf(f)}
		names[k] = f.sheet
		queue = append(queue, f)
	}
	var start []*formulaAt
	for f := range computing {
		start = append(start, f)
	}
	sort.Slice(start, func(i, j int) bool {
		a, b := start[i], start[j]
		if a.sheet != b.sheet {
			return a.sheet < b.sheet
		}
		if a.row != b.row {
			return a.row < b.row
		}
		return a.col < b.col
	})
	for _, f := range start {
		include(f)
	}
	for len(queue) > 0 {
		f := queue[0]
		queue = queue[1:]
		text := calcOf(f)
		for _, fn := range formula.CalledFunctions(text) {
			if unfollowed[fn] && engine.Has(fn) {
				plan.veto[f] = fn + " looks at cells its arguments do not name, which a streaming edit does not read"
			}
		}
		if plan.veto[f] != "" {
			continue
		}
		d := formula.DepsOf(text)
		refs, ok := rs.ranges(f.sheet, f.row, f.col, d, 0)
		if !ok {
			refs, ok = rs.tableRanges(f, d)
		}
		if !ok || d.Multi {
			plan.veto[f] = "its references cannot be read off it"
			continue
		}
		for _, ref := range refs {
			sheet := ref.Sheet
			if sheet == "" {
				sheet = f.sheet
			}
			sh := s.wb.Sheet(sheet)
			if sh == nil {
				continue
			}
			sheet = sh.Name
			used := s.usedNow(sheet)
			rg := ref.Range.Intersect(used)
			if rg.Empty() {
				continue
			}
			if hasFn(text, "SUBTOTAL", "AGGREGATE") && len(sh.HiddenRows) > 0 {
				plan.veto[f] = "it leaves out hidden rows, which the sheets calculator does not see in a file this large"
				break
			}
			lower := strings.ToLower(sheet)
			reads[f] = append(reads[f], sheetRange{lower, rg})
			// A sample of the stored results read, spread over each
			// column, is computed again and checked (see crossCheck).
			step := max(1, used.Rows()/crossCheck)
			for c := rg.C1; c <= rg.C2; c++ {
				ck := colKey{lower, c}
				if covered[ck] == nil {
					covered[ck] = &rowSpans{}
				}
				for _, part := range covered[ck].missing(rg.R1, rg.R2) {
					s.eachNow(sheet, grid.Range{R1: part[0], C1: c, R2: part[1], C2: c}, func(r, c int, v grid.Cell) {
						k := cellKey{lower, r, c}
						if v.Formula == "" {
							cells[k], names[k] = entry{v: v}, sheet
							return
						}
						g := at[k]
						if g == nil {
							return
						}
						if computing[g] || g.pending || !freezable(g.stored) {
							include(g)
							return
						}
						// A stored result to take; or one of the sample, of
						// the formulas the calculator can compute (it takes
						// the others at their results anyway).
						if checked[ck] < crossCheck && (r-used.R1)%step == 0 {
							if ok, _ := computable(calcOf(g), rs, g.sheet); ok {
								checked[ck]++
								include(g)
								return
							}
						}
						if !included[g] {
							plan.frozen[g] = true
							cells[k], names[k] = entry{v: g.stored}, sheet
						}
					})
					covered[ck].add(part[0], part[1])
				}
			}
		}
	}

	book := &cellsBook{cells: map[cellKey]cellEntry{}, sheets: map[string]string{}, used: s.usedNow, rs: rs}
	for _, sh := range s.wb.Sheets {
		book.sheets[strings.ToLower(sh.Name)] = sh.Name
	}
	for k, e := range cells {
		book.cells[k] = cellEntry{v: e.v, formula: e.formula}
	}
	plan.native = nativeAggregates(s.usedNow, func(j *judge, sheet string, r, c int) (grid.Cell, bool) {
		k := keyOf(sheet, r, c)
		e, ok := cells[k]
		switch {
		case !ok:
			return grid.Cell{}, true
		case e.formula == "":
			return e.v, true
		}
		if g := at[k]; g != nil && included[g] {
			res := j.result(g)
			return res.v, res.err == nil
		}
		return grid.Cell{}, false
	})
	return book, plan, nil
}

// cellsBook is a calculation workbook of the cells formulas read (see
// calcBook): values, and the formulas being computed.
type cellsBook struct {
	cells  map[cellKey]cellEntry
	sheets map[string]string // lower-cased name to name
	used   func(sheet string) grid.Range
	rs     *resolver
}

type cellEntry struct {
	v       grid.Cell
	formula string
}

func (b *cellsBook) Cell(sheet string, row, col int) (engine.Value, string) {
	e := b.cells[keyOf(sheet, row, col)]
	if e.formula != "" {
		return engine.Value{}, e.formula
	}
	return cellToCalc(e.v), ""
}

func (b *cellsBook) Extent(sheet string) (int, int, bool) {
	name, ok := b.sheets[strings.ToLower(sheet)]
	if !ok {
		return 0, 0, false
	}
	u := b.used(name)
	return max(u.R2, 0), max(u.C2, 0), true
}

func (b *cellsBook) Name(name, sheet string) (string, bool) {
	if def := b.rs.lookup(name, sheet); def != nil {
		return def.text, true
	}
	return "", false
}

// rowSpans are the rows of a column read so far, as sorted, separate
// spans.
type rowSpans struct{ spans [][2]int }

// missing lists the parts of rows r1..r2 not read yet.
func (sp *rowSpans) missing(r1, r2 int) [][2]int {
	var out [][2]int
	at := r1
	i := sort.Search(len(sp.spans), func(i int) bool { return sp.spans[i][1] >= r1 })
	for ; i < len(sp.spans) && at <= r2; i++ {
		x := sp.spans[i]
		if x[0] > r2 {
			break
		}
		if x[0] > at {
			out = append(out, [2]int{at, x[0] - 1})
		}
		at = max(at, x[1]+1)
	}
	if at <= r2 {
		out = append(out, [2]int{at, r2})
	}
	return out
}

// add marks rows r1..r2 read.
func (sp *rowSpans) add(r1, r2 int) {
	i := sort.Search(len(sp.spans), func(i int) bool { return sp.spans[i][1] >= r1-1 })
	j := i
	for j < len(sp.spans) && sp.spans[j][0] <= r2+1 {
		r1, r2 = min(r1, sp.spans[j][0]), max(r2, sp.spans[j][1])
		j++
	}
	sp.spans = append(sp.spans[:i], append([][2]int{{r1, r2}}, sp.spans[j:]...)...)
}

// nameOf spells a defined name as the workbook does.
func nameOf(wb *grid.Workbook, lower string) string {
	for _, n := range wb.Names {
		if strings.ToLower(n.Name) == lower {
			return n.Name
		}
	}
	return lower
}

func hasFn(f string, fns ...string) bool {
	for _, c := range formula.CalledFunctions(f) {
		for _, fn := range fns {
			if c == fn {
				return true
			}
		}
	}
	return false
}

// eachNow calls fn with every non-empty cell of a range as it is now.
func (s *StreamEditor) eachNow(sheet string, rg grid.Range, fn func(r, c int, v grid.Cell)) {
	sh := s.wb.Sheet(sheet)
	edits := s.edits[strings.ToLower(sheet)]
	for r := rg.R1; r <= rg.R2; r++ {
		orig, ok := s.rowThen(sheet, r)
		for c := rg.C1; c <= rg.C2; c++ {
			if e, ok := edits[[2]int{r, c}]; ok {
				if v := s.cellNow(sheet, r, c); !v.IsEmpty() || e.v.Formula != "" {
					fn(r, c, v)
				}
				continue
			}
			if !ok || sh == nil {
				continue
			}
			if c > sh.RowLen(orig) {
				if len(edits) == 0 {
					break
				}
				continue
			}
			x := sh.At(orig, c)
			if x == nil || x.IsEmpty() {
				continue
			}
			v := *x
			if v.Formula != "" && len(s.ops) > 0 {
				v.Formula = s.rewrite(v.Formula, sh.Name)
			}
			fn(r, c, v)
		}
	}
}

// aggregate is a call of an aggregate function on references alone:
// =SUM(I2:I200001), =MAX(B2:B9,D2:D9).
var aggregate = regexp.MustCompile(`^(?i)(SUM|COUNT|COUNTA|AVERAGE|MIN|MAX)\((.+)\)$`)

// aggregateOf reads a formula that is an aggregate of references alone.
func aggregateOf(text string) (string, []formula.RefUse, bool) {
	text = dropSpaces(text)
	m := aggregate.FindStringSubmatch(text)
	if m == nil {
		return "", nil, false
	}
	uses := formula.RefUses(m[2])
	// Nothing but the references, separated by commas.
	var parts []string
	for _, u := range uses {
		parts = append(parts, strings.ReplaceAll(u.Text, "$", ""))
	}
	if len(uses) == 0 || !strings.EqualFold(strings.Join(parts, ","), strings.ReplaceAll(m[2], "$", "")) {
		return "", nil, false
	}
	return strings.ToUpper(m[1]), uses, true
}

// dropSpaces drops the spaces of a formula outside quoted sheet names and
// strings.
func dropSpaces(s string) string {
	if !strings.Contains(s, " ") {
		return s
	}
	var b strings.Builder
	var quote byte
	for i := 0; i < len(s); i++ {
		c := s[i]
		switch {
		case quote != 0:
			if c == quote {
				quote = 0
			}
		case c == '\'' || c == '"':
			quote = c
		case c == ' ':
			continue
		}
		b.WriteByte(c)
	}
	return b.String()
}

// nativeAggregates computes aggregates of references straight from the
// values the calculation workbook holds: for the totals of large tables,
// the calculator walks every cell of its range, slowly. value gives a
// cell's value, false for a formula whose result is unknown.
//
// Formulas are computed top to bottom, so ranges that start alike and grow
// (running totals, =SUM($B$2:B9)) are each scanned on from where the one
// before stopped.
func nativeAggregates(usedNow func(sheet string) grid.Range, value func(j *judge, sheet string, r, c int) (grid.Cell, bool)) func(j *judge, f *formulaAt) (grid.Cell, bool) {
	type scanKey struct {
		sheet      string
		r1, c1, c2 int
	}
	scans := map[scanKey]*aggScan{}
	return func(j *judge, f *formulaAt) (grid.Cell, bool) {
		fn, uses, ok := aggregateOf(calcOf(f))
		if !ok {
			return grid.Cell{}, false
		}
		var total aggScan
		for _, u := range uses {
			sheet := u.Sheet
			if sheet == "" {
				sheet = f.sheet
			}
			rg := u.Range.Intersect(usedNow(sheet))
			if rg.Empty() {
				continue
			}
			k := scanKey{strings.ToLower(sheet), rg.R1, rg.C1, rg.C2}
			sc := scans[k]
			if sc == nil || sc.r2 > rg.R2 {
				sc = &aggScan{r2: rg.R1 - 1}
			}
			if !sc.extend(rg, func(r, c int) (grid.Cell, bool) { return value(j, sheet, r, c) }) {
				delete(scans, k) // read in part: of no use to the next
				return grid.Cell{}, false
			}
			scans[k] = sc
			total.add(sc, fn)
		}
		return total.result(fn)
	}
}

// aggScan is what an aggregate has read of a range, rows R1 to r2.
type aggScan struct {
	r2        int
	sum       float64
	low, high float64
	n, filled int        // numbers; cells holding anything
	err       *grid.Cell // the first error, in reading order
}

// extend reads on to the last row of rg; false when a value is unknown.
func (sc *aggScan) extend(rg grid.Range, value func(r, c int) (grid.Cell, bool)) bool {
	for r := sc.r2 + 1; r <= rg.R2; r++ {
		for c := rg.C1; c <= rg.C2; c++ {
			v, ok := value(r, c)
			if !ok {
				return false
			}
			switch v.Kind {
			case grid.Error:
				if sc.err == nil {
					e := v
					sc.err = &e
				}
				sc.filled++
			case grid.Number:
				if sc.n == 0 || v.Num < sc.low {
					sc.low = v.Num
				}
				if sc.n == 0 || v.Num > sc.high {
					sc.high = v.Num
				}
				sc.sum += v.Num
				sc.n++
				sc.filled++
			case grid.Text, grid.Bool:
				sc.filled++
			}
		}
		sc.r2 = r
	}
	return true
}

// add takes in another range's scan, for an aggregate of several.
func (sc *aggScan) add(o *aggScan, fn string) {
	if sc.err == nil && o.err != nil {
		sc.err = o.err
	}
	if o.n > 0 {
		if sc.n == 0 || o.low < sc.low {
			sc.low = o.low
		}
		if sc.n == 0 || o.high > sc.high {
			sc.high = o.high
		}
	}
	sc.sum += o.sum
	sc.n += o.n
	sc.filled += o.filled
}

// result is the aggregate's value, as Excel gives it: errors pass COUNT
// by and count for COUNTA; the others give the first.
func (sc *aggScan) result(fn string) (grid.Cell, bool) {
	if sc.err != nil && fn != "COUNT" && fn != "COUNTA" {
		return *sc.err, true
	}
	switch fn {
	case "SUM":
		return grid.Cell{Kind: grid.Number, Num: sc.sum}, true
	case "COUNT":
		return grid.Cell{Kind: grid.Number, Num: float64(sc.n)}, true
	case "COUNTA":
		return grid.Cell{Kind: grid.Number, Num: float64(sc.filled)}, true
	case "AVERAGE":
		if sc.n == 0 {
			return grid.Cell{Kind: grid.Error, Text: "#DIV/0!"}, true
		}
		return grid.Cell{Kind: grid.Number, Num: sc.sum / float64(sc.n)}, true
	case "MIN":
		return grid.Cell{Kind: grid.Number, Num: sc.low}, true
	case "MAX":
		return grid.Cell{Kind: grid.Number, Num: sc.high}, true
	}
	return grid.Cell{}, false
}

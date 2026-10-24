package xlsx

import (
	"fmt"
	"regexp"
	"sort"
	"strings"

	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
	"github.com/xuri/excelize/v2"
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
// workbook does not hold what they would look at.
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
func (s *StreamEditor) calcBook(list []*formulaAt, rs *resolver, computing map[*formulaAt]bool) (*excelize.File, calcPlan, func(), error) {
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
	reads := map[*formulaAt][]cellKey{} // the non-empty cells each formula included reads
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
	scanned := map[string]*formulaAt{} // ranges read, by the formula that read them first
	for len(queue) > 0 {
		f := queue[0]
		queue = queue[1:]
		text := calcOf(f)
		for _, fn := range formula.CalledFunctions(text) {
			if unfollowed[fn] {
				plan.veto[f] = fn + " looks at cells or the workbook in a way the sheets calculator does not follow in a file this large"
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
			rg := ref.Range.Intersect(s.usedNow(sheet))
			if rg.Empty() {
				continue
			}
			if hasFn(text, "SUBTOTAL", "AGGREGATE") && len(sh.HiddenRows) > 0 {
				plan.veto[f] = "it leaves out hidden rows, which the sheets calculator does not see in a file this large"
				break
			}
			key := sheet + "!" + rg.String()
			if prev, ok := scanned[key]; ok {
				reads[f] = append(reads[f], reads[prev]...)
				continue
			}
			scanned[key] = f
			step := max(1, rg.Rows()/crossCheck)
			checked := 0
			s.eachNow(sheet, rg, func(r, c int, v grid.Cell) {
				k := keyOf(sheet, r, c)
				reads[f] = append(reads[f], k)
				if v.Formula == "" {
					if v.Kind == grid.Error {
						if ef, ok := errorFormula[v.Text]; ok {
							cells[k], names[k] = entry{formula: ef}, sheet
						} else {
							plan.veto[f] = fmt.Sprintf("it reads %s, an error the sheets calculator cannot hold", grid.FormatRef(sheet, grid.Range{R1: r, C1: c, R2: r, C2: c}))
						}
						return
					}
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
				// A stored result to take; a sample is computed again and
				// checked, of the formulas the calculator can compute (it
				// takes the others at their results anyway).
				if checked < crossCheck && (r-rg.R1)%step == 0 {
					if ok, _ := computable(calcOf(g), rs, g.sheet); ok {
						checked++
						include(g)
						return
					}
				}
				if !included[g] {
					plan.frozen[g] = true
					cells[k], names[k] = entry{v: g.stored}, sheet
				}
			})
		}
	}

	// The calculator gets what the formulas it computes read; aggregates
	// of references are computed without it (see nativeAggregates), from
	// cells it need not hold.
	inBook := map[cellKey]bool{}
	var needs []*formulaAt
	for g := range included {
		if _, _, ok := aggregateOf(calcOf(g)); !ok {
			needs = append(needs, g)
		}
	}
	excel := map[*formulaAt]bool{}
	for len(needs) > 0 {
		g := needs[len(needs)-1]
		needs = needs[:len(needs)-1]
		if excel[g] {
			continue
		}
		excel[g] = true
		inBook[keyOf(g.sheet, g.row, g.col)] = true
		for _, k := range reads[g] {
			inBook[k] = true
			if h := at[k]; h != nil && included[h] && !excel[h] {
				needs = append(needs, h)
			}
		}
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
			if res.err != nil {
				return grid.Cell{}, false
			}
			v := res.v
			if v.Kind != grid.Error {
				// What a reader of the cell sees: the result, typed.
				if typed, ok := formula.Typed(res.raw, j.ty.kindsOf(g)); ok {
					v = typed
				}
			}
			return v, true
		}
		// An error written as a formula the calculator answers with.
		for text, ef := range errorFormula {
			if ef == e.formula {
				return grid.Cell{Kind: grid.Error, Text: text}, true
			}
		}
		return grid.Cell{}, false
	})
	f := excelize.NewFile()
	release := func() { f.Close() }
	for i, sh := range s.wb.Sheets {
		var err error
		if i == 0 {
			err = f.SetSheetName(f.GetSheetName(0), sh.Name)
		} else {
			_, err = f.NewSheet(sh.Name)
		}
		if err != nil {
			release()
			return nil, plan, nil, err
		}
	}
	for name, defs := range rs.names {
		for _, d := range defs {
			scope := ""
			if d.scope != "" {
				if sh := s.wb.Sheet(d.scope); sh != nil {
					scope = sh.Name
				}
			}
			// Names a calculator cannot hold (pointing nowhere) are left
			// out; formulas using them fail and are not trusted.
			_ = f.SetDefinedName(&excelize.DefinedName{Name: nameOf(s.wb, name), RefersTo: d.text, Scope: scope})
		}
	}
	bySheet := map[string][][2]int{}
	for k := range inBook {
		if _, ok := cells[k]; ok {
			bySheet[names[k]] = append(bySheet[names[k]], [2]int{k.row, k.col})
		}
	}
	for sheet, pos := range bySheet {
		sort.Slice(pos, func(i, j int) bool { return pos[i][0] < pos[j][0] || pos[i][0] == pos[j][0] && pos[i][1] < pos[j][1] })
		sw, err := f.NewStreamWriter(sheet)
		if err != nil {
			release()
			return nil, plan, nil, err
		}
		for i := 0; i < len(pos); {
			row := pos[i][0]
			j := i
			for j < len(pos) && pos[j][0] == row {
				j++
			}
			first := pos[i][1]
			vals := make([]interface{}, pos[j-1][1]-first+1)
			for _, p := range pos[i:j] {
				e := cells[keyOf(sheet, p[0], p[1])]
				var v interface{}
				switch {
				case e.formula != "":
					v = excelize.Cell{Formula: formula.ToFile(e.formula)}
				case e.v.Kind == grid.Number:
					v = e.v.Num
				case e.v.Kind == grid.Bool:
					v = e.v.Bool
				case e.v.Kind == grid.Text:
					v = e.v.Text
				}
				vals[p[1]-first] = v
			}
			if err := sw.SetRow(grid.CellName(row, first), vals); err != nil {
				release()
				return nil, plan, nil, err
			}
			i = j
		}
		if err := sw.Flush(); err != nil {
			release()
			return nil, plan, nil, err
		}
	}
	return f, plan, release, nil
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
	text = strings.ReplaceAll(text, " ", "")
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

// nativeAggregates computes aggregates of references straight from the
// values the calculation workbook holds: for the totals of large tables,
// the calculator walks every cell of its range, slowly. value gives a
// cell's value, false for a formula whose result is unknown.
func nativeAggregates(usedNow func(sheet string) grid.Range, value func(j *judge, sheet string, r, c int) (grid.Cell, bool)) func(j *judge, f *formulaAt) (grid.Cell, bool) {
	return func(j *judge, f *formulaAt) (grid.Cell, bool) {
		fn, uses, ok := aggregateOf(calcOf(f))
		if !ok {
			return grid.Cell{}, false
		}
		var sum, low, high float64
		n, filled := 0, 0
		for _, u := range uses {
			sheet := u.Sheet
			if sheet == "" {
				sheet = f.sheet
			}
			rg := u.Range.Intersect(usedNow(sheet))
			for r := rg.R1; r <= rg.R2 && !rg.Empty(); r++ {
				for c := rg.C1; c <= rg.C2; c++ {
					v, ok := value(j, sheet, r, c)
					if !ok {
						return grid.Cell{}, false
					}
					switch v.Kind {
					case grid.Error:
						return v, true
					case grid.Number:
						if n == 0 || v.Num < low {
							low = v.Num
						}
						if n == 0 || v.Num > high {
							high = v.Num
						}
						sum += v.Num
						n++
						filled++
					case grid.Text, grid.Bool:
						filled++
					}
				}
			}
		}
		switch fn {
		case "SUM":
			return grid.Cell{Kind: grid.Number, Num: sum}, true
		case "COUNT":
			return grid.Cell{Kind: grid.Number, Num: float64(n)}, true
		case "COUNTA":
			return grid.Cell{Kind: grid.Number, Num: float64(filled)}, true
		case "AVERAGE":
			if n == 0 {
				return grid.Cell{Kind: grid.Error, Text: "#DIV/0!"}, true
			}
			return grid.Cell{Kind: grid.Number, Num: sum / float64(n)}, true
		case "MIN":
			return grid.Cell{Kind: grid.Number, Num: low}, true
		case "MAX":
			return grid.Cell{Kind: grid.Number, Num: high}, true
		}
		return grid.Cell{}, false
	}
}

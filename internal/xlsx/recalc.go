package xlsx

import (
	"bytes"
	"fmt"
	"sort"
	"strconv"
	"strings"
	"time"

	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
	"github.com/xuri/excelize/v2"
)

// Formula results kept in the file. Excel recalculates when it opens a
// file sheets saved, but everything else that reads it (pandas, previews,
// other tools and agents) shows the results the file stores. So when
// saving, sheets recomputes every formula an edit can have changed and
// stores the result; a formula it cannot compute loses its old result
// rather than keep a wrong one. A file sheets saves never shows a stale
// number: each result is current, or absent until Excel computes it.

// recalcBudget bounds the time spent computing results when saving; the
// formulas left over lose their stored results.
var recalcBudget = 15 * time.Second

// cellKey is a cell; sheet is lower-cased, as Excel compares sheet names.
type cellKey struct {
	sheet    string
	row, col int
}

func keyOf(sheet string, row, col int) cellKey { return cellKey{strings.ToLower(sheet), row, col} }

// formulaAt is a formula where it stands now.
type formulaAt struct {
	sheet    string // as the workbook spells it
	row, col int
	text     string
	pending  bool      // its stored result was missing or stale when opened
	stored   grid.Cell // its stored result when opened, if any (Empty if none)
	// calc is the text the calculator computes: text with table
	// references and LET translated (see formula/translate.go), when
	// they could be.
	calc string
}

// noteFresh notes, before a cell is first written in this session,
// whether it is empty.
func (e *Editor) noteFresh(sheet string, row, col int) {
	k := keyOf(sheet, row, col)
	if _, seen := e.changed[k]; seen {
		return
	}
	if e.wasEmpty == nil {
		e.wasEmpty = map[cellKey]bool{}
	}
	cell := grid.CellName(row, col)
	v, _ := e.f.GetCellValue(sheet, cell, excelize.Options{RawCellValue: true})
	f, _ := e.f.GetCellFormula(sheet, cell)
	e.wasEmpty[k] = v == "" && f == ""
}

// touch records that a cell's content changed in this session.
func (e *Editor) touch(sheet string, row, col int) {
	if e.changed == nil {
		e.changed = map[cellKey]string{}
	}
	e.changed[keyOf(sheet, row, col)] = sheet
}

// mustRecompute records a formula an edit of rows, columns or sheets may
// have given a new result.
func (e *Editor) mustRecompute(sheet string, row, col int) {
	if e.recompute == nil {
		e.recompute = map[cellKey]bool{}
	}
	e.recompute[keyOf(sheet, row, col)] = true
}

// formulasNow lists the workbook's formulas as they stand: those read at
// opening (or moved by a structural edit), with the cells this session
// wrote read back from excelize.
func (e *Editor) formulasNow() (map[cellKey]*formulaAt, error) {
	all := map[cellKey]*formulaAt{}
	if e.lines != nil {
		for i := range e.lines {
			l := &e.lines[i]
			all[keyOf(l.sheet, l.row, l.col)] = l
		}
	} else {
		for _, s := range e.wb.Sheets {
			s.Each(s.Used(), func(r, c int, cell *grid.Cell) {
				if cell.Formula != "" {
					all[keyOf(s.Name, r, c)] = &formulaAt{sheet: s.Name, row: r, col: c, text: cell.Formula, pending: cell.Pending, stored: valueOf(*cell)}
				}
			})
		}
	}
	for k, sheet := range e.changed {
		if i, err := e.f.GetSheetIndex(sheet); err != nil || i < 0 {
			delete(all, k)
			continue
		}
		text, err := e.f.GetCellFormula(sheet, grid.CellName(k.row, k.col))
		if err != nil {
			return nil, err
		}
		if text == "" {
			delete(all, k)
		} else {
			all[k] = &formulaAt{sheet: sheet, row: k.row, col: k.col, text: formula.FromFile(text)}
		}
	}
	return all, nil
}

// depIndex finds the formulas that read a cell.
type depIndex struct {
	ids     []*formulaAt
	cells   map[cellKey][]int
	columns map[cellKey]map[grid.Range][]int // narrow ranges, by sheet and column (row 0)
	wide    map[string]map[grid.Range][]int  // ranges over many columns, by sheet
	always  []int                            // volatile: read anything
}

func (x *depIndex) add(id int, sheet string, rg grid.Range) {
	sheet = strings.ToLower(sheet)
	switch {
	case rg.R1 == rg.R2 && rg.C1 == rg.C2:
		k := cellKey{sheet, rg.R1, rg.C1}
		x.cells[k] = append(x.cells[k], id)
	case rg.Cols() <= 16:
		for c := rg.C1; c <= rg.C2; c++ {
			k := cellKey{sheet, 0, c}
			if x.columns[k] == nil {
				x.columns[k] = map[grid.Range][]int{}
			}
			x.columns[k][rg] = append(x.columns[k][rg], id)
		}
	default:
		if x.wide[sheet] == nil {
			x.wide[sheet] = map[grid.Range][]int{}
		}
		x.wide[sheet][rg] = append(x.wide[sheet][rg], id)
	}
}

// readers calls fn with every formula that reads cell k.
func (x *depIndex) readers(k cellKey, fn func(id int)) {
	for _, id := range x.cells[k] {
		fn(id)
	}
	for rg, ids := range x.columns[cellKey{k.sheet, 0, k.col}] {
		if rg.Contains(k.row, k.col) {
			for _, id := range ids {
				fn(id)
			}
		}
	}
	for rg, ids := range x.wide[k.sheet] {
		if rg.Contains(k.row, k.col) {
			for _, id := range ids {
				fn(id)
			}
		}
	}
}

// resolver turns the names and tables a formula uses into ranges.
type resolver struct {
	names  map[string][]nameDef // lower-cased name → definitions
	tables []tableNow
	// renames gives a formula the names of table columns renamed after
	// their headers this session, which it takes when saved (nil: none).
	renames func(text, sheet string, row, col int) string
}

type nameDef struct {
	scope string // lower-cased sheet, "" for the workbook
	text  string
}

func (e *Editor) resolver() *resolver {
	rs := &resolver{names: map[string][]nameDef{}}
	for _, n := range e.f.GetDefinedName() {
		if strings.HasPrefix(n.Name, "_xlnm.") {
			continue
		}
		scope := ""
		if n.Scope != "" && !strings.EqualFold(n.Scope, "Workbook") {
			scope = strings.ToLower(n.Scope)
		}
		k := strings.ToLower(n.Name)
		rs.names[k] = append(rs.names[k], nameDef{scope, strings.TrimPrefix(n.RefersTo, "=")})
	}
	for _, t := range e.patches.tables {
		if tn, ok := t.now(e.patches.rangeRewrites); ok {
			// Column names as saved: after their headers.
			if x, ok := e.patches.tableXML[strings.ToLower(t.part)]; ok {
				tn.columns = columnNames(x)
			}
			rs.tables = append(rs.tables, tn)
		}
	}
	if len(e.patches.columnRenames) > 0 {
		rs.renames = func(text, sheet string, row, col int) string {
			return e.patches.renameColumns(text, func(cr *columnRename) bool {
				return strings.EqualFold(cr.sheet, sheet) && cr.ref.Contains(row, col)
			})
		}
	}
	return rs
}

// wbResolver is the resolver of a workbook as read.
func wbResolver(wb *grid.Workbook) *resolver {
	rs := &resolver{names: map[string][]nameDef{}}
	for _, n := range wb.Names {
		k := strings.ToLower(n.Name)
		rs.names[k] = append(rs.names[k], nameDef{strings.ToLower(n.Scope), n.Ref})
	}
	for _, s := range wb.Sheets {
		for _, t := range s.Tables {
			rs.tables = append(rs.tables, tableNow{name: t.Name, sheet: s.Name, ref: t.Range, headers: 1, totals: t.Totals, columns: t.Columns})
		}
	}
	return rs
}

// ranges lists what a formula on sheet reads; ok is false when that cannot
// be told (a volatile formula, a table that is gone).
func (rs *resolver) ranges(sheet string, row, col int, d formula.Deps, depth int) (out []formula.Ref, ok bool) {
	if d.Volatile || depth > 8 {
		return nil, false
	}
	for _, r := range d.Refs {
		if r.Sheet == "" {
			r.Sheet = sheet
		}
		out = append(out, r)
	}
	for _, n := range d.Names {
		defs := rs.names[strings.ToLower(n)]
		var def *nameDef
		for i := range defs {
			if defs[i].scope == strings.ToLower(sheet) || defs[i].scope == "" && def == nil {
				def = &defs[i]
			}
		}
		if def == nil {
			if t := rs.table(n); t != nil {
				out = append(out, formula.Ref{Sheet: t.sheet, Range: t.ref}) // a table's name: its cells
			}
			continue // else not a name: #NAME?, which no edit changes
		}
		nd := formula.DepsOf(def.text)
		for _, r := range nd.Refs {
			if r.Sheet == "" {
				return nil, false // a name relative to where it is used
			}
		}
		sub, ok := rs.ranges(sheet, row, col, nd, depth+1)
		if !ok {
			return nil, false
		}
		out = append(out, sub...)
	}
	if len(d.Tables) > 0 {
		// Structured references are read once translated into the cells
		// they stand for (see calcText); one that is not cannot be told.
		return nil, false
	}
	return out, true
}

// tableRanges is what a formula's table references can read: the whole of
// each table; ok is false when that cannot be told either.
func (rs *resolver) tableRanges(f *formulaAt, d formula.Deps) ([]formula.Ref, bool) {
	if len(d.Tables) == 0 || d.Volatile || d.External {
		return nil, false
	}
	plain := d
	plain.Tables = nil
	out, ok := rs.ranges(f.sheet, f.row, f.col, plain, 0)
	if !ok {
		return nil, false
	}
	for _, name := range d.Tables {
		var t *tableNow
		if name == "" {
			for i := range rs.tables {
				if strings.EqualFold(rs.tables[i].sheet, f.sheet) && rs.tables[i].ref.Contains(f.row, f.col) {
					t = &rs.tables[i]
				}
			}
		} else {
			t = rs.table(name)
		}
		if t == nil {
			return nil, false
		}
		out = append(out, formula.Ref{Sheet: t.sheet, Range: t.ref})
	}
	return out, true
}

// calcState is what working out results needs from an editing session.
type calcState struct {
	formulas  map[cellKey]*formulaAt // the workbook's formulas as they stand
	rs        *resolver
	changed   map[cellKey]string // cells written, by sheet name
	recompute map[cellKey]bool   // formulas an edit of rows, columns or sheets reached
	edited    bool               // rows, columns or sheets changed
	// restore stores again the results of formulas not recomputed: an
	// edit of rows took them off with their formulas (see relines).
	restore bool
	order   func(sheet string) int
	// calcFile is the workbook to compute in, set up by the plan for the
	// formulas being computed (see calcplan.go), and what releases it.
	calcFile func(list []*formulaAt, rs *resolver, computing map[*formulaAt]bool) (*excelize.File, calcPlan, func(), error)
	typer    func(all map[cellKey]*formulaAt, computing map[*formulaAt]bool, rs *resolver) *typer
}

// calcResult is what working out results gives.
type calcResult struct {
	results  map[string]map[string]grid.Cell // by lower-cased sheet and cell; Empty removes a result
	dropped  map[cellKey]string              // why a formula has no stored result, or what its error means
	warnings []string
	touched  map[string][][2]int // cells whose values changed, by lower-cased sheet
}

// recalc works out the results to store: every formula a change in this
// session can have reached, and the ones whose results were missing or
// stale when the file was opened, as far as time allows.
func (e *Editor) recalc() error {
	all, err := e.formulasNow()
	if err != nil {
		return err
	}
	st := calcState{formulas: all, rs: e.resolver(), changed: e.changed, recompute: e.recompute,
		edited: e.structural || e.sheetsChanged, restore: e.lines != nil, order: e.sheetIndex, calcFile: e.calcFile,
		typer: func(all map[cellKey]*formulaAt, computing map[*formulaAt]bool, rs *resolver) *typer {
			return e.newTyper(all, computing, rs)
		}}
	res, err := computeResults(st)
	if err != nil {
		return err
	}
	e.patches.results, e.dropped, e.patches.touched = res.results, res.dropped, res.touched
	e.warnings = append(e.warnings, res.warnings...)
	return nil
}

// calcFile is the workbook the calculator works in: the one being edited,
// or a copy when the plan changes formulas for it, or when sheets were
// renamed or deleted: formulas follow those only as the file is written,
// and the copy is made from it.
func (e *Editor) calcFile(list []*formulaAt, rs *resolver, computing map[*formulaAt]bool) (*excelize.File, calcPlan, func(), error) {
	f, plan, release, err := e.calcIn(list, rs, computing)
	if err != nil {
		return nil, plan, nil, err
	}
	// Aggregates of references are computed from the cells' values, as
	// Excel does: the calculator's SUM, say, skips errors where Excel
	// gives the error.
	at := map[cellKey]*formulaAt{}
	for _, x := range list {
		at[keyOf(x.sheet, x.row, x.col)] = x
	}
	plan.native = nativeAggregates(e.used, func(j *judge, sheet string, r, c int) (grid.Cell, bool) {
		if g := at[keyOf(sheet, r, c)]; g != nil {
			res := j.result(g)
			if res.err != nil {
				return grid.Cell{}, false
			}
			v := res.v
			if v.Kind != grid.Error {
				if typed, ok := formula.Typed(res.raw, j.ty.kindsOf(g)); ok {
					v = typed
				}
			}
			return v, true
		}
		return fileValue(f, sheet, r, c), true
	})
	return f, plan, release, nil
}

// fileValue reads a value cell as excelize holds it.
func fileValue(f *excelize.File, sheet string, row, col int) grid.Cell {
	cell := grid.CellName(row, col)
	v, err := f.GetCellValue(sheet, cell, excelize.Options{RawCellValue: true})
	if err != nil || v == "" {
		return grid.Cell{}
	}
	switch typ, _ := f.GetCellType(sheet, cell); typ {
	case excelize.CellTypeBool:
		return grid.Cell{Kind: grid.Bool, Bool: v == "1" || strings.EqualFold(v, "TRUE")}
	case excelize.CellTypeError:
		return grid.Cell{Kind: grid.Error, Text: v}
	case excelize.CellTypeNumber, excelize.CellTypeUnset, excelize.CellTypeDate:
		if n, err := strconv.ParseFloat(v, 64); err == nil {
			return grid.Cell{Kind: grid.Number, Num: n}
		}
	}
	return grid.Cell{Kind: grid.Text, Text: v}
}

// calcIn is the workbook the calculator works in (see calcFile).
func (e *Editor) calcIn(list []*formulaAt, rs *resolver, computing map[*formulaAt]bool) (*excelize.File, calcPlan, func(), error) {
	plan := newPlan(list, rs, computing)
	if !plan.needed() && len(e.patches.sheetEdits) == 0 {
		return e.f, plan, func() {}, nil
	}
	var data []byte
	if len(e.patches.sheetEdits) > 0 {
		pt := *e.patches
		pt.stamp, pt.results, pt.touched = false, nil, nil
		b, err := serialize(e.f, &pt)
		if err != nil {
			return nil, plan, nil, err
		}
		data = b
	} else {
		b, err := e.f.WriteToBuffer()
		if err != nil {
			return nil, plan, nil, err
		}
		data = b.Bytes()
	}
	f, err := excelize.OpenReader(bytes.NewReader(data))
	if err != nil {
		return nil, plan, nil, err
	}
	if err := plan.apply(f, list); err != nil {
		f.Close()
		return nil, plan, nil, err
	}
	return f, plan, func() { f.Close() }, nil
}

// computeResults works out the results to store for a session's state.
func computeResults(st calcState) (calcResult, error) {
	out := calcResult{results: map[string]map[string]grid.Cell{}, dropped: map[cellKey]string{}}
	all, rs := st.formulas, st.rs

	// The formulas to recompute: those written or moved by an edit, and
	// everything that reads what changed, transitively.
	type depsOf struct {
		refs []formula.Ref
		ok   bool
	}
	// Dependencies are read off the text the calculator computes, where
	// table references stand translated into their cells.
	for _, f := range all {
		f.calc = rs.calcText(f)
	}
	depCache := map[*formulaAt]depsOf{}
	deps := func(f *formulaAt) ([]formula.Ref, bool) {
		d, seen := depCache[f]
		if !seen {
			d.refs, d.ok = rs.ranges(f.sheet, f.row, f.col, formula.DepsOf(calcOf(f)), 0)
			depCache[f] = d
		}
		return d.refs, d.ok
	}
	must := map[*formulaAt]bool{}
	if len(st.changed) > 0 || len(st.recompute) > 0 || st.edited {
		x := &depIndex{cells: map[cellKey][]int{}, columns: map[cellKey]map[grid.Range][]int{}, wide: map[string]map[grid.Range][]int{}}
		for _, f := range all {
			id := len(x.ids)
			x.ids = append(x.ids, f)
			refs, ok := deps(f)
			if !ok {
				// A table reference not translated: it may read anything
				// in its table, nothing more.
				refs, ok = rs.tableRanges(f, formula.DepsOf(calcOf(f)))
			}
			if !ok {
				x.always = append(x.always, id)
				continue
			}
			for _, r := range refs {
				x.add(id, r.Sheet, r.Range)
			}
		}
		var queue []cellKey
		reach := func(f *formulaAt) {
			if !must[f] {
				must[f] = true
				queue = append(queue, keyOf(f.sheet, f.row, f.col))
			}
		}
		for k := range st.changed {
			if f := all[k]; f != nil {
				reach(f)
			} else {
				queue = append(queue, k)
			}
		}
		for k := range st.recompute {
			if f := all[k]; f != nil {
				reach(f)
			}
		}
		for _, id := range x.always {
			reach(x.ids[id])
		}
		for len(queue) > 0 {
			k := queue[0]
			queue = queue[1:]
			x.readers(k, func(id int) { reach(x.ids[id]) })
		}
	}

	// Compute in a stable order: the ones that must change first, then
	// the missing ones while time lasts.
	var first, then []*formulaAt
	for _, f := range all {
		switch {
		case must[f]:
			first = append(first, f)
		case f.pending:
			then = append(then, f)
		}
	}
	order := func(fs []*formulaAt) {
		sort.Slice(fs, func(i, j int) bool {
			a, b := fs[i], fs[j]
			if a.sheet != b.sheet {
				return st.order(a.sheet) < st.order(b.sheet)
			}
			if a.row != b.row {
				return a.row < b.row
			}
			return a.col < b.col
		})
	}
	order(first)
	order(then)

	// The judge decides which of the calculator's answers hold (see
	// judge.go); those that do are typed by what each formula can produce.
	computing := map[*formulaAt]bool{}
	for _, f := range first {
		computing[f] = true
	}
	for _, f := range then {
		computing[f] = true
	}
	var list []*formulaAt
	for _, f := range all {
		list = append(list, f)
	}
	calcIn, plan, release, err := st.calcFile(list, rs, computing)
	if err != nil {
		return out, err
	}
	defer release()
	ty := st.typer(all, computing, rs)
	deadline := time.Now().Add(recalcBudget)
	j := newJudge(calcIn, newCycleGraph(list, deps), rs, ty, computing, plan, func(f *formulaAt) (grid.Cell, bool) {
		return f.stored, f.stored.Kind != grid.Empty && !f.pending
	}, deadline)

	var unsupported []string
	left := 0
	store := func(f *formulaAt, v grid.Cell) {
		s := strings.ToLower(f.sheet)
		if out.results[s] == nil {
			out.results[s] = map[string]grid.Cell{}
		}
		out.results[s][grid.CellName(f.row, f.col)] = v
	}
	var linked []string
	for i, f := range append(first, then...) {
		ref := grid.FormatRef(f.sheet, grid.Range{R1: f.row, C1: f.col, R2: f.row, C2: f.col})
		if formula.DepsOf(f.text).External {
			// Only Excel can update a link to another workbook; the
			// stored result is all there is until it does.
			if i < len(first) {
				linked = append(linked, ref)
			}
			continue
		}
		key := keyOf(f.sheet, f.row, f.col)
		_, written := st.changed[key]
		ok, why := j.holds(f)
		var v grid.Cell
		if ok {
			r := j.result(f)
			v = r.v
			if r.note != "" {
				out.dropped[key] = r.note
			}
			if v.Kind != grid.Error {
				if v, ok = formula.Typed(r.raw, ty.kindsOf(f)); !ok {
					why = "the sheets calculator cannot tell whether " + strconv.Quote(r.raw) + " is a number or text here"
				}
			}
		}
		if !ok {
			// No stored result rather than a wrong or stale one.
			store(f, grid.Cell{})
			out.dropped[key] = why
			switch {
			case i >= len(first) || written:
			case strings.HasPrefix(why, "not computed in the time"):
				left++
			case !strings.HasPrefix(why, "circular reference"):
				unsupported = append(unsupported, fmt.Sprintf("%s (%s)", ref, why))
			}
			continue
		}
		store(f, v)
	}
	if cycle := j.members; len(cycle) > 0 {
		var list []string
		for _, f := range cycle[:min(len(cycle), 8)] {
			list = append(list, grid.FormatRef(f.sheet, grid.Range{R1: f.row, C1: f.col, R2: f.row, C2: f.col}))
		}
		if len(cycle) > 8 {
			list = append(list, fmt.Sprintf("and %d more", len(cycle)-8))
		}
		if len(cycle) == 1 {
			out.warnings = append(out.warnings, fmt.Sprintf("circular reference: %s refers to itself, so Excel cannot compute it; it has no stored result", list[0]))
		} else {
			out.warnings = append(out.warnings, fmt.Sprintf("circular reference: %d formulas refer to each other, so Excel cannot compute them; they have no stored result: %s",
				len(cycle), strings.Join(list, ", ")))
		}
	}
	if st.restore {
		for _, f := range all {
			if computing[f] || f.pending || f.stored.Kind == grid.Empty {
				continue
			}
			if s := strings.ToLower(f.sheet); out.results[s] == nil || out.results[s][grid.CellName(f.row, f.col)].Kind == grid.Empty {
				if _, done := out.results[s][grid.CellName(f.row, f.col)]; !done {
					store(f, f.stored)
				}
			}
		}
	}
	// Cells whose values changed, for charts drawing them.
	out.touched = map[string][][2]int{}
	for k := range st.changed {
		out.touched[k.sheet] = append(out.touched[k.sheet], [2]int{k.row, k.col})
	}
	for sheet, cells := range out.results {
		for cell := range cells {
			r, c, _ := grid.ParseCell(cell)
			out.touched[sheet] = append(out.touched[sheet], [2]int{r, c})
		}
	}
	if len(unsupported) > 0 {
		list := unsupported
		if len(list) > 5 {
			list = append(list[:5:5], fmt.Sprintf("and %d more", len(unsupported)-5))
		}
		out.warnings = append(out.warnings, fmt.Sprintf("%s that this change reaches %s no stored result now; Excel computes %s on open: %s",
			countOf(len(unsupported), "formula"), map[bool]string{true: "has", false: "have"}[len(unsupported) == 1],
			map[bool]string{true: "it", false: "them"}[len(unsupported) == 1], strings.Join(list, "; ")))
	}
	if len(linked) > 0 {
		out.warnings = append(out.warnings, fmt.Sprintf("%s that this change reaches %s to another workbook; %s stored %s may be out of date until Excel updates the links: %s",
			countOf(len(linked), "formula"), map[bool]string{true: "links", false: "link"}[len(linked) == 1],
			map[bool]string{true: "its", false: "their"}[len(linked) == 1], map[bool]string{true: "result", false: "results"}[len(linked) == 1],
			strings.Join(linked[:min(len(linked), 5)], ", ")))
	}
	if left > 0 {
		out.warnings = append(out.warnings, fmt.Sprintf("stopped computing formula results after %s; %s that this change reaches %s no stored result until Excel computes them",
			recalcBudget, countOf(left, "formula"), map[bool]string{true: "has", false: "have"}[left == 1]))
	}
	return out, nil
}

// valueOf is the stored result of a formula cell as read (Empty if none).
func valueOf(c grid.Cell) grid.Cell {
	if c.Pending {
		return grid.Cell{}
	}
	return grid.Cell{Kind: c.Kind, Bool: c.Bool, Num: c.Num, Text: c.Text}
}

func (e *Editor) sheetIndex(name string) int {
	i, _ := e.f.GetSheetIndex(name)
	return i
}

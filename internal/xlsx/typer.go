package xlsx

import (
	"sort"
	"strconv"
	"strings"

	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
	"github.com/xuri/excelize/v2"
)

// typer tells what the cells of a workbook hold, for typing the
// calculator's results (see formula.ResultKinds): as read, or as an
// editing session left them. Where cells moved in a session, it answers
// for the whole sheet: wider, never wrong.
type typer struct {
	wb       *grid.Workbook                                 // as read
	changed  map[cellKey]string                             // cells written since, by sheet name
	now      func(sheet string, row, col int) formula.Kinds // what such a cell holds
	formulas map[cellKey]*formulaAt
	recalc   map[*formulaAt]bool // results being recomputed: stored ones do not hold
	rs       *resolver
	memo     map[*formulaAt]formula.Kinds
	busy     map[*formulaAt]bool
	columns  map[cellKey]formula.Kinds // value cells by sheet and column (row 0)
	inColumn map[cellKey][]*formulaAt  // formulas by sheet and column (row 0)
	sheets   map[string]formula.Kinds  // value cells by sheet
	moved    map[string]bool           // sheets whose cells moved
	names    map[string]formula.Kinds  // looked-up names
	colMemo  map[cellKey]formula.Kinds
}

func (e *Editor) newTyper(formulas map[cellKey]*formulaAt, recalc map[*formulaAt]bool, rs *resolver) *typer {
	return newTyper(e.wb, e.changed, e.current, e.lines != nil, formulas, recalc, rs)
}

// newTyper types the formulas of wb; recalc are those whose stored results
// do not hold. Cells changed since reading are answered by now; moved says
// rows or columns moved, so positions read no longer hold.
func newTyper(wb *grid.Workbook, changed map[cellKey]string, now func(string, int, int) formula.Kinds, moved bool,
	formulas map[cellKey]*formulaAt, recalc map[*formulaAt]bool, rs *resolver) *typer {
	t := &typer{wb: wb, changed: changed, now: now, formulas: formulas, recalc: recalc, rs: rs,
		memo: map[*formulaAt]formula.Kinds{}, busy: map[*formulaAt]bool{},
		columns: map[cellKey]formula.Kinds{}, inColumn: map[cellKey][]*formulaAt{},
		sheets: map[string]formula.Kinds{}, moved: map[string]bool{}, names: map[string]formula.Kinds{}, colMemo: map[cellKey]formula.Kinds{}}
	if moved {
		for _, s := range wb.Sheets {
			t.moved[strings.ToLower(s.Name)] = true
		}
	}
	for _, s := range wb.Sheets {
		sheet := strings.ToLower(s.Name)
		s.Each(s.Used(), func(r, c int, cell *grid.Cell) {
			if cell.Formula != "" {
				return
			}
			k := formula.KindsOf(*cell)
			t.columns[cellKey{sheet, 0, c}] |= k
			t.sheets[sheet] |= k
		})
	}
	for k, sheet := range changed {
		if _, ok := formulas[k]; ok {
			continue
		}
		v := now(sheet, k.row, k.col)
		t.columns[cellKey{k.sheet, 0, k.col}] |= v
		t.sheets[k.sheet] |= v
	}
	for k, f := range formulas {
		col := cellKey{k.sheet, 0, k.col}
		t.inColumn[col] = append(t.inColumn[col], f)
	}
	for _, fs := range t.inColumn {
		sort.Slice(fs, func(i, j int) bool { return fs[i].row < fs[j].row })
	}
	return t
}

// sheetKinds is what the formulas of a sheet can produce, all together.
func (t *typer) sheetKinds(sheet string) formula.Kinds {
	key := cellKey{sheet, -1, -1}
	if k, ok := t.colMemo[key]; ok {
		return k
	}
	t.colMemo[key] = formula.KAny // while working it out
	var k formula.Kinds
	for col, fs := range t.inColumn {
		if col.sheet == sheet && len(fs) > 0 {
			k |= t.columnKinds(col)
		}
	}
	t.colMemo[key] = k
	return k
}

// columnKinds is what the formulas of a column can produce, all together:
// for ranges taking in so many that typing each would cost too much.
func (t *typer) columnKinds(col cellKey) formula.Kinds {
	if k, ok := t.colMemo[col]; ok {
		return k
	}
	t.colMemo[col] = formula.KAny // while working it out
	var k formula.Kinds
	for _, f := range t.inColumn[col] {
		k |= t.kindsOf(f)
	}
	t.colMemo[col] = k
	return k
}

// current reads a value cell as excelize holds it now.
func (e *Editor) current(sheet string, row, col int) formula.Kinds {
	cell := grid.CellName(row, col)
	typ, _ := e.f.GetCellType(sheet, cell)
	v, _ := e.f.GetCellValue(sheet, cell, excelize.Options{RawCellValue: true})
	switch {
	case v == "":
		return formula.KNumber
	case typ == excelize.CellTypeBool:
		return formula.KBool
	case typ == excelize.CellTypeError:
		return 0
	case typ == excelize.CellTypeNumber || typ == excelize.CellTypeUnset:
		if _, err := strconv.ParseFloat(v, 64); err == nil {
			return formula.KNumber
		}
	}
	return formula.TextKinds(v)
}

// kindsOf is what a formula can produce: its stored result's type when
// that still holds, else what its text can produce.
func (t *typer) kindsOf(f *formulaAt) formula.Kinds {
	if k, ok := t.memo[f]; ok {
		return k
	}
	if t.busy[f] {
		return formula.KAny // a cycle
	}
	t.busy[f] = true
	defer delete(t.busy, f)
	k := formula.Kinds(0)
	key := keyOf(f.sheet, f.row, f.col)
	if _, written := t.changed[key]; !written && !t.recalc[f] && !f.pending && !t.moved[key.sheet] {
		if s := t.wb.Sheet(f.sheet); s != nil {
			if c := s.At(f.row, f.col); c != nil && c.Formula != "" {
				k = formula.KindsOf(grid.Cell{Kind: c.Kind, Bool: c.Bool, Num: c.Num, Text: c.Text})
			}
		}
	}
	if k == 0 {
		k = formula.ResultKinds(calcOf(f), &sheetTyper{t, f.sheet})
	}
	t.memo[f] = k
	return k
}

// sheetTyper is the typer seen from a formula on sheet.
type sheetTyper struct {
	t     *typer
	sheet string
}

func (s *sheetTyper) RangeKinds(sheet string, rg grid.Range) formula.Kinds {
	t := s.t
	if sheet == "" {
		sheet = s.sheet
	}
	lower := strings.ToLower(sheet)
	if rg.R1 == rg.R2 && rg.C1 == rg.C2 && !t.moved[lower] {
		k := cellKey{lower, rg.R1, rg.C1}
		if f := t.formulas[k]; f != nil {
			return t.kindsOf(f)
		}
		if name, ok := t.changed[k]; ok {
			return t.now(name, rg.R1, rg.C1)
		}
		if ws := t.wb.Sheet(sheet); ws != nil {
			if c := ws.At(rg.R1, rg.C1); c != nil {
				return formula.KindsOf(*c)
			}
		}
		return formula.KNumber // empty: 0
	}
	k := formula.KNumber // ranges hold empty cells, read as 0
	if t.moved[lower] || rg.Cols() > 64 {
		return k | t.sheets[lower] | t.sheetKinds(lower)
	}
	for c := rg.C1; c <= rg.C2; c++ {
		col := cellKey{lower, 0, c}
		k |= t.columns[col]
		fs := t.inColumn[col]
		lo := sort.Search(len(fs), func(i int) bool { return fs[i].row >= rg.R1 })
		hi := sort.Search(len(fs), func(i int) bool { return fs[i].row > rg.R2 })
		if hi-lo > 256 {
			k |= t.columnKinds(col)
			continue
		}
		for _, f := range fs[lo:hi] {
			k |= t.kindsOf(f)
		}
	}
	return k
}

func (s *sheetTyper) NameKinds(name string) (formula.Kinds, bool) {
	t := s.t
	defs := t.rs.names[strings.ToLower(name)]
	var def *nameDef
	for i := range defs {
		if defs[i].scope == strings.ToLower(s.sheet) || defs[i].scope == "" && def == nil {
			def = &defs[i]
		}
	}
	if def == nil {
		return 0, false
	}
	key := strings.ToLower(s.sheet) + "!" + strings.ToLower(name)
	if k, ok := t.names[key]; ok {
		return k, true
	}
	t.names[key] = formula.KAny // while working it out: a name defined through itself
	k := formula.ResultKinds(def.text, s)
	t.names[key] = k
	return k, true
}

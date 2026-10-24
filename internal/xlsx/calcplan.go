package xlsx

import (
	"sort"
	"strings"

	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
	"github.com/xuri/excelize/v2"
)

// A calculation plan sets the calculator up for a workbook's formulas,
// in a copy of it: never in the file being saved.
//
//   - Formulas with table references or LET are computed in translated
//     form (formula.ExpandTables, formula.InlineLet), which it can
//     evaluate.
//   - A formula it cannot compute even so, whose stored result holds (it
//     is not being recomputed), is taken at that result: what reads it
//     can then be computed, from Excel's own answer.
type calcPlan struct {
	translated []*formulaAt
	frozen     map[*formulaAt]bool
	// veto are formulas the calculator is not to be trusted with here,
	// and why.
	veto map[*formulaAt]string
	// native computes a formula without the calculator, where it can
	// (see StreamEditor.native).
	native func(j *judge, f *formulaAt) (grid.Cell, bool)
}

// calcText is the text the calculator is to compute for f.
func (rs *resolver) calcText(f *formulaAt) string {
	text := f.text
	if strings.Contains(strings.ToUpper(text), "LET(") {
		if t, ok := formula.InlineLet(text); ok {
			text = t
		}
	}
	if rs.usesTables(text) {
		if rs.renames != nil {
			text = rs.renames(text, f.sheet, f.row, f.col)
		}
		tables := make([]formula.Table, len(rs.tables))
		for i, t := range rs.tables {
			tables[i] = t.formulaTable()
		}
		if t, ok := formula.ExpandTables(text, f.sheet, f.row, f.col, tables); ok {
			text = t
		}
	}
	return text
}

// usesTables tells whether a formula names an Excel table.
func (rs *resolver) usesTables(text string) bool {
	if len(rs.tables) == 0 {
		return false
	}
	if !strings.Contains(text, "[") {
		// A table named alone (=SUM(Sales)): its name must be in the text.
		named := false
		upper := strings.ToUpper(text)
		for _, t := range rs.tables {
			named = named || strings.Contains(upper, strings.ToUpper(t.name))
		}
		if !named {
			return false
		}
	}
	d := formula.DepsOf(text)
	if len(d.Tables) > 0 {
		return true
	}
	for _, n := range d.Names {
		if rs.table(n) != nil {
			return true
		}
	}
	return false
}

func (rs *resolver) table(name string) *tableNow {
	for i := range rs.tables {
		if strings.EqualFold(rs.tables[i].name, name) {
			return &rs.tables[i]
		}
	}
	return nil
}

// newPlan works out the plan for fs; fresh are the formulas being
// recomputed, whose stored results do not hold.
func newPlan(fs []*formulaAt, rs *resolver, fresh map[*formulaAt]bool) calcPlan {
	p := calcPlan{frozen: map[*formulaAt]bool{}}
	for _, f := range fs {
		if f.calc == "" {
			f.calc = rs.calcText(f)
		}
		if f.calc != f.text {
			p.translated = append(p.translated, f)
		}
		if !fresh[f] && !f.pending && freezable(f.stored) {
			if ok, _ := computable(f.calc, rs, f.sheet); !ok {
				p.frozen[f] = true
			}
		}
	}
	return p
}

func (p calcPlan) needed() bool { return len(p.translated) > 0 || len(p.frozen) > 0 }

// freezable tells whether a stored result can stand in for its formula in
// the calculator: a value, or an error it can produce.
func freezable(v grid.Cell) bool {
	switch v.Kind {
	case grid.Number, grid.Text, grid.Bool:
		return true
	case grid.Error:
		_, ok := errorFormula[v.Text]
		return ok
	}
	return false
}

// errorFormula gives each error a formula the calculator answers with it.
var errorFormula = map[string]string{"#N/A": "NA()", "#DIV/0!": "1/0", "#NUM!": "SQRT(-1)"}

// apply sets the calculator up in f, a copy of the workbook. Every formula
// on a sheet the plan changes is rewritten as a formula of its own first,
// so that no formula filled down from a changed cell (a shared formula)
// changes with it.
func (p calcPlan) apply(f *excelize.File, fs []*formulaAt) error {
	sheets := map[string]bool{}
	for _, x := range p.translated {
		sheets[strings.ToLower(x.sheet)] = true
	}
	for x := range p.frozen {
		sheets[strings.ToLower(x.sheet)] = true
	}
	var on []*formulaAt
	for _, x := range fs {
		if sheets[strings.ToLower(x.sheet)] {
			on = append(on, x)
		}
	}
	sort.Slice(on, func(i, j int) bool {
		a, b := on[i], on[j]
		if a.sheet != b.sheet {
			return a.sheet < b.sheet
		}
		if a.row != b.row {
			return a.row > b.row
		}
		return a.col > b.col
	})
	for _, x := range on {
		cell := grid.CellName(x.row, x.col)
		if p.frozen[x] {
			var err error
			switch v := x.stored; v.Kind {
			case grid.Number:
				err = f.SetCellFloat(x.sheet, cell, v.Num, -1, 64)
			case grid.Text:
				err = setText(f, x.sheet, cell, v.Text)
			case grid.Bool:
				err = f.SetCellBool(x.sheet, cell, v.Bool)
			case grid.Error:
				err = f.SetCellFormula(x.sheet, cell, errorFormula[v.Text])
			}
			if err != nil {
				return err
			}
			continue
		}
		text := x.text
		if x.calc != "" {
			text = x.calc
		}
		stored := formula.ToFile(text)
		var opts []excelize.FormulaOpts
		if formula.Spills(stored) {
			t, ref := excelize.STCellFormulaTypeArray, cell
			opts = append(opts, excelize.FormulaOpts{Type: &t, Ref: &ref})
		}
		if err := f.SetCellFormula(x.sheet, cell, stored, opts...); err != nil {
			return err
		}
	}
	return nil
}

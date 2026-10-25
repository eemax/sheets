package xlsx

import (
	"strings"

	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
)

// A calculation plan sets the engine up for a workbook's formulas:
//
//   - Formulas with table references are computed in translated form
//     (formula.ExpandTables), with the table's cells.
//   - A formula it cannot compute even so, whose stored result holds (it
//     is not being recomputed), is taken at that result: what reads it
//     can then be computed, from Excel's own answer.
type calcPlan struct {
	frozen map[*formulaAt]bool
	// veto are formulas the engine is not to be trusted with here, and
	// why.
	veto map[*formulaAt]string
	// native computes aggregates of long ranges incrementally (see
	// nativeAggregates).
	native func(j *judge, f *formulaAt) (grid.Cell, bool)
}

// calcText is the text the engine is to compute for f.
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
		if !fresh[f] && !f.pending && freezable(f.stored) {
			if ok, _ := computable(f.calc, rs, f.sheet); !ok {
				p.frozen[f] = true
			}
		}
	}
	return p
}

// freezable tells whether a stored result can stand in for its formula: a
// value, or an error.
func freezable(v grid.Cell) bool {
	return v.Kind != grid.Empty
}

package xlsx

import (
	"bytes"
	"errors"
	"strconv"
	"strings"
	"sync"
	"time"

	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
	"github.com/xuri/excelize/v2"
)

// calc evaluates formulas with excelize's engine. The workbook is opened
// lazily, only when a pending value is actually needed. As when storing
// results, an answer counts only when the judge says it holds (see
// judge.go), and is typed by what the formula can produce, where that
// tells.
type calc struct {
	data []byte
	wb   *grid.Workbook
	once sync.Once
	f    *excelize.File
	err  error

	at map[cellKey]*formulaAt
	j  *judge
}

func (c *calc) Resolve(sheet string, row, col int) (grid.Cell, error) {
	c.once.Do(func() {
		c.f, c.err = excelize.OpenReader(bytes.NewReader(c.data))
		c.data = nil
		if c.err == nil {
			c.analyse()
		}
	})
	if c.err != nil {
		return grid.Cell{}, c.err
	}
	f := c.at[keyOf(sheet, row, col)]
	if f == nil {
		v, _, err := evaluate(c.f, sheet, grid.CellName(row, col))
		return v, err
	}
	if ok, _ := c.j.holds(f); !ok {
		r := c.j.evals[f]
		if r != nil && errors.Is(r.err, errUncertain) {
			return r.v, nil // shown as computed, which marks it for what it is
		}
		return grid.Cell{}, errUnsupported
	}
	r := c.j.result(f)
	if r.v.Kind == grid.Error {
		return r.v, nil
	}
	if typed, ok := formula.Typed(r.raw, c.j.ty.kindsOf(f)); ok {
		return typed, nil
	}
	return r.v, nil // as the calculator gave it: shown, not stored
}

// analyse prepares the judge over the workbook's formulas.
func (c *calc) analyse() {
	c.at = map[cellKey]*formulaAt{}
	pending := map[*formulaAt]bool{}
	var list []*formulaAt
	if c.wb == nil {
		c.wb = &grid.Workbook{}
	}
	for _, s := range c.wb.Sheets {
		s.Each(s.Used(), func(r, col int, cell *grid.Cell) {
			if cell.Formula == "" {
				return
			}
			f := &formulaAt{sheet: s.Name, row: r, col: col, text: cell.Formula, pending: cell.Pending, stored: valueOf(*cell)}
			c.at[keyOf(s.Name, r, col)] = f
			list = append(list, f)
			if cell.Pending {
				pending[f] = true
			}
		})
	}
	rs := wbResolver(c.wb)
	for _, f := range list {
		f.calc = rs.calcText(f)
	}
	g := newCycleGraph(list, func(f *formulaAt) ([]formula.Ref, bool) {
		return rs.ranges(f.sheet, f.row, f.col, formula.DepsOf(calcOf(f)), 0)
	})
	// This copy of the workbook is the calculator's own: the plan changes
	// it in place.
	plan := newPlan(list, rs, pending)
	if plan.needed() {
		if err := plan.apply(c.f, list); err != nil {
			c.err = err
			return
		}
	}
	ty := newTyper(c.wb, nil, nil, false, c.at, pending, rs)
	c.j = newJudge(c.f, g, rs, ty, pending, plan.frozen, func(f *formulaAt) (grid.Cell, bool) {
		return f.stored, f.stored.Kind != grid.Empty && !f.pending
	}, time.Now().Add(24*time.Hour))
}

// errUnsupported marks formulas the calculator cannot evaluate although
// Excel can (functions it does not implement, such as LET or FILTER).
var errUnsupported = errors.New("not supported by the sheets calculator")

// errUncertain marks a result the calculator may get wrong: #VALUE! from
// text it cannot read as a number, where Excel may (1,200 is 1200 in an
// English locale). It is shown, flagged, but never stored in a file.
var errUncertain = errors.New("result may differ in Excel")

// evaluate calculates one cell and types the result. Errors in the formula
// itself (#DIV/0!, #VALUE!) come back as Error cells. When the calculator,
// not the formula, is the problem, err is set and note says why.
func evaluate(f *excelize.File, sheet, cell string) (v grid.Cell, note string, err error) {
	v, _, note, err = evaluateWith(f, definedNames(f), sheet, cell)
	return v, note, err
}

// evaluateWith is evaluate with the workbook's names already looked up. It
// also returns the calculator's result as text, raw, which says less than
// it seems: "5" can be the number or the text (see formula.Typed).
func evaluateWith(f *excelize.File, names func(string) bool, sheet, cell string) (v grid.Cell, raw string, note string, err error) {
	// A call to a function Excel does not have is #NAME? there, whatever
	// the calculator makes of it.
	if text, ferr := f.GetCellFormula(sheet, cell); ferr == nil && text != "" {
		// The calculator reads an error written into a formula as nothing
		// (SUM(Z1,#REF!) is Z1); only without functions or references is
		// the result sure: the error itself.
		if errs := formula.ErrorLiterals(text); len(errs) > 0 {
			if len(formula.Calls(text)) == 0 && len(formula.Refs(text)) == 0 {
				return grid.Cell{Kind: grid.Error, Text: errs[0]}, "", "", nil
			}
			return grid.Cell{}, "", "it holds " + errs[0] + ", which the sheets calculator cannot evaluate (Excel computes it on open)", errUnsupported
		}
		if unknown := formula.UnknownFunctions(text, names); len(unknown) > 0 {
			return grid.Cell{Kind: grid.Error, Text: "#NAME?"}, "", "unknown function " + strings.Join(unknown, ", "), nil
		}
	}
	s, cerr := f.CalcCellValue(sheet, cell, excelize.Options{RawCellValue: true})
	if cerr == nil {
		return typed(s), s, "", nil
	}
	msg := cerr.Error()
	switch {
	case strings.HasPrefix(msg, "#"):
		return grid.Cell{Kind: grid.Error, Text: msg}, "", "", nil
	case strings.HasPrefix(msg, "not support "):
		fn := strings.TrimSuffix(strings.TrimPrefix(msg, "not support "), " function")
		fn = strings.TrimPrefix(formula.FromFile(strings.ReplaceAll(fn, "dot", ".")), "_xlws.")
		return grid.Cell{}, "", fn + " is not supported by the sheets calculator (Excel computes it on open)", errUnsupported
	case strings.Contains(msg, "strconv."):
		return grid.Cell{Kind: grid.Error, Text: "#VALUE!"}, "", "#VALUE! here from text read as a number, which Excel may manage", errUncertain
	}
	if text, ferr := f.GetCellFormula(sheet, cell); ferr == nil {
		if fn := formula.LambdaFamily(text); fn != "" {
			return grid.Cell{}, "", fn + " is not supported by the sheets calculator (Excel computes it on open)", errUnsupported
		}
	}
	return grid.Cell{}, "", "the sheets calculator could not evaluate this formula (" + msg + "); Excel computes it on open", errUnsupported
}

// definedNames reports workbook names, which formulas may call when they
// hold a LAMBDA.
func definedNames(f *excelize.File) func(string) bool {
	names := map[string]bool{}
	for _, n := range f.GetDefinedName() {
		names[strings.ToUpper(n.Name)] = true
	}
	return func(s string) bool { return names[strings.ToUpper(s)] }
}

func typed(s string) grid.Cell {
	switch {
	case s == "":
		return grid.Cell{Kind: grid.Text}
	case s == "TRUE" || s == "FALSE":
		return grid.Cell{Kind: grid.Bool, Bool: s == "TRUE"}
	case strings.HasPrefix(s, "#") && strings.ContainsAny(s, "!?/") || s == "#N/A":
		return grid.Cell{Kind: grid.Error, Text: s}
	}
	if n, err := strconv.ParseFloat(s, 64); err == nil {
		return grid.Cell{Kind: grid.Number, Num: n}
	}
	return grid.Cell{Kind: grid.Text, Text: s}
}

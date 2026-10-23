package xlsx

import (
	"bytes"
	"errors"
	"strconv"
	"strings"
	"sync"

	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
	"github.com/xuri/excelize/v2"
)

// calc evaluates formulas with excelize's engine. The workbook is opened
// lazily, only when a pending value is actually needed.
type calc struct {
	data []byte
	once sync.Once
	f    *excelize.File
	err  error
}

func (c *calc) Resolve(sheet string, row, col int) (grid.Cell, error) {
	c.once.Do(func() {
		c.f, c.err = excelize.OpenReader(bytes.NewReader(c.data))
		c.data = nil
	})
	if c.err != nil {
		return grid.Cell{}, c.err
	}
	v, _, err := evaluate(c.f, sheet, grid.CellName(row, col))
	return v, err
}

// errUnsupported marks formulas the calculator cannot evaluate although
// Excel can (functions it does not implement, such as LET or FILTER).
var errUnsupported = errors.New("not supported by the sheets calculator")

// evaluate calculates one cell and types the result. Errors in the formula
// itself (#DIV/0!, #VALUE!) come back as Error cells. When the calculator,
// not the formula, is the problem, err is set and note says why.
func evaluate(f *excelize.File, sheet, cell string) (v grid.Cell, note string, err error) {
	// A call to a function Excel does not have is #NAME? there, whatever
	// the calculator makes of it.
	if text, ferr := f.GetCellFormula(sheet, cell); ferr == nil && text != "" {
		if unknown := formula.UnknownFunctions(text, definedNames(f)); len(unknown) > 0 {
			return grid.Cell{Kind: grid.Error, Text: "#NAME?"}, "unknown function " + strings.Join(unknown, ", "), nil
		}
	}
	s, cerr := f.CalcCellValue(sheet, cell, excelize.Options{RawCellValue: true})
	if cerr == nil {
		return typed(s), "", nil
	}
	msg := cerr.Error()
	switch {
	case strings.HasPrefix(msg, "#"):
		return grid.Cell{Kind: grid.Error, Text: msg}, "", nil
	case strings.HasPrefix(msg, "not support "):
		fn := strings.TrimSuffix(strings.TrimPrefix(msg, "not support "), " function")
		fn = strings.TrimPrefix(formula.FromFile(strings.ReplaceAll(fn, "dot", ".")), "_xlws.")
		return grid.Cell{}, fn + " is not supported by the sheets calculator (Excel computes it on open)", errUnsupported
	case strings.Contains(msg, "strconv."):
		return grid.Cell{Kind: grid.Error, Text: "#VALUE!"}, "", nil
	}
	if text, ferr := f.GetCellFormula(sheet, cell); ferr == nil {
		if fn := formula.LambdaFamily(text); fn != "" {
			return grid.Cell{}, fn + " is not supported by the sheets calculator (Excel computes it on open)", errUnsupported
		}
	}
	return grid.Cell{}, "the sheets calculator could not evaluate this formula (" + msg + "); Excel computes it on open", errUnsupported
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

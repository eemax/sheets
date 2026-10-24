package xlsx

import (
	"strings"
	"sync"

	"github.com/eemax/sheets/internal/formula"
	"github.com/xuri/excelize/v2"
)

var (
	lacksOnce sync.Once
	lacks     map[string]bool
)

// calcLacks reports whether the calculator lacks an Excel function. It
// asks the calculator, once: a call to a function it lacks fails with "not
// support", whatever the arguments.
func calcLacks(fn string) bool {
	lacksOnce.Do(func() {
		lacks = map[string]bool{}
		f := excelize.NewFile()
		defer f.Close()
		for _, name := range formula.Functions() {
			func() {
				defer func() {
					if recover() != nil {
						lacks[name] = true
					}
				}()
				if err := f.SetCellFormula("Sheet1", "A1", formula.ToFile(name+"()")); err != nil {
					lacks[name] = true
					return
				}
				if _, err := f.CalcCellValue("Sheet1", "A1", excelize.Options{RawCellValue: true}); err != nil && strings.Contains(err.Error(), "not support") {
					lacks[name] = true
				}
			}()
		}
	})
	return lacks[strings.ToUpper(fn)]
}

// computable tells whether the calculator can compute a formula at all; if
// not, why. (Whether it then does, and agrees with Excel, is the judge's
// to find out.)
func computable(text string, rs *resolver, sheet string) (bool, string) {
	return computableDepth(text, rs, sheet, 0)
}

func computableDepth(text string, rs *resolver, sheet string, depth int) (bool, string) {
	if depth > 8 {
		return false, "its names are defined through each other"
	}
	d := formula.DepsOf(text)
	switch {
	case d.External:
		return false, "it links to another workbook"
	case d.Multi:
		return false, "it reads several sheets at once, which the sheets calculator cannot"
	case len(d.Tables) > 0:
		// Those translated into cells do not get here (see calcText).
		name := d.Tables[0]
		if name == "" {
			name = "a table"
		}
		return false, "it uses a reference into " + name + " that the sheets calculator cannot place (outside the table's rows, or to a column it does not have)"
	}
	if errs := formula.ErrorLiterals(text); len(errs) > 0 && (len(formula.Calls(text)) > 0 || len(formula.Refs(text)) > 0) {
		return false, "it holds " + errs[0] + ", which the sheets calculator cannot evaluate"
	}
	if fn := formula.LambdaFamily(text); fn != "" {
		return false, fn + " is not supported by the sheets calculator"
	}
	if why := formula.ArrayUse(text, func(name string) bool { return rs.multiCell(name, sheet) }); why != "" {
		return false, why
	}
	for _, fn := range formula.CalledFunctions(text) {
		if formula.IsFunction(fn) {
			if calcLacks(fn) {
				return false, fn + " is not supported by the sheets calculator"
			}
			continue
		}
		// A workbook name called as a function holds a LAMBDA.
		if def := rs.lookup(fn, sheet); def != nil {
			return false, fn + " is a LAMBDA, which the sheets calculator cannot run"
		}
	}
	for _, n := range d.Names {
		if def := rs.lookup(n, sheet); def != nil {
			// A name whose cells were deleted points at #REF!, which
			// the calculator reads as nothing.
			if errs := formula.ErrorLiterals(def.text); len(errs) > 0 {
				return false, "it uses the name " + n + ", which points at " + errs[0]
			}
			if ok, why := computableDepth(def.text, rs, sheet, depth+1); !ok {
				return false, "it uses the name " + n + ": " + why
			}
		}
	}
	return true, ""
}

// lookup finds the definition of a workbook name as seen from sheet.
func (rs *resolver) lookup(name, sheet string) *nameDef {
	defs := rs.names[strings.ToLower(name)]
	var def *nameDef
	for i := range defs {
		if defs[i].scope == strings.ToLower(sheet) || defs[i].scope == "" && def == nil {
			def = &defs[i]
		}
	}
	return def
}

// multiCell tells whether a workbook name stands for a range of several
// cells.
func (rs *resolver) multiCell(name, sheet string) bool {
	def := rs.lookup(name, sheet)
	if def == nil {
		return false
	}
	for _, r := range formula.Refs(def.text) {
		if r.Range.Rows() > 1 || r.Range.Cols() > 1 {
			return true
		}
	}
	return false
}

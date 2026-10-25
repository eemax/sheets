package xlsx

import (
	"strings"

	"github.com/eemax/sheets/internal/engine"
	"github.com/eemax/sheets/internal/formula"
)

// computable tells whether the engine can compute a formula at all; if
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
	if fn := formula.LambdaFamily(text); fn != "" {
		return false, fn + " is not supported by the sheets calculator"
	}
	for _, fn := range formula.CalledFunctions(text) {
		if formula.IsFunction(fn) {
			if !engine.Has(fn) {
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

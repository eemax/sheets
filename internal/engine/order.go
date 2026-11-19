package engine

import "strings"

// Which references a formula reads as a set: given whole to a function
// whose result no reordering of their cells changes (SUM, MAX, COUNT, an
// average), as opposed to by position (a lookup, a cell in arithmetic, a
// range spilled). check --against asks it of a formula reading rows a sort
// moved: a total over all of them did not move for the sort.

// OrderRead is a reference a formula reads.
type OrderRead struct {
	Sheet string // "" for the formula's own sheet
	Area  Area
	// Name is the defined name it was given as ("" for a reference).
	Name string
	// Set: given whole to a function that reads its cells as a set.
	Set bool
	// Group numbers the call that pairs its ranges cell by cell (SUMIF's
	// range and sum range, SUMIFS's, SUMPRODUCT's arrays): the ranges of one
	// group are a set only together, moved alike. 0 for none.
	Group int
}

// setFns read every argument as a set of values.
var setFns = map[string]bool{"SUM": true, "SUMSQ": true, "MIN": true, "MAX": true, "MINA": true, "MAXA": true,
	"COUNT": true, "COUNTA": true, "COUNTBLANK": true, "AVERAGE": true, "AVERAGEA": true, "PRODUCT": true,
	"MEDIAN": true, "STDEV": true, "STDEV.S": true, "STDEV.P": true, "STDEVA": true, "STDEVP": true, "STDEVPA": true,
	"VAR": true, "VAR.S": true, "VAR.P": true, "VARA": true, "VARP": true, "VARPA": true, "GEOMEAN": true,
	"HARMEAN": true, "AVEDEV": true, "DEVSQ": true, "AND": true, "OR": true}

// firstSetFns read their first argument as a set (the rest are not).
var firstSetFns = map[string]bool{"LARGE": true, "SMALL": true, "PERCENTILE": true, "PERCENTILE.INC": true,
	"PERCENTILE.EXC": true, "QUARTILE": true, "QUARTILE.INC": true, "QUARTILE.EXC": true, "COUNTIF": true}

// OrderReads lists the references a formula reads, each telling whether it
// is read as a set. name resolves a defined name to the reference it stands
// for (false: not a reference, or unknown). ok is false when the formula is
// not one the parser reads, or a name it uses cannot be resolved.
func OrderReads(text string, name func(sheet, n string) (RefArg, bool)) ([]OrderRead, bool) {
	n, err := Parse(text)
	if err != nil {
		return nil, false
	}
	var out []OrderRead
	ok := true
	groups := 0
	// ref records a reference argument (set, in group g) and tells whether a
	// was one.
	ref := func(a *node, set bool, g int) bool {
		if r, isRef := refArg(a); isRef {
			out = append(out, OrderRead{Sheet: r.Sheet, Area: r.Area, Set: set, Group: g})
			return true
		}
		if a.kind == nName {
			r, found := name(a.sheet, a.str)
			if !found {
				ok = false
				return true
			}
			out = append(out, OrderRead{Sheet: r.Sheet, Area: r.Area, Name: a.str, Set: set, Group: g})
			return true
		}
		return false
	}
	var walk func(n *node)
	// arg walks an argument: a reference alone is read as set says.
	arg := func(a *node, set bool, g int) {
		if !ref(a, set, g) {
			walk(a)
		}
	}
	walk = func(n *node) {
		switch n.kind {
		case nRef, nName:
			ref(n, false, 0)
			return
		case nBinary:
			if n.str == ":" {
				if ref(n, false, 0) {
					return
				}
			}
		case nCall:
			fn := strings.ToUpper(bareName(n.str))
			switch {
			case setFns[fn]:
				for _, a := range n.args {
					arg(a, true, 0)
				}
				return
			case firstSetFns[fn]:
				for i, a := range n.args {
					arg(a, i == 0, 0)
				}
				return
			case fn == "SUBTOTAL":
				for i, a := range n.args {
					arg(a, i > 0, 0)
				}
				return
			case fn == "SUMIF" || fn == "AVERAGEIF" || fn == "SUMIFS" || fn == "AVERAGEIFS" || fn == "MAXIFS" ||
				fn == "MINIFS" || fn == "COUNTIFS" || fn == "SUMPRODUCT":
				groups++
				for i, a := range n.args {
					paired := true
					switch fn {
					case "SUMIF", "AVERAGEIF":
						paired = i != 1
					case "COUNTIFS":
						paired = i%2 == 0
					case "SUMIFS", "AVERAGEIFS", "MAXIFS", "MINIFS":
						paired = i == 0 || i%2 == 1
					}
					if paired {
						arg(a, true, groups)
					} else {
						arg(a, false, 0)
					}
				}
				return
			}
		}
		for _, a := range n.args {
			walk(a)
		}
	}
	walk(n)
	return out, ok
}

// Outer gives the call a formula is, alone (=MAX(B2:B9), not
// =MAX(B2:B9)*2 or =IF(A1,MAX(B2:B9))), with the references given as its
// arguments: ok is false when the formula is no call alone, or an argument
// of it is not a reference.
func Outer(text string) (Call, bool) {
	n, err := Parse(text)
	if err != nil || n.kind != nCall {
		return Call{}, false
	}
	c := Call{Func: strings.ToUpper(bareName(n.str))}
	for _, a := range n.args {
		r, ok := refArg(a)
		if !ok {
			return Call{}, false
		}
		c.Refs = append(c.Refs, r)
	}
	return c, len(c.Refs) > 0
}

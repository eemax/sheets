package engine

import (
	"strings"

	"github.com/eemax/sheets/internal/grid"
)

// What a formula's text says, for checks that read formulas (lint) rather
// than compute them.

// A RefArg is a reference given directly as an argument.
type RefArg struct {
	Sheet string // "" for the formula's own sheet
	Area  Area
}

// A Call is a function call with the references given directly as its
// arguments: SUM(B2:B9,D4) gives B2:B9 and D4.
type Call struct {
	Func string // upper case, without storage prefixes
	Refs []RefArg
}

// Calls lists the function calls of a formula, outermost first.
func Calls(text string) ([]Call, error) {
	n, err := Parse(text)
	if err != nil {
		return nil, err
	}
	var out []Call
	var walk func(n *node)
	walk = func(n *node) {
		if n.kind == nCall {
			c := Call{Func: strings.ToUpper(bareName(n.str))}
			for _, a := range n.args {
				if r, ok := refArg(a); ok {
					c.Refs = append(c.Refs, r)
				}
			}
			out = append(out, c)
		}
		for _, a := range n.args {
			walk(a)
		}
	}
	walk(n)
	return out, nil
}

// Reference tells whether a formula is a reference alone (=Budget!N9).
func Reference(text string) (RefArg, bool) {
	n, err := Parse(text)
	if err != nil {
		return RefArg{}, false
	}
	return refArg(n)
}

// rangeResult tells whether a call gives a reference of several cells, and
// its size where known: INDEX with its row or column left out or 0
// (INDEX(T2:T600,,3), a column), OFFSET of more than one cell.
func rangeResult(name string, args []*node) (Area, bool) {
	zero := func(i int) bool {
		return i >= len(args) || args[i].kind == nMissing || args[i].kind == nNum && args[i].num == 0
	}
	big := func(i int) bool { return i < len(args) && args[i].kind == nNum && args[i].num > 1 }
	switch name {
	case "INDEX":
		if len(args) == 0 {
			return Area{}, false
		}
		r, ok := refArg(args[0])
		if !ok {
			return Area{}, false
		}
		rows, cols := r.Area.Rows(), r.Area.Cols()
		n := 1
		switch {
		case len(args) == 1:
			n = rows * cols
		case len(args) == 2 && (rows == 1 || cols == 1):
			// INDEX(A2:A9, 3): one cell of a line; 0 is all of it.
			if zero(1) {
				n = rows * cols
			}
		case len(args) == 2:
			n = cols // a row of the area, or all of it
			if zero(1) {
				n = rows * cols
			}
		default:
			if zero(1) {
				n *= rows
			}
			if zero(2) {
				n *= cols
			}
		}
		return Area{R1: 1, C1: 1, R2: n, C2: 1}, n > 1
	case "OFFSET":
		if big(3) || big(4) {
			return Area{R1: 1, C1: 1, R2: 2, C2: 1}, true
		}
		if len(args) > 0 && !(len(args) > 3 && args[3].kind != nMissing) && !(len(args) > 4 && args[4].kind != nMissing) {
			if r, ok := refArg(args[0]); ok && r.Area.Rows()*r.Area.Cols() > 1 {
				return r.Area, true
			}
		}
	}
	return Area{}, false
}

// refArg reads a reference node: an area, or two joined by a colon on one
// sheet.
func refArg(n *node) (RefArg, bool) {
	switch {
	case n.kind == nRef:
		return RefArg{Sheet: n.sheet, Area: n.area}, true
	case n.kind == nBinary && n.str == ":" && len(n.args) == 2:
		a, ok1 := refArg(n.args[0])
		b, ok2 := refArg(n.args[1])
		if !ok1 || !ok2 || !strings.EqualFold(a.Sheet, b.Sheet) && b.Sheet != "" {
			return RefArg{}, false
		}
		return RefArg{Sheet: a.Sheet, Area: norm(Area{
			R1: min(a.Area.R1, b.Area.R1), C1: min(a.Area.C1, b.Area.C1),
			R2: max(a.Area.R2, b.Area.R2), C2: max(a.Area.C2, b.Area.C2),
		})}, true
	}
	return RefArg{}, false
}

// ArrayUse tells whether a formula reads a range of several cells where
// one value goes (in arithmetic, a comparison, or an argument that takes
// one value), outside the functions that take arrays: =SUM(IF(A2:A9>5,
// B2:B9)). Excel 365 computes such a formula over the whole range; older
// Excel, and Excel 365 reading a formula stored as theirs, over one cell
// of it. spills tells whether the array reaches the result (=A2:A9*2),
// rather than a function taking it whole. Whole columns and rows count
// only when a function takes them whole: as a result they mean the cell
// in the formula's row, as written for older Excel.
func ArrayUse(text string) (array, spills bool) {
	n, err := Parse(text)
	if err != nil {
		return false, false
	}
	var walk func(n *node, arr, taken, value bool)
	found := func(a Area, arr, taken, value bool) {
		if !value || arr || a.Rows()*a.Cols() < 2 {
			return
		}
		if taken {
			array = true
		} else if a.R1 > 1 || a.R2 < grid.MaxRows {
			if a.C1 > 1 || a.C2 < grid.MaxCols {
				array, spills = true, true
			}
		}
	}
	walk = func(n *node, arr, taken, value bool) {
		switch n.kind {
		case nRef:
			found(n.area, arr, taken, value)
		case nBinary:
			if n.str == ":" || n.str == " " {
				if r, ok := refArg(n); ok {
					found(r.Area, arr, taken, value)
				}
				return
			}
			for _, a := range n.args {
				walk(a, arr, taken, true)
			}
		case nUnary, nPercent:
			walk(n.args[0], arr, taken, true)
		case nCall:
			name := strings.ToUpper(bareName(n.str))
			if r, ok := rangeResult(name, n.args); ok {
				found(r, arr, taken, value)
			}
			for i, a := range n.args {
				switch name {
				case "SINGLE":
					// @A2:A9 asks for one cell.
				case "ANCHORARRAY":
					// D2#: the cells a dynamic array fills, a range.
					found(Area{R1: 1, C1: 1, R2: 2, C2: 1}, arr, taken, value)
				case "IF", "CHOOSE":
					walk(a, arr, taken, value || i == 0)
				case "IFS":
					walk(a, arr, taken, value || i%2 == 0)
				case "IFERROR", "IFNA", "LET", "SWITCH":
					walk(a, arr, taken, value || name == "SWITCH" && i == 0)
				case "COUNTIF", "SUMIF", "AVERAGEIF", "COUNTIFS", "SUMIFS", "AVERAGEIFS", "MAXIFS", "MINIFS":
					// A range of criteria asks for a count or sum each.
					crit := i == 1 && strings.HasSuffix(name, "IF") ||
						name == "COUNTIFS" && i%2 == 1 || name != "COUNTIFS" && strings.HasSuffix(name, "IFS") && i > 0 && i%2 == 0
					walk(a, arr, taken || !crit, crit)
				default:
					s := fns[name]
					switch {
					case s == nil:
					case s.kind(i) == aArr:
						walk(a, true, true, false)
					case s.kind(i) == aRef:
						walk(a, arr, true, false)
					default:
						walk(a, arr, taken, true)
					}
				}
			}
		default:
			for _, a := range n.args {
				walk(a, arr, taken, false)
			}
		}
	}
	walk(n, false, false, true)
	return array, spills
}

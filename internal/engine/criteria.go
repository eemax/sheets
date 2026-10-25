package engine

import (
	"strconv"
	"strings"
)

// A criterion of COUNTIF, SUMIF and the *IFS functions: "<5", ">=1/1/2024",
// "<>x", "a*", "", or a value. As in Excel: text compares without regard
// to case; * ? and ~ are wildcards in = and <> tests on text; text that
// reads as a number is matched as that number; a blank criterion matches
// blank cells.
type criterion struct {
	op    string // = <> < > <= >=
	kind  Kind   // Number, Text, Bool, Error, or Empty for "" / "<>"
	num   float64
	text  string // folded
	wild  bool
	exact Value // a value given as the criterion (not text)
	// unknown, when set, receives a cell's text whose match depends on
	// Excel's language setting.
	unknown *string
}

func parseCriterion(v Value) criterion {
	switch v.Kind {
	case Number:
		return criterion{op: "=", kind: Number, num: v.Num}
	case Bool:
		return criterion{op: "=", kind: Bool, num: b2f(v.Bool)}
	case Error:
		return criterion{op: "=", kind: Error, text: v.Str}
	case Empty:
		// A blank cell as the criterion matches cells holding 0? Excel
		// reads it as "=0".
		return criterion{op: "=", kind: Number, num: 0}
	}
	s := v.Str
	op := "="
	for _, p := range []string{"<=", ">=", "<>", "<", ">", "="} {
		if strings.HasPrefix(s, p) {
			op, s = p, s[len(p):]
			break
		}
	}
	if s == "" {
		return criterion{op: op, kind: Empty}
	}
	if n, ok := textNumber(s); ok {
		return criterion{op: op, kind: Number, num: n, text: foldCase(s)}
	}
	switch strings.ToUpper(s) {
	case "TRUE":
		return criterion{op: op, kind: Bool, num: 1}
	case "FALSE":
		return criterion{op: op, kind: Bool, num: 0}
	}
	if code, ok := errorAt(s); ok && code == strings.ToUpper(s) {
		return criterion{op: op, kind: Error, text: code}
	}
	return criterion{op: op, kind: Text, text: foldCase(s), wild: strings.ContainsAny(s, "*?~")}
}

func b2f(b bool) float64 {
	if b {
		return 1
	}
	return 0
}

// match tests a cell's value.
func (cr criterion) match(v Value) bool {
	switch cr.kind {
	case Empty:
		blank := v.Kind == Empty || v.Kind == Text && v.Str == ""
		switch cr.op {
		case "=":
			return blank
		case "<>":
			return v.Kind != Empty
		}
		return false
	case Number:
		var x float64
		switch v.Kind {
		case Number:
			x = v.Num
		case Text:
			if cr.op != "=" && cr.op != "<>" {
				return false // only = and <> read text as a number
			}
			n, ok, dependent := readNumber(v.Str)
			if dependent && cr.unknown != nil {
				// Read by Excel's language: whether it matches cannot
				// be told.
				*cr.unknown = v.Str
			}
			if !ok {
				return cr.op == "<>"
			}
			x = n
		default:
			return cr.op == "<>"
		}
		return cmpOp(cr.op, cmpFloat(x, cr.num))
	case Bool:
		if v.Kind != Bool {
			return cr.op == "<>"
		}
		return cmpOp(cr.op, cmpFloat(b2f(v.Bool), cr.num))
	case Error:
		if v.Kind != Error {
			return cr.op == "<>"
		}
		return cmpOp(cr.op, strings.Compare(v.Str, cr.text))
	}
	// Text.
	if v.Kind != Text {
		if cr.op == "<>" {
			return true
		}
		if (cr.op == "=") && v.Kind == Empty {
			return false
		}
		return false
	}
	t := foldCase(v.Str)
	switch cr.op {
	case "=":
		if cr.wild {
			return wildMatch(cr.text, t)
		}
		return t == cr.text
	case "<>":
		if cr.wild {
			return !wildMatch(cr.text, t)
		}
		return t != cr.text
	}
	return cmpOp(cr.op, strings.Compare(t, cr.text))
}

func cmpOp(op string, r int) bool {
	switch op {
	case "=":
		return r == 0
	case "<>":
		return r != 0
	case "<":
		return r < 0
	case ">":
		return r > 0
	case "<=":
		return r <= 0
	}
	return r >= 0
}

// wildMatch matches text against a pattern with * (any run), ? (one
// character) and ~ (the next character as itself); both folded.
func wildMatch(pattern, text string) bool {
	p, t := []rune(pattern), []rune(text)
	var match func(i, j int) bool
	memo := map[[2]int]bool{}
	match = func(i, j int) bool {
		k := [2]int{i, j}
		if v, ok := memo[k]; ok {
			return v
		}
		var r bool
		switch {
		case i == len(p):
			r = j == len(t)
		case p[i] == '*':
			r = match(i+1, j) || j < len(t) && match(i, j+1)
		case p[i] == '?':
			r = j < len(t) && match(i+1, j+1)
		case p[i] == '~' && i+1 < len(p) && (p[i+1] == '*' || p[i+1] == '?' || p[i+1] == '~'):
			r = j < len(t) && t[j] == p[i+1] && match(i+2, j+1)
		default:
			r = j < len(t) && t[j] == p[i] && match(i+1, j+1)
		}
		memo[k] = r
		return r
	}
	return match(0, 0)
}

// ifs is the engine of the conditional aggregates: op over target's cells
// (or a count) where every range's cell meets its criterion. Criteria
// that are arrays (in array evaluation) give an array of results.
func ifs(c *ctx, op string, target Value, ranges, crits []Value) Value {
	for _, r := range append(append([]Value{}, ranges...), target) {
		if r.Kind == Error {
			return r
		}
	}
	for _, r := range ranges {
		if r.Kind != Ref {
			return errv(ErrValue)
		}
	}
	if op != "count" && target.Kind != Ref {
		return errv(ErrValue)
	}
	// Each range is a grid of the same shape as the first; SUMIF's sum
	// range stretches from its top-left cell to that shape.
	shape := ranges[0].Ref.Areas[0]
	for _, r := range ranges[1:] {
		a := r.Ref.Areas[0]
		if a.Rows() != shape.Rows() || a.Cols() != shape.Cols() {
			return errv(ErrValue)
		}
	}
	// Only the cells up to the last filled row of any range can match
	// anything but a blank criterion; past it every range is blank.
	rows := 0
	for _, r := range append(append([]Value{}, ranges...), target) {
		if r.Kind != Ref {
			continue
		}
		a := r.Ref.Areas[0]
		b := c.bounded(r.Ref.Sheet, Area{a.R1, a.C1, a.R1 + shape.Rows() - 1, a.C1 + shape.Cols() - 1})
		rows = max(rows, b.R2-a.R1+1)
	}
	rows = min(rows, shape.Rows())
	run := func(cs []criterion) Value {
		var unknown string
		for k := range cs {
			cs[k].unknown = &unknown
		}
		defer func() {
			if unknown != "" {
				c.fail(unsupported("it compares " + strconv.Quote(unknown) + " with a number, which depends on Excel's language setting"))
			}
		}()
		var acc []float64
		count := 0
		for i := 0; i < rows; i++ {
			for j := 0; j < shape.Cols(); j++ {
				ok := true
				for k, r := range ranges {
					a := r.Ref.Areas[0]
					if !cs[k].match(c.cell(r.Ref.Sheet, a.R1+i, a.C1+j)) {
						ok = false
						break
					}
				}
				if !ok {
					continue
				}
				count++
				if op == "count" {
					continue
				}
				a := target.Ref.Areas[0]
				v := c.cell(target.Ref.Sheet, a.R1+i, a.C1+j)
				switch v.Kind {
				case Error:
					return v
				case Number:
					acc = append(acc, v.Num)
				}
			}
		}
		// Blank rows past the data match a blank criterion too.
		if rows < shape.Rows() && op == "count" {
			blank := true
			for _, cr := range cs {
				if !cr.match(Value{}) {
					blank = false
				}
			}
			if blank {
				count += (shape.Rows() - rows) * shape.Cols()
			}
		}
		switch op {
		case "count":
			return num(float64(count))
		case "sum":
			return num(sum(acc))
		case "average":
			if len(acc) == 0 {
				return errv(ErrDiv0)
			}
			return num(mean(acc))
		case "max", "min":
			if len(acc) == 0 {
				return num(0)
			}
			m := acc[0]
			for _, x := range acc[1:] {
				if op == "max" && x > m || op == "min" && x < m {
					m = x
				}
			}
			return num(m)
		}
		return errv(ErrValue)
	}
	// Criteria that are arrays apply one by one, giving an array. A range
	// of criteria does so in array evaluation; in a normal formula it
	// gives the cell in the formula's row or column, and only when there
	// is none does it apply as an array (Excel then shows the first).
	var arrs []*Arr
	rr, cc := 1, 1
	for _, cv := range crits {
		var a *Arr
		if c.lifts(cv) || cv.Kind == Ref && isMulti(cv) && !c.intersects(cv) {
			a = c.toArr(cv)
			rr, cc = max(rr, a.Rows), max(cc, a.Cols)
		}
		arrs = append(arrs, a)
	}
	if rr == 1 && cc == 1 {
		cs := make([]criterion, len(crits))
		for k, cv := range crits {
			cv = c.scalar(cv)
			if cv.Kind == Error {
				return cv
			}
			cs[k] = parseCriterion(cv)
		}
		return run(cs)
	}
	out := newArr(rr, cc)
	for i := 0; i < rr; i++ {
		for j := 0; j < cc; j++ {
			cs := make([]criterion, len(crits))
			for k, cv := range crits {
				if arrs[k] != nil {
					cv = pick(arrs[k], i, j)
				} else {
					cv = c.scalar(cv)
				}
				cs[k] = parseCriterion(cv)
			}
			out.V[i*cc+j] = run(cs)
		}
	}
	return arrOf(out)
}

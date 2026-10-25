package engine

import (
	"math"
	"strings"

	"github.com/eemax/sheets/internal/formula"
)

// argKind is how a function takes an argument.
type argKind uint8

const (
	// aVal is one value: a reference reads its cell (outside array
	// evaluation, a range gives the cell in the formula's row or column);
	// in array evaluation an array here makes the function apply to each
	// element.
	aVal argKind = iota
	// aRef is a reference or array as given (SUM's numbers, COUNTIF's
	// range).
	aRef
	// aArr is aRef evaluated as an array even in a normal formula
	// (SUMPRODUCT's arrays, INDEX's array).
	aArr
)

type fnSpec struct {
	min, max int       // max < 0: no limit
	args     []argKind // per argument; the last one repeats
	f        func(c *ctx, a []Value) Value
	lazy     func(c *ctx, a []*node) Value // gets the argument nodes instead
	noLift   bool                          // takes arrays in aVal places itself
}

func (s *fnSpec) kind(i int) argKind {
	if len(s.args) == 0 {
		return aVal
	}
	if i < len(s.args) {
		return s.args[i]
	}
	return s.args[len(s.args)-1]
}

var fns = map[string]*fnSpec{}

func reg(names string, s *fnSpec) {
	for _, n := range strings.Fields(names) {
		fns[n] = s
	}
}

// Has reports whether the engine computes a function.
func Has(name string) bool {
	_, ok := fns[strings.ToUpper(name)]
	return ok
}

func (c *ctx) call(n *node) Value {
	s := fns[n.str]
	if s == nil {
		switch {
		case formula.IsFunction(n.str):
			return c.fail(unsupported(n.str + " is not supported by the sheets calculator"))
		case c.hasName(n.str):
			return c.fail(unsupported(n.str + " is a LAMBDA, which the sheets calculator cannot run"))
		}
		// A function Excel does not have: #NAME?, as there.
		return errv(ErrName)
	}
	if len(n.args) < s.min || s.max >= 0 && len(n.args) > s.max {
		return errv(ErrValue)
	}
	if s.lazy != nil {
		return s.lazy(c, n.args)
	}
	vals := make([]Value, len(n.args))
	lift := false
	for i, a := range n.args {
		switch s.kind(i) {
		case aArr:
			vals[i] = c.child(true).eval(a)
		case aRef:
			vals[i] = c.eval(a)
		default:
			v := c.eval(a)
			if !s.noLift && c.lifts(v) {
				lift = true
				vals[i] = v
			} else {
				vals[i] = c.scalar(v)
			}
		}
	}
	for _, v := range vals {
		if v.Kind == Error && (v.Str == ErrUnsupported || v.Str == ErrCycle) {
			return c.fail(v)
		}
	}
	if lift {
		return c.liftCall(s, n, vals)
	}
	return c.known(s.f(c, vals))
}

// known passes a function's result on, noting it when it is unknown.
func (c *ctx) known(v Value) Value {
	if v.Kind == Error && v.Str == ErrUnsupported {
		return c.fail(v)
	}
	return v
}

// liftCall applies a function element by element over the arrays given
// in its single-value places.
func (c *ctx) liftCall(s *fnSpec, n *node, vals []Value) Value {
	arrs := make([]*Arr, len(vals))
	rows, cols := 1, 1
	for i, v := range vals {
		if s.kind(i) == aVal && c.lifts(v) {
			arrs[i] = c.toArr(v)
			rows, cols = max(rows, arrs[i].Rows), max(cols, arrs[i].Cols)
		}
	}
	out := newArr(rows, cols)
	args := make([]Value, len(vals))
	for i := 0; i < rows; i++ {
		for j := 0; j < cols; j++ {
			for k, v := range vals {
				if arrs[k] != nil {
					args[k] = pick(arrs[k], i, j)
				} else {
					args[k] = v
				}
			}
			out.V[i*cols+j] = c.known(s.f(c, args))
		}
	}
	return arrOf(out)
}

// fail records that the formula met something the engine cannot compute;
// the formula's result is then unknown, whatever wraps it.
func (c *ctx) fail(v Value) Value {
	if c.failed != nil && *c.failed == nil {
		x := v
		*c.failed = &x
	}
	return v
}

// ---- argument helpers ----

func opt(a []Value, i int) (Value, bool) {
	if i >= len(a) || isMissing(a[i]) {
		return Value{}, false
	}
	return a[i], true
}

// numArg reads argument i as a number; def when left out.
func numArg(a []Value, i int, def float64) (float64, *Value) {
	v, ok := opt(a, i)
	if !ok {
		return def, nil
	}
	return toNumber(v)
}

func textArg(a []Value, i int, def string) (string, *Value) {
	v, ok := opt(a, i)
	if !ok {
		return def, nil
	}
	return toText(v)
}

func boolArg(a []Value, i int, def bool) (bool, *Value) {
	v, ok := opt(a, i)
	if !ok {
		return def, nil
	}
	return toBool(v)
}

// fn1 makes a function of one number.
func fn1(f func(x float64) Value) *fnSpec {
	return &fnSpec{min: 1, max: 1, f: func(c *ctx, a []Value) Value {
		x, err := toNumber(a[0])
		if err != nil {
			return *err
		}
		return f(x)
	}}
}

// fn2 makes a function of two numbers, the second optional (def).
func fn2(min int, def float64, f func(x, y float64) Value) *fnSpec {
	return &fnSpec{min: min, max: 2, f: func(c *ctx, a []Value) Value {
		x, err := toNumber(a[0])
		if err != nil {
			return *err
		}
		y, err := numArg(a, 1, def)
		if err != nil {
			return *err
		}
		return f(x, y)
	}}
}

// ---- logical and information functions ----

func init() {
	reg("IF", &fnSpec{min: 1, max: 3, lazy: lazyIf})
	reg("IFERROR", &fnSpec{min: 2, max: 2, lazy: func(c *ctx, a []*node) Value { return lazyIfError(c, a, false) }})
	reg("IFNA", &fnSpec{min: 2, max: 2, lazy: func(c *ctx, a []*node) Value { return lazyIfError(c, a, true) }})
	reg("IFS", &fnSpec{min: 2, max: -1, lazy: lazyIfs})
	reg("CHOOSE", &fnSpec{min: 2, max: -1, lazy: lazyChoose})
	reg("SWITCH", &fnSpec{min: 3, max: -1, lazy: lazySwitch})
	reg("LET", &fnSpec{min: 3, max: -1, lazy: lazyLet})
	reg("AND", &fnSpec{min: 1, max: -1, args: []argKind{aRef}, f: func(c *ctx, a []Value) Value { return logic(c, a, "and") }})
	reg("OR", &fnSpec{min: 1, max: -1, args: []argKind{aRef}, f: func(c *ctx, a []Value) Value { return logic(c, a, "or") }})
	reg("XOR", &fnSpec{min: 1, max: -1, args: []argKind{aRef}, f: func(c *ctx, a []Value) Value { return logic(c, a, "xor") }})
	reg("NOT", &fnSpec{min: 1, max: 1, f: func(c *ctx, a []Value) Value {
		b, err := toBool(a[0])
		if err != nil {
			return *err
		}
		return boolean(!b)
	}})
	reg("TRUE", &fnSpec{min: 0, max: 0, f: func(*ctx, []Value) Value { return boolean(true) }})
	reg("FALSE", &fnSpec{min: 0, max: 0, f: func(*ctx, []Value) Value { return boolean(false) }})
	reg("NA", &fnSpec{min: 0, max: 0, f: func(*ctx, []Value) Value { return errv(ErrNA) }})
	is := func(test func(v Value) bool) *fnSpec {
		return &fnSpec{min: 1, max: 1, f: func(c *ctx, a []Value) Value { return boolean(test(a[0])) }}
	}
	reg("ISBLANK", is(func(v Value) bool { return v.Kind == Empty }))
	reg("ISNUMBER", is(func(v Value) bool { return v.Kind == Number }))
	reg("ISTEXT", is(func(v Value) bool { return v.Kind == Text }))
	reg("ISNONTEXT", is(func(v Value) bool { return v.Kind != Text }))
	reg("ISLOGICAL", is(func(v Value) bool { return v.Kind == Bool }))
	reg("ISERROR", is(func(v Value) bool { return v.Kind == Error }))
	reg("ISERR", is(func(v Value) bool { return v.Kind == Error && v.Str != ErrNA }))
	reg("ISNA", is(func(v Value) bool { return v.Kind == Error && v.Str == ErrNA }))
	reg("ISEVEN", &fnSpec{min: 1, max: 1, f: func(c *ctx, a []Value) Value { return parity(a[0], 0) }})
	reg("ISODD", &fnSpec{min: 1, max: 1, f: func(c *ctx, a []Value) Value { return parity(a[0], 1) }})
	reg("ISREF", &fnSpec{min: 1, max: 1, args: []argKind{aRef}, f: func(c *ctx, a []Value) Value { return boolean(a[0].Kind == Ref) }})
	reg("ISFORMULA", &fnSpec{min: 1, max: 1, args: []argKind{aRef}, f: func(c *ctx, a []Value) Value {
		if a[0].Kind != Ref {
			return errv(ErrValue)
		}
		ar := a[0].Ref.Areas[0]
		_, f := c.e.book.Cell(a[0].Ref.Sheet, ar.R1, ar.C1)
		return boolean(f != "")
	}})
	reg("ERROR.TYPE", &fnSpec{min: 1, max: 1, f: func(c *ctx, a []Value) Value {
		if a[0].Kind != Error {
			return errv(ErrNA)
		}
		for i, code := range []string{ErrNull, ErrDiv0, ErrValue, ErrRef, ErrName, ErrNum, ErrNA, "#GETTING_DATA", ErrSpill} {
			if a[0].Str == code {
				return num(float64(i + 1))
			}
		}
		if a[0].Str == ErrCalc {
			return num(14)
		}
		return errv(ErrNA)
	}})
	reg("N", &fnSpec{min: 1, max: 1, f: func(c *ctx, a []Value) Value {
		switch v := a[0]; v.Kind {
		case Number, Error:
			return v
		case Bool:
			if v.Bool {
				return num(1)
			}
		}
		return num(0)
	}})
	reg("T", &fnSpec{min: 1, max: 1, f: func(c *ctx, a []Value) Value {
		switch v := a[0]; v.Kind {
		case Text, Error:
			return v
		}
		return str("")
	}})
	reg("TYPE", &fnSpec{min: 1, max: 1, noLift: true, args: []argKind{aRef}, f: func(c *ctx, a []Value) Value {
		v := a[0]
		if isMulti(v) {
			return num(64)
		}
		v = c.scalar(v)
		switch v.Kind {
		case Text:
			return num(2)
		case Bool:
			return num(4)
		case Error:
			return num(16)
		}
		return num(1)
	}})
	// ANCHORARRAY(D2), written D2#: the range D2's dynamic array fills.
	reg("ANCHORARRAY", &fnSpec{min: 1, max: 1, lazy: func(c *ctx, a []*node) Value {
		v := c.eval(a[0])
		if v.Kind != Ref || len(v.Ref.Areas) != 1 {
			return errv(ErrRef)
		}
		at := v.Ref.Areas[0]
		ab, ok := c.e.book.(ArrayBook)
		if !ok {
			return errv(ErrRef)
		}
		ar, ac, fill, ok := ab.ArrayAt(v.Ref.Sheet, at.R1, at.C1)
		if !ok || ar != at.R1 || ac != at.C1 {
			return errv(ErrRef) // not a dynamic array's cell
		}
		c.e.Compute(v.Ref.Sheet, ar, ac)
		return Value{Kind: Ref, Ref: &RefV{Sheet: v.Ref.Sheet, Areas: []Area{fill}}}
	}})
	reg("SINGLE", &fnSpec{min: 1, max: 1, lazy: func(c *ctx, a []*node) Value {
		return c.child(false).scalar(c.eval(a[0]))
	}})
	reg("ROW", &fnSpec{min: 0, max: 1, args: []argKind{aRef}, f: func(c *ctx, a []Value) Value { return rowCol(c, a, true) }})
	reg("COLUMN", &fnSpec{min: 0, max: 1, args: []argKind{aRef}, f: func(c *ctx, a []Value) Value { return rowCol(c, a, false) }})
	reg("ROWS", &fnSpec{min: 1, max: 1, args: []argKind{aRef}, f: func(c *ctx, a []Value) Value { return dims(c, a[0], true) }})
	reg("COLUMNS", &fnSpec{min: 1, max: 1, args: []argKind{aRef}, f: func(c *ctx, a []Value) Value { return dims(c, a[0], false) }})
}

func parity(v Value, want int) Value {
	x, err := toNumber(v)
	if err != nil {
		return *err
	}
	return boolean(int64(math.Trunc(x))%2 == int64(want) || int64(math.Trunc(x))%2 == -int64(want))
}

func rowCol(c *ctx, a []Value, row bool) Value {
	if len(a) == 0 || isMissing(a[0]) {
		if row {
			return num(float64(c.row))
		}
		return num(float64(c.col))
	}
	v := a[0]
	if v.Kind == Error {
		return v
	}
	if v.Kind != Ref {
		return errv(ErrValue)
	}
	ar := v.Ref.Areas[0]
	n := ar.Rows()
	if !row {
		n = ar.Cols()
	}
	if c.arr && n > 1 {
		// ROW(A1:A5) evaluated as an array: each row's number.
		var out *Arr
		if row {
			out = newArr(n, 1)
		} else {
			out = newArr(1, n)
		}
		for i := 0; i < n; i++ {
			if row {
				out.V[i] = num(float64(ar.R1 + i))
			} else {
				out.V[i] = num(float64(ar.C1 + i))
			}
		}
		return arrOf(out)
	}
	if row {
		return num(float64(ar.R1))
	}
	return num(float64(ar.C1))
}

func dims(c *ctx, v Value, rows bool) Value {
	switch v.Kind {
	case Error:
		return v
	case Ref:
		if len(v.Ref.Areas) != 1 {
			return errv(ErrRef)
		}
		if rows {
			return num(float64(v.Ref.Areas[0].Rows()))
		}
		return num(float64(v.Ref.Areas[0].Cols()))
	case Array:
		if rows {
			return num(float64(v.Arr.Rows))
		}
		return num(float64(v.Arr.Cols))
	}
	return num(1)
}

func lazyIf(c *ctx, a []*node) Value {
	cond := c.eval(a[0])
	if c.lifts(cond) {
		ca := c.toArr(cond)
		var ta, fa *Arr
		rows, cols := ca.Rows, ca.Cols
		if len(a) > 1 {
			ta = c.toArr(c.branch(a[1]))
			rows, cols = max(rows, ta.Rows), max(cols, ta.Cols)
		}
		if len(a) > 2 {
			fa = c.toArr(c.branch(a[2]))
			rows, cols = max(rows, fa.Rows), max(cols, fa.Cols)
		}
		out := newArr(rows, cols)
		for i := 0; i < rows; i++ {
			for j := 0; j < cols; j++ {
				cv := pick(ca, i, j)
				var r Value
				if cv.Kind == Error {
					r = cv
				} else if b, err := toBool(cv); err != nil {
					r = *err
				} else if b {
					r = num(0)
					if ta != nil {
						r = pick(ta, i, j)
					}
				} else {
					r = boolean(false)
					if fa != nil {
						r = pick(fa, i, j)
					}
				}
				out.V[i*cols+j] = r
			}
		}
		return arrOf(out)
	}
	cv := c.scalar(cond)
	if cv.Kind == Error {
		return cv
	}
	b, err := toBool(cv)
	if err != nil {
		return *err
	}
	if b {
		if len(a) < 2 {
			return boolean(true)
		}
		return c.branch(a[1])
	}
	if len(a) < 3 {
		return boolean(false)
	}
	return c.branch(a[2])
}

// branch evaluates an argument that a function returns as is: left out,
// it is 0 (IF(A1,,1)).
func (c *ctx) branch(n *node) Value {
	if n.kind == nMissing {
		return num(0)
	}
	return c.eval(n)
}

func lazyIfError(c *ctx, a []*node, onlyNA bool) Value {
	v := c.eval(a[0])
	catches := func(x Value) bool {
		return x.Kind == Error && x.Str != ErrUnsupported && x.Str != ErrCycle && (!onlyNA || x.Str == ErrNA)
	}
	if c.lifts(v) {
		va := c.toArr(v)
		var alt *Arr
		out := newArr(va.Rows, va.Cols)
		for i, x := range va.V {
			if catches(x) {
				if alt == nil {
					alt = c.toArr(c.branch(a[1]))
				}
				out.V[i] = pick(alt, i/va.Cols, i%va.Cols)
			} else {
				out.V[i] = x
			}
		}
		return arrOf(out)
	}
	s := c.scalar(v)
	if catches(s) {
		return c.branch(a[1])
	}
	if s.Kind == Empty {
		return num(0)
	}
	return s
}

func lazyIfs(c *ctx, a []*node) Value {
	if len(a)%2 != 0 {
		return errv(ErrValue)
	}
	for i := 0; i < len(a); i += 2 {
		cv := c.scalar(c.eval(a[i]))
		if cv.Kind == Error {
			return cv
		}
		b, err := toBool(cv)
		if err != nil {
			return *err
		}
		if b {
			return c.branch(a[i+1])
		}
	}
	return errv(ErrNA)
}

func lazyChoose(c *ctx, a []*node) Value {
	iv := c.eval(a[0])
	if c.lifts(iv) {
		ia := c.toArr(iv)
		out := newArr(ia.Rows, ia.Cols)
		cache := map[int]*Arr{}
		for k, x := range ia.V {
			n, err := toNumber(x)
			if err != nil {
				out.V[k] = *err
				continue
			}
			i := int(n)
			if i < 1 || i >= len(a) {
				out.V[k] = errv(ErrValue)
				continue
			}
			if cache[i] == nil {
				cache[i] = c.toArr(c.branch(a[i]))
			}
			out.V[k] = pick(cache[i], k/ia.Cols, k%ia.Cols)
		}
		return arrOf(out)
	}
	n, err := toNumber(c.scalar(iv))
	if err != nil {
		return *err
	}
	i := int(n)
	if i < 1 || i >= len(a) {
		return errv(ErrValue)
	}
	return c.branch(a[i])
}

func lazySwitch(c *ctx, a []*node) Value {
	x := c.scalar(c.eval(a[0]))
	if x.Kind == Error {
		return x
	}
	rest := a[1:]
	for len(rest) >= 2 {
		v := c.scalar(c.eval(rest[0]))
		if v.Kind == Error {
			return v
		}
		if compare(x, v) == 0 {
			return c.branch(rest[1])
		}
		rest = rest[2:]
	}
	if len(rest) == 1 {
		return c.branch(rest[0])
	}
	return errv(ErrNA)
}

func lazyLet(c *ctx, a []*node) Value {
	if len(a)%2 == 0 {
		return errv(ErrValue)
	}
	d := *c
	d.locals = map[string]Value{}
	for k, v := range c.locals {
		d.locals[k] = v
	}
	for i := 0; i+1 < len(a); i += 2 {
		if a[i].kind != nName {
			return errv(ErrValue)
		}
		d.locals[strings.ToUpper(a[i].str)] = d.eval(a[i+1])
	}
	return d.eval(a[len(a)-1])
}

// logic is AND, OR and XOR over values and references: logical values
// and numbers count; text and blanks in references are skipped.
func logic(c *ctx, a []Value, op string) Value {
	seen := false
	acc := op == "and"
	trues := 0
	var bad *Value
	for _, v := range a {
		if isMissing(v) {
			continue
		}
		c.each(v, func(x Value, fromRef bool) bool {
			var b bool
			switch x.Kind {
			case Error:
				bad = &x
				return false
			case Bool:
				b = x.Bool
			case Number:
				b = x.Num != 0
			case Text:
				if fromRef {
					return true
				}
				t, err := toBool(x)
				if err != nil {
					bad = err
					return false
				}
				b = t
			default:
				if fromRef {
					return true
				}
				b = false
			}
			seen = true
			switch op {
			case "and":
				acc = acc && b
			case "or":
				acc = acc || b
			default:
				if b {
					trues++
				}
			}
			return true
		})
		if bad != nil {
			return *bad
		}
	}
	if !seen {
		return errv(ErrValue)
	}
	if op == "xor" {
		return boolean(trues%2 == 1)
	}
	return boolean(acc)
}

// hasName tells whether the workbook defines a name.
func (c *ctx) hasName(name string) bool {
	_, ok := c.e.book.Name(name, c.sheet)
	return ok
}

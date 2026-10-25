package engine

import (
	"math"
	"strings"
	"time"

	"github.com/eemax/sheets/internal/grid"
)

// Book is the workbook the engine reads.
type Book interface {
	// Cell gives a cell's content: its value (for a formula, the result
	// the file stored, if any) and its formula ("" for a constant or a
	// blank).
	Cell(sheet string, row, col int) (v Value, formula string)
	// Extent bounds a sheet's cells: references to whole columns or rows
	// read no further. ok is false when there is no such sheet.
	Extent(sheet string) (rows, cols int, ok bool)
	// Name gives the formula a defined name stands for, as seen from a
	// sheet (a name scoped to that sheet first).
	Name(name, sheet string) (string, bool)
}

// ArrayBook is a Book that knows array formulas.
type ArrayBook interface {
	// ArrayAt gives the array formula covering a cell: its anchor and the
	// range its result fills. ok is false outside array formulas.
	ArrayAt(sheet string, row, col int) (anchorRow, anchorCol int, fill Area, ok bool)
}

// Engine computes formulas of a Book, each once, until Reset.
type Engine struct {
	book     Book
	Date1904 bool
	// Now is the time NOW and TODAY give.
	Now func() time.Time
	// Override, when set, gives a formula cell's value instead of
	// computing it (results taken as stored, say).
	Override func(sheet string, row, col int) (Value, bool)

	memo   map[cellKey]Value
	arrays map[cellKey]*Arr
	busy   map[cellKey]bool
	parsed map[string]parsed
	// cycle is set when a circular reference was met.
	Cycle bool
}

type parsed struct {
	n   *node
	err error
}

type cellKey struct {
	sheet    string
	row, col int
}

func key(sheet string, row, col int) cellKey {
	return cellKey{strings.ToLower(sheet), row, col}
}

// New returns an engine over a book.
func New(b Book) *Engine {
	return &Engine{book: b, Now: time.Now, memo: map[cellKey]Value{}, arrays: map[cellKey]*Arr{}, busy: map[cellKey]bool{}, parsed: map[string]parsed{}}
}

// Reset forgets computed results, after the book changed.
func (e *Engine) Reset() {
	e.memo = map[cellKey]Value{}
	e.arrays = map[cellKey]*Arr{}
	e.busy = map[cellKey]bool{}
}

// ErrCycle is the value of a formula in a circular reference.
const ErrCycle = "#CYCLE!"

// Unsupported is the error value of a formula the engine cannot compute
// (a function it does not have, a link to another workbook); Str holds
// "#UNSUPPORTED" and Note says why.
const ErrUnsupported = "#UNSUPPORTED"

// Result is a computed cell.
type Result struct {
	V Value
	// Why says what the engine could not do, for ErrUnsupported.
	Why string
}

// Compute gives the value of a formula cell, as Excel shows it.
func (e *Engine) Compute(sheet string, row, col int) Value {
	k := key(sheet, row, col)
	if v, ok := e.memo[k]; ok {
		return v
	}
	if e.Override != nil {
		if v, ok := e.Override(sheet, row, col); ok {
			e.memo[k] = v
			return v
		}
	}
	_, text := e.book.Cell(sheet, row, col)
	if text == "" {
		v, _ := e.book.Cell(sheet, row, col)
		return v
	}
	if e.busy[k] {
		e.Cycle = true
		return errv(ErrCycle)
	}
	e.busy[k] = true
	defer delete(e.busy, k)
	v := e.evalCell(sheet, row, col, text)
	e.memo[k] = v
	return v
}

// ArrayOf gives the whole result of an array formula computed so far.
func (e *Engine) ArrayOf(sheet string, row, col int) (*Arr, bool) {
	a, ok := e.arrays[key(sheet, row, col)]
	return a, ok
}

// Eval computes a formula as if it stood in a cell, without storing it.
func (e *Engine) Eval(sheet string, row, col int, text string) Value {
	return e.evalCell(sheet, row, col, text)
}

func (e *Engine) parse(text string) (*node, error) {
	if p, ok := e.parsed[text]; ok {
		return p.n, p.err
	}
	n, err := Parse(text)
	e.parsed[text] = parsed{n, err}
	return n, err
}

func (e *Engine) evalCell(sheet string, row, col int, text string) Value {
	n, err := e.parse(text)
	if err != nil {
		return unsupported(err.Error())
	}
	var fill Area
	array := false
	if ab, ok := e.book.(ArrayBook); ok {
		if ar, ac, f, ok := ab.ArrayAt(sheet, row, col); ok && ar == row && ac == col {
			array, fill = true, f
		}
	}
	var failed *Value
	c := &ctx{e: e, sheet: sheet, row: row, col: col, arr: array, failed: &failed}
	v := c.eval(n)
	if failed != nil {
		return *failed
	}
	if v.Kind == Error && v.Str == ErrUnsupported {
		return v
	}
	if array {
		a := c.toArr(v)
		e.arrays[key(sheet, row, col)] = a
		_ = fill
		return finalScalar(a.At(0, 0))
	}
	return finalScalar(c.scalar(v))
}

// finalScalar is a formula's result as a cell holds it: a blank read
// through a reference shows as 0.
func finalScalar(v Value) Value {
	if v.Kind == Empty {
		return num(0)
	}
	return v
}

func unsupported(why string) Value { return Value{Kind: Error, Str: ErrUnsupported}.withWhy(why) }

func (v Value) withWhy(why string) Value {
	v.Arr = &Arr{V: []Value{{Kind: Text, Str: why}}}
	return v
}

// Why says why a value is ErrUnsupported.
func (v Value) Why() string {
	if v.Kind == Error && v.Str == ErrUnsupported && v.Arr != nil && len(v.Arr.V) > 0 {
		return v.Arr.V[0].Str
	}
	return ""
}

// cellValue is what a reference to a cell reads: a constant, or a
// formula's result, computed when needed.
func (e *Engine) cellValue(sheet string, row, col int) Value {
	v, text := e.book.Cell(sheet, row, col)
	if text != "" {
		return e.Compute(sheet, row, col)
	}
	if ab, ok := e.book.(ArrayBook); ok {
		if ar, ac, fill, ok := ab.ArrayAt(sheet, row, col); ok && (ar != row || ac != col) {
			e.Compute(sheet, ar, ac)
			if a := e.arrays[key(sheet, ar, ac)]; a != nil {
				i, j := row-fill.R1, col-fill.C1
				if i < a.Rows && j < a.Cols {
					return finalScalar(a.At(i, j))
				}
				if a.Rows == 1 && a.Cols == 1 {
					return finalScalar(a.At(0, 0))
				}
				return errv(ErrNA)
			}
		}
	}
	return v
}

// ctx evaluates one formula: its cell (for relative names, implicit
// intersection and ROW()), and whether it is evaluated as an array.
type ctx struct {
	e        *Engine
	sheet    string
	row, col int
	arr      bool
	locals   map[string]Value // LET names
	// failed holds the first thing met that the engine cannot compute,
	// shared by the whole formula.
	failed **Value
}

// cell reads a cell for the formula, noting a result the engine cannot
// know.
func (c *ctx) cell(sheet string, row, col int) Value {
	v := c.e.cellValue(sheet, row, col)
	if v.Kind == Error && (v.Str == ErrUnsupported || v.Str == ErrCycle) {
		c.fail(v)
	}
	return v
}

func (c *ctx) child(arr bool) *ctx {
	d := *c
	d.arr = arr
	return &d
}

func (c *ctx) eval(n *node) Value {
	switch n.kind {
	case nNum:
		return num(n.num)
	case nStr:
		return str(n.str)
	case nBool:
		return boolean(n.b)
	case nErr:
		return errv(n.str)
	case nMissing:
		return Value{Kind: Empty, Str: missing}
	case nArray:
		return arrOf(n.arr)
	case nRef:
		sheet := n.sheet
		if sheet == "" {
			sheet = c.sheet
		} else if _, _, ok := c.e.book.Extent(sheet); !ok {
			return errv(ErrRef)
		}
		return Value{Kind: Ref, Ref: &RefV{Sheet: sheet, Areas: []Area{n.area}}}
	case nName:
		return c.name(n)
	case nUnary:
		return c.lift1(c.eval(n.args[0]), func(v Value) Value {
			x, err := toNumber(v)
			if err != nil {
				return c.known(*err)
			}
			return num(-x)
		})
	case nPercent:
		return c.lift1(c.eval(n.args[0]), func(v Value) Value {
			x, err := toNumber(v)
			if err != nil {
				return *err
			}
			return num(x / 100)
		})
	case nBinary:
		switch n.str {
		case ":":
			return c.rangeOp(c.eval(n.args[0]), c.eval(n.args[1]))
		case " ":
			return c.intersect(c.eval(n.args[0]), c.eval(n.args[1]))
		}
		return c.binary(n.str, c.eval(n.args[0]), c.eval(n.args[1]))
	case nUnion:
		var areas []Area
		sheet := ""
		for _, a := range n.args {
			v := c.eval(a)
			if v.Kind == Error {
				return v
			}
			if v.Kind != Ref {
				return errv(ErrValue)
			}
			if sheet != "" && !strings.EqualFold(sheet, v.Ref.Sheet) {
				return errv(ErrValue)
			}
			sheet = v.Ref.Sheet
			areas = append(areas, v.Ref.Areas...)
		}
		return Value{Kind: Ref, Ref: &RefV{Sheet: sheet, Areas: areas}}
	case nCall:
		return c.call(n)
	}
	return errv(ErrValue)
}

// missing marks an argument left out.
const missing = "\x00missing"

func isMissing(v Value) bool { return v.Kind == Empty && v.Str == missing }

func (c *ctx) name(n *node) Value {
	up := strings.ToUpper(n.str)
	if n.sheet == "" {
		if v, ok := c.locals[up]; ok {
			return v
		}
	}
	sheet := c.sheet
	if n.sheet != "" {
		sheet = n.sheet
	}
	text, ok := c.e.book.Name(n.str, sheet)
	if !ok {
		return errv(ErrName)
	}
	p, err := c.e.parse(text)
	if err != nil {
		return c.fail(unsupported("the name " + n.str + " " + err.Error()))
	}
	return c.eval(p)
}

// ---- values of references ----

// bounded is an area clipped to its sheet's cells.
func (c *ctx) bounded(sheet string, a Area) Area {
	rows, cols, ok := c.e.book.Extent(sheet)
	if !ok {
		return Area{1, 1, 0, 0}
	}
	// Whole columns and rows stop at the sheet's last cell.
	if a.R2 == grid.MaxRows && rows < a.R2 {
		a.R2 = max(rows, a.R1-1)
	}
	if a.C2 == grid.MaxCols && cols < a.C2 {
		a.C2 = max(cols, a.C1-1)
	}
	return a
}

// scalar turns a value into one value: a reference to one cell reads it;
// a reference to several, in a formula not evaluated as an array, takes
// the cell in the formula's row or column (implicit intersection); an
// array gives its first value.
func (c *ctx) scalar(v Value) Value {
	switch v.Kind {
	case Ref:
		r := v.Ref
		if len(r.Areas) != 1 {
			return errv(ErrValue)
		}
		a := r.Areas[0]
		if a.R1 == a.R2 && a.C1 == a.C2 {
			return c.cell(r.Sheet, a.R1, a.C1)
		}
		if c.arr {
			// In an array formula a range in a scalar slot is an array;
			// the caller lifts it. Here only its first value fits.
			return c.cell(r.Sheet, a.R1, a.C1)
		}
		// Implicit intersection: the cell of a column in the formula's
		// row, or of a row in its column (whatever the sheet).
		switch {
		case a.C1 == a.C2 && c.row >= a.R1 && c.row <= a.R2:
			return c.cell(r.Sheet, c.row, a.C1)
		case a.R1 == a.R2 && c.col >= a.C1 && c.col <= a.C2:
			return c.cell(r.Sheet, a.R1, c.col)
		}
		return errv(ErrValue)
	case Array:
		if len(v.Arr.V) == 0 {
			return errv(ErrValue)
		}
		return c.scalar(v.Arr.V[0])
	}
	return v
}

// toArr reads a value as an array: a reference's cells (bounded to its
// sheet), an array as is, a scalar as a 1×1 array.
func (c *ctx) toArr(v Value) *Arr {
	switch v.Kind {
	case Array:
		return v.Arr
	case Ref:
		r := v.Ref
		if len(r.Areas) != 1 {
			a := newArr(1, 1)
			a.V[0] = errv(ErrValue)
			return a
		}
		ar := c.bounded(r.Sheet, r.Areas[0])
		rows, cols := ar.Rows(), ar.Cols()
		if rows <= 0 || cols <= 0 {
			// A whole column of an empty sheet: one blank.
			a := newArr(1, 1)
			return a
		}
		a := newArr(rows, cols)
		for i := 0; i < rows; i++ {
			for j := 0; j < cols; j++ {
				a.V[i*cols+j] = c.cell(r.Sheet, ar.R1+i, ar.C1+j)
			}
		}
		return a
	}
	a := newArr(1, 1)
	a.V[0] = v
	return a
}

// lifts tells whether a value applies element by element where one value
// is expected: an array always (an array constant works so in any
// formula), a reference to several cells only in array evaluation (in a
// normal formula it gives the cell in the formula's row or column).
func (c *ctx) lifts(v Value) bool {
	if v.Kind == Array {
		return v.Arr.Rows*v.Arr.Cols != 1
	}
	return c.arr && isMulti(v)
}

// intersects tells whether implicit intersection finds a cell of a
// reference: a column the formula's row crosses, or a row its column
// crosses.
func (c *ctx) intersects(v Value) bool {
	if v.Kind != Ref || len(v.Ref.Areas) != 1 {
		return false
	}
	a := v.Ref.Areas[0]
	return a.C1 == a.C2 && c.row >= a.R1 && c.row <= a.R2 || a.R1 == a.R2 && c.col >= a.C1 && c.col <= a.C2
}

// isMulti tells whether a value holds several values (an array, or a
// reference to more than one cell).
func isMulti(v Value) bool {
	switch v.Kind {
	case Array:
		return v.Arr.Rows*v.Arr.Cols != 1
	case Ref:
		if len(v.Ref.Areas) != 1 {
			return true
		}
		a := v.Ref.Areas[0]
		return a.R1 != a.R2 || a.C1 != a.C2
	}
	return false
}

// each walks the values of an argument that takes references: the cells
// of a reference (bounded to their sheet, row by row), an array's values,
// or the value itself. fromRef tells the callback whether a value came
// from a cell (where text and logical values are skipped by SUM and the
// like) or was given directly.
func (c *ctx) each(v Value, fn func(x Value, fromRef bool) bool) {
	switch v.Kind {
	case Ref:
		for _, area := range v.Ref.Areas {
			a := c.bounded(v.Ref.Sheet, area)
			for r := a.R1; r <= a.R2; r++ {
				for col := a.C1; col <= a.C2; col++ {
					if !fn(c.cell(v.Ref.Sheet, r, col), true) {
						return
					}
				}
			}
		}
	case Array:
		for _, x := range v.Arr.V {
			if !fn(x, true) {
				return
			}
		}
	default:
		fn(v, false)
	}
}

// ---- operators ----

func (c *ctx) lift1(v Value, fn func(Value) Value) Value {
	if c.lifts(v) {
		a := c.toArr(v)
		out := newArr(a.Rows, a.Cols)
		for i, x := range a.V {
			out.V[i] = fn(x)
		}
		return arrOf(out)
	}
	v = c.scalar(v)
	if v.Kind == Error {
		return v
	}
	return fn(v)
}

// lift2 applies fn element by element: in an array evaluation, arrays and
// ranges broadcast against each other (a single row or column stretches;
// where one is short, #N/A); otherwise both sides are single values.
func (c *ctx) lift2(a, b Value, fn func(x, y Value) Value) Value {
	if c.lifts(a) || c.lifts(b) {
		x, y := c.toArr(a), c.toArr(b)
		rows, cols := max(x.Rows, y.Rows), max(x.Cols, y.Cols)
		out := newArr(rows, cols)
		for i := 0; i < rows; i++ {
			for j := 0; j < cols; j++ {
				out.V[i*cols+j] = fn(pick(x, i, j), pick(y, i, j))
			}
		}
		return arrOf(out)
	}
	return fn(c.scalar(a), c.scalar(b))
}

// pick is the value of a broadcast array at (i, j).
func pick(a *Arr, i, j int) Value {
	if a.Rows == 1 {
		i = 0
	}
	if a.Cols == 1 {
		j = 0
	}
	if i >= a.Rows || j >= a.Cols {
		return errv(ErrNA)
	}
	return a.At(i, j)
}

func (c *ctx) binary(op string, a, b Value) Value {
	return c.lift2(a, b, func(x, y Value) Value { return c.known(binop(op, x, y)) })
}

func binop(op string, x, y Value) Value {
	switch op {
	case "&":
		if x.Kind == Error {
			return x
		}
		if y.Kind == Error {
			return y
		}
		s, err := toText(x)
		if err != nil {
			return *err
		}
		t, err := toText(y)
		if err != nil {
			return *err
		}
		return str(s + t)
	case "=", "<>", "<", ">", "<=", ">=":
		if x.Kind == Error {
			return x
		}
		if y.Kind == Error {
			return y
		}
		r := compare(x, y)
		switch op {
		case "=":
			return boolean(r == 0)
		case "<>":
			return boolean(r != 0)
		case "<":
			return boolean(r < 0)
		case ">":
			return boolean(r > 0)
		case "<=":
			return boolean(r <= 0)
		}
		return boolean(r >= 0)
	}
	p, err := toNumber(x)
	if err != nil {
		return *err
	}
	q, err := toNumber(y)
	if err != nil {
		return *err
	}
	switch op {
	case "+":
		return numOrErr(addRound(p, q))
	case "-":
		return numOrErr(addRound(p, -q))
	case "*":
		return numOrErr(p * q)
	case "/":
		if q == 0 {
			return errv(ErrDiv0)
		}
		return numOrErr(p / q)
	case "^":
		if p == 0 && q == 0 {
			return errv(ErrNum)
		}
		if p == 0 && q < 0 {
			return errv(ErrDiv0)
		}
		r := math.Pow(p, q)
		if math.IsNaN(r) {
			return errv(ErrNum)
		}
		return numOrErr(r)
	}
	return errv(ErrValue)
}

// addRound adds as Excel does: a sum that comes within 15 significant
// digits of zero relative to its terms is zero (0.1+0.2-0.3 is 0).
func addRound(p, q float64) float64 {
	s := p + q
	if s != 0 && math.Abs(s) < math.Max(math.Abs(p), math.Abs(q))*1e-15 {
		return 0
	}
	return s
}

// rangeOp is A1:B2 between references (or functions giving them): the
// smallest area holding both.
func (c *ctx) rangeOp(a, b Value) Value {
	if a.Kind == Error {
		return a
	}
	if b.Kind == Error {
		return b
	}
	if a.Kind != Ref || b.Kind != Ref || len(a.Ref.Areas) != 1 || len(b.Ref.Areas) != 1 {
		return errv(ErrValue)
	}
	if !strings.EqualFold(a.Ref.Sheet, b.Ref.Sheet) {
		return errv(ErrValue)
	}
	x, y := a.Ref.Areas[0], b.Ref.Areas[0]
	ar := Area{min(x.R1, y.R1), min(x.C1, y.C1), max(x.R2, y.R2), max(x.C2, y.C2)}
	return Value{Kind: Ref, Ref: &RefV{Sheet: a.Ref.Sheet, Areas: []Area{ar}}}
}

// intersect is the space operator: the cells two references share.
func (c *ctx) intersect(a, b Value) Value {
	if a.Kind == Error {
		return a
	}
	if b.Kind == Error {
		return b
	}
	if a.Kind != Ref || b.Kind != Ref || !strings.EqualFold(a.Ref.Sheet, b.Ref.Sheet) {
		return errv(ErrValue)
	}
	var out []Area
	for _, x := range a.Ref.Areas {
		for _, y := range b.Ref.Areas {
			ar := Area{max(x.R1, y.R1), max(x.C1, y.C1), min(x.R2, y.R2), min(x.C2, y.C2)}
			if ar.R1 <= ar.R2 && ar.C1 <= ar.C2 {
				out = append(out, ar)
			}
		}
	}
	if len(out) == 0 {
		return errv(ErrNull)
	}
	return Value{Kind: Ref, Ref: &RefV{Sheet: a.Ref.Sheet, Areas: out}}
}

func floorDiv(a, b int) int {
	q := a / b
	if (a%b != 0) && ((a < 0) != (b < 0)) {
		q--
	}
	return q
}

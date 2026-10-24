package formula

import (
	"strconv"
	"strings"

	"github.com/eemax/sheets/internal/grid"
)

// Kinds is a set of the types a value can have. A calculator that returns
// results as text (as excelize does) cannot say whether "005" is text or
// 5 a number, or "1" the boolean TRUE; what a formula can produce, read
// off its text and the cells it reads, can.
type Kinds uint8

const (
	KNumber Kinds = 1 << iota
	KBool
	// KText is text that never looks like a number or a boolean
	// ("North"); KNumText is text that may ("007", "TRUE").
	KText
	KNumText
	// KAny marks a value that can be anything.
	KAny = KNumber | KBool | KText | KNumText
)

// KindsOf classifies a constant cell value; empty cells read as 0.
func KindsOf(c grid.Cell) Kinds {
	switch c.Kind {
	case grid.Number, grid.Empty:
		return KNumber
	case grid.Bool:
		return KBool
	case grid.Text:
		return TextKinds(c.Text)
	}
	return 0 // errors travel apart
}

// TextKinds classifies a text value. Excel reads text as a number where it
// can ("1,200", "$5", "15%", "2026-07-01", " 42 "), more widely than a
// parser of numbers: text with any digit counts as possibly numeric.
func TextKinds(s string) Kinds {
	if looksScalar(s) || strings.ContainsAny(s, "0123456789") {
		return KNumText
	}
	return KText
}

// InputKinds is every kind of value a formula reads: its references, the
// names it uses, and the text it holds.
func InputKinds(f string, ty Typer) Kinds {
	var k Kinds
	d := DepsOf(f)
	for _, r := range d.Refs {
		k |= ty.RangeKinds(r.Sheet, r.Range)
	}
	for _, n := range d.Names {
		if nk, ok := ty.NameKinds(n); ok {
			k |= nk
		}
	}
	for _, t := range lex(f) {
		if strings.HasPrefix(t.text, `"`) && len(t.text) >= 2 {
			k |= TextKinds(strings.ReplaceAll(t.text[1:len(t.text)-1], `""`, `"`))
		}
	}
	if len(d.Tables) > 0 || d.Volatile || d.External {
		k |= KAny
	}
	return k
}

func looksScalar(s string) bool {
	if _, err := strconv.ParseFloat(strings.TrimSpace(s), 64); err == nil {
		return true
	}
	u := strings.ToUpper(strings.TrimSpace(s))
	return u == "TRUE" || u == "FALSE"
}

// Typer answers what the cells and names a formula reads hold.
type Typer interface {
	// RangeKinds is every kind the cells of a range hold (or produce);
	// sheet "" is the formula's own.
	RangeKinds(sheet string, rg grid.Range) Kinds
	// NameKinds is what a workbook name stands for; ok is false for a
	// word that names nothing.
	NameKinds(name string) (Kinds, bool)
}

// ResultKinds is every kind of value a formula can produce.
func ResultKinds(f string, ty Typer) Kinds {
	p := &typeParser{ty: ty}
	for _, t := range lex(f) {
		if strings.TrimSpace(t.text) != "" || strings.HasPrefix(t.text, `"`) {
			p.ts = append(p.ts, t)
		}
	}
	k := p.expr()
	if p.i < len(p.ts) || p.bad {
		return KAny
	}
	return k
}

type typeParser struct {
	ts  []token
	i   int
	ty  Typer
	bad bool
	// quiet is set while reading arguments whose kinds cannot matter (of
	// SUM, which gives a number whatever it sums): references in them are
	// not looked up, which on large sheets is most of the work.
	quiet int
}

func (p *typeParser) peek(n int) string {
	if p.i+n < len(p.ts) {
		return p.ts[p.i+n].text
	}
	return ""
}

func (p *typeParser) next() token {
	if p.i >= len(p.ts) {
		p.bad = true
		return token{}
	}
	t := p.ts[p.i]
	p.i++
	return t
}

func (p *typeParser) expr() Kinds {
	k := p.concat()
	for {
		switch op := p.peek(0); {
		case op == "=" || op == "<" || op == ">":
			p.i++
			if p.peek(0) == "=" || op == "<" && p.peek(0) == ">" {
				p.i++
			}
			p.concat()
			k = KBool
		default:
			return k
		}
	}
}

func (p *typeParser) concat() Kinds {
	k := p.additive()
	for p.peek(0) == "&" {
		p.i++
		p.additive()
		k = KText | KNumText
	}
	return k
}

func (p *typeParser) additive() Kinds {
	k := p.term()
	for p.peek(0) == "+" || p.peek(0) == "-" {
		p.i++
		p.term()
		k = KNumber
	}
	return k
}

func (p *typeParser) term() Kinds {
	k := p.power()
	for p.peek(0) == "*" || p.peek(0) == "/" {
		p.i++
		p.power()
		k = KNumber
	}
	return k
}

func (p *typeParser) power() Kinds {
	k := p.unary()
	for p.peek(0) == "^" {
		p.i++
		p.unary()
		k = KNumber
	}
	return k
}

func (p *typeParser) unary() Kinds {
	if p.peek(0) == "-" || p.peek(0) == "+" {
		p.i++
		p.unary()
		return KNumber
	}
	k := p.primary()
	for p.peek(0) == "%" {
		p.i++
		k = KNumber
	}
	return k
}

func (p *typeParser) primary() Kinds {
	t := p.next()
	switch {
	case p.bad:
		return KAny
	case t.text == "(":
		k := p.expr()
		if p.next().text != ")" {
			p.bad = true
		}
		return k
	case t.text == "{":
		var k Kinds
		for p.peek(0) != "}" && !p.bad {
			if s := p.peek(0); s == "," || s == ";" {
				p.i++
				continue
			}
			k |= p.unary()
		}
		p.next()
		return k
	case strings.HasPrefix(t.text, `"`):
		return TextKinds(strings.ReplaceAll(t.text[1:len(t.text)-1], `""`, `"`))
	case t.text == "#":
		p.i = errorLiteral(p.ts, p.i-1) + 1
		return 0
	case strings.HasPrefix(t.text, "["), strings.HasPrefix(t.text, "'") && strings.Contains(t.text, "["):
		p.skipReference()
		return KAny // structured references, other workbooks
	case strings.HasPrefix(t.text, "'"):
		return p.reference(strings.ReplaceAll(t.text[1:len(t.text)-1], "''", "'"))
	case !t.word:
		p.bad = true
		return KAny
	case t.call:
		return p.call(strings.ToUpper(bare(t.text)))
	}
	w := t.text
	if c := w[0]; c >= '0' && c <= '9' || c == '.' {
		if p.peek(0) == ":" { // a row range, 3:5
			p.i--
			return p.reference("")
		}
		return KNumber
	}
	switch strings.ToUpper(w) {
	case "TRUE", "FALSE":
		return KBool
	}
	if p.peek(0) == "!" || p.peek(0) == "[" {
		if p.peek(0) == "[" {
			p.skipReference()
			return KAny
		}
		return p.reference(w)
	}
	if ref, ok := parseRef(w); ok && (ref.kind == cellRef || p.peek(0) == ":") {
		p.i--
		return p.reference("")
	}
	if p.quiet > 0 {
		return KAny
	}
	k, ok := p.ty.NameKinds(bare(w))
	if !ok {
		return KAny
	}
	return k
}

// reference reads A1, A1:B9, A:A or 3:3 (after Sheet! when sheet is set)
// and returns the kinds its cells hold.
func (p *typeParser) reference(sheet string) Kinds {
	if sheet != "" && p.next().text != "!" {
		p.bad = true
		return KAny
	}
	first := p.next()
	a, ok := parseRef(strings.TrimSpace(first.text))
	if !ok {
		// Sheet!Name: a name scoped to a sheet.
		if k, found := p.ty.NameKinds(bare(first.text)); found {
			return k
		}
		return KAny
	}
	rg := bounds(a)
	if p.peek(0) == ":" {
		p.i++
		b, ok := parseRef(p.next().text)
		if !ok {
			p.bad = true
			return KAny
		}
		rg = rg.Union(bounds(b))
	}
	if p.quiet > 0 {
		return KAny
	}
	return p.ty.RangeKinds(sheet, rg)
}

func (p *typeParser) skipReference() {
	for p.i < len(p.ts) {
		s := p.peek(0)
		if s == "," || s == ")" || s == "&" || s == "+" || s == "-" || s == "*" || s == "/" || s == "^" || s == "=" || s == "<" || s == ">" || s == ";" || s == "}" {
			return
		}
		p.i++
	}
}

func bounds(r part) grid.Range {
	switch r.kind {
	case colRef:
		return grid.Range{R1: 1, C1: r.col, R2: grid.MaxRows, C2: r.col}
	case rowRef:
		return grid.Range{R1: r.row, C1: 1, R2: r.row, C2: grid.MaxCols}
	}
	return grid.Range{R1: r.row, C1: r.col, R2: r.row, C2: r.col}
}

// call reads a function's arguments (the name is read, the "(" next) and
// returns what the function gives.
func (p *typeParser) call(fn string) Kinds {
	if p.next().text != "(" {
		p.bad = true
		return KAny
	}
	fixed := textFns[fn] || boolFns[fn] || numberFns[fn] || fn == "MAX" || fn == "MIN" || fn == "MAXA" || fn == "MINA" || fn == "LARGE" || fn == "SMALL"
	if fixed {
		p.quiet++
		defer func() { p.quiet-- }()
	}
	var args []Kinds
	for p.peek(0) != ")" && !p.bad {
		if p.peek(0) == "," {
			args = append(args, KNumber) // an empty argument
			p.i++
			continue
		}
		args = append(args, p.expr())
		if p.peek(0) == "," {
			p.i++
			if p.peek(0) == ")" {
				args = append(args, KNumber)
			}
		}
	}
	if p.next().text != ")" {
		p.bad = true
	}
	arg := func(i int) Kinds {
		if i < len(args) {
			return args[i]
		}
		return 0
	}
	any := func(from, step int) Kinds {
		var k Kinds
		for i := from; i < len(args); i += step {
			k |= args[i]
		}
		return k
	}
	switch {
	case textFns[fn]:
		return KText | KNumText
	case boolFns[fn]:
		return KBool
	case numberFns[fn]:
		return KNumber
	}
	switch fn {
	case "IF":
		if len(args) < 3 {
			return arg(1) | KBool // no else: FALSE
		}
		return arg(1) | arg(2)
	case "IFERROR", "IFNA":
		return arg(0) | arg(1)
	case "IFS":
		return any(1, 2)
	case "CHOOSE":
		return any(1, 1)
	case "SWITCH":
		k := any(2, 2)
		if len(args)%2 == 0 {
			k |= args[len(args)-1] // the default
		}
		return k
	case "INDEX":
		return arg(0)
	case "XLOOKUP":
		return arg(2) | arg(3)
	case "LOOKUP":
		if len(args) > 2 {
			return arg(2)
		}
		return arg(1)
	case "VLOOKUP", "HLOOKUP":
		// The table's cells; narrower would need the column picked.
		return arg(1)
	case "MAX", "MIN", "MAXA", "MINA", "LARGE", "SMALL":
		return KNumber
	}
	return KAny
}

var (
	textFns = setOf(`ASC BAHTTEXT CHAR CLEAN CONCAT CONCATENATE DBCS DOLLAR FIXED LEFT LEFTB LOWER MID MIDB
		PHONETIC PROPER REPLACE REPLACEB REPT RIGHT RIGHTB ROMAN SUBSTITUTE T TEXT TEXTAFTER TEXTBEFORE
		TEXTJOIN TRIM UNICHAR UPPER VALUETOTEXT ARRAYTOTEXT ENCODEURL DEC2BIN DEC2HEX DEC2OCT BIN2HEX
		BIN2OCT HEX2BIN HEX2OCT OCT2BIN OCT2HEX BASE ADDRESS FORMULATEXT`)
	boolFns = setOf(`AND OR NOT XOR TRUE FALSE ISBLANK ISERR ISERROR ISEVEN ISODD ISFORMULA ISLOGICAL ISNA
		ISNONTEXT ISNUMBER ISREF ISTEXT EXACT`)
	numberFns = setOf(`ABS ACOS ACOSH ASIN ASINH ATAN ATAN2 ATANH AVEDEV AVERAGE AVERAGEA AVERAGEIF AVERAGEIFS
		CEILING COMBIN COS COSH COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS DATE DATEDIF DATEVALUE DAY DAYS
		DAYS360 DEGREES EDATE EOMONTH EVEN EXP FACT FLOOR FV GCD HOUR INT IPMT IRR LCM LEN LN LOG LOG10
		MATCH MAXIFS MEDIAN MINIFS MINUTE MOD MONTH MROUND NETWORKDAYS NOW NPER NPV ODD PERCENTILE PI PMT
		POWER PPMT PRODUCT PV QUARTILE QUOTIENT RADIANS RAND RANDBETWEEN RANK RATE ROUND ROUNDDOWN ROUNDUP
		ROW ROWS COLUMN COLUMNS SECOND SIGN SIN SINH SQRT STDEV STDEVA STDEVP SUBTOTAL SUM SUMIF SUMIFS
		SUMPRODUCT SUMSQ TAN TANH TIME TIMEVALUE TODAY TRUNC VAR VARA VARP WEEKDAY WEEKNUM WORKDAY XIRR
		XNPV XMATCH YEAR YEARFRAC FIND SEARCH CODE N VALUE NUMBERVALUE CORREL SLOPE INTERCEPT FORECAST MODE`)
)

// Typed reads a calculator's result, given as text, by what the formula
// can produce. ok is false when the text fits more than one type ("7"
// from a lookup that can return the number 7 or the text "7") or none.
func Typed(raw string, k Kinds) (v grid.Cell, ok bool) {
	_, err := strconv.ParseFloat(raw, 64)
	number := err == nil
	upper := strings.ToUpper(raw)
	boolean := upper == "TRUE" || upper == "FALSE" || raw == "1" || raw == "0"
	var cands []grid.Cell
	if k&KNumber != 0 && number {
		n, _ := strconv.ParseFloat(raw, 64)
		cands = append(cands, grid.Cell{Kind: grid.Number, Num: n})
	}
	if k&KBool != 0 && boolean {
		cands = append(cands, grid.Cell{Kind: grid.Bool, Bool: upper == "TRUE" || raw == "1"})
	}
	if k&KNumText != 0 || k&KText != 0 && !looksScalar(raw) {
		cands = append(cands, grid.Cell{Kind: grid.Text, Text: raw})
	}
	if len(cands) != 1 {
		return grid.Cell{}, false
	}
	return cands[0], true
}

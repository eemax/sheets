package engine

import (
	"errors"
	"strconv"
	"strings"
	"unicode"
	"unicode/utf8"

	"github.com/eemax/sheets/internal/grid"
)

// The parser turns formula text (without its leading =) into a tree of
// nodes. Excel's operators bind, from loosest: comparisons, &, + -, * /, ^,
// %, negation, then the reference operators: , (union, inside
// parentheses), space (intersection) and : (range).

type nodeKind uint8

const (
	nNum nodeKind = iota
	nStr
	nBool
	nErr
	nArray   // {1,2;3,4}
	nRef     // a cell or area, with its sheet
	nName    // a defined name, or a LET/LAMBDA parameter
	nCall    // a function call
	nUnary   // - or +
	nPercent // x%
	nBinary  // arithmetic, &, comparisons, and : (range) and ' ' (intersection)
	nUnion   // (A1,B2)
	nMissing // an argument left out: IF(A1,,1)
)

type node struct {
	kind  nodeKind
	num   float64
	str   string // text, error code, operator, function or name
	b     bool
	arr   *Arr
	sheet string // nRef: "" for the formula's own sheet
	area  Area
	args  []*node // operands or arguments
}

// ErrSyntax is returned for text that is not a formula the parser reads.
var ErrSyntax = errors.New("formula not understood")

type tokKind uint8

const (
	tEOF tokKind = iota
	tNum
	tStr
	tBool
	tErr
	tRef
	tName
	tFunc // a name followed by (
	tOp
	tComma
	tSemi
	tLParen
	tRParen
	tLBrace
	tRBrace
	tSpace // whitespace between two operands: the intersection operator
)

type token struct {
	kind  tokKind
	text  string
	num   float64
	b     bool
	sheet string
	area  Area
	spill bool // a reference with the spill operator: D2#
}

// Parse reads a formula.
func Parse(text string) (*node, error) {
	toks, err := lex(strings.TrimPrefix(text, "="))
	if err != nil {
		return nil, err
	}
	p := &parser{toks: toks}
	n, err := p.expr(0)
	if err != nil {
		return nil, err
	}
	if p.peek().kind != tEOF {
		return nil, ErrSyntax
	}
	return n, nil
}

// ---- lexer ----

func lex(s string) ([]token, error) {
	var out []token
	i := 0
	for i < len(s) {
		c := s[i]
		switch {
		case c == ' ' || c == '\n' || c == '\r' || c == '\t' || c == 0xA0:
			j := i
			for j < len(s) && (s[j] == ' ' || s[j] == '\n' || s[j] == '\r' || s[j] == '\t' || s[j] == 0xA0) {
				j++
			}
			// Whitespace is the intersection operator between two
			// operands; elsewhere it means nothing.
			if len(out) > 0 && endsOperand(out[len(out)-1]) && j < len(s) && startsOperand(s[j:]) {
				out = append(out, token{kind: tSpace})
			}
			i = j
		case c == '"':
			j := i + 1
			var b strings.Builder
			for {
				if j >= len(s) {
					return nil, ErrSyntax
				}
				if s[j] == '"' {
					if j+1 < len(s) && s[j+1] == '"' {
						b.WriteByte('"')
						j += 2
						continue
					}
					break
				}
				b.WriteByte(s[j])
				j++
			}
			out = append(out, token{kind: tStr, text: b.String()})
			i = j + 1
		case c == '#' && len(out) > 0 && out[len(out)-1].kind == tRef && !out[len(out)-1].spill:
			// D2#: the cells D2's dynamic array spills into.
			out[len(out)-1].spill = true
			i++
		case c == '#':
			code, ok := errorAt(s[i:])
			if !ok {
				return nil, ErrSyntax
			}
			out = append(out, token{kind: tErr, text: code})
			i += len(code)
		case c == '(':
			out = append(out, token{kind: tLParen})
			i++
		case c == ')':
			out = append(out, token{kind: tRParen})
			i++
		case c == '{':
			out = append(out, token{kind: tLBrace})
			i++
		case c == '}':
			out = append(out, token{kind: tRBrace})
			i++
		case c == ',':
			out = append(out, token{kind: tComma})
			i++
		case c == ';':
			out = append(out, token{kind: tSemi})
			i++
		case c == '<' || c == '>':
			if i+1 < len(s) && (s[i+1] == '=' || c == '<' && s[i+1] == '>') {
				out = append(out, token{kind: tOp, text: s[i : i+2]})
				i += 2
			} else {
				out = append(out, token{kind: tOp, text: string(c)})
				i++
			}
		case strings.IndexByte("+-*/^&=%:", c) >= 0:
			out = append(out, token{kind: tOp, text: string(c)})
			i++
		case c == '\'':
			// A quoted sheet name: 'Q1 Sales'!A1.
			j := i + 1
			var b strings.Builder
			for {
				if j >= len(s) {
					return nil, ErrSyntax
				}
				if s[j] == '\'' {
					if j+1 < len(s) && s[j+1] == '\'' {
						b.WriteByte('\'')
						j += 2
						continue
					}
					break
				}
				b.WriteByte(s[j])
				j++
			}
			j++
			if j >= len(s) || s[j] != '!' {
				return nil, ErrSyntax
			}
			t, n, err := refAfterSheet(b.String(), s[j+1:])
			if err != nil {
				return nil, err
			}
			out = append(out, t)
			i = j + 1 + n
		case c == '[':
			// [1]Sheet!A1 links to another workbook; Table[Column] is a
			// table reference left untranslated.
			if len(out) > 0 && out[len(out)-1].kind == tName {
				return nil, errStructured
			}
			return nil, errExternal
		case c >= '0' && c <= '9' || c == '.':
			// A row range (1:3) or a number.
			if t, n, ok := rowRange(s[i:]); ok {
				out = append(out, t)
				i += n
				continue
			}
			j := i
			for j < len(s) && (s[j] >= '0' && s[j] <= '9' || s[j] == '.') {
				j++
			}
			if j < len(s) && (s[j] == 'e' || s[j] == 'E') {
				k := j + 1
				if k < len(s) && (s[k] == '+' || s[k] == '-') {
					k++
				}
				if k < len(s) && s[k] >= '0' && s[k] <= '9' {
					for k < len(s) && s[k] >= '0' && s[k] <= '9' {
						k++
					}
					j = k
				}
			}
			n, err := strconv.ParseFloat(s[i:j], 64)
			if err != nil {
				return nil, ErrSyntax
			}
			out = append(out, token{kind: tNum, num: n})
			i = j
		default:
			// $3:$3, $1:5: whole rows written with $.
			if c == '$' {
				if t, n, ok := rowRange(s[i:]); ok {
					out = append(out, t)
					i += n
					continue
				}
			}
			r, size := utf8.DecodeRuneInString(s[i:])
			if !(unicode.IsLetter(r) || r == '_' || r == '\\' || r == '$') {
				return nil, ErrSyntax
			}
			j := i + size
			for j < len(s) {
				r, size := utf8.DecodeRuneInString(s[j:])
				if unicode.IsLetter(r) || unicode.IsDigit(r) || r == '_' || r == '.' || r == '\\' || r == '$' || r == '?' {
					j += size
					continue
				}
				break
			}
			word := s[i:j]
			switch {
			case j < len(s) && s[j] == '!':
				t, n, err := refAfterSheet(word, s[j+1:])
				if err != nil {
					return nil, err
				}
				out = append(out, t)
				i = j + 1 + n
				continue
			case j < len(s) && s[j] == '(':
				out = append(out, token{kind: tFunc, text: bareName(word)})
				i = j
				continue
			}
			if t, n, ok := refAt(s[i:]); ok {
				out = append(out, t)
				i += n
				continue
			}
			switch strings.ToUpper(word) {
			case "TRUE":
				out = append(out, token{kind: tBool, b: true})
			case "FALSE":
				out = append(out, token{kind: tBool, b: false})
			default:
				out = append(out, token{kind: tName, text: bareName(word)})
			}
			i = j
		}
	}
	return append(out, token{kind: tEOF}), nil
}

var (
	errExternal   = errors.New("links to another workbook")
	errStructured = errors.New("uses a table reference the engine cannot place")
)

// bareName drops the prefixes files carry for newer functions (_xlfn.)
// and LAMBDA parameters (_xlpm.).
func bareName(w string) string {
	for _, p := range []string{"_xlfn.", "_xlws.", "_xlpm.", "_xlfn._xlws."} {
		if len(w) > len(p) && strings.EqualFold(w[:len(p)], p) {
			w = w[len(p):]
		}
	}
	return w
}

func endsOperand(t token) bool {
	switch t.kind {
	case tRef, tName, tRParen, tNum, tStr, tBool, tErr, tRBrace:
		return true
	}
	return false
}

func startsOperand(s string) bool {
	if s == "" {
		return false
	}
	c := s[0]
	if c == '(' || c == '\'' || c == '$' {
		return true
	}
	r, _ := utf8.DecodeRuneInString(s)
	return unicode.IsLetter(r) || c >= '0' && c <= '9'
}

func errorAt(s string) (string, bool) {
	up := strings.ToUpper(s)
	for _, code := range errorCodes {
		if strings.HasPrefix(up, code) {
			return code, true
		}
	}
	return "", false
}

// refAfterSheet reads the reference after Sheet!: a cell, an area, whole
// columns or rows. A sheet range (Sheet1:Sheet3!A1) is not supported.
func refAfterSheet(sheet, s string) (token, int, error) {
	if strings.HasPrefix(sheet, "[") {
		return token{}, 0, errExternal // '[1]May 2021'!A1
	}
	if strings.Contains(sheet, ":") {
		return token{}, 0, errors.New("reads several sheets at once")
	}
	if t, n, ok := rowRange(s); ok {
		t.sheet = sheet
		return t, n, nil
	}
	if t, n, ok := refAt(s); ok {
		t.sheet = sheet
		return t, n, nil
	}
	if strings.HasPrefix(strings.ToUpper(s), "#REF!") {
		return token{kind: tErr, text: ErrRef}, 5, nil
	}
	// A sheet-scoped name: Sheet1!Rates.
	j := 0
	for j < len(s) {
		r, size := utf8.DecodeRuneInString(s[j:])
		if unicode.IsLetter(r) || unicode.IsDigit(r) || r == '_' || r == '.' || r == '\\' {
			j += size
			continue
		}
		break
	}
	if j > 0 {
		return token{kind: tName, text: s[:j], sheet: sheet}, j, nil
	}
	return token{}, 0, ErrSyntax
}

// refAt reads a cell (A1, $A$1), an area (A1:B2) or whole columns (A:C) at
// the start of s, if one is there, ending where a name could not go on.
func refAt(s string) (token, int, bool) {
	r1, c1, n1, ok := cellAt(s)
	if ok {
		if n1 < len(s) && s[n1] == ':' {
			if r2, c2, n2, ok := cellAt(s[n1+1:]); ok {
				return token{kind: tRef, area: norm(Area{r1, c1, r2, c2})}, n1 + 1 + n2, true
			}
		}
		return token{kind: tRef, area: Area{r1, c1, r1, c1}}, n1, true
	}
	ca, na, ok := colAt(s)
	if ok && na < len(s) && s[na] == ':' {
		if cb, nb, ok := colAt(s[na+1:]); ok {
			return token{kind: tRef, area: norm(Area{1, ca, grid.MaxRows, cb})}, na + 1 + nb, true
		}
	}
	return token{}, 0, false
}

// rowRange reads whole rows (1:3, $2:$2).
func rowRange(s string) (token, int, bool) {
	ra, na, ok := rowAt(s)
	if !ok || na >= len(s) || s[na] != ':' {
		return token{}, 0, false
	}
	rb, nb, ok := rowAt(s[na+1:])
	if !ok {
		return token{}, 0, false
	}
	return token{kind: tRef, area: norm(Area{ra, 1, rb, grid.MaxCols})}, na + 1 + nb, true
}

func norm(a Area) Area {
	if a.R1 > a.R2 {
		a.R1, a.R2 = a.R2, a.R1
	}
	if a.C1 > a.C2 {
		a.C1, a.C2 = a.C2, a.C1
	}
	return a
}

// cellAt reads $A$1 at the start of s; the next character must not carry
// on a name.
func cellAt(s string) (row, col, n int, ok bool) {
	col, i, ok := colAt(s)
	if !ok {
		return 0, 0, 0, false
	}
	row, j, ok := rowAt(s[i:])
	if !ok {
		return 0, 0, 0, false
	}
	n = i + j
	if n < len(s) {
		r, _ := utf8.DecodeRuneInString(s[n:])
		if unicode.IsLetter(r) || unicode.IsDigit(r) || r == '_' || r == '.' || r == '(' {
			return 0, 0, 0, false
		}
	}
	return row, col, n, true
}

func colAt(s string) (col, n int, ok bool) {
	i := 0
	if i < len(s) && s[i] == '$' {
		i++
	}
	j := i
	for j < len(s) && j-i < 4 && (s[j] >= 'A' && s[j] <= 'Z' || s[j] >= 'a' && s[j] <= 'z') {
		j++
	}
	if j == i || j-i > 3 {
		return 0, 0, false
	}
	c, ok := grid.ColIndex(strings.ToUpper(s[i:j]))
	if !ok || c > grid.MaxCols {
		return 0, 0, false
	}
	return c, j, true
}

func rowAt(s string) (row, n int, ok bool) {
	i := 0
	if i < len(s) && s[i] == '$' {
		i++
	}
	j := i
	for j < len(s) && s[j] >= '0' && s[j] <= '9' {
		j++
	}
	if j == i || j-i > 7 {
		return 0, 0, false
	}
	r, err := strconv.Atoi(s[i:j])
	if err != nil || r < 1 || r > grid.MaxRows {
		return 0, 0, false
	}
	return r, j, true
}

// ---- parser ----

type parser struct {
	toks []token
	pos  int
}

func (p *parser) peek() token { return p.toks[p.pos] }
func (p *parser) next() token {
	t := p.toks[p.pos]
	if p.pos < len(p.toks)-1 {
		p.pos++
	}
	return t
}

// Binding powers of the infix operators.
func infixPower(t token) int {
	switch t.kind {
	case tSpace:
		return 80
	case tOp:
		switch t.text {
		case ":":
			return 90
		case "^":
			return 50
		case "*", "/":
			return 40
		case "+", "-":
			return 30
		case "&":
			return 20
		case "=", "<>", "<", ">", "<=", ">=":
			return 10
		case "%":
			return 60
		}
	}
	return -1
}

func (p *parser) expr(minPower int) (*node, error) {
	left, err := p.prefix()
	if err != nil {
		return nil, err
	}
	for {
		t := p.peek()
		power := infixPower(t)
		if power <= minPower {
			return left, nil
		}
		if t.kind == tOp && t.text == "%" {
			p.next()
			left = &node{kind: nPercent, args: []*node{left}}
			continue
		}
		p.next()
		op := t.text
		if t.kind == tSpace {
			op = " "
		}
		// All binary operators are left-associative in Excel (2^3^2 is
		// 64).
		right, err := p.expr(power)
		if err != nil {
			return nil, err
		}
		left = &node{kind: nBinary, str: op, args: []*node{left, right}}
	}
}

func (p *parser) prefix() (*node, error) {
	t := p.next()
	switch t.kind {
	case tNum:
		return &node{kind: nNum, num: t.num}, nil
	case tStr:
		return &node{kind: nStr, str: t.text}, nil
	case tBool:
		return &node{kind: nBool, b: t.b}, nil
	case tErr:
		return &node{kind: nErr, str: t.text}, nil
	case tRef:
		n := &node{kind: nRef, sheet: t.sheet, area: t.area}
		if t.spill {
			return &node{kind: nCall, str: "ANCHORARRAY", args: []*node{n}}, nil
		}
		return n, nil
	case tName:
		return &node{kind: nName, str: t.text, sheet: t.sheet}, nil
	case tOp:
		switch t.text {
		case "-", "+":
			// Negation binds tighter than ^ (-2^2 is 4), looser than the
			// reference operators.
			x, err := p.expr(70)
			if err != nil {
				return nil, err
			}
			if t.text == "+" {
				return x, nil
			}
			return &node{kind: nUnary, str: "-", args: []*node{x}}, nil
		}
		return nil, ErrSyntax
	case tFunc:
		return p.call(t.text)
	case tLParen:
		first, err := p.expr(0)
		if err != nil {
			return nil, err
		}
		items := []*node{first}
		for p.peek().kind == tComma {
			p.next()
			n, err := p.expr(0)
			if err != nil {
				return nil, err
			}
			items = append(items, n)
		}
		if p.next().kind != tRParen {
			return nil, ErrSyntax
		}
		if len(items) > 1 {
			return &node{kind: nUnion, args: items}, nil
		}
		return first, nil
	case tLBrace:
		return p.arrayConst()
	}
	return nil, ErrSyntax
}

func (p *parser) call(name string) (*node, error) {
	if p.next().kind != tLParen {
		return nil, ErrSyntax
	}
	n := &node{kind: nCall, str: strings.ToUpper(name)}
	if p.peek().kind == tRParen {
		p.next()
		return n, nil
	}
	for {
		switch p.peek().kind {
		case tComma:
			n.args = append(n.args, &node{kind: nMissing})
		case tRParen:
			n.args = append(n.args, &node{kind: nMissing})
		default:
			a, err := p.expr(0)
			if err != nil {
				return nil, err
			}
			n.args = append(n.args, a)
		}
		switch p.next().kind {
		case tComma:
			if p.peek().kind == tRParen {
				p.next()
				n.args = append(n.args, &node{kind: nMissing})
				return n, nil
			}
			continue
		case tRParen:
			return n, nil
		}
		return nil, ErrSyntax
	}
}

func (p *parser) arrayConst() (*node, error) {
	var rows [][]Value
	row := []Value{}
	for {
		t := p.next()
		var v Value
		neg := false
		if t.kind == tOp && (t.text == "-" || t.text == "+") {
			neg = t.text == "-"
			t = p.next()
		}
		switch t.kind {
		case tNum:
			v = num(t.num)
			if neg {
				v.Num = -v.Num
			}
		case tStr:
			v = str(t.text)
		case tBool:
			v = boolean(t.b)
		case tErr:
			v = errv(t.text)
		default:
			return nil, ErrSyntax
		}
		row = append(row, v)
		switch p.next().kind {
		case tComma:
		case tSemi:
			rows = append(rows, row)
			row = []Value{}
		case tRBrace:
			rows = append(rows, row)
			cols := len(rows[0])
			a := newArr(len(rows), cols)
			for i, r := range rows {
				if len(r) != cols {
					return nil, ErrSyntax
				}
				copy(a.V[i*cols:], r)
			}
			return &node{kind: nArray, arr: a}, nil
		default:
			return nil, ErrSyntax
		}
	}
}

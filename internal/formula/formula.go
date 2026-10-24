// Package formula finds cell references inside A1-style formulas so they can
// be shifted (to expand shared formulas) or normalised into patterns (to
// summarise a column of formulas as one line).
package formula

import (
	"errors"
	"fmt"
	"regexp"
	"strconv"
	"strings"

	"github.com/eemax/sheets/internal/grid"
)

type refKind uint8

const (
	cellRef refKind = iota
	colRef          // the A in A:C
	rowRef          // the 3 in 3:5
)

// part is either literal text or a reference.
type part struct {
	text   string
	isRef  bool
	kind   refKind
	row    int
	col    int
	rowAbs bool
	colAbs bool
	// qualified marks a reference into another sheet (Sheet!A1); sheet
	// names it.
	qualified bool
	sheet     string
}

// tokenize splits a formula into literal text and references. It skips
// string literals, quoted sheet names, structured references and function
// names; everything it does not recognise stays literal.
func tokenize(f string) []part {
	var parts []part
	lit := 0 // start of pending literal text
	flush := func(end int) {
		if end > lit {
			parts = append(parts, part{text: f[lit:end]})
		}
	}
	sheet := "" // the sheet named before the last '!'
	for i := 0; i < len(f); {
		c := f[i]
		switch {
		case c == '"':
			i = skipQuoted(f, i, '"')
			continue
		case c == '\'':
			j := skipQuoted(f, i, '\'')
			if j < len(f) && f[j] == '!' && j-i >= 2 {
				sheet = strings.ReplaceAll(f[i+1:j-1], "''", "'")
			}
			i = j
			continue
		case c == '[':
			depth := 0
			for i < len(f) {
				if f[i] == '[' {
					depth++
				} else if f[i] == ']' {
					depth--
					if depth == 0 {
						i++
						break
					}
				}
				i++
			}
			continue
		}
		if !(c == '$' || isAlpha(c) || isDigit(c)) || (i > 0 && isWordChar(f[i-1])) {
			i++
			continue
		}
		// Scan a whole word (letters, digits, $, _, .) to decide what it is.
		j := i
		for j < len(f) && (isWordChar(f[j]) || f[j] == '$') {
			j++
		}
		word := f[i:j]
		next := byte(0)
		if j < len(f) {
			next = f[j]
		}
		if next == '(' || next == '!' {
			if next == '!' {
				sheet = word
			}
			i = j // function name or sheet name
			continue
		}
		if p, ok := parseRef(word); ok {
			// A row or column reference only counts as part of a range.
			if p.kind != cellRef && !(next == ':' || i > 0 && f[i-1] == ':') {
				i = j
				continue
			}
			flush(i)
			p.qualified = i > 0 && f[i-1] == '!'
			if p.qualified {
				p.sheet = sheet
			}
			parts = append(parts, p)
			lit = j
		}
		i = j
	}
	flush(len(f))
	return parts
}

func skipQuoted(f string, i int, q byte) int {
	i++
	for i < len(f) {
		if f[i] == q {
			if i+1 < len(f) && f[i+1] == q {
				i += 2
				continue
			}
			return i + 1
		}
		i++
	}
	return i
}

func isAlpha(c byte) bool    { return c >= 'A' && c <= 'Z' || c >= 'a' && c <= 'z' }
func isDigit(c byte) bool    { return c >= '0' && c <= '9' }
func isWordChar(c byte) bool { return isAlpha(c) || isDigit(c) || c == '_' || c == '.' || c > 127 }

// parseRef recognises $A$1, A1, $A (column) and $3 (row).
func parseRef(w string) (part, bool) {
	p := part{isRef: true}
	i := 0
	if i < len(w) && w[i] == '$' {
		p.colAbs = true
		i++
	}
	ls := i
	for i < len(w) && isAlpha(w[i]) {
		i++
	}
	letters := w[ls:i]
	if letters == "" {
		// Row reference such as 3 or $3.
		n, err := strconv.Atoi(w[ls:])
		if err != nil || n < 1 || n > grid.MaxRows {
			return p, false
		}
		p.kind, p.row, p.rowAbs, p.colAbs = rowRef, n, p.colAbs, false
		return p, true
	}
	col, ok := grid.ColIndex(letters)
	if !ok {
		return p, false
	}
	p.col = col
	if i == len(w) {
		p.kind = colRef
		return p, true
	}
	if w[i] == '$' {
		p.rowAbs = true
		i++
	}
	n, err := strconv.Atoi(w[i:])
	if err != nil || n < 1 || n > grid.MaxRows || strings.ContainsAny(w[i:], "+-") {
		return p, false
	}
	p.row = n
	return p, true
}

func (p part) render() string {
	var b strings.Builder
	if p.kind != rowRef {
		if p.colAbs {
			b.WriteByte('$')
		}
		b.WriteString(grid.ColName(p.col))
	}
	if p.kind != colRef {
		if p.rowAbs {
			b.WriteByte('$')
		}
		b.WriteString(strconv.Itoa(p.row))
	}
	return b.String()
}

// Shift moves every relative reference by (dr, dc), as copying the formula
// that many rows and columns away would. References pushed off the sheet
// become #REF!.
func Shift(f string, dr, dc int) string {
	if dr == 0 && dc == 0 {
		return f
	}
	var b strings.Builder
	for _, p := range tokenize(f) {
		if !p.isRef {
			b.WriteString(p.text)
			continue
		}
		if p.kind != rowRef && !p.colAbs {
			p.col += dc
		}
		if p.kind != colRef && !p.rowAbs {
			p.row += dr
		}
		if p.kind != rowRef && (p.col < 1 || p.col > grid.MaxCols) || p.kind != colRef && (p.row < 1 || p.row > grid.MaxRows) {
			b.WriteString("#REF!")
			continue
		}
		b.WriteString(p.render())
	}
	return b.String()
}

// Pattern describes a formula relative to the cell holding it, so that
// formulas copied down a column share one Key. Display shows it in A1 style
// with {r} standing for the current row, e.g. "D{r}-E{r}".
func Pattern(f string, row, col int) (key, display string) {
	var k, d strings.Builder
	for _, p := range tokenize(f) {
		if !p.isRef {
			k.WriteString(p.text)
			d.WriteString(p.text)
			continue
		}
		if p.kind != rowRef {
			if p.colAbs {
				k.WriteString("C" + strconv.Itoa(p.col))
				d.WriteString("$" + grid.ColName(p.col))
			} else {
				k.WriteString("C[" + strconv.Itoa(p.col-col) + "]")
				d.WriteString(grid.ColName(p.col))
			}
		}
		if p.kind != colRef {
			if p.rowAbs {
				k.WriteString("R" + strconv.Itoa(p.row))
				d.WriteString("$" + strconv.Itoa(p.row))
			} else {
				off := p.row - row
				k.WriteString("R[" + strconv.Itoa(off) + "]")
				switch {
				case off == 0:
					d.WriteString("{r}")
				case off > 0:
					d.WriteString("{r+" + strconv.Itoa(off) + "}")
				default:
					d.WriteString("{r" + strconv.Itoa(off) + "}")
				}
			}
		}
	}
	return k.String(), d.String()
}

var rowPlaceholder = regexp.MustCompile(`\{r([+-]\d+)?\}`)

// HasRowPlaceholder reports whether f uses {r}, the notation Pattern
// produces.
func HasRowPlaceholder(f string) bool {
	return strings.Contains(f, "{r") && rowPlaceholder.MatchString(f)
}

// FillRow replaces {r}, {r+1}, {r-2}, ... with row numbers counted from
// row, so a pattern from peek can be written back to any row.
func FillRow(f string, row int) string {
	if !strings.Contains(f, "{r") {
		return f
	}
	return rowPlaceholder.ReplaceAllStringFunc(f, func(m string) string {
		off := 0
		if len(m) > 3 {
			off, _ = strconv.Atoi(m[2 : len(m)-1])
		}
		return strconv.Itoa(row + off)
	})
}

// CrossSheet reports whether the formula refers to another sheet or file.
func CrossSheet(f string) bool {
	for i := 0; i < len(f); i++ {
		switch f[i] {
		case '"':
			i = skipQuoted(f, i, '"') - 1
		case '!':
			return true
		}
	}
	return false
}

// SameRowColumns lists the columns a formula in the given row reads from its
// own row on the same sheet (D{r} and E{r} in =D{r}*E{r}).
func SameRowColumns(f string, row int) []int {
	var cols []int
	for _, p := range tokenize(f) {
		if p.isRef && p.kind == cellRef && !p.qualified && !p.rowAbs && p.row == row {
			cols = append(cols, p.col)
		}
	}
	return cols
}

// Divides reports whether a formula uses the division operator.
func Divides(f string) bool {
	for _, t := range lex(f) {
		if t.text == "/" {
			return true
		}
	}
	return false
}

// Multiplies reports whether a formula uses the multiplication operator.
func Multiplies(f string) bool {
	for _, t := range lex(f) {
		if t.text == "*" {
			return true
		}
	}
	return false
}

// Calls lists the functions a formula calls, upper case, without storage
// prefixes.
func Calls(f string) []string {
	var out []string
	for _, t := range lex(f) {
		if t.call {
			out = append(out, strings.ToUpper(bare(t.text)))
		}
	}
	return out
}

// DeleteLines rewrites a formula for deleting rows (or columns, with cols
// set) from..to of sheet target, as Excel does: references into the
// deleted lines become #REF!, ranges partly inside shrink, and references
// beyond move up (or left). home is the sheet holding the formula ("" for
// names, which are always qualified).
func DeleteLines(f, home, target string, cols bool, from, to int) string {
	ps := tokenize(f)
	changed := false
	sheetOf := func(p part) string {
		if p.qualified {
			return p.sheet
		}
		return home
	}
	line := func(p part) (int, bool) {
		switch {
		case cols && p.kind != rowRef:
			return p.col, true
		case !cols && p.kind != colRef:
			return p.row, true
		}
		return 0, false // a whole row is untouched by deleting columns, and vice versa
	}
	in := func(n int) bool { return n >= from && n <= to }
	n := to - from + 1
	set := func(p *part, v int) {
		if cols {
			p.col = v
		} else {
			p.row = v
		}
	}
	// shift moves a reference beyond the deleted block back by its size.
	shift := func(p *part) {
		if v, ok := line(*p); ok && v > to {
			set(p, v-n)
			changed = true
		}
	}
	for i := 0; i < len(ps); i++ {
		p := &ps[i]
		if !p.isRef || !strings.EqualFold(sheetOf(*p), target) {
			continue
		}
		// A range: ref ':' ref.
		if i+2 < len(ps) && !ps[i+1].isRef && ps[i+1].text == ":" && ps[i+2].isRef {
			q := &ps[i+2]
			a, aok := line(*p)
			b, bok := line(*q)
			switch {
			case !aok || !bok:
			case in(a) && in(b):
				*p = part{text: "#REF!"}
				ps[i+1].text, *q = "", part{}
				changed = true
			case in(a):
				set(p, to+1)
				changed = true
			case in(b):
				set(q, from-1)
				changed = true
			}
			if p.isRef {
				shift(p)
				shift(q)
			}
			i += 2
			continue
		}
		if v, ok := line(*p); ok && in(v) && p.kind == cellRef {
			*p = part{text: "#REF!"}
			changed = true
			continue
		}
		shift(p)
	}
	if !changed {
		return f
	}
	var b strings.Builder
	for _, p := range ps {
		if p.isRef {
			b.WriteString(p.render())
		} else {
			b.WriteString(p.text)
		}
	}
	return b.String()
}

// InsertLines rewrites a formula for inserting n rows (or columns, with
// cols set) before line at of sheet target: references at or beyond it move
// down (or right), so ranges spanning the insertion grow. home is the sheet
// holding the formula ("" for names).
func InsertLines(f, home, target string, cols bool, at, n int) string {
	ps := tokenize(f)
	changed := false
	for i := range ps {
		p := &ps[i]
		if !p.isRef {
			continue
		}
		sheet := home
		if p.qualified {
			sheet = p.sheet
		} else if i >= 2 && !ps[i-1].isRef && ps[i-1].text == ":" && ps[i-2].isRef && ps[i-2].qualified {
			sheet = ps[i-2].sheet // second half of Sheet!A1:B2
		}
		if !strings.EqualFold(sheet, target) {
			continue
		}
		switch {
		case cols && p.kind != rowRef && p.col >= at:
			p.col += n
			changed = true
		case !cols && p.kind != colRef && p.row >= at:
			p.row += n
			changed = true
		}
	}
	if !changed {
		return f
	}
	var b strings.Builder
	for _, p := range ps {
		if p.isRef {
			if p.col > grid.MaxCols || p.row > grid.MaxRows {
				b.WriteString("#REF!")
				continue
			}
			b.WriteString(p.render())
		} else {
			b.WriteString(p.text)
		}
	}
	return b.String()
}

// ExtendRanges moves the end of ranges on sheet target that end at row
// from down to row to: a total's =SUM(B4:B9) becomes =SUM(B4:B12) when rows
// are added right above it. home is the sheet holding the formula.
func ExtendRanges(f, home, target string, from, to int) string {
	return ExtendRangesIn(f, home, target, from, to, 1, grid.MaxCols)
}

// ExtendRangesIn is ExtendRanges for the ranges within columns c1..c2 only.
func ExtendRangesIn(f, home, target string, from, to, c1, c2 int) string {
	ps := tokenize(f)
	changed := false
	for i := 0; i+2 < len(ps); i++ {
		a, sep, b := &ps[i], ps[i+1], &ps[i+2]
		if !a.isRef || sep.isRef || sep.text != ":" || !b.isRef || a.kind != cellRef || b.kind != cellRef {
			continue
		}
		sheet := home
		if a.qualified {
			sheet = a.sheet
		}
		if strings.EqualFold(sheet, target) && b.row == from && a.row <= from && min(a.col, b.col) >= c1 && max(a.col, b.col) <= c2 {
			b.row = to
			changed = true
		}
		i += 2
	}
	if !changed {
		return f
	}
	var out strings.Builder
	for _, p := range ps {
		if p.isRef {
			out.WriteString(p.render())
		} else {
			out.WriteString(p.text)
		}
	}
	return out.String()
}

// Ref is a cell or range a formula refers to.
type Ref struct {
	Sheet string // "" for the formula's own sheet
	Range grid.Range
}

// Refs lists the cells and ranges a formula refers to. Whole rows and
// columns extend to the sheet limits.
func Refs(f string) []Ref {
	ps := tokenize(f)
	var out []Ref
	bounds := func(p part) grid.Range {
		switch p.kind {
		case colRef:
			return grid.Range{R1: 1, C1: p.col, R2: grid.MaxRows, C2: p.col}
		case rowRef:
			return grid.Range{R1: p.row, C1: 1, R2: p.row, C2: grid.MaxCols}
		}
		return grid.Range{R1: p.row, C1: p.col, R2: p.row, C2: p.col}
	}
	for i := 0; i < len(ps); i++ {
		p := ps[i]
		if !p.isRef {
			continue
		}
		r := Ref{Range: bounds(p)}
		if p.qualified {
			r.Sheet = p.sheet
		}
		if i+2 < len(ps) && !ps[i+1].isRef && ps[i+1].text == ":" && ps[i+2].isRef {
			r.Range = r.Range.Union(bounds(ps[i+2]))
			i += 2
		}
		out = append(out, r)
	}
	return out
}

// Numbers lists the numeric constants written in a formula (not row
// numbers of references or ranges like 3:5).
func Numbers(f string) []float64 {
	ts := lex(f)
	var out []float64
	for i, t := range ts {
		if !t.word || t.call {
			continue
		}
		if i > 0 && ts[i-1].text == ":" || i+1 < len(ts) && ts[i+1].text == ":" {
			continue
		}
		if n, err := strconv.ParseFloat(t.text, 64); err == nil {
			out = append(out, n)
		}
	}
	return out
}

// Check finds what makes a formula unreadable to Excel, which offers to
// repair a file holding one: unclosed quotes or brackets, or an operator
// with nothing after it. It does not know the functions, so =NOPE(1)
// passes (Excel answers #NAME?).
func Check(f string) error {
	f = strings.TrimPrefix(strings.TrimSpace(f), "=")
	if strings.TrimSpace(f) == "" {
		return errors.New("it is empty")
	}
	var stack []byte
	var at []int // where each open bracket is
	closer := map[byte]byte{'(': ')', '{': '}', '[': ']'}
	last := byte(0) // the last character outside quotes, not a blank
	for i := 0; i < len(f); i++ {
		c := f[i]
		if len(stack) > 0 && stack[len(stack)-1] == '[' {
			// Inside a structured reference: ' escapes the next character.
			switch c {
			case '\'':
				i++
			case '[':
				stack, at = append(stack, c), append(at, i)
			case ']':
				stack, at = stack[:len(stack)-1], at[:len(at)-1]
			}
			last = c
			continue
		}
		switch c {
		case '"', '\'':
			j := i + 1
			for ; j < len(f); j++ {
				if f[j] == c {
					if j+1 < len(f) && f[j+1] == c { // "" inside a string
						j++
						continue
					}
					break
				}
			}
			if j >= len(f) {
				return fmt.Errorf("the %c at position %d is not closed", c, i+2)
			}
			i, last = j, c
			continue
		case '(', '{', '[':
			stack, at = append(stack, c), append(at, i)
		case ')', '}', ']':
			if len(stack) == 0 || closer[stack[len(stack)-1]] != c {
				return fmt.Errorf("the %c at position %d closes nothing", c, i+2)
			}
			stack, at = stack[:len(stack)-1], at[:len(at)-1]
		}
		if c != ' ' && c != '\t' && c != '\n' && c != '\r' {
			last = c
		}
	}
	if n := len(stack); n > 0 {
		return fmt.Errorf("the %c at position %d is not closed", stack[n-1], at[n-1]+2)
	}
	if strings.IndexByte("+-*/^&=<>,(:", last) >= 0 {
		return fmt.Errorf("it ends with %q", string(last))
	}
	return nil
}

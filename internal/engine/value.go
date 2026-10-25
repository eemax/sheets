// Package engine is the sheets formula engine: it computes Excel formulas
// the way Excel does, with typed values (a number is never confused with
// text that looks like one), references that stay references until a
// function or operator needs their values, arrays, and Excel's rules for
// coercing, comparing and propagating errors. It reads a workbook through
// Book and keeps no copy of it.
//
// Its yardstick is Excel itself: the corpus test in internal/xlsx compares
// it with the results Excel stored in real workbooks.
package engine

import (
	"math"
	"strconv"
	"strings"
)

// Kind is the type of a value.
type Kind uint8

const (
	Empty  Kind = iota // a blank cell, or a missing argument
	Number             // numbers, dates and times
	Text
	Bool
	Error
	Array // a grid of values, from an array constant or an array result
	Ref   // a reference: one or more areas of a sheet
)

// Value is a formula value. Errors keep their code ("#N/A") in Str.
type Value struct {
	Kind Kind
	Num  float64
	Str  string
	Bool bool
	Arr  *Arr
	Ref  *RefV
}

// Arr is an array, row by row.
type Arr struct {
	Rows, Cols int
	V          []Value
}

func (a *Arr) At(r, c int) Value { return a.V[r*a.Cols+c] }

// RefV is a reference: one area, or several for a union (A1:B2,D4).
type RefV struct {
	Sheet string
	Areas []Area
}

// Area is a rectangle of cells, 1-based and inclusive.
type Area struct{ R1, C1, R2, C2 int }

func (a Area) Rows() int { return a.R2 - a.R1 + 1 }
func (a Area) Cols() int { return a.C2 - a.C1 + 1 }

// Error codes.
const (
	ErrNull  = "#NULL!"
	ErrDiv0  = "#DIV/0!"
	ErrValue = "#VALUE!"
	ErrRef   = "#REF!"
	ErrName  = "#NAME?"
	ErrNum   = "#NUM!"
	ErrNA    = "#N/A"
	ErrSpill = "#SPILL!"
	ErrCalc  = "#CALC!"
)

var errorCodes = []string{ErrNull, ErrDiv0, ErrValue, ErrRef, ErrName, ErrNum, ErrNA, "#GETTING_DATA", ErrSpill, ErrCalc}

func num(n float64) Value { return Value{Kind: Number, Num: n} }
func str(s string) Value  { return Value{Kind: Text, Str: s} }
func boolean(b bool) Value {
	return Value{Kind: Bool, Bool: b}
}
func errv(code string) Value { return Value{Kind: Error, Str: code} }

// numOrErr is a number, or #NUM! when it is not finite, as Excel gives.
func numOrErr(n float64) Value {
	if math.IsNaN(n) || math.IsInf(n, 0) {
		return errv(ErrNum)
	}
	return num(n)
}

func newArr(rows, cols int) *Arr { return &Arr{Rows: rows, Cols: cols, V: make([]Value, rows*cols)} }

func arrOf(a *Arr) Value { return Value{Kind: Array, Arr: a} }

// IsError reports whether v is an error value.
func (v Value) IsError() bool { return v.Kind == Error }

// toNumber coerces a scalar to a number as arithmetic does: text that reads
// as a number (" 1,200 ", "$5", "10%", "1/2/2024", "12:30") converts; other
// text is #VALUE!.
func toNumber(v Value) (float64, *Value) {
	switch v.Kind {
	case Number:
		return v.Num, nil
	case Empty:
		return 0, nil
	case Bool:
		if v.Bool {
			return 1, nil
		}
		return 0, nil
	case Text:
		n, ok, dependent := readNumber(v.Str)
		switch {
		case dependent:
			e := languageDependent(v.Str)
			return 0, &e
		case ok:
			return n, nil
		}
		e := errv(ErrValue)
		return 0, &e
	case Error:
		return 0, &v
	}
	e := errv(ErrValue)
	return 0, &e
}

// textNumber reads text as Excel does in arithmetic: numbers with thousands
// separators, a currency sign, a percent sign, scientific notation, dates and
// times in the forms Excel understands in an English locale.
func textNumber(s string) (float64, bool) {
	t := strings.TrimSpace(s)
	if t == "" {
		return 0, false
	}
	neg := false
	if strings.HasPrefix(t, "(") && strings.HasSuffix(t, ")") {
		neg, t = true, strings.TrimSpace(t[1:len(t)-1])
	}
	pct := false
	if strings.HasSuffix(t, "%") {
		pct, t = true, strings.TrimSpace(t[:len(t)-1])
	}
	sign := ""
	if strings.HasPrefix(t, "-") || strings.HasPrefix(t, "+") {
		sign, t = t[:1], strings.TrimSpace(t[1:])
	}
	if strings.HasPrefix(t, "$") {
		t = strings.TrimSpace(t[1:])
		if sign == "" && (strings.HasPrefix(t, "-") || strings.HasPrefix(t, "+")) {
			sign, t = t[:1], strings.TrimSpace(t[1:])
		}
	}
	if n, ok := plainNumber(t); ok {
		if sign == "-" {
			n = -n
		}
		if neg {
			n = -n
		}
		if pct {
			n /= 100
		}
		return n, true
	}
	if sign != "" || neg || pct {
		return 0, false
	}
	if n, ok := parseDateTime(s); ok {
		return n, true
	}
	return 0, false
}

// plainNumber reads digits with optional thousands separators (groups of
// three), a decimal point and an exponent.
func plainNumber(t string) (float64, bool) {
	if t == "" {
		return 0, false
	}
	mant, exp := t, ""
	if i := strings.IndexAny(t, "eE"); i > 0 {
		mant, exp = t[:i], t[i:]
	}
	if strings.Contains(mant, ",") {
		intPart, frac := mant, ""
		if i := strings.IndexByte(mant, '.'); i >= 0 {
			intPart, frac = mant[:i], mant[i:]
		}
		groups := strings.Split(intPart, ",")
		if len(groups[0]) == 0 || len(groups[0]) > 3 {
			return 0, false
		}
		for _, g := range groups[1:] {
			if len(g) != 3 {
				return 0, false
			}
		}
		mant = strings.Join(groups, "") + frac
	}
	for i := 0; i < len(mant); i++ {
		c := mant[i]
		if (c < '0' || c > '9') && c != '.' {
			return 0, false
		}
	}
	if strings.Count(mant, ".") > 1 || mant == "." {
		return 0, false
	}
	n, err := strconv.ParseFloat(mant+exp, 64)
	if err != nil {
		return 0, false
	}
	return n, true
}

// toText coerces a scalar to text as concatenation does.
func toText(v Value) (string, *Value) {
	switch v.Kind {
	case Text:
		return v.Str, nil
	case Number:
		return NumberText(v.Num), nil
	case Bool:
		if v.Bool {
			return "TRUE", nil
		}
		return "FALSE", nil
	case Empty:
		return "", nil
	case Error:
		return "", &v
	}
	e := errv(ErrValue)
	return "", &e
}

// NumberText writes a number as Excel's General format does when a
// formula turns it into text: up to 15 significant digits, no trailing
// zeros, scientific notation for very large or small magnitudes.
func NumberText(n float64) string {
	if n == 0 {
		return "0"
	}
	if math.IsInf(n, 0) || math.IsNaN(n) {
		return ErrNum
	}
	// Round to 15 significant digits first, as Excel stores no more.
	s := strconv.FormatFloat(n, 'e', 14, 64)
	mant, expPart, _ := strings.Cut(s, "e")
	exp, _ := strconv.Atoi(expPart)
	neg := strings.HasPrefix(mant, "-")
	mant = strings.TrimPrefix(mant, "-")
	digits := strings.TrimRight(strings.Replace(mant, ".", "", 1), "0")
	if digits == "" {
		return "0"
	}
	sign := ""
	if neg {
		sign = "-"
	}
	if exp >= 15 || exp < -9 {
		m := digits[:1]
		if len(digits) > 1 {
			m += "." + digits[1:]
		}
		es := "+"
		if exp < 0 {
			es, exp = "-", -exp
		}
		e := strconv.Itoa(exp)
		if len(e) < 2 {
			e = "0" + e
		}
		return sign + m + "E" + es + e
	}
	var b strings.Builder
	b.WriteString(sign)
	if exp < 0 {
		b.WriteString("0.")
		b.WriteString(strings.Repeat("0", -exp-1))
		b.WriteString(digits)
		return b.String()
	}
	if len(digits) <= exp+1 {
		b.WriteString(digits)
		b.WriteString(strings.Repeat("0", exp+1-len(digits)))
		return b.String()
	}
	b.WriteString(digits[:exp+1])
	b.WriteByte('.')
	b.WriteString(digits[exp+1:])
	return b.String()
}

// toBool coerces a scalar to a logical value.
func toBool(v Value) (bool, *Value) {
	switch v.Kind {
	case Bool:
		return v.Bool, nil
	case Number:
		return v.Num != 0, nil
	case Empty:
		return false, nil
	case Text:
		switch strings.ToUpper(strings.TrimSpace(v.Str)) {
		case "TRUE":
			return true, nil
		case "FALSE":
			return false, nil
		}
		e := errv(ErrValue)
		return false, &e
	case Error:
		return false, &v
	}
	e := errv(ErrValue)
	return false, &e
}

// compare orders two scalars as Excel's comparison operators do: numbers
// before text before logical values; text without regard to case; a blank
// equal to 0, "" or FALSE. It returns -1, 0 or 1.
func compare(a, b Value) int {
	if a.Kind == Empty {
		a = blankLike(b)
	}
	if b.Kind == Empty {
		b = blankLike(a)
	}
	ra, rb := typeRank(a.Kind), typeRank(b.Kind)
	if ra != rb {
		if ra < rb {
			return -1
		}
		return 1
	}
	switch a.Kind {
	case Number:
		return cmpFloat(a.Num, b.Num)
	case Text:
		return cmpText(a.Str, b.Str)
	case Bool:
		switch {
		case a.Bool == b.Bool:
			return 0
		case !a.Bool:
			return -1
		}
		return 1
	}
	return 0
}

// cmpFloat compares numbers as Excel does, to 15 significant digits.
func cmpFloat(a, b float64) int {
	if a == b {
		return 0
	}
	if roundSig(a, 15) == roundSig(b, 15) {
		return 0
	}
	if a < b {
		return -1
	}
	return 1
}

func cmpText(a, b string) int {
	return strings.Compare(foldCase(a), foldCase(b))
}

// foldCase folds text for comparison without regard to case.
func foldCase(s string) string { return strings.ToLower(s) }

func blankLike(o Value) Value {
	switch o.Kind {
	case Text:
		return str("")
	case Bool:
		return boolean(false)
	}
	return num(0)
}

func typeRank(k Kind) int {
	switch k {
	case Number, Empty:
		return 0
	case Text:
		return 1
	case Bool:
		return 2
	}
	return 3
}

// roundSig rounds to n significant digits.
func roundSig(x float64, n int) float64 {
	if x == 0 || math.IsInf(x, 0) || math.IsNaN(x) {
		return x
	}
	s := strconv.FormatFloat(x, 'e', n-1, 64)
	r, _ := strconv.ParseFloat(s, 64)
	return r
}

// readNumber reads text as a number the way arithmetic does. Excel reads
// such text by its language setting, so the engine reads only what every
// language reads alike: digits with a sign, spaces around, a percent sign
// or an exponent; ISO dates (2024-07-01); times (13:30). dependent is set
// for text whose reading depends on the language: decimal and thousands
// marks (1,200 is 1200 in English, 1.2 in German; 1.5 is no number in
// German), currency signs, and dates written with slashes, dots or month
// names. The engine does not guess at those.
func readNumber(s string) (n float64, ok, dependent bool) {
	t := strings.TrimSpace(s)
	if t == "" {
		return 0, false, false
	}
	if n, ok := universalNumber(t); ok {
		return n, true, false
	}
	if _, ok := textNumber(t); ok {
		return 0, false, true
	}
	if otherLanguageNumber(t) || dayFirstDate(t) {
		return 0, false, true
	}
	return 0, false, false
}

// universalNumber reads text every Excel language reads alike.
func universalNumber(t string) (float64, bool) {
	body := strings.TrimSpace(strings.TrimSuffix(t, "%"))
	pct := body != t
	neg := false
	if strings.HasPrefix(body, "-") || strings.HasPrefix(body, "+") {
		neg, body = body[0] == '-', strings.TrimSpace(body[1:])
	}
	digits := body
	if i := strings.IndexAny(body, "eE"); i > 0 {
		digits = body[:i]
		exp := strings.TrimPrefix(strings.TrimPrefix(body[i+1:], "+"), "-")
		if exp == "" || strings.Trim(exp, "0123456789") != "" {
			digits = ""
		}
	}
	if digits != "" && strings.Trim(digits, "0123456789") == "" {
		n, err := strconv.ParseFloat(body, 64)
		if err == nil {
			if neg {
				n = -n
			}
			if pct {
				n /= 100
			}
			return n, true
		}
	}
	// ISO dates, times, and the two together.
	if isISODate(t) || isTime(t) || isISODateTime(t) {
		if n, ok := parseDateTime(t); ok {
			return n, true
		}
	}
	return 0, false
}

func isISODate(t string) bool {
	return len(t) == 10 && t[4] == '-' && t[7] == '-' && strings.Trim(t[:4]+t[5:7]+t[8:], "0123456789") == ""
}

func isTime(t string) bool {
	fields := strings.Fields(t)
	if len(fields) == 0 || len(fields) > 2 || !strings.Contains(fields[0], ":") {
		return false
	}
	if strings.Trim(fields[0], "0123456789:.") != "" {
		return false
	}
	if len(fields) == 2 {
		up := strings.ToUpper(fields[1])
		return up == "AM" || up == "PM"
	}
	return true
}

func isISODateTime(t string) bool {
	d, rest, ok := strings.Cut(t, " ")
	if !ok {
		d, rest, ok = strings.Cut(t, "T")
	}
	return ok && isISODate(d) && isTime(rest)
}

// otherLanguageNumber tells whether text reads as a number with the
// marks swapped (1.200,50 or 1,5): a number to Excel in some languages.
func otherLanguageNumber(t string) bool {
	swapped := strings.Map(func(r rune) rune {
		switch r {
		case ',':
			return '.'
		case '.':
			return ','
		}
		return r
	}, t)
	_, ok := plainNumber(strings.TrimPrefix(strings.TrimPrefix(swapped, "-"), "+"))
	return ok
}

// languageDependent is the unknown value of text whose reading as a
// number or date depends on Excel's language setting.
func languageDependent(s string) Value {
	return unsupported("it reads " + strconv.Quote(strings.TrimSpace(s)) + " as a number or date, which depends on Excel's language setting")
}

// notNumber is the error for text a function cannot read as a number or
// date: unknown when another language might read it, else #VALUE!.
func notNumber(s string) Value {
	if _, _, dependent := readNumber(s); dependent {
		return languageDependent(s)
	}
	return errv(ErrValue)
}

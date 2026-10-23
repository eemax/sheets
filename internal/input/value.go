// Package input turns text typed by an agent (a cell value on the command
// line, or TSV/CSV/JSON on stdin) into typed cell values.
package input

import (
	"encoding/json"
	"math"
	"regexp"
	"strconv"
	"strings"
	"time"

	"github.com/eemax/sheets/internal/grid"
)

// Value is a typed value ready to be written to a cell.
type Value struct {
	Kind    grid.Kind
	Text    string
	Num     float64
	Bool    bool
	Formula string      // without '='
	Fmt     grid.Format // number format to apply; Kind General means leave as is
}

// IsEmpty reports whether writing v clears the cell.
func (v Value) IsEmpty() bool { return v.Kind == grid.Empty && v.Formula == "" }

var (
	reNumber   = regexp.MustCompile(`^[+-]?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?$|^[+-]?\.\d+$`)
	reParenNeg = regexp.MustCompile(`^\((\d{1,3}(,\d{3})+|\d+)(\.\d+)?\)$`)
	rePercent  = regexp.MustCompile(`^[+-]?(\d+(\.\d+)?|\.\d+)\s?%$`)
	reDate     = regexp.MustCompile(`^\d{4}-\d{2}-\d{2}$`)
	reDateTime = regexp.MustCompile(`^\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}(:\d{2}(\.\d+)?)?(Z|[+-]\d{2}:?\d{2})?$`)
	reTime     = regexp.MustCompile(`^\d{1,2}:\d{2}(:\d{2})?$`)
)

// prefix and suffix currency symbols recognised in input.
var (
	prefixSymbols = []string{"R$", "$", "€", "£", "¥", "₹", "₽", "₩", "₺", "₪", "₫", "₱", "฿", "₴", "₦"}
	suffixSymbols = []string{"€", "kr", "zł", "Kč", "CHF", "Ft", "lei", "$", "£"}
)

// Parse infers the type of a text value. With raw set, everything except
// the empty string is text and nothing is a formula.
func Parse(s string, raw bool) Value {
	if s == "" {
		return Value{}
	}
	if raw {
		return Value{Kind: grid.Text, Text: s}
	}
	if strings.HasPrefix(s, "'") {
		return Value{Kind: grid.Text, Text: s[1:]}
	}
	if len(s) > 1 && s[0] == '=' {
		return Value{Formula: s[1:]}
	}
	t := strings.TrimSpace(s)
	switch strings.ToUpper(t) {
	case "TRUE":
		return Value{Kind: grid.Bool, Bool: true}
	case "FALSE":
		return Value{Kind: grid.Bool}
	}
	if n, dec, ok := number(t); ok {
		k := grid.General
		if strings.Contains(t, ",") {
			k = grid.Integer
			if dec > 0 {
				k = grid.Decimal
			}
		}
		return Value{Kind: grid.Number, Num: n, Fmt: grid.Format{Kind: k, Decimals: dec}}
	}
	if m := reParenNeg.FindString(t); m != "" {
		if n, dec, ok := number(t[1 : len(t)-1]); ok {
			k := grid.Integer
			if dec > 0 {
				k = grid.Decimal
			}
			return Value{Kind: grid.Number, Num: -n, Fmt: grid.Format{Kind: k, Decimals: dec}}
		}
	}
	if rePercent.MatchString(t) {
		body := strings.TrimSpace(strings.TrimSuffix(t, "%"))
		if n, err := strconv.ParseFloat(body, 64); err == nil {
			return Value{Kind: grid.Number, Num: round15(n / 100), Fmt: grid.Format{Kind: grid.Percent, Decimals: decimals(body)}}
		}
	}
	if v, ok := currency(t); ok {
		return v
	}
	if reDate.MatchString(t) {
		if d, err := time.Parse("2006-01-02", t); err == nil {
			return Value{Kind: grid.Number, Num: grid.TimeToSerial(d), Fmt: grid.Format{Kind: grid.Date}}
		}
	}
	if reDateTime.MatchString(t) {
		if d, ok := parseDateTime(t); ok {
			return Value{Kind: grid.Number, Num: grid.TimeToSerial(d), Fmt: grid.Format{Kind: grid.DateTime}}
		}
	}
	if reTime.MatchString(t) {
		parts := strings.Split(t, ":")
		h, _ := strconv.Atoi(parts[0])
		m, _ := strconv.Atoi(parts[1])
		sec := 0
		if len(parts) == 3 {
			sec, _ = strconv.Atoi(parts[2])
		}
		if h < 24 && m < 60 && sec < 60 {
			f := grid.Format{Kind: grid.Time}
			if len(parts) == 3 {
				f.Decimals = 1
			}
			return Value{Kind: grid.Number, Num: float64(h*3600+m*60+sec) / 86400, Fmt: f}
		}
	}
	return Value{Kind: grid.Text, Text: s}
}

// number parses plain numbers with optional thousands separators. Values
// that would lose information as a number stay text: leading zeros (codes,
// zip codes) and more than 15 significant digits (ids, card numbers).
func number(t string) (float64, int, bool) {
	if !reNumber.MatchString(t) {
		return 0, 0, false
	}
	body := strings.TrimLeft(t, "+-")
	if len(body) > 1 && body[0] == '0' && body[1] != '.' {
		return 0, 0, false
	}
	digits := strings.Trim(strings.NewReplacer(",", "", ".", "").Replace(body), "0")
	if len(digits) > 15 {
		return 0, 0, false
	}
	n, err := strconv.ParseFloat(strings.ReplaceAll(t, ",", ""), 64)
	if err != nil {
		return 0, 0, false
	}
	return n, decimals(body), true
}

func decimals(s string) int {
	if i := strings.IndexByte(s, '.'); i >= 0 {
		return len(s) - i - 1
	}
	return 0
}

func currency(t string) (Value, bool) {
	neg := false
	s := t
	if strings.HasPrefix(s, "-") {
		neg, s = true, s[1:]
	}
	for _, sym := range prefixSymbols {
		if rest, ok := strings.CutPrefix(s, sym); ok {
			rest = strings.TrimSpace(rest)
			if strings.HasPrefix(rest, "-") && !neg {
				neg, rest = true, rest[1:]
			}
			if n, dec, ok := number(rest); ok {
				return currencyValue(n, dec, sym, neg), true
			}
		}
	}
	for _, sym := range suffixSymbols {
		if rest, ok := strings.CutSuffix(s, sym); ok {
			if n, dec, ok := number(strings.TrimSpace(rest)); ok {
				return currencyValue(n, dec, sym, neg), true
			}
		}
	}
	return Value{}, false
}

func currencyValue(n float64, dec int, sym string, neg bool) Value {
	if neg {
		n = -n
	}
	return Value{Kind: grid.Number, Num: n, Fmt: grid.Format{Kind: grid.Currency, Decimals: dec, Symbol: sym}}
}

func parseDateTime(t string) (time.Time, bool) {
	t = strings.Replace(t, " ", "T", 1)
	for _, layout := range []string{time.RFC3339Nano, "2006-01-02T15:04:05.999999999", "2006-01-02T15:04", "2006-01-02T15:04Z07:00", "2006-01-02T15:04:05Z0700"} {
		if d, err := time.Parse(layout, t); err == nil {
			return d, true
		}
	}
	return time.Time{}, false
}

// round15 trims floating point noise beyond 15 significant digits.
func round15(v float64) float64 {
	if v == 0 || math.IsInf(v, 0) || math.IsNaN(v) {
		return v
	}
	r, _ := strconv.ParseFloat(strconv.FormatFloat(v, 'g', 15, 64), 64)
	return r
}

// FromJSON converts a decoded JSON scalar. Strings go through Parse so that
// formulas and dates work the same way as in TSV input.
func FromJSON(v any, raw bool) Value {
	switch x := v.(type) {
	case nil:
		return Value{}
	case bool:
		return Value{Kind: grid.Bool, Bool: x}
	case float64:
		return Value{Kind: grid.Number, Num: x}
	case json.Number:
		// Integers with more than 15 significant digits are ids, not
		// quantities: keep them exact as text.
		n, err := x.Float64()
		lit := x.String()
		if err != nil || !strings.ContainsAny(lit, "eE") && len(strings.Trim(strings.NewReplacer("-", "", ".", "").Replace(lit), "0")) > 15 {
			return Value{Kind: grid.Text, Text: lit}
		}
		return Value{Kind: grid.Number, Num: n}
	case string:
		return Parse(x, raw)
	}
	b, _ := json.Marshal(v)
	return Value{Kind: grid.Text, Text: string(b)}
}

package grid

import (
	"strings"
	"unicode"
)

// FormatKind is the meaning of a number format, independent of its exact
// code. It is what survives when sheets strips and reapplies formatting.
type FormatKind uint8

const (
	General FormatKind = iota
	Integer
	Decimal
	Currency
	Percent
	Date
	DateTime
	Time
	TextFmt // "@": force text
)

var kindNames = [...]string{"general", "integer", "decimal", "currency", "percent", "date", "datetime", "time", "text"}

func (k FormatKind) String() string { return kindNames[k] }

// ParseFormatKind parses a kind name as accepted on the command line.
func ParseFormatKind(s string) (FormatKind, bool) {
	s = strings.ToLower(strings.TrimSpace(s))
	switch s {
	case "number":
		return Decimal, true
	case "int":
		return Integer, true
	case "money":
		return Currency, true
	case "pct":
		return Percent, true
	}
	for i, n := range kindNames {
		if n == s {
			return FormatKind(i), true
		}
	}
	return General, false
}

// IsNumeric reports whether the kind formats plain quantities.
func (k FormatKind) IsNumeric() bool {
	return k == Integer || k == Decimal || k == Currency || k == Percent
}

// IsTemporal reports whether the kind formats dates or times.
func (k FormatKind) IsTemporal() bool { return k == Date || k == DateTime || k == Time }

// Format is a classified number format.
type Format struct {
	Kind     FormatKind
	Decimals int    // digits after the decimal point; for Time, 1 means seconds shown
	Symbol   string // currency symbol for Currency
	Code     string // original format code
}

// builtinFormats maps the implicit xlsx number format ids to their codes.
var builtinFormats = map[int]string{
	0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00",
	5: `"$"#,##0_);("$"#,##0)`, 6: `"$"#,##0_);[Red]("$"#,##0)`,
	7: `"$"#,##0.00_);("$"#,##0.00)`, 8: `"$"#,##0.00_);[Red]("$"#,##0.00)`,
	9: "0%", 10: "0.00%", 11: "0.00E+00", 12: "# ?/?", 13: "# ??/??",
	14: "yyyy-mm-dd", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy",
	18: "h:mm AM/PM", 19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss", 22: "yyyy-mm-dd h:mm",
	37: "#,##0_);(#,##0)", 38: "#,##0_);[Red](#,##0)",
	39: "#,##0.00_);(#,##0.00)", 40: "#,##0.00_);[Red](#,##0.00)",
	41: `_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)`,
	42: `_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)`,
	43: `_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)`,
	44: `_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)`,
	45: "mm:ss", 46: "[h]:mm:ss", 47: "mm:ss.0", 48: "##0.0E+0", 49: "@",
}

// BuiltinFormatCode returns the code of a built-in format id. Ids 27-36 and
// 50-81 are locale-specific (mostly East Asian dates) and map to a date.
func BuiltinFormatCode(id int) (string, bool) {
	if c, ok := builtinFormats[id]; ok {
		return c, true
	}
	if id >= 27 && id <= 36 || id >= 50 && id <= 81 {
		return "yyyy-mm-dd", true
	}
	return "", false
}

var currencySymbols = []string{"$", "€", "£", "¥", "₹", "₽", "₩", "₺", "₪", "₫", "₱", "฿", "₴", "₦", "R$", "kr", "zł", "Kč", "CHF", "Fr.", "Ft", "lei", "R", "USD", "EUR", "GBP", "SEK", "NOK", "DKK", "JPY"}

// ClassifyFormat works out what a number format code means.
func ClassifyFormat(code string) Format {
	f := Format{Code: code}
	if code == "" || strings.EqualFold(code, "General") {
		return f
	}
	section := firstSection(code)
	var plain strings.Builder // code with literals removed
	for i := 0; i < len(section); i++ {
		c := section[i]
		switch c {
		case '"':
			j := strings.IndexByte(section[i+1:], '"')
			lit := section[i+1:]
			if j >= 0 {
				lit = section[i+1 : i+1+j]
				i += j + 1
			} else {
				i = len(section)
			}
			if f.Symbol == "" {
				f.Symbol = currencyIn(lit)
			}
		case '\\', '_', '*':
			if c == '\\' && i+1 < len(section) && f.Symbol == "" {
				f.Symbol = currencyIn(section[i+1 : i+2])
			}
			i++ // skip the escaped, padding or fill character
		case '[':
			j := strings.IndexByte(section[i:], ']')
			if j < 0 {
				i = len(section)
				break
			}
			inner := section[i+1 : i+j]
			i += j
			if strings.HasPrefix(inner, "$") {
				sym, _, _ := strings.Cut(inner[1:], "-")
				if sym != "" && f.Symbol == "" {
					f.Symbol = sym
				}
			} else if l := strings.ToLower(inner); l == "h" || l == "hh" || l == "m" || l == "mm" || l == "s" || l == "ss" {
				plain.WriteString(l) // elapsed time
			}
		default:
			plain.WriteByte(c)
		}
	}
	p := strings.ToLower(plain.String())
	if f.Symbol == "" {
		f.Symbol = currencyIn(p)
	}

	hasDigits := strings.ContainsAny(p, "0#?")
	isSci := strings.Contains(p, "e+") || strings.Contains(p, "e-")
	if !isSci {
		date := strings.ContainsAny(p, "yd")
		clock := strings.ContainsAny(p, "hs")
		if !date && !clock && strings.Contains(p, "m") && !hasDigits {
			date = true
		}
		switch {
		case date && clock:
			f.Kind = DateTime
			return f
		case date:
			f.Kind = Date
			return f
		case clock:
			f.Kind = Time
			if strings.Contains(p, "s") {
				f.Decimals = 1
			}
			return f
		}
	}
	if strings.Contains(p, "@") && !hasDigits {
		f.Kind = TextFmt
		return f
	}
	if !hasDigits {
		return f
	}
	if dot := strings.IndexByte(p, '.'); dot >= 0 {
		for _, c := range p[dot+1:] {
			if c != '0' && c != '#' && c != '?' {
				break
			}
			f.Decimals++
		}
	}
	switch {
	case strings.Contains(p, "%"):
		f.Kind = Percent
	case f.Symbol != "":
		f.Kind = Currency
	case f.Decimals == 0:
		f.Kind = Integer
	default:
		f.Kind = Decimal
	}
	return f
}

// firstSection returns the positive-number section of a format code.
func firstSection(code string) string {
	inQuote := false
	for i := 0; i < len(code); i++ {
		switch code[i] {
		case '"':
			inQuote = !inQuote
		case '\\':
			i++
		case ';':
			if !inQuote {
				return code[:i]
			}
		}
	}
	return code
}

func currencyIn(s string) string {
	t := strings.TrimSpace(s)
	if t == "" {
		return ""
	}
	for _, sym := range currencySymbols {
		if t == sym {
			return sym
		}
	}
	// Alphabetic symbols must match the whole literal so that, e.g., "R" in
	// a unit label is not taken for rand.
	for _, sym := range currencySymbols {
		if !unicode.IsLetter([]rune(sym)[0]) && strings.Contains(t, sym) {
			return sym
		}
	}
	return ""
}

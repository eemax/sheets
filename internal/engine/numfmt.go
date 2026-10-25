package engine

import (
	"math"
	"strconv"
	"strings"
)

// FormatNumber renders a number with an Excel format code, as TEXT does:
// sections for positive, negative, zero and text; 0 # ? placeholders,
// thousands separators, scaling commas, percent, exponents, literals, and
// date and time codes.
func FormatNumber(v float64, code string, date1904 bool) string {
	secs := splitSections(code)
	sec := secs[0]
	neg := v < 0
	if op, lim, ok := sectionCondition(secs[0]); ok {
		// [<3]...;...: the first section when its test holds; else the
		// second (with its own test, if any), else the third.
		switch {
		case cmpOp(op, cmpFloat(v, lim)):
			sec = secs[0]
		case len(secs) >= 2:
			sec = secs[1]
			if op2, lim2, ok := sectionCondition(secs[1]); ok && !cmpOp(op2, cmpFloat(v, lim2)) && len(secs) >= 3 {
				sec = secs[2]
			}
		default:
			return strings.Repeat("#", 8)
		}
		out := formatSection(math.Abs(v), stripBrackets(sec), date1904)
		if neg && strings.ContainsAny(out, "123456789") {
			out = "-" + out
		}
		return out
	}
	switch {
	case len(secs) >= 2 && neg:
		sec, v = secs[1], -v
	case len(secs) >= 3 && v == 0:
		sec = secs[2]
	}
	sec = stripBrackets(sec)
	if strings.TrimSpace(sec) == "" && len(secs) == 1 {
		return NumberText(v)
	}
	out := formatSection(v, sec, date1904)
	if neg && len(secs) < 2 && v != 0 {
		// One section: the sign goes in front, unless rounding left zero.
		if strings.ContainsAny(out, "123456789") {
			out = "-" + out
		}
	}
	return out
}

// formatSection renders a number with one section of a format code.
func formatSection(v float64, sec string, date1904 bool) string {
	if strings.EqualFold(strings.TrimSpace(sec), "General") {
		return NumberText(v)
	}
	if strings.TrimSpace(sec) == "" {
		return ""
	}
	if isDateCode(sec) {
		return formatDate(v, sec, date1904)
	}
	return formatDigits(v, sec)
}

// sectionCondition reads a section's leading [<3] test.
func sectionCondition(sec string) (string, float64, bool) {
	if !strings.HasPrefix(sec, "[") {
		return "", 0, false
	}
	i := strings.IndexByte(sec, ']')
	if i < 0 {
		return "", 0, false
	}
	inner := sec[1:i]
	for _, op := range []string{"<=", ">=", "<>", "<", ">", "="} {
		if strings.HasPrefix(inner, op) {
			n, err := strconv.ParseFloat(strings.TrimSpace(inner[len(op):]), 64)
			if err != nil {
				return "", 0, false
			}
			return op, n, true
		}
	}
	return "", 0, false
}

// FormatText renders text with a format code's text section (@ is the
// text).
func FormatText(s, code string) string {
	secs := splitSections(code)
	if len(secs) < 4 {
		if len(secs) == 1 && strings.Contains(secs[0], "@") {
			return renderLiteral(stripBrackets(secs[0]), s)
		}
		return s
	}
	return renderLiteral(stripBrackets(secs[3]), s)
}

func splitSections(code string) []string {
	var out []string
	var b strings.Builder
	inQ := false
	for i := 0; i < len(code); i++ {
		ch := code[i]
		switch {
		case ch == '"':
			inQ = !inQ
		case ch == '\\' && !inQ && i+1 < len(code):
			b.WriteByte(ch)
			i++
			ch = code[i]
		case ch == ';' && !inQ:
			out = append(out, b.String())
			b.Reset()
			continue
		}
		b.WriteByte(ch)
	}
	return append(out, b.String())
}

// stripBrackets drops colours, conditions and locale tags ([Red],
// [>100], [$-409]), keeping elapsed-time codes ([h], [mm]) and currency
// symbols ([$€-407] gives €).
func stripBrackets(sec string) string {
	var b strings.Builder
	inQ := false
	for i := 0; i < len(sec); i++ {
		ch := sec[i]
		if ch == '"' {
			inQ = !inQ
		}
		if ch == '[' && !inQ {
			j := strings.IndexByte(sec[i:], ']')
			if j < 0 {
				break
			}
			inner := sec[i+1 : i+j]
			lo := strings.ToLower(inner)
			switch {
			case lo == "h" || lo == "hh" || lo == "m" || lo == "mm" || lo == "s" || lo == "ss":
				b.WriteString("[" + lo + "]")
			case strings.HasPrefix(inner, "$"):
				sym := strings.TrimPrefix(inner, "$")
				if k := strings.IndexByte(sym, '-'); k >= 0 {
					sym = sym[:k]
				}
				b.WriteString(`"` + sym + `"`)
			}
			i += j
			continue
		}
		b.WriteByte(ch)
	}
	return b.String()
}

func isDateCode(sec string) bool {
	inQ := false
	for i := 0; i < len(sec); i++ {
		ch := sec[i]
		switch {
		case ch == '"':
			inQ = !inQ
		case inQ:
		case ch == '\\':
			i++
		case strings.IndexByte("yYdDhHsS", ch) >= 0:
			return true
		case ch == 'm' || ch == 'M':
			// m alone could be a literal in a number format only when
			// quoted; unquoted it is a month or minute.
			return true
		case ch == '[' && i+1 < len(sec) && strings.IndexByte("hHmMsS", sec[i+1]) >= 0:
			return true
		}
	}
	return false
}

// renderLiteral writes a section's literals, with s in place of @.
func renderLiteral(sec, s string) string {
	var b strings.Builder
	for i := 0; i < len(sec); i++ {
		ch := sec[i]
		switch ch {
		case '"':
			j := strings.IndexByte(sec[i+1:], '"')
			if j < 0 {
				b.WriteString(sec[i+1:])
				return b.String()
			}
			b.WriteString(sec[i+1 : i+1+j])
			i += j + 1
		case '\\':
			if i+1 < len(sec) {
				b.WriteByte(sec[i+1])
				i++
			}
		case '@':
			b.WriteString(s)
		case '_':
			b.WriteByte(' ')
			i++
		case '*':
			i++
		default:
			b.WriteByte(ch)
		}
	}
	return b.String()
}

type fmtTok struct {
	kind byte // 'd' digit placeholder, 'l' literal, '.' point, ',' comma, '%', 'E' exponent, '/' fraction
	ch   byte // placeholder: 0 # ?
	lit  string
	sign string // exponent: "+" or "-"
}

func tokenizeNumber(sec string) []fmtTok {
	var ts []fmtTok
	for i := 0; i < len(sec); i++ {
		ch := sec[i]
		switch ch {
		case '0', '#', '?':
			ts = append(ts, fmtTok{kind: 'd', ch: ch})
		case '.':
			ts = append(ts, fmtTok{kind: '.'})
		case ',':
			ts = append(ts, fmtTok{kind: ','})
		case '%':
			ts = append(ts, fmtTok{kind: '%'})
		case 'E', 'e':
			if i+1 < len(sec) && (sec[i+1] == '+' || sec[i+1] == '-') {
				ts = append(ts, fmtTok{kind: 'E', sign: string(sec[i+1])})
				i++
			} else {
				ts = append(ts, fmtTok{kind: 'l', lit: sec[i : i+1]})
			}
		case '/':
			ts = append(ts, fmtTok{kind: '/'})
		case '"':
			j := strings.IndexByte(sec[i+1:], '"')
			if j < 0 {
				ts = append(ts, fmtTok{kind: 'l', lit: sec[i+1:]})
				i = len(sec)
				continue
			}
			ts = append(ts, fmtTok{kind: 'l', lit: sec[i+1 : i+1+j]})
			i += j + 1
		case '\\':
			if i+1 < len(sec) {
				ts = append(ts, fmtTok{kind: 'l', lit: string(sec[i+1])})
				i++
			}
		case '_':
			ts = append(ts, fmtTok{kind: 'l', lit: " "})
			i++
		case '*':
			i++
		case '@':
		default:
			ts = append(ts, fmtTok{kind: 'l', lit: sec[i : i+1]})
		}
	}
	return ts
}

func formatDigits(v float64, sec string) string {
	ts := tokenizeNumber(sec)
	// Fractions (# ?/?) are rare; show the number with as many decimals.
	for _, t := range ts {
		if t.kind == '/' {
			return formatFraction(v, ts)
		}
	}
	point, exp := -1, -1
	for i, t := range ts {
		if t.kind == '.' && point < 0 {
			point = i
		}
		if t.kind == 'E' && exp < 0 {
			exp = i
		}
	}
	intEnd := len(ts)
	if point >= 0 {
		intEnd = point
	} else if exp >= 0 {
		intEnd = exp
	}
	fracEnd := len(ts)
	if exp >= 0 {
		fracEnd = exp
	}
	// Commas: between digit placeholders they group thousands; right
	// after the last integer placeholder each one divides by 1,000.
	grouping := false
	lastDigit := -1
	for i := 0; i < intEnd; i++ {
		if ts[i].kind == 'd' {
			lastDigit = i
		}
	}
	scale := 0
	for i := 0; i < intEnd; i++ {
		if ts[i].kind != ',' {
			continue
		}
		if i > lastDigit {
			scale++
		} else {
			grouping = true
		}
	}
	for _, t := range ts {
		if t.kind == '%' {
			v *= 100
		}
	}
	v /= math.Pow(1000, float64(scale))
	decimals := 0
	if point >= 0 {
		for i := point + 1; i < fracEnd; i++ {
			if ts[i].kind == 'd' {
				decimals++
			}
		}
	}
	expVal := 0
	if exp >= 0 {
		intDigits := 0
		for i := 0; i < intEnd; i++ {
			if ts[i].kind == 'd' {
				intDigits++
			}
		}
		if intDigits == 0 {
			intDigits = 1
		}
		if v != 0 {
			e := int(math.Floor(math.Log10(math.Abs(v))))
			// The exponent is a multiple of the integer digits' count
			// (##0.0E+0 gives engineering notation).
			expVal = e - (((e % intDigits) + intDigits) % intDigits)
			if intDigits == 1 {
				expVal = e
			}
			v /= math.Pow(10, float64(expVal))
		}
	}
	r := roundDigits(math.Abs(v), decimals, 0)
	s := strconv.FormatFloat(r, 'f', decimals, 64)
	intPart, fracPart := s, ""
	if k := strings.IndexByte(s, '.'); k >= 0 {
		intPart, fracPart = s[:k], s[k+1:]
	}
	if intPart == "0" {
		intPart = ""
	}
	var b strings.Builder
	// Integer placeholders, filled from the right; the first takes any
	// digits left over.
	var slots []int
	for i := 0; i < intEnd; i++ {
		if ts[i].kind == 'd' {
			slots = append(slots, i)
		}
	}
	digitAt := map[int]string{}
	if len(slots) > 0 {
		pos := len(intPart)
		for k := len(slots) - 1; k >= 0; k-- {
			t := ts[slots[k]]
			var d string
			if k == 0 {
				d = intPart[:pos]
				pos = 0
			} else if pos > 0 {
				d = intPart[pos-1 : pos]
				pos--
			}
			if d == "" {
				switch t.ch {
				case '0':
					d = "0"
				case '?':
					d = " "
				}
			}
			digitAt[slots[k]] = d
		}
	}
	// Place values for grouping: count digits from the right.
	totalDigits := 0
	for _, i := range slots {
		totalDigits += len(strings.TrimLeft(digitAt[i], " "))
	}
	written := 0
	for i := 0; i < intEnd; i++ {
		t := ts[i]
		switch t.kind {
		case 'd':
			d := digitAt[i]
			for _, ch := range d {
				b.WriteRune(ch)
				if ch >= '0' && ch <= '9' {
					written++
					place := totalDigits - written
					if grouping && place > 0 && place%3 == 0 {
						b.WriteByte(',')
					}
				}
			}
		case 'l':
			b.WriteString(t.lit)
		case '%':
			b.WriteByte('%')
		}
	}
	if point >= 0 {
		b.WriteByte('.')
		fi := 0
		// Trailing # drop zeros; trailing ? become spaces.
		lastSig := len(strings.TrimRight(fracPart, "0"))
		for i := point + 1; i < fracEnd; i++ {
			t := ts[i]
			switch t.kind {
			case 'd':
				if fi < len(fracPart) {
					ch := fracPart[fi]
					if fi >= lastSig && t.ch == '#' {
					} else if fi >= lastSig && t.ch == '?' {
						b.WriteByte(' ')
					} else {
						b.WriteByte(ch)
					}
				}
				fi++
			case 'l':
				b.WriteString(t.lit)
			case '%':
				b.WriteByte('%')
			}
		}
	}
	if exp >= 0 {
		b.WriteByte('E')
		e := expVal
		if e < 0 {
			b.WriteByte('-')
			e = -e
		} else if ts[exp].sign == "+" {
			b.WriteByte('+')
		}
		digits := 0
		for i := exp + 1; i < len(ts) && ts[i].kind == 'd'; i++ {
			digits++
		}
		es := strconv.Itoa(e)
		for len(es) < digits {
			es = "0" + es
		}
		b.WriteString(es)
		for i := exp + 1 + digits; i < len(ts); i++ {
			if ts[i].kind == 'l' {
				b.WriteString(ts[i].lit)
			}
		}
	}
	return b.String()
}

func formatFraction(v float64, ts []fmtTok) string {
	whole := math.Trunc(v)
	frac := math.Abs(v - whole)
	den := 0
	for i, t := range ts {
		if t.kind == '/' {
			for j := i + 1; j < len(ts) && ts[j].kind == 'd'; j++ {
				den++
			}
		}
	}
	maxDen := int(math.Pow(10, float64(den))) - 1
	if maxDen < 1 {
		maxDen = 9
	}
	bestN, bestD, bestErr := 0, 1, math.Inf(1)
	for d := 1; d <= maxDen; d++ {
		n := math.Round(frac * float64(d))
		if e := math.Abs(frac - n/float64(d)); e < bestErr-1e-12 {
			bestN, bestD, bestErr = int(n), d, e
		}
	}
	if bestN == bestD {
		whole += math.Copysign(1, v)
		bestN = 0
	}
	if bestN == 0 {
		return strconv.Itoa(int(whole))
	}
	if whole == 0 {
		return strconv.Itoa(bestN) + "/" + strconv.Itoa(bestD)
	}
	return strconv.Itoa(int(whole)) + " " + strconv.Itoa(bestN) + "/" + strconv.Itoa(bestD)
}

var dayNames = []string{"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"}
var monthTitles = []string{"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"}

// formatDate renders a serial with date and time codes.
func formatDate(v float64, sec string, date1904 bool) string {
	if v < 0 {
		return strings.Repeat("#", 8)
	}
	if date1904 {
		v += 1462
	}
	// Round to the shown precision: whole seconds unless the code shows
	// fractions.
	fracDigits := 0
	if i := strings.Index(strings.ToLower(sec), "s."); i >= 0 {
		for j := i + 2; j < len(sec) && sec[j] == '0'; j++ {
			fracDigits++
		}
	}
	// Count in whole ticks (seconds, or their shown fractions), so that
	// 0.0219444... days is 31:37, not 31:36.999.
	scale := math.Pow(10, float64(fracDigits))
	ticks := math.Round(v * 86400 * scale)
	days := math.Floor(ticks / (86400 * scale))
	rem := ticks - days*86400*scale
	v = ticks / (86400 * scale)
	y, mo, d := civil(days)
	secs := int(math.Floor(rem / scale))
	h := secs / 3600
	mi := secs / 60 % 60
	s := secs % 60
	sub := (rem - float64(secs)*scale) / scale
	ampm := strings.Contains(strings.ToUpper(sec), "AM/PM") || strings.Contains(strings.ToUpper(sec), "A/P")
	// Tokens.
	type tok struct {
		kind byte // y m d h s a(ampm) l(literal) e(elapsed) f(fraction)
		n    int
		lit  string
	}
	var ts []tok
	for i := 0; i < len(sec); i++ {
		ch := sec[i]
		lower := ch | 0x20
		switch {
		case ch == '"':
			j := strings.IndexByte(sec[i+1:], '"')
			if j < 0 {
				ts = append(ts, tok{kind: 'l', lit: sec[i+1:]})
				i = len(sec)
				continue
			}
			ts = append(ts, tok{kind: 'l', lit: sec[i+1 : i+1+j]})
			i += j + 1
		case ch == '\\':
			if i+1 < len(sec) {
				ts = append(ts, tok{kind: 'l', lit: string(sec[i+1])})
				i++
			}
		case ch == '[':
			j := strings.IndexByte(sec[i:], ']')
			if j > 0 {
				inner := strings.ToLower(sec[i+1 : i+j])
				ts = append(ts, tok{kind: 'e', lit: inner[:1], n: len(inner)})
				i += j
			}
		case lower == 'y' || lower == 'm' || lower == 'd' || lower == 'h' || lower == 's':
			j := i
			for j < len(sec) && sec[j]|0x20 == lower {
				j++
			}
			ts = append(ts, tok{kind: lower, n: j - i})
			i = j - 1
		case strings.HasPrefix(strings.ToUpper(sec[i:]), "AM/PM"):
			ts = append(ts, tok{kind: 'a', n: 5})
			i += 4
		case strings.HasPrefix(strings.ToUpper(sec[i:]), "A/P"):
			ts = append(ts, tok{kind: 'a', n: 3})
			i += 2
		case ch == '.' && i+1 < len(sec) && sec[i+1] == '0':
			j := i + 1
			for j < len(sec) && sec[j] == '0' {
				j++
			}
			ts = append(ts, tok{kind: 'f', n: j - i - 1})
			i = j - 1
		case ch == '_':
			ts = append(ts, tok{kind: 'l', lit: " "})
			i++
		case ch == '*':
			i++
		default:
			ts = append(ts, tok{kind: 'l', lit: sec[i : i+1]})
		}
	}
	// m after h, or before s, is minutes.
	for i, t := range ts {
		if t.kind != 'm' {
			continue
		}
		minute := false
		for j := i - 1; j >= 0; j-- {
			if ts[j].kind == 'l' {
				continue
			}
			minute = ts[j].kind == 'h' || ts[j].kind == 'e' && ts[j].lit == "h"
			break
		}
		for j := i + 1; j < len(ts) && !minute; j++ {
			if ts[j].kind == 'l' {
				continue
			}
			minute = ts[j].kind == 's' || ts[j].kind == 'e' && ts[j].lit == "s"
			break
		}
		if minute {
			ts[i].kind = 'n'
		}
	}
	pad := func(x, n int) string {
		s := strconv.Itoa(x)
		for len(s) < n {
			s = "0" + s
		}
		return s
	}
	var b strings.Builder
	for _, t := range ts {
		switch t.kind {
		case 'l':
			b.WriteString(t.lit)
		case 'y':
			if t.n <= 2 {
				b.WriteString(pad(y%100, 2))
			} else {
				b.WriteString(pad(y, 4))
			}
		case 'm':
			switch {
			case t.n == 1:
				b.WriteString(strconv.Itoa(mo))
			case t.n == 2:
				b.WriteString(pad(mo, 2))
			case t.n == 3:
				b.WriteString(monthTitles[mo-1][:3])
			case t.n == 5:
				b.WriteString(monthTitles[mo-1][:1])
			default:
				b.WriteString(monthTitles[mo-1])
			}
		case 'd':
			switch {
			case t.n == 1:
				b.WriteString(strconv.Itoa(d))
			case t.n == 2:
				b.WriteString(pad(d, 2))
			case t.n == 3:
				b.WriteString(dayNames[weekday(v)][:3])
			default:
				b.WriteString(dayNames[weekday(v)])
			}
		case 'h':
			hh := h
			if ampm {
				hh = h % 12
				if hh == 0 {
					hh = 12
				}
			}
			if t.n >= 2 {
				b.WriteString(pad(hh, 2))
			} else {
				b.WriteString(strconv.Itoa(hh))
			}
		case 'n':
			if t.n >= 2 {
				b.WriteString(pad(mi, 2))
			} else {
				b.WriteString(strconv.Itoa(mi))
			}
		case 's':
			if t.n >= 2 {
				b.WriteString(pad(s, 2))
			} else {
				b.WriteString(strconv.Itoa(s))
			}
		case 'f':
			f := strconv.FormatFloat(sub, 'f', t.n, 64)
			b.WriteString(f[strings.IndexByte(f, '.'):])
		case 'e':
			var x int
			switch t.lit {
			case "h":
				x = int(math.Floor(v*24 + 1e-9))
			case "m":
				x = int(math.Floor(v*1440 + 1e-9))
			default:
				x = int(math.Floor(v*86400 + 1e-9))
			}
			b.WriteString(pad(x, t.n))
		case 'a':
			if t.n == 5 {
				if h < 12 {
					b.WriteString("AM")
				} else {
					b.WriteString("PM")
				}
			} else if h < 12 {
				b.WriteString("A")
			} else {
				b.WriteString("P")
			}
		}
	}
	return b.String()
}

// weekday gives 0 for Sunday .. 6 for Saturday. Excel counts serial 1 as
// a Sunday and every serial on from it, the phantom 29 February included.
func weekday(serial float64) int {
	n := int(math.Floor(serial))
	return ((n-1)%7 + 7) % 7
}

package engine

import (
	"strings"
	"unicode"
	"unicode/utf8"

	"golang.org/x/text/encoding/charmap"
)

// textFn makes a function whose arguments are read as text.
func textFn(min, max int, f func(c *ctx, s []string, a []Value) Value) *fnSpec {
	return &fnSpec{min: min, max: max, f: func(c *ctx, a []Value) Value {
		s := make([]string, len(a))
		for i, v := range a {
			if isMissing(v) {
				continue
			}
			t, err := toText(v)
			if err != nil {
				return *err
			}
			s[i] = t
		}
		return f(c, s, a)
	}}
}

func runes(s string) []rune { return []rune(s) }

func init() {
	reg("LEN", textFn(1, 1, func(c *ctx, s []string, a []Value) Value { return num(float64(utf8.RuneCountInString(s[0]))) }))
	reg("LOWER", textFn(1, 1, func(c *ctx, s []string, a []Value) Value { return str(strings.ToLower(s[0])) }))
	reg("UPPER", textFn(1, 1, func(c *ctx, s []string, a []Value) Value { return str(strings.ToUpper(s[0])) }))
	reg("PROPER", textFn(1, 1, func(c *ctx, s []string, a []Value) Value {
		var b strings.Builder
		prevLetter := false
		for _, r := range s[0] {
			if unicode.IsLetter(r) {
				if prevLetter {
					b.WriteRune(unicode.ToLower(r))
				} else {
					b.WriteRune(unicode.ToUpper(r))
				}
				prevLetter = true
			} else {
				b.WriteRune(r)
				prevLetter = false
			}
		}
		return str(b.String())
	}))
	reg("TRIM", textFn(1, 1, func(c *ctx, s []string, a []Value) Value {
		return str(strings.Join(strings.FieldsFunc(s[0], func(r rune) bool { return r == ' ' }), " "))
	}))
	reg("CLEAN", textFn(1, 1, func(c *ctx, s []string, a []Value) Value {
		return str(strings.Map(func(r rune) rune {
			if r < 32 {
				return -1
			}
			return r
		}, s[0]))
	}))
	reg("LEFT", textFn(1, 2, func(c *ctx, s []string, a []Value) Value {
		n, err := numArg(a, 1, 1)
		if err != nil {
			return *err
		}
		if n < 0 {
			return errv(ErrValue)
		}
		r := runes(s[0])
		return str(string(r[:min(int(n), len(r))]))
	}))
	reg("RIGHT", textFn(1, 2, func(c *ctx, s []string, a []Value) Value {
		n, err := numArg(a, 1, 1)
		if err != nil {
			return *err
		}
		if n < 0 {
			return errv(ErrValue)
		}
		r := runes(s[0])
		k := min(int(n), len(r))
		return str(string(r[len(r)-k:]))
	}))
	reg("MID", textFn(3, 3, func(c *ctx, s []string, a []Value) Value {
		start, err := toNumber(a[1])
		if err != nil {
			return *err
		}
		n, err := toNumber(a[2])
		if err != nil {
			return *err
		}
		if start < 1 || n < 0 {
			return errv(ErrValue)
		}
		r := runes(s[0])
		i := int(start) - 1
		if i >= len(r) {
			return str("")
		}
		return str(string(r[i:min(len(r), i+int(n))]))
	}))
	reg("REPT", textFn(2, 2, func(c *ctx, s []string, a []Value) Value {
		n, err := toNumber(a[1])
		if err != nil {
			return *err
		}
		if n < 0 || float64(len(s[0]))*n > 32767 {
			return errv(ErrValue)
		}
		return str(strings.Repeat(s[0], int(n)))
	}))
	reg("EXACT", textFn(2, 2, func(c *ctx, s []string, a []Value) Value { return boolean(s[0] == s[1]) }))
	reg("CONCATENATE", textFn(1, -1, func(c *ctx, s []string, a []Value) Value { return str(strings.Join(s, "")) }))
	reg("CONCAT", &fnSpec{min: 1, max: -1, args: []argKind{aRef}, f: func(c *ctx, a []Value) Value {
		var b strings.Builder
		var bad *Value
		for _, v := range a {
			c.each(v, func(x Value, _ bool) bool {
				t, err := toText(x)
				if err != nil {
					bad = err
					return false
				}
				b.WriteString(t)
				return true
			})
			if bad != nil {
				return *bad
			}
		}
		return str(b.String())
	}})
	reg("TEXTJOIN", &fnSpec{min: 3, max: -1, args: []argKind{aVal, aVal, aRef}, f: func(c *ctx, a []Value) Value {
		delim, err := toText(a[0])
		if err != nil {
			return *err
		}
		// TEXTJOIN(",",,A1:A5): an empty ignore_empty is TRUE.
		skip := true
		if !isMissing(a[1]) {
			b, err := toBool(a[1])
			if err != nil {
				return *err
			}
			skip = b
		}
		var parts []string
		var bad *Value
		for _, v := range a[2:] {
			c.each(v, func(x Value, _ bool) bool {
				t, err := toText(x)
				if err != nil {
					bad = err
					return false
				}
				if t == "" && skip {
					return true
				}
				parts = append(parts, t)
				return true
			})
			if bad != nil {
				return *bad
			}
		}
		out := strings.Join(parts, delim)
		if len(out) > 32767 {
			return errv(ErrValue)
		}
		return str(out)
	}})
	reg("SUBSTITUTE", textFn(3, 4, func(c *ctx, s []string, a []Value) Value {
		text, old, new := s[0], s[1], s[2]
		if old == "" {
			return str(text)
		}
		inst, ok := opt(a, 3)
		if !ok {
			return str(strings.ReplaceAll(text, old, new))
		}
		n, err := toNumber(inst)
		if err != nil {
			return *err
		}
		if n < 1 {
			return errv(ErrValue)
		}
		k := int(n)
		idx := -1
		from := 0
		for i := 0; i < k; i++ {
			j := strings.Index(text[from:], old)
			if j < 0 {
				return str(text)
			}
			idx = from + j
			from = idx + len(old)
		}
		return str(text[:idx] + new + text[idx+len(old):])
	}))
	reg("REPLACE", textFn(4, 4, func(c *ctx, s []string, a []Value) Value {
		start, err := toNumber(a[1])
		if err != nil {
			return *err
		}
		n, err := toNumber(a[2])
		if err != nil {
			return *err
		}
		if start < 1 || n < 0 {
			return errv(ErrValue)
		}
		r := runes(s[0])
		i := min(int(start)-1, len(r))
		j := min(i+int(n), len(r))
		return str(string(r[:i]) + s[3] + string(r[j:]))
	}))
	find := func(insensitive bool) *fnSpec {
		return textFn(2, 3, func(c *ctx, s []string, a []Value) Value {
			start, err := numArg(a, 2, 1)
			if err != nil {
				return *err
			}
			within := runes(s[1])
			if start < 1 || int(start) > len(within)+1 {
				return errv(ErrValue)
			}
			from := int(start) - 1
			if s[0] == "" {
				return num(float64(from + 1))
			}
			if insensitive {
				pat := foldCase(s[0])
				text := runes(foldCase(s[1]))
				if strings.ContainsAny(s[0], "*?~") {
					for i := from; i < len(text); i++ {
						if wildPrefix(pat, string(text[i:])) {
							return num(float64(i + 1))
						}
					}
					return errv(ErrValue)
				}
				if k := strings.Index(string(text[from:]), pat); k >= 0 {
					return num(float64(from + utf8.RuneCountInString(string(text[from:])[:k]) + 1))
				}
				return errv(ErrValue)
			}
			rest := string(within[from:])
			if k := strings.Index(rest, s[0]); k >= 0 {
				return num(float64(from + utf8.RuneCountInString(rest[:k]) + 1))
			}
			return errv(ErrValue)
		})
	}
	reg("FIND", find(false))
	reg("SEARCH", find(true))
	reg("VALUE", &fnSpec{min: 1, max: 1, f: func(c *ctx, a []Value) Value {
		switch v := a[0]; v.Kind {
		case Number:
			return v
		case Empty:
			return num(0)
		case Text:
			n, ok, dependent := readNumber(v.Str)
			switch {
			case dependent:
				return languageDependent(v.Str)
			case ok:
				return num(n)
			}
			return errv(ErrValue)
		case Error:
			return v
		}
		return errv(ErrValue)
	}})
	reg("NUMBERVALUE", textFn(1, 3, func(c *ctx, s []string, a []Value) Value {
		dec, err := textArg(a, 1, ".")
		if err != nil {
			return *err
		}
		group, err := textArg(a, 2, ",")
		if err != nil {
			return *err
		}
		t := strings.ReplaceAll(strings.ReplaceAll(s[0], " ", ""), group, "")
		if dec != "." {
			t = strings.ReplaceAll(t, dec, ".")
		}
		if t == "" {
			return num(0)
		}
		if n, ok := textNumber(t); ok {
			return num(n)
		}
		return errv(ErrValue)
	}))
	reg("TEXT", &fnSpec{min: 2, max: 2, f: func(c *ctx, a []Value) Value {
		code, err := toText(a[1])
		if err != nil {
			return *err
		}
		switch v := a[0]; v.Kind {
		case Error:
			return v
		case Text:
			n, ok, dependent := readNumber(v.Str)
			switch {
			case dependent:
				return languageDependent(v.Str)
			case ok:
				return str(FormatNumber(n, code, c.e.Date1904))
			}
			return str(FormatText(v.Str, code))
		case Bool:
			s, _ := toText(v)
			return str(FormatText(s, code))
		}
		n, _ := toNumber(a[0])
		return str(FormatNumber(n, code, c.e.Date1904))
	}})
	reg("FIXED", &fnSpec{min: 1, max: 3, f: func(c *ctx, a []Value) Value {
		x, err := toNumber(a[0])
		if err != nil {
			return *err
		}
		d, err := numArg(a, 1, 2)
		if err != nil {
			return *err
		}
		noCommas, err := boolArg(a, 2, false)
		if err != nil {
			return *err
		}
		code := "#,##0"
		if noCommas {
			code = "0"
		}
		if d > 0 {
			code += "." + strings.Repeat("0", int(d))
		}
		x = roundDigits(x, int(d), 0)
		return str(FormatNumber(x, code, false))
	}})
	reg("DOLLAR", &fnSpec{min: 1, max: 2, f: func(c *ctx, a []Value) Value {
		x, err := toNumber(a[0])
		if err != nil {
			return *err
		}
		d, err := numArg(a, 1, 2)
		if err != nil {
			return *err
		}
		code := "$#,##0"
		if d > 0 {
			code += "." + strings.Repeat("0", int(d))
		}
		x = roundDigits(x, int(d), 0)
		return str(FormatNumber(x, code+";("+code+")", false))
	}})
	reg("CHAR", fn1(func(x float64) Value {
		n := int(x)
		if n < 1 || n > 255 {
			return errv(ErrValue)
		}
		r := charmap.Windows1252.DecodeByte(byte(n))
		return str(string(r))
	}))
	reg("UNICHAR", fn1(func(x float64) Value {
		n := int(x)
		if n < 1 || n > 0x10FFFF {
			return errv(ErrValue)
		}
		return str(string(rune(n)))
	}))
	reg("CODE", textFn(1, 1, func(c *ctx, s []string, a []Value) Value {
		if s[0] == "" {
			return errv(ErrValue)
		}
		r, _ := utf8.DecodeRuneInString(s[0])
		if b, ok := charmap.Windows1252.EncodeRune(r); ok {
			return num(float64(b))
		}
		return num(63)
	}))
	reg("UNICODE", textFn(1, 1, func(c *ctx, s []string, a []Value) Value {
		if s[0] == "" {
			return errv(ErrValue)
		}
		r, _ := utf8.DecodeRuneInString(s[0])
		return num(float64(r))
	}))
	before := func(after bool) *fnSpec {
		return &fnSpec{min: 2, max: 6, f: func(c *ctx, a []Value) Value {
			text, err := toText(a[0])
			if err != nil {
				return *err
			}
			delim, err := toText(a[1])
			if err != nil {
				return *err
			}
			inst, err := numArg(a, 2, 1)
			if err != nil {
				return *err
			}
			mode, err := numArg(a, 3, 0)
			if err != nil {
				return *err
			}
			hay, needle := text, delim
			if mode == 1 {
				hay, needle = foldCase(text), foldCase(delim)
			}
			k := int(inst)
			if k == 0 {
				return errv(ErrValue)
			}
			var idx []int
			for from := 0; ; {
				j := strings.Index(hay[from:], needle)
				if j < 0 || needle == "" && from > len(hay) {
					break
				}
				idx = append(idx, from+j)
				from += j + max(len(needle), 1)
				if from > len(hay) {
					break
				}
			}
			var at int
			switch {
			case k > 0 && k <= len(idx):
				at = idx[k-1]
			case k < 0 && -k <= len(idx):
				at = idx[len(idx)+k]
			default:
				if v, ok := opt(a, 5); ok {
					return v
				}
				return errv(ErrNA)
			}
			if after {
				return str(text[at+len(delim):])
			}
			return str(text[:at])
		}}
	}
	reg("TEXTBEFORE", before(false))
	reg("TEXTAFTER", before(true))
	reg("TEXTSPLIT", &fnSpec{min: 2, max: 6, f: func(c *ctx, a []Value) Value {
		text, err := toText(a[0])
		if err != nil {
			return *err
		}
		colD, err := textArg(a, 1, "")
		if err != nil {
			return *err
		}
		rowD, err := textArg(a, 2, "")
		if err != nil {
			return *err
		}
		ignore, err := boolArg(a, 3, false)
		if err != nil {
			return *err
		}
		split := func(s, d string) []string {
			if d == "" {
				return []string{s}
			}
			parts := strings.Split(s, d)
			if ignore {
				var out []string
				for _, p := range parts {
					if p != "" {
						out = append(out, p)
					}
				}
				return out
			}
			return parts
		}
		rows := split(text, rowD)
		var grid [][]string
		cols := 0
		for _, r := range rows {
			cs := split(r, colD)
			grid = append(grid, cs)
			cols = max(cols, len(cs))
		}
		if len(grid) == 0 || cols == 0 {
			return errv(ErrCalc)
		}
		out := newArr(len(grid), cols)
		for i, r := range grid {
			for j := 0; j < cols; j++ {
				if j < len(r) {
					out.V[i*cols+j] = str(r[j])
				} else {
					out.V[i*cols+j] = errv(ErrNA)
				}
			}
		}
		return arrOf(out)
	}})
}

// wildPrefix tells whether a wildcard pattern matches some start of text.
func wildPrefix(pattern, text string) bool {
	return wildMatch(pattern+"*", text)
}

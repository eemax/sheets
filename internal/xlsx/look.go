package xlsx

import (
	"fmt"
	"math"
	"sort"
	"strconv"
	"strings"

	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/style"
)

// A Look is styling to lay over what cells already have: what sheets style
// sets, and what fmt replays from the record of it. Unset fields leave the
// cell as it is.
type Look struct {
	Reset                           bool // clear the cells' look first (number formats stay)
	Bold, Italic, Underline, Strike *bool
	Font                            string
	Size                            float64
	Color, Fill                     string // RRGGBB; Fill "none" removes the fill
	NumFmt                          string // a number format code; "General" clears it
	Align, VAlign                   string // as excelize names them
	Wrap                            *bool
	Indent, Rotate                  *int
	Borders                         []Border
	Width                           float64 // column width in characters; -1 fits the content
	Height                          float64 // row height in points
	Merge                           *bool   // merge the range (true) or undo merges in it (false)

	words []string // the properties as given
}

// A Border sets the lines of some sides of each cell.
type Border struct {
	Sides string // all, outline, inner, top, bottom, left, right, none
	Style int    // excelize line style; 0 removes the line
	Color string // RRGGBB
}

// Words returns the properties as given.
func (l *Look) Words() []string { return l.words }

// CellStyle reports whether the look changes cell styles (rather than only
// widths, heights or merges).
func (l *Look) CellStyle() bool {
	return l.Reset || l.Bold != nil || l.Italic != nil || l.Underline != nil || l.Strike != nil ||
		l.Font != "" || l.Size > 0 || l.Color != "" || l.Fill != "" || l.NumFmt != "" ||
		l.Align != "" || l.VAlign != "" || l.Wrap != nil || l.Indent != nil || l.Rotate != nil || len(l.Borders) > 0
}

// keys names what a look sets, to tell when a later one sets all of it
// again.
func (l *Look) keys() map[string]bool {
	k := map[string]bool{}
	add := func(set bool, name string) {
		if set {
			k[name] = true
		}
	}
	add(l.Reset, "reset")
	add(l.Bold != nil, "bold")
	add(l.Italic != nil, "italic")
	add(l.Underline != nil, "underline")
	add(l.Strike != nil, "strike")
	add(l.Font != "", "font")
	add(l.Size > 0, "size")
	add(l.Color != "", "color")
	add(l.Fill != "", "fill")
	add(l.NumFmt != "", "format")
	add(l.Align != "", "align")
	add(l.VAlign != "", "valign")
	add(l.Wrap != nil, "wrap")
	add(l.Indent != nil, "indent")
	add(l.Rotate != nil, "rotate")
	add(l.Width != 0, "width")
	add(l.Height != 0, "height")
	add(l.Merge != nil, "merge")
	for _, b := range l.Borders {
		k["border="+b.Sides] = true
	}
	return k
}

// lookFlags are the properties written as a bare word.
var lookFlags = map[string]func(*Look){
	"bold": func(l *Look) { l.Bold = ptr(true) }, "nobold": func(l *Look) { l.Bold = ptr(false) },
	"italic": func(l *Look) { l.Italic = ptr(true) }, "noitalic": func(l *Look) { l.Italic = ptr(false) },
	"underline": func(l *Look) { l.Underline = ptr(true) }, "nounderline": func(l *Look) { l.Underline = ptr(false) },
	"strike": func(l *Look) { l.Strike = ptr(true) }, "nostrike": func(l *Look) { l.Strike = ptr(false) },
	"wrap": func(l *Look) { l.Wrap = ptr(true) }, "nowrap": func(l *Look) { l.Wrap = ptr(false) },
	"merge": func(l *Look) { l.Merge = ptr(true) }, "unmerge": func(l *Look) { l.Merge = ptr(false) },
	"reset": func(l *Look) { l.Reset = true },
}

// lookKeys are the properties written key=value.
var lookKeys = []string{"bold", "italic", "underline", "strike", "wrap", "font", "size", "color", "fill", "format", "decimals",
	"align", "valign", "indent", "rotate", "border", "width", "height"}

func ptr[T any](v T) *T { return &v }

// IsLookWord tells whether a word is a style property.
func IsLookWord(w string) bool {
	if _, ok := lookFlags[strings.ToLower(w)]; ok {
		return true
	}
	k, _, ok := strings.Cut(w, "=")
	if !ok {
		return false
	}
	k = strings.ToLower(k)
	for _, name := range lookKeys {
		if k == name {
			return true
		}
	}
	return false
}

// ParseLook reads style properties: bold, italic, underline, strike, wrap
// (and nobold and so on), merge, unmerge, reset, and key=value for font,
// size, color, fill, format, decimals, align, valign, indent, rotate,
// border, width and height.
func ParseLook(ws []string) (Look, error) {
	l := Look{words: append([]string(nil), ws...)}
	var (
		kind     grid.Format
		hasKind  bool
		decimals = -1
	)
	for _, w := range ws {
		if set, ok := lookFlags[strings.ToLower(w)]; ok {
			set(&l)
			continue
		}
		key, val, ok := strings.Cut(w, "=")
		if !ok {
			return l, fmt.Errorf("%q is not a style property (%s)", w, lookHelp)
		}
		key = strings.ToLower(strings.TrimSpace(key))
		var err error
		switch key {
		case "bold":
			l.Bold, err = parseBool(w, val)
		case "italic":
			l.Italic, err = parseBool(w, val)
		case "underline":
			l.Underline, err = parseBool(w, val)
		case "strike":
			l.Strike, err = parseBool(w, val)
		case "wrap":
			l.Wrap, err = parseBool(w, val)
		case "font":
			if l.Font = strings.TrimSpace(val); l.Font == "" {
				err = fmt.Errorf("%s: give a font name", w)
			}
		case "size":
			l.Size, err = parseNumber(w, val, 1, 409)
		case "color":
			l.Color, err = ParseColor(val)
		case "fill":
			if strings.EqualFold(val, "none") {
				l.Fill = "none"
			} else {
				l.Fill, err = ParseColor(val)
			}
		case "format":
			if k, sym, ok := parseKind(val); ok {
				kind, hasKind = grid.Format{Kind: k, Symbol: sym}, true
			} else if strings.EqualFold(val, "general") {
				l.NumFmt = "General"
			} else if err = CheckNumFmt(val); err == nil {
				l.NumFmt = val
			} else {
				err = fmt.Errorf("format=%s: %v (or a type: %s)", val, err, kindHelp)
			}
		case "decimals":
			var n float64
			if n, err = parseNumber(w, val, 0, 10); err == nil {
				decimals = int(n)
			}
		case "align":
			l.Align, err = parseAlign(val)
		case "valign":
			l.VAlign, err = parseVAlign(val)
		case "indent":
			var n float64
			if n, err = parseNumber(w, val, 0, 15); err == nil {
				l.Indent = ptr(int(n))
			}
		case "rotate":
			var n float64
			if n, err = parseNumber(w, val, -90, 90); err == nil {
				r := int(n)
				if r < 0 {
					r = 90 - r // Excel counts clockwise turns from 91 up
				}
				l.Rotate = &r
			}
		case "border":
			var b []Border
			if b, err = parseBorders(val); err == nil {
				l.Borders = append(l.Borders, b...)
			}
		case "width":
			if strings.EqualFold(val, "auto") {
				l.Width = -1
			} else {
				l.Width, err = parseNumber(w, val, 0.5, 255)
			}
		case "height":
			l.Height, err = parseNumber(w, val, 1, 409)
		default:
			return l, fmt.Errorf("%q is not a style property (%s)", w, lookHelp)
		}
		if err != nil {
			return l, err
		}
	}
	switch {
	case hasKind:
		if decimals >= 0 {
			kind.Decimals = decimals
		} else {
			kind.Decimals = map[grid.FormatKind]int{grid.Decimal: 2, grid.Currency: 2}[kind.Kind]
		}
		if kind.Kind.IsTemporal() && decimals > 0 {
			kind.Decimals = 1 // times with seconds
		}
		l.NumFmt = style.NumFmt(kind, false)
		if l.NumFmt == "" {
			l.NumFmt = "General"
		}
	case decimals >= 0:
		return l, fmt.Errorf("decimals= goes with a format type, like format=currency decimals=0")
	}
	if l.Merge != nil && *l.Merge && l.Align == "centerContinuous" {
		return l, fmt.Errorf("merge and align=center-across do the same; pick one (center-across keeps the cells separate)")
	}
	return l, nil
}

const (
	lookHelp = "bold, italic, underline, strike, wrap, merge, unmerge, reset, or font=, size=, color=, fill=, format=, decimals=, align=, valign=, indent=, rotate=, border=, width=, height="
	kindHelp = "integer, decimal, currency[:symbol], percent, date, datetime, time, text, general"
)

func parseKind(v string) (grid.FormatKind, string, bool) {
	name, sym, _ := strings.Cut(v, ":")
	k, ok := grid.ParseFormatKind(name)
	if !ok || k == grid.General {
		return 0, "", false
	}
	if sym != "" && k != grid.Currency {
		return 0, "", false
	}
	return k, sym, true
}

func parseBool(w, v string) (*bool, error) {
	switch strings.ToLower(v) {
	case "true", "yes", "on", "1":
		return ptr(true), nil
	case "false", "no", "off", "0":
		return ptr(false), nil
	}
	return nil, fmt.Errorf("%s: expected true or false", w)
}

func parseNumber(w, v string, lo, hi float64) (float64, error) {
	n, err := strconv.ParseFloat(strings.TrimSpace(v), 64)
	if err != nil || math.IsNaN(n) || n < lo || n > hi {
		return 0, fmt.Errorf("%s: expected a number from %s to %s", w, grid.FormatNumber(lo), grid.FormatNumber(hi))
	}
	return n, nil
}

func parseAlign(v string) (string, error) {
	switch strings.ToLower(v) {
	case "left", "center", "right", "justify", "fill", "general", "distributed":
		return strings.ToLower(v), nil
	case "centre":
		return "center", nil
	case "center-across", "centre-across", "centeracross", "center-across-selection":
		return "centerContinuous", nil
	}
	return "", fmt.Errorf("align=%s: expected left, center, right, center-across, justify, fill or general", v)
}

func parseVAlign(v string) (string, error) {
	switch strings.ToLower(v) {
	case "top", "bottom", "justify", "distributed":
		return strings.ToLower(v), nil
	case "middle", "center", "centre":
		return "center", nil
	}
	return "", fmt.Errorf("valign=%s: expected top, middle or bottom", v)
}

var borderStyles = map[string]int{"none": 0, "thin": 1, "medium": 2, "dashed": 3, "dotted": 4, "thick": 5, "double": 6, "hair": 7}

// parseBorders reads SIDES[:STYLE][:COLOR], sides joined with +
// (top+bottom), style thin by default, colour black.
func parseBorders(v string) ([]Border, error) {
	parts := strings.Split(v, ":")
	if len(parts) > 3 {
		return nil, fmt.Errorf("border=%s: expected SIDES[:STYLE][:COLOR], like bottom:double or outline:medium:#1F2937", v)
	}
	b := Border{Style: 1, Color: "000000"}
	if len(parts) > 1 {
		s, ok := borderStyles[strings.ToLower(parts[1])]
		if !ok {
			return nil, fmt.Errorf("border=%s: line style %q is not one of thin, medium, thick, double, dashed, dotted, hair, none", v, parts[1])
		}
		b.Style = s
	}
	if len(parts) > 2 {
		c, err := ParseColor(parts[2])
		if err != nil {
			return nil, fmt.Errorf("border=%s: %v", v, err)
		}
		b.Color = c
	}
	var out []Border
	for _, side := range strings.Split(strings.ToLower(parts[0]), "+") {
		switch side {
		case "none":
			b.Style = 0
			side = "all"
		case "all", "outline", "inner", "top", "bottom", "left", "right":
		case "box":
			side = "outline"
		default:
			return nil, fmt.Errorf("border=%s: side %q is not one of all, outline, inner, top, bottom, left, right, none", v, side)
		}
		x := b
		x.Sides = side
		out = append(out, x)
	}
	return out, nil
}

// colorNames are the colour names sheets knows: Excel's standard colours,
// and the light fills and dark text of its highlight presets.
var colorNames = map[string]string{
	"black": "000000", "white": "FFFFFF",
	"gray": "808080", "lightgray": "D9D9D9", "darkgray": "404040",
	"red": "FF0000", "lightred": "FFC7CE", "darkred": "9C0006", "pink": "FFC7CE",
	"green": "00B050", "lightgreen": "C6EFCE", "darkgreen": "006100",
	"blue": "0070C0", "lightblue": "DDEBF7", "darkblue": "1F4E79", "navy": "1F3864",
	"yellow": "FFFF00", "lightyellow": "FFEB9C", "darkyellow": "9C5700",
	"orange": "ED7D31", "lightorange": "FCE4D6",
	"purple": "7030A0", "lightpurple": "E4DFEC",
	"teal": "008080",
}

// ColorNames lists the colour names, for help text.
func ColorNames() []string {
	var out []string
	for n := range colorNames {
		out = append(out, n)
	}
	sort.Strings(out)
	return out
}

// ParseColor reads #RRGGBB, RRGGBB, #RGB, or a colour name, and returns
// RRGGBB.
func ParseColor(v string) (string, error) {
	name := strings.ToLower(strings.NewReplacer(" ", "", "-", "", "_", "", "grey", "gray").Replace(v))
	if hex, ok := colorNames[name]; ok {
		return hex, nil
	}
	h := strings.TrimPrefix(strings.TrimSpace(v), "#")
	if len(h) == 3 {
		h = string([]byte{h[0], h[0], h[1], h[1], h[2], h[2]})
	}
	if len(h) == 8 { // AARRGGBB
		h = h[2:]
	}
	if len(h) == 6 {
		if _, err := strconv.ParseUint(h, 16, 32); err == nil {
			return strings.ToUpper(h), nil
		}
	}
	return "", fmt.Errorf("color %q: expected #RRGGBB or a name (%s)", v, strings.Join(ColorNames(), ", "))
}

// CheckNumFmt finds what Excel would refuse in a number format code.
func CheckNumFmt(code string) error {
	if strings.TrimSpace(code) == "" {
		return fmt.Errorf("empty format")
	}
	sections := 1
	upper := strings.ToUpper(code)
	for i := 0; i < len(code); i++ {
		c := code[i]
		switch {
		case c == '"':
			j := strings.IndexByte(code[i+1:], '"')
			if j < 0 {
				return fmt.Errorf("the \" at position %d is not closed", i+1)
			}
			i += j + 1
		case c == '\\' || c == '_' || c == '*':
			if i+1 >= len(code) {
				return fmt.Errorf("%c at the end needs a character after it", c)
			}
			i++
		case c == '[':
			j := strings.IndexByte(code[i+1:], ']')
			if j <= 0 {
				return fmt.Errorf("the [ at position %d is not closed", i+1)
			}
			i += j + 1
		case c == ';':
			if sections++; sections > 4 {
				return fmt.Errorf("at most 4 sections (positive;negative;zero;text)")
			}
		case strings.HasPrefix(upper[i:], "GENERAL"):
			i += len("GENERAL") - 1
		case strings.HasPrefix(upper[i:], "AM/PM"):
			i += len("AM/PM") - 1
		case strings.HasPrefix(upper[i:], "A/P"):
			i += len("A/P") - 1
		case c >= 'a' && c <= 'z' || c >= 'A' && c <= 'Z':
			if !strings.ContainsRune("yYmMdDhHsSeEbBgG", rune(c)) {
				return fmt.Errorf("the letter %c must be quoted (\"%c\") or escaped (\\%c)", c, c, c)
			}
		}
	}
	return nil
}

// luminance is the relative luminance of an RRGGBB colour (WCAG).
func luminance(hex string) float64 {
	v, err := strconv.ParseUint(hex, 16, 32)
	if err != nil {
		return 0
	}
	ch := func(x uint64) float64 {
		c := float64(x) / 255
		if c <= 0.03928 {
			return c / 12.92
		}
		return math.Pow((c+0.055)/1.055, 2.4)
	}
	return 0.2126*ch(v>>16&0xFF) + 0.7152*ch(v>>8&0xFF) + 0.0722*ch(v&0xFF)
}

// contrast is the WCAG contrast ratio of two RRGGBB colours.
func contrast(a, b string) float64 {
	la, lb := luminance(a), luminance(b)
	if la < lb {
		la, lb = lb, la
	}
	return (la + 0.05) / (lb + 0.05)
}

// readableOn picks black or white text for a fill.
func readableOn(fill string) string {
	if contrast("000000", fill) >= contrast("FFFFFF", fill) {
		return "000000"
	}
	return "FFFFFF"
}

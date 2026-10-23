// Package style holds the house style: presets for fonts and colours, and
// the number formats sheets applies for each kind of value.
package style

import (
	"math"
	"strconv"
	"strings"
	"unicode"

	"github.com/eemax/sheets/internal/grid"
)

// Preset is a complete, opinionated look for a workbook.
type Preset struct {
	Name        string
	Description string
	Font        string
	Size        float64
	TitleSize   float64
	TitleColor  string
	HeaderFill  string
	HeaderFont  string
	BandFill    string // alternate data rows; empty disables banding
	Border      string // thin borders around table cells
	TotalBorder string // rule above a total row
	Gridlines   bool
	// Finance colour coding for body cells. Empty strings disable it.
	InputColor   string // typed-in numbers
	FormulaColor string // formulas on the same sheet
	LinkColor    string // formulas referring to other sheets or files
	MinWidth     float64
	MaxWidth     float64
}

// Default is the house style.
var Default = Preset{
	Name:        "default",
	Description: "dark header, banded rows, thin borders, no gridlines",
	Font:        "Calibri",
	Size:        11,
	TitleSize:   14,
	TitleColor:  "1F2937",
	HeaderFill:  "1F2937",
	HeaderFont:  "FFFFFF",
	BandFill:    "F3F4F6",
	Border:      "D1D5DB",
	TotalBorder: "1F2937",
	MinWidth:    8,
	MaxWidth:    60,
}

// Finance adds the banking colour convention to the house style.
var Finance = func() Preset {
	p := Default
	p.Name = "finance"
	p.Description = "house style plus blue inputs, black formulas, green cross-sheet links"
	p.InputColor = "0000FF"
	p.FormulaColor = "000000"
	p.LinkColor = "008000"
	return p
}()

// Presets lists the built-in presets.
var Presets = []Preset{Default, Finance}

// Lookup finds a preset by name.
func Lookup(name string) (Preset, bool) {
	for _, p := range Presets {
		if strings.EqualFold(p.Name, name) {
			return p, true
		}
	}
	return Preset{}, false
}

// NumFmt returns the house number format code for f. plain drops the
// thousands separator, for ids and years. General returns "".
func NumFmt(f grid.Format, plain bool) string {
	d := zeros(f.Decimals)
	switch f.Kind {
	case grid.Integer:
		if plain {
			return "0"
		}
		return "#,##0_);(#,##0)"
	case grid.Decimal:
		if plain {
			return "0" + d
		}
		return "#,##0" + d + "_);(#,##0" + d + ")"
	case grid.Currency:
		sym := f.Symbol
		if sym == "" {
			sym = "$"
		}
		q := `"` + sym + `"`
		if isWordSymbol(sym) {
			return "#,##0" + d + " " + q + "_);(#,##0" + d + " " + q + ")"
		}
		return q + "#,##0" + d + "_);(" + q + "#,##0" + d + ")"
	case grid.Percent:
		return "0" + d + "%"
	case grid.Date:
		return "yyyy-mm-dd"
	case grid.DateTime:
		return "yyyy-mm-dd hh:mm"
	case grid.Time:
		if f.Decimals > 0 {
			return "hh:mm:ss"
		}
		return "hh:mm"
	case grid.TextFmt:
		return "@"
	}
	return ""
}

func zeros(n int) string {
	if n <= 0 {
		return ""
	}
	return "." + strings.Repeat("0", n)
}

func isWordSymbol(sym string) bool {
	for _, r := range sym {
		if !unicode.IsLetter(r) && r != '.' {
			return false
		}
	}
	return true
}

// Render shows v as the house format for f would, for estimating column
// widths. It need not match Excel to the pixel.
func Render(v float64, f grid.Format, plain bool, date1904 bool) string {
	switch f.Kind {
	case grid.Date, grid.DateTime, grid.Time:
		return grid.FormatTemporal(v, f.Kind, date1904)
	case grid.Percent:
		return strconv.FormatFloat(v*100, 'f', f.Decimals, 64) + "%"
	case grid.Integer, grid.Decimal, grid.Currency:
		s := strconv.FormatFloat(math.Abs(v), 'f', max(f.Decimals, 0), 64)
		if !plain {
			s = group(s)
		}
		if f.Kind == grid.Currency {
			sym := f.Symbol
			if sym == "" {
				sym = "$"
			}
			if isWordSymbol(sym) {
				s += " " + sym
			} else {
				s = sym + s
			}
		}
		if v < 0 && !plain {
			return "(" + s + ")"
		}
		if v < 0 {
			return "-" + s
		}
		if plain {
			return s
		}
		return s + " " // padding for the closing parenthesis
	}
	return grid.FormatNumber(v)
}

// group inserts thousands separators into a plain decimal string.
func group(s string) string {
	intPart, frac, hasFrac := strings.Cut(s, ".")
	if len(intPart) <= 3 {
		return s
	}
	var b strings.Builder
	pre := len(intPart) % 3
	if pre > 0 {
		b.WriteString(intPart[:pre])
	}
	for i := pre; i < len(intPart); i += 3 {
		if b.Len() > 0 {
			b.WriteByte(',')
		}
		b.WriteString(intPart[i : i+3])
	}
	if hasFrac {
		b.WriteByte('.')
		b.WriteString(frac)
	}
	return b.String()
}

// TextWidth estimates the display width of s in character units, counting
// wide (East Asian) characters double.
func TextWidth(s string) float64 {
	w := 0.0
	for _, r := range s {
		switch {
		case r >= 0x1100 && (r <= 0x115F || r >= 0x2E80 && r <= 0xA4CF || r >= 0xAC00 && r <= 0xD7A3 || r >= 0xF900 && r <= 0xFAFF || r >= 0xFE30 && r <= 0xFE4F || r >= 0xFF00 && r <= 0xFF60 || r >= 0xFFE0 && r <= 0xFFE6 || r >= 0x1F300):
			w += 2
		case unicode.IsUpper(r):
			w += 1.15
		default:
			w++
		}
	}
	return w
}

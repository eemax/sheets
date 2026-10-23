package grid

import (
	"fmt"
	"strconv"
	"strings"
)

// Spreadsheet limits (xlsx). Other formats are smaller, never larger.
const (
	MaxRows = 1048576
	MaxCols = 16384
)

// ColName converts a 1-based column index to letters: 1 → A, 27 → AA.
func ColName(col int) string {
	var b [4]byte
	i := len(b)
	for col > 0 {
		col--
		i--
		b[i] = byte('A' + col%26)
		col /= 26
	}
	return string(b[i:])
}

// ColIndex converts column letters to a 1-based index. It accepts lower case.
func ColIndex(s string) (int, bool) {
	if s == "" || len(s) > 3 {
		return 0, false
	}
	n := 0
	for i := 0; i < len(s); i++ {
		c := s[i]
		switch {
		case c >= 'A' && c <= 'Z':
			n = n*26 + int(c-'A') + 1
		case c >= 'a' && c <= 'z':
			n = n*26 + int(c-'a') + 1
		default:
			return 0, false
		}
	}
	if n > MaxCols {
		return 0, false
	}
	return n, true
}

// CellName formats a 1-based (row, col) as A1 notation.
func CellName(row, col int) string {
	return ColName(col) + strconv.Itoa(row)
}

// ParseCell parses A1 notation (with optional $ markers) into 1-based row, col.
func ParseCell(s string) (row, col int, ok bool) {
	s = strings.ReplaceAll(s, "$", "")
	i := 0
	for i < len(s) && isLetter(s[i]) {
		i++
	}
	if i == 0 || i == len(s) {
		return 0, 0, false
	}
	col, ok = ColIndex(s[:i])
	if !ok {
		return 0, 0, false
	}
	row, err := strconv.Atoi(s[i:])
	if err != nil || row < 1 || row > MaxRows || s[i] == '+' || s[i] == '-' {
		return 0, 0, false
	}
	return row, col, true
}

func isLetter(c byte) bool { return c >= 'A' && c <= 'Z' || c >= 'a' && c <= 'z' }

// Range is a rectangle of cells, 1-based and inclusive.
type Range struct {
	R1, C1, R2, C2 int
}

// Rows and Cols report the size of the range.
func (r Range) Rows() int { return r.R2 - r.R1 + 1 }
func (r Range) Cols() int { return r.C2 - r.C1 + 1 }

// Empty reports whether the range covers no cells.
func (r Range) Empty() bool { return r.R2 < r.R1 || r.C2 < r.C1 }

// Contains reports whether (row, col) lies inside the range.
func (r Range) Contains(row, col int) bool {
	return row >= r.R1 && row <= r.R2 && col >= r.C1 && col <= r.C2
}

// Intersect returns the overlap of two ranges, which may be Empty.
func (r Range) Intersect(o Range) Range {
	return Range{max(r.R1, o.R1), max(r.C1, o.C1), min(r.R2, o.R2), min(r.C2, o.C2)}
}

// Union returns the bounding box of both ranges.
func (r Range) Union(o Range) Range {
	if r.Empty() {
		return o
	}
	if o.Empty() {
		return r
	}
	return Range{min(r.R1, o.R1), min(r.C1, o.C1), max(r.R2, o.R2), max(r.C2, o.C2)}
}

func (r Range) String() string {
	if r.R1 == r.R2 && r.C1 == r.C2 {
		return CellName(r.R1, r.C1)
	}
	return CellName(r.R1, r.C1) + ":" + CellName(r.R2, r.C2)
}

// ParseRange parses "A1", "A1:C5", "A:C" (whole columns) or "3:10" (whole
// rows). Whole rows/columns extend to the sheet limits; callers clip them
// against the used range.
func ParseRange(s string) (Range, error) {
	s = strings.TrimSpace(s)
	a, b, isPair := strings.Cut(s, ":")
	if !isPair {
		row, col, ok := ParseCell(a)
		if !ok {
			return Range{}, fmt.Errorf("invalid cell or range %q", s)
		}
		return Range{row, col, row, col}, nil
	}
	a, b = strings.ReplaceAll(a, "$", ""), strings.ReplaceAll(b, "$", "")
	if r1, c1, ok := ParseCell(a); ok {
		r2, c2, ok := ParseCell(b)
		if !ok {
			return Range{}, fmt.Errorf("invalid range %q", s)
		}
		return normalize(Range{r1, c1, r2, c2}), nil
	}
	if c1, ok := ColIndex(a); ok {
		c2, ok := ColIndex(b)
		if !ok {
			return Range{}, fmt.Errorf("invalid column range %q", s)
		}
		return normalize(Range{1, c1, MaxRows, c2}), nil
	}
	r1, err1 := strconv.Atoi(a)
	r2, err2 := strconv.Atoi(b)
	if err1 != nil || err2 != nil || r1 < 1 || r2 < 1 || r1 > MaxRows || r2 > MaxRows {
		return Range{}, fmt.Errorf("invalid range %q", s)
	}
	return normalize(Range{r1, 1, r2, MaxCols}), nil
}

func normalize(r Range) Range {
	if r.R1 > r.R2 {
		r.R1, r.R2 = r.R2, r.R1
	}
	if r.C1 > r.C2 {
		r.C1, r.C2 = r.C2, r.C1
	}
	return r
}

// Ref is a parsed user reference such as "Sales!A1:C9". Sheet is empty when
// the reference did not name one, and Range is nil when only a sheet was
// named.
type Ref struct {
	Sheet string
	Range *Range
}

// ParseRef splits "Sheet!A1:B2", "'My Sheet'!A1", "A1:B2" or "Sheet". Because
// a bare word like "Q1" may be either a sheet or a cell, isSheet is consulted
// first for bare words.
func ParseRef(s string, isSheet func(string) bool) (Ref, error) {
	s = strings.TrimSpace(s)
	if s == "" {
		return Ref{}, nil
	}
	if i := strings.LastIndex(s, "!"); i >= 0 {
		sheet := unquoteSheet(s[:i])
		rest := s[i+1:]
		if rest == "" {
			return Ref{Sheet: sheet}, nil
		}
		r, err := ParseRange(rest)
		if err != nil {
			return Ref{}, err
		}
		return Ref{Sheet: sheet, Range: &r}, nil
	}
	if isSheet != nil && isSheet(unquoteSheet(s)) {
		return Ref{Sheet: unquoteSheet(s)}, nil
	}
	if r, err := ParseRange(s); err == nil {
		return Ref{Range: &r}, nil
	}
	return Ref{Sheet: unquoteSheet(s)}, nil
}

func unquoteSheet(s string) string {
	if len(s) >= 2 && s[0] == '\'' && s[len(s)-1] == '\'' {
		return strings.ReplaceAll(s[1:len(s)-1], "''", "'")
	}
	return s
}

// QuoteSheet returns the sheet name quoted for use in a reference when it
// contains characters that need quoting.
func QuoteSheet(name string) string {
	plain := name != ""
	for _, r := range name {
		if !(r == '_' || r == '.' || r >= '0' && r <= '9' || r >= 'A' && r <= 'Z' || r >= 'a' && r <= 'z' || r > 127) {
			plain = false
			break
		}
	}
	if plain {
		if _, _, ok := ParseCell(name); ok {
			plain = false
		} else if name[0] >= '0' && name[0] <= '9' {
			plain = false
		}
	}
	if plain {
		return name
	}
	return "'" + strings.ReplaceAll(name, "'", "''") + "'"
}

// FormatRef renders a sheet-qualified range, e.g. Sales!A1:C9.
func FormatRef(sheet string, r Range) string {
	return QuoteSheet(sheet) + "!" + r.String()
}

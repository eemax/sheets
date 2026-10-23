package analyze

import (
	"math"
	"sort"
	"strconv"
	"strings"

	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
)

// Column summarises one column of a region's body.
type Column struct {
	Col      int
	Header   string
	Type     string          // integer, decimal, currency, percent, date, datetime, time, text, bool, mixed, empty
	Kind     grid.FormatKind // number format kind for numeric and date columns
	Symbol   string          // currency symbol
	Decimals int             // decimal places the values need (capped at 6)
	Filled   int
	Blank    int
	Errors   int
	Pending  int // formula values that could not be evaluated
	Distinct int // distinct displayed values (capped at maxDistinct)
	Min, Max float64
	Sum      float64
	Enum     bool     // few distinct values that repeat
	Values   []string // enum values by frequency, or a few samples
	MaxLen   int      // longest text value, in runes
	Mix      string   // for mixed columns, e.g. "number 95%, text 5%"
	Formulas []FormulaGroup
	// FormulaGroupsMore counts groups beyond those listed.
	FormulaGroupsMore int
}

// FormulaGroup is a run of cells in a column sharing one relative formula.
type FormulaGroup struct {
	Range   grid.Range
	Pattern string
}

const (
	maxDistinct  = 100000
	maxEnum      = 8
	maxSamples   = 3
	maxFormulaGr = 5
)

// Profile summarises the body columns of a region.
func Profile(wb *grid.Workbook, s *grid.Sheet, reg Region) []Column {
	cols := make([]Column, 0, reg.Range.Cols())
	for col := reg.Range.C1; col <= reg.Range.C2; col++ {
		c := Column{Col: col}
		if reg.HeaderRow > 0 {
			c.Header = Label(wb, s, reg.HeaderRow, col)
		}
		profileColumn(wb, s, reg.Body, &c)
		cols = append(cols, c)
	}
	return cols
}

func profileColumn(wb *grid.Workbook, s *grid.Sheet, body grid.Range, out *Column) {
	var (
		nums, texts, bools int
		kinds              = map[grid.FormatKind]int{}
		symbols            = map[string]int{}
		counts             = map[string]int{}
		order              []string
		integral           = true
		group              *FormulaGroup
		groupKey           string
	)
	out.Min, out.Max = math.Inf(1), math.Inf(-1)
	closeGroup := func() {
		if group == nil {
			return
		}
		if len(out.Formulas) < maxFormulaGr {
			out.Formulas = append(out.Formulas, *group)
		} else {
			out.FormulaGroupsMore++
		}
		group = nil
	}
	for row := body.R1; row <= body.R2; row++ {
		raw := s.At(row, out.Col)
		if raw == nil || raw.IsEmpty() {
			out.Blank++
			closeGroup()
			continue
		}
		if raw.Formula != "" {
			key, disp := formula.Pattern(raw.Formula, row, out.Col)
			if group != nil && key == groupKey && group.Range.R2 == row-1 {
				group.Range.R2 = row
			} else {
				closeGroup()
				group = &FormulaGroup{Range: grid.Range{R1: row, C1: out.Col, R2: row, C2: out.Col}, Pattern: disp}
				groupKey = key
			}
		} else {
			closeGroup()
		}
		c := wb.Value(s, row, out.Col)
		out.Filled++
		if c.Pending {
			out.Pending++
			continue
		}
		f := wb.FormatOf(c)
		switch c.Kind {
		case grid.Error:
			out.Errors++
			continue
		case grid.Number:
			nums++
			if f.Kind != grid.General && f.Kind != grid.TextFmt {
				kinds[f.Kind]++
				if f.Kind == grid.Currency {
					symbols[f.Symbol]++
				}
			}
			out.Min = math.Min(out.Min, c.Num)
			out.Max = math.Max(out.Max, c.Num)
			out.Sum += c.Num
			scale := 1.0
			if f.Kind == grid.Percent {
				scale = 100
			}
			if d := DecimalsNeeded(c.Num * scale); d > out.Decimals {
				out.Decimals = d
			}
			if c.Num != math.Trunc(c.Num) {
				integral = false
			}
		case grid.Text:
			texts++
			out.MaxLen = max(out.MaxLen, len([]rune(c.Text)))
		case grid.Bool:
			bools++
		}
		d := grid.Display(c, f, wb.Date1904)
		if _, seen := counts[d]; !seen {
			if len(counts) >= maxDistinct {
				continue
			}
			order = append(order, d)
		}
		counts[d]++
	}
	closeGroup()
	out.Distinct = len(counts)

	typed := nums + texts + bools
	switch {
	case typed == 0:
		out.Type = "empty"
	case nums == typed:
		out.Kind = majority(kinds)
		if out.Kind == grid.General {
			out.Kind = grid.Decimal
			if integral {
				out.Kind = grid.Integer
			}
		}
		if out.Kind == grid.Currency {
			out.Symbol = majorityString(symbols)
		}
		out.Type = out.Kind.String()
	case texts == typed:
		out.Type = "text"
	case bools == typed:
		out.Type = "bool"
	default:
		out.Type = "mixed"
		out.Mix = mix(typed, nums, texts, bools)
	}
	if nums == 0 {
		out.Min, out.Max = 0, 0
	}

	// Enum when few values repeat; otherwise a few samples.
	if out.Distinct > 0 && out.Distinct <= maxEnum && out.Filled >= 2*out.Distinct && out.Type != "empty" {
		out.Enum = true
		out.Values = append(out.Values, order...)
		sort.SliceStable(out.Values, func(i, j int) bool { return counts[out.Values[i]] > counts[out.Values[j]] })
	} else {
		for _, v := range order {
			if len(out.Values) == maxSamples {
				break
			}
			out.Values = append(out.Values, v)
		}
	}
}

// DecimalsNeeded returns how many decimal places represent v exactly
// enough (up to 6).
func DecimalsNeeded(v float64) int {
	v = math.Abs(v)
	for d := 0; d < 6; d++ {
		p := v * math.Pow10(d)
		if math.Abs(p-math.Round(p)) < 1e-9*math.Max(1, p) {
			return d
		}
	}
	return 6
}

func majority(m map[grid.FormatKind]int) grid.FormatKind {
	best, n := grid.General, 0
	for k, c := range m {
		if c > n || c == n && k < best {
			best, n = k, c
		}
	}
	return best
}

func majorityString(m map[string]int) string {
	best, n := "", 0
	for k, c := range m {
		if c > n || c == n && k < best {
			best, n = k, c
		}
	}
	return best
}

// mix describes the shares of each kind in a mixed column.
func mix(total, nums, texts, bools int) string {
	var parts []string
	for _, p := range []struct {
		name string
		n    int
	}{{"number", nums}, {"text", texts}, {"bool", bools}} {
		if p.n > 0 {
			pct := int(math.Round(float64(p.n) * 100 / float64(total)))
			parts = append(parts, p.name+" "+strconv.Itoa(pct)+"%")
		}
	}
	return strings.Join(parts, ", ")
}

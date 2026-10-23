package analyze

import (
	"math"
	"strings"

	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
)

// Functions whose result has the type of their arguments.
var (
	keepsAmount = setOf("SUM ROUND ROUNDUP ROUNDDOWN MROUND CEILING FLOOR CEILING.MATH FLOOR.MATH MAX MIN ABS IF IFS IFERROR IFNA AVERAGE SUMIF SUMIFS AVERAGEIF AVERAGEIFS")
	keepsDate   = setOf("EDATE EOMONTH WORKDAY WORKDAY.INTL MAX MIN IF IFS IFERROR IFNA")
)

func setOf(s string) map[string]bool {
	m := map[string]bool{}
	for _, f := range strings.Fields(s) {
		m[f] = true
	}
	return m
}

func callsOnly(f string, allowed map[string]bool) bool {
	for _, fn := range formula.Calls(f) {
		if !allowed[fn] {
			return false
		}
	}
	return true
}

// Propagate gives formula columns the type of what they compute from:
// =Units*Price is money when Price is, =Due+30 is a date. A division is
// left alone (a ratio of two amounts is not an amount), and so is date
// arithmetic other than shifting one date (=End-Start counts days).
//
// fmts holds each column's format kind and is updated in place; locked
// columns keep theirs. Only integer and decimal columns change. It reports
// which columns changed.
func Propagate(s *grid.Sheet, cols []Column, fmts []grid.Format, locked []bool) []bool {
	changed := make([]bool, len(cols))
	index := map[int]int{}
	for i, c := range cols {
		index[c.Col] = i
	}
	for pass := 0; pass < 3; pass++ {
		again := false
		for i, c := range cols {
			if locked != nil && locked[i] || len(c.Formulas) == 0 || c.FormulaGroupsMore > 0 {
				continue
			}
			if k := fmts[i].Kind; k != grid.Integer && k != grid.Decimal {
				continue
			}
			var currency, dates int
			symbol := ""
			divides := false
			for _, g := range c.Formulas {
				f := s.Get(g.Range.R1, c.Col).Formula
				if formula.Divides(f) {
					divides = true
					break
				}
				for _, rc := range formula.SameRowColumns(f, g.Range.R1) {
					j, in := index[rc]
					if !in || rc == c.Col {
						continue
					}
					switch fmts[j].Kind {
					case grid.Currency:
						currency++
						if symbol == "" {
							symbol = fmts[j].Symbol
						}
					case grid.Date, grid.DateTime:
						dates++
					}
				}
			}
			if divides {
				continue
			}
			first := s.Get(c.Formulas[0].Range.R1, c.Col).Formula
			switch {
			case currency > 0 && dates == 0 && callsOnly(first, keepsAmount):
				fmts[i] = grid.Format{Kind: grid.Currency, Symbol: symbol}
			case dates == 1 && currency == 0 && !formula.Multiplies(first) && callsOnly(first, keepsDate):
				fmts[i] = grid.Format{Kind: grid.Date}
				if c.Min != math.Trunc(c.Min) || c.Max != math.Trunc(c.Max) {
					fmts[i].Kind = grid.DateTime
				}
			default:
				continue
			}
			changed[i], again = true, true
		}
		if !again {
			break
		}
	}
	return changed
}

// typed returns the format kind a profiled column holds, General for text
// and mixed columns.
func typed(c Column) grid.Format {
	switch c.Type {
	case "integer", "decimal", "currency", "percent", "date", "datetime", "time":
		return grid.Format{Kind: c.Kind, Symbol: c.Symbol}
	}
	return grid.Format{}
}

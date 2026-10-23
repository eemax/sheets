package cli

import (
	"fmt"
	"math"
	"strconv"
	"strings"
	"text/tabwriter"

	"github.com/eemax/sheets/internal/analyze"
	"github.com/eemax/sheets/internal/grid"
	"github.com/spf13/cobra"
)

type peekResult struct {
	Sheet     string       `json:"sheet"`
	Tables    []peekTable  `json:"tables"`
	Other     []regionInfo `json:"other_regions,omitempty"`
	More      int          `json:"more_tables,omitempty"`
	MoreOther int          `json:"more_other_regions,omitempty"`
}

type peekTable struct {
	regionInfo
	TotalRow int          `json:"total_row_number,omitempty"`
	Profile  []peekColumn `json:"profile"`
	Sample   []readRow    `json:"sample"`
}

type peekColumn struct {
	Col      string   `json:"col"`
	Header   string   `json:"header,omitempty"`
	Type     string   `json:"type"`
	Mix      string   `json:"mix,omitempty"`
	Filled   int      `json:"filled"`
	Blank    int      `json:"blank,omitempty"`
	Errors   int      `json:"errors,omitempty"`
	Pending  int      `json:"unevaluated,omitempty"`
	Distinct int      `json:"distinct"`
	Min      any      `json:"min,omitempty"`
	Max      any      `json:"max,omitempty"`
	Sum      *float64 `json:"sum,omitempty"`
	Values   []string `json:"values,omitempty"` // enum values or samples
	Enum     bool     `json:"enum,omitempty"`
	Formulas []string `json:"formulas,omitempty"` // "D{r}-E{r} @ F2:F99"
	Summary  string   `json:"-"`
}

const maxPeekTables = 5

func (a *app) peekCmd() *cobra.Command {
	var sample int
	cmd := &cobra.Command{
		Use:   "peek <file> [sheet|range]",
		Short: "Profile tables: column types, ranges, distinct values, formula patterns",
		Long: `Summarise the tables on a sheet (or a given range) in a size that does not
grow with the data: one line per column with its type, how many cells are
filled, the value range or the distinct values, and formula patterns
written with {r} for the current row (D{r}-E{r}). A few sample rows follow.`,
		Args: cobra.RangeArgs(1, 2),
		RunE: func(cmd *cobra.Command, args []string) error {
			wb, err := open(args[0])
			if err != nil {
				return err
			}
			ref := ""
			if len(args) > 1 {
				ref = args[1]
			}
			s, rg, explicit, err := target(wb, ref)
			if err != nil {
				return err
			}
			res := peekResult{Sheet: s.Name, Tables: []peekTable{}}
			var regions []analyze.Region
			if explicit {
				regions = []analyze.Region{analyze.Classify(wb, s, rg)}
			} else {
				regions = analyze.Regions(wb, s)
			}
			for _, reg := range regions {
				if !reg.IsTable() && !(explicit || len(regions) == 1) {
					if len(res.Other) < maxRegionsListed {
						res.Other = append(res.Other, describeRegion(wb, s, reg))
					} else {
						res.MoreOther++
					}
					continue
				}
				if len(res.Tables) == maxPeekTables {
					res.More++
					continue
				}
				res.Tables = append(res.Tables, profileTable(wb, s, reg, sample))
			}
			if a.json {
				return a.printJSON(res)
			}
			a.printPeek(args[0], wb, s, res, sample)
			return nil
		},
	}
	cmd.Flags().IntVar(&sample, "sample", 3, "number of sample rows")
	return cmd
}

func profileTable(wb *grid.Workbook, s *grid.Sheet, reg analyze.Region, sample int) peekTable {
	pt := peekTable{regionInfo: describeRegion(wb, s, reg), TotalRow: reg.TotalRow}
	if pt.Kind == "cell" {
		pt.Kind = "block"
	}
	if !reg.IsTable() {
		pt.Rows = reg.Body.Rows()
	}
	for _, c := range analyze.Profile(wb, s, reg) {
		pt.Profile = append(pt.Profile, peekCol(wb, c))
	}
	last := min(reg.Body.R2, reg.Body.R1+sample-1)
	for r := reg.Body.R1; r <= last && sample > 0; r++ {
		row := readRow{Row: r}
		for c := reg.Range.C1; c <= reg.Range.C2; c++ {
			row.Values = append(row.Values, jsonValue(wb, s, r, c, false))
		}
		pt.Sample = append(pt.Sample, row)
	}
	return pt
}

func peekCol(wb *grid.Workbook, c analyze.Column) peekColumn {
	pc := peekColumn{
		Col: grid.ColName(c.Col), Header: c.Header, Type: c.Type, Mix: c.Mix,
		Filled: c.Filled, Blank: c.Blank, Errors: c.Errors, Pending: c.Pending,
		Distinct: c.Distinct, Values: c.Values, Enum: c.Enum,
	}
	if c.Type == "currency" && c.Symbol != "" {
		pc.Type += " " + c.Symbol
	}
	for _, g := range c.Formulas {
		pc.Formulas = append(pc.Formulas, "="+g.Pattern+" @ "+g.Range.String())
	}
	if c.FormulaGroupsMore > 0 {
		pc.Formulas = append(pc.Formulas, fmt.Sprintf("+%d more patterns", c.FormulaGroupsMore))
	}

	var parts []string
	if len(c.Formulas) == 1 && c.FormulaGroupsMore == 0 && c.Formulas[0].Range.Rows() == c.Filled {
		parts = append(parts, "="+c.Formulas[0].Pattern)
	} else if len(c.Formulas) > 0 {
		var fs []string
		for _, g := range c.Formulas[:min(len(c.Formulas), 3)] {
			fs = append(fs, "="+g.Pattern+" ("+g.Range.String()+")")
		}
		if n := len(c.Formulas) - 3 + c.FormulaGroupsMore; n > 0 {
			fs = append(fs, fmt.Sprintf("+%d more", n))
		}
		parts = append(parts, strings.Join(fs, ", "))
	}

	kind := c.Kind
	show := func(v float64) string {
		if kind.IsTemporal() {
			return grid.FormatTemporal(v, kind, wb.Date1904)
		}
		if kind == grid.Percent {
			return roundStat(v*100) + "%"
		}
		return roundStat(v)
	}
	numeric := kind.IsNumeric() || kind.IsTemporal()
	switch {
	case c.Enum:
		parts = append(parts, enumList(c.Values))
	case numeric:
		pc.Min, pc.Max = jsonNum(c.Min, kind, wb), jsonNum(c.Max, kind, wb)
		if c.Min == c.Max {
			parts = append(parts, show(c.Min))
		} else {
			parts = append(parts, show(c.Min)+" … "+show(c.Max))
		}
		if (kind == grid.Integer || kind == grid.Decimal || kind == grid.Currency) && !analyze.IDLike(c.Header) && !(c.Min >= 1900 && c.Max <= 2100 && kind == grid.Integer) {
			sum := c.Sum
			pc.Sum = &sum
			parts = append(parts, "sum "+roundStat(sum))
		}
	case len(c.Values) > 0:
		var q []string
		for _, v := range c.Values {
			q = append(q, fmt.Sprintf("%q", grid.Truncate(v, 30)))
		}
		s := strings.Join(q, ", ")
		if c.Distinct > len(c.Values) {
			s += fmt.Sprintf(", … %s distinct", grid.FormatNumber(float64(c.Distinct)))
		}
		parts = append(parts, s)
	}
	if c.Errors > 0 {
		parts = append(parts, plural(c.Errors, "error"))
	}
	if c.Pending > 0 {
		parts = append(parts, fmt.Sprintf("%d not evaluated", c.Pending))
	}
	pc.Summary = strings.Join(parts, " · ")
	if c.Mix != "" {
		pc.Type += " (" + c.Mix + ")"
	}
	return pc
}

// roundStat shortens a summary number: four decimals at most, or six
// significant digits for small values.
func roundStat(v float64) string {
	if math.Abs(v) >= 1 {
		return grid.FormatNumber(math.Round(v*1e4) / 1e4)
	}
	r, _ := strconv.ParseFloat(strconv.FormatFloat(v, 'g', 6, 64), 64)
	return grid.FormatNumber(r)
}

func jsonNum(v float64, k grid.FormatKind, wb *grid.Workbook) any {
	if k.IsTemporal() {
		return grid.FormatTemporal(v, k, wb.Date1904)
	}
	return v
}

func enumList(values []string) string {
	q := make([]string, len(values))
	for i, v := range values {
		q[i] = grid.Truncate(v, 30)
	}
	return strings.Join(q, ", ")
}

func (a *app) printPeek(file string, wb *grid.Workbook, s *grid.Sheet, res peekResult, sample int) {
	if len(res.Tables) == 0 && len(res.Other) == 0 {
		a.printf("%s is empty\n", grid.QuoteSheet(s.Name))
		return
	}
	for i, t := range res.Tables {
		if i > 0 {
			a.printf("\n")
		}
		head := []string{grid.QuoteSheet(s.Name) + "!" + t.Range, t.Kind}
		if t.Title != "" {
			head[1] += fmt.Sprintf(" %q", t.Title)
		}
		if t.Header > 0 {
			head = append(head, fmt.Sprintf("header row %d", t.Header))
		}
		rows := plural(t.Rows, "row")
		if t.TotalRow > 0 {
			rows += fmt.Sprintf(" + total row %d", t.TotalRow)
		}
		head = append(head, rows)
		a.printf("%s\n", strings.Join(head, " · "))

		tw := tabwriter.NewWriter(a.stdout, 0, 0, 2, ' ', 0)
		fmt.Fprintln(tw, "col\theader\ttype\tfilled\tvalues")
		for _, c := range t.Profile {
			filled := grid.FormatNumber(float64(c.Filled))
			if c.Blank > 0 {
				filled += "/" + grid.FormatNumber(float64(c.Filled+c.Blank))
			}
			fmt.Fprintf(tw, "%s\t%s\t%s\t%s\t%s\n", c.Col, grid.Escape(grid.Truncate(c.Header, 30)), c.Type, filled, grid.Escape(c.Summary))
		}
		tw.Flush()

		if len(t.Sample) > 0 {
			rg, _ := grid.ParseRange(t.Range)
			first := t.Sample[0].Row
			if t.Header > 0 {
				first = t.Header
			}
			sr := grid.Range{R1: first, C1: rg.C1, R2: t.Sample[len(t.Sample)-1].Row, C2: rg.C2}
			a.printf("sample:\n")
			a.printGrid(wb, s, sr, false, 40)
		}
	}
	if res.More > 0 {
		a.printf("\n… %s\n", plural(res.More, "more table"))
	}
	if len(res.Other) > 0 {
		a.printf("\nother regions:\n")
		for _, r := range res.Other {
			a.printf("  %s\n", regionLine(r))
		}
		if res.MoreOther > 0 {
			a.printf("  … %s\n", plural(res.MoreOther, "more region"))
		}
	}
	if len(res.Tables) > 0 {
		t := res.Tables[0]
		rg, _ := grid.ParseRange(t.Range)
		rg.R2 = min(rg.R2, rg.R1+19)
		a.hint("sheets read %s %s", shellQuote(file), cliRef(s.Name, rg))
	}
}

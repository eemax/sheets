package cli

import (
	"encoding/csv"
	"fmt"
	"io"
	"strings"
	"text/tabwriter"

	"github.com/eemax/sheets/internal/grid"
	"github.com/spf13/cobra"
)

type readResult struct {
	Sheet   string    `json:"sheet"`
	Range   string    `json:"range"`
	Columns []string  `json:"columns"`
	Rows    []readRow `json:"rows"`
	Next    string    `json:"next,omitempty"`
}

type readRow struct {
	Row    int   `json:"row"`
	Values []any `json:"values"`
}

type readOpts struct {
	limit    int
	maxCols  int
	width    int
	full     bool
	formulas bool
	format   string
}

func (a *app) readCmd() *cobra.Command {
	var o readOpts
	cmd := &cobra.Command{
		Use:   "read <file> [range]",
		Short: "Print cells with row numbers and column letters",
		Long: `Print a range of cells. The range may be Sheet!A1:D20, Sheet, A1:D20 (first
sheet), whole columns (Sheet!A:C) or whole rows (Sheet!5:20). Output is
capped (--limit rows, --max-cols columns) and ends with the command for the
next page.

Values are shown plainly: ISO dates, percentages with %, numbers without
thousands separators. --formulas shows formulas instead of their values.`,
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
			s, rg, _, err := target(wb, ref)
			if err != nil {
				return err
			}
			return a.read(args[0], wb, s, rg, o)
		},
	}
	f := cmd.Flags()
	f.IntVarP(&o.limit, "limit", "n", 50, "maximum rows to print")
	f.IntVar(&o.maxCols, "max-cols", 30, "maximum columns to print")
	f.IntVar(&o.width, "width", 80, "truncate longer values (text output)")
	f.BoolVar(&o.full, "full", false, "never truncate values")
	f.BoolVarP(&o.formulas, "formulas", "f", false, "show formulas instead of values")
	f.StringVar(&o.format, "format", "grid", "grid, tsv, csv or json")
	return cmd
}

func (a *app) read(file string, wb *grid.Workbook, s *grid.Sheet, rg grid.Range, o readOpts) error {
	if a.json {
		o.format = "json"
	}
	if rg.Empty() {
		if o.format == "json" {
			return a.printJSON(readResult{Sheet: s.Name, Rows: []readRow{}})
		}
		a.printf("%s is empty\n", grid.QuoteSheet(s.Name))
		return nil
	}
	shown := rg
	if o.limit > 0 && shown.Rows() > o.limit {
		shown.R2 = shown.R1 + o.limit - 1
	}
	if o.maxCols > 0 && shown.Cols() > o.maxCols {
		shown.C2 = shown.C1 + o.maxCols - 1
	}
	next := ""
	switch {
	case shown.R2 < rg.R2:
		nr := grid.Range{R1: shown.R2 + 1, C1: shown.C1, R2: min(rg.R2, shown.R2+max(o.limit, 1)), C2: shown.C2}
		next = fmt.Sprintf("sheets read %s %s", shellQuote(file), cliRef(s.Name, nr))
	case shown.C2 < rg.C2:
		nr := grid.Range{R1: shown.R1, C1: shown.C2 + 1, R2: shown.R2, C2: min(rg.C2, shown.C2+max(o.maxCols, 1))}
		next = fmt.Sprintf("sheets read %s %s", shellQuote(file), cliRef(s.Name, nr))
	}
	width := o.width
	if o.full || o.format != "grid" {
		width = 0
	}

	switch o.format {
	case "json":
		res := readResult{Sheet: s.Name, Range: shown.String(), Next: next}
		for c := shown.C1; c <= shown.C2; c++ {
			res.Columns = append(res.Columns, grid.ColName(c))
		}
		for r := shown.R1; r <= shown.R2; r++ {
			row := readRow{Row: r, Values: make([]any, shown.Cols())}
			for c := shown.C1; c <= shown.C2; c++ {
				row.Values[c-shown.C1] = jsonValue(wb, s, r, c, o.formulas)
			}
			res.Rows = append(res.Rows, row)
		}
		return a.printJSON(res)
	case "tsv", "csv":
		var cw *csv.Writer
		if o.format == "csv" {
			cw = csv.NewWriter(a.stdout)
		}
		for r := shown.R1; r <= shown.R2; r++ {
			fields := make([]string, shown.Cols())
			for c := shown.C1; c <= shown.C2; c++ {
				fields[c-shown.C1] = cellText(wb, s, r, c, o.formulas, 0, cw == nil)
			}
			if cw != nil {
				cw.Write(fields)
			} else {
				a.printf("%s\n", strings.Join(fields, "\t"))
			}
		}
		if cw != nil {
			cw.Flush()
		}
		if next != "" {
			fmt.Fprintf(a.stderr, "truncated; next: %s\n", next)
		}
		return nil
	case "grid":
	default:
		return fmt.Errorf("unknown --format %q (grid, tsv, csv or json)", o.format)
	}

	head := grid.FormatRef(s.Name, shown)
	if shown != rg {
		head += fmt.Sprintf(" (of %s)", rg.String())
	}
	a.printf("%s\n", head)
	a.printGrid(wb, s, shown, o.formulas, width)
	if next != "" {
		if more := rg.R2 - shown.R2; more > 0 {
			a.printf("… %s\n", plural(more, "more row"))
		} else {
			a.printf("… %s\n", plural(rg.C2-shown.C2, "more column"))
		}
		a.hint("%s", next)
	}
	return nil
}

// printGrid prints rows prefixed with their number under a line of column
// letters. Tabs separate cells; on a terminal the columns are aligned.
func (a *app) printGrid(wb *grid.Workbook, s *grid.Sheet, rg grid.Range, formulas bool, width int) {
	var w io.Writer = a.stdout
	var tw *tabwriter.Writer
	if a.tty {
		tw = tabwriter.NewWriter(a.stdout, 0, 0, 2, ' ', 0)
		w = tw
	}
	var b strings.Builder
	for c := rg.C1; c <= rg.C2; c++ {
		b.WriteString("\t" + grid.ColName(c))
	}
	fmt.Fprintln(w, b.String())
	for r := rg.R1; r <= rg.R2; r++ {
		b.Reset()
		fmt.Fprintf(&b, "%d", r)
		fields := make([]string, 0, rg.Cols())
		for c := rg.C1; c <= rg.C2; c++ {
			fields = append(fields, cellText(wb, s, r, c, formulas, width, true))
		}
		for len(fields) > 0 && fields[len(fields)-1] == "" && tw == nil {
			fields = fields[:len(fields)-1]
		}
		for _, f := range fields {
			b.WriteString("\t" + f)
		}
		fmt.Fprintln(w, b.String())
	}
	if tw != nil {
		tw.Flush()
	}
}

// cellText renders one cell for text output.
func cellText(wb *grid.Workbook, s *grid.Sheet, row, col int, formulas bool, width int, escape bool) string {
	c := s.At(row, col)
	if c == nil || c.IsEmpty() {
		return ""
	}
	var t string
	if formulas && c.Formula != "" {
		t = "=" + c.Formula
	} else {
		c = wb.Value(s, row, col)
		if c.Pending {
			t = "?"
		} else {
			t = grid.Display(c, wb.FormatOf(c), wb.Date1904)
		}
	}
	if escape {
		t = grid.Escape(t)
	}
	return grid.Truncate(t, width)
}

// jsonValue renders one cell as a typed JSON value.
func jsonValue(wb *grid.Workbook, s *grid.Sheet, row, col int, formulas bool) any {
	c := s.At(row, col)
	if c == nil || c.IsEmpty() {
		return nil
	}
	if formulas && c.Formula != "" {
		return "=" + c.Formula
	}
	c = wb.Value(s, row, col)
	switch c.Kind {
	case grid.Number:
		f := wb.FormatOf(c)
		if f.Kind.IsTemporal() {
			return grid.FormatTemporal(c.Num, f.Kind, wb.Date1904)
		}
		return c.Num
	case grid.Bool:
		return c.Bool
	case grid.Text, grid.Error:
		return c.Text
	}
	return nil
}

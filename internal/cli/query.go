package cli

import (
	"encoding/csv"
	"fmt"
	"io"
	"regexp"
	"strings"
	"text/tabwriter"

	"github.com/eemax/sheets/internal/book"
	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/input"
	"github.com/eemax/sheets/internal/query"
	"github.com/spf13/cobra"
)

type queryResult struct {
	*query.Result
	Into *writeResult `json:"into,omitempty"`
}

func (a *app) queryCmd() *cobra.Command {
	var (
		limit  int
		format string
		into   string
		extra  []string
		schema bool
	)
	cmd := &cobra.Command{
		Use:   "query <file> [sql]",
		Short: "Run SQL over the workbook's tables",
		Long: `Run a SQL query (SQLite dialect) over the tables in a workbook. Every table
sheets detects (a block with a header row) is a SQL table named after its
sheet ("Sales"; a second table on that sheet is "Sales_2") or after its
Excel table name. Columns are the header labels; quote names with spaces:
"Order ID". Without SQL, or with --schema, the tables are listed.

  sheets query book.xlsx 'SELECT Region, SUM(Revenue) AS Revenue FROM Sales GROUP BY Region ORDER BY 2 DESC'
  sheets query book.xlsx 'SELECT _row, * FROM Sales WHERE Units > 100'

Values: numbers are numbers (percentages as fractions: 12.5% is 0.125),
dates are ISO text ('2026-07-01', so date(), strftime() and comparisons
work), errors and blanks are NULL. Total rows are left out. Each table has
a _row column holding the sheet row of the record.

--into Sheet writes the result there as a new table (replacing what the
sheet held); --into Sheet!C5 writes it at that cell. Result columns named
like a source column keep its number format.`,
		Args: cobra.RangeArgs(1, 2),
		RunE: func(cmd *cobra.Command, args []string) error {
			wb, err := open(args[0])
			if err != nil {
				return err
			}
			var extras []query.Extra
			for _, e := range extra {
				x, err := parseExtra(wb, e)
				if err != nil {
					return err
				}
				extras = append(extras, x)
			}
			tables, err := query.Tables(wb, extras)
			if err != nil {
				return err
			}
			if len(args) == 1 || schema {
				return a.printSchema(tables)
			}
			n := limit
			if into != "" {
				n = 0 // write every row
			}
			res, err := query.Run(wb, tables, args[1], n)
			if err != nil {
				return err
			}
			if into != "" {
				w, err := a.queryInto(args[0], into, res, tables)
				if err != nil {
					return err
				}
				if a.json {
					return a.printJSON(queryResult{Into: w})
				}
				a.printDryRun()
				a.printWrite(w)
				return nil
			}
			if a.json {
				format = "json"
			}
			return a.printRows(res, format)
		},
	}
	f := cmd.Flags()
	f.IntVarP(&limit, "limit", "n", 100, "maximum rows to print")
	f.StringVar(&format, "format", "grid", "grid, tsv, csv or json")
	f.StringVar(&into, "into", "", "write the result to Sheet (replacing it) or Sheet!A1")
	f.StringArrayVar(&extra, "table", nil, "define a table: name=Sheet!A1:D20 (first row is the header)")
	f.BoolVar(&schema, "schema", false, "list tables and columns")
	return cmd
}

func parseExtra(wb *grid.Workbook, spec string) (query.Extra, error) {
	name, ref, ok := strings.Cut(spec, "=")
	if !ok || strings.TrimSpace(name) == "" {
		return query.Extra{}, fmt.Errorf("--table %q: expected name=Sheet!A1:D20", spec)
	}
	s, rg, explicit, err := target(wb, ref)
	if err != nil {
		return query.Extra{}, fmt.Errorf("--table %s: %w", name, err)
	}
	if !explicit {
		rg = s.Used()
	}
	return query.Extra{Name: strings.TrimSpace(name), Sheet: s.Name, Range: rg}, nil
}

func (a *app) printSchema(tables []*query.Table) error {
	if a.json {
		if tables == nil {
			tables = []*query.Table{}
		}
		return a.printJSON(tables)
	}
	if len(tables) == 0 {
		a.printf("no tables (blocks with a header row); define one with --table name=Sheet!A1:D20\n")
		return nil
	}
	for _, t := range tables {
		line := fmt.Sprintf("%s  %s · %s", query.Quote(t.Name), grid.QuoteSheet(t.Sheet)+"!"+t.Range, plural(t.Rows, "row"))
		if t.Total {
			line += " (total row left out)"
		}
		a.printf("%s\n", line)
		cols := []string{query.RowColumn + " INTEGER"}
		for _, c := range t.Columns {
			col := query.Quote(c.Name)
			if c.Type != "" {
				col += " " + c.Type
			}
			switch c.Kind {
			case "currency", "percent", "date", "datetime", "time", "bool", "mixed", "empty":
				col += " (" + c.Kind + ")"
			}
			cols = append(cols, col)
		}
		a.printf("  %s\n", strings.Join(cols, ", "))
	}
	return nil
}

func (a *app) printRows(res *query.Result, format string) error {
	switch format {
	case "json":
		return a.printJSON(res)
	case "csv":
		w := csv.NewWriter(a.stdout)
		w.Write(res.Columns)
		for _, r := range res.Rows {
			w.Write(sqlStrings(r))
		}
		w.Flush()
		a.moreRows(res, a.stderr)
		return w.Error()
	case "tsv", "grid":
	default:
		return fmt.Errorf("unknown --format %q (grid, tsv, csv or json)", format)
	}
	var w io.Writer = a.stdout
	var tw *tabwriter.Writer
	if a.tty && format == "grid" {
		tw = tabwriter.NewWriter(a.stdout, 0, 0, 2, ' ', 0)
		w = tw
	}
	fmt.Fprintln(w, strings.Join(res.Columns, "\t"))
	for _, r := range res.Rows {
		fields := sqlStrings(r)
		for i, f := range fields {
			fields[i] = grid.Escape(f)
			if format == "grid" {
				fields[i] = grid.Truncate(fields[i], 80)
			}
		}
		fmt.Fprintln(w, strings.Join(fields, "\t"))
	}
	if tw != nil {
		tw.Flush()
	}
	if format == "tsv" {
		a.moreRows(res, a.stderr)
	} else {
		a.moreRows(res, a.stdout)
	}
	return nil
}

func (a *app) moreRows(res *query.Result, w io.Writer) {
	if res.More > 0 {
		fmt.Fprintf(w, "… %s; narrow the query or raise --limit\n", plural(res.More, "more row"))
	} else if len(res.Rows) == 0 && w == a.stdout {
		fmt.Fprintln(w, "(no rows)")
	}
}

func sqlStrings(r []any) []string {
	out := make([]string, len(r))
	for i, v := range r {
		switch x := v.(type) {
		case nil:
		case float64:
			out[i] = grid.FormatNumber(x)
		case int64:
			out[i] = fmt.Sprint(x)
		case bool:
			out[i] = map[bool]string{true: "1", false: "0"}[x]
		default:
			out[i] = fmt.Sprint(x)
		}
	}
	return out
}

var isoDate = regexp.MustCompile(`^\d{4}-\d{2}-\d{2}([ T]\d{2}:\d{2}(:\d{2})?)?$`)

// queryInto writes a result into the workbook it came from.
func (a *app) queryInto(file, into string, res *query.Result, tables []*query.Table) (*writeResult, error) {
	// Number formats of source columns, by name, for result columns that
	// keep a source column's name (SUM(Revenue) AS Revenue).
	formats := map[string]grid.Format{}
	for _, t := range tables {
		for _, c := range t.Columns {
			if k, ok := grid.ParseFormatKind(c.Kind); ok {
				formats[strings.ToLower(c.Name)] = grid.Format{Kind: k, Symbol: c.Symbol}
			}
		}
	}
	rows := [][]input.Value{make([]input.Value, len(res.Columns))}
	for i, c := range res.Columns {
		rows[0][i] = input.Value{Kind: grid.Text, Text: c}
	}
	for _, r := range res.Rows {
		row := make([]input.Value, len(r))
		for i, v := range r {
			src := formats[strings.ToLower(res.Columns[i])]
			switch x := v.(type) {
			case int64:
				row[i] = input.Value{Kind: grid.Number, Num: float64(x)}
			case float64:
				row[i] = input.Value{Kind: grid.Number, Num: x}
			case string:
				if isoDate.MatchString(x) {
					row[i] = input.Parse(x, false)
				} else {
					row[i] = input.Value{Kind: grid.Text, Text: x}
				}
			}
			if row[i].Kind == grid.Number && row[i].Fmt.Kind == grid.General && (src.Kind == grid.Currency || src.Kind == grid.Percent) {
				row[i].Fmt = grid.Format{Kind: src.Kind, Symbol: src.Symbol, Decimals: 2}
				if src.Kind == grid.Percent {
					row[i].Fmt.Decimals = 1
				}
			}
		}
		rows = append(rows, row)
	}
	ed, err := book.Edit(file, false)
	if err != nil {
		return nil, err
	}
	defer ed.Close()
	r, err := grid.ParseRef(into, func(n string) bool { return ed.Workbook().Sheet(n) != nil })
	if err != nil {
		return nil, err
	}
	if r.Sheet == "" {
		return nil, fmt.Errorf("--into needs a sheet: --into Summary (replaces that sheet) or --into Summary!A1")
	}
	clear := r.Range == nil // a whole sheet: replace its contents
	w, err := writeRows(ed, file, into, rows, clear && ed.Workbook().Sheet(r.Sheet) != nil, false)
	if err != nil {
		return nil, err
	}
	if err := ed.Save(a.dryRun); err != nil {
		return nil, err
	}
	w.saved(ed)
	return w, nil
}

package cli

import (
	"fmt"
	"math"
	"strconv"
	"strings"

	"github.com/eemax/sheets/internal/analyze"
	"github.com/eemax/sheets/internal/book"
	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/query"
	"github.com/eemax/sheets/internal/xlsx"
	"github.com/spf13/cobra"
)

type chartResult struct {
	File     string        `json:"file"`
	Sheet    string        `json:"sheet"`
	Anchor   string        `json:"anchor"`
	Type     string        `json:"type"`
	Title    string        `json:"title"`
	X        string        `json:"x"`
	Y        []string      `json:"y"`
	Points   int           `json:"points"`
	DryRun   bool          `json:"dry_run,omitempty"`
	Warnings []string      `json:"warnings,omitempty"`
	Charts   []sheetCharts `json:"charts,omitempty"`
}

type sheetCharts struct {
	Sheet  string       `json:"sheet"`
	Charts []grid.Chart `json:"charts"`
}

func (a *app) chartCmd() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "chart add|list|delete <file> ...",
		Short: "Add, list or delete charts",
		Long: `Add charts in the house style, list the charts in a workbook, or delete one.

  sheets chart add book.xlsx Sales                         # first label column by the number columns
  sheets chart add book.xlsx Sales --x Region --y Revenue --type bar
  sheets chart add book.xlsx Sales --x Date --y Units --type line --at Summary!B2
  sheets chart list book.xlsx
  sheets chart delete book.xlsx Sales!H2

The source is a table: a sheet name, an Excel table name, Sheet_2 for a
second table on a sheet (as in sheets query), or a range whose first row
is the header. --x names the category column (default: the first column of
labels or dates), --y the value columns (default: every number column).
Total rows are left out.

The type defaults to a line for dates, a bar chart for many or long
labels, and columns otherwise. Series take the house palette in order (it
stays distinguishable with colour-vision deficiency); a single series gets
no legend; gridlines are hairlines; columns and bars stay thin. Charts that
would mislead are refused: more than 8 series, pies over 6 slices, value
columns in different units on one axis.`,
	}
	var (
		x, typ, title, at, size string
		y                       []string
		stacked                 bool
	)
	add := &cobra.Command{
		Use:   "add <file> <table>",
		Short: "Add a chart of a table",
		Args:  cobra.ExactArgs(2),
		RunE: func(cmd *cobra.Command, args []string) error {
			ed, err := book.Edit(args[0], false)
			if err != nil {
				return err
			}
			defer ed.Close()
			spec, res, err := chartSpec(ed.Workbook(), args[1], x, y, typ, stacked, title, at, size)
			if err != nil {
				return err
			}
			if err := ed.AddChart(spec); err != nil {
				return err
			}
			if err := ed.Save(a.dryRun); err != nil {
				return err
			}
			res.File, res.DryRun = args[0], a.dryRun
			res.Warnings = append(repeatedCategories(ed.Workbook(), spec), ed.Warnings()...)
			if a.json {
				return a.printJSON(res)
			}
			a.printDryRun()
			a.printf("added %s chart %q at %s: %s by %s, %s\n", res.Type, res.Title, grid.QuoteSheet(res.Sheet)+"!"+res.Anchor, strings.Join(res.Y, ", "), res.X, plural(res.Points, "point"))
			a.printWarnings(res.Warnings)
			a.hint("sheets render %s %s", shellQuote(args[0]), shellQuote(res.Sheet))
			return nil
		},
	}
	f := add.Flags()
	f.StringVar(&x, "x", "", "category column (header or letter)")
	f.StringSliceVar(&y, "y", nil, "value columns (headers or letters; repeat or separate with commas)")
	f.StringVarP(&typ, "type", "t", "", "column, bar, line, area, pie or scatter (default: chosen from the data)")
	f.BoolVar(&stacked, "stacked", false, "stack the series (column, bar, area)")
	f.StringVar(&title, "title", "", "chart title (default: the columns plotted)")
	f.StringVar(&at, "at", "", "top-left cell, Sheet!H2 or H2 (default: right of the table)")
	f.StringVar(&size, "size", "640x360", "width x height in pixels")

	list := &cobra.Command{
		Use:   "list <file> [sheet]",
		Short: "List charts with their type, title and data",
		Args:  cobra.RangeArgs(1, 2),
		RunE: func(cmd *cobra.Command, args []string) error {
			wb, err := open(args[0])
			if err != nil {
				return err
			}
			var out []sheetCharts
			for _, s := range wb.Sheets {
				if len(args) > 1 && !strings.EqualFold(s.Name, args[1]) {
					continue
				}
				if len(s.Charts) > 0 {
					sc := sheetCharts{Sheet: s.Name}
					for _, c := range s.Charts {
						sc.Charts = append(sc.Charts, c.Shown())
					}
					out = append(out, sc)
				}
			}
			if len(args) > 1 && wb.Sheet(args[1]) == nil {
				return fmt.Errorf("sheet %q not found (sheets: %s)", args[1], strings.Join(wb.SheetNames(), ", "))
			}
			if a.json {
				if out == nil {
					out = []sheetCharts{}
				}
				return a.printJSON(out)
			}
			if len(out) == 0 {
				a.printf("no charts\n")
				return nil
			}
			for _, sc := range out {
				for _, c := range sc.Charts {
					a.printf("%s  %s\n", grid.QuoteSheet(sc.Sheet)+"!"+c.Anchor, chartLine(c))
				}
			}
			return nil
		},
	}
	del := &cobra.Command{
		Use:   "delete <file> <Sheet!cell>",
		Short: "Delete the chart anchored at a cell",
		Args:  cobra.ExactArgs(2),
		RunE: func(cmd *cobra.Command, args []string) error {
			ed, err := book.Edit(args[0], false)
			if err != nil {
				return err
			}
			defer ed.Close()
			r, err := grid.ParseRef(args[1], func(n string) bool { return ed.Workbook().Sheet(n) != nil })
			if err != nil {
				return err
			}
			if r.Sheet == "" || r.Range == nil {
				return fmt.Errorf("give the chart's sheet and anchor cell, like Sales!H2 (sheets chart list shows them)")
			}
			cell := r.Range.String()
			if err := ed.DeleteChart(r.Sheet, cell); err != nil {
				return err
			}
			if err := ed.Save(a.dryRun); err != nil {
				return err
			}
			a.printDryRun()
			a.printf("deleted the chart at %s\n", grid.QuoteSheet(r.Sheet)+"!"+cell)
			a.printWarnings(ed.Warnings())
			return nil
		},
	}
	cmd.AddCommand(add, list, del)
	return cmd
}

// chartLine describes a chart in one line.
func chartLine(c grid.Chart) string {
	line := c.Type
	if line == "" {
		line = "chart"
	}
	if c.Title != "" {
		line += fmt.Sprintf(" %q", c.Title)
	}
	var series []string
	for _, s := range c.Series {
		series = append(series, s.Values)
	}
	if len(series) > 0 {
		line += " · " + plural(len(series), "series") + ": " + strings.Join(series, ", ")
	}
	return line
}

// chartSpec works out a chart of a table from the flags.
func chartSpec(wb *grid.Workbook, source, x string, y []string, typ string, stacked bool, title, at, size string) (book.ChartSpec, *chartResult, error) {
	var spec book.ChartSpec
	t, err := chartTable(wb, source)
	if err != nil {
		return spec, nil, err
	}
	if t.Rows < 1 {
		return spec, nil, fmt.Errorf("table %s has no data rows", query.Quote(t.Name))
	}
	find := func(name string) (int, bool) {
		for i, c := range t.Columns {
			if strings.EqualFold(c.Name, strings.TrimSpace(name)) {
				return i, true
			}
		}
		if col, ok := grid.ColIndex(strings.TrimSpace(name)); ok {
			for i, c := range t.Columns {
				if c.Index() == col {
					return i, true
				}
			}
		}
		return 0, false
	}
	columns := func() string {
		var names []string
		for _, c := range t.Columns {
			names = append(names, c.Name)
		}
		return strings.Join(names, ", ")
	}
	numeric := func(c query.Column) bool {
		switch c.Kind {
		case "integer", "decimal", "currency", "percent":
			return true
		}
		return false
	}

	// Categories.
	xi := -1
	if x != "" {
		i, ok := find(x)
		if !ok {
			return spec, nil, fmt.Errorf("--x %s: no such column in %s (columns: %s)", x, query.Quote(t.Name), columns())
		}
		xi = i
	} else {
		for i, c := range t.Columns {
			if !numeric(c) || typ == "scatter" {
				xi = i
				break
			}
		}
		if xi < 0 {
			xi = 0
		}
	}
	xc := t.Columns[xi]

	// Values.
	var ys []int
	if len(y) > 0 {
		for _, name := range y {
			i, ok := find(name)
			if !ok {
				return spec, nil, fmt.Errorf("--y %s: no such column in %s (columns: %s)", name, query.Quote(t.Name), columns())
			}
			if !numeric(t.Columns[i]) {
				return spec, nil, fmt.Errorf("--y %s holds %s, not numbers", t.Columns[i].Name, t.Columns[i].Kind)
			}
			ys = append(ys, i)
		}
	} else {
		for i, c := range t.Columns {
			if i == xi || !numeric(c) || analyze.IDLike(c.Name) {
				continue
			}
			ys = append(ys, i)
		}
		if len(ys) == 0 {
			return spec, nil, fmt.Errorf("%s has no number columns to plot (columns: %s)", query.Quote(t.Name), columns())
		}
		if len(ys) > xlsx.MaxSeries {
			return spec, nil, fmt.Errorf("%s has %d number columns; pick at most %d with --y", query.Quote(t.Name), len(ys), xlsx.MaxSeries)
		}
	}
	// One axis, one unit: different units on one axis mislead.
	unit := func(c query.Column) string {
		switch c.Kind {
		case "currency":
			return "currency " + c.Symbol
		case "percent":
			return "percent"
		}
		return "number"
	}
	if len(y) == 0 {
		groups := map[string][]string{}
		var order []string
		for _, i := range ys {
			u := unit(t.Columns[i])
			if groups[u] == nil {
				order = append(order, u)
			}
			groups[u] = append(groups[u], t.Columns[i].Name)
		}
		if len(order) > 1 {
			var parts []string
			example := groups[order[0]][0]
			for _, u := range order {
				parts = append(parts, u+": "+strings.Join(groups[u], ", "))
				if strings.HasPrefix(u, "currency") {
					example = groups[u][len(groups[u])-1] // money is usually the story
				}
			}
			return spec, nil, fmt.Errorf("%s has number columns in different units (%s); pick the ones to plot with --y, e.g. --y %s", query.Quote(t.Name), strings.Join(parts, "; "), shellQuote(example))
		}
	}
	for _, i := range ys[1:] {
		if unit(t.Columns[i]) != unit(t.Columns[ys[0]]) {
			return spec, nil, fmt.Errorf("%s (%s) and %s (%s) have different units; one axis cannot show both honestly, so chart them separately", t.Columns[ys[0]].Name, unit(t.Columns[ys[0]]), t.Columns[i].Name, unit(t.Columns[i]))
		}
	}

	body := t.Body()
	points := body.Rows()
	if typ == "" {
		switch {
		case xc.Kind == "date" || xc.Kind == "datetime":
			typ = "line"
		case points > 12 || longLabels(wb, t, xc):
			typ = "bar"
		default:
			typ = "column"
		}
	}
	if typ == "scatter" && !numeric(xc) {
		return spec, nil, fmt.Errorf("a scatter plot needs numbers on the x axis; %s holds %s", xc.Name, xc.Kind)
	}

	w, h, err := parseSize(size)
	if err != nil {
		return spec, nil, err
	}
	s := wb.Sheet(t.Sheet)
	placeSheet := s.Name
	anchor := ""
	if at != "" {
		r, err := grid.ParseRef(at, func(n string) bool { return wb.Sheet(n) != nil })
		if err != nil || r.Range == nil {
			return spec, nil, fmt.Errorf("--at %q: give a cell, like H2 or Summary!B2", at)
		}
		if r.Sheet != "" {
			ps := wb.Sheet(r.Sheet)
			if ps == nil {
				return spec, nil, fmt.Errorf("--at %s: sheet %q not found", at, r.Sheet)
			}
			placeSheet = ps.Name
		}
		anchor = r.Range.String()
	} else {
		tr, _ := grid.ParseRange(t.Range)
		row, col := tr.R1, tr.C2+2
		taken := map[string]bool{}
		for _, c := range s.Charts {
			taken[c.Anchor] = true
		}
		for taken[grid.CellName(row, col)] {
			row += 20 // below the chart already there
		}
		anchor = grid.CellName(row, col)
	}

	var names []string
	for _, i := range ys {
		names = append(names, t.Columns[i].Name)
	}
	if title == "" {
		title = strings.Join(names, " and ")
		if len(names) > 2 {
			title = strings.Join(names[:len(names)-1], ", ") + " and " + names[len(names)-1]
		}
		if typ != "pie" {
			title += " by " + xc.Name
		}
	}

	column := func(c query.Column) grid.Range {
		return grid.Range{R1: body.R1, C1: c.Index(), R2: body.R2, C2: c.Index()}
	}
	spec = book.ChartSpec{
		Sheet: placeSheet, Anchor: anchor, Type: typ, Stacked: stacked, Title: title,
		Width: w, Height: h, Points: points,
		Categories:  xlsx.ChartRef(s.Name, column(xc)),
		ValueFormat: valueFormat(t.Columns[ys[0]], wb, s, t, ys),
	}
	switch xc.Kind {
	case "date":
		spec.CategoryFormat = "yyyy-mm-dd"
	case "datetime":
		spec.CategoryFormat = "yyyy-mm-dd hh:mm"
	}
	if typ == "scatter" {
		spec.CategoryFormat = valueFormat(xc, wb, s, t, []int{xi})
	}
	for _, i := range ys {
		c := t.Columns[i]
		header := grid.Range{R1: t.HeaderRow(), C1: c.Index(), R2: t.HeaderRow(), C2: c.Index()}
		spec.Series = append(spec.Series, book.SeriesSpec{Name: xlsx.ChartRef(s.Name, header), Values: xlsx.ChartRef(s.Name, column(c))})
	}
	kind := typ
	if stacked {
		kind = "stacked " + typ
	}
	return spec, &chartResult{Sheet: placeSheet, Anchor: anchor, Type: kind, Title: title, X: xc.Name, Y: names, Points: points}, nil
}

// chartTable finds the table a chart plots: by table name (as sheets query
// names them) or by range, whose first row is the header.
func chartTable(wb *grid.Workbook, source string) (*query.Table, error) {
	tables, err := query.Tables(wb, nil)
	if err != nil {
		return nil, err
	}
	for _, t := range tables {
		if strings.EqualFold(t.Name, source) {
			return t, nil
		}
	}
	r, err := grid.ParseRef(source, func(n string) bool { return wb.Sheet(n) != nil })
	if err == nil && r.Range != nil {
		sheet := r.Sheet
		if sheet == "" {
			sheet = wb.FirstVisible().Name
		}
		ts, err := query.Tables(wb, []query.Extra{{Name: "chart source", Sheet: sheet, Range: *r.Range}})
		if err != nil {
			return nil, err
		}
		return ts[0], nil
	}
	if s := wb.Sheet(source); s != nil {
		return nil, fmt.Errorf("sheet %s has no table (a block with a header row); give a range such as %s!A1:D20", grid.QuoteSheet(s.Name), grid.QuoteSheet(s.Name))
	}
	var names []string
	for _, t := range tables {
		names = append(names, query.Quote(t.Name))
	}
	return nil, fmt.Errorf("no table %q (tables: %s)", source, strings.Join(names, ", "))
}

func longLabels(wb *grid.Workbook, t *query.Table, c query.Column) bool {
	s := wb.Sheet(t.Sheet)
	body := t.Body()
	total := 0
	for r := body.R1; r <= body.R2; r++ {
		total += len([]rune(s.Get(r, c.Index()).Text))
	}
	return total/max(body.Rows(), 1) > 12
}

// valueFormat picks the axis number format of the plotted columns: their
// unit, no decimals, thousands as K and millions as M.
func valueFormat(c query.Column, wb *grid.Workbook, s *grid.Sheet, t *query.Table, cols []int) string {
	if c.Kind == "percent" {
		return "0%"
	}
	peak := 0.0
	body := t.Body()
	for _, i := range cols {
		col := t.Columns[i].Index()
		for r := body.R1; r <= body.R2; r++ {
			if v := wb.Value(s, r, col); v != nil && v.Kind == grid.Number {
				peak = math.Max(peak, math.Abs(v.Num))
			}
		}
	}
	prefix := ""
	if c.Kind == "currency" {
		sym := c.Symbol
		if sym == "" {
			sym = "$"
		}
		prefix = `"` + sym + `"`
	}
	switch {
	case peak >= 10_000_000:
		return prefix + `#,##0,,"M"`
	case peak >= 10_000:
		return prefix + `#,##0,"K"`
	case peak < 10 && c.Kind == "decimal":
		return prefix + "#,##0.0"
	}
	return prefix + "#,##0"
}

func parseSize(s string) (uint, uint, error) {
	w, h, ok := strings.Cut(strings.ToLower(s), "x")
	wi, err1 := strconv.Atoi(strings.TrimSpace(w))
	hi, err2 := strconv.Atoi(strings.TrimSpace(h))
	if !ok || err1 != nil || err2 != nil || wi < 200 || hi < 150 || wi > 4000 || hi > 4000 {
		return 0, 0, fmt.Errorf("--size %q: give width x height in pixels, like 640x360 (at least 200x150)", s)
	}
	return uint(wi), uint(hi), nil
}

// repeatedCategories warns when a chart's categories repeat: charts plot
// rows, so "revenue by region" over orders shows each order, not each
// region's total.
func repeatedCategories(wb *grid.Workbook, spec book.ChartSpec) []string {
	if spec.Type == "line" || spec.Type == "scatter" || spec.Type == "area" {
		return nil
	}
	r, err := grid.ParseRef(spec.Categories, nil)
	if err != nil || r.Range == nil {
		return nil
	}
	s := wb.Sheet(r.Sheet)
	if s == nil {
		return nil
	}
	seen := map[string]bool{}
	repeats := 0
	for row := r.Range.R1; row <= r.Range.R2; row++ {
		v := wb.Value(s, row, r.Range.C1)
		if v == nil {
			continue
		}
		k := grid.Display(v, wb.FormatOf(v), wb.Date1904)
		if seen[k] {
			repeats++
		}
		seen[k] = true
	}
	if repeats == 0 {
		return nil
	}
	return []string{fmt.Sprintf("the categories repeat (%s); the chart shows each row, not a total per category. For totals, aggregate first: sheets query FILE 'SELECT ..., SUM(...) AS ... FROM ... GROUP BY 1' --into Summary, then chart Summary", plural(repeats, "repeated value"))}
}

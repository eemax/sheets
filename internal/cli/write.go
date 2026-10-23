package cli

import (
	"errors"
	"fmt"
	"io"
	"math"
	"os"
	"sort"
	"strings"

	"github.com/eemax/sheets/internal/analyze"
	"github.com/eemax/sheets/internal/book"
	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/input"
	"github.com/spf13/cobra"
)

type writeResult struct {
	File         string            `json:"file"`
	Sheet        string            `json:"sheet"`
	Range        string            `json:"range"`
	Rows         int               `json:"rows"`
	Cols         int               `json:"cols"`
	CreatedFile  bool              `json:"created_file,omitempty"`
	CreatedSheet bool              `json:"created_sheet,omitempty"`
	Types        map[string]string `json:"types,omitempty"`
	Formulas     *formulaReport    `json:"formulas,omitempty"`
	Warnings     []string          `json:"warnings,omitempty"`
}

func (a *app) writeCmd() *cobra.Command {
	var (
		from, inFile      string
		clear, appendRows bool
		raw               bool
	)
	cmd := &cobra.Command{
		Use:   "write <file> [sheet!cell]",
		Short: "Write TSV, CSV or JSON from stdin into a sheet",
		Long: `Write a block of values read from stdin (or -i file) starting at the given
cell (default A1 of the first sheet). The file and sheet are created when
missing.

Input is TSV, CSV, or JSON (an array of arrays, or an array of objects whose
keys become the header row); the format is detected unless --from is given.

Values are typed: 2026-07-01 is a date, 12.5% a percentage, $1,200 currency,
=SUM(A1:A9) a formula, TRUE a boolean. Leading zeros and long digit strings
stay text. Prefix a value with ' or pass --raw to keep it as text.

In formulas, {r} stands for the row the value lands on, so the patterns
'sheets peek' prints can be written back: =C{r}*D{r} on row 7 becomes
=C7*D7. {r-1} and {r+1} work too.

--append writes below the last used row. When the first input row matches
the sheet's header names, columns are matched by name and that row is
skipped.`,
		Args: cobra.RangeArgs(1, 2),
		RunE: func(cmd *cobra.Command, args []string) error {
			var r io.Reader = a.stdin
			if inFile != "" {
				f, err := os.Open(inFile)
				if err != nil {
					return err
				}
				defer f.Close()
				r = f
			} else if f, ok := a.stdin.(*os.File); ok && isTerminal(f) {
				return errors.New("no input: pipe data on stdin or pass -i <file>")
			}
			rows, err := input.Table(r, from, raw)
			if err != nil {
				return err
			}
			if len(rows) == 0 {
				return errors.New("no input rows")
			}
			if clear && appendRows {
				return errors.New("--clear and --append contradict each other; pick one")
			}
			ref := ""
			if len(args) > 1 {
				ref = args[1]
			}
			ed, err := book.Edit(args[0], true)
			if err != nil {
				return err
			}
			defer ed.Close()
			res, err := writeRows(ed, args[0], ref, rows, clear, appendRows)
			if err != nil {
				return err
			}
			res.Warnings = ed.Warnings()
			if err := ed.Save(a.dryRun); err != nil {
				return err
			}
			if a.json {
				return a.printJSON(res)
			}
			a.printDryRun()
			a.printWrite(res)
			return nil
		},
	}
	f := cmd.Flags()
	f.StringVar(&from, "from", "auto", "input format: auto, tsv, csv or json")
	f.StringVarP(&inFile, "input", "i", "", "read input from this file instead of stdin")
	f.BoolVar(&clear, "clear", false, "empty the sheet before writing")
	f.BoolVar(&appendRows, "append", false, "write below the last used row")
	f.BoolVar(&raw, "raw", false, "write every value as text")
	return cmd
}

// sheetFor resolves the sheet and anchor cell of a write target.
func sheetFor(ed book.Editor, ref string) (name string, created bool, anchor *grid.Range, err error) {
	wb := ed.Workbook()
	r, err := grid.ParseRef(ref, func(n string) bool { return wb.Sheet(n) != nil })
	if err != nil {
		return "", false, nil, err
	}
	name = r.Sheet
	if name == "" {
		if s := wb.FirstVisible(); s != nil {
			name = s.Name
		} else {
			name = "Sheet1"
		}
	}
	name, created, err = ed.EnsureSheet(name)
	return name, created, r.Range, err
}

func writeRows(ed book.Editor, file, ref string, rows [][]input.Value, clear, appendRows bool) (*writeResult, error) {
	sheet, created, anchor, err := sheetFor(ed, ref)
	if err != nil {
		return nil, err
	}
	res := &writeResult{File: file, Sheet: sheet, CreatedFile: ed.Created(), CreatedSheet: created && !ed.Created()}
	wb := ed.Workbook()
	s := wb.Sheet(sheet)
	row0, col0 := 1, 1
	if anchor != nil {
		row0, col0 = anchor.R1, anchor.C1
	}
	width := 0
	for _, r := range rows {
		width = max(width, len(r))
	}
	cols := make([]int, width) // input column → sheet column
	for i := range cols {
		cols[i] = col0 + i
	}
	skipped := 0 // input rows consumed as a header
	if clear {
		if err := ed.Clear(sheet); err != nil {
			return nil, err
		}
	} else if appendRows {
		if used := s.Used(); !used.Empty() {
			row0 = used.R2 + 1
			if anchor == nil {
				col0 = used.C1
				for i := range cols {
					cols[i] = col0 + i
				}
			}
			if mapped, ok, err := matchHeader(wb, s, rows[0]); err != nil {
				return nil, err
			} else if ok {
				cols = mapped
				rows = rows[1:]
				skipped = 1
				if len(rows) == 0 {
					return nil, errors.New("input has a header row but no data rows")
				}
			}
		}
	}

	var formulas []cellRef
	types := map[int]map[string]int{}
	typedFrom := 0
	if len(rows) > 1 && allText(rows[0]) {
		typedFrom = 1 // a header row says nothing about types
	}
	minCol, maxCol := grid.MaxCols, 0
	for i, r := range rows {
		if len(r) > len(cols) {
			return nil, fmt.Errorf("input row %d has %d values but the header names %d columns", i+1+skipped, len(r), len(cols))
		}
		for j, v := range r {
			row, col := row0+i, cols[j]
			if v.Formula != "" {
				v.Formula = formula.FillRow(v.Formula, row)
			}
			if err := ed.Set(sheet, row, col, v); err != nil {
				return nil, err
			}
			minCol, maxCol = min(minCol, col), max(maxCol, col)
			if v.Formula != "" {
				formulas = append(formulas, cellRef{sheet, row, col, formula.Spills(v.Formula)})
			}
			if i >= typedFrom && !v.IsEmpty() {
				if types[col] == nil {
					types[col] = map[string]int{}
				}
				types[col][valueType(v)]++
			}
		}
	}
	res.Rows, res.Cols = len(rows), maxCol-minCol+1
	res.Range = grid.Range{R1: row0, C1: minCol, R2: row0 + len(rows) - 1, C2: maxCol}.String()
	res.Types = map[string]string{}
	for col, counts := range types {
		// Whole numbers among decimals are still a decimal column.
		if counts["decimal"] > 0 {
			counts["decimal"] += counts["integer"]
			delete(counts, "integer")
		}
		best, n := "", 0
		for t, c := range counts {
			if c > n || c == n && t < best {
				best, n = t, c
			}
		}
		if best != "text" {
			res.Types[grid.ColName(col)] = best
		}
	}

	res.Formulas = evaluateFormulas(ed, formulas, false)
	return res, nil
}

// matchHeader maps input columns to sheet columns by header name when the
// first input row consists of the sheet's header labels.
func matchHeader(wb *grid.Workbook, s *grid.Sheet, first []input.Value) ([]int, bool, error) {
	if !allText(first) {
		return nil, false, nil
	}
	var header *analyze.Region
	for _, reg := range analyze.Regions(wb, s) {
		if reg.HeaderRow > 0 {
			r := reg
			header = &r
		}
	}
	if header == nil {
		return nil, false, nil
	}
	byName := map[string]int{}
	var names []string
	for col := header.Range.C1; col <= header.Range.C2; col++ {
		h := strings.TrimSpace(analyze.Label(wb, s, header.HeaderRow, col))
		if h != "" {
			byName[strings.ToLower(h)] = col
			names = append(names, h)
		}
	}
	cols := make([]int, len(first))
	var unknown []string
	for i, v := range first {
		col, ok := byName[strings.ToLower(strings.TrimSpace(v.Text))]
		if !ok {
			unknown = append(unknown, v.Text)
		}
		cols[i] = col
	}
	if len(unknown) == len(first) {
		return nil, false, nil // plain data, not a header
	}
	if len(unknown) > 0 {
		return nil, false, fmt.Errorf("input header has columns the sheet lacks: %s (sheet headers: %s)", strings.Join(unknown, ", "), strings.Join(names, ", "))
	}
	return cols, true, nil
}

func allText(r []input.Value) bool {
	for _, v := range r {
		if v.Kind != grid.Text || v.Formula != "" {
			return false
		}
	}
	return len(r) > 0
}

func valueType(v input.Value) string {
	switch {
	case v.Formula != "":
		return "formula"
	case v.Kind == grid.Bool:
		return "bool"
	case v.Kind == grid.Text:
		return "text"
	case v.Fmt.Kind != grid.General:
		return v.Fmt.Kind.String()
	case v.Num != math.Trunc(v.Num):
		return "decimal"
	}
	return "integer"
}

func plainValue(c grid.Cell) any {
	switch c.Kind {
	case grid.Number:
		return c.Num
	case grid.Bool:
		return c.Bool
	case grid.Text, grid.Error:
		return c.Text
	}
	return nil
}

func showValue(v any) string {
	switch x := v.(type) {
	case float64:
		return grid.FormatNumber(x)
	case bool:
		if x {
			return "TRUE"
		}
		return "FALSE"
	case string:
		return fmt.Sprintf("%q", grid.Truncate(x, 40))
	}
	return "empty"
}

func (a *app) printWrite(res *writeResult) {
	var notes []string
	if res.CreatedFile {
		notes = append(notes, "new file")
	}
	if res.CreatedSheet {
		notes = append(notes, "new sheet")
	}
	line := fmt.Sprintf("wrote %s × %s to %s", plural(res.Rows, "row"), plural(res.Cols, "col"), grid.QuoteSheet(res.Sheet)+"!"+res.Range)
	if len(notes) > 0 {
		line += " (" + strings.Join(notes, ", ") + ")"
	}
	a.printf("%s\n", line)
	if len(res.Types) > 0 {
		var cols []string
		for c := range res.Types {
			cols = append(cols, c)
		}
		sort.Slice(cols, func(i, j int) bool {
			ci, _ := grid.ColIndex(cols[i])
			cj, _ := grid.ColIndex(cols[j])
			return ci < cj
		})
		var parts []string
		for _, c := range cols {
			parts = append(parts, c+" "+res.Types[c])
		}
		a.printf("types: %s\n", strings.Join(parts, ", "))
	}
	a.printFormulaReport(res.Formulas)
	a.printWarnings(res.Warnings)
	if res.Formulas == nil || len(res.Formulas.Errors) == 0 {
		a.hint("sheets fmt %s", shellQuote(res.File))
	}
}

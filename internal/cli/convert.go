package cli

import (
	"fmt"
	"path/filepath"
	"strings"

	"github.com/eemax/sheets/internal/book"
	"github.com/eemax/sheets/internal/delim"
	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/xlsx"
	"github.com/spf13/cobra"
)

type convertResult struct {
	From     string   `json:"from"`
	To       string   `json:"to"`
	Sheet    string   `json:"sheet,omitempty"`
	Rows     int      `json:"rows"`
	Cols     int      `json:"cols"`
	Warnings []string `json:"warnings,omitempty"`
	DryRun   bool     `json:"dry_run,omitempty"`
}

func (a *app) convertCmd() *cobra.Command {
	var sheet, delimiter string
	var bom bool
	cmd := &cobra.Command{
		Use:   "convert <from> <to>",
		Short: "Convert between .xlsx and .csv/.tsv",
		Long: `Convert a workbook between formats, by the files' extensions.

  sheets convert data.csv data.xlsx        # values typed: numbers, dates, %, currency
  sheets convert book.xlsx sales.csv --sheet Sales
  sheets convert export.csv export.tsv

From CSV, the sheet is named after the file, and values are typed as sheets
types input (1,200 is a number, 2026-03-01 a date, 00123 stays text); run
sheets fmt on the result for the house style. To CSV, one sheet is written
(--sheet, or the first), with values as sheets shows them: ISO dates,
percentages with %, formulas as their results. CSV is written as UTF-8,
with commas (tabs for .tsv) unless --delimiter; --bom adds the byte order
mark Excel needs to read accented letters in a UTF-8 CSV.

sheets reads, edits and writes CSV files directly too (read, peek, query,
lint, diff, write, set, sort, replace, ...), keeping their delimiter,
encoding and line breaks; formatting needs .xlsx.`,
		Args: cobra.ExactArgs(2),
		RunE: func(cmd *cobra.Command, args []string) error {
			from, to := args[0], args[1]
			if _, err := book.FormatOf(from); err != nil {
				return err
			}
			kind, err := book.FormatOf(to)
			if err != nil {
				return err
			}
			wb, err := open(from)
			if err != nil {
				return err
			}
			res := convertResult{From: from, To: to, DryRun: a.dryRun}
			var out []byte
			switch kind {
			case "xlsx":
				if wb.Format == "xlsx" {
					return fmt.Errorf("%s is already an xlsx workbook; copy it instead", from)
				}
				for _, s := range wb.Sheets {
					u := s.Used()
					res.Sheet, res.Rows, res.Cols = s.Name, u.R2, u.C2
					if u.Empty() {
						res.Rows, res.Cols = 0, 0
					}
				}
				if out, err = xlsx.FromWorkbook(wb); err != nil {
					return err
				}
			default:
				s := wb.FirstVisible()
				if sheet != "" {
					if s = wb.Sheet(sheet); s == nil {
						return fmt.Errorf("sheet %q not found (sheets: %s)", sheet, strings.Join(wb.SheetNames(), ", "))
					}
				}
				if s == nil {
					return fmt.Errorf("%s has no sheets", from)
				}
				res.Sheet = s.Name
				d := delim.DefaultDialect(to)
				d.BOM, d.Plain = bom, true
				if delimiter != "" {
					r := []rune(strings.ReplaceAll(delimiter, `\t`, "\t"))
					if len(r) != 1 || r[0] == '"' || r[0] == '\n' || r[0] == '\r' {
						return fmt.Errorf("--delimiter %q: give one character, like ';' or '\\t'", delimiter)
					}
					d.Comma = r[0]
				}
				u := s.Used()
				if !u.Empty() {
					res.Rows, res.Cols = u.R2, u.C2
				}
				var unknown []string
				value := func(r, c int) (grid.Cell, grid.Format, error) {
					cell := wb.Value(s, r, c)
					if cell == nil || cell.IsEmpty() {
						return grid.Cell{}, grid.Format{}, nil
					}
					if cell.Pending {
						unknown = append(unknown, grid.CellName(r, c))
						return grid.Cell{}, grid.Format{}, nil
					}
					return *cell, wb.FormatOf(cell), nil
				}
				if out, err = delim.Encode(res.Rows, res.Cols, value, d, nil, wb.Date1904); err != nil {
					return err
				}
				if len(unknown) > 0 {
					res.Warnings = append(res.Warnings, fmt.Sprintf("the sheets calculator cannot compute %s, left empty: %s (open the workbook in Excel and save it, so it stores their results)",
						plural(len(unknown), "formula"), listCells(unknown, 8)))
				}
				if others := len(wb.Sheets) - 1; others > 0 && sheet == "" {
					res.Warnings = append(res.Warnings, fmt.Sprintf("wrote sheet %s only; a CSV file holds one sheet (the others: --sheet NAME)", s.Name))
				}
			}
			if !a.dryRun {
				if err := xlsx.WriteFile(to, out); err != nil {
					return err
				}
			}
			if a.json {
				return a.printJSON(res)
			}
			a.printDryRun()
			a.printf("converted %s → %s (%s, %s × %s)\n", from, to, grid.QuoteSheet(res.Sheet), plural(res.Rows, "row"), plural(res.Cols, "col"))
			a.printWarnings(res.Warnings)
			if kind == "xlsx" {
				a.hint("sheets fmt %s", shellQuote(to))
			} else if filepath.Ext(to) != "" {
				a.hint("sheets peek %s", shellQuote(to))
			}
			return nil
		},
	}
	f := cmd.Flags()
	f.StringVar(&sheet, "sheet", "", "the sheet to write to CSV (default: the first)")
	f.StringVar(&delimiter, "delimiter", "", `field delimiter for CSV output: ',', ';', '\t' or '|'`)
	f.BoolVar(&bom, "bom", false, "start a UTF-8 CSV with a byte order mark, for Excel")
	return cmd
}

// listCells shows the first few cells of a list.
func listCells(cells []string, n int) string {
	if len(cells) <= n {
		return strings.Join(cells, ", ")
	}
	return strings.Join(cells[:n], ", ") + fmt.Sprintf(" and %d more", len(cells)-n)
}

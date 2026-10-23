package cli

import (
	"errors"
	"fmt"

	"github.com/eemax/sheets/internal/book"
	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/input"
	"github.com/spf13/cobra"
)

type setResult struct {
	File     string         `json:"file"`
	Cells    []setCell      `json:"cells"`
	Formulas *formulaReport `json:"formulas,omitempty"`
	Warnings []string       `json:"warnings,omitempty"`
}

type setCell struct {
	Ref   string `json:"ref"`
	Input string `json:"input"`
	Cells int    `json:"cells"`
}

func (a *app) setCmd() *cobra.Command {
	var raw bool
	cmd := &cobra.Command{
		Use:   "set <file> <cell> <value> [<cell> <value> ...]",
		Short: "Set cells to values or formulas",
		Long: `Set one or more cells. A cell is Sheet!B4 or B4 (first sheet). Values are
typed like 'sheets write' input; start a formula with =.

A range fills every cell. Formulas are filled the way Excel fills down or
right, shifting relative references:

  sheets set book.xlsx 'Data!E2:E100' '=C2*D2'    # E3 gets =C3*D3, ...
  sheets set book.xlsx 'Data!E2:E100' '=C{r}*D{r}' # the same; {r} is the row

The file and sheet are created when missing. An empty value clears a cell.`,
		Args: func(cmd *cobra.Command, args []string) error {
			if len(args) < 3 || len(args)%2 == 0 {
				return errors.New("expected <file> followed by <cell> <value> pairs")
			}
			return nil
		},
		RunE: func(cmd *cobra.Command, args []string) error {
			ed, err := book.Edit(args[0], true)
			if err != nil {
				return err
			}
			defer ed.Close()
			res := setResult{File: args[0]}
			var toEval []cellRef
			for i := 1; i < len(args); i += 2 {
				sheet, _, rg, err := sheetFor(ed, args[i])
				if err != nil {
					return err
				}
				if rg == nil {
					return fmt.Errorf("%q: give a cell such as B4 or Sheet!B4", args[i])
				}
				if rg.R2 == grid.MaxRows || rg.C2 == grid.MaxCols {
					return fmt.Errorf("%q: whole rows or columns cannot be set; give a bounded range", args[i])
				}
				v := input.Parse(args[i+1], raw)
				for r := rg.R1; r <= rg.R2; r++ {
					for c := rg.C1; c <= rg.C2; c++ {
						cv := v
						if v.Formula != "" && formula.HasRowPlaceholder(v.Formula) {
							cv.Formula = formula.FillRow(v.Formula, r)
							toEval = append(toEval, cellRef{sheet, r, c, formula.Spills(cv.Formula)})
						} else if v.Formula != "" {
							cv.Formula = formula.Shift(v.Formula, r-rg.R1, c-rg.C1)
							toEval = append(toEval, cellRef{sheet, r, c, formula.Spills(cv.Formula)})
						}
						if err := ed.Set(sheet, r, c, cv); err != nil {
							return err
						}
					}
				}
				res.Cells = append(res.Cells, setCell{Ref: grid.FormatRef(sheet, *rg), Input: args[i+1], Cells: rg.Rows() * rg.Cols()})
			}
			res.Formulas = evaluateFormulas(ed, toEval, true)
			if err := ed.Save(a.dryRun); err != nil {
				return err
			}
			res.Warnings = ed.Warnings()
			if a.json {
				return a.printJSON(res)
			}
			a.printDryRun()
			for _, c := range res.Cells {
				line := fmt.Sprintf("%s ← %s", c.Ref, c.Input)
				if c.Cells > 1 {
					line += fmt.Sprintf(" (%d cells)", c.Cells)
				}
				a.printf("%s\n", line)
			}
			a.printFormulaReport(res.Formulas)
			a.printWarnings(res.Warnings)
			return nil
		},
	}
	cmd.Flags().BoolVar(&raw, "raw", false, "store values as text, never as numbers or formulas")
	return cmd
}

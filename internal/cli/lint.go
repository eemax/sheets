package cli

import (
	"fmt"
	"strings"
	"text/tabwriter"

	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/lint"
	"github.com/spf13/cobra"
)

type lintResult struct {
	File     string         `json:"file"`
	Errors   int            `json:"errors"`
	Warnings int            `json:"warnings"`
	Notes    int            `json:"notes"`
	Findings []lint.Finding `json:"findings"`
}

const maxCellsListed = 8

func (a *app) lintCmd() *cobra.Command {
	var (
		sheets []string
		failOn string
		rules  bool
	)
	cmd := &cobra.Command{
		Use:   "lint <file>",
		Short: "Find errors, broken references, inconsistent formulas and other problems",
		Long: `Check a workbook for problems: error values, references to deleted cells,
unknown functions, cells that break their column's formula pattern, totals
that leave out rows, numbers stored as text, and more (--rules lists them).

Exits 1 when it finds errors, so it can gate a pipeline; --fail-on warning
is stricter, --fail-on never only reports.`,
		Args: func(cmd *cobra.Command, args []string) error {
			if rules {
				return nil
			}
			return cobra.ExactArgs(1)(cmd, args)
		},
		RunE: func(cmd *cobra.Command, args []string) error {
			if rules {
				tw := tabwriter.NewWriter(a.stdout, 0, 0, 2, ' ', 0)
				for _, r := range lint.Rules {
					fmt.Fprintf(tw, "%s\t%s\t%s\n", r.Name, r.Severity, r.What)
				}
				return tw.Flush()
			}
			threshold := lint.Severity(99)
			if failOn != "never" {
				s, ok := lint.ParseSeverity(failOn)
				if !ok {
					return fmt.Errorf("--fail-on %q: use error, warning, note or never", failOn)
				}
				threshold = s
			}
			wb, err := open(args[0])
			if err != nil {
				return err
			}
			for _, name := range sheets {
				if wb.Sheet(name) == nil {
					return fmt.Errorf("sheet %q not found (sheets: %s)", name, strings.Join(wb.SheetNames(), ", "))
				}
			}
			res := lintResult{File: args[0], Findings: lint.Check(wb, lint.Options{Sheets: sheets})}
			if res.Findings == nil {
				res.Findings = []lint.Finding{}
			}
			fail := false
			for _, f := range res.Findings {
				switch f.Severity {
				case lint.Error:
					res.Errors++
				case lint.Warning:
					res.Warnings++
				default:
					res.Notes++
				}
				fail = fail || f.Severity >= threshold
			}
			if a.json {
				if err := a.printJSON(res); err != nil {
					return err
				}
			} else {
				a.printLint(res)
			}
			if fail {
				return errSilent
			}
			return nil
		},
	}
	f := cmd.Flags()
	f.StringArrayVar(&sheets, "sheet", nil, "check only this sheet (repeatable)")
	f.StringVar(&failOn, "fail-on", "error", "exit 1 on findings this severe: error, warning, note or never")
	f.BoolVar(&rules, "rules", false, "list the checks")
	return cmd
}

func (a *app) printLint(res lintResult) {
	if len(res.Findings) == 0 {
		a.printf("%s: no problems found\n", res.File)
		return
	}
	a.printf("%s: %s, %s, %s\n", res.File, plural(res.Errors, "error"), plural(res.Warnings, "warning"), plural(res.Notes, "note"))
	tw := tabwriter.NewWriter(a.stdout, 0, 0, 2, ' ', 0)
	for _, f := range res.Findings {
		fmt.Fprintf(tw, "%s\t%s\t%s\n", f.Severity, cellList(f.Sheet, f.Cells), f.Message)
	}
	tw.Flush()
}

// cellList renders Sheet!A1, B2, C3 (+5 more).
func cellList(sheet string, cells []string) string {
	if len(cells) == 0 {
		return "(workbook)"
	}
	shown := cells[:min(len(cells), maxCellsListed)]
	s := grid.QuoteSheet(sheet) + "!" + strings.Join(shown, ", ")
	if len(cells) > maxCellsListed {
		s += fmt.Sprintf(" (+%d more)", len(cells)-maxCellsListed)
	}
	return s
}

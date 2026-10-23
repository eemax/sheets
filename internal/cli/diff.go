package cli

import (
	"fmt"
	"path/filepath"
	"strings"
	"time"

	"github.com/eemax/sheets/internal/book"
	"github.com/eemax/sheets/internal/diff"
	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/undo"
	"github.com/spf13/cobra"
)

type diffResult struct {
	Old string `json:"old"`
	New string `json:"new"`
	diff.Result
}

func (a *app) diffCmd() *cobra.Command {
	var (
		since    int
		limit    int
		exitCode bool
	)
	cmd := &cobra.Command{
		Use:   "diff <old> <new> | diff <file>",
		Short: "Show what changed between two workbooks, or since an earlier change",
		Long: `Compare two workbooks: sheets added, removed or renamed, rows and columns
added or removed, and cells whose formula, value or number format changed.
Rows and columns are aligned first, so an inserted row shows as one added
row, not as every row below it changing.

With one file, compare it with its state before the last change sheets made
(--since 3: before the last three), from the undo history. That is the way
to check what a command did.

  sheets diff book.xlsx
  sheets diff before.xlsx after.xlsx`,
		Args: cobra.RangeArgs(1, 2),
		RunE: func(cmd *cobra.Command, args []string) error {
			var oldWB, newWB *grid.Workbook
			res := diffResult{}
			var head string
			if len(args) == 2 {
				var err error
				if oldWB, err = open(args[0]); err != nil {
					return err
				}
				if newWB, err = open(args[1]); err != nil {
					return err
				}
				res.Old, res.New = args[0], args[1]
				head = fmt.Sprintf("%s → %s", args[0], args[1])
			} else {
				data, e, current, err := undo.Before(args[0], since)
				if err == undo.ErrNothing {
					return fmt.Errorf("no recorded changes to %s; give two files to compare", args[0])
				}
				if err != nil {
					return err
				}
				if newWB, err = open(args[0]); err != nil {
					return err
				}
				if e.Existed {
					if oldWB, err = book.OpenBytes(filepath.Base(args[0]), data); err != nil {
						return err
					}
				} else {
					oldWB = &grid.Workbook{Format: newWB.Format, Formats: []grid.Format{{}}}
				}
				when := e.Time.Local().Format(time.DateTime)
				res.Old = fmt.Sprintf("before: %s (%s)", e.Command, when)
				res.New = args[0]
				head = fmt.Sprintf("%s: before %s (%s) → now", args[0], e.Command, when)
				if !e.Existed {
					head += "; the file did not exist before"
				}
				if !current {
					head += "\nnote: the file was changed by something else after sheets last wrote it"
				}
			}
			res.Result = diff.Workbooks(oldWB, newWB, diff.Options{MaxCells: limit})
			if a.json {
				if err := a.printJSON(res); err != nil {
					return err
				}
			} else {
				a.printDiff(head, res.Result)
			}
			if exitCode && !res.Same {
				return errSilent
			}
			return nil
		},
	}
	f := cmd.Flags()
	f.IntVar(&since, "since", 1, "with one file: compare with before the last N changes")
	f.IntVarP(&limit, "limit", "n", 30, "cell changes listed per sheet (0: all)")
	f.BoolVar(&exitCode, "exit-code", false, "exit 1 when the workbooks differ")
	return cmd
}

func (a *app) printDiff(head string, r diff.Result) {
	a.printf("%s\n", head)
	if r.Same {
		a.printf("no differences\n")
		return
	}
	for _, s := range r.Sheets {
		name := grid.QuoteSheet(s.Name)
		switch s.Status {
		case "added", "removed":
			used := ""
			if s.Used != "" {
				used = " (" + s.Used + ")"
			}
			a.printf("%s: %s%s\n", name, s.Status, used)
			continue
		}
		a.printf("%s: %s\n", name, s.Summary())
		for _, l := range s.ColsRemoved {
			a.printf("  - %s\n", span("column", l, true))
		}
		for _, l := range s.ColsAdded {
			a.printf("  + %s\n", span("column", l, true))
		}
		for _, l := range s.RowsRemoved {
			a.printf("  - %s%s\n", span("row", l, false), content(l))
		}
		for _, l := range s.RowsAdded {
			a.printf("  + %s%s\n", span("row", l, false), content(l))
		}
		for _, c := range s.Cells {
			cell := c.Cell
			if c.OldCell != "" {
				cell += " (was " + c.OldCell + ")"
			}
			line := fmt.Sprintf("  %s  %s → %s", cell, orEmpty(c.Old), orEmpty(c.New))
			if c.Note != "" {
				line += "  (" + c.Note + ")"
			}
			a.printf("%s\n", line)
		}
		if s.MoreCells > 0 {
			a.printf("  … %s (raise --limit)\n", plural(s.MoreCells, "more changed cell"))
		}
		for _, f := range s.Formats {
			a.printf("  format %s  %s → %s\n", f.Range, f.Old, f.New)
		}
	}
	if len(r.Names) > 0 {
		var parts []string
		for _, n := range r.Names {
			switch {
			case n.Old == "":
				parts = append(parts, "+ "+n.Name+" = "+n.New)
			case n.New == "":
				parts = append(parts, "- "+n.Name)
			default:
				parts = append(parts, n.Name+": "+n.Old+" → "+n.New)
			}
		}
		a.printf("names: %s\n", strings.Join(parts, "; "))
	}
	if r.Order != nil {
		a.printf("sheet order: %s → %s\n", strings.Join(r.Order.Old, ", "), strings.Join(r.Order.New, ", "))
	}
}

func span(word string, l diff.Lines, cols bool) string {
	from, to := fmt.Sprint(l.From), fmt.Sprint(l.To)
	if cols {
		from, to = grid.ColName(l.From), grid.ColName(l.To)
	}
	if l.From == l.To {
		return word + " " + from
	}
	return word + "s " + from + "-" + to
}

func content(l diff.Lines) string {
	if l.Content == "" {
		return ""
	}
	prefix := ": "
	if l.To > l.From {
		prefix = ", first: "
	}
	return prefix + l.Content
}

func orEmpty(s string) string {
	if s == "" {
		return "(empty)"
	}
	return s
}

package cli

import (
	"errors"
	"fmt"
	"strconv"
	"strings"
	"time"

	"github.com/eemax/sheets/internal/book"
	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/undo"
	"github.com/spf13/cobra"
)

type structResult struct {
	File     string   `json:"file"`
	Done     string   `json:"done"`
	Sheets   []string `json:"sheets,omitempty"` // sheet order afterwards
	Removed  int      `json:"removed_values,omitempty"`
	Warnings []string `json:"warnings,omitempty"`
	DryRun   bool     `json:"dry_run,omitempty"`
}

// edit opens a file, runs fn, saves (unless dry run) and reports.
func (a *app) edit(file string, create bool, fn func(ed book.Editor, res *structResult) error) error {
	ed, err := book.Edit(file, create)
	if err != nil {
		return err
	}
	defer ed.Close()
	res := &structResult{File: file, DryRun: a.dryRun}
	if err := fn(ed, res); err != nil {
		return err
	}
	if err := ed.Save(a.dryRun); err != nil {
		return err
	}
	res.Warnings = append(res.Warnings, ed.Warnings()...)
	if a.json {
		return a.printJSON(res)
	}
	a.printDryRun()
	a.printf("%s\n", res.Done)
	if len(res.Sheets) > 0 {
		a.printf("sheets: %s\n", strings.Join(res.Sheets, ", "))
	}
	a.printWarnings(res.Warnings)
	return nil
}

func (a *app) printDryRun() {
	if a.dryRun {
		a.printf("dry run: nothing was saved\n")
	}
}

// lineSpec parses Sheet!5, Sheet!5:9, 5 (rows) or Sheet!C, C:E (columns).
func lineSpec(wb *grid.Workbook, spec string, cols bool) (string, int, int, error) {
	sheet, part := "", spec
	if i := strings.LastIndex(spec, "!"); i >= 0 {
		sheet, part = strings.Trim(spec[:i], "'"), spec[i+1:]
	}
	if sheet == "" {
		if s := wb.FirstVisible(); s != nil {
			sheet = s.Name
		}
	} else if s := wb.Sheet(sheet); s != nil {
		sheet = s.Name
	} else {
		return "", 0, 0, fmt.Errorf("sheet %q not found (sheets: %s)", sheet, strings.Join(wb.SheetNames(), ", "))
	}
	lo, hi, ok := strings.Cut(part, ":")
	if !ok {
		hi = lo
	}
	parse := func(s string) (int, bool) {
		s = strings.TrimPrefix(strings.TrimSpace(s), "$")
		if cols {
			return grid.ColIndex(s)
		}
		n, err := strconv.Atoi(s)
		return n, err == nil && n >= 1 && n <= grid.MaxRows
	}
	a, ok1 := parse(lo)
	b, ok2 := parse(hi)
	if !ok1 || !ok2 {
		if cols {
			return "", 0, 0, fmt.Errorf("%q: give columns like Sheet!C or Sheet!C:E", spec)
		}
		return "", 0, 0, fmt.Errorf("%q: give rows like Sheet!5 or Sheet!5:9", spec)
	}
	return sheet, min(a, b), max(a, b), nil
}

func (a *app) linesCmd(cols bool) *cobra.Command {
	unit, example, other := "rows", "Sales!5", "Sales!5:9"
	if cols {
		unit, example, other = "cols", "Sales!C", "Sales!C:E"
	}
	word := map[bool]string{false: "row", true: "column"}[cols]
	cmd := &cobra.Command{
		Use:   unit + " insert|delete <file> <" + word + "s>",
		Short: "Insert or delete " + word + "s, keeping formulas pointed at the right cells",
		Long: fmt.Sprintf(`Insert or delete %[1]ss. References move with the cells, on every sheet:
formulas, names, merged cells, conditional formats, validations, filters and
tables. Formulas that pointed into deleted %[1]ss become #REF!, and sheets
lists them.

  sheets %[2]s insert book.xlsx %[3]s -n 2   # two new %[1]ss before %[3]s
  sheets %[2]s delete book.xlsx %[4]s`, word, unit, example, other),
	}
	var count int
	insert := &cobra.Command{
		Use:   "insert <file> <" + word + ">",
		Short: "Insert empty " + word + "s before the given one",
		Args:  cobra.ExactArgs(2),
		RunE: func(cmd *cobra.Command, args []string) error {
			if count < 1 {
				return errors.New("-n must be at least 1")
			}
			return a.edit(args[0], false, func(ed book.Editor, res *structResult) error {
				sheet, at, _, err := lineSpec(ed.Workbook(), args[1], cols)
				if err != nil {
					return err
				}
				if cols {
					err = ed.InsertCols(sheet, at, count)
				} else {
					err = ed.InsertRows(sheet, at, count)
				}
				if err != nil {
					return err
				}
				res.Done = fmt.Sprintf("inserted %s before %s (new: %s)", plural(count, word), lineRef(sheet, at, at, cols), lineRef(sheet, at, at+count-1, cols))
				return nil
			})
		},
	}
	insert.Flags().IntVarP(&count, "count", "n", 1, "how many "+word+"s")
	del := &cobra.Command{
		Use:   "delete <file> <" + word + "s>",
		Short: "Delete " + word + "s",
		Args:  cobra.ExactArgs(2),
		RunE: func(cmd *cobra.Command, args []string) error {
			return a.edit(args[0], false, func(ed book.Editor, res *structResult) error {
				wb := ed.Workbook()
				sheet, from, to, err := lineSpec(wb, args[1], cols)
				if err != nil {
					return err
				}
				rg := grid.Range{R1: from, C1: 1, R2: to, C2: grid.MaxCols}
				if cols {
					rg = grid.Range{R1: 1, C1: from, R2: grid.MaxRows, C2: to}
				}
				wb.Sheet(sheet).Each(rg, func(int, int, *grid.Cell) { res.Removed++ })
				if cols {
					err = ed.DeleteCols(sheet, from, to)
				} else {
					err = ed.DeleteRows(sheet, from, to)
				}
				if err != nil {
					return err
				}
				res.Done = fmt.Sprintf("deleted %s (%s)", lineRef(sheet, from, to, cols), plural(res.Removed, "value"))
				if res.Removed > 0 && !a.dryRun && !undo.Disabled() {
					res.Done += fmt.Sprintf("; undo with: sheets undo %s", shellQuote(args[0]))
				}
				return nil
			})
		},
	}
	cmd.AddCommand(insert, del)
	return cmd
}

func lineRef(sheet string, a, b int, cols bool) string {
	s, e := strconv.Itoa(a), strconv.Itoa(b)
	if cols {
		s, e = grid.ColName(a), grid.ColName(b)
	}
	ref := grid.QuoteSheet(sheet) + "!" + s
	if a != b {
		ref += ":" + e
	}
	return ref
}

func (a *app) sheetCmd() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "sheet add|rename|delete|copy|move|hide|show <file> ...",
		Short: "Add, rename, delete, copy, move, hide or show sheets",
		Long: `Manage sheets. Renaming updates every reference to the sheet: formulas on
all sheets, names, conditional formats, validations, chart ranges and pivot
sources. Deleting refuses while formulas elsewhere still refer to the sheet,
unless --force (those references then become #REF!, as in Excel).`,
	}
	var before, after string
	var first, last, force bool
	report := func(ed book.Editor, res *structResult) {
		res.Sheets = a.sheetList(ed)
	}
	add := &cobra.Command{
		Use:   "add <file> <name>",
		Short: "Add an empty sheet (creates the file if needed)",
		Args:  cobra.ExactArgs(2),
		RunE: func(cmd *cobra.Command, args []string) error {
			return a.edit(args[0], true, func(ed book.Editor, res *structResult) error {
				if err := ed.AddSheet(args[1], before, after); err != nil {
					return err
				}
				res.Done = "added sheet " + grid.QuoteSheet(args[1])
				report(ed, res)
				return nil
			})
		},
	}
	add.Flags().StringVar(&before, "before", "", "place before this sheet")
	add.Flags().StringVar(&after, "after", "", "place after this sheet")
	rename := &cobra.Command{
		Use:   "rename <file> <sheet> <new-name>",
		Short: "Rename a sheet and every reference to it",
		Args:  cobra.ExactArgs(3),
		RunE: func(cmd *cobra.Command, args []string) error {
			return a.edit(args[0], false, func(ed book.Editor, res *structResult) error {
				refs := len(ed.References(args[1]))
				if err := ed.RenameSheet(args[1], args[2]); err != nil {
					return err
				}
				res.Done = fmt.Sprintf("renamed %s to %s", grid.QuoteSheet(args[1]), grid.QuoteSheet(args[2]))
				if refs > 0 {
					res.Done += fmt.Sprintf(" (%s updated)", plural(refs, "reference"))
				}
				report(ed, res)
				return nil
			})
		},
	}
	del := &cobra.Command{
		Use:   "delete <file> <sheet>",
		Short: "Delete a sheet",
		Args:  cobra.ExactArgs(2),
		RunE: func(cmd *cobra.Command, args []string) error {
			return a.edit(args[0], false, func(ed book.Editor, res *structResult) error {
				s := ed.Workbook().Sheet(args[1])
				if s == nil {
					return fmt.Errorf("sheet %q not found (sheets: %s)", args[1], strings.Join(ed.Workbook().SheetNames(), ", "))
				}
				name := s.Name
				if refs := ed.References(name); len(refs) > 0 && !force {
					list := strings.Join(refs[:min(len(refs), 10)], ", ")
					if len(refs) > 10 {
						list += fmt.Sprintf(", … %d more", len(refs)-10)
					}
					return fmt.Errorf("%s is still used in %s: %s; fix those first or pass --force (they become #REF!)", grid.QuoteSheet(name), plural(len(refs), "place"), list)
				}
				s.Each(s.Used(), func(int, int, *grid.Cell) { res.Removed++ })
				if err := ed.DeleteSheet(name); err != nil {
					return err
				}
				res.Done = fmt.Sprintf("deleted sheet %s (%s)", grid.QuoteSheet(name), plural(res.Removed, "value"))
				if !a.dryRun && !undo.Disabled() {
					res.Done += fmt.Sprintf("; undo with: sheets undo %s", shellQuote(args[0]))
				}
				report(ed, res)
				return nil
			})
		},
	}
	del.Flags().BoolVar(&force, "force", false, "delete even if other formulas refer to the sheet")
	cp := &cobra.Command{
		Use:   "copy <file> <sheet> <new-name>",
		Short: "Duplicate a sheet's cells and formatting",
		Args:  cobra.ExactArgs(3),
		RunE: func(cmd *cobra.Command, args []string) error {
			return a.edit(args[0], false, func(ed book.Editor, res *structResult) error {
				notes, err := ed.CopySheet(args[1], args[2])
				if err != nil {
					return err
				}
				res.Done = fmt.Sprintf("copied %s to %s", grid.QuoteSheet(args[1]), grid.QuoteSheet(args[2]))
				res.Warnings = append(res.Warnings, notes...)
				report(ed, res)
				return nil
			})
		},
	}
	move := &cobra.Command{
		Use:   "move <file> <sheet>",
		Short: "Change a sheet's position",
		Args:  cobra.ExactArgs(2),
		RunE: func(cmd *cobra.Command, args []string) error {
			n := 0
			for _, set := range []bool{before != "", after != "", first, last} {
				if set {
					n++
				}
			}
			if n != 1 {
				return errors.New("give exactly one of --before, --after, --first, --last")
			}
			return a.edit(args[0], false, func(ed book.Editor, res *structResult) error {
				if err := ed.MoveSheet(args[1], before, after, first, last); err != nil {
					return err
				}
				res.Done = "moved " + grid.QuoteSheet(args[1])
				report(ed, res)
				return nil
			})
		},
	}
	move.Flags().StringVar(&before, "before", "", "place before this sheet")
	move.Flags().StringVar(&after, "after", "", "place after this sheet")
	move.Flags().BoolVar(&first, "first", false, "make it the first sheet")
	move.Flags().BoolVar(&last, "last", false, "make it the last sheet")
	visibility := func(show bool) *cobra.Command {
		verb := map[bool]string{true: "show", false: "hide"}[show]
		return &cobra.Command{
			Use:   verb + " <file> <sheet>",
			Short: strings.ToUpper(verb[:1]) + verb[1:] + " a sheet",
			Args:  cobra.ExactArgs(2),
			RunE: func(cmd *cobra.Command, args []string) error {
				return a.edit(args[0], false, func(ed book.Editor, res *structResult) error {
					if err := ed.SetVisible(args[1], show); err != nil {
						return err
					}
					res.Done = map[bool]string{true: "showed ", false: "hid "}[show] + grid.QuoteSheet(args[1])
					report(ed, res)
					return nil
				})
			},
		}
	}
	cmd.AddCommand(add, rename, del, cp, move, visibility(false), visibility(true))
	return cmd
}

// sheetList is the sheet order as the file now has it, hidden ones marked.
func (a *app) sheetList(ed book.Editor) []string {
	return ed.SheetList()
}

func (a *app) undoCmd() *cobra.Command {
	var list, force bool
	cmd := &cobra.Command{
		Use:   "undo <file>",
		Short: "Restore a file to before the last change sheets made to it",
		Long: `Every write by sheets keeps the previous version of the file (the last 20
per file, in the user cache directory; SHEETS_UNDO=off disables this).
undo restores the last one; run it again to step further back. It refuses
when the file was changed by something else since, unless --force.`,
		Args: cobra.ExactArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			if list {
				entries, err := undo.List(args[0])
				if err != nil {
					return err
				}
				if a.json {
					if entries == nil {
						entries = []undo.Entry{}
					}
					return a.printJSON(entries)
				}
				if len(entries) == 0 {
					a.printf("no recorded changes to %s\n", args[0])
					return nil
				}
				for i := len(entries) - 1; i >= 0; i-- {
					e := entries[i]
					a.printf("%s  %s\n", e.Time.Local().Format(time.DateTime), e.Command)
				}
				return nil
			}
			if a.dryRun {
				entries, err := undo.List(args[0])
				if err != nil {
					return err
				}
				if len(entries) == 0 {
					return undo.ErrNothing
				}
				a.printDryRun()
				a.printf("would undo: %s\n", entries[len(entries)-1].Command)
				return nil
			}
			e, moved, err := undo.Undo(args[0], force)
			if err != nil {
				return err
			}
			if a.json {
				return a.printJSON(map[string]any{"undone": e, "moved_to": moved})
			}
			if moved != "" {
				a.printf("undid %s (%s): the file did not exist before, so it was moved to %s\n", e.Command, e.Time.Local().Format(time.DateTime), moved)
				return nil
			}
			a.printf("undid %s (%s)\n", e.Command, e.Time.Local().Format(time.DateTime))
			return nil
		},
	}
	cmd.Flags().BoolVar(&list, "list", false, "list recorded changes, newest first")
	cmd.Flags().BoolVar(&force, "force", false, "restore even if the file changed since")
	return cmd
}

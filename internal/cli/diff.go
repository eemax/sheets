package cli

import (
	"fmt"
	"os"
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
		Use:   "diff <old> <new> | diff <file> | diff <git's 7 arguments>",
		Short: "Show what changed between two workbooks, or since an earlier change",
		Long: `Compare two workbooks: sheets added, removed or renamed, rows and columns
added or removed, and cells whose formula, value or number format changed.
Rows and columns are aligned first, so an inserted row shows as one added
row, not as every row below it changing.

With one file, compare it with its state before the last change sheets made
(--since 3: before the last three), from the undo history. That is the way
to check what a command did.

It is git's diff driver for workbooks too: git diff then prints what
changed in a .xlsx file, not "Binary files differ" (git show and git log -p
with --ext-diff). Set it up once in a repository:

  git config diff.sheets.command 'sheets diff'
  echo '*.xlsx diff=sheets' >> .gitattributes

  sheets diff book.xlsx
  sheets diff before.xlsx after.xlsx`,
		Args: func(cmd *cobra.Command, args []string) error {
			if _, ok := gitDiffArgs(args); ok {
				return nil
			}
			return cobra.RangeArgs(1, 2)(cmd, args)
		},
		RunE: func(cmd *cobra.Command, args []string) error {
			var oldWB, newWB *grid.Workbook
			res := diffResult{}
			var head string
			if len(args) == 1 && os.Getenv("GIT_DIFF_PATH_TOTAL") != "" {
				// git gives an unmerged path alone.
				a.printf("* Unmerged path %s\n", args[0])
				return nil
			}
			if g, ok := gitDiffArgs(args); ok {
				var err error
				if oldWB, newWB, err = g.open(); err != nil {
					return err
				}
				res.Old, res.New, head = g.oldName(), g.newName(), g.head()
			} else if len(args) == 2 {
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
			if exitCode && !res.Same && !isGit(args) {
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

// gitDiff is how git runs an external diff driver: the path, then the old
// file, its blob's hash and mode, then the new file's (/dev/null for a
// file added or deleted, with . for its hash and mode); a file renamed or
// copied adds its new path and git's message of it.
type gitDiff struct {
	path, oldFile, oldHex, newFile, newHex, newPath string
}

// gitDiffArgs tells git's arguments from a person's.
func gitDiffArgs(args []string) (gitDiff, bool) {
	if len(args) != 7 && len(args) != 9 {
		return gitDiff{}, false
	}
	if !gitHex(args[2]) || !gitMode(args[3]) || !gitHex(args[5]) || !gitMode(args[6]) {
		return gitDiff{}, false
	}
	g := gitDiff{path: args[0], oldFile: args[1], oldHex: args[2], newFile: args[4], newHex: args[5]}
	if len(args) == 9 {
		g.newPath = args[7]
	}
	return g, true
}

func isGit(args []string) bool { _, ok := gitDiffArgs(args); return ok }

func gitHex(s string) bool {
	if s == "." {
		return true
	}
	if len(s) != 40 && len(s) != 64 {
		return false
	}
	return strings.Trim(s, "0123456789abcdef") == ""
}

func gitMode(s string) bool {
	return s == "." || len(s) == 6 && strings.Trim(s, "01234567") == ""
}

// gitNothing is git's name for the side of a file added or deleted.
const gitNothing = "/dev/null"

// open reads both sides; a side git gives as /dev/null is a workbook with
// nothing in it.
func (g gitDiff) open() (*grid.Workbook, *grid.Workbook, error) {
	read := func(file string) (*grid.Workbook, error) {
		if file == gitNothing {
			return nil, nil
		}
		return open(file)
	}
	oldWB, err := read(g.oldFile)
	if err != nil {
		return nil, nil, fmt.Errorf("%s (before): %w", g.path, err)
	}
	newWB, err := read(g.newFile)
	if err != nil {
		return nil, nil, fmt.Errorf("%s: %w", g.path, err)
	}
	if oldWB == nil && newWB == nil {
		return nil, nil, fmt.Errorf("%s: git gave no file on either side", g.path)
	}
	if oldWB == nil {
		oldWB = &grid.Workbook{Format: newWB.Format, Formats: []grid.Format{{}}}
	}
	if newWB == nil {
		newWB = &grid.Workbook{Format: oldWB.Format, Formats: []grid.Format{{}}}
	}
	return oldWB, newWB, nil
}

func (g gitDiff) oldName() string { return "a/" + g.path }

func (g gitDiff) newName() string {
	if g.newPath != "" {
		return "b/" + g.newPath
	}
	return "b/" + g.path
}

// head is the first line git shows of a file: its name, as git names the
// sides, and what happened to it.
func (g gitDiff) head() string {
	h := "sheets diff " + g.oldName() + " " + g.newName()
	switch {
	case g.oldFile == gitNothing:
		h += " (new file)"
	case g.newFile == gitNothing:
		h += " (deleted)"
	case g.newPath != "" && g.newPath != g.path:
		h += " (renamed)"
	}
	return h
}

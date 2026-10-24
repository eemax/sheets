package cli

import (
	"errors"
	"fmt"
	"regexp"
	"strings"

	"github.com/eemax/sheets/internal/analyze"
	"github.com/eemax/sheets/internal/book"
	"github.com/eemax/sheets/internal/grid"
	"github.com/spf13/cobra"
)

type rangeResult struct {
	File string `json:"file"`
	book.RangeResult
	Warnings []string `json:"warnings,omitempty"`
	DryRun   bool     `json:"dry_run,omitempty"`
}

// report saves an edit and prints what a range operation did.
func (a *app) report(ed book.Editor, file, done string, res book.RangeResult) error {
	if err := ed.Save(a.dryRun); err != nil {
		return err
	}
	out := rangeResult{File: file, RangeResult: res, Warnings: ed.Warnings(), DryRun: a.dryRun}
	if a.json {
		return a.printJSON(out)
	}
	a.printDryRun()
	a.printf("%s\n", done)
	for _, n := range res.Notes {
		a.printf("note: %s\n", n)
	}
	a.printWarnings(out.Warnings)
	return nil
}

// whole is a range covering a sheet.
var whole = grid.Range{R1: 1, C1: 1, R2: grid.MaxRows, C2: grid.MaxCols}

func (a *app) clearCmd() *cobra.Command {
	var format, notes, links, all bool
	cmd := &cobra.Command{
		Use:   "clear <file> <target>",
		Short: "Clear cells: contents, or formats, notes and links",
		Long: `Empty cells, as Excel's Clear does. By default the values and formulas go
and formatting stays (the Delete key). --format clears formatting instead
(cell styles, merges, the conditional formats lying within, and recorded
styling over the cells), --notes and --links those, and --all everything.
Data validation stays (sheets validate --clear removes it).

  sheets clear book.xlsx Sales!A2:D100
  sheets clear book.xlsx 'Sales[Notes]'
  sheets clear book.xlsx Sales!F1:H20 --all

` + targetHelp + ` A dynamic array formula is cleared with its result; part of
its result cannot be cleared alone. An Excel table keeps its header row.`,
		Args: cobra.ExactArgs(2),
		RunE: func(cmd *cobra.Command, args []string) error {
			ed, sheet, r, err := editTarget(args[0], args[1])
			if err != nil {
				return err
			}
			defer ed.Close()
			w := book.ClearWhat{Contents: !format && !notes && !links, Formats: format, Notes: notes, Links: links}
			if all {
				w = book.ClearWhat{Contents: true, Formats: true, Notes: true, Links: true}
			}
			at := whole
			if r != nil {
				at = *r
			}
			res, err := ed.ClearRange(sheet, at, w)
			if err != nil {
				return err
			}
			var what []string
			for _, x := range []struct {
				on   bool
				name string
			}{{w.Contents, "contents"}, {w.Formats, "formats"}, {w.Notes, "notes"}, {w.Links, "links"}} {
				if x.on {
					what = append(what, x.name)
				}
			}
			done := fmt.Sprintf("cleared %s of %s (%s)", strings.Join(what, ", "), res.To, plural(res.Cells, "cell"))
			if res.Cells == 0 {
				done = fmt.Sprintf("nothing to clear in %s (%s)", res.To, strings.Join(what, ", "))
			}
			return a.report(ed, args[0], done, res)
		},
	}
	f := cmd.Flags()
	f.BoolVar(&format, "format", false, "clear formatting, not contents")
	f.BoolVar(&notes, "notes", false, "clear notes")
	f.BoolVar(&links, "links", false, "clear hyperlinks")
	f.BoolVar(&all, "all", false, "clear contents, formatting, notes and links")
	return cmd
}

// destination reads where a copy or move goes: a cell or range, on the
// source's sheet unless it names another (created when missing).
func destination(ed book.Editor, fromSheet, ref string) (string, grid.Range, bool, error) {
	wb := ed.Workbook()
	p, err := grid.ParseRef(ref, func(n string) bool { return wb.Sheet(n) != nil })
	if err != nil {
		return "", grid.Range{}, false, err
	}
	if p.Range == nil {
		return "", grid.Range{}, false, fmt.Errorf("%q: give the cell to go to, like Sheet!B2", ref)
	}
	if p.Range.R2 == grid.MaxRows || p.Range.C2 == grid.MaxCols {
		r := *p.Range
		r.R2, r.C2 = min(r.R2, r.R1), min(r.C2, r.C1)
		p.Range = &r
	}
	sheet := p.Sheet
	if sheet == "" {
		sheet = fromSheet
	}
	name, created, err := ed.EnsureSheet(sheet)
	return name, *p.Range, created, err
}

func (a *app) copyCmd() *cobra.Command {
	var values, format bool
	cmd := &cobra.Command{
		Use:   "copy <file> <from> <to>",
		Short: "Copy cells, as Excel's copy and paste",
		Long: `Copy cells to another place, as Excel's copy and paste: values, formulas
(their relative references shift, A2 becoming A12 ten rows down), formats,
merged cells, notes, links, data validation, and the styling sheets style
and cond recorded. --values pastes values only (a formula gives its result,
with its number format); --format pastes formats only.

  sheets copy book.xlsx Sales!A1:D9 Summary!A1
  sheets copy book.xlsx Sales!E2 Sales!E3:E50     # fill a formula down
  sheets copy book.xlsx Model!B2:B20 Model!H2 --values

<to> is the top-left cell, or a range the copy fills evenly (as Excel
tiles a paste); without a sheet it is on <from>'s sheet. A missing sheet
is added. ` + strings.Replace(targetHelp, "The target is", "<from> is", 1),
		Args: cobra.ExactArgs(3),
		RunE: func(cmd *cobra.Command, args []string) error {
			if values && format {
				return errors.New("--values and --format exclude each other")
			}
			ed, from, src, err := editTarget(args[0], args[1])
			if err != nil {
				return err
			}
			defer ed.Close()
			r := whole
			if src != nil {
				r = *src
			}
			to, at, created, err := destination(ed, from, args[2])
			if err != nil {
				return err
			}
			res, err := ed.CopyRange(from, r, to, at, book.CopyOptions{Values: values, Formats: format})
			if err != nil {
				return err
			}
			what := ""
			switch {
			case values:
				what = " (values)"
			case format:
				what = " (formats)"
			}
			done := fmt.Sprintf("copied %s to %s%s", res.From, res.To, what)
			if created {
				done += " (new sheet)"
			}
			return a.report(ed, args[0], done, res)
		},
	}
	cmd.Flags().BoolVar(&values, "values", false, "paste values only: formulas give their results")
	cmd.Flags().BoolVar(&format, "format", false, "paste formats only")
	return cmd
}

func (a *app) moveCmd() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "move <file> <from> <to>",
		Short: "Move cells, as Excel's cut and paste; references follow",
		Long: `Move cells, as Excel's cut and paste. Everything goes with them: values,
formulas (which keep reading the cells they read), formats, merged cells,
notes, links, data validation, recorded styling, and conditional formats
lying within them. Formulas, names and charts anywhere that refer to the
moved cells follow them; references to the cells pasted over become
#REF!, as in Excel, and sheets lists them.

  sheets move book.xlsx Sales!H1:J20 Sales!A30
  sheets move book.xlsx Sales!A1:D9 Archive!A1

<to> is the top-left cell; without a sheet it is on <from>'s sheet. A
missing sheet is added. To move whole rows or columns between others, use
sheets rows/cols insert and delete; to reorder rows, sheets sort.`,
		Args: cobra.ExactArgs(3),
		RunE: func(cmd *cobra.Command, args []string) error {
			ed, from, src, err := editTarget(args[0], args[1])
			if err != nil {
				return err
			}
			defer ed.Close()
			r := whole
			if src != nil {
				r = *src
			}
			to, at, created, err := destination(ed, from, args[2])
			if err != nil {
				return err
			}
			res, err := ed.MoveRange(from, r, to, at)
			if err != nil {
				return err
			}
			done := fmt.Sprintf("moved %s to %s", res.From, res.To)
			if created {
				done += " (new sheet)"
			}
			return a.report(ed, args[0], done, res)
		},
	}
	return cmd
}

type sortResult struct {
	File string   `json:"file"`
	By   []string `json:"by"`
	book.SortResult
	Header   string   `json:"header,omitempty"`
	Warnings []string `json:"warnings,omitempty"`
	DryRun   bool     `json:"dry_run,omitempty"`
}

func (a *app) sortCmd() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "sort <file> <target> <column>[:desc]...",
		Short: "Sort the rows of a table or range by columns",
		Long: `Sort rows, as Excel does: by each column in turn, ascending unless :desc
follows its name. Numbers and dates come first, then text (ignoring case),
then FALSE and TRUE, then errors; blank cells always go last. Rows that
tie keep their order.

  sheets sort book.xlsx Sales Revenue:desc
  sheets sort book.xlsx Sales Region Revenue:desc
  sheets sort book.xlsx Sales!A5:F40 C

The target is a sheet holding one table or an Excel table (its data rows
are sorted; the header and total rows stay), or a range: its first row
stays when it is a table's header row. Columns are header names or
letters.

Whole rows of the target move, with their formats, notes and links.
Formulas move as in Excel, their relative references shifting with the
row; a sort that would change what a formula means (it reads another
row by position, or a column outside the target in its row) is refused
with the reason. Formulas elsewhere reading single cells of the sorted
rows are listed, since they now read whatever row is there.`,
		Args: func(cmd *cobra.Command, args []string) error {
			if len(args) < 3 {
				return errors.New("expected <file> <target> <column>..., like: sheets sort book.xlsx Sales Revenue:desc")
			}
			return nil
		},
		RunE: func(cmd *cobra.Command, args []string) error {
			ed, err := book.Edit(args[0], false)
			if err != nil {
				return err
			}
			defer ed.Close()
			wb := ed.Workbook()
			sheet, r, header, err := sortTarget(wb, args[1])
			if err != nil {
				return err
			}
			s := wb.Sheet(sheet)
			var keys []book.SortKey
			var by []string
			for _, k := range args[2:] {
				// The direction follows the last colon; a header may hold
				// colons of its own (Ratio: A/B).
				name, desc := k, false
				if i := strings.LastIndex(k, ":"); i >= 0 {
					switch strings.ToLower(k[i+1:]) {
					case "asc", "ascending":
						name = k[:i]
					case "desc", "descending":
						name, desc = k[:i], true
					}
				}
				col, label, err := sortColumn(wb, s, r, header, name)
				if err != nil && name != k {
					// A header such as "x:desc" itself.
					if c, l, err2 := sortColumn(wb, s, r, header, k); err2 == nil {
						col, label, err, desc = c, l, nil, false
					}
				}
				if err != nil && name == k && strings.Contains(k, ":") {
					return fmt.Errorf("%v; a direction goes after a colon: asc or desc", err)
				}
				if err != nil {
					return err
				}
				keys = append(keys, book.SortKey{Col: col, Desc: desc})
				by = append(by, label+map[bool]string{true: " descending", false: ""}[desc])
			}
			res, err := ed.SortRange(sheet, r, keys)
			if err != nil {
				return err
			}
			if err := ed.Save(a.dryRun); err != nil {
				return err
			}
			out := sortResult{File: args[0], By: by, SortResult: res, Warnings: ed.Warnings(), DryRun: a.dryRun}
			if header > 0 {
				out.Header = fmt.Sprint(header)
			}
			if a.json {
				return a.printJSON(out)
			}
			a.printDryRun()
			moved := "already in order"
			if res.Moved > 0 {
				moved = plural(res.Moved, "row") + " moved"
			}
			a.printf("sorted %s by %s (%s; %s)\n", res.Range, strings.Join(by, ", "), plural(res.Rows, "row"), moved)
			for _, n := range res.Notes {
				a.printf("note: %s\n", n)
			}
			a.printWarnings(out.Warnings)
			return nil
		},
	}
	return cmd
}

// sortTarget resolves what to sort: the data rows of a table (header and
// total rows stay), or a range, less its first row when that is a table's
// header row and its last when that is the total row. header is the
// header row, 0 when there is none.
func sortTarget(wb *grid.Workbook, ref string) (sheet string, r grid.Range, header int, err error) {
	name, rg, err := analyze.Resolve(wb, ref)
	if err != nil || rg == nil {
		// A sheet or an Excel table named alone: its data rows.
		n2, r2, err2 := analyze.Resolve(wb, strings.TrimSpace(ref)+"[#Data]")
		if err2 != nil {
			if err != nil {
				return "", grid.Range{}, 0, err
			}
			return "", grid.Range{}, 0, err2
		}
		name, rg = n2, r2
		h, _ := headerAbove(wb, name, *rg)
		return name, *rg, h, nil
	}
	r = *rg
	s := wb.Sheet(name)
	if r.R2 == grid.MaxRows || r.C2 == grid.MaxCols {
		used := s.Used()
		if r.R2 == grid.MaxRows {
			r.R1, r.R2 = max(r.R1, used.R1), used.R2
		}
		if r.C2 == grid.MaxCols {
			r.C1, r.C2 = max(r.C1, used.C1), used.C2
		}
	}
	for _, t := range s.Tables {
		if t.Range.R1 == r.R1 && t.Range.C1 <= r.C1 && r.C2 <= t.Range.C2 {
			header = r.R1
			r.R1++
		}
		if t.Totals > 0 && r.R2 == t.Range.R2 && t.Range.C1 <= r.C1 && r.C2 <= t.Range.C2 {
			r.R2--
		}
	}
	if header == 0 {
		for _, reg := range analyze.Regions(wb, s) {
			if !reg.IsTable() || reg.Range.Intersect(r).Empty() {
				continue
			}
			if reg.HeaderRow == r.R1 {
				header = r.R1
				r.R1++
			}
			if reg.TotalRow > 0 && reg.TotalRow == r.R2 {
				r.R2--
			}
		}
	}
	if header == 0 {
		// Data rows given without their header: name columns by it.
		if h, ok := headerAbove(wb, name, r); ok {
			header = h
		}
	}
	if r.Empty() {
		return "", grid.Range{}, 0, fmt.Errorf("%q holds no rows to sort", ref)
	}
	return name, r, header, nil
}

// headerAbove finds the header row of the table whose data rows are r.
func headerAbove(wb *grid.Workbook, sheet string, r grid.Range) (int, bool) {
	s := wb.Sheet(sheet)
	for _, t := range s.Tables {
		if t.Range.R1+1 == r.R1 && t.Range.C1 <= r.C1 && r.C2 <= t.Range.C2 {
			return t.Range.R1, true
		}
	}
	for _, reg := range analyze.Regions(wb, s) {
		if reg.IsTable() && reg.Body.R1 == r.R1 && !reg.Range.Intersect(r).Empty() {
			return reg.HeaderRow, true
		}
	}
	return 0, false
}

// sortColumn finds a column to sort by: a header name, or a letter.
func sortColumn(wb *grid.Workbook, s *grid.Sheet, r grid.Range, header int, name string) (int, string, error) {
	var labels []string
	if header > 0 {
		for c := r.C1; c <= r.C2; c++ {
			l := analyze.Label(wb, s, header, c)
			labels = append(labels, l)
			if strings.EqualFold(strings.TrimSpace(l), strings.TrimSpace(name)) {
				return c, l, nil
			}
		}
	}
	if c, ok := grid.ColIndex(strings.ToUpper(strings.TrimSpace(name))); ok {
		if c < r.C1 || c > r.C2 {
			return 0, "", fmt.Errorf("column %s is outside %s", strings.ToUpper(name), r)
		}
		label := grid.ColName(c)
		if header > 0 {
			if l := analyze.Label(wb, s, header, c); l != "" {
				label = l
			}
		}
		return c, label, nil
	}
	if len(labels) > 0 {
		return 0, "", fmt.Errorf("no column %q; the columns are %s (or give letters %s to %s)", name, strings.Join(labels, ", "), grid.ColName(r.C1), grid.ColName(r.C2))
	}
	return 0, "", fmt.Errorf("no column %q: %s has no header row, so give letters %s to %s", name, r, grid.ColName(r.C1), grid.ColName(r.C2))
}

type replaceResult struct {
	File     string             `json:"file"`
	Changed  []book.Replacement `json:"changed"`
	Total    int                `json:"total"`
	Warnings []string           `json:"warnings,omitempty"`
	DryRun   bool               `json:"dry_run,omitempty"`
}

func (a *app) replaceCmd() *cobra.Command {
	var useRegex, exact, caseSens, fx bool
	var limit int
	cmd := &cobra.Command{
		Use:   "replace <file> <find> <with> [target]",
		Short: "Find and replace text in cell values (or formulas)",
		Long: `Replace text in cells, as Excel's Replace does: in every sheet, or in a
target (a sheet, range or table reference). Matching ignores case unless
--case; --exact matches whole values; --regex takes a Go regular
expression, whose groups <with> can use as $1.

Values are matched as sheets shows them (ISO dates, 12%), and a changed
value is read again as typed input: "1,200" becomes a number, an empty
result clears the cell. Formulas are left alone unless --formulas, which
replaces in formula text instead (=SUM(B2:B9)).

  sheets replace book.xlsx 'N/A' '' Sales --exact
  sheets replace book.xlsx 'Acme Inc.' 'Acme Corp'
  sheets replace book.xlsx '^(\w+), (\w+)$' '$2 $1' 'Staff[Name]' --regex
  sheets replace book.xlsx 'Sheet1!' 'Data!' --formulas`,
		Args: cobra.RangeArgs(3, 4),
		RunE: func(cmd *cobra.Command, args []string) error {
			fn, err := replacer(args[1], args[2], useRegex, exact, caseSens)
			if err != nil {
				return err
			}
			var (
				ed    book.Editor
				sheet string
				r     *grid.Range
			)
			if len(args) == 4 {
				ed, sheet, r, err = editTarget(args[0], args[3])
			} else {
				ed, err = book.Edit(args[0], false)
			}
			if err != nil {
				return err
			}
			defer ed.Close()
			changed, err := ed.Replace(sheet, r, fn, fx)
			if err != nil {
				return err
			}
			if err := ed.Save(a.dryRun); err != nil {
				return err
			}
			out := replaceResult{File: args[0], Changed: changed, Total: len(changed), Warnings: ed.Warnings(), DryRun: a.dryRun}
			if limit > 0 && len(out.Changed) > limit {
				out.Changed = out.Changed[:limit]
			}
			if out.Changed == nil {
				out.Changed = []book.Replacement{}
			}
			if a.json {
				return a.printJSON(out)
			}
			a.printDryRun()
			if out.Total == 0 {
				a.printf("no cells matched\n")
				return nil
			}
			a.printf("replaced in %s:\n", plural(out.Total, "cell"))
			for _, c := range out.Changed {
				a.printf("%s\t%s → %s\n", c.Ref, grid.Escape(grid.Truncate(c.Old, 60)), grid.Escape(grid.Truncate(c.New, 60)))
			}
			if out.Total > len(out.Changed) {
				a.printf("… and %d more\n", out.Total-len(out.Changed))
			}
			a.printWarnings(out.Warnings)
			return nil
		},
	}
	f := cmd.Flags()
	f.BoolVar(&useRegex, "regex", false, "<find> is a regular expression")
	f.BoolVar(&exact, "exact", false, "match whole values only")
	f.BoolVar(&caseSens, "case", false, "match case")
	f.BoolVar(&fx, "formulas", false, "replace in formula text instead of values")
	f.IntVarP(&limit, "limit", "n", 20, "most changes to list")
	return cmd
}

// numberedGroups writes $1 as ${1}, leaving $$ (a dollar sign) and ${name}
// alone.
func numberedGroups(s string) string {
	var b strings.Builder
	for i := 0; i < len(s); i++ {
		if s[i] != '$' || i+1 >= len(s) {
			b.WriteByte(s[i])
			continue
		}
		if s[i+1] == '$' {
			b.WriteString("$$")
			i++
			continue
		}
		j := i + 1
		for j < len(s) && s[j] >= '0' && s[j] <= '9' {
			j++
		}
		if j == i+1 {
			b.WriteByte('$')
			continue
		}
		b.WriteString("${" + s[i+1:j] + "}")
		i = j - 1
	}
	return b.String()
}

// replacer builds the function replace runs over each value.
func replacer(find, with string, useRegex, exact, caseSens bool) (func(string) (string, bool), error) {
	if find == "" {
		return nil, errors.New("<find> is empty")
	}
	pattern := find
	if !useRegex {
		pattern = regexp.QuoteMeta(find)
		with = strings.ReplaceAll(with, "$", "$$")
	} else {
		// Checked as given, so an error shows the pattern as written.
		if _, err := regexp.Compile(find); err != nil {
			return nil, fmt.Errorf("--regex %q: %v", find, err)
		}
		// $1 is group 1 whatever follows, as in most tools; Go alone would
		// read $1TH as a group named 1TH.
		with = numberedGroups(with)
	}
	if exact {
		pattern = "^(?:" + pattern + ")$"
	}
	if !caseSens {
		pattern = "(?i)" + pattern
	}
	re, err := regexp.Compile(pattern)
	if err != nil {
		return nil, fmt.Errorf("--regex %q: %v", find, err)
	}
	return func(s string) (string, bool) {
		if !re.MatchString(s) {
			return s, false
		}
		return re.ReplaceAllString(s, with), true
	}, nil
}

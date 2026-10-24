package cli

import (
	"errors"
	"fmt"
	"strconv"
	"strings"

	"github.com/eemax/sheets/internal/analyze"
	"github.com/eemax/sheets/internal/book"
	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/words"
	"github.com/spf13/cobra"
)

// editTarget opens a workbook and resolves a target in it: a range, a
// sheet, or a table reference (Sales[Revenue]). r is nil for a sheet.
func editTarget(file, ref string) (book.Editor, string, *grid.Range, error) {
	ed, err := book.Edit(file, false)
	if err != nil {
		return nil, "", nil, err
	}
	sheet, r, err := analyze.Resolve(ed.Workbook(), ref)
	if err != nil {
		ed.Close()
		return nil, "", nil, err
	}
	return ed, sheet, r, nil
}

// save saves an edit (or not, on a dry run) and prints its warnings.
func (a *app) save(ed book.Editor) error {
	if err := ed.Save(a.dryRun); err != nil {
		return err
	}
	if !a.json {
		a.printDryRun()
	}
	return nil
}

const targetHelp = `The target is a range (Sales!B2:D20, B:D, 2:2), a sheet, or a table
reference in Excel's syntax, where the name is a sheet holding one table or
an Excel table: Sales[Revenue] is that column's data rows,
Sales[[Units]:[Revenue]] several columns, Sales[#Headers] the header row
(also #Data, #Totals, #All), Sales[[#Totals],[Units]] one total cell.`

func (a *app) styleCmd() *cobra.Command {
	var (
		list, show bool
		unrecord   []string
	)
	cmd := &cobra.Command{
		Use:   "style <file> <target> <property>...",
		Short: "Style cells: fill, font, number format, borders, alignment, widths",
		Long: `Lay style properties over cells, keeping what they do not set.

  sheets style book.xlsx 'Sales[#Headers]' fill=lightblue color=black bold
  sheets style book.xlsx 'Sales[Revenue]' format=currency decimals=0
  sheets style book.xlsx Sales!A1 size=16 bold
  sheets style book.xlsx 'Sales[#Totals]' border=top:double
  sheets style book.xlsx Sales!B:D width=14

` + targetHelp + ` Whole rows and columns style their used part. A number format
given a column (Sales[Revenue]) reaches the table's total row too.

Properties:
  bold italic underline strike wrap      (nobold, nowrap, ... undo them)
  font=Arial size=12
  color=red fill=lightblue fill=none     #RRGGBB or a name: ` + strings.Join(book.ColorNames(), ", ") + `
  format=TYPE [decimals=N]               integer, decimal, currency[:€], percent, date,
                                         datetime, time, text, general
  format='0.0"x"'                        or an Excel number format code
  align=left|center|right|center-across  valign=top|middle|bottom  indent=N  rotate=-90..90
  border=SIDES[:STYLE][:COLOR]           sides all, outline, inner, top, bottom, left,
                                         right, none (top+bottom); style thin, medium,
                                         thick, double, dashed, dotted
  width=N|auto (columns) height=N (rows)
  merge unmerge                          prefer align=center-across over merge
  reset                                  clear the look (number formats stay)

A fill without color= gets black or white text, whichever reads better.

fmt keeps what style sets: style records it in the workbook, as hidden
names whose ranges follow inserted and deleted rows and columns (in Excel
too), and fmt lays it over the house style again. Rows written below a
styled table get its look. --list shows the record; --unrecord N removes
entries (the styling stays until fmt runs again); fmt --reset clears it.

--show prints how cells look, in blocks, in the same properties:

  sheets style book.xlsx Sales!A1:D3 --show`,
		Args: func(cmd *cobra.Command, args []string) error {
			if list || len(unrecord) > 0 {
				return cobra.RangeArgs(1, 2)(cmd, args)
			}
			if show {
				return cobra.ExactArgs(2)(cmd, args)
			}
			if len(args) < 3 {
				return errors.New("expected <file> <target> <property>..., like: sheets style book.xlsx 'Sales[#Headers]' fill=lightblue bold")
			}
			return nil
		},
		RunE: func(cmd *cobra.Command, args []string) error {
			if list || len(unrecord) > 0 {
				return a.records(args, list, unrecord)
			}
			if show {
				return a.showLooks(args[0], args[1])
			}
			look, err := book.ParseLook(args[2:])
			if err != nil {
				return err
			}
			ed, sheet, r, err := editTarget(args[0], args[1])
			if err != nil {
				return err
			}
			defer ed.Close()
			at := grid.Range{R1: 1, C1: 1, R2: grid.MaxRows, C2: grid.MaxCols}
			if r != nil {
				at = *r
			}
			res, err := ed.Style(sheet, at, &look)
			if err != nil {
				return err
			}
			// A total is in its column's units: a number format given a
			// column's data rows (Sales[Revenue]) reaches the total row.
			var totals *book.StyleResult
			if look.NumFmt != "" {
				if ts, tr, ok := analyze.TotalsUnder(ed.Workbook(), args[1]); ok {
					var fw []string
					for _, w := range args[2:] {
						if k, _, _ := strings.Cut(strings.ToLower(w), "="); k == "format" || k == "decimals" {
							fw = append(fw, w)
						}
					}
					fl, err := book.ParseLook(fw)
					if err != nil {
						return err
					}
					tres, err := ed.Style(ts, tr, &fl)
					if err != nil {
						return err
					}
					totals = &tres
				}
			}
			if err := a.save(ed); err != nil {
				return err
			}
			if a.json {
				out := map[string]any{"range": res.Range, "recorded": res.Recorded, "notes": res.Notes}
				if totals != nil {
					out["totals"] = totals.Range
				}
				return a.printJSON(out)
			}
			a.printf("styled %s: %s\n", res.Range, words.Join(args[2:]))
			if totals != nil {
				a.printf("and the total row %s: the number format\n", totals.Range)
			}
			for _, n := range res.Notes {
				a.printf("note: %s\n", n)
			}
			a.printWarnings(ed.Warnings())
			return nil
		},
	}
	cmd.Flags().BoolVar(&list, "list", false, "list the recorded styling (of a sheet, if given)")
	cmd.Flags().BoolVar(&show, "show", false, "show how the target's cells look, as style properties")
	cmd.Flags().StringSliceVar(&unrecord, "unrecord", nil, "remove entries of the record by number (see --list)")
	return cmd
}

// showLooks prints the looks of a target's cells.
func (a *app) showLooks(file, ref string) error {
	ed, sheet, r, err := editTarget(file, ref)
	if err != nil {
		return err
	}
	defer ed.Close()
	at := grid.Range{R1: 1, C1: 1, R2: grid.MaxRows, C2: grid.MaxCols}
	if r != nil {
		at = *r
	}
	runs, err := ed.Looks(sheet, at)
	if err != nil {
		return err
	}
	if a.json {
		return a.printJSON(runs)
	}
	w := 0
	for _, l := range runs {
		w = max(w, len(l.Range))
	}
	for _, l := range runs {
		look := words.Join(l.Look)
		if len(l.Look) == 0 {
			look = "(no style)"
		}
		a.printf("%-*s  %s\n", w, l.Range, look)
	}
	return nil
}

// records lists or edits the record of styling.
func (a *app) records(args []string, list bool, unrecord []string) error {
	ed, err := book.Edit(args[0], false)
	if err != nil {
		return err
	}
	defer ed.Close()
	if len(unrecord) > 0 {
		var seqs []int
		for _, u := range unrecord {
			n, err := strconv.Atoi(strings.TrimPrefix(strings.TrimSpace(u), "#"))
			if err != nil {
				return fmt.Errorf("--unrecord %s: expected the number --list shows", u)
			}
			seqs = append(seqs, n)
		}
		if err := ed.Unrecord(seqs); err != nil {
			return err
		}
		if err := a.save(ed); err != nil {
			return err
		}
		if a.json {
			return a.printJSON(map[string]any{"file": args[0], "removed": seqs})
		}
		a.printf("removed %s from the record; the styling stays until fmt runs\n", plural(len(seqs), "entry"))
		return nil
	}
	sheet := ""
	if len(args) > 1 {
		s, _, err := analyze.Resolve(ed.Workbook(), args[1])
		if err != nil {
			return err
		}
		sheet = s
	}
	var recs []book.Record
	for _, r := range ed.Records() {
		if sheet == "" || strings.EqualFold(r.Sheet, sheet) {
			recs = append(recs, r)
		}
	}
	if a.json {
		if recs == nil {
			recs = []book.Record{}
		}
		return a.printJSON(recs)
	}
	if len(recs) == 0 {
		a.printf("no recorded styling\n")
		return nil
	}
	w, n := 0, 0
	for _, r := range recs {
		w, n = max(w, len(r.Ref)), max(n, len(strconv.Itoa(r.Seq)))
	}
	for _, r := range recs {
		a.printf("#%-*d  %-*s  %s\n", n, r.Seq, w, r.Ref, r.Text())
	}
	return nil
}

func (a *app) condCmd() *cobra.Command {
	var list, clear bool
	cmd := &cobra.Command{
		Use:   "cond <file> <target> <condition> [property]...",
		Short: "Highlight cells that meet a condition; colour scales, data bars, icons",
		Long: `Add a conditional format: cells meeting the condition take the look the
properties give (default: Excel's light red fill with dark red text).

  sheets cond book.xlsx 'Sales[Margin]' '< 0' color=red
  sheets cond book.xlsx 'Sales[Revenue]' 'top 10' fill=lightgreen bold
  sheets cond book.xlsx Sales!A2:F200 '=$E2="Late"' fill=lightyellow   # whole rows
  sheets cond book.xlsx 'Sales[Units]' bars
  sheets cond book.xlsx 'Sales[Growth]' scale red-yellow-green

` + targetHelp + `

Conditions:
  '< 0'  '>= 100'  '= East'  '<> 0'          compare each cell (quote: < and > are shell syntax)
  'between 1 and 10'  'not between 1 and 10'
  'top 10'  'bottom 5'  'top 10%'  'above average'  'below average'
  duplicates  unique  blank  'not blank'  errors  'no errors'
  'contains late'  'not contains late'  'begins with A'  'ends with .pdf'
  '=$E2="Late"'                             a formula for the target's first cell; $ fixes a column
  scale [LOW-HIGH | LOW-MID-HIGH]            colour scale, default red-yellow-green (highest green)
  bars [COLOR]                               data bars
  icons [arrows|traffic|flags|symbols|ratings]

Properties: color, fill, bold, italic, underline, strike, border, format (as in
sheets style). Values: numbers, dates (2026-07-01), 10%, text, or =formula.

fmt keeps conditional formats added here (and drops the rest, as it strips
formatting). --list shows a sheet's conditional formats; --clear removes
those within the target.`,
		Args: func(cmd *cobra.Command, args []string) error {
			if list || clear {
				return cobra.RangeArgs(1, 2)(cmd, args)
			}
			if len(args) < 3 {
				return errors.New("expected <file> <target> <condition> [property]..., like: sheets cond book.xlsx 'Sales[Margin]' '< 0' color=red")
			}
			return nil
		},
		RunE: func(cmd *cobra.Command, args []string) error {
			ref := ""
			if len(args) > 1 {
				ref = args[1]
			}
			ed, sheet, r, err := editTarget(args[0], ref)
			if err != nil {
				return err
			}
			defer ed.Close()
			switch {
			case list:
				cfs, err := ed.CondFormats(sheet)
				if err != nil {
					return err
				}
				if r != nil {
					var in []book.CondFormat
					for _, c := range cfs {
						if rr, err := grid.ParseRange(strings.Fields(c.Range)[0]); err == nil && !rr.Intersect(*r).Empty() {
							in = append(in, c)
						}
					}
					cfs = in
				}
				if a.json {
					if cfs == nil {
						cfs = []book.CondFormat{}
					}
					return a.printJSON(cfs)
				}
				if len(cfs) == 0 {
					a.printf("%s: no conditional formats\n", sheet)
					return nil
				}
				w := 0
				for _, c := range cfs {
					w = max(w, len(c.Range))
				}
				for _, c := range cfs {
					tag := ""
					switch {
					case c.Banding:
						tag = "  (fmt banding)"
					case c.Recorded:
						tag = "  (kept by fmt)"
					}
					a.printf("%s!%-*s  %s%s\n", grid.QuoteSheet(sheet), w, c.Range, c.Rule, tag)
				}
				return nil
			case clear:
				n, err := ed.ClearCond(sheet, r)
				if err != nil {
					return err
				}
				if err := a.save(ed); err != nil {
					return err
				}
				if a.json {
					return a.printJSON(map[string]any{"sheet": sheet, "removed": n})
				}
				a.printf("removed %s from %s\n", plural(n, "conditional format"), targetText(sheet, r))
				return nil
			}
			if r == nil {
				return fmt.Errorf("%q is a sheet; give the cells, like %s!B2:B100 or '%s[Column]'", args[1], sheet, sheet)
			}
			at, err := ed.Cond(sheet, *r, args[2:])
			if err != nil {
				return err
			}
			if err := a.save(ed); err != nil {
				return err
			}
			if a.json {
				return a.printJSON(map[string]any{"range": at, "rule": args[2:]})
			}
			a.printf("%s: %s\n", at, words.Join(args[2:]))
			a.printWarnings(ed.Warnings())
			return nil
		},
	}
	cmd.Flags().BoolVar(&list, "list", false, "list the conditional formats of the sheet (within the target, if a range)")
	cmd.Flags().BoolVar(&clear, "clear", false, "remove the conditional formats within the target")
	return cmd
}

func targetText(sheet string, r *grid.Range) string {
	if r == nil {
		return grid.QuoteSheet(sheet)
	}
	return grid.FormatRef(sheet, *r)
}

func (a *app) validateCmd() *cobra.Command {
	var (
		list, clear, warn, noBlank bool
		prompt, message            string
	)
	cmd := &cobra.Command{
		Use:   "validate <file> <target> <rule>...",
		Short: "Restrict what cells accept: dropdown lists, number or date ranges",
		Long: `Add a data validation rule: what cells accept, with a dropdown for lists.
It replaces the rules within the target.

  sheets validate book.xlsx 'Sales[Region]' list North,South,East,West
  sheets validate book.xlsx 'Sales[Region]' list =Lists!$A$2:$A$9
  sheets validate book.xlsx 'Sales[Units]' whole '>= 0' --error 'Units cannot be negative'
  sheets validate book.xlsx 'Sales[Date]' date between 2026-01-01 and 2026-12-31
  sheets validate book.xlsx 'Sales[Code]' length '<= 8' --prompt 'Up to 8 characters'
  sheets validate book.xlsx 'Sales[ID]' formula '=COUNTIF($A:$A,A2)=1' --error 'IDs are unique'

` + targetHelp + `

Rules:
  list A,B,C | list =Sheet!$A$2:$A$9     a dropdown; the range form for long lists
  whole|number|date|time|length OP X     OP: < <= > >= = <>  (quote: '>= 0')
  whole|number|date|time|length between X and Y
  formula =...                            true for the target's first cell
  any                                     accept anything; with --prompt, a hint

Values are numbers, dates (2026-07-01, today), times (17:30) or =formulas.
A refused value gets --error's message (or Excel's own); --warn lets it in
after a warning. --list shows a sheet's rules; --clear removes those
within the target. fmt leaves validation alone.`,
		Args: func(cmd *cobra.Command, args []string) error {
			if list || clear {
				return cobra.RangeArgs(1, 2)(cmd, args)
			}
			if len(args) < 3 {
				return errors.New("expected <file> <target> <rule>..., like: sheets validate book.xlsx 'Sales[Region]' list North,South")
			}
			return nil
		},
		RunE: func(cmd *cobra.Command, args []string) error {
			ref := ""
			if len(args) > 1 {
				ref = args[1]
			}
			ed, sheet, r, err := editTarget(args[0], ref)
			if err != nil {
				return err
			}
			defer ed.Close()
			switch {
			case list:
				vs, err := ed.Validations(sheet)
				if err != nil {
					return err
				}
				if a.json {
					if vs == nil {
						vs = []book.Validation{}
					}
					return a.printJSON(vs)
				}
				if len(vs) == 0 {
					a.printf("%s: no validation rules\n", sheet)
					return nil
				}
				w := 0
				for _, v := range vs {
					w = max(w, len(v.Range))
				}
				for _, v := range vs {
					line := fmt.Sprintf("%s!%-*s  %s", grid.QuoteSheet(sheet), w, v.Range, v.Rule)
					if v.Prompt != "" {
						line += fmt.Sprintf("  prompt %q", v.Prompt)
					}
					if v.Error != "" {
						line += fmt.Sprintf("  error %q", v.Error)
					}
					a.printf("%s\n", line)
				}
				return nil
			case clear:
				n, err := ed.ClearValidation(sheet, r)
				if err != nil {
					return err
				}
				if err := a.save(ed); err != nil {
					return err
				}
				if a.json {
					return a.printJSON(map[string]any{"sheet": sheet, "removed": n})
				}
				a.printf("removed %s from %s\n", plural(n, "validation rule"), targetText(sheet, r))
				return nil
			}
			if r == nil {
				return fmt.Errorf("%q is a sheet; give the cells, like %s!B2:B100 or '%s[Column]'", args[1], sheet, sheet)
			}
			at, err := ed.Validate(sheet, *r, book.ValidationRule{Words: args[2:], Prompt: prompt, Message: message, Warn: warn, NoBlank: noBlank})
			if err != nil {
				return err
			}
			if err := a.save(ed); err != nil {
				return err
			}
			if a.json {
				return a.printJSON(map[string]any{"range": at, "rule": args[2:]})
			}
			a.printf("%s: %s\n", at, words.Join(args[2:]))
			return nil
		},
	}
	f := cmd.Flags()
	f.BoolVar(&list, "list", false, "list the validation rules of the sheet")
	f.BoolVar(&clear, "clear", false, "remove the validation rules within the target")
	f.StringVar(&prompt, "prompt", "", "a hint shown when a cell is selected")
	f.StringVar(&message, "error", "", "the message when a value is refused")
	f.BoolVar(&warn, "warn", false, "warn about a value instead of refusing it")
	f.BoolVar(&noBlank, "no-blank", false, "refuse empty cells too")
	return cmd
}

func (a *app) viewCmd() *cobra.Command {
	var (
		freeze, tab, gridlines, filter string
		zoom                           int
		hide, show                     []string
		active                         bool
	)
	cmd := &cobra.Command{
		Use:   "view <file> [sheet]",
		Short: "Freeze panes, tab colour, gridlines, zoom, hidden rows and columns, filter",
		Long: `Change how a sheet shows, or with no flags, report it.

  sheets view book.xlsx Sales --freeze B2        # row 1 and column A stay in view
  sheets view book.xlsx Sales --freeze none
  sheets view book.xlsx Sales --tab green --gridlines off --zoom 120
  sheets view book.xlsx Sales --hide C:D --hide 40:45 --show F
  sheets view book.xlsx Sales --filter A1:F200   # filter buttons; --filter off
  sheets view book.xlsx Sales --active           # the sheet Excel opens on

--freeze takes the first cell that scrolls (B2 keeps row 1 and column A),
or a number of rows (--freeze 1 is the same as A2). fmt freezes a table's
header and sets gridlines and filters itself; it keeps what view sets.`,
		Args: cobra.RangeArgs(1, 2),
		RunE: func(cmd *cobra.Command, args []string) error {
			ref := ""
			if len(args) > 1 {
				ref = args[1]
			}
			ed, sheet, r, err := editTarget(args[0], ref)
			if err != nil {
				return err
			}
			defer ed.Close()
			if r != nil {
				return fmt.Errorf("%q: give a sheet name; view works on whole sheets", ref)
			}
			var o book.ViewOptions
			changed := false
			fl := cmd.Flags()
			if fl.Changed("freeze") {
				at, err := freezeCell(freeze)
				if err != nil {
					return err
				}
				o.Freeze, changed = &at, true
			}
			if fl.Changed("tab") {
				c := ""
				if !strings.EqualFold(tab, "none") {
					if c, err = book.ParseColor(tab); err != nil {
						return fmt.Errorf("--tab: %v", err)
					}
				}
				o.Tab, changed = &c, true
			}
			if fl.Changed("gridlines") {
				on, err := onOff("--gridlines", gridlines)
				if err != nil {
					return err
				}
				o.Gridlines, changed = &on, true
			}
			if fl.Changed("zoom") {
				if zoom < 10 || zoom > 400 {
					return errors.New("--zoom: expected 10 to 400 (percent)")
				}
				o.Zoom, changed = &zoom, true
			}
			for _, h := range hide {
				l, err := lines(h)
				if err != nil {
					return fmt.Errorf("--hide %s: %v", h, err)
				}
				o.Hide, changed = append(o.Hide, l), true
			}
			for _, s := range show {
				l, err := lines(s)
				if err != nil {
					return fmt.Errorf("--show %s: %v", s, err)
				}
				o.Show, changed = append(o.Show, l), true
			}
			if fl.Changed("filter") {
				if strings.EqualFold(filter, "off") || strings.EqualFold(filter, "none") {
					o.NoFilter, changed = true, true
				} else {
					fr := filter
					if !strings.Contains(fr, "!") && !strings.Contains(fr, "[") {
						fr = grid.QuoteSheet(sheet) + "!" + fr
					}
					fs, rr, err := analyze.Resolve(ed.Workbook(), fr)
					if err != nil || rr == nil || !strings.EqualFold(fs, sheet) {
						return fmt.Errorf("--filter %s: give a range on %s, like A1:F200, or off", filter, sheet)
					}
					o.Filter, changed = rr, true
				}
			}
			if active {
				o.Active, changed = true, true
			}
			if changed {
				if err := ed.View(sheet, o); err != nil {
					return err
				}
				if err := a.save(ed); err != nil {
					return err
				}
			}
			v, err := ed.ViewOf(sheet)
			if err != nil {
				return err
			}
			if a.json {
				return a.printJSON(v)
			}
			a.printf("%s\n", viewText(v))
			return nil
		},
	}
	f := cmd.Flags()
	f.StringVar(&freeze, "freeze", "", "the first cell that scrolls (B2), a number of rows, or none")
	f.StringVar(&tab, "tab", "", "tab colour (#RRGGBB or a name), or none")
	f.StringVar(&gridlines, "gridlines", "", "on or off")
	f.IntVar(&zoom, "zoom", 0, "zoom in percent")
	f.StringArrayVar(&hide, "hide", nil, "hide rows (5:9) or columns (C:D) (repeatable)")
	f.StringArrayVar(&show, "show", nil, "show hidden rows or columns again (repeatable)")
	f.StringVar(&filter, "filter", "", "put filter buttons on a range's first row, or off")
	f.BoolVar(&active, "active", false, "make this the sheet Excel opens on")
	return cmd
}

func onOff(flag, v string) (bool, error) {
	switch strings.ToLower(v) {
	case "on", "true", "yes", "show":
		return true, nil
	case "off", "false", "no", "hide":
		return false, nil
	}
	return false, fmt.Errorf("%s: expected on or off", flag)
}

// freezeCell reads --freeze: a cell, a number of rows, or none.
func freezeCell(v string) (grid.Range, error) {
	v = strings.TrimSpace(v)
	if strings.EqualFold(v, "none") || strings.EqualFold(v, "off") {
		return grid.Range{R1: 1, C1: 1, R2: 1, C2: 1}, nil
	}
	if n, err := strconv.Atoi(v); err == nil && n >= 0 && n < grid.MaxRows {
		return grid.Range{R1: n + 1, C1: 1, R2: n + 1, C2: 1}, nil
	}
	if r, c, ok := grid.ParseCell(strings.ReplaceAll(v, "$", "")); ok {
		return grid.Range{R1: r, C1: c, R2: r, C2: c}, nil
	}
	return grid.Range{}, fmt.Errorf("--freeze %s: expected the first cell that scrolls (B2), a number of rows, or none", v)
}

// lines reads whole rows (5, 5:9) or columns (C, C:E).
func lines(v string) (grid.Range, error) {
	v = strings.TrimSpace(v)
	if !strings.Contains(v, ":") {
		v = v + ":" + v
	}
	r, err := grid.ParseRange(v)
	if err != nil || !(r.R1 == 1 && r.R2 == grid.MaxRows) && !(r.C1 == 1 && r.C2 == grid.MaxCols) {
		return grid.Range{}, errors.New("expected rows (5 or 5:9) or columns (C or C:E)")
	}
	return r, nil
}

func viewText(v book.ViewState) string {
	var parts []string
	if v.Freeze != "" {
		parts = append(parts, "frozen at "+v.Freeze)
	} else {
		parts = append(parts, "not frozen")
	}
	parts = append(parts, map[bool]string{true: "gridlines on", false: "gridlines off"}[v.Gridlines])
	parts = append(parts, fmt.Sprintf("zoom %d%%", v.Zoom))
	if v.Tab != "" {
		parts = append(parts, "tab "+v.Tab)
	}
	if v.Filter != "" {
		parts = append(parts, "filter "+v.Filter)
	}
	if len(v.HiddenRows) > 0 {
		parts = append(parts, "hidden rows "+strings.Join(v.HiddenRows, ", "))
	}
	if len(v.HiddenCols) > 0 {
		parts = append(parts, "hidden columns "+strings.Join(v.HiddenCols, ", "))
	}
	if v.Active {
		parts = append(parts, "active")
	}
	return v.Sheet + ": " + strings.Join(parts, ", ")
}

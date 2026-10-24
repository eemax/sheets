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

type noteItem struct {
	Ref    string `json:"ref"`
	Author string `json:"author,omitempty"`
	Text   string `json:"text"`
	Thread bool   `json:"thread,omitempty"`
}

type linkItem struct {
	Ref      string `json:"ref"`
	Text     string `json:"text,omitempty"`
	URL      string `json:"url,omitempty"`
	Location string `json:"location,omitempty"`
	Tooltip  string `json:"tooltip,omitempty"`
}

// listScope resolves the sheets and range a --list covers: every sheet, or
// a target.
func listScope(wb *grid.Workbook, args []string) ([]*grid.Sheet, *grid.Range, error) {
	if len(args) < 2 {
		return wb.Sheets, nil, nil
	}
	name, r, err := analyze.Resolve(wb, args[1])
	if err != nil {
		return nil, nil, err
	}
	return []*grid.Sheet{wb.Sheet(name)}, r, nil
}

func (a *app) noteCmd() *cobra.Command {
	var list bool
	var author string
	cmd := &cobra.Command{
		Use:   "note <file> <cell> <text>",
		Short: "Put a note on a cell, or list notes",
		Long: `Put a note on a cell (Excel's notes, shown on hovering), replacing the one
there; empty text removes it. --list prints the notes of the workbook or a
target, comment threads from Excel 365 included (sheets cannot change
those).

  sheets note book.xlsx Sales!D7 'Includes the March refund'
  sheets note book.xlsx Sales!D7 ''
  sheets note book.xlsx --list
  sheets note book.xlsx Sales --list`,
		Args: func(cmd *cobra.Command, args []string) error {
			if list {
				return cobra.RangeArgs(1, 2)(cmd, args)
			}
			if len(args) != 3 {
				return errors.New("expected <file> <cell> <text>, like: sheets note book.xlsx Sales!D7 'Checked'; or --list")
			}
			return nil
		},
		RunE: func(cmd *cobra.Command, args []string) error {
			if list {
				return a.listNotes(args)
			}
			ed, err := book.Edit(args[0], false)
			if err != nil {
				return err
			}
			defer ed.Close()
			sheet, row, col, err := oneCell(ed, args[1])
			if err != nil {
				return err
			}
			if err := ed.SetNote(sheet, row, col, args[2], author); err != nil {
				return err
			}
			ref := grid.FormatRef(sheet, grid.Range{R1: row, C1: col, R2: row, C2: col})
			done := "note on " + ref + ": " + grid.Escape(grid.Truncate(args[2], 60))
			if args[2] == "" {
				done = "removed the note on " + ref
			}
			return a.report(ed, args[0], done, book.RangeResult{To: ref, Cells: 1})
		},
	}
	cmd.Flags().BoolVar(&list, "list", false, "list notes")
	cmd.Flags().StringVar(&author, "author", "sheets", "the note's author")
	return cmd
}

// oneCell resolves a single cell to change.
func oneCell(ed book.Editor, ref string) (string, int, int, error) {
	sheet, r, err := analyze.Resolve(ed.Workbook(), ref)
	if err != nil {
		return "", 0, 0, err
	}
	if r == nil || r.Rows() != 1 || r.Cols() != 1 {
		return "", 0, 0, fmt.Errorf("%q: give one cell, like Sheet!B4", ref)
	}
	return sheet, r.R1, r.C1, nil
}

func (a *app) listNotes(args []string) error {
	wb, err := open(args[0])
	if err != nil {
		return err
	}
	sheets, r, err := listScope(wb, args)
	if err != nil {
		return err
	}
	items := []noteItem{}
	for _, s := range sheets {
		for _, n := range s.Notes {
			if r == nil || r.Contains(n.Row, n.Col) {
				items = append(items, noteItem{Ref: grid.FormatRef(s.Name, grid.Range{R1: n.Row, C1: n.Col, R2: n.Row, C2: n.Col}), Author: n.Author, Text: n.Text, Thread: n.Thread})
			}
		}
	}
	if a.json {
		return a.printJSON(items)
	}
	if len(items) == 0 {
		a.printf("no notes\n")
		return nil
	}
	for _, n := range items {
		who := ""
		if n.Author != "" {
			who = "[" + n.Author + "] "
		}
		if n.Thread {
			who += "(thread) "
		}
		a.printf("%s\t%s%s\n", n.Ref, who, grid.Escape(n.Text))
	}
	return nil
}

var urlScheme = regexp.MustCompile(`^[a-zA-Z][a-zA-Z0-9+.-]*:`)

func (a *app) linkCmd() *cobra.Command {
	var list bool
	var text, tip string
	cmd := &cobra.Command{
		Use:   "link <file> <cell> <address>",
		Short: "Put a hyperlink on a cell, or list hyperlinks",
		Long: `Put a hyperlink on a cell: to a web address (https://...), an e-mail
(mailto:...), a file, or a place in the workbook (Sheet!A1, or a name). An
empty cell shows the address, or --text; the cell gets the look of a link,
which fmt keeps. An empty address removes the link. --list prints the
links of the workbook or a target.

  sheets link book.xlsx Sales!A1 https://example.com/q3 --text 'Q3 report'
  sheets link book.xlsx Summary!B2 Sales!A1 --tip 'Go to the data'
  sheets link book.xlsx Sales!A1 ''
  sheets link book.xlsx --list`,
		Args: func(cmd *cobra.Command, args []string) error {
			if list {
				return cobra.RangeArgs(1, 2)(cmd, args)
			}
			if len(args) != 3 {
				return errors.New("expected <file> <cell> <address>, like: sheets link book.xlsx Sales!A1 https://example.com; or --list")
			}
			return nil
		},
		RunE: func(cmd *cobra.Command, args []string) error {
			if list {
				return a.listLinks(args)
			}
			ed, err := book.Edit(args[0], false)
			if err != nil {
				return err
			}
			defer ed.Close()
			sheet, row, col, err := oneCell(ed, args[1])
			if err != nil {
				return err
			}
			ref := grid.FormatRef(sheet, grid.Range{R1: row, C1: col, R2: row, C2: col})
			addr := strings.TrimSpace(args[2])
			if addr == "" {
				if err := ed.SetLink(sheet, row, col, nil, ""); err != nil {
					return err
				}
				return a.report(ed, args[0], "removed the link on "+ref, book.RangeResult{To: ref, Cells: 1})
			}
			l, err := linkTarget(ed.Workbook(), addr)
			if err != nil {
				return err
			}
			l.Tooltip = tip
			if err := ed.SetLink(sheet, row, col, l, text); err != nil {
				return err
			}
			to := l.URL
			if to == "" {
				to = l.Location
			}
			return a.report(ed, args[0], "link on "+ref+" to "+to, book.RangeResult{To: ref, Cells: 1})
		},
	}
	f := cmd.Flags()
	f.BoolVar(&list, "list", false, "list hyperlinks")
	f.StringVar(&text, "text", "", "text the cell shows")
	f.StringVar(&tip, "tip", "", "tooltip shown on hovering")
	return cmd
}

// linkTarget reads a link's address: a place in the workbook, or a web
// address, e-mail or file.
func linkTarget(wb *grid.Workbook, addr string) (*grid.Link, error) {
	place := strings.TrimPrefix(addr, "#")
	if p, err := grid.ParseRef(place, func(n string) bool { return wb.Sheet(n) != nil }); err == nil && p.Sheet != "" && p.Range != nil {
		if s := wb.Sheet(p.Sheet); s != nil {
			return &grid.Link{Location: grid.QuoteSheet(s.Name) + "!" + p.Range.String()}, nil
		}
	}
	for _, n := range wb.Names {
		if strings.EqualFold(n.Name, place) {
			return &grid.Link{Location: n.Name}, nil
		}
	}
	switch {
	case urlScheme.MatchString(addr) && !strings.HasPrefix(addr, "#"):
		return &grid.Link{URL: addr}, nil
	case strings.HasPrefix(strings.ToLower(addr), "www."):
		return &grid.Link{URL: "https://" + addr}, nil
	case strings.Contains(addr, "@") && !strings.ContainsAny(addr, " /"):
		return &grid.Link{URL: "mailto:" + addr}, nil
	case strings.Contains(place, "!"):
		sheet := place[:strings.LastIndex(place, "!")]
		return nil, fmt.Errorf("%q: no sheet %s (sheets: %s)", addr, sheet, strings.Join(wb.SheetNames(), ", "))
	}
	return nil, fmt.Errorf("%q: give a web address (https://...), mailto:, a file (file:///...), or a place in the workbook (Sheet!A1, or a name)", addr)
}

func (a *app) listLinks(args []string) error {
	wb, err := open(args[0])
	if err != nil {
		return err
	}
	sheets, r, err := listScope(wb, args)
	if err != nil {
		return err
	}
	items := []linkItem{}
	for _, s := range sheets {
		for _, l := range s.Links {
			if r != nil && !r.Contains(l.Row, l.Col) {
				continue
			}
			it := linkItem{Ref: grid.FormatRef(s.Name, grid.Range{R1: l.Row, C1: l.Col, R2: l.Row, C2: l.Col}), URL: l.URL, Location: l.Location, Tooltip: l.Tooltip}
			if c := wb.Value(s, l.Row, l.Col); c != nil && !c.IsEmpty() {
				it.Text = grid.Display(c, wb.FormatOf(c), wb.Date1904)
			}
			items = append(items, it)
		}
	}
	if a.json {
		return a.printJSON(items)
	}
	if len(items) == 0 {
		a.printf("no links\n")
		return nil
	}
	for _, l := range items {
		to := l.URL
		if to == "" {
			to = "#" + l.Location
		}
		line := fmt.Sprintf("%s\t%s → %s", l.Ref, grid.Escape(l.Text), to)
		if l.Tooltip != "" {
			line += fmt.Sprintf(" (%s)", l.Tooltip)
		}
		a.printf("%s\n", line)
	}
	return nil
}

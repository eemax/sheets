package cli

import (
	"fmt"
	"strings"

	"github.com/eemax/sheets/internal/book"
	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/style"
	"github.com/spf13/cobra"
)

type fmtResult struct {
	File      string `json:"file"`
	Preset    string `json:"preset"`
	Formatted bool   `json:"formatted"` // for --check: already formatted
	*book.FormatResult
}

func (a *app) fmtCmd() *cobra.Command {
	var (
		preset string
		check  bool
		sheets []string
		as     []string
		list   bool
	)
	cmd := &cobra.Command{
		Use:   "fmt <file>",
		Short: "Strip formatting and apply the house style",
		Long: `Remove existing cell formatting and apply the house style: a dark header
row with filters, banded data rows, thin borders, a bold total row, frozen
header, fitted column widths and consistent number formats (thousands
separators, negatives in parentheses, ISO dates).

Formatting depends only on the cell contents, so running fmt twice changes
nothing and --check can verify a file in CI (exit status 1 when it would
change).

Column types are inferred. Override them with --as, by header or letter:

  sheets fmt book.xlsx --as Revenue=currency --as D=percent --as 'Q1!Cost=currency:€'

Types: integer, decimal, currency[:symbol], percent, date, datetime, time, text.`,
		Args: func(cmd *cobra.Command, args []string) error {
			if list {
				return nil
			}
			return cobra.ExactArgs(1)(cmd, args)
		},
		RunE: func(cmd *cobra.Command, args []string) error {
			if list {
				for _, p := range style.Presets {
					a.printf("%s\t%s\n", p.Name, p.Description)
				}
				return nil
			}
			p, ok := style.Lookup(preset)
			if !ok {
				var names []string
				for _, p := range style.Presets {
					names = append(names, p.Name)
				}
				return fmt.Errorf("unknown preset %q (presets: %s)", preset, strings.Join(names, ", "))
			}
			opts := book.FormatOptions{Preset: p, Sheets: sheets, DryRun: a.dryRun}
			for _, s := range as {
				o, err := parseOverride(s)
				if err != nil {
					return err
				}
				opts.Overrides = append(opts.Overrides, o)
			}
			file := args[0]
			if check {
				formatted, fr, err := book.Check(file, opts)
				if err != nil {
					return err
				}
				if a.json {
					if err := a.printJSON(fmtResult{File: file, Preset: p.Name, Formatted: formatted, FormatResult: fr}); err != nil {
						return err
					}
				} else if formatted {
					a.printf("%s is formatted\n", file)
				} else {
					a.printf("%s needs formatting\n", file)
					a.hint("sheets fmt %s", shellQuote(file))
				}
				if !formatted {
					return errSilent
				}
				return nil
			}
			fr, err := book.Format(file, opts)
			if err != nil {
				return err
			}
			if a.json {
				return a.printJSON(fmtResult{File: file, Preset: p.Name, Formatted: true, FormatResult: fr})
			}
			a.printDryRun()
			a.printf("formatted %s · preset %s\n", file, p.Name)
			for _, s := range fr.Sheets {
				for _, n := range s.Notes {
					a.printf("%s  note: %s\n", s.Name, n)
				}
				if len(s.Tables) == 0 {
					a.printf("%s  %s\n", s.Name, plural(s.Blocks, "block"))
					continue
				}
				for _, t := range s.Tables {
					var cols []string
					for _, c := range t.Columns {
						label := c.Header
						if label == "" {
							label = c.Col
						}
						cols = append(cols, label+" "+c.Type)
					}
					a.printf("%s  table %s: %s\n", s.Name, t.Range, strings.Join(cols, ", "))
				}
			}
			return nil
		},
	}
	f := cmd.Flags()
	f.StringVarP(&preset, "preset", "p", "default", "style preset (see --list-presets)")
	f.BoolVar(&check, "check", false, "exit 1 if the file is not formatted; change nothing")
	f.StringArrayVar(&sheets, "sheet", nil, "format only this sheet (repeatable)")
	f.StringArrayVar(&as, "as", nil, "force a column type: COLUMN=TYPE (repeatable)")
	f.BoolVar(&list, "list-presets", false, "list style presets")
	return cmd
}

// parseOverride parses [Sheet!]Column=type[:symbol].
func parseOverride(s string) (book.Override, error) {
	i := strings.LastIndex(s, "=")
	if i <= 0 {
		return book.Override{}, fmt.Errorf("--as %q: expected COLUMN=TYPE", s)
	}
	col, typ := s[:i], s[i+1:]
	o := book.Override{}
	if j := strings.LastIndex(col, "!"); j >= 0 {
		o.Sheet = strings.Trim(col[:j], "'")
		col = col[j+1:]
	}
	o.Column = col
	kind, sym, _ := strings.Cut(typ, ":")
	k, ok := grid.ParseFormatKind(kind)
	if !ok || k == grid.General {
		return o, fmt.Errorf("--as %q: unknown type %q (integer, decimal, currency[:symbol], percent, date, datetime, time, text)", s, kind)
	}
	o.Format = grid.Format{Kind: k, Symbol: sym}
	return o, nil
}

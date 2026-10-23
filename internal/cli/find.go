package cli

import (
	"fmt"
	"regexp"
	"strings"
	"text/tabwriter"

	"github.com/eemax/sheets/internal/analyze"
	"github.com/eemax/sheets/internal/grid"
	"github.com/spf13/cobra"
)

type findResult struct {
	Matches []findMatch `json:"matches"`
	Total   int         `json:"total"`
}

type findMatch struct {
	Ref    string `json:"ref"`
	Header string `json:"header,omitempty"`
	Value  string `json:"value"`
}

func (a *app) findCmd() *cobra.Command {
	var (
		sheet                         string
		useRegex, exact, caseSens, fx bool
		limit                         int
	)
	cmd := &cobra.Command{
		Use:   "find <file> <text>",
		Short: "Find cells by value; prints their addresses and column headers",
		Long: `Search every sheet (or --sheet) for cells whose displayed value contains the
text, ignoring case. --exact matches whole values, --regex takes a Go
regular expression, --formulas searches formula text too.`,
		Args: cobra.ExactArgs(2),
		RunE: func(cmd *cobra.Command, args []string) error {
			wb, err := open(args[0])
			if err != nil {
				return err
			}
			match, err := matcher(args[1], useRegex, exact, caseSens)
			if err != nil {
				return err
			}
			sheets := wb.Sheets
			if sheet != "" {
				s := wb.Sheet(sheet)
				if s == nil {
					return fmt.Errorf("sheet %q not found (sheets: %s)", sheet, strings.Join(wb.SheetNames(), ", "))
				}
				sheets = []*grid.Sheet{s}
			}
			res := findResult{Matches: []findMatch{}}
			for _, s := range sheets {
				regions := analyze.Regions(wb, s)
				s.Each(s.Used(), func(row, col int, raw *grid.Cell) {
					formula := raw.Formula
					c := wb.Value(s, row, col)
					v := grid.Display(c, wb.FormatOf(c), wb.Date1904)
					if !match(v) && !(fx && formula != "" && match("="+formula)) {
						return
					}
					res.Total++
					if limit > 0 && len(res.Matches) >= limit {
						return
					}
					m := findMatch{Ref: grid.FormatRef(s.Name, grid.Range{R1: row, C1: col, R2: row, C2: col}), Value: v}
					if fx && formula != "" {
						m.Value = "=" + formula
					}
					for _, reg := range regions {
						if reg.HeaderRow > 0 && row > reg.HeaderRow && reg.Range.Contains(row, col) {
							m.Header = analyze.ColumnLabel(wb, s, reg, col)
						}
					}
					res.Matches = append(res.Matches, m)
				})
			}
			if a.json {
				return a.printJSON(res)
			}
			if res.Total == 0 {
				a.printf("no matches\n")
				return nil
			}
			tw := tabwriter.NewWriter(a.stdout, 0, 0, 2, ' ', 0)
			for _, m := range res.Matches {
				header := ""
				if m.Header != "" {
					header = "[" + m.Header + "]"
				}
				fmt.Fprintf(tw, "%s\t%s\t%s\n", m.Ref, header, grid.Escape(grid.Truncate(m.Value, 80)))
			}
			tw.Flush()
			if res.Total > len(res.Matches) {
				a.printf("%d of %d matches shown; narrow with --sheet or a more specific pattern, or raise --limit\n", len(res.Matches), res.Total)
			} else {
				a.printf("%s\n", plural(res.Total, "match"))
			}
			return nil
		},
	}
	f := cmd.Flags()
	f.StringVar(&sheet, "sheet", "", "search only this sheet")
	f.BoolVar(&useRegex, "regex", false, "treat text as a regular expression")
	f.BoolVar(&exact, "exact", false, "match whole values only")
	f.BoolVar(&caseSens, "case", false, "case-sensitive")
	f.BoolVar(&fx, "formulas", false, "also search formula text")
	f.IntVarP(&limit, "limit", "n", 50, "maximum matches to print")
	return cmd
}

func matcher(pattern string, useRegex, exact, caseSens bool) (func(string) bool, error) {
	if useRegex {
		if !caseSens {
			pattern = "(?i)" + pattern
		}
		if exact {
			pattern = "^(?:" + pattern + ")$"
		}
		re, err := regexp.Compile(pattern)
		if err != nil {
			return nil, err
		}
		return re.MatchString, nil
	}
	if !caseSens {
		pattern = strings.ToLower(pattern)
	}
	return func(v string) bool {
		if !caseSens {
			v = strings.ToLower(v)
		}
		if exact {
			return v == pattern
		}
		return strings.Contains(v, pattern)
	}, nil
}

package cli

import (
	"fmt"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/eemax/sheets/internal/analyze"
	"github.com/eemax/sheets/internal/grid"
	"github.com/spf13/cobra"
)

type infoResult struct {
	File   string      `json:"file"`
	Format string      `json:"format"`
	Bytes  int64       `json:"bytes"`
	Sheets []sheetInfo `json:"sheets"`
	Names  []nameInfo  `json:"names,omitempty"`
	// Styled counts the entries of recorded styling (sheets style --list).
	Styled int `json:"styled,omitempty"`
}

type sheetInfo struct {
	Name     string `json:"name"`
	Hidden   bool   `json:"hidden,omitempty"`
	Used     string `json:"used"`
	Formulas int    `json:"formulas"`
	Merged   int    `json:"merged_ranges,omitempty"`
	// Hidden rows and columns within the used range, as "5:9" and "C:D".
	HiddenRows []string     `json:"hidden_rows,omitempty"`
	HiddenCols []string     `json:"hidden_cols,omitempty"`
	Charts     []grid.Chart `json:"charts,omitempty"`
	Regions    []regionInfo `json:"regions"`
	More       int          `json:"more_regions,omitempty"`
}

type regionInfo struct {
	Range   string   `json:"range"`
	Kind    string   `json:"kind"` // table, block, cell or title
	Named   string   `json:"excel_table,omitempty"`
	Title   string   `json:"title,omitempty"`
	Header  int      `json:"header_row,omitempty"`
	Group   int      `json:"group_row,omitempty"`
	Rows    int      `json:"rows"`
	Cols    int      `json:"cols"`
	Total   bool     `json:"total_row,omitempty"`
	Columns []string `json:"columns,omitempty"`
	Text    string   `json:"text,omitempty"`
}

type nameInfo struct {
	Name  string `json:"name"`
	Ref   string `json:"ref"`
	Scope string `json:"scope,omitempty"`
}

const maxRegionsListed = 10

func (a *app) infoCmd() *cobra.Command {
	return &cobra.Command{
		Use:   "info <file>",
		Short: "Overview: sheets, used ranges, tables and named ranges",
		Args:  cobra.ExactArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			wb, err := open(args[0])
			if err != nil {
				return err
			}
			res := infoResult{File: args[0], Format: wb.Format, Bytes: wb.Size, Styled: wb.Styled}
			for _, s := range wb.Sheets {
				res.Sheets = append(res.Sheets, describeSheet(wb, s))
			}
			for _, n := range wb.Names {
				res.Names = append(res.Names, nameInfo(n))
			}
			if a.json {
				return a.printJSON(res)
			}
			a.printInfo(res)
			return nil
		},
	}
}

func describeSheet(wb *grid.Workbook, s *grid.Sheet) sheetInfo {
	si := sheetInfo{Name: s.Name, Hidden: s.Hidden}
	used := s.Used()
	if used.Empty() {
		return si
	}
	si.Used = used.String()
	si.Merged = len(s.Merges)
	si.HiddenRows = hiddenRuns(s.HiddenRows, used.R1, used.R2, strconv.Itoa)
	si.HiddenCols = hiddenRuns(s.HiddenCols, used.C1, used.C2, grid.ColName)
	for _, c := range s.Charts {
		si.Charts = append(si.Charts, c.Shown())
	}
	s.Each(used, func(_, _ int, c *grid.Cell) {
		if c.Formula != "" {
			si.Formulas++
		}
	})
	regions := analyze.Regions(wb, s)
	for i, reg := range regions {
		if i == maxRegionsListed {
			si.More = len(regions) - i
			break
		}
		si.Regions = append(si.Regions, describeRegion(wb, s, reg))
	}
	return si
}

func describeRegion(wb *grid.Workbook, s *grid.Sheet, reg analyze.Region) regionInfo {
	ri := regionInfo{Range: reg.Range.String(), Rows: reg.Range.Rows(), Cols: reg.Range.Cols()}
	if t := s.TableAt(reg.Range); t != nil {
		ri.Named = t.Name
	}
	switch {
	case reg.Range.Rows() == 1 && reg.Range.Cols() == 1:
		ri.Kind = "cell"
		if reg.IsTitle {
			ri.Kind = "title"
		}
		c := wb.Value(s, reg.Range.R1, reg.Range.C1)
		ri.Text = grid.Truncate(grid.Display(c, wb.FormatOf(c), wb.Date1904), 60)
	case reg.IsTable():
		ri.Kind = "table"
		ri.Header = reg.HeaderRow
		ri.Group = reg.GroupRow
		ri.Rows = reg.Body.Rows()
		ri.Total = reg.TotalRow > 0
		if reg.TitleRows > 0 {
			s.Each(grid.Range{R1: reg.Range.R1, C1: reg.Range.C1, R2: reg.Range.R1, C2: reg.Range.C2}, func(_, _ int, c *grid.Cell) {
				ri.Title = grid.Truncate(c.Text, 60)
			})
		}
		for col := reg.Range.C1; col <= reg.Range.C2; col++ {
			label := analyze.ColumnLabel(wb, s, reg, col)
			if label == "" {
				label = "(" + grid.ColName(col) + ")"
			}
			ri.Columns = append(ri.Columns, label)
		}
	default:
		ri.Kind = "block"
	}
	return ri
}

func (a *app) printInfo(res infoResult) {
	a.printf("%s · %s · %s · %s\n", filepath.Base(res.File), res.Format, humanBytes(res.Bytes), plural(len(res.Sheets), "sheet"))
	best, bestCells := "", -1
	for _, s := range res.Sheets {
		var meta []string
		if s.Hidden {
			meta = append(meta, "hidden")
		}
		if s.Used == "" {
			meta = append(meta, "empty")
		} else {
			meta = append(meta, s.Used)
		}
		if s.Formulas > 0 {
			meta = append(meta, plural(s.Formulas, "formula"))
		}
		if s.Merged > 0 {
			meta = append(meta, plural(s.Merged, "merged range"))
		}
		if len(s.Charts) > 0 {
			meta = append(meta, plural(len(s.Charts), "chart"))
		}
		if len(s.HiddenRows) > 0 {
			meta = append(meta, "hidden rows "+listRuns(s.HiddenRows))
		}
		if len(s.HiddenCols) > 0 {
			meta = append(meta, "hidden columns "+listRuns(s.HiddenCols))
		}
		a.printf("%s  %s\n", s.Name, strings.Join(meta, " · "))
		for _, r := range s.Regions {
			a.printf("  %s\n", regionLine(r))
			if r.Kind == "table" && !s.Hidden && r.Rows*r.Cols > bestCells {
				best, bestCells = s.Name, r.Rows*r.Cols
			}
		}
		if s.More > 0 {
			a.printf("  … %s\n", plural(s.More, "more region"))
		}
		for _, c := range s.Charts {
			a.printf("  %s  %s\n", c.Anchor, chartLine(c))
		}
		if best == "" && !s.Hidden && s.Used != "" {
			best = s.Name
		}
	}
	if len(res.Names) > 0 {
		var parts []string
		for _, n := range res.Names {
			p := n.Name + "=" + n.Ref
			if n.Scope != "" {
				p += " (" + n.Scope + ")"
			}
			parts = append(parts, p)
		}
		a.printf("names: %s\n", strings.Join(parts, ", "))
	}
	if res.Styled > 0 {
		a.printf("styling: %s recorded, which fmt keeps (sheets style --list)\n", plural(res.Styled, "entry"))
	}
	if best != "" {
		a.hint("sheets peek %s %s", shellQuote(res.File), shellQuote(best))
	}
}

func regionLine(r regionInfo) string {
	switch r.Kind {
	case "cell":
		return fmt.Sprintf("%s  %q", r.Range, r.Text)
	case "title":
		return fmt.Sprintf("%s  title %q", r.Range, r.Text)
	case "table":
		parts := []string{r.Range + "  table"}
		if r.Named != "" {
			parts[0] = r.Range + "  Excel table " + r.Named
		}
		if r.Title != "" {
			parts[0] += fmt.Sprintf(" %q", r.Title)
		}
		rows := plural(r.Rows, "row")
		if r.Total {
			rows += " + total"
		}
		parts = append(parts, rows, headerList(r.Columns, 8))
		return strings.Join(parts, " · ")
	}
	return fmt.Sprintf("%s  block %d×%d", r.Range, r.Rows, r.Cols)
}

func headerList(cols []string, n int) string {
	if len(cols) <= n {
		return strings.Join(cols, ", ")
	}
	return strings.Join(cols[:n], ", ") + fmt.Sprintf(", … (+%d)", len(cols)-n)
}

func humanBytes(n int64) string {
	switch {
	case n >= 1<<20:
		return fmt.Sprintf("%.1f MB", float64(n)/(1<<20))
	case n >= 1<<10:
		return fmt.Sprintf("%d KB", (n+512)>>10)
	}
	return fmt.Sprintf("%d B", n)
}

// hiddenRuns names the hidden runs that fall within lo..hi.
func hiddenRuns(runs []grid.Run, lo, hi int, name func(int) string) []string {
	var out []string
	for _, r := range runs {
		from, to := max(r.From, lo), min(r.To, hi)
		if from > to {
			continue
		}
		if from == to {
			out = append(out, name(from))
		} else {
			out = append(out, name(from)+":"+name(to))
		}
	}
	return out
}

// listRuns shows the first few runs.
func listRuns(runs []string) string {
	if len(runs) > 4 {
		return strings.Join(runs[:4], ", ") + fmt.Sprintf(" and %d more", len(runs)-4)
	}
	return strings.Join(runs, ", ")
}

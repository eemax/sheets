package cli

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strings"

	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/render"
	"github.com/spf13/cobra"
)

type renderResult struct {
	*render.Result
	File string `json:"file"`
}

func (a *app) renderCmd() *cobra.Command {
	var (
		out, format      string
		scale            float64
		noHeaders        bool
		maxRows, maxCols int
	)
	cmd := &cobra.Command{
		Use:   "render <file> [sheet|range]",
		Short: "Draw a sheet as a PNG image (or HTML) to look at its formatting and charts",
		Long: `Draw a sheet, or a range of it, the way a spreadsheet app shows it: values as
formatted in the file, fonts, fills, borders, alignment, wrapped text,
merged cells, banding, and charts. Column letters and row numbers frame it
so what you see maps back to cell addresses.

The PNG is for agents that can look at images (check a format or a chart
before handing a file over); --format html gives a page for people.
Without a range, the used range and its charts are drawn (up to 60 rows and
26 columns).

  sheets render book.xlsx Sales            # writes book-Sales.png next to book.xlsx
  sheets render book.xlsx 'Sales!A1:H30' -o check.png

The drawing approximates Excel: fonts stand in for Calibri, and conditional
formats other than banding and Excel table styles are not drawn (the
output says so).`,
		Args: cobra.RangeArgs(1, 2),
		RunE: func(cmd *cobra.Command, args []string) error {
			if err := xlsxOnly(args[0], "render draws workbooks with their formatting, which a CSV file does not have; read it with sheets read, or convert it"); err != nil {
				return err
			}
			wb, err := open(args[0])
			if err != nil {
				return err
			}
			ref := ""
			if len(args) > 1 {
				ref = args[1]
			}
			s, rg, explicit, err := target(wb, ref)
			if err != nil {
				return err
			}
			opts := render.Options{Sheet: s.Name, Scale: scale, Headers: !noHeaders, MaxRows: maxRows, MaxCols: maxCols}
			if explicit {
				opts.Range = &rg
			}
			if format == "" {
				format = "png"
				if strings.EqualFold(filepath.Ext(out), ".html") {
					format = "html"
				}
			}
			if format != "png" && format != "html" {
				return fmt.Errorf("unknown --format %q (png or html)", format)
			}
			if out == "" {
				base := strings.TrimSuffix(args[0], filepath.Ext(args[0]))
				out = base + "-" + safeName(s.Name) + "." + format
			}
			if scale < 1 || scale > 4 {
				return errors.New("--scale must be between 1 and 4")
			}
			tmp := out + ".tmp"
			fh, err := os.Create(tmp)
			if err != nil {
				return err
			}
			var res *render.Result
			if format == "html" {
				res, err = render.HTML(args[0], fh, opts)
			} else {
				res, err = render.PNG(args[0], fh, opts)
			}
			if cerr := fh.Close(); err == nil {
				err = cerr
			}
			if err != nil {
				os.Remove(tmp)
				return err
			}
			if err := os.Rename(tmp, out); err != nil {
				return err
			}
			if a.json {
				return a.printJSON(renderResult{Result: res, File: out})
			}
			line := fmt.Sprintf("drew %s → %s (%dx%d", grid.QuoteSheet(res.Sheet)+"!"+res.Range, out, res.Width, res.Height)
			if format == "png" {
				line += fmt.Sprintf(" at %gx", scale)
			}
			line += ")"
			if res.Charts > 0 {
				line += ", " + plural(res.Charts, "chart")
			}
			a.printf("%s\n", line)
			if res.Of != "" {
				a.printf("only part of %s: give a range to draw another part\n", res.Of)
			}
			if res.Missing != "" {
				a.printf("not drawn: %s\n", res.Missing)
			}
			return nil
		},
	}
	f := cmd.Flags()
	f.StringVarP(&out, "output", "o", "", "file to write (default: <file>-<sheet>.png next to the workbook)")
	f.StringVar(&format, "format", "", "png or html (default: from the output name, else png)")
	f.Float64Var(&scale, "scale", 2, "pixel density of the PNG")
	f.BoolVar(&noHeaders, "no-headers", false, "leave out column letters and row numbers")
	f.IntVar(&maxRows, "max-rows", 60, "rows to draw at most")
	f.IntVar(&maxCols, "max-cols", 26, "columns to draw at most")
	return cmd
}

var unsafeChars = regexp.MustCompile(`[^A-Za-z0-9._-]+`)

func safeName(s string) string {
	n := strings.Trim(unsafeChars.ReplaceAllString(s, "_"), "_")
	if n == "" {
		return "sheet"
	}
	return n
}

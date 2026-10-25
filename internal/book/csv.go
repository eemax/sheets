package book

import (
	"errors"
	"fmt"
	"io/fs"
	"path/filepath"
	"strings"

	"github.com/eemax/sheets/internal/delim"
	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/xlsx"
)

// A CSV file is a workbook of one sheet named after the file. It is edited
// through an xlsx workbook in memory, written back as the file's own
// dialect when saved; rows no edit changed stay as they were.

func openCSV(path string, data []byte) (*grid.Workbook, error) {
	f, err := delim.Parse(data, path)
	if err != nil {
		return nil, err
	}
	wb := f.Workbook(delim.SheetName(path))
	wb.Path = path
	wb.Size = int64(len(data))
	return wb, nil
}

// errCSV refuses what a CSV file cannot hold.
func errCSV(path, what string) error {
	return fmt.Errorf("%s is a CSV file, which holds values only, not %s; convert it first: sheets convert %s %s.xlsx",
		path, what, path, strings.TrimSuffix(path, filepath.Ext(path)))
}

type csvEditor struct {
	*xlsx.Editor
	path     string
	sheet    string
	orig     *delim.File // nil for a new file
	dialect  delim.Dialect
	created  bool
	warnings []string
}

func editCSV(path string, create bool) (Editor, error) {
	c := &csvEditor{path: path, sheet: delim.SheetName(path), dialect: delim.DefaultDialect(path)}
	data, err := xlsx.ReadFile(path)
	var wb *grid.Workbook
	switch {
	case err == nil:
		if c.orig, err = delim.Parse(data, path); err != nil {
			return nil, err
		}
		c.dialect = c.orig.Dialect
		wb = c.orig.Editing(c.sheet)
	case errors.Is(err, fs.ErrNotExist) && create:
		c.created = true
		wb = &grid.Workbook{Format: "csv", Formats: []grid.Format{{}}, Sheets: []*grid.Sheet{grid.NewSheet(c.sheet)}}
	default:
		return nil, err
	}
	xb, err := xlsx.FromWorkbook(wb)
	if err != nil {
		return nil, err
	}
	if c.Editor, err = xlsx.OpenEditorData(path, xb); err != nil {
		return nil, err
	}
	return c, nil
}

func (c *csvEditor) Created() bool { return c.created }

// HoldsValues is true: a CSV file holds no formulas.
func (c *csvEditor) HoldsValues() bool { return true }

// EnsureSheet accepts the file's one sheet.
func (c *csvEditor) EnsureSheet(name string) (string, bool, error) {
	if strings.EqualFold(name, c.sheet) {
		return c.sheet, false, nil
	}
	return "", false, fmt.Errorf("%s is a CSV file, which holds one sheet, %s: write to %s!A1, or just A1", c.path, grid.QuoteSheet(c.sheet), grid.QuoteSheet(c.sheet))
}

func (c *csvEditor) AddSheet(name, before, after string) error {
	return errCSV(c.path, "more sheets")
}
func (c *csvEditor) RenameSheet(from, to string) error { return errCSV(c.path, "sheet names") }
func (c *csvEditor) DeleteSheet(name string) error     { return errCSV(c.path, "more sheets") }
func (c *csvEditor) CopySheet(from, to string) ([]string, error) {
	return nil, errCSV(c.path, "more sheets")
}
func (c *csvEditor) MoveSheet(name, before, after string, first, last bool) error {
	return errCSV(c.path, "more sheets")
}
func (c *csvEditor) SetVisible(name string, visible bool) error {
	return errCSV(c.path, "hidden sheets")
}
func (c *csvEditor) AddChart(spec ChartSpec) error        { return errCSV(c.path, "charts") }
func (c *csvEditor) DeleteChart(sheet, cell string) error { return errCSV(c.path, "charts") }
func (c *csvEditor) Style(sheet string, r grid.Range, l *Look) (StyleResult, error) {
	return StyleResult{}, errCSV(c.path, "styling")
}
func (c *csvEditor) Cond(sheet string, r grid.Range, rule []string) (string, error) {
	return "", errCSV(c.path, "conditional formats")
}
func (c *csvEditor) Validate(sheet string, r grid.Range, rule ValidationRule) (string, error) {
	return "", errCSV(c.path, "data validation")
}
func (c *csvEditor) View(sheet string, o ViewOptions) error { return errCSV(c.path, "view settings") }
func (c *csvEditor) SetNote(sheet string, row, col int, text, author string) error {
	return errCSV(c.path, "notes")
}
func (c *csvEditor) SetLink(sheet string, row, col int, l *grid.Link, text string) error {
	return errCSV(c.path, "hyperlinks")
}

// Save writes the sheet's values in the file's dialect. A formula is
// saved as its result: a CSV file holds no formulas.
func (c *csvEditor) Save(dryRun bool) error {
	xb, err := c.Editor.Render()
	if err != nil {
		return err
	}
	wb, err := xlsx.ReadBytes(xb)
	if err != nil {
		return err
	}
	s := wb.Sheet(c.sheet)
	if s == nil {
		return fmt.Errorf("internal error: the sheet of %s went missing", c.path)
	}
	used := s.Used()
	rows, cols := 0, 0
	if !used.Empty() {
		rows, cols = used.R2, used.C2
	}
	var formulas, unknown []string
	value := func(r, col int) (grid.Cell, grid.Format, error) {
		cell := s.At(r, col)
		if cell == nil || cell.IsEmpty() {
			return grid.Cell{}, grid.Format{}, nil
		}
		if cell.Formula != "" {
			formulas = append(formulas, grid.CellName(r, col))
			if v := wb.Value(s, r, col); v.Pending {
				unknown = append(unknown, grid.CellName(r, col)+" (="+cell.Formula+")")
				return grid.Cell{}, grid.Format{}, nil
			}
		}
		return *cell, wb.FormatOf(cell), nil
	}
	out, err := delim.Encode(rows, cols, value, c.dialect, c.orig, wb.Date1904)
	if err != nil {
		return fmt.Errorf("%s: %w", c.path, err)
	}
	if len(unknown) > 0 {
		list := unknown
		if len(list) > 5 {
			list = append(list[:5:5], fmt.Sprintf("and %d more", len(unknown)-5))
		}
		return fmt.Errorf("%s is a CSV file, which holds values, and the sheets calculator cannot compute %s; nothing was saved (write the values, or convert to .xlsx)", c.path, strings.Join(list, ", "))
	}
	c.warnings = nil
	if len(formulas) > 0 {
		list := formulas
		if len(list) > 5 {
			list = append(list[:5:5], fmt.Sprintf("and %d more", len(formulas)-5))
		}
		c.warnings = append(c.warnings, fmt.Sprintf("%s is a CSV file, which holds values: %s saved as %s, not as %s (convert the file to .xlsx to keep formulas)",
			c.path, strings.Join(list, ", "), map[bool]string{true: "its result", false: "their results"}[len(formulas) == 1], map[bool]string{true: "a formula", false: "formulas"}[len(formulas) == 1]))
	}
	if dryRun {
		return nil
	}
	return xlsx.WriteFile(c.path, out)
}

func (c *csvEditor) Warnings() []string {
	return append(append([]string{}, c.Editor.Warnings()...), c.warnings...)
}

package xlsx

import (
	"fmt"

	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/style"
	"github.com/xuri/excelize/v2"
)

// ReadFile reads a file sheets works on: from disk, or from the batch
// holding it (see files.go).
func ReadFile(path string) ([]byte, error) { return readFile(path) }

// WriteFile replaces a file atomically, keeping its undo history (or
// hands it to the batch holding it).
func WriteFile(path string, data []byte) error { return writeAtomic(path, data) }

// FromWorkbook writes the values of a workbook (sheets, cells and number
// formats; no formulas, styles or anything else) as an xlsx package. It
// streams the rows, so large tables convert quickly.
func FromWorkbook(wb *grid.Workbook) ([]byte, error) {
	f := excelize.NewFile()
	defer f.Close()
	if len(wb.Sheets) == 0 {
		return nil, fmt.Errorf("the workbook has no sheets")
	}
	styles := map[uint16]int{}
	styleOf := func(c *grid.Cell) (int, error) {
		if c.Kind != grid.Number || c.Fmt == 0 {
			return 0, nil
		}
		if id, ok := styles[c.Fmt]; ok {
			return id, nil
		}
		fm := wb.FormatOf(c)
		code := fm.Code
		if code == "" {
			code = style.NumFmt(fm, false)
		}
		if code == "" {
			styles[c.Fmt] = 0
			return 0, nil
		}
		id, err := f.NewStyle(&excelize.Style{CustomNumFmt: &code})
		if err != nil {
			return 0, err
		}
		styles[c.Fmt] = id
		return id, nil
	}
	for i, s := range wb.Sheets {
		if i == 0 {
			if err := f.SetSheetName(f.GetSheetName(0), s.Name); err != nil {
				return nil, fmt.Errorf("sheet name %q: %w", s.Name, err)
			}
		} else if _, err := f.NewSheet(s.Name); err != nil {
			return nil, fmt.Errorf("sheet name %q: %w", s.Name, err)
		}
		sw, err := f.NewStreamWriter(s.Name)
		if err != nil {
			return nil, err
		}
		used := s.Used()
		for r := 1; r <= used.R2 && !used.Empty(); r++ {
			n := s.RowLen(r)
			if n == 0 {
				continue
			}
			row := make([]interface{}, n)
			for c := 1; c <= n; c++ {
				cell := s.At(r, c)
				if cell == nil || cell.IsEmpty() {
					continue
				}
				id, err := styleOf(cell)
				if err != nil {
					return nil, err
				}
				var v interface{}
				switch cell.Kind {
				case grid.Number:
					v = cell.Num
				case grid.Bool:
					v = cell.Bool
				default:
					v = cell.Text
				}
				row[c-1] = excelize.Cell{StyleID: id, Value: v}
			}
			if err := sw.SetRow(grid.CellName(r, 1), row); err != nil {
				return nil, err
			}
		}
		if err := sw.Flush(); err != nil {
			return nil, err
		}
	}
	if wb.Date1904 {
		on := true
		if err := f.SetWorkbookProps(&excelize.WorkbookPropsOptions{Date1904: &on}); err != nil {
			return nil, err
		}
	}
	buf, err := f.WriteToBuffer()
	if err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

// OpenEditorData opens a workbook held in memory for editing; path is
// where Save writes it.
func OpenEditorData(path string, data []byte) (*Editor, error) {
	e := &Editor{path: path, styles: map[string]int{}, unshare: map[string]bool{}, extent: map[string]grid.Range{}, patches: newPatches(), hasDrawings: map[string]bool{}}
	var err error
	if data, err = normalizeNotes(data); err != nil {
		return nil, err
	}
	e.data = data
	if e.wb, e.layouts, _, err = parse(data); err != nil {
		return nil, fmt.Errorf("%s: %w", path, err)
	}
	if e.f, err = e.openExcelize(data); err != nil {
		return nil, fmt.Errorf("%s: %w", path, err)
	}
	for name, l := range e.layouts {
		e.hasDrawings[name] = l.drawing
		for _, t := range l.tables {
			e.patches.tables = append(e.patches.tables, tableSource{name, t})
		}
	}
	return e, nil
}

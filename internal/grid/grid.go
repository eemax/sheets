// Package grid is the format-neutral model of a workbook. Readers for a file
// format (xlsx today) produce a Workbook; everything that analyses or
// displays data works on this model only.
package grid

import "strings"

// Kind is the type of a cell's value.
type Kind uint8

const (
	Empty Kind = iota
	Text
	Number
	Bool
	Error
)

func (k Kind) String() string {
	return [...]string{"empty", "text", "number", "bool", "error"}[k]
}

// Cell is a single cell. Dates and times are numbers carrying a date Format,
// exactly as spreadsheets store them. Fields are ordered to keep the struct
// small: sheets hold millions of these.
type Cell struct {
	Kind Kind
	Bool bool
	// Pending marks a formula whose value is unknown: the file holds no
	// cached result, or the cached result is stale. See Workbook.Value.
	Pending bool
	// Computed marks a formula value calculated by sheets rather than read
	// from the file (or an attempt to, when Pending is still set).
	Computed bool
	// Shared marks a formula stored as part of a filled-down group.
	Shared  bool
	Fmt     uint16  // index into Workbook.Formats
	Num     float64 // Number value
	Text    string  // Text value, or the error code for Error
	Formula string  // without the leading '='; empty for constants
}

// IsEmpty reports whether the cell has neither value nor formula.
func (c *Cell) IsEmpty() bool { return c.Kind == Empty && c.Formula == "" }

// Sheet is one worksheet. Cells are stored in ragged rows: rows[r-1] holds
// columns 1..len(rows[r-1]).
type Sheet struct {
	Name   string
	Hidden bool
	Merges []Range
	// Tables are named tables defined in the file (Excel "Format as
	// Table"), which carry their own style and filter.
	Tables []NamedTable
	rows   [][]Cell
	maxCol int
}

// NamedTable is a table object stored in the file.
type NamedTable struct {
	Name  string
	Range Range
}

// TableAt returns the named table overlapping r, if any.
func (s *Sheet) TableAt(r Range) *NamedTable {
	for i := range s.Tables {
		if !s.Tables[i].Range.Intersect(r).Empty() {
			return &s.Tables[i]
		}
	}
	return nil
}

// NewSheet returns an empty sheet.
func NewSheet(name string) *Sheet { return &Sheet{Name: name} }

// Set stores a cell, growing the sheet as needed.
func (s *Sheet) Set(row, col int, c Cell) {
	for len(s.rows) < row {
		s.rows = append(s.rows, nil)
	}
	r := s.rows[row-1]
	if len(r) < col {
		if cap(r) >= col {
			r = r[:col]
		} else {
			nr := make([]Cell, col, max(col, 2*cap(r)))
			copy(nr, r)
			r = nr
		}
		s.rows[row-1] = r
	}
	r[col-1] = c
	s.maxCol = max(s.maxCol, col)
}

// At returns a pointer to the cell, or nil when it lies outside the stored
// area (which means it is empty).
func (s *Sheet) At(row, col int) *Cell {
	if row < 1 || row > len(s.rows) || col < 1 {
		return nil
	}
	r := s.rows[row-1]
	if col > len(r) {
		return nil
	}
	return &r[col-1]
}

// Get returns the cell value, the zero Cell when empty.
func (s *Sheet) Get(row, col int) Cell {
	if c := s.At(row, col); c != nil {
		return *c
	}
	return Cell{}
}

// RowLen returns the number of stored columns in a row.
func (s *Sheet) RowLen(row int) int {
	if row < 1 || row > len(s.rows) {
		return 0
	}
	return len(s.rows[row-1])
}

// Used returns the bounding box of all non-empty cells.
func (s *Sheet) Used() Range {
	used := Range{R1: 1, C1: 1}
	minCol := 0
	for r := range s.rows {
		for c := range s.rows[r] {
			if s.rows[r][c].IsEmpty() {
				continue
			}
			if used.R2 == 0 {
				used.R1 = r + 1
			}
			used.R2 = r + 1
			used.C2 = max(used.C2, c+1)
			if minCol == 0 || c+1 < minCol {
				minCol = c + 1
			}
		}
	}
	if used.R2 == 0 {
		return Range{1, 1, 0, 0}
	}
	used.C1 = minCol
	return used
}

// Each calls fn for every non-empty cell inside r in row-major order.
func (s *Sheet) Each(r Range, fn func(row, col int, c *Cell)) {
	for row := max(r.R1, 1); row <= min(r.R2, len(s.rows)); row++ {
		cells := s.rows[row-1]
		for col := max(r.C1, 1); col <= min(r.C2, len(cells)); col++ {
			if c := &cells[col-1]; !c.IsEmpty() {
				fn(row, col, c)
			}
		}
	}
}

// DefinedName is a named range or constant.
type DefinedName struct {
	Name  string
	Ref   string
	Scope string // sheet name for sheet-local names, empty for global
}

// Workbook is a whole file.
type Workbook struct {
	Path     string
	Format   string // file format, e.g. "xlsx"
	Size     int64  // bytes on disk
	Sheets   []*Sheet
	Names    []DefinedName
	Formats  []Format // indexed by Cell.Fmt; Formats[0] is General
	Date1904 bool
	// Resolver computes values of Pending formula cells. Nil when the
	// format cannot evaluate formulas.
	Resolver Resolver
}

// Resolver evaluates formulas.
type Resolver interface {
	Resolve(sheet string, row, col int) (Cell, error)
}

// Sheet finds a sheet by name, case-insensitively.
func (w *Workbook) Sheet(name string) *Sheet {
	for _, s := range w.Sheets {
		if s.Name == name {
			return s
		}
	}
	for _, s := range w.Sheets {
		if strings.EqualFold(s.Name, name) {
			return s
		}
	}
	return nil
}

// SheetNames lists sheet names in workbook order.
func (w *Workbook) SheetNames() []string {
	names := make([]string, len(w.Sheets))
	for i, s := range w.Sheets {
		names[i] = s.Name
	}
	return names
}

// FirstVisible returns the first sheet that is not hidden.
func (w *Workbook) FirstVisible() *Sheet {
	for _, s := range w.Sheets {
		if !s.Hidden {
			return s
		}
	}
	if len(w.Sheets) > 0 {
		return w.Sheets[0]
	}
	return nil
}

// FormatOf returns the number format of a cell.
func (w *Workbook) FormatOf(c *Cell) Format {
	if int(c.Fmt) < len(w.Formats) {
		return w.Formats[c.Fmt]
	}
	return Format{}
}

// Value returns the cell with its formula value resolved when needed. The
// cell in the sheet is updated so each formula is evaluated at most once.
// A formula the resolver cannot evaluate stays Pending (shown as "?").
func (w *Workbook) Value(s *Sheet, row, col int) *Cell {
	c := s.At(row, col)
	if c == nil || !c.Pending || c.Computed || w.Resolver == nil {
		return c
	}
	c.Computed = true // attempted, whatever the outcome
	v, err := w.Resolver.Resolve(s.Name, row, col)
	if err != nil {
		return c
	}
	c.Pending = false
	c.Kind, c.Text, c.Num, c.Bool = v.Kind, v.Text, v.Num, v.Bool
	return c
}

// ResolveRange resolves all pending formulas in a range.
func (w *Workbook) ResolveRange(s *Sheet, r Range) {
	if w.Resolver == nil {
		return
	}
	s.Each(r, func(row, col int, c *Cell) {
		if c.Pending {
			w.Value(s, row, col)
		}
	})
}

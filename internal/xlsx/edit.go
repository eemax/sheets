package xlsx

import (
	"bytes"
	"errors"
	"fmt"
	"io/fs"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"strings"

	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/input"
	"github.com/eemax/sheets/internal/style"
	"github.com/xuri/excelize/v2"
)

// Editor changes cell contents. It keeps a parsed copy of the workbook so
// callers can look at what is there before writing.
type Editor struct {
	path     string
	f        *excelize.File
	wb       *grid.Workbook // contents as of opening
	created  bool
	styles   map[string]int // existing style + number format → style id
	unshare  map[string]bool
	extent   map[string]grid.Range // used range per touched sheet
	layouts  map[string]*layout
	patches  *patches
	warnings []string
	data     []byte // the file as opened; nil for a new file
	// structural is set once rows, columns or sheets were inserted,
	// deleted or renamed.
	structural  bool
	hasDrawings map[string]bool

	// What this session changed, for recomputing stored results (see
	// recalc.go): cells written, formulas an edit of rows, columns or
	// sheets may have changed, and, after such an edit of rows or
	// columns, every formula where it now stands.
	changed       map[cellKey]string
	recompute     map[cellKey]bool
	lines         []formulaAt
	sheetsChanged bool
	// dropped notes, after saving, why a formula got no stored result,
	// or what its error means.
	dropped map[cellKey]string
	st      *styler // for targeted styling (see styling.go)
}

// OpenEditor opens path for editing, or starts a new workbook when the file
// does not exist and create is set.
func OpenEditor(path string, create bool) (*Editor, error) {
	e := &Editor{path: path, styles: map[string]int{}, unshare: map[string]bool{}, extent: map[string]grid.Range{}, patches: newPatches(), hasDrawings: map[string]bool{}}
	data, err := readFile(path)
	switch {
	case err == nil:
		e.data = data
		if e.f, err = excelize.OpenReader(bytes.NewReader(data)); err != nil {
			return nil, fmt.Errorf("%s: %w", path, err)
		}
		if e.wb, e.layouts, _, err = parse(data); err != nil {
			return nil, fmt.Errorf("%s: %w", path, err)
		}
		for name, l := range e.layouts {
			e.hasDrawings[name] = l.drawing
			for _, t := range l.tables {
				e.patches.tables = append(e.patches.tables, tableSource{name, t})
			}
		}
	case errors.Is(err, fs.ErrNotExist) && create:
		e.f = excelize.NewFile()
		e.wb = &grid.Workbook{Format: "xlsx", Formats: []grid.Format{{}}}
		e.created = true
	default:
		return nil, err
	}
	return e, nil
}

// Created reports whether the editor started a new file.
func (e *Editor) Created() bool { return e.created }

// Workbook returns the contents as they were when the file was opened.
func (e *Editor) Workbook() *grid.Workbook { return e.wb }

// EnsureSheet returns the actual name of the sheet, creating it if needed.
// A brand-new file's placeholder sheet is renamed instead of kept.
func (e *Editor) EnsureSheet(name string) (string, bool, error) {
	if s := e.wb.Sheet(name); s != nil {
		return s.Name, false, nil
	}
	if e.created && len(e.wb.Sheets) == 0 {
		first := e.f.GetSheetName(0)
		if err := e.f.SetSheetName(first, name); err != nil {
			return "", false, fmt.Errorf("sheet name %q: %w", name, err)
		}
	} else if _, err := e.f.NewSheet(name); err != nil {
		return "", false, fmt.Errorf("sheet name %q: %w", name, err)
	}
	e.wb.Sheets = append(e.wb.Sheets, grid.NewSheet(name))
	return name, true, nil
}

// Set writes one value. Typed values (dates, percentages, currency) get the
// house number format while the rest of the cell's style is kept.
func (e *Editor) Set(sheet string, row, col int, v input.Value) error {
	if v.Formula != "" {
		if err := formula.Check(v.Formula); err != nil {
			return fmt.Errorf("%s: =%s is not a valid formula: %v", grid.FormatRef(sheet, grid.Range{R1: row, C1: col, R2: row, C2: col}), strings.TrimPrefix(v.Formula, "="), err)
		}
	}
	if err := e.unshareFormulas(sheet, row, col); err != nil {
		return err
	}
	cell := grid.CellName(row, col)
	if e.wb.Date1904 && (v.Fmt.Kind == grid.Date || v.Fmt.Kind == grid.DateTime) {
		v.Num -= 1462 // days between the 1900 and 1904 epochs
	}
	if err := e.clearSpill(sheet, row, col); err != nil {
		return err
	}
	e.warnIfInSpill(sheet, row, col, v)
	var err error
	switch {
	case v.Formula != "":
		err = e.setFormula(sheet, cell, v.Formula)
	case v.Kind == grid.Empty:
		err = e.f.SetCellDefault(sheet, cell, "")
	case v.Kind == grid.Text:
		err = e.f.SetCellStr(sheet, cell, v.Text)
	case v.Kind == grid.Bool:
		err = e.f.SetCellBool(sheet, cell, v.Bool)
	case v.Kind == grid.Number:
		err = e.f.SetCellFloat(sheet, cell, v.Num, -1, 64)
	}
	if err != nil {
		return fmt.Errorf("%s: %w", grid.FormatRef(sheet, grid.Range{R1: row, C1: col, R2: row, C2: col}), err)
	}
	e.touch(sheet, row, col)
	if !v.IsEmpty() {
		e.grow(sheet, grid.Range{R1: row, C1: col, R2: row, C2: col})
	}
	if v.Fmt.Kind == grid.General {
		return nil
	}
	return e.applyNumFmt(sheet, cell, style.NumFmt(v.Fmt, false))
}

// setFormula stores a formula as Excel expects it: newer functions get their
// _xlfn. prefix, and formulas returning arrays become dynamic array
// formulas that spill.
func (e *Editor) setFormula(sheet, cell, text string) error {
	stored := formula.ToFile(text)
	// Clear first: excelize keeps a stale cached value otherwise.
	if err := e.f.SetCellDefault(sheet, cell, ""); err != nil {
		return err
	}
	dynamic := formula.Spills(stored)
	var opts []excelize.FormulaOpts
	if dynamic {
		t, ref := excelize.STCellFormulaTypeArray, cell
		opts = append(opts, excelize.FormulaOpts{Type: &t, Ref: &ref})
	}
	if err := e.f.SetCellFormula(sheet, cell, stored, opts...); err != nil {
		return err
	}
	if dynamic || e.wasDynamic(sheet, cell) {
		e.patches.markCell(sheet, cell, dynamic)
	}
	return nil
}

func (e *Editor) wasDynamic(sheet, cell string) bool {
	if e.patches.dynamic[sheet][cell] {
		return true
	}
	row, col, _ := grid.ParseCell(cell)
	if lay := e.layouts[sheet]; lay != nil {
		_, ok := lay.spills[[2]int{row, col}]
		return ok
	}
	return false
}

// clearSpill empties the old result of a dynamic array formula before its
// anchor cell is overwritten; otherwise the spilled values would stay
// behind as constants.
func (e *Editor) clearSpill(sheet string, row, col int) error {
	lay := e.layouts[sheet]
	if lay == nil {
		return nil
	}
	spill, ok := lay.spills[[2]int{row, col}]
	if !ok {
		return nil
	}
	delete(lay.spills, [2]int{row, col})
	e.patches.markCell(sheet, grid.CellName(row, col), false)
	s := e.wb.Sheet(sheet)
	var err error
	s.Each(spill, func(r, c int, cell *grid.Cell) {
		if err == nil && (r != row || c != col) && cell.Formula == "" {
			err = e.f.SetCellDefault(sheet, grid.CellName(r, c), "")
			e.touch(sheet, r, c)
		}
	})
	return err
}

// warnIfInSpill notes a write into the result area of a dynamic array,
// which Excel answers with #SPILL!.
func (e *Editor) warnIfInSpill(sheet string, row, col int, v input.Value) {
	lay := e.layouts[sheet]
	if lay == nil || v.IsEmpty() {
		return
	}
	for anchor, spill := range lay.spills {
		if spill.Contains(row, col) && (anchor[0] != row || anchor[1] != col) {
			e.warnings = append(e.warnings, fmt.Sprintf("%s is inside the spill range %s of the dynamic array formula in %s; Excel will show #SPILL! there",
				grid.FormatRef(sheet, grid.Range{R1: row, C1: col, R2: row, C2: col}), spill, grid.CellName(anchor[0], anchor[1])))
			return
		}
	}
}

// Warnings lists problems noticed while editing that did not stop it.
func (e *Editor) Warnings() []string { return e.warnings }

// grow extends the recorded used range of a sheet.
func (e *Editor) grow(sheet string, r grid.Range) {
	ext, ok := e.extent[sheet]
	if !ok {
		ext = grid.Range{R1: 1, C1: 1}
		if s := e.wb.Sheet(sheet); s != nil {
			ext = s.Used()
		}
	}
	e.extent[sheet] = ext.Union(r)
}

// applyNumFmt swaps the number format of a cell, keeping its other style.
func (e *Editor) applyNumFmt(sheet, cell, code string) error {
	cur, err := e.f.GetCellStyle(sheet, cell)
	if err != nil {
		return err
	}
	key := strconv.Itoa(cur) + "|" + code
	id, ok := e.styles[key]
	if !ok {
		st, err := e.f.GetStyle(cur)
		if err != nil {
			return err
		}
		st.NumFmt = 0
		st.DecimalPlaces = nil
		st.CustomNumFmt = &code
		normalizeFill(st)
		if id, err = e.f.NewStyle(st); err != nil {
			return fmt.Errorf("%s: keeping its style while setting the number format: %w", cell, err)
		}
		e.styles[key] = id
	}
	return e.f.SetCellStyle(sheet, cell, cell, id)
}

// unshareFormulas rewrites a sheet's filled-down (shared) formulas as plain
// formulas before the first edit that touches one. excelize drops every
// formula of a shared group when its first cell is overwritten.
func (e *Editor) unshareFormulas(sheet string, row, col int) error {
	s := e.wb.Sheet(sheet)
	if s == nil || e.unshare[s.Name] {
		return nil
	}
	if c := s.At(row, col); c == nil || !c.Shared {
		return nil
	}
	return e.unshareSheet(s.Name)
}

func (e *Editor) unshareSheet(sheet string) error {
	s := e.wb.Sheet(sheet)
	if s == nil || e.unshare[s.Name] {
		return nil
	}
	e.unshare[s.Name] = true
	type fc struct {
		row, col int
		f        string
	}
	var cells []fc
	s.Each(s.Used(), func(r, c int, cell *grid.Cell) {
		if cell.Shared {
			cells = append(cells, fc{r, c, cell.Formula})
		}
	})
	// Children first: a group's anchor is its top-left cell, and rewriting
	// it first would erase the children.
	sort.Slice(cells, func(i, j int) bool {
		if cells[i].row != cells[j].row {
			return cells[i].row > cells[j].row
		}
		return cells[i].col > cells[j].col
	})
	for _, c := range cells {
		if err := e.f.SetCellFormula(sheet, grid.CellName(c.row, c.col), formula.ToFile(c.f)); err != nil {
			return err
		}
	}
	return nil
}

// Clear empties every cell of a sheet's used range, including styles and
// merges.
func (e *Editor) Clear(sheet string) error {
	s := e.wb.Sheet(sheet)
	if s == nil {
		return nil
	}
	if ts := e.tablesOn(sheet); len(ts) > 0 {
		return fmt.Errorf("%s holds Excel table %s, which clearing would leave without its header; write into it, or use another sheet", grid.QuoteSheet(sheet), ts[0])
	}
	used := s.Used()
	if used.Empty() {
		return nil
	}
	var err error
	s.Each(used, func(r, c int, _ *grid.Cell) {
		if err == nil {
			err = e.f.SetCellDefault(sheet, grid.CellName(r, c), "")
			e.touch(sheet, r, c)
		}
	})
	if err != nil {
		return err
	}
	for _, m := range s.Merges {
		if err := e.f.UnmergeCell(sheet, grid.CellName(m.R1, m.C1), grid.CellName(m.R2, m.C2)); err != nil {
			return err
		}
	}
	e.extent[sheet] = grid.Range{R1: 1, C1: 1}
	if lay := e.layouts[sheet]; lay != nil {
		for anchor := range lay.spills {
			e.patches.markCell(sheet, grid.CellName(anchor[0], anchor[1]), false)
		}
		lay.spills = map[[2]int]grid.Range{}
		e.patches.plainCols[sheet] = true
		return stripStyles(e.f, sheet, lay)
	}
	return nil
}

// Evaluate calculates the current value of a cell after edits. Formula
// errors are Error cells. ok is false when the calculator cannot evaluate
// the formula; note then says why.
//
// After saving, it is the result stored in the file, or why there is none.
func (e *Editor) Evaluate(sheet string, row, col int) (v grid.Cell, note string, ok bool) {
	if v, ok := e.patches.results[strings.ToLower(sheet)][grid.CellName(row, col)]; ok {
		note := e.dropped[keyOf(sheet, row, col)]
		if v.Kind == grid.Empty && !strings.HasPrefix(note, "circular reference") && !strings.Contains(note, "Excel computes it") {
			note += "; Excel computes it on open"
		}
		return v, note, v.Kind != grid.Empty
	}
	v, note, err := evaluate(e.f, sheet, grid.CellName(row, col))
	return v, note, err == nil || errors.Is(err, errUncertain)
}

// Save writes the workbook atomically, or with dryRun only works out what
// would change (warnings included). Excel and LibreOffice are told to
// recalculate on open, since cached formula results may now be stale.
func (e *Editor) Save(dryRun bool) error {
	on := true
	if err := e.f.SetCalcProps(&excelize.CalcPropsOptions{FullCalcOnLoad: &on}); err != nil {
		return err
	}
	// excelize does not maintain <dimension>; some readers trust it.
	for sheet, ext := range e.extent {
		if ext.Empty() {
			ext = grid.Range{R1: 1, C1: 1, R2: 1, C2: 1}
		}
		if err := e.f.SetSheetDimension(sheet, ext.String()); err != nil {
			return err
		}
	}
	if err := e.restyleWritten(); err != nil {
		return err
	}
	if err := e.fixTables(); err != nil {
		return err
	}
	if err := e.recalc(); err != nil {
		return err
	}
	e.patches.stamp = true
	buf, err := serialize(e.f, e.patches)
	if err != nil {
		return err
	}
	e.warnings = append(e.warnings, e.patches.renameNotes()...)
	if e.structural && e.data != nil {
		before, err1 := ReadBytes(e.data)
		after, err2 := ReadBytes(buf)
		if err1 == nil && err2 == nil {
			if refs := newRefErrors(before, after); len(refs) > 0 {
				e.warnings = append(e.warnings, fmt.Sprintf("%s now %s #REF! (a reference pointed into what was removed): %s",
					countOf(len(refs), "formula"), map[bool]string{true: "contains", false: "contain"}[len(refs) == 1], strings.Join(refs[:min(len(refs), 10)], ", ")))
			}
		}
	}
	if dryRun {
		return nil
	}
	return writeAtomic(e.path, buf)
}

func countOf(n int, word string) string {
	if n == 1 {
		return "1 " + word
	}
	return fmt.Sprintf("%d %ss", n, word)
}

// Journal, when set, is told about every file sheets replaces, with its
// previous contents (nil when it did not exist). The CLI keeps an undo
// history with it.
var Journal func(path string, before, after []byte)

func save(f *excelize.File, path string, pt *patches) error {
	buf, err := serialize(f, pt)
	if err != nil {
		return err
	}
	return writeAtomic(path, buf)
}

func writeAtomic(path string, buf []byte) error {
	if b := held(path); b != nil {
		b.data, b.changed = buf, true
		return nil
	}
	dir := filepath.Dir(path)
	if _, err := os.Stat(dir); err != nil {
		return fmt.Errorf("cannot write %s: %w", path, err)
	}
	before, err := os.ReadFile(path)
	if err != nil && !errors.Is(err, fs.ErrNotExist) {
		return err
	}
	tmp, err := os.CreateTemp(dir, ".sheets-*.xlsx")
	if err != nil {
		return fmt.Errorf("cannot write %s: %w", path, err)
	}
	defer os.Remove(tmp.Name())
	if _, err := tmp.Write(buf); err != nil {
		tmp.Close()
		return err
	}
	if err := tmp.Close(); err != nil {
		return err
	}
	mode := fs.FileMode(0o644)
	if st, err := os.Stat(path); err == nil {
		mode = st.Mode().Perm()
	}
	if err := os.Chmod(tmp.Name(), mode); err != nil {
		return err
	}
	if err := os.Rename(tmp.Name(), path); err != nil {
		return err
	}
	if Journal != nil {
		Journal(path, before, buf)
	}
	return nil
}

// Close releases resources.
func (e *Editor) Close() error { return e.f.Close() }

package xlsx

import (
	"fmt"

	"github.com/xuri/excelize/v2"
	"sort"
	"strings"

	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
)

// Structural edits. excelize shifts formulas (on every sheet), names,
// merges, conditional formats, validations, filters and tables when rows
// or columns move; renames and deletions of sheets are completed here.

// InsertRows inserts n empty rows before row at.
func (e *Editor) InsertRows(sheet string, at, n int) error {
	err := e.relines(sheet, false,
		func(f, home string) string { return formula.InsertLines(f, home, sheet, false, at, n) },
		func(r, c int) (int, int, bool) {
			if r >= at {
				r += n
			}
			return r, c, true
		},
		func() error { return e.f.InsertRows(sheet, at, n) })
	if err != nil {
		return err
	}
	e.shiftExtent(sheet, func(r grid.Range) grid.Range {
		if r.R2 >= at {
			r.R2 += n
			if r.R1 >= at {
				r.R1 += n
			}
		}
		return r
	})
	return nil
}

// DeleteRows deletes rows from..to (inclusive).
func (e *Editor) DeleteRows(sheet string, from, to int) error {
	if err := e.deleteLines(sheet, false, from, to); err != nil {
		return err
	}
	n := to - from + 1
	e.shiftExtent(sheet, func(r grid.Range) grid.Range {
		switch {
		case r.R1 > to:
			r.R1 -= n
			r.R2 -= n
		case r.R2 >= from:
			r.R2 = max(from-1, r.R2-n)
		}
		return r
	})
	return nil
}

// InsertCols inserts n empty columns before column at.
func (e *Editor) InsertCols(sheet string, at, n int) error {
	err := e.relines(sheet, true,
		func(f, home string) string { return formula.InsertLines(f, home, sheet, true, at, n) },
		func(r, c int) (int, int, bool) {
			if c >= at {
				c += n
			}
			return r, c, true
		},
		func() error { return e.f.InsertCols(sheet, grid.ColName(at), n) })
	if err != nil {
		return err
	}
	e.shiftExtent(sheet, func(r grid.Range) grid.Range {
		if r.C2 >= at {
			r.C2 += n
			if r.C1 >= at {
				r.C1 += n
			}
		}
		return r
	})
	return nil
}

// DeleteCols deletes columns from..to (inclusive).
func (e *Editor) DeleteCols(sheet string, from, to int) error {
	if err := e.deleteLines(sheet, true, from, to); err != nil {
		return err
	}
	n := to - from + 1
	e.shiftExtent(sheet, func(r grid.Range) grid.Range {
		switch {
		case r.C1 > to:
			r.C1 -= n
			r.C2 -= n
		case r.C2 >= from:
			r.C2 = max(from-1, r.C2-n)
		}
		return r
	})
	return nil
}

// deleteLines removes rows or columns from..to.
func (e *Editor) deleteLines(target string, cols bool, from, to int) error {
	n := to - from + 1
	return e.relines(target, cols,
		func(f, home string) string { return formula.DeleteLines(f, home, target, cols, from, to) },
		func(r, c int) (int, int, bool) {
			line := r
			if cols {
				line = c
			}
			switch {
			case line >= from && line <= to:
				return 0, 0, false // deleted with its line
			case line > to && cols:
				c -= n
			case line > to:
				r -= n
			}
			return r, c, true
		},
		func() error {
			for i := to; i >= from; i-- {
				var err error
				if cols {
					err = e.f.RemoveCol(target, grid.ColName(i))
				} else {
					err = e.f.RemoveRow(target, i)
				}
				if err != nil {
					return err
				}
			}
			return nil
		})
}

// relines inserts or deletes rows or columns of target through op, keeping
// formulas right. excelize's own formula adjustment points references into
// deleted lines at their neighbours instead of #REF!, fails on LET and
// LAMBDA names, and re-parses every formula once per deleted line (minutes
// for a block of rows in a big sheet). So formulas are taken off first,
// rewritten in one pass (rewrite gets the formula and the sheet holding it),
// and put back where move says their cell went once op has moved cells,
// names, merges, conditional formats, validations, filters and tables.
func (e *Editor) relines(target string, cols bool, rewrite func(f, home string) string, move func(r, c int) (int, int, bool), op func() error) error {
	type moved struct {
		sheet    string
		row, col int
		f        string
	}
	var keep []moved
	for _, s := range e.wb.Sheets {
		var cells [][2]int
		s.Each(s.Used(), func(r, c int, cell *grid.Cell) {
			if cell.Formula == "" {
				return
			}
			cells = append(cells, [2]int{r, c})
			m := moved{sheet: s.Name, row: r, col: c, f: rewrite(cell.Formula, s.Name)}
			if s.Name == target {
				var ok bool
				if m.row, m.col, ok = move(r, c); !ok {
					return
				}
			}
			keep = append(keep, m)
		})
		// Bottom-right first, so a shared formula's anchor goes last.
		for i := len(cells) - 1; i >= 0; i-- {
			if err := e.f.SetCellFormula(s.Name, grid.CellName(cells[i][0], cells[i][1]), ""); err != nil {
				return err
			}
		}
	}
	// Names: excelize shifts them, but cannot parse LET names either; and
	// it would point names into deleted lines at their neighbours. Built-in
	// names (_xlnm.: filter, print area) are plain ranges excelize adjusts
	// with what they belong to, and must keep their hidden flag, which the
	// API cannot express; they stay.
	var names []excelize.DefinedName
	for _, n := range e.f.GetDefinedName() {
		if strings.HasPrefix(n.Name, "_xlnm.") {
			continue
		}
		names = append(names, n)
		if err := e.f.DeleteDefinedName(&excelize.DefinedName{Name: n.Name, Scope: n.Scope}); err != nil {
			return err
		}
	}
	if err := op(); err != nil {
		return err
	}
	for _, n := range names {
		n.RefersTo = rewrite(n.RefersTo, "")
		if err := e.f.SetDefinedName(&n); err != nil {
			return err
		}
	}
	for _, m := range keep {
		if err := e.setFormula(m.sheet, grid.CellName(m.row, m.col), m.f); err != nil {
			return err
		}
	}
	e.structural = true
	e.forgetPositions()
	return nil
}

// forgetPositions drops cell positions recorded at open, which have moved;
// later edits in the same session work from what excelize holds.
func (e *Editor) forgetPositions() {
	for _, l := range e.layouts {
		l.spills = map[[2]int]grid.Range{}
	}
	e.unshare = map[string]bool{}
	for _, s := range e.wb.Sheets {
		s.Each(s.Used(), func(_, _ int, c *grid.Cell) { c.Shared = false })
	}
}

func (e *Editor) shiftExtent(sheet string, fn func(grid.Range) grid.Range) {
	ext, ok := e.extent[sheet]
	if !ok {
		ext = grid.Range{R1: 1, C1: 1}
		if s := e.wb.Sheet(sheet); s != nil {
			ext = s.Used()
		}
	}
	if !ext.Empty() {
		ext = fn(ext)
	}
	e.extent[sheet] = ext
}

// AddSheet adds an empty sheet, placed before or after another one when
// given.
func (e *Editor) AddSheet(name, before, after string) error {
	if e.wb.Sheet(name) != nil {
		return fmt.Errorf("sheet %q already exists", name)
	}
	if _, _, err := e.EnsureSheet(name); err != nil {
		return err
	}
	return e.place(name, before, after)
}

// place moves a sheet before or after another; with neither it stays.
func (e *Editor) place(name, before, after string) error {
	switch {
	case before != "":
		target, err := e.sheetName(before)
		if err != nil {
			return err
		}
		return e.f.MoveSheet(name, target)
	case after != "":
		target, err := e.sheetName(after)
		if err != nil {
			return err
		}
		list := e.f.GetSheetList()
		i := indexOf(list, target)
		if i == len(list)-1 {
			// MoveSheet places before its target: move the last sheet
			// before ours instead.
			if list[i] == name {
				return nil
			}
			return e.f.MoveSheet(list[i], name)
		}
		next := list[i+1]
		if next == name {
			return nil
		}
		return e.f.MoveSheet(name, next)
	}
	return nil
}

// MoveSheet changes a sheet's position: before or after another sheet, or
// to the first or last place.
func (e *Editor) MoveSheet(name, before, after string, first, last bool) error {
	name, err := e.sheetName(name)
	if err != nil {
		return err
	}
	list := e.f.GetSheetList()
	switch {
	case first:
		if list[0] == name {
			return nil
		}
		return e.f.MoveSheet(name, list[0])
	case last:
		return e.place(name, "", list[len(list)-1])
	}
	return e.place(name, before, after)
}

func indexOf(list []string, s string) int {
	for i, v := range list {
		if v == s {
			return i
		}
	}
	return -1
}

func (e *Editor) sheetName(name string) (string, error) {
	for _, n := range e.f.GetSheetList() {
		if strings.EqualFold(n, name) {
			return n, nil
		}
	}
	return "", fmt.Errorf("sheet %q not found (sheets: %s)", name, strings.Join(e.f.GetSheetList(), ", "))
}

// RenameSheet renames a sheet and every reference to it: formulas on all
// sheets, names, conditional formats, validations, charts and pivot
// sources. excelize itself only updates names.
func (e *Editor) RenameSheet(from, to string) error {
	from, err := e.sheetName(from)
	if err != nil {
		return err
	}
	if existing, err := e.sheetName(to); err == nil && existing != from {
		return fmt.Errorf("sheet %q already exists", existing)
	}
	if err := e.f.SetSheetName(from, to); err != nil {
		return fmt.Errorf("sheet name %q: %w", to, err)
	}
	e.patches.rename(from, to)
	e.patches.sheetEdits = append(e.patches.sheetEdits, sheetEdit{from: from, to: to})
	e.renameState(from, to)
	return nil
}

func (e *Editor) renameState(from, to string) {
	if s := e.wb.Sheet(from); s != nil {
		s.Name = to
	}
	if l, ok := e.layouts[from]; ok {
		e.layouts[to] = l
		delete(e.layouts, from)
	}
	if x, ok := e.extent[from]; ok {
		e.extent[to] = x
		delete(e.extent, from)
	}
}

// References lists the formula cells on other sheets that refer to sheet.
func (e *Editor) References(sheet string) []string {
	var refs []string
	for _, s := range e.wb.Sheets {
		if strings.EqualFold(s.Name, sheet) {
			continue
		}
		s.Each(s.Used(), func(r, c int, cell *grid.Cell) {
			if cell.Formula != "" && formula.RefersToSheet(cell.Formula, sheet) {
				refs = append(refs, grid.FormatRef(s.Name, grid.Range{R1: r, C1: c, R2: r, C2: c}))
			}
		})
	}
	for _, n := range e.wb.Names {
		if formula.RefersToSheet(n.Ref, sheet) {
			refs = append(refs, "name "+n.Name)
		}
	}
	return refs
}

// DeleteSheet deletes a sheet. References to it from elsewhere become
// #REF!, as in Excel; callers check References first.
func (e *Editor) DeleteSheet(name string) error {
	name, err := e.sheetName(name)
	if err != nil {
		return err
	}
	if len(e.f.GetSheetList()) == 1 {
		return fmt.Errorf("cannot delete %q: a workbook needs at least one sheet", name)
	}
	if err := e.f.DeleteSheet(name); err != nil {
		return err
	}
	e.patches.sheetEdits = append(e.patches.sheetEdits, sheetEdit{from: name})
	for i, s := range e.wb.Sheets {
		if s.Name == name {
			e.wb.Sheets = append(e.wb.Sheets[:i], e.wb.Sheets[i+1:]...)
			break
		}
	}
	delete(e.extent, name)
	delete(e.layouts, name)
	return nil
}

// CopySheet duplicates a sheet's cells and formatting under a new name,
// placed right after the original.
func (e *Editor) CopySheet(from, to string) ([]string, error) {
	from, err := e.sheetName(from)
	if err != nil {
		return nil, err
	}
	if s := e.wb.Sheet(from); s != nil && len(s.Tables) > 0 {
		return nil, fmt.Errorf("cannot copy %q: it holds Excel table %q, which cannot be duplicated", from, s.Tables[0].Name)
	}
	if e.wb.Sheet(to) != nil {
		return nil, fmt.Errorf("sheet %q already exists", to)
	}
	src, _ := e.f.GetSheetIndex(from)
	idx, err := e.f.NewSheet(to)
	if err != nil {
		return nil, fmt.Errorf("sheet name %q: %w", to, err)
	}
	if err := e.f.CopySheet(src, idx); err != nil {
		return nil, err
	}
	if err := e.place(to, "", from); err != nil {
		return nil, err
	}
	var notes []string
	if e.hasDrawings[from] {
		notes = append(notes, "charts and pictures are not copied")
	}
	if s := e.wb.Sheet(from); s != nil {
		cp := *s
		cp.Name = to
		e.wb.Sheets = append(e.wb.Sheets, &cp)
		e.extent[to] = s.Used()
	}
	return notes, nil
}

// SetVisible hides or shows a sheet.
func (e *Editor) SetVisible(name string, visible bool) error {
	name, err := e.sheetName(name)
	if err != nil {
		return err
	}
	if !visible {
		n := 0
		for _, s := range e.f.GetSheetList() {
			if v, _ := e.f.GetSheetVisible(s); v {
				n++
			}
		}
		if v, _ := e.f.GetSheetVisible(name); v && n == 1 {
			return fmt.Errorf("cannot hide %q: a workbook needs a visible sheet", name)
		}
	}
	return e.f.SetSheetVisible(name, visible)
}

// newRefErrors compares formulas before and after structural edits and
// lists cells that gained a #REF!.
func newRefErrors(before, after *grid.Workbook) []string {
	had := map[string]bool{}
	for _, s := range before.Sheets {
		s.Each(s.Used(), func(_, _ int, c *grid.Cell) {
			if strings.Contains(c.Formula, "#REF!") {
				had[s.Name+"\x00"+c.Formula] = true
			}
		})
	}
	var out []string
	for _, s := range after.Sheets {
		s.Each(s.Used(), func(r, c int, cell *grid.Cell) {
			if strings.Contains(cell.Formula, "#REF!") && !had[s.Name+"\x00"+cell.Formula] {
				out = append(out, grid.FormatRef(s.Name, grid.Range{R1: r, C1: c, R2: r, C2: c}))
			}
		})
	}
	sort.Strings(out)
	return out
}

// SheetList is the current sheet order, hidden sheets marked.
func (e *Editor) SheetList() []string {
	var out []string
	for _, s := range e.f.GetSheetList() {
		if v, _ := e.f.GetSheetVisible(s); !v {
			s += " (hidden)"
		}
		out = append(out, s)
	}
	return out
}

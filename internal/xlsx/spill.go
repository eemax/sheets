package xlsx

import (
	"github.com/eemax/sheets/internal/engine"
	"github.com/eemax/sheets/internal/grid"
	"github.com/xuri/excelize/v2"
)

// Dynamic arrays (Excel 365 formulas whose result spills into the cells
// below and to the right) are saved as Excel saves them: the formula's
// range covers its result, and the other cells of that range hold their
// values. So every reader of the file (pandas, LibreOffice, previews) sees
// the whole result, not its first cell alone. A result that would land on
// other cells is #SPILL!, as in Excel.

// maxSpill bounds the results laid out, in cells; a larger one keeps its
// first cell, and Excel spills it on open.
const maxSpill = 100_000

// spillResults lays out the results of the dynamic arrays computed over
// their spill ranges. It reports whether a range or a value changed, so
// that what reads them needs working out again.
func (e *Editor) spillResults(res calcResult) (bool, error) {
	if len(res.arrays) == 0 {
		return false, nil
	}
	wb, err := e.stateNow()
	if err != nil {
		return false, err
	}
	changed := false
	for k, sp := range res.arrays {
		a := sp.arr
		if a == nil || a.Rows*a.Cols > maxSpill {
			continue
		}
		sheet, anchor := sp.sheet, [2]int{k.row, k.col}
		s := wb.Sheet(sheet)
		if s == nil {
			continue
		}
		if e.layouts == nil {
			e.layouts = map[string]*layout{}
		}
		lay := e.layouts[sheet]
		if lay == nil {
			lay = &layout{spills: map[[2]int]grid.Range{}}
			e.layouts[sheet] = lay
		}
		if lay.spills == nil {
			lay.spills = map[[2]int]grid.Range{}
		}
		here := grid.Range{R1: k.row, C1: k.col, R2: k.row, C2: k.col}
		old, ok := lay.spills[anchor]
		if !ok {
			old = here
		}
		want := grid.Range{R1: k.row, C1: k.col, R2: k.row + a.Rows - 1, C2: k.col + a.Cols - 1}
		blocked := want.R2 > grid.MaxRows || want.C2 > grid.MaxCols
		if !blocked {
			s.Each(want, func(r, c int, cell *grid.Cell) {
				if r == k.row && c == k.col || cell.IsEmpty() || old.Contains(r, c) && cell.Formula == "" {
					return
				}
				blocked = true // another value in the way
			})
		}
		if blocked {
			want = here
			res.results[k.sheet][grid.CellName(k.row, k.col)] = grid.Cell{Kind: grid.Error, Text: "#SPILL!"}
			res.dropped[k] = "#SPILL!: the result would cover cells that hold values; clear them"
		}
		// The cells the old result filled and the new one leaves.
		var clearErr error
		s.Each(old, func(r, c int, cell *grid.Cell) {
			if clearErr != nil || r == k.row && c == k.col || want.Contains(r, c) || cell.Formula != "" || cell.IsEmpty() {
				return
			}
			clearErr = e.f.SetCellDefault(sheet, grid.CellName(r, c), "")
			e.touch(sheet, r, c)
			changed = true
		})
		if clearErr != nil {
			return false, clearErr
		}
		if !blocked {
			for i := 0; i < a.Rows; i++ {
				for j := 0; j < a.Cols; j++ {
					if i == 0 && j == 0 {
						continue
					}
					r, c := k.row+i, k.col+j
					v := spillValue(a.At(i, j))
					if now := s.Get(r, c); now.Kind == v.Kind && now.Num == v.Num && now.Text == v.Text && now.Bool == v.Bool && now.Formula == "" {
						continue
					}
					if err := e.putValue(sheet, r, c, v); err != nil {
						return false, err
					}
					e.touch(sheet, r, c)
					changed = true
				}
			}
		}
		if want != old {
			cell := grid.CellName(k.row, k.col)
			f, err := e.f.GetCellFormula(sheet, cell)
			if err != nil {
				return false, err
			}
			t, ref := excelize.STCellFormulaTypeArray, want.String()
			if err := e.f.SetCellFormula(sheet, cell, f, excelize.FormulaOpts{Type: &t, Ref: &ref}); err != nil {
				return false, err
			}
			e.patches.markCell(sheet, cell, true)
			lay.spills[anchor] = want
			e.edits++
			changed = true
		}
	}
	return changed, nil
}

// keepSpills adds the values of the spill ranges to the results stored,
// so that saving keeps them in the range of their formula.
func (e *Editor) keepSpills(res calcResult) {
	for k, sp := range res.arrays {
		lay := e.layouts[sp.sheet]
		if lay == nil {
			continue
		}
		rg, ok := lay.spills[[2]int{k.row, k.col}]
		if !ok || sp.arr == nil || rg.Rows() != sp.arr.Rows || rg.Cols() != sp.arr.Cols {
			continue
		}
		for i := 0; i < sp.arr.Rows; i++ {
			for j := 0; j < sp.arr.Cols; j++ {
				if i > 0 || j > 0 {
					e.patches.results[k.sheet][grid.CellName(k.row+i, k.col+j)] = spillValue(sp.arr.At(i, j))
				}
			}
		}
	}
}

// spillValue is an element of a result as a cell of its spill range holds
// it: a blank shows as 0, as in Excel.
func spillValue(v engine.Value) grid.Cell {
	if v.Kind == engine.Empty {
		return grid.Cell{Kind: grid.Number}
	}
	return calcToCell(v)
}

// putValue writes a value into a cell, an error value as one.
func (e *Editor) putValue(sheet string, row, col int, v grid.Cell) error {
	cell := grid.CellName(row, col)
	delete(e.patches.errorCells[sheet], cell)
	switch v.Kind {
	case grid.Number:
		return e.f.SetCellFloat(sheet, cell, v.Num, -1, 64)
	case grid.Text:
		return setText(e.f, sheet, cell, v.Text)
	case grid.Bool:
		return e.f.SetCellBool(sheet, cell, v.Bool)
	case grid.Error:
		if err := e.f.SetCellDefault(sheet, cell, v.Text); err != nil {
			return err
		}
		if e.patches.errorCells[sheet] == nil {
			e.patches.errorCells[sheet] = map[string]bool{}
		}
		e.patches.errorCells[sheet][cell] = true
		return nil
	}
	return e.f.SetCellDefault(sheet, cell, "")
}

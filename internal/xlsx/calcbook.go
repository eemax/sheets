package xlsx

import (
	"strings"

	"github.com/eemax/sheets/internal/engine"
	"github.com/eemax/sheets/internal/grid"
)

// gridBook lets the formula engine read a workbook as sheets parsed it.
// text, when set, gives the text the engine computes for a formula (table
// references translated into cells, say).
type gridBook struct {
	wb   *grid.Workbook
	text func(sheet string, row, col int, formula string) string
}

func (b *gridBook) sheet(name string) *grid.Sheet {
	if s := b.wb.Sheet(name); s != nil {
		return s
	}
	for _, s := range b.wb.Sheets {
		if strings.EqualFold(s.Name, name) {
			return s
		}
	}
	return nil
}

func (b *gridBook) Cell(sheet string, row, col int) (engine.Value, string) {
	s := b.sheet(sheet)
	if s == nil {
		return engine.Value{}, ""
	}
	c := s.Get(row, col)
	f := c.Formula
	if f != "" && b.text != nil {
		f = b.text(s.Name, row, col, f)
	}
	if c.Pending {
		return engine.Value{}, f
	}
	return cellToCalc(c), f
}

func (b *gridBook) Extent(sheet string) (int, int, bool) {
	s := b.sheet(sheet)
	if s == nil {
		return 0, 0, false
	}
	u := s.Used()
	return max(u.R2, 0), max(u.C2, 0), true
}

func (b *gridBook) Name(name, sheet string) (string, bool) {
	var global *grid.DefinedName
	for i := range b.wb.Names {
		n := &b.wb.Names[i]
		if !strings.EqualFold(n.Name, name) {
			continue
		}
		if n.Scope != "" && strings.EqualFold(n.Scope, sheet) {
			return strings.TrimPrefix(n.Ref, "="), true
		}
		if n.Scope == "" {
			global = n
		}
	}
	if global != nil {
		return strings.TrimPrefix(global.Ref, "="), true
	}
	return "", false
}

func (b *gridBook) ArrayAt(sheet string, row, col int) (int, int, engine.Area, bool) {
	s := b.sheet(sheet)
	if s == nil {
		return 0, 0, engine.Area{}, false
	}
	for _, a := range s.Arrays {
		if a.Contains(row, col) {
			return a.R1, a.C1, engine.Area{R1: a.R1, C1: a.C1, R2: a.R2, C2: a.C2}, true
		}
	}
	return 0, 0, engine.Area{}, false
}

// cellToCalc is a cell's value for the engine.
func cellToCalc(c grid.Cell) engine.Value {
	switch c.Kind {
	case grid.Number:
		return engine.Value{Kind: engine.Number, Num: c.Num}
	case grid.Text:
		return engine.Value{Kind: engine.Text, Str: c.Text}
	case grid.Bool:
		return engine.Value{Kind: engine.Bool, Bool: c.Bool}
	case grid.Error:
		return engine.Value{Kind: engine.Error, Str: c.Text}
	}
	return engine.Value{}
}

// calcToCell is an engine value as a cell holds it.
func calcToCell(v engine.Value) grid.Cell {
	switch v.Kind {
	case engine.Number:
		return grid.Cell{Kind: grid.Number, Num: v.Num}
	case engine.Text:
		return grid.Cell{Kind: grid.Text, Text: v.Str}
	case engine.Bool:
		return grid.Cell{Kind: grid.Bool, Bool: v.Bool}
	case engine.Error:
		return grid.Cell{Kind: grid.Error, Text: v.Str}
	}
	return grid.Cell{}
}

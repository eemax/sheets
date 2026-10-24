package xlsx

import (
	"fmt"

	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/input"
)

// Replacement is one cell replace changed.
type Replacement struct {
	Ref string `json:"ref"`
	Old string `json:"old"`
	New string `json:"new"`
}

// Replace runs fn over the cells of a sheet (r limits it; every sheet when
// sheet is ""), as Excel's Replace does: each value as it reads (ISO dates,
// 12%), and with formulas set, each formula's text (=SUM(B2:B9)). Where fn
// changes the text, the cell is set to it as typed input, so "12" becomes a
// number; a value that comes to start with = stays text, unless formulas
// are being replaced. Formulas' results and the results of dynamic arrays
// are not values of their own and are left alone.
func (e *Editor) Replace(sheet string, r *grid.Range, fn func(string) (string, bool), formulas bool) ([]Replacement, error) {
	if !e.fresh() {
		return nil, errNotFresh
	}
	var sheets []*grid.Sheet
	if sheet == "" {
		sheets = e.wb.Sheets
	} else {
		name, err := e.sheetName(sheet)
		if err != nil {
			return nil, err
		}
		sheets = []*grid.Sheet{e.wb.Sheet(name)}
	}
	var out []Replacement
	for _, s := range sheets {
		if s == nil {
			continue
		}
		area := s.Used()
		if r != nil {
			area = area.Intersect(*r)
		}
		if area.Empty() {
			continue
		}
		spills := e.spillsOn(s.Name)
		headers := map[[2]int]string{} // header cells of Excel tables
		for _, t := range e.tablesOn(s.Name) {
			if t.headers > 0 {
				for c := t.ref.C1; c <= t.ref.C2; c++ {
					headers[[2]int{t.ref.R1, c}] = t.name
				}
			}
		}
		type change struct {
			row, col int
			v        input.Value
			rep      Replacement
		}
		var todo []change
		var err error
		s.Each(area, func(row, col int, c *grid.Cell) {
			if err != nil {
				return
			}
			if _, hidden, _ := hiddenIn(s.Merges, row, col); hidden {
				return // not shown: Excel's Replace passes it by
			}
			old := ""
			switch {
			case c.Formula != "":
				if !formulas {
					return
				}
				old = "=" + c.Formula
			case ghost(spills, row, col):
				return
			default:
				old = grid.Display(c, e.wb.FormatOf(c), e.wb.Date1904)
			}
			text, ok := fn(old)
			if !ok || text == old {
				return
			}
			v := input.Parse(text, false)
			if v.Formula != "" && !formulas {
				v = input.Value{Kind: grid.Text, Text: text}
			}
			ref := grid.FormatRef(s.Name, grid.Range{R1: row, C1: col, R2: row, C2: col})
			if t, ok := headers[[2]int{row, col}]; ok && (v.IsEmpty() || v.Formula != "") {
				err = fmt.Errorf("%s is a column name of Excel table %q, which cannot be empty or a formula", ref, t)
				return
			}
			todo = append(todo, change{row, col, v, Replacement{Ref: ref, Old: old, New: text}})
		})
		if err != nil {
			return nil, err
		}
		for _, ch := range todo {
			if err := e.Set(s.Name, ch.row, ch.col, ch.v); err != nil {
				return nil, err
			}
			out = append(out, ch.rep)
		}
	}
	return out, nil
}

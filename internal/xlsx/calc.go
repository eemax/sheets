package xlsx

import (
	"errors"
	"sync"
	"time"

	"github.com/eemax/sheets/internal/engine"
	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
)

// calc computes formulas whose results a file does not hold, with the
// sheets calculator (internal/engine) over the workbook as read. As when storing results, an
// answer counts only when the judge says it holds (see judge.go).
type calc struct {
	wb   *grid.Workbook
	once sync.Once

	at map[cellKey]*formulaAt
	j  *judge
}

func (c *calc) Resolve(sheet string, row, col int) (grid.Cell, error) {
	c.once.Do(c.analyse)
	f := c.at[keyOf(sheet, row, col)]
	if f == nil {
		return grid.Cell{}, errUnsupported
	}
	if ok, _ := c.j.holds(f); !ok {
		return grid.Cell{}, errUnsupported
	}
	return c.j.result(f).v, nil
}

// analyse prepares the engine and the judge over the workbook's formulas.
func (c *calc) analyse() {
	c.at = map[cellKey]*formulaAt{}
	pending := map[*formulaAt]bool{}
	var list []*formulaAt
	if c.wb == nil {
		c.wb = &grid.Workbook{}
	}
	for _, s := range c.wb.Sheets {
		s.Each(s.Used(), func(r, col int, cell *grid.Cell) {
			if cell.Formula == "" {
				return
			}
			f := &formulaAt{sheet: s.Name, row: r, col: col, text: cell.Formula, pending: cell.Pending, stored: valueOf(*cell), unprefixed: cell.Unprefixed}
			c.at[keyOf(s.Name, r, col)] = f
			list = append(list, f)
			if cell.Pending {
				pending[f] = true
			}
		})
	}
	rs := wbResolver(c.wb)
	for _, f := range list {
		f.calc = rs.calcText(f)
	}
	g := newCycleGraph(list, func(f *formulaAt) ([]formula.Ref, bool) {
		return rs.ranges(f.sheet, f.row, f.col, formula.DepsOf(calcOf(f)), 0)
	})
	plan := newPlan(list, rs, pending)
	book := &gridBook{wb: c.wb, text: func(sheet string, row, col int, text string) string {
		if f := c.at[keyOf(sheet, row, col)]; f != nil {
			return calcOf(f)
		}
		return text
	}}
	ty := newTyper(c.wb, nil, nil, false, c.at, pending, rs)
	used := func(sheet string) grid.Range {
		if s := c.wb.Sheet(sheet); s != nil {
			return s.Used()
		}
		return grid.Range{R1: 1, C1: 1}
	}
	plan.native = nativeAggregates(used, func(j *judge, sheet string, r, col int) (grid.Cell, bool) {
		if f := c.at[keyOf(sheet, r, col)]; f != nil {
			res := j.result(f)
			return res.v, res.err == nil
		}
		if s := c.wb.Sheet(sheet); s != nil {
			return s.Get(r, col), true
		}
		return grid.Cell{}, true
	})
	eng := engine.New(book)
	eng.Date1904 = c.wb.Date1904
	c.j = newJudge(eng, g, rs, ty, pending, plan, func(f *formulaAt) (grid.Cell, bool) {
		return f.stored, f.stored.Kind != grid.Empty && !f.pending
	}, time.Now().Add(24*time.Hour))
}

// errUnsupported marks formulas the engine cannot compute although Excel
// can (functions it does not have, links to other workbooks).
var errUnsupported = errors.New("not supported by the sheets calculator")

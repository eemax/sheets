package xlsx

import (
	"fmt"
	"sort"
	"testing"

	"github.com/eemax/sheets/internal/grid"
)

// After any edit of rows or columns, in either editor, the array formulas
// (legacy and dynamic) are the ones there were, where their cells went,
// and no other formula has become one.
func TestStructureKeepsArrays(t *testing.T) {
	type op struct {
		name    string
		cols    bool
		del     bool
		at      int
		applyTo func(ed interface {
			InsertRows(string, int, int) error
			DeleteRows(string, int, int) error
			InsertCols(string, int, int) error
			DeleteCols(string, int, int) error
		}, sheet string) error
	}
	var ops []op
	for at := 1; at <= 7; at++ {
		at := at
		ops = append(ops,
			op{fmt.Sprintf("rows insert %d", at), false, false, at, func(ed interface {
				InsertRows(string, int, int) error
				DeleteRows(string, int, int) error
				InsertCols(string, int, int) error
				DeleteCols(string, int, int) error
			}, s string) error {
				return ed.InsertRows(s, at, 2)
			}},
			op{fmt.Sprintf("rows delete %d", at), false, true, at, func(ed interface {
				InsertRows(string, int, int) error
				DeleteRows(string, int, int) error
				InsertCols(string, int, int) error
				DeleteCols(string, int, int) error
			}, s string) error {
				return ed.DeleteRows(s, at, at)
			}},
			op{fmt.Sprintf("cols insert %d", at), true, false, at, func(ed interface {
				InsertRows(string, int, int) error
				DeleteRows(string, int, int) error
				InsertCols(string, int, int) error
				DeleteCols(string, int, int) error
			}, s string) error {
				return ed.InsertCols(s, at, 2)
			}},
			op{fmt.Sprintf("cols delete %d", at), true, true, at, func(ed interface {
				InsertRows(string, int, int) error
				DeleteRows(string, int, int) error
				InsertCols(string, int, int) error
				DeleteCols(string, int, int) error
			}, s string) error {
				return ed.DeleteCols(s, at, at)
			}})
	}
	// Where a cell goes: ok false when deleted.
	move := func(o op, r, c int) (int, int, bool) {
		x := &r
		if o.cols {
			x = &c
		}
		switch {
		case o.del && *x == o.at:
			return 0, 0, false
		case o.del && *x > o.at:
			*x--
		case !o.del && *x >= o.at:
			*x += 2
		}
		return r, c, true
	}
	anchors := func(wb *grid.Workbook) []string {
		var out []string
		for _, s := range wb.Sheets {
			for _, a := range s.Arrays {
				out = append(out, s.Name+"!"+grid.CellName(a.R1, a.C1))
			}
		}
		sort.Strings(out)
		return out
	}
	for _, name := range []string{"arrays_rows.xlsx", "arrays_openpyxl.xlsx", "arrays_xlsxwriter.xlsx", "dynamic_xlsxwriter.xlsx", "refs.xlsx"} {
		fixed := checkData(t, name, openFixture(t, name))
		data := fixed.Data
		if data == nil {
			data = openFixture(t, name)
		}
		before, err := ReadBytes(data)
		must(t, err)
		sheet := before.Sheets[0].Name
		for _, mode := range []string{"always", "never"} {
			t.Setenv("SHEETS_STREAM", mode)
			for _, o := range ops {
				var ed interface {
					InsertRows(string, int, int) error
					DeleteRows(string, int, int) error
					InsertCols(string, int, int) error
					DeleteCols(string, int, int) error
					Render() ([]byte, error)
				}
				if mode == "never" {
					e, err := OpenEditorData(name, data)
					must(t, err)
					ed = e
				} else {
					a, err := OpenAuto(name, data)
					must(t, err)
					ed = a
				}
				if err := o.applyTo(ed, sheet); err != nil {
					continue // refused: it cuts an array
				}
				out, err := ed.Render()
				if err != nil {
					t.Fatalf("%s %s %s: %v", name, mode, o.name, err)
				}
				after, err := ReadBytes(out)
				must(t, err)
				var want []string
				for _, s := range before.Sheets {
					for _, a := range s.Arrays {
						r, c, ok := a.R1, a.C1, true
						if s.Name == sheet {
							r, c, ok = move(o, r, c)
						}
						if ok {
							want = append(want, s.Name+"!"+grid.CellName(r, c))
						}
					}
				}
				sort.Strings(want)
				if got := anchors(after); fmt.Sprint(got) != fmt.Sprint(want) {
					t.Errorf("%s, %s, %s: arrays at %v, want %v", name, mode, o.name, got, want)
				}
			}
		}
	}
}

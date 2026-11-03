package xlsx

import (
	"bytes"
	"os"
	"path/filepath"
	"regexp"
	"testing"

	"github.com/eemax/sheets/internal/input"
	"github.com/xuri/excelize/v2"
)

// A formula of its own inside an array formula's range, which Excel does
// not allow: neither it nor the array gets a result, in either editor, and
// a full-editor edit leaves it as it is.
func TestFormulaInsideArray(t *testing.T) {
	data := openFixture(t, "arrays_inside.xlsx") // E1:E3 {=A1:A3*3}, E3 =A1+100
	bothEditors(t, func(t *testing.T) {
		res := checkData(t, "i.xlsx", data)
		if len(res.Unknown) != 2 {
			t.Errorf("unknown: %+v", res.Unknown)
		}
		fixedTwice(t, data, res)
	})
	path := filepath.Join(t.TempDir(), "i.xlsx")
	must(t, os.WriteFile(path, data, 0o644))
	e, err := OpenEditorData(path, data)
	must(t, err)
	must(t, e.Set("S", 1, 1, input.Parse("5", false)))
	must(t, e.Save(false))
	wb := read(t, path)
	if c := wb.Sheet("S").Get(3, 5); c.Formula != "A1+100" {
		t.Errorf("E3 now %+v", c)
	}
}

// A link to another workbook that an edit reaches loses its result (only
// Excel can update it), so nothing is stored that is computed from it.
func TestEditReachesLink(t *testing.T) {
	f := excelize.NewFile()
	f.SetCellValue("Sheet1", "A1", 2)
	f.SetCellFormula("Sheet1", "B1", "[1]Rates!A1*A1")
	f.SetCellFormula("Sheet1", "B3", "B1+1")
	var buf bytes.Buffer
	must(t, f.Write(&buf))
	// As Excel saved it: results vouched for.
	data := patchZip(t, buf.Bytes(), dataPart, func(x string) string {
		for cell, v := range map[string]string{"B1": "7", "B3": "8"} {
			x = regexp.MustCompile(`<c r="`+cell+`"[^>]*>(<f>[^<]*</f>)(<v>[^<]*</v>)?`).ReplaceAllString(x, `<c r="`+cell+`">${1}<v>`+v+`</v>`)
		}
		return x
	})
	for _, mode := range []string{"always", "never"} {
		t.Run(mode, func(t *testing.T) {
			t.Setenv("SHEETS_STREAM", mode)
			path := filepath.Join(t.TempDir(), "l.xlsx")
			must(t, os.WriteFile(path, data, 0o644))
			var ed interface {
				Set(string, int, int, input.Value) error
				Save(bool) error
			}
			if mode == "never" {
				e, err := OpenEditorData(path, data)
				must(t, err)
				ed = e
			} else {
				a, err := OpenAuto(path, data)
				must(t, err)
				ed = a
			}
			must(t, ed.Set("Sheet1", 1, 1, input.Parse("5", false)))
			must(t, ed.Save(false))
			saved, err := os.ReadFile(path)
			must(t, err)
			wb, _ := ReadBytes(saved)
			s := wb.Sheet("Sheet1")
			if b1, b3 := s.Get(1, 2), s.Get(3, 2); !b1.Pending || !b3.Pending {
				t.Errorf("after the edit B1 %+v, B3 %+v", b1, b3)
			}
			res := checkData(t, "l.xlsx", saved)
			if res.Data != nil {
				t.Errorf("check --fix stores %+v", res.Changes)
			}
		})
	}
}

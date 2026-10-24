package xlsx

import (
	"fmt"
	"os"
	"regexp"
	"strings"
	"testing"

	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/input"
)

// condsOn lists a sheet's conditional formats and validations as range and
// rule, for comparing.
func condsOn(t *testing.T, path, sheet string) string {
	t.Helper()
	e, err := OpenEditor(path, false)
	must(t, err)
	defer e.Close()
	cfs, err := e.CondFormats(sheet)
	must(t, err)
	var out []string
	for _, cf := range cfs {
		out = append(out, cf.Range+" "+cf.Rule)
	}
	dvs, err := e.Validations(sheet)
	must(t, err)
	for _, dv := range dvs {
		out = append(out, "valid "+dv.Range+" "+dv.Rule)
	}
	return strings.Join(out, "; ")
}

// ruleBook has a conditional format and a validation whose formulas read
// their own row, and a sheet whose rules read the first.
func ruleBook(t *testing.T) string {
	path := writeBook(t, [][]string{
		{"Region", "Units", "Price"},
		{"East", "5", "30"},
		{"West", "2", "40"},
		{"North", "9", "10"},
		{"South", "1", "100"},
	})
	change(t, path, func(e *Editor) {
		_, err := e.Cond("Data", rng("A2:C5"), []string{`=$B2>4`, "fill=lightyellow"})
		must(t, err)
		_, err = e.Validate("Data", rng("C2:C5"), ValidationRule{Words: []string{"formula", "=C2>$B2"}})
		must(t, err)
		_, _, err = e.EnsureSheet("Other")
		must(t, err)
		_, err = e.Cond("Other", rng("A1:A3"), []string{`=Data!$B$3>1`})
		must(t, err)
	})
	return path
}

func TestRuleFormulasFollowFirstCell(t *testing.T) {
	cases := []struct {
		name        string
		edit        func(e *Editor) error
		data, other string
	}{
		{"first row deleted", func(e *Editor) error { return e.DeleteRows("Data", 2, 2) },
			"A2:C4 =$B2>4; valid C2:C4 formula =C2>$B2", "A1:A3 =Data!$B$2>1"},
		{"first two rows deleted", func(e *Editor) error { return e.DeleteRows("Data", 2, 3) },
			"A2:C3 =$B2>4; valid C2:C3 formula =C2>$B2", "A1:A3 =Data!#REF!>1"},
		{"header deleted", func(e *Editor) error { return e.DeleteRows("Data", 1, 1) },
			"A1:C4 =$B1>4; valid C1:C4 formula =C1>$B1", "A1:A3 =Data!$B$2>1"},
		{"row inserted at the first", func(e *Editor) error { return e.InsertRows("Data", 2, 1) },
			"A3:C6 =$B3>4; valid C3:C6 formula =C3>$B3", "A1:A3 =Data!$B$4>1"},
		{"row inserted within", func(e *Editor) error { return e.InsertRows("Data", 3, 2) },
			"A2:C7 =$B2>4; valid C2:C7 formula =C2>$B2", "A1:A3 =Data!$B$5>1"},
		{"last row deleted", func(e *Editor) error { return e.DeleteRows("Data", 5, 5) },
			"A2:C4 =$B2>4; valid C2:C4 formula =C2>$B2", "A1:A3 =Data!$B$3>1"},
		{"first column deleted", func(e *Editor) error { return e.DeleteCols("Data", 1, 1) },
			"A2:B5 =$A2>4; valid B2:B5 formula =B2>$A2", "A1:A3 =Data!$A$3>1"},
		{"read column deleted", func(e *Editor) error { return e.DeleteCols("Data", 2, 2) },
			"A2:B5 =#REF!>4; valid B2:B5 formula =B2>#REF!", "A1:A3 =Data!#REF!>1"},
		{"column inserted", func(e *Editor) error { return e.InsertCols("Data", 1, 1) },
			"B2:D5 =$C2>4; valid D2:D5 formula =D2>$C2", "A1:A3 =Data!$C$3>1"},
		{"all rows deleted", func(e *Editor) error { return e.DeleteRows("Data", 2, 5) },
			"", "A1:A3 =Data!#REF!>1"},
		{"block moved", func(e *Editor) error {
			_, err := e.MoveRange("Data", rng("A2:C5"), "Data", rng("E2"))
			return err
		}, "E2:G5 =$F2>4; valid G2:G5 formula =G2>$F2", "A1:A3 =Data!$F$3>1"},
		{"sheet renamed", func(e *Editor) error { return e.RenameSheet("Data", "Facts") },
			"", "A1:A3 =Facts!$B$3>1"},
	}
	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			path := ruleBook(t)
			change(t, path, func(e *Editor) { must(t, c.edit(e)) })
			validPackage(t, path)
			data := "Data"
			if c.name == "sheet renamed" {
				data = "Facts"
				c.data = "A2:C5 =$B2>4; valid C2:C5 formula =C2>$B2"
			}
			if got := condsOn(t, path, data); got != c.data {
				t.Errorf("Data: got %q, want %q", got, c.data)
			}
			if got := condsOn(t, path, "Other"); got != c.other {
				t.Errorf("Other: got %q, want %q", got, c.other)
			}
		})
	}
}

func TestRulesAfterEditsInOneSession(t *testing.T) {
	path := ruleBook(t)
	change(t, path, func(e *Editor) {
		must(t, e.InsertRows("Data", 1, 1))
		// Given for the cells as they are now: it must not move again.
		_, err := e.Cond("Data", rng("B3:B6"), []string{`=$B3>1`, "bold"})
		must(t, err)
		cfs, err := e.CondFormats("Data")
		must(t, err)
		if len(cfs) != 2 {
			t.Fatalf("listed in session: %v", cfs)
		}
		must(t, e.DeleteRows("Data", 3, 3))
		n, err := e.ClearCond("Data", &grid.Range{R1: 1, C1: 1, R2: 100, C2: 1})
		must(t, err)
		if n != 0 {
			t.Errorf("cleared %d formats within column A, which none lies within", n)
		}
	})
	if got, want := condsOn(t, path, "Data"), "A3:C5 =$B3>4; B3:B5 =$B3>1; valid C3:C5 formula =C3>$B3"; got != want {
		t.Errorf("got %q, want %q", got, want)
	}
	change(t, path, func(e *Editor) {
		n, err := e.ClearCond("Data", &grid.Range{R1: 3, C1: 2, R2: 5, C2: 2})
		must(t, err)
		if n != 1 {
			t.Errorf("cleared %d, want 1", n)
		}
	})
	if got, want := condsOn(t, path, "Data"), "A3:C5 =$B3>4; valid C3:C5 formula =C3>$B3"; got != want {
		t.Errorf("after clearing: got %q, want %q", got, want)
	}
}

func TestSheetNamesSwapped(t *testing.T) {
	path := ruleBook(t)
	change(t, path, func(e *Editor) {
		must(t, e.RenameSheet("Data", "Tmp"))
		must(t, e.RenameSheet("Other", "Data"))
		must(t, e.RenameSheet("Tmp", "Other"))
	})
	if got, want := condsOn(t, path, "Data"), "A1:A3 =Other!$B$3>1"; got != want {
		t.Errorf("got %q, want %q", got, want)
	}
}

var x14ID = regexp.MustCompile(`<x14:id>(\{[^}]+\})</x14:id>`)

func TestDataBarsKeepTheirParts(t *testing.T) {
	path := ruleBook(t)
	change(t, path, func(e *Editor) {
		_, err := e.Cond("Data", rng("B2:B5"), []string{"bars"})
		must(t, err)
		_, err = e.Cond("Data", rng("C2:C5"), []string{"bars", "green"})
		must(t, err)
	})
	change(t, path, func(e *Editor) {
		must(t, e.DeleteRows("Data", 2, 2))
		_, err := e.Cond("Data", rng("A2:A4"), []string{"bars"})
		must(t, err)
	})
	validPackage(t, path)
	part := sheetXML(t, path, "xl/worksheets/sheet1.xml")
	ids := x14ID.FindAllStringSubmatch(part, -1)
	seen := map[string]bool{}
	for _, id := range ids {
		if seen[id[1]] {
			t.Errorf("id %s used twice", id[1])
		}
		seen[id[1]] = true
		if !strings.Contains(part, `<x14:cfRule type="dataBar" id="`+id[1]+`"`) {
			t.Errorf("no extension rule for %s", id[1])
		}
	}
	if len(ids) != 3 || strings.Count(part, `<x14:cfRule type="dataBar"`) != 3 {
		t.Fatalf("%d data bars, want 3:\n%s", len(ids), part)
	}
	// The extension in the sheet's list, after the main blocks, once.
	if strings.Count(part, "<x14:conditionalFormattings>") != 1 ||
		strings.Index(part, "<x14:conditionalFormattings>") < strings.LastIndex(part, "</conditionalFormatting>") {
		t.Errorf("extension misplaced:\n%s", part)
	}
	for _, want := range []string{"<xm:sqref>B2:B4</xm:sqref>", "<xm:sqref>C2:C4</xm:sqref>", "<xm:sqref>A2:A4</xm:sqref>"} {
		if !strings.Contains(part, want) {
			t.Errorf("no %s", want)
		}
	}
	// Priorities distinct, in the order the rules were made.
	prios := regexp.MustCompile(`<cfRule [^>]*priority="(\d+)"`).FindAllStringSubmatch(part, -1)
	var got []string
	for _, p := range prios {
		got = append(got, p[1])
	}
	if fmt.Sprint(got) != "[1 2 3 4]" {
		t.Errorf("priorities %v", got)
	}
}

func TestMergedCellsHideTheirCells(t *testing.T) {
	path := writeBook(t, [][]string{
		{"Title", "", "", ""},
		{"a", "b", "c", "d"},
		{"1", "2", "3", "4"},
	})
	change(t, path, func(e *Editor) { must(t, e.f.MergeCell("Data", "A1", "D1")) })
	for _, stream := range []bool{false, true} {
		t.Run(fmt.Sprint("stream=", stream), func(t *testing.T) {
			var ed editor
			if stream {
				data, err := os.ReadFile(path)
				must(t, err)
				a, err := OpenAuto(path, data)
				must(t, err)
				ed = a
			} else {
				e, err := OpenEditor(path, false)
				must(t, err)
				defer e.Close()
				ed = e
			}
			err := ed.Set("Data", 1, 3, input.Parse("x", false))
			if err == nil || !strings.Contains(err.Error(), "Data!C1 lies inside merged cells A1:D1, which show only A1") {
				t.Errorf("write into a merge's hidden cell: %v", err)
			}
			must(t, ed.Set("Data", 1, 2, input.Value{})) // nothing to clear: no error
			must(t, ed.InsertRows("Data", 1, 1))         // the merge moves to row 2
			must(t, ed.Set("Data", 1, 3, input.Parse("y", false)))
			if err := ed.Set("Data", 2, 4, input.Parse("x", false)); err == nil || !strings.Contains(err.Error(), "A2:D2") {
				t.Errorf("after the merge moved: %v", err)
			}
			must(t, ed.Set("Data", 2, 1, input.Parse("New title", false)))
		})
	}
}

package fuzz

import (
	"fmt"
	"strings"
	"testing"

	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/xlsx"
)

// objBook is a workbook whose sheet Data holds a table over A1:C5, a merge
// over A8:C8, a note on B3, a link on A2, a chart at H2 reading the table,
// row 7 and column E hidden. Rows from ins on (row 3) stand n rows down: the
// workbook after rows are inserted before row 3.
func objBook(n int) *grid.Workbook {
	row := func(r int) int {
		if r >= 3 {
			return r + n
		}
		return r
	}
	cells := map[string]string{}
	for i, name := range []string{"Item", "Units", "Cost"} {
		cells[fmt.Sprintf("Data!%c1", 'A'+i)] = "'" + name
	}
	for r := 2; r <= 5; r++ {
		cells[fmt.Sprintf("Data!A%d", row(r))] = "'x"
		cells[fmt.Sprintf("Data!B%d", row(r))] = fmt.Sprint(r)
		cells[fmt.Sprintf("Data!C%d", row(r))] = fmt.Sprint(r * 2)
	}
	cells[fmt.Sprintf("Data!A%d", row(8))] = "'Footer"
	wb := reachBook(cells)
	s := wb.Sheet("Data")
	s.Tables = []grid.NamedTable{{Name: "Sales1", Range: grid.Range{R1: 1, C1: 1, R2: row(5), C2: 3}, Filter: true, Columns: []string{"Item", "Units", "Cost"}}}
	s.Merges = []grid.Range{{R1: row(8), C1: 1, R2: row(8), C2: 3}}
	s.Notes = []grid.Note{{Row: row(3), Col: 2, Author: "Ann", Text: "check"}}
	s.Links = []grid.Link{{Row: 2, Col: 1, URL: "https://example.com/a"}}
	s.Charts = []grid.Chart{{Anchor: "H2", Type: "column", Series: []grid.ChartSeries{{Name: "Data!$B$1", Categories: fmt.Sprintf("Data!$A$2:$A$%d", row(5)), Values: fmt.Sprintf("Data!$B$2:$B$%d", row(5))}}}}
	s.HiddenRows = []grid.Run{{From: row(7), To: row(7)}}
	s.HiddenCols = []grid.Run{{From: 5, To: 5}}
	return wb
}

// TestReachObjects: what a sheet holds besides cells (tables, merges, notes,
// links, charts, hidden rows and columns, a sheet's visibility) must stand as
// it did outside the edit's reach, followed through the edit's shifts; and the
// objects an edit reaches may change (fuzzer-long's audit, f2.1: a table
// grown by three rows, a merge made out of nothing, a note made, were no
// finding).
func TestReachObjects(t *testing.T) {
	edit := func(args ...string) Edit {
		e := Edit{Class: args[0], Args: append([]string{args[0], FileArg}, args[1:]...)}
		switch args[0] {
		case "rows", "cols", "sheet":
			e.Class = args[0] + "-" + args[1]
			e.Args = append([]string{args[0], args[1], FileArg}, args[2:]...)
		}
		return e
	}
	far := edit("set", "Data!ZZ900", "1")
	for _, tc := range []struct {
		name    string
		edit    Edit
		after   func() *grid.Workbook
		slip    func(wb *grid.Workbook)
		wantErr string // part of the finding; "" when the edit kept to its reach
	}{
		{name: "nothing changed", edit: far},
		{name: "a table grown", edit: far, slip: func(wb *grid.Workbook) { wb.Sheet("Data").Tables[0].Range.R2 += 3 },
			wantErr: "the table Sales1 at A1:C5 is not there as it was: the edit's shifts give the table Sales1 at A1:C5, and the sheet holds table Sales1 at A1:C8"},
		{name: "a table lost", edit: far, slip: func(wb *grid.Workbook) { wb.Sheet("Data").Tables = nil },
			wantErr: "the table Sales1 at A1:C5 is not there as it was"},
		{name: "a table renamed", edit: far, slip: func(wb *grid.Workbook) { wb.Sheet("Data").Tables[0].Name = "Sales2" },
			wantErr: "the table Sales1 at A1:C5 is not there as it was"},
		{name: "a table's filter lost", edit: far, slip: func(wb *grid.Workbook) { wb.Sheet("Data").Tables[0].Filter = false },
			wantErr: "the table Sales1 at A1:C5 is not there as it was"},
		{name: "a table's column renamed by an edit far from its header", edit: far, slip: func(wb *grid.Workbook) { wb.Sheet("Data").Tables[0].Columns[1] = "Qty" },
			wantErr: "the table Sales1 at A1:C5 is not there as it was"},
		{name: "a merge made", edit: far, slip: func(wb *grid.Workbook) {
			wb.Sheet("Data").Merges = append(wb.Sheet("Data").Merges, grid.Range{R1: 1, C1: 5, R2: 2, C2: 6})
		}, wantErr: "the edit made a merge at E1:F2 it does not name"},
		{name: "a merge lost", edit: far, slip: func(wb *grid.Workbook) { wb.Sheet("Data").Merges = nil },
			wantErr: "the merge at A8:C8 is not there as it was"},
		{name: "a merge moved", edit: far, slip: func(wb *grid.Workbook) { wb.Sheet("Data").Merges[0].R1, wb.Sheet("Data").Merges[0].R2 = 9, 9 },
			wantErr: "the merge at A8:C8 is not there as it was"},
		{name: "a note made", edit: far, slip: func(wb *grid.Workbook) {
			wb.Sheet("Data").Notes = append(wb.Sheet("Data").Notes, grid.Note{Row: 1, Col: 1, Text: "slip"})
		}, wantErr: `the edit made a note "slip" on A1 it does not name`},
		{name: "a note changed", edit: far, slip: func(wb *grid.Workbook) { wb.Sheet("Data").Notes[0].Text = "other" },
			wantErr: `the note "check" on B3 is not there as it was`},
		{name: "a note lost", edit: far, slip: func(wb *grid.Workbook) { wb.Sheet("Data").Notes = nil },
			wantErr: `the note "check" on B3 is not there as it was`},
		{name: "a link moved", edit: far, slip: func(wb *grid.Workbook) { wb.Sheet("Data").Links[0].Row = 3 },
			wantErr: "the link https://example.com/a on A2 is not there as it was"},
		{name: "a chart's series changed", edit: far, slip: func(wb *grid.Workbook) { wb.Sheet("Data").Charts[0].Series[0].Values = "Data!$B$2:$B$9" },
			wantErr: "the chart column on H2 is not there as it was"},
		{name: "a chart's type changed", edit: far, slip: func(wb *grid.Workbook) { wb.Sheet("Data").Charts[0].Type = "line" },
			wantErr: "the chart column on H2 is not there as it was"},
		{name: "a chart made", edit: far, slip: func(wb *grid.Workbook) {
			wb.Sheet("Data").Charts = append(wb.Sheet("Data").Charts, wb.Sheet("Data").Charts[0])
		}, wantErr: "the edit made a chart column"},
		{name: "a row hidden", edit: far, slip: func(wb *grid.Workbook) {
			wb.Sheet("Data").HiddenRows = append(wb.Sheet("Data").HiddenRows, grid.Run{From: 3, To: 3})
		}, wantErr: "row 3 was shown, and is hidden now"},
		{name: "a row shown", edit: far, slip: func(wb *grid.Workbook) { wb.Sheet("Data").HiddenRows = nil },
			wantErr: "row 7 was hidden, and is shown now"},
		{name: "a column shown", edit: far, slip: func(wb *grid.Workbook) { wb.Sheet("Data").HiddenCols = nil },
			wantErr: "column E was hidden, and is shown now"},
		{name: "a sheet hidden", edit: far, slip: func(wb *grid.Workbook) { wb.Sheet("Data").Hidden = true },
			wantErr: "the sheet was hidden: false, and is hidden: true now"},
		{name: "a filter made", edit: far, slip: func(wb *grid.Workbook) { wb.Sheet("Data").Filter = grid.Range{R1: 1, C1: 1, R2: 5, C2: 3} },
			wantErr: "the edit made a filter at A1:C5 it does not name"},

		// Edits that reach an object may change it; one that does not name it may not.
		{name: "a set into a merge may unmerge it", edit: edit("set", "Data!A8", "x"), slip: func(wb *grid.Workbook) { wb.Sheet("Data").Merges = nil }},
		{name: "a clear over a note may drop it", edit: edit("clear", "Data!B3"), slip: func(wb *grid.Workbook) { wb.Sheet("Data").Notes = nil }},
		{name: "a clear over a note may not drop a link", edit: edit("clear", "Data!B3"), slip: func(wb *grid.Workbook) {
			wb.Sheet("Data").Notes, wb.Sheet("Data").Links = nil, nil
		}, wantErr: "the link https://example.com/a on A2 is not there as it was"},
		{name: "a set in a table does not reshape it", edit: edit("set", "Data!B2", "9"), slip: func(wb *grid.Workbook) { wb.Sheet("Data").Tables[0].Range.R2++ },
			wantErr: "the table Sales1 at A1:C5 is not there as it was"},
		{name: "a set in a table's header renames its column", edit: edit("set", "Data!B1", "Qty"), slip: func(wb *grid.Workbook) {
			wb.Sheet("Data").Tables[0].Columns[1] = "Qty"
			wb.Sheet("Data").Set(1, 2, grid.Cell{Kind: grid.Text, Text: "Qty"})
		}},
		{name: "a sheet made as a copy is new", edit: edit("sheet", "copy", "Data", "Copy"), after: func() *grid.Workbook {
			wb := objBook(0)
			s := wb.Sheet("Data")
			c := grid.NewSheet("Copy")
			for r := 1; r <= 8; r++ {
				for cc := 1; cc <= 3; cc++ {
					c.Set(r, cc, s.Get(r, cc))
				}
			}
			c.Tables = []grid.NamedTable{{Name: "Sales2", Range: s.Tables[0].Range, Columns: s.Tables[0].Columns}}
			c.Merges = s.Merges
			wb.Sheets = append(wb.Sheets, c)
			return wb
		}},

		// Shifted objects.
		{name: "rows inserted: the objects below move, a table grows", edit: edit("rows", "insert", "Data!3", "-n", "2"), after: func() *grid.Workbook { return objBook(2) }},
		{name: "rows inserted: a table not grown", edit: edit("rows", "insert", "Data!3", "-n", "2"), after: func() *grid.Workbook {
			wb := objBook(2)
			wb.Sheet("Data").Tables[0].Range.R2 = 5
			return wb
		}, wantErr: "the table Sales1 at A1:C7 is not there as it was: the edit's shifts give the table Sales1 at A1:C7, and the sheet holds table Sales1 at A1:C5"},
		{name: "rows inserted: a merge left where it was", edit: edit("rows", "insert", "Data!3", "-n", "2"), after: func() *grid.Workbook {
			wb := objBook(2)
			wb.Sheet("Data").Merges[0].R1, wb.Sheet("Data").Merges[0].R2 = 8, 8
			return wb
		}, wantErr: "the merge at A10:C10 is not there as it was"},
		{name: "rows inserted: a note left where it was", edit: edit("rows", "insert", "Data!3", "-n", "2"), after: func() *grid.Workbook {
			wb := objBook(2)
			wb.Sheet("Data").Notes[0].Row = 3
			return wb
		}, wantErr: `the note "check" on B5 is not there as it was`},
		{name: "rows inserted: a hidden row left where it was", edit: edit("rows", "insert", "Data!3", "-n", "2"), after: func() *grid.Workbook {
			wb := objBook(2)
			wb.Sheet("Data").HiddenRows = []grid.Run{{From: 7, To: 7}}
			return wb
		}, wantErr: "row 7 was shown, and is hidden now"},
		{name: "rows inserted: a chart's series not shifted", edit: edit("rows", "insert", "Data!3", "-n", "2"), after: func() *grid.Workbook {
			wb := objBook(2)
			wb.Sheet("Data").Charts[0].Series[0].Values = "Data!$B$2:$B$5"
			return wb
		}, wantErr: "the chart column on H2 is not there as it was"},
		{name: "rows inserted: the new rows are not hidden", edit: edit("rows", "insert", "Data!3", "-n", "2"), after: func() *grid.Workbook {
			wb := objBook(2)
			wb.Sheet("Data").HiddenRows = append(wb.Sheet("Data").HiddenRows, grid.Run{From: 3, To: 4})
			return wb
		}},
		{name: "rows deleted: a table shrinks, a note goes, the rest moves up", edit: edit("rows", "delete", "Data!3:4"), after: func() *grid.Workbook {
			wb := reachBook(map[string]string{"Data!A1": "'Item", "Data!B1": "'Units", "Data!C1": "'Cost", "Data!A2": "'x", "Data!B2": "2", "Data!C2": "4",
				"Data!A3": "'x", "Data!B3": "5", "Data!C3": "10", "Data!A6": "'Footer"})
			s := wb.Sheet("Data")
			s.Tables = []grid.NamedTable{{Name: "Sales1", Range: grid.Range{R1: 1, C1: 1, R2: 3, C2: 3}, Filter: true, Columns: []string{"Item", "Units", "Cost"}}}
			s.Merges = []grid.Range{{R1: 6, C1: 1, R2: 6, C2: 3}}
			s.Links = []grid.Link{{Row: 2, Col: 1, URL: "https://example.com/a"}}
			s.Charts = []grid.Chart{{Anchor: "H2", Type: "column", Series: []grid.ChartSeries{{Name: "Data!$B$1", Categories: "Data!$A$2:$A$3", Values: "Data!$B$2:$B$3"}}}}
			s.HiddenRows = []grid.Run{{From: 5, To: 5}}
			s.HiddenCols = []grid.Run{{From: 5, To: 5}}
			return wb
		}},
		{name: "columns inserted: a hidden column moves", edit: edit("cols", "insert", "Data!B", "-n", "1"), after: func() *grid.Workbook {
			wb := reachBook(map[string]string{"Data!A1": "'Item", "Data!C1": "'Units", "Data!D1": "'Cost",
				"Data!A2": "'x", "Data!C2": "2", "Data!D2": "4", "Data!A3": "'x", "Data!C3": "3", "Data!D3": "6", "Data!A4": "'x", "Data!C4": "4", "Data!D4": "8",
				"Data!A5": "'x", "Data!C5": "5", "Data!D5": "10", "Data!A8": "'Footer"})
			s := wb.Sheet("Data")
			s.Tables = []grid.NamedTable{{Name: "Sales1", Range: grid.Range{R1: 1, C1: 1, R2: 5, C2: 4}, Filter: true, Columns: []string{"Item", "Units", "Cost"}}}
			s.Merges = []grid.Range{{R1: 8, C1: 1, R2: 8, C2: 4}}
			s.Notes = []grid.Note{{Row: 3, Col: 3, Author: "Ann", Text: "check"}}
			s.Links = []grid.Link{{Row: 2, Col: 1, URL: "https://example.com/a"}}
			s.Charts = []grid.Chart{{Anchor: "I2", Type: "column", Series: []grid.ChartSeries{{Name: "Data!$C$1", Categories: "Data!$A$2:$A$5", Values: "Data!$C$2:$C$5"}}}}
			s.HiddenRows = []grid.Run{{From: 7, To: 7}}
			s.HiddenCols = []grid.Run{{From: 6, To: 6}}
			return wb
		}},
		{name: "a sheet renamed: a chart reads it by its new name", edit: edit("sheet", "rename", "Data", "Old"), after: func() *grid.Workbook {
			wb := objBook(0)
			wb.Sheets[0].Name = "Old"
			wb.Sheets[0].Charts[0].Series[0] = grid.ChartSeries{Name: "Old!$B$1", Categories: "Old!$A$2:$A$5", Values: "Old!$B$2:$B$5"}
			return wb
		}},
		{name: "a sheet renamed: a chart still reads the old name", edit: edit("sheet", "rename", "Data", "Old"), after: func() *grid.Workbook {
			wb := objBook(0)
			wb.Sheets[0].Name = "Old"
			return wb
		}, wantErr: "the chart column on H2 is not there as it was"},
	} {
		before := objBook(0)
		var after *grid.Workbook
		switch {
		case tc.after != nil:
			after = tc.after()
		default:
			after = objBook(0)
		}
		if tc.slip != nil {
			tc.slip(after)
		}
		mo := newModel(before)
		mo.r = newReach(before)
		mo.apply(tc.edit)
		got := mo.reached(after)
		if tc.wantErr == "" && got != "" || tc.wantErr != "" && !strings.Contains(got, tc.wantErr) {
			t.Errorf("%s: the model says %q, want %q", tc.name, got, tc.wantErr)
		}
	}
}

// TestReachObjectsPool: the objects the committed pool holds, read from its
// files, go through an edit that touches nothing near them as they were, and
// each slip of the audit's repro (f2.1) is found: a table grown, a merge, a
// note.
func TestReachObjectsPool(t *testing.T) {
	wb := loadPool(t)[0] // a sheet Data with a table
	before, err := xlsx.ReadBytes(wb.Data)
	if err != nil {
		t.Fatal(err)
	}
	for _, c := range []string{"", "table", "merge", "note"} {
		after, _ := xlsx.ReadBytes(wb.Data)
		s := after.Sheet("Data")
		switch c {
		case "table":
			s.Tables[0].Range.R2 += 3
		case "merge":
			s.Merges = append(s.Merges, grid.Range{R1: 1, C1: 1, R2: 2, C2: 2})
		case "note":
			s.Notes = append(s.Notes, grid.Note{Row: 1, Col: 1, Text: "slip"})
		}
		m := newModel(before)
		m.r = newReach(before)
		m.apply(Edit{Class: "set", Args: []string{"set", FileArg, "Data!ZZ900", "1"}})
		msg := m.reached(after)
		if c == "" && msg != "" || c != "" && msg == "" {
			t.Errorf("slip %q: the reach invariant says %q", c, msg)
		}
	}
}

// fileBook is objBook's file side: a validation over B2:B5 (a list), a
// conditional format over B2:C5 reading $B2, and the look of each cell.
func fileBook(n int) *fileObjects {
	row := func(r int) int {
		if r >= 3 {
			return r + n
		}
		return r
	}
	f := &fileObjects{objects: map[string][]object{}, looks: map[string]map[[2]int]string{}}
	f.objects["data"] = []object{
		{kind: "validation", rg: grid.Range{R1: 2, C1: 2, R2: row(5), C2: 2}, free: true, relative: true, cover: true, label: "list", refs: "list a,b"},
		{kind: "conditional format", rg: grid.Range{R1: 2, C1: 2, R2: row(5), C2: 3}, free: true, relative: true, cover: true, refs: "=$B2>10"},
	}
	looks := map[[2]int]string{}
	for r := 1; r <= 5; r++ {
		for c := 1; c <= 3; c++ {
			looks[[2]int{row(r), c}] = ""
		}
	}
	looks[[2]int{1, 1}] = "bold"
	f.looks["data"] = looks
	return f
}

// TestReachFile: validations, conditional formats and the look of a cell are
// judged like the rest: as they were outside the reach, followed through the
// edit; one an edit cuts into leaves what remains of it.
func TestReachFile(t *testing.T) {
	far := Edit{Class: "set", Args: []string{"set", FileArg, "Data!ZZ900", "1"}}
	for _, tc := range []struct {
		name    string
		edit    Edit
		after   func() *fileObjects
		wantErr string
	}{
		{name: "nothing changed", edit: far, after: func() *fileObjects { return fileBook(0) }},
		{name: "a validation grown", edit: far, after: func() *fileObjects {
			f := fileBook(0)
			f.objects["data"][0].rg.R2 = 9
			return f
		}, wantErr: "the validation list is not there as it was: the edit's shifts give the validation list reading list a,b over B2:B5, and the sheet holds it over B2:B9 (they differ at B6)"},
		{name: "a validation made", edit: far, after: func() *fileObjects {
			f := fileBook(0)
			f.objects["data"] = append(f.objects["data"], object{kind: "validation", rg: grid.Range{R1: 9, C1: 1, R2: 9, C2: 1}, free: true, relative: true, cover: true, refs: "whole"})
			return f
		}, wantErr: "the edit made a validation at A9 it does not name"},
		{name: "a conditional format lost", edit: far, after: func() *fileObjects {
			f := fileBook(0)
			f.objects["data"] = f.objects["data"][:1]
			return f
		}, wantErr: "the conditional format is not there as it was: the edit's shifts give the conditional format reading =$B2>10 over B2:C5, and the sheet holds it over nothing (they differ at B2)"},
		{name: "a conditional format's rule rewritten", edit: far, after: func() *fileObjects {
			f := fileBook(0)
			f.objects["data"][1].refs = "=$B3>10"
			return f
		}, wantErr: "the conditional format is not there as it was: the edit's shifts give the conditional format reading =$B2>10 over B2:C5, and the sheet holds it over nothing (they differ at B2), and the sheet holds one reading =$B3>10 over B2:C5"},
		{name: "a look changed", edit: far, after: func() *fileObjects {
			f := fileBook(0)
			f.looks["data"][[2]int{1, 1}] = ""
			return f
		}, wantErr: "Data!A1: outside the edit's reach, its look was [bold], and is [] now"},
		{name: "a look changed in the reach", edit: Edit{Class: "set", Args: []string{"set", FileArg, "Data!A1", "x"}}, after: func() *fileObjects {
			f := fileBook(0)
			f.looks["data"][[2]int{1, 1}] = ""
			return f
		}},
		{name: "a validation under a clear may go", edit: Edit{Class: "clear", Args: []string{"clear", FileArg, "Data!B2:B5"}}, after: func() *fileObjects {
			f := fileBook(0)
			f.objects["data"] = f.objects["data"][1:]
			return f
		}},
		{name: "a validation cut by a clear leaves its remainder", edit: Edit{Class: "clear", Args: []string{"clear", FileArg, "Data!B3"}}, after: func() *fileObjects {
			f := fileBook(0)
			f.objects["data"][0].rg = grid.Range{R1: 2, C1: 2, R2: 2, C2: 2}
			f.objects["data"] = append(f.objects["data"], object{kind: "validation", rg: grid.Range{R1: 4, C1: 2, R2: 5, C2: 2}, free: true, relative: true, cover: true, label: "list", refs: "list a,b"})
			return f
		}},
		{name: "a validation written as two ranges is the same", edit: far, after: func() *fileObjects {
			f := fileBook(0)
			v := f.objects["data"][0]
			v.rg = grid.Range{R1: 2, C1: 2, R2: 3, C2: 2}
			w := v
			w.rg = grid.Range{R1: 4, C1: 2, R2: 5, C2: 2}
			f.objects["data"] = append([]object{v, w}, f.objects["data"][1:]...)
			return f
		}},
		{name: "rows inserted: both move and grow", edit: Edit{Class: "rows-insert", Args: []string{"rows", "insert", FileArg, "Data!3", "-n", "2"}}, after: func() *fileObjects { return fileBook(2) }},
		{name: "rows inserted: a rule left as it was", edit: Edit{Class: "rows-insert", Args: []string{"rows", "insert", FileArg, "Data!3", "-n", "2"}}, after: func() *fileObjects {
			f := fileBook(2)
			f.objects["data"][1].rg.R2 = 5
			return f
		}, wantErr: "the conditional format is not there as it was: the edit's shifts give the conditional format reading =$B2>10 over B2:C7, and the sheet holds it over B2:C5 (they differ at B6)"},
	} {
		before := objBook(0)
		mo := newModel(before)
		mo.r = newReach(before)
		mo.apply(tc.edit)
		after := objBook(0)
		switch tc.edit.Class {
		case "rows-insert":
			after = objBook(2)
		}
		got := mo.reachedFile(after, fileBook(0), tc.after())
		if tc.wantErr == "" && got != "" || tc.wantErr != "" && !strings.Contains(got, tc.wantErr) {
			t.Errorf("%s: the model says %q, want %q", tc.name, got, tc.wantErr)
		}
	}
}

// TestReadFileObjects: the file's validations, conditional formats and looks
// are read from a pool file, and a look is each cell's own.
func TestReadFileObjects(t *testing.T) {
	var found bool
	for _, w := range loadPool(t) {
		wb, err := xlsx.ReadBytes(w.Data)
		if err != nil {
			t.Fatal(err)
		}
		f, err := readFileObjects(w.Data, wb)
		if err != nil {
			t.Fatalf("%s: %v", w.Name, err)
		}
		for k, s := range wb.Sheets {
			_ = k
			if u := s.Used(); u.R2 >= 1 && len(f.looks[strings.ToLower(s.Name)]) == 0 {
				t.Errorf("%s: sheet %s holds cells and no look is read", w.Name, s.Name)
			}
		}
		for _, os := range f.objects {
			for _, o := range os {
				found = found || o.kind == "validation" || o.kind == "conditional format"
			}
		}
	}
	if !found {
		t.Error("no pool file's validation or conditional format is read")
	}
}

// TestReachObjectsKept: a link's place is the text it was typed as (Excel does
// not follow a hyperlink to a cell when rows are inserted or a sheet is
// renamed), a merge cut to one cell goes, and a note's author is not judged (the long run's first findings of objects).
func TestReachObjectsKept(t *testing.T) {
	judge := func(edit Edit, mutate func(wb *grid.Workbook), after *grid.Workbook) string {
		before := objBook(0)
		mutate(before)
		mo := newModel(before)
		mo.r = newReach(before)
		mo.apply(edit)
		return mo.reached(after)
	}
	link := func(wb *grid.Workbook) {
		wb.Sheet("Data").Links = append(wb.Sheet("Data").Links, grid.Link{Row: 9, Col: 1, Location: "Data!B2"})
	}
	// A rename leaves the link as it was.
	after := objBook(0)
	link(after)
	after.Sheets[0].Name = "Old"
	after.Sheets[0].Charts[0].Series[0] = grid.ChartSeries{Name: "Old!$B$1", Categories: "Old!$A$2:$A$5", Values: "Old!$B$2:$B$5"}
	if got := judge(Edit{Class: "sheet-rename", Args: []string{"sheet", "rename", FileArg, "Data", "Old"}}, link, after); got != "" {
		t.Errorf("a link's place after a rename: %s", got)
	}
	// Rows inserted above it: it moves, and its place does not.
	after = objBook(2)
	after.Sheet("Data").Links = append(after.Sheet("Data").Links, grid.Link{Row: 11, Col: 1, Location: "Data!B2"})
	if got := judge(Edit{Class: "rows-insert", Args: []string{"rows", "insert", FileArg, "Data!3", "-n", "2"}}, link, after); got != "" {
		t.Errorf("a link's place after rows are inserted: %s", got)
	}
	// A merge over three columns, two of them deleted, is one cell: gone.
	cut := objBook(0)
	cut.Sheet("Data").Merges = nil
	if got := judge(Edit{Class: "cols-delete", Args: []string{"cols", "delete", FileArg, "Data!B:C"}}, func(wb *grid.Workbook) {}, cut); strings.Contains(got, "merge") {
		t.Errorf("a merge cut to one cell: %s", got)
	}
	// A note's author is not judged (the full editor writes "Author"), its text is.
	other := objBook(0)
	other.Sheet("Data").Notes[0].Author = "Author"
	if got := judge(Edit{Class: "set", Args: []string{"set", FileArg, "Data!ZZ900", "1"}}, func(wb *grid.Workbook) {}, other); got != "" {
		t.Errorf("a note by another author: %q", got)
	}
	other.Sheet("Data").Notes[0].Text = "other"
	got := judge(Edit{Class: "set", Args: []string{"set", FileArg, "Data!ZZ900", "1"}}, func(wb *grid.Workbook) {}, other)
	for _, want := range []string{`the note "check" on B3 is not there as it was`, `the sheet holds note "other" on B3`} {
		if !strings.Contains(got, want) {
			t.Errorf("a note with other text: %q lacks %q", got, want)
		}
	}
	// And a note gone says so.
	gone := objBook(0)
	gone.Sheet("Data").Notes = nil
	if got := judge(Edit{Class: "set", Args: []string{"set", FileArg, "Data!ZZ900", "1"}}, func(wb *grid.Workbook) {}, gone); !strings.Contains(got, "and the sheet holds nothing of its kind") {
		t.Errorf("a note lost: %q", got)
	}
}

// TestCompareFileObjects: the editors invariant compares the validations,
// conditional formats and looks of the two files: the cells a rule covers (a
// sqref written as two ranges is the same), its rule, and each cell's look.
func TestCompareFileObjects(t *testing.T) {
	name := func(k string) string { return "Data" }
	two := func(f *fileObjects) *fileObjects {
		v := f.objects["data"][0]
		v.rg = grid.Range{R1: 2, C1: 2, R2: 3, C2: 2}
		w := v
		w.rg = grid.Range{R1: 4, C1: 2, R2: 5, C2: 2}
		f.objects["data"] = append([]object{v, w}, f.objects["data"][1:]...)
		return f
	}
	for _, tc := range []struct {
		name    string
		other   func() *fileObjects
		wantErr string
	}{
		{"the same", func() *fileObjects { return fileBook(0) }, ""},
		{"a range in two", func() *fileObjects { return two(fileBook(0)) }, ""},
		{"a validation over fewer cells", func() *fileObjects {
			f := fileBook(0)
			f.objects["data"][0].rg.R2 = 4
			return f
		}, `Data!validation "list a,b": the streaming editor's file holds it over B2:B5, the full editor's over B2:B4 (they differ at B5)`},
		{"a rule read another way", func() *fileObjects {
			f := fileBook(0)
			f.objects["data"][1].refs = "=$B3>10"
			return f
		}, `Data!conditional format "=$B2>10": the streaming editor's file holds it over B2:C5, the full editor's over nothing (they differ at B2)`},
		{"a look", func() *fileObjects {
			f := fileBook(0)
			f.looks["data"][[2]int{1, 1}] = "bold fill=#DDEBF7"
			return f
		}, `Data!A1: the streaming editor's file shows it as [bold], the full editor's as [bold fill=#DDEBF7]`},
	} {
		got := compareFileObjects(fileBook(0), tc.other(), name)
		if tc.wantErr == "" && got != "" || tc.wantErr != "" && !strings.Contains(got, tc.wantErr) {
			t.Errorf("%s: %q, want %q", tc.name, got, tc.wantErr)
		}
	}
}

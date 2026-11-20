package fuzz

import (
	"sort"
	"strings"
	"testing"

	"github.com/eemax/sheets/internal/grid"
)

// formulaBook builds a workbook that holds formulas of the given kinds at
// the cells named ("Data!G2": Plain), and a value in each cell named in
// values.
func formulaBook(cells map[string]string, tables ...string) *grid.Workbook {
	wb := &grid.Workbook{}
	sheet := func(name string) *grid.Sheet {
		if s := wb.Sheet(name); s != nil {
			return s
		}
		s := grid.NewSheet(name)
		wb.Sheets = append(wb.Sheets, s)
		return s
	}
	var names []string
	for ref := range cells {
		names = append(names, ref)
	}
	sort.Strings(names)
	for _, ref := range names {
		name, cell, _ := strings.Cut(ref, "!")
		r, c, _ := grid.ParseCell(cell)
		kind := cells[ref]
		s := sheet(name)
		if kind == "" {
			s.Set(r, c, grid.Cell{Kind: grid.Number, Num: 1})
			continue
		}
		s.Set(r, c, grid.Cell{Kind: grid.Number, Num: 1, Formula: `IF(N("` + kind + `")=0,B2)`})
	}
	for _, t := range tables { // "Data!A1:C5"
		name, rg, _ := strings.Cut(t, "!")
		r, _ := grid.ParseRange(rg)
		sheet(name).Tables = append(sheet(name).Tables, grid.NamedTable{Name: "T", Range: r})
	}
	return wb
}

// TestModel: the model of where formulas go follows each edit, and refuses a
// saved workbook that lost a formula, kept one in a cell an edit emptied, or
// made one the edit did not.
func TestModel(t *testing.T) {
	P, L, D := Plain, Legacy, Dynamic
	edit := func(args ...string) Edit {
		e := Edit{Class: args[0], Args: append([]string{args[0], FileArg}, args[1:]...)}
		switch args[0] {
		case "rows", "cols":
			e.Class = args[0] + "-" + args[1]
			e.Args = append([]string{args[0], args[1], FileArg}, args[2:]...)
		case "sheet":
			e.Class = "sheet-" + args[1]
			e.Args = append([]string{args[0], args[1], FileArg}, args[2:]...)
		}
		return e
	}
	type m = map[string]string
	for _, tc := range []struct {
		name          string
		before        m
		edit          Edit
		after         m
		tables        []string
		want, wantErr string // wantErr: part of the first difference; "" when the model holds
	}{
		{name: "copy", before: m{"Data!G2": P, "Data!G3": L, "Data!G4": D, "Data!A1": ""}, edit: edit("copy", "Data!G2:G4", "Data!W2"),
			after: m{"Data!G2": P, "Data!G3": L, "Data!G4": D, "Data!W2": P, "Data!W3": L, "Data!W4": D, "Data!A1": ""}},
		{name: "copy loses the second row", before: m{"Data!G2": P, "Data!G3": P, "Data!G4": P}, edit: edit("copy", "Data!G2:G4", "Data!W2"),
			after: m{"Data!G2": P, "Data!G3": P, "Data!G4": P, "Data!W2": P, "Data!W4": P}, wantErr: "Data!W3: the edit should leave a plain formula there, and it holds no formula (it was lost)"},
		{name: "copy makes a formula a dynamic array", before: m{"Data!G2": P}, edit: edit("copy", "Data!G2", "Data!W2"),
			after: m{"Data!G2": P, "Data!W2": D}, wantErr: "Data!W2: the edit should leave a plain formula there, and it holds a dynamic array"},
		{name: "copy over a formula from a cell with none", before: m{"Data!A1": P, "Data!B1": ""}, edit: edit("copy", "Data!B1", "Data!A1"),
			after: m{"Data!B1": ""}},
		{name: "copy over a formula from a cell with none, kept", before: m{"Data!A1": P, "Data!B1": ""}, edit: edit("copy", "Data!B1", "Data!A1"),
			after: m{"Data!A1": P, "Data!B1": ""}, wantErr: "Data!A1: the edit should leave no formula there"},
		{name: "copy --values", before: m{"Data!G2": P}, edit: edit("copy", "Data!G2", "Data!W2", "--values"),
			after: m{"Data!G2": P}},
		{name: "copy --values pasting a formula", before: m{"Data!G2": P}, edit: edit("copy", "Data!G2", "Data!W2", "--values"),
			after: m{"Data!G2": P, "Data!W2": P}, wantErr: "Data!W2: the edit should leave no formula there"},
		{name: "copy --format", before: m{"Data!G2": P}, edit: edit("copy", "Data!G2", "Data!W2", "--format"),
			after: m{"Data!G2": P}},
		{name: "copy to another sheet, new", before: m{"Data!G2": L}, edit: edit("copy", "Data!G2", "Plan!A1"),
			after: m{"Data!G2": L, "Plan!A1": L}},
		{name: "move", before: m{"Data!B3": P, "Data!B4": D, "Data!C1": P}, edit: edit("move", "Data!B3:B4", "Data!Y30"),
			after: m{"Data!Y30": P, "Data!Y31": D, "Data!C1": P}},
		{name: "move leaves one behind", before: m{"Data!B3": P}, edit: edit("move", "Data!B3", "Data!Y30"),
			after: m{"Data!B3": P, "Data!Y30": P}, wantErr: "Data!B3: the edit should leave no formula there"},
		{name: "move onto itself shifted", before: m{"Data!B3": P, "Data!B4": P}, edit: edit("move", "Data!B3:B4", "Data!B4"),
			after: m{"Data!B4": P, "Data!B5": P}},
		{name: "clear", before: m{"Data!B3": P, "Data!B4": L, "Data!C1": P}, edit: edit("clear", "Data!B3:B4"),
			after: m{"Data!C1": P}},
		{name: "clear keeps one", before: m{"Data!B3": P}, edit: edit("clear", "Data!B3"),
			after: m{"Data!B3": P}, wantErr: "Data!B3: the edit should leave no formula there"},
		{name: "set a formula and a value", before: m{"Data!B3": P, "Data!B4": P}, edit: edit("set", "Data!B3", "5", "Data!C1", `=IF(N("d")=0,SEQUENCE(2))`),
			after: m{"Data!B4": P, "Data!C1": D}},
		{name: "set a formula that is not kept", before: m{}, edit: edit("set", "Data!C1", `=IF(N("d")=0,SEQUENCE(2))`),
			after: m{"Data!A1": ""}, wantErr: "Data!C1: the edit should leave a dynamic array there, and it holds no formula"},
		{name: "set a formula in a table's header", before: m{}, edit: edit("set", "Data!A1", `=IF(N("p")=0,B2)`),
			after: m{}, tables: []string{"Data!A1:C5"}},
		{name: "set a range", before: m{}, edit: edit("set", "Data!C1:C3", `=IF(N("p")=0,B2)`),
			after: m{"Data!C1": P, "Data!C2": P, "Data!C3": P}},
		{name: "write", before: m{"Data!B3": P}, edit: Edit{Class: "write", Args: []string{"write", FileArg, "Data!A3", "--from", "tsv"}, Stdin: "1\t=IF(N(\"p\")=0,B2)\n2\n"},
			after: m{"Data!B3": P}},
		{name: "replace", before: m{"Data!B3": P, "Data!C3": L}, edit: edit("replace", "a", "b"),
			after: m{"Data!B3": P, "Data!C3": L}},
		{name: "replace loses one", before: m{"Data!B3": P, "Data!C3": L}, edit: edit("replace", "a", "b"),
			after: m{"Data!B3": P}, wantErr: "Data!C3: the edit should leave a legacy array there, and it holds no formula"},
		{name: "rows insert", before: m{"Data!B3": P, "Data!B4": L, "Data!C1": P}, edit: edit("rows", "insert", "Data!4", "-n", "2"),
			after: m{"Data!B3": P, "Data!B6": L, "Data!C1": P}},
		{name: "rows delete", before: m{"Data!B3": P, "Data!B4": L, "Data!B5": P, "Data!B8": D}, edit: edit("rows", "delete", "Data!4:5"),
			after: m{"Data!B3": P, "Data!B6": D}},
		{name: "rows delete shifts by one", before: m{"Data!B3": P, "Data!B8": D}, edit: edit("rows", "delete", "Data!4:5"),
			after: m{"Data!B3": P, "Data!B7": D}, wantErr: "Data!B6: the edit should leave a dynamic array there"},
		{name: "cols insert", before: m{"Data!B3": P, "Data!D3": L}, edit: edit("cols", "insert", "Data!C", "-n", "1"),
			after: m{"Data!B3": P, "Data!E3": L}},
		{name: "cols delete", before: m{"Data!B3": P, "Data!C3": P, "Data!D3": L}, edit: edit("cols", "delete", "Data!B:C"),
			after: m{"Data!B3": L}},
		{name: "sheet rename", before: m{"Data!B3": P, "Plan!A1": L}, edit: edit("sheet", "rename", "Data", "Old"),
			after: m{"Old!B3": P, "Plan!A1": L}},
		{name: "sheet copy", before: m{"Data!B3": P, "Data!B4": L}, edit: edit("sheet", "copy", "Data", "Data 2"),
			after: m{"Data!B3": P, "Data!B4": L, "Data 2!B3": P, "Data 2!B4": L}},
		{name: "sheet copy loses one", before: m{"Data!B3": P, "Data!B4": L}, edit: edit("sheet", "copy", "Data", "Data 2"),
			after: m{"Data!B3": P, "Data!B4": L, "Data 2!B3": P}, wantErr: "Data 2!B4: the edit should leave a legacy array there"},
		{name: "sort moves formulas in their columns", before: m{"Data!B2": P, "Data!B3": "", "Data!C2": L},
			edit:  edit("sort", "Data", "B"),
			after: m{"Data!B3": P, "Data!B2": "", "Data!C2": L}},
		{name: "sort loses a formula", before: m{"Data!B2": P, "Data!B3": P},
			edit:  edit("sort", "Data!A1:C3", "B"),
			after: m{"Data!B3": P}, wantErr: "after a sort column B holds 1 formulas of a plain formula, it held 2"},
		{name: "sort changes a kind", before: m{"Data!B2": P, "Data!B3": P},
			edit:  edit("sort", "Data!A1:C3", "B"),
			after: m{"Data!B3": P, "Data!B2": D}, wantErr: "after a sort column B holds"},
		{name: "a batch makes a sheet and writes to it",
			before: m{"Data!B3": P},
			edit:   Edit{Class: "do", Args: []string{"do", FileArg}, Stdin: "sheet copy Data Copy\nset Copy!B3 5\nset Copy!C8 '=IF(N(\"l\")=0,B2)'\ncopy Copy!C8 Copy!D8\n"},
			after:  m{"Data!B3": P, "Copy!C8": L, "Copy!D8": L}},
		{name: "a batch loses the copy",
			before: m{"Data!B3": P},
			edit:   Edit{Class: "do", Args: []string{"do", FileArg}, Stdin: "sheet copy Data Copy\nset Copy!C8 '=IF(N(\"p\")=0,B2)'\n"},
			after:  m{"Data!B3": P, "Copy!B3": P}, wantErr: "Copy!C8: the edit should leave a plain formula there"},
		{name: "a batch sorts and is not followed",
			before: m{"Data!B2": P, "Data!B3": ""},
			edit:   Edit{Class: "do", Args: []string{"do", FileArg}, Stdin: "sort Data B\nset Data!F1 5\ncopy Data!B2 Data!E1\n"},
			after:  m{"Data!B3": P, "Data!E1": D}},
		{name: "a script",
			before: m{"Data!B3": P, "Data!B4": L},
			edit: Edit{Class: "script", Args: []string{"script", FileArg, "-e", "const ws = workbook.getWorksheet(\"Data\");\n" +
				"ws.getRange(\"A1\").setFormula(\"=IF(N(\\\"d\\\")=0,SEQUENCE(2))\");\n" +
				"ws.getRange(\"B3\").moveTo(\"C3\");\n" +
				"ws.getRange(\"E1\").copyFrom(\"A1:A2\");\n" +
				"ws.getRange(\"2:2\").insert(\"Down\");\n" +
				"ws.getRange(\"C:C\").delete(\"Left\");\n"}},
			after: m{"Data!A1": D, "Data!D1": D, "Data!B5": L}},
		{name: "a script loses one",
			before: m{"Data!B3": P},
			edit: Edit{Class: "script", Args: []string{"script", FileArg, "-e", "const ws = workbook.getWorksheet(\"Data\");\n" +
				"ws.getRange(\"B3\").copyFrom(\"B3\");\n" +
				"ws.getRange(\"E1\").setValue(5);\n" +
				"ws.getRange(\"2:2\").delete(\"Up\");\n"}},
			after: m{"Data!B3": P}, wantErr: "Data!B2: the edit should leave a plain formula there, and it holds no formula"},
		{name: "sort by an Excel table's name", before: m{"Data!B3": P, "Data!B4": L}, tables: []string{"Data!A1:C5"}, edit: edit("sort", "T", "B"),
			after: m{"Data!B3": P, "Data!B4": L}},
		{name: "sort by an Excel table's name loses one", before: m{"Data!B3": P, "Data!B4": L}, tables: []string{"Data!A1:C5"}, edit: edit("sort", "T", "B"),
			after: m{"Data!B3": P}, wantErr: "Data"},
		{name: "a command with no rule is not judged", before: m{"Data!B3": P}, edit: edit("fmt"), after: m{}},
	} {
		t.Run(tc.name, func(t *testing.T) {
			before := formulaBook(tc.before, tc.tables...)
			mo := newModel(before)
			mo.apply(tc.edit)
			got := mo.check(formulaBook(tc.after, tc.tables...))
			if tc.wantErr == "" && got != "" {
				t.Errorf("the model refuses a right result: %s", got)
			}
			if tc.wantErr != "" && !strings.Contains(got, tc.wantErr) {
				t.Errorf("the model says %q, want it to say %q", got, tc.wantErr)
			}
		})
	}
}

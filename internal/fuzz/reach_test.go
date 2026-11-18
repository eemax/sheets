package fuzz

import (
	"sort"
	"strconv"
	"strings"
	"testing"

	"github.com/eemax/sheets/internal/grid"
)

// reachBook builds a workbook from cells named "Sheet!A1": "5" is a number,
// "'x" a text, "=F" a formula with no result (pending), "=F|7" one whose
// result is 7. spills ("Data!M2:M4") are dynamic arrays, names "Units=Data!B2:B4"
// (one scoped to a sheet: "Rate@Copy=Copy!$B$2").
func reachBook(cells map[string]string, extra ...string) *grid.Workbook {
	wb := &grid.Workbook{}
	sheet := func(name string) *grid.Sheet {
		if s := wb.Sheet(name); s != nil {
			return s
		}
		s := grid.NewSheet(name)
		wb.Sheets = append(wb.Sheets, s)
		return s
	}
	if len(cells) == 0 {
		sheet("Data")
	}
	var refs []string
	for ref := range cells {
		refs = append(refs, ref)
	}
	sort.Strings(refs)
	for _, ref := range refs {
		name, at, _ := strings.Cut(ref, "!")
		r, c, _ := grid.ParseCell(at)
		v := cells[ref]
		var cl grid.Cell
		switch {
		case strings.HasPrefix(v, "="):
			f, res, ok := strings.Cut(v[1:], "|")
			cl = grid.Cell{Formula: f, Pending: !ok}
			if ok {
				cl.Kind, cl.Num = grid.Number, mustFloat(res)
			}
		case strings.HasPrefix(v, "'"):
			cl = grid.Cell{Kind: grid.Text, Text: v[1:]}
		default:
			cl = grid.Cell{Kind: grid.Number, Num: mustFloat(v)}
		}
		sheet(name).Set(r, c, cl)
	}
	for _, x := range extra {
		if n, ref, ok := strings.Cut(x, "="); ok {
			n, scope, _ := strings.Cut(n, "@")
			wb.Names = append(wb.Names, grid.DefinedName{Name: n, Ref: ref, Scope: scope})
			continue
		}
		name, rg, _ := strings.Cut(x, "!")
		r, _ := grid.ParseRange(rg)
		sheet(name).Spills = append(sheet(name).Spills, r)
	}
	return wb
}

func mustFloat(s string) float64 {
	f, err := strconv.ParseFloat(s, 64)
	if err != nil {
		panic(err)
	}
	return f
}

// TestReach: the model of an edit's reach passes what an edit may change
// (its cells, references it rewrites, recomputed dependents, arrays that
// fill again) and refuses anything else that changed: a cell it does not
// name, one left behind, a formula rewritten or recomputed with nothing it
// reads changed, a sheet made out of nothing.
func TestReach(t *testing.T) {
	edit := func(args ...string) Edit {
		e := Edit{Class: args[0], Args: append([]string{args[0], FileArg}, args[1:]...)}
		switch args[0] {
		case "rows", "cols", "sheet":
			e.Class = args[0] + "-" + args[1]
			e.Args = append([]string{args[0], args[1], FileArg}, args[2:]...)
		}
		return e
	}
	script := func(lines ...string) Edit {
		return Edit{Class: "script", Args: []string{"script", FileArg, "-e", "const ws = workbook.getWorksheet(\"Data\");\n" + strings.Join(lines, "\n") + "\n"}}
	}
	type m = map[string]string
	for _, tc := range []struct {
		name          string
		before, after m
		extra         []string // spills and names, before and after alike
		afterExtra    []string // the spills and names after, when other
		edit          Edit
		wantErr       string // part of the finding; "" when the edit kept to its reach
	}{
		{name: "set", before: m{"Data!A1": "1", "Data!B1": "2"}, edit: edit("set", "Data!A1", "5"),
			after: m{"Data!A1": "5", "Data!B1": "2"}},
		{name: "set writes a neighbour", before: m{"Data!A1": "1", "Data!B1": "2"}, edit: edit("set", "Data!A1", "5"),
			after: m{"Data!A1": "5", "Data!B1": "3"}, wantErr: "Data!B1: outside the edit's reach, it held number 2, and holds number 3 now"},
		{name: "set fills an empty cell it does not name", before: m{"Data!A1": "1"}, edit: edit("set", "Data!A1", "5"),
			after: m{"Data!A1": "5", "Data!Z9": "'x"}, wantErr: `Data!Z9: outside the edit's reach, it was empty, and holds text "x" now`},
		{name: "set loses a cell", before: m{"Data!A1": "1", "Data!C3": "'x"}, edit: edit("set", "Data!A1", "5"),
			after: m{"Data!A1": "5"}, wantErr: `Data!C3: outside the edit's reach, it held text "x", and holds nothing now`},
		{name: "set a range", before: m{"Data!A1": "1", "Data!A3": "1", "Data!A4": "1"}, edit: edit("set", "Data!A1:A3", "5"),
			after: m{"Data!A1": "5", "Data!A2": "5", "Data!A3": "5", "Data!A4": "1"}},
		{name: "a dependent recomputed", before: m{"Data!A1": "1", "Data!B1": "=A1*2|2", "Data!C1": "=SUM(B1:B2)|2"}, edit: edit("set", "Data!A1", "5"),
			after: m{"Data!A1": "5", "Data!B1": "=A1*2|10", "Data!C1": "=SUM(B1:B2)|10"}},
		{name: "a formula recomputed with nothing it reads changed", before: m{"Data!A1": "1", "Data!B1": "=C1*2|2", "Data!C1": "1"}, edit: edit("set", "Data!A1", "5"),
			after: m{"Data!A1": "5", "Data!B1": "=C1*2|4", "Data!C1": "1"}, wantErr: "Data!B1: outside the edit's reach, it held =C1*2 (number 2), and holds =C1*2 (number 4) now, and nothing it reads may have changed"},
		{name: "a pending result computed", before: m{"Data!A1": "1", "Data!B1": "=C1*2", "Data!C1": "1"}, edit: edit("set", "Data!A1", "5"),
			after: m{"Data!A1": "5", "Data!B1": "=C1*2|2", "Data!C1": "1"}},
		{name: "a formula rewritten by an edit that moves nothing", before: m{"Data!A1": "1", "Data!B1": "=C1*2|2"}, edit: edit("set", "Data!A1", "5"),
			after: m{"Data!A1": "5", "Data!B1": "=C2*2|2"}, wantErr: "its formula =C1*2 is rewritten as =C2*2 by an edit that moves no cells"},
		{name: "a formula lost", before: m{"Data!A1": "1", "Data!B1": "=C1*2|2"}, edit: edit("set", "Data!A1", "5"),
			after: m{"Data!A1": "5", "Data!B1": "2"}, wantErr: "Data!B1: outside the edit's reach, it held =C1*2 (number 2), and holds number 2 now"},
		{name: "a value turned formula", before: m{"Data!A1": "1", "Data!B1": "2"}, edit: edit("set", "Data!A1", "5"),
			after: m{"Data!A1": "5", "Data!B1": "=1+1|2"}, wantErr: "Data!B1: outside the edit's reach, it held number 2, and holds =1+1 (number 2) now"},
		{name: "a name's input", before: m{"Data!B2": "1", "Data!H1": "=SUM(Units)|1"}, extra: []string{"Units=Data!$B$2:$B$4"}, edit: edit("set", "Data!B3", "5"),
			after: m{"Data!B2": "1", "Data!B3": "5", "Data!H1": "=SUM(Units)|6"}},
		{name: "a name's range shrunk by rows deleted", before: m{"Data!B2": "1", "Data!B3": "5", "Data!H1": "=SUM(Units)|6"}, extra: []string{"Units=Data!$B$2:$B$4"},
			afterExtra: []string{"Units=Data!$B$2:$B$3"}, edit: edit("rows", "delete", "Data!3"),
			after: m{"Data!B2": "1", "Data!H1": "=SUM(Units)|1"}},
		{name: "a name's value with its range unchanged", before: m{"Data!B2": "1", "Data!H1": "=SUM(Units)|1", "Data!Z1": "3"}, extra: []string{"Units=Data!$B$2:$B$4"}, edit: edit("set", "Data!Z1", "5"),
			after: m{"Data!B2": "1", "Data!H1": "=SUM(Units)|9", "Data!Z1": "5"}, wantErr: "Data!H1: outside the edit's reach"},
		// Defined names: kept, or rewritten as the edit's shifts put them
		// (fuzzer-long's audit, f1.3: a set that rewrote Units, and the
		// SUM(Units) it changed, were no finding).
		{name: "a set rewrites a name", before: m{"Data!B2": "1", "Data!B3": "5", "Data!H1": "=SUM(Units)|1", "Data!Z1": "3"}, extra: []string{"Units=Data!$B$2:$B$2"},
			afterExtra: []string{"Units=Data!$B$2:$B$3"}, edit: edit("set", "Data!Z1", "5"),
			after:   m{"Data!B2": "1", "Data!B3": "5", "Data!H1": "=SUM(Units)|6", "Data!Z1": "5"},
			wantErr: "the defined name Units, outside the edit's reach, referred to =Data!$B$2:$B$2, and refers to =Data!$B$2:$B$3 now, after an edit that moves no cells"},
		{name: "a set loses a name", before: m{"Data!B2": "1", "Data!Z1": "3"}, extra: []string{"Rate=Data!$B$2"},
			afterExtra: []string{}, edit: edit("set", "Data!Z1", "5"), after: m{"Data!B2": "1", "Data!Z1": "5"},
			wantErr: "the defined name Rate, outside the edit's reach, is gone (it referred to =Data!$B$2)"},
		{name: "a set defines a name", before: m{"Data!B2": "1", "Data!Z1": "3"}, extra: []string{"Rate=Data!$B$2"},
			afterExtra: []string{"Rate=Data!$B$2", "Extra=Data!$Z$1"}, edit: edit("set", "Data!Z1", "5"), after: m{"Data!B2": "1", "Data!Z1": "5"},
			wantErr: "the edit defined the name Extra (=Data!$Z$1), and names no such name"},
		{name: "rows insert grows a name", before: m{"Data!B2": "1", "Data!B3": "5", "Data!H1": "=SUM(Units)|6"}, extra: []string{"Units=Data!$B$2:$B$3"},
			afterExtra: []string{"Units=Data!$B$2:$B$4"}, edit: edit("rows", "insert", "Data!3"),
			after: m{"Data!B2": "1", "Data!B4": "5", "Data!H1": "=SUM(Units)|6"}},
		{name: "rows insert leaves a name", before: m{"Data!B2": "1", "Data!B3": "5", "Data!H1": "=SUM(Units)|6"}, extra: []string{"Units=Data!$B$2:$B$3"},
			edit: edit("rows", "insert", "Data!1"), after: m{"Data!B3": "1", "Data!B4": "5", "Data!H2": "=SUM(Units)|5"},
			wantErr: "the defined name Units, outside the edit's reach, refers to =Data!$B$2:$B$3 as it did, and the edit's shift of references gives =Data!$B$3:$B$4"},
		{name: "rows insert on another sheet shifts a name", before: m{"Data!B2": "1", "Plan!A1": "1"}, extra: []string{"Rate=Data!$B$2"},
			afterExtra: []string{"Rate=Data!$B$3"}, edit: edit("rows", "insert", "Plan!1"), after: m{"Data!B2": "1", "Plan!A2": "1"},
			wantErr: "the defined name Rate, outside the edit's reach, referred to =Data!$B$2, and refers to =Data!$B$3 now, and the edit's shift of references gives =Data!$B$2"},
		{name: "a move takes a name along", before: m{"Data!B2": "1", "Data!H1": "=Rate*2|2"}, extra: []string{"Rate=Data!$B$2"},
			afterExtra: []string{"Rate=Data!$E$2"}, edit: edit("move", "Data!B2", "Data!E2"), after: m{"Data!E2": "1", "Data!H1": "=Rate*2|2"}},
		{name: "a sheet renamed renames a name's references", before: m{"Data!B2": "1"}, extra: []string{"Rate='Data'!$B$2"},
			afterExtra: []string{"Rate=Old!$B$2"}, edit: edit("sheet", "rename", "Data", "Old"), after: m{"Old!B2": "1"}},
		{name: "a sheet copied takes its names along", before: m{"Data!B2": "1"}, extra: []string{"Rate=Data!$B$2"},
			afterExtra: []string{"Rate=Data!$B$2", "Rate@Copy=Copy!$B$2"}, edit: edit("sheet", "copy", "Data", "Copy"), after: m{"Data!B2": "1", "Copy!B2": "1"}},
		{name: "a sheet copied, and a name of another sheet", before: m{"Data!B2": "1", "Plan!A1": "1"}, extra: []string{"Rate=Data!$B$2"},
			afterExtra: []string{"Rate=Data!$B$2", "Rate@Plan=Copy!$B$2"}, edit: edit("sheet", "copy", "Data", "Copy"), after: m{"Data!B2": "1", "Copy!B2": "1", "Plan!A1": "1"},
			wantErr: "the edit defined the name Rate (of Plan) (=Copy!$B$2)"},
		{name: "clear", before: m{"Data!A1": "1", "Data!A2": "2", "Data!A3": "3"}, edit: edit("clear", "Data!A1:A2"),
			after: m{"Data!A3": "3"}},
		{name: "clear one more", before: m{"Data!A1": "1", "Data!A2": "2", "Data!A3": "3"}, edit: edit("clear", "Data!A1:A2"),
			after: m{}, wantErr: "Data!A3: outside the edit's reach, it held number 3, and holds nothing now"},
		{name: "write", before: m{"Data!A1": "1", "Data!C2": "3"}, edit: Edit{Class: "write", Args: []string{"write", FileArg, "Data!A1", "--from", "tsv"}, Stdin: "7\t8\n9\n"},
			after: m{"Data!A1": "7", "Data!B1": "8", "Data!A2": "9", "Data!C2": "3"}},
		{name: "write past its block", before: m{"Data!A1": "1"}, edit: Edit{Class: "write", Args: []string{"write", FileArg, "Data!A1", "--from", "tsv"}, Stdin: "7\t8\n9\n"},
			after: m{"Data!A1": "7", "Data!B1": "8", "Data!A2": "9", "Data!B2": "8"}, wantErr: "Data!B2: outside the edit's reach, it was empty"},
		{name: "copy", before: m{"Data!A1": "1", "Data!A2": "2", "Data!D1": "4"}, edit: edit("copy", "Data!A1:A2", "Data!C1"),
			after: m{"Data!A1": "1", "Data!A2": "2", "Data!C1": "1", "Data!C2": "2", "Data!D1": "4"}},
		{name: "copy moves", before: m{"Data!A1": "1", "Data!A2": "2"}, edit: edit("copy", "Data!A1:A2", "Data!C1"),
			after: m{"Data!C1": "1", "Data!C2": "2"}, wantErr: "Data!A1: outside the edit's reach, it held number 1, and holds nothing now"},
		{name: "move, and what reads it follows", before: m{"Data!A1": "1", "Data!B5": "=A1*2|2"}, edit: edit("move", "Data!A1", "Data!C1"),
			after: m{"Data!C1": "1", "Data!B5": "=C1*2|2"}},
		{name: "move rewrites a formula other than its references", before: m{"Data!A1": "1", "Data!B5": "=A1*2|2"}, edit: edit("move", "Data!A1", "Data!C1"),
			after: m{"Data!C1": "1", "Data!B5": "=C1*3|3"}, wantErr: "its formula =A1*2 is rewritten as =C1*3, other than in its references"},
		{name: "rows insert", before: m{"Data!A1": "1", "Data!A2": "2", "Data!A3": "3"}, edit: edit("rows", "insert", "Data!2", "-n", "2"),
			after: m{"Data!A1": "1", "Data!A4": "2", "Data!A5": "3"}},
		{name: "rows insert, and the rows inserted filled", before: m{"Data!A1": "1", "Data!A2": "2"}, edit: edit("rows", "insert", "Data!2"),
			after: m{"Data!A1": "1", "Data!B2": "=1|1", "Data!A3": "2"}},
		{name: "rows insert leaves a row behind", before: m{"Data!A1": "1", "Data!A2": "2", "Data!A3": "3"}, edit: edit("rows", "insert", "Data!2", "-n", "2"),
			after: m{"Data!A1": "1", "Data!A4": "2", "Data!A3": "3"}, wantErr: "Data!A5: outside the edit's reach, it held number 3, and holds nothing now"},
		{name: "rows delete, and the references", before: m{"Data!A1": "1", "Data!A2": "2", "Data!A3": "3", "Data!B9": "=SUM(A1:A3)|6", "Data!C9": "=A2|2"}, edit: edit("rows", "delete", "Data!2"),
			after: m{"Data!A1": "1", "Data!A2": "3", "Data!B8": "=SUM(A1:A2)|4", "Data!C8": "=#REF!|0"}},
		{name: "rows delete shifts by one too many", before: m{"Data!A1": "1", "Data!A3": "3", "Data!A4": "4"}, edit: edit("rows", "delete", "Data!2"),
			after: m{"Data!A1": "1", "Data!A2": "4"}, wantErr: "Data!A2: outside the edit's reach, it held number 3, and holds number 4 now"},
		{name: "cols delete", before: m{"Data!A1": "1", "Data!B1": "2", "Data!C1": "3"}, edit: edit("cols", "delete", "Data!B"),
			after: m{"Data!A1": "1", "Data!B1": "3"}},
		{name: "a dynamic array spills again", before: m{"Data!A1": "1", "Data!A2": "2", "Data!M2": "=A1:A2*2|2", "Data!M3": "4"}, extra: []string{"Data!M2:M3"}, edit: edit("set", "Data!A2", "5"),
			after: m{"Data!A1": "1", "Data!A2": "5", "Data!M2": "=A1:A2*2|2", "Data!M3": "10"}},
		{name: "a dynamic array's results change with nothing it reads changed", before: m{"Data!A1": "1", "Data!A2": "2", "Data!M2": "=A1:A2*2|2", "Data!M3": "4", "Data!Z1": "0"}, extra: []string{"Data!M2:M3"}, edit: edit("set", "Data!Z1", "5"),
			after: m{"Data!A1": "1", "Data!A2": "2", "Data!M2": "=A1:A2*2|2", "Data!M3": "7", "Data!Z1": "5"}, wantErr: "Data!M3: outside the edit's reach, it held number 4, and holds number 7 now"},
		{name: "a set in a dynamic array's results", before: m{"Data!M2": "=SEQUENCE(2)|1", "Data!M3": "2"}, extra: []string{"Data!M2:M3"}, edit: edit("set", "Data!M3", "'x"),
			afterExtra: []string{}, after: m{"Data!M2": "=SEQUENCE(2)|0", "Data!M3": "'x"}},
		{name: "rows inserted through a dynamic array", before: m{"Data!M2": "=SEQUENCE(2)|1", "Data!M3": "2"}, extra: []string{"Data!M2:M3"}, edit: edit("rows", "insert", "Data!3"),
			after: m{"Data!M2": "=SEQUENCE(2)|1", "Data!M3": "2"}},
		{name: "replace", before: m{"Data!A1": "'Apple pie", "Data!A2": "'Pear", "Data!A3": "5"}, edit: edit("replace", "apple", "Fig"),
			after: m{"Data!A1": "'Fig pie", "Data!A2": "'Pear", "Data!A3": "5"}},
		{name: "replace a text it does not find", before: m{"Data!A1": "'Apple pie", "Data!A2": "'Pear"}, edit: edit("replace", "apple", "Fig"),
			after: m{"Data!A1": "'Fig pie", "Data!A2": "'Fig"}, wantErr: `Data!A2: outside the edit's reach, it held text "Pear"`},
		// The numbers a replace may change are those whose text holds the
		// whole find (fuzzer-long's audit, f1.4: any find with a letter e, a
		// space or a dot put every number in the reach).
		{name: "replace wipes numbers it cannot match", before: m{"Data!A1": "'Apple", "Data!B1": "1200", "Data!C1": "7"}, edit: edit("replace", "Apple", "Pear"),
			after: m{"Data!A1": "'Pear"}, wantErr: "Data!B1: outside the edit's reach, it held number 1200, and holds nothing now"},
		{name: "replace a digit in a number", before: m{"Data!A1": "1200", "Data!A2": "7"}, edit: edit("replace", "2", "5"),
			after: m{"Data!A1": "1500", "Data!A2": "7"}},
		{name: "replace a digit wipes a number without it", before: m{"Data!A1": "1200", "Data!A2": "7"}, edit: edit("replace", "2", "5"),
			after: m{"Data!A1": "1500"}, wantErr: "Data!A2: outside the edit's reach, it held number 7, and holds nothing now"},
		{name: "replace in formulas", before: m{"Data!A1": "=SUM(B1:B2)|3", "Data!A2": "=MAX(B1:B2)|2"}, edit: edit("replace", "SUM(", "PRODUCT(", "--formulas"),
			after: m{"Data!A1": "=PRODUCT(B1:B2)|2", "Data!A2": "=MAX(B1:B2)|2"}},
		{name: "sort a range", before: m{"Data!A2": "2", "Data!A3": "1", "Data!B9": "=A2|2"}, edit: edit("sort", "Data!A2:A3", "A"),
			after: m{"Data!A2": "1", "Data!A3": "2", "Data!B9": "=A2|1"}},
		{name: "sort past its range", before: m{"Data!A2": "2", "Data!A3": "1", "Data!A4": "0"}, edit: edit("sort", "Data!A2:A3", "A"),
			after: m{"Data!A2": "0", "Data!A3": "1", "Data!A4": "2"}, wantErr: "Data!A4: outside the edit's reach, it held number 0, and holds number 2 now"},
		{name: "sheet copy", before: m{"Data!A1": "1"}, edit: edit("sheet", "copy", "Data", "Copy"),
			after: m{"Data!A1": "1", "Copy!A1": "1"}},
		{name: "sheet add, with something in it", before: m{"Data!A1": "1"}, edit: edit("sheet", "add", "New"),
			after: m{"Data!A1": "1", "New!A1": "1"}, wantErr: "New!A1: outside the edit's reach, it was empty"},
		{name: "a sheet the edit does not name", before: m{"Data!A1": "1"}, edit: edit("set", "Data!A1", "5"),
			after: m{"Data!A1": "5", "Other!A1": "1"}, wantErr: "the edit made a sheet, Other, it does not name"},
		{name: "sheet rename, and what reads it", before: m{"Data!A1": "1", "Plan!A1": "=Data!A1|1"}, edit: edit("sheet", "rename", "Data", "Old"),
			after: m{"Old!A1": "1", "Plan!A1": "=Old!A1|1"}},
		// References rewritten: only as the edit's shift puts them (the
		// auditor's slip: unqualified references read as on the moved block's
		// sheet, fuzzer-long's audit f0.2).
		{name: "a move on another sheet leaves a formula's own references", before: m{"Data!B6": "5", "Data!G6": "=B6*2|10", "Plan!B6": "1"}, edit: edit("move", "Plan!B6", "Plan!E6"),
			after: m{"Data!B6": "5", "Data!G6": "=B6*2|10", "Plan!E6": "1"}},
		{name: "a move on another sheet rewrites a formula's own references", before: m{"Data!B6": "5", "Data!G6": "=B6*2|10", "Plan!B6": "1"}, edit: edit("move", "Plan!B6", "Plan!E6"),
			after: m{"Data!B6": "5", "Data!G6": "=E6*2|0", "Plan!E6": "1"}, wantErr: "Data!G6: outside the edit's reach, its formula =B6*2 is rewritten as =E6*2, and the edit's shift of references gives =B6*2"},
		{name: "a move to another sheet, and what reads it", before: m{"Data!A1": "1", "Data!B5": "=A1*2+Data!A1|3"}, edit: edit("move", "Data!A1", "Plan!C3"),
			after: m{"Plan!C3": "1", "Data!B5": "=Plan!C3*2+'Plan'!C3|3"}},
		{name: "a move leaves what reads it behind", before: m{"Data!A1": "1", "Data!B5": "=A1*2|2"}, edit: edit("move", "Data!A1", "Data!C1"),
			after: m{"Data!C1": "1", "Data!B5": "=A1*2|0"}, wantErr: "Data!B5: outside the edit's reach, its formula =A1*2 is left as it was, and the edit's shift of references gives =C1*2"},
		{name: "a move lands on cells a formula reads", before: m{"Data!A1": "1", "Data!C1": "3", "Data!B5": "=C1*2+$C$1|9"}, edit: edit("move", "Data!A1", "Data!C1"),
			after: m{"Data!C1": "1", "Data!B5": "=#REF!*2+#REF!|0"}},
		{name: "a move takes part of a range a formula reads", before: m{"Data!A1": "1", "Data!A2": "2", "Data!B5": "=SUM(A1:A2)|3"}, edit: edit("move", "Data!A1", "Data!C1"),
			after: m{"Data!A2": "2", "Data!C1": "1", "Data!B5": "=SUM(A1:A2)|2"}},
		{name: "a move takes a range a formula reads, its dollars too", before: m{"Data!A1": "1", "Data!A2": "2", "Data!B5": "=SUM($A$1:A2)|3"}, edit: edit("move", "Data!A1:A2", "Data!C3"),
			after: m{"Data!C3": "1", "Data!C4": "2", "Data!B5": "=SUM($C$3:C4)|3"}},
		{name: "a move drops a reference's dollars", before: m{"Data!A1": "1", "Data!B5": "=$A$1*2|2"}, edit: edit("move", "Data!A1", "Data!C1"),
			after: m{"Data!C1": "1", "Data!B5": "=C1*2|2"}, wantErr: "is rewritten as =C1*2, and the edit's shift of references gives =$C$1*2"},
		{name: "rows insert grows a range across it", before: m{"Data!A1": "1", "Data!A2": "2", "Data!B9": "=SUM(A1:A2)|3", "Plan!A1": "=Data!A2+A2|2"}, edit: edit("rows", "insert", "Data!2"),
			after: m{"Data!A1": "1", "Data!A3": "2", "Data!B10": "=SUM(A1:A3)|3", "Plan!A1": "=Data!A3+A2|2"}},
		{name: "rows insert leaves a reference", before: m{"Data!A1": "1", "Data!A2": "2", "Data!B9": "=SUM(A1:A2)|3"}, edit: edit("rows", "insert", "Data!2"),
			after: m{"Data!A1": "1", "Data!A3": "2", "Data!B10": "=SUM(A1:A2)|1"}, wantErr: "Data!B10: outside the edit's reach, its formula =SUM(A1:A2) is left as it was, and the edit's shift of references gives =SUM(A1:A3)"},
		{name: "rows insert shifts another sheet's own references", before: m{"Data!A1": "1", "Data!A2": "2", "Plan!A1": "=Data!A2+A2|2"}, edit: edit("rows", "insert", "Data!2"),
			after: m{"Data!A1": "1", "Data!A3": "2", "Plan!A1": "=Data!A3+A3|2"}, wantErr: "the edit's shift of references gives =Data!A3+A2"},
		{name: "rows delete shrinks a range by one too many", before: m{"Data!A1": "1", "Data!A2": "2", "Data!A3": "3", "Data!B9": "=SUM(A1:A3)|6"}, edit: edit("rows", "delete", "Data!2"),
			after: m{"Data!A1": "1", "Data!A2": "3", "Data!B8": "=SUM(A1:A1)|1"}, wantErr: "is rewritten as =SUM(A1:A1), and the edit's shift of references gives =SUM(A1:A2)"},
		{name: "rows delete takes a range whole", before: m{"Data!A1": "1", "Data!A2": "2", "Data!A3": "3", "Data!B9": "=SUM(A2:A3)+A1|6"}, edit: edit("rows", "delete", "Data!2:3"),
			after: m{"Data!A1": "1", "Data!B7": "=SUM(#REF!)+A1|0"}},
		{name: "cols delete, and another sheet's references", before: m{"Data!A1": "1", "Data!B1": "2", "Plan!B1": "3", "Plan!C1": "=Data!B1+B1|5"}, edit: edit("cols", "delete", "Plan!A"),
			after: m{"Data!A1": "1", "Data!B1": "2", "Plan!A1": "3", "Plan!B1": "=Data!B1+A1|5"}},
		{name: "cols delete shifts a reference to another sheet", before: m{"Data!A1": "1", "Data!B1": "2", "Plan!B1": "3", "Plan!C1": "=Data!B1+B1|5"}, edit: edit("cols", "delete", "Plan!A"),
			after: m{"Data!A1": "1", "Data!B1": "2", "Plan!A1": "3", "Plan!B1": "=Data!A1+A1|4"}, wantErr: "the edit's shift of references gives =Data!B1+A1"},
		{name: "sheet rename renames a reference to another sheet", before: m{"Data!A1": "1", "Plan!A1": "=Data!A1+Plan!A2|1", "Plan!A2": "0"}, edit: edit("sheet", "rename", "Data", "Old"),
			after: m{"Old!A1": "1", "Plan!A1": "=Old!A1+Old!A2|1", "Plan!A2": "0"}, wantErr: "the edit's shift of references gives =Old!A1+Plan!A2"},
		{name: "a batch moves, then inserts", before: m{"Data!A1": "1", "Data!B5": "=A1*2|2"},
			edit:  Edit{Class: "do", Args: []string{"do", FileArg}, Stdin: "move Data!A1 Data!C1\nrows insert Data!1\n"},
			after: m{"Data!C2": "1", "Data!B6": "=C2*2|2"}},
		{name: "a script moves, and what reads it is left", before: m{"Data!A1": "1", "Data!B5": "=A1*2|2"},
			edit:  script(`ws.getRange("A1").moveTo("E1");`),
			after: m{"Data!E1": "1", "Data!B5": "=A1*2|0"}, wantErr: "the edit's shift of references gives =E1*2"},
		{name: "a batch", before: m{"Data!A1": "1", "Data!A2": "2", "Data!A5": "5"},
			edit:  Edit{Class: "do", Args: []string{"do", FileArg}, Stdin: "rows insert Data!2\nset Data!A2 9\nclear Data!A1\n"},
			after: m{"Data!A2": "9", "Data!A3": "2", "Data!A6": "5"}},
		{name: "a batch leaves one in place", before: m{"Data!A1": "1", "Data!A2": "2", "Data!A5": "5"},
			edit:  Edit{Class: "do", Args: []string{"do", FileArg}, Stdin: "rows insert Data!2\nset Data!A2 9\n"},
			after: m{"Data!A1": "1", "Data!A2": "9", "Data!A3": "2", "Data!A5": "5"}, wantErr: "Data!A6: outside the edit's reach, it held number 5, and holds nothing now"},
		{name: "a script", before: m{"Data!A1": "1", "Data!A2": "2", "Data!C1": "3"},
			edit:  script(`ws.getRange("A1").moveTo("E1");`, `ws.getRange("B:B").delete("Left");`, `ws.getRange("A2").setValue(5);`),
			after: m{"Data!A2": "5", "Data!B1": "3", "Data!D1": "1"}},
		{name: "a script copies past its target", before: m{"Data!A1": "1", "Data!A2": "2"},
			edit:  script(`ws.getRange("D1").copyFrom("A1:A2");`),
			after: m{"Data!A1": "1", "Data!A2": "2", "Data!D1": "1", "Data!D2": "2", "Data!D3": "2"}, wantErr: "Data!D3: outside the edit's reach, it was empty"},
	} {
		t.Run(tc.name, func(t *testing.T) {
			before := reachBook(tc.before, tc.extra...)
			afterExtra := tc.extra
			if tc.afterExtra != nil {
				afterExtra = tc.afterExtra
			}
			mo := newModel(before)
			mo.r = newReach(before)
			mo.apply(tc.edit)
			if mo.why != "" {
				t.Fatalf("the model cannot read the edit: %s", mo.why)
			}
			got := mo.reached(reachBook(tc.after, afterExtra...))
			if tc.wantErr == "" && got != "" {
				t.Errorf("the model refuses an edit that kept to its reach: %s", got)
			}
			if tc.wantErr != "" && !strings.Contains(got, tc.wantErr) {
				t.Errorf("the model says %q, want it to say %q", got, tc.wantErr)
			}
		})
	}
}

// TestRefToken: a formula's references are read as one token each, so a
// formula rewritten in its references only reads the same; and what it reads
// is found past its strings, in a sheet named with quotes too.
func TestRefToken(t *testing.T) {
	for _, tc := range []struct{ a, b string }{
		{`IF(N("p")=0,B3*2)`, `IF(N("p")=0,B5*2)`},
		{`SUM(Data!B2:B10)`, `SUM('Q 2027'!B2:B12)`},
		{`'Bob''s Plan'!A1+1`, `#REF!+1`},
		{`Data!$C$2*2`, `Data!#REF!*2`},
	} {
		if masked(tc.a) != masked(tc.b) {
			t.Errorf("%s and %s read %s and %s", tc.a, tc.b, masked(tc.a), masked(tc.b))
		}
	}
	for _, tc := range []struct{ in, want string }{
		{`IF(N("p")=0,A2&"-x")`, `IF(N()=0,A2&)`},
		{`IF(N("p")=0,'Q1 "draft"'!B9*2)`, `IF(N()=0,'Q1 "draft"'!B9*2)`},
		{`'Bob''s Plan'!A1&"it's"`, `'Bob''s Plan'!A1&`},
		{`"say ""hi"""&B2`, `&B2`},
	} {
		if got := stringsOut(tc.in); got != tc.want {
			t.Errorf("stringsOut(%s) = %s, want %s", tc.in, got, tc.want)
		}
	}
	if ins, ok := inputs(`IF(N("p")=0,'Q1 "draft"'!B9*2)`, "summary", nil); !ok || len(ins) != 1 || ins[0].sheet != `q1 "draft"` || ins[0].rg.String() != "B9" {
		t.Errorf("the reference to a sheet with quotes in its name: %v %v", ins, ok)
	}
	if masked("B3*2") == masked("B3*3") || masked("SUM(B2:B4)") == masked("MAX(B2:B4)") {
		t.Error("a formula rewritten other than in its references reads the same")
	}
	// What refsOf reads as a reference, and what not: a string, a function or
	// a name holding letters and digits.
	for _, tc := range []struct{ f, want string }{
		{`LOG10(B2)+ATAN2(1,2)`, `B2`},
		{`"A1"&B2&"x""C3"`, `B2`},
		{`'Bob''s Plan'!$A$1:B3*2`, `bob's plan!$A$1:B3`},
		{`Data!#REF!+#REF!`, `data!#REF! #REF!`},
		{`_xlfn.SEQUENCE(2)+SUM(Units)+Rate`, ``},
		{`Åäö!B2+'Q1 "draft"'!C9`, `åäö!B2 q1 "draft"!C9`},
		{`b2+XFD1+XFE1+A1048577`, `B2 XFD1`},
	} {
		var got []string
		for _, r := range refsOf(tc.f) {
			s := ""
			if r.sheet != "" {
				s = r.sheet + "!"
			}
			switch {
			case r.bad:
				s += "#REF!"
			case r.rng:
				s += r.a.String() + ":" + r.b.String()
			default:
				s += r.a.String()
			}
			got = append(got, s)
		}
		if g := strings.Join(got, " "); g != tc.want {
			t.Errorf("refsOf(%s) = %s, want %s", tc.f, g, tc.want)
		}
	}
}

// TestReachFormats: a cell outside the edit's reach keeps its number format
// (fuzzer-long's audit, f0.5: a date shown as 46022 passed); one in it may
// take another.
func TestReachFormats(t *testing.T) {
	book := func(a1, b1 uint16) *grid.Workbook {
		wb := reachBook(map[string]string{"Data!A1": "46022", "Data!B1": "1"})
		wb.Formats = []grid.Format{{}, {Kind: grid.Date, Code: "yyyy-mm-dd"}, {Kind: grid.Decimal, Decimals: 2, Code: "0.00"}}
		s := wb.Sheet("Data")
		for col, f := range []uint16{a1, b1} {
			c := s.Get(1, col+1)
			c.Fmt = f
			s.Set(1, col+1, c)
		}
		return wb
	}
	set := Edit{Class: "set", Args: []string{"set", FileArg, "Data!B1", "1"}}
	for _, tc := range []struct {
		name    string
		after   *grid.Workbook
		wantErr string
	}{
		{"kept", book(1, 0), ""},
		{"the cell set takes another", book(1, 2), ""},
		{"a date loses its format", book(0, 0), "Data!A1: outside the edit's reach, it held number 46022 shown as yyyy-mm-dd, and is shown as General now"},
		{"a date takes another", book(2, 0), "shown as yyyy-mm-dd, and is shown as 0.00 now"},
	} {
		mo := newModel(book(1, 0))
		mo.r = newReach(book(1, 0))
		mo.apply(set)
		got := mo.reached(tc.after)
		if tc.wantErr == "" && got != "" || tc.wantErr != "" && !strings.Contains(got, tc.wantErr) {
			t.Errorf("%s: the model says %q, want %q", tc.name, got, tc.wantErr)
		}
	}
	// A replace reads a date as sheets shows it: 46022 shown as 2025-12-31
	// is in the reach of a find of 2025-12, and not of one it cannot hold.
	for _, tc := range []struct {
		find    string
		wantErr string
	}{
		{"2025-12", ""},
		{"Apple", "Data!A1: outside the edit's reach, it held number 46022"},
	} {
		mo := newModel(book(1, 0))
		mo.r = newReach(book(1, 0))
		mo.apply(Edit{Class: "replace", Args: []string{"replace", FileArg, tc.find, "x"}})
		after := book(1, 0)
		after.Sheet("Data").Set(1, 1, grid.Cell{Kind: grid.Text, Text: "x"})
		if got := mo.reached(after); tc.wantErr == "" && got != "" || tc.wantErr != "" && !strings.Contains(got, tc.wantErr) {
			t.Errorf("replace %s over a date: the model says %q, want %q", tc.find, got, tc.wantErr)
		}
	}
	// Formats renumbered (another index, the same format) are the same.
	after := book(2, 0)
	after.Formats = []grid.Format{{}, {Kind: grid.Decimal, Decimals: 2, Code: "0.00"}, {Kind: grid.Date, Code: "yyyy-mm-dd"}}
	mo := newModel(book(1, 0))
	mo.r = newReach(book(1, 0))
	mo.apply(set)
	if got := mo.reached(after); got != "" {
		t.Errorf("formats renumbered: %s", got)
	}
}

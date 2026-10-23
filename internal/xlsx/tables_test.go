package xlsx

import (
	"regexp"
	"strings"
	"testing"

	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/input"
	"github.com/xuri/excelize/v2"
)

var (
	tableRef  = regexp.MustCompile(`<table [^>]*\sref="([^"]+)"`)
	filterRef = regexp.MustCompile(`<autoFilter ref="([^"]+)"`)
)

// tableState reads the table part: its range, filter range, and whether it
// still counts a totals row.
func tableState(t *testing.T, path string) (ref, filter string, totals bool) {
	t.Helper()
	x := sheetXML(t, path, "xl/tables/table1.xml")
	m, f := tableRef.FindStringSubmatch(x), filterRef.FindStringSubmatch(x)
	if m == nil || f == nil {
		t.Fatalf("table part: %s", x)
	}
	return m[1], f[1], strings.Contains(x, `totalsRowCount="1"`)
}

func edit(t *testing.T, path string, fn func(e *Editor) error) *Editor {
	t.Helper()
	e, err := OpenEditor(path, false)
	if err != nil {
		t.Fatal(err)
	}
	if err := fn(e); err != nil {
		t.Fatal(err)
	}
	if err := e.Save(false); err != nil {
		t.Fatal(err)
	}
	return e
}

// An Excel table (stored under an absolute path, as openpyxl writes it,
// which excelize does not adjust) follows rows added above its totals row,
// keeps that totals row, and loses it only when it is deleted.
func TestTablesFollowRows(t *testing.T) {
	path := fixture(t, "tables.xlsx")

	// Append one row the way write --append does.
	edit(t, path, func(e *Editor) error {
		if err := e.InsertRows("Sales", 5, 1); err != nil {
			return err
		}
		e.ExtendTable("Sales", grid.Range{R1: 2, C1: 1, R2: 4, C2: 3}, 1)
		return e.Set("Sales", 5, 1, input.Parse("West", false))
	})
	if ref, filter, totals := tableState(t, path); ref != "A1:C6" || filter != "A1:C5" || !totals {
		t.Errorf("after append: ref %s, filter %s, totals row %v", ref, filter, totals)
	}

	// What a table cannot do without is refused.
	e, _ := OpenEditor(path, false)
	for what, err := range map[string]error{
		"header row":      e.DeleteRows("Sales", 1, 1),
		"all data rows":   e.DeleteRows("Sales", 2, 5),
		"column inside":   e.InsertCols("Sales", 2, 1),
		"columns overlap": e.DeleteCols("Sales", 3, 4),
		"clear":           e.Clear("Sales"),
	} {
		if err == nil || !strings.Contains(err.Error(), "SalesT") {
			t.Errorf("%s: err = %v", what, err)
		}
	}

	// Rows above and columns to the left move it; deleting its totals row
	// leaves a table without one.
	edit(t, path, func(e *Editor) error {
		if err := e.InsertCols("Sales", 1, 1); err != nil {
			return err
		}
		return e.DeleteRows("Sales", 6, 6)
	})
	if ref, filter, totals := tableState(t, path); ref != "B1:D5" || filter != "B1:D5" || totals {
		t.Errorf("after moves: ref %s, filter %s, totals row %v", ref, filter, totals)
	}
}

// Column names follow the header cells, as Excel requires; an emptied
// header gets a name, as Excel would give it.
func TestTableColumnNames(t *testing.T) {
	path := fixture(t, "tables.xlsx")
	e := edit(t, path, func(e *Editor) error {
		if err := e.Set("Sales", 1, 2, input.Parse("Qty", false)); err != nil {
			return err
		}
		return e.Set("Sales", 1, 3, input.Value{})
	})
	x := sheetXML(t, path, "xl/tables/table1.xml")
	if !strings.Contains(x, `name="Region"`) || !strings.Contains(x, `name="Qty"`) || !strings.Contains(x, `name="Column3"`) {
		t.Errorf("columns: %s", x)
	}
	if !strings.Contains(x, `<tableStyleInfo name="TableStyleMedium2"`) {
		t.Errorf("table style lost: %s", x)
	}
	f, _ := excelize.OpenFile(path)
	defer f.Close()
	if v, _ := f.GetCellValue("Sales", "C1"); v != "Column3" {
		t.Errorf("C1 = %q", v)
	}
	// The totals row and E2 name SalesT[Units]: they follow the rename.
	w := strings.Join(e.Warnings(), "\n")
	if !strings.Contains(w, `columns of Excel table "SalesT" renamed after their headers: "Units" → "Qty", "Price" → "Column3"; formulas naming them now use the new names`) ||
		!strings.Contains(w, `header C1 of Excel table "SalesT" was empty; it is now "Column3"`) {
		t.Errorf("warnings: %s", w)
	}
	for cell, want := range map[string]string{"B5": "SUBTOTAL(109,SalesT[Qty])", "E2": "SUM(SalesT[Qty])"} {
		if got, _ := f.GetCellFormula("Sales", cell); got != want {
			t.Errorf("%s = %s, want %s", cell, got, want)
		}
	}
}

// Formulas inside a table name its columns without the table's name, in
// cells and in the table's own definitions; they follow renames too.
func TestRenameInsideTable(t *testing.T) {
	path := writeBook(t, [][]string{{"Name", "Units", "Price", "Total"}, {"Ada", "2", "3", "=[@Units]*[@Price]"}, {"Bob", "4", "5", "=[@Units]*[@Price]"}})
	f, _ := excelize.OpenFile(path)
	if err := f.AddTable("Data", &excelize.Table{Range: "A1:D3", Name: "Orders"}); err != nil {
		t.Fatal(err)
	}
	f.SetCellFormula("Data", "F1", "SUM(Orders[Units])")
	f.Save()
	f.Close()
	edit(t, path, func(e *Editor) error { return e.Set("Data", 1, 2, input.Parse("Qty", false)) })
	f, _ = excelize.OpenFile(path)
	defer f.Close()
	for cell, want := range map[string]string{"D2": "[@Qty]*[@Price]", "D3": "[@Qty]*[@Price]", "F1": "SUM(Orders[Qty])"} {
		if got, _ := f.GetCellFormula("Data", cell); got != want {
			t.Errorf("%s = %s, want %s", cell, got, want)
		}
	}
}

// A table excelize made (a relative part path, which excelize does adjust,
// rewriting the part its own way) comes out as sheets writes it.
func TestTablesExcelizeMade(t *testing.T) {
	path := writeBook(t, [][]string{{"Name", "Score"}, {"Ada", "9"}, {"Bob", "4"}})
	f, _ := excelize.OpenFile(path)
	if err := f.AddTable("Data", &excelize.Table{Range: "A1:B3", Name: "Scores", StyleName: "TableStyleLight9"}); err != nil {
		t.Fatal(err)
	}
	if err := f.Save(); err != nil {
		t.Fatal(err)
	}
	f.Close()
	edit(t, path, func(e *Editor) error { return e.InsertRows("Data", 3, 2) })
	if ref, filter, _ := tableState(t, path); ref != "A1:B5" || filter != "A1:B5" {
		t.Errorf("ref %s, filter %s", ref, filter)
	}
	if x := sheetXML(t, path, "xl/tables/table1.xml"); !strings.Contains(x, "TableStyleLight9") {
		t.Errorf("table style lost: %s", x)
	}
}

// Formulas that use a sheet's tables by name use the sheet.
func TestReferencesThroughTables(t *testing.T) {
	path := fixture(t, "tables.xlsx")
	edit(t, path, func(e *Editor) error {
		if _, _, err := e.EnsureSheet("Report"); err != nil {
			return err
		}
		for c, f := range []string{"=SUM(SalesT[Units])", "=ROWS(SalesT)", `="SalesTx"`, "=SalesTotal*2"} {
			if err := e.Set("Report", 1, c+1, input.Parse(f, false)); err != nil {
				return err
			}
		}
		return nil
	})
	e, _ := OpenEditor(path, false)
	if refs := strings.Join(e.References("Sales"), ", "); refs != "Report!A1, Report!B1" {
		t.Errorf("references = %s", refs)
	}
}

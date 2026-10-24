package analyze

import (
	"strings"
	"testing"

	"github.com/eemax/sheets/internal/grid"
)

func TestResolve(t *testing.T) {
	wb, s := sheet(
		[]any{"Region", "Units", "Revenue", "Margin's"},
		[]any{"East", 2, 30, 0.1},
		[]any{"West", 4, 50, 0.2},
		[]any{"Total", "=SUM(B2:B3)", "=SUM(C2:C3)", nil},
	)
	s.Name = "Sales"
	q := grid.NewSheet("Q1 Plan")
	q.Set(1, 1, grid.Cell{Kind: grid.Text, Text: "x"})
	t2 := grid.NewSheet("Tables")
	t2.Tables = []grid.NamedTable{{Name: "Orders", Range: grid.Range{R1: 3, C1: 2, R2: 7, C2: 4}, Totals: 1, Columns: []string{"Id", "Qty", "Price"}}}
	wb.Sheets = append(wb.Sheets, q, t2)
	for ref, want := range map[string]string{
		"Sales!B2:C3":                  "Sales B2:C3",
		"B:D":                          "Sales B1:D1048576",
		"Sales":                        "Sales -",
		"'Q1 Plan'!A1":                 "Q1 Plan A1",
		"Q1 Plan":                      "Q1 Plan -",
		"Sales[Revenue]":               "Sales C2:C3",
		"Sales[revenue]":               "Sales C2:C3",
		"Sales[[Units]:[Revenue]]":     "Sales B2:C3",
		"Sales[[Revenue]:[Units]]":     "Sales B2:C3",
		"Sales[#Headers]":              "Sales A1:D1",
		"Sales[#Totals]":               "Sales A4:D4",
		"Sales[#All]":                  "Sales A1:D4",
		"Sales[#Data]":                 "Sales A2:D3",
		"Sales[[#Totals],[Units]]":     "Sales B4",
		"Sales[[#Headers],[#Data]]":    "Sales A1:D3",
		"Sales[Margin''s]":             "Sales D2:D3",
		"Sales[[Margin''s]]":           "Sales D2:D3",
		"Orders[Qty]":                  "Tables C4:C6",
		"orders[#Totals]":              "Tables B7:D7",
		"Orders[[#Headers],[Id]]":      "Tables B3",
		"Orders[[#All],[Qty]:[Price]]": "Tables C3:D7",
	} {
		sheet, r, err := Resolve(wb, ref)
		if err != nil {
			t.Errorf("Resolve(%q): %v", ref, err)
			continue
		}
		got := sheet + " -"
		if r != nil {
			got = sheet + " " + r.String()
		}
		if got != want {
			t.Errorf("Resolve(%q) = %s, want %s", ref, got, want)
		}
	}
	for ref, msg := range map[string]string{
		"Sales[Revnue]":               `no column "Revnue" (columns: Region, Units, Revenue, Margin's)`,
		"Nope[Units]":                 "neither a sheet nor an Excel table (sheets: Sales, Q1 Plan, Tables; Excel tables: Orders)",
		"Sales[#Bogus]":               "#Bogus is not one of",
		"'Q1 Plan'[x]":                "has no table with a header row",
		"Sales[]":                     "empty brackets",
		"Sales[[Units],[Revenue]]":    "give one column or one span",
		"Sales[[#Headers],[#Totals]]": "not next to each other",
		"Nope!A1":                     `sheet "Nope" not found`,
		"Sales!ZZZZ1":                 "give a range",
	} {
		_, _, err := Resolve(wb, ref)
		if err == nil || !strings.Contains(err.Error(), msg) {
			t.Errorf("Resolve(%q) = %v, want an error with %q", ref, err, msg)
		}
	}
}

// A sheet with two tables: the one with the column named.
func TestResolveTwoTables(t *testing.T) {
	wb, s := sheet(
		[]any{"Region", "Units"},
		[]any{"East", 2},
		[]any{"West", 4},
		[]any{nil, nil},
		[]any{"Product", "Price"},
		[]any{"Pen", 1.5},
		[]any{"Ink", 3.0},
	)
	s.Name = "S"
	_, r, err := Resolve(wb, "S[Price]")
	if err != nil || r.String() != "B6:B7" {
		t.Errorf("S[Price] = %v, %v", r, err)
	}
	if _, _, err := Resolve(wb, "S[#Headers]"); err == nil || !strings.Contains(err.Error(), "holds 2 tables (A1:B3, A5:B7)") {
		t.Errorf("S[#Headers]: %v", err)
	}
}

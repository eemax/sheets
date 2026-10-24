package formula

import (
	"testing"

	"github.com/eemax/sheets/internal/grid"
)

func TestMoveRefs(t *testing.T) {
	rg := func(s string) grid.Range {
		r, err := grid.ParseRange(s)
		if err != nil {
			t.Fatal(err)
		}
		return r
	}
	down5 := Move{From: "Sheet1", Src: rg("B2:C3"), To: "Sheet1", DR: 5}
	tests := []struct {
		f, home, newHome string
		m                Move
		want             string
	}{
		// References into the moved block follow it.
		{"B2*2", "Sheet1", "Sheet1", down5, "B7*2"},
		{"SUM(B2:C3)", "Sheet1", "Sheet1", down5, "SUM(B7:C8)"},
		{"$B$2+B$3+$C2", "Sheet1", "Sheet1", down5, "$B$7+B$8+$C7"},
		{`"B2"&B2`, "Sheet1", "Sheet1", down5, `"B2"&B7`},
		// Ranges only partly moved stay.
		{"SUM(B1:B3)", "Sheet1", "Sheet1", down5, "SUM(B1:B3)"},
		{"SUM(B:B)", "Sheet1", "Sheet1", down5, "SUM(B:B)"},
		// Other sheets are other cells.
		{"Data!B2", "Sheet1", "Sheet1", down5, "Data!B2"},
		{"B2", "Data", "Data", down5, "B2"},
		{"Sheet1!B2", "Data", "Data", down5, "Sheet1!B7"},
		// References to cells pasted over become #REF!.
		{"B8+1", "Sheet1", "Sheet1", down5, "#REF!+1"},
		{"SUM(B7:C8)", "Sheet1", "Sheet1", down5, "SUM(#REF!)"},
		{"Sheet1!C7", "Data", "Data", down5, "Sheet1!#REF!"},
		{"SUM(B6:B8)", "Sheet1", "Sheet1", down5, "SUM(B6:B8)"},
		// A block moved onto itself: cells that moved follow, the rest of
		// the landing area is overwritten.
		{"A1+A2+A3", "Sheet1", "Sheet1", Move{From: "Sheet1", Src: rg("A1:A2"), To: "Sheet1", DR: 1}, "A2+A3+#REF!"},
		// Across sheets.
		{"B2*2", "Sheet1", "Sheet1", Move{From: "Sheet1", Src: rg("B2"), To: "Data"}, "Data!B2*2"},
		{"B2*2", "Sheet1", "Sheet1", Move{From: "Sheet1", Src: rg("B2"), To: "My Data", DC: 1}, "'My Data'!C2*2"},
		{"'Q1 Sales'!B2", "Summary", "Summary", Move{From: "Q1 Sales", Src: rg("B2"), To: "Data"}, "Data!B2"},
		{"Data!B2", "Sheet1", "Sheet1", Move{From: "Data", Src: rg("B2"), To: "Sheet1", DR: 1}, "Sheet1!B3"},
		// A formula moved with the block to another sheet keeps reading
		// the cells it read.
		{"B2+Z9", "Sheet1", "Data", Move{From: "Sheet1", Src: rg("B2:C3"), To: "Data"}, "B2+Sheet1!Z9"},
		{"SUM(A1:A5)", "Sheet1", "Data", Move{From: "Sheet1", Src: rg("B2:C3"), To: "Data"}, "SUM(Sheet1!A1:A5)"},
		// Names are always qualified.
		{"Sheet1!$B$2:$C$3", "", "", down5, "Sheet1!$B$7:$C$8"},
		{"Sheet1!$B$1:$C$3", "", "", down5, "Sheet1!$B$1:$C$3"},
		// Left alone: other workbooks, references across sheets.
		{"[1]Sheet1!B2", "Data", "Data", down5, "[1]Sheet1!B2"},
		{"'[Book.xlsx]Sheet1'!B2", "Data", "Data", down5, "'[Book.xlsx]Sheet1'!B2"},
		{"SUM(Jan:Dec!B2)", "Sheet1", "Sheet1", down5, "SUM(Jan:Dec!B2)"},
		// Off the sheet.
		{"B2", "Sheet1", "Sheet1", Move{From: "Sheet1", Src: rg("B2"), To: "Sheet1", DR: grid.MaxRows}, "#REF!"},
	}
	for _, tt := range tests {
		if got := MoveRefs(tt.f, tt.home, tt.newHome, tt.m); got != tt.want {
			t.Errorf("MoveRefs(%q, %s→%s, %+v) = %q, want %q", tt.f, tt.home, tt.newHome, tt.m, got, tt.want)
		}
	}
}

func TestQualifierOf(t *testing.T) {
	for s, want := range map[string]string{
		"SUM(Sheet1!":     "Sheet1!",
		"SUM('My Sheet'!": "'My Sheet'!",
		"x+'It''s here'!": "'It''s here'!",
		"SUM(":            "",
		"+$Weird_Name.1!": "$Weird_Name.1!",
		"'a'&'Q1 ''24'!":  "'Q1 ''24'!",
	} {
		if got := qualifierOf(s); got != want {
			t.Errorf("qualifierOf(%q) = %q, want %q", s, got, want)
		}
	}
}

func TestRefUses(t *testing.T) {
	got := RefUses(`B2*$C$2+SUM($D$2:$D$9)+SUM(E:E)+Other!F2+SUM($C$2:C2)+"G5"+'My Sheet'!$H$1`)
	want := []RefUse{
		{Range: grid.Range{R1: 2, C1: 2, R2: 2, C2: 2}, RelRows: true, Text: "B2"},
		{Range: grid.Range{R1: 2, C1: 3, R2: 2, C2: 3}, AbsRows: true, Text: "$C$2"},
		{Range: grid.Range{R1: 2, C1: 4, R2: 9, C2: 4}, AbsRows: true, Text: "$D$2:$D$9"},
		{Range: grid.Range{R1: 1, C1: 5, R2: grid.MaxRows, C2: 5}, Text: "E:E"},
		{Sheet: "Other", Range: grid.Range{R1: 2, C1: 6, R2: 2, C2: 6}, RelRows: true, Text: "Other!F2"},
		{Range: grid.Range{R1: 2, C1: 3, R2: 2, C2: 3}, Text: "$C$2:C2"},
		{Sheet: "My Sheet", Range: grid.Range{R1: 1, C1: 8, R2: 1, C2: 8}, AbsRows: true, Text: "'My Sheet'!$H$1"},
	}
	if len(got) != len(want) {
		t.Fatalf("got %d refs, want %d: %+v", len(got), len(want), got)
	}
	for i := range want {
		if got[i] != want[i] {
			t.Errorf("ref %d = %+v, want %+v", i, got[i], want[i])
		}
	}
}

// The second half of a range is on its first half's sheet: deleting rows
// of the formula's own sheet leaves a range on another sheet alone.
func TestDeleteLinesOtherSheetRange(t *testing.T) {
	for _, tt := range []struct{ f, home, target, want string }{
		{"MAX(Data!D2:D8)", "Summary", "Summary", "MAX(Data!D2:D8)"},
		{"MAX(Data!D2:D8)+D9", "Summary", "Summary", "MAX(Data!D2:D8)+D7"},
		{"SUM('My Data'!B2:B9)", "Summary", "Summary", "SUM('My Data'!B2:B9)"},
		{"SUM(Data!D2:D8)", "Summary", "Data", "SUM(Data!D2:D6)"},
		{"SUM(D2:D8)", "Data", "Data", "SUM(D2:D6)"},
	} {
		if got := DeleteLines(tt.f, tt.home, tt.target, false, 3, 4); got != tt.want {
			t.Errorf("DeleteLines(%q on %s, rows 3:4 of %s) = %q, want %q", tt.f, tt.home, tt.target, got, tt.want)
		}
	}
}

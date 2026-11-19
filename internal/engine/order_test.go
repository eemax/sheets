package engine

import (
	"strconv"
	"testing"
)

// OrderReads tells which references a formula reads as a set: given whole to
// a function that no reordering of whole rows changes, alone or as cell-by-cell
// arithmetic or a comparison of ranges (SUMPRODUCT((A2:A6="East")*B2:B6),
// SUM(B2:B7*C2:C7)), those of several cells together in one group.
func TestOrderReadsElementwise(t *testing.T) {
	none := func(string, string) (RefArg, bool) { return RefArg{}, false }
	type read struct {
		area  string
		set   bool
		group bool // in a group
	}
	for _, c := range []struct {
		formula string
		want    []read
	}{
		{"=SUM(B2:B7)", []read{{"B2:B7", true, false}}},
		{"=SUMPRODUCT(B2:B6,C2:C6)", []read{{"B2:B6", true, true}, {"C2:C6", true, true}}},
		{`=SUMPRODUCT((A2:A6="East")*B2:B6)`, []read{{"A2:A6", true, true}, {"B2:B6", true, true}}},
		{"=SUMPRODUCT(B2:B6*C2:C6)", []read{{"B2:B6", true, true}, {"C2:C6", true, true}}},
		{"=SUM(B2:B7*C2:C7)", []read{{"B2:B7", true, true}, {"C2:C7", true, true}}},
		{"=SUMPRODUCT(--(B2:B6>4),C2:C6)", []read{{"B2:B6", true, true}, {"C2:C6", true, true}}},
		{"=MAX(B2:B6*2)", []read{{"B2:B6", true, true}}},
		// A single cell is a value every row is computed with: read by position.
		{"=SUMPRODUCT((B2:B6>$F$1)*C2:C6)", []read{{"B2:B6", true, true}, {"F1", false, false}, {"C2:C6", true, true}}},
		// Outside a function that reads a set, a range in arithmetic is by position.
		{"=B2:B6*C2:C6", []read{{"B2:B6", false, false}, {"C2:C6", false, false}}},
		{"=SUM(IF(B2:B6>4,C2:C6))", []read{{"B2:B6", false, false}, {"C2:C6", false, false}}},
		{"=SUM(B2*C2)", []read{{"B2", false, false}, {"C2", false, false}}},
	} {
		got, ok := OrderReads(c.formula, none)
		if !ok || len(got) != len(c.want) {
			t.Errorf("%s: %+v (ok %v), want %d reads", c.formula, got, ok, len(c.want))
			continue
		}
		for i, w := range c.want {
			g := got[i]
			area := Area{R1: g.Area.R1, C1: g.Area.C1, R2: g.Area.R2, C2: g.Area.C2}
			if rendered := areaText(area); rendered != w.area || g.Set != w.set || (g.Group > 0) != w.group {
				t.Errorf("%s: read %d is %s set %v group %d, want %+v", c.formula, i, rendered, g.Set, g.Group, w)
			}
		}
	}
	// Two arguments of SUM computed apart are two groups.
	got, _ := OrderReads("=SUM(B2:B6*C2:C6,D2:D6*E2:E6)", none)
	if len(got) != 4 || got[0].Group == 0 || got[0].Group != got[1].Group || got[2].Group != got[3].Group || got[0].Group == got[2].Group {
		t.Errorf("groups: %+v", got)
	}
}

// A call that pairs ranges cell by cell reads them as a set only when each of
// them is a reference, or references combined cell by cell (report-causes
// f5.3): an array constant or a call among them weighs the others' cells by
// position, so none of them is a set, and no group holds them.
func TestAgainstRegressOrderReadsConstant(t *testing.T) {
	none := func(string, string) (RefArg, bool) { return RefArg{}, false }
	for _, c := range []struct {
		formula string
		set     []bool // of each read, in order
	}{
		{"=SUMPRODUCT(Data!B2:B6,{1;2;3;4;5})", []bool{false}},
		{"=SUMPRODUCT({1;2;3;4;5},B2:B6,C2:C6)", []bool{false, false}},
		{"=SUMPRODUCT(B2:B6,ROW(B2:B6))", []bool{false, false}},
		{"=SUMPRODUCT(B2:B6,C2:C6*{1;2;3;4;5})", []bool{false, false}},
		{"=SUMPRODUCT(B2:B6,2)", []bool{false}},
		{"=SUMIF(A2:A6,\"East\",OFFSET(B2,0,0,5,1))", []bool{false, false}},
		// A criterion is no range paired: a constant or a call there leaves the
		// ranges a set.
		{"=SUMIF(A2:A6,\"East\",B2:B6)", []bool{true, true}},
		{"=SUMIFS(C2:C6,A2:A6,UPPER(\"east\"))", []bool{true, true}},
		{"=COUNTIFS(A2:A6,{\"East\"},B2:B6,\">3\")", []bool{true, true}},
		// References combined cell by cell are still read together.
		{`=SUMPRODUCT((A2:A6="East")*B2:B6,C2:C6)`, []bool{true, true, true}},
	} {
		got, ok := OrderReads(c.formula, none)
		if !ok || len(got) != len(c.set) {
			t.Errorf("%s: %+v (ok %v), want %d reads", c.formula, got, ok, len(c.set))
			continue
		}
		for i, set := range c.set {
			if got[i].Set != set || (got[i].Group > 0) != set {
				t.Errorf("%s: read %d set %v group %d, want set %v", c.formula, i, got[i].Set, got[i].Group, set)
			}
		}
	}
}

// areaText writes an area over the columns A to Z as A1 or B2:C9.
func areaText(a Area) string {
	cell := func(r, c int) string { return string(rune('A'+c-1)) + strconv.Itoa(r) }
	if a.R1 == a.R2 && a.C1 == a.C2 {
		return cell(a.R1, a.C1)
	}
	return cell(a.R1, a.C1) + ":" + cell(a.R2, a.C2)
}

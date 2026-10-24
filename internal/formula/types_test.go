package formula

import (
	"testing"

	"github.com/eemax/sheets/internal/grid"
)

// sheetTyper serves a fixed sheet: A1 the text "007", A2 the number 7, B1
// the text "k1", B2 "Units", C1 TRUE; the name Rate is a number.
type sheetTyper struct{}

func (sheetTyper) RangeKinds(sheet string, rg grid.Range) Kinds {
	cells := map[[2]int]Kinds{{1, 1}: KNumText, {2, 1}: KNumber, {1, 2}: KText, {2, 2}: KText, {1, 3}: KBool}
	var k Kinds
	for r := rg.R1; r <= min(rg.R2, 2); r++ {
		for c := rg.C1; c <= min(rg.C2, 3); c++ {
			if v, ok := cells[[2]int{r, c}]; ok {
				k |= v
			} else {
				k |= KNumber
			}
		}
	}
	if rg.R2 > 2 || rg.C2 > 3 {
		k |= KNumber // empty cells read as 0
	}
	return k
}

func (sheetTyper) NameKinds(name string) (Kinds, bool) {
	if name == "Rate" {
		return KNumber, true
	}
	return 0, false
}

func TestResultKinds(t *testing.T) {
	for f, want := range map[string]Kinds{
		"A2*2":                    KNumber,
		"A1":                      KNumText,
		`A1&""`:                   KText | KNumText,
		`TEXT(A2,"000")`:          KText | KNumText,
		"A2>3":                    KBool,
		"IF(A2>3,TRUE)":           KBool,
		`IF(A2>3,"big",A2)`:       KText | KNumber,
		"IF(A2,B1,C1)":            KText | KBool,
		"INDEX(A1:A2,1)":          KNumText | KNumber,
		"VLOOKUP(\"k1\",B1:B2,1)": KText,
		"SUM(A:A)*Rate":           KNumber,
		"Rate":                    KNumber,
		"Nope":                    KAny,
		"LET(x,1,x)":              KAny,
		"SUM(Sales[Units])":       KNumber,
		"Sales[Units]":            KAny,
		"-A2%":                    KNumber,
		`"x"`:                     KText,
		`"5"`:                     KNumText,
		"{1,2;3,4}":               KNumber,
		"Sheet2!A1:A2":            KNumText | KNumber,
		"IFERROR(1/0,#N/A)":       KNumber,
		"A1 +":                    KAny, // not a formula: anything
	} {
		if got := ResultKinds(f, sheetTyper{}); got != want {
			t.Errorf("ResultKinds(%q) = %04b, want %04b", f, got, want)
		}
	}
}

func TestTyped(t *testing.T) {
	for _, tt := range []struct {
		raw  string
		k    Kinds
		want grid.Cell
		ok   bool
	}{
		{"5", KNumber, grid.Cell{Kind: grid.Number, Num: 5}, true},
		{"005", KText | KNumText, grid.Cell{Kind: grid.Text, Text: "005"}, true},
		{"1", KBool, grid.Cell{Kind: grid.Bool, Bool: true}, true},
		{"North", KText | KNumber, grid.Cell{Kind: grid.Text, Text: "North"}, true},
		{"7", KText | KNumber, grid.Cell{Kind: grid.Number, Num: 7}, true},
		{"7", KNumText | KNumber, grid.Cell{}, false}, // 7 or "7"
		{"", KNumber, grid.Cell{}, false},             // an empty reference: 0 in Excel
		{"", KText, grid.Cell{Kind: grid.Text}, true},
		{"TRUE", KBool | KText, grid.Cell{Kind: grid.Bool, Bool: true}, true},
	} {
		got, ok := Typed(tt.raw, tt.k)
		if ok != tt.ok || got != tt.want {
			t.Errorf("Typed(%q, %04b) = %+v %v", tt.raw, tt.k, got, ok)
		}
	}
}

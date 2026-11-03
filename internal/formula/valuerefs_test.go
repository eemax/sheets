package formula

import (
	"fmt"
	"testing"
)

// ROWS, COLUMNS, ROW, COLUMN and ISREF read where a reference is, not
// what it holds.
func TestValueRefs(t *testing.T) {
	for f, want := range map[string]string{
		"ROWS($A$2:A2)":           "[]",
		"ROWS($A$2:A2)+B1":        "[B1]",
		"SUM(ROW(A1:A3)*C1)":      "[C1]",
		"COLUMNS('S x'!A1:B2)":    "[]",
		"ISREF(A1)":               "[]",
		"ROWS(OFFSET(A1,0,0,2))":  "[A1]",
		"ROWS(A1:A3,B1)":          "[]",
		"IF(ROWS(A1:A2)>1,C3,D4)": "[C3 D4]",
	} {
		var got []string
		for _, r := range ValueRefs(f) {
			got = append(got, r.Range.String())
		}
		if s := fmt.Sprint(got); s != want && !(s == "[]" && want == "[]") {
			t.Errorf("%s: %s, want %s", f, s, want)
		}
	}
}

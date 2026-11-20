package cli

import (
	"strings"
	"testing"
)

// TestReceiptMarkdown: values with Markdown's marks in them read as they
// are, in a table cell and in prose.
func TestReceiptMarkdown(t *testing.T) {
	for _, c := range []struct{ in, code, text string }{
		{"", "*empty*", ""},
		{`=IF(A1="a|b",1,2)`, "`=IF(A1=\"a\\|b\",1,2)`", `=IF(A1="a\|b",1,2)`},
		{"=B4*C3", "`=B4*C3`", `=B4\*C3`},
		{"a`b", "``a`b``", "a\\`b"},
		{"`x`", "`` `x` ``", "\\`x\\`"},
		{"two\nlines", "`two lines`", "two lines"},
		{"[link](x) <b> #1 _u_", "`[link](x) <b> #1 _u_`", `\[link\](x) \<b\> \#1 \_u\_`},
	} {
		if got := mdCode(c.in); got != c.code {
			t.Errorf("mdCode(%q) = %q, want %q", c.in, got, c.code)
		}
		if got := mdText(c.in); got != c.text {
			t.Errorf("mdText(%q) = %q, want %q", c.in, got, c.text)
		}
	}
	if got := mdLink("my receipt-1-before.png"); got != "<my receipt-1-before.png>" {
		t.Errorf("mdLink: %q", got)
	}
	if got := mdLink("r-1-before.png"); got != "r-1-before.png" {
		t.Errorf("mdLink: %q", got)
	}
}

// TestReceiptRegions: changes gather into at most three ranges, each of at
// most receiptRows by receiptCols, every sheet first; the rest are counted.
func TestReceiptRegions(t *testing.T) {
	var spots []spot
	for r := 2; r <= 200; r += 2 {
		spots = append(spots, spot{after: place{"Data", r, 2}, before: place{"Data", r, 2}, kind: "input changed"})
	}
	spots = append(spots, spot{after: place{"Summary", 3, 2}, before: place{"Summary", 3, 2}, kind: "value moved"})
	// Added on a sheet the earlier version has not: pictured after only.
	spots = append(spots, spot{after: place{"New", 1, 1}, kind: "input changed"})
	regs, left, where := regions(spots)
	if len(regs) != receiptPreviews {
		t.Fatalf("%d regions", len(regs))
	}
	var got []string
	for _, g := range regs {
		if g.after.Rows() > receiptRows+2 || g.after.Cols() > receiptCols+2 {
			t.Errorf("%s!%s is too large", g.sheet, g.after)
		}
		was := "nothing"
		if g.beforeSheet != "" {
			was = g.beforeSheet + "!" + g.before.String()
		}
		got = append(got, g.sheet+"!"+g.after.String()+" was "+was)
	}
	want := "Data!A1:C31 was Data!A1:C31; Summary!A1:C4 was Summary!A1:C4; New!A1:B2 was nothing"
	if strings.Join(got, "; ") != want {
		t.Errorf("regions:\n got %s\nwant %s", strings.Join(got, "; "), want)
	}
	if left != 100-15 || len(where) != causeListed || where[0] != "Data!B32" {
		t.Errorf("left %d, where %v", left, where)
	}
}

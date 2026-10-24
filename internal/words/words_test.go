package words

import (
	"reflect"
	"testing"
)

func TestSplit(t *testing.T) {
	for line, want := range map[string][]string{
		`set A1 5`:                    {"set", "A1", "5"},
		"  set   A1\t5  ":             {"set", "A1", "5"},
		`set A1 '=SUM(A1:A3)'`:        {"set", "A1", "=SUM(A1:A3)"},
		`set A1 "=IF(B1=\"x\",1,0)"`:  {"set", "A1", `=IF(B1="x",1,0)`},
		`style A1 format='"$"#,##0'`:  {"style", "A1", `format="$"#,##0`},
		`set A1 it\'s`:                {"set", "A1", "it's"},
		`set A1 'a b'c"d e"`:          {"set", "A1", "a bcd e"},
		`set A1 5 # a comment`:        {"set", "A1", "5"},
		`set A1 a#b`:                  {"set", "A1", "a#b"},
		`# only a comment`:            nil,
		`set A1 ''`:                   {"set", "A1", ""},
		`set A1 "back\slash"`:         {"set", "A1", `back\slash`},
		`view 'Q1 Sales' --freeze B2`: {"view", "Q1 Sales", "--freeze", "B2"},
		`set "Sheet 1!A1" "it's \$5"`: {"set", "Sheet 1!A1", "it's $5"},
	} {
		got, err := Texts(line)
		if err != nil {
			t.Errorf("Split(%q): %v", line, err)
			continue
		}
		if len(got) == 0 && len(want) == 0 {
			continue
		}
		if !reflect.DeepEqual(got, want) {
			t.Errorf("Split(%q) = %q, want %q", line, got, want)
		}
	}
	for _, bad := range []string{`set A1 'x`, `set A1 "x`} {
		if _, err := Split(bad); err == nil {
			t.Errorf("Split(%q): no error", bad)
		}
	}
}

func TestRaw(t *testing.T) {
	ws, _ := Split(`write A1 <<'EOF'`)
	if ws[2].Text != "<<EOF" || ws[2].Raw != `<<'EOF'` {
		t.Errorf("got %+v", ws[2])
	}
	ws, _ = Split(`set A1 '<<EOF'`)
	if ws[2].Raw != `'<<EOF'` {
		t.Errorf("got %+v", ws[2])
	}
}

func TestQuoteRoundTrip(t *testing.T) {
	if q := Quote("fill=#DDEBF7"); q != "fill=#DDEBF7" {
		t.Errorf("Quote quoted %s", q)
	}
	for _, w := range []string{"plain", "", "a b", "fill=#DDEBF7", `"$"#,##0.00_);("$"#,##0.00)`, "it's", `back\slash`, "x=1", "#hash", "Sales!A1:F1", "tab\there"} {
		got, err := Texts(Join([]string{"style", w}))
		if err != nil || len(got) != 2 || got[1] != w {
			t.Errorf("round trip of %q: %q (%v)", w, got, err)
		}
	}
}

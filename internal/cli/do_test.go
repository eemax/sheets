package cli

import (
	"reflect"
	"strings"
	"testing"
)

func TestParseScript(t *testing.T) {
	src := `# build
sheets write Data!A1 -i in.csv

set Data!F1 'Margin %' Data!F2 '=E2/D2'
write Data!G1 <<END
a	b
=1+2	x
END
write Data!H1 <<'EOF'
literal $x
EOF
write Data!I1 << DONE
1
DONE
	write Data!J1 <<-T
		tabbed
	T
write Data!K1 < data.tsv
write Data!L1 <data.tsv
style A1:F1 \
  bold fill=lightblue
set A1 '<<not a heredoc'
`
	steps, err := parseScript(src)
	if err != nil {
		t.Fatal(err)
	}
	type want struct {
		line  int
		args  string
		input string
		from  string
	}
	exp := []want{
		{2, "write|Data!A1|-i|in.csv", "", ""},
		{4, "set|Data!F1|Margin %|Data!F2|=E2/D2", "", ""},
		{5, "write|Data!G1", "a\tb\n=1+2\tx\n", ""},
		{9, "write|Data!H1", "literal $x\n", ""},
		{12, "write|Data!I1", "1\n", ""},
		{15, "write|Data!J1", "tabbed\n", ""},
		{18, "write|Data!K1", "", "data.tsv"},
		{19, "write|Data!L1", "", "data.tsv"},
		{20, "style|A1:F1|bold|fill=lightblue", "", ""},
		{22, "set|A1|<<not a heredoc", "", ""},
	}
	if len(steps) != len(exp) {
		t.Fatalf("%d steps, want %d: %+v", len(steps), len(exp), steps)
	}
	for i, st := range steps {
		got := want{st.line, strings.Join(st.args, "|"), string(st.input), st.from}
		if !reflect.DeepEqual(got, exp[i]) {
			t.Errorf("step %d = %+v, want %+v", i, got, exp[i])
		}
	}
	if steps[8].text != "style A1:F1" {
		t.Errorf("text of a continued line = %q", steps[8].text)
	}
	for _, bad := range []string{"write A1 <<END\nno end\n", "set A1 'open\n", "write A1 <<\n"} {
		if _, err := parseScript(bad); err == nil {
			t.Errorf("parseScript(%q): no error", bad)
		}
	}
}

func TestNegatives(t *testing.T) {
	a := &app{}
	root := a.root()
	for in, want := range map[string]string{
		"set b.xlsx A1 -5":                        "set b.xlsx A1 -5",
		"set b.xlsx A1 5":                         "set b.xlsx A1 5",
		"set --raw b.xlsx A1 -5":                  "set --raw -- b.xlsx A1 -5",
		"set b.xlsx A1 -5 --json":                 "set --json -- b.xlsx A1 -5",
		"cond b.xlsx B2:B9 between -5 and 5 bold": "cond -- b.xlsx B2:B9 between -5 and 5 bold",
		"validate b.xlsx B2 whole > -1 --error x": "validate --error x -- b.xlsx B2 whole > -1",
		"write b.xlsx -i f.tsv":                   "write b.xlsx -i f.tsv",
		"set b.xlsx A1 -- -5":                     "set b.xlsx A1 -- -5",
		"sheet add b.xlsx -1":                     "sheet add -- b.xlsx -1",
	} {
		got := strings.Join(negatives(root, strings.Fields(in)), " ")
		if in == "set b.xlsx A1 -5" {
			want = "set -- b.xlsx A1 -5"
		}
		if got != want {
			t.Errorf("negatives(%q) = %q, want %q", in, got, want)
		}
	}
}

func TestNumberedGroups(t *testing.T) {
	for in, want := range map[string]string{
		"$1TH":     "${1}TH",
		"$2 $1":    "${2} ${1}",
		"$$1":      "$$1",
		"${name}x": "${name}x",
		"cost: $":  "cost: $",
		"$12a":     "${12}a",
		"a$$$1":    "a$$${1}",
	} {
		if got := numberedGroups(in); got != want {
			t.Errorf("numberedGroups(%q) = %q, want %q", in, got, want)
		}
	}
}

package fuzz

import (
	"archive/zip"
	"bytes"
	"math/rand/v2"
	"os"
	"slices"
	"strings"
	"testing"

	"github.com/eemax/sheets/internal/grid"
)

func TestEscapeKinds(t *testing.T) {
	for _, tc := range []struct {
		text string
		want []string
	}{
		{"hello", nil},
		{"R&D", []string{"xml"}},
		{"<b>x</b>", []string{"xml"}},
		{`say "hi"`, []string{"quote"}},
		{"it's", []string{"quote"}},
		{" lead", []string{"space"}},
		{"trail ", []string{"space"}},
		{"two  spaces", nil},
		{"_x0041_", []string{"_x"}},
		{"ctl\x01x", []string{"control"}},
		{"line1\nline2", nil},
		{"tab\there", nil},
		{"Åäö 日本語", []string{"non-ASCII"}},
		{"<\x1f&>", []string{"xml", "control"}},
		{"Müller & Söhne", []string{"xml", "non-ASCII"}},
	} {
		if got := escapeKinds(tc.text); !slices.Equal(got, tc.want) {
			t.Errorf("escapeKinds(%q) = %v, want %v", tc.text, got, tc.want)
		}
	}
}

// TestValuesNeedEscaping: what the edits type, and the names they give
// sheets, hold each thing a writer must escape or quote, or the streaming
// editor's text writer is never put to it (the first version typed none).
func TestValuesNeedEscaping(t *testing.T) {
	seen := map[string]int{}
	for _, v := range append(slices.Clone(values), oddValues...) {
		for _, k := range escapeKinds(v) {
			seen[k]++
		}
	}
	for _, k := range EscapeKinds {
		if seen[k] == 0 {
			t.Errorf("no value typed by a set holds %s", k)
		}
	}
	// A batch line and a tab-separated row cannot hold a line break or a tab.
	for _, v := range values {
		if strings.ContainsAny(v, "\n\t") {
			t.Errorf("%q is in values, which a write and a batch use: it holds a line break or a tab", v)
		}
	}
	names := strings.Join(newNames, "")
	for _, c := range []string{"&", "'", "<", `"`, " "} {
		if !strings.Contains(names, c) {
			t.Errorf("no new sheet's name holds %q", c)
		}
	}
	for _, c := range []string{"&", "<", `"`, "'", " "} {
		if !slices.Contains(finds, c) {
			t.Errorf("replace never finds %q", c)
		}
	}
	if !slices.ContainsFunc(replaces, func(s string) bool { return strings.ContainsAny(s, "&<") }) {
		t.Error("replace never puts in & or <")
	}
}

// TestEditsTypeEscapes: the edits the generator makes put such text in set,
// write and scripts, and odd texts (a line break, a control character) only
// where one line or one row can hold them.
func TestEditsTypeEscapes(t *testing.T) {
	wb := poolBook(t, 0)
	rng := rand.New(rand.NewPCG(5, 6))
	seen := map[string]int{}
	odd := map[string]int{}
	for range 3000 {
		cp := *wb
		cp.Sheets = slices.Clone(wb.Sheets)
		g := &gen{rng: rng, wb: &cp}
		e := g.edit(pick(g, []string{"set", "write", "do", "script"}))
		for _, text := range typed(e) {
			for _, k := range escapeKinds(text) {
				seen[k]++
			}
		}
		if e.Class == "do" || e.Class == "write" {
			for _, line := range strings.Split(e.Stdin, "\n") {
				if strings.ContainsAny(line, "\x01\x1f") {
					odd[e.Class]++
				}
			}
		}
		if e.Class == "set" {
			for _, a := range e.Args {
				if strings.ContainsAny(a, "\n\x01\x1f\t") {
					odd["set"]++
				}
			}
		}
	}
	for _, k := range EscapeKinds {
		if seen[k] < 5 {
			t.Errorf("of 3000 edits, set and write typed %d texts with %s in them", seen[k], k)
		}
	}
	if odd["set"] == 0 {
		t.Error("no set typed a line break, a tab or a control character")
	}
	if odd["do"]+odd["write"] > 0 {
		t.Errorf("a batch or a write got a control character: %v", odd)
	}
}

func poolBook(t *testing.T, i int) *grid.Workbook {
	t.Helper()
	wb, err := readWorkbook(loadPool(t)[i].Data)
	if err != nil {
		t.Fatal(err)
	}
	return wb
}

// TestPoolEscapes: the pool holds texts and sheet names that need escaping,
// in files from both producers.
func TestPoolEscapes(t *testing.T) {
	texts := map[string]map[string]bool{}
	names := map[string]bool{}
	for _, w := range loadPool(t) {
		producer := strings.SplitN(w.Name, "-", 2)[0]
		wb, err := readWorkbook(w.Data)
		if err != nil {
			t.Fatal(err)
		}
		for _, s := range wb.Sheets {
			for _, k := range escapeKinds(s.Name) {
				names[k] = true
			}
			s.Each(used(s), func(_, _ int, c *grid.Cell) {
				if c.Kind == grid.Text && c.Formula == "" {
					for _, k := range escapeKinds(c.Text) {
						if texts[k] == nil {
							texts[k] = map[string]bool{}
						}
						texts[k][producer] = true
					}
				}
			})
		}
	}
	for _, k := range []string{"xml", "quote", "space", "non-ASCII"} {
		for _, p := range []string{"openpyxl", "xlsxwriter"} {
			if !texts[k][p] {
				t.Errorf("no %s file holds a text with %s in it", p, k)
			}
		}
	}
	for _, k := range []string{"xml", "quote"} {
		if !names[k] {
			t.Errorf("no sheet in the pool has %s in its name", k)
		}
	}
}

// TestShellWord: an argument with a control character is quoted in the
// $'...' form, which a report's indentation cannot change; others as words.Quote.
func TestShellWord(t *testing.T) {
	for _, tc := range []struct{ in, want string }{
		{"Data", "Data"},
		{"two words", "'two words'"},
		{"line1\nline2", `$'line1\nline2'`},
		{"tab\there", `$'tab\there'`},
		{"ctl\x01x", `$'ctl\x01x'`},
		{"it's\n", `$'it\'s\n'`},
		{`a\b` + "\n", `$'a\\b\n'`},
	} {
		if got := shellWord(tc.in); got != tc.want {
			t.Errorf("shellWord(%q) = %s, want %s", tc.in, got, tc.want)
		}
	}
	e := Edit{Class: "set", Args: []string{"set", FileArg, "Data!A1", "line1\nline2"}}
	if got, want := e.String(), `sheets set book.xlsx 'Data!A1' $'line1\nline2'`; got != want {
		t.Errorf("Edit.String() = %s, want %s", got, want)
	}
}

// TestSortOK: the model of what sheets refuses to sort.
func TestSortOK(t *testing.T) {
	s := grid.NewSheet("Data")
	for r := 1; r <= 8; r++ {
		s.Set(r, 1, grid.Cell{Kind: grid.Text, Text: "x"})
		s.Set(r, 2, grid.Cell{Kind: grid.Number, Num: float64(r)})
		s.Set(r, 3, grid.Cell{Kind: grid.Number, Num: 1, Formula: `IF(N("p")=0,B` + string(rune('0'+r)) + `*2)`})
	}
	other := grid.NewSheet("Calc")
	for r := 2; r <= 4; r++ {
		other.Set(r, 1, grid.Cell{Kind: grid.Text, Text: "k"})
		other.Set(r, 2, grid.Cell{Kind: grid.Number, Num: 1, Formula: `IF(N("p")=0,Data!B2*2)`})
	}
	wb := &grid.Workbook{Sheets: []*grid.Sheet{s, other}, Names: []grid.DefinedName{{Name: "Rate"}}}
	rg := func(s string) grid.Range { r, _ := grid.ParseRange(s); return r }
	check := func(name string, sh *grid.Sheet, r string, want bool) {
		t.Helper()
		if got := sortOK(wb, sh, rg(r)); got != want {
			t.Errorf("%s: sortOK(%s, %s) = %v, want %v", name, sh.Name, r, got, want)
		}
	}
	check("rows with formulas reading their own row", s, "A2:C8", true)
	check("columns without formulas", s, "A2:B8", true)
	check("a single row", s, "A2:C2", false)
	check("a formula reading another sheet by its row", other, "A2:B4", false)
	check("the column without formulas of that sheet", other, "A2:A4", true)
	s.Set(5, 4, grid.Cell{Kind: grid.Number, Num: 1, Formula: `IF(N("p")=0,SUM(B2:B4))`})
	check("a formula reading a range", s, "A2:D8", false)
	check("... left out of the columns", s, "A2:C8", true)
	s.Set(5, 4, grid.Cell{Kind: grid.Number, Num: 1, Formula: `IF(N("p")=0,B4*2)`})
	check("a formula reading another row", s, "A2:D8", false)
	s.Set(5, 4, grid.Cell{Kind: grid.Number, Num: 1, Formula: `IF(N("p")=0,B5*Rate)`})
	check("a formula using a name", s, "A2:D8", false)
	s.Set(5, 4, grid.Cell{})
	s.Spills = []grid.Range{rg("F3:F4")}
	check("a spill", s, "A2:F8", false)
	check("... beside it", s, "A2:C8", true)
	s.Spills = []grid.Range{rg("F3")}
	s.Set(3, 6, grid.Cell{Kind: grid.Number, Num: 1, Formula: `IF(N("d")=0,B3*2)`})
	check("a dynamic array in one cell", s, "A2:F8", true)
	s.Spills = []grid.Range{rg("F3:F4")}
	s.Set(3, 6, grid.Cell{})
	s.Arrays = []grid.Range{rg("E2:E3"), rg("G4")}
	check("a legacy array over several cells", s, "A2:E8", false)
	check("a legacy array over one", s, "A2:C8", true)
	s.Tables = []grid.NamedTable{{Name: "T", Range: rg("A1:C8")}}
	check("part of a table's rows", s, "A2:C5", true)
	check("part of a table's columns", s, "A2:B8", false)
	check("the header row", s, "A1:C8", false)
	check("all of its data rows and columns", s, "A2:C8", true)
}

// TestSortsAreDone: the sorts the generator picks as sortable are done by
// sheets, in nearly all cases: on each pool file, 8 sorts, run with both
// editors, are refused rarely. (Of the first version's random sorts 97 in 137
// were refused, so the sort class was barely judged.)
func TestSortsAreDone(t *testing.T) {
	if raceOn {
		t.Skip("whether sheets does a sort is no matter of data races, and 320 sorts take 39 s under -race")
	}
	pool := loadPool(t)
	rng := rand.New(rand.NewPCG(9, 10))
	done, refused, none := 0, 0, 0
	for _, w := range pool {
		wb, err := readWorkbook(w.Data)
		if err != nil {
			t.Fatal(err)
		}
		s, err := newSession(t.TempDir(), w.Data)
		if err != nil {
			t.Fatal(err)
		}
		for range 8 {
			cp := *wb
			cp.Sheets = slices.Clone(wb.Sheets)
			g := &gen{rng: rng, wb: &cp}
			args := g.sortable(pick(g, wb.Sheets))
			if args == nil {
				none++
				continue
			}
			e := Edit{Class: "sort", Args: append([]string{"sort", FileArg}, args...)}
			for k, ed := range editors {
				if err := os.WriteFile(s.paths[k], w.Data, 0o644); err != nil {
					t.Fatal(err)
				}
				code, msg := run(ed.env, false, e, s.paths[k])
				if k == 1 {
					continue // one editor suffices: sort is full-editor code in both
				}
				if code == 0 {
					done++
				} else {
					refused++
					t.Logf("%s: %s: %s", w.Name, e, firstLine(msg))
				}
			}
		}
		s.close()
	}
	t.Logf("sortable sorts: %d done, %d refused; %d times no sort to pick", done, refused, none)
	if done < 150 || refused*5 > done {
		t.Errorf("of %d sorts the generator picked as sortable, sheets refused %d", done+refused, refused)
	}
}

// TestCheckSpace: a text that starts or ends with white space must say
// xml:space="preserve".
func TestCheckSpace(t *testing.T) {
	const head = `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData>`
	part := func(cells string) *zip.File {
		var buf bytes.Buffer
		zw := zip.NewWriter(&buf)
		w, _ := zw.Create("xl/worksheets/sheet1.xml")
		w.Write([]byte(head + `<row r="1">` + cells + `</row></sheetData></worksheet>`))
		zw.Close()
		zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
		if err != nil {
			t.Fatal(err)
		}
		return zr.File[0]
	}
	for _, tc := range []struct {
		name, cells string
		ok          bool
	}{
		{"plain text", `<c r="A1" t="inlineStr"><is><t>hello</t></is></c>`, true},
		{"spaces inside", `<c r="A1" t="inlineStr"><is><t>two  spaces</t></is></c>`, true},
		{"a leading space, preserved", `<c r="A1" t="inlineStr"><is><t xml:space="preserve"> lead</t></is></c>`, true},
		{"a trailing space, preserved", `<c r="A1" t="inlineStr"><is><t xml:space="preserve">trail </t></is></c>`, true},
		{"a leading space", `<c r="A1" t="inlineStr"><is><t> lead</t></is></c>`, false},
		{"a trailing space", `<c r="A1" t="inlineStr"><is><t>trail </t></is></c>`, false},
		{"a trailing line break", `<c r="A1" t="inlineStr"><is><t>x&#xA;</t></is></c>`, false},
		{"a space in a run of rich text", `<c r="A1" t="inlineStr"><is><r><t>a</t></r><r><t> b</t></r></is></c>`, false},
		{"a run, preserved", `<c r="A1" t="inlineStr"><is><r><t>a</t></r><r><t xml:space="preserve"> b</t></r></is></c>`, true},
		{"a formula's result with spaces is no <t>", `<c r="A1" t="str"><f>x</f><v> a </v></c>`, true},
	} {
		err := checkSpace(part(tc.cells))
		if tc.ok && err != nil {
			t.Errorf("%s: refused: %v", tc.name, err)
		}
		if !tc.ok && err == nil {
			t.Errorf("%s: not refused", tc.name)
		}
	}
}

// TestCheckTyped: a text a set or write typed is stored as typed.
func TestCheckTyped(t *testing.T) {
	s := grid.NewSheet("Data")
	wb := &grid.Workbook{Sheets: []*grid.Sheet{s}}
	put := func(ref, text string) {
		r, c, _ := grid.ParseCell(ref)
		s.Set(r, c, grid.Cell{Kind: grid.Text, Text: text})
	}
	put("A2", "R&D")
	put("A3", "A") // typed _x0041_, stored unescaped: Excel reads "A"
	put("B2", " lead")
	put("C2", "5")
	s.Set(2, 4, grid.Cell{Kind: grid.Number, Num: 5})
	s.Tables = []grid.NamedTable{{Name: "T", Range: grid.Range{R1: 1, C1: 6, R2: 4, C2: 7}}}
	put("F1", "Name2") // a table's header: sheets may name the column otherwise
	set := func(args ...string) Edit {
		return Edit{Class: "set", Args: append([]string{"set", FileArg}, args...)}
	}
	for _, tc := range []struct {
		name string
		e    Edit
		want string // part of the message; "" when the text is stored as typed
	}{
		{"text", set("Data!A2", "R&D"), ""},
		{"several cells", set("Data!A2", "R&D", "Data!B2", " lead"), ""},
		{"an escape stored unescaped", set("Data!A3", "_x0041_"), `typed the text "_x0041_", and the cell holds text "A"`},
		{"a number is no text to judge", set("Data!D2", "5"), ""},
		{"a formula is no text to judge", set("Data!A2", "=IF(N(\"p\")=0,1)"), ""},
		{"the last typing wins", set("Data!A3", "_x0041_", "Data!A3", "5"), ""},
		{"a table's header is not judged", set("Data!F1", "Name"), ""},
		{"a text the cell lost", set("Data!E2", "R&D"), `typed the text "R&D", and the cell holds nothing`},
		{"a range", set("Data!A2:A3", "R&D"), `Data!A3: typed the text "R&D"`},
		{"write", Edit{Class: "write", Args: []string{"write", FileArg, "Data!A2", "--from", "tsv"}, Stdin: "R&D\t lead\n=1\t\n"}, ""},
		{"write loses a cell", Edit{Class: "write", Args: []string{"write", FileArg, "Data!A2", "--from", "tsv"}, Stdin: "R&D\t lead\n_x0041_\n"}, `Data!A3: typed the text "_x0041_"`},
		{"other edits are not judged", Edit{Class: "clear", Args: []string{"clear", FileArg, "Data!A2"}}, ""},
	} {
		got := checkTyped(tc.e, wb)
		if tc.want == "" && got != "" {
			t.Errorf("%s: refused: %s", tc.name, got)
		}
		if tc.want != "" && !strings.Contains(got, tc.want) {
			t.Errorf("%s: says %q, want it to say %q", tc.name, got, tc.want)
		}
	}
}

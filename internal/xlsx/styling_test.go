package xlsx

import (
	"regexp"
	"sort"
	"strings"
	"testing"

	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/input"
	"github.com/eemax/sheets/internal/style"
	"github.com/xuri/excelize/v2"
)

func TestParseLook(t *testing.T) {
	for _, c := range []struct {
		in   string
		want func(Look) bool
	}{
		{"bold fill=lightblue", func(l Look) bool { return *l.Bold && l.Fill == "DDEBF7" }},
		{"nobold color=#c00", func(l Look) bool { return !*l.Bold && l.Color == "CC0000" }},
		{"fill=none", func(l Look) bool { return l.Fill == "none" }},
		{"color=Light-Blue", func(l Look) bool { return l.Color == "DDEBF7" }},
		{"format=currency decimals=0", func(l Look) bool { return l.NumFmt == `"$"#,##0_);("$"#,##0)` }},
		{"format=currency", func(l Look) bool { return l.NumFmt == `"$"#,##0.00_);("$"#,##0.00)` }},
		{"format=currency:€ decimals=0", func(l Look) bool { return l.NumFmt == `"€"#,##0_);("€"#,##0)` }},
		{"format=percent decimals=1", func(l Look) bool { return l.NumFmt == "0.0%" }},
		{"format=percent", func(l Look) bool { return l.NumFmt == "0%" }},
		{"format=date", func(l Look) bool { return l.NumFmt == "yyyy-mm-dd" }},
		{"format=general", func(l Look) bool { return l.NumFmt == "General" }},
		{`format=0.0"x"`, func(l Look) bool { return l.NumFmt == `0.0"x"` }},
		{"format=#,##0;[Red](#,##0)", func(l Look) bool { return l.NumFmt == "#,##0;[Red](#,##0)" }},
		{"align=center-across valign=middle", func(l Look) bool { return l.Align == "centerContinuous" && l.VAlign == "center" }},
		{"rotate=-45", func(l Look) bool { return *l.Rotate == 135 }},
		{"rotate=90", func(l Look) bool { return *l.Rotate == 90 }},
		{"border=top+bottom:double:navy", func(l Look) bool {
			return len(l.Borders) == 2 && l.Borders[0].Sides == "top" && l.Borders[1].Style == 6 && l.Borders[1].Color == "1F3864"
		}},
		{"border=none", func(l Look) bool {
			return len(l.Borders) == 1 && l.Borders[0].Style == 0 && l.Borders[0].Sides == "all"
		}},
		{"width=auto height=20", func(l Look) bool { return l.Width == -1 && l.Height == 20 }},
		{"merge", func(l Look) bool { return *l.Merge && !l.CellStyle() }},
		{"reset", func(l Look) bool { return l.Reset && l.CellStyle() }},
		{"BOLD Size=14 font=Arial", func(l Look) bool { return *l.Bold && l.Size == 14 && l.Font == "Arial" }},
	} {
		l, err := ParseLook(strings.Fields(c.in))
		if err != nil || !c.want(l) {
			t.Errorf("ParseLook(%q) = %+v, %v", c.in, l, err)
		}
	}
	for _, bad := range []string{"blod", "fill=bluish", "size=0", "size=500", "format=abc", `format="x`, "decimals=2",
		"format=date decimals=x", "align=middle", "border=diagonal", "border=top:heavy", "rotate=120", "width=-3", "merge align=center-across", "color="} {
		if _, err := ParseLook(strings.Fields(bad)); err == nil {
			t.Errorf("ParseLook(%q): no error", bad)
		}
	}
}

func TestCheckNumFmt(t *testing.T) {
	for code, ok := range map[string]bool{
		"#,##0.00": true, `"$"#,##0_);("$"#,##0)`: true, "yyyy-mm-dd hh:mm:ss": true, "0.0%": true, "@": true,
		"[Red]0;[Blue]-0": true, `0 "units"`: true, `\x0`: true, "h:mm AM/PM": true, "[h]:mm": true, "General": true,
		"0;0;0;@": true, "#,##0 €": true, "mmm d, yyyy": true,
		"abc": false, "0;0;0;0;0": false, `"x`: false, "[Red": false, "0 units": false, `0\`: false, "": false,
	} {
		if err := CheckNumFmt(code); (err == nil) != ok {
			t.Errorf("CheckNumFmt(%q) = %v, want ok=%v", code, err, ok)
		}
	}
}

func TestContrast(t *testing.T) {
	if readableOn("1F2937") != "FFFFFF" || readableOn("DDEBF7") != "000000" || readableOn("FFFF00") != "000000" {
		t.Error("readableOn picked the wrong text colour")
	}
	if c := contrast("000000", "FFFFFF"); c < 20.9 || c > 21.1 {
		t.Errorf("contrast(black, white) = %v", c)
	}
}

// change opens path, runs fn and saves.
func change(t *testing.T, path string, fn func(e *Editor)) {
	t.Helper()
	e, err := OpenEditor(path, false)
	if err != nil {
		t.Fatal(err)
	}
	defer e.Close()
	fn(e)
	if err := e.Save(false); err != nil {
		t.Fatal(err)
	}
}

func must(t *testing.T, err error) {
	t.Helper()
	if err != nil {
		t.Fatal(err)
	}
}

func look(t *testing.T, s string) *Look {
	t.Helper()
	l, err := ParseLook(strings.Fields(s))
	must(t, err)
	return &l
}

func rng(s string) grid.Range {
	r, err := grid.ParseRange(s)
	if err != nil {
		panic(err)
	}
	return r
}

func format(t *testing.T, path string) {
	t.Helper()
	if _, err := Format(path, FormatOptions{Preset: style.Default}); err != nil {
		t.Fatal(err)
	}
}

// cellLook reads a cell's style through excelize.
func cellLook(t *testing.T, path, sheet, cell string) *excelize.Style {
	t.Helper()
	f, err := excelize.OpenFile(path)
	must(t, err)
	defer f.Close()
	id, err := f.GetCellStyle(sheet, cell)
	must(t, err)
	s, err := f.GetStyle(id)
	must(t, err)
	return s
}

func styleFill(s *excelize.Style) string {
	if len(s.Fill.Color) == 0 {
		return ""
	}
	c := strings.TrimPrefix(s.Fill.Color[0], "#")
	return strings.ToUpper(c[len(c)-6:])
}

func budget(t *testing.T) string {
	return writeBook(t, [][]string{
		{"Item", "Budget", "Spent", "Remaining"},
		{"Rent", "1200", "1100", "=B{r}-C{r}"},
		{"Staff", "800", "850", "=B{r}-C{r}"},
		{"IT", "150", "90.5", "=B{r}-C{r}"},
	})
}

func TestStyleKeptByFmt(t *testing.T) {
	path := budget(t)
	format(t, path)
	change(t, path, func(e *Editor) {
		res, err := e.Style("data", rng("A1:D1"), look(t, "bold fill=lightblue"))
		must(t, err)
		if !res.Recorded || res.Range != "Data!A1:D1" || len(res.Notes) != 1 || !strings.Contains(res.Notes[0], "black") {
			t.Errorf("Style = %+v", res)
		}
		_, err = e.Style("Data", rng("B:D"), look(t, "format=currency decimals=0"))
		must(t, err)
	})
	check := func(when string) {
		h := cellLook(t, path, "Data", "B1")
		if styleFill(h) != "DDEBF7" || h.Font == nil || !h.Font.Bold || !strings.HasSuffix(strings.ToUpper(h.Font.Color), "000000") {
			t.Errorf("%s: header style = fill %s font %+v", when, styleFill(h), h.Font)
		}
		for _, c := range []string{"B2", "C4", "D3"} {
			if s := cellLook(t, path, "Data", c); s.CustomNumFmt == nil || *s.CustomNumFmt != `"$"#,##0_);("$"#,##0)` {
				t.Errorf("%s: %s format = %v", when, c, s.CustomNumFmt)
			}
		}
	}
	check("after style")
	format(t, path)
	check("after fmt")
	if same, _, err := Check(path, FormatOptions{Preset: style.Default}); err != nil || !same {
		t.Errorf("fmt is not idempotent with recorded styling: %v %v", same, err)
	}
	// The header keeps the house font size and border under the record.
	if h := cellLook(t, path, "Data", "A1"); len(h.Border) == 0 || h.Alignment == nil || h.Alignment.Horizontal != "left" {
		t.Errorf("house header style lost: %+v %+v", h.Border, h.Alignment)
	}
	// --reset drops the record.
	if _, err := Format(path, FormatOptions{Preset: style.Default, Reset: true}); err != nil {
		t.Fatal(err)
	}
	if h := cellLook(t, path, "Data", "B1"); styleFill(h) != style.Default.HeaderFill {
		t.Errorf("after fmt --reset the header fill is %s", styleFill(h))
	}
	e, err := OpenEditor(path, false)
	must(t, err)
	defer e.Close()
	if recs := e.Records(); len(recs) != 0 {
		t.Errorf("records after reset: %+v", recs)
	}
}

func TestRecordNamesHidden(t *testing.T) {
	path := budget(t)
	change(t, path, func(e *Editor) {
		_, err := e.Style("Data", rng("A1:D1"), look(t, "bold"))
		must(t, err)
	})
	wbx := sheetXML(t, path, "xl/workbook.xml")
	if !regexp.MustCompile(`<definedName hidden="1" comment="style bold" name="_sheets\.000001">Data!\$A\$1:\$D\$1</definedName>`).MatchString(wbx) {
		t.Errorf("record name not hidden or wrong:\n%s", wbx)
	}
	// It survives another edit (excelize re-reads hidden as "true").
	change(t, path, func(e *Editor) { must(t, e.Set("Data", 9, 1, inputText("x"))) })
	if !strings.Contains(sheetXML(t, path, "xl/workbook.xml"), `name="_sheets.000001"`) || strings.Count(sheetXML(t, path, "xl/workbook.xml"), "hidden=") != 1 {
		t.Errorf("record name after another edit:\n%s", sheetXML(t, path, "xl/workbook.xml"))
	}
}

func TestSupersede(t *testing.T) {
	path := budget(t)
	change(t, path, func(e *Editor) {
		for _, s := range []struct{ r, l string }{
			{"A1:D1", "fill=lightblue"},
			{"A1:B1", "bold"},
			{"A1:D1", "fill=lightgreen bold"}, // replaces both
			{"B2:B4", "format=integer"},
			{"B2:B4", "width=12"}, // different property: stays
			{"A1:D9", "reset"},    // clears looks: drops fill/bold and format, keeps width
			{"A1:D1", "color=red"},
		} {
			_, err := e.Style("Data", rng(s.r), look(t, s.l))
			must(t, err)
		}
		var got []string
		for _, r := range e.Records() {
			got = append(got, r.Range().String()+" "+r.Text())
		}
		want := []string{"B2:B4 style width=12", "A1:D9 style reset", "A1:D1 style color=red"}
		if strings.Join(got, "|") != strings.Join(want, "|") {
			t.Errorf("records = %q, want %q", got, want)
		}
	})
}

func TestCondKeptByFmtAndOrdered(t *testing.T) {
	path := budget(t)
	format(t, path)
	change(t, path, func(e *Editor) {
		_, err := e.Cond("Data", rng("D2:D4"), []string{"< 0", "color=red", "bold"})
		must(t, err)
		// The same again changes nothing.
		_, err = e.Cond("Data", rng("D2:D4"), []string{"< 0", "color=red", "bold"})
		must(t, err)
		if n := len(e.Records()); n != 1 {
			t.Errorf("%d records after the same cond twice", n)
		}
	})
	rules := func() []string {
		return regexp.MustCompile(`<cfRule [^>]*>`).FindAllString(sheetXML(t, path, dataPart), -1)
	}
	want := func(when string) {
		x := sheetXML(t, path, dataPart)
		if !strings.Contains(x, `<conditionalFormatting sqref="D2:D4"><cfRule type="cellIs" dxfId="1" priority="1" operator="lessThan"><formula>0</formula></cfRule></conditionalFormatting>`) {
			t.Errorf("%s: highlight rule missing or not first:\n%s", when, x)
		}
		if !strings.Contains(x, `<conditionalFormatting sqref="A2:D4"><cfRule type="expression" dxfId="0" priority="2"><formula>MOD(ROW()-2,2)=1</formula>`) {
			t.Errorf("%s: banding missing or not last:\n%s", when, x)
		}
		if n := len(rules()); n != 2 {
			t.Errorf("%s: %d rules", when, n)
		}
	}
	want("after cond")
	format(t, path)
	want("after fmt")
	format(t, path)
	want("after fmt twice")
	change(t, path, func(e *Editor) {
		cfs, err := e.CondFormats("Data")
		must(t, err)
		if len(cfs) != 2 || cfs[0].Rule != "=MOD(ROW()-2,2)=1" || !cfs[0].Banding || cfs[1].Rule != "< 0" || !cfs[1].Recorded {
			t.Errorf("CondFormats = %+v", cfs)
		}
		// Only what lies within the target goes: not the banding over it.
		n, err := e.ClearCond("Data", &grid.Range{R1: 3, C1: 4, R2: 3, C2: 4})
		must(t, err)
		if n != 0 {
			t.Errorf("ClearCond(D3) removed %d rules", n)
		}
		n, err = e.ClearCond("Data", &grid.Range{R1: 1, C1: 4, R2: 9, C2: 4})
		must(t, err)
		if n != 1 {
			t.Errorf("ClearCond(D1:D9) removed %d rules", n)
		}
		if len(e.Records()) != 0 {
			t.Errorf("records after clear: %+v", e.Records())
		}
	})
	if x := sheetXML(t, path, dataPart); strings.Contains(x, `sqref="D2:D4"`) || !strings.Contains(x, `sqref="A2:D4"`) {
		t.Errorf("after clearing D1:D9:\n%s", x)
	}
}

// Every kind of rule survives fmt removing a rule that overlaps it.
func TestCondRoundTrip(t *testing.T) {
	for _, rule := range [][]string{
		{"< 0", "color=red"}, {"between 1 and 10"}, {"not between -5 and 5", "fill=lightyellow"}, {"= East"},
		{"top 10%"}, {"bottom 3", "bold"}, {"above average"}, {"below average"}, {"duplicates"}, {"unique"},
		{"blank"}, {"not blank"}, {"errors"}, {"no errors"}, {"contains", "x"}, {"not contains", "y"},
		{"begins with", "A"}, {"ends with", ".pdf"}, {"=$B2>1000", "fill=lightgreen"}, {"scale"}, {"scale", "white-blue"},
		{"scale", "green-yellow-red"}, {"bars"}, {"bars", "orange"}, {"icons"}, {"icons", "traffic"}, {"> 2026-01-01"},
		{">= 10%", "format=0.0%"}, {"<> 0", "border=bottom:thin:red"},
	} {
		path := budget(t)
		change(t, path, func(e *Editor) {
			_, err := e.Cond("Data", rng("C2:C4"), rule)
			must(t, err)
		})
		block := regexp.MustCompile(`(?s)<conditionalFormatting sqref="C2:C4">.*?</conditionalFormatting>`)
		before := block.FindString(sheetXML(t, path, dataPart))
		sheetExt := regexp.MustCompile(`(?s)<extLst><ext uri="\{78C0D931[^"]*">.*?</ext></extLst>`)
		ext := sheetExt.FindString(sheetXML(t, path, dataPart))
		if before == "" {
			t.Errorf("%q: no rule written", rule)
			continue
		}
		format(t, path) // adds banding over A2:D4, removing and restoring the rule
		x := sheetXML(t, path, dataPart)
		after := block.FindString(x)
		norm := regexp.MustCompile(` priority="\d+"`)
		if norm.ReplaceAllString(before, "") != norm.ReplaceAllString(after, "") {
			t.Errorf("%q changed through fmt:\n%s\n%s", rule, before, after)
		}
		if a := sheetExt.FindString(x); a != ext {
			t.Errorf("%q: extension changed through fmt:\n%s\n%s", rule, ext, a)
		}
	}
}

func TestCondBadRules(t *testing.T) {
	r := rng("B2:B9")
	for _, bad := range [][]string{{}, {"<"}, {"between", "1"}, {"top", "x"}, {"top", "200%"}, {"contains"}, {"=SUM("},
		{"scale", "purplish-blue"}, {"icons", "hearts"}, {"bars", "bold"}, {"< 0", "size=20"}, {"weird"}, {"< 0", "blod"}} {
		if _, err := parseRule(bad, r); err == nil {
			t.Errorf("parseRule(%q): no error", bad)
		}
	}
	ru, err := parseRule([]string{"contains", `say "hi"`}, r)
	if err != nil || ru.opts.Criteria != `ISNUMBER(SEARCH("say ""hi""",B2))` {
		t.Errorf("contains with quotes: %+v %v", ru.opts, err)
	}
	ru, err = parseRule([]string{"<0"}, r)
	if err != nil || ru.opts.Criteria != "<" || ru.opts.Value != "0" {
		t.Errorf("<0: %+v %v", ru.opts, err)
	}
	ru, err = parseRule([]string{"value", ">=", "=AVERAGE($B$2:$B$9)"}, r)
	if err != nil || ru.opts.Criteria != ">=" || ru.opts.Value != "AVERAGE($B$2:$B$9)" {
		t.Errorf("value >= formula: %+v %v", ru.opts, err)
	}
}

// Inserting rows moves the record's range, the rule's range and its
// formula together, as Excel does.
func TestCondFollowsRows(t *testing.T) {
	path := budget(t)
	change(t, path, func(e *Editor) {
		_, err := e.Cond("Data", rng("A2:D4"), []string{"=$D2<0", "fill=lightred"})
		must(t, err)
		_, err = e.Style("Data", rng("A1:D1"), look(t, "bold"))
		must(t, err)
	})
	change(t, path, func(e *Editor) { must(t, e.InsertRows("Data", 1, 2)) })
	x := sheetXML(t, path, dataPart)
	if !strings.Contains(x, `<conditionalFormatting sqref="A4:D6"><cfRule type="expression" dxfId="0" priority="1"><formula>$D4&lt;0</formula>`) {
		t.Errorf("rule after inserting rows:\n%s", x)
	}
	e, err := OpenEditor(path, false)
	must(t, err)
	defer e.Close()
	recs := e.Records()
	if len(recs) != 2 || recs[0].Range().String() != "A4:D6" || recs[1].Range().String() != "A3:D3" {
		t.Errorf("records after inserting rows: %+v", recs)
	}
	cfs, _ := e.CondFormats("Data")
	if len(cfs) != 1 || !cfs[0].Recorded {
		t.Errorf("the moved rule is no longer recognised as recorded: %+v", cfs)
	}
	// Deleting the rows a record covers drops it.
	must(t, e.DeleteRows("Data", 3, 3))
	must(t, e.Save(false))
	e2, err := OpenEditor(path, false)
	must(t, err)
	defer e2.Close()
	if recs := e2.Records(); len(recs) != 1 || recs[0].Kind != "cond" {
		t.Errorf("records after deleting the styled row: %+v", recs)
	}
}

func TestValidate(t *testing.T) {
	path := budget(t)
	change(t, path, func(e *Editor) {
		_, err := e.Validate("Data", rng("A2:A10"), ValidationRule{Words: []string{"list", "Rent,Staff,IT"}})
		must(t, err)
		_, err = e.Validate("Data", rng("B2:B10"), ValidationRule{Words: []string{"whole", ">=", "0"}, Message: "No negatives"})
		must(t, err)
		_, err = e.Validate("Data", rng("C2:C10"), ValidationRule{Words: []string{"number", "between", "0", "and", "1e6"}, Warn: true, Prompt: "Spent so far"})
		must(t, err)
		_, err = e.Validate("Data", rng("D2:D10"), ValidationRule{Words: []string{"formula", `=AND(D2<>"",D2<5)`}})
		must(t, err)
		// A rule inside another's range takes those cells from it.
		_, err = e.Validate("Data", rng("A5:A6"), ValidationRule{Words: []string{"date", ">=", "2026-01-01"}})
		must(t, err)
		vs, err := e.Validations("Data")
		must(t, err)
		got := map[string]string{}
		for _, v := range vs {
			got[v.Range] = v.Rule
		}
		want := map[string]string{"A2:A4 A7:A10": "list Rent,Staff,IT", "B2:B10": "whole >= 0", "C2:C10": "number between 0 and 1e6",
			"D2:D10": `formula =AND(D2<>"",D2<5)`, "A5:A6": "date >= DATE(2026,1,1)"}
		for k, v := range want {
			if got[k] != v {
				t.Errorf("rule on %s = %q, want %q (all: %v)", k, got[k], v, got)
			}
		}
	})
	x := sheetXML(t, path, dataPart)
	for _, want := range []string{`showErrorMessage="true"`, `errorStyle="warning"`, `prompt="Spent so far"`,
		`<formula1>AND(D2&lt;&gt;&#34;&#34;,D2&lt;5)</formula1>`, `error="No negatives"`} {
		if !strings.Contains(x, want) {
			t.Errorf("validation XML lacks %s:\n%s", want, x)
		}
	}
	for _, bad := range [][]string{{"list"}, {"whole", ">", "x"}, {"date", ">", "tomorrow"}, {"length", "between", "1"}, {"formula", "=SUM("}, {"any"}, {"maybe"}, {"list", `a"b,c`}} {
		if _, err := (ValidationRule{Words: bad}).dataValidation(rng("A1")); err == nil {
			t.Errorf("validation %q: no error", bad)
		}
	}
}

func TestViewKeptByFmt(t *testing.T) {
	path := budget(t)
	format(t, path)
	change(t, path, func(e *Editor) {
		at := rng("B2")
		tab := "FF0000"
		z := 120
		must(t, e.View("Data", ViewOptions{Freeze: &at, Gridlines: ptr(true), Tab: &tab, Zoom: &z, Hide: []grid.Range{rng("C:C")}}))
	})
	format(t, path)
	e, err := OpenEditor(path, false)
	must(t, err)
	defer e.Close()
	v, err := e.ViewOf("Data")
	must(t, err)
	if v.Freeze != "B2" || !v.Gridlines || v.Tab != "#FF0000" || v.Zoom != 120 || len(v.HiddenCols) != 1 || v.HiddenCols[0] != "C" {
		t.Errorf("view after fmt = %+v", v)
	}
	must(t, e.View("Data", ViewOptions{Freeze: &grid.Range{R1: 1, C1: 1, R2: 1, C2: 1}, Tab: ptr(""), NoFilter: true}))
	must(t, e.Save(false))
	e2, err := OpenEditor(path, false)
	must(t, err)
	defer e2.Close()
	v, _ = e2.ViewOf("Data")
	if v.Freeze != "" || v.Tab != "" || v.Filter != "" {
		t.Errorf("view after unfreeze, no tab, no filter = %+v", v)
	}
	if strings.Contains(sheetXML(t, path, dataPart), "tabColor") {
		t.Error("tab colour left in the sheet")
	}
}

func TestMergeSafety(t *testing.T) {
	path := budget(t)
	change(t, path, func(e *Editor) {
		if _, err := e.Style("Data", rng("A1:B1"), look(t, "merge")); err == nil || !strings.Contains(err.Error(), "would hide B1") {
			t.Errorf("merge over values: %v", err)
		}
		_, err := e.Style("Data", rng("F1:H1"), look(t, "merge align=center"))
		must(t, err)
		// Merges are not recorded: they stay through fmt on their own.
		if recs := e.Records(); len(recs) != 1 || recs[0].Text() != "style align=center" {
			t.Errorf("records = %+v", recs)
		}
	})
	if !strings.Contains(sheetXML(t, path, dataPart), `<mergeCell ref="F1:H1">`) {
		t.Error("no merge written")
	}
}

var order = map[string]int{"top": 0, "bottom": 1, "left": 2, "right": 3}

func TestBorders(t *testing.T) {
	path := budget(t)
	change(t, path, func(e *Editor) {
		_, err := e.Style("Data", rng("A1:D4"), look(t, "border=outline:medium"))
		must(t, err)
	})
	sides := func(cell string) string {
		var out []string
		for _, b := range cellLook(t, path, "Data", cell).Border {
			if b.Style == 2 {
				out = append(out, b.Type)
			}
		}
		sort.Slice(out, func(i, j int) bool { return order[out[i]] < order[out[j]] })
		return strings.Join(out, ",")
	}
	for cell, want := range map[string]string{"A1": "top,left", "D1": "top,right", "B2": "", "A4": "bottom,left", "D4": "bottom,right", "C4": "bottom"} {
		if got := sides(cell); got != want {
			t.Errorf("%s medium borders = %q, want %q", cell, got, want)
		}
	}
}

func inputText(s string) input.Value { return input.Value{Kind: grid.Text, Text: s} }

// Frozen rows stay frozen across inserted and deleted rows, as in Excel.
func TestPanesMove(t *testing.T) {
	path := budget(t)
	freezeAt := func(cell string) {
		change(t, path, func(e *Editor) {
			at := rng(cell)
			must(t, e.View("Data", ViewOptions{Freeze: &at}))
		})
	}
	frozen := func() string {
		e, err := OpenEditor(path, false)
		must(t, err)
		defer e.Close()
		v, err := e.ViewOf("Data")
		must(t, err)
		return v.Freeze
	}
	freezeAt("B2")
	for _, c := range []struct {
		op   func(e *Editor) error
		want string
	}{
		{func(e *Editor) error { return e.InsertRows("Data", 1, 2) }, "B4"}, // above the split: it moves
		{func(e *Editor) error { return e.InsertRows("Data", 4, 1) }, "B4"}, // right below it: it stays
		{func(e *Editor) error { return e.DeleteRows("Data", 2, 3) }, "B2"}, // two of three frozen rows
		{func(e *Editor) error { return e.InsertCols("Data", 1, 1) }, "C2"}, // a column before the frozen one
		{func(e *Editor) error { return e.DeleteCols("Data", 4, 4) }, "C2"}, // after the split
		{func(e *Editor) error { return e.DeleteCols("Data", 1, 2) }, "A2"}, // all frozen columns
	} {
		change(t, path, func(e *Editor) { must(t, c.op(e)) })
		if got := frozen(); got != c.want {
			t.Errorf("frozen at %s, want %s", got, c.want)
		}
	}
}

// A recorded rule over the banding's own range does not make fmt keep the
// old banding beside the new.
func TestCondOnBandingRange(t *testing.T) {
	path := budget(t)
	format(t, path)
	change(t, path, func(e *Editor) {
		_, err := e.Cond("Data", rng("A2:D4"), []string{`=$D2<0`, "fill=lightyellow"})
		must(t, err)
	})
	format(t, path)
	format(t, path)
	x := sheetXML(t, path, dataPart)
	if n := strings.Count(x, "MOD(ROW()-2,2)=1"); n != 1 {
		t.Errorf("%d banding rules after fmt:\n%s", n, x)
	}
	if n := strings.Count(x, "$D2&lt;0"); n != 1 {
		t.Errorf("%d highlight rules after fmt:\n%s", n, x)
	}
	change(t, path, func(e *Editor) {
		cfs, err := e.CondFormats("Data")
		must(t, err)
		if len(cfs) != 2 || cfs[0].Recorded == cfs[0].Banding || cfs[1].Recorded == cfs[1].Banding {
			t.Errorf("CondFormats = %+v", cfs)
		}
	})
}

// A number format that would show ##### widens its column, once.
func TestNumFmtWidens(t *testing.T) {
	path := writeBook(t, [][]string{{"Item", "Budget"}, {"Rent", "1200000"}, {"IT", "150"}})
	var first, second []string
	change(t, path, func(e *Editor) {
		res, err := e.Style("Data", rng("B:B"), look(t, "format=currency"))
		must(t, err)
		first = res.Notes
	})
	change(t, path, func(e *Editor) {
		res, err := e.Style("Data", rng("B:B"), look(t, "format=currency"))
		must(t, err)
		second = res.Notes
	})
	if len(first) != 1 || !strings.Contains(first[0], "widened column B") || len(second) != 0 {
		t.Errorf("notes: %q then %q", first, second)
	}
}

package xlsx

import (
	"bytes"
	"image"
	"image/png"
	"os"
	"path/filepath"
	"reflect"
	"sort"
	"testing"

	"github.com/eemax/sheets/internal/grid"
	"github.com/xuri/excelize/v2"
)

func featuresOf(t *testing.T, name string) []Feature {
	t.Helper()
	data, err := os.ReadFile(filepath.Join("testdata", name))
	if err != nil {
		t.Fatal(err)
	}
	fs, err := Features(data)
	if err != nil {
		t.Fatal(err)
	}
	return fs
}

func TestFeatures(t *testing.T) {
	// XlsxWriter's file, as Excel saves it: a chart grouped with a text
	// box, a list validation in the 2010 extension list.
	want := []Feature{
		{Kind: FeatureValidation, Sheet: "Data", Where: "B2:B6", What: "whole validation", Type: "whole", Rule: "between|0|1000"},
		{Kind: FeatureValidation, Sheet: "Data", Where: "C2:C6", What: "decimal validation", Type: "decimal", Rule: "greaterThanOrEqual|0"},
		{Kind: FeatureValidation, Sheet: "Data", Where: "A2:A6", What: "list validation", Type: "list", Rule: "Summary!$D$2:$D$6"},
		{Kind: FeatureCond, Sheet: "Data", Where: "D2:D6", What: "cell value rule", Type: "cellIs", Rule: "greaterThan|140"},
		{Kind: FeatureChart, Sheet: "Data", Where: "F2", What: "column chart", Type: "column", Series: []string{"Data!$D$2:$D$6"}},
		{Kind: FeatureShape, Sheet: "Data", Where: "F2", What: "TextBox 2"},
		{Kind: FeatureName, Where: "TaxRate", What: "Summary!$B$1"},
	}
	if got := featuresOf(t, "against_before.xlsx"); !reflect.DeepEqual(got, want) {
		t.Errorf("against_before.xlsx:\n got %+v\nwant %+v", got, want)
	}
	// openpyxl's round trip keeps neither the group nor the extension list.
	want = []Feature{want[0], want[1], want[3], want[6]}
	if got := featuresOf(t, "against_openpyxl.xlsx"); !reflect.DeepEqual(got, want) {
		t.Errorf("against_openpyxl.xlsx:\n got %+v\nwant %+v", got, want)
	}
	// A data bar with its Excel 2010 part is one rule; a rule only in the
	// extension list counts.
	var conds []string
	for _, f := range featuresOf(t, "rules.xlsx") {
		if f.Kind == FeatureCond {
			conds = append(conds, f.Sheet+"!"+f.Where+" "+f.What)
		}
	}
	sort.Strings(conds)
	wantConds := []string{"Data!A3:D7 formula rule", "Data!B3:B7 data bar", "Data!C3:C7 cell value rule", "Data!C3:C7 icon set", "Other!A1:A3 formula rule"}
	if !reflect.DeepEqual(conds, wantConds) {
		t.Errorf("rules.xlsx: %v", conds)
	}
	// openpyxl's pictures-free chart and its Excel table.
	got := featuresOf(t, "rich.xlsx")
	kinds := map[string]int{}
	for _, f := range got {
		kinds[f.Kind]++
	}
	if kinds[FeatureChart] != 1 || kinds[FeatureTable] != 1 || kinds[FeatureValidation] != 1 {
		t.Errorf("rich.xlsx: %+v", got)
	}
	// What a rule says, and the columns of a table.
	for _, f := range got {
		switch f.Kind {
		case FeatureTable:
			if f.Rule != "k|v" {
				t.Errorf("rich.xlsx table columns: %q", f.Rule)
			}
		case FeatureValidation:
			if f.Type != "list" || f.Rule != `"North,South,East,West"` {
				t.Errorf("rich.xlsx validation: %+v", f)
			}
		}
	}
	for _, f := range featuresOf(t, "rules.xlsx") {
		if f.Kind == FeatureCond && f.Type == "cellIs" && f.Rule != "greaterThan|$C$3" || f.Kind == FeatureValidation && f.Type == "custom" && f.Rule != "C3>$B3" {
			t.Errorf("rules.xlsx: %+v", f)
		}
	}
	// An Excel table's columns, before and after a rename.
	for name, want := range map[string]Feature{
		"against_table.xlsx":         {Kind: FeatureTable, Sheet: "Data", Where: "A1:B3", What: "Sales", Rule: "Region|Units"},
		"against_table_renamed.xlsx": {Kind: FeatureTable, Sheet: "Data", Where: "A1:B3", What: "Sales2", Rule: "Region|Units"},
	} {
		if got := featuresOf(t, name); len(got) != 1 || !reflect.DeepEqual(got[0], want) {
			t.Errorf("%s: %+v", name, got)
		}
	}
	if _, err := Features([]byte("not a zip")); err == nil {
		t.Error("not a zip: no error")
	}
}

func TestTracer(t *testing.T) {
	wb, err := Read(filepath.Join("testdata", "against_before.xlsx"))
	if err != nil {
		t.Fatal(err)
	}
	tr := NewTracer(wb)
	got := tr.Spread([]Source{
		{Sheet: "Data", Range: grid.Range{R1: 2, C1: 2, R2: 2, C2: 2}},    // B2
		{Sheet: "Summary", Range: grid.Range{R1: 1, C1: 2, R2: 1, C2: 2}}, // the tax rate
	})
	want := map[At]int{
		{"Data", 2, 4}: 0, {"Data", 7, 2}: 0, {"Data", 7, 4}: 0,
		{"Summary", 2, 2}: 0, {"Summary", 3, 2}: 0, {"Summary", 6, 2}: 0, {"Summary", 7, 2}: 0,
		// The nearest cause: the tax reads the rate.
		{"Summary", 4, 2}: 1, {"Summary", 5, 2}: 1,
	}
	if !reflect.DeepEqual(got, want) {
		t.Errorf("from Data!B2:\n got %v\nwant %v", got, want)
	}
	// Through the name only: the tax, then the net.
	got = tr.Spread([]Source{{Sheet: "Summary", Range: grid.Range{R1: 1, C1: 2, R2: 1, C2: 2}}})
	want = map[At]int{{"Summary", 4, 2}: 0, {"Summary", 5, 2}: 0}
	if !reflect.DeepEqual(got, want) {
		t.Errorf("from the tax rate: %v", got)
	}
	if u := tr.Using("taxrate"); len(u) != 1 || u[0] != (At{"Summary", 4, 2}) {
		t.Errorf("using TaxRate: %v", u)
	}
	// A formula given as a source is reached itself.
	got = tr.Spread([]Source{{Sheet: "Data", Range: grid.Range{R1: 2, C1: 4, R2: 2, C2: 4}}})
	want = map[At]int{{"Data", 2, 4}: 0, {"Data", 7, 4}: 0, {"Summary", 3, 2}: 0, {"Summary", 4, 2}: 0, {"Summary", 5, 2}: 0}
	if !reflect.DeepEqual(got, want) {
		t.Errorf("from Data!D2: %v", got)
	}
	// Readers only: the totals in row 7 are not reached for being there,
	// and nothing reads them; D2 is, for reading B2 and C2.
	got = tr.Spread([]Source{{Sheet: "Data", Range: grid.Range{R1: 7, C1: 1, R2: 7, C2: 4}, Readers: true}})
	if len(got) != 0 {
		t.Errorf("readers of Data!7: %v", got)
	}
	got = tr.Spread([]Source{{Sheet: "Data", Range: grid.Range{R1: 2, C1: 1, R2: 2, C2: 4}, Readers: true}})
	if _, ok := got[At{"Data", 2, 4}]; !ok || len(got) != 9 {
		t.Errorf("readers of Data!2: %v", got)
	}
	if tr.Opaque("Summary", 4, 2) != "" || !tr.Formula("Data", 2, 4) || tr.Formula("Data", 2, 2) {
		t.Error("formulas")
	}
}

// SpreadAll tells of every source reaching a cell, up to a limit, the lowest
// first.
func TestSpreadAll(t *testing.T) {
	wb, err := Read(filepath.Join("testdata", "against_before.xlsx"))
	if err != nil {
		t.Fatal(err)
	}
	tr := NewTracer(wb)
	src := func(sheet string, row, col int) Source {
		return Source{Sheet: sheet, Range: grid.Range{R1: row, C1: col, R2: row, C2: col}}
	}
	// B2 and B3 of Data, and the tax rate.
	sources := []Source{src("Data", 2, 2), src("Data", 3, 2), src("Summary", 1, 2)}
	got := tr.SpreadAll(sources, 2)
	want := map[At]Reach{
		{"Data", 2, 4}:    {Sources: []int{0}},
		{"Data", 3, 4}:    {Sources: []int{1}},
		{"Data", 7, 2}:    {Sources: []int{0, 1}},
		{"Summary", 2, 2}: {Sources: []int{0, 1}},
		{"Summary", 6, 2}: {Sources: []int{0, 1}},             // the most units: both reach it
		{"Summary", 4, 2}: {Sources: []int{0, 1}, More: true}, // the tax: the rate too
		{"Summary", 5, 2}: {Sources: []int{0, 1}, More: true},
		{"Data", 7, 4}:    {Sources: []int{0, 1}},
		{"Summary", 3, 2}: {Sources: []int{0, 1}},
		{"Summary", 7, 2}: {Sources: []int{0}}, // North units: B2 alone
	}
	if !reflect.DeepEqual(got, want) {
		t.Errorf("SpreadAll:\n got %v\nwant %v", got, want)
	}
	// The same again: it keeps the index it reads.
	if again := tr.SpreadAll(sources, 2); !reflect.DeepEqual(again, want) {
		t.Errorf("SpreadAll again: %v", again)
	}
	// With room for all, nobody is left out.
	if all := tr.SpreadAll(sources, 3); all[At{"Summary", 4, 2}].More || len(all[At{"Summary", 4, 2}].Sources) != 3 {
		t.Errorf("room for all: %v", all[At{"Summary", 4, 2}])
	}
}

// A chart on a chart sheet (a sheet that is one chart, with no anchor) is a
// feature of that sheet; deleting the sheet takes it away.
func TestFeaturesChartsheet(t *testing.T) {
	want := []Feature{{Kind: FeatureChart, Sheet: "Chart1", What: "bar chart", Chartsheet: true, Type: "bar", Series: []string{"Data!$A$1:$A$3"}}}
	if got := featuresOf(t, "against_chartsheet.xlsx"); !reflect.DeepEqual(got, want) {
		t.Errorf("against_chartsheet.xlsx:\n got %+v\nwant %+v", got, want)
	}
	if got := featuresOf(t, "against_chartsheet_deleted.xlsx"); len(got) != 0 {
		t.Errorf("against_chartsheet_deleted.xlsx: %+v", got)
	}
}

// Touching finds what reads the lines (the ones a later version removed),
// whether a reference spans them, ends on one, or is nothing else.
func TestTracerTouching(t *testing.T) {
	wb, err := Read(filepath.Join("testdata", "against_before.xlsx"))
	if err != nil {
		t.Fatal(err)
	}
	tr := NewTracer(wb)
	// Data!6, the last row of B2:B6 and D2:D6, and the row of D6 itself.
	got := tr.Touching("Data", [][2]int{{6, 6}}, false)
	want := map[At]int{
		{"Data", 6, 4}: 0, {"Data", 7, 2}: 0, {"Data", 7, 4}: 0,
		{"Summary", 2, 2}: 0, {"Summary", 3, 2}: 0, {"Summary", 6, 2}: 0,
	}
	if !reflect.DeepEqual(got, want) {
		t.Errorf("reading Data!6:\n got %v\nwant %v", got, want)
	}
	// The first row, B2 itself: its own row, the ranges it starts, and the
	// single cell Summary!B7 reads.
	got = tr.Touching("Data", [][2]int{{2, 2}}, false)
	if _, ok := got[At{"Summary", 7, 2}]; !ok || got[At{"Summary", 2, 2}] != 0 || len(got) != 7 {
		t.Errorf("reading Data!2: %v", got)
	}
	// Runs: each formula maps to the first it reads in.
	got = tr.Touching("Data", [][2]int{{1, 1}, {3, 3}, {6, 6}}, false)
	if got[At{"Summary", 2, 2}] != 1 || got[At{"Data", 6, 4}] != 2 {
		t.Errorf("reading Data!1, 3 and 6: %v", got)
	}
	// Columns: Revenue (D) is read by its total and the summary's.
	got = tr.Touching("data", [][2]int{{4, 4}}, true)
	if _, ok := got[At{"Summary", 3, 2}]; !ok || len(got) != 2 {
		t.Errorf("reading column D: %v", got)
	}
	if got := tr.Touching("Data", nil, false); len(got) != 0 {
		t.Errorf("no lines: %v", got)
	}
}

func TestTracerOpaqueAndArrays(t *testing.T) {
	s := grid.NewSheet("S")
	s.Set(1, 1, grid.Cell{Kind: grid.Number, Num: 1})
	s.Set(1, 2, grid.Cell{Formula: "A1+NOW()"})
	s.Set(1, 3, grid.Cell{Formula: "'[other.xlsx]Data'!A1"})
	s.Set(1, 4, grid.Cell{Formula: "SEQUENCE(A1+1)", Kind: grid.Number, Num: 1})
	s.Set(2, 4, grid.Cell{Kind: grid.Number, Num: 2})
	s.Set(1, 5, grid.Cell{Formula: "D2*10"})
	s.Spills = []grid.Range{{R1: 1, C1: 4, R2: 2, C2: 4}}
	wb := &grid.Workbook{Sheets: []*grid.Sheet{s}}
	tr := NewTracer(wb)
	if w := tr.Opaque("S", 1, 2); w == "" || !contains(w, "NOW") {
		t.Errorf("volatile: %q", w)
	}
	if w := tr.Opaque("S", 1, 3); w != linkNote {
		t.Errorf("link: %q", w)
	}
	if a, ok := tr.Anchor("S", 2, 4); !ok || a != (At{"S", 1, 4}) {
		t.Errorf("anchor of D2: %v %v", a, ok)
	}
	got := tr.Spread([]Source{{Sheet: "S", Range: grid.Range{R1: 1, C1: 1, R2: 1, C2: 1}}})
	want := map[At]int{{"S", 1, 4}: 0, {"S", 2, 4}: 0, {"S", 1, 5}: 0}
	if !reflect.DeepEqual(got, want) {
		t.Errorf("spread through a spill: %v", got)
	}
}

// Functions whose result comes from outside the workbook (a server, the
// web, a data service) are opaque, but what they read is still followed: a
// change there is their cause.
func TestTracerOutside(t *testing.T) {
	s := grid.NewSheet("S")
	s.Set(1, 1, grid.Cell{Kind: grid.Text, Text: "https://example.com/rate"})
	s.Set(1, 2, grid.Cell{Formula: `RTD("my.server",,"EURUSD")`})
	s.Set(1, 3, grid.Cell{Formula: `_xlfn.WEBSERVICE(A1)`})
	s.Set(1, 4, grid.Cell{Formula: `_xlfn.STOCKHISTORY("MSFT",A1)`})
	s.Set(1, 5, grid.Cell{Formula: `CUBEVALUE("Sales",A1)`})
	s.Set(1, 6, grid.Cell{Formula: `A1&"x"`})
	tr := NewTracer(&grid.Workbook{Sheets: []*grid.Sheet{s}})
	for col, fn := range map[int]string{2: "RTD", 3: "WEBSERVICE", 4: "STOCKHISTORY", 5: "CUBEVALUE"} {
		if w := tr.Opaque("S", 1, col); !contains(w, "it calls "+fn+", whose result comes from outside the workbook") {
			t.Errorf("%s: %q", fn, w)
		}
	}
	if w := tr.Opaque("S", 1, 6); w != "" {
		t.Errorf("a formula of the workbook's own: %q", w)
	}
	got := tr.Spread([]Source{{Sheet: "S", Range: grid.Range{R1: 1, C1: 1, R2: 1, C2: 1}}})
	for _, col := range []int{3, 4, 5, 6} {
		if _, ok := got[At{"S", 1, col}]; !ok {
			t.Errorf("what reads A1 does not follow it: column %d in %v", col, got)
		}
	}
}

func contains(s, sub string) bool {
	for i := 0; i+len(sub) <= len(s); i++ {
		if s[i:i+len(sub)] == sub {
			return true
		}
	}
	return false
}

// A picture, a shape and a chart, each in an anchor of its own, as excelize
// writes them.
func TestFeaturesDrawn(t *testing.T) {
	f := excelize.NewFile()
	defer f.Close()
	var img bytes.Buffer
	if err := png.Encode(&img, image.NewRGBA(image.Rect(0, 0, 2, 2))); err != nil {
		t.Fatal(err)
	}
	if err := f.AddPictureFromBytes("Sheet1", "B3", &excelize.Picture{Extension: ".png", File: img.Bytes(), Format: &excelize.GraphicOptions{AltText: "logo"}}); err != nil {
		t.Fatal(err)
	}
	if err := f.AddShape("Sheet1", &excelize.Shape{Cell: "E5", Type: "rect", Width: 80, Height: 40}); err != nil {
		t.Fatal(err)
	}
	f.SetSheetRow("Sheet1", "A1", &[]any{1, 2, 3})
	if err := f.AddChart("Sheet1", "H2", &excelize.Chart{Type: excelize.Line, Series: []excelize.ChartSeries{{Values: "Sheet1!$A$1:$C$1"}}}); err != nil {
		t.Fatal(err)
	}
	buf, err := f.WriteToBuffer()
	if err != nil {
		t.Fatal(err)
	}
	fs, err := Features(buf.Bytes())
	if err != nil {
		t.Fatal(err)
	}
	got := map[string]string{}
	for _, x := range fs {
		got[x.Kind] = x.Sheet + "!" + x.Where
	}
	want := map[string]string{FeaturePicture: "Sheet1!B3", FeatureShape: "Sheet1!E5", FeatureChart: "Sheet1!H2"}
	if !reflect.DeepEqual(got, want) {
		t.Errorf("drawn: %+v", fs)
	}
}

// A workbook with more rules than the bound leaves them out and counts them;
// its names, tables and the rest are listed as ever.
func TestFeaturesWithin(t *testing.T) {
	f := excelize.NewFile()
	defer f.Close()
	for r := 1; r <= 30; r++ {
		cell := "B" + string(rune('0'+r/10)) + string(rune('0'+r%10))
		dv := excelize.NewDataValidation(true)
		dv.SetSqref(cell)
		dv.SetRange(0, 9, excelize.DataValidationTypeWhole, excelize.DataValidationOperatorBetween)
		if err := f.AddDataValidation("Sheet1", dv); err != nil {
			t.Fatal(err)
		}
	}
	if err := f.SetDefinedName(&excelize.DefinedName{Name: "Rate", RefersTo: "Sheet1!$A$1"}); err != nil {
		t.Fatal(err)
	}
	buf, err := f.WriteToBuffer()
	if err != nil {
		t.Fatal(err)
	}
	kinds := func(fs []Feature) map[string]int {
		m := map[string]int{}
		for _, x := range fs {
			m[x.Kind]++
		}
		return m
	}
	fs, over, err := FeaturesWithin(buf.Bytes(), 30)
	if err != nil || over != 0 || !reflect.DeepEqual(kinds(fs), map[string]int{FeatureValidation: 30, FeatureName: 1}) {
		t.Errorf("at the bound: %v, over %d, %v", kinds(fs), over, err)
	}
	fs, over, err = FeaturesWithin(buf.Bytes(), 29)
	if err != nil || over != 30 || !reflect.DeepEqual(kinds(fs), map[string]int{FeatureName: 1}) {
		t.Errorf("past the bound: %v, over %d, %v", kinds(fs), over, err)
	}
	all, err := Features(buf.Bytes())
	if err != nil || len(all) != 31 {
		t.Errorf("no bound: %d features, %v", len(all), err)
	}
}

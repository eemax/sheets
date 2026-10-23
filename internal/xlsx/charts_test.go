package xlsx

import (
	"strings"
	"testing"

	"github.com/eemax/sheets/internal/grid"
)

func TestReadCharts(t *testing.T) {
	wb, err := Read("testdata/rich.xlsx")
	if err != nil {
		t.Fatal(err)
	}
	cs := wb.Sheet("Data").Charts
	if len(cs) != 1 {
		t.Fatalf("charts = %+v", cs)
	}
	c := cs[0]
	if c.Anchor != "E2" || c.Type != "column" || len(c.Series) != 1 || c.Series[0].Values != "'Data'!$B$2:$B$4" || c.Series[0].Name != "'Data'!B1" {
		t.Errorf("chart = %+v", c)
	}
}

func addChart(t *testing.T, path string, spec ChartSpec) {
	t.Helper()
	e, err := OpenEditor(path, false)
	if err != nil {
		t.Fatal(err)
	}
	if err := e.AddChart(spec); err != nil {
		t.Fatal(err)
	}
	if err := e.Save(false); err != nil {
		t.Fatal(err)
	}
}

func TestAddChart(t *testing.T) {
	path := fixture(t, "rich.xlsx") // has a chart from openpyxl already
	spec := ChartSpec{
		Sheet: "Data", Anchor: "H2", Type: "column", Title: "Units", Width: 640, Height: 360, Points: 3,
		Categories: "Data!$A$2:$A$4", Series: []SeriesSpec{{Name: "Data!$B$1", Values: "Data!$B$2:$B$4"}},
		ValueFormat: "#,##0",
	}
	addChart(t, path, spec)
	xml := sheetXML(t, path, "xl/charts/chart2.xml")
	for _, want := range []string{`val="2A78D6"`, `<a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF">`, `formatCode="#,##0"`} {
		if !strings.Contains(xml, want) {
			t.Errorf("new chart lacks %s", want)
		}
	}
	if strings.Contains(xml, "<legend>") {
		t.Error("a single series got a legend")
	}
	if old := sheetXML(t, path, "xl/charts/chart1.xml"); strings.Contains(old, `w="19050"`) {
		t.Error("existing chart restyled")
	}
	wb, _ := Read(path)
	if cs := wb.Sheet("Data").Charts; len(cs) != 2 || cs[1].Anchor != "H2" || cs[1].Title != "Units" || cs[1].Type != "column" {
		t.Errorf("charts = %+v", cs)
	}

	// Ranges follow inserted rows, appends, and renames.
	e, _ := OpenEditor(path, false)
	if err := e.InsertRows("Data", 3, 2); err != nil {
		t.Fatal(err)
	}
	e.ExtendTable("Data", grid.Range{R1: 4, C1: 1, R2: 6, C2: 2}, 1)
	if err := e.RenameSheet("Data", "Regions"); err != nil {
		t.Fatal(err)
	}
	if err := e.Save(false); err != nil {
		t.Fatal(err)
	}
	wb, _ = Read(path)
	for _, c := range wb.Sheet("Regions").Charts {
		if v := c.Series[0].Values; !strings.HasSuffix(v, "!$B$2:$B$7") || !strings.Contains(v, "Regions") {
			t.Errorf("series values = %s", v)
		}
	}
	if d := sheetXML(t, path, "xl/drawings/drawing1.xml"); !strings.Contains(d, `<xdr:wsDr xmlns:c=`) {
		t.Error("drawing lost its chart namespace")
	}
}

func TestChartRefusals(t *testing.T) {
	path := fixture(t, "rich.xlsx")
	e, _ := OpenEditor(path, false)
	base := ChartSpec{Sheet: "Data", Anchor: "H2", Width: 640, Height: 360, Points: 3, Categories: "Data!$A$2:$A$4"}
	many := make([]SeriesSpec, 9)
	cases := map[string]ChartSpec{
		"too many series":   func() ChartSpec { s := base; s.Type = "column"; s.Series = many; return s }(),
		"pie of two series": func() ChartSpec { s := base; s.Type = "pie"; s.Series = many[:2]; return s }(),
		"big pie":           func() ChartSpec { s := base; s.Type = "pie"; s.Series = many[:1]; s.Points = 7; return s }(),
		"stacked line":      func() ChartSpec { s := base; s.Type = "line"; s.Stacked = true; s.Series = many[:2]; return s }(),
		"crowded scatter":   func() ChartSpec { s := base; s.Type = "scatter"; s.Series = many[:4]; return s }(),
		"overlapping areas": func() ChartSpec { s := base; s.Type = "area"; s.Series = many[:2]; return s }(),
	}
	for name, spec := range cases {
		if err := e.AddChart(spec); err == nil {
			t.Errorf("%s: accepted", name)
		}
	}
}

// Series and slice colours come from fills; the white outlines sheets adds
// between marks, and marker colours, must not be taken for them.
func TestChartColors(t *testing.T) {
	path := fixture(t, "rich.xlsx")
	addChart(t, path, ChartSpec{Sheet: "Data", Anchor: "H2", Type: "pie", Title: "Units", Width: 640, Height: 360, Points: 3,
		Categories: "Data!$A$2:$A$4", Series: []SeriesSpec{{Name: "Data!$B$1", Values: "Data!$B$2:$B$4"}}})
	addChart(t, path, ChartSpec{Sheet: "Data", Anchor: "H22", Type: "line", Title: "Units", Width: 640, Height: 360, Points: 3,
		Categories: "Data!$A$2:$A$4", Series: []SeriesSpec{{Name: "Data!$B$1", Values: "Data!$B$2:$B$4"}}})
	wb, _ := Read(path)
	cs := wb.Sheet("Data").Charts
	pie, line := cs[1], cs[2]
	if got := strings.Join(pie.Series[0].PointColors, ","); got != "2A78D6,EB6834,1BAF7A" {
		t.Errorf("pie colours = %s", got)
	}
	if line.Series[0].Color != "2A78D6" {
		t.Errorf("line colour = %s", line.Series[0].Color)
	}
}

// sheets draws a single area as the area plus a line along its edge; it
// reads back as one area series, the edge kept only for drawing.
func TestAreaReadsAsArea(t *testing.T) {
	path := fixture(t, "rich.xlsx")
	addChart(t, path, ChartSpec{
		Sheet: "Data", Anchor: "H2", Type: "area", Title: "Units", Width: 640, Height: 360, Points: 3,
		Categories: "Data!$A$2:$A$4", Series: []SeriesSpec{{Name: "Data!$B$1", Values: "Data!$B$2:$B$4"}},
	})
	wb, _ := Read(path)
	c := wb.Sheet("Data").Charts[1]
	if c.Type != "area" || len(c.Series) != 2 || len(c.Shown().Series) != 1 || c.Shown().Series[0].Kind != "area" {
		t.Errorf("chart = %+v", c)
	}
}

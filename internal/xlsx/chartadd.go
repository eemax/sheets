package xlsx

import (
	"fmt"
	"strconv"
	"strings"

	"github.com/eemax/sheets/internal/grid"
	"github.com/xuri/excelize/v2"
)

// ChartSpec describes a chart to add. References are sheet-qualified
// absolute ranges ('Sales'!$B$2:$B$9).
type ChartSpec struct {
	Sheet          string // sheet the chart is placed on
	Anchor         string // its top-left cell
	Type           string // column, bar, line, area, pie, scatter
	Stacked        bool
	Title          string
	Width, Height  uint // pixels
	Categories     string
	Series         []SeriesSpec
	Points         int    // number of categories
	ValueFormat    string // number format of the value axis
	CategoryFormat string // number format of the category axis (dates, scatter x)
}

// SeriesSpec is one series: its name (a header cell) and values.
type SeriesSpec struct {
	Name, Values string
}

// The house chart palette: a categorical order validated for colour-vision
// deficiency and contrast (adjacent pairs), assigned to series in order and
// never cycled. Scatter plots, where every pair can sit side by side, use
// at most the first three.
var Palette = []string{"2A78D6", "EB6834", "1BAF7A", "EDA100", "E87BA4", "008300", "6250D6", "E34948"}

// Ink for chart text, never a series colour.
const (
	inkPrimary   = "0B0B0B"
	inkSecondary = "52514E"
	inkMuted     = "898781"
	surface      = "FFFFFF"
)

// Limits on series and slices.
const (
	MaxSeries        = 8
	MaxScatterSeries = 3
	MaxPieSlices     = 6
)

// AddChart adds a chart in the house style.
func (e *Editor) AddChart(spec ChartSpec) error {
	n := len(spec.Series)
	switch {
	case n == 0:
		return fmt.Errorf("a chart needs at least one series")
	case n > MaxSeries:
		return fmt.Errorf("%d series is too many to tell apart (at most %d); sum the rest into an \"Other\" column first (sheets query can), or make several charts", n, MaxSeries)
	case spec.Type == "scatter" && n > MaxScatterSeries:
		return fmt.Errorf("a scatter plot keeps at most %d series distinguishable; make several charts", MaxScatterSeries)
	case spec.Type == "pie" && n > 1:
		return fmt.Errorf("a pie shows one series; use a stacked column to compare several")
	case spec.Type == "pie" && spec.Points > MaxPieSlices:
		return fmt.Errorf("%d slices are too many for a pie (at most %d); use a bar chart", spec.Points, MaxPieSlices)
	case spec.Type == "area" && !spec.Stacked && n > 1:
		return fmt.Errorf("overlapping areas hide each other; use a line chart for several series, or --stacked for parts of a whole")
	case spec.Type == "line" && spec.Stacked:
		return fmt.Errorf("stacked lines read as separate trends but are not; use a stacked area")
	case spec.Type == "scatter" && spec.Stacked:
		return fmt.Errorf("a scatter plot cannot be stacked")
	}
	types := map[string][2]excelize.ChartType{
		"column":  {excelize.Col, excelize.ColStacked},
		"bar":     {excelize.Bar, excelize.BarStacked},
		"line":    {excelize.Line, excelize.Line},
		"area":    {excelize.Area, excelize.AreaStacked},
		"pie":     {excelize.Pie, excelize.Pie},
		"scatter": {excelize.Scatter, excelize.Scatter},
	}
	pair, ok := types[spec.Type]
	if !ok {
		return fmt.Errorf("unknown chart type %q (column, bar, line, area, pie, scatter)", spec.Type)
	}
	typ := pair[0]
	if spec.Stacked {
		typ = pair[1]
	}
	if e.wb.Sheet(spec.Sheet) == nil {
		return fmt.Errorf("sheet %q not found", spec.Sheet)
	}

	font := &excelize.Font{Family: "Calibri", Size: 13, Color: inkPrimary}
	ch := &excelize.Chart{
		Type:         typ,
		Title:        excelize.ChartTitle{Paragraph: []excelize.RichTextRun{{Text: spec.Title, Font: font}}},
		Dimension:    excelize.ChartDimension{Width: spec.Width, Height: spec.Height},
		Legend:       excelize.ChartLegend{Position: "none"},
		Border:       excelize.LineOptions{Type: excelize.LineNone},
		ShowBlanksAs: "gap",
		Format:       excelize.GraphicOptions{Name: "sheets chart"},
		XAxis:        excelize.ChartAxis{Font: excelize.Font{Family: "Calibri", Size: 9, Color: inkSecondary}},
		YAxis: excelize.ChartAxis{
			MajorGridLines: true,
			Font:           excelize.Font{Family: "Calibri", Size: 9, Color: inkMuted},
			NumFmt:         excelize.ChartNumFmt{CustomNumFmt: spec.ValueFormat},
		},
	}
	if spec.CategoryFormat != "" {
		ch.XAxis.NumFmt = excelize.ChartNumFmt{CustomNumFmt: spec.CategoryFormat}
	}
	no := false
	ch.VaryColors = &no
	if n > 1 || spec.Type == "pie" {
		ch.Legend = excelize.ChartLegend{Position: "top", Font: &excelize.Font{Family: "Calibri", Size: 9, Color: inkSecondary}}
	}
	if spec.Type == "bar" {
		ch.XAxis.ReverseOrder = true // first category at the top, as in the table
	}
	if spec.Type == "column" || spec.Type == "bar" {
		gap := barGap(spec)
		ch.GapWidth = &gap
	}
	if spec.Type == "pie" {
		yes := true
		ch.VaryColors = &yes
		ch.PlotArea = excelize.ChartPlotArea{ShowPercent: true}
		ch.Legend.Position = "right"
	}
	markers := spec.Points <= 12
	for i, s := range spec.Series {
		color := Palette[i]
		cs := excelize.ChartSeries{Name: s.Name, Categories: spec.Categories, Values: s.Values}
		switch spec.Type {
		case "line":
			cs.Line = excelize.LineOptions{Type: excelize.LineSolid, Width: 1.5, Fill: solid(color)}
			cs.Fill = solid(color)
			cs.Marker = excelize.ChartMarker{Symbol: "none"}
			if markers {
				cs.Marker = excelize.ChartMarker{Symbol: "circle", Size: 6, Fill: solid(color), Border: excelize.LineOptions{Type: excelize.LineSolid, Width: 1.5, Fill: solid(surface)}}
			}
		case "scatter":
			cs.Marker = excelize.ChartMarker{Symbol: "circle", Size: 7, Fill: solid(color), Border: excelize.LineOptions{Type: excelize.LineSolid, Width: 1.5, Fill: solid(surface)}}
			cs.Line = excelize.LineOptions{Type: excelize.LineNone}
		case "area":
			cs.Fill = solid(color)
			cs.Fill.Transparency = 85 // a wash, with the line drawn over it
			if spec.Stacked {
				cs.Fill.Transparency = 30
			}
		case "pie":
			for p := 0; p < spec.Points && p < len(Palette); p++ {
				cs.DataPoint = append(cs.DataPoint, excelize.ChartDataPoint{Index: p, Fill: solid(Palette[p])})
			}
		default:
			cs.Fill = solid(color)
		}
		ch.Series = append(ch.Series, cs)
	}
	var combo []*excelize.Chart
	if spec.Type == "area" && !spec.Stacked {
		// A single area: the wash, and its line on top.
		s := spec.Series[0]
		combo = append(combo, &excelize.Chart{Type: excelize.Line, XAxis: ch.XAxis, YAxis: ch.YAxis, Series: []excelize.ChartSeries{{
			Name: s.Name, Categories: spec.Categories, Values: s.Values,
			Line:   excelize.LineOptions{Type: excelize.LineSolid, Width: 1.5, Fill: solid(Palette[0])},
			Fill:   solid(Palette[0]),
			Marker: excelize.ChartMarker{Symbol: "none"},
		}}})
	}

	before := chartParts(e.f)
	next := fmt.Sprintf("xl/charts/chart%d.xml", len(before)+1)
	if before[next] {
		// excelize names new charts by count and would overwrite this one.
		return fmt.Errorf("cannot add a chart: the workbook's chart parts are numbered with gaps (%s exists); open and save it in Excel first", next)
	}
	if err := e.f.AddChart(spec.Sheet, spec.Anchor, ch, combo...); err != nil {
		return err
	}
	for part := range chartParts(e.f) {
		if !before[part] {
			// Bar charts list categories top-down (reversed), which would put
			// the value axis at the top unless it crosses at the far end.
			e.patches.newCharts[part] = chartFix{valueAxisAtEnd: spec.Type == "bar"}
		}
	}
	return nil
}

func solid(color string) excelize.Fill {
	return excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{color}}
}

// barGap sizes the gap between categories so bars stay thin (about 24px)
// instead of filling the plot.
func barGap(spec ChartSpec) uint {
	plot := float64(spec.Width) - 90
	if spec.Type == "bar" {
		plot = float64(spec.Height) - 70
	}
	bars := float64(len(spec.Series))
	if spec.Stacked {
		bars = 1
	}
	perCategory := plot / float64(max(spec.Points, 1))
	gap := (perCategory/24 - bars) * 100
	return uint(min(max(gap, 50), 500))
}

// chartParts lists the chart parts in the package.
func chartParts(f *excelize.File) map[string]bool {
	out := map[string]bool{}
	f.Pkg.Range(func(k, _ any) bool {
		if name, ok := k.(string); ok && strings.HasPrefix(name, "xl/charts/chart") && strings.HasSuffix(name, ".xml") {
			out[name] = true
		}
		return true
	})
	return out
}

// DeleteChart removes the chart anchored at a cell.
func (e *Editor) DeleteChart(sheet, cell string) error {
	s := e.wb.Sheet(sheet)
	if s == nil {
		return fmt.Errorf("sheet %q not found", sheet)
	}
	var anchors []string
	for _, c := range s.Charts {
		anchors = append(anchors, c.Anchor)
		if strings.EqualFold(c.Anchor, cell) {
			return e.f.DeleteChart(s.Name, c.Anchor)
		}
	}
	if len(anchors) == 0 {
		return fmt.Errorf("%s has no charts", grid.QuoteSheet(s.Name))
	}
	return fmt.Errorf("no chart anchored at %s (charts on %s are at %s)", cell, grid.QuoteSheet(s.Name), strings.Join(anchors, ", "))
}

// ChartRef renders an absolute, sheet-qualified range for chart series.
func ChartRef(sheet string, r grid.Range) string {
	abs := func(row, col int) string { return "$" + grid.ColName(col) + "$" + strconv.Itoa(row) }
	ref := grid.QuoteSheet(sheet) + "!" + abs(r.R1, r.C1)
	if r.R1 != r.R2 || r.C1 != r.C2 {
		ref += ":" + abs(r.R2, r.C2)
	}
	return ref
}

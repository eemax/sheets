package render

import (
	"bytes"
	"image"
	"image/png"
	"path/filepath"
	"strings"
	"testing"

	"github.com/eemax/sheets/internal/grid"
	"github.com/xuri/excelize/v2"
)

const rich = "../xlsx/testdata/rich.xlsx" // a table and a chart, made by openpyxl

func TestPNG(t *testing.T) {
	var buf bytes.Buffer
	res, err := PNG(rich, &buf, Options{Sheet: "Data", Scale: 1, Headers: true})
	if err != nil {
		t.Fatal(err)
	}
	if res.Charts != 1 || res.Range == "" {
		t.Errorf("result = %+v", res)
	}
	img, err := png.Decode(&buf)
	if err != nil {
		t.Fatal(err)
	}
	if b := img.Bounds(); b.Dx() != res.Width || b.Dy() != res.Height {
		t.Errorf("image %v, result %dx%d", b, res.Width, res.Height)
	}
	// The chart (anchored at E2) puts its default blue bars on the page.
	if !hasColor(img, func(r, g, b uint32) bool { return b > r+0x2000 && b > g }) {
		t.Error("no chart colour drawn")
	}
	// Text is drawn: some dark pixels in the first data cell.
	if !hasColorIn(img, image.Rect(headerW, headerH+20, headerW+64, headerH+40), func(r, g, b uint32) bool { return r < 0x6000 && g < 0x6000 && b < 0x6000 }) {
		t.Error("no text in A2")
	}
}

func TestRangeAndScale(t *testing.T) {
	var buf bytes.Buffer
	rg, _ := grid.ParseRange("A1:B2")
	res, err := PNG(rich, &buf, Options{Sheet: "Data", Range: &rg, Scale: 2})
	if err != nil {
		t.Fatal(err)
	}
	img, _ := png.Decode(&buf)
	if res.Range != "A1:B2" || res.Charts != 0 || img.Bounds().Dx() != 2*res.Width {
		t.Errorf("result = %+v, image %v", res, img.Bounds())
	}
}

func TestHTML(t *testing.T) {
	var buf bytes.Buffer
	if _, err := HTML(rich, &buf, Options{Sheet: "Data", Headers: true}); err != nil {
		t.Fatal(err)
	}
	out := buf.String()
	for _, want := range []string{"<table", "<td>North</td>", `<img class="chart"`, "data:image/png;base64,"} {
		if !strings.Contains(out, want) {
			t.Errorf("html lacks %s", want)
		}
	}
}

func TestAxisLabel(t *testing.T) {
	for _, tt := range []struct {
		v    float64
		code string
		want string
	}{
		{25000, `"$"#,##0,"K"`, "$25K"},
		{1500, "#,##0", "1,500"},
		{2e6, `#,##0,,"M"`, "2M"},
		{0.25, "0%", "25%"},
		{7, "General", "7"},
	} {
		if got := axisLabel(tt.v, tt.code); got != tt.want {
			t.Errorf("axisLabel(%v, %s) = %s, want %s", tt.v, tt.code, got, tt.want)
		}
	}
	if got := niceTicks(0, 37425, 5); got[0] != 0 || got[len(got)-1] < 37425 {
		t.Errorf("niceTicks = %v", got)
	}
}

func hasColor(img image.Image, ok func(r, g, b uint32) bool) bool {
	return hasColorIn(img, img.Bounds(), ok)
}

func hasColorIn(img image.Image, area image.Rectangle, ok func(r, g, b uint32) bool) bool {
	area = area.Intersect(img.Bounds())
	for y := area.Min.Y; y < area.Max.Y; y++ {
		for x := area.Min.X; x < area.Max.X; x++ {
			r, g, b, _ := img.At(x, y).RGBA()
			if ok(r, g, b) {
				return true
			}
		}
	}
	return false
}

// Hidden rows and columns take no room, rows without a height of their own
// fit large type, and a centre-across label spans the empty cells after it.
func TestGeometry(t *testing.T) {
	f := excelize.NewFile()
	f.SetCellValue("Sheet1", "A1", "Big title")
	big, _ := f.NewStyle(&excelize.Style{Font: &excelize.Font{Size: 22}})
	f.SetCellStyle("Sheet1", "A1", "A1", big)
	f.SetCellValue("Sheet1", "B2", "Q1")
	cc, _ := f.NewStyle(&excelize.Style{Alignment: &excelize.Alignment{Horizontal: "centerContinuous"}})
	f.SetCellStyle("Sheet1", "B2", "D2", cc)
	for r := 3; r <= 5; r++ {
		f.SetCellValue("Sheet1", "A"+itoa(r), r)
	}
	f.SetRowVisible("Sheet1", 4, false)
	f.SetRowHeight("Sheet1", 5, 30)
	f.SetColVisible("Sheet1", "E", false)
	f.SetCellValue("Sheet1", "F1", "x")
	if err := f.AutoFilter("Sheet1", "A3:B5", nil); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(t.TempDir(), "geo.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	sh, err := load(path, Options{})
	if err != nil {
		t.Fatal(err)
	}
	height := func(r int) int { return sh.rowY[r] - sh.rowY[r-1] }
	if h := height(1); h < 36 {
		t.Errorf("22pt row is %dpx high", h)
	}
	if h := height(3); h != 20 {
		t.Errorf("default row is %dpx high", h)
	}
	if h := height(4); h != 0 {
		t.Errorf("hidden row is %dpx high", h)
	}
	if h := height(5); h != 40 {
		t.Errorf("30pt row is %dpx high", h)
	}
	if w := sh.colX[5] - sh.colX[4]; w != 0 {
		t.Errorf("hidden column is %dpx wide", w)
	}
	if last := sh.continuous(2, 2); last != 4 {
		t.Errorf("centre-across label spans to column %d, want 4", last)
	}
	if len(sh.buttons) != 2 || !sh.buttons[[2]int{3, 1}] || !sh.buttons[[2]int{3, 2}] {
		t.Errorf("filter buttons at %v, want A3 and B3", sh.buttons)
	}
	if last := sh.runsOn(1, 1, "left"); last != 2 {
		t.Errorf("title runs on to column %d, want 2", last)
	}
}

// Banding is drawn when another rule shares its range (excelize's
// GetConditionalFormats keeps one rule per range).
func TestBandingSharingItsRange(t *testing.T) {
	f := excelize.NewFile()
	for r := 1; r <= 6; r++ {
		f.SetCellValue("Sheet1", grid.CellName(r, 1), "x")
		f.SetCellValue("Sheet1", grid.CellName(r, 2), r)
	}
	band, err := f.NewConditionalStyle(&excelize.Style{Fill: excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{"F2F2F2"}}})
	if err != nil {
		t.Fatal(err)
	}
	hi, err := f.NewConditionalStyle(&excelize.Style{Font: &excelize.Font{Bold: true}})
	if err != nil {
		t.Fatal(err)
	}
	// Two formats over one range, as fmt's banding and a cond make.
	for _, o := range []excelize.ConditionalFormatOptions{
		{Type: "formula", Criteria: "MOD(ROW()-2,2)=1", Format: &band},
		{Type: "formula", Criteria: "$B2>4", Format: &hi},
	} {
		if err := f.SetConditionalFormat("Sheet1", "A2:B6", []excelize.ConditionalFormatOptions{o}); err != nil {
			t.Fatal(err)
		}
	}
	path := filepath.Join(t.TempDir(), "band.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	sh, err := load(path, Options{Sheet: "Sheet1"})
	if err != nil {
		t.Fatal(err)
	}
	for r := 2; r <= 6; r++ {
		ce := sh.cell(r, 1)
		banded := ce.look != nil && ce.look.fill != nil
		if banded != (r%2 == 1) {
			t.Errorf("row %d banded = %v", r, banded)
		}
	}
	if !strings.Contains(strings.Join(sh.missing, ";"), "other than banding") {
		t.Errorf("the other rule is not reported: %v", sh.missing)
	}
}

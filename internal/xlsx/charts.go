package xlsx

import (
	"encoding/xml"
	"io"
	"path"
	"strconv"
	"strings"

	"github.com/eemax/sheets/internal/grid"
)

// charts reads the charts drawn on a worksheet: its drawing part says
// where each chart is anchored, and each chart part what it plots. Parts
// that cannot be read are skipped: charts are reported, not edited, here.
func (p *pkg) charts(drawingPart string) []grid.Chart {
	dir := path.Dir(drawingPart)
	rels, err := p.rels(path.Join(dir, "_rels", path.Base(drawingPart)+".rels"), dir)
	if err != nil {
		return nil
	}
	target := map[string]string{}
	for _, r := range rels {
		if strings.HasSuffix(r.Type, "/chart") {
			target[r.ID] = r.Target
		}
	}
	r, err := p.open(drawingPart)
	if err != nil {
		return nil
	}
	defer r.Close()
	dec := xml.NewDecoder(r)
	var out []grid.Chart
	var in, col, row, toCol, toRow string // in: element whose text we collect
	var cx, cy int
	inFrom, inTo := false, false
	for {
		tok, err := dec.RawToken()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "from":
				inFrom = true
			case "to":
				inTo = true
			case "col", "row":
				switch {
				case inFrom:
					in = t.Name.Local
				case inTo:
					in = "to" + t.Name.Local
				}
			case "ext":
				if v, err := strconv.Atoi(attr(t, "cx")); err == nil {
					cx = v
				}
				if v, err := strconv.Atoi(attr(t, "cy")); err == nil {
					cy = v
				}
			case "twoCellAnchor", "oneCellAnchor", "absoluteAnchor":
				col, row, toCol, toRow, cx, cy = "", "", "", "", 0, 0
			case "chart":
				id := attr(t, "id")
				part, ok := target[id]
				if !ok {
					continue
				}
				c := p.chart(part)
				ci, err1 := strconv.Atoi(col)
				ri, err2 := strconv.Atoi(row)
				if err1 == nil && err2 == nil {
					c.Anchor = grid.CellName(ri+1, ci+1)
				}
				if tc, err1 := strconv.Atoi(toCol); err1 == nil {
					if tr, err2 := strconv.Atoi(toRow); err2 == nil {
						c.To = grid.CellName(tr+1, tc+1)
					}
				}
				const emuPerPixel = 9525
				c.Width, c.Height = cx/emuPerPixel, cy/emuPerPixel
				out = append(out, c)
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "from":
				inFrom = false
			case "to":
				inTo = false
			}
			in = ""
		case xml.CharData:
			switch in {
			case "col":
				col += strings.TrimSpace(string(t))
			case "row":
				row += strings.TrimSpace(string(t))
			case "tocol":
				toCol += strings.TrimSpace(string(t))
			case "torow":
				toRow += strings.TrimSpace(string(t))
			}
		}
	}
	return out
}

// chart reads one chart part.
func (p *pkg) chart(part string) grid.Chart {
	c := grid.Chart{}
	r, err := p.open(part)
	if err != nil {
		return c
	}
	defer r.Close()
	dec := xml.NewDecoder(r)
	var (
		types      []string
		title      strings.Builder
		inTitle    int
		inText     bool
		ser        *grid.ChartSeries
		field      string // tx, cat, val (or xVal, yVal) while inside a series
		inF        bool
		f          strings.Builder
		barDir     string
		grouping   string
		inPlotArea bool
		inDPt      bool
		current    string // plot type of the series being read
		inLn       int    // inside a line's properties
		skip       int    // inside a marker or data labels
		pointSet   bool   // the current data point's fill was read
		fillColor  string
		lineColor  string
		inValAx    bool
		inCatAx    bool
		inLegend   bool
	)
	for {
		tok, err := dec.RawToken()
		if err == io.EOF || err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			name := t.Name.Local
			switch {
			case name == "plotArea":
				inPlotArea = true
			case name == "title" && ser == nil && !inPlotArea:
				inTitle++ // the chart's title; axis titles live in the plot area
			case name == "t" && inTitle > 0:
				inText = true
			case inPlotArea && strings.HasSuffix(name, "Chart") && name != "chart":
				types = append(types, chartKind(name))
				current = chartKind(name)
			case name == "barDir":
				barDir = attr(t, "val")
			case name == "grouping":
				grouping = attr(t, "val")
			case name == "ser":
				ser = &grid.ChartSeries{Kind: current}
			case ser != nil && (name == "tx" || name == "cat" || name == "val" || name == "xVal" || name == "yVal"):
				field = name
			case ser != nil && name == "f":
				inF = true
				f.Reset()
			case ser != nil && name == "dPt":
				inDPt, pointSet = true, false
			case ser != nil && name == "ln":
				inLn++
			case ser != nil && (name == "marker" || name == "dLbls"):
				skip++ // marker and label colours are not the series colour
			case ser != nil && name == "srgbClr" && field == "" && skip == 0:
				// A shape's colour is its fill; a line series has only a line.
				val := strings.ToUpper(attr(t, "val"))
				switch {
				case inDPt && inLn == 0 && !pointSet:
					ser.PointColors = append(ser.PointColors, val)
					pointSet = true
				case !inDPt && inLn == 0:
					fillColor = val
				case !inDPt && inLn > 0 && lineColor == "":
					lineColor = val
				}
			case name == "valAx":
				inValAx = true
			case name == "catAx" || name == "dateAx":
				inCatAx = true
			case name == "orientation" && inCatAx:
				c.Reversed = attr(t, "val") == "maxMin"
			case name == "numFmt" && inValAx && c.ValueFormat == "":
				c.ValueFormat = attr(t, "formatCode")
			case name == "crosses" && inValAx:
				c.CrossesMax = attr(t, "val") == "max"
			case name == "legend":
				inLegend = true
				c.Legend = "r" // Excel's default position
			case name == "legendPos" && inLegend:
				c.Legend = attr(t, "val")
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "title":
				if ser == nil && !inPlotArea && inTitle > 0 {
					inTitle--
				}
			case "t":
				inText = false
			case "f":
				if inF && ser != nil {
					switch field {
					case "tx":
						ser.Name = f.String()
					case "cat", "xVal":
						ser.Categories = f.String()
					case "val", "yVal":
						ser.Values = f.String()
					}
				}
				inF = false
			case "tx", "cat", "val", "xVal", "yVal":
				field = ""
			case "ser":
				if ser != nil {
					ser.Color = fillColor
					if ser.Color == "" {
						ser.Color = lineColor
					}
					c.Series = append(c.Series, *ser)
				}
				ser, fillColor, lineColor = nil, "", ""
			case "ln":
				if ser != nil && inLn > 0 {
					inLn--
				}
			case "marker", "dLbls":
				if ser != nil && skip > 0 {
					skip--
				}
			case "plotArea":
				inPlotArea = false
			case "dPt":
				inDPt = false
			case "valAx":
				inValAx = false
			case "catAx", "dateAx":
				inCatAx = false
			case "legend":
				inLegend = false
			}
		case xml.CharData:
			switch {
			case inF:
				f.Write(t)
			case inText:
				title.Write(t)
			}
		}
	}
	for i, k := range types {
		if k == "bar" && barDir == "col" {
			k = "column"
		}
		if (k == "bar" || k == "column" || k == "line" || k == "area") && (grouping == "stacked" || grouping == "percentStacked") {
			k = map[string]string{"stacked": "stacked ", "percentStacked": "100% stacked "}[grouping] + k
		}
		types[i] = k
	}
	for i := range c.Series {
		if c.Series[i].Kind == "bar" && barDir == "col" {
			c.Series[i].Kind = "column"
		}
	}
	// A line on the same data as an area is the area's edge (sheets draws
	// areas so): an area chart, not an area and a line.
	lines := false
	for i := range c.Series {
		s := &c.Series[i]
		for _, a := range c.Series {
			if s.Kind == "line" && a.Kind == "area" && a.Values == s.Values && a.Categories == s.Categories {
				s.Outline = true
			}
		}
		lines = lines || s.Kind == "line" && !s.Outline
	}
	if !lines {
		kept := types[:0]
		for _, k := range types {
			if k != "line" || len(types) == 1 {
				kept = append(kept, k)
			}
		}
		types = kept
	}
	c.Type = strings.Join(types, "+")
	c.Title = strings.TrimSpace(title.String())
	return c
}

// chartKind names a plot type element: barChart, line3DChart, ...
func chartKind(element string) string {
	k := strings.TrimSuffix(element, "Chart")
	k = strings.TrimSuffix(k, "3D")
	switch k {
	case "ofPie":
		return "pie"
	}
	return k
}

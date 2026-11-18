package xlsx

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"path"
	"regexp"
	"strconv"
	"strings"

	"github.com/eemax/sheets/internal/grid"
)

// What a workbook carries besides its cells, for telling what a change
// lost (sheets check --against): charts, data validations, conditional
// formats, shapes, pictures, defined names and Excel tables.

// Feature is one thing a workbook carries besides its cells.
type Feature struct {
	// Kind is chart, validation, conditional_format, shape, picture, name
	// or table.
	Kind string `json:"kind"`
	// Sheet is the sheet it is on; for a name, the sheet it is scoped to
	// ("" for the workbook).
	Sheet string `json:"sheet,omitempty"`
	// Where is the range it covers (validations, conditional formats,
	// tables), the cell its top-left corner is anchored to (charts,
	// shapes, pictures), or the name.
	Where string `json:"where"`
	// What says what it is: "list validation", "column chart", the name
	// of a shape.
	What string `json:"what,omitempty"`
	// Chartsheet: it is on a chart sheet (a sheet that is one chart), which
	// has no cells and no anchor.
	Chartsheet bool `json:"chart_sheet,omitempty"`
}

// Feature kinds.
const (
	FeatureChart      = "chart"
	FeatureValidation = "validation"
	FeatureCond       = "conditional_format"
	FeatureShape      = "shape"
	FeaturePicture    = "picture"
	FeatureName       = "name"
	FeatureTable      = "table"
)

var (
	validationTag = regexp.MustCompile(`<(?:x:)?dataValidation\b[^>]*>`)
	x14Validation = regexp.MustCompile(`(?s)<x14:dataValidation\b([^>]*)>.*?</x14:dataValidation>`)
	condTag       = regexp.MustCompile(`(?s)<(?:x:)?conditionalFormatting\b([^>]*)>(.*?)</(?:x:)?conditionalFormatting>`)
	ruleTag       = regexp.MustCompile(`<(?:x:)?cfRule\b[^>]*>`)
	x14CfRule     = regexp.MustCompile(`<x14:cfRule\b[^>]*>`)
	x14RuleID     = regexp.MustCompile(`<x14:id>([^<]*)</x14:id>`)
	attrRe        = regexp.MustCompile(`\b([\w:]+)="([^"]*)"`)
)

// attrsOf reads the attributes of a start tag as written.
func attrsOf(tag string) map[string]string {
	out := map[string]string{}
	for _, m := range attrRe.FindAllStringSubmatch(tag, -1) {
		out[m[1]] = m[2]
	}
	return out
}

// Features lists what a workbook (an xlsx package held in memory)
// carries besides its cells, sheet by sheet in the workbook's order.
func Features(data []byte) ([]Feature, error) {
	if bytes.HasPrefix(data, oleMagic) {
		return nil, fmt.Errorf("file is password-protected or in the legacy .xls format; save it as an unencrypted .xlsx")
	}
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, fmt.Errorf("not an xlsx file: %w", err)
	}
	p := &pkg{files: map[string]*zip.File{}}
	for _, f := range zr.File {
		p.files[strings.ToLower(f.Name)] = f
	}
	wbPath := "xl/workbook.xml"
	if rels, err := p.rels("_rels/.rels", ""); err == nil {
		for _, r := range rels {
			if strings.HasSuffix(r.Type, "/officeDocument") {
				wbPath = r.Target
			}
		}
	}
	var wbx xWorkbook
	if err := p.decode(wbPath, &wbx); err != nil {
		return nil, fmt.Errorf("reading workbook: %w", err)
	}
	wbDir := path.Dir(wbPath)
	rels, err := p.rels(path.Join(wbDir, "_rels", path.Base(wbPath)+".rels"), wbDir)
	if err != nil {
		return nil, fmt.Errorf("reading workbook relationships: %w", err)
	}
	relByID := map[string]xRel{}
	for _, r := range rels {
		relByID[r.ID] = r
	}
	var out []Feature
	for _, s := range wbx.Sheets {
		rel, ok := relByID[s.relID()]
		if ok && strings.HasSuffix(rel.Type, "/chartsheet") {
			// A sheet that is one chart: sheets does not read it as a sheet,
			// but losing it loses the chart.
			if drawing := p.sheetParts(rel.Target).drawing; drawing != "" {
				for _, f := range p.drawn(s.Name, drawing) {
					f.Chartsheet = true
					out = append(out, f)
				}
			}
			continue
		}
		if !ok || !strings.HasSuffix(rel.Type, "/worksheet") {
			continue
		}
		b, err := p.read(rel.Target)
		if err != nil {
			return nil, fmt.Errorf("reading sheet %q: %w", s.Name, err)
		}
		out = append(out, sheetRules(s.Name, b)...)
		parts := p.sheetParts(rel.Target)
		if parts.drawing != "" {
			out = append(out, p.drawn(s.Name, parts.drawing)...)
		}
		for _, t := range parts.tables {
			out = append(out, Feature{Kind: FeatureTable, Sheet: s.Name, Where: t.ref.String(), What: t.name})
		}
	}
	for _, n := range wbx.Names {
		if strings.HasPrefix(n.Name, "_xlnm.") || strings.HasPrefix(n.Name, recordPrefix) {
			continue
		}
		f := Feature{Kind: FeatureName, Where: n.Name, What: strings.TrimSpace(n.Ref)}
		if n.LocalSheetID != "" {
			if i, err := strconv.Atoi(n.LocalSheetID); err == nil && i < len(wbx.Sheets) {
				f.Sheet = wbx.Sheets[i].Name
			}
		}
		out = append(out, f)
	}
	return out, nil
}

// sheetRules lists a sheet's data validations and conditional formats,
// those in Excel 2010's extension list too (where Excel keeps a list read
// from another sheet, and the parts of a rule older Excel cannot show).
func sheetRules(sheet string, b []byte) []Feature {
	s := string(b)
	var out []Feature
	for _, tag := range validationTag.FindAllString(s, -1) {
		a := attrsOf(tag)
		out = append(out, Feature{Kind: FeatureValidation, Sheet: sheet, Where: a["sqref"], What: validationWhat(a["type"])})
	}
	for _, m := range x14Validation.FindAllStringSubmatch(s, -1) {
		a := attrsOf("<v " + m[1] + ">")
		where := ""
		if sq := x14Sqref.FindStringSubmatch(m[0]); sq != nil {
			where = strings.TrimSpace(sq[1])
		}
		out = append(out, Feature{Kind: FeatureValidation, Sheet: sheet, Where: where, What: validationWhat(a["type"])})
	}
	// A rule of Excel 2010's list that one of the main list points to
	// (x14:id) is part of that rule, not one of its own.
	linked := map[string]bool{}
	for _, m := range condTag.FindAllStringSubmatch(s, -1) {
		a := attrsOf("<c " + m[1] + ">")
		for _, rule := range ruleTag.FindAllString(m[2], -1) {
			out = append(out, Feature{Kind: FeatureCond, Sheet: sheet, Where: a["sqref"], What: condWhat(attrsOf(rule)["type"])})
		}
		for _, id := range x14RuleID.FindAllStringSubmatch(m[2], -1) {
			linked[strings.TrimSpace(id[1])] = true
		}
	}
	for _, blk := range x14Block.FindAllString(s, -1) {
		where := ""
		if sq := x14Sqref.FindStringSubmatch(blk); sq != nil {
			where = strings.TrimSpace(sq[1])
		}
		for _, rule := range x14CfRule.FindAllString(blk, -1) {
			a := attrsOf(rule)
			if linked[a["id"]] {
				continue
			}
			out = append(out, Feature{Kind: FeatureCond, Sheet: sheet, Where: where, What: condWhat(a["type"])})
		}
	}
	return out
}

func validationWhat(typ string) string {
	switch typ {
	case "", "none":
		return "validation (any value, with a message)"
	case "textLength":
		return "text length validation"
	}
	return typ + " validation"
}

func condWhat(typ string) string {
	switch typ {
	case "cellIs":
		return "cell value rule"
	case "expression":
		return "formula rule"
	case "dataBar":
		return "data bar"
	case "colorScale":
		return "color scale"
	case "iconSet":
		return "icon set"
	case "":
		return "rule"
	}
	return typ + " rule"
}

// drawn lists the charts, shapes and pictures a sheet's drawing holds,
// each by the cell its top-left corner is anchored to.
func (p *pkg) drawn(sheet, drawingPart string) []Feature {
	charts := map[string]grid.Chart{}
	for _, c := range p.charts(drawingPart) {
		if _, seen := charts[c.Anchor]; !seen {
			charts[c.Anchor] = c
		}
	}
	r, err := p.open(drawingPart)
	if err != nil {
		return nil
	}
	defer r.Close()
	dec := xml.NewDecoder(r)
	var out []Feature
	// Each chart, shape or picture, those in a group too (at the group's
	// anchor).
	var (
		inFrom          bool
		in, col, row    string
		kind, name      string
		member, isChart bool
		depth, at       int
	)
	where := func() string {
		if ci, err1 := strconv.Atoi(col); err1 == nil {
			if ri, err2 := strconv.Atoi(row); err2 == nil {
				return grid.CellName(ri+1, ci+1)
			}
		}
		return ""
	}
	for {
		tok, err := dec.RawToken()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			switch t.Name.Local {
			case "twoCellAnchor", "oneCellAnchor", "absoluteAnchor":
				col, row, member = "", "", false
			case "from":
				inFrom = true
			case "col", "row":
				if inFrom {
					in = t.Name.Local
				}
			case "sp", "cxnSp", "pic", "graphicFrame", "contentPart":
				if !member {
					member, kind, name, isChart, at = true, t.Name.Local, "", false, depth
				}
			case "cNvPr":
				if member && name == "" {
					name = attr(t, "name")
				}
			case "chart":
				if member && kind == "graphicFrame" {
					isChart = true
				}
			}
		case xml.EndElement:
			if member && depth == at {
				member = false
				switch w := where(); {
				case isChart:
					what := "chart"
					if c, ok := charts[w]; ok && c.Type != "" {
						what = c.Type + " chart"
						if c.Title != "" {
							what += " " + strconv.Quote(c.Title)
						}
					}
					out = append(out, Feature{Kind: FeatureChart, Sheet: sheet, Where: w, What: what})
				case kind == "pic":
					out = append(out, Feature{Kind: FeaturePicture, Sheet: sheet, Where: w, What: name})
				default:
					out = append(out, Feature{Kind: FeatureShape, Sheet: sheet, Where: w, What: name})
				}
			}
			depth--
			if t.Name.Local == "from" {
				inFrom = false
			}
			in = ""
		case xml.CharData:
			switch in {
			case "col":
				col += strings.TrimSpace(string(t))
			case "row":
				row += strings.TrimSpace(string(t))
			}
		}
	}
	return out
}

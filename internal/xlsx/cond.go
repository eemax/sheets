package xlsx

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"regexp"
	"sort"
	"strconv"
	"strings"

	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
	"github.com/xuri/excelize/v2"
)

// A rule is a conditional format: a condition and the look of the cells
// that meet it, or colour scales, data bars or icons.
type rule struct {
	opts excelize.ConditionalFormatOptions
	look *Look // the cells' look when the condition holds; nil for scales, bars and icons
}

// The look of a rule given none: Excel's light red fill with dark red text.
var defaultHighlight = Look{Fill: "FFC7CE", Color: "9C0006"}

const condHelp = `'< 0', '>= 100', '= East', 'between 1 and 10', 'top 10', 'bottom 5%', 'above average', duplicates, unique, blank, 'not blank', errors, 'contains x', 'begins with x', 'ends with x', '=$E2="Late"' (a formula for the range's first cell), scale [red-yellow-green], bars [color], icons [arrows|traffic|flags|ratings]`

// parseRule reads a condition from the start of ws and the look after it.
// r is the range it applies to: formulas are written for its first cell.
func parseRule(ws []string, r grid.Range) (rule, error) {
	if len(ws) == 0 {
		return rule{}, fmt.Errorf("no condition: give one of %s", condHelp)
	}
	// A whole condition in one word, as it is quoted for the shell:
	// 'between 1 and 10', 'contains past due'.
	if first := strings.TrimSpace(ws[0]); strings.ContainsAny(first, " \t") && !strings.HasPrefix(first, "=") || strings.HasPrefix(first, "= ") {
		fs := strings.Fields(first)
		// The text of a text condition stays one value.
		for _, lead := range [][]string{{"not", "contains"}, {"contains"}, {"begins", "with"}, {"starts", "with"}, {"ends", "with"}} {
			if len(fs) > len(lead) && strings.EqualFold(strings.Join(fs[:len(lead)], " "), strings.Join(lead, " ")) {
				rest := strings.TrimSpace(first)
				for range lead {
					rest = strings.TrimSpace(rest[strings.IndexAny(rest, " \t"):])
				}
				fs = append(fs[:len(lead):len(lead)], rest)
				break
			}
		}
		ws = append(fs, ws[1:]...)
	}
	// An operator glued to its value: '<0', '>=100'.
	var ts []string
	for i, w := range ws {
		if i == 0 && !strings.HasPrefix(w, "=") || i == 0 && strings.HasPrefix(w, "==") || i == 0 && strings.HasPrefix(w, "= ") {
			if op := leadingOp(w); op != "" {
				ts = append(ts, op)
				if rest := strings.TrimSpace(w[len(op):]); rest != "" {
					ts = append(ts, rest)
				}
				continue
			}
		}
		ts = append(ts, w)
	}
	cell := grid.CellName(r.R1, r.C1)
	var (
		ru rule
		n  int // words the condition took
	)
	word := func(i int) string {
		if i < len(ts) {
			return strings.ToLower(ts[i])
		}
		return ""
	}
	value := func(i int) (string, error) {
		if i >= len(ts) {
			return "", fmt.Errorf("%s: a value is missing", strings.Join(ts, " "))
		}
		return valueExpr(ts[i])
	}
	textRule := func(i int, make func(q string) string) error {
		if i >= len(ts) {
			return fmt.Errorf("%s: the text is missing", strings.Join(ts, " "))
		}
		q := `"` + strings.ReplaceAll(ts[i], `"`, `""`) + `"`
		ru.opts = excelize.ConditionalFormatOptions{Type: "formula", Criteria: make(q)}
		n = i + 1
		return nil
	}
	formulaRule := func(f string) {
		ru.opts = excelize.ConditionalFormatOptions{Type: "formula", Criteria: f}
	}
	var err error
	switch w := word(0); {
	case w == "value" && leadingOp(word(1)) != "":
		ts = ts[1:]
		return parseRule(ts, r)
	case leadingOp(w) == w && w != "":
		op := map[string]string{"==": "=", "!=": "<>"}[w]
		if op == "" {
			op = w
		}
		var v string
		if v, err = value(1); err == nil {
			ru.opts = excelize.ConditionalFormatOptions{Type: "cell", Criteria: op, Value: v}
			n = 2
		}
	case w == "between" || w == "not" && word(1) == "between":
		i := 1
		crit := "between"
		if w == "not" {
			i, crit = 2, "not between"
		}
		if word(i+1) != "and" {
			return rule{}, fmt.Errorf("expected: %s X and Y", crit)
		}
		lo, err1 := value(i)
		hi, err2 := value(i + 2)
		if err1 != nil || err2 != nil {
			return rule{}, fmt.Errorf("expected: %s X and Y", crit)
		}
		ru.opts = excelize.ConditionalFormatOptions{Type: "cell", Criteria: crit, MinValue: lo, MaxValue: hi}
		n = i + 3
	case w == "top" || w == "bottom":
		count := strings.TrimSpace(word(1))
		pct := strings.HasSuffix(count, "%")
		k, perr := strconv.Atoi(strings.TrimSuffix(count, "%"))
		if perr != nil || k < 1 || k > 1000 || pct && k > 100 {
			return rule{}, fmt.Errorf("expected: %s N or %s N%%, like %s 10", w, w, w)
		}
		ru.opts = excelize.ConditionalFormatOptions{Type: w, Criteria: "=", Value: strconv.Itoa(k), Percent: pct}
		n = 2
	case (w == "above" || w == "below") && word(1) == "average":
		ru.opts = excelize.ConditionalFormatOptions{Type: "average", Criteria: "=", AboveAverage: w == "above"}
		n = 2
	case w == "duplicates" || w == "duplicate":
		ru.opts, n = excelize.ConditionalFormatOptions{Type: "duplicate", Criteria: "="}, 1
	case w == "unique":
		ru.opts, n = excelize.ConditionalFormatOptions{Type: "unique", Criteria: "="}, 1
	case w == "blank" || w == "blanks":
		formulaRule("LEN(TRIM(" + cell + "))=0")
		n = 1
	case (w == "not" || w == "no") && (word(1) == "blank" || word(1) == "blanks"):
		formulaRule("LEN(TRIM(" + cell + "))>0")
		n = 2
	case w == "errors" || w == "error":
		formulaRule("ISERROR(" + cell + ")")
		n = 1
	case (w == "not" || w == "no") && (word(1) == "errors" || word(1) == "error"):
		formulaRule("NOT(ISERROR(" + cell + "))")
		n = 2
	case w == "contains":
		err = textRule(1, func(q string) string { return "ISNUMBER(SEARCH(" + q + "," + cell + "))" })
	case w == "not" && word(1) == "contains":
		err = textRule(2, func(q string) string { return "NOT(ISNUMBER(SEARCH(" + q + "," + cell + ")))" })
	case (w == "begins" || w == "starts") && word(1) == "with":
		err = textRule(2, func(q string) string { return "LEFT(" + cell + ",LEN(" + q + "))=" + q })
	case w == "ends" && word(1) == "with":
		err = textRule(2, func(q string) string { return "RIGHT(" + cell + ",LEN(" + q + "))=" + q })
	case w == "formula" || strings.HasPrefix(ts[0], "="):
		f := ts[0]
		n = 1
		if w == "formula" {
			if len(ts) < 2 {
				return rule{}, fmt.Errorf("formula: the formula is missing")
			}
			f, n = ts[1], 2
		}
		f = strings.TrimPrefix(strings.TrimSpace(f), "=")
		if err := formula.Check(f); err != nil {
			return rule{}, fmt.Errorf("=%s is not a valid formula: %v", f, err)
		}
		formulaRule(formula.ToFile(f))
	case w == "scale":
		lo, mid, hi, k, err := parseScale(ts[1:])
		if err != nil {
			return rule{}, err
		}
		n = 1 + k
		ru.opts = excelize.ConditionalFormatOptions{Type: "2_color_scale", Criteria: "=", MinType: "min", MaxType: "max", MinColor: "#" + lo, MaxColor: "#" + hi}
		if mid != "" {
			ru.opts.Type, ru.opts.MidType, ru.opts.MidValue, ru.opts.MidColor = "3_color_scale", "percentile", "50", "#"+mid
		}
	case w == "bars" || w == "bar" || w == "databars":
		color := "638EC6"
		n = 1
		if len(ts) > 1 && !IsLookWord(ts[1]) {
			if color, err = ParseColor(ts[1]); err != nil {
				return rule{}, err
			}
			n = 2
		}
		ru.opts = excelize.ConditionalFormatOptions{Type: "data_bar", Criteria: "=", MinType: "min", MaxType: "max", BarColor: "#" + color, BarSolid: true}
	case w == "icons":
		set := "3Arrows"
		n = 1
		if len(ts) > 1 && !IsLookWord(ts[1]) {
			s, ok := iconSets[strings.ToLower(ts[1])]
			if !ok {
				return rule{}, fmt.Errorf("icons %s: expected arrows, traffic, flags, symbols or ratings", ts[1])
			}
			set, n = s, 2
		}
		ru.opts = excelize.ConditionalFormatOptions{Type: "icon_set", IconStyle: set}
	default:
		return rule{}, fmt.Errorf("%q is not a condition: give one of %s", strings.Join(ts, " "), condHelp)
	}
	if err != nil {
		return rule{}, err
	}
	rest := ts[n:]
	switch ru.opts.Type {
	case "2_color_scale", "3_color_scale", "data_bar", "icon_set":
		if len(rest) > 0 {
			return rule{}, fmt.Errorf("%s draws its own look; %q does not apply", ts[0], strings.Join(rest, " "))
		}
		return ru, nil
	}
	l := defaultHighlight
	if len(rest) > 0 {
		if l, err = ParseLook(rest); err != nil {
			return rule{}, err
		}
		if l.Reset || l.Font != "" || l.Size > 0 || l.Align != "" || l.VAlign != "" || l.Wrap != nil || l.Indent != nil ||
			l.Rotate != nil || l.Width != 0 || l.Height != 0 || l.Merge != nil {
			return rule{}, fmt.Errorf("a conditional format can set color, fill, bold, italic, underline, strike, border and format; not %s", strings.Join(rest, " "))
		}
	}
	ru.look = &l
	return ru, nil
}

var iconSets = map[string]string{"arrows": "3Arrows", "traffic": "3TrafficLights1", "lights": "3TrafficLights1", "flags": "3Flags", "symbols": "3Symbols", "ratings": "5Rating", "stars": "3Stars"}

// parseScale reads an optional LOW-[MID-]HIGH colour list; the default is
// Excel's red-yellow-green, highest best.
func parseScale(ws []string) (lo, mid, hi string, n int, err error) {
	if len(ws) == 0 || IsLookWord(ws[0]) {
		return "F8696B", "FFEB84", "63BE7B", 0, nil
	}
	named := map[string]string{"red": "F8696B", "yellow": "FFEB84", "green": "63BE7B", "white": "FFFFFF", "blue": "5A8AC6"}
	parts := strings.Split(ws[0], "-")
	if len(parts) < 2 || len(parts) > 3 {
		return "", "", "", 0, fmt.Errorf("scale %s: expected LOW-HIGH or LOW-MID-HIGH colours, like red-yellow-green or white-blue", ws[0])
	}
	var cs []string
	for _, p := range parts {
		c, ok := named[strings.ToLower(p)]
		if !ok {
			if c, err = ParseColor(p); err != nil {
				return "", "", "", 0, fmt.Errorf("scale %s: %v", ws[0], err)
			}
		}
		cs = append(cs, c)
	}
	if len(cs) == 2 {
		return cs[0], "", cs[1], 1, nil
	}
	return cs[0], cs[1], cs[2], 1, nil
}

// leadingOp returns the comparison operator a word starts with.
func leadingOp(w string) string {
	for _, op := range []string{"<=", ">=", "<>", "!=", "==", "<", ">", "="} {
		if strings.HasPrefix(w, op) {
			return op
		}
	}
	return ""
}

var isoDate = regexp.MustCompile(`^(\d{4})-(\d{2})-(\d{2})$`)

// valueExpr turns a value as given into a formula: numbers and dates as
// such, =formula as written, anything else as text.
func valueExpr(v string) (string, error) {
	v = strings.TrimSpace(v)
	if strings.HasPrefix(v, "=") {
		f := strings.TrimPrefix(v, "=")
		if err := formula.Check(f); err != nil {
			return "", fmt.Errorf("%s is not a valid formula: %v", v, err)
		}
		return formula.ToFile(f), nil
	}
	if _, err := strconv.ParseFloat(v, 64); err == nil {
		return v, nil
	}
	if p, ok := strings.CutSuffix(v, "%"); ok {
		if x, err := strconv.ParseFloat(p, 64); err == nil {
			return strconv.FormatFloat(x/100, 'f', -1, 64), nil
		}
	}
	if m := isoDate.FindStringSubmatch(v); m != nil {
		return fmt.Sprintf("DATE(%s,%s,%s)", m[1], strings.TrimLeft(m[2], "0"), strings.TrimLeft(m[3], "0")), nil
	}
	return `"` + strings.ReplaceAll(v, `"`, `""`) + `"`, nil
}

// dxf is the differential style of a rule's look.
func (l *Look) dxf() *excelize.Style {
	s := &excelize.Style{}
	if l.Bold != nil || l.Italic != nil || l.Underline != nil || l.Strike != nil || l.Color != "" {
		s.Font = &excelize.Font{Color: l.Color}
		if l.Bold != nil {
			s.Font.Bold = *l.Bold
		}
		if l.Italic != nil {
			s.Font.Italic = *l.Italic
		}
		if l.Underline != nil && *l.Underline {
			s.Font.Underline = "single"
		}
		if l.Strike != nil {
			s.Font.Strike = *l.Strike
		}
	}
	if l.Fill != "" && l.Fill != "none" {
		s.Fill = excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{l.Fill}}
	}
	for _, b := range l.Borders {
		for _, side := range b.sidesAt(edgeTop | edgeBottom | edgeLeft | edgeRight) {
			if b.Style > 0 {
				s.Border = append(s.Border, excelize.Border{Type: side, Color: b.Color, Style: b.Style})
			}
		}
	}
	if l.NumFmt != "" && l.NumFmt != "General" {
		code := l.NumFmt
		s.CustomNumFmt = &code
	}
	return s
}

// addRule adds a conditional format to a range.
func (e *Editor) addRule(sheet string, r grid.Range, ru rule) error {
	opts := ru.opts
	if ru.look != nil {
		id, err := e.f.NewConditionalStyle(ru.look.dxf())
		if err != nil {
			return err
		}
		opts.Format = &id
	}
	return e.addConds(sheet, r, []excelize.ConditionalFormatOptions{opts})
}

// A CondFormat is a conditional format as it stands in a sheet.
type CondFormat struct {
	Range    string `json:"range"`
	Rule     string `json:"rule"`
	Recorded bool   `json:"recorded,omitempty"` // fmt keeps it
	Banding  bool   `json:"banding,omitempty"`  // fmt's banded rows
}

var bandFormula = regexp.MustCompile(`^MOD\(ROW\(\)-\d+,2\)=1$`)

// xCond is a conditional format as a sheet part stores it.
type xCond struct {
	Sqref string `xml:"sqref,attr"`
	Rules []struct {
		Type         string   `xml:"type,attr"`
		Operator     string   `xml:"operator,attr"`
		Rank         string   `xml:"rank,attr"`
		Percent      string   `xml:"percent,attr"`
		Bottom       string   `xml:"bottom,attr"`
		AboveAverage string   `xml:"aboveAverage,attr"`
		Text         string   `xml:"text,attr"`
		TimePeriod   string   `xml:"timePeriod,attr"`
		Formulas     []string `xml:"formula"`
		Dxf          string   `xml:"dxfId,attr"`
		ColorScale   *struct {
			Colors []struct {
				RGB string `xml:"rgb,attr"`
			} `xml:"color"`
		} `xml:"colorScale"`
		DataBar *struct {
			Color struct {
				RGB string `xml:"rgb,attr"`
			} `xml:"color"`
		} `xml:"dataBar"`
		IconSet *struct {
			Set string `xml:"iconSet,attr"`
		} `xml:"iconSet"`
	} `xml:"cfRule"`
}

// A CondRule is one rule of a conditional format, as the file holds it.
type CondRule struct {
	Range    string   // the cells, as an sqref: "A2:F9 H2:H9"
	Type     string   // cellIs, expression, colorScale, dataBar, ...
	Formulas []string // as stored (_xlfn. prefixes and all)
	Dxf      int      // the differential format it applies; -1 for none
}

// CondRules lists the rules of a sheet's conditional formats, in the order
// of the file. (excelize's GetConditionalFormats keeps one of the rules
// sharing a range.)
func CondRules(data []byte, sheet string) ([]CondRule, error) {
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, err
	}
	p := &pkg{files: map[string]*zip.File{}}
	for _, zf := range zr.File {
		p.files[strings.ToLower(zf.Name)] = zf
	}
	_, _, sheets, err := p.workbookParts()
	if err != nil {
		return nil, err
	}
	var part string
	for name, pt := range sheets {
		if strings.EqualFold(name, sheet) {
			part = pt
		}
	}
	if part == "" {
		return nil, fmt.Errorf("sheet %q not found", sheet)
	}
	b, err := p.read(part)
	if err != nil {
		return nil, err
	}
	if i := bytes.LastIndex(b, []byte("</sheetData>")); i >= 0 {
		b = b[i:]
	}
	var out []CondRule
	for _, m := range condBlock.FindAll(b, -1) {
		var c xCond
		if err := xml.Unmarshal(m, &c); err != nil {
			return nil, err
		}
		for _, r := range c.Rules {
			dxf := -1
			if n, err := strconv.Atoi(r.Dxf); err == nil {
				dxf = n
			}
			out = append(out, CondRule{Range: c.Sqref, Type: r.Type, Formulas: r.Formulas, Dxf: dxf})
		}
	}
	return out, nil
}

// condFormats lists the conditional formats of a sheet part, top-left
// first. (excelize's GetConditionalFormats keys them by range, and loses
// all but one of the formats sharing a range.)
func condFormats(sheetXML []byte, sheet string, recs []Record) ([]CondFormat, error) {
	var out []CondFormat
	for _, m := range condBlock.FindAll(sheetXML, -1) {
		var c xCond
		if err := xml.Unmarshal(m, &c); err != nil {
			return nil, err
		}
		for _, r := range c.Rules {
			cf := CondFormat{Range: c.Sqref, Rule: describeRule(r.Type, r.Operator, r.Rank, truthy(r.Percent), truthy(r.Bottom), r.AboveAverage, r.Text, r.TimePeriod, r.Formulas)}
			switch {
			case r.ColorScale != nil:
				var cs []string
				for _, col := range r.ColorScale.Colors {
					cs = append(cs, rgb6(col.RGB))
				}
				cf.Rule = "scale " + strings.Join(cs, "-")
			case r.DataBar != nil:
				cf.Rule = "bars " + rgb6(r.DataBar.Color.RGB)
			case r.IconSet != nil:
				set := r.IconSet.Set
				if set == "" {
					set = "3TrafficLights1"
				}
				cf.Rule = "icons " + set
			}
			cf.Banding = r.Type == "expression" && len(r.Formulas) == 1 && bandFormula.MatchString(r.Formulas[0])
			cf.Recorded = !cf.Banding && recordedCond(recs, sheet, c.Sqref)
			out = append(out, cf)
		}
	}
	sort.SliceStable(out, func(i, j int) bool {
		a, _ := firstRange(out[i].Range)
		b, _ := firstRange(out[j].Range)
		if a.R1 != b.R1 {
			return a.R1 < b.R1
		}
		return a.C1 < b.C1
	})
	return out, nil
}

func rgb6(c string) string {
	if len(c) > 6 {
		return strings.ToUpper(c[len(c)-6:])
	}
	return strings.ToUpper(c)
}

// describeRule says what a rule does, roughly as it was asked for.
func describeRule(typ, op, rank string, pct, bottom bool, above, text, period string, fs []string) string {
	ops := map[string]string{"lessThan": "<", "lessThanOrEqual": "<=", "greaterThan": ">", "greaterThanOrEqual": ">=",
		"equal": "=", "notEqual": "<>", "between": "between", "notBetween": "not between"}
	f := func(i int) string {
		if i < len(fs) {
			return formula.FromFile(fs[i])
		}
		return ""
	}
	switch typ {
	case "cellIs":
		o := ops[op]
		if o == "" {
			o = op
		}
		if strings.HasSuffix(o, "between") {
			return fmt.Sprintf("%s %s and %s", o, f(0), f(1))
		}
		return o + " " + f(0)
	case "top10":
		s := "top "
		if bottom {
			s = "bottom "
		}
		s += rank
		if pct {
			s += "%"
		}
		return s
	case "aboveAverage":
		if above == "0" || above == "false" {
			return "below average"
		}
		return "above average"
	case "duplicateValues":
		return "duplicates"
	case "uniqueValues":
		return "unique"
	case "expression":
		return "=" + f(0)
	case "containsText", "notContainsText", "beginsWith", "endsWith":
		return map[string]string{"containsText": "contains", "notContainsText": "not contains", "beginsWith": "begins with", "endsWith": "ends with"}[typ] + " " + text
	case "containsBlanks":
		return "blank"
	case "notContainsBlanks":
		return "not blank"
	case "containsErrors":
		return "errors"
	case "notContainsErrors":
		return "no errors"
	case "timePeriod":
		return "date " + period
	}
	return typ
}

func firstRange(sqref string) (grid.Range, error) {
	part, _, _ := strings.Cut(strings.TrimSpace(sqref), " ")
	return grid.ParseRange(strings.ReplaceAll(part, "$", ""))
}

// sqrefRanges parses a space-separated list of ranges.
func sqrefRanges(sqref string) []grid.Range {
	var out []grid.Range
	for _, p := range strings.Fields(sqref) {
		if r, err := grid.ParseRange(strings.ReplaceAll(p, "$", "")); err == nil {
			out = append(out, r)
		}
	}
	return out
}

// recordedCond tells whether the conditional format on sqref is one sheets
// cond recorded, which fmt keeps.
func recordedCond(recs []Record, sheet, sqref string) bool {
	rs := sqrefRanges(sqref)
	if len(rs) != 1 {
		return false
	}
	for _, rec := range recs {
		if rec.Kind == "cond" && strings.EqualFold(rec.Sheet, sheet) && rec.Range() == rs[0] {
			return true
		}
	}
	return false
}

var (
	condBlock   = regexp.MustCompile(`(?s)<conditionalFormatting\b[^>]*?\bsqref="([^"]*)"[^>]*>.*?</conditionalFormatting>`)
	x14Block    = regexp.MustCompile(`(?s)<x14:conditionalFormatting\b[^>]*>.*?</x14:conditionalFormatting>`)
	x14Sqref    = regexp.MustCompile(`<xm:sqref>([^<]*)</xm:sqref>`)
	x14CondsExt = regexp.MustCompile(`(?s)<ext\b[^>]*\buri="\{78C0D931-6437-407d-A8EE-F0AAD7539E65\}"[^>]*>\s*<x14:conditionalFormattings\b[^>]*>\s*</x14:conditionalFormattings>\s*</ext>`)
	emptyExtLst = regexp.MustCompile(`<extLst>\s*</extLst>`)
)

// dropCondBlocks removes from a sheet part the conditional formats whose
// range drop accepts, whole, with their Excel 2010 parts. (excelize's
// UnsetConditionalFormat takes cells out of every format's range, one
// cell at a time: seconds on a large sheet.) It returns the ranges of the
// formats it found.
func dropCondBlocks(b []byte, drop func(sqref string, banding bool) bool) ([]byte, []string) {
	var found []string
	b = condBlock.ReplaceAllFunc(b, func(m []byte) []byte {
		sq := string(condBlock.FindSubmatch(m)[1])
		banding := bandRule.Match(m)
		if !banding {
			found = append(found, sq)
		}
		if drop(sq, banding) {
			return nil
		}
		return m
	})
	b = x14Block.ReplaceAllFunc(b, func(m []byte) []byte {
		if sq := x14Sqref.FindSubmatch(m); sq != nil && drop(string(sq[1]), false) {
			return nil
		}
		return m
	})
	b = x14CondsExt.ReplaceAll(b, nil)
	return emptyExtLst.ReplaceAll(b, nil), found
}

// stripConds removes, before excelize opens a package, the conditional
// formats of the sheets fmt formats but the ones keep accepts. It returns
// the package and the ranges of the formats each sheet had.
func stripConds(data []byte, formatted func(sheet string) bool, keep func(sheet, sqref string) bool) ([]byte, map[string][]string, error) {
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, nil, err
	}
	p := &pkg{files: map[string]*zip.File{}}
	for _, zf := range zr.File {
		p.files[strings.ToLower(zf.Name)] = zf
	}
	_, _, sheets, err := p.workbookParts()
	if err != nil {
		return nil, nil, err
	}
	sheetOf := map[string]string{}
	for name, part := range sheets {
		if formatted(name) {
			sheetOf[strings.ToLower(part)] = name
		}
	}
	found := map[string][]string{}
	changed := false
	var out bytes.Buffer
	zw := zip.NewWriter(&out)
	for _, zf := range zr.File {
		r, err := zf.Open()
		if err != nil {
			return nil, nil, err
		}
		b, err := io.ReadAll(r)
		r.Close()
		if err != nil {
			return nil, nil, err
		}
		if sheet, ok := sheetOf[strings.ToLower(zf.Name)]; ok && bytes.Contains(b, []byte("<conditionalFormatting")) {
			n := len(b)
			// fmt's banding is fmt's to lay again, whatever shares its range.
			b, found[sheet] = dropCondBlocks(b, func(sq string, banding bool) bool { return banding || !keep(sheet, sq) })
			changed = changed || len(b) != n
		}
		w, err := zw.CreateHeader(&zip.FileHeader{Name: zf.Name, Method: zip.Deflate})
		if err != nil {
			return nil, nil, err
		}
		if _, err := w.Write(b); err != nil {
			return nil, nil, err
		}
	}
	if err := zw.Close(); err != nil {
		return nil, nil, err
	}
	if !changed {
		return data, found, nil
	}
	return out.Bytes(), found, nil
}

func overlaps(a, b []grid.Range) bool {
	for _, x := range a {
		for _, y := range b {
			if !x.Intersect(y).Empty() {
				return true
			}
		}
	}
	return false
}

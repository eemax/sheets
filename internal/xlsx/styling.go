package xlsx

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"regexp"
	"sort"
	"strconv"
	"strings"

	"github.com/eemax/sheets/internal/grid"
	"github.com/xuri/excelize/v2"
)

// This file is the editor's side of targeted styling: sheets style, cond,
// validate and view. What fmt would strip is recorded (see records.go).

// used is the used range of a sheet as edited so far.
func (e *Editor) used(sheet string) grid.Range {
	if ext, ok := e.extent[sheet]; ok {
		return ext
	}
	if s := e.wb.Sheet(sheet); s != nil {
		if u := s.Used(); !u.Empty() {
			return u
		}
	}
	return grid.Range{R1: 1, C1: 1, R2: 1, C2: 1}
}

func (e *Editor) styler() *styler {
	if e.st == nil {
		e.st = newStyler(e.f)
	}
	return e.st
}

// StyleResult is what sheets style did.
type StyleResult struct {
	Range    string   `json:"range"`    // the cells it styled
	Recorded bool     `json:"recorded"` // fmt keeps it
	Notes    []string `json:"notes,omitempty"`
}

// Style lays a look over a range (whole rows or columns: their used part)
// and records it, so that fmt keeps it.
func (e *Editor) Style(sheet string, r grid.Range, l *Look) (StyleResult, error) {
	sheet, err := e.sheetName(sheet)
	if err != nil {
		return StyleResult{}, err
	}
	at := clip(r, e.used(sheet))
	st := e.styler()
	before := len(st.Notes())
	if err := st.apply(sheet, at, l, nil); err != nil {
		return StyleResult{}, err
	}
	res := StyleResult{Range: grid.FormatRef(sheet, at), Notes: st.Notes()[before:]}
	// Merges stay through fmt on their own.
	var kept []string
	for _, w := range l.Words() {
		if lw := strings.ToLower(w); lw != "merge" && lw != "unmerge" {
			kept = append(kept, w)
		}
	}
	if len(kept) == 0 {
		return res, nil
	}
	res.Recorded, err = addRecord(e.f, Record{Kind: "style", Sheet: sheet, Range: r, Args: kept})
	if err == nil && !res.Recorded {
		res.Notes = append(res.Notes, "too long to record: fmt will not keep it")
	}
	return res, err
}

// Records lists the record, in order.
func (e *Editor) Records() []Record {
	recs, _ := readRecords(e.f)
	return recs
}

// Unrecord removes entries of the record by number; the styling stays
// until fmt runs.
func (e *Editor) Unrecord(seqs []int) error {
	recs, _ := readRecords(e.f)
	for _, seq := range seqs {
		found := false
		for _, r := range recs {
			if r.Seq == seq {
				found = true
				if err := dropRecords(e.f, r.name); err != nil {
					return err
				}
			}
		}
		if !found {
			return fmt.Errorf("no recorded styling numbered %d (sheets style --list shows them)", seq)
		}
	}
	return nil
}

// Cond adds a conditional format to a range, and records it so that fmt
// keeps it.
func (e *Editor) Cond(sheet string, r grid.Range, ws []string) (string, error) {
	sheet, err := e.sheetName(sheet)
	if err != nil {
		return "", err
	}
	at := clip(r, e.used(sheet))
	ru, err := parseRule(ws, at)
	if err != nil {
		return "", err
	}
	text := Record{Kind: "cond", Args: ws}.Text()
	for _, rec := range e.Records() {
		if rec.Kind == "cond" && strings.EqualFold(rec.Sheet, sheet) && rec.Range == at && rec.Text() == text {
			return grid.FormatRef(sheet, at), nil // there already
		}
	}
	if err := addRule(e.f, sheet, at, ru); err != nil {
		return "", err
	}
	if ok, err := addRecord(e.f, Record{Kind: "cond", Sheet: sheet, Range: at, Args: ws}); err != nil || !ok {
		return grid.FormatRef(sheet, at), err
	}
	return grid.FormatRef(sheet, at), nil
}

// CondFormats lists the conditional formats of a sheet as the file holds
// them (as opened).
func (e *Editor) CondFormats(sheet string) ([]CondFormat, error) {
	sheet, err := e.sheetName(sheet)
	if err != nil {
		return nil, err
	}
	x, err := e.sheetPart(sheet)
	if err != nil {
		return nil, err
	}
	return condFormats(x, sheet, e.Records())
}

// sheetPart returns a sheet's part as the file held it when opened.
func (e *Editor) sheetPart(sheet string) ([]byte, error) {
	if e.data == nil {
		return nil, nil
	}
	zr, err := zip.NewReader(bytes.NewReader(e.data), int64(len(e.data)))
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
	for name, part := range sheets {
		if strings.EqualFold(name, sheet) {
			r, err := p.open(part)
			if err != nil {
				return nil, err
			}
			defer r.Close()
			return io.ReadAll(r)
		}
	}
	return nil, nil
}

// ClearCond removes the conditional formats that lie within r (all of the
// sheet's when r is nil), and their records. It returns how many rules it
// removed.
func (e *Editor) ClearCond(sheet string, r *grid.Range) (int, error) {
	sheet, err := e.sheetName(sheet)
	if err != nil {
		return 0, err
	}
	cfs, err := e.CondFormats(sheet)
	if err != nil {
		return 0, err
	}
	hit := func(ref string) bool { return r == nil || within(sqrefRanges(ref), *r) }
	n := 0
	for _, cf := range cfs {
		if hit(cf.Range) {
			n++
			if e.patches.dropConds[sheet] == nil {
				e.patches.dropConds[sheet] = map[string]bool{}
			}
			e.patches.dropConds[sheet][cf.Range] = true
		}
	}
	for _, rec := range e.Records() {
		if rec.Kind == "cond" && strings.EqualFold(rec.Sheet, sheet) && hit(rec.Range.String()) {
			if err := dropRecords(e.f, rec.name); err != nil {
				return n, err
			}
		}
	}
	return n, nil
}

// within tells whether every range lies inside r.
func within(rs []grid.Range, r grid.Range) bool {
	for _, x := range rs {
		if !r.Contains(x.R1, x.C1) || !r.Contains(x.R2, x.C2) {
			return false
		}
	}
	return len(rs) > 0
}

// Validate adds a data validation rule to a range, replacing the rules
// that lie within it.
func (e *Editor) Validate(sheet string, r grid.Range, rule ValidationRule) (string, error) {
	sheet, err := e.sheetName(sheet)
	if err != nil {
		return "", err
	}
	at := clip(r, e.used(sheet))
	dv, err := rule.dataValidation(at)
	if err != nil {
		return "", err
	}
	// Rules must not overlap: the target leaves the ranges of others.
	if _, err := e.ClearValidation(sheet, &at); err != nil {
		return "", err
	}
	return grid.FormatRef(sheet, at), e.f.AddDataValidation(sheet, dv)
}

// Validations lists the data validation rules of a sheet.
func (e *Editor) Validations(sheet string) ([]Validation, error) {
	sheet, err := e.sheetName(sheet)
	if err != nil {
		return nil, err
	}
	return validations(e.f, sheet)
}

// ClearValidation takes the cells of r out of every rule's range, as
// Excel's Clear All does for a selection (all of the sheet's rules when r
// is nil). It returns how many rules it changed or removed.
func (e *Editor) ClearValidation(sheet string, r *grid.Range) (int, error) {
	sheet, err := e.sheetName(sheet)
	if err != nil {
		return 0, err
	}
	dvs, err := e.f.GetDataValidations(sheet)
	if err != nil {
		return 0, err
	}
	n := 0
	for _, dv := range dvs {
		if r == nil || overlaps(sqrefRanges(dv.Sqref), []grid.Range{*r}) {
			n++
		}
	}
	if n == 0 {
		return 0, nil
	}
	if r == nil {
		return n, e.f.DeleteDataValidation(sheet)
	}
	return n, e.f.DeleteDataValidation(sheet, r.String())
}

// ViewOptions are what sheets view changes; nil fields stay as they are.
type ViewOptions struct {
	Freeze    *grid.Range // the first cell that scrolls; A1 unfreezes
	Tab       *string     // RRGGBB, or "" to remove the colour
	Gridlines *bool
	Zoom      *int
	Hide      []grid.Range // whole rows or columns
	Show      []grid.Range
	Filter    *grid.Range // filter buttons on this range's first row
	NoFilter  bool        // remove the sheet's filter
	Active    bool
}

// ViewState is how a sheet shows.
type ViewState struct {
	Sheet      string   `json:"sheet"`
	Freeze     string   `json:"freeze,omitempty"` // the first cell that scrolls
	Tab        string   `json:"tab,omitempty"`
	Gridlines  bool     `json:"gridlines"`
	Zoom       int      `json:"zoom"`
	HiddenRows []string `json:"hidden_rows,omitempty"`
	HiddenCols []string `json:"hidden_cols,omitempty"`
	Filter     string   `json:"filter,omitempty"`
	Active     bool     `json:"active"`
}

// View changes how a sheet shows. Freeze, gridlines and filter are
// recorded, since fmt sets them too.
func (e *Editor) View(sheet string, o ViewOptions) error {
	sheet, err := e.sheetName(sheet)
	if err != nil {
		return err
	}
	f := e.f
	if o.Freeze != nil {
		if err := freeze(f, sheet, *o.Freeze); err != nil {
			return err
		}
		arg := "freeze"
		if o.Freeze.R1 == 1 && o.Freeze.C1 == 1 {
			arg = "nofreeze"
		}
		if _, err := addRecord(f, Record{Kind: "view", Sheet: sheet, Range: *o.Freeze, Args: []string{arg}}); err != nil {
			return err
		}
	}
	if o.Tab != nil {
		if *o.Tab == "" {
			e.patches.noTabColor[sheet] = true
		} else {
			delete(e.patches.noTabColor, sheet)
			rgb := "FF" + *o.Tab
			if err := f.SetSheetProps(sheet, &excelize.SheetPropsOptions{TabColorRGB: &rgb}); err != nil {
				return err
			}
		}
	}
	if o.Gridlines != nil {
		if err := f.SetSheetView(sheet, 0, &excelize.ViewOptions{ShowGridLines: o.Gridlines}); err != nil {
			return err
		}
		arg := map[bool]string{true: "gridlines=on", false: "gridlines=off"}[*o.Gridlines]
		if _, err := addRecord(f, Record{Kind: "view", Sheet: sheet, Range: grid.Range{R1: 1, C1: 1, R2: 1, C2: 1}, Args: []string{arg}}); err != nil {
			return err
		}
	}
	if o.Zoom != nil {
		z := float64(*o.Zoom)
		if err := f.SetSheetView(sheet, 0, &excelize.ViewOptions{ZoomScale: &z}); err != nil {
			return err
		}
	}
	for _, r := range o.Hide {
		if err := setVisible(f, sheet, r, false); err != nil {
			return err
		}
	}
	for _, r := range o.Show {
		if err := setVisible(f, sheet, r, true); err != nil {
			return err
		}
	}
	if o.Filter != nil || o.NoFilter {
		rec := Record{Kind: "view", Sheet: sheet, Range: grid.Range{R1: 1, C1: 1, R2: 1, C2: 1}, Args: []string{"nofilter"}}
		if o.Filter != nil {
			rec.Range, rec.Args = *o.Filter, []string{"filter"}
		}
		if err := e.filter(sheet, o.Filter); err != nil {
			return err
		}
		if _, err := addRecord(f, rec); err != nil {
			return err
		}
	}
	if o.Active {
		idx, err := f.GetSheetIndex(sheet)
		if err != nil {
			return err
		}
		f.SetActiveSheet(idx)
	}
	return nil
}

// freeze keeps the rows above and the columns left of a cell in view; A1
// unfreezes.
func freeze(f *excelize.File, sheet string, at grid.Range) error {
	if at.R1 == 1 && at.C1 == 1 {
		return f.SetPanes(sheet, &excelize.Panes{})
	}
	p := &excelize.Panes{Freeze: true, XSplit: at.C1 - 1, YSplit: at.R1 - 1, TopLeftCell: grid.CellName(at.R1, at.C1)}
	switch {
	case p.XSplit > 0 && p.YSplit > 0:
		p.ActivePane = "bottomRight"
	case p.YSplit > 0:
		p.ActivePane = "bottomLeft"
	default:
		p.ActivePane = "topRight"
	}
	return f.SetPanes(sheet, p)
}

// setVisible hides or shows whole rows or columns.
func setVisible(f *excelize.File, sheet string, r grid.Range, visible bool) error {
	if r.R1 == 1 && r.R2 == grid.MaxRows {
		return f.SetColVisible(sheet, grid.ColName(r.C1)+":"+grid.ColName(r.C2), visible)
	}
	for row := r.R1; row <= r.R2; row++ {
		if err := f.SetRowVisible(sheet, row, visible); err != nil {
			return err
		}
	}
	return nil
}

// filter puts filter buttons on a range's first row, or removes the
// sheet's filter when r is nil.
func (e *Editor) filter(sheet string, at *grid.Range) error {
	f := e.f
	if at == nil {
		e.patches.noFilter[sheet] = true
		if err := f.DeleteDefinedName(&excelize.DefinedName{Name: "_xlnm._FilterDatabase", Scope: sheet}); err != nil && !strings.Contains(err.Error(), "no defined name") {
			return err
		}
		return nil
	}
	r := *at
	for _, t := range e.tablesOn(sheet) {
		if !t.ref.Intersect(r).Empty() {
			return fmt.Errorf("%s overlaps Excel table %s, which has its own filter buttons", r, t)
		}
	}
	delete(e.patches.noFilter, sheet)
	// AutoFilter replaces all sheet properties (tab colour, fit to page,
	// ...); put them back.
	props, err := f.GetSheetProps(sheet)
	if err != nil {
		return err
	}
	if err := f.AutoFilter(sheet, r.String(), nil); err != nil {
		return err
	}
	return f.SetSheetProps(sheet, &props)
}

// ViewOf reports how a sheet shows.
func (e *Editor) ViewOf(sheet string) (ViewState, error) {
	sheet, err := e.sheetName(sheet)
	if err != nil {
		return ViewState{}, err
	}
	f := e.f
	v := ViewState{Sheet: sheet, Gridlines: true, Zoom: 100}
	if p, err := f.GetPanes(sheet); err == nil && p.Freeze && (p.XSplit > 0 || p.YSplit > 0) {
		v.Freeze = grid.CellName(p.YSplit+1, p.XSplit+1)
	}
	if props, err := f.GetSheetProps(sheet); err == nil && props.TabColorRGB != nil && *props.TabColorRGB != "" && !e.patches.noTabColor[sheet] {
		c := *props.TabColorRGB
		v.Tab = "#" + c[max(len(c)-6, 0):]
	}
	if so, err := f.GetSheetView(sheet, 0); err == nil {
		if so.ShowGridLines != nil {
			v.Gridlines = *so.ShowGridLines
		}
		if so.ZoomScale != nil && *so.ZoomScale > 0 {
			v.Zoom = int(*so.ZoomScale)
		}
	}
	if s := e.wb.Sheet(sheet); s != nil {
		for _, h := range s.HiddenRows {
			v.HiddenRows = append(v.HiddenRows, runText(h, strconv.Itoa))
		}
		for _, h := range s.HiddenCols {
			v.HiddenCols = append(v.HiddenCols, runText(h, grid.ColName))
		}
		if s.Filter.R1 > 0 && !e.patches.noFilter[sheet] {
			v.Filter = s.Filter.String()
		}
	}
	if idx, err := f.GetSheetIndex(sheet); err == nil {
		v.Active = idx == f.GetActiveSheetIndex()
	}
	return v, nil
}

func runText(r grid.Run, name func(int) string) string {
	if r.From == r.To {
		return name(r.From)
	}
	return name(r.From) + ":" + name(r.To)
}

// restyleWritten lays the recorded styling over the cells this session
// wrote that were empty before: rows appended to a styled table get its
// look, as they would in Excel.
func (e *Editor) restyleWritten() error {
	if len(e.changed) == 0 {
		return nil
	}
	recs := e.Records()
	if len(recs) == 0 {
		return nil
	}
	fresh := map[cellKey]bool{}
	box := map[string]grid.Range{}
	for k := range e.changed {
		if s := e.wb.Sheet(k.sheet); s != nil {
			if c := s.At(k.row, k.col); c != nil && !c.IsEmpty() {
				continue
			}
		}
		fresh[k] = true
		box[k.sheet] = box[k.sheet].Union(grid.Range{R1: k.row, C1: k.col, R2: k.row, C2: k.col})
	}
	for _, rec := range recs {
		b, ok := box[strings.ToLower(rec.Sheet)]
		if rec.Kind != "style" || !ok {
			continue
		}
		l, err := ParseLook(rec.Args)
		if err != nil || !l.CellStyle() {
			continue
		}
		r := clip(rec.Range, e.used(rec.Sheet))
		if r.Intersect(b).Empty() {
			continue
		}
		sheet := rec.Sheet
		err = e.styler().apply(sheet, r, &l, func(row, col int) bool { return fresh[keyOf(sheet, row, col)] })
		if err != nil {
			return err
		}
	}
	return nil
}

// growRecords extends recorded ranges, conditional formats and data
// validations that end at a table's last data row over n rows appended
// below it.
func (e *Editor) growRecords(sheet string, body grid.Range, n int) error {
	for _, rec := range e.Records() {
		r := rec.Range
		if !strings.EqualFold(rec.Sheet, sheet) || r.R2 != body.R2 || r.C1 < body.C1 || r.C2 > body.C2 {
			continue
		}
		r.R2 += n
		if err := dropRecords(e.f, rec.name); err != nil {
			return err
		}
		if err := e.f.SetDefinedName(&excelize.DefinedName{Name: rec.name, Comment: rec.Text(), RefersTo: absRef(rec.Sheet, r)}); err != nil {
			return err
		}
	}
	e.patches.grow = append(e.patches.grow, growth{sheet: sheet, from: body.R2, to: body.R2 + n, c1: body.C1, c2: body.C2})
	return nil
}

// growth extends ranges ending at row from to row to, within columns
// c1..c2 (see growRecords).
type growth struct {
	sheet    string
	from, to int
	c1, c2   int
}

var (
	sqrefAttr = regexp.MustCompile(`(<(?:conditionalFormatting|dataValidation)\b[^>]*?\bsqref=")([^"]*)(")`)
	xmSqref   = regexp.MustCompile(`(<xm:sqref>)([^<]*)(</xm:sqref>)`)
)

// growSqrefs applies growths to the conditional formats and data
// validations of a sheet part.
func growSqrefs(b []byte, gs []growth) []byte {
	fix := func(m []byte, re *regexp.Regexp) []byte {
		p := re.FindSubmatch(m)
		var out []string
		for _, part := range strings.Fields(string(p[2])) {
			r, err := grid.ParseRange(strings.ReplaceAll(part, "$", ""))
			if err != nil {
				out = append(out, part)
				continue
			}
			for _, g := range gs {
				if r.R2 == g.from && r.C1 >= g.c1 && r.C2 <= g.c2 {
					r.R2 = g.to
				}
			}
			out = append(out, r.String())
		}
		return append(append(append([]byte{}, p[1]...), strings.Join(out, " ")...), p[3]...)
	}
	b = sqrefAttr.ReplaceAllFunc(b, func(m []byte) []byte { return fix(m, sqrefAttr) })
	return xmSqref.ReplaceAllFunc(b, func(m []byte) []byte { return fix(m, xmSqref) })
}

var (
	cfRuleTag   = regexp.MustCompile(`<(?:x14:)?cfRule\b[^>]*>`)
	priorityAt  = regexp.MustCompile(` priority="(\d+)"`)
	cfFormula   = regexp.MustCompile(`(?s)(<cfRule\b[^>]*>)(.*?)(</cfRule>)`)
	ruleFormula = regexp.MustCompile(`(<formula>)([^<]*)(</formula>)`)
	bandRule    = regexp.MustCompile(`<formula>MOD\(ROW\(\)-\d+,2\)=1</formula>`)
)

// orderRules gives a sheet's conditional formats distinct priorities, with
// fmt's banding last so that a highlight wins over it. excelize numbers a
// new rule by the count of rules before it, which can repeat a number and
// puts it below banding. A sheet whose priorities are distinct with banding
// last is left as it is.
func orderRules(b []byte) []byte {
	type tag struct {
		at, end int
		prio    int
		band    bool
	}
	var tags []tag
	for _, loc := range cfRuleTag.FindAllIndex(b, -1) {
		m := priorityAt.FindSubmatch(b[loc[0]:loc[1]])
		if m == nil {
			continue
		}
		p, _ := strconv.Atoi(string(m[1]))
		end := loc[1] + 400
		if end > len(b) {
			end = len(b)
		}
		body := b[loc[1]:end]
		if i := strings.Index(string(body), "</cfRule>"); i >= 0 {
			body = body[:i]
		}
		tags = append(tags, tag{at: loc[0], end: loc[1], prio: p, band: bandRule.Match(body)})
	}
	if len(tags) == 0 {
		return b
	}
	seen := map[int]bool{}
	ok := true
	maxOther, minBand := 0, int(^uint(0)>>1)
	for _, t := range tags {
		if seen[t.prio] {
			ok = false
		}
		seen[t.prio] = true
		if t.band {
			minBand = min(minBand, t.prio)
		} else {
			maxOther = max(maxOther, t.prio)
		}
	}
	if ok && minBand > maxOther {
		return b
	}
	order := make([]int, len(tags))
	for i := range order {
		order[i] = i
	}
	sort.SliceStable(order, func(i, j int) bool {
		a, c := tags[order[i]], tags[order[j]]
		if a.band != c.band {
			return !a.band
		}
		return a.prio < c.prio
	})
	newPrio := make([]int, len(tags))
	for rank, i := range order {
		newPrio[i] = rank + 1
	}
	var out []byte
	last := 0
	for i, t := range tags {
		out = append(out, b[last:t.at]...)
		out = append(out, priorityAt.ReplaceAll(b[t.at:t.end], []byte(` priority="`+strconv.Itoa(newPrio[i])+`"`))...)
		last = t.end
	}
	return append(out, b[last:]...)
}

// rewriteRuleFormulas applies the row and column edits of a session to
// the formulas of a sheet's conditional formats, which excelize leaves as
// they were.
func rewriteRuleFormulas(b []byte, sheet string, rewrites []func(f, home string) string) []byte {
	return cfFormula.ReplaceAllFunc(b, func(m []byte) []byte {
		p := cfFormula.FindSubmatch(m)
		body := ruleFormula.ReplaceAllFunc(p[2], func(fm []byte) []byte {
			q := ruleFormula.FindSubmatch(fm)
			f := unescapeXML(string(q[2]))
			for _, rw := range rewrites {
				f = rw(f, sheet)
			}
			return append(append(append([]byte{}, q[1]...), xmlEscape(f)...), q[3]...)
		})
		return append(append(append([]byte{}, p[1]...), body...), p[3]...)
	})
}

var tabColor = regexp.MustCompile(`<tabColor\b[^>]*/>|<tabColor\b[^>]*>\s*</tabColor>`)

// A LookRun is a block of cells with the same look.
type LookRun struct {
	Range string   `json:"range"`
	Look  []string `json:"look"` // style properties, as sheets style takes them
}

// Looks describes the looks of the cells of r, in blocks of equal look.
func (e *Editor) Looks(sheet string, r grid.Range) ([]LookRun, error) {
	sheet, err := e.sheetName(sheet)
	if err != nil {
		return nil, err
	}
	r = clip(r, e.used(sheet))
	if r.Rows()*r.Cols() > 100000 {
		return nil, fmt.Errorf("%s has %d cells; show a smaller range", r, r.Rows()*r.Cols())
	}
	font, _ := e.f.GetDefaultFont()
	memo := map[int][]string{}
	desc := func(row, col int) (string, error) {
		id, err := e.f.GetCellStyle(sheet, grid.CellName(row, col))
		if err != nil {
			return "", err
		}
		if _, ok := memo[id]; !ok {
			s, err := e.f.GetStyle(id)
			if err != nil {
				return "", err
			}
			memo[id] = describeStyle(s, font)
		}
		return strings.Join(memo[id], "\x00"), nil
	}
	// Runs along each row, then runs of rows with the same runs.
	type run struct {
		c1, c2 int
		d      string
	}
	var rows [][]run
	for row := r.R1; row <= r.R2; row++ {
		var rs []run
		for col := r.C1; col <= r.C2; col++ {
			d, err := desc(row, col)
			if err != nil {
				return nil, err
			}
			if n := len(rs); n > 0 && rs[n-1].d == d {
				rs[n-1].c2 = col
				continue
			}
			rs = append(rs, run{col, col, d})
		}
		rows = append(rows, rs)
	}
	var out []LookRun
	type open struct {
		r1, r2 int
		run
	}
	var blocks []open
	for i, rs := range rows {
		row := r.R1 + i
		for _, x := range rs {
			grown := false
			for j := range blocks {
				b := &blocks[j]
				if b.r2 == row-1 && b.c1 == x.c1 && b.c2 == x.c2 && b.d == x.d {
					b.r2, grown = row, true
					break
				}
			}
			if !grown {
				blocks = append(blocks, open{row, row, x})
			}
		}
	}
	sort.SliceStable(blocks, func(i, j int) bool {
		if blocks[i].r1 != blocks[j].r1 {
			return blocks[i].r1 < blocks[j].r1
		}
		return blocks[i].c1 < blocks[j].c1
	})
	for _, b := range blocks {
		look := []string{}
		if b.d != "" {
			look = strings.Split(b.d, "\x00")
		}
		out = append(out, LookRun{Range: grid.FormatRef(sheet, grid.Range{R1: b.r1, C1: b.c1, R2: b.r2, C2: b.c2}), Look: look})
	}
	return out, nil
}

// builtinFormats are the codes of Excel's built-in number formats.
var builtinFormats = map[int]string{1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 9: "0%", 10: "0.00%", 11: "0.00E+00",
	12: "# ?/?", 13: "# ??/??", 14: "m/d/yyyy", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy", 18: "h:mm AM/PM", 19: "h:mm:ss AM/PM",
	20: "h:mm", 21: "h:mm:ss", 22: "m/d/yyyy h:mm", 37: "#,##0_);(#,##0)", 38: "#,##0_);[Red](#,##0)", 39: "#,##0.00_);(#,##0.00)",
	40: "#,##0.00_);[Red](#,##0.00)", 45: "mm:ss", 46: "[h]:mm:ss", 47: "mm:ss.0", 48: "##0.0E+0", 49: "@"}

var borderNames = map[int]string{1: "thin", 2: "medium", 3: "dashed", 4: "dotted", 5: "thick", 6: "double", 7: "hair",
	8: "mediumDashed", 9: "dashDot", 10: "mediumDashDot", 11: "dashDotDot", 12: "mediumDashDotDot", 13: "slantDashDot"}

// describeStyle writes a style as sheets style properties, leaving out
// what a cell shows without a style.
func describeStyle(s *excelize.Style, defaultFont string) []string {
	var out []string
	hex := func(c string) string {
		c = strings.ToUpper(strings.TrimPrefix(c, "#"))
		if len(c) > 6 {
			c = c[len(c)-6:]
		}
		return "#" + c
	}
	if f := s.Font; f != nil {
		if f.Bold {
			out = append(out, "bold")
		}
		if f.Italic {
			out = append(out, "italic")
		}
		if f.Underline != "" && f.Underline != "none" {
			out = append(out, "underline")
		}
		if f.Strike {
			out = append(out, "strike")
		}
		if f.Family != "" && !strings.EqualFold(f.Family, defaultFont) {
			out = append(out, "font="+f.Family)
		}
		if f.Size > 0 && f.Size != 11 {
			out = append(out, "size="+strconv.FormatFloat(f.Size, 'f', -1, 64))
		}
		if f.Color != "" && hex(f.Color) != "#000000" {
			out = append(out, "color="+hex(f.Color))
		}
	}
	if s.Fill.Type == "pattern" && s.Fill.Pattern == 1 && len(s.Fill.Color) > 0 {
		out = append(out, "fill="+hex(s.Fill.Color[0]))
	}
	switch {
	case s.CustomNumFmt != nil && *s.CustomNumFmt != "" && !strings.EqualFold(*s.CustomNumFmt, "general"):
		out = append(out, "format="+*s.CustomNumFmt)
	case s.NumFmt != 0:
		if code, ok := builtinFormats[s.NumFmt]; ok {
			out = append(out, "format="+code)
		} else {
			out = append(out, fmt.Sprintf("format=builtin:%d", s.NumFmt))
		}
	}
	if a := s.Alignment; a != nil {
		if a.Horizontal != "" && a.Horizontal != "general" {
			h := a.Horizontal
			if h == "centerContinuous" {
				h = "center-across"
			}
			out = append(out, "align="+h)
		}
		switch a.Vertical {
		case "", "bottom":
		case "center":
			out = append(out, "valign=middle")
		default:
			out = append(out, "valign="+a.Vertical)
		}
		if a.WrapText {
			out = append(out, "wrap")
		}
		if a.Indent > 0 {
			out = append(out, "indent="+strconv.Itoa(a.Indent))
		}
		if a.TextRotation != 0 {
			r := a.TextRotation
			if r > 90 {
				r = 90 - r
			}
			out = append(out, "rotate="+strconv.Itoa(r))
		}
	}
	sides := map[string]string{}
	for _, b := range s.Border {
		if b.Style > 0 && (b.Type == "top" || b.Type == "bottom" || b.Type == "left" || b.Type == "right") {
			sides[b.Type] = borderNames[b.Style] + ":" + hex(b.Color)
		}
	}
	if len(sides) == 4 && sides["top"] == sides["bottom"] && sides["top"] == sides["left"] && sides["top"] == sides["right"] {
		out = append(out, "border=all:"+sides["top"])
	} else {
		for _, side := range []string{"top", "bottom", "left", "right"} {
			if v, ok := sides[side]; ok {
				out = append(out, "border="+side+":"+v)
			}
		}
	}
	return out
}

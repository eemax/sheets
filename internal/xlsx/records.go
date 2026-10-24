package xlsx

import (
	"archive/zip"
	"bytes"
	"fmt"
	"regexp"
	"sort"
	"strconv"
	"strings"

	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/words"
	"github.com/xuri/excelize/v2"
)

// The record keeps the styling sheets applied on request (style, cond,
// and the freeze, gridlines and filter of view) that fmt would otherwise
// strip, so that fmt can lay it over the house style again. Each entry is a
// hidden defined name: its range follows rows and columns as they are
// inserted and deleted, by sheets and by Excel alike, and its comment says
// what was applied:
//
//	_sheets.000003  Sales!$A$1:$F$1  "style fill=lightblue bold"
const recordPrefix = "_sheets."

// maxRecordText is the longest comment Excel keeps on a name.
const maxRecordText = 255

// A Record is one entry of the record.
type Record struct {
	Seq   int    `json:"seq"`
	Kind  string `json:"kind"` // style, cond, view
	Sheet string `json:"sheet"`
	// Areas are the cells it applies to: one range, or for styling that
	// moved or was cleared in part, several (Excel keeps a name over
	// several areas, Sales!$A$1:$F$1,Sales!$A$3:$F$3).
	Areas []grid.Range `json:"-"`
	Ref   string       `json:"range"`
	Args  []string     `json:"args"`
	name  string
}

// Range is the first of the record's areas, the only one of a cond or view
// entry.
func (r Record) Range() grid.Range {
	if len(r.Areas) == 0 {
		return grid.Range{}
	}
	return r.Areas[0]
}

// Text is the record as a command line: kind and arguments.
func (r Record) Text() string { return words.Join(append([]string{r.Kind}, r.Args...)) }

// absRef writes a range for a defined name: Sales!$B$2:$D$9, Sales!$B:$D,
// Sales!$2:$2.
func absRef(sheet string, r grid.Range) string {
	q := grid.QuoteSheet(sheet) + "!"
	switch {
	case r.R1 == 1 && r.R2 == grid.MaxRows:
		return q + "$" + grid.ColName(r.C1) + ":$" + grid.ColName(r.C2)
	case r.C1 == 1 && r.C2 == grid.MaxCols:
		return q + "$" + strconv.Itoa(r.R1) + ":$" + strconv.Itoa(r.R2)
	}
	a := "$" + grid.ColName(r.C1) + "$" + strconv.Itoa(r.R1)
	if r.R1 == r.R2 && r.C1 == r.C2 {
		return q + a
	}
	return q + a + ":$" + grid.ColName(r.C2) + "$" + strconv.Itoa(r.R2)
}

// areasRef writes several areas of a sheet as one reference, joined by
// commas as Excel writes a union.
func areasRef(sheet string, areas []grid.Range) string {
	parts := make([]string, len(areas))
	for i, a := range areas {
		parts[i] = absRef(sheet, a)
	}
	return strings.Join(parts, ",")
}

// parseAbsRef reads what absRef writes; ok is false for a reference that
// lost its range (#REF!).
func parseAbsRef(s string) (string, grid.Range, bool) {
	s = strings.TrimPrefix(strings.TrimSpace(s), "=")
	i := strings.LastIndex(s, "!")
	if i < 0 || strings.Contains(s, "#REF!") {
		return "", grid.Range{}, false
	}
	sheet := s[:i]
	if len(sheet) >= 2 && sheet[0] == '\'' && sheet[len(sheet)-1] == '\'' {
		sheet = strings.ReplaceAll(sheet[1:len(sheet)-1], "''", "'")
	}
	r, err := grid.ParseRange(strings.ReplaceAll(s[i+1:], "$", ""))
	if err != nil || sheet == "" {
		return "", grid.Range{}, false
	}
	return sheet, r, true
}

// parseAreas reads what areasRef writes. Areas that lost their range
// (#REF!, after Excel deleted their rows) are left out; ok is false when
// none is left, or the areas name different sheets.
func parseAreas(s string) (string, []grid.Range, bool) {
	s = strings.TrimPrefix(strings.TrimSpace(s), "=")
	sheet := ""
	var areas []grid.Range
	for _, part := range splitUnion(s) {
		sh, r, ok := parseAbsRef(part)
		if !ok {
			continue
		}
		if sheet != "" && !strings.EqualFold(sh, sheet) {
			return "", nil, false
		}
		sheet = sh
		areas = append(areas, r)
	}
	return sheet, areas, len(areas) > 0
}

// splitUnion splits a reference at the commas between its areas, outside
// quoted sheet names.
func splitUnion(s string) []string {
	var out []string
	quoted, from := false, 0
	for i := 0; i < len(s); i++ {
		switch s[i] {
		case '\'':
			quoted = !quoted // '' inside a name toggles twice
		case ',':
			if !quoted {
				out = append(out, s[from:i])
				from = i + 1
			}
		}
	}
	return append(out, s[from:])
}

// refText shows a record's areas: Sales!A1:F1, or Sales!A1:F1,A3:F3.
func refText(sheet string, areas []grid.Range) string {
	if len(areas) == 0 {
		return ""
	}
	t := grid.FormatRef(sheet, areas[0])
	for _, a := range areas[1:] {
		t += "," + a.String()
	}
	return t
}

// readRecords returns the record in order, and the names of entries whose
// range is gone (their rows, columns or sheet were deleted).
func readRecords(f *excelize.File) (recs []Record, gone []string) {
	return recordsOf(f.GetDefinedName(), func(sheet string) bool {
		idx, err := f.GetSheetIndex(sheet)
		return err == nil && idx >= 0
	})
}

// rawRecords reads the record straight from a package, before excelize
// opens it.
func rawRecords(data []byte) (recs []Record, gone []string, err error) {
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, nil, err
	}
	p := &pkg{files: map[string]*zip.File{}}
	for _, zf := range zr.File {
		p.files[strings.ToLower(zf.Name)] = zf
	}
	wbPath, _, sheets, err := p.workbookParts()
	if err != nil {
		return nil, nil, err
	}
	var wbx xWorkbook
	if err := p.decode(wbPath, &wbx); err != nil {
		return nil, nil, err
	}
	var defs []excelize.DefinedName
	for _, n := range wbx.Names {
		defs = append(defs, excelize.DefinedName{Name: n.Name, Comment: n.Comment, RefersTo: strings.TrimSpace(n.Ref)})
	}
	recs, gone = recordsOf(defs, func(sheet string) bool {
		for name := range sheets {
			if strings.EqualFold(name, sheet) {
				return true
			}
		}
		return false
	})
	return recs, gone, nil
}

func recordsOf(defs []excelize.DefinedName, exists func(sheet string) bool) (recs []Record, gone []string) {
	for _, d := range defs {
		if !strings.HasPrefix(d.Name, recordPrefix) {
			continue
		}
		seq, err := strconv.Atoi(strings.TrimPrefix(d.Name, recordPrefix))
		ws, err2 := words.Texts(d.Comment)
		sheet, areas, ok := parseAreas(d.RefersTo)
		if err != nil || err2 != nil || len(ws) == 0 || !ok || !exists(sheet) || ws[0] != "style" && len(areas) > 1 {
			gone = append(gone, d.Name)
			continue
		}
		recs = append(recs, Record{Seq: seq, Kind: ws[0], Sheet: sheet, Areas: areas, Ref: refText(sheet, areas), Args: ws[1:], name: d.Name})
	}
	sort.Slice(recs, func(i, j int) bool { return recs[i].Seq < recs[j].Seq })
	return recs, gone
}

// dropRecords deletes entries by name.
func dropRecords(f *excelize.File, names ...string) error {
	for _, n := range names {
		if err := f.DeleteDefinedName(&excelize.DefinedName{Name: n}); err != nil && !strings.Contains(err.Error(), "no defined name") {
			return err
		}
	}
	return nil
}

// addRecord appends an entry, first dropping older ones that it makes
// pointless (see supersedes). It reports false, with nothing changed, when
// the entry is too long for Excel to keep.
func addRecord(f *excelize.File, rec Record) (bool, error) {
	text := rec.Text()
	if len(text) > maxRecordText {
		return false, nil
	}
	recs, gone := readRecords(f)
	if err := dropRecords(f, gone...); err != nil {
		return false, err
	}
	seq := 0
	for _, old := range recs {
		seq = max(seq, old.Seq)
		if supersedes(rec, old) {
			if err := dropRecords(f, old.name); err != nil {
				return false, err
			}
		}
	}
	return true, f.SetDefinedName(&excelize.DefinedName{
		Name:     fmt.Sprintf("%s%06d", recordPrefix, seq+1),
		Comment:  text,
		RefersTo: areasRef(rec.Sheet, rec.Areas),
	})
}

// setAreas points an entry at other areas of its sheet, keeping its place
// in the record; with none left, the entry goes.
func setAreas(f *excelize.File, rec Record, areas []grid.Range) error {
	if err := dropRecords(f, rec.name); err != nil {
		return err
	}
	if len(areas) == 0 {
		return nil
	}
	return f.SetDefinedName(&excelize.DefinedName{Name: rec.name, Comment: rec.Text(), RefersTo: areasRef(rec.Sheet, areas)})
}

// subtract returns the parts of r outside c, as up to four ranges: the
// rows above c, the rows below, and to the left and right of c beside it.
func subtract(r, c grid.Range) []grid.Range {
	x := r.Intersect(c)
	if x.Empty() {
		return []grid.Range{r}
	}
	var out []grid.Range
	if r.R1 < x.R1 {
		out = append(out, grid.Range{R1: r.R1, C1: r.C1, R2: x.R1 - 1, C2: r.C2})
	}
	if x.R2 < r.R2 {
		out = append(out, grid.Range{R1: x.R2 + 1, C1: r.C1, R2: r.R2, C2: r.C2})
	}
	if r.C1 < x.C1 {
		out = append(out, grid.Range{R1: x.R1, C1: r.C1, R2: x.R2, C2: x.C1 - 1})
	}
	if x.C2 < r.C2 {
		out = append(out, grid.Range{R1: x.R1, C1: x.C2 + 1, R2: x.R2, C2: r.C2})
	}
	return out
}

// supersedes tells whether a new entry makes an old one pointless: the same
// view setting on the same sheet, or styling that sets again everything the
// old one set, over all of its range.
func supersedes(rec, old Record) bool {
	if !strings.EqualFold(rec.Sheet, old.Sheet) || rec.Kind != old.Kind {
		return false
	}
	switch rec.Kind {
	case "view":
		return len(rec.Args) > 0 && len(old.Args) > 0 && viewKey(rec.Args[0]) == viewKey(old.Args[0])
	case "style":
		for _, oa := range old.Areas {
			covered := false
			for _, na := range rec.Areas {
				covered = covered || na.Contains(oa.R1, oa.C1) && na.Contains(oa.R2, oa.C2)
			}
			if !covered {
				return false
			}
		}
		nl, err1 := ParseLook(rec.Args)
		ol, err2 := ParseLook(old.Args)
		if err1 != nil || err2 != nil {
			return false
		}
		nk := nl.keys()
		for k := range ol.keys() {
			if !nk[k] && !(nl.Reset && isCellProperty(k)) {
				return false
			}
		}
		return true
	}
	return false
}

// isCellProperty tells whether a look key is a cell style property (which
// reset clears), rather than a width, height or merge.
func isCellProperty(k string) bool { return k != "width" && k != "height" && k != "merge" }

// viewKey groups view settings that replace each other.
// viewMoves tells whether a view entry's reference follows its cells as
// lines move: a filter's does; a freeze moves as its pane does, and the
// others (gridlines, no freeze, no filter) are not cells.
func viewMoves(args []string) bool { return len(args) > 0 && args[0] == "filter" }

func viewKey(arg string) string {
	k, _, _ := strings.Cut(arg, "=")
	return strings.TrimPrefix(k, "no")
}

// clip limits a recorded range to the used part of a sheet where it spans
// whole rows or columns.
func clip(r, used grid.Range) grid.Range {
	if r.R1 == 1 && r.R2 == grid.MaxRows {
		r.R1, r.R2 = max(used.R1, 1), max(used.R2, 1)
	}
	if r.C1 == 1 && r.C2 == grid.MaxCols {
		r.C1, r.C2 = max(used.C1, 1), max(used.C2, 1)
	}
	return r
}

var recordName = regexp.MustCompile(`<definedName\b[^>]*\bname="(?:` + regexp.QuoteMeta(recordPrefix) + `[^"]*|_xlnm\._FilterDatabase)"[^>]*>`)

// hideRecords marks the record's names hidden, which excelize cannot set:
// Excel then leaves them out of the Name Manager and the name box. A
// filter's name, which relines puts back, is hidden as Excel keeps it.
func hideRecords(wb []byte) []byte {
	return recordName.ReplaceAllFunc(wb, func(tag []byte) []byte {
		if strings.Contains(string(tag), ` hidden="`) {
			return tag
		}
		return []byte(strings.Replace(string(tag), "<definedName ", `<definedName hidden="1" `, 1))
	})
}

// condFormulas rewrites the formulas in a cond entry's words: a whole rule
// (=$C2="Late"), or values (> =$B$1).
func condFormulas(args []string, fn func(string) string) []string {
	out := make([]string, len(args))
	for i, a := range args {
		t := strings.TrimSpace(a)
		switch {
		case strings.HasPrefix(t, "=") && !strings.HasPrefix(t, "==") && !strings.HasPrefix(t, "= "):
			out[i] = "=" + fn(t[1:])
		case strings.Contains(a, " ="):
			fs := strings.Fields(a)
			for j, f := range fs {
				if strings.HasPrefix(f, "=") && len(f) > 1 && f[1] != '=' {
					fs[j] = "=" + fn(f[1:])
				}
			}
			out[i] = strings.Join(fs, " ")
		default:
			out[i] = a
		}
	}
	return out
}

// rewriteCondRecords keeps the formulas of recorded conditional formats in
// step with their rules, which fn rewrites (for moved cells or renamed
// sheets); home is the entry's sheet. A rule elsewhere re-made from its
// entry (copied, or moved to another sheet) then reads the right cells.
func (e *Editor) rewriteCondRecords(fn func(f, home string) string) error {
	for _, rec := range e.Records() {
		if rec.Kind != "cond" {
			continue
		}
		args := condFormulas(rec.Args, func(f string) string { return fn(f, rec.Sheet) })
		if words.Join(args) == words.Join(rec.Args) {
			continue
		}
		rec.Args = args
		if err := setAreas(e.f, rec, rec.Areas); err != nil {
			return err
		}
	}
	return nil
}

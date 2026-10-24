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
	Seq   int        `json:"seq"`
	Kind  string     `json:"kind"` // style, cond, view
	Sheet string     `json:"sheet"`
	Range grid.Range `json:"-"`
	Ref   string     `json:"range"`
	Args  []string   `json:"args"`
	name  string
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
		sheet, r, ok := parseAbsRef(d.RefersTo)
		if err != nil || err2 != nil || len(ws) == 0 || !ok || !exists(sheet) {
			gone = append(gone, d.Name)
			continue
		}
		recs = append(recs, Record{Seq: seq, Kind: ws[0], Sheet: sheet, Range: r, Ref: grid.FormatRef(sheet, r), Args: ws[1:], name: d.Name})
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
		RefersTo: absRef(rec.Sheet, rec.Range),
	})
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
		if !rec.Range.Contains(old.Range.R1, old.Range.C1) || !rec.Range.Contains(old.Range.R2, old.Range.C2) {
			return false
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

var recordName = regexp.MustCompile(`<definedName\b[^>]*\bname="` + regexp.QuoteMeta(recordPrefix) + `[^"]*"[^>]*>`)

// hideRecords marks the record's names hidden, which excelize cannot set:
// Excel then leaves them out of the Name Manager and the name box.
func hideRecords(wb []byte) []byte {
	return recordName.ReplaceAllFunc(wb, func(tag []byte) []byte {
		if strings.Contains(string(tag), ` hidden="`) {
			return tag
		}
		return []byte(strings.Replace(string(tag), "<definedName ", `<definedName hidden="1" `, 1))
	})
}

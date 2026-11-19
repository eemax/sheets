package fuzz

import (
	"fmt"
	"sort"
	"strings"

	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/xlsx"
)

// The objects of a sheet: what it holds besides cells. A merge, a table, a
// note, a link, a chart, the sheet's filter, a data validation, a conditional
// format, the rows and columns it hides, whether the sheet is hidden, and the
// look of each cell. The reach invariant judged only cells at first (values,
// formulas, number formats, defined names), so a slip that grew a table by
// three rows, merged cells the edit never named or dropped a note went by,
// though the code that moves, copies and shifts them is the code most likely
// to slip (fuzzer-long's audit, f2.1).
//
// An object outside the edit's reach must stand as it did, followed through
// the edit as Excel follows it (reach.place):
//
//   - a span (a merge, a table, a filter, the range of a validation or a
//     conditional format) goes with the lines inserted or deleted before it,
//     grows when lines are inserted inside it, and shrinks, or goes, when
//     lines are deleted from it;
//   - a cell's object (a note, a link, a chart's anchor) moves with its cell,
//     and goes when its row or column is deleted;
//   - what it holds that reads cells (a table's calculated column, a link's
//     place, a chart's series, a validation's or a conditional format's
//     rule) is rewritten as a formula is (reach.shifted), and the rest
//     letter for letter;
//   - an object on a sheet made as a copy is new, and not judged.
//
// An object that an edit's reach lies on may be changed by it (a set into a
// merge, a clear over a validation, a copy onto a note), but not a table, a
// filter, a chart, a hidden row or column, nor the sheet's visibility: no edit
// the fuzzer makes changes those but by shifting them. A script's own
// commands that are not modelled give no judgement (model.why), here as for
// cells.
type object struct {
	kind  string
	rg    grid.Range // where it stands: a span, or the one cell it is on
	cell  bool       // rg is the cell the object is on (it moves, or goes), not a span
	fixed string     // what it holds that no edit rewrites
	refs  string     // what it holds that reads cells (rewritten as a formula is)
	label string     // how a finding names it
	free  bool       // an edit whose reach lies on it may change it
	// soft is what it holds that an edit reaching the header row of its place
	// may change (a table's column names follow its header cells).
	soft string
	// relative says its refs read from the first cell of its place, as a
	// validation's and a conditional format's formulas do: when lines are
	// deleted from the front of its place, they read from the first cell
	// that stays, and the formula is not judged (place sets skipRefs).
	relative, skipRefs bool
	// anywhere says where it stands is not judged (a chart's anchor: the
	// editors disagree, see the README), so it is compared by what it holds.
	anywhere bool
}

// String is the object as a finding names it.
func (o object) String() string {
	s := o.kind
	if o.label != "" {
		s += " " + o.label
	}
	if o.cell {
		return s + " on " + grid.CellName(o.rg.R1, o.rg.C1)
	}
	return s + " at " + o.rg.String()
}

// gridObjects are the objects a sheet holds that the reader gives (the file's
// validations, conditional formats and looks are read by fileObjects).
func gridObjects(s *grid.Sheet) []object {
	var out []object
	for _, m := range s.Merges {
		out = append(out, object{kind: "merge", rg: m, free: true})
	}
	for _, t := range s.Tables {
		out = append(out, object{kind: "table", rg: t.Range, label: t.Name,
			fixed: fmt.Sprintf("%s totals=%d filter=%v", t.Name, t.Totals, t.Filter), soft: fmt.Sprintf("columns=%q", t.Columns),
			refs: strings.Join(t.Calculated, "\n")})
	}
	for _, n := range s.Notes {
		out = append(out, object{kind: "note", rg: grid.Range{R1: n.Row, C1: n.Col, R2: n.Row, C2: n.Col}, cell: true, free: true,
			label: fmt.Sprintf("%q", n.Text), fixed: fmt.Sprintf("%q %q %v", n.Author, n.Text, n.Thread)})
	}
	for _, l := range s.Links {
		out = append(out, object{kind: "link", rg: grid.Range{R1: l.Row, C1: l.Col, R2: l.Row, C2: l.Col}, cell: true, free: true,
			label: l.URL + l.Location, fixed: fmt.Sprintf("%q %q", l.URL, l.Tooltip), refs: l.Location})
	}
	for _, c := range s.Charts {
		var refs []string
		fixed := fmt.Sprintf("%s %q legend=%s %dx%d format=%q reversed=%v crossesMax=%v", c.Type, c.Title, c.Legend, c.Width, c.Height, c.ValueFormat, c.Reversed, c.CrossesMax)
		for _, se := range c.Series {
			refs = append(refs, se.Name, se.Categories, se.Values)
			fixed += fmt.Sprintf(" [%s %v %s %v]", se.Color, se.PointColors, se.Kind, se.Outline)
		}
		r, cc, _ := grid.ParseCell(c.Anchor)
		out = append(out, object{kind: "chart", rg: grid.Range{R1: r, C1: cc, R2: r, C2: cc}, cell: true, anywhere: true,
			label: c.Type, fixed: fixed, refs: strings.Join(refs, "\n")})
	}
	if f := s.Filter; f.R2 >= f.R1 && f.R1 >= 1 {
		out = append(out, object{kind: "filter", rg: f})
	}
	return out
}

// lines are the rows (or columns) a sheet hides, one by one; false when there
// are too many to follow (a sheet whose columns F onward are hidden).
func hiddenLines(runs []grid.Run) (map[int]bool, bool) {
	out := map[int]bool{}
	for _, r := range runs {
		if len(out)+r.To-r.From+1 > 20000 {
			return nil, false
		}
		for p := r.From; p <= r.To; p++ {
			out[p] = true
		}
	}
	return out, true
}

// compareObjects says how the objects two editors saved differ: the first
// difference, "" when they hold the same.
func compareObjects(a, b *grid.Sheet) string {
	list := func(s *grid.Sheet) []string {
		var out []string
		for _, o := range gridObjects(s) {
			at := o.rg.String()
			if o.anywhere {
				at = "somewhere"
			}
			out = append(out, fmt.Sprintf("%s %s %s %q %q %q", o.kind, at, o.label, o.fixed, o.soft, o.refs))
		}
		for _, side := range []struct {
			what string
			runs []grid.Run
		}{{"hidden rows", s.HiddenRows}, {"hidden columns", s.HiddenCols}} {
			out = append(out, fmt.Sprintf("%s %v", side.what, side.runs))
		}
		out = append(out, fmt.Sprintf("sheet hidden %v", s.Hidden))
		sort.Strings(out)
		return out
	}
	la, lb := list(a), list(b)
	for i := 0; i < len(la) || i < len(lb); i++ {
		switch {
		case i >= len(lb) || i < len(la) && la[i] < lb[i]:
			return "the streaming editor's file holds " + la[i] + ", the full editor's does not"
		case i >= len(la) || la[i] > lb[i]:
			return "the full editor's file holds " + lb[i] + ", the streaming editor's does not"
		}
	}
	return ""
}

// fileObjects are the objects of a file the reader does not give (grid has
// no validations, conditional formats or looks): read from the file with
// sheets' own readers, by sheet (lower case).
type fileObjects struct {
	objects map[string][]object
	looks   map[string]map[[2]int]string
}

// readFileObjects reads the validations, conditional formats and cell looks
// of a file. The looks are those of the cells in the sheet's used range, one
// string each (the properties sheets style would take).
func readFileObjects(data []byte, wb *grid.Workbook) (*fileObjects, error) {
	e, err := xlsx.OpenEditorData("book.xlsx", data)
	if err != nil {
		return nil, err
	}
	out := &fileObjects{objects: map[string][]object{}, looks: map[string]map[[2]int]string{}}
	for _, s := range wb.Sheets {
		k := strings.ToLower(s.Name)
		vs, err := e.Validations(s.Name)
		if err != nil {
			return nil, err
		}
		for _, v := range vs {
			for _, rg := range sqrefParts(v.Range) {
				out.objects[k] = append(out.objects[k], object{kind: "validation", rg: rg, free: true, relative: true, label: strings.Fields(v.Rule + " ")[0],
					fixed: fmt.Sprintf("%q %q %v", v.Prompt, v.Error, v.Warn), refs: v.Rule})
			}
		}
		cs, err := e.CondFormats(s.Name)
		if err != nil {
			return nil, err
		}
		for _, c := range cs {
			for _, rg := range sqrefParts(c.Range) {
				out.objects[k] = append(out.objects[k], object{kind: "conditional format", rg: rg, free: true, relative: true,
					fixed: fmt.Sprintf("%v %v", c.Recorded, c.Banding), refs: c.Rule})
			}
		}
		looks := map[[2]int]string{}
		if u := s.Used(); u.R2 >= 1 {
			runs, err := e.Looks(s.Name, u)
			if err != nil {
				return nil, err
			}
			for _, run := range runs {
				_, at, _ := cutLast(run.Range, "!")
				rg, err := grid.ParseRange(at)
				if err != nil {
					continue
				}
				look := strings.Join(run.Look, " ")
				for r := rg.R1; r <= rg.R2; r++ {
					for c := rg.C1; c <= rg.C2; c++ {
						looks[[2]int{r, c}] = look
					}
				}
			}
		}
		out.looks[k] = looks
	}
	return out, nil
}

func cutLast(s, sep string) (string, string, bool) {
	i := strings.LastIndex(s, sep)
	if i < 0 {
		return "", s, false
	}
	return s[:i], s[i+len(sep):], true
}

// sqrefParts reads the ranges of a validation's or a conditional format's sqref.
func sqrefParts(sqref string) []grid.Range {
	var out []grid.Range
	for _, p := range strings.Fields(strings.ReplaceAll(sqref, "$", "")) {
		if rg, err := grid.ParseRange(p); err == nil {
			out = append(out, rg)
		}
	}
	return out
}

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
// Two things are not judged, and are bugs the first long runs found (the
// README names them): a note's author (the full editor writes every note again
// with the author "Author", where the streaming editor leaves the file's own),
// and where a chart stands (the streaming editor leaves a chart where it was
// when rows are inserted or deleted above it, the full editor moves it with
// its cells). Judged, each ends every sequence that reaches it, at the first
// edit the full editor takes on a workbook with a note or a chart.
//
// A frozen pane (the rows and columns kept in view) is judged by where its
// split stands (reach.judgePanes): lines inserted or deleted among the frozen
// ones move it, as Excel moves it, and no other edit does (the house style
// freezes the header of every sheet fmt touches, and every rows or columns
// edit moves a split: an off-by-one in the code that does, left the frozen
// header showing the row inserted above it, with no pool file holding a pane:
// fuzzer-long's audit, f4.3).
//
// An object that an edit's reach lies on may be changed by it (a set into a
// merge, a clear --all over a validation, a copy onto a note), but not a
// table, a filter, a chart, a hidden row or column, nor the sheet's
// visibility: no edit the fuzzer makes changes those but by shifting them.
// The reach of a clear without --all, and of the target of copy --values, is
// the cells' contents alone (reach.touchContents): the notes, links, merges,
// validations, conditional formats and the look of its cells stay, and are
// judged as outside the reach, as sheets' own help says (fuzzer-long's audit,
// f4.2: a clear that acted as --all, dropping them, went by). A script's own
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
	// vanishes says it goes when a shift cuts it to one cell (a merge).
	vanishes bool
	// cover says sheets may write its cells as other ranges (a validation's
	// sqref cut in two): the cells it covers are judged, not its ranges.
	cover bool
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
		out = append(out, object{kind: "merge", rg: m, free: true, vanishes: true})
	}
	for _, t := range s.Tables {
		out = append(out, object{kind: "table", rg: t.Range, label: t.Name,
			fixed: fmt.Sprintf("%s totals=%d filter=%v", t.Name, t.Totals, t.Filter), soft: fmt.Sprintf("columns=%q", t.Columns),
			refs: strings.Join(t.Calculated, "\n")})
	}
	for _, n := range s.Notes {
		out = append(out, object{kind: "note", rg: grid.Range{R1: n.Row, C1: n.Col, R2: n.Row, C2: n.Col}, cell: true, free: true,
			label: fmt.Sprintf("%q", n.Text), fixed: fmt.Sprintf("%q %v", n.Text, n.Thread)})
	}
	for _, l := range s.Links {
		out = append(out, object{kind: "link", rg: grid.Range{R1: l.Row, C1: l.Col, R2: l.Row, C2: l.Col}, cell: true, free: true,
			label: l.URL + l.Location, fixed: fmt.Sprintf("%q %q %q", l.URL, l.Location, l.Tooltip)})
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

// hiddenLines are the rows (or columns) a sheet hides, one by one; false when
// there are too many to follow (a sheet whose columns F onward are hidden:
// those are not judged).
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
	// panes are the frozen rows and columns of each sheet (the rows kept in
	// view, then the columns; none for a sheet that freezes nothing).
	panes map[string][2]int
}

// readFileObjects reads the validations, conditional formats and cell looks
// of a file. The looks are those of the cells in the sheet's used range, one
// string each (the properties sheets style would take).
func readFileObjects(data []byte, wb *grid.Workbook) (*fileObjects, error) {
	e, err := xlsx.OpenEditorData("book.xlsx", data)
	if err != nil {
		return nil, err
	}
	out := &fileObjects{objects: map[string][]object{}, looks: map[string]map[[2]int]string{}, panes: map[string][2]int{}}
	for _, s := range wb.Sheets {
		k := strings.ToLower(s.Name)
		view, err := e.ViewOf(s.Name)
		if err != nil {
			return nil, err
		}
		if r, c, ok := grid.ParseCell(view.Freeze); ok {
			out.panes[k] = [2]int{r - 1, c - 1} // Freeze is the first cell that scrolls
		}
		vs, err := e.Validations(s.Name)
		if err != nil {
			return nil, err
		}
		for _, v := range vs {
			for _, rg := range sqrefParts(v.Range) {
				out.objects[k] = append(out.objects[k], object{kind: "validation", rg: rg, free: true, relative: true, cover: true, label: strings.Fields(v.Rule + " ")[0],
					fixed: fmt.Sprintf("%q %q %v", v.Prompt, v.Error, v.Warn), refs: v.Rule})
			}
		}
		cs, err := e.CondFormats(s.Name)
		if err != nil {
			return nil, err
		}
		for _, c := range cs {
			for _, rg := range sqrefParts(c.Range) {
				out.objects[k] = append(out.objects[k], object{kind: "conditional format", rg: rg, free: true, relative: true, cover: true,
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

// paneName says how many rows and columns a pane freezes: "2 rows and 1
// column", "no rows or columns" for none.
func paneName(p [2]int) string {
	plural := func(n int, what string) string {
		if n == 1 {
			return "1 " + what
		}
		return fmt.Sprintf("%d %ss", n, what)
	}
	switch {
	case p[0] == 0 && p[1] == 0:
		return "no rows or columns"
	case p[1] == 0:
		return plural(p[0], "row")
	case p[0] == 0:
		return plural(p[1], "column")
	}
	return plural(p[0], "row") + " and " + plural(p[1], "column")
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

// compareFileObjects says how the validations, conditional formats and looks
// two editors saved differ: the first difference, "" when they hold the same.
// A rule's cells are compared, not its ranges (see object.cover), and the
// look of each cell one by one.
func compareFileObjects(stream, full *fileObjects, name func(k string) string) string {
	var keys []string
	seen := map[string]bool{}
	for _, f := range []*fileObjects{stream, full} {
		for k := range f.objects {
			if !seen[k] {
				seen[k] = true
				keys = append(keys, k)
			}
		}
		for k := range f.looks {
			if !seen[k] {
				seen[k] = true
				keys = append(keys, k)
			}
		}
		for k := range f.panes {
			if !seen[k] {
				seen[k] = true
				keys = append(keys, k)
			}
		}
	}
	sort.Strings(keys)
	for _, k := range keys {
		if a, b := stream.panes[k], full.panes[k]; a != b {
			return fmt.Sprintf("%s: the streaming editor's file freezes %s, the full editor's %s", name(k), paneName(a), paneName(b))
		}
		group := func(f *fileObjects) map[string][]grid.Range {
			out := map[string][]grid.Range{}
			for _, o := range f.objects[k] {
				key := o.kind + "\x00" + o.refs + "\x00" + o.fixed
				out[key] = append(out[key], o.rg)
			}
			return out
		}
		a, b := group(stream), group(full)
		var gs []string
		for g := range a {
			gs = append(gs, g)
		}
		for g := range b {
			if _, ok := a[g]; !ok {
				gs = append(gs, g)
			}
		}
		sort.Strings(gs)
		for _, g := range gs {
			if at, ok := sameCover(a[g], b[g], nil); !ok {
				kind, rest, _ := strings.Cut(g, "\x00")
				rule, _, _ := strings.Cut(rest, "\x00")
				return fmt.Sprintf("%s!%s %q: the streaming editor's file holds it over %s, the full editor's over %s (they differ at %s)",
					name(k), kind, rule, listRanges(a[g]), listRanges(b[g]), grid.CellName(at[0], at[1]))
			}
		}
		la, lb := stream.looks[k], full.looks[k]
		var ps [][2]int
		for p := range la {
			ps = append(ps, p)
		}
		for p := range lb {
			if _, ok := la[p]; !ok {
				ps = append(ps, p)
			}
		}
		sort.Slice(ps, func(i, j int) bool {
			if ps[i][0] != ps[j][0] {
				return ps[i][0] < ps[j][0]
			}
			return ps[i][1] < ps[j][1]
		})
		for _, p := range ps {
			if la[p] != lb[p] {
				return fmt.Sprintf("%s!%s: the streaming editor's file shows it as [%s], the full editor's as [%s]", name(k), grid.CellName(p[0], p[1]), la[p], lb[p])
			}
		}
	}
	return ""
}

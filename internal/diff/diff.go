// Package diff compares two versions of a workbook: sheets added, removed
// and renamed, rows and columns inserted and deleted, and the cells that
// changed (formulas, values, number formats). Rows and columns are aligned
// first, so an inserted row does not make every row below it look changed.
package diff

import (
	"fmt"
	"hash/fnv"
	"math"
	"sort"
	"strconv"
	"strings"

	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
)

// Result is the difference between two workbooks.
type Result struct {
	Sheets []Sheet      `json:"sheets"`
	Names  []NameChange `json:"names,omitempty"`
	Order  *OrderChange `json:"order,omitempty"`
	Same   bool         `json:"same"`
}

// Sheet is what changed on one sheet.
type Sheet struct {
	Name        string         `json:"name"`   // as in the new workbook; the old name when removed
	Status      string         `json:"status"` // added, removed, changed
	RenamedFrom string         `json:"renamed_from,omitempty"`
	Hidden      string         `json:"hidden,omitempty"` // "hidden" or "shown" when that changed
	Used        string         `json:"used,omitempty"`   // used range of an added or removed sheet
	RowsAdded   []Lines        `json:"rows_added,omitempty"`
	RowsRemoved []Lines        `json:"rows_removed,omitempty"`
	ColsAdded   []Lines        `json:"cols_added,omitempty"`
	ColsRemoved []Lines        `json:"cols_removed,omitempty"`
	Cells       []Change       `json:"cells,omitempty"`
	Formats     []FormatChange `json:"formats,omitempty"`
	MoreCells   int            `json:"more_cells,omitempty"`
}

// Lines is a run of added or removed rows or columns, with what they held.
type Lines struct {
	From    int    `json:"from"` // row number, or column index
	To      int    `json:"to"`
	Content string `json:"content,omitempty"` // first line's values, for rows
}

// Change is one cell that changed.
type Change struct {
	Cell    string `json:"cell"`               // in the new workbook
	OldCell string `json:"old_cell,omitempty"` // when the cell moved
	Old     string `json:"old"`                // "" when empty; formulas start with =, text is quoted
	New     string `json:"new"`
	Note    string `json:"note,omitempty"` // "result of =..." when only a formula's value changed
}

// FormatChange is a run of cells in one column whose number format changed.
type FormatChange struct {
	Range string `json:"range"`
	Old   string `json:"old"`
	New   string `json:"new"`
}

// NameChange is a defined name added, removed or redefined.
type NameChange struct {
	Name string `json:"name"`
	Old  string `json:"old,omitempty"`
	New  string `json:"new,omitempty"`
}

// OrderChange records a new sheet order.
type OrderChange struct {
	Old []string `json:"old"`
	New []string `json:"new"`
}

// Options limit the result.
type Options struct {
	MaxCells int // cell changes listed per sheet; 0 means all
}

// Workbooks compares two versions of a workbook.
func Workbooks(a, b *grid.Workbook, opts Options) Result {
	res := Result{}
	matchA := map[*grid.Sheet]*grid.Sheet{}
	usedB := map[*grid.Sheet]bool{}
	for _, sa := range a.Sheets {
		if sb := b.Sheet(sa.Name); sb != nil && strings.EqualFold(sb.Name, sa.Name) {
			matchA[sa] = sb
			usedB[sb] = true
		}
	}
	// Renamed sheets: an unmatched old sheet whose rows mostly reappear in
	// an unmatched new one.
	for _, sa := range a.Sheets {
		if matchA[sa] != nil {
			continue
		}
		best, score, samePlace := (*grid.Sheet)(nil), 0.0, false
		for _, sb := range b.Sheets {
			if usedB[sb] {
				continue
			}
			if sc := rowOverlap(a, sa, b, sb); sc > score {
				best, score = sb, sc
				samePlace = indexOf(a, sa) == indexOf(b, sb)
			}
		}
		// Most rows in common; half will do for a sheet in the same place.
		if best != nil && (score >= 0.8 || samePlace && score >= 0.5) {
			matchA[sa] = best
			usedB[best] = true
		}
	}

	for _, sb := range b.Sheets {
		var sa *grid.Sheet
		for k, v := range matchA {
			if v == sb {
				sa = k
			}
		}
		if sa == nil {
			res.Sheets = append(res.Sheets, Sheet{Name: sb.Name, Status: "added", Used: usedString(sb)})
			continue
		}
		d := sheet(a, sa, b, sb, opts)
		if sa.Name != sb.Name {
			d.RenamedFrom = sa.Name
		}
		if sa.Hidden != sb.Hidden {
			d.Hidden = map[bool]string{true: "hidden", false: "shown"}[sb.Hidden]
		}
		if d.RenamedFrom != "" || d.Hidden != "" || !d.empty() {
			d.Status = "changed"
			res.Sheets = append(res.Sheets, d)
		}
	}
	for _, sa := range a.Sheets {
		if matchA[sa] == nil {
			res.Sheets = append(res.Sheets, Sheet{Name: sa.Name, Status: "removed", Used: usedString(sa)})
		}
	}

	res.Names = names(a, b)
	oa, ob := commonOrder(a, b, matchA)
	if strings.Join(oa, "\x00") != strings.Join(ob, "\x00") {
		res.Order = &OrderChange{Old: a.SheetNames(), New: b.SheetNames()}
	}
	res.Same = len(res.Sheets) == 0 && len(res.Names) == 0 && res.Order == nil
	if res.Sheets == nil {
		res.Sheets = []Sheet{}
	}
	return res
}

func indexOf(wb *grid.Workbook, s *grid.Sheet) int {
	for i, x := range wb.Sheets {
		if x == s {
			return i
		}
	}
	return -1
}

func (d Sheet) empty() bool {
	return len(d.RowsAdded)+len(d.RowsRemoved)+len(d.ColsAdded)+len(d.ColsRemoved)+len(d.Cells)+len(d.Formats)+d.MoreCells == 0
}

func usedString(s *grid.Sheet) string {
	if u := s.Used(); !u.Empty() {
		return u.String()
	}
	return ""
}

// commonOrder lists matched sheets in each workbook's order, by new name.
func commonOrder(a, b *grid.Workbook, match map[*grid.Sheet]*grid.Sheet) ([]string, []string) {
	var oa, ob []string
	in := map[*grid.Sheet]bool{}
	for _, sa := range a.Sheets {
		if sb := match[sa]; sb != nil {
			oa = append(oa, sb.Name)
			in[sb] = true
		}
	}
	for _, sb := range b.Sheets {
		if in[sb] {
			ob = append(ob, sb.Name)
		}
	}
	return oa, ob
}

func names(a, b *grid.Workbook) []NameChange {
	key := func(n grid.DefinedName) string { return strings.ToLower(n.Scope + "!" + n.Name) }
	old := map[string]grid.DefinedName{}
	for _, n := range a.Names {
		old[key(n)] = n
	}
	var out []NameChange
	seen := map[string]bool{}
	for _, n := range b.Names {
		k := key(n)
		seen[k] = true
		if o, ok := old[k]; !ok {
			out = append(out, NameChange{Name: n.Name, New: n.Ref})
		} else if o.Ref != n.Ref {
			out = append(out, NameChange{Name: n.Name, Old: o.Ref, New: n.Ref})
		}
	}
	for _, n := range a.Names {
		if !seen[key(n)] {
			out = append(out, NameChange{Name: n.Name, Old: n.Ref})
		}
	}
	return out
}

// side is one version of a sheet being compared.
type side struct {
	wb *grid.Workbook
	s  *grid.Sheet
}

// key describes a cell for alignment: constants by value, formulas by their
// pattern relative to the cell (so a formula moved with its row matches).
func (x side) key(row, col int) string {
	c := x.s.At(row, col)
	if c == nil || c.IsEmpty() {
		return ""
	}
	if c.Formula != "" {
		k, _ := formula.Pattern(c.Formula, row, col)
		return "=" + k
	}
	return strconv.Itoa(int(c.Kind)) + ":" + grid.Display(c, grid.Format{}, false)
}

func hash(parts ...string) uint64 {
	h := fnv.New64a()
	for _, p := range parts {
		h.Write([]byte(p))
		h.Write([]byte{0})
	}
	return h.Sum64()
}

// alike reports two formulas doing the same kind of thing (both a SUM,
// say), so a total whose range grew still pairs with its old self.
func alike(A side, ra, ca int, B side, rb, cb int) bool {
	x, y := A.s.At(ra, ca), B.s.At(rb, cb)
	if x == nil || y == nil || x.Formula == "" || y.Formula == "" {
		return false
	}
	fx, fy := formula.Calls(x.Formula), formula.Calls(y.Formula)
	return len(fx) > 0 && len(fy) > 0 && fx[0] == fy[0]
}

// rowOverlap is the share of rows two sheets have in common, for spotting
// a renamed sheet. Rows are compared by their contents regardless of
// column, so a column inserted at the same time does not hide the match.
func rowOverlap(wa *grid.Workbook, sa *grid.Sheet, wb *grid.Workbook, sb *grid.Sheet) float64 {
	ua, ub := sa.Used(), sb.Used()
	if ua.Empty() || ub.Empty() {
		return 0
	}
	rows := func(x side, u grid.Range) map[uint64]int {
		m := map[uint64]int{}
		for r := u.R1; r <= u.R2; r++ {
			var parts []string
			for c := u.C1; c <= u.C2; c++ {
				if k := x.key(r, c); k != "" {
					parts = append(parts, k)
				}
			}
			if len(parts) == 0 {
				continue // blank rows say nothing
			}
			sort.Strings(parts)
			m[hash(parts...)]++
		}
		return m
	}
	ma, mb := rows(side{wa, sa}, ua), rows(side{wb, sb}, ub)
	common, na, nb := 0, 0, 0
	for h, n := range ma {
		common += min(n, mb[h])
		na += n
	}
	for _, n := range mb {
		nb += n
	}
	// Measured against the smaller sheet, so rows added or changed since
	// do not hide that most of the old sheet is still there.
	if smaller := min(na, nb); smaller > 0 {
		return float64(common) / float64(smaller)
	}
	return 0
}

func sheet(wa *grid.Workbook, sa *grid.Sheet, wb *grid.Workbook, sb *grid.Sheet, opts Options) Sheet {
	d := Sheet{Name: sb.Name}
	A, B := side{wa, sa}, side{wb, sb}
	ua, ub := sa.Used(), sb.Used()
	ra, rb := max(ua.R2, 0), max(ub.R2, 0)
	ca, cb := max(ua.C2, 0), max(ub.C2, 0)

	// Columns: aligned only when the count changed; otherwise they pair up
	// in place.
	var colSteps []step
	if ca == cb {
		for c := 0; c < ca; c++ {
			colSteps = append(colSteps, step{c, c})
		}
	} else {
		colSteps = alignColumns(A, B, ra, rb, ca, cb)
	}
	var pairs [][2]int // matched columns, 1-based
	for _, st := range colSteps {
		switch {
		case st.a >= 0 && st.b >= 0:
			pairs = append(pairs, [2]int{st.a + 1, st.b + 1})
		case st.b >= 0:
			d.ColsAdded = addLine(d.ColsAdded, st.b+1, "")
		default:
			d.ColsRemoved = addLine(d.ColsRemoved, st.a+1, "")
		}
	}

	// Rows, over the matched columns.
	sig := func(x side, rows int, pick int) []uint64 {
		out := make([]uint64, rows)
		parts := make([]string, len(pairs))
		for r := 1; r <= rows; r++ {
			for k, p := range pairs {
				parts[k] = x.key(r, p[pick])
			}
			out[r-1] = hash(parts...)
		}
		return out
	}
	sa1, sb1 := sig(A, ra, 0), sig(B, rb, 1)
	similar := func(i, j int) bool {
		same, used := 0, 0
		for _, p := range pairs {
			ka, kb := A.key(i+1, p[0]), B.key(j+1, p[1])
			if ka == "" && kb == "" {
				continue
			}
			used++
			if ka == kb || alike(A, i+1, p[0], B, j+1, p[1]) {
				same++
			}
		}
		return used > 0 && same*2 >= used
	}
	rowSteps := pairUp(ra, rb, align(sa1, sb1), similar)

	type colRun struct {
		r1, r2   int
		old, new string
	}
	runs := map[int]*colRun{}
	flushRun := func(c int) {
		if run := runs[c]; run != nil {
			rg := grid.Range{R1: run.r1, C1: c, R2: run.r2, C2: c}
			d.Formats = append(d.Formats, FormatChange{Range: rg.String(), Old: run.old, New: run.new})
			delete(runs, c)
		}
	}
	listed := 0
	for _, st := range rowSteps {
		switch {
		case st.a < 0:
			d.RowsAdded = addLine(d.RowsAdded, st.b+1, rowContent(B, st.b+1, cb))
			continue
		case st.b < 0:
			d.RowsRemoved = addLine(d.RowsRemoved, st.a+1, rowContent(A, st.a+1, ca))
			continue
		}
		oldRow, newRow := st.a+1, st.b+1
		for _, p := range pairs {
			ch, fmtOld, fmtNew, changed := compare(A, oldRow, p[0], B, newRow, p[1])
			if changed {
				if opts.MaxCells > 0 && listed >= opts.MaxCells {
					d.MoreCells++
				} else {
					d.Cells = append(d.Cells, ch)
					listed++
				}
			}
			if fmtOld != fmtNew {
				if run := runs[p[1]]; run != nil && run.r2 == newRow-1 && run.old == fmtOld && run.new == fmtNew {
					run.r2 = newRow
				} else {
					flushRun(p[1])
					runs[p[1]] = &colRun{newRow, newRow, fmtOld, fmtNew}
				}
			} else {
				flushRun(p[1])
			}
		}
	}
	cols := make([]int, 0, len(runs))
	for c := range runs {
		cols = append(cols, c)
	}
	sort.Ints(cols)
	for _, c := range cols {
		flushRun(c)
	}
	sort.SliceStable(d.Formats, func(i, j int) bool {
		ri, _ := grid.ParseRange(d.Formats[i].Range)
		rj, _ := grid.ParseRange(d.Formats[j].Range)
		if ri.C1 != rj.C1 {
			return ri.C1 < rj.C1
		}
		return ri.R1 < rj.R1
	})
	return d
}

// alignColumns aligns columns by their first values (usually the header),
// which inserting rows further down does not disturb.
func alignColumns(A, B side, ra, rb, ca, cb int) []step {
	head := func(x side, rows, c int) []string {
		var vals []string
		for r := 1; r <= rows && len(vals) < 20; r++ {
			if k := x.key(r, c); k != "" {
				vals = append(vals, k)
			}
		}
		return vals
	}
	ha := make([][]string, ca)
	hb := make([][]string, cb)
	sa := make([]uint64, ca)
	sb := make([]uint64, cb)
	for c := 0; c < ca; c++ {
		ha[c] = head(A, ra, c+1)
		sa[c] = hash(ha[c][:min(3, len(ha[c]))]...)
	}
	for c := 0; c < cb; c++ {
		hb[c] = head(B, rb, c+1)
		sb[c] = hash(hb[c][:min(3, len(hb[c]))]...)
	}
	similar := func(i, j int) bool {
		count := map[string]int{}
		for _, v := range ha[i] {
			count[v]++
		}
		same := 0
		for _, v := range hb[j] {
			if count[v] > 0 {
				count[v]--
				same++
			}
		}
		n := max(len(ha[i]), len(hb[j]))
		return n > 0 && same*2 >= n
	}
	return pairUp(ca, cb, align(sa, sb), similar)
}

func addLine(list []Lines, n int, content string) []Lines {
	if k := len(list); k > 0 && list[k-1].To == n-1 {
		list[k-1].To = n
		return list
	}
	return append(list, Lines{From: n, To: n, Content: content})
}

func rowContent(x side, row, cols int) string {
	var parts []string
	last := 0
	for c := 1; c <= cols; c++ {
		r := repr(x, row, c)
		parts = append(parts, r)
		if r != "" {
			last = c
		}
	}
	return grid.Truncate(strings.Join(parts[:last], ", "), 120)
}

// repr shows a cell for a reader: formulas with =, text quoted.
func repr(x side, row, col int) string {
	c := x.s.At(row, col)
	if c == nil || c.IsEmpty() {
		return ""
	}
	if c.Formula != "" {
		return "=" + c.Formula
	}
	return value(x, row, col)
}

func value(x side, row, col int) string {
	c := x.wb.Value(x.s, row, col)
	if c == nil {
		return ""
	}
	if c.Pending {
		return "?"
	}
	v := grid.Display(c, x.wb.FormatOf(c), x.wb.Date1904)
	if c.Kind == grid.Text {
		return strconv.Quote(v)
	}
	return v
}

func formatName(x side, row, col int) string {
	c := x.s.At(row, col)
	if c == nil || c.IsEmpty() || c.Kind != grid.Number {
		return ""
	}
	f := x.wb.FormatOf(c)
	name := f.Kind.String()
	if f.Kind == grid.Currency && f.Symbol != "" {
		name += " " + f.Symbol
	}
	return name
}

// compare reports whether a cell changed, and the number formats on both
// sides (compared separately: a cell can keep its value and change format).
func compare(A side, ra, ca int, B side, rb, cb int) (Change, string, string, bool) {
	ch := Change{Cell: grid.CellName(rb, cb)}
	if ra != rb || ca != cb {
		ch.OldCell = grid.CellName(ra, ca)
	}
	fa, fb := formatName(A, ra, ca), formatName(B, rb, cb)
	if fa == "" || fb == "" {
		fa, fb = "", "" // a format only matters on numbers present on both sides
	}
	x, y := A.s.At(ra, ca), B.s.At(rb, cb)
	xe, ye := x == nil || x.IsEmpty(), y == nil || y.IsEmpty()
	switch {
	case xe && ye:
		return ch, fa, fb, false
	case xe || ye || (x.Formula == "") != (y.Formula == ""):
		ch.Old, ch.New = repr(A, ra, ca), repr(B, rb, cb)
		return ch, fa, fb, true
	case x.Formula != "":
		ka, _ := formula.Pattern(x.Formula, ra, ca)
		kb, _ := formula.Pattern(y.Formula, rb, cb)
		if ka != kb {
			ch.Old, ch.New = "="+x.Formula, "="+y.Formula
			return ch, fa, fb, true
		}
		va, vb := value(A, ra, ca), value(B, rb, cb)
		if va != vb && va != "?" && vb != "?" && !sameNumber(A, ra, ca, B, rb, cb) {
			ch.Old, ch.New, ch.Note = va, vb, "result of ="+y.Formula
			return ch, fa, fb, true
		}
		return ch, fa, fb, false
	}
	if sameNumber(A, ra, ca, B, rb, cb) {
		return ch, fa, fb, false
	}
	va, vb := value(A, ra, ca), value(B, rb, cb)
	if va == vb && x.Kind == y.Kind {
		return ch, fa, fb, false
	}
	ch.Old, ch.New = va, vb
	return ch, fa, fb, true
}

func sameNumber(A side, ra, ca int, B side, rb, cb int) bool {
	x, y := A.wb.Value(A.s, ra, ca), B.wb.Value(B.s, rb, cb)
	if x == nil || y == nil || x.Kind != grid.Number || y.Kind != grid.Number || x.Pending || y.Pending {
		return false
	}
	return math.Abs(x.Num-y.Num) <= 1e-9*math.Max(1, math.Max(math.Abs(x.Num), math.Abs(y.Num)))
}

// Summary is a one-line account of a sheet's changes.
func (d Sheet) Summary() string {
	var parts []string
	count := func(ls []Lines, word string, added bool) {
		n := 0
		for _, l := range ls {
			n += l.To - l.From + 1
		}
		if n == 0 {
			return
		}
		verb := map[bool]string{true: "added", false: "removed"}[added]
		parts = append(parts, fmt.Sprintf("%s %s", plural(n, word), verb))
	}
	count(d.RowsAdded, "row", true)
	count(d.RowsRemoved, "row", false)
	count(d.ColsAdded, "column", true)
	count(d.ColsRemoved, "column", false)
	if n := len(d.Cells) + d.MoreCells; n > 0 {
		parts = append(parts, plural(n, "cell")+" changed")
	}
	if n := len(d.Formats); n > 0 {
		parts = append(parts, plural(n, "number format change"))
	}
	if d.RenamedFrom != "" {
		parts = append([]string{"renamed from " + grid.QuoteSheet(d.RenamedFrom)}, parts...)
	}
	if d.Hidden != "" {
		parts = append(parts, d.Hidden)
	}
	return strings.Join(parts, ", ")
}

func plural(n int, word string) string {
	if n == 1 {
		return "1 " + word
	}
	return strconv.Itoa(n) + " " + word + "s"
}

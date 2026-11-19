package fuzz

import (
	"fmt"
	"sort"
	"strings"

	"github.com/eemax/sheets/internal/grid"
)

// follow is where a sheet (lower case) stands after the edit's renames.
func (rc *reach) follow(k string) string {
	for _, x := range rc.xforms {
		if x.kind == 'r' && k == x.sheet {
			k = x.to
		}
	}
	return k
}

// place follows an object through the edit's shifts of lines: the sheet it is
// on and where it stands now; false when the lines it stood on were deleted.
func (rc *reach) place(k string, o object) (string, object, bool) {
	for _, x := range rc.xforms {
		switch x.kind {
		case 'r':
			if k == x.sheet {
				k = x.to
			}
		case 'l':
			if k != x.sheet || o.anywhere {
				continue
			}
			lo, hi := &o.rg.R1, &o.rg.R2
			limit := grid.MaxRows
			if x.cols {
				lo, hi, limit = &o.rg.C1, &o.rg.C2, grid.MaxCols
			}
			if o.cell {
				q, ok := shiftLine(*lo, x.at, x.end, x.n)
				if !ok {
					return k, o, false
				}
				*lo, *hi = q, q
				continue
			}
			a, b, ok := shiftSpan(*lo, *hi, x.at, x.end, x.n)
			if !ok {
				return k, o, false
			}
			if o.relative && x.n == 0 && x.at <= *lo && *lo <= x.end {
				o.skipRefs = true
			}
			*lo, *hi = a, min(b, limit)
		}
	}
	return k, o, true
}

// hits tells whether the cells an edit reaches (not the lines it inserts) lie
// on a range of a sheet (lower case, as it stands now).
func (rc *reach) hits(k string, rg grid.Range) bool {
	for _, r := range rc.solid[k] {
		if !r.Intersect(rg).Empty() {
			return true
		}
	}
	return false
}

// sameRefs tells whether what an object read before reads, after the edit's
// shifts, what it does now (see shifted).
func (rc *reach) sameRefs(before, after, home, now string, display func(string) string) (string, bool) {
	if before == "" && after == "" {
		return "", true
	}
	return rc.shifted(before, after, home, now, display)
}

// judgeObjects judges the objects of the workbook before the edit (by sheet,
// lower case) against those after it: each outside the reach is still there,
// where the edit's shifts put it and holding what it held, and none is made
// out of nothing. exists says which sheets the workbook holds now. It returns
// the first difference, "" when there is none.
func (rc *reach) judgeObjects(before, after map[string][]object, exists func(k string) bool, display func(string) string) string {
	type want struct {
		from string // the sheet it was on
		o    object
	}
	expected := map[string][]want{}
	// reached are the places of objects an edit's reach lies on, where they
	// stand now: what is left of them (a validation cut in two) is theirs.
	reached := map[string][]grid.Range{}
	var keys []string
	for k0 := range before {
		keys = append(keys, k0)
	}
	sort.Strings(keys)
	for _, k0 := range keys {
		for _, o := range before[k0] {
			k, p, ok := rc.place(k0, o)
			switch {
			case !ok, rc.copies[k], !exists(k):
				continue
			case o.vanishes && p.rg.R1 == p.rg.R2 && p.rg.C1 == p.rg.C2 && !(o.rg.R1 == o.rg.R2 && o.rg.C1 == o.rg.C2):
				continue // cut to one cell: it goes (Excel keeps no merge of one cell)
			case o.free && rc.hits(k, p.rg):
				reached[k] = append(reached[k], p.rg)
				continue // the edit reaches it: it may change it
			}
			expected[k] = append(expected[k], want{k0, p})
		}
	}
	var ks []string
	for k := range expected {
		ks = append(ks, k)
	}
	for k := range after {
		if _, ok := expected[k]; !ok {
			ks = append(ks, k)
		}
	}
	sort.Strings(ks)
	for _, k := range ks {
		if rc.copies[k] {
			continue
		}
		// The objects now, each taken once by the one it matches.
		used := make([]bool, len(after[k]))
		// Objects whose cells are judged, not their ranges: by their group.
		type group struct {
			from string
			o    object
			want []grid.Range // the cells the shifts give
			got  []grid.Range // the cells the sheet holds
		}
		var groups []*group
		for _, w := range expected[k] {
			if w.o.cover {
				var g *group
				for _, h := range groups {
					if h.o.kind == w.o.kind && h.o.fixed == w.o.fixed && h.o.refs == w.o.refs && h.from == w.from {
						g = h
					}
				}
				if g == nil {
					g = &group{from: w.from, o: w.o}
					groups = append(groups, g)
				}
				g.want = append(g.want, w.o.rg)
				continue
			}
			found := false
			for i, a := range after[k] {
				if used[i] || a.cover || a.kind != w.o.kind || a.fixed != w.o.fixed {
					continue
				}
				if a.cell != w.o.cell {
					continue
				}
				if a.rg != w.o.rg && !w.o.anywhere {
					continue
				}
				if a.soft != w.o.soft && !rc.hits(k, grid.Range{R1: w.o.rg.R1, C1: w.o.rg.C1, R2: w.o.rg.R1, C2: w.o.rg.C2}) {
					continue
				}
				if _, ok := rc.sameRefs(w.o.refs, a.refs, w.from, k, display); !ok && !w.o.skipRefs {
					continue
				}
				used[i], found = true, true
				break
			}
			if found {
				continue
			}
			// What the sheet holds instead: the nearest of its kind (at the
			// same place, else with the same fixed text).
			instead := "nothing of its kind"
			for pass := 0; pass < 2 && instead == "nothing of its kind"; pass++ {
				for i, a := range after[k] {
					if used[i] || a.cover || a.kind != w.o.kind || pass == 0 && a.rg != w.o.rg || pass == 1 && a.fixed != w.o.fixed {
						continue
					}
					instead = a.String()
					if a.fixed != w.o.fixed {
						instead += " holding " + a.fixed + " (it held " + w.o.fixed + ")"
					}
					if a.refs != "" && a.refs != w.o.refs {
						instead += " reading " + strings.ReplaceAll(a.refs, "\n", "; ")
					}
					break
				}
			}
			gives := w.o.String()
			if w.o.refs != "" {
				if g, _ := rc.sameRefs(w.o.refs, w.o.refs, w.from, k, display); g != "" {
					gives += " reading " + strings.ReplaceAll(g, "\n", "; ")
				}
			}
			return fmt.Sprintf("%s: outside the edit's reach, the %s is not there as it was: the edit's shifts give the %s, and the sheet holds %s", display(k), w.o, gives, instead)
		}
		// The objects the sheet holds: a group's, or none of the expected.
		mask := append(append([]grid.Range{}, rc.solid[k]...), reached[k]...)
		var unmatched []object
		for i, a := range after[k] {
			if used[i] {
				continue
			}
			if a.cover {
				var in *group
				for _, g := range groups {
					if g.o.kind != a.kind || g.o.fixed != a.fixed {
						continue
					}
					if _, ok := rc.sameRefs(g.o.refs, a.refs, g.from, k, display); ok || g.o.skipRefs {
						in = g
						break
					}
				}
				if in != nil {
					in.got = append(in.got, a.rg)
					continue
				}
			}
			if a.free && (rc.hits(k, a.rg) || overlaps(reached[k], a.rg)) {
				continue
			}
			unmatched = append(unmatched, a)
		}
		for _, g := range groups {
			if at, ok := sameCover(g.want, g.got, mask); !ok {
				name := strings.TrimSpace(g.o.kind + " " + g.o.label)
				reads := ""
				if g.o.refs != "" {
					if want, _ := rc.sameRefs(g.o.refs, g.o.refs, g.from, k, display); want != "" {
						reads = " reading " + want
					}
				}
				// A rule of its own, read another way (not in the group).
				other := ""
				for _, a := range unmatched {
					if a.cover && a.kind == g.o.kind && a.fixed == g.o.fixed {
						other = fmt.Sprintf(", and the sheet holds one reading %s over %s", a.refs, a.rg)
						break
					}
				}
				return fmt.Sprintf("%s: outside the edit's reach, the %s is not there as it was: the edit's shifts give the %s%s over %s, and the sheet holds it over %s (they differ at %s)%s",
					display(k), name, name, reads, listRanges(g.want), listRanges(g.got), grid.CellName(at[0], at[1]), other)
			}
		}
		if len(unmatched) > 0 {
			return fmt.Sprintf("%s: outside the edit's reach, the edit made a %s it does not name", display(k), unmatched[0])
		}
	}
	return ""
}

// listRanges names ranges, "nothing" for none.
func listRanges(rs []grid.Range) string {
	if len(rs) == 0 {
		return "nothing"
	}
	var out []string
	for _, r := range rs {
		out = append(out, r.String())
	}
	return strings.Join(out, " ")
}

// sameCover tells whether two lists of ranges cover the same cells outside a
// mask; if not, a cell they differ on. It compares the compressed grid their
// edges make, so a range of a whole column costs no more than a cell.
func sameCover(a, b, mask []grid.Range) ([2]int, bool) {
	rows, cols := map[int]bool{1: true}, map[int]bool{1: true}
	for _, list := range [][]grid.Range{a, b, mask} {
		for _, r := range list {
			rows[r.R1], rows[r.R2+1] = true, true
			cols[r.C1], cols[r.C2+1] = true, true
		}
	}
	var rs, cs []int
	for r := range rows {
		rs = append(rs, r)
	}
	for c := range cols {
		cs = append(cs, c)
	}
	sort.Ints(rs)
	sort.Ints(cs)
	covers := func(list []grid.Range, r, c int) bool {
		for _, g := range list {
			if g.Contains(r, c) {
				return true
			}
		}
		return false
	}
	for _, r := range rs {
		for _, c := range cs {
			if covers(mask, r, c) {
				continue
			}
			if covers(a, r, c) != covers(b, r, c) {
				return [2]int{r, c}, false
			}
		}
	}
	return [2]int{}, true
}

// wholeLine tells whether a range holds the whole of a row (or column).
func wholeLine(rg grid.Range, p int, cols bool) bool {
	if cols {
		return rg.C1 <= p && p <= rg.C2 && rg.R1 <= 1 && rg.R2 >= grid.MaxRows
	}
	return rg.R1 <= p && p <= rg.R2 && rg.C1 <= 1 && rg.C2 >= grid.MaxCols
}

// judgeHidden judges the rows (columns) a sheet hides: those outside the
// lines the edit inserts and the whole lines it reaches (a sort of the sheet)
// are hidden after as they were before, each where the shifts put it.
func (rc *reach) judgeHidden(k0 string, was []grid.Run, k string, now []grid.Run, cols bool, display func(string) string) string {
	before, ok1 := hiddenLines(was)
	after, ok2 := hiddenLines(now)
	if !ok1 || !ok2 {
		return ""
	}
	free := func(p int) bool {
		for _, r := range rc.fresh[k] {
			if wholeLine(r, p, cols) {
				return true
			}
		}
		for _, r := range rc.solid[k] {
			if wholeLine(r, p, cols) {
				return true
			}
		}
		return false
	}
	what := "row"
	if cols {
		what = "column"
	}
	name := func(p int) string {
		if cols {
			return grid.ColName(p)
		}
		return fmt.Sprint(p)
	}
	want := map[int]bool{}
	for p := range before {
		cur := k0
	follow:
		for _, x := range rc.xforms {
			switch {
			case x.kind == 'r' && x.sheet == cur:
				cur = x.to
			case x.kind == 'l' && x.sheet == cur && x.cols == cols:
				q, ok := shiftLine(p, x.at, x.end, x.n)
				if !ok {
					p = 0
					break follow
				}
				p = q
			}
		}
		if p >= 1 && !free(p) {
			want[p] = true
		}
	}
	var lines []int
	for p := range want {
		lines = append(lines, p)
	}
	for p := range after {
		lines = append(lines, p)
	}
	sort.Ints(lines)
	for _, p := range lines {
		switch w, a := want[p], after[p] && !free(p); {
		case w && !a:
			return fmt.Sprintf("%s: outside the edit's reach, %s %s was hidden, and is shown now", display(k), what, name(p))
		case a && !w:
			return fmt.Sprintf("%s: outside the edit's reach, %s %s was shown, and is hidden now", display(k), what, name(p))
		}
	}
	return ""
}

// checkObjects judges what the sheets hold besides cells (objects.go): the
// grid reader's objects, the rows and columns they hide, and their
// visibility.
func (rc *reach) checkObjects(now map[string]*grid.Sheet, display func(string) string) string {
	before, after := map[string][]object{}, map[string][]object{}
	for k, s := range rc.sheets {
		before[k] = gridObjects(s)
	}
	for k, s := range now {
		after[k] = gridObjects(s)
	}
	exists := func(k string) bool { return now[k] != nil }
	if msg := rc.judgeObjects(before, after, exists, display); msg != "" {
		return msg
	}
	var keys []string
	for k := range rc.sheets {
		keys = append(keys, k)
	}
	sort.Strings(keys)
	for _, k0 := range keys {
		k := rc.follow(k0)
		s, n := rc.sheets[k0], now[k]
		if n == nil || rc.copies[k] {
			continue
		}
		for _, cols := range []bool{false, true} {
			was, nowRuns := s.HiddenRows, n.HiddenRows
			if cols {
				was, nowRuns = s.HiddenCols, n.HiddenCols
			}
			if msg := rc.judgeHidden(k0, was, k, nowRuns, cols, display); msg != "" {
				return msg
			}
		}
		if s.Hidden != n.Hidden {
			return fmt.Sprintf("%s: outside the edit's reach, the sheet was hidden: %v, and is hidden: %v now", display(k), s.Hidden, n.Hidden)
		}
	}
	return ""
}

// checkFile judges what the file holds that the reader does not give
// (validations, conditional formats, the look of each cell), before and after
// the edit, as checkObjects does the rest.
func (rc *reach) checkFile(after *grid.Workbook, before, afterObjs *fileObjects, order []string) string {
	now := map[string]*grid.Sheet{}
	for _, s := range after.Sheets {
		now[strings.ToLower(s.Name)] = s
	}
	display := func(k string) string {
		if s := now[k]; s != nil {
			return s.Name
		}
		if s := rc.sheets[k]; s != nil {
			return s.Name
		}
		return k
	}
	exists := func(k string) bool { return now[k] != nil }
	if msg := rc.judgeObjects(before.objects, afterObjs.objects, exists, display); msg != "" {
		return msg
	}
	// The look of each cell outside the reach, where it stands now.
	for _, k := range order {
		var ps [][2]int
		for p := range rc.at[k] {
			ps = append(ps, p)
		}
		// A contents-only reach keeps the look of its cells (f4.2).
		for p := range rc.kept[k] {
			ps = append(ps, p)
		}
		sort.Slice(ps, func(i, j int) bool {
			if ps[i][0] != ps[j][0] {
				return ps[i][0] < ps[j][0]
			}
			return ps[i][1] < ps[j][1]
		})
		for _, p := range ps {
			o, ok := rc.at[k][p]
			if !ok {
				o = rc.kept[k][p]
			}
			was, ok1 := before.looks[o.sheet][[2]int{o.r, o.c}]
			is, ok2 := afterObjs.looks[k][p]
			if !ok1 || !ok2 || was == is {
				continue
			}
			name := k + "!" + grid.CellName(p[0], p[1])
			if n := now[k]; n != nil {
				name = n.Name + "!" + grid.CellName(p[0], p[1])
			}
			return fmt.Sprintf("%s: outside the edit's reach, its look was [%s], and is [%s] now", name, was, is)
		}
	}
	return ""
}

// overlaps tells whether a range meets any of a list.
func overlaps(list []grid.Range, rg grid.Range) bool {
	for _, r := range list {
		if !r.Intersect(rg).Empty() {
			return true
		}
	}
	return false
}

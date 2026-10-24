package xlsx

import (
	"fmt"
	"sort"
	"strings"

	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/words"
	"golang.org/x/text/collate"
	"golang.org/x/text/language"
)

// SortKey is a column to sort rows by.
type SortKey struct {
	Col  int // the sheet's column
	Desc bool
}

// SortResult reports a sort.
type SortResult struct {
	Range string   `json:"range"`
	Rows  int      `json:"rows"`
	Moved int      `json:"moved"` // rows that changed place
	Notes []string `json:"notes,omitempty"`
}

// sortValue is a cell as Excel orders it: numbers (dates and times
// included), then text, then FALSE and TRUE, then errors; blanks come last
// whichever the direction.
type sortValue struct {
	class int // 0 number, 1 text, 2 logical, 3 error, 4 blank
	num   float64
	key   []byte // text's collation key: case and width do not count
}

// SortRange sorts the rows of r by keys, as Excel sorts: whole rows of r
// move together, with their formats, notes and links; ties keep their
// order. Formulas move as Excel moves them, their relative references
// shifting with the row; a sort that would change what one means (it
// reads another row of the sorted ones by position, or a column outside
// them in its row) is refused with the reason.
func (e *Editor) SortRange(sheet string, r grid.Range, keys []SortKey) (SortResult, error) {
	if !e.fresh() {
		return SortResult{}, errNotFresh
	}
	sheet, err := e.sheetName(sheet)
	if err != nil {
		return SortResult{}, err
	}
	r = e.stretch(sheet, r)
	res := SortResult{Range: grid.FormatRef(sheet, r), Rows: r.Rows()}
	s := e.wb.Sheet(sheet)
	if s == nil || r.Rows() < 2 {
		return res, nil
	}
	for _, k := range keys {
		if k.Col < r.C1 || k.Col > r.C2 {
			return res, fmt.Errorf("column %s is outside %s", grid.ColName(k.Col), r)
		}
	}
	if err := e.checkSort(s, r); err != nil {
		return res, err
	}

	// The order.
	coll := collate.New(language.Und, collate.IgnoreCase, collate.IgnoreWidth)
	var buf collate.Buffer
	vals := make([][]sortValue, len(keys))
	var unknown []string
	for i, k := range keys {
		vals[i] = make([]sortValue, r.Rows())
		for row := r.R1; row <= r.R2; row++ {
			v := sortValue{class: 4}
			if c := e.wb.Value(s, row, k.Col); c != nil && !c.IsEmpty() {
				if c.Pending {
					got, _, ok := e.Evaluate(sheet, row, k.Col)
					if !ok {
						unknown = append(unknown, grid.CellName(row, k.Col))
						continue
					}
					c = &got
				}
				switch c.Kind {
				case grid.Number:
					v = sortValue{class: 0, num: c.Num}
				case grid.Text:
					v = sortValue{class: 1, key: append([]byte{}, coll.KeyFromString(&buf, c.Text)...)}
					buf.Reset()
				case grid.Bool:
					v = sortValue{class: 2}
					if c.Bool {
						v.num = 1
					}
				case grid.Error:
					v = sortValue{class: 3}
				}
			}
			vals[i][row-r.R1] = v
		}
	}
	if len(unknown) > 0 {
		sort.Strings(unknown)
		return res, fmt.Errorf("cannot sort: the sheets calculator cannot compute %s, so the order is unknown until Excel computes %s", listOf(unknown, 5), map[bool]string{true: "it", false: "them"}[len(unknown) == 1])
	}
	perm := make([]int, r.Rows()) // perm[i]: the row (offset) that lands at offset i
	for i := range perm {
		perm[i] = i
	}
	sort.SliceStable(perm, func(a, b int) bool {
		for i, k := range keys {
			x, y := vals[i][perm[a]], vals[i][perm[b]]
			c := compareSort(x, y)
			if c == 0 {
				continue
			}
			if k.Desc && x.class != 4 && y.class != 4 {
				c = -c
			}
			return c < 0
		}
		return false
	})
	to := make([]int, len(perm)) // to[i]: where the row at offset i goes
	for i, p := range perm {
		to[p] = i
		if p != i {
			res.Moved++
		}
	}
	if res.Moved == 0 {
		return res, nil
	}
	if err := e.checkSortFormulas(s, r, to); err != nil {
		return res, err
	}
	res.Notes = append(res.Notes, e.sortReaders(s, r)...)

	// Move the rows.
	if err := e.unshareSheet(sheet); err != nil {
		return res, err
	}
	b, err := e.take(sheet, r)
	if err != nil {
		return res, err
	}
	moved := map[[2]int]*taken{}
	for k, h := range b.cells {
		moved[[2]int{to[k[0]], k[1]}] = h
	}
	if err := e.wipe(sheet, r, moved, pasteAll); err != nil {
		return res, err
	}
	for k, h := range moved {
		row := r.R1 + k[0]
		text := ""
		if h.formula != "" {
			text = formula.Shift(h.formula, row-(r.R1+perm[k[0]]), 0)
		}
		if err := e.put(sheet, row, r.C1+k[1], h, text, pasteAll); err != nil {
			return res, err
		}
	}
	// Comment threads follow their rows when saved.
	e.patches.threadMoves = append(e.patches.threadMoves, threadMove{sheet: sheet, move: func(row, col int) (int, int, bool) {
		if r.Contains(row, col) {
			return r.R1 + to[row-r.R1], col, true
		}
		return row, col, true
	}})
	return res, nil
}

func compareSort(x, y sortValue) int {
	switch {
	case x.class != y.class:
		return x.class - y.class
	case x.class == 1:
		return strings.Compare(string(x.key), string(y.key))
	case x.num < y.num:
		return -1
	case x.num > y.num:
		return 1
	}
	return 0
}

// checkSort refuses ranges Excel would not sort, or that sorting would
// break: merged cells, hidden rows, dynamic arrays, and part of an Excel
// table.
func (e *Editor) checkSort(s *grid.Sheet, r grid.Range) error {
	sheet := s.Name
	ms, err := e.f.GetMergeCells(sheet)
	if err != nil {
		return err
	}
	for _, m := range ms {
		mr, err := grid.ParseRange(m.GetStartAxis() + ":" + m.GetEndAxis())
		if err == nil && !mr.Intersect(r).Empty() {
			return fmt.Errorf("%s holds the merged cells %s, which Excel cannot sort either; unmerge them first (sheets style %s unmerge)", grid.FormatRef(sheet, r), mr, cliRange(sheet, mr))
		}
	}
	var hidden []string
	for row := r.R1; row <= r.R2; row++ {
		if s.RowHidden(row) {
			hidden = append(hidden, fmt.Sprint(row))
		}
	}
	if len(hidden) > 0 {
		return fmt.Errorf("row%s %s %s hidden (by a filter, or by hand); show %s first (sheets view %s --show %s:%s), since a sort would move %s among the others",
			map[bool]string{true: "", false: "s"}[len(hidden) == 1], listOf(hidden, 5), map[bool]string{true: "is", false: "are"}[len(hidden) == 1],
			map[bool]string{true: "it", false: "them"}[len(hidden) == 1], shellWord(sheet), hidden[0], hidden[len(hidden)-1], map[bool]string{true: "it", false: "them"}[len(hidden) == 1])
	}
	for a, sp := range e.spillsOn(sheet) {
		if !sp.Intersect(r).Empty() {
			return fmt.Errorf("%s holds the dynamic array formula in %s or its result (%s), which cannot be sorted; sort its source instead, or wrap it in SORT()", grid.FormatRef(sheet, r), grid.CellName(a[0], a[1]), sp)
		}
	}
	for _, t := range e.tablesOn(sheet) {
		if t.ref.Intersect(r).Empty() {
			continue
		}
		data := grid.Range{R1: t.ref.R1 + t.headers, C1: t.ref.C1, R2: t.ref.R2 - t.totals, C2: t.ref.C2}
		if !within([]grid.Range{r}, data) {
			return fmt.Errorf("%s reaches beyond the data rows of Excel table %s (%s); sort the table's data rows", grid.FormatRef(sheet, r), t, data)
		}
		if r.C1 != data.C1 || r.C2 != data.C2 {
			return fmt.Errorf("%s is part of the columns of Excel table %s; sort all of them (%s), so rows stay whole", grid.FormatRef(sheet, r), t, data)
		}
	}
	return nil
}

// checkSortFormulas refuses a sort that would change what a formula in the
// sorted rows means. Excel shifts a moved formula's relative references by
// the rows it moved, so the sort is safe when each formula reads cells of
// its own row within the sorted columns, cells outside the sorted rows by
// absolute reference, or whole ranges of them, and when formulas reading
// other rows (a running total, a change from the row above) fill a column
// alike, so the column means the same in the new order.
func (e *Editor) checkSortFormulas(s *grid.Sheet, r grid.Range, to []int) error {
	sheet := s.Name
	pattern := map[int]bool{} // columns filled alike with formulas
	for col := r.C1; col <= r.C2; col++ {
		key, alike := "", true
		for row := r.R1; row <= r.R2 && alike; row++ {
			c := s.At(row, col)
			if c == nil || c.Formula == "" {
				alike = false
				break
			}
			k, _ := formula.Pattern(c.Formula, row, col)
			if key == "" {
				key = k
			}
			alike = k == key
		}
		pattern[col] = alike
	}
	rs := e.resolver()
	var bad error
	s.Each(r, func(row, col int, c *grid.Cell) {
		if bad != nil || c.Formula == "" {
			return
		}
		moves := to[row-r.R1] != row-r.R1
		at := grid.FormatRef(sheet, grid.Range{R1: row, C1: col, R2: row, C2: col})
		why := func(problem, fix string) {
			bad = fmt.Errorf("sorting %s would change what %s means: =%s %s. %s", r, at, c.Formula, problem, fix)
		}
		for _, u := range formula.RefUses(c.Formula) {
			if u.Sheet != "" && !strings.EqualFold(u.Sheet, sheet) {
				if !u.AbsRows && u.Range.Rows() < grid.MaxRows {
					why("reads "+u.Text+" by its row, so after the sort it would pair this row with another row there",
						"Look the value up instead (XLOOKUP on a key column), or make the reference absolute")
					return
				}
				continue
			}
			rg := u.Range
			switch {
			case rg.R1 == 1 && rg.R2 == grid.MaxRows: // whole columns
			case u.AbsRows:
				if rg.R2 < r.R1 || rg.R1 > r.R2 || rg.R1 <= r.R1 && rg.R2 >= r.R2 {
					continue
				}
				why("points at "+u.Text+", some of the sorted rows, which will hold other rows' values",
					"Point it at all of the sorted rows, or at cells outside them")
				return
			case u.RelRows && rg.R1 == row && rg.R2 == row:
				if moves && (rg.C1 < r.C1 || rg.C2 > r.C2) {
					why("reads "+u.Text+" in its own row, outside the sorted columns, which stay put",
						"Sort a range that takes in that column too, so rows stay whole")
					return
				}
			case pattern[col]:
				// A column of alike formulas reading other rows: Excel's
				// shifted formulas fill it the same way in the new order.
				if rg.C1 < r.C1 || rg.C2 > r.C2 {
					why("reads "+u.Text+", outside the sorted columns, which stay put",
						"Sort a range that takes in that column too, so rows stay whole")
					return
				}
			default:
				why("reads "+u.Text+", another row, by its position, so after the sort it would read whatever row lands there",
					"Make the reference absolute ($), or clear the formula and add it again after sorting")
				return
			}
		}
		// Names pointing into the sorted rows.
		for _, n := range formula.DepsOf(c.Formula).Names {
			for _, d := range rs.names[strings.ToLower(n)] {
				for _, ref := range formula.Refs(d.text) {
					if ref.Sheet != "" && !strings.EqualFold(ref.Sheet, sheet) {
						continue
					}
					rg := ref.Range
					if rg.R2 < r.R1 || rg.R1 > r.R2 || rg.R1 <= r.R1 && rg.R2 >= r.R2 {
						continue
					}
					why("uses the name "+n+", which points at some of the sorted rows",
						"Point the name at all of the sorted rows, or at cells outside them")
					return
				}
			}
		}
	})
	return bad
}

// sortReaders lists the formulas elsewhere that read single cells or part
// of the sorted rows: after the sort, they read whatever row is there.
func (e *Editor) sortReaders(s *grid.Sheet, r grid.Range) []string {
	var hits []string
	for _, sh := range e.wb.Sheets {
		sh.Each(sh.Used(), func(row, col int, c *grid.Cell) {
			if c.Formula == "" || strings.EqualFold(sh.Name, s.Name) && r.Contains(row, col) {
				return
			}
			for _, u := range formula.RefUses(c.Formula) {
				sheet := u.Sheet
				if sheet == "" {
					sheet = sh.Name
				}
				rg := u.Range
				if !strings.EqualFold(sheet, s.Name) || rg.Intersect(r).Empty() || rg.R1 <= r.R1 && rg.R2 >= r.R2 {
					continue
				}
				hits = append(hits, fmt.Sprintf("%s (=%s)", grid.FormatRef(sh.Name, grid.Range{R1: row, C1: col, R2: row, C2: col}), c.Formula))
				return
			}
		})
	}
	if len(hits) == 0 {
		return nil
	}
	return []string{fmt.Sprintf("%s %s single cells of the sorted rows and now %s whatever row is there: %s",
		countOf(len(hits), "formula"), map[bool]string{true: "reads", false: "read"}[len(hits) == 1], map[bool]string{true: "reads", false: "read"}[len(hits) == 1], listOf(hits, 5))}
}

// listOf joins up to n items, saying how many more there are.
func listOf(items []string, n int) string {
	if len(items) <= n {
		return strings.Join(items, ", ")
	}
	return strings.Join(items[:n], ", ") + fmt.Sprintf(" and %d more", len(items)-n)
}

// shellWord quotes a word for a suggested command when it needs it.
func shellWord(s string) string { return words.Quote(s) }

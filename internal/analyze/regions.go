// Package analyze finds structure in a sheet: blocks of data, their title,
// header and total rows, and what each column holds.
package analyze

import (
	"regexp"
	"sort"
	"strings"

	"github.com/eemax/sheets/internal/grid"
)

// Region is a block of connected non-empty cells.
type Region struct {
	Range     grid.Range
	TitleRows int // leading rows holding a single title cell
	// GroupRow is a row of labels above the header that group its columns
	// (Q1 over Units and Revenue); 0 when there is none.
	GroupRow  int
	HeaderRow int        // 0 when the region has no header
	TotalRow  int        // 0 when the region has no total row
	Body      grid.Range // data rows: excludes title, group, header and total rows
	// IsTitle marks a lone text cell that captions the table just below it.
	IsTitle bool
}

// IsTable reports whether the region has a header and at least one data row.
func (r Region) IsTable() bool { return r.HeaderRow > 0 && !r.Body.Empty() }

// Regions finds the blocks of data on a sheet, top to bottom, left to right.
func Regions(wb *grid.Workbook, s *grid.Sheet) []Region {
	boxes := components(s)
	out := make([]Region, 0, len(boxes))
	for _, b := range boxes {
		out = append(out, Classify(wb, s, b))
	}
	// A single text cell up to two blank rows above a table, within its
	// columns, is that table's title.
	for i := range out {
		r := out[i].Range
		if r.Rows() != 1 || r.Cols() != 1 {
			continue
		}
		if c := s.At(r.R1, r.C1); c == nil || c.Kind != grid.Text || c.Formula != "" {
			continue
		}
		for _, t := range out {
			if t.IsTable() && t.Range.R1 > r.R1 && t.Range.R1-r.R1 <= 3 && r.C1 >= t.Range.C1 && r.C1 <= t.Range.C2 {
				out[i].IsTitle = true
				break
			}
		}
	}
	return out
}

type run struct{ c1, c2, id int }

// components groups non-empty cells that touch (including diagonally) and
// returns the bounding boxes, merged until none overlap.
func components(s *grid.Sheet) []grid.Range {
	used := s.Used()
	if used.Empty() {
		return nil
	}
	parent := []int{}
	find := func(x int) int {
		for parent[x] != x {
			parent[x] = parent[parent[x]]
			x = parent[x]
		}
		return x
	}
	type rowRuns struct {
		row  int
		runs []run
	}
	var all []rowRuns
	var prev []run
	prevRow := -1
	for row := used.R1; row <= used.R2; row++ {
		var cur []run
		n := s.RowLen(row)
		for col := 1; col <= n; col++ {
			if c := s.At(row, col); c == nil || c.IsEmpty() {
				continue
			}
			if len(cur) > 0 && cur[len(cur)-1].c2 == col-1 {
				cur[len(cur)-1].c2 = col
				continue
			}
			id := len(parent)
			parent = append(parent, id)
			cur = append(cur, run{col, col, id})
		}
		if prevRow == row-1 {
			for _, a := range cur {
				for _, b := range prev {
					if b.c1 <= a.c2+1 && b.c2 >= a.c1-1 {
						parent[find(a.id)] = find(b.id)
					}
				}
			}
		}
		if len(cur) > 0 {
			all = append(all, rowRuns{row, cur})
			prev, prevRow = cur, row
		}
	}
	boxes := map[int]grid.Range{}
	for _, rr := range all {
		for _, r := range rr.runs {
			root := find(r.id)
			b := grid.Range{R1: rr.row, C1: r.c1, R2: rr.row, C2: r.c2}
			if old, ok := boxes[root]; ok {
				b = old.Union(b)
			}
			boxes[root] = b
		}
	}
	list := make([]grid.Range, 0, len(boxes))
	for _, b := range boxes {
		list = append(list, b)
	}
	// Merge overlapping boxes until stable.
	for merged := true; merged; {
		merged = false
		for i := 0; i < len(list) && !merged; i++ {
			for j := i + 1; j < len(list); j++ {
				if !list[i].Intersect(list[j]).Empty() {
					list[i] = list[i].Union(list[j])
					list = append(list[:j], list[j+1:]...)
					merged = true
					break
				}
			}
		}
	}
	sort.Slice(list, func(i, j int) bool {
		if list[i].R1 != list[j].R1 {
			return list[i].R1 < list[j].R1
		}
		return list[i].C1 < list[j].C1
	})
	return list
}

var (
	totalLabel = regexp.MustCompile(`(?i)^\s*(grand\s+)?(total|totals|sum|subtotal|sub-total)\b`)
	idHeader   = regexp.MustCompile(`(?i)(^|[^a-z])(id|year|yr|fy|code|zip|postal|phone|no\.?|#)([^a-z]|$)`)
)

// Label returns the header text of a column; numeric labels such as years
// are shown as written.
func Label(wb *grid.Workbook, s *grid.Sheet, row, col int) string {
	c := s.At(row, col)
	if c == nil || c.IsEmpty() {
		return ""
	}
	return grid.Display(c, wb.FormatOf(c), wb.Date1904)
}

// IDLike reports whether a header names identifiers or years: numbers that
// are labels, not quantities (no thousands separators, no sums).
func IDLike(header string) bool { return idHeader.MatchString(header) }

// Classify finds the title, header and total rows of a block.
func Classify(wb *grid.Workbook, s *grid.Sheet, b grid.Range) Region {
	reg := Region{Range: b, Body: b}
	width := b.Cols()
	// Title rows: a single text cell, at the left edge, above a wider row.
	// (A lone label further right is a crosstab header with a blank corner.)
	for r := b.R1; r < b.R2 && reg.TitleRows < 3 && width > 1; r++ {
		c := s.At(r, b.C1)
		if c == nil || c.Kind != grid.Text || c.Formula != "" || filled(s, r, b) != 1 || filled(s, r+1, b) < 2 {
			break
		}
		reg.TitleRows++
	}
	h := b.R1 + reg.TitleRows
	switch {
	case h+1 < b.R2 && isGroupRow(wb, s, h, b) && isHeader(wb, s, h+1, b):
		reg.GroupRow, reg.HeaderRow = h, h+1
	case h < b.R2 && isHeader(wb, s, h, b):
		reg.HeaderRow = h
	}
	first := h
	if reg.HeaderRow > 0 {
		first = reg.HeaderRow + 1
	}
	reg.Body = grid.Range{R1: first, C1: b.C1, R2: b.R2, C2: b.C2}
	if reg.Body.Rows() >= 2 && (isTotal(s, b.R2, b) || excelTotals(s, b)) {
		reg.TotalRow = b.R2
		reg.Body.R2--
	}
	return reg
}

// isGroupRow reports whether row h groups the labels of the row below:
// text labels (or years) only, with fewer distinct labels than the row
// below has, because each spans several columns.
func isGroupRow(wb *grid.Workbook, s *grid.Sheet, h int, b grid.Range) bool {
	if b.Cols() < 2 {
		return false
	}
	labels := map[string]bool{}
	ok := true
	s.Each(grid.Range{R1: h, C1: b.C1, R2: h, C2: b.C2}, func(_, _ int, c *grid.Cell) {
		switch {
		case c.Formula != "":
			ok = false
		case c.Kind == grid.Text && strings.TrimSpace(c.Text) != "":
			labels[strings.ToLower(strings.TrimSpace(c.Text))] = true
		case c.Kind == grid.Number && c.Num == float64(int(c.Num)) && c.Num >= 1900 && c.Num <= 2100:
			labels[grid.FormatNumber(c.Num)] = true
		default:
			ok = false
		}
	})
	return ok && len(labels) > 0 && len(labels) < filled(s, h+1, b)
}

// ColumnLabel names a column of a region by its header, prefixed with its
// group when the region has a group row: "Q1 / Units".
func ColumnLabel(wb *grid.Workbook, s *grid.Sheet, reg Region, col int) string {
	if reg.HeaderRow == 0 {
		return ""
	}
	label := Label(wb, s, reg.HeaderRow, col)
	if reg.GroupRow == 0 {
		return label
	}
	group := GroupLabel(wb, s, reg, col)
	switch {
	case group == "" || strings.EqualFold(group, label):
		return label
	case label == "":
		return group
	}
	return group + " / " + label
}

// GroupLabel returns the group a column falls under: the label above it, the
// label of a merged range covering it, or (when the row has no merges) the
// nearest label to its left.
func GroupLabel(wb *grid.Workbook, s *grid.Sheet, reg Region, col int) string {
	g := reg.GroupRow
	if g == 0 {
		return ""
	}
	if l := Label(wb, s, g, col); l != "" {
		return l
	}
	merged := false
	for _, m := range s.Merges {
		if m.R1 <= g && m.R2 >= g {
			merged = true
			if m.Contains(g, col) {
				return Label(wb, s, m.R1, m.C1)
			}
		}
	}
	if merged {
		return ""
	}
	for c := col - 1; c >= reg.Range.C1; c-- {
		if l := Label(wb, s, g, c); l != "" {
			return l
		}
	}
	return ""
}

func filled(s *grid.Sheet, row int, b grid.Range) int {
	n := 0
	s.Each(grid.Range{R1: row, C1: b.C1, R2: row, C2: b.C2}, func(int, int, *grid.Cell) { n++ })
	return n
}

// isHeader decides whether row h labels the columns below it. Labels are
// text, or years (Region | 2024 | 2025); the top-left cell of a crosstab
// may be blank.
func isHeader(wb *grid.Workbook, s *grid.Sheet, h int, b grid.Range) bool {
	width := b.Cols()
	labels := map[string]bool{}
	n, texts := 0, 0
	unique := true
	for col := b.C1; col <= b.C2; col++ {
		c := s.At(h, col)
		if c == nil || c.IsEmpty() {
			continue
		}
		if c.Formula != "" {
			return false
		}
		var key string
		switch {
		case c.Kind == grid.Text:
			key = strings.ToLower(strings.TrimSpace(c.Text))
			if key == "" {
				return false
			}
			texts++
		case c.Kind == grid.Number && c.Num == float64(int(c.Num)) && c.Num >= 1900 && c.Num <= 2100:
			key = grid.FormatNumber(c.Num)
		default:
			return false
		}
		if labels[key] {
			unique = false
		}
		labels[key] = true
		n++
	}
	if texts == 0 || n*2 < width {
		return false
	}
	// Strong evidence: some column holds non-text data below a text label.
	last := min(b.R2, h+50)
	for col := b.C1; col <= b.C2; col++ {
		if c := s.At(h, col); c == nil || c.IsEmpty() {
			continue
		}
		for r := h + 1; r <= last; r++ {
			c := wb.Value(s, r, col)
			if c != nil && (c.Kind == grid.Number || c.Kind == grid.Bool) {
				return true
			}
		}
	}
	// All text: accept a fully labelled row of distinct labels that do not
	// recur in at least two rows below.
	if n != width || !unique || b.R2-h < 2 {
		return false
	}
	for col := b.C1; col <= b.C2; col++ {
		label := strings.ToLower(strings.TrimSpace(s.Get(h, col).Text))
		for r := h + 1; r <= last; r++ {
			if c := s.At(r, col); c != nil && c.Kind == grid.Text && strings.ToLower(strings.TrimSpace(c.Text)) == label {
				return false
			}
		}
	}
	return true
}

// excelTotals reports whether b is an Excel table whose last row is its
// totals row.
func excelTotals(s *grid.Sheet, b grid.Range) bool {
	for _, t := range s.Tables {
		if t.Range == b && t.Totals > 0 {
			return true
		}
	}
	return false
}

func isTotal(s *grid.Sheet, row int, b grid.Range) bool {
	sums, other := 0, 0
	labelled := false
	firstSeen := false
	s.Each(grid.Range{R1: row, C1: b.C1, R2: row, C2: b.C2}, func(_, _ int, c *grid.Cell) {
		if !firstSeen {
			firstSeen = true
			labelled = c.Kind == grid.Text && c.Formula == "" && totalLabel.MatchString(c.Text)
		}
		if c.Formula != "" {
			u := strings.ToUpper(strings.TrimSpace(c.Formula))
			if strings.HasPrefix(u, "SUM(") || strings.HasPrefix(u, "SUBTOTAL(") {
				sums++
				return
			}
		}
		if c.Kind == grid.Number {
			other++
		}
	})
	return labelled || sums > 0 && other == 0
}

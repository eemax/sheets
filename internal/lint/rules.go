package lint

import (
	"fmt"
	"slices"
	"sort"
	"strings"

	"github.com/eemax/sheets/internal/analyze"
	"github.com/eemax/sheets/internal/engine"
	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/input"
)

// Rules that read across the workbook: what a formula adds up, what its
// label says, which tables it names.

// formulaCell is a formula and where it is.
type formulaCell struct {
	sheet      *grid.Sheet
	row, col   int
	text       string
	unprefixed bool // see grid.Cell.Unprefixed
}

func allFormulas(wb *grid.Workbook) []formulaCell {
	var out []formulaCell
	for _, s := range wb.Sheets {
		s.Each(s.Used(), func(r, c int, cell *grid.Cell) {
			if cell.Formula != "" {
				out = append(out, formulaCell{s, r, c, cell.Formula, cell.Unprefixed})
			}
		})
	}
	return out
}

// sheetOf resolves the sheet a reference names, "" being the formula's own.
func sheetOf(wb *grid.Workbook, home *grid.Sheet, name string) *grid.Sheet {
	if name == "" {
		return home
	}
	return wb.Sheet(name)
}

func areaRange(a engine.Area) grid.Range { return grid.Range{R1: a.R1, C1: a.C1, R2: a.R2, C2: a.C2} }

// subtotal is a cell totalling a range of its own sheet: =SUM(N4:N27).
type subtotal struct {
	row    int
	ranges []grid.Range
}

// doubleCounts finds sums that add a total together with the rows it
// totals: =SUM(N4:N28) where N28 is =SUM(N4:N27). SUBTOTAL leaves out the
// subtotals it meets, so only SUM counts.
func doubleCounts(wb *grid.Workbook, formulas []formulaCell, only func(*grid.Sheet) bool) []Finding {
	// Cells holding a total of a range, by sheet and column.
	totals := map[*grid.Sheet]map[int][]subtotal{}
	for _, f := range formulas {
		u := strings.ToUpper(strings.TrimSpace(f.text))
		if !strings.HasPrefix(u, "SUM(") && !strings.HasPrefix(u, "SUBTOTAL(") {
			continue
		}
		calls, err := engine.Calls(f.text)
		if err != nil || len(calls) == 0 || len(calls[0].Refs) == 0 {
			continue
		}
		st := subtotal{row: f.row}
		for _, r := range calls[0].Refs {
			if sheetOf(wb, f.sheet, r.Sheet) != f.sheet {
				st.ranges = nil
				break
			}
			st.ranges = append(st.ranges, areaRange(r.Area))
		}
		if len(st.ranges) == 0 || st.ranges[0].Rows()*st.ranges[0].Cols() < 2 {
			continue
		}
		if totals[f.sheet] == nil {
			totals[f.sheet] = map[int][]subtotal{}
		}
		totals[f.sheet][f.col] = append(totals[f.sheet][f.col], st)
	}
	for _, byCol := range totals {
		for _, list := range byCol {
			sort.Slice(list, func(i, j int) bool { return list[i].row < list[j].row })
		}
	}

	var out []Finding
	for _, f := range formulas {
		if !only(f.sheet) || !strings.Contains(strings.ToUpper(f.text), "SUM(") {
			continue
		}
		calls, err := engine.Calls(f.text)
		if err != nil {
			continue
		}
		var hits []string
		var covered []string
		for _, c := range calls {
			if c.Func != "SUM" {
				continue
			}
			for _, r := range c.Refs {
				t := sheetOf(wb, f.sheet, r.Sheet)
				if t == nil || totals[t] == nil {
					continue
				}
				area := areaRange(r.Area).Intersect(t.Used())
				for col := area.C1; col <= area.C2 && !area.Empty(); col++ {
					list := totals[t][col]
					i := sort.Search(len(list), func(i int) bool { return list[i].row >= area.R1 })
					for ; i < len(list) && list[i].row <= area.R2; i++ {
						st := list[i]
						if t == f.sheet && st.row == f.row && col == f.col {
							continue
						}
						inside := true
						for _, sr := range st.ranges {
							inside = inside && sr.Intersect(areaRange(r.Area)) == sr
						}
						if !inside {
							continue
						}
						hits = append(hits, cellOn(f.sheet, t, st.row, col))
						covered = append(covered, rangeOn(f.sheet, t, st.ranges[0]))
					}
				}
			}
		}
		if len(hits) == 0 {
			continue
		}
		what := fmt.Sprintf("%s, a total of %s", hits[0], covered[0])
		if len(hits) > 1 {
			what = fmt.Sprintf("%s (%s), totals of rows", strings.Join(hits[:min(len(hits), 4)], ", "), map[bool]string{true: "and more", false: "each"}[len(hits) > 4])
		}
		out = append(out, Finding{Rule: "double-count", Severity: Warning, Sheet: f.sheet.Name, Cells: []string{ref(f.row, f.col)}, row: f.row, col: f.col,
			Message: fmt.Sprintf("=%s adds %s it already adds: double counting", f.text, what)})
	}
	return out
}

// cellOn names a cell of sheet t as a formula on home writes it.
func cellOn(home, t *grid.Sheet, row, col int) string {
	return rangeOn(home, t, grid.Range{R1: row, C1: col, R2: row, C2: col})
}

func rangeOn(home, t *grid.Sheet, r grid.Range) string {
	if home == t {
		return r.String()
	}
	return grid.QuoteSheet(t.Name) + "!" + r.String()
}

// labelMismatches finds formulas reading one cell of a line other than the
// one their label names, when the label names another line of that
// table: Summary "Marketing" =Budget!N9, which is the Events line, where
// Marketing is row 8. Columns likewise, by their headers.
func labelMismatches(wb *grid.Workbook, formulas []formulaCell, regions func(*grid.Sheet) []analyze.Region, only func(*grid.Sheet) bool) []Finding {
	var out []Finding
	for _, f := range formulas {
		if !only(f.sheet) {
			continue
		}
		r, ok := engine.Reference(f.text)
		if !ok || r.Area.R1 != r.Area.R2 || r.Area.C1 != r.Area.C2 {
			continue
		}
		t := sheetOf(wb, f.sheet, r.Sheet)
		if t == nil {
			continue
		}
		tr, tc := r.Area.R1, r.Area.C1
		reg, ok := regionAt(regions(t), tr, tc)
		// A formula reading its own block computes from it; one reading
		// another block picks a value out, by its label.
		if !ok || t == f.sheet && reg.Range.Contains(f.row, f.col) {
			continue
		}
		target := grid.FormatRef(t.Name, grid.Range{R1: tr, C1: tc, R2: tr, C2: tc})
		if t == f.sheet {
			target = ref(tr, tc)
		}
		// By rows: the label to the left of the formula against the
		// label of the line it reads.
		if own := labelLeft(f.sheet, f.row, f.col); own != "" && tc > reg.Range.C1 {
			if read := plainText(t, tr, reg.Range.C1); read != "" && !sameLabel(own, read) {
				if r2, ok := onlyMatch(reg.Range.R1, reg.Range.R2, func(r int) string { return plainText(t, r, reg.Range.C1) }, own); ok && !empty(t, r2, tc) {
					out = append(out, Finding{Rule: "label-mismatch", Severity: Warning, Sheet: f.sheet.Name, Cells: []string{ref(f.row, f.col)}, row: f.row, col: f.col,
						Message: fmt.Sprintf("labelled %q but reads %s, the %q line; the %s line is %s", own, target, read, own, cellOn(f.sheet, t, r2, tc))})
				}
				continue
			}
		}
		// By columns: the header above the formula against the header of
		// the column it reads.
		h := reg.Headed().HeaderRow
		if own := labelAbove(f.sheet, regions(f.sheet), f.row, f.col); own != "" && h > 0 && tr > h {
			if read := plainText(t, h, tc); read != "" && !sameLabel(own, read) {
				if c2, ok := onlyMatch(reg.Range.C1, reg.Range.C2, func(c int) string { return plainText(t, h, c) }, own); ok && !empty(t, tr, c2) {
					out = append(out, Finding{Rule: "label-mismatch", Severity: Warning, Sheet: f.sheet.Name, Cells: []string{ref(f.row, f.col)}, row: f.row, col: f.col,
						Message: fmt.Sprintf("headed %q but reads %s, in the %q column; the %s column is %s", own, target, read, own, cellOn(f.sheet, t, tr, c2))})
				}
			}
		}
	}
	return out
}

func empty(s *grid.Sheet, row, col int) bool {
	c := s.At(row, col)
	return c == nil || c.IsEmpty()
}

// onlyMatch finds the one line from..to whose label is label; ok is false
// when none or several are.
func onlyMatch(from, to int, labelOf func(int) string, label string) (int, bool) {
	at := 0
	for i := from; i <= to; i++ {
		if sameLabel(labelOf(i), label) {
			if at > 0 {
				return 0, false
			}
			at = i
		}
	}
	return at, at > 0
}

// regionAt finds the block holding a cell.
func regionAt(regs []analyze.Region, row, col int) (analyze.Region, bool) {
	for _, r := range regs {
		if r.Range.Contains(row, col) {
			return r, true
		}
	}
	return analyze.Region{}, false
}

// plainText is the text a cell holds as typed, trimmed.
func plainText(s *grid.Sheet, row, col int) string {
	c := s.At(row, col)
	if c == nil || c.Kind != grid.Text || c.Formula != "" {
		return ""
	}
	return strings.TrimSpace(c.Text)
}

// labelLeft is the label of a formula's line: the nearest value to its
// left, when that is text.
func labelLeft(s *grid.Sheet, row, col int) string {
	for c := col - 1; c >= 1; c-- {
		if cell := s.At(row, c); cell != nil && !cell.IsEmpty() {
			return plainText(s, row, c)
		}
	}
	return ""
}

// labelAbove is the header of a formula's column: its block's header, or
// the nearest text above it in the block.
func labelAbove(s *grid.Sheet, regs []analyze.Region, row, col int) string {
	reg, ok := regionAt(regs, row, col)
	if !ok {
		return ""
	}
	if h := reg.Headed().HeaderRow; h > 0 && h < row {
		return plainText(s, h, col)
	}
	for r := row - 1; r >= reg.Range.R1; r-- {
		if cell := s.At(r, col); cell != nil && !cell.IsEmpty() {
			return plainText(s, r, col)
		}
	}
	return ""
}

// sameLabel compares labels loosely: case, spacing, and one naming the
// other within it ("Marketing" and "Marketing costs").
func sameLabel(a, b string) bool {
	a, b = strings.Join(strings.Fields(strings.ToLower(a)), " "), strings.Join(strings.Fields(strings.ToLower(b)), " ")
	if a == "" || b == "" {
		return false
	}
	return a == b || strings.Contains(a, b) || strings.Contains(b, a)
}

// unknownTables finds structured references to Excel tables the workbook
// does not have, which Excel cannot compute.
func unknownTables(wb *grid.Workbook, formulas []formulaCell, only func(*grid.Sheet) bool) []Finding {
	have := map[string]bool{}
	var names []string
	for _, s := range wb.Sheets {
		for _, t := range s.Tables {
			have[strings.ToLower(t.Name)] = true
			names = append(names, t.Name)
		}
	}
	tables := "the workbook has no Excel tables"
	if len(names) > 0 {
		tables = "its Excel tables are " + strings.Join(names, ", ")
	}
	var out []Finding
	for _, f := range formulas {
		if !only(f.sheet) || !strings.Contains(f.text, "[") {
			continue
		}
		for _, name := range formula.DepsOf(f.text).Tables {
			var why string
			switch {
			case name == "" && f.sheet.TableAt(grid.Range{R1: f.row, C1: f.col, R2: f.row, C2: f.col}) == nil:
				why = "uses a column reference like [Units] outside an Excel table"
			case name != "" && !have[strings.ToLower(name)]:
				why = fmt.Sprintf("names table %s, which does not exist (%s)", name, tables)
			default:
				continue
			}
			out = append(out, Finding{Rule: "unknown-table", Severity: Error, Sheet: f.sheet.Name, Cells: []string{ref(f.row, f.col)}, row: f.row, col: f.col,
				Message: fmt.Sprintf("=%s %s; Excel cannot compute it: write the range instead", f.text, why)})
			break
		}
	}
	return out
}

// textNumbers finds numbers stored as text outside tables, between numbers
// in their row or column: a report's figure typed in as text, which SUM
// leaves out.
func textNumbers(wb *grid.Workbook, s *grid.Sheet, regs []analyze.Region) []Finding {
	inTable := func(row, col int) bool {
		for _, r := range regs {
			if (r.IsTable() || r.Continues > 0) && r.Range.Contains(row, col) {
				return true
			}
		}
		return false
	}
	isNumber := func(row, col int) bool {
		c := wb.Value(s, row, col)
		return c != nil && c.Kind == grid.Number
	}
	var cells []string
	first, firstCol := 0, 0
	s.Each(s.Used(), func(r, c int, cell *grid.Cell) {
		if cell.Kind != grid.Text || cell.Formula != "" || inTable(r, c) || !numberText(cell.Text) {
			return
		}
		if !(isNumber(r, c-1) && isNumber(r, c+1) || isNumber(r-1, c) && isNumber(r+1, c)) {
			return
		}
		if len(cells) == 0 {
			first, firstCol = r, c
		}
		cells = append(cells, ref(r, c))
	})
	if len(cells) == 0 {
		return nil
	}
	return []Finding{{Rule: "number-as-text", Severity: Warning, Sheet: s.Name, Cells: cells, row: first, col: firstCol,
		Message: fmt.Sprintf("%s between numbers %s stored as text; SUM and sorting ignore them (rewrite them with sheets set)", countOf(len(cells), "value"), map[bool]string{true: "is a number", false: "are numbers"}[len(cells) == 1])}}
}

// numberText reports whether text reads as a number, not as an identifier
// with leading zeros or a long code.
func numberText(text string) bool {
	t := strings.TrimSpace(text)
	if t == "" || len(t) > 15 || len(t) > 1 && t[0] == '0' && t[1] != '.' && t[1] != ',' {
		return false
	}
	return input.Parse(t, false).Kind == grid.Number
}

// legacyArrays finds formulas that compute over a range where one value
// goes but are stored as plain formulas, as tools writing files for older
// Excel store them: Excel reads one cell of the range (Excel 365 shows
// =SUM(IF(@A2:A9>5,B2:B9))), not the whole range their author meant.
func legacyArrays(formulas []formulaCell, only func(*grid.Sheet) bool) []Finding {
	bySheet := map[*grid.Sheet][]string{}
	first := map[*grid.Sheet]formulaCell{}
	var order []*grid.Sheet
	for _, f := range formulas {
		if !only(f.sheet) || inArray(f.sheet, f.row, f.col) {
			continue
		}
		if array, _ := engine.ArrayUse(f.text); !array {
			continue
		}
		if bySheet[f.sheet] == nil {
			order = append(order, f.sheet)
			first[f.sheet] = f
		}
		bySheet[f.sheet] = append(bySheet[f.sheet], ref(f.row, f.col))
	}
	var out []Finding
	for _, s := range order {
		f, cells := first[s], bySheet[s]
		out = append(out, Finding{Rule: "implicit-intersection", Severity: Warning, Sheet: s.Name, Cells: cells, row: f.row, col: f.col,
			Message: fmt.Sprintf("%s like =%s %s over a range where one value goes, but %s stored for older Excel, which reads one cell of it (Excel 365 shows @); set %s again with sheets set to store %s as Excel 365 does",
				countOf(len(cells), "formula"), f.text, map[bool]string{true: "computes", false: "compute"}[len(cells) == 1], map[bool]string{true: "is", false: "are"}[len(cells) == 1],
				map[bool]string{true: "it", false: "them"}[len(cells) == 1], map[bool]string{true: "it", false: "them"}[len(cells) == 1])})
	}
	return out
}

// unprefixedFunctions finds formulas stored calling functions added after
// Excel 2007 without the prefix Excel needs, as openpyxl writes them:
// Excel shows #NAME? for them.
func unprefixedFunctions(formulas []formulaCell, only func(*grid.Sheet) bool) []Finding {
	bySheet := map[*grid.Sheet][]string{}
	fns := map[*grid.Sheet][]string{}
	first := map[*grid.Sheet]formulaCell{}
	var order []*grid.Sheet
	for _, f := range formulas {
		if !f.unprefixed || !only(f.sheet) {
			continue
		}
		if bySheet[f.sheet] == nil {
			order = append(order, f.sheet)
			first[f.sheet] = f
		}
		bySheet[f.sheet] = append(bySheet[f.sheet], ref(f.row, f.col))
		for _, fn := range formula.Unprefixed(f.text) {
			if !slices.Contains(fns[f.sheet], fn) {
				fns[f.sheet] = append(fns[f.sheet], fn)
			}
		}
	}
	var out []Finding
	for _, s := range order {
		f, cells, one := first[s], bySheet[s], len(bySheet[s]) == 1
		out = append(out, Finding{Rule: "unprefixed-function", Severity: Error, Sheet: s.Name, Cells: cells, row: f.row, col: f.col,
			Message: fmt.Sprintf("%s like =%s %s %s without the _xlfn. prefix Excel needs for functions added after 2007 (openpyxl leaves it off), so Excel shows #NAME? until %s entered again; set %s again with sheets set, which adds it",
				countOf(len(cells), "formula"), f.text, map[bool]string{true: "calls", false: "call"}[one], strings.Join(fns[s], ", "),
				map[bool]string{true: "it is", false: "they are"}[one], map[bool]string{true: "it", false: "them"}[one])})
	}
	return out
}

// inArray tells whether a cell holds an array formula (or a dynamic one).
func inArray(s *grid.Sheet, row, col int) bool {
	for _, a := range s.Arrays {
		if a.Contains(row, col) {
			return true
		}
	}
	return false
}

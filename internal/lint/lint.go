// Package lint finds problems in a workbook: errors, broken references,
// formulas that break their column's pattern, totals that miss rows,
// numbers stored as text. It works on the format-neutral grid.
package lint

import (
	"fmt"
	"math"
	"regexp"
	"sort"
	"strings"

	"github.com/eemax/sheets/internal/analyze"
	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/input"
)

// Severity orders findings.
type Severity int

const (
	Note Severity = iota
	Warning
	Error
)

func (s Severity) String() string { return [...]string{"note", "warning", "error"}[s] }

// MarshalText makes severities read as words in JSON.
func (s Severity) MarshalText() ([]byte, error) { return []byte(s.String()), nil }

// ParseSeverity reads "note", "warning" or "error".
func ParseSeverity(s string) (Severity, bool) {
	for i, n := range []string{"note", "warning", "error"} {
		if strings.EqualFold(s, n) {
			return Severity(i), true
		}
	}
	return 0, false
}

// Finding is one problem, possibly covering several cells.
type Finding struct {
	Rule     string   `json:"rule"`
	Severity Severity `json:"severity"`
	Sheet    string   `json:"sheet,omitempty"`
	Cells    []string `json:"cells,omitempty"` // A1 refs on Sheet
	Message  string   `json:"message"`
	row, col int      // position for sorting
}

// Rules describes each rule, for documentation and --help.
var Rules = []struct{ Name, Severity, What string }{
	{"error-value", "error", "a cell shows an error value (#DIV/0!, #N/A, #VALUE!, ...)"},
	{"broken-reference", "error", "a formula or name refers to deleted cells (#REF!)"},
	{"unknown-function", "error", "a formula calls a function Excel does not have (#NAME?)"},
	{"inconsistent-formula", "warning", "a cell breaks the formula pattern of its column, or holds a typed-in value where the column computes"},
	{"total-range", "warning", "a total row formula leaves out rows of its table"},
	{"number-as-text", "warning", "numbers or dates stored as text in a number or date column (SUM ignores them)"},
	{"duplicate-header", "warning", "two columns of a table share a header"},
	{"merged-in-table", "warning", "merged cells inside a table's data (sorting and filtering break)"},
	{"mixed-types", "note", "a number column also holds text"},
	{"stray-whitespace", "note", "text with leading or trailing spaces (lookups and grouping miss it)"},
	{"repeated-constant", "note", "the same number typed into many formulas; a cell would make it visible and editable"},
	{"external-link", "note", "a formula refers to another workbook"},
	{"volatile-function", "note", "a formula uses a function that recalculates on every change (NOW, TODAY, RAND, OFFSET, INDIRECT)"},
}

// Options narrow the check.
type Options struct {
	Sheets []string // only these sheets; empty means all
}

// Check lints a workbook.
func Check(wb *grid.Workbook, opts Options) []Finding {
	var out []Finding
	isName := names(wb)
	for _, s := range wb.Sheets {
		if len(opts.Sheets) > 0 && !nameIn(s.Name, opts.Sheets) {
			continue
		}
		out = append(out, cells(wb, s, isName)...)
		for _, reg := range analyze.Regions(wb, s) {
			if reg.IsTable() {
				out = append(out, table(wb, s, reg)...)
			}
		}
		out = append(out, constants(s)...)
	}
	for _, n := range wb.Names {
		if strings.Contains(n.Ref, "#REF!") {
			out = append(out, Finding{Rule: "broken-reference", Severity: Error, Message: fmt.Sprintf("name %s refers to deleted cells: %s", n.Name, n.Ref), row: math.MaxInt})
		}
	}
	order := map[string]int{}
	for i, s := range wb.Sheets {
		order[s.Name] = i
	}
	sort.SliceStable(out, func(i, j int) bool {
		a, b := out[i], out[j]
		if a.Severity != b.Severity {
			return a.Severity > b.Severity
		}
		if order[a.Sheet] != order[b.Sheet] {
			return order[a.Sheet] < order[b.Sheet]
		}
		if a.row != b.row {
			return a.row < b.row
		}
		return a.col < b.col
	})
	return out
}

func nameIn(name string, list []string) bool {
	for _, n := range list {
		if strings.EqualFold(n, name) {
			return true
		}
	}
	return false
}

func names(wb *grid.Workbook) func(string) bool {
	set := map[string]bool{}
	for _, n := range wb.Names {
		set[strings.ToUpper(n.Name)] = true
	}
	return func(s string) bool { return set[strings.ToUpper(s)] }
}

func ref(row, col int) string { return grid.CellName(row, col) }

// uncertain groups #VALUE! results of the sheets calculator, which rejects
// number-like text that Excel converts.
const uncertain = "#VALUE?"

var volatile = map[string]bool{"NOW": true, "TODAY": true, "RAND": true, "RANDBETWEEN": true, "RANDARRAY": true, "OFFSET": true, "INDIRECT": true}

// cells checks each cell on its own: error values and formula problems.
func cells(wb *grid.Workbook, s *grid.Sheet, isName func(string) bool) []Finding {
	type group struct {
		cells    []string
		row, col int
	}
	errs := map[string]*group{}    // error code → cells
	unknown := map[string]*group{} // function → cells
	var refErr, external, vol group
	volFns := map[string]bool{}
	add := func(g *group, r, c int) {
		if len(g.cells) == 0 {
			g.row, g.col = r, c
		}
		g.cells = append(g.cells, ref(r, c))
	}
	s.Each(s.Used(), func(r, c int, raw *grid.Cell) {
		if raw.Formula != "" {
			f := raw.Formula
			if strings.Contains(f, "#REF!") {
				add(&refErr, r, c)
				return // its error value says nothing more
			}
			for _, fn := range formula.UnknownFunctions(f, isName) {
				if unknown[fn] == nil {
					unknown[fn] = &group{}
				}
				add(unknown[fn], r, c)
			}
			if externalRef.MatchString(f) {
				add(&external, r, c)
			}
			for _, fn := range formula.Calls(f) {
				if volatile[fn] {
					add(&vol, r, c)
					volFns[fn] = true
					break
				}
			}
		}
		v := wb.Value(s, r, c)
		if v != nil && v.Kind == grid.Error && !v.Pending {
			if raw.Formula != "" && len(formula.UnknownFunctions(raw.Formula, isName)) > 0 {
				return // reported as unknown-function
			}
			key := v.Text
			if v.Computed && v.Text == "#VALUE!" {
				key = uncertain // Excel converts number-like text that the calculator rejects
			}
			if errs[key] == nil {
				errs[key] = &group{}
			}
			add(errs[key], r, c)
		}
	})
	var out []Finding
	for code, g := range errs {
		verb := map[bool]string{true: "shows", false: "show"}[len(g.cells) == 1]
		if code == uncertain {
			out = append(out, Finding{Rule: "error-value", Severity: Warning, Sheet: s.Name, Cells: g.cells, row: g.row, col: g.col,
				Message: fmt.Sprintf("%s %s #VALUE! as computed by sheets; Excel may accept the inputs, but check for text where a number belongs", countOf(len(g.cells), "cell"), verb)})
			continue
		}
		out = append(out, Finding{Rule: "error-value", Severity: Error, Sheet: s.Name, Cells: g.cells, row: g.row, col: g.col,
			Message: fmt.Sprintf("%s %s %s", countOf(len(g.cells), "cell"), verb, code)})
	}
	if len(refErr.cells) > 0 {
		out = append(out, Finding{Rule: "broken-reference", Severity: Error, Sheet: s.Name, Cells: refErr.cells, row: refErr.row, col: refErr.col,
			Message: fmt.Sprintf("%s refer to deleted cells (#REF!)", countOf(len(refErr.cells), "formula"))})
	}
	for fn, g := range unknown {
		out = append(out, Finding{Rule: "unknown-function", Severity: Error, Sheet: s.Name, Cells: g.cells, row: g.row, col: g.col,
			Message: fmt.Sprintf("%s is not an Excel function (#NAME? in Excel)", fn)})
	}
	if len(external.cells) > 0 {
		out = append(out, Finding{Rule: "external-link", Severity: Note, Sheet: s.Name, Cells: external.cells, row: external.row, col: external.col,
			Message: fmt.Sprintf("%s refer to other workbooks; their values only update when those files are available", countOf(len(external.cells), "formula"))})
	}
	if len(vol.cells) > 0 {
		var fns []string
		for fn := range volFns {
			fns = append(fns, fn)
		}
		sort.Strings(fns)
		out = append(out, Finding{Rule: "volatile-function", Severity: Note, Sheet: s.Name, Cells: vol.cells, row: vol.row, col: vol.col,
			Message: fmt.Sprintf("%s use %s, which recalculate on every change", countOf(len(vol.cells), "formula"), strings.Join(fns, ", "))})
	}
	return out
}

// externalRef matches a reference into another workbook: [1]Sheet!A1 or
// '[book.xlsx]Sheet'!A1, but not a structured reference (Table1[Col]).
var externalRef = regexp.MustCompile(`\[[^\]]*\][^!+\-*/,()&=<>^\s"\[\]]*!`)

// table checks a table's columns.
func table(wb *grid.Workbook, s *grid.Sheet, reg analyze.Region) []Finding {
	var out []Finding
	cols := analyze.Profile(wb, s, reg)

	// Duplicate headers.
	seen := map[string][]string{}
	var order []string
	for _, c := range cols {
		key := strings.ToLower(strings.TrimSpace(c.Header))
		if key == "" {
			continue
		}
		if seen[key] == nil {
			order = append(order, key)
		}
		seen[key] = append(seen[key], ref(reg.HeaderRow, c.Col))
	}
	for _, key := range order {
		if refs := seen[key]; len(refs) > 1 {
			r, c, _ := grid.ParseCell(refs[0])
			out = append(out, Finding{Rule: "duplicate-header", Severity: Warning, Sheet: s.Name, Cells: refs, row: r, col: c,
				Message: fmt.Sprintf("header %q appears %d times; lookups and queries cannot tell the columns apart", analyze.Label(wb, s, r, c), len(refs))})
		}
	}

	// Merged cells in the data.
	for _, m := range s.Merges {
		if !m.Intersect(reg.Body).Empty() {
			out = append(out, Finding{Rule: "merged-in-table", Severity: Warning, Sheet: s.Name, Cells: []string{m.String()}, row: m.R1, col: m.C1,
				Message: "merged cells inside the table's data; sorting and filtering will fail or scramble rows"})
		}
	}

	for _, c := range cols {
		out = append(out, columnFormulas(s, reg, c)...)
		out = append(out, columnValues(wb, s, reg, c)...)
	}
	out = append(out, totals(s, reg)...)
	return out
}

// columnFormulas finds cells that break the dominant formula of a column.
func columnFormulas(s *grid.Sheet, reg analyze.Region, c analyze.Column) []Finding {
	type pat struct {
		display string
		n       int
	}
	pats := map[string]*pat{}
	formulas := 0
	for r := reg.Body.R1; r <= reg.Body.R2; r++ {
		cell := s.At(r, c.Col)
		if cell == nil || cell.Formula == "" {
			continue
		}
		formulas++
		key, disp := formula.Pattern(cell.Formula, r, c.Col)
		if pats[key] == nil {
			pats[key] = &pat{display: disp}
		}
		pats[key].n++
	}
	if formulas < 3 {
		return nil
	}
	var domKey string
	var dom *pat
	for k, p := range pats {
		if dom == nil || p.n > dom.n || p.n == dom.n && k < domKey {
			domKey, dom = k, p
		}
	}
	filled := 0
	for r := reg.Body.R1; r <= reg.Body.R2; r++ {
		if cell := s.At(r, c.Col); cell != nil && !cell.IsEmpty() {
			filled++
		}
	}
	// A pattern is established when most of the column follows it.
	if dom.n*2 <= filled || dom.n < 3 {
		return nil
	}
	var out []Finding
	for r := reg.Body.R1; r <= reg.Body.R2; r++ {
		cell := s.At(r, c.Col)
		if cell == nil || cell.IsEmpty() {
			continue
		}
		if cell.Formula == "" {
			if r == reg.Body.R1 && strings.Contains(dom.display, "{r-1}") {
				continue // the opening value of a running total
			}
			v := grid.Display(cell, grid.Format{}, false)
			out = append(out, Finding{Rule: "inconsistent-formula", Severity: Warning, Sheet: s.Name, Cells: []string{ref(r, c.Col)}, row: r, col: c.Col,
				Message: fmt.Sprintf("typed-in %s where the column computes =%s", quoteValue(v), dom.display)})
			continue
		}
		if key, _ := formula.Pattern(cell.Formula, r, c.Col); key != domKey {
			out = append(out, Finding{Rule: "inconsistent-formula", Severity: Warning, Sheet: s.Name, Cells: []string{ref(r, c.Col)}, row: r, col: c.Col,
				Message: fmt.Sprintf("=%s breaks the column pattern =%s", cell.Formula, dom.display)})
		}
	}
	return out
}

func quoteValue(v string) string {
	if _, err := fmt.Sscanf(v, "%g", new(float64)); err == nil {
		return v
	}
	return fmt.Sprintf("%q", grid.Truncate(v, 30))
}

// columnValues checks the values of a column: numbers stored as text,
// text among numbers, stray spaces.
func columnValues(wb *grid.Workbook, s *grid.Sheet, reg analyze.Region, c analyze.Column) []Finding {
	var asText, otherText, spaced []string
	var firstText, firstOther, firstSpace int
	numbers := 0
	for r := reg.Body.R1; r <= reg.Body.R2; r++ {
		cell := wb.Value(s, r, c.Col)
		if cell == nil || cell.Pending {
			continue
		}
		switch cell.Kind {
		case grid.Number:
			numbers++
		case grid.Text:
			if cell.Text != strings.TrimSpace(cell.Text) && strings.TrimSpace(cell.Text) != "" {
				if len(spaced) == 0 {
					firstSpace = r
				}
				spaced = append(spaced, ref(r, c.Col))
			}
			if cell.Formula != "" {
				continue
			}
			if v := input.Parse(strings.TrimSpace(cell.Text), false); v.Kind == grid.Number {
				if len(asText) == 0 {
					firstText = r
				}
				asText = append(asText, ref(r, c.Col))
			} else if strings.TrimSpace(cell.Text) != "" {
				if len(otherText) == 0 {
					firstOther = r
				}
				otherText = append(otherText, ref(r, c.Col))
			}
		}
	}
	var out []Finding
	label := c.Header
	if label == "" {
		label = grid.ColName(c.Col)
	}
	// Numbers stored as text matter in a column that is otherwise numbers.
	if len(asText) > 0 && numbers >= len(asText) {
		out = append(out, Finding{Rule: "number-as-text", Severity: Warning, Sheet: s.Name, Cells: asText, row: firstText, col: c.Col,
			Message: fmt.Sprintf("%s in %s %s stored as text; SUM and sorting ignore them (rewrite them with sheets set)", countOf(len(asText), "value"), label, map[bool]string{true: "is a number", false: "are numbers"}[len(asText) == 1])})
	}
	if len(otherText) > 0 && numbers > 0 && numbers >= 4*len(otherText) {
		v := s.Get(firstOther, c.Col).Text
		out = append(out, Finding{Rule: "mixed-types", Severity: Note, Sheet: s.Name, Cells: otherText, row: firstOther, col: c.Col,
			Message: fmt.Sprintf("%s holds text in %s (like %q) among numbers; calculations skip them", label, countOf(len(otherText), "cell"), grid.Truncate(v, 20))})
	}
	if len(spaced) > 0 {
		out = append(out, Finding{Rule: "stray-whitespace", Severity: Note, Sheet: s.Name, Cells: spaced, row: firstSpace, col: c.Col,
			Message: fmt.Sprintf("%s in %s %s leading or trailing spaces; lookups and grouping treat them as different values", countOf(len(spaced), "value"), label, map[bool]string{true: "has", false: "have"}[len(spaced) == 1])})
	}
	return out
}

// totals checks that total row formulas cover the whole table.
func totals(s *grid.Sheet, reg analyze.Region) []Finding {
	if reg.TotalRow == 0 {
		return nil
	}
	var out []Finding
	for col := reg.Range.C1; col <= reg.Range.C2; col++ {
		cell := s.At(reg.TotalRow, col)
		if cell == nil || cell.Formula == "" {
			continue
		}
		for _, r := range formula.Refs(cell.Formula) {
			if r.Sheet != "" && !strings.EqualFold(r.Sheet, s.Name) {
				continue
			}
			rg := r.Range
			if rg.C1 != col || rg.C2 != col || rg.Rows() < 2 || rg.R2 == grid.MaxRows {
				continue // not a run down this column
			}
			if rg.R1 > reg.Body.R2 || rg.R2 < reg.Body.R1 {
				continue
			}
			var missing []string
			if rg.R1 > reg.Body.R1 {
				missing = append(missing, rowSpan(reg.Body.R1, rg.R1-1))
			}
			if rg.R2 < reg.Body.R2 {
				missing = append(missing, rowSpan(rg.R2+1, reg.Body.R2))
			}
			if len(missing) == 0 {
				continue
			}
			body := grid.Range{R1: reg.Body.R1, C1: col, R2: reg.Body.R2, C2: col}
			out = append(out, Finding{Rule: "total-range", Severity: Warning, Sheet: s.Name, Cells: []string{ref(reg.TotalRow, col)}, row: reg.TotalRow, col: col,
				Message: fmt.Sprintf("total =%s covers %s but the column's data runs %s: %s left out", cell.Formula, rg, body, strings.Join(missing, " and "))})
		}
	}
	return out
}

func rowSpan(a, b int) string {
	if a == b {
		return fmt.Sprintf("row %d", a)
	}
	return fmt.Sprintf("rows %d-%d", a, b)
}

// constants finds the same number typed into many formulas.
func constants(s *grid.Sheet) []Finding {
	skip := map[float64]bool{0: true, 1: true, 2: true, 3: true, 4: true, 10: true, 12: true, 24: true, 60: true, 100: true, 365: true, 1000: true, 0.5: true}
	type use struct {
		cells    []string
		row, col int
	}
	uses := map[float64]*use{}
	s.Each(s.Used(), func(r, c int, cell *grid.Cell) {
		if cell.Formula == "" {
			return
		}
		seen := map[float64]bool{}
		for _, n := range formula.Numbers(cell.Formula) {
			if skip[math.Abs(n)] || seen[n] {
				continue
			}
			seen[n] = true
			if uses[n] == nil {
				uses[n] = &use{row: r, col: c}
			}
			uses[n].cells = append(uses[n].cells, ref(r, c))
		}
	})
	var out []Finding
	for n, u := range uses {
		// A filled-down column repeats its constant by design; count the
		// columns it appears in.
		cols := map[string]bool{}
		for _, a := range u.cells {
			_, c, _ := grid.ParseCell(a)
			cols[grid.ColName(c)] = true
		}
		if len(cols) < 2 && len(u.cells) < 10 {
			continue
		}
		if len(u.cells) < 3 {
			continue
		}
		out = append(out, Finding{Rule: "repeated-constant", Severity: Note, Sheet: s.Name, Cells: u.cells, row: u.row, col: u.col,
			Message: fmt.Sprintf("%s is typed into %d formulas; put it in a cell and refer to that, so it is visible and changes in one place", grid.FormatNumber(n), len(u.cells))})
	}
	return out
}

func countOf(n int, word string) string {
	if n == 1 {
		return "1 " + word
	}
	return fmt.Sprintf("%d %ss", n, word)
}

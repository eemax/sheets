package fuzz

import (
	"fmt"
	"regexp"
	"sort"
	"strconv"
	"strings"

	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/input"
	"github.com/eemax/sheets/internal/words"
)

// The model of where formulas go. checkKinds judges the formulas a file
// holds, one by one, against their markers; it cannot see a formula an edit
// loses, or one it makes out of nothing. This model follows every formula
// cell through the edit: it starts from the cells that hold formulas before
// the edit, applies what the edit does to them (the model of kinds, in
// check.go), and the result must be what the saved file holds, cell for cell
// and kind for kind (the marker of each formula).
//
//   - set, write, and a script's setValue and setFormula: a cell takes a
//     formula when the value typed is one, else it holds none;
//   - clear: the cells hold none;
//   - copy: each cell of the target holds what its source cell did, a
//     formula of the same kind or none (--values: none, --format: unchanged);
//   - move: the same, and the source cells it empties hold none;
//   - rows and columns: formulas go where their cells go, and those in the
//     rows or columns deleted are gone;
//   - sheet rename and copy: the formulas stay, and the copy holds the same;
//   - replace: no cell gains or loses a formula;
//   - sort: no column gains or loses a formula of a kind (the rows move, and
//     the model does not say where). A sheet sorted in a batch or script,
//     and edited at all afterwards, is not judged.
//
// A do batch or a script is the model of each of its commands, in order.
// An edit the model cannot read (a command it has no rule for) is not
// judged.
type model struct {
	sheets map[string]map[[2]int]string // formula cells by sheet (lower case): the marker of each, "" for none
	names  map[string]string            // the sheet's name as last given
	order  []string
	sorted map[string]bool // positions on the sheet are unknown after a sort; only the count of each kind per column holds
	lost   map[string]bool // the model gave up on the sheet
	// why is what the model cannot read, when it cannot.
	why string
}

func newModel(wb *grid.Workbook) *model {
	m := &model{sheets: map[string]map[[2]int]string{}, names: map[string]string{}, sorted: map[string]bool{}, lost: map[string]bool{}}
	for _, s := range wb.Sheets {
		k := strings.ToLower(s.Name)
		m.names[k] = s.Name
		m.order = append(m.order, k)
		cells := map[[2]int]string{}
		if u := s.Used(); u.R2 >= 1 {
			s.Each(u, func(r, c int, cell *grid.Cell) {
				if cell.Formula != "" {
					cells[[2]int{r, c}] = marker(cell.Formula)
				}
			})
		}
		m.sheets[k] = cells
	}
	return m
}

func (m *model) isSheet(name string) bool {
	_, ok := m.sheets[strings.ToLower(name)]
	return ok
}

// ensure makes the sheet if it is new (a command that names one makes it).
func (m *model) ensure(name string) string {
	k := strings.ToLower(name)
	if _, ok := m.sheets[k]; !ok {
		m.sheets[k] = map[[2]int]string{}
		m.names[k] = name
		m.order = append(m.order, k)
	}
	return k
}

// target resolves a reference to a sheet (made when new) and a range.
func (m *model) target(ref, inSheet string) (string, grid.Range, error) {
	p, err := grid.ParseRef(ref, m.isSheet)
	if err != nil {
		return "", grid.Range{}, err
	}
	name := p.Sheet
	if name == "" {
		name = inSheet
	}
	if name == "" {
		return "", grid.Range{}, fmt.Errorf("%q names no sheet", ref)
	}
	var rg grid.Range
	if p.Range != nil {
		rg = *p.Range
	} else {
		rg = grid.Range{R1: 1, C1: 1, R2: grid.MaxRows, C2: grid.MaxCols}
	}
	return m.ensure(name), rg, nil
}

// touch tells the model a sheet is edited in a way it follows: a sheet that
// was sorted no longer has known positions.
func (m *model) touch(k string) {
	if m.sorted[k] {
		m.lost[k] = true
	}
}

func (m *model) put(k string, r, c int, formula bool, mk string) {
	m.touch(k)
	if formula {
		m.sheets[k][[2]int{r, c}] = mk
	} else {
		delete(m.sheets[k], [2]int{r, c})
	}
}

func (m *model) clearRect(k string, rg grid.Range) {
	m.touch(k)
	for cell := range m.sheets[k] {
		if rg.Contains(cell[0], cell[1]) {
			delete(m.sheets[k], cell)
		}
	}
}

type placed struct {
	dr, dc int
	f      bool
	mk     string
}

// read takes the formulas and non-formulas of a range, relative to its
// corner, before anything is written.
func (m *model) read(k string, rg grid.Range) []placed {
	m.touch(k)
	var out []placed
	for r := rg.R1; r <= rg.R2; r++ {
		for c := rg.C1; c <= rg.C2; c++ {
			mk, ok := m.sheets[k][[2]int{r, c}]
			out = append(out, placed{r - rg.R1, c - rg.C1, ok, mk})
		}
	}
	return out
}

// bounded is false for a range of whole rows or columns, which the model
// does not unroll.
func bounded(rg grid.Range) bool { return rg.R2 < grid.MaxRows && rg.C2 < grid.MaxCols }

// command applies one command, as typed without sheets and the file.
func (m *model) command(w []string, stdin string) {
	if m.why != "" || len(w) == 0 {
		return
	}
	fail := func(format string, a ...any) { m.why = fmt.Sprintf(format, a...) }
	switch w[0] {
	case "set":
		if len(w) < 3 || len(w)%2 == 0 {
			fail("set: %v", w)
			return
		}
		for i := 1; i+1 < len(w); i += 2 {
			k, rg, err := m.target(w[i], "")
			if err != nil || !bounded(rg) {
				fail("set %s: %v", w[i], err)
				return
			}
			v := input.Parse(w[i+1], false)
			for r := rg.R1; r <= rg.R2; r++ {
				for c := rg.C1; c <= rg.C2; c++ {
					m.put(k, r, c, v.Formula != "", marker(v.Formula))
				}
			}
		}
	case "write":
		if len(w) < 2 {
			fail("write: %v", w)
			return
		}
		k, rg, err := m.target(w[1], "")
		if err != nil {
			fail("write %s: %v", w[1], err)
			return
		}
		for i, line := range strings.Split(strings.TrimSuffix(stdin, "\n"), "\n") {
			for j, text := range strings.Split(line, "\t") {
				if text == "" {
					continue
				}
				v := input.Parse(text, false)
				m.put(k, rg.R1+i, rg.C1+j, v.Formula != "", marker(v.Formula))
			}
		}
	case "clear":
		if len(w) < 2 {
			fail("clear: %v", w)
			return
		}
		k, rg, err := m.target(w[1], "")
		if err != nil {
			fail("clear %s: %v", w[1], err)
			return
		}
		m.clearRect(k, rg)
	case "copy", "move":
		if len(w) < 3 {
			fail("%s: %v", w[0], w)
			return
		}
		values, format := false, false
		for _, a := range w[3:] {
			values = values || a == "--values"
			format = format || a == "--format"
		}
		src, from, err := m.target(w[1], "")
		if err != nil || !bounded(from) {
			fail("%s %s: %v", w[0], w[1], err)
			return
		}
		dst, to, err := m.target(w[2], m.names[src])
		if err != nil || !bounded(to) {
			fail("%s %s: %v", w[0], w[2], err)
			return
		}
		if format {
			return
		}
		cells := m.read(src, from)
		if w[0] == "move" {
			m.clearRect(src, from)
		}
		// The target is the source's size from its corner; a larger one is
		// tiled, which the generator never asks for.
		if to.Rows() > 1 || to.Cols() > 1 {
			fail("%s to a range %s", w[0], w[2])
			return
		}
		for _, p := range cells {
			m.put(dst, to.R1+p.dr, to.C1+p.dc, p.f && !values, p.mk)
		}
	case "sort":
		if len(w) < 2 {
			fail("sort: %v", w)
			return
		}
		name := w[1]
		if !m.isSheet(name) {
			p, err := grid.ParseRef(name, m.isSheet)
			if err != nil || p.Sheet == "" {
				fail("sort %s: %v", w[1], err)
				return
			}
			name = p.Sheet
		}
		m.sort(m.ensure(name))
	case "replace":
	case "rows", "cols":
		if len(w) < 3 {
			fail("%s: %v", w[0], w)
			return
		}
		n := 1
		for i, a := range w {
			if a == "-n" && i+1 < len(w) {
				n, _ = strconv.Atoi(w[i+1])
			}
		}
		k, rg, err := m.target(w[2], "")
		if err != nil {
			fail("%s %s: %v", w[0], w[2], err)
			return
		}
		from, to := rg.R1, rg.R2
		if w[0] == "cols" {
			from, to = rg.C1, rg.C2
		}
		switch w[1] {
		case "insert":
			m.shift(k, w[0] == "cols", from, 0, n)
		case "delete":
			m.shift(k, w[0] == "cols", from, to, 0)
		default:
			fail("%s %s", w[0], w[1])
		}
	case "sheet":
		if len(w) < 4 {
			fail("sheet: %v", w)
			return
		}
		from, to := strings.ToLower(w[2]), w[3]
		if !m.isSheet(from) {
			fail("sheet %s %s: no sheet", w[1], w[2])
			return
		}
		switch w[1] {
		case "rename":
			m.sheets[strings.ToLower(to)] = m.sheets[from]
			m.names[strings.ToLower(to)] = to
			if m.sorted[from] {
				m.sorted[strings.ToLower(to)] = true
			}
			for i, k := range m.order {
				if k == from {
					m.order[i] = strings.ToLower(to)
				}
			}
			delete(m.sheets, from)
			delete(m.names, from)
		case "copy":
			k := m.ensure(to)
			for cell, mk := range m.sheets[from] {
				m.sheets[k][cell] = mk
			}
			m.sorted[k] = m.sorted[from]
			m.lost[k] = m.lost[from]
		default:
			fail("sheet %s", w[1])
		}
	default:
		fail("no rule for %s", w[0])
	}
}

// sort: the rows move, and where is not modelled.
func (m *model) sort(k string) {
	m.touch(k)
	m.sorted[k] = true
}

// shift inserts n rows or columns before row (column) at, or deletes from..to.
func (m *model) shift(k string, cols bool, at, to, n int) {
	m.touch(k)
	next := map[[2]int]string{}
	i := 0
	if cols {
		i = 1
	}
	for cell, mk := range m.sheets[k] {
		p := cell[i]
		switch {
		case n > 0 && p >= at:
			p += n
		case n == 0 && p > to:
			p -= to - at + 1
		case n == 0 && p >= at:
			continue
		}
		cell[i] = p
		next[cell] = mk
	}
	m.sheets[k] = next
}

var (
	scriptRange = `"((?:[^"\\]|\\.)*)"`
	scriptCall  = regexp.MustCompile(`^ws\.getRange\(` + scriptRange + `\)\.(\w+)\((.*)\);$`)
	scriptSort  = regexp.MustCompile(`^ws\.getRange\(` + scriptRange + `\)\.getSort\(\)\.apply\(`)
	scriptArg   = regexp.MustCompile(`^` + scriptRange + `$`)
)

// script applies the statements of a script made by the fuzzer: the first
// names the sheet, and each other is one range operation on it.
func (m *model) script(code string) {
	lines := strings.Split(strings.TrimSuffix(code, "\n"), "\n")
	sheet := ""
	if len(lines) > 0 {
		if _, rest, ok := strings.Cut(lines[0], "workbook.getWorksheet("); ok {
			if q, err := strconv.Unquote(strings.TrimSuffix(rest, ");")); err == nil {
				sheet = q
			}
		}
	}
	if sheet == "" || !m.isSheet(sheet) {
		m.why = "script: no sheet named in " + lines[0]
		return
	}
	k := strings.ToLower(sheet)
	unq := func(s string) string { u, _ := strconv.Unquote(`"` + s + `"`); return u }
	for _, line := range lines[1:] {
		if m.why != "" {
			return
		}
		if sm := scriptSort.FindStringSubmatch(line); sm != nil {
			m.sort(k)
			continue
		}
		sm := scriptCall.FindStringSubmatch(line)
		if sm == nil {
			m.why = "script: " + line
			return
		}
		rg, err := grid.ParseRange(unq(sm[1]))
		if err != nil {
			m.why = "script: " + line
			return
		}
		arg := ""
		if a := scriptArg.FindStringSubmatch(sm[3]); a != nil {
			arg = unq(a[1])
		}
		switch sm[2] {
		case "setValue":
			m.put(k, rg.R1, rg.C1, false, "")
		case "setFormula":
			m.put(k, rg.R1, rg.C1, true, marker(arg))
		case "clear":
			m.clearRect(k, rg)
		case "copyFrom", "moveTo":
			from, to := rg, rg
			if sm[2] == "copyFrom" {
				var err error
				if from, err = grid.ParseRange(arg); err != nil {
					m.why = "script: " + line
					return
				}
				to = rg
			} else {
				var err error
				if to, err = grid.ParseRange(arg); err != nil {
					m.why = "script: " + line
					return
				}
			}
			cells := m.read(k, from)
			if sm[2] == "moveTo" {
				m.clearRect(k, from)
			}
			for _, p := range cells {
				m.put(k, to.R1+p.dr, to.C1+p.dc, p.f, p.mk)
			}
		case "insert", "delete":
			cols := arg == "Right" || arg == "Left"
			from, to := rg.R1, rg.R2
			if cols {
				from, to = rg.C1, rg.C2
			}
			if sm[2] == "insert" {
				m.shift(k, cols, from, 0, to-from+1)
			} else {
				m.shift(k, cols, from, to, 0)
			}
		default:
			m.why = "script: " + line
		}
	}
}

// apply applies an edit.
func (m *model) apply(e Edit) {
	switch e.Class {
	case "do":
		for _, line := range strings.Split(e.Stdin, "\n") {
			if strings.TrimSpace(line) == "" {
				continue
			}
			w, err := words.Texts(line)
			if err != nil {
				m.why = "do: " + err.Error()
				return
			}
			m.command(w, "")
		}
	case "script":
		m.script(e.Args[len(e.Args)-1])
	default:
		var w []string
		for _, a := range e.Args {
			if a != FileArg {
				w = append(w, a)
			}
		}
		m.command(w, e.Stdin)
	}
}

// check compares the model with the workbook the edit saved, and returns the
// first difference, "" when there is none.
func (m *model) check(got *grid.Workbook) string {
	if m.why != "" {
		return ""
	}
	have := newModel(got)
	got0 := got
	var keys []string
	for k := range m.sheets {
		keys = append(keys, k)
	}
	sort.Strings(keys)
	for _, k := range keys {
		if _, ok := have.sheets[k]; !ok {
			return fmt.Sprintf("sheet %s is gone", m.names[k])
		}
	}
	for _, k := range keys {
		if m.lost[k] {
			continue
		}
		want, got := m.sheets[k], have.sheets[k]
		if m.sorted[k] {
			if d := byColumn(want, got); d != "" {
				return fmt.Sprintf("%s: %s", m.names[k], d)
			}
			continue
		}
		cells := map[[2]int]bool{}
		for c := range want {
			cells[c] = true
		}
		for c := range got {
			cells[c] = true
		}
		var order [][2]int
		for c := range cells {
			order = append(order, c)
		}
		sort.Slice(order, func(i, j int) bool {
			if order[i][0] != order[j][0] {
				return order[i][0] < order[j][0]
			}
			return order[i][1] < order[j][1]
		})
		for _, c := range order {
			where := m.names[k] + "!" + grid.CellName(c[0], c[1])
			wmk, wok := want[c]
			gmk, gok := got[c]
			switch {
			case wok && !gok && tableHeader(got0, k, c):
				// A table's header is text: sheets names the column after a
				// formula typed there, as Excel does, and keeps no formula.
			case wok && !gok:
				return fmt.Sprintf("%s: the edit should leave %s there, and it holds no formula (it was lost)", where, formulaName(wmk))
			case !wok && gok:
				return fmt.Sprintf("%s: the edit should leave no formula there, and it holds %s", where, formulaName(gmk))
			case wmk != "" && gmk != "" && wmk != gmk:
				return fmt.Sprintf("%s: the edit should leave %s there, and it holds %s", where, formulaName(wmk), formulaName(gmk))
			}
		}
	}
	return ""
}

// tableHeader tells whether a cell is in the header row of an Excel table.
func tableHeader(wb *grid.Workbook, sheet string, cell [2]int) bool {
	for _, s := range wb.Sheets {
		if strings.ToLower(s.Name) != sheet {
			continue
		}
		for _, t := range s.Tables {
			if t.Range.R1 == cell[0] && t.Range.C1 <= cell[1] && cell[1] <= t.Range.C2 {
				return true
			}
		}
	}
	return false
}

func formulaName(mk string) string {
	if n, ok := kindName[mk]; ok {
		return "a " + n
	}
	return "a formula"
}

// byColumn compares, for each column, how many formulas of each kind there
// are: a sort moves rows, so the cells of a column only change places.
func byColumn(want, got map[[2]int]string) string {
	count := func(cells map[[2]int]string) map[[2]interface{}]int {
		out := map[[2]interface{}]int{}
		for c, mk := range cells {
			out[[2]interface{}{c[1], mk}]++
		}
		return out
	}
	w, g := count(want), count(got)
	var keys [][2]interface{}
	for k := range w {
		keys = append(keys, k)
	}
	for k := range g {
		if _, ok := w[k]; !ok {
			keys = append(keys, k)
		}
	}
	sort.Slice(keys, func(i, j int) bool {
		if keys[i][0].(int) != keys[j][0].(int) {
			return keys[i][0].(int) < keys[j][0].(int)
		}
		return keys[i][1].(string) < keys[j][1].(string)
	})
	for _, k := range keys {
		if w[k] != g[k] {
			return fmt.Sprintf("after a sort column %s holds %d formulas of %s, it held %d", grid.ColName(k[0].(int)), g[k], formulaName(k[1].(string)), w[k])
		}
	}
	return ""
}

package fuzz

import (
	"fmt"
	"math/rand/v2"
	"regexp"
	"slices"
	"strconv"
	"strings"

	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/words"
)

// Classes are the edit classes, with their weights.
var Classes = []struct {
	Name   string
	Weight int
}{
	{"set", 12}, {"write", 6}, {"clear", 7}, {"copy", 12}, {"move", 12}, {"sort", 8}, {"replace", 6},
	{"rows-insert", 6}, {"rows-delete", 6}, {"cols-insert", 4}, {"cols-delete", 4},
	{"sheet-rename", 3}, {"sheet-copy", 3}, {"sheet-add", 2}, {"do", 6}, {"script", 6},
}

// streamBoost multiplies the weight of the classes the streaming editor does
// itself (see streamable) while the workbook is one it can take them on: the
// other classes are full-editor code in both copies.
var streamBoost = map[string]int{"set": 3, "write": 2, "rows-insert": 3, "rows-delete": 3, "script": 2}

// The formulas edits add, by the kind Excel 365 stores each typed in: {r}
// is the row of its cell, {s} a sheet of the workbook. Each carries its kind
// as the always-true condition of an IF, IF(N("p")=0,...), which keeps the
// type of the result: numbers, texts, empty texts and booleans are all here
// (a marker added to the result would turn the others into errors).
var (
	plainFormulas = marked(Plain,
		`B{r}*2`, `SUM(B2:B6)`, `IF(B{r}>3,1,0)`, `MAX(C2:C5)`,
		`MIN(B2:D2)`, `{s}!B2+1`, `COUNTIF(B2:B9,">3")`, `ROUND(C{r}*1.1,2)`,
		// texts
		`A{r}&"-x"`, `IF(B{r}>3,"hi","lo")`, `UPPER(A{r})`, `IF(A{r}="","none",A{r})`, `{s}!A2&"!"`,
		// empty texts
		`IF(B{r}>1000,1,"")`, `IF(A{r}="Fig","x","")`, `IF(F{r},"yes","")`, `IF(B{r}>3,"big","")`,
		// booleans
		`B{r}>3`, `AND(B{r}>3,C{r}<50)`, `ISNUMBER(B{r})`, `NOT(F{r})`, `OR(F{r},D{r}>250)`,
	)
	dynamicFormulas = marked(Dynamic,
		`SEQUENCE(2)`, `B2:B4*2`, `SUM(IF(B2:B6>3,B2:B6))`, `SEQUENCE(1,2)`,
		`MAX(B2:B6-D2:D6)`,
		// texts and booleans in an array
		`UPPER(A2:A4)`, `A2:A4&"!"`, `B2:B5>3`,
	)
	// What set, write and scripts type into a cell. The second half is text
	// the writers must escape or quote: &, < and > and quotes (XML), spaces at
	// the ends (xml:space="preserve"), what Excel reads as an escape
	// (_x0041_), and non-ASCII. The streaming editor's own text writer
	// (xlsx/stream_sheet.go newCell) is reached with them, and the full
	// editor's. None is a number, a date, a boolean or a formula.
	values = []string{"5", "-3", "12.5", "0", "Apple", "N/A", "Acme Inc.", "2026-07-01", "TRUE", "hello", "1,200",
		"R&D", "<b>x</b>", "5 > 3 & 2 < 4", `say "hi"`, "it's", " lead", "trail ", "two  spaces", "_x0041_", "A_x0020_B", "_x005F_x0041_", "Åäö 日本語 😀", "&amp;", "Müller & Söhne"}
	// Text only a set or a script types, since a batch is a line to a command
	// and a write a row of a tab-separated table: a line break, a tab and a
	// control character, which XML cannot hold (Excel stores it as _x0001_).
	oddValues = []string{"line1\nline2", "tab\there", "ctl\x01x", "<\x1f&>"}
	finds     = []string{"Apple", "N/A", "Fig", "Total", "5", "apple", "Acme Inc.", "Pear", "&", "<", `"`, "'", " ", "R&D", "Åäö"}
	replaces  = []string{"Apple", "Pear", "", "1,200", "7", "Acme Corp", "2026-01-31", "&", "<b>", `"q"`, "'", " ", "_x0041_", "Ä & Ö"}
	// Replacing in formula text renames functions only, which keeps kinds.
	formulaReplaces = [][2]string{{"SUM(", "PRODUCT("}, {"PRODUCT(", "SUM("}, {"MIN(", "MAX("}, {"MAX(", "MIN("}}
	// New sheets' names: some need escaping in the workbook's parts
	// (workbook.xml's attribute, the relationships, the quotes of a formula).
	newNames = []string{"Copy", "Data 2", "Q 2027", "Old", "Backup", "Plan", "R&D 2", "Bob's", "A<B", `Q1 "draft"`}
)

// marked puts each formula in an IF named by its kind's marker.
func marked(kind string, formulas ...string) []string {
	out := make([]string, len(formulas))
	for i, f := range formulas {
		out[i] = `=IF(N("` + kind + `")=0,` + f + `)`
	}
	return out
}

type gen struct {
	rng    *rand.Rand
	wb     *grid.Workbook // the workbook, and the sheets this edit makes (a copy: the session's is not changed)
	stream bool           // the workbook is one the streaming editor takes edits on (see streamable)
	fresh  bool           // the edit makes a new sheet (set, copy)
	made   *grid.Sheet    // the sheet the last edit made, until a batch picks it up
	focus  *grid.Sheet    // the sheet a batch edits: the one it made
	batch  bool           // the edit is a line of a do batch: one line, no tab or line break in what it types
}

// generate makes a random edit for a workbook.
func generate(rng *rand.Rand, wb *grid.Workbook) Edit {
	cp := *wb
	cp.Sheets = slices.Clone(wb.Sheets)
	g := &gen{rng: rng, wb: &cp, stream: streamable(wb)}
	return g.edit(g.class(true))
}

// streamable tells whether the streaming editor can take a workbook's edits
// itself. It does only set and the insert and delete of rows, and hands
// over to the full editor for those as soon as the workbook holds a dynamic
// array, and for rows when it holds a legacy array over several cells
// (xlsx/stream.go: anySpills, anyLegacy). Every other edit is full-editor
// code in both copies.
func streamable(wb *grid.Workbook) bool {
	for _, s := range wb.Sheets {
		if len(s.Spills) > 0 {
			return false
		}
		for _, a := range s.Arrays {
			if a.Rows()*a.Cols() > 1 {
				return false
			}
		}
	}
	return true
}

func (g *gen) weight(name string, w int) int {
	if g.stream {
		if b := streamBoost[name]; b > 0 {
			return w * b
		}
	}
	return w
}

func (g *gen) class(all bool) string {
	skip := func(name string) bool { return !all && (name == "do" || name == "script" || name == "write") }
	total := 0
	for _, c := range Classes {
		if !skip(c.Name) {
			total += g.weight(c.Name, c.Weight)
		}
	}
	n := g.rng.IntN(total)
	for _, c := range Classes {
		if skip(c.Name) {
			continue
		}
		if n -= g.weight(c.Name, c.Weight); n < 0 {
			return c.Name
		}
	}
	panic("unreachable")
}

func (g *gen) chance(p float64) bool { return g.rng.Float64() < p }

func pick[T any](g *gen, xs []T) T { return xs[g.rng.IntN(len(xs))] }

// sheet picks a sheet: in a batch, mostly the one it made.
func (g *gen) sheet() *grid.Sheet {
	if g.focus != nil && g.chance(0.7) {
		return g.focus
	}
	return pick(g, g.wb.Sheets)
}

// addSheet makes a sheet of the workbook the edit is made for: a copy of
// from, or an empty one.
func (g *gen) addSheet(name string, from *grid.Sheet) *grid.Sheet {
	s := grid.NewSheet(name)
	if from != nil {
		cp := *from
		cp.Name = name
		s = &cp
	}
	g.wb.Sheets = append(g.wb.Sheets, s)
	g.made = s
	return s
}

// renamed makes the rename visible to the edits after it in a batch.
func (g *gen) renamed(s *grid.Sheet, name string) {
	cp := *s
	cp.Name = name
	for i, o := range g.wb.Sheets {
		if o == s {
			g.wb.Sheets[i] = &cp
		}
	}
	if g.focus == s {
		g.focus = &cp
	}
}

// target picks the sheet of a set, write or copy: now and then one the edit
// makes (these commands make the sheet they name when it is new).
func (g *gen) target(s *grid.Sheet) *grid.Sheet {
	if g.fresh || g.chance(0.06) {
		return g.addSheet(g.newName(), nil)
	}
	return s
}

// used is a sheet's used range, at least A1; a sheet with nothing in it is
// taken as a few rows and columns, so edits on a new sheet land in several
// rows.
func used(s *grid.Sheet) grid.Range {
	u := s.Used()
	if u.R2 < 1 {
		return grid.Range{R1: 1, C1: 1, R2: 8, C2: 5}
	}
	return u
}

// spot picks a cell: often a formula's, else one in or near the used
// range. A formula is picked by its group first (plain, plain reading a
// range, legacy array, dynamic array), so the few arrays and range
// formulas are picked as often as the many filled-down ones.
func (g *gen) spot(s *grid.Sheet) (int, int) {
	if g.chance(0.4) {
		kind := map[[2]int]string{}
		for _, a := range s.Arrays {
			kind[[2]int{a.R1, a.C1}] = Legacy
		}
		for _, a := range s.Spills {
			kind[[2]int{a.R1, a.C1}] = Dynamic
		}
		groups := map[string][][2]int{}
		s.Each(used(s), func(r, c int, cell *grid.Cell) {
			if cell.Formula == "" {
				return
			}
			k := kind[[2]int{r, c}]
			if k == "" {
				k = Plain
				if strings.Contains(cell.Formula, ":") {
					k = "range"
				}
			}
			groups[k] = append(groups[k], [2]int{r, c})
		})
		var names []string
		for _, k := range []string{Plain, "range", Legacy, Dynamic} {
			if len(groups[k]) > 0 {
				names = append(names, k)
			}
		}
		if len(names) > 0 {
			p := pick(g, groups[pick(g, names)])
			return p[0], p[1]
		}
	}
	u := used(s)
	return 1 + g.rng.IntN(u.R2+2), 1 + g.rng.IntN(u.C2+2)
}

// area picks a range starting at a spot, mostly small.
func (g *gen) area(s *grid.Sheet) grid.Range {
	r, c := g.spot(s)
	h, w := 1, 1
	if g.chance(0.6) {
		h, w = 1+g.rng.IntN(4), 1+g.rng.IntN(3)
	}
	return grid.Range{R1: r, C1: c, R2: r + h - 1, C2: c + w - 1}
}

// The references in a formula's text: an optional sheet, a cell or a range.
var refRE = regexp.MustCompile(`(?:(?:'((?:[^']|'')+)'|([A-Za-z0-9_]+))!)?\$?([A-Z]{1,3})\$?([0-9]+)(?::\$?([A-Z]{1,3})\$?([0-9]+))?`)

// aggregate functions give one value from a range: a formula that has none
// of them and reads a range is one that gives an array where one value goes.
var aggregate = regexp.MustCompile(`\b(SUM|MAX|MIN|COUNTIF|PRODUCT)\(`)

// readArea picks cells some formula reads: the range the formula names, or
// part of it. Moving or copying the cells a formula reads rewrites the
// formula (its references follow), which is where a formula can change its
// kind; random cells rarely are those. Formulas that read a range where one
// value goes are picked first, then any. It returns false when no formula
// reads a cell.
func (g *gen) readArea() (*grid.Sheet, grid.Range, bool) {
	type read struct {
		s  *grid.Sheet
		rg grid.Range
	}
	var arrayish, all []read
	for _, s := range g.wb.Sheets {
		if u := s.Used(); u.R2 < 1 {
			continue
		}
		s.Each(s.Used(), func(_, _ int, c *grid.Cell) {
			if c.Formula == "" {
				return
			}
			for _, m := range refRE.FindAllStringSubmatch(c.Formula, -1) {
				at := s
				if name := m[1] + m[2]; name != "" {
					if at = g.wb.Sheet(strings.ReplaceAll(m[1], "''", "'") + m[2]); at == nil {
						continue
					}
				}
				ref := m[3] + m[4]
				if m[5] != "" {
					ref += ":" + m[5] + m[6]
				}
				rg, err := grid.ParseRange(ref)
				if err != nil {
					continue
				}
				all = append(all, read{at, rg})
				if m[5] != "" && !aggregate.MatchString(c.Formula) {
					arrayish = append(arrayish, read{at, rg})
				}
			}
		})
	}
	from := all
	if len(arrayish) > 0 && g.chance(0.6) {
		from = arrayish
	}
	if len(from) == 0 {
		return nil, grid.Range{}, false
	}
	r := pick(g, from)
	rg := r.rg
	if g.chance(0.4) && (rg.Rows() > 1 || rg.Cols() > 1) {
		// Part of it.
		rg.R1 += g.rng.IntN(rg.Rows())
		rg.R2 = min(rg.R2, rg.R1+g.rng.IntN(3))
		rg.C1 += g.rng.IntN(rg.Cols())
		rg.C2 = min(rg.C2, rg.C1+g.rng.IntN(2))
	}
	return r.s, rg, true
}

func ref(s *grid.Sheet, rg grid.Range) string { return s.Name + "!" + rg.String() }

func cell(s *grid.Sheet, r, c int) string { return s.Name + "!" + grid.CellName(r, c) }

// formulaSheet is how a formula names a sheet.
func formulaSheet(name string) string {
	for _, r := range name {
		if !(r >= 'A' && r <= 'Z' || r >= 'a' && r <= 'z' || r >= '0' && r <= '9' || r == '_') {
			return "'" + strings.ReplaceAll(name, "'", "''") + "'"
		}
	}
	return name
}

// text picks a value to type: odd lets in the texts with a line break, a tab
// or a control character (see oddValues).
func (g *gen) text(odd bool) string {
	if odd && g.chance(0.25) {
		return pick(g, oddValues)
	}
	return pick(g, values)
}

// value picks what set or write puts in a cell: a value or a formula; row is
// the row of the cell it is typed in, which fills in {r}: always, since a
// formula with a literal {r} in it is none Excel can read, and sheets accepts
// it. For a range, the cell is its first: sheets fills the rest of the range
// with the formula, its relative references shifted, as Excel does.
func (g *gen) value(row int, odd bool) string {
	var v string
	switch n := g.rng.IntN(10); {
	case n < 4:
		v = g.text(odd)
	case n < 8 || g.stream && g.chance(0.85):
		// A dynamic array makes the workbook one the streaming editor hands
		// over: in one it does not, they are rarer.
		v = pick(g, plainFormulas)
	default:
		v = pick(g, dynamicFormulas)
	}
	v = strings.ReplaceAll(v, "{s}", formulaSheet(g.sheet().Name))
	return strings.ReplaceAll(v, "{r}", strconv.Itoa(row))
}

func (g *gen) newName() string {
	base := pick(g, newNames)
	name := base
	for i := 2; g.wb.Sheet(name) != nil; i++ {
		name = fmt.Sprintf("%s %d", base, i)
	}
	return name
}

func (g *gen) edit(class string) Edit {
	e := Edit{Class: class}
	s := g.sheet()
	u := used(s)
	switch class {
	case "set":
		e.Args = []string{"set", FileArg}
		s = g.target(s)
		u = used(s)
		if g.chance(0.3) {
			rg := g.area(s)
			rg.R2 = rg.R1 + g.rng.IntN(4)
			e.Args = append(e.Args, ref(s, rg), g.value(rg.R1, !g.batch))
			break
		}
		for range 1 + g.rng.IntN(3) {
			r, c := g.spot(s)
			e.Args = append(e.Args, cell(s, r, c), g.value(r, !g.batch))
		}
	case "write":
		s = g.target(s)
		r, c := g.spot(s)
		var b strings.Builder
		for i := range 1 + g.rng.IntN(3) {
			var row []string
			for range 1 + g.rng.IntN(3) {
				row = append(row, g.value(r+i, false))
			}
			b.WriteString(strings.Join(row, "\t") + "\n")
		}
		e.Args, e.Stdin = []string{"write", FileArg, cell(s, r, c), "--from", "tsv"}, b.String()
	case "clear":
		e.Args = []string{"clear", FileArg, ref(s, g.area(s))}
		if g.chance(0.2) {
			e.Args = append(e.Args, "--all")
		}
	case "copy", "move":
		var from grid.Range
		read := false
		if g.chance(0.3) {
			// The cells a formula reads, as a whole or in part.
			var rs *grid.Sheet
			if rs, from, read = g.readArea(); read {
				s = rs
			}
		}
		if !read {
			from = g.area(s)
		}
		t := s
		if g.chance(0.2) {
			t = g.sheet()
		}
		t = g.target(t)
		r, c := g.spot(t)
		if g.chance(0.5) {
			tu := used(t)
			r, c = 1+g.rng.IntN(tu.R2+3), 1+g.rng.IntN(tu.C2+3)
		}
		e.Args = []string{class, FileArg, ref(s, from), cell(t, r, c)}
		if class == "copy" && g.chance(0.2) {
			e.Args = append(e.Args, "--values")
		}
	case "sort":
		e.Args = append([]string{"sort", FileArg}, g.sort(s)...)
	case "replace":
		if g.chance(0.2) {
			p := pick(g, formulaReplaces)
			e.Args = []string{"replace", FileArg, p[0], p[1], "--formulas"}
			break
		}
		find := pick(g, finds)
		var texts []string
		s.Each(u, func(_, _ int, c *grid.Cell) {
			if c.Kind == grid.Text && c.Formula == "" && c.Text != "" {
				texts = append(texts, c.Text)
			}
		})
		if len(texts) > 0 && g.chance(0.5) {
			find = pick(g, texts)
		}
		e.Args = []string{"replace", FileArg, find, pick(g, replaces)}
		if g.chance(0.3) {
			e.Args = append(e.Args, s.Name)
		}
		if g.chance(0.3) {
			e.Args = append(e.Args, "--exact")
		}
	case "rows-insert":
		e.Args = []string{"rows", "insert", FileArg, fmt.Sprintf("%s!%d", s.Name, 1+g.rng.IntN(u.R2+1)), "-n", strconv.Itoa(1 + g.rng.IntN(3))}
	case "rows-delete":
		r := 1 + g.rng.IntN(u.R2+1)
		e.Args = []string{"rows", "delete", FileArg, fmt.Sprintf("%s!%d:%d", s.Name, r, r+g.rng.IntN(3))}
	case "cols-insert":
		e.Args = []string{"cols", "insert", FileArg, s.Name + "!" + grid.ColName(1+g.rng.IntN(u.C2+1)), "-n", strconv.Itoa(1 + g.rng.IntN(2))}
	case "cols-delete":
		c := 1 + g.rng.IntN(u.C2+1)
		e.Args = []string{"cols", "delete", FileArg, s.Name + "!" + grid.ColName(c) + ":" + grid.ColName(c+g.rng.IntN(2))}
	case "sheet-rename":
		name := g.newName()
		e.Args = []string{"sheet", "rename", FileArg, s.Name, name}
		g.renamed(s, name)
	case "sheet-copy":
		name := g.newName()
		e.Args = []string{"sheet", "copy", FileArg, s.Name, name}
		g.addSheet(name, s)
	case "sheet-add":
		name := g.newName()
		e.Args = []string{"sheet", "add", FileArg, name}
		g.addSheet(name, nil)
	case "do":
		// The lines are made from the workbook as it is before the batch and
		// the sheets its own lines make: a batch that starts by making a
		// sheet (sheet add or copy, or a set or copy to a new name) goes on
		// to edit it, in one session.
		var b strings.Builder
		cp := *g.wb
		cp.Sheets = slices.Clone(g.wb.Sheets)
		bg := &gen{rng: g.rng, wb: &cp, stream: g.stream, batch: true}
		for i := range 2 + g.rng.IntN(3) {
			var sub Edit
			if i == 0 && bg.chance(0.5) {
				bg.fresh = true
				sub = bg.edit(pick(bg, []string{"sheet-add", "sheet-copy", "set", "copy"}))
				bg.fresh = false
			} else {
				sub = bg.edit(bg.class(false))
			}
			if bg.made != nil {
				bg.focus, bg.made = bg.made, nil
			}
			var line []string
			for _, a := range sub.Args {
				if a != FileArg {
					line = append(line, words.Quote(a))
				}
			}
			b.WriteString(strings.Join(line, " ") + "\n")
		}
		e.Args, e.Stdin = []string{"do", FileArg}, b.String()
	case "script":
		var b strings.Builder
		fmt.Fprintf(&b, "const ws = workbook.getWorksheet(%s);\n", strconv.Quote(s.Name))
		for range 1 + g.rng.IntN(3) {
			b.WriteString(g.scriptOp(s) + "\n")
		}
		e.Args = []string{"script", FileArg, "-e", b.String()}
	}
	return e
}

// sort picks what follows `sheets sort book.xlsx`. Most of the sorts the
// first version made, at random, were refused by both editors: 97 of 137 (the
// sheet holds several tables, the range takes in a dynamic array or a legacy
// array, a formula in it reads another sheet by its row, or other rows by
// position), so the sort class was barely judged. So most are now sorts
// sheets can do, from sortables; the rest are at random, and exercise the
// refusals.
func (g *gen) sort(s *grid.Sheet) []string {
	if g.chance(0.85) {
		if args := g.sortable(s); args != nil {
			return args
		}
	}
	u := used(s)
	var headers []string
	for c := u.C1; c <= u.C2; c++ {
		if h := s.Get(1, c); h.Kind == grid.Text && h.Formula == "" && h.Text != "" {
			headers = append(headers, h.Text)
		}
	}
	if len(headers) > 0 && g.chance(0.5) {
		args := []string{s.Name}
		for range 1 + g.rng.IntN(2) {
			args = append(args, g.desc(pick(g, headers)))
		}
		return args
	}
	r1 := 1 + g.rng.IntN(3)
	rg := grid.Range{R1: r1, C1: u.C1, R2: max(r1+1, u.R2-g.rng.IntN(3)), C2: u.C2}
	if g.chance(0.3) {
		rg.C2 = u.C1 + g.rng.IntN(max(1, u.Cols()))
	}
	return []string{ref(s, rg), g.desc(grid.ColName(rg.C1 + g.rng.IntN(rg.Cols())))}
}

// sortable picks a sort of the sheet that sheets does, as far as the model
// can tell (sortOK), or nil when there is none: the rows of an Excel table
// (by its name, with its column names, or as a range of its data rows and
// columns), or the rows under the header of the block of data at the sheet's
// top left, over some of its leftmost columns.
func (g *gen) sortable(s *grid.Sheet) []string {
	type spec struct {
		rg    grid.Range
		table *grid.NamedTable
	}
	var specs []spec
	for i, t := range s.Tables {
		data := grid.Range{R1: t.Range.R1 + 1, C1: t.Range.C1, R2: t.Range.R2 - t.Totals, C2: t.Range.C2}
		if sortOK(g.wb, s, data) {
			specs = append(specs, spec{data, &s.Tables[i]})
		}
	}
	// The block: the rows from row 2 down to the first with nothing in the
	// sheet's leftmost columns, and each width that holds.
	last := 1
	for r := 2; r <= used(s).R2; r++ {
		empty := true
		for c := 1; c <= 6; c++ {
			if cl := s.Get(r, c); !cl.IsEmpty() {
				empty = false
			}
		}
		if empty {
			break
		}
		last = r
	}
	for c2 := 1; c2 <= min(used(s).C2, 12); c2++ {
		if rg := (grid.Range{R1: 2, C1: 1, R2: last, C2: c2}); sortOK(g.wb, s, rg) {
			specs = append(specs, spec{rg, nil})
		}
	}
	if len(specs) == 0 {
		return nil
	}
	sp := pick(g, specs)
	// Keys: columns without a formula first, since a column of formulas that
	// the calculator cannot compute (a text, a result sheets does not know)
	// refuses the sort.
	var keys, formulaKeys []int
	for c := sp.rg.C1; c <= sp.rg.C2; c++ {
		plain := true
		for r := sp.rg.R1; r <= sp.rg.R2; r++ {
			if s.Get(r, c).Formula != "" {
				plain = false
			}
		}
		if plain {
			keys = append(keys, c)
		} else {
			formulaKeys = append(formulaKeys, c)
		}
	}
	if len(keys) == 0 || g.chance(0.1) {
		keys = append(keys, formulaKeys...)
	}
	n := 1 + g.rng.IntN(2)
	var args []string
	if sp.table != nil && g.chance(0.5) && len(sp.table.Columns) >= sp.rg.Cols() {
		args = []string{sp.table.Name}
		for range n {
			c := pick(g, keys)
			if name := sp.table.Columns[c-sp.rg.C1]; name != "" {
				args = append(args, g.desc(name))
			}
		}
		if len(args) > 1 {
			return args
		}
	}
	args = []string{ref(s, sp.rg)}
	for range n {
		args = append(args, g.desc(grid.ColName(pick(g, keys))))
	}
	return args
}

// sortOK is the model of what sheets refuses to sort (xlsx/sort.go): a
// range with fewer than two rows; one that takes in a dynamic array, its
// results, or part of a legacy array over several cells; one that takes in
// part of an Excel table (it must be the table's data rows, and its
// columns); and one holding a formula that reads another sheet, a range, a
// name, another row, or a column outside the range. (Sheets does some of
// those: a column of formulas filled alike may read other rows. The model
// leaves them out: it is cautious, so what it passes is sorted.)
func sortOK(wb *grid.Workbook, s *grid.Sheet, rg grid.Range) bool {
	if rg.Rows() < 2 {
		return false
	}
	for _, a := range s.Spills {
		if !a.Intersect(rg).Empty() {
			return false
		}
	}
	for _, a := range s.Arrays {
		if a.Rows()*a.Cols() > 1 && !a.Intersect(rg).Empty() {
			return false
		}
	}
	for _, t := range s.Tables {
		if t.Range.Intersect(rg).Empty() {
			continue
		}
		data := grid.Range{R1: t.Range.R1 + 1, C1: t.Range.C1, R2: t.Range.R2 - t.Totals, C2: t.Range.C2}
		if rg.R1 < data.R1 || rg.R2 > data.R2 || rg.C1 != data.C1 || rg.C2 != data.C2 {
			return false
		}
	}
	for _, m := range s.Merges {
		if !m.Intersect(rg).Empty() {
			return false
		}
	}
	for r := rg.R1; r <= rg.R2; r++ {
		if s.RowHidden(r) {
			return false
		}
	}
	var names []string
	for _, n := range wb.Names {
		names = append(names, regexp.QuoteMeta(n.Name))
	}
	var nameRE *regexp.Regexp
	if len(names) > 0 {
		nameRE = regexp.MustCompile(`(?i)\b(` + strings.Join(names, "|") + `)\b`)
	}
	ok := true
	s.Each(rg, func(r, _ int, c *grid.Cell) {
		if c.Formula == "" || !ok {
			return
		}
		if nameRE != nil && nameRE.MatchString(c.Formula) {
			ok = false
		}
		for _, m := range refRE.FindAllStringSubmatch(c.Formula, -1) {
			if m[1]+m[2] != "" || m[5] != "" {
				ok = false // another sheet, or a range
				continue
			}
			col, _ := grid.ColIndex(m[3])
			row, _ := strconv.Atoi(m[4])
			if row != r || col < rg.C1 || col > rg.C2 {
				ok = false
			}
		}
	})
	return ok
}

func (g *gen) desc(key string) string {
	if g.chance(0.4) {
		return key + ":desc"
	}
	return key
}

// scriptOp is one range operation in Office Scripts.
func (g *gen) scriptOp(s *grid.Sheet) string {
	u := used(s)
	q := strconv.Quote
	switch g.rng.IntN(9) {
	case 0:
		r, c := g.spot(s)
		v := g.text(true)
		if n, err := strconv.ParseFloat(v, 64); err == nil {
			return fmt.Sprintf("ws.getRange(%s).setValue(%v);", q(grid.CellName(r, c)), n)
		}
		return fmt.Sprintf("ws.getRange(%s).setValue(%s);", q(grid.CellName(r, c)), q(v))
	case 1, 2:
		r, c := g.spot(s)
		f := g.value(r, false)
		if !strings.HasPrefix(f, "=") {
			f = pick(g, plainFormulas)
			f = strings.ReplaceAll(strings.ReplaceAll(f, "{r}", strconv.Itoa(r)), "{s}", formulaSheet(s.Name))
		}
		return fmt.Sprintf("ws.getRange(%s).setFormula(%s);", q(grid.CellName(r, c)), q(f))
	case 3:
		return fmt.Sprintf("ws.getRange(%s).clear();", q(g.area(s).String()))
	case 4:
		r, c := g.spot(s)
		return fmt.Sprintf("ws.getRange(%s).copyFrom(%s);", q(grid.CellName(r+g.rng.IntN(3), c+g.rng.IntN(3))), q(g.area(s).String()))
	case 5:
		return fmt.Sprintf("ws.getRange(%s).moveTo(%s);", q(g.area(s).String()), q(grid.CellName(1+g.rng.IntN(u.R2+3), 1+g.rng.IntN(u.C2+3))))
	case 6:
		r := 1 + g.rng.IntN(u.R2+1)
		if g.chance(0.5) {
			return fmt.Sprintf("ws.getRange(%s).insert(%s);", q(fmt.Sprintf("%d:%d", r, r+g.rng.IntN(2))), q("Down"))
		}
		return fmt.Sprintf("ws.getRange(%s).delete(%s);", q(fmt.Sprintf("%d:%d", r, r+g.rng.IntN(2))), q("Up"))
	case 7:
		c := grid.ColName(1 + g.rng.IntN(u.C2+1))
		if g.chance(0.5) {
			return fmt.Sprintf("ws.getRange(%s).insert(%s);", q(c+":"+c), q("Right"))
		}
		return fmt.Sprintf("ws.getRange(%s).delete(%s);", q(c+":"+c), q("Left"))
	default:
		rg := grid.Range{R1: 2, C1: u.C1, R2: max(3, u.R2-g.rng.IntN(3)), C2: u.C2}
		return fmt.Sprintf("ws.getRange(%s).getSort().apply([{key: %d, ascending: %v}]);", q(rg.String()), g.rng.IntN(rg.Cols()), g.chance(0.5))
	}
}

// EscapeKinds name what a text can hold that a writer must escape or quote,
// in the order escapeKinds gives them.
var EscapeKinds = []string{"xml", "quote", "space", "_x", "control", "non-ASCII"}

// escapeKinds tells what in a text a writer must escape or quote: "xml" (&, <
// or >), "quote" (a double quote or an apostrophe), "space" (spaces at its
// ends, which need xml:space="preserve"), "_x" (what Excel reads as an
// escape, _x0041_), "control" (a character XML cannot hold, which Excel
// stores as _x0001_) and "non-ASCII".
func escapeKinds(s string) []string {
	var kinds []string
	for _, k := range []struct {
		name string
		has  bool
	}{
		{"xml", strings.ContainsAny(s, "&<>")},
		{"quote", strings.ContainsAny(s, `"'`)},
		{"space", strings.TrimSpace(s) != s},
		{"_x", strings.Contains(s, "_x0")},
		{"control", strings.ContainsFunc(s, func(r rune) bool { return r < 0x20 && r != '\t' && r != '\n' && r != '\r' })},
		{"non-ASCII", strings.ContainsFunc(s, func(r rune) bool { return r > 0x7e })},
	} {
		if k.has {
			kinds = append(kinds, k.name)
		}
	}
	return kinds
}

// typed are the texts a set or write edit types into cells: the values that
// are no formulas.
func typed(e Edit) []string {
	var texts []string
	switch e.Class {
	case "set":
		for i := 3; i < len(e.Args); i += 2 {
			texts = append(texts, e.Args[i])
		}
	case "write":
		for _, row := range strings.Split(e.Stdin, "\n") {
			texts = append(texts, strings.Split(row, "\t")...)
		}
	}
	return slices.DeleteFunc(texts, func(t string) bool { return strings.HasPrefix(t, "=") })
}

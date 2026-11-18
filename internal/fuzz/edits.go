package fuzz

import (
	"fmt"
	"math/rand/v2"
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
	{"sheet-rename", 3}, {"sheet-copy", 3}, {"do", 5}, {"script", 6},
}

// The formulas edits add, by the kind Excel 365 stores each typed in: {r}
// is the row of its cell, {s} a sheet of the workbook.
var (
	plainFormulas = []string{
		`=B{r}*2+N("p")`, `=SUM(B2:B6)+N("p")`, `=IF(B{r}>3,1,0)+N("p")`, `=MAX(C2:C5)+N("p")`,
		`=MIN(B2:D2)+N("p")`, `={s}!B2+1+N("p")`, `=COUNTIF(B2:B9,">3")+N("p")`, `=ROUND(C{r}*1.1,2)+N("p")`,
	}
	dynamicFormulas = []string{
		`=SEQUENCE(2)+N("d")`, `=B2:B4*2+N("d")`, `=SUM(IF(B2:B6>3,B2:B6))+N("d")`, `=SEQUENCE(1,2)+N("d")`,
		`=MAX(B2:B6-D2:D6)+N("d")`,
	}
	values   = []string{"5", "-3", "12.5", "0", "Apple", "N/A", "Acme Inc.", "2026-07-01", "TRUE", "hello", "1,200"}
	finds    = []string{"Apple", "N/A", "Fig", "Total", "5", "apple", "Acme Inc.", "Pear"}
	replaces = []string{"Apple", "Pear", "", "1,200", "7", "Acme Corp", "2026-01-31"}
	// Replacing in formula text renames functions only, which keeps kinds.
	formulaReplaces = [][2]string{{"SUM(", "PRODUCT("}, {"PRODUCT(", "SUM("}, {"MIN(", "MAX("}, {"MAX(", "MIN("}}
	newNames        = []string{"Copy", "Data 2", "Q 2027", "Old", "Backup", "Plan"}
)

type gen struct {
	rng *rand.Rand
	wb  *grid.Workbook
}

// generate makes a random edit for a workbook.
func generate(rng *rand.Rand, wb *grid.Workbook) Edit {
	g := &gen{rng: rng, wb: wb}
	return g.edit(g.class(true))
}

func (g *gen) class(all bool) string {
	total := 0
	for _, c := range Classes {
		if all || c.Name != "do" && c.Name != "script" && c.Name != "write" {
			total += c.Weight
		}
	}
	n := g.rng.IntN(total)
	for _, c := range Classes {
		if !all && (c.Name == "do" || c.Name == "script" || c.Name == "write") {
			continue
		}
		if n -= c.Weight; n < 0 {
			return c.Name
		}
	}
	panic("unreachable")
}

func (g *gen) chance(p float64) bool { return g.rng.Float64() < p }

func pick[T any](g *gen, xs []T) T { return xs[g.rng.IntN(len(xs))] }

func (g *gen) sheet() *grid.Sheet { return pick(g, g.wb.Sheets) }

// used is a sheet's used range, at least A1.
func used(s *grid.Sheet) grid.Range {
	u := s.Used()
	if u.R2 < 1 {
		return grid.Range{R1: 1, C1: 1, R2: 1, C2: 1}
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

// value picks what set or write puts in a cell: a value or a formula; row
// fills in {r} unless 0.
func (g *gen) value(row int) string {
	var v string
	switch n := g.rng.IntN(10); {
	case n < 4:
		v = pick(g, values)
	case n < 8:
		v = pick(g, plainFormulas)
	default:
		v = pick(g, dynamicFormulas)
	}
	v = strings.ReplaceAll(v, "{s}", formulaSheet(g.sheet().Name))
	if row > 0 {
		v = strings.ReplaceAll(v, "{r}", strconv.Itoa(row))
	}
	return v
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
		if g.chance(0.3) {
			rg := g.area(s)
			rg.R2 = rg.R1 + g.rng.IntN(4)
			e.Args = append(e.Args, ref(s, rg), g.value(0))
			break
		}
		for range 1 + g.rng.IntN(3) {
			r, c := g.spot(s)
			e.Args = append(e.Args, cell(s, r, c), g.value(r))
		}
	case "write":
		r, c := g.spot(s)
		var b strings.Builder
		for i := range 1 + g.rng.IntN(3) {
			var row []string
			for range 1 + g.rng.IntN(3) {
				row = append(row, g.value(r+i))
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
		from := g.area(s)
		t := s
		if g.chance(0.2) {
			t = g.sheet()
		}
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
		e.Args = []string{"sort", FileArg}
		var headers []string
		for c := u.C1; c <= u.C2; c++ {
			if h := s.Get(1, c); h.Kind == grid.Text && h.Formula == "" && h.Text != "" {
				headers = append(headers, h.Text)
			}
		}
		if len(headers) > 0 && g.chance(0.5) {
			e.Args = append(e.Args, s.Name)
			for range 1 + g.rng.IntN(2) {
				e.Args = append(e.Args, g.desc(pick(g, headers)))
			}
			break
		}
		r1 := 1 + g.rng.IntN(3)
		rg := grid.Range{R1: r1, C1: u.C1, R2: max(r1+1, u.R2-g.rng.IntN(3)), C2: u.C2}
		if g.chance(0.3) {
			rg.C2 = u.C1 + g.rng.IntN(max(1, u.Cols()))
		}
		e.Args = append(e.Args, ref(s, rg), g.desc(grid.ColName(rg.C1+g.rng.IntN(rg.Cols()))))
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
		e.Args = []string{"sheet", "rename", FileArg, s.Name, g.newName()}
	case "sheet-copy":
		e.Args = []string{"sheet", "copy", FileArg, s.Name, g.newName()}
	case "do":
		var b strings.Builder
		for range 2 + g.rng.IntN(3) {
			sub := g.edit(g.class(false))
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
		v := pick(g, values)
		if n, err := strconv.ParseFloat(v, 64); err == nil {
			return fmt.Sprintf("ws.getRange(%s).setValue(%v);", q(grid.CellName(r, c)), n)
		}
		return fmt.Sprintf("ws.getRange(%s).setValue(%s);", q(grid.CellName(r, c)), q(v))
	case 1, 2:
		r, c := g.spot(s)
		f := g.value(r)
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

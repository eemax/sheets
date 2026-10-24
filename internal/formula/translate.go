package formula

import (
	"strconv"
	"strings"

	"github.com/eemax/sheets/internal/grid"
)

// Translations of formulas into forms a calculator without table
// references or LET can evaluate, with the same result. They are for
// calculating only; files keep the formulas as written.

// Table is an Excel table as structured references see it.
type Table struct {
	Name    string
	Sheet   string
	Ref     grid.Range
	Headers int // header rows at the top of Ref
	Totals  int // totals rows at the bottom
	Columns []string
}

// ExpandTables replaces the structured references of a formula in cell
// (row, col) of sheet with the cells they stand for. ok is false when one
// cannot be read, or stands for several cells where a plain range would
// not mean the same (a column used in arithmetic, which Excel 365
// spills).
func ExpandTables(f, sheet string, row, col int, tables []Table) (out string, ok bool) {
	find := func(name string) *Table {
		for i := range tables {
			if strings.EqualFold(tables[i].Name, name) {
				return &tables[i]
			}
		}
		return nil
	}
	var b strings.Builder
	changed := false
	ts := lex(f)
	for i := 0; i < len(ts); i++ {
		t := ts[i]
		var tbl *Table
		var spec string
		switch {
		case t.word && !t.call && i+1 < len(ts) && strings.HasPrefix(ts[i+1].text, "[") && (i == 0 || ts[i-1].text != "!"):
			if tbl = find(bare(t.text)); tbl == nil {
				return f, false // a link to another workbook, or an unknown table
			}
			spec = ts[i+1].text
			i++
		case strings.HasPrefix(t.text, "["):
			if i+2 < len(ts) && ts[i+1].word && ts[i+2].text == "!" {
				b.WriteString(t.text) // [1]Sheet!A1: another workbook
				continue
			}
			for k := range tables {
				if strings.EqualFold(tables[k].Sheet, sheet) && tables[k].Ref.Contains(row, col) {
					tbl = &tables[k]
				}
			}
			if tbl == nil {
				return f, false
			}
			spec = t.text
		case t.word && !t.call && (i == 0 || ts[i-1].text != "!") && (i+1 >= len(ts) || ts[i+1].text != "!") && find(bare(t.text)) != nil:
			tbl, spec = find(bare(t.text)), "[#Data]" // the table's name alone: its data
		default:
			b.WriteString(t.text)
			continue
		}
		rg, err := tbl.cells(spec, row)
		if err != nil {
			return f, false
		}
		if rg.Rows() > 1 || rg.Cols() > 1 {
			// A range: the same only as a whole argument of a function.
			prev, next := lastNonSpace(b.String()), ""
			for k := i + 1; k < len(ts); k++ {
				if strings.TrimSpace(ts[k].text) != "" {
					next = ts[k].text
					break
				}
			}
			if prev != "(" && prev != "," || next != ")" && next != "," {
				return f, false
			}
		}
		b.WriteString(absRef(tbl.Sheet, rg))
		changed = true
	}
	if !changed {
		return f, true
	}
	return b.String(), true
}

func lastNonSpace(s string) string {
	s = strings.TrimRight(s, " \t\r\n")
	if s == "" {
		return ""
	}
	return s[len(s)-1:]
}

func absRef(sheet string, rg grid.Range) string {
	cell := func(r, c int) string { return "$" + grid.ColName(c) + "$" + strconv.Itoa(r) }
	out := grid.QuoteSheet(sheet) + "!" + cell(rg.R1, rg.C1)
	if rg.Rows() > 1 || rg.Cols() > 1 {
		out += ":" + cell(rg.R2, rg.C2)
	}
	return out
}

type specError struct{}

func (specError) Error() string { return "structured reference not understood" }

// cells reads a structured reference specifier ("[Units]", "[@Units]",
// "[[#Totals],[Units]]", "[[Units]:[Price]]") into the cells it means for a
// formula on row.
func (t *Table) cells(spec string, row int) (grid.Range, error) {
	inner := spec[1 : len(spec)-1]
	first, last := t.Ref.R1+t.Headers, t.Ref.R2-t.Totals
	var rows []string // special items
	var cols []string // column names, one or a span of two
	switch {
	case strings.HasPrefix(inner, "@"):
		rows = []string{"#This Row"}
		if rest := inner[1:]; rest != "" {
			if strings.HasPrefix(rest, "[") {
				items, err := specItems(rest)
				if err != nil {
					return grid.Range{}, err
				}
				cols = items
			} else {
				cols = []string{unescapeColumn(rest)}
			}
		}
	case strings.HasPrefix(inner, "["):
		items, err := specItems(inner)
		if err != nil {
			return grid.Range{}, err
		}
		for _, it := range items {
			if strings.HasPrefix(it, "#") {
				rows = append(rows, it)
			} else {
				cols = append(cols, it)
			}
		}
	case strings.HasPrefix(inner, "#"):
		rows = []string{inner}
	case inner == "":
		rows = []string{"#Data"}
	default:
		cols = []string{unescapeColumn(inner)}
	}
	if len(rows) == 0 {
		rows = []string{"#Data"}
	}
	rg := grid.Range{R1: -1}
	add := func(r1, r2 int) {
		if rg.R1 < 0 {
			rg.R1, rg.R2 = r1, r2
		} else {
			rg.R1, rg.R2 = min(rg.R1, r1), max(rg.R2, r2)
		}
	}
	for _, r := range rows {
		switch strings.ToUpper(r) {
		case "#ALL":
			add(t.Ref.R1, t.Ref.R2)
		case "#DATA":
			if first > last {
				return grid.Range{}, specError{}
			}
			add(first, last)
		case "#HEADERS":
			if t.Headers == 0 {
				return grid.Range{}, specError{}
			}
			add(t.Ref.R1, t.Ref.R1+t.Headers-1)
		case "#TOTALS":
			if t.Totals == 0 {
				return grid.Range{}, specError{}
			}
			add(t.Ref.R2-t.Totals+1, t.Ref.R2)
		case "#THIS ROW":
			if row < first || row > last {
				return grid.Range{}, specError{} // #VALUE! in Excel
			}
			add(row, row)
		default:
			return grid.Range{}, specError{}
		}
	}
	rg.C1, rg.C2 = t.Ref.C1, t.Ref.C2
	if len(cols) > 0 {
		c1, ok1 := t.column(cols[0])
		c2, ok2 := t.column(cols[len(cols)-1])
		if !ok1 || !ok2 || len(cols) > 2 {
			return grid.Range{}, specError{}
		}
		rg.C1, rg.C2 = min(c1, c2), max(c1, c2)
	}
	// Specials must make one block: #Headers with #Totals skips the data.
	if len(rows) == 2 && strings.EqualFold(rows[0], "#Headers") && strings.EqualFold(rows[1], "#Totals") {
		return grid.Range{}, specError{}
	}
	return rg, nil
}

func (t *Table) column(name string) (int, bool) {
	for i, c := range t.Columns {
		if strings.EqualFold(c, name) {
			return t.Ref.C1 + i, true
		}
	}
	return 0, false
}

// specItems splits "[#Data],[Units]" or "[Units]:[Price]" into its items,
// unescaped.
func specItems(s string) ([]string, error) {
	var out []string
	for i := 0; i < len(s); {
		switch s[i] {
		case '[':
			end := closeBracket(s, i)
			if end < 0 {
				return nil, specError{}
			}
			out = append(out, unescapeColumn(s[i+1:end-1]))
			i = end
		case ',', ':', ' ':
			i++
		default:
			return nil, specError{}
		}
	}
	return out, nil
}

// InlineLet rewrites LET(name, value, ..., body) as body with each name
// replaced by its value, innermost LET first. ok is false where that
// would change the result: a value calling a volatile function (computed
// once by LET, each time when inlined), or a LAMBDA anywhere.
func InlineLet(f string) (string, bool) {
	for _, t := range lex(f) {
		if n := strings.ToUpper(bare(t.text)); t.call && lambdaFamily[n] && n != "LET" {
			return f, false
		}
	}
	for range 64 {
		ts := lex(f)
		at := -1
		for i, t := range ts {
			if t.call && strings.EqualFold(bare(t.text), "LET") {
				at = i // the last one: nothing after it is an outer LET
			}
		}
		if at < 0 {
			return f, true
		}
		// The innermost: no LET among its arguments, as the last one.
		args := topLevelArgs(ts, at+1)
		if len(args) < 3 || len(args)%2 == 0 {
			return f, false
		}
		names := map[string]string{} // upper-cased name → value text
		for k := 0; k+1 < len(args); k += 2 {
			if len(args[k]) != 1 || !args[k][0].word {
				return f, false
			}
			value := substitute(args[k+1], names)
			if d := DepsOf(value); d.Volatile {
				return f, false
			}
			names[strings.ToUpper(bare(args[k][0].text))] = "(" + value + ")"
		}
		body := "(" + substitute(args[len(args)-1], names) + ")"
		// Replace from the LET name to its closing parenthesis.
		end := closeParen(ts, at+1)
		if end < 0 {
			return f, false
		}
		f = join(ts[:at]) + body + join(ts[end+1:])
	}
	return f, false
}

// substitute writes tokens with each word naming a LET variable replaced.
func substitute(ts []token, names map[string]string) string {
	var b strings.Builder
	for i, t := range ts {
		if t.word && !t.call && (i == 0 || ts[i-1].text != "!") && (i+1 >= len(ts) || ts[i+1].text != "!" && !strings.HasPrefix(ts[i+1].text, "[")) {
			if v, ok := names[strings.ToUpper(bare(t.text))]; ok {
				b.WriteString(v)
				continue
			}
		}
		b.WriteString(t.text)
	}
	return b.String()
}

// closeParen returns the index of the parenthesis closing the one at
// ts[open], or -1.
func closeParen(ts []token, open int) int {
	depth := 0
	for i := open; i < len(ts); i++ {
		switch ts[i].text {
		case "(":
			depth++
		case ")":
			if depth--; depth == 0 {
				return i
			}
		}
	}
	return -1
}

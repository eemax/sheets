package analyze

import (
	"fmt"
	"strings"

	"github.com/eemax/sheets/internal/grid"
)

// Resolve reads a reference to cells: a range (Sales!B2:D9, B:D, 2:2), a
// sheet name alone (r is then nil), or a table reference in Excel's
// syntax, where the name is an Excel table or a sheet holding one table:
//
//	Sales[Revenue]               the column's data rows
//	Sales[[Units]:[Revenue]]     several columns
//	Sales[#Headers]              the header row (#Data, #Totals, #All)
//	Sales[[#Totals],[Units]]     one cell of the total row
func Resolve(wb *grid.Workbook, ref string) (sheet string, r *grid.Range, err error) {
	ref = strings.TrimSpace(ref)
	if name, inner, ok := structured(ref); ok && wb.Sheet(ref) == nil {
		t, err := findTable(wb, name, inner)
		if err != nil {
			return "", nil, err
		}
		rg, err := t.cells(inner)
		if err != nil {
			return "", nil, err
		}
		return t.sheet, &rg, nil
	}
	p, err := grid.ParseRef(ref, func(n string) bool { return wb.Sheet(n) != nil })
	if err != nil {
		return "", nil, fmt.Errorf("%q: %v (give a range like Sales!B2:D9, a sheet, or a column like 'Sales[Revenue]')", ref, err)
	}
	if p.Sheet == "" {
		s := wb.FirstVisible()
		if s == nil {
			return "", nil, fmt.Errorf("the workbook has no sheets")
		}
		return s.Name, p.Range, nil
	}
	s := wb.Sheet(p.Sheet)
	if s == nil {
		return "", nil, fmt.Errorf("sheet %q not found (sheets: %s)", p.Sheet, strings.Join(wb.SheetNames(), ", "))
	}
	return s.Name, p.Range, nil
}

// structured splits Name[...] into the name and what the brackets hold.
func structured(ref string) (name, inner string, ok bool) {
	i := strings.IndexByte(ref, '[')
	if i <= 0 || !strings.HasSuffix(ref, "]") {
		return "", "", false
	}
	name = ref[:i]
	if len(name) >= 2 && name[0] == '\'' && name[len(name)-1] == '\'' {
		name = strings.ReplaceAll(name[1:len(name)-1], "''", "'")
	}
	return name, ref[i+1 : len(ref)-1], true
}

// tableArea is a table's rows and columns.
type tableArea struct {
	name, sheet string
	header      int // 0: none
	first, last int // data rows
	totals      int // 0: none
	c1, c2      int
	cols        []string // column names, from c1
}

func findTable(wb *grid.Workbook, name, inner string) (tableArea, error) {
	for _, s := range wb.Sheets {
		for _, t := range s.Tables {
			if !strings.EqualFold(t.Name, name) {
				continue
			}
			a := tableArea{name: t.Name, sheet: s.Name, header: t.Range.R1, first: t.Range.R1 + 1, last: t.Range.R2 - t.Totals, c1: t.Range.C1, c2: t.Range.C2, cols: t.Columns}
			if t.Totals > 0 {
				a.totals = t.Range.R2
			}
			return a, nil
		}
	}
	s := wb.Sheet(name)
	if s == nil {
		var tables []string
		for _, sh := range wb.Sheets {
			for _, t := range sh.Tables {
				tables = append(tables, t.Name)
			}
		}
		hint := "sheets: " + strings.Join(wb.SheetNames(), ", ")
		if len(tables) > 0 {
			hint += "; Excel tables: " + strings.Join(tables, ", ")
		}
		return tableArea{}, fmt.Errorf("%q is neither a sheet nor an Excel table (%s)", name, hint)
	}
	var found []tableArea
	for _, reg := range Regions(wb, s) {
		if !reg.IsTable() {
			continue
		}
		a := tableArea{name: s.Name, sheet: s.Name, header: reg.HeaderRow, first: reg.Body.R1, last: reg.Body.R2, totals: reg.TotalRow, c1: reg.Range.C1, c2: reg.Range.C2}
		for c := a.c1; c <= a.c2; c++ {
			a.cols = append(a.cols, Label(wb, s, reg.HeaderRow, c))
		}
		found = append(found, a)
	}
	switch {
	case len(found) == 1:
		return found[0], nil
	case len(found) == 0:
		return tableArea{}, fmt.Errorf("sheet %q has no table with a header row; give a range instead", s.Name)
	}
	// Several tables: the one with the columns named.
	var match []tableArea
	for _, a := range found {
		if _, err := a.cells(inner); err == nil {
			match = append(match, a)
		}
	}
	if len(match) == 1 {
		return match[0], nil
	}
	var where []string
	for _, a := range found {
		where = append(where, grid.Range{R1: a.header, C1: a.c1, R2: max(a.last, a.totals), C2: a.c2}.String())
	}
	return tableArea{}, fmt.Errorf("sheet %q holds %d tables (%s); give a range instead", s.Name, len(found), strings.Join(where, ", "))
}

// cells resolves what a table reference's brackets hold.
func (t tableArea) cells(inner string) (grid.Range, error) {
	items, err := splitItems(inner)
	if err != nil {
		return grid.Range{}, err
	}
	rows := map[string]bool{}
	c1, c2 := t.c1, t.c2
	cols := false
	for _, it := range items {
		if strings.HasPrefix(it.from, "#") && it.to == "" {
			switch k := strings.ToLower(it.from); k {
			case "#headers", "#data", "#totals", "#all":
				rows[k] = true
				continue
			case "#this row":
				return grid.Range{}, fmt.Errorf("%s[%s]: @ and #This Row only mean something inside a formula", t.name, inner)
			default:
				return grid.Range{}, fmt.Errorf("%s[%s]: %s is not one of #Headers, #Data, #Totals, #All", t.name, inner, it.from)
			}
		}
		if cols {
			return grid.Range{}, fmt.Errorf("%s[%s]: give one column or one span [A]:[B]", t.name, inner)
		}
		a, err := t.col(it.from)
		if err != nil {
			return grid.Range{}, err
		}
		b := a
		if it.to != "" {
			if b, err = t.col(it.to); err != nil {
				return grid.Range{}, err
			}
		}
		c1, c2, cols = min(a, b), max(a, b), true
	}
	if len(rows) == 0 {
		rows["#data"] = true
	}
	if rows["#all"] {
		rows["#headers"], rows["#data"], rows["#totals"] = true, true, t.totals > 0
	}
	if rows["#totals"] && t.totals == 0 {
		return grid.Range{}, fmt.Errorf("%s has no total row", t.name)
	}
	if rows["#headers"] && t.header == 0 {
		return grid.Range{}, fmt.Errorf("%s has no header row", t.name)
	}
	r1, r2 := 0, 0
	add := func(a, b int) {
		if r1 == 0 || a < r1 {
			r1 = a
		}
		r2 = max(r2, b)
	}
	if rows["#headers"] {
		add(t.header, t.header)
	}
	if rows["#data"] {
		add(t.first, t.last)
	}
	if rows["#totals"] {
		add(t.totals, t.totals)
	}
	if rows["#headers"] && rows["#totals"] && !rows["#data"] {
		return grid.Range{}, fmt.Errorf("%s[%s]: #Headers and #Totals are not next to each other; style them one at a time", t.name, inner)
	}
	return grid.Range{R1: r1, C1: c1, R2: r2, C2: c2}, nil
}

func (t tableArea) col(name string) (int, error) {
	want := strings.TrimSpace(name)
	for i, c := range t.cols {
		if strings.EqualFold(strings.TrimSpace(c), want) {
			return t.c1 + i, nil
		}
	}
	var names []string
	for _, c := range t.cols {
		if c != "" {
			names = append(names, c)
		}
	}
	return 0, fmt.Errorf("%s has no column %q (columns: %s)", t.name, name, strings.Join(names, ", "))
}

type item struct{ from, to string }

// splitItems reads "Revenue", "#Headers", "[#Totals],[Units]" or
// "[Units]:[Revenue]". A ' escapes the next character in a name.
func splitItems(s string) ([]item, error) {
	s = strings.TrimSpace(s)
	if s == "" {
		return nil, fmt.Errorf("empty brackets: give a column, like Sales[Revenue]")
	}
	if !strings.HasPrefix(s, "[") {
		return []item{{from: unescapeItem(s)}}, nil
	}
	var out []item
	for s != "" {
		name, rest, err := bracketed(s)
		if err != nil {
			return nil, err
		}
		it := item{from: name}
		if strings.HasPrefix(rest, ":") {
			if it.to, rest, err = bracketed(strings.TrimSpace(rest[1:])); err != nil {
				return nil, err
			}
		}
		out = append(out, it)
		rest = strings.TrimSpace(rest)
		if rest != "" && rest[0] != ',' {
			return nil, fmt.Errorf("[%s]: expected , between items", s)
		}
		s = strings.TrimSpace(strings.TrimPrefix(rest, ","))
	}
	return out, nil
}

// bracketed reads [name] from the start of s.
func bracketed(s string) (string, string, error) {
	if !strings.HasPrefix(s, "[") {
		return "", "", fmt.Errorf("expected [ at %q", s)
	}
	var b strings.Builder
	for i := 1; i < len(s); i++ {
		switch s[i] {
		case '\'':
			if i+1 < len(s) {
				i++
				b.WriteByte(s[i])
			}
		case ']':
			return b.String(), s[i+1:], nil
		default:
			b.WriteByte(s[i])
		}
	}
	return "", "", fmt.Errorf("[ not closed in %q", s)
}

func unescapeItem(s string) string {
	var b strings.Builder
	for i := 0; i < len(s); i++ {
		if s[i] == '\'' && i+1 < len(s) {
			i++
		}
		b.WriteByte(s[i])
	}
	return b.String()
}

// TotalsUnder returns the total row's cells under the columns a table
// reference names, when it names columns' data rows (Sales[Revenue]) of a
// table with a total row.
func TotalsUnder(wb *grid.Workbook, ref string) (sheet string, r grid.Range, ok bool) {
	name, inner, isTable := structured(strings.TrimSpace(ref))
	if !isTable || wb.Sheet(ref) != nil {
		return "", grid.Range{}, false
	}
	items, err := splitItems(inner)
	if err != nil {
		return "", grid.Range{}, false
	}
	for _, it := range items {
		if strings.HasPrefix(it.from, "#") {
			return "", grid.Range{}, false
		}
	}
	t, err := findTable(wb, name, inner)
	if err != nil || t.totals == 0 {
		return "", grid.Range{}, false
	}
	cols, err := t.cells(inner)
	if err != nil {
		return "", grid.Range{}, false
	}
	return t.sheet, grid.Range{R1: t.totals, C1: cols.C1, R2: t.totals, C2: cols.C2}, true
}

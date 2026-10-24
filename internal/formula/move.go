package formula

import (
	"strings"

	"github.com/eemax/sheets/internal/grid"
)

// Move is a block of cells cut from one place and pasted at another.
type Move struct {
	From   string     // the sheet the cells come from
	Src    grid.Range // the cells cut
	To     string     // the sheet they land on (From when the same)
	DR, DC int        // how far they go
}

// Dst is where the cells land.
func (m Move) Dst() grid.Range {
	return grid.Range{R1: m.Src.R1 + m.DR, C1: m.Src.C1 + m.DC, R2: m.Src.R2 + m.DR, C2: m.Src.C2 + m.DC}
}

// MoveRefs rewrites a formula for m as Excel does when cells are cut and
// pasted. A reference to a cell, or a range, wholly inside the moved block
// follows it; one wholly inside the cells it lands on, which did not move
// themselves, becomes #REF!; the rest stay. home is the sheet holding the
// formula ("" for names and other always-qualified text), newHome the sheet
// holding it afterwards: a formula moved with the block to another sheet
// keeps pointing at the cells it read, so its references to its old sheet
// get the sheet's name. A formula whose references cannot be told apart
// (across sheets, Jan:Dec!B2) is returned as it is.
func MoveRefs(f, home, newHome string, m Move) string {
	if strings.Contains(f, ":") && DepsOf(f).Multi {
		return f
	}
	ps := tokenize(f)
	dst := m.Dst()
	sameSheet := strings.EqualFold(m.From, m.To)
	changed := false
	// qual is the sheet a reference is written with, "" for none.
	type out struct {
		qual    string
		requal  bool // the qualifier is to be written anew
		ref     bool // the parts become #REF!
		skipped bool
	}
	plan := make([]out, len(ps))
	for i := 0; i < len(ps); i++ {
		p := ps[i]
		if !p.isRef {
			continue
		}
		j := i // last part of the reference
		if i+2 < len(ps) && !ps[i+1].isRef && ps[i+1].text == ":" && ps[i+2].isRef {
			j = i + 2
		}
		if p.qualified && external(ps, i) {
			i = j
			continue
		}
		sheet := home
		if p.qualified {
			sheet = p.sheet
		}
		rg := bounds(p)
		if j > i {
			rg = rg.Union(bounds(ps[j]))
		}
		target := sheet
		moved := strings.EqualFold(sheet, m.From) && within(rg, m.Src)
		switch {
		case moved:
			target = m.To
			for k := i; k <= j; k += 2 {
				q := &ps[k]
				if q.kind != rowRef {
					q.col += m.DC
				}
				if q.kind != colRef {
					q.row += m.DR
				}
			}
			changed = true
		case strings.EqualFold(sheet, m.To) && within(rg, dst) && !(sameSheet && within(rg, m.Src)):
			plan[i].ref = true
			for k := i + 1; k <= j; k++ {
				plan[k].skipped = true
			}
			changed = true
			i = j
			continue
		}
		// Qualify a reference that now points away from the sheet holding
		// the formula.
		switch {
		case p.qualified && moved && !strings.EqualFold(m.From, m.To):
			plan[i].qual, plan[i].requal = m.To, true
		case !p.qualified && home != "" && !strings.EqualFold(target, newHome):
			plan[i].qual, plan[i].requal = target, true
			changed = true
		}
		i = j
	}
	if !changed {
		return f
	}
	var b strings.Builder
	for i, p := range ps {
		switch {
		case plan[i].skipped:
			continue
		case !p.isRef:
			b.WriteString(p.text)
			continue
		case plan[i].ref:
			b.WriteString("#REF!")
			continue
		}
		if plan[i].requal {
			if p.qualified {
				s := b.String()
				b.Reset()
				b.WriteString(strings.TrimSuffix(s, qualifierOf(s)))
			}
			b.WriteString(grid.QuoteSheet(plan[i].qual) + "!")
		}
		if p.kind != rowRef && (p.col < 1 || p.col > grid.MaxCols) || p.kind != colRef && (p.row < 1 || p.row > grid.MaxRows) {
			b.WriteString("#REF!")
			continue
		}
		b.WriteString(p.render())
	}
	return b.String()
}

func within(r, outer grid.Range) bool {
	return r.R1 >= outer.R1 && r.R2 <= outer.R2 && r.C1 >= outer.C1 && r.C2 <= outer.C2
}

// external tells whether the qualified reference at ps[i] names a sheet of
// another workbook: [1]Sheet!A1 or '[Book.xlsx]Sheet'!A1.
func external(ps []part, i int) bool {
	if strings.HasPrefix(ps[i].sheet, "[") {
		return true
	}
	if i == 0 || ps[i-1].isRef {
		return false
	}
	lit := ps[i-1].text
	q := qualifierOf(lit)
	before := strings.TrimSuffix(lit, q)
	return strings.HasSuffix(before, "]")
}

// qualifierOf returns the sheet prefix (Sheet! or 'My Sheet'!) that s
// ends with, or "".
func qualifierOf(s string) string {
	if !strings.HasSuffix(s, "!") {
		return ""
	}
	body := s[:len(s)-1]
	if strings.HasSuffix(body, "'") {
		// Back to the opening quote; '' inside is an escaped quote.
		i := len(body) - 2
		for i >= 0 {
			if body[i] == '\'' {
				if i > 0 && body[i-1] == '\'' {
					i -= 2
					continue
				}
				return s[i:]
			}
			i--
		}
		return ""
	}
	i := len(body)
	for i > 0 && (isWordChar(body[i-1]) || body[i-1] == '$') {
		i--
	}
	return s[i:]
}

// RefUse is a reference as a formula writes it, for telling what moving
// the formula does to it.
type RefUse struct {
	Sheet string     // "" for the formula's own sheet
	Range grid.Range // whole rows and columns reach the sheet's edges
	// RelRows and AbsRows tell whether all of its rows are relative (they
	// shift when the formula moves) or all absolute ($2); neither for a
	// mix like $C$2:C2. Whole columns are neither.
	RelRows, AbsRows bool
	Text             string // as written, with its sheet
}

// RefUses lists the references a formula writes, outside strings, table
// references and names. External references are left out.
func RefUses(f string) []RefUse {
	ps := tokenize(f)
	var out []RefUse
	for i := 0; i < len(ps); i++ {
		p := ps[i]
		if !p.isRef {
			continue
		}
		j := i
		if i+2 < len(ps) && !ps[i+1].isRef && ps[i+1].text == ":" && ps[i+2].isRef {
			j = i + 2
		}
		if p.qualified && external(ps, i) {
			i = j
			continue
		}
		u := RefUse{Range: bounds(p)}
		text := p.render()
		if p.qualified {
			u.Sheet = p.sheet
			text = grid.QuoteSheet(p.sheet) + "!" + text
		}
		rel, abs := true, true
		for k := i; k <= j; k += 2 {
			q := ps[k]
			u.Range = u.Range.Union(bounds(q))
			if q.kind == colRef {
				rel, abs = false, false
				continue
			}
			rel = rel && !q.rowAbs
			abs = abs && q.rowAbs
		}
		if j > i {
			text += ":" + ps[j].render()
		}
		u.RelRows, u.AbsRows, u.Text = rel, abs, text
		out = append(out, u)
		i = j
	}
	return out
}

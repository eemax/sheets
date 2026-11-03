package formula

import "strings"

// Deps is what a formula's result depends on, read off its text.
type Deps struct {
	// Refs are the cells and ranges it reads; Sheet "" is its own sheet.
	Refs []Ref
	// Names are words that may be workbook names (Rate, or a LAMBDA
	// called like a function): not functions, references, LET or LAMBDA
	// parameters, or booleans.
	Names []string
	// Tables are the Excel tables named in structured references
	// (Sales[Units]); "" stands for a bare [Units], which inside a table
	// means that table.
	Tables []string
	// Volatile marks a result that can change while its references do
	// not, or whose references cannot be read off the formula: NOW,
	// RAND, INDIRECT, OFFSET, references spanning several sheets.
	Volatile bool
	// External marks links to other workbooks, whose stored results only
	// Excel can bring up to date.
	External bool
	// Multi marks references across several sheets (Jan:Dec!B2).
	Multi bool
	// Position marks a result that depends on where cells are (ROW,
	// COLUMN, ...), which inserting or deleting rows or columns changes.
	Position bool
}

var (
	volatileFns = setOf(`NOW TODAY RAND RANDBETWEEN RANDARRAY INDIRECT OFFSET CELL INFO GETPIVOTDATA`)
	positionFns = setOf(`ROW ROWS COLUMN COLUMNS ADDRESS`)
)

// DepsOf reads a formula's dependencies.
func DepsOf(f string) Deps {
	d := Deps{Refs: Refs(f)}
	ts := lex(f)
	params := lambdaParams(ts)
	text := func(i int) string {
		if i < 0 || i >= len(ts) {
			return ""
		}
		return ts[i].text
	}
	seen := map[string]bool{}
	skip := -1 // tokens up to here belong to an error literal
	for i, t := range ts {
		if i <= skip {
			continue
		}
		if t.text == "#" {
			skip = errorLiteral(ts, i)
			continue
		}
		switch {
		case strings.HasPrefix(t.text, "'") && strings.Contains(t.text, "["):
			d.External = true // '[Book.xlsx]Sheet'!A1: another workbook
		case strings.HasPrefix(t.text, "["):
			switch {
			case i+2 < len(ts) && ts[i+1].word && ts[i+2].text == "!":
				d.External = true // [1]Sheet!A1: another workbook
			case i > 0 && ts[i-1].word:
				d.Tables = append(d.Tables, bare(ts[i-1].text))
			default:
				d.Tables = append(d.Tables, "")
			}
		case t.call:
			fn := strings.ToUpper(bare(t.text))
			switch {
			case volatileFns[fn]:
				d.Volatile = true
			case positionFns[fn]:
				d.Position = true
			case !IsFunction(fn) && !params[fn]:
				d.addName(t.text, seen) // a LAMBDA kept in a name
			}
		case t.word:
			w := t.text
			switch {
			case text(i+1) == "!" || strings.HasPrefix(text(i+1), "["):
				// A sheet name, or a table name (handled with its bracket).
			case w[0] >= '0' && w[0] <= '9' || w[0] == '.':
				// A number, or a row of a row range.
			case text(i+1) == ":" && i+3 < len(ts) && ts[i+2].word && text(i+3) == "!":
				d.Volatile, d.Multi = true, true // Jan:Dec!B2, across sheets
			default:
				// A column (A) or row (3) alone is a reference only in a
				// range; Tax is a name though TAX is a column.
				if p, ok := parseRef(w); ok && (p.kind == cellRef || text(i+1) == ":" || text(i-1) == ":") {
					continue
				}
				u := strings.ToUpper(bare(w))
				if u == "TRUE" || u == "FALSE" || params[u] {
					continue
				}
				d.addName(w, seen)
			}
		}
	}
	return d
}

func (d *Deps) addName(w string, seen map[string]bool) {
	n := bare(w)
	if k := strings.ToUpper(n); !seen[k] {
		seen[k] = true
		d.Names = append(d.Names, n)
	}
}

// shapeFns read where their reference arguments are, not what they hold.
var shapeFns = setOf(`ROW ROWS COLUMN COLUMNS ISREF AREAS`)

// ValueRefs are the references whose values a formula reads: Refs, but
// for a reference given whole to a function that reads only where it is
// (=ROWS($A$2:A2) in A2 reads no cell, and refers to itself harmlessly).
func ValueRefs(f string) []Ref {
	ts := lex(f)
	changed := false
	for i := 0; i < len(ts); i++ {
		if !ts[i].call || !shapeFns[strings.ToUpper(bare(ts[i].text))] || i+1 >= len(ts) || ts[i+1].text != "(" {
			continue
		}
		// Its arguments, at depth 1.
		depth, start := 0, i+2
	args:
		for j := i + 1; j < len(ts); j++ {
			switch ts[j].text {
			case "(":
				depth++
				continue
			case ")":
				depth--
				if depth > 0 {
					continue
				}
			case ",":
				if depth != 1 {
					continue
				}
			default:
				continue
			}
			if start < j && pureRef(ts[start:j]) {
				for k := start; k < j; k++ {
					ts[k].text = ""
				}
				ts[start].text = "0"
				changed = true
			}
			start = j + 1
			if depth == 0 {
				break args
			}
		}
	}
	if !changed {
		return Refs(f)
	}
	return Refs(join(ts))
}

// pureRef tells whether tokens are a reference alone: Sheet!$A$2:B9.
func pureRef(ts []token) bool {
	has := false
	for _, t := range ts {
		switch {
		case strings.TrimSpace(t.text) == "":
		case t.call:
			return false
		case t.word:
			has = true
		case t.text == ":" || t.text == "!" || strings.HasPrefix(t.text, "'"):
		default:
			return false
		}
	}
	return has
}

var errorCodes = []string{"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A",
	"#GETTING_DATA", "#SPILL!", "#CALC!", "#FIELD!", "#BLOCKED!", "#CONNECT!", "#BUSY!", "#UNKNOWN!"}

// errorLiteral returns the index of the last token of the error literal
// (#N/A, #DIV/0!) that starts at ts[i], or i when none does.
func errorLiteral(ts []token, i int) int {
	end, text := i, "#"
	for j := i + 1; j < len(ts); j++ {
		text += strings.ToUpper(ts[j].text)
		prefix := false
		for _, code := range errorCodes {
			if text == code {
				end = j
			}
			prefix = prefix || strings.HasPrefix(code, text)
		}
		if !prefix {
			break
		}
	}
	return end
}

// ErrorLiterals lists the error values written into a formula (#REF!,
// #N/A), outside strings.
func ErrorLiterals(f string) []string {
	ts := lex(f)
	var out []string
	for i := 0; i < len(ts); i++ {
		if ts[i].text != "#" {
			continue
		}
		end := errorLiteral(ts, i)
		if end > i {
			out = append(out, strings.ToUpper(join(ts[i:end+1])))
			i = end
		}
	}
	return out
}

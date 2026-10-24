package formula

import (
	"strconv"
	"strings"
)

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

// rangeArgs gives, for functions that take ranges of several cells as
// such, which of their arguments (from 0) do: "*" all, "1+" the second
// on, "odd" the second, fourth and so on, "even" the first, third and so
// on. A range given to any other argument is read as an array.
var rangeArgs = func() map[string]string {
	m := map[string]string{}
	for spec, fns := range map[string]string{
		"*": `SUM SUMSQ SUMPRODUCT AVERAGE AVERAGEA COUNT COUNTA MAX MAXA MIN MINA MEDIAN MODE MODE.SNGL
			MODE.MULT STDEV STDEV.S STDEV.P STDEVA STDEVP STDEVPA VAR VAR.S VAR.P VARA VARP VARPA PRODUCT
			AVEDEV DEVSQ GEOMEAN HARMEAN KURT SKEW CORREL COVAR COVARIANCE.P COVARIANCE.S SLOPE INTERCEPT
			RSQ PEARSON STEYX CONCAT AND OR XOR SUMX2MY2 SUMX2PY2 SUMXMY2 VSTACK HSTACK MMULT FREQUENCY`,
		"0": `COUNTIF COUNTBLANK LARGE SMALL PERCENTILE PERCENTILE.INC PERCENTILE.EXC QUARTILE QUARTILE.INC
			QUARTILE.EXC PERCENTRANK PERCENTRANK.INC PERCENTRANK.EXC TRIMMEAN INDEX ROWS COLUMNS IRR MIRR
			UNIQUE SORT TRANSPOSE TAKE DROP CHOOSECOLS CHOOSEROWS TOCOL TOROW WRAPROWS WRAPCOLS EXPAND
			MINVERSE MDETERM OFFSET Z.TEST ZTEST`,
		"0 1":   `FILTER XIRR LINEST LOGEST T.TEST TTEST F.TEST FTEST CHISQ.TEST CHITEST PROB`,
		"0 1 2": `TREND GROWTH`,
		"0 2": `SUMIF AVERAGEIF DSUM DAVERAGE DCOUNT DCOUNTA DGET DMAX DMIN DPRODUCT DSTDEV DSTDEVP DVAR
			DVARP`,
		"even":  `COUNTIFS`,
		"0 odd": `SUMIFS AVERAGEIFS MAXIFS MINIFS SORTBY`,
		"1":     `VLOOKUP HLOOKUP MATCH XMATCH RANK RANK.EQ RANK.AVG`,
		"1 2":   `LOOKUP XLOOKUP FORECAST FORECAST.LINEAR XNPV`,
		"1+":    `NPV SUBTOTAL`,
		"2+":    `TEXTJOIN AGGREGATE`,
		"2":     `NETWORKDAYS WORKDAY`,
		"3":     `NETWORKDAYS.INTL WORKDAY.INTL SERIESSUM`,
	} {
		for fn := range setOf(fns) {
			m[fn] = spec
		}
	}
	return m
}()

// takesRange tells whether argument i (from 0) of fn takes a range of
// several cells as such.
func takesRange(fn string, i int) bool {
	spec, ok := rangeArgs[fn]
	if !ok {
		return false
	}
	for _, s := range strings.Fields(spec) {
		switch {
		case s == "*":
			return true
		case s == "odd" && i%2 == 1, s == "even" && i%2 == 0:
			return true
		case strings.HasSuffix(s, "+"):
			if n, err := strconv.Atoi(strings.TrimSuffix(s, "+")); err == nil && i >= n {
				return true
			}
		default:
			if n, err := strconv.Atoi(s); err == nil && i == n {
				return true
			}
		}
	}
	return false
}

var operators = setOf(`+ - * / ^ & = < > %`)

// ArrayUse tells whether a formula reads a range of several cells the way
// only array evaluation does: in arithmetic or a comparison
// ((A2:A9="East")*B2:B9), alone (=A2:A9, which spills), or as an argument
// of a function that takes single values. It returns why, or "".
// multi tells whether a workbook name stands for several cells.
func ArrayUse(f string, multi func(name string) bool) string {
	var ts []token
	for _, t := range lex(f) {
		if strings.TrimSpace(t.text) != "" || strings.HasPrefix(t.text, `"`) {
			ts = append(ts, t)
		}
	}
	text := func(i int) string {
		if i < 0 || i >= len(ts) {
			return ""
		}
		return ts[i].text
	}
	for i := 0; i < len(ts); i++ {
		start, end := -1, -1 // the tokens of a range of several cells
		t := ts[i]
		switch {
		case t.word && !t.call && text(i+1) == ":" && i+2 < len(ts) && ts[i+2].word:
			a, ok1 := parseRef(t.text)
			b, ok2 := parseRef(ts[i+2].text)
			if ok1 && ok2 && (a.kind != cellRef || a.row != b.row || a.col != b.col) {
				start, end = i, i+2
				if text(i-1) == "!" {
					start = i - 2
				}
			}
		case t.word && !t.call && text(i+1) != "!" && text(i+1) != ":" && text(i-1) != ":" && multi != nil && multi(bare(t.text)):
			start, end = i, i
		}
		if start < 0 {
			continue
		}
		i = end
		// Out of parentheses that only wrap it: (A1:A9)*2.
		for text(start-1) == "(" && text(end+1) == ")" && !(start >= 2 && ts[start-2].call) {
			start, end = start-1, end+1
		}
		prev, next := text(start-1), text(end+1)
		switch {
		case operators[prev] || operators[next]:
			return "it computes with a range of cells as a whole (array arithmetic), which the sheets calculator cannot (it can SUMIFS, COUNTIFS, and SUMPRODUCT of plain ranges)"
		case prev == "" && next == "":
			return "it is a range of cells alone, which Excel spills"
		case prev == "(" || prev == ",":
			fn, arg := enclosingCall(ts, start)
			if !takesRange(fn, arg) {
				to := ""
				if fn != "" {
					to = " to " + fn
				}
				return "it passes a range of cells" + to + " where a single value is expected, which Excel answers with an array"
			}
		}
	}
	return ""
}

// enclosingCall names the function whose parentheses hold token i, and
// which of its arguments (from 0) holds it.
func enclosingCall(ts []token, i int) (string, int) {
	depth, arg := 0, 0
	for k := i - 1; k >= 0; k-- {
		switch ts[k].text {
		case ")", "}":
			depth++
		case "{":
			depth--
		case ",":
			if depth == 0 {
				arg++
			}
		case "(":
			if depth == 0 {
				if k > 0 && ts[k-1].call {
					return strings.ToUpper(bare(ts[k-1].text)), arg
				}
				return "", arg
			}
			depth--
		}
	}
	return "", arg
}

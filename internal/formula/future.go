package formula

import (
	"strings"

	"github.com/eemax/sheets/internal/grid"
)

// Functions added to Excel after the 2007 file format are stored with a
// prefix; without it Excel shows #NAME? until the cell is re-entered.
// FILTER and SORT carry an extra _xlws. part.
var future = func() map[string]string {
	m := map[string]string{"FILTER": "_xlfn._xlws.", "SORT": "_xlfn._xlws."}
	for _, f := range strings.Fields(`
		ACOT ACOTH AGGREGATE ARABIC ARRAYTOTEXT BASE BETA.DIST BETA.INV BINOM.DIST
		BINOM.DIST.RANGE BINOM.INV BITAND BITLSHIFT BITOR BITRSHIFT BITXOR BYCOL BYROW
		CEILING.MATH CEILING.PRECISE CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT
		CHISQ.TEST CHOOSECOLS CHOOSEROWS COMBINA CONCAT CONFIDENCE.NORM CONFIDENCE.T COT
		COTH COVARIANCE.P COVARIANCE.S CSC CSCH DAYS DECIMAL DROP ECMA.CEILING ERF.PRECISE
		ERFC.PRECISE EXPAND EXPON.DIST F.DIST F.DIST.RT F.INV F.INV.RT F.TEST FIELDVALUE
		FILTERXML FLOOR.MATH FLOOR.PRECISE FORECAST.ETS FORECAST.ETS.CONFINT
		FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT FORECAST.LINEAR FORMULATEXT GAMMA
		GAMMA.DIST GAMMA.INV GAMMALN.PRECISE GAUSS GROUPBY HSTACK HYPGEOM.DIST IFNA IFS
		IMAGE IMCOSH IMCOT IMCSC IMCSCH IMSEC IMSECH IMSINH IMTAN ISFORMULA ISOMITTED
		ISOWEEKNUM LAMBDA LET LOGNORM.DIST LOGNORM.INV MAKEARRAY MAP MAXIFS MINIFS
		MODE.MULT MODE.SNGL MUNIT NEGBINOM.DIST NETWORKDAYS.INTL NORM.DIST NORM.INV
		NORM.S.DIST NORM.S.INV NUMBERVALUE PDURATION PERCENTILE.EXC PERCENTILE.INC
		PERCENTOF PERCENTRANK.EXC PERCENTRANK.INC PERMUTATIONA PHI PIVOTBY POISSON.DIST
		QUARTILE.EXC QUARTILE.INC QUERYSTRING RANDARRAY RANK.AVG RANK.EQ REDUCE
		REGEXEXTRACT REGEXREPLACE REGEXTEST RRI SCAN SEC SECH SEQUENCE SHEET SHEETS
		SKEW.P SORTBY STDEV.P STDEV.S SWITCH T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T
		T.TEST TAKE TEXTAFTER TEXTBEFORE TEXTJOIN TEXTSPLIT TOCOL TOROW TRIMRANGE UNICHAR
		UNICODE UNIQUE VALUETOTEXT VAR.P VAR.S VSTACK WEBSERVICE WEIBULL.DIST
		WORKDAY.INTL WRAPCOLS WRAPROWS XLOOKUP XMATCH XOR Z.TEST`) {
		m[f] = "_xlfn."
	}
	return m
}()

// spilling functions return arrays; Excel 365 stores formulas using them as
// dynamic arrays so the result spills into neighbouring cells.
var spilling = setOf(`
	BYCOL BYROW CHOOSECOLS CHOOSEROWS DROP EXPAND FILTER FREQUENCY GROUPBY GROWTH
	HSTACK LINEST LOGEST MAKEARRAY MAP MINVERSE MMULT MUNIT PIVOTBY RANDARRAY
	REGEXEXTRACT SCAN SEQUENCE SORT SORTBY TAKE TEXTSPLIT TOCOL TOROW TRANSPOSE TREND
	UNIQUE VSTACK WRAPCOLS WRAPROWS`)

func setOf(s string) map[string]bool {
	m := map[string]bool{}
	for _, f := range strings.Fields(s) {
		m[f] = true
	}
	return m
}

// Prefixes used in stored formulas.
const (
	xlfn = "_xlfn."
	xlws = "_xlws."
	xlpm = "_xlpm."
)

type token struct {
	text string
	word bool // identifier, function name, reference or number
	call bool // word directly followed by '('
}

// lex splits a formula into words and everything else, keeping string
// literals, quoted sheet names and structured references whole.
func lex(f string) []token {
	var out []token
	for i := 0; i < len(f); {
		c := f[i]
		switch {
		case c == '"' || c == '\'':
			j := skipQuoted(f, i, c)
			out = append(out, token{text: f[i:j]})
			i = j
		case c == '[':
			j, depth := i, 0
			for j < len(f) {
				if f[j] == '[' {
					depth++
				} else if f[j] == ']' {
					depth--
					if depth == 0 {
						j++
						break
					}
				}
				j++
			}
			out = append(out, token{text: f[i:j]})
			i = j
		case isWordChar(c) || c == '$' || c == '\\':
			j := i
			for j < len(f) && (isWordChar(f[j]) || f[j] == '$' || f[j] == '\\') {
				j++
			}
			out = append(out, token{text: f[i:j], word: true, call: j < len(f) && f[j] == '('})
			i = j
		default:
			out = append(out, token{text: f[i : i+1]})
			i++
		}
	}
	return out
}

func join(ts []token) string {
	var b strings.Builder
	for _, t := range ts {
		b.WriteString(t.text)
	}
	return b.String()
}

// bare strips storage prefixes from a word.
func bare(w string) string {
	for _, p := range []string{xlfn, xlws, xlpm} {
		if len(w) > len(p) && strings.EqualFold(w[:len(p)], p) {
			w = w[len(p):]
		}
	}
	return w
}

// ToFile converts a formula as written by a person into the form the file
// format stores: future functions get their _xlfn. prefix and LET/LAMBDA
// parameter names their _xlpm. prefix. Already-prefixed text is kept.
func ToFile(f string) string {
	ts := lex(f)
	params := lambdaParams(ts)
	changed := false
	for i, t := range ts {
		if !t.word || i > 0 && ts[i-1].text == "!" {
			continue
		}
		name := strings.ToUpper(bare(t.text))
		if params[name] { // also when called: LET(f, LAMBDA(x, x+1), f(2))
			ts[i].text = xlpm + bare(t.text)
			changed = true
		} else if p, ok := future[name]; ok && t.call {
			ts[i].text = p + name
			changed = true
		}
	}
	if !changed {
		return f
	}
	return join(ts)
}

// FromFile removes storage prefixes so formulas read as typed in Excel.
func FromFile(f string) string {
	if !strings.Contains(f, "_xl") {
		return f
	}
	ts := lex(f)
	for i, t := range ts {
		if t.word {
			ts[i].text = bare(t.text)
		}
	}
	return join(ts)
}

// Spills reports whether a formula uses a function that returns an array,
// so it must be stored as a dynamic array formula.
func Spills(f string) bool {
	for _, t := range lex(f) {
		if t.call && spilling[strings.ToUpper(bare(t.text))] {
			return true
		}
	}
	return false
}

// lambdaParams collects the names LET and LAMBDA declare: for LET every
// other argument but the last, for LAMBDA every argument but the last.
func lambdaParams(ts []token) map[string]bool {
	names := map[string]bool{}
	for i, t := range ts {
		if !t.call {
			continue
		}
		fn := strings.ToUpper(bare(t.text))
		if fn != "LET" && fn != "LAMBDA" {
			continue
		}
		args := topLevelArgs(ts, i+1)
		for k, arg := range args {
			if k == len(args)-1 || fn == "LET" && k%2 == 1 {
				continue
			}
			if len(arg) == 1 && arg[0].word && !arg[0].call {
				names[strings.ToUpper(bare(arg[0].text))] = true
			}
		}
	}
	return names
}

// topLevelArgs splits the arguments of the call whose '(' is ts[open],
// dropping surrounding whitespace tokens.
func topLevelArgs(ts []token, open int) [][]token {
	var args [][]token
	var cur []token
	depth := 0
	for i := open; i < len(ts); i++ {
		switch ts[i].text {
		case "(", "{":
			depth++
			if depth == 1 {
				continue
			}
		case ")", "}":
			depth--
			if depth == 0 {
				return append(args, trim(cur))
			}
		case ",", ";":
			if depth == 1 {
				args = append(args, trim(cur))
				cur = nil
				continue
			}
		}
		cur = append(cur, ts[i])
	}
	return append(args, trim(cur))
}

func trim(ts []token) []token {
	for len(ts) > 0 && strings.TrimSpace(ts[0].text) == "" {
		ts = ts[1:]
	}
	for len(ts) > 0 && strings.TrimSpace(ts[len(ts)-1].text) == "" {
		ts = ts[:len(ts)-1]
	}
	return ts
}

// sheetPrefix finds the sheet-name token before each '!' (Data!A1 or
// 'My Data'!A1), skipping external references ([1]Data!A1).
func sheetPrefix(ts []token, i int) (string, bool) {
	if i+1 >= len(ts) || ts[i+1].text != "!" {
		return "", false
	}
	if i > 0 && strings.HasSuffix(ts[i-1].text, "]") {
		return "", false
	}
	t := ts[i].text
	if strings.HasPrefix(t, "'") && strings.HasSuffix(t, "'") && len(t) >= 2 {
		return strings.ReplaceAll(t[1:len(t)-1], "''", "'"), true
	}
	if ts[i].word {
		return t, true
	}
	return "", false
}

// RenameSheet rewrites references to sheet from as references to sheet to.
func RenameSheet(f, from, to string) string {
	if !strings.Contains(f, "!") {
		return f
	}
	ts := lex(f)
	changed := false
	for i := range ts {
		if name, ok := sheetPrefix(ts, i); ok && strings.EqualFold(name, from) {
			ts[i].text = grid.QuoteSheet(to)
			changed = true
		}
	}
	if !changed {
		return f
	}
	return join(ts)
}

// RefersToSheet reports whether a formula references the named sheet.
func RefersToSheet(f, sheet string) bool {
	if !strings.Contains(f, "!") {
		return false
	}
	ts := lex(f)
	for i := range ts {
		if name, ok := sheetPrefix(ts, i); ok && strings.EqualFold(name, sheet) {
			return true
		}
	}
	return false
}

// DropSheet replaces references to a deleted sheet with #REF!, as Excel
// does.
func DropSheet(f, sheet string) string {
	if !strings.Contains(f, "!") {
		return f
	}
	ts := lex(f)
	var out []token
	changed := false
	for i := 0; i < len(ts); i++ {
		if name, ok := sheetPrefix(ts, i); ok && strings.EqualFold(name, sheet) {
			// Swallow Sheet ! ref [: ref].
			j := i + 2
			if j < len(ts) && ts[j].word {
				j++
				if j+1 < len(ts) && ts[j].text == ":" && ts[j+1].word {
					j += 2
				}
			}
			out = append(out, token{text: "#REF!"})
			i = j - 1
			changed = true
			continue
		}
		out = append(out, ts[i])
	}
	if !changed {
		return f
	}
	return join(out)
}

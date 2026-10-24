package xlsx

import (
	"fmt"
	"math"
	"regexp"
	"strconv"
	"strings"

	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
	"github.com/xuri/excelize/v2"
)

// A ValidationRule is a data validation rule: what a cell accepts.
type ValidationRule struct {
	Words   []string // the rule as given
	Prompt  string   // shown when the cell is selected
	Message string   // shown when a value is refused
	Warn    bool     // warn instead of refusing
	NoBlank bool     // refuse an empty cell too
}

const validateHelp = `list A,B,C | list =Sheet!$A$1:$A$9 | whole|number|date|time|length OP X (OP: < <= > >= = <>) | whole|number|date|time|length between X and Y | formula '=COUNTIF($A:$A,A2)=1' | any`

// dataValidation builds the rule for range r.
func (c ValidationRule) dataValidation(r grid.Range) (*excelize.DataValidation, error) {
	ws := c.Words
	if len(ws) == 0 {
		return nil, fmt.Errorf("no rule: give one of %s", validateHelp)
	}
	dv := excelize.NewDataValidation(!c.NoBlank)
	dv.Sqref = r.String()
	kind := strings.ToLower(ws[0])
	rest := ws[1:]
	switch kind {
	case "list":
		items := rest
		if len(items) == 1 && strings.Contains(items[0], ",") {
			items = strings.Split(items[0], ",")
		}
		if len(items) == 0 {
			return nil, fmt.Errorf("list: give the choices (list North,South,East) or where they are (list =Lists!$A$2:$A$9)")
		}
		if len(items) == 1 && (strings.HasPrefix(items[0], "=") || strings.Contains(items[0], "!") || isRangeRef(items[0])) {
			src := strings.TrimPrefix(items[0], "=")
			if err := formula.Check(src); err != nil {
				return nil, fmt.Errorf("list =%s: %v", src, err)
			}
			dv.SetSqrefDropList(string(xmlEscape(src)))
			break
		}
		for i, it := range items {
			items[i] = strings.TrimSpace(it)
			if strings.Contains(items[i], `"`) {
				return nil, fmt.Errorf("list: choice %q holds a double quote, which a list cannot; put the choices in cells and give their range", it)
			}
		}
		if err := dv.SetDropList(items); err != nil {
			return nil, fmt.Errorf("list: the choices come to more than 255 characters; put them in cells and give their range (list =Lists!$A$2:$A$40)")
		}
	case "whole", "integer", "number", "decimal", "date", "time", "length", "textlength":
		t := map[string]excelize.DataValidationType{"whole": excelize.DataValidationTypeWhole, "integer": excelize.DataValidationTypeWhole,
			"number": excelize.DataValidationTypeDecimal, "decimal": excelize.DataValidationTypeDecimal, "date": excelize.DataValidationTypeDate,
			"time": excelize.DataValidationTypeTime, "length": excelize.DataValidationTypeTextLength, "textlength": excelize.DataValidationTypeTextLength}[kind]
		op, f1, f2, err := parseBounds(kind, rest)
		if err != nil {
			return nil, err
		}
		if err := dv.SetRange(string(xmlEscape(f1)), string(xmlEscape(f2)), t, op); err != nil {
			return nil, err
		}
		if f2 == "" {
			dv.Formula2 = ""
		}
	case "formula", "custom":
		if len(rest) != 1 {
			return nil, fmt.Errorf("formula: give one formula for the range's first cell, like formula '=COUNTIF($A:$A,A2)=1'")
		}
		f := strings.TrimPrefix(strings.TrimSpace(rest[0]), "=")
		if err := formula.Check(f); err != nil {
			return nil, fmt.Errorf("=%s is not a valid formula: %v", f, err)
		}
		dv.Type, dv.Formula1 = "custom", string(xmlEscape(formula.ToFile(f)))
	case "any":
		if c.Prompt == "" {
			return nil, fmt.Errorf("any accepts every value; it only makes sense with --prompt")
		}
	default:
		return nil, fmt.Errorf("%q is not a rule: give one of %s", ws[0], validateHelp)
	}
	if c.Prompt != "" {
		dv.SetInput("", c.Prompt)
	}
	// Without an error alert, Excel lets any value in.
	style := excelize.DataValidationErrorStyleStop
	if c.Warn {
		style = excelize.DataValidationErrorStyleWarning
	}
	dv.SetError(style, "", c.Message)
	if c.Message == "" {
		dv.Error, dv.ErrorTitle = nil, nil
	}
	return dv, nil
}

func isRangeRef(s string) bool {
	_, err := grid.ParseRange(strings.ReplaceAll(s, "$", ""))
	return err == nil && strings.Contains(s, ":")
}

// parseBounds reads OP X or between X and Y, for a rule of kind.
func parseBounds(kind string, ws []string) (excelize.DataValidationOperator, string, string, error) {
	usage := fmt.Errorf("%s: expected OP X (OP: < <= > >= = <>) or between X and Y", kind)
	if len(ws) == 1 {
		if op := leadingOp(ws[0]); op != "" && len(ws[0]) > len(op) {
			ws = []string{op, strings.TrimSpace(ws[0][len(op):])}
		}
	}
	val := func(s string) (string, error) {
		s = strings.TrimSpace(s)
		switch {
		case strings.HasPrefix(s, "="):
			f := strings.TrimPrefix(s, "=")
			if err := formula.Check(f); err != nil {
				return "", fmt.Errorf("%s is not a valid formula: %v", s, err)
			}
			return formula.ToFile(f), nil
		case kind == "date":
			if m := isoDate.FindStringSubmatch(s); m != nil {
				return fmt.Sprintf("DATE(%s,%s,%s)", m[1], strings.TrimLeft(m[2], "0"), strings.TrimLeft(m[3], "0")), nil
			}
			if strings.EqualFold(s, "today") {
				return "TODAY()", nil
			}
			return "", fmt.Errorf("date: %q is not a date (YYYY-MM-DD, today, or =formula)", s)
		case kind == "time":
			if t, ok := parseClock(s); ok {
				return strconv.FormatFloat(t, 'f', -1, 64), nil
			}
			return "", fmt.Errorf("time: %q is not a time (HH:MM or HH:MM:SS)", s)
		}
		if _, err := strconv.ParseFloat(s, 64); err != nil {
			return "", fmt.Errorf("%s: %q is not a number", kind, s)
		}
		return s, nil
	}
	ops := map[string]excelize.DataValidationOperator{"<": excelize.DataValidationOperatorLessThan, "<=": excelize.DataValidationOperatorLessThanOrEqual,
		">": excelize.DataValidationOperatorGreaterThan, ">=": excelize.DataValidationOperatorGreaterThanOrEqual,
		"=": excelize.DataValidationOperatorEqual, "==": excelize.DataValidationOperatorEqual,
		"<>": excelize.DataValidationOperatorNotEqual, "!=": excelize.DataValidationOperatorNotEqual}
	switch {
	case len(ws) == 2:
		op, ok := ops[ws[0]]
		if !ok {
			return 0, "", "", usage
		}
		v, err := val(ws[1])
		return op, v, "", err
	case len(ws) == 4 && strings.EqualFold(ws[0], "between") && strings.EqualFold(ws[2], "and"),
		len(ws) == 5 && strings.EqualFold(ws[0], "not") && strings.EqualFold(ws[1], "between") && strings.EqualFold(ws[3], "and"):
		op := excelize.DataValidationOperatorBetween
		if len(ws) == 5 {
			op, ws = excelize.DataValidationOperatorNotBetween, ws[1:]
		}
		lo, err := val(ws[1])
		if err != nil {
			return 0, "", "", err
		}
		hi, err := val(ws[3])
		return op, lo, hi, err
	}
	return 0, "", "", usage
}

// parseClock reads HH:MM[:SS] as a fraction of a day.
func parseClock(s string) (float64, bool) {
	parts := strings.Split(s, ":")
	if len(parts) < 2 || len(parts) > 3 {
		return 0, false
	}
	var v [3]int
	for i, p := range parts {
		n, err := strconv.Atoi(p)
		if err != nil || n < 0 || i == 0 && n > 23 || i > 0 && n > 59 {
			return 0, false
		}
		v[i] = n
	}
	return float64(v[0]*3600+v[1]*60+v[2]) / 86400, true
}

// A Validation is a data validation rule as it stands in a sheet.
type Validation struct {
	Range  string `json:"range"`
	Rule   string `json:"rule"`
	Prompt string `json:"prompt,omitempty"`
	Error  string `json:"error,omitempty"`
	Warn   bool   `json:"warn,omitempty"` // warns, letting the value in
}

var dateCall = regexp.MustCompile(`^DATE\((\d{4}),(\d{1,2}),(\d{1,2})\)$`)

// boundText shows a rule's bound as validate takes it: a date as
// 2026-03-01, a time as 08:30.
func boundText(kind, f string) string {
	switch kind {
	case "date":
		if m := dateCall.FindStringSubmatch(f); m != nil {
			mo, _ := strconv.Atoi(m[2])
			d, _ := strconv.Atoi(m[3])
			return fmt.Sprintf("%s-%02d-%02d", m[1], mo, d)
		}
		if strings.EqualFold(f, "TODAY()") {
			return "today"
		}
	case "time":
		if v, err := strconv.ParseFloat(f, 64); err == nil && v >= 0 && v < 1 {
			s := int(math.Round(v * 86400))
			if s%60 != 0 {
				return fmt.Sprintf("%02d:%02d:%02d", s/3600, s/60%60, s%60)
			}
			return fmt.Sprintf("%02d:%02d", s/3600, s/60%60)
		}
	}
	if strings.HasPrefix(f, "=") || strings.ContainsAny(f, "(!$") || !isNumber(f) {
		return "=" + formula.FromFile(strings.TrimPrefix(f, "="))
	}
	return f
}

func isNumber(s string) bool {
	_, err := strconv.ParseFloat(s, 64)
	return err == nil
}

// validations lists a sheet's data validation rules.
func validations(f *excelize.File, sheet string) ([]Validation, error) {
	dvs, err := f.GetDataValidations(sheet)
	if err != nil {
		return nil, err
	}
	ops := map[string]string{"lessThan": "<", "lessThanOrEqual": "<=", "greaterThan": ">", "greaterThanOrEqual": ">=", "equal": "=", "notEqual": "<>"}
	kinds := map[string]string{"whole": "whole", "decimal": "number", "date": "date", "time": "time", "textLength": "length"}
	var out []Validation
	for _, dv := range dvs {
		v := Validation{Range: dv.Sqref}
		f1, f2 := unescapeXML(dv.Formula1), unescapeXML(dv.Formula2)
		switch dv.Type {
		case "list":
			if strings.HasPrefix(f1, `"`) {
				v.Rule = "list " + strings.Trim(f1, `"`)
			} else {
				v.Rule = "list =" + f1
			}
		case "custom":
			v.Rule = "formula =" + formula.FromFile(f1)
		case "":
			v.Rule = "any"
		default:
			k := kinds[dv.Type]
			if k == "" {
				k = dv.Type
			}
			switch dv.Operator {
			case "between", "", "notBetween":
				not := ""
				if dv.Operator == "notBetween" {
					not = "not "
				}
				v.Rule = fmt.Sprintf("%s %sbetween %s and %s", k, not, boundText(k, f1), boundText(k, f2))
			default:
				v.Rule = fmt.Sprintf("%s %s %s", k, ops[dv.Operator], boundText(k, f1))
			}
		}
		if dv.Prompt != nil {
			v.Prompt = *dv.Prompt
		}
		if dv.Error != nil {
			v.Error = *dv.Error
		}
		v.Warn = dv.ErrorStyle != nil && *dv.ErrorStyle == "warning"
		out = append(out, v)
	}
	return out, nil
}

func unescapeXML(s string) string {
	return strings.NewReplacer("&lt;", "<", "&gt;", ">", "&quot;", `"`, "&apos;", "'", "&#34;", `"`, "&#39;", "'", "&amp;", "&").Replace(s)
}

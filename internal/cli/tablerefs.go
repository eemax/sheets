package cli

import (
	"fmt"
	"strings"

	"github.com/eemax/sheets/internal/analyze"
	"github.com/eemax/sheets/internal/grid"
)

// sheetTableRefs turns references to sheets' own tables in a formula
// (Sales[Revenue], where Sales is a sheet holding one table but no Excel
// table is called Sales) into the ranges they stand for, absolute. Excel
// reads Sales[Revenue] only as a column of an Excel table (Insert ›
// Table), and cannot evaluate a formula naming a table that does not
// exist. References to Excel tables are left as they are.
func sheetTableRefs(wb *grid.Workbook, text string) (string, error) {
	if !strings.Contains(text, "[") {
		return text, nil
	}
	var out strings.Builder
	i := 0
	for i < len(text) {
		c := text[i]
		if c == '"' {
			j := i + 1
			for j < len(text) {
				if text[j] == '"' {
					if j+1 < len(text) && text[j+1] == '"' {
						j += 2
						continue
					}
					break
				}
				j++
			}
			end := min(j+1, len(text))
			out.WriteString(text[i:end])
			i = end
			continue
		}
		if !isNameStart(c) || i > 0 && (isNameChar(text[i-1]) || text[i-1] == '!' || text[i-1] == '\'') {
			out.WriteByte(c)
			i++
			continue
		}
		j := i
		for j < len(text) && isNameChar(text[j]) {
			j++
		}
		if j >= len(text) || text[j] != '[' {
			out.WriteString(text[i:j])
			i = j
			continue
		}
		end := closingBracket(text, j)
		if end < 0 {
			return "", fmt.Errorf("=%s: a [ is not closed", text)
		}
		name, ref := text[i:j], text[i:end+1]
		if excelTable(wb, name) {
			out.WriteString(ref)
			i = end + 1
			continue
		}
		sheet, rg, err := analyze.Resolve(wb, ref)
		if err != nil || rg == nil {
			return "", fmt.Errorf("=%s: %s is not a column of an Excel table, and %v; write the range instead (Sales!F2:F61)", text, ref, err)
		}
		out.WriteString(absoluteRef(sheet, *rg))
		i = end + 1
	}
	return out.String(), nil
}

func isNameStart(c byte) bool {
	return c >= 'A' && c <= 'Z' || c >= 'a' && c <= 'z' || c == '_' || c >= 0x80
}

func isNameChar(c byte) bool {
	return isNameStart(c) || c >= '0' && c <= '9' || c == '.'
}

// closingBracket finds the ] closing the [ at i, past nested brackets.
func closingBracket(s string, i int) int {
	depth := 0
	for j := i; j < len(s); j++ {
		switch s[j] {
		case '[':
			depth++
		case ']':
			depth--
			if depth == 0 {
				return j
			}
		}
	}
	return -1
}

func excelTable(wb *grid.Workbook, name string) bool {
	for _, s := range wb.Sheets {
		for _, t := range s.Tables {
			if strings.EqualFold(t.Name, name) {
				return true
			}
		}
	}
	return false
}

// absoluteRef writes a range as $A$1:$B$2, with its sheet.
func absoluteRef(sheet string, r grid.Range) string {
	cell := func(row, col int) string { return "$" + grid.ColName(col) + "$" + fmt.Sprint(row) }
	ref := cell(r.R1, r.C1)
	if r.R1 != r.R2 || r.C1 != r.C2 {
		ref += ":" + cell(r.R2, r.C2)
	}
	return grid.QuoteSheet(sheet) + "!" + ref
}

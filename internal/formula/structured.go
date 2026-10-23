package formula

import (
	"strings"
	"unicode/utf8"
)

// RenameTableColumns rewrites structured references to the columns of an
// Excel table after some of them were renamed, as Excel does when a header
// changes: Sales[Units], Sales[[#This Row],[Units]], Sales[[Units]:[Price]]
// and, in formulas inside the table (inTable), the bare [Units] and
// [@Units]. renames maps lower-cased old names to new ones. Text in strings
// and quoted sheet names is left alone, as are references to workbooks
// ([1]Sheet1!A1).
func RenameTableColumns(f, table string, inTable bool, renames map[string]string) string {
	var out strings.Builder
	changed := false
	for i := 0; i < len(f); {
		switch ch := f[i]; ch {
		case '"', '\'':
			j := closeQuote(f, i)
			out.WriteString(f[i:j])
			i = j
		case '[':
			end := closeBracket(f, i)
			if end < 0 {
				out.WriteString(f[i:])
				i = len(f)
				continue
			}
			j := i
			for j > 0 && identByte(f[j-1]) {
				j--
			}
			group := f[i:end]
			var mine bool
			if j < i {
				mine = strings.EqualFold(f[j:i], table)
			} else {
				// Unqualified: this table's, unless it names a workbook.
				mine = inTable && (end == len(f) || !identByte(f[end]) && f[end] != '\'' && f[end] != '!')
			}
			if mine {
				if g := renameGroup(group, renames); g != group {
					group, changed = g, true
				}
			}
			out.WriteString(group)
			i = end
		default:
			out.WriteByte(ch)
			i++
		}
	}
	if !changed {
		return f
	}
	return out.String()
}

// closeQuote returns the index after the quoted text starting at i; a
// doubled quote inside is an escaped one.
func closeQuote(f string, i int) int {
	q := f[i]
	for j := i + 1; j < len(f); j++ {
		if f[j] == q {
			if j+1 < len(f) && f[j+1] == q {
				j++
				continue
			}
			return j + 1
		}
	}
	return len(f)
}

// closeBracket returns the index after the bracket matching the one at i,
// or -1. Inside, ' escapes the next character.
func closeBracket(f string, i int) int {
	depth := 0
	for j := i; j < len(f); j++ {
		switch f[j] {
		case '\'':
			j++
		case '[':
			depth++
		case ']':
			if depth--; depth == 0 {
				return j + 1
			}
		}
	}
	return -1
}

// identByte reports whether b can be part of a table name.
func identByte(b byte) bool {
	return b == '_' || b == '.' || b == '\\' || b >= '0' && b <= '9' || b >= 'A' && b <= 'Z' || b >= 'a' && b <= 'z' || b >= utf8.RuneSelf
}

// renameGroup renames the columns in one bracketed specifier: [Units],
// [@Units], [@[Unit price]], [[#Totals],[Units]], [[Units]:[Price]].
func renameGroup(group string, renames map[string]string) string {
	inner := group[1 : len(group)-1]
	switch {
	case strings.HasPrefix(inner, "@["):
		return "[@" + renameItems(inner[1:], renames) + "]"
	case strings.HasPrefix(inner, "@"):
		if to, ok := renames[strings.ToLower(unescapeColumn(inner[1:]))]; ok {
			if plainColumn(to) {
				return "[@" + to + "]"
			}
			return "[@[" + escapeColumn(to) + "]]"
		}
		return group
	case strings.HasPrefix(inner, "["):
		return "[" + renameItems(inner, renames) + "]"
	case strings.HasPrefix(inner, "#"):
		return group
	}
	if to, ok := renames[strings.ToLower(unescapeColumn(inner))]; ok {
		return "[" + escapeColumn(to) + "]"
	}
	return group
}

// renameItems renames the columns among the bracketed items of a list,
// leaving separators and specials (#Totals) as they are.
func renameItems(s string, renames map[string]string) string {
	var out strings.Builder
	for i := 0; i < len(s); {
		if s[i] != '[' {
			out.WriteByte(s[i])
			i++
			continue
		}
		end := closeBracket(s, i)
		if end < 0 {
			out.WriteString(s[i:])
			break
		}
		item := s[i:end]
		if name := item[1 : len(item)-1]; !strings.HasPrefix(name, "#") {
			if to, ok := renames[strings.ToLower(unescapeColumn(name))]; ok {
				item = "[" + escapeColumn(to) + "]"
			}
		}
		out.WriteString(item)
		i = end
	}
	return out.String()
}

// Column names in structured references escape [ ] # and ' with a '.
func escapeColumn(s string) string {
	return strings.NewReplacer("'", "''", "[", "'[", "]", "']", "#", "'#").Replace(s)
}

func unescapeColumn(s string) string {
	var b strings.Builder
	for i := 0; i < len(s); i++ {
		if s[i] == '\'' && i+1 < len(s) {
			i++
		}
		b.WriteByte(s[i])
	}
	return b.String()
}

// plainColumn reports whether a name can follow @ without brackets.
func plainColumn(s string) bool {
	for i := 0; i < len(s); i++ {
		if !identByte(s[i]) {
			return false
		}
	}
	return s != ""
}

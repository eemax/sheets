package cli

import (
	"regexp"
	"sort"
	"strings"

	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/xlsx"
)

// What check --against calls a feature lost. A feature of the old version
// is kept when the new one has one of its kind and its identity: a
// validation or conditional format with the same type and formula, a chart
// of the same type at the same place or on the same data, a shape or picture
// in the same place or by the same name, a name or table by its name. Never
// one of its kind alone: another validation, another chart on the sheet is
// no reason to say nothing of the one that is gone. Where a feature was is
// where the rows and columns added and removed since put it.

// shiftFunc gives a reference of the old version (a range, a cell, or a
// chart's series as Sheet!$B$2:$B$9; unqualified ones are on home, a sheet
// as now named) as the rows and columns added and removed since move it:
// "" when it was removed.
type shiftFunc func(home, ref string) string

// lostFeatures are what the old version carries that the new one lost. renamed
// maps each old sheet (lower-cased) to its new name ("" for a sheet removed),
// tables each Excel table renamed (old name lower-cased) to its new name.
func lostFeatures(before, after []xlsx.Feature, renamed, tables map[string]string, shift shiftFunc) []lostFeature {
	if shift == nil {
		shift = func(_, ref string) string { return ref }
	}
	used := make([]bool, len(after))
	// nowOf gives the sheet a feature of the old version is on, as now named.
	nowOf := func(f xlsx.Feature) (string, bool) {
		if f.Sheet == "" {
			return "", true
		}
		to, ok := renamed[strings.ToLower(f.Sheet)]
		if ok && to == "" {
			return "", false // its sheet was removed
		}
		if !ok {
			to = f.Sheet
		}
		return to, true
	}
	// moved gives the places a feature of the old version is now, by the
	// rows and columns added and removed ("" when they took it).
	moved := func(f xlsx.Feature, home string) string {
		var out []string
		for _, w := range strings.Fields(f.Where) {
			if s := shift(home, w); s != "" && !strings.Contains(s, "#REF!") {
				out = append(out, s)
			}
		}
		return placeKey(strings.Join(out, " "))
	}
	seriesMoved := func(f xlsx.Feature, home string) []string {
		out := make([]string, len(f.Series))
		for i, s := range f.Series {
			if m := shift(home, s); !strings.Contains(m, "#REF!") {
				out[i] = refKey(m)
			}
		}
		return out
	}
	// samePlace: the feature is where the lines added and removed put it (where
	// it was, when none were).
	samePlace := func(f, g xlsx.Feature, home string) bool {
		p := moved(f, home)
		return p != "" && p == placeKey(g.Where)
	}
	sameSeries := func(f, g xlsx.Feature, home string) bool {
		if len(f.Series) == 0 || len(f.Series) != len(g.Series) {
			return false
		}
		then := seriesMoved(f, home)
		for i, s := range g.Series {
			if refKey(f.Series[i]) != refKey(s) && then[i] != refKey(s) {
				return false
			}
		}
		return true
	}
	// pass 0 matches the same place, 1 the same feature anywhere on the sheet
	// (it moved or grew).
	match := func(f xlsx.Feature, pass int) bool {
		home, ok := nowOf(f)
		if !ok {
			return false
		}
		for i, g := range after {
			if used[i] || g.Kind != f.Kind || g.Chartsheet != f.Chartsheet {
				continue
			}
			here := strings.EqualFold(g.Sheet, home)
			switch {
			case f.Kind == xlsx.FeatureName:
				// By its name, in the scope it had.
				if pass > 0 || !here || !strings.EqualFold(f.Where, g.Where) {
					continue
				}
			case f.Kind == xlsx.FeatureTable:
				name := f.What
				if to, ok := tables[strings.ToLower(f.What)]; ok {
					name = to
				}
				if pass > 0 || !here || !strings.EqualFold(name, g.What) {
					continue
				}
			case f.Chartsheet:
				// A chart sheet is no worksheet, so a rename of it is not
				// followed: its chart is the one chart sheet left of its type.
				if typeKey(f.Type) != typeKey(g.Type) || pass == 0 && !here {
					continue
				}
			case !here:
				continue
			case f.Kind == xlsx.FeatureValidation || f.Kind == xlsx.FeatureCond:
				if typeKey(f.Type) != typeKey(g.Type) || ruleKey(f.Rule) != ruleKey(g.Rule) || pass == 0 && !samePlace(f, g, home) {
					continue
				}
			case f.Kind == xlsx.FeatureChart:
				if typeKey(f.Type) != typeKey(g.Type) {
					continue
				}
				if pass == 0 && !samePlace(f, g, home) || pass > 0 && !sameSeries(f, g, home) {
					continue
				}
			default: // a shape or a picture
				if pass == 0 && !samePlace(f, g, home) || pass > 0 && !strings.EqualFold(f.What, g.What) {
					continue
				}
			}
			used[i] = true
			return true
		}
		return false
	}
	kept := make([]bool, len(before))
	for pass := 0; pass < 2; pass++ {
		for i, f := range before {
			if !kept[i] {
				kept[i] = match(f, pass)
			}
		}
	}
	var out []lostFeature
	for i, f := range before {
		if kept[i] {
			continue
		}
		where, what := f.Where, f.What // a name of the workbook
		if f.Sheet != "" {
			where = grid.QuoteSheet(f.Sheet) + "!" + f.Where
		}
		if f.Chartsheet {
			where, what = grid.QuoteSheet(f.Sheet), strings.TrimSpace(what+" on a chart sheet")
		}
		out = append(out, lostFeature{Kind: f.Kind, Where: where, What: what})
	}
	if out == nil {
		out = []lostFeature{}
	}
	return out
}

// typeKey names a rule's type: a validation of no type allows any value.
func typeKey(t string) string {
	if t == "" {
		return "none"
	}
	return t
}

// refKey writes a reference to compare it: quotes, dollar signs and case
// aside.
func refKey(ref string) string {
	return strings.ToLower(strings.NewReplacer("$", "", "'", "").Replace(ref))
}

// placeKey writes where a feature is to compare it: its ranges, in order.
func placeKey(where string) string {
	parts := strings.Fields(refKey(where))
	sort.Strings(parts)
	return strings.Join(parts, " ")
}

// A reference at the start of some text (a cell, a range, whole columns or
// rows, with the sheet it names), and a structured reference's name.
var (
	refAt = regexp.MustCompile(`^(?:(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(?:\$?[A-Za-z]{1,3}\$?[0-9]+(?::\$?[A-Za-z]{1,3}\$?[0-9]+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[0-9]+:\$?[0-9]+)`)
	tblAt = regexp.MustCompile(`^[A-Za-z_][A-Za-z0-9_.]*\[`)
)

func nameChar(b byte) bool {
	return b >= 'a' && b <= 'z' || b >= 'A' && b <= 'Z' || b >= '0' && b <= '9' || b == '_' || b == '.' || b >= 0x80
}

// ruleKey writes what a validation or conditional format says to compare it:
// what its formulas test, with the cells they read left out (rows and columns
// added or removed rewrite those, a sheet renamed the sheet's name), their
// case and spaces too. A text in quotes stays as it is.
func ruleKey(rule string) string {
	var b strings.Builder
	for i := 0; i < len(rule); {
		c := rule[i]
		switch {
		case c == '"':
			j := i + 1
			for j < len(rule) {
				if rule[j] == '"' {
					if j+1 < len(rule) && rule[j+1] == '"' {
						j += 2
						continue
					}
					break
				}
				j++
			}
			b.WriteString(rule[i:min(j+1, len(rule))])
			i = j + 1
			continue
		case c == ' ' || c == '\t' || c == '\n':
			i++
			continue
		case i == 0 || !nameChar(rule[i-1]) && rule[i-1] != '!':
			if m := refAt.FindString(rule[i:]); m != "" && (i+len(m) == len(rule) || !nameChar(rule[i+len(m)]) && rule[i+len(m)] != '(' && rule[i+len(m)] != '[') {
				b.WriteString("{ref}")
				i += len(m)
				continue
			}
			if m := tblAt.FindString(rule[i:]); m != "" {
				// A structured reference: the table's name and what it picks.
				depth, j := 0, i+len(m)-1
				for ; j < len(rule); j++ {
					if rule[j] == '[' {
						depth++
					} else if rule[j] == ']' {
						if depth--; depth == 0 {
							break
						}
					}
				}
				b.WriteString("{table}")
				i = j + 1
				continue
			}
		}
		b.WriteString(strings.ToLower(rule[i : i+1]))
		i++
	}
	return b.String()
}

// tableRenames finds the Excel tables renamed: those of the old version
// that the new one has under no name, where the new version has a table of
// the same columns on the same sheet (as now named) under a name the old
// version had no table by, in the same range first. It maps each old name
// (lower-cased) to the new one.
func tableRenames(before, after []xlsx.Feature, renamed map[string]string) map[string]string {
	hadBefore, hasNow := map[string]bool{}, map[string]bool{}
	for _, f := range before {
		if f.Kind == xlsx.FeatureTable {
			hadBefore[strings.ToLower(f.What)] = true
		}
	}
	for _, g := range after {
		if g.Kind == xlsx.FeatureTable {
			hasNow[strings.ToLower(g.What)] = true
		}
	}
	out := map[string]string{}
	taken := map[string]bool{}
	for _, sameRange := range []bool{true, false} {
		for _, f := range before {
			if f.Kind != xlsx.FeatureTable || hasNow[strings.ToLower(f.What)] || out[strings.ToLower(f.What)] != "" {
				continue
			}
			home := f.Sheet
			if to, ok := renamed[strings.ToLower(f.Sheet)]; ok {
				home = to
			}
			for _, g := range after {
				key := strings.ToLower(g.What)
				if g.Kind != xlsx.FeatureTable || hadBefore[key] || taken[key] || !strings.EqualFold(g.Sheet, home) || g.Rule != f.Rule {
					continue
				}
				if sameRange && placeKey(g.Where) != placeKey(f.Where) {
					continue
				}
				out[strings.ToLower(f.What)] = g.What
				taken[key] = true
				break
			}
		}
	}
	return out
}

// renameTable rewrites a formula's structured references to the table old
// (any case) as to now.
func renameTable(f, old, now string) string {
	if !strings.Contains(strings.ToLower(f), strings.ToLower(old)+"[") {
		return f
	}
	var b strings.Builder
	in := false
	for i := 0; i < len(f); {
		c := f[i]
		if c == '"' {
			in = !in
		} else if !in && i+len(old) < len(f) && f[i+len(old)] == '[' && strings.EqualFold(f[i:i+len(old)], old) &&
			(i == 0 || !nameChar(f[i-1]) && f[i-1] != '!') {
			b.WriteString(now)
			i += len(old)
			continue
		}
		b.WriteByte(c)
		i++
	}
	return b.String()
}

// readsTables tells whether a sheet's formulas use one of the tables (by
// lower-cased name).
func readsTables(s *grid.Sheet, tables map[string]string) bool {
	found := false
	s.Each(s.Used(), func(_, _ int, c *grid.Cell) {
		if found || c.Formula == "" {
			return
		}
		lf := strings.ToLower(c.Formula)
		for name := range tables {
			if strings.Contains(lf, name+"[") {
				found = true
				return
			}
		}
	})
	return found
}

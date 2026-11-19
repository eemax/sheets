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
// where the rows and columns added and removed since put it. A validation
// or conditional format whose cells hold another of its kind now, its rule
// edited (a bound, a list's items, a threshold: the same type, and a
// formula's own of the same shape), is changed, not lost; another type, or a
// formula of another shape (the banding sheets fmt writes over a highlight),
// is the old rule lost.

// changedFeature is a validation or conditional format the new version
// holds on the same cells as one of the old version (where the rows and
// columns added and removed put them), of the same kind, whose rule is
// another: an edit of it (a dropdown's items, a bound, a threshold), not a
// loss.
type changedFeature struct {
	Kind   string `json:"kind"` // validation, conditional_format
	Where  string `json:"where"`
	What   string `json:"what,omitempty"` // as it is now
	Before string `json:"before"`         // its rule, as it was
	After  string `json:"after"`
}

// ruleText shows a rule: its type, operator and formulas; a validation of no
// type allows any value.
func ruleText(f xlsx.Feature) string {
	if f.Kind == xlsx.FeatureValidation && typeKey(f.Type) == "none" {
		return "any value"
	}
	var parts []string
	for _, p := range strings.Split(f.Rule, "|") {
		if p = strings.TrimSpace(p); p != "" {
			parts = append(parts, p)
		}
	}
	return strings.TrimSpace(typeKey(f.Type) + " " + strings.Join(parts, " "))
}

// shiftFunc gives a reference of the old version (a range, a cell, or a
// chart's series as Sheet!$B$2:$B$9; unqualified ones are on home, a sheet
// as now named) as the rows and columns added and removed since move it:
// "" when it was removed.
type shiftFunc func(home, ref string) string

// lostFeatures are what the old version carries that the new one lost (see
// compareFeatures).
func lostFeatures(before, after []xlsx.Feature, renamed, tables map[string]string, shift shiftFunc) []lostFeature {
	lost, _ := compareFeatures(before, after, renamed, tables, shift)
	return lost
}

// compareFeatures gives what the old version carries that the new one lost,
// and the validations and conditional formats changed. renamed maps each old
// sheet (lower-cased) to its new name ("" for a sheet removed), tables each
// Excel table renamed (old name lower-cased) to its new name.
func compareFeatures(before, after []xlsx.Feature, renamed, tables map[string]string, shift shiftFunc) ([]lostFeature, []changedFeature) {
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
	// movedTo gives the sheet (lower-cased) and places a feature of the old
	// version is on now, by the rows and columns added and removed, and the
	// cells cut and pasted, perhaps to another sheet ("" when they took it).
	movedTo := func(f xlsx.Feature, home string) (string, string) {
		var out []string
		sheet := strings.ToLower(home)
		for _, w := range strings.Fields(f.Where) {
			s := shift(home, w)
			if s == "" || strings.Contains(s, "#REF!") {
				continue
			}
			on := strings.ToLower(home)
			if k := strings.LastIndex(s, "!"); k >= 0 && !strings.Contains(w, "!") {
				on, s = strings.ToLower(strings.ReplaceAll(strings.Trim(s[:k], "'"), "''", "'")), s[k+1:]
			}
			if len(out) == 0 {
				sheet = on
			} else if on != sheet {
				return "", "" // split between sheets: not the feature it was
			}
			out = append(out, s)
		}
		return sheet, placeKey(strings.Join(out, " "))
	}
	moved := func(f xlsx.Feature, home string) string {
		_, p := movedTo(f, home)
		return p
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
	// A feature of the new version is found by a key of what it must have
	// in common with the old one (its kind, sheet, type, rule and, on the
	// first pass, place), so each old feature costs one lookup, not a scan
	// of the new ones: a sheet with tens of thousands of rules is the case.
	// pass 0 matches the same place, 1 the same feature anywhere on the
	// sheet (it moved or grew).
	type bucket struct {
		idx  []int
		head int
	}
	var index [2]map[string]*bucket
	cs := func(f xlsx.Feature) string {
		if f.Chartsheet {
			return "cs"
		}
		return "-"
	}
	// afterKey is the key of a feature of the new version.
	afterKey := func(g xlsx.Feature, pass int) string {
		sheet := strings.ToLower(g.Sheet)
		head := g.Kind + "\x00" + cs(g) + "\x00"
		switch {
		case g.Kind == xlsx.FeatureName:
			return head + sheet + "\x00" + strings.ToLower(g.Where)
		case g.Kind == xlsx.FeatureTable:
			return head + sheet + "\x00" + strings.ToLower(g.What)
		case g.Chartsheet:
			if pass == 0 {
				return head + sheet + "\x00" + typeKey(g.Type)
			}
			return head + typeKey(g.Type)
		case g.Kind == xlsx.FeatureValidation || g.Kind == xlsx.FeatureCond:
			k := head + sheet + "\x00" + typeKey(g.Type) + "\x00" + ruleKey(g.Rule)
			if pass == 0 {
				k += "\x00" + placeKey(g.Where)
			}
			return k
		case g.Kind == xlsx.FeatureChart:
			return head + sheet + "\x00" + typeKey(g.Type)
		}
		if pass == 0 {
			return head + sheet + "\x00" + placeKey(g.Where)
		}
		return head + sheet + "\x00" + strings.ToLower(g.What)
	}
	// beforeKey is the key an old feature looks for, and what else must hold
	// of the one it finds (nil: nothing).
	beforeKey := func(f xlsx.Feature, home string, pass int) (string, func(xlsx.Feature) bool, bool) {
		sheet := strings.ToLower(home)
		head := f.Kind + "\x00" + cs(f) + "\x00"
		switch {
		case f.Kind == xlsx.FeatureName:
			// By its name, in the scope it had.
			return head + sheet + "\x00" + strings.ToLower(f.Where), nil, pass == 0
		case f.Kind == xlsx.FeatureTable:
			name := f.What
			if to, ok := tables[strings.ToLower(f.What)]; ok {
				name = to
			}
			return head + sheet + "\x00" + strings.ToLower(name), nil, pass == 0
		case f.Chartsheet:
			// A chart sheet is no worksheet, so a rename of it is not
			// followed: its chart is the one chart sheet left of its type.
			if pass == 0 {
				return head + sheet + "\x00" + typeKey(f.Type), nil, true
			}
			return head + typeKey(f.Type), nil, true
		case f.Kind == xlsx.FeatureValidation || f.Kind == xlsx.FeatureCond:
			if pass == 0 {
				on, p := movedTo(f, home)
				return head + on + "\x00" + typeKey(f.Type) + "\x00" + ruleKey(f.Rule) + "\x00" + p, nil, p != ""
			}
			return head + sheet + "\x00" + typeKey(f.Type) + "\x00" + ruleKey(f.Rule), nil, true
		case f.Kind == xlsx.FeatureChart:
			return head + sheet + "\x00" + typeKey(f.Type), func(g xlsx.Feature) bool {
				if pass == 0 {
					return samePlace(f, g, home)
				}
				return sameSeries(f, g, home)
			}, true
		}
		if pass == 0 {
			p := moved(f, home)
			return head + sheet + "\x00" + p, nil, p != ""
		}
		return head + sheet + "\x00" + strings.ToLower(f.What), nil, true
	}
	match := func(f xlsx.Feature, pass int) bool {
		home, ok := nowOf(f)
		if !ok {
			return false
		}
		key, pred, ok := beforeKey(f, home, pass)
		if !ok {
			return false
		}
		if index[pass] == nil {
			index[pass] = map[string]*bucket{}
			for i, g := range after {
				k := afterKey(g, pass)
				bk := index[pass][k]
				if bk == nil {
					bk = &bucket{}
					index[pass][k] = bk
				}
				bk.idx = append(bk.idx, i)
			}
		}
		bk := index[pass][key]
		if bk == nil {
			return false
		}
		for bk.head < len(bk.idx) && used[bk.idx[bk.head]] {
			bk.head++
		}
		for _, i := range bk.idx[bk.head:] {
			if used[i] || pred != nil && !pred(after[i]) {
				continue
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
	// A rule not kept whose cells hold another of its kind now was edited:
	// changed, the other of its type first.
	var changed []changedFeature
	byPlace := map[string][]int{}
	for i, g := range after {
		if !used[i] && (g.Kind == xlsx.FeatureValidation || g.Kind == xlsx.FeatureCond) {
			k := g.Kind + "\x00" + strings.ToLower(g.Sheet) + "\x00" + placeKey(g.Where)
			byPlace[k] = append(byPlace[k], i)
		}
	}
	for _, sameType := range []bool{true, false} {
		for i, f := range before {
			if kept[i] || f.Kind != xlsx.FeatureValidation && f.Kind != xlsx.FeatureCond {
				continue
			}
			home, ok := nowOf(f)
			on, p := movedTo(f, home)
			if !ok || p == "" {
				continue
			}
			for _, j := range byPlace[f.Kind+"\x00"+on+"\x00"+p] {
				g := after[j]
				if used[j] || !ruleEdited(f, g, sameType) {
					continue
				}
				used[j], kept[i] = true, true
				changed = append(changed, changedFeature{Kind: f.Kind, Where: grid.QuoteSheet(g.Sheet) + "!" + g.Where, What: g.What, Before: ruleText(f), After: ruleText(g)})
				break
			}
		}
	}
	if changed == nil {
		changed = []changedFeature{}
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
	return out, changed
}

// ruleEdited tells whether a rule of the new version, on the cells of one of
// the old, is the old one edited: of its type, or, when sameType is not
// asked, of a kindred one (a whole-number validation made a decimal one, a
// date, a text length: all a bound on a value) or a validation of no type,
// which allows any value and stands in for any (the rule cleared, its message
// kept); and, for a rule that is a formula of its own (a conditional format
// by expression, a custom validation), of the same shape, as a threshold
// edited is and the banding that replaced a highlight is not. Another rule on
// the same cells (a list where a number was, data bars for a highlight) is
// the old one lost.
func ruleEdited(f, g xlsx.Feature, sameType bool) bool {
	tf, tg := typeKey(f.Type), typeKey(g.Type)
	if tf != tg {
		if sameType || f.Kind != xlsx.FeatureValidation {
			return false
		}
		return tf == "none" || tg == "none" || bounded[tf] && bounded[tg]
	}
	if tf == "expression" || tf == "custom" {
		return ruleShape(f.Rule) == ruleShape(g.Rule)
	}
	return true
}

// bounded are the validations that bound a value (operator and limits).
var bounded = map[string]bool{"whole": true, "decimal": true, "date": true, "time": true, "textLength": true}

// ruleShape writes a rule's formulas with what an edit changes left out: the
// numbers and texts in them and the comparison made, as well as what ruleKey
// leaves out (the cells read, case, spaces).
func ruleShape(rule string) string {
	k := ruleKey(rule)
	var b strings.Builder
	for i := 0; i < len(k); {
		c := k[i]
		switch {
		case c == '"':
			j := i + 1
			for j < len(k) {
				if k[j] == '"' {
					if j+1 < len(k) && k[j+1] == '"' {
						j += 2
						continue
					}
					break
				}
				j++
			}
			b.WriteString(`"s"`)
			i = j + 1
		case c >= '0' && c <= '9' && (i == 0 || !nameChar(k[i-1])):
			for i < len(k) && (k[i] >= '0' && k[i] <= '9' || k[i] == '.') {
				i++
			}
			b.WriteByte('#')
		case c == '<' || c == '>' || c == '=':
			for i < len(k) && (k[i] == '<' || k[i] == '>' || k[i] == '=') {
				i++
			}
			b.WriteByte('~')
		default:
			b.WriteByte(c)
			i++
		}
	}
	return b.String()
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

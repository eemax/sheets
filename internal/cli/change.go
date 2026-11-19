package cli

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"fmt"
	"sort"
	"strings"
	"sync"

	"github.com/eemax/sheets/internal/book"
	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/lint"
	"github.com/eemax/sheets/internal/xlsx"
)

// The short form of the change report, after every changing command: the
// cells it wrote, the computed values downstream that moved, the errors and
// unknown values it brought, the features it lost and how the lint
// findings changed. It is built from what the edit worked out anyway (the
// results its save recomputed, before and after: see xlsx.Effects), so it
// costs no second calculation; the features are read off the file before
// and after, and the lint is run on both when the workbook is small enough
// for that to be quick. It is at most a handful of lines, whatever the
// edit; sheets check --against gives the whole report.

// changeReport is the short form.
type changeReport struct {
	// ValuesOnly: a CSV file, which holds values; only cells written are
	// counted.
	ValuesOnly bool `json:"values_only,omitempty"`
	// Uncounted: an undo, which restores the file as it was and does not
	// compare the cells (the counts are left out, not given as 0); the
	// features lost and the lint are compared.
	Uncounted bool `json:"counts_omitted,omitempty"`
	// Cells is the number of cells the command wrote (set, cleared, pasted,
	// sorted into place).
	Cells int `json:"cells"`
	// Moved are the computed values the command did not write whose stored
	// results changed: what it moved downstream.
	Moved     int           `json:"moved"`
	MovedList []changeValue `json:"moved_values,omitempty"`
	// Errors are formulas showing an error they did not show before.
	Errors    int           `json:"new_errors"`
	ErrorList []changeValue `json:"errors,omitempty"`
	// Unknown are formulas that had a result before (or were written) and
	// have none sheets can stand by now, grouped by why.
	Unknown     int             `json:"new_unknown"`
	UnknownList []changeUnknown `json:"unknown,omitempty"`
	// FeaturesLost: charts, validations, conditional formats, shapes,
	// pictures, names and tables the file had and no longer has.
	FeaturesLost     int           `json:"features_lost"`
	FeaturesLostList []lostFeature `json:"features_lost_list,omitempty"`
	// FeaturesChanged: validations and conditional formats the command
	// replaced on the same cells with another of their kind (a dropdown's
	// items, a bound, a threshold): edited, not lost.
	FeaturesChanged     int              `json:"features_changed,omitempty"`
	FeaturesChangedList []changedFeature `json:"features_changed_list,omitempty"`
	// RulesNotCompared: the validations and conditional formats were too
	// many to compare after every edit (that many, at most, on the larger
	// side); the other features were.
	RulesNotCompared int         `json:"rules_not_compared,omitempty"`
	Lint             *changeLint `json:"lint,omitempty"`
}

// MarshalJSON leaves the counts out of an undo's report.
func (r changeReport) MarshalJSON() ([]byte, error) {
	if !r.Uncounted {
		type plain changeReport
		return marshalPlain(plain(r))
	}
	return marshalPlain(struct {
		Uncounted           bool             `json:"counts_omitted"`
		FeaturesLost        int              `json:"features_lost"`
		FeaturesLostList    []lostFeature    `json:"features_lost_list,omitempty"`
		FeaturesChanged     int              `json:"features_changed,omitempty"`
		FeaturesChangedList []changedFeature `json:"features_changed_list,omitempty"`
		RulesNotCompared    int              `json:"rules_not_compared,omitempty"`
		Lint                *changeLint      `json:"lint,omitempty"`
	}{true, r.FeaturesLost, r.FeaturesLostList, r.FeaturesChanged, r.FeaturesChangedList, r.RulesNotCompared, r.Lint})
}

type changeValue struct {
	Cell   string `json:"cell"`
	Before string `json:"before"` // "" when the formula was written (or had no result)
	After  string `json:"after"`
	// NoResultBefore: the formula stored no result before, so whether it
	// showed this error already is not known.
	NoResultBefore bool `json:"no_result_before,omitempty"`
}

type changeUnknown struct {
	Why   string   `json:"why"`
	Count int      `json:"count"`
	Cells []string `json:"cells"`
}

// changeLint is how the lint findings changed, counted by cells.
type changeLint struct {
	Compared bool   `json:"compared"`
	Why      string `json:"why,omitempty"` // when not compared
	Errors   int    `json:"errors"`        // after less before
	Warnings int    `json:"warnings"`
	Notes    int    `json:"notes"`
	// Added and Removed count the findings that rose and fell, by severity,
	// rule by rule (Errors and the others net them: a new finding of one
	// rule does not hide behind one that went away in another).
	Added   lintCounts `json:"added"`
	Removed lintCounts `json:"removed"`
	// Rules are the rules whose counts changed, those that rose first.
	Rules []changeRule `json:"rules,omitempty"`
}

type lintCounts struct {
	Errors   int `json:"errors"`
	Warnings int `json:"warnings"`
	Notes    int `json:"notes"`
}

func (c *lintCounts) add(sev lint.Severity, n int) {
	switch sev {
	case lint.Error:
		c.Errors += n
	case lint.Warning:
		c.Warnings += n
	default:
		c.Notes += n
	}
}

type changeRule struct {
	Rule     string        `json:"rule"`
	Severity lint.Severity `json:"severity"`
	Before   int           `json:"before"`
	After    int           `json:"after"`
	// Message is a finding of the rule now that was not there before.
	Message string `json:"message,omitempty"`
}

const (
	changeListed = 5 // values, errors, cells per reason, features
	changeWhys   = 3 // reasons for unknown values
	changeRules  = 3 // lint rules
	// lintBytes bounds the sheets' XML linted before and after an edit:
	// reading and linting a side takes about 34 ms per MB (both sides run
	// side by side, on an M4), which on a streaming edit of 10,000 rows
	// of four columns (1.5 MB) was 1.7 times the time without it. At
	// 512 KB (about 3,500 such rows) it is about 17 ms. Larger workbooks
	// say so, and sheets lint does it.
	lintBytes = 512 << 10
	// featureRules bounds the validations and conditional formats read
	// before and after an edit: reading a rule takes about 10 µs, which
	// on a sheet with one validation per cell (tens of thousands) was
	// 5 times the time of the edit. More rules than this on either side
	// are not compared, and the report says so; sheets check --against
	// compares them.
	featureRules = 1000
)

// cellPos is a cell where it stands now; sheet lower-cased.
type cellPos struct {
	sheet    string
	row, col int
}

type resultTrack struct {
	sheet   string // as spelled
	written bool
	carried bool // written, an error it showed where it came from
	blank   bool // written, over a formula that had no result
	reached bool // the change reached it (not only computed for lack of a result)
	before  grid.Cell
	was     bool
	after   grid.Cell
	now     bool
	why     string
}

// session follows a file through the editing sessions of one command (a
// batch's steps, the streaming editor and the full one it hands over to),
// keeping each cell's first result before and last result after.
type session struct {
	path       string
	before     []byte // the file before the command; nil when it is new
	after      []byte
	started    bool
	valuesOnly bool
	uncounted  bool               // an undo: nothing counted, only features and lint compared
	written    map[cellPos]string // sheet as spelled
	results    map[cellPos]*resultTrack
	ops        []xlsx.LineOp
	order      []string // sheet names, in order, for listing
}

func newSession(path string) *session {
	return &session{path: path, written: map[cellPos]string{}, results: map[cellPos]*resultTrack{}}
}

// batchSession follows a batch (sheets do, sheets script) from the file as
// it began.
func batchSession(path string, start []byte) *session {
	s := newSession(path)
	s.before, s.started = start, true
	return s
}

// ended closes a batch's session on the file it left: none when it saved
// nothing.
func (a *app) ended(b *xlsx.Batch) {
	if a.session == nil {
		return
	}
	if !b.Changed() {
		a.session = nil
		return
	}
	a.session.after = b.Data()
	a.changed = nil
}

// add takes in what an editor did, once it saved (or rendered, in a dry
// run).
func (s *session) add(ed book.Editor) {
	t, ok := ed.(book.Tracked)
	if !ok {
		return
	}
	if !s.started {
		s.before, s.started = t.Before(), true
	}
	s.after = t.After()
	if vo, ok := ed.(book.ValuesOnly); ok && vo.HoldsValues() {
		s.valuesOnly = true
	}
	if wb := ed.Workbook(); wb != nil {
		s.order = wb.SheetNames()
	}
	for _, ef := range t.Effects() {
		s.effects(ef)
	}
}

// effects takes in one part of a session: its moves of rows, columns and
// sheets carry what earlier parts found to where it is now.
func (s *session) effects(ef xlsx.Effects) {
	for _, op := range ef.Ops {
		s.ops = append(s.ops, op)
		written := map[cellPos]string{}
		for p, sheet := range s.written {
			if q, name, ok := opMove(op, p, sheet); ok {
				written[q] = name
			}
		}
		results := map[cellPos]*resultTrack{}
		for p, r := range s.results {
			if q, name, ok := opMove(op, p, r.sheet); ok {
				r.sheet = name
				results[q] = r
			}
		}
		s.written, s.results = written, results
	}
	mine := map[cellPos]bool{}
	for _, rc := range ef.Results {
		p := cellPos{strings.ToLower(rc.Sheet), rc.Row, rc.Col}
		mine[p] = true
		if r := s.results[p]; r != nil {
			r.after, r.now, r.why = rc.After, rc.Now, rc.Why
			if rc.Written {
				r.carried, r.blank = rc.Carried, rc.Blank
			}
			r.written = r.written || rc.Written
			r.reached = r.reached || rc.Reached
			continue
		}
		s.results[p] = &resultTrack{sheet: rc.Sheet, written: rc.Written, carried: rc.Carried, blank: rc.Blank, reached: rc.Reached, before: rc.Before, was: rc.Was, after: rc.After, now: rc.Now, why: rc.Why}
	}
	for _, w := range ef.Written {
		p := cellPos{strings.ToLower(w.Sheet), w.Row, w.Col}
		s.written[p] = w.Sheet
		if !mine[p] {
			// Written over with something else: what it computed before is
			// gone with it.
			delete(s.results, p)
		}
	}
}

// follows tells whether the session holds a result for the cell (s may be
// nil).
func (s *session) follows(sheet string, row, col int) bool {
	if s == nil {
		return false
	}
	_, ok := s.results[cellPos{strings.ToLower(sheet), row, col}]
	return ok
}

// opMove is where a change of lines puts a cell (false: it took it).
func opMove(op xlsx.LineOp, p cellPos, sheet string) (cellPos, string, bool) {
	if p.sheet != strings.ToLower(op.Sheet) {
		return p, sheet, true
	}
	switch {
	case op.Gone:
		return p, sheet, false
	case op.Rename != "":
		p.sheet = strings.ToLower(op.Rename)
		return p, op.Rename, true
	}
	at := &p.row
	if op.Cols {
		at = &p.col
	}
	n, ok := lineMove(op, *at)
	*at = n
	return p, sheet, ok
}

// lineMove is where a change of lines puts line n.
func lineMove(op xlsx.LineOp, n int) (int, bool) {
	switch {
	case op.N > 0 && n >= op.At:
		return n + op.N, true
	case op.N < 0 && n >= op.At-op.N:
		return n + op.N, true
	case op.N < 0 && n >= op.At:
		return 0, false
	}
	return n, true
}

// shiftRef gives a reference of the file before (a range or cell, perhaps
// with its sheet; unqualified ones are on home) where the session's lines
// moved it: "" when they took it.
func (s *session) shiftRef(home, ref string) string {
	sheet, body, qualified := home, ref, false
	if i := strings.LastIndex(ref, "!"); i >= 0 {
		sheet, body, qualified = strings.Trim(ref[:i], "'"), ref[i+1:], true
		sheet = strings.ReplaceAll(sheet, "''", "'")
	}
	rg, err := grid.ParseRange(strings.ReplaceAll(body, "$", ""))
	if err != nil {
		return ref
	}
	was, renamed := rg, false
	for _, op := range s.ops {
		if !strings.EqualFold(op.Sheet, sheet) {
			continue
		}
		switch {
		case op.Gone:
			return ""
		case op.Rename != "":
			sheet, renamed = op.Rename, true
			continue
		}
		lo, hi, top := &rg.R1, &rg.R2, grid.MaxRows
		if op.Cols {
			lo, hi, top = &rg.C1, &rg.C2, grid.MaxCols
		}
		if *lo == 1 && *hi == top {
			continue // whole columns (or rows) stay whole
		}
		a, okA := lineMove(op, *lo)
		if !okA {
			a = op.At // the first line left after the ones deleted
		}
		b, okB := lineMove(op, *hi)
		if !okB {
			b = op.At - 1
		}
		if b < a {
			return ""
		}
		*lo, *hi = a, b
	}
	if rg == was && !renamed {
		return ref
	}
	out := rg.String()
	if rg == was {
		out = strings.ReplaceAll(body, "$", "")
	} else if rg.R1 == rg.R2 && rg.C1 == rg.C2 && !strings.Contains(body, ":") {
		out = grid.CellName(rg.R1, rg.C1)
	}
	if qualified {
		return grid.QuoteSheet(sheet) + "!" + out
	}
	return out
}

// renamed maps each sheet of the file before (lower-cased) to its name now,
// "" for one deleted.
func (s *session) renamed() map[string]string {
	out := map[string]string{}
	now := map[string]string{} // name now (lower-cased) → first name, lower-cased
	for _, op := range s.ops {
		k := strings.ToLower(op.Sheet)
		first, ok := now[k]
		if !ok {
			first = k
		}
		switch {
		case op.Gone:
			out[first] = ""
			delete(now, k)
		case op.Rename != "":
			out[first] = op.Rename
			delete(now, k)
			now[strings.ToLower(op.Rename)] = first
		}
	}
	return out
}

// report works out the short form.
func (s *session) report() *changeReport {
	r := &changeReport{ValuesOnly: s.valuesOnly, Uncounted: s.uncounted, Cells: len(s.written)}
	if s.valuesOnly {
		return r
	}
	order := map[string]int{}
	for i, n := range s.order {
		order[strings.ToLower(n)] = i + 1
	}
	var keys []cellPos
	for p := range s.results {
		keys = append(keys, p)
	}
	sort.Slice(keys, func(i, j int) bool {
		a, b := keys[i], keys[j]
		if a.sheet != b.sheet {
			oa, ob := order[a.sheet], order[b.sheet]
			if oa == 0 || ob == 0 || oa == ob {
				return oa != 0 && ob == 0 || (oa == 0) == (ob == 0) && a.sheet < b.sheet
			}
			return oa < ob
		}
		if a.row != b.row {
			return a.row < b.row
		}
		return a.col < b.col
	})
	whys := map[string]*changeUnknown{}
	var whyOrder []string
	for _, p := range keys {
		t := s.results[p]
		_, w := s.written[p]
		w = w || t.written
		ref := grid.FormatRef(t.sheet, grid.Range{R1: p.row, C1: p.col, R2: p.row, C2: p.col})
		before := ""
		if !w && t.was {
			before = showCell(t.before)
		}
		cv := changeValue{Cell: ref, Before: before, After: showCell(t.after)}
		switch {
		case !t.now:
			if !w && !t.was || w && t.blank {
				// Nothing known before, or a formula that had no result
				// and still has none: nothing new.
				continue
			}
			r.Unknown++
			root, _ := rootReason(t.why)
			u := whys[root]
			if u == nil {
				u = &changeUnknown{Why: root}
				whys[root] = u
				whyOrder = append(whyOrder, root)
			}
			u.Count++
			if len(u.Cells) < changeListed {
				u.Cells = append(u.Cells, ref)
			}
			continue
		case !w && t.was && !xlsx.SameResult(t.before, t.after):
			r.Moved++
			if len(r.MovedList) < changeListed {
				r.MovedList = append(r.MovedList, cv)
			}
		}
		// An error is new unless the formula showed one before: where it
		// came from (a sort or move), or the result it had before a rewrite.
		// A formula the change reached that stored no result (openpyxl's
		// and XlsxWriter's) shows one now: nothing says it showed one
		// before, so it is counted, and told so.
		if t.after.Kind == grid.Error && !t.carried && (w || t.was || t.reached) && !(t.was && t.before.Kind == grid.Error) {
			r.Errors++
			if !w && !t.was {
				cv.NoResultBefore = true
			}
			if len(r.ErrorList) < changeListed {
				r.ErrorList = append(r.ErrorList, cv)
			}
		}
	}
	sort.SliceStable(whyOrder, func(i, j int) bool { return whys[whyOrder[i]].Count > whys[whyOrder[j]].Count })
	for _, w := range whyOrder {
		r.UnknownList = append(r.UnknownList, *whys[w])
	}
	s.compare(r)
	return r
}

// compare reads the features lost and the lint off the file before and
// after, side by side.
func (s *session) compare(r *changeReport) {
	if s.after == nil || !isXLSX(s.after) {
		return
	}
	var wg sync.WaitGroup
	var fa, fb []xlsx.Feature
	var na, nb int // rules past the bound, on each side
	var lintA, lintB []lint.Finding
	lintable := sheetBytes(s.before) <= lintBytes && sheetBytes(s.after) <= lintBytes
	run := func(fn func()) {
		wg.Add(1)
		go func() { defer wg.Done(); fn() }()
	}
	if s.before != nil && isXLSX(s.before) {
		run(func() { fa, na, _ = xlsx.FeaturesWithin(s.before, featureRules) })
		if lintable {
			run(func() {
				if wb, err := xlsx.ReadBytes(s.before); err == nil {
					lintA = lint.Check(wb, lint.Options{})
				}
			})
		}
	}
	run(func() { fb, nb, _ = xlsx.FeaturesWithin(s.after, featureRules) })
	failed := false
	if lintable {
		run(func() {
			wb, err := xlsx.ReadBytes(s.after)
			if err != nil {
				failed = true
				return
			}
			lintB = lint.Check(wb, lint.Options{})
		})
	}
	wg.Wait()
	if na > 0 || nb > 0 {
		// Rules are compared on both sides or on neither.
		fa, fb = withoutRules(fa), withoutRules(fb)
		r.RulesNotCompared = max(na, nb)
	}
	if len(fa) > 0 {
		lost, changed := compareFeatures(fa, fb, s.renamed(), nil, s.shiftRef)
		r.FeaturesLost = len(lost)
		r.FeaturesLostList = lost[:min(len(lost), changeListed)]
		if len(changed) > 0 {
			r.FeaturesChanged = len(changed)
			r.FeaturesChangedList = changed[:min(len(changed), changeListed)]
		}
	}
	switch {
	case !lintable:
		r.Lint = &changeLint{Why: "the workbook is too large to lint after every edit; sheets lint " + shellQuote(s.path) + " does"}
	case failed:
		r.Lint = &changeLint{Why: "the file saved could not be read back"}
	default:
		r.Lint = lintDelta(lintA, lintB)
	}
}

// withoutRules is the features less the validations and conditional formats.
func withoutRules(fs []xlsx.Feature) []xlsx.Feature {
	var out []xlsx.Feature
	for _, f := range fs {
		if f.Kind != xlsx.FeatureValidation && f.Kind != xlsx.FeatureCond {
			out = append(out, f)
		}
	}
	return out
}

// lintDelta compares two lint runs, rule by rule, by the cells found.
func lintDelta(before, after []lint.Finding) *changeLint {
	type count struct {
		sev           lint.Severity
		before, after int
		msg           string
	}
	cells := func(f lint.Finding) int { return max(1, len(f.Cells)) }
	seen := map[string]bool{}
	by := map[string]*count{}
	var rules []string
	at := func(f lint.Finding) *count {
		c := by[f.Rule]
		if c == nil {
			c = &count{sev: f.Severity}
			by[f.Rule] = c
			rules = append(rules, f.Rule)
		}
		return c
	}
	for _, f := range before {
		at(f).before += cells(f)
		seen[f.Rule+"\x00"+f.Sheet+"\x00"+f.Message] = true
	}
	for _, f := range after {
		c := at(f)
		c.after += cells(f)
		if c.msg == "" && !seen[f.Rule+"\x00"+f.Sheet+"\x00"+f.Message] {
			c.msg = f.Message
			if f.Sheet != "" && len(f.Cells) > 0 {
				c.msg = grid.QuoteSheet(f.Sheet) + "!" + f.Cells[0] + ": " + f.Message
			}
		}
	}
	out := &changeLint{Compared: true}
	for _, rule := range rules {
		c := by[rule]
		d := c.after - c.before
		if d == 0 {
			continue
		}
		(&out.Added).add(c.sev, max(d, 0))
		(&out.Removed).add(c.sev, max(-d, 0))
		switch c.sev {
		case lint.Error:
			out.Errors += d
		case lint.Warning:
			out.Warnings += d
		default:
			out.Notes += d
		}
		cr := changeRule{Rule: rule, Severity: c.sev, Before: c.before, After: c.after}
		if d > 0 {
			cr.Message = c.msg
		}
		out.Rules = append(out.Rules, cr)
	}
	sort.SliceStable(out.Rules, func(i, j int) bool {
		a, b := out.Rules[i], out.Rules[j]
		if (a.After > a.Before) != (b.After > b.Before) {
			return a.After > a.Before
		}
		return a.Severity > b.Severity
	})
	if len(out.Rules) > changeRules {
		out.Rules = out.Rules[:changeRules]
	}
	return out
}

// isXLSX tells a workbook package from a CSV file's bytes.
func isXLSX(b []byte) bool { return bytes.HasPrefix(b, []byte("PK")) }

// sheetBytes is the size of a package's worksheets, unpacked: what reading
// it for the lint costs.
func sheetBytes(b []byte) uint64 {
	if b == nil {
		return 0
	}
	zr, err := zip.NewReader(bytes.NewReader(b), int64(len(b)))
	if err != nil {
		return 0
	}
	var n uint64
	for _, f := range zr.File {
		if strings.HasPrefix(strings.ToLower(f.Name), "xl/worksheets/") {
			n += f.UncompressedSize64
		}
	}
	return n
}

// track takes in what an editor did, once it saved: a command's own
// session, or the batch's when the command is one of its steps.
func (a *app) track(path string, ed book.Editor) {
	if a.session == nil {
		a.session = newSession(path)
	}
	a.session.add(ed)
	a.changed = nil
}

// trackBytes takes in a change made without an editor (fmt, check --fix):
// the file before and after, and the results it stored.
func (a *app) trackBytes(path string, before, after []byte, results ...xlsx.ResultChange) {
	if a.session == nil {
		a.session = newSession(path)
	}
	s := a.session
	if !s.started {
		s.before, s.started = before, true
	}
	s.after = after
	s.valuesOnly = !isXLSX(after)
	s.effects(xlsx.Effects{Results: results})
	a.changed = nil
}

// printChange prints the short form: at most 13 lines (the head, what
// moved, the new errors, three reasons for unknowns and the rest counted,
// the features lost, those changed, the rules not compared, three lint
// rules).
func (a *app) printChange(r *changeReport) {
	if r == nil {
		return
	}
	head := "change: " + plural(r.Cells, "cell") + " written"
	if r.ValuesOnly && !r.Uncounted {
		a.printf("%s (a CSV file holds values only)\n", head)
		return
	}
	lost := plural(r.FeaturesLost, "feature") + " lost"
	if r.FeaturesChanged > 0 {
		lost += fmt.Sprintf(" (%d changed)", r.FeaturesChanged)
	}
	if r.Uncounted {
		head = "change: cells not compared (an undo restores the file as it was), " + lost
	} else {
		head += fmt.Sprintf(", %s moved, %s, %s, %s", plural(r.Moved, "value"), plural(r.Errors, "new error"), plural(r.Unknown, "new unknown"), lost)
	}
	if r.RulesNotCompared > 0 {
		head += " (rules not compared)"
	}
	if l := r.Lint; l != nil {
		head += ", lint " + lintHead(l)
	}
	a.printf("%s\n", head)
	if len(r.MovedList) > 0 {
		var parts []string
		for _, v := range r.MovedList {
			parts = append(parts, fmt.Sprintf("%s %s → %s", v.Cell, v.Before, v.After))
		}
		a.printf("moved: %s%s\n", strings.Join(parts, ", "), more(r.Moved, len(r.MovedList)))
	}
	if len(r.ErrorList) > 0 {
		var parts []string
		for _, v := range r.ErrorList {
			p := v.Cell + " " + v.After
			if v.Before != "" {
				p += " (was " + v.Before + ")"
			} else if v.NoResultBefore {
				p += " (no result stored before)"
			}
			parts = append(parts, p)
		}
		a.printf("new errors: %s%s\n", strings.Join(parts, ", "), more(r.Errors, len(r.ErrorList)))
	}
	for i, u := range r.UnknownList {
		if i == changeWhys {
			rest := 0
			for _, v := range r.UnknownList[i:] {
				rest += v.Count
			}
			a.printf("unknown: %d more for other reasons\n", rest)
			break
		}
		a.printf("unknown: %s%s: %s\n", strings.Join(u.Cells, ", "), more(u.Count, len(u.Cells)), u.Why)
	}
	if len(r.FeaturesLostList) > 0 {
		var parts []string
		for _, f := range r.FeaturesLostList {
			p := strings.ReplaceAll(f.Kind, "_", " ") + " " + f.Where
			if f.What != "" {
				p += " (" + f.What + ")"
			}
			parts = append(parts, p)
		}
		a.printf("lost: %s%s\n", strings.Join(parts, ", "), more(r.FeaturesLost, len(r.FeaturesLostList)))
	}
	if len(r.FeaturesChangedList) > 0 {
		var parts []string
		for _, f := range r.FeaturesChangedList {
			parts = append(parts, fmt.Sprintf("%s %s (%s → %s)", strings.ReplaceAll(f.Kind, "_", " "), f.Where, grid.Truncate(f.Before, 60), grid.Truncate(f.After, 60)))
		}
		a.printf("changed: %s%s\n", strings.Join(parts, ", "), more(r.FeaturesChanged, len(r.FeaturesChangedList)))
	}
	if r.RulesNotCompared > 0 {
		a.printf("lost: validations and conditional formats not compared (%s, more than %s: sheets check --against compares them)\n", plural(r.RulesNotCompared, "rule"), grid.FormatNumber(featureRules))
	}
	if l := r.Lint; l != nil {
		if !l.Compared {
			a.printf("lint: not compared: %s\n", l.Why)
		}
		for _, c := range l.Rules {
			line := fmt.Sprintf("lint: %s %d → %d", c.Rule, c.Before, c.After)
			if c.Message != "" {
				line += ": " + grid.Truncate(c.Message, 120)
			}
			a.printf("%s\n", line)
		}
	}
}

// lintHead says how the lint changed, in a few words.
func lintHead(l *changeLint) string {
	if !l.Compared {
		return "not compared"
	}
	var parts []string
	for _, side := range []struct {
		sign string
		c    lintCounts
	}{{"+", l.Added}, {"-", l.Removed}} {
		for _, x := range []struct {
			n    int
			word string
		}{{side.c.Errors, "error"}, {side.c.Warnings, "warning"}, {side.c.Notes, "note"}} {
			if x.n != 0 {
				parts = append(parts, side.sign+plural(x.n, x.word))
			}
		}
	}
	if len(parts) == 0 {
		return "as before"
	}
	return strings.Join(parts, " ")
}

func abs(n int) int {
	if n < 0 {
		return -n
	}
	return n
}

// more says how many were not listed.
func more(total, listed int) string {
	if total > listed {
		return fmt.Sprintf(" and %s more", grid.FormatNumber(float64(total-listed)))
	}
	return ""
}

// withChange adds the short form to a command's JSON object as "change".
func withChange(v any, r *changeReport) any {
	if r == nil {
		return v
	}
	b, err := marshalPlain(v)
	if err != nil || len(b) < 2 || b[0] != '{' {
		return v
	}
	c, err := marshalPlain(r)
	if err != nil {
		return v
	}
	var out []byte
	out = append(out, b[:len(b)-1]...)
	if len(b) > 2 {
		out = append(out, ',')
	}
	out = append(out, `"change":`...)
	out = append(out, c...)
	out = append(out, '}')
	return json.RawMessage(out)
}

// changeReport is the command's change report: nil when it changed nothing,
// or is a step of a batch.
func (a *app) changeReport() *changeReport {
	if a.batch || a.session == nil || !a.session.started {
		return nil
	}
	if a.changed == nil {
		a.changed = a.session.report()
	}
	return a.changed
}

// marshalPlain is JSON as printJSON writes it: <, > and & as they are.
func marshalPlain(v any) ([]byte, error) {
	var buf bytes.Buffer
	enc := json.NewEncoder(&buf)
	enc.SetEscapeHTML(false)
	if err := enc.Encode(v); err != nil {
		return nil, err
	}
	return bytes.TrimRight(buf.Bytes(), "\n"), nil
}

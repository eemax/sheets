package cli

import (
	"bytes"
	"fmt"
	"math"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"text/tabwriter"

	"github.com/eemax/sheets/internal/book"
	"github.com/eemax/sheets/internal/diff"
	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/xlsx"
)

// check --against: what changed from one version of a workbook to another,
// whoever wrote either. Inputs and formulas that changed; computed values
// that moved, each traced to its cause upstream (through what each formula
// reads); values that moved with no cause, flagged; features lost; values
// sheets cannot know; and check's own findings on the new version. Rows
// moved, as by a sort, are moved rows, not changes.

// againstResult is what check --against reports.
type againstResult struct {
	File    string `json:"file"`
	Against string `json:"against"`
	// ValuesOnly: a CSV file is one of the two, so only values are
	// compared (no formulas, causes or features).
	ValuesOnly   bool             `json:"values_only,omitempty"`
	Summary      againstSummary   `json:"summary"`
	Changed      []changedCell    `json:"changed"`
	Moved        []movedValue     `json:"moved"`
	Unexplained  []movedValue     `json:"unexplained"`
	FeaturesLost []lostFeature    `json:"features_lost"`
	Unknown      []unknownValue   `json:"unknown"`
	Findings     []checkFinding   `json:"findings"`
	MovedRows    []movedRows      `json:"moved_rows"`
	Structure    []structureEntry `json:"structure"`
	Errors       int              `json:"errors"`
	Warnings     int              `json:"warnings"`
	Notes        int              `json:"notes"`
}

type againstSummary struct {
	Inputs       int `json:"inputs_changed"`
	Formulas     int `json:"formulas_changed"`
	Flagged      int `json:"flagged"`
	Moved        int `json:"moved"`
	Unexplained  int `json:"unexplained"`
	FeaturesLost int `json:"features_lost"`
	Unknown      int `json:"unknown"`
	MovedRows    int `json:"moved_rows"`
}

// changedCell is an input or formula that changed.
type changedCell struct {
	Cell string `json:"cell"`          // Sheet!A1, in the new version
	Was  string `json:"was,omitempty"` // where it was, when elsewhere
	// Kind: input (a value typed in), formula (a formula changed, added
	// or removed), or value (a CSV file's cell).
	Kind   string `json:"kind"`
	Before string `json:"before"` // "" when empty; formulas start with =, text is quoted
	After  string `json:"after"`
	// Flag says why a change looks like a mistake: one row of a formula
	// column unlike the rows around it.
	Flag string `json:"flag,omitempty"`
}

// movedValue is a computed value that moved, with its cause: an input or
// formula change, rows or columns added, removed or moved, a name changed.
type movedValue struct {
	Cell   string `json:"cell"`
	Was    string `json:"was,omitempty"`
	Before string `json:"before"`
	After  string `json:"after"`
	Cause  string `json:"cause,omitempty"`
	// Note says where a value came from when sheets computed it (the file
	// stores none, or one marked as out of date).
	Note string `json:"note,omitempty"`
}

// unknownValue is a value sheets cannot compare: unknown on one side, or
// moved with a cause sheets cannot trace.
type unknownValue struct {
	Cell   string `json:"cell"`
	Before string `json:"before"`
	After  string `json:"after"`
	Why    string `json:"why"`
}

type lostFeature struct {
	Kind  string `json:"kind"` // chart, validation, conditional_format, shape, picture, name, table
	Where string `json:"where"`
	What  string `json:"what,omitempty"`
}

type movedRows struct {
	Sheet string `json:"sheet"`
	From  int    `json:"from"` // the first row's number in the old version
	To    int    `json:"to"`   // and in the new one
	Rows  int    `json:"rows"`
}

// structureEntry is a sheet added, removed or renamed, rows or columns
// added or removed, a name added or changed.
type structureEntry struct {
	Change string `json:"change"`
	Where  string `json:"where"`
}

// againstListed bounds the entries text output lists per section; --json
// lists them all.
const againstListed = 10

// cellAt names a cell: Sheet!A1.
func cellAt(sheet string, row, col int) string {
	return grid.QuoteSheet(sheet) + "!" + grid.CellName(row, col)
}

// valueOf is a cell's value as a reader sees it, and whether sheets knows
// it.
type valueOf struct {
	text  string
	cell  grid.Cell
	known bool
}

// value reads a cell's value, computing a formula's when the file holds
// none.
func value(wb *grid.Workbook, s *grid.Sheet, row, col int) valueOf {
	c := wb.Value(s, row, col)
	if c == nil {
		return valueOf{known: true}
	}
	if c.Pending {
		return valueOf{text: "?"}
	}
	return valueOf{text: valueText(wb, c), cell: *c, known: true}
}

func valueText(wb *grid.Workbook, c *grid.Cell) string {
	if c.Kind == grid.Empty {
		return ""
	}
	v := grid.Display(c, wb.FormatOf(c), wb.Date1904)
	if c.Kind == grid.Text {
		return strconv.Quote(v)
	}
	return v
}

// sameValues tells whether two known values are the same, numbers to
// within rounding.
func sameValues(a, b valueOf) bool {
	x, y := a.cell, b.cell
	if x.Kind != y.Kind {
		return false
	}
	switch x.Kind {
	case grid.Number:
		return math.Abs(x.Num-y.Num) <= 1e-9*math.Max(1, math.Max(math.Abs(x.Num), math.Abs(y.Num)))
	case grid.Text, grid.Error:
		return x.Text == y.Text
	case grid.Bool:
		return x.Bool == y.Bool
	}
	return true
}

// arrays finds the array formulas whose results hold cells, sheet by
// sheet.
type arrays map[*grid.Sheet][]grid.Range

// anchor gives the cell of the array formula whose result holds a cell,
// when that is another cell.
func (x arrays) anchor(s *grid.Sheet, row, col int) (int, int, bool) {
	rgs, ok := x[s]
	if !ok {
		for _, rg := range append(append([]grid.Range{}, s.Spills...), s.Arrays...) {
			if rg.Rows()*rg.Cols() > 1 {
				rgs = append(rgs, rg)
			}
		}
		x[s] = rgs
	}
	for _, rg := range rgs {
		if rg.Contains(row, col) && (rg.R1 != row || rg.C1 != col) {
			return rg.R1, rg.C1, true
		}
	}
	return 0, 0, false
}

// sideOf is one version: its file, workbook as fixed (results sheets can
// compute stored), and why results are unknown.
type sideOf struct {
	path   string
	data   []byte
	wb     *grid.Workbook
	why    map[string]string // by cellAt, lower-cased
	stored map[string]string // results sheets computed: by cellAt, lower-cased, a note
	csv    bool
}

func (x *sideOf) unknownWhy(sheet string, row, col int) string {
	if w := x.why[strings.ToLower(cellAt(sheet, row, col))]; w != "" {
		return w
	}
	return "the sheets calculator cannot compute it"
}

// computed tells whether a cell holds a computed value: a formula, or
// part of an array formula's result.
func (x arrays) computed(s *grid.Sheet, row, col int) bool {
	if c := s.At(row, col); c != nil && c.Formula != "" {
		return true
	}
	_, _, ok := x.anchor(s, row, col)
	return ok
}

// readSide reads one version for --against.
func (a *app) readSide(path string, cr *xlsx.CheckResult) (*sideOf, error) {
	x := &sideOf{path: path, why: map[string]string{}, stored: map[string]string{}}
	kind, err := book.FormatOf(path)
	if err != nil {
		return nil, err
	}
	if kind == "csv" {
		x.csv = true
		x.wb, err = open(path)
		return x, err
	}
	if x.data, err = xlsx.ReadFile(path); err != nil {
		return nil, err
	}
	if cr == nil {
		opts := checkTime
		opts.NoCompare = true
		if cr, err = xlsx.CheckResults(path, x.data, opts); err != nil {
			return nil, err
		}
	}
	x.wb = cr.After
	for _, u := range cr.Unknown {
		x.why[strings.ToLower(cellAt(u.Sheet, u.Row, u.Col))] = u.Why
	}
	if cr.FixError != "" {
		x.wb = cr.Before // values computed as read
	}
	for _, c := range cr.Changes {
		switch c.Kind {
		case "stored":
			x.stored[strings.ToLower(cellAt(c.Sheet, c.Row, c.Col))] = "the file stores no result"
		case "replaced":
			x.stored[strings.ToLower(cellAt(c.Sheet, c.Row, c.Col))] = "the file's result, marked as out of date, is " + showCell(c.Before)
		}
	}
	return x, nil
}

// against reports what changed from before to path. res is check's report
// on path, cr its results (nil for a CSV file).
func (a *app) against(path, before string, res *checkResult, cr *xlsx.CheckResult) (*againstResult, error) {
	out := &againstResult{File: path, Against: before, Findings: res.Findings,
		Changed: []changedCell{}, Moved: []movedValue{}, Unexplained: []movedValue{}, FeaturesLost: []lostFeature{},
		Unknown: []unknownValue{}, MovedRows: []movedRows{}, Structure: []structureEntry{}}
	B, err := a.readSide(path, cr)
	if err != nil {
		return nil, err
	}
	A, err := a.readSide(before, nil)
	if err != nil {
		return nil, err
	}
	out.ValuesOnly = A.csv || B.csv
	r := &report{out: out, A: A, B: B, arrays: arrays{}}
	r.compare()
	if !out.ValuesOnly {
		fa, err := xlsx.Features(A.data)
		if err != nil {
			return nil, fmt.Errorf("%s: %w", before, err)
		}
		fb, err := xlsx.Features(B.data)
		if err != nil {
			return nil, fmt.Errorf("%s: %w", path, err)
		}
		out.FeaturesLost = lostFeatures(fa, fb, r.renamed)
	}
	s := &out.Summary
	for _, c := range out.Changed {
		if c.Kind == "formula" {
			s.Formulas++
		} else {
			s.Inputs++
		}
		if c.Flag != "" {
			s.Flagged++
		}
	}
	s.Moved, s.Unexplained, s.FeaturesLost, s.Unknown = len(out.Moved), len(out.Unexplained), len(out.FeaturesLost), len(out.Unknown)
	for _, m := range out.MovedRows {
		s.MovedRows += m.Rows
	}
	return out, nil
}

// report works out an againstResult.
type report struct {
	out  *againstResult
	A, B *sideOf
	// renamed maps each old sheet (lower-cased) to its new name; a sheet
	// removed maps to "".
	renamed map[string]string

	sources []xlsx.Source
	causes  []string
	// lineSources are rows, columns and sheets added or moved, which come
	// after the cells changed.
	lineSources []xlsx.Source
	lineCauses  []string
	// candidates are computed values that moved, to trace.
	candidates []candidate
	// unknownSrc are formulas whose values are unknown, and why.
	unknownSrc []xlsx.Source
	unknownWhy []string
	// formulaSrc are formulas that are causes themselves (rows removed
	// across their references, a name they use changed).
	formulaSrc map[xlsx.At]string
	// edits are the lines added and removed on each sheet (new names), and
	// renames the sheets renamed (old name lower-cased, new name).
	edits   []lineEdit
	renames [][2]string
	arrays  arrays
	// replayed counts the runs of lines replayed on formulas, against
	// replayLimit.
	replayed int
	// lines pairs each old sheet's rows and columns with the new version's
	// (by the old name, lower-cased); cur is the sheet being compared.
	lines map[string]*lineMap
	cur   *lineMap
}

// lineMap gives the new number of each row and column kept, and its sheet's
// new name.
type lineMap struct {
	sheet    string
	row, col map[int]int
	// moved are the old rows a sort moved: a formula in one moved with it,
	// its references rewritten as for a copy.
	moved map[int]bool
}

// removal is rows or columns removed from a sheet, as runs numbered as
// before, each with its label.
type removal struct {
	sheet  string // as the old version names it
	cols   bool
	runs   [][2]int
	labels []string
}

// lineEdit is the rows and columns added (numbered as now) and removed
// (numbered as before) on a sheet, as runs.
type lineEdit struct {
	sheet                  string
	rowsRemoved, rowsAdded [][2]int
	colsRemoved, colsAdded [][2]int
	// movedOld and movedNew span the rows a sort moved (as numbered before
	// and now; zero when none moved), with the rows removed or added joining
	// them. Where rows were added or removed within them cannot be told: the
	// sort took them elsewhere.
	movedOld, movedNew [2]int
}

type candidate struct {
	at     xlsx.At
	before valueOf
	after  valueOf
	was    string
	note   string
	// via is where an array formula now is whose result held the cell
	// before and no longer does (a spill that shrank): what moved it moved
	// the cell.
	via *xlsx.At
}

// cause adds a change whose cells others follow: a cell's own (an input
// or formula), which is the nearer cause, or rows and columns added,
// removed or moved (structural).
func (r *report) cause(sheet string, rg grid.Range, label string) {
	r.sources = append(r.sources, xlsx.Source{Sheet: sheet, Range: rg})
	r.causes = append(r.causes, label)
}

func (r *report) structural(sheet string, rg grid.Range, label string) {
	r.lineSources = append(r.lineSources, xlsx.Source{Sheet: sheet, Range: rg, Readers: true})
	r.lineCauses = append(r.lineCauses, label)
}

// wasRows says where moved rows were, for a cause: nothing when they kept
// their numbers (rows added or removed around them moved the rest).
func wasRows(sheet string, m diff.Moved) string {
	if m.From == m.To {
		return ""
	}
	return " (were " + lineRange(sheet, m.From, m.From+m.Rows-1, false) + ")"
}

// sameNumber notes rows that moved yet kept their numbers: the rows around
// them moved, or rows were added or removed above them.
func sameNumber(m movedRows) string {
	if m.From == m.To {
		return " (its number held; the rows around it moved)"
	}
	return ""
}

func lineRange(sheet string, from, to int, cols bool) string {
	q := grid.QuoteSheet(sheet) + "!"
	if cols {
		if from == to {
			return q + grid.ColName(from)
		}
		return q + grid.ColName(from) + ":" + grid.ColName(to)
	}
	if from == to {
		return q + strconv.Itoa(from)
	}
	return q + strconv.Itoa(from) + ":" + strconv.Itoa(to)
}

// pairReplaced pairs, in order, the rows removed and added in the same
// place: a row whose every value changed is changed, not replaced.
func pairReplaced(rows []diff.Pair) []diff.Pair {
	var out []diff.Pair
	for i := 0; i < len(rows); {
		if rows[i].Old != 0 && rows[i].New != 0 {
			out = append(out, rows[i])
			i++
			continue
		}
		var olds, news []int
		for ; i < len(rows) && (rows[i].Old == 0 || rows[i].New == 0); i++ {
			if rows[i].New == 0 {
				olds = append(olds, rows[i].Old)
			} else {
				news = append(news, rows[i].New)
			}
		}
		n := min(len(olds), len(news))
		for k := 0; k < n; k++ {
			out = append(out, diff.Pair{Old: olds[k], New: news[k]})
		}
		for _, o := range olds[n:] {
			out = append(out, diff.Pair{Old: o})
		}
		for _, w := range news[n:] {
			out = append(out, diff.Pair{New: w})
		}
	}
	return out
}

// alikeLimit bounds the rows pairAlike compares, each with each.
const alikeLimit = 500

// pairAlike pairs the rows removed with the rows added that are mostly
// alike (at least half their cells the same, formulas by their pattern):
// rows a sort moved and an edit changed. It gives the rows, with those
// pairs.
func (r *report) pairAlike(sa, sb *grid.Sheet, rows []diff.Pair, cols [][2]int) []diff.Pair {
	var olds, news []int
	for _, q := range rows {
		switch {
		case q.New == 0:
			olds = append(olds, q.Old)
		case q.Old == 0:
			news = append(news, q.New)
		}
	}
	if len(olds) == 0 || len(news) == 0 || len(olds)*len(news) > alikeLimit*alikeLimit {
		return rows
	}
	key := func(wb *grid.Workbook, s *grid.Sheet, row, col int) string {
		c := s.At(row, col)
		if c == nil || c.IsEmpty() {
			return ""
		}
		if c.Formula != "" {
			k, _ := formula.Pattern(folded(c.Formula), row, col)
			return "=" + k
		}
		return value(wb, s, row, col).text
	}
	keys := func(wb *grid.Workbook, s *grid.Sheet, rows []int, pick int) map[int][]string {
		out := make(map[int][]string, len(rows))
		for _, row := range rows {
			ks := make([]string, len(cols))
			for i, c := range cols {
				ks[i] = key(wb, s, row, c[pick])
			}
			out[row] = ks
		}
		return out
	}
	ka, kb := keys(r.A.wb, sa, olds, 0), keys(r.B.wb, sb, news, 1)
	taken := map[int]bool{}
	paired := map[int]int{} // new row → old row
	for _, nr := range news {
		best, score := 0, 0
		for _, or := range olds {
			if taken[or] {
				continue
			}
			same, used := 0, 0
			for i := range cols {
				x, y := ka[or][i], kb[nr][i]
				if x == "" && y == "" {
					continue
				}
				used++
				if x == y {
					same++
				}
			}
			if used > 0 && same*2 >= used && same > score {
				best, score = or, same
			}
		}
		if best > 0 {
			taken[best], paired[nr] = true, best
		}
	}
	if len(paired) == 0 {
		return rows
	}
	var out []diff.Pair
	for _, q := range rows {
		switch {
		case q.New == 0 && taken[q.Old]:
			continue
		case q.Old == 0 && paired[q.New] > 0:
			out = append(out, diff.Pair{Old: paired[q.New], New: q.New})
			continue
		}
		out = append(out, q)
	}
	return out
}

// boolCalls finds TRUE() and FALSE(), which LibreOffice writes for every
// TRUE and FALSE typed in a formula (Excel reads them as the same).
var boolCalls = regexp.MustCompile(`(^|[^A-Z0-9_.])(TRUE|FALSE)\(\)`)

// folded is a formula as Excel compares it: letters outside text in double
// quotes in one case (function names, references, sheet names), and TRUE()
// and FALSE() as TRUE and FALSE.
func folded(f string) string {
	var b, seg strings.Builder // seg is a stretch outside quotes
	flush := func() {
		b.WriteString(boolCalls.ReplaceAllString(seg.String(), "${1}${2}"))
		seg.Reset()
	}
	in := false
	for _, r := range f {
		switch {
		case r == '"':
			if !in {
				flush()
			}
			in = !in
			b.WriteRune(r)
		case in:
			b.WriteRune(r)
		default:
			seg.WriteRune([]rune(strings.ToUpper(string(r)))[0])
		}
	}
	flush()
	return b.String()
}

// spillOnly tells whether a row (of a sheet that uses u) holds, apart from
// empty cells, nothing but the cells of array formulas' results (not the
// formulas themselves): a row a result grew into or shrank out of.
func (r *report) spillOnly(s *grid.Sheet, u grid.Range, row int) bool {
	n := 0
	for c := u.C1; c <= u.C2; c++ {
		cell := s.At(row, c)
		if cell == nil || cell.IsEmpty() {
			continue
		}
		if cell.Formula != "" {
			return false
		}
		if _, _, ok := r.arrays.anchor(s, row, c); !ok {
			return false
		}
		n++
	}
	return n > 0
}

// rowEmpty tells whether a row of a sheet that uses u holds nothing.
func rowEmpty(s *grid.Sheet, u grid.Range, row int) bool {
	if row < u.R1 || row > u.R2 {
		return true
	}
	for c := u.C1; c <= u.C2; c++ {
		if cell := s.At(row, c); cell != nil && !cell.IsEmpty() {
			return false
		}
	}
	return true
}

// pairSpilled pairs, in place, the rows added or removed that hold only an
// array formula's results where the other version has nothing: a result
// that grew or shrank at the bottom of a sheet, not rows added or removed.
func (r *report) pairSpilled(sa, sb *grid.Sheet, rows []diff.Pair) []diff.Pair {
	if len(sa.Spills)+len(sa.Arrays)+len(sb.Spills)+len(sb.Arrays) == 0 {
		return rows // nothing spills: Used scans a sheet, so not for each row
	}
	ua, ub := sa.Used(), sb.Used()
	oldTaken, newTaken := map[int]bool{}, map[int]bool{}
	for _, q := range rows {
		if q.Old > 0 {
			oldTaken[q.Old] = true
		}
		if q.New > 0 {
			newTaken[q.New] = true
		}
	}
	out := make([]diff.Pair, 0, len(rows))
	for _, q := range rows {
		switch {
		case q.New == 0 && !newTaken[q.Old] && rowEmpty(sb, ub, q.Old) && r.spillOnly(sa, ua, q.Old):
			newTaken[q.Old] = true
			q.New = q.Old
		case q.Old == 0 && !oldTaken[q.New] && rowEmpty(sa, ua, q.New) && r.spillOnly(sb, ub, q.New):
			oldTaken[q.New] = true
			q.Old = q.New
		}
		out = append(out, q)
	}
	return out
}

// runs joins line numbers into runs.
func runs(lines []int) [][2]int {
	sort.Ints(lines)
	var out [][2]int
	for _, n := range lines {
		if k := len(out) - 1; k >= 0 && out[k][1] == n-1 {
			out[k][1] = n
			continue
		}
		out = append(out, [2]int{n, n})
	}
	return out
}

func (r *report) compare() {
	A, B, out := r.A, r.B, r.out
	if out.ValuesOnly && len(A.wb.Sheets) > 0 && len(B.wb.Sheets) > 0 {
		// A CSV file's one sheet is the workbook's first: it is named
		// after the file.
		switch {
		case A.csv && !B.csv:
			A.wb.Sheets[0].Name = B.wb.Sheets[0].Name
		case B.csv && !A.csv:
			B.wb.Sheets[0].Name = A.wb.Sheets[0].Name
		}
	}
	al := diff.Align(A.wb, B.wb, diff.Options{Moves: true, LostRefs: true})
	r.renamed = map[string]string{}
	r.formulaSrc = map[xlsx.At]string{}
	var tr *xlsx.Tracer
	if !out.ValuesOnly {
		tr = xlsx.NewTracer(B.wb)
	}
	var removals []removal
	type job struct {
		sa, sb *grid.Sheet
		rows   []diff.Pair
		cols   [][2]int
		moved  []diff.Moved
	}
	var jobs []job
	for _, p := range al.Sheets {
		switch {
		case p.Old == nil:
			out.Structure = append(out.Structure, structureEntry{"sheet added", grid.QuoteSheet(p.New.Name)})
			if u := p.New.Used(); !u.Empty() {
				r.structural(p.New.Name, u, "sheet added: "+grid.QuoteSheet(p.New.Name))
			}
			continue
		case p.New == nil:
			out.Structure = append(out.Structure, structureEntry{"sheet removed", grid.QuoteSheet(p.Old.Name)})
			r.renamed[strings.ToLower(p.Old.Name)] = ""
			continue
		}
		r.renamed[strings.ToLower(p.Old.Name)] = p.New.Name
		if p.Old.Name != p.New.Name && !out.ValuesOnly {
			out.Structure = append(out.Structure, structureEntry{"sheet renamed", grid.QuoteSheet(p.Old.Name) + " → " + grid.QuoteSheet(p.New.Name)})
		}
		sa, sb := p.Old, p.New
		p.Rows = pairReplaced(p.Rows)
		var rowsAdded, rowsRemoved, colsAdded, colsRemoved []int
		var cols [][2]int
		for _, c := range p.Cols {
			switch {
			case c.Old == 0:
				colsAdded = append(colsAdded, c.New)
			case c.New == 0:
				colsRemoved = append(colsRemoved, c.Old)
			default:
				cols = append(cols, [2]int{c.Old, c.New})
			}
		}
		if len(p.Moved) > 0 {
			// Rows sorted and changed too: alike, they moved. Of all the rows
			// paired now, those whose rank among them changed are the rows
			// moved.
			p.Rows = r.pairAlike(sa, sb, p.Rows, cols)
			ua, ub := sa.Used(), sb.Used()
			p.Moved = diff.MovedRows(p.Rows, func(q diff.Pair) bool { return rowEmpty(sa, ua, q.Old) && rowEmpty(sb, ub, q.New) })
		}
		p.Rows = r.pairSpilled(sa, sb, p.Rows)
		removalOf := func(removed []int, byCol bool) removal {
			g := removal{sheet: sa.Name, cols: byCol, runs: runs(removed)}
			for _, run := range g.runs {
				g.labels = append(g.labels, lineRange(sa.Name, run[0], run[1], byCol)+" in "+A.path)
			}
			return g
		}
		for _, q := range p.Rows {
			switch {
			case q.Old == 0:
				rowsAdded = append(rowsAdded, q.New)
			case q.New == 0:
				rowsRemoved = append(rowsRemoved, q.Old)
			}
		}
		for _, run := range runs(rowsAdded) {
			w := lineRange(sb.Name, run[0], run[1], false)
			out.Structure = append(out.Structure, structureEntry{map[bool]string{true: "row added", false: "rows added"}[run[0] == run[1]], w})
			r.structural(sb.Name, grid.Range{R1: run[0], C1: 1, R2: run[1], C2: grid.MaxCols}, "rows added: "+w)
		}
		for _, run := range runs(rowsRemoved) {
			out.Structure = append(out.Structure, structureEntry{map[bool]string{true: "row removed", false: "rows removed"}[run[0] == run[1]], lineRange(sa.Name, run[0], run[1], false) + " in " + A.path})
		}
		for _, run := range runs(colsAdded) {
			w := lineRange(sb.Name, run[0], run[1], true)
			out.Structure = append(out.Structure, structureEntry{map[bool]string{true: "column added", false: "columns added"}[run[0] == run[1]], w})
			r.structural(sb.Name, grid.Range{R1: 1, C1: run[0], R2: grid.MaxRows, C2: run[1]}, "columns added: "+w)
		}
		for _, run := range runs(colsRemoved) {
			out.Structure = append(out.Structure, structureEntry{map[bool]string{true: "column removed", false: "columns removed"}[run[0] == run[1]], lineRange(sa.Name, run[0], run[1], true) + " in " + A.path})
		}
		if len(rowsRemoved) > 0 {
			removals = append(removals, removalOf(rowsRemoved, false))
		}
		if len(colsRemoved) > 0 {
			removals = append(removals, removalOf(colsRemoved, true))
		}
		lo, hi := 0, 0
		var spanOld, spanNew [2]int
		for _, m := range p.Moved {
			spanOld, spanNew = widen(spanOld, m.From, m.From+m.Rows-1), widen(spanNew, m.To, m.To+m.Rows-1)
			out.MovedRows = append(out.MovedRows, movedRows{Sheet: sb.Name, From: m.From, To: m.To, Rows: m.Rows})
			r.structural(sb.Name, grid.Range{R1: m.To, C1: 1, R2: m.To + m.Rows - 1, C2: grid.MaxCols},
				"rows moved: "+lineRange(sb.Name, m.To, m.To+m.Rows-1, false)+wasRows(sa.Name, m))
			if lo == 0 || min(m.From, m.To) < lo {
				lo = min(m.From, m.To)
			}
			hi = max(hi, m.From+m.Rows-1, m.To+m.Rows-1)
		}
		if lo > 0 {
			// What a sort spans: every cell there may hold another row's now.
			r.structural(sb.Name, grid.Range{R1: lo, C1: 1, R2: hi, C2: grid.MaxCols}, "rows moved: "+lineRange(sb.Name, lo, hi, false))
		}
		jobs = append(jobs, job{sa, sb, p.Rows, cols, p.Moved})
		spanOld, spanNew = reach(spanOld, rowsRemoved), reach(spanNew, rowsAdded)
		r.edits = append(r.edits, lineEdit{sheet: sb.Name, rowsRemoved: runs(rowsRemoved), rowsAdded: runs(rowsAdded), colsRemoved: runs(colsRemoved), colsAdded: runs(colsAdded), movedOld: spanOld, movedNew: spanNew})
	}
	for from, to := range r.renamed {
		if to != "" && !strings.EqualFold(from, to) {
			r.renames = append(r.renames, [2]string{from, to})
		}
	}
	sort.Slice(r.renames, func(i, j int) bool { return r.renames[i][0] < r.renames[j][0] })
	r.lines = map[string]*lineMap{}
	for _, j := range jobs {
		m := &lineMap{sheet: j.sb.Name, row: map[int]int{}, col: map[int]int{}, moved: map[int]bool{}}
		for _, mv := range j.moved {
			for k := 0; k < mv.Rows; k++ {
				m.moved[mv.From+k] = true
			}
		}
		for _, q := range j.rows {
			if q.Old > 0 && q.New > 0 {
				m.row[q.Old] = q.New
			}
		}
		for _, c := range j.cols {
			m.col[c[0]] = c[1]
		}
		r.lines[strings.ToLower(j.sa.Name)] = m
	}
	for _, j := range jobs {
		r.cur = r.lines[strings.ToLower(j.sa.Name)]
		for _, q := range j.rows {
			if q.Old == 0 || q.New == 0 {
				continue
			}
			for _, c := range j.cols {
				r.cell(j.sa, q.Old, c[0], j.sb, q.New, c[1])
			}
		}
	}
	if out.ValuesOnly {
		return
	}
	// Names added or changed: the formulas using them follow.
	old := map[string]grid.DefinedName{}
	for _, n := range A.wb.Names {
		old[strings.ToLower(r.renamed[strings.ToLower(n.Scope)]+"!"+n.Name)] = n
	}
	for _, n := range B.wb.Names {
		o, ok := old[strings.ToLower(n.Scope+"!"+n.Name)]
		var label string
		switch {
		case !ok:
			label = "name " + n.Name + " added"
			out.Structure = append(out.Structure, structureEntry{"name added", n.Name + " = " + n.Ref})
		case !r.sameRef(o.Ref, n.Ref):
			label = "name " + n.Name + " changed"
			out.Structure = append(out.Structure, structureEntry{"name changed", n.Name + ": " + o.Ref + " → " + n.Ref})
		default:
			continue
		}
		for _, at := range tr.Using(n.Name) {
			if _, seen := r.formulaSrc[at]; !seen {
				r.formulaSrc[at] = label
			}
		}
	}
	// Formulas that read lines removed (in the old version): what they
	// read changed, whether Excel shrank a reference or, with nothing left
	// to read, wrote #REF!.
	if len(removals) > 0 {
		ta := xlsx.NewTracer(A.wb)
		for _, g := range removals {
			for at, k := range ta.Touching(g.sheet, g.runs, g.cols) {
				m := r.lines[strings.ToLower(at.Sheet)]
				if m == nil || m.row[at.Row] == 0 || m.col[at.Col] == 0 {
					continue // its sheet, row or column was removed
				}
				here := xlsx.At{Sheet: m.sheet, Row: m.row[at.Row], Col: m.col[at.Col]}
				if _, seen := r.formulaSrc[here]; !seen && tr.Formula(here.Sheet, here.Row, here.Col) {
					r.formulaSrc[here] = map[bool]string{true: "columns removed: ", false: "rows removed: "}[g.cols] + g.labels[k]
				}
			}
		}
	}
	ats := make([]xlsx.At, 0, len(r.formulaSrc))
	for at := range r.formulaSrc {
		ats = append(ats, at)
	}
	sort.Slice(ats, func(i, j int) bool { return lessAt(ats[i], ats[j]) })
	for _, at := range ats {
		r.cause(at.Sheet, grid.Range{R1: at.Row, C1: at.Col, R2: at.Row, C2: at.Col}, r.formulaSrc[at])
	}
	r.trace(tr)
}

// reach takes in the rows added or removed (lines) that join a span of
// moved rows without a row kept between: the sort may have taken them too.
func reach(span [2]int, lines []int) [2]int {
	if span[0] == 0 {
		return span
	}
	in := make(map[int]bool, len(lines))
	for _, n := range lines {
		in[n] = true
	}
	for in[span[0]-1] {
		span[0]--
	}
	for in[span[1]+1] {
		span[1]++
	}
	return span
}

// widen extends a span of lines (zero for none) to take in from..to.
func widen(span [2]int, from, to int) [2]int {
	if span[0] == 0 {
		return [2]int{from, to}
	}
	return [2]int{min(span[0], from), max(span[1], to)}
}

func lessAt(a, b xlsx.At) bool {
	if a.Sheet != b.Sheet {
		return a.Sheet < b.Sheet
	}
	if a.Row != b.Row {
		return a.Row < b.Row
	}
	return a.Col < b.Col
}

// sameRef tells whether a name refers to the same thing, as Excel rewrites
// it for the sheets renamed and the rows and columns added and removed.
func (r *report) sameRef(old, now string) bool {
	if strings.EqualFold(old, now) {
		return true
	}
	f, ok := r.replay(old, "", false)
	return ok && strings.EqualFold(f, now)
}

// cell compares one cell in its two versions.
func (r *report) cell(sa *grid.Sheet, ra, ca int, sb *grid.Sheet, rb, cb int) {
	A, B, out := r.A, r.B, r.out
	x, y := sa.At(ra, ca), sb.At(rb, cb)
	xe, ye := x == nil || x.IsEmpty(), y == nil || y.IsEmpty()
	if xe && ye {
		return
	}
	here := cellAt(sb.Name, rb, cb)
	// A CSV file's one sheet is named after the file: the same sheet.
	renamed := !strings.EqualFold(sa.Name, sb.Name) && !(out.ValuesOnly && len(A.wb.Sheets) == 1 && len(B.wb.Sheets) == 1)
	was := ""
	if ra != rb || ca != cb || renamed {
		was = grid.CellName(ra, ca)
		if renamed {
			was = cellAt(sa.Name, ra, ca)
		}
	}
	va, vb := value(A.wb, sa, ra, ca), value(B.wb, sb, rb, cb)
	if r.out.ValuesOnly {
		if va.known && vb.known && !sameValues(va, vb) {
			out.Changed = append(out.Changed, changedCell{Cell: here, Was: was, Kind: "value", Before: va.text, After: vb.text})
		}
		return
	}
	fa, fb := !xe && x.Formula != "", !ye && y.Formula != ""
	ka, kb := r.arrays.computed(sa, ra, ca), r.arrays.computed(sb, rb, cb)
	one := grid.Range{R1: rb, C1: cb, R2: rb, C2: cb}
	var via *xlsx.At
	switch {
	case !ka && !kb:
		// Inputs.
		if (xe != ye) || !sameValues(va, vb) {
			out.Changed = append(out.Changed, changedCell{Cell: here, Was: was, Kind: "input", Before: va.text, After: vb.text})
			r.cause(sb.Name, one, "input "+here)
		}
		return
	case fa && fb:
		if folded(x.Formula) == folded(y.Formula) {
			// The same text: the formula is as it was, whatever rows were
			// moved around it, unless rows or columns added or removed moved
			// what it reads and it was not rewritten for them.
			// (A row a sort moved took its formula with it, rewritten as for a
			// copy: where it reads is its own business.)
			if want, ok := r.stale(x.Formula, sb.Name); ok && !r.cur.moved[ra] {
				c := changedCell{Cell: here, Was: was, Kind: "formula", Before: "=" + x.Formula, After: "=" + y.Formula,
					Flag: "its text is as it was, though rows or columns added or removed moved what it reads: Excel writes =" + want}
				out.Changed = append(out.Changed, c)
				r.cause(sb.Name, one, "formula "+here)
				return
			}
			break
		}
		pa, _ := formula.Pattern(folded(x.Formula), ra, ca)
		pb, _ := formula.Pattern(folded(y.Formula), rb, cb)
		if pa != pb && !r.adjusted(folded(x.Formula), sb.Name, folded(y.Formula), rb, cb) {
			c := changedCell{Cell: here, Was: was, Kind: "formula", Before: "=" + x.Formula, After: "=" + y.Formula}
			c.Flag = brokenRow(sb, rb, cb, pa, pb)
			out.Changed = append(out.Changed, c)
			r.cause(sb.Name, one, "formula "+here)
			return
		}
	case fa != fb && (fa || fb) && !(ka && kb):
		// A formula added or removed.
		c := changedCell{Cell: here, Was: was, Kind: "formula", Before: repr(x, va), After: repr(y, vb)}
		if fa && !kb {
			pa, _ := formula.Pattern(folded(x.Formula), ra, ca)
			c.Flag = brokenRow(sb, rb, cb, pa, "")
		}
		out.Changed = append(out.Changed, c)
		r.cause(sb.Name, one, "formula "+here)
		return
	case ka && !kb && ye:
		// An array's result that shrank: the cell it left moved, and
		// what moved the array moved it.
		if ar, ac, ok := r.arrays.anchor(sa, ra, ca); ok && r.cur.row[ar] > 0 && r.cur.col[ac] > 0 {
			via = &xlsx.At{Sheet: sb.Name, Row: r.cur.row[ar], Col: r.cur.col[ac]}
		}
	case ka && !kb:
		// A value where an array's result was.
		out.Changed = append(out.Changed, changedCell{Cell: here, Was: was, Kind: "input", Before: va.text, After: vb.text})
		r.cause(sb.Name, one, "input "+here)
		return
	}
	// A computed value, its formula unchanged: did it move?
	if va.known && vb.known && sameValues(va, vb) {
		return
	}
	at := xlsx.At{Sheet: sb.Name, Row: rb, Col: cb}
	if !va.known || !vb.known {
		why := "unknown in " + A.path + ": " + A.unknownWhy(sa.Name, ra, ca)
		if !vb.known {
			why = B.unknownWhy(sb.Name, rb, cb)
		}
		out.Unknown = append(out.Unknown, unknownValue{Cell: here, Before: va.text, After: vb.text, Why: why})
		r.unknownSrc = append(r.unknownSrc, xlsx.Source{Sheet: sb.Name, Range: one})
		r.unknownWhy = append(r.unknownWhy, here+", whose value sheets cannot know ("+why+")")
		return
	}
	note := B.stored[strings.ToLower(here)]
	if n := A.stored[strings.ToLower(cellAt(sa.Name, ra, ca))]; n != "" && note == "" {
		note = "in " + A.path + ", " + n
	}
	if note != "" {
		note = "computed by sheets: " + note
	}
	r.candidates = append(r.candidates, candidate{at: at, before: va, after: vb, was: was, note: note, via: via})
}

// repr shows a cell: its formula, or its value.
func repr(c *grid.Cell, v valueOf) string {
	if c != nil && c.Formula != "" {
		return "=" + c.Formula
	}
	return v.text
}

// replayLimit bounds the work of replaying lines added and removed on
// formulas: each run of lines on each formula it can reach.
const replayLimit = 2_000_000

// within tells whether a run of lines meets a span (zero for none).
func within(run, span [2]int) bool {
	return span[0] > 0 && run[0] <= span[1] && run[1] >= span[0]
}

// replay rewrites a formula, or a name's reference (home ""), as Excel does
// for the sheets renamed and the rows and columns added and removed. It
// says false when there is nothing to replay, or too much. With sure set,
// the rows added or removed within what a sort moved are left out: where
// they were cannot be told, as the sort took the rows elsewhere.
func (r *report) replay(old, home string, sure bool) (string, bool) {
	n := 0
	for _, e := range r.edits {
		n += len(e.rowsRemoved) + len(e.rowsAdded) + len(e.colsRemoved) + len(e.colsAdded)
	}
	if n == 0 && len(r.renames) == 0 || r.replayed > replayLimit {
		return "", false
	}
	f := old
	for _, rn := range r.renames {
		f = formula.RenameSheet(f, rn[0], rn[1])
	}
	// How far the formula reads on each sheet: lines past that touch nothing
	// of it.
	reach := map[string][2]int{}
	for _, ref := range formula.Refs(f) {
		sheet := strings.ToLower(ref.Sheet)
		if ref.Sheet == "" {
			sheet = strings.ToLower(home)
		}
		x := reach[sheet]
		reach[sheet] = [2]int{max(x[0], ref.Range.R2), max(x[1], ref.Range.C2)}
	}
	for _, e := range r.edits {
		x := reach[strings.ToLower(e.sheet)]
		for _, dim := range []struct {
			cols           bool
			removed, added [][2]int
			movedOld       [2]int
			movedNew       [2]int
			far            int
		}{{false, e.rowsRemoved, e.rowsAdded, e.movedOld, e.movedNew, x[0]}, {true, e.colsRemoved, e.colsAdded, [2]int{}, [2]int{}, x[1]}} {
			far := dim.far
			for i := len(dim.removed) - 1; i >= 0; i-- {
				if run := dim.removed[i]; run[0] <= far && !(sure && within(run, dim.movedOld)) {
					r.replayed++
					f = formula.DeleteLines(f, home, e.sheet, dim.cols, run[0], run[1])
				}
			}
			for _, a := range dim.added {
				if a[0] <= far && !(sure && within(a, dim.movedNew)) {
					r.replayed++
					f = formula.InsertLines(f, home, e.sheet, dim.cols, a[0], a[1]-a[0]+1)
					far += a[1] - a[0] + 1
				}
			}
		}
	}
	return f, true
}

// adjusted tells whether a formula (now in row, col of sheet home) is its
// old self rewritten as Excel does for the sheets renamed and the rows and
// columns added and removed.
func (r *report) adjusted(old, home, now string, row, col int) bool {
	f, ok := r.replay(old, home, false)
	if !ok {
		return false
	}
	// A sheet's new name goes in as typed: fold it, as old and now are.
	pf, _ := formula.Pattern(folded(f), row, col)
	pn, _ := formula.Pattern(now, row, col)
	return pf == pn
}

// stale gives what Excel would have written for a formula whose text is as
// it was, when rows or columns added or removed moved what it reads: Excel
// rewrites such a formula, so one left as it was reads other cells now (the
// usual slip of a script that inserts rows without rewriting formulas).
func (r *report) stale(old, home string) (string, bool) {
	f, ok := r.replay(old, home, true)
	return f, ok && folded(f) != folded(old)
}

// brokenRow flags one row of a formula column changed unlike the rows
// around it, which keep its old pattern: now is its new pattern ("" for a
// value).
func brokenRow(s *grid.Sheet, row, col int, old, now string) string {
	keep, follow := 0, 0
	display := ""
	for _, r := range []int{row - 1, row + 1} {
		if r < 1 {
			continue
		}
		c := s.At(r, col)
		if c == nil || c.Formula == "" {
			continue
		}
		k, _ := formula.Pattern(folded(c.Formula), r, col)
		_, d := formula.Pattern(c.Formula, r, col)
		switch k {
		case old:
			keep++
			display = d
		case now:
			follow++
		}
	}
	if keep == 0 || follow > 0 {
		return ""
	}
	if now == "" {
		return "a value where the rows around it compute =" + display
	}
	return "unlike the rows around it, which compute =" + display
}

// trace finds each moved value's cause.
func (r *report) trace(tr *xlsx.Tracer) {
	out := r.out
	r.sources = append(r.sources, r.lineSources...)
	r.causes = append(r.causes, r.lineCauses...)
	causeOf := tr.Spread(r.sources)
	// A value traces through the cell itself, or through the array formula
	// whose result it was.
	places := func(c candidate) []xlsx.At {
		if c.via != nil {
			return []xlsx.At{c.at, *c.via}
		}
		return []xlsx.At{c.at}
	}
	cause := func(m map[xlsx.At]int, c candidate) (int, bool) {
		for _, at := range places(c) {
			if k, ok := m[at]; ok {
				return k, true
			}
		}
		return 0, false
	}
	// What sheets cannot follow: formulas whose reads cannot be told, and
	// those whose values are unknown.
	var unknownOf map[xlsx.At]int
	for _, c := range r.candidates {
		if _, ok := cause(causeOf, c); ok {
			continue
		}
		if unknownOf == nil {
			unknownOf = tr.Spread(r.unknownSources(tr))
		}
		break
	}
	for _, c := range r.candidates {
		here := cellAt(c.at.Sheet, c.at.Row, c.at.Col)
		at := c.at
		if c.via != nil {
			at = *c.via
		} else if ar, ac, ok := r.arrays.anchor(r.B.wb.Sheet(at.Sheet), at.Row, at.Col); ok {
			at = xlsx.At{Sheet: at.Sheet, Row: ar, Col: ac}
		}
		if k, ok := cause(causeOf, c); ok {
			out.Moved = append(out.Moved, movedValue{Cell: here, Was: c.was, Before: c.before.text, After: c.after.text, Cause: r.causes[k], Note: c.note})
			continue
		}
		if why := tr.Opaque(at.Sheet, at.Row, at.Col); why != "" {
			out.Unknown = append(out.Unknown, unknownValue{Cell: here, Before: c.before.text, After: c.after.text, Why: "moved, and sheets cannot tell why: " + why})
			continue
		}
		if k, ok := cause(unknownOf, c); ok {
			out.Unknown = append(out.Unknown, unknownValue{Cell: here, Before: c.before.text, After: c.after.text, Why: "moved, and sheets cannot tell why: it reads " + r.unknownWhy[k]})
			continue
		}
		out.Unexplained = append(out.Unexplained, movedValue{Cell: here, Was: c.was, Before: c.before.text, After: c.after.text, Note: c.note})
	}
}

// unknownSources are the formulas sheets cannot follow: whose reads
// cannot be told, and whose values are unknown.
func (r *report) unknownSources(tr *xlsx.Tracer) []xlsx.Source {
	srcs := append([]xlsx.Source{}, r.unknownSrc...)
	for _, s := range r.B.wb.Sheets {
		s.Each(s.Used(), func(row, col int, c *grid.Cell) {
			if c.Formula == "" {
				return
			}
			if why := tr.Opaque(s.Name, row, col); why != "" {
				srcs = append(srcs, xlsx.Source{Sheet: s.Name, Range: grid.Range{R1: row, C1: col, R2: row, C2: col}})
				r.unknownWhy = append(r.unknownWhy, cellAt(s.Name, row, col)+", whose reads sheets cannot follow ("+why+")")
			}
		})
	}
	return srcs
}

// lostFeatures are what the old version carries that the new one lost: a
// feature matches one of its kind on the same sheet (renamed or not) in
// the same place, or, failing that, anywhere on that sheet (it moved or
// grew); names and tables match by name.
func lostFeatures(before, after []xlsx.Feature, renamed map[string]string) []lostFeature {
	used := make([]bool, len(after))
	sheetOf := func(f xlsx.Feature) (string, bool) {
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
		return strings.ToLower(to), true
	}
	identity := func(f xlsx.Feature) string {
		switch f.Kind {
		case xlsx.FeatureName:
			return strings.ToLower(f.Where)
		case xlsx.FeatureTable:
			return strings.ToLower(f.What)
		}
		return strings.ToLower(f.Where)
	}
	match := func(f xlsx.Feature, loose bool) bool {
		sheet, ok := sheetOf(f)
		if !ok {
			return false
		}
		for i, g := range after {
			if used[i] || g.Kind != f.Kind {
				continue
			}
			if loose && f.Chartsheet && g.Chartsheet {
				// A chart sheet is no worksheet, so a rename of it is not
				// followed: its chart is the one chart sheet left.
				used[i] = true
				return true
			}
			if strings.ToLower(g.Sheet) != sheet {
				continue
			}
			named := f.Kind == xlsx.FeatureName || f.Kind == xlsx.FeatureTable
			if identity(g) == identity(f) || loose && !named {
				used[i] = true
				return true
			}
		}
		return false
	}
	kept := make([]bool, len(before))
	for i, f := range before {
		kept[i] = match(f, false)
	}
	var out []lostFeature
	for i, f := range before {
		if kept[i] || match(f, true) {
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

// againstFails tells whether --against exits 1: values moved with no
// cause, features lost.
func againstFails(r *againstResult) bool {
	return len(r.Unexplained) > 0 || len(r.FeaturesLost) > 0
}

func (a *app) printAgainst(r *againstResult) {
	s := r.Summary
	head := fmt.Sprintf("%s against %s: ", r.File, r.Against)
	var parts []string
	if r.ValuesOnly {
		parts = append(parts, countOf(s.Inputs, "value")+" changed")
	} else {
		parts = append(parts, countOf(s.Inputs, "input")+" changed", countOf(s.Formulas, "formula")+" changed")
		moved := countOf(s.Moved, "computed value") + " moved"
		if s.Moved > 0 {
			moved += " (each with its cause)"
		}
		parts = append(parts, moved, grid.FormatNumber(float64(s.Unexplained))+" unexplained", countOf(s.FeaturesLost, "feature")+" lost", grid.FormatNumber(float64(s.Unknown))+" unknown")
	}
	if s.MovedRows > 0 {
		parts = append(parts, countOf(s.MovedRows, "row")+" moved")
	}
	a.printf("%s%s\n", head, strings.Join(parts, ", "))
	if r.ValuesOnly {
		a.printf("a CSV file holds values only: formulas, causes and features are not compared\n")
	}
	more := func(n int) {
		if n > againstListed {
			a.printf("  … %s more (--json lists them all)\n", grid.FormatNumber(float64(n-againstListed)))
		}
	}
	section := func(title string, n int) bool {
		if n == 0 {
			return false
		}
		a.printf("%s:\n", title)
		return true
	}
	if section("unexplained: values that moved with nothing upstream changed", len(r.Unexplained)) {
		tw := a.table()
		for _, m := range r.Unexplained[:min(len(r.Unexplained), againstListed)] {
			tw.row("  "+withWas(m.Cell, m.Was), orEmpty(m.Before)+" → "+orEmpty(m.After), m.Note)
		}
		tw.Flush()
		more(len(r.Unexplained))
	}
	if section("features lost", len(r.FeaturesLost)) {
		tw := a.table()
		for _, f := range r.FeaturesLost[:min(len(r.FeaturesLost), againstListed)] {
			tw.row("  "+f.Kind, f.Where, f.What)
		}
		tw.Flush()
		more(len(r.FeaturesLost))
	}
	var flagged, plain []changedCell
	for _, c := range r.Changed {
		if c.Flag != "" {
			flagged = append(flagged, c)
		} else {
			plain = append(plain, c)
		}
	}
	if section("flagged: formula changes that look like mistakes", len(flagged)) {
		for _, c := range flagged[:min(len(flagged), againstListed)] {
			a.printf("  %s  %s → %s  (%s)\n", withWas(c.Cell, c.Was), orEmpty(c.Before), orEmpty(c.After), c.Flag)
		}
		more(len(flagged))
	}
	if section("changed", len(plain)) {
		tw := a.table()
		for _, c := range plain[:min(len(plain), againstListed)] {
			tw.row("  "+withWas(c.Cell, c.Was), c.Kind, orEmpty(c.Before)+" → "+orEmpty(c.After))
		}
		tw.Flush()
		more(len(plain))
	}
	if section("moved, each with its cause", len(r.Moved)) {
		tw := a.table()
		for _, m := range r.Moved[:min(len(r.Moved), againstListed)] {
			tw.row("  "+withWas(m.Cell, m.Was), orEmpty(m.Before)+" → "+orEmpty(m.After), m.Cause)
		}
		tw.Flush()
		more(len(r.Moved))
	}
	if section("unknown: values sheets cannot compare", len(r.Unknown)) {
		tw := a.table()
		for _, u := range r.Unknown[:min(len(r.Unknown), againstListed)] {
			tw.row("  "+u.Cell, orEmpty(u.Before)+" → "+orEmpty(u.After), u.Why)
		}
		tw.Flush()
		more(len(r.Unknown))
	}
	if section("rows moved (as by a sort; not changes)", len(r.MovedRows)) {
		for _, m := range r.MovedRows[:min(len(r.MovedRows), againstListed)] {
			a.printf("  %s, were %s%s\n", lineRange(m.Sheet, m.To, m.To+m.Rows-1, false), strings.TrimPrefix(lineRange(m.Sheet, m.From, m.From+m.Rows-1, false), grid.QuoteSheet(m.Sheet)+"!"), sameNumber(m))
		}
		more(len(r.MovedRows))
	}
	if section("structure", len(r.Structure)) {
		for _, e := range r.Structure[:min(len(r.Structure), againstListed)] {
			a.printf("  %s: %s\n", e.Change, e.Where)
		}
		more(len(r.Structure))
	}
	a.printf("check of %s: ", r.File)
	if len(r.Findings) == 0 {
		a.printf("no problems found\n")
	} else {
		a.printf("%s, %s, %s\n", plural(r.Errors, "error"), plural(r.Warnings, "warning"), plural(r.Notes, "note"))
		// The worst first, as check lists them; a file can have thousands.
		fs := append([]checkFinding{}, r.Findings...)
		sort.SliceStable(fs, func(i, j int) bool { return fs[i].Severity > fs[j].Severity })
		tw := a.table()
		for _, f := range fs[:min(len(fs), againstListed)] {
			tw.row(f.Severity.String(), findingCells(f), f.Message)
		}
		tw.Flush()
		more(len(fs))
	}
}

// table lines up columns, as check's report does, without trailing
// spaces where a line's last column is empty.
type table struct {
	a   *app
	buf bytes.Buffer
	tw  *tabwriter.Writer
}

func (a *app) table() *table {
	t := &table{a: a}
	t.tw = tabwriter.NewWriter(&t.buf, 0, 0, 2, ' ', 0)
	return t
}

func (t *table) row(cells ...string) {
	fmt.Fprintln(t.tw, strings.Join(cells, "\t"))
}

func (t *table) Flush() {
	t.tw.Flush()
	for _, l := range strings.SplitAfter(t.buf.String(), "\n") {
		if l != "" {
			t.a.printf("%s\n", strings.TrimRight(l, " \n"))
		}
	}
}

func withWas(cell, was string) string {
	if was == "" {
		return cell
	}
	return cell + " (was " + was + ")"
}

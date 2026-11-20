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
	// Receipt is what --receipt wrote.
	Receipt *receiptInfo `json:"receipt,omitempty"`
	// sides are the two versions (earlier, then the file checked), and
	// renamed the sheets renamed (old name lower-cased, new name; "" for
	// a sheet removed), for the receipt.
	sides   [2]*sideOf
	renamed map[string]string
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
	Was    string `json:"was,omitempty"` // where it was, when not the same place
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
	// Columns are the columns that moved, as A:D, when a sort of a range
	// left the cells beside it where they were; "" for whole rows.
	Columns string `json:"columns,omitempty"`
	// WasColumns are the columns they were in, when cells were cut and
	// pasted into other columns (Columns are where they are now).
	WasColumns   string `json:"was_columns,omitempty"`
	c1, c2       int
	wasC1, wasC2 int
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

// causeListed bounds the cells changed a moved value names as its cause (then
// "and more"), in text and --json alike.
const causeListed = 5

// againstWidth bounds what text output shows of a value or formula,
// againstNote of a cause, flag or reason, and againstFinding of the check's
// own findings, which end with what to do (as sheets diff cuts values);
// --json gives them whole.
const (
	againstWidth   = 120
	againstNote    = 240
	againstFinding = 400
)

// clip cuts a value or formula for text output; clipNote a cause, flag or
// reason.
func clip(s string) string     { return grid.Truncate(s, againstWidth) }
func clipNote(s string) string { return grid.Truncate(s, againstNote) }

// change shows a value or formula before and after, each cut.
func change(before, after string) string {
	return orEmpty(clip(before)) + " → " + orEmpty(clip(after))
}

// cellAt names a cell: Sheet!A1.
func cellAt(sheet string, row, col int) string {
	return grid.QuoteSheet(sheet) + "!" + grid.CellName(row, col)
}

// valueOf is a cell's value as a reader sees it, and whether sheets knows
// it.
type valueOf struct {
	text string
	cell grid.Cell
	// num is a number as the 1900 date system counts it: a date in a 1904
	// workbook is 1,462 days earlier by its serial, which says nothing of
	// the date. shift is what was added to the stored number for that.
	num   float64
	shift float64
	known bool
}

// epochDays are the days between the 1900 and the 1904 date systems.
const epochDays = 1462

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
	return valueFrom(wb, c)
}

// valueFrom reads a cell that holds a value.
func valueFrom(wb *grid.Workbook, c *grid.Cell) valueOf {
	v := valueOf{text: valueText(wb, c), cell: *c, num: c.Num, known: true}
	if f := wb.FormatOf(c); c.Kind == grid.Number && wb.Date1904 && (f.Kind == grid.Date || f.Kind == grid.DateTime) {
		v.num += epochDays
		v.shift = epochDays
	}
	return v
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

// sameResults tells whether two known computed values are the same,
// numbers to Excel's 15 significant digits: what its calculator and ours,
// or two producers writing one result, may differ in is the digits past
// them. A result under 1e-15 is as good as zero: what cancellation leaves
// of a sum or a difference.
func sameResults(a, b valueOf) bool { return sameValues(a, b, false) }

// sameInputs tells whether two known typed values (a number, text, a CSV
// file's cell) are the same, exactly: a barcode, an order number or an
// amount in the millions differs in digits a tolerance would take for
// rounding. Numbers compare as the files store them; across the two date
// systems, as the days they count to 15 digits.
func sameInputs(a, b valueOf) bool { return sameValues(a, b, true) }

func sameValues(a, b valueOf, exact bool) bool {
	x, y := a.cell, b.cell
	if x.Kind != y.Kind {
		return false
	}
	switch x.Kind {
	case grid.Number:
		if exact && a.shift == b.shift {
			return x.Num == y.Num
		}
		d := math.Abs(a.num - b.num)
		return d <= 5e-15*math.Max(math.Abs(a.num), math.Abs(b.num)) || d < 1e-15
	case grid.Text, grid.Error:
		return x.Text == y.Text
	case grid.Bool:
		return x.Bool == y.Bool
	}
	return true
}

// inputTexts show two inputs that differ as they read: when both read the
// same (a number differing past the digits its format shows), by the digits
// the file stores.
func inputTexts(a, b valueOf) (string, string) {
	if a.text == b.text && a.cell.Kind == grid.Number && b.cell.Kind == grid.Number {
		return strconv.FormatFloat(a.cell.Num, 'g', -1, 64), strconv.FormatFloat(b.cell.Num, 'g', -1, 64)
	}
	return a.text, b.text
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
	// differ are the results the file stores that the sheets calculator,
	// from what each formula reads as stored, gets otherwise (check keeps
	// Excel's): by cellAt, lower-cased.
	differ map[string]xlsx.Difference
	csv    bool
	// fixed is the file with the results sheets computes stored (nil when
	// that changes nothing), which the receipt's pictures are drawn from.
	fixed []byte
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
	x := &sideOf{path: path, why: map[string]string{}, stored: map[string]string{}, differ: map[string]xlsx.Difference{}}
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
		// Compared as check compares, so a result it stores that the sheets
		// calculator disagrees with is known for what it is.
		if cr, err = xlsx.CheckResults(path, x.data, checkTime); err != nil {
			return nil, err
		}
	}
	for _, d := range cr.Differ {
		x.differ[strings.ToLower(cellAt(d.Sheet, d.Row, d.Col))] = d
	}
	x.wb = cr.After
	for _, u := range cr.Unknown {
		x.why[strings.ToLower(cellAt(u.Sheet, u.Row, u.Col))] = u.Why
	}
	if cr.FixError != "" {
		x.wb = cr.Before // values computed as read
	} else {
		x.fixed = cr.Data
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
	if !out.ValuesOnly {
		if r.fa, err = xlsx.Features(A.data); err != nil {
			return nil, fmt.Errorf("%s: %w", before, err)
		}
		if r.fb, err = xlsx.Features(B.data); err != nil {
			return nil, fmt.Errorf("%s: %w", path, err)
		}
	}
	r.compare()
	out.sides, out.renamed = [2]*sideOf{A, B}, r.renamed
	if !out.ValuesOnly {
		out.FeaturesLost = lostFeatures(r.fa, r.fb, r.renamed, r.tables, r.shiftRef)
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
	// cells is how many of the sources come first as cells changed (inputs
	// and formulas).
	cells int
	// lineSources are rows, columns and sheets added or moved, which come
	// after the cells changed.
	lineSources []xlsx.Source
	lineCauses  []string
	// candidates are computed values that moved, to trace; held those that
	// did not, whose results the sheets calculator disagrees with.
	candidates []candidate
	held       []heldValue
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
	// fa and fb are the features each version carries; tables the Excel
	// tables renamed (old name lower-cased to new), tableRenames the same in
	// order.
	fa, fb       []xlsx.Feature
	tables       map[string]string
	tableRenames [][2]string
	// anyLines: some sheet had lines added or removed; anyCut: cells cut and
	// pasted.
	anyLines, anyCut bool
	// removedSheets are the sheets removed, as the old version names them;
	// removedRefs find what DeleteLines leaves of a reference to each.
	removedSheets []string
	removedRefs   []*regexp.Regexp
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
	// blocks are columns whose rows a sort moved on its own.
	blocks []blockMap
	// pasted are the old rows cut and pasted (job.pasted), in the columns
	// pasteC1 to pasteC2 (0 for whole rows).
	pasted           map[int]bool
	pasteC1, pasteC2 int
	// pasteDC is how many columns right the cells cut and pasted went; cells
	// is the block of them when it is among others that stayed (as numbered
	// before).
	pasteDC int
	cells   grid.Range
}

// pastedAt tells whether an old cell was cut and pasted elsewhere.
func (m *lineMap) pastedAt(row, col int) bool {
	if m.cells.R1 > 0 {
		return m.cells.Contains(row, col)
	}
	return m.pasted[row]
}

// blockMap is lineMap's for the columns of a block (diff.Block): the old
// columns, the new number of each row kept there, the rows a sort moved,
// and the new columns it spans.
type blockMap struct {
	cols   map[int]bool
	row    map[int]int
	moved  map[int]bool
	c1, c2 int
}

// block gives the block an old column is in, or nil.
func (m *lineMap) block(col int) *blockMap {
	for i := range m.blocks {
		if m.blocks[i].cols[col] {
			return &m.blocks[i]
		}
	}
	return nil
}

// rowOf gives the new number of an old cell's row (0: removed).
func (m *lineMap) rowOf(row, col int) int {
	if b := m.block(col); b != nil {
		return b.row[row]
	}
	return m.row[row]
}

// movedAt tells whether a sort moved an old cell's row.
func (m *lineMap) movedAt(row, col int) bool {
	if b := m.block(col); b != nil {
		return b.moved[row]
	}
	return m.moved[row]
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
	// cut is cells cut and pasted on the sheet (job.cut).
	cut []formula.Move
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
	// cause is the cell's own change, when it is a formula changed.
	cause string
	// shift is a formula that a sort or lines added or removed moved, or
	// whose references Excel rewrote for them: what moved it, when nothing
	// else did, moved its value (ROW(), an implicit intersection).
	shift *shiftOf
}

// heldValue is a computed value that did not move, whose result the new
// version stores and the sheets calculator gets otherwise (calc): where it
// was (cellAt, lower-cased), and its value.
type heldValue struct {
	at          xlsx.At
	old         string
	was         string // as given for a changed place, else empty
	value, calc valueOf
}

// shiftOf is where a formula was, and its text before and now.
type shiftOf struct {
	sheet    string // the old version's name
	row, col int
	old, now string
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

// heldNumber notes rows that moved yet kept their numbers: the rows around
// them moved, or rows were added or removed above them.
func heldNumber(m movedRows) string {
	if m.From == m.To && m.c1 == m.wasC1 {
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

// resultPair is an array formula in both versions whose result spans
// several cells in either: its formula's cell before and now, and the
// larger of the results' extents.
type resultPair struct {
	ra, ca, rb, cb int
	rows, cols     int
}

// multi gives a sheet's array formulas' results over several cells, by
// their formula's cell.
func multi(s *grid.Sheet) map[[2]int]grid.Range {
	out := map[[2]int]grid.Range{}
	for _, rg := range append(append([]grid.Range{}, s.Spills...), s.Arrays...) {
		if rg.Rows()*rg.Cols() > 1 {
			k := [2]int{rg.R1, rg.C1}
			if o, ok := out[k]; !ok || rg.Rows()*rg.Cols() > o.Rows()*o.Cols() {
				out[k] = rg
			}
		}
	}
	return out
}

// results pairs a sheet's array formulas over several cells by their
// formula's cell (paired as its row and column are), so their results
// compare cell by cell from there: a FILTER that lost a row, a SORT sorted
// again, rows inserted or deleted through a SEQUENCE. It gives the pairs,
// and the cells of the results they compare, before and now.
func (r *report) results(j *job, m *lineMap) ([]resultPair, map[[2]int]bool, map[[2]int]bool) {
	ma, mb := multi(j.sa), multi(j.sb)
	if len(ma)+len(mb) == 0 {
		return nil, nil, nil
	}
	formulaAt := func(s *grid.Sheet, row, col int) bool {
		c := s.At(row, col)
		return c != nil && c.Formula != ""
	}
	// Each new cell's old one, for the formulas only the new version
	// spills (the lines' maps the other way, made once).
	var colInv, rowInv map[int]int
	blockInv := map[int]map[int]int{}
	oldOf := func(row, col int) (int, int) {
		if colInv == nil {
			colInv, rowInv = map[int]int{}, map[int]int{}
			for o, n := range m.col {
				colInv[n] = o
			}
			for o, n := range m.row {
				rowInv[n] = o
			}
		}
		oc := colInv[col]
		if oc == 0 {
			return 0, 0
		}
		inv := rowInv
		if b := m.block(oc); b != nil {
			if inv = blockInv[b.c1]; inv == nil {
				inv = map[int]int{}
				for o, n := range b.row {
					inv[n] = o
				}
				blockInv[b.c1] = inv
			}
		}
		return inv[row], oc
	}
	var out []resultPair
	seen := map[[2]int]bool{} // new cells
	add := func(ra, ca, rb, cb int) {
		if seen[[2]int{rb, cb}] || !formulaAt(j.sa, ra, ca) || !formulaAt(j.sb, rb, cb) {
			return
		}
		seen[[2]int{rb, cb}] = true
		p := resultPair{ra: ra, ca: ca, rb: rb, cb: cb, rows: 1, cols: 1}
		if rg, ok := ma[[2]int{ra, ca}]; ok {
			p.rows, p.cols = max(p.rows, rg.Rows()), max(p.cols, rg.Cols())
		}
		if rg, ok := mb[[2]int{rb, cb}]; ok {
			p.rows, p.cols = max(p.rows, rg.Rows()), max(p.cols, rg.Cols())
		}
		out = append(out, p)
	}
	for k := range ma {
		if rb, cb := m.rowOf(k[0], k[1]), m.col[k[1]]; rb > 0 && cb > 0 {
			add(k[0], k[1], rb, cb)
		}
	}
	for k := range mb {
		if !seen[k] {
			if ra, ca := oldOf(k[0], k[1]); ra > 0 {
				add(ra, ca, k[0], k[1])
			}
		}
	}
	sort.Slice(out, func(x, y int) bool {
		if out[x].rb != out[y].rb {
			return out[x].rb < out[y].rb
		}
		return out[x].cb < out[y].cb
	})
	oldIn, newIn := map[[2]int]bool{}, map[[2]int]bool{}
	for _, p := range out {
		for dr := 0; dr < p.rows; dr++ {
			for dc := 0; dc < p.cols; dc++ {
				if dr > 0 || dc > 0 {
					oldIn[[2]int{p.ra + dr, p.ca + dc}] = true
					newIn[[2]int{p.rb + dr, p.cb + dc}] = true
				}
			}
		}
	}
	return out, oldIn, newIn
}

// empty tells whether a cell holds nothing.
func empty(s *grid.Sheet, row, col int) bool {
	c := s.At(row, col)
	return c == nil || c.IsEmpty()
}

// pasted pairs again rows cut and pasted where nothing was (sheets move,
// Excel's cut and paste): a run of rows that pairs in order elsewhere,
// leaving nothing where it was and finding nothing where it went, while the
// rows around it stayed (a total reading it, its formula rewritten to follow
// it). Pairing rows as they align, such a sheet has rows added where the run
// was and removed where the total was. The run pairs as it moved, every
// other row in place, when each row paired otherwise than before holds what
// it held, its formulas as Excel rewrites them for the move; it gives the
// rows, and records the run in j.
func (r *report) pasted(j *job, rows []diff.Pair) []diff.Pair {
	sa, sb := j.sa, j.sb
	added := false
	for _, q := range rows {
		added = added || q.Old == 0 || q.New == 0
	}
	if !added {
		return rows
	}
	// The one run of rows with something in them paired in order elsewhere.
	ua, ub := sa.Used(), sb.Used()
	var paired []diff.Pair
	for _, q := range rows {
		if q.Old > 0 && q.New > 0 && q.Old != q.New && !rowEmpty(sa, ua, q.Old) {
			paired = append(paired, q)
		}
	}
	if len(paired) == 0 {
		return rows
	}
	sort.Slice(paired, func(x, y int) bool { return paired[x].Old < paired[y].Old })
	o1, d := paired[0].Old, paired[0].New-paired[0].Old
	for k, q := range paired {
		if q.Old != o1+k || q.New-q.Old != d {
			return rows // more than one run
		}
	}
	o2 := o1 + len(paired) - 1
	inSrc := func(row int) bool { return row >= o1 && row <= o2 }
	inDst := func(row int) bool { return row >= o1+d && row <= o2+d }
	// The columns it spans, and where it was and went empty there.
	c1, c2 := 0, 0
	for row := o1; row <= o2; row++ {
		for c := ua.C1; c <= ua.C2; c++ {
			if !empty(sa, row, c) {
				if c1 == 0 || c < c1 {
					c1 = c
				}
				c2 = max(c2, c)
			}
		}
	}
	cols := func(s *grid.Sheet, row int) bool {
		for c := c1; c <= c2; c++ {
			if !empty(s, row, c) {
				return false
			}
		}
		return true
	}
	var vacated, filled []int // new rows it left, old rows it took
	for row := o1; row <= o2; row++ {
		if !inDst(row) {
			if !cols(sb, row) {
				return rows
			}
			vacated = append(vacated, row)
		}
		if !inSrc(row + d) {
			if !cols(sa, row+d) {
				return rows
			}
			filled = append(filled, row+d)
		}
	}
	// Every other row in place; the rows it left and took pair, empty.
	was := make(map[int]int, len(rows)) // old row → new, as aligned
	for _, q := range rows {
		if q.Old > 0 {
			was[q.Old] = q.New
		}
	}
	move := formula.Move{From: sb.Name, Src: grid.Range{R1: o1, C1: c1, R2: o2, C2: c2}, To: sb.Name, DR: d}
	n := max(ua.R2, ub.R2, o2+d)
	out := make([]diff.Pair, 0, n)
	other := map[int]int{}
	for k, row := range filled {
		other[row] = vacated[k]
	}
	// A row with something in it that it passed, which stayed: else rows
	// added or removed above it moved it, as aligned.
	stayed := false
	for row := min(o1, o1+d); row <= max(o2, o2+d) && !stayed; row++ {
		stayed = !inSrc(row) && !rowEmpty(sa, ua, row)
	}
	if !stayed {
		return rows
	}
	for row := 1; row <= n; row++ {
		q := diff.Pair{Old: row, New: row}
		switch {
		case inSrc(row):
			q.New = row + d
		case other[row] > 0:
			q.New = other[row]
		}
		if was[row] != q.New && !r.holds(sa, q.Old, sb, q.New, j.cols, move) {
			return rows
		}
		out = append(out, q)
	}
	sort.Slice(out, func(x, y int) bool { return out[x].New < out[y].New })
	j.pasted = []diff.Moved{{From: o1, To: o1 + d, Rows: o2 - o1 + 1}}
	j.cut = []formula.Move{move}
	if c1 > max(ua.C1, ub.C1, 1) || c2 < max(ua.C2, ub.C2) {
		j.pasteC1, j.pasteC2 = c1, c2 // the columns of the rows, not whole rows
	}
	return out
}

// holds tells whether a row holds what it held, its formulas as Excel
// rewrites them for a move, or as they were.
func (r *report) holds(sa *grid.Sheet, ra int, sb *grid.Sheet, rb int, cols [][2]int, m formula.Move) bool {
	for _, c := range cols {
		if !r.holdsCell(sa, ra, c[0], sb, rb, c[1], m) {
			return false
		}
	}
	return true
}

// holdsCell tells whether a cell holds what another held, a formula as
// Excel rewrites it for a move, or as it was.
func (r *report) holdsCell(sa *grid.Sheet, ra, ca int, sb *grid.Sheet, rb, cb int, m formula.Move) bool {
	x, y := sa.At(ra, ca), sb.At(rb, cb)
	xe, ye := x == nil || x.IsEmpty(), y == nil || y.IsEmpty()
	switch {
	case xe && ye:
	case xe != ye:
		return false
	case x.Formula != "" && y.Formula != "":
		// As Excel rewrites it, or as it was (openpyxl's move_range: the
		// check of its formula flags that).
		f := formula.MoveRefs(r.renamedIn(x.Formula), sb.Name, sb.Name, m)
		pa, _ := formula.Pattern(folded(f), rb, cb)
		pb, _ := formula.Pattern(folded(y.Formula), rb, cb)
		if pa != pb && folded(x.Formula) != folded(y.Formula) {
			return false
		}
	case x.Formula != "" || y.Formula != "":
		return false
	case !sameResults(value(r.A.wb, sa, ra, ca), value(r.B.wb, sb, rb, cb)):
		return false
	}
	return true
}

// cutCells finds cells cut and pasted among others that stayed (sheets move,
// Excel's cut and paste of a block beside a table, or below it): cells that
// vacated one rectangle of the sheet and filled another of the same size,
// each holding what it held, formulas as Excel rewrites them for the move,
// while the cells around stayed where they were (a few may have changed).
// Pairing rows as they align, such a sheet has cells emptied where the block
// was and rows added where it went, or the rows of the block taken for a
// sort. It records the block in j, and pairs the rows and columns as they
// are numbered.
func (r *report) cutCells(j *job) bool {
	sa, sb := j.sa, j.sb
	ua, ub := sa.Used(), sb.Used()
	if ua.R2 == 0 || ub.R2 == 0 {
		return false
	}
	// The cells that went empty, and those that were filled.
	var src, dst grid.Range
	grow := func(rg *grid.Range, row, col int) {
		if rg.R1 == 0 {
			*rg = grid.Range{R1: row, C1: col, R2: row, C2: col}
			return
		}
		rg.R1, rg.C1, rg.R2, rg.C2 = min(rg.R1, row), min(rg.C1, col), max(rg.R2, row), max(rg.C2, col)
	}
	sa.Each(ua, func(row, col int, _ *grid.Cell) {
		if empty(sb, row, col) {
			grow(&src, row, col)
		}
	})
	sb.Each(ub, func(row, col int, _ *grid.Cell) {
		if empty(sa, row, col) {
			grow(&dst, row, col)
		}
	})
	if src.R1 == 0 || dst.R1 == 0 || src.Rows() != dst.Rows() || src.Cols() != dst.Cols() {
		return false
	}
	dr, dc := dst.R1-src.R1, dst.C1-src.C1
	if src.R1 <= dst.R2 && dst.R1 <= src.R2 && src.C1 <= dst.C2 && dst.C1 <= src.C2 {
		return false // the two overlap: not told from cells changed
	}
	move := formula.Move{From: sb.Name, Src: src, To: sb.Name, DR: dr, DC: dc}
	for row := src.R1; row <= src.R2; row++ {
		for col := src.C1; col <= src.C2; col++ {
			if !empty(sb, row, col) || !empty(sa, row+dr, col+dc) || !r.holdsCell(sa, row, col, sb, row+dr, col+dc, move) {
				return false
			}
		}
	}
	// The cells around stayed: a sort moves the cells of many rows, and a
	// few others changed is a block cut among edits.
	others := 0
	changed := map[[2]int]bool{} // where they are
	sa.Each(ua, func(row, col int, x *grid.Cell) {
		if others > cutOthers || src.Contains(row, col) || dst.Contains(row, col) {
			return
		}
		if _, _, ok := r.arrays.anchor(sa, row, col); ok {
			return // part of an array formula's result: computed, and moved with what it reads
		}
		y := sb.At(row, col)
		if y == nil || y.IsEmpty() {
			others = cutOthers + 1
			return
		}
		same := x.Formula == y.Formula && (x.Formula != "" || x.Kind == y.Kind && x.Num == y.Num && x.Text == y.Text && x.Bool == y.Bool)
		if !same && x.Formula != "" && y.Formula != "" {
			f := formula.MoveRefs(r.renamedIn(x.Formula), sb.Name, sb.Name, move)
			same = folded(f) == folded(y.Formula)
		}
		if !same {
			others++
			changed[[2]int{row, col}] = true
		}
	})
	if others > cutOthers {
		return false
	}
	// Straight down or up, a cut leaves the other cells of the rows it left and
	// went to, and of its columns, as they were; a sort of some columns that
	// carried a blank cell does not (the rows the blank left and went to hold
	// other labels, and the column's cells between them are others).
	if dc == 0 {
		for at := range changed {
			if at[0] >= src.R1 && at[0] <= src.R2 || at[0] >= dst.R1 && at[0] <= dst.R2 || at[1] >= src.C1 && at[1] <= src.C2 {
				return false
			}
		}
	}
	// Moved straight down or up with nothing else in the rows it spans or
	// below (or straight across, in the columns): that is as well rows (or
	// columns) added or removed above it, which move all there is.
	if dc == 0 || dr == 0 {
		lo := map[bool]int{true: min(src.R1, dst.R1), false: min(src.C1, dst.C1)}[dc == 0]
		beyond := func(s *grid.Sheet, u grid.Range, block grid.Range) bool {
			found := false
			s.Each(u, func(row, col int, _ *grid.Cell) {
				if found || block.Contains(row, col) {
					return
				}
				found = dc == 0 && row >= lo || dc != 0 && col >= lo
			})
			return found
		}
		if !beyond(sa, ua, src) && !beyond(sb, ub, dst) {
			return false
		}
	}
	n := max(ua.R2, ub.R2)
	j.rows = make([]diff.Pair, 0, n)
	for row := 1; row <= n; row++ {
		j.rows = append(j.rows, diff.Pair{Old: row, New: row})
	}
	j.cols, j.colsAdded, j.colsRemoved, j.moved, j.blocks = nil, nil, nil, nil, nil
	for col := 1; col <= max(ua.C2, ub.C2); col++ {
		j.cols = append(j.cols, [2]int{col, col})
	}
	j.pasted = []diff.Moved{{From: src.R1, To: dst.R1, Rows: src.Rows()}}
	if dc != 0 || src.C1 > min(ua.C1, ub.C1) || src.C2 < max(ua.C2, ub.C2) {
		j.pasteC1, j.pasteC2, j.pasteDC = dst.C1, dst.C2, dc // the cells, not whole rows
	}
	j.cut = []formula.Move{move}
	j.cells = &cellCut{src: src, dst: dst, dr: dr, dc: dc}
	return true
}

// cutOthers is how many cells beside a block cut and pasted may have changed
// (edits made with it) for it still to be the block that moved.
const cutOthers = 5

// cellCut is a block of cells cut and pasted among others that stayed (the
// cells compare through it, not as the rows pair): where it was and went, and
// how far.
type cellCut struct {
	src, dst grid.Range // as numbered before and now
	dr, dc   int
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

// job is a sheet in both versions, paired: what compare works through.
type job struct {
	sa, sb *grid.Sheet
	// rows pair the sheet's rows; cols are the columns compared through
	// them (those of blocks aside).
	rows                                           []diff.Pair
	cols                                           [][2]int
	rowsAdded, rowsRemoved, colsAdded, colsRemoved []int
	moved                                          []diff.Moved
	blocks                                         []blockJob
	// pasted is rows cut and pasted where nothing was (sheets move, Excel's
	// cut and paste), the rows around them where they were: the run, the
	// columns it spans (0 for whole rows), and the move as formulas follow
	// it (on the sheet as now named).
	pasted           []diff.Moved
	pasteC1, pasteC2 int
	// pasteDC is how many columns right the cells went (negative: left), and
	// cells the block, when it is cells cut and pasted among others that
	// stayed rather than rows.
	pasteDC int
	cells   *cellCut
	cut     []formula.Move
}

// blockJob is columns whose rows a sort moved on its own (diff.Block): its
// column pairs, how its rows pair, the rows moved, and the new columns it
// spans.
type blockJob struct {
	cols   [][2]int
	rows   []diff.Pair
	moved  []diff.Moved
	c1, c2 int
}

// lineCount counts the lines a job has added and removed.
func (j *job) lineCount() int {
	return len(j.rowsAdded) + len(j.rowsRemoved) + len(j.colsAdded) + len(j.colsRemoved)
}

// rowBlank tells whether a row of a sheet holds nothing in some columns
// (pick: which of each pair).
func rowBlank(s *grid.Sheet, row int, cols [][2]int, pick int) bool {
	for _, c := range cols {
		if cell := s.At(row, c[pick]); cell != nil && !cell.IsEmpty() {
			return false
		}
	}
	return true
}

// prepare works out a sheet's job from its alignment: rows replaced and
// rows a sort moved and changed paired, the lines added and removed.
func (r *report) prepare(p diff.SheetPair) *job {
	sa, sb := p.Old, p.New
	j := &job{sa: sa, sb: sb}
	p.Rows = pairReplaced(p.Rows)
	inBlock := map[int]bool{}
	for _, b := range p.Blocks {
		for _, c := range b.Cols {
			inBlock[c.Old] = true
		}
	}
	for _, c := range p.Cols {
		switch {
		case c.Old == 0:
			j.colsAdded = append(j.colsAdded, c.New)
		case c.New == 0:
			j.colsRemoved = append(j.colsRemoved, c.Old)
		case !inBlock[c.Old]:
			j.cols = append(j.cols, [2]int{c.Old, c.New})
		}
	}
	if len(p.Moved) > 0 {
		// Rows sorted and changed too: alike, they moved. Of all the rows
		// paired now, those whose rank among them changed are the rows
		// moved.
		p.Rows = r.pairAlike(sa, sb, p.Rows, j.cols)
		ua, ub := sa.Used(), sb.Used()
		p.Moved = diff.MovedRows(p.Rows, func(q diff.Pair) bool { return rowEmpty(sa, ua, q.Old) && rowEmpty(sb, ub, q.New) })
	}
	p.Rows = r.pairSpilled(sa, sb, p.Rows)
	if len(p.Blocks) == 0 && len(j.colsAdded)+len(j.colsRemoved) == 0 {
		if p.Rows = r.pasted(j, p.Rows); len(j.pasted) > 0 {
			p.Moved = nil // the rows cut and pasted, not sorted
		}
	}
	if len(j.pasted) == 0 && len(p.Moved) == 0 && r.cutCells(j) {
		return j // its rows and columns pair as numbered: nothing added or removed
	}
	j.rows, j.moved = p.Rows, p.Moved
	for _, b := range p.Blocks {
		bj := blockJob{}
		for _, c := range b.Cols {
			bj.cols = append(bj.cols, [2]int{c.Old, c.New})
			if bj.c1 == 0 || c.New < bj.c1 {
				bj.c1 = c.New
			}
			bj.c2 = max(bj.c2, c.New)
		}
		bj.rows = r.pairAlike(sa, sb, pairReplaced(b.Rows), bj.cols)
		bj.moved = diff.MovedRows(bj.rows, func(q diff.Pair) bool {
			return rowBlank(sa, q.Old, bj.cols, 0) && rowBlank(sb, q.New, bj.cols, 1)
		})
		j.blocks = append(j.blocks, bj)
	}
	for _, q := range j.rows {
		switch {
		case q.Old == 0:
			j.rowsAdded = append(j.rowsAdded, q.New)
		case q.New == 0:
			j.rowsRemoved = append(j.rowsRemoved, q.Old)
		}
	}
	return j
}

// edit gives the lines a job added and removed, with the span of the rows
// a sort moved.
func (j *job) edit() lineEdit {
	var spanOld, spanNew [2]int
	for _, m := range j.moved {
		spanOld, spanNew = widen(spanOld, m.From, m.From+m.Rows-1), widen(spanNew, m.To, m.To+m.Rows-1)
	}
	for _, b := range j.blocks {
		for _, m := range b.moved {
			spanOld, spanNew = widen(spanOld, m.From, m.From+m.Rows-1), widen(spanNew, m.To, m.To+m.Rows-1)
		}
	}
	spanOld, spanNew = spanWith(spanOld, j.rowsRemoved), spanWith(spanNew, j.rowsAdded)
	return lineEdit{sheet: j.sb.Name, rowsRemoved: runs(j.rowsRemoved), rowsAdded: runs(j.rowsAdded),
		colsRemoved: runs(j.colsRemoved), colsAdded: runs(j.colsAdded), movedOld: spanOld, movedNew: spanNew, cut: j.cut}
}

// movedRange names rows that moved: whole rows, or the columns c1..c2 of
// them (c1 0 for whole rows).
func movedRange(sheet string, from, to, c1, c2 int) string {
	if c1 == 0 {
		return lineRange(sheet, from, to, false)
	}
	return grid.QuoteSheet(sheet) + "!" + grid.Range{R1: from, C1: c1, R2: to, C2: c2}.String()
}

// record reports a job's structure (rows and columns added, removed or
// moved) and adds it to the causes, giving what it removed.
func (r *report) record(j *job) []removal {
	A, out := r.A, r.out
	sa, sb := j.sa, j.sb
	var removals []removal
	removalOf := func(removed []int, byCol bool) removal {
		g := removal{sheet: sa.Name, cols: byCol, runs: runs(removed)}
		for _, run := range g.runs {
			g.labels = append(g.labels, lineRange(sa.Name, run[0], run[1], byCol)+" in "+A.path)
		}
		return g
	}
	for _, run := range runs(j.rowsAdded) {
		w := lineRange(sb.Name, run[0], run[1], false)
		out.Structure = append(out.Structure, structureEntry{map[bool]string{true: "row added", false: "rows added"}[run[0] == run[1]], w})
		r.structural(sb.Name, grid.Range{R1: run[0], C1: 1, R2: run[1], C2: grid.MaxCols}, "rows added: "+w)
	}
	for _, run := range runs(j.rowsRemoved) {
		out.Structure = append(out.Structure, structureEntry{map[bool]string{true: "row removed", false: "rows removed"}[run[0] == run[1]], lineRange(sa.Name, run[0], run[1], false) + " in " + A.path})
	}
	for _, run := range runs(j.colsAdded) {
		w := lineRange(sb.Name, run[0], run[1], true)
		out.Structure = append(out.Structure, structureEntry{map[bool]string{true: "column added", false: "columns added"}[run[0] == run[1]], w})
		r.structural(sb.Name, grid.Range{R1: 1, C1: run[0], R2: grid.MaxRows, C2: run[1]}, "columns added: "+w)
	}
	for _, run := range runs(j.colsRemoved) {
		out.Structure = append(out.Structure, structureEntry{map[bool]string{true: "column removed", false: "columns removed"}[run[0] == run[1]], lineRange(sa.Name, run[0], run[1], true) + " in " + A.path})
	}
	if len(j.rowsRemoved) > 0 {
		removals = append(removals, removalOf(j.rowsRemoved, false))
	}
	if len(j.colsRemoved) > 0 {
		removals = append(removals, removalOf(j.colsRemoved, true))
	}
	// Rows moved, whole or in a block of columns, as causes: what reads
	// them, and what a sort spans (every cell there may hold another row's
	// now).
	moves := func(ms []diff.Moved, c1, c2, dc int) {
		lo, hi := 0, 0
		cols, wasCols := "", ""
		if c1 > 0 {
			cols = grid.ColName(c1) + ":" + grid.ColName(c2)
			if dc != 0 {
				wasCols = grid.ColName(c1-dc) + ":" + grid.ColName(c2-dc)
			}
		}
		for _, m := range ms {
			out.MovedRows = append(out.MovedRows, movedRows{Sheet: sb.Name, From: m.From, To: m.To, Rows: m.Rows, Columns: cols, WasColumns: wasCols,
				c1: c1, c2: c2, wasC1: c1 - dc, wasC2: c2 - dc})
			was := ""
			if m.From != m.To || dc != 0 {
				was = " (were " + movedRange(sa.Name, m.From, m.From+m.Rows-1, c1-dc, c2-dc) + ")"
			}
			r.structural(sb.Name, grid.Range{R1: m.To, C1: max(c1, 1), R2: m.To + m.Rows - 1, C2: map[bool]int{true: grid.MaxCols, false: c2}[c1 == 0]},
				"rows moved: "+movedRange(sb.Name, m.To, m.To+m.Rows-1, c1, c2)+was)
			if lo == 0 || min(m.From, m.To) < lo {
				lo = min(m.From, m.To)
			}
			hi = max(hi, m.From+m.Rows-1, m.To+m.Rows-1)
		}
		if lo > 0 {
			r.structural(sb.Name, grid.Range{R1: lo, C1: max(min(c1, c1-dc), 1), R2: hi, C2: map[bool]int{true: grid.MaxCols, false: max(c2, c2-dc)}[c1 == 0]},
				"rows moved: "+movedRange(sb.Name, lo, hi, min(c1, c1-dc), max(c2, c2-dc)))
		}
	}
	moves(j.moved, 0, 0, 0)
	for _, b := range j.blocks {
		moves(b.moved, b.c1, b.c2, 0)
	}
	moves(j.pasted, j.pasteC1, j.pasteC2, j.pasteDC)
	return removals
}

// readsOthers tells whether a sheet's formulas name one of the sheets
// (lower-cased) other than its own.
func readsOthers(s *grid.Sheet, sheets map[string]bool) bool {
	found := false
	s.Each(s.Used(), func(_, _ int, c *grid.Cell) {
		if found || c.Formula == "" {
			return
		}
		lf := strings.ToLower(c.Formula)
		for name := range sheets {
			if name != strings.ToLower(s.Name) && (strings.Contains(lf, name) || strings.Contains(lf, strings.ReplaceAll(name, "'", "''"))) {
				found = true
				return
			}
		}
	})
	return found
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
	opts := diff.Options{Moves: true, LostRefs: true, ColumnShifts: true, Dates: true, Spills: true}
	al := diff.Align(A.wb, B.wb, opts)
	r.renamed = map[string]string{}
	r.formulaSrc = map[xlsx.At]string{}
	var tr *xlsx.Tracer
	if !out.ValuesOnly {
		tr = xlsx.NewTracer(B.wb)
	}
	jobs := make([]*job, len(al.Sheets)) // nil for a sheet added or removed
	for i, p := range al.Sheets {
		switch {
		case p.Old == nil:
		case p.New == nil:
			r.renamed[strings.ToLower(p.Old.Name)] = ""
			r.removedSheets = append(r.removedSheets, p.Old.Name)
		default:
			r.renamed[strings.ToLower(p.Old.Name)] = p.New.Name
			jobs[i] = r.prepare(p)
		}
	}
	for from, to := range r.renamed {
		if to != "" && !strings.EqualFold(from, to) {
			r.renames = append(r.renames, [2]string{from, to})
		}
	}
	sort.Slice(r.renames, func(i, j int) bool { return r.renames[i][0] < r.renames[j][0] })
	if !out.ValuesOnly {
		// An Excel table renamed: its structured references are rewritten.
		r.tables = tableRenames(r.fa, r.fb, r.renamed)
		for from, to := range r.tables {
			r.tableRenames = append(r.tableRenames, [2]string{from, to})
		}
		sort.Slice(r.tableRenames, func(i, j int) bool { return r.tableRenames[i][0] < r.tableRenames[j][0] })
		for _, rn := range r.tableRenames {
			from := rn[0]
			for _, f := range r.fa {
				if f.Kind == xlsx.FeatureTable && strings.ToLower(f.What) == rn[0] {
					from = f.What
				}
			}
			out.Structure = append(out.Structure, structureEntry{"table renamed", from + " → " + rn[1]})
		}
	}
	edits := func() []lineEdit {
		var es []lineEdit
		for _, j := range jobs {
			if j != nil {
				es = append(es, j.edit())
			}
		}
		return es
	}
	r.setEdits(edits())
	if !out.ValuesOnly && (r.anyLines || r.anyCut || len(r.renames) > 0 || len(r.tables) > 0) {
		// A sheet whose formulas read another that lines were added to or
		// removed from, or that was renamed: its rows are keyed by their
		// formulas, which Excel rewrote, so they pair again with the old
		// formulas as Excel would have rewritten them. Kept when that leaves
		// fewer lines added and removed.
		touched := map[string]bool{}
		for _, j := range jobs {
			if j != nil && (j.lineCount() > 0 || len(j.cut) > 0 || !strings.EqualFold(j.sa.Name, j.sb.Name)) {
				touched[strings.ToLower(j.sa.Name)] = true
			}
		}
		again := opts
		again.Rewrite = func(sheet, f string) string {
			home := r.renamed[strings.ToLower(sheet)]
			return r.shifted(r.renamedIn(f), home, false, home)
		}
		changed := false
		for i, j := range jobs {
			if j == nil || !readsOthers(j.sa, touched) {
				continue
			}
			if k := r.prepare(diff.AlignSheet(A.wb, j.sa, B.wb, j.sb, again)); k.lineCount() < j.lineCount() {
				jobs[i], changed = k, true
			}
		}
		if changed {
			r.setEdits(edits())
		}
	}
	var removals []removal
	for i, p := range al.Sheets {
		switch {
		case p.Old == nil:
			out.Structure = append(out.Structure, structureEntry{"sheet added", grid.QuoteSheet(p.New.Name)})
			if u := p.New.Used(); !u.Empty() {
				r.structural(p.New.Name, u, "sheet added: "+grid.QuoteSheet(p.New.Name))
			}
		case p.New == nil:
			out.Structure = append(out.Structure, structureEntry{"sheet removed", grid.QuoteSheet(p.Old.Name)})
		default:
			if p.Old.Name != p.New.Name && !out.ValuesOnly {
				out.Structure = append(out.Structure, structureEntry{"sheet renamed", grid.QuoteSheet(p.Old.Name) + " → " + grid.QuoteSheet(p.New.Name)})
			}
			removals = append(removals, r.record(jobs[i])...)
		}
	}
	r.lines = map[string]*lineMap{}
	for _, j := range jobs {
		if j == nil {
			continue
		}
		m := &lineMap{sheet: j.sb.Name, row: map[int]int{}, col: map[int]int{}, moved: map[int]bool{}, pasted: map[int]bool{},
			pasteC1: j.pasteC1, pasteC2: j.pasteC2, pasteDC: j.pasteDC}
		if j.cells != nil {
			m.cells = j.cells.src
		}
		for _, mv := range j.moved {
			for k := 0; k < mv.Rows; k++ {
				m.moved[mv.From+k] = true
			}
		}
		for _, mv := range j.pasted {
			for k := 0; k < mv.Rows; k++ {
				m.pasted[mv.From+k] = true
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
		for _, b := range j.blocks {
			bm := blockMap{cols: map[int]bool{}, row: map[int]int{}, moved: map[int]bool{}, c1: b.c1, c2: b.c2}
			for _, c := range b.cols {
				bm.cols[c[0]] = true
				m.col[c[0]] = c[1]
			}
			for _, q := range b.rows {
				if q.Old > 0 && q.New > 0 {
					bm.row[q.Old] = q.New
				}
			}
			for _, mv := range b.moved {
				for k := 0; k < mv.Rows; k++ {
					bm.moved[mv.From+k] = true
				}
			}
			m.blocks = append(m.blocks, bm)
		}
		r.lines[strings.ToLower(j.sa.Name)] = m
	}
	for _, j := range jobs {
		if j == nil {
			continue
		}
		r.cur = r.lines[strings.ToLower(j.sa.Name)]
		first := len(r.candidates)
		var res []resultPair
		var oldIn, newIn map[[2]int]bool
		if !out.ValuesOnly {
			res, oldIn, newIn = r.results(j, r.cur)
		}
		// pair compares two cells paired by their rows, unless the result of
		// an array formula holds either and its formula's cell compares it.
		pair := func(ra, ca, rb, cb int) {
			if c := j.cells; c != nil && (c.src.Contains(ra, ca) || c.dst.Contains(rb, cb)) {
				return // cells cut and pasted: they compare where they went
			}
			if a, b := oldIn[[2]int{ra, ca}], newIn[[2]int{rb, cb}]; a || b {
				if a && b || a && empty(j.sb, rb, cb) || b && empty(j.sa, ra, ca) {
					return
				}
			}
			r.cell(j.sa, ra, ca, j.sb, rb, cb)
		}
		for _, q := range j.rows {
			if q.Old == 0 || q.New == 0 {
				continue
			}
			for _, c := range j.cols {
				pair(q.Old, c[0], q.New, c[1])
			}
		}
		for _, b := range j.blocks {
			for _, q := range b.rows {
				if q.Old == 0 || q.New == 0 {
					continue
				}
				for _, c := range b.cols {
					pair(q.Old, c[0], q.New, c[1])
				}
			}
		}
		if c := j.cells; c != nil {
			for row := c.src.R1; row <= c.src.R2; row++ {
				for col := c.src.C1; col <= c.src.C2; col++ {
					r.cell(j.sa, row, col, j.sb, row+c.dr, col+c.dc)
				}
			}
		}
		for _, p := range res {
			for dr := 0; dr < p.rows; dr++ {
				for dc := 0; dc < p.cols; dc++ {
					if dr > 0 || dc > 0 {
						r.cell(j.sa, p.ra+dr, p.ca+dc, j.sb, p.rb+dr, p.cb+dc)
					}
				}
			}
		}
		if len(res) > 0 {
			// The results' cells among the rest, row by row.
			cs := r.candidates[first:]
			sort.SliceStable(cs, func(x, y int) bool {
				if cs[x].at.Row != cs[y].at.Row {
					return cs[x].at.Row < cs[y].at.Row
				}
				return cs[x].at.Col < cs[y].at.Col
			})
		}
	}
	if out.ValuesOnly {
		return
	}
	// The sources so far are cells changed: inputs and formulas.
	r.cells = len(r.sources)
	// Names added, changed or removed: the formulas using them follow.
	old := map[string]grid.DefinedName{}
	for _, n := range A.wb.Names {
		old[strings.ToLower(r.renamed[strings.ToLower(n.Scope)]+"!"+n.Name)] = n
	}
	seen := map[string]bool{}
	using := func(name, label string) {
		for _, at := range tr.Using(name) {
			if _, seen := r.formulaSrc[at]; !seen {
				r.formulaSrc[at] = label
			}
		}
	}
	for _, n := range B.wb.Names {
		key := strings.ToLower(n.Scope + "!" + n.Name)
		seen[key] = true
		o, ok := old[key]
		switch {
		case !ok:
			out.Structure = append(out.Structure, structureEntry{"name added", n.Name + " = " + n.Ref})
			using(n.Name, "name "+n.Name+" added")
		case !r.sameRef(o.Ref, n.Ref):
			out.Structure = append(out.Structure, structureEntry{"name changed", n.Name + ": " + o.Ref + " → " + n.Ref})
			using(n.Name, "name "+n.Name+" changed")
		}
	}
	for _, n := range A.wb.Names {
		if !seen[strings.ToLower(r.renamed[strings.ToLower(n.Scope)]+"!"+n.Name)] {
			// Excel leaves the formulas using it as they were, now #NAME?.
			out.Structure = append(out.Structure, structureEntry{"name removed", n.Name + " = " + n.Ref})
			using(n.Name, "name "+n.Name+" removed")
		}
	}
	// Formulas that read lines removed, or a sheet removed (in the old
	// version): what they read changed, whether Excel shrank a reference
	// or, with nothing left to read, wrote #REF!.
	var ta *xlsx.Tracer
	touching := func(sheet string, runs [][2]int, cols bool, label func(k int) string) {
		if ta == nil {
			ta = xlsx.NewTracer(A.wb)
		}
		for at, k := range ta.Touching(sheet, runs, cols) {
			m := r.lines[strings.ToLower(at.Sheet)]
			if m == nil || m.rowOf(at.Row, at.Col) == 0 || m.col[at.Col] == 0 {
				continue // its sheet, row or column was removed
			}
			here := xlsx.At{Sheet: m.sheet, Row: m.rowOf(at.Row, at.Col), Col: m.col[at.Col]}
			if _, seen := r.formulaSrc[here]; !seen && tr.Formula(here.Sheet, here.Row, here.Col) {
				r.formulaSrc[here] = label(k)
			}
		}
	}
	for _, g := range removals {
		touching(g.sheet, g.runs, g.cols, func(k int) string {
			return map[bool]string{true: "columns removed: ", false: "rows removed: "}[g.cols] + g.labels[k]
		})
	}
	for _, name := range r.removedSheets {
		touching(name, [][2]int{{1, grid.MaxRows}}, false, func(int) string { return "sheet removed: " + grid.QuoteSheet(name) })
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

// spanWith takes in the rows added or removed (lines) that join a span of
// moved rows without a row kept between: the sort may have taken them too.
func spanWith(span [2]int, lines []int) [2]int {
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
		if va.known && vb.known && !sameInputs(va, vb) {
			before, after := inputTexts(va, vb)
			out.Changed = append(out.Changed, changedCell{Cell: here, Was: was, Kind: "value", Before: before, After: after})
		}
		return
	}
	fa, fb := !xe && x.Formula != "", !ye && y.Formula != ""
	ka, kb := r.arrays.computed(sa, ra, ca), r.arrays.computed(sb, rb, cb)
	one := grid.Range{R1: rb, C1: cb, R2: rb, C2: cb}
	at := xlsx.At{Sheet: sb.Name, Row: rb, Col: cb}
	note := func() string {
		note := B.stored[strings.ToLower(here)]
		if n := A.stored[strings.ToLower(cellAt(sa.Name, ra, ca))]; n != "" && note == "" {
			note = "in " + A.path + ", " + n
		}
		if note != "" {
			note = "computed by sheets: " + note
		}
		return note
	}
	unknownWhy := func() string {
		if !vb.known {
			return B.unknownWhy(sb.Name, rb, cb)
		}
		return "unknown in " + A.path + ": " + A.unknownWhy(sa.Name, ra, ca)
	}
	// changed records a formula changed, a cause of what follows it, and
	// what it computes now when that is not what the cell held.
	changed := func(c changedCell) {
		out.Changed = append(out.Changed, c)
		r.cause(sb.Name, one, "formula "+here)
		switch {
		case !kb:
			// A value now: the change shows it.
		case !va.known || !vb.known:
			out.Unknown = append(out.Unknown, unknownValue{Cell: here, Was: was, Before: va.text, After: vb.text, Why: unknownWhy()})
		case xe || !sameResults(va, vb):
			r.candidates = append(r.candidates, candidate{at: at, before: va, after: vb, was: was, note: note(), cause: "formula " + here})
		}
	}
	var via *xlsx.At
	var shift *shiftOf
	switch {
	case !ka && !kb:
		// Inputs.
		if (xe != ye) || !sameInputs(va, vb) {
			before, after := inputTexts(va, vb)
			out.Changed = append(out.Changed, changedCell{Cell: here, Was: was, Kind: "input", Before: before, After: after})
			r.cause(sb.Name, one, "input "+here)
		}
		return
	case fa && fb:
		if folded(x.Formula) == folded(y.Formula) {
			// The same text: the formula is as it was, whatever rows were
			// moved around it, unless rows or columns added or removed (or a
			// sheet renamed) moved what it reads and it was not rewritten for
			// them. (A row a sort moved took its formula with it, rewritten as
			// for a copy: where it reads is its own business.)
			if want, ok := r.stale(x.Formula, sb.Name); ok && !r.cur.movedAt(ra, ca) {
				changed(changedCell{Cell: here, Was: was, Kind: "formula", Before: "=" + x.Formula, After: "=" + y.Formula,
					Flag: "its text is as it was, though " + r.staleWhy(x.Formula, sb.Name) + ": Excel writes =" + want})
				return
			}
			if ra != rb || ca != cb {
				shift = &shiftOf{sheet: sa.Name, row: ra, col: ca, old: x.Formula, now: y.Formula}
			}
			break
		}
		pa, _ := formula.Pattern(folded(x.Formula), ra, ca)
		pb, _ := formula.Pattern(folded(y.Formula), rb, cb)
		if pa != pb && !r.adjusted(folded(x.Formula), sb.Name, folded(y.Formula), rb, cb) {
			c := changedCell{Cell: here, Was: was, Kind: "formula", Before: "=" + x.Formula, After: "=" + y.Formula}
			// The rows around it keep what Excel would have written for the
			// old formula, columns inserted or deleted included.
			want := pa
			if f, ok := r.replay(folded(x.Formula), sb.Name, false); ok {
				want, _ = formula.Pattern(folded(f), rb, cb)
			}
			c.Flag = brokenRow(sb, rb, cb, want, pb)
			changed(c)
			return
		}
		// Moved with its row, or rewritten by Excel for what moved.
		if ra != rb || ca != cb || pa != pb {
			shift = &shiftOf{sheet: sa.Name, row: ra, col: ca, old: x.Formula, now: y.Formula}
		}
	case fa != fb && (fa || fb) && !(ka && kb):
		// A formula added or removed.
		c := changedCell{Cell: here, Was: was, Kind: "formula", Before: repr(x, va), After: repr(y, vb)}
		if fa && !kb {
			pa, _ := formula.Pattern(folded(x.Formula), ra, ca)
			c.Flag = brokenRow(sb, rb, cb, pa, "")
		}
		changed(c)
		return
	case ka && !kb && ye:
		// An array's result that shrank: the cell it left moved, and
		// what moved the array moved it.
		if ar, ac, ok := r.arrays.anchor(sa, ra, ca); ok && r.cur.rowOf(ar, ac) > 0 && r.cur.col[ac] > 0 {
			via = &xlsx.At{Sheet: sb.Name, Row: r.cur.rowOf(ar, ac), Col: r.cur.col[ac]}
		}
	case ka && !kb:
		// A value where an array's result was.
		out.Changed = append(out.Changed, changedCell{Cell: here, Was: was, Kind: "input", Before: va.text, After: vb.text})
		r.cause(sb.Name, one, "input "+here)
		return
	}
	// A computed value, its formula unchanged: did it move?
	if va.known && vb.known && sameResults(va, vb) {
		if len(B.differ) > 0 {
			if d, ok := B.differ[strings.ToLower(here)]; ok {
				// It holds what it held, a result the sheets calculator
				// disagrees with: out of date, if what it reads changed.
				r.held = append(r.held, heldValue{at: at, was: was, old: strings.ToLower(cellAt(sa.Name, ra, ca)), value: va, calc: calcValue(B.wb, y, d.Calc)})
			}
		}
		return
	}
	why := ""
	if !va.known || !vb.known {
		why = unknownWhy()
	} else {
		why = r.calculators(sa, ra, ca, sb, rb, cb, va, vb)
	}
	if why != "" {
		out.Unknown = append(out.Unknown, unknownValue{Cell: here, Was: was, Before: va.text, After: vb.text, Why: why})
		r.unknownSrc = append(r.unknownSrc, xlsx.Source{Sheet: sb.Name, Range: one})
		r.unknownWhy = append(r.unknownWhy, here+", whose value sheets cannot know ("+why+")")
		return
	}
	r.candidates = append(r.candidates, candidate{at: at, before: va, after: vb, was: was, note: note(), via: via, shift: shift})
}

// calculators says why a value that moved may have moved only because one
// version holds the result Excel stored and the other the one sheets
// computed, where the sheets calculator disagrees with Excel's (check keeps
// Excel's): "" when that is not so.
func (r *report) calculators(sa *grid.Sheet, ra, ca int, sb *grid.Sheet, rb, cb int, va, vb valueOf) string {
	A, B := r.A, r.B
	old, now := strings.ToLower(cellAt(sa.Name, ra, ca)), strings.ToLower(cellAt(sb.Name, rb, cb))
	if d, ok := A.differ[old]; ok && B.stored[now] != "" && sameResults(calcValue(A.wb, sa.At(ra, ca), d.Calc), vb) {
		return "the difference is the calculator's, not a change: sheets gets " + vb.text + " from what it reads in " + A.path +
			", which stores " + va.text + " (Excel's result, kept), and computed it in " + B.path + " (" + B.stored[now] + ")"
	}
	if d, ok := B.differ[now]; ok && A.stored[old] != "" && sameResults(calcValue(B.wb, sb.At(rb, cb), d.Calc), va) {
		return "the difference is the calculator's, not a change: sheets gets " + va.text + " from what it reads in " + B.path +
			", which stores " + vb.text + " (Excel's result, kept), and computed it in " + A.path + " (" + A.stored[old] + ")"
	}
	return ""
}

// calcValue is what the sheets calculator got for a cell (stored, as the
// file holds it), as value reads a cell.
func calcValue(wb *grid.Workbook, stored *grid.Cell, calc grid.Cell) valueOf {
	if stored != nil {
		calc.Fmt = stored.Fmt
	}
	return valueFrom(wb, &calc)
}

// repr shows a cell: its formula, or its value.
func repr(c *grid.Cell, v valueOf) string {
	if c != nil && c.Formula != "" {
		return "=" + c.Formula
	}
	return v.text
}

// within tells whether a run of lines meets a span (zero for none).
func within(run, span [2]int) bool {
	return span[0] > 0 && run[0] <= span[1] && run[1] >= span[0]
}

// setEdits takes the lines added and removed on each sheet.
func (r *report) setEdits(edits []lineEdit) {
	r.edits, r.anyLines, r.anyCut = edits, false, false
	for _, e := range edits {
		if len(e.rowsRemoved)+len(e.rowsAdded)+len(e.colsRemoved)+len(e.colsAdded) > 0 {
			r.anyLines = true
		}
		r.anyCut = r.anyCut || len(e.cut) > 0
	}
}

// replay rewrites a formula, or a name's reference (home ""), as Excel does
// for the sheets renamed, the rows and columns added and removed, and the
// cells cut and pasted. It says false when there is nothing to replay. With
// sure set, the rows added or removed within what a sort moved are left
// out: where they were cannot be told, as the sort took the rows elsewhere.
func (r *report) replay(old, home string, sure bool) (string, bool) {
	if !r.anyLines && !r.anyCut && len(r.renames) == 0 && len(r.tables) == 0 {
		return "", false
	}
	return r.shifted(r.renamedIn(old), home, sure, ""), true
}

// renamedIn rewrites a formula for the sheets and Excel tables renamed.
func (r *report) renamedIn(f string) string {
	for _, rn := range r.renames {
		f = formula.RenameSheet(f, rn[0], rn[1])
	}
	for _, rn := range r.tableRenames {
		f = renameTable(f, rn[0], rn[1])
	}
	return f
}

// shiftRef gives a reference of the old version as the rows and columns added
// and removed move it (lostFeatures): "" when they removed it.
func (r *report) shiftRef(home, ref string) string {
	if !r.anyLines {
		return ref
	}
	g := r.lined(r.renamedIn(ref), home, false, "")
	if strings.Contains(g, "#REF!") {
		return ""
	}
	return g
}

// shifted rewrites a formula (its sheets' names as now) for the lines added
// and removed, then the cells cut and pasted, on every sheet but skip (""
// for none).
func (r *report) shifted(f, home string, sure bool, skip string) string {
	return r.cutIn(r.lined(f, home, sure, skip), home, skip)
}

// lined rewrites a formula for the lines added and removed, as shifted.
func (r *report) lined(f, home string, sure bool, skip string) string {
	for _, e := range r.edits {
		if skip != "" && strings.EqualFold(e.sheet, skip) {
			continue
		}
		removed, added := e.rowsRemoved, e.rowsAdded
		if sure {
			removed, added = outside(removed, e.movedOld), outside(added, e.movedNew)
		}
		f = shiftLines(f, home, e.sheet, false, removed, added)
		f = shiftLines(f, home, e.sheet, true, e.colsRemoved, e.colsAdded)
	}
	return f
}

// cutIn rewrites a formula for the cells cut and pasted, as shifted.
func (r *report) cutIn(f, home, skip string) string {
	if !r.anyCut {
		return f
	}
	for _, e := range r.edits {
		if skip != "" && strings.EqualFold(e.sheet, skip) {
			continue
		}
		for _, m := range e.cut {
			f = formula.MoveRefs(f, home, home, m)
		}
	}
	return f
}

// adjusted tells whether a formula (now in row, col of sheet home) is its
// old self rewritten as Excel does for the sheets renamed and removed and
// the rows and columns added and removed.
func (r *report) adjusted(old, home, now string, row, col int) bool {
	pn, _ := formula.Pattern(now, row, col)
	for _, f := range r.rewrites(old, home) {
		// A sheet's new name goes in as typed: fold it, as old and now are.
		if pf, _ := formula.Pattern(folded(f), row, col); pf == pn {
			return true
		}
	}
	return false
}

// removedMark stands for a sheet removed while its references are rewritten
// as Excel writes them.
const removedMark = "SHEETSREMOVEDMARK"

// rewrites are what Excel, or another tool, may have written for a formula
// for the sheets renamed and removed and the lines added and removed:
// replayed, and with what it read of a sheet removed as #REF! (Excel
// writes =SUM(#REF!B2:B6), sheets =SUM(#REF!)).
func (r *report) rewrites(old, home string) []string {
	f, ok := r.replay(old, home, false)
	if !ok {
		if len(r.removedSheets) == 0 {
			return nil
		}
		f = old
	}
	out := []string{f}
	lost, excel := f, f
	for i, name := range r.removedSheets {
		if !mentions(f, "", name) {
			continue
		}
		if r.removedRefs == nil {
			for _, n := range r.removedSheets {
				r.removedRefs = append(r.removedRefs, regexp.MustCompile(`(?i)(`+regexp.QuoteMeta(n)+`|'`+regexp.QuoteMeta(strings.ReplaceAll(n, "'", "''"))+`')!#REF!`))
			}
		}
		// DeleteLines leaves the sheet's name before #REF!; sheets writes none.
		lost = formula.DeleteLines(formula.DeleteLines(lost, home, name, false, 1, grid.MaxRows), home, name, true, 1, grid.MaxCols)
		lost = r.removedRefs[i].ReplaceAllString(lost, "#REF!")
		excel = strings.ReplaceAll(formula.RenameSheet(excel, strings.ToLower(name), removedMark), removedMark+"!", "#REF!")
	}
	if lost != f {
		out = append(out, lost, excel)
	}
	return out
}

// stale gives what Excel would have written for a formula whose text is as
// it was, when rows or columns added or removed moved what it reads: Excel
// rewrites such a formula, so one left as it was reads other cells now (the
// usual slip of a script that inserts rows without rewriting formulas).
func (r *report) stale(old, home string) (string, bool) {
	f, ok := r.replay(old, home, true)
	return f, ok && folded(f) != folded(old)
}

// staleWhy says what moved what a stale formula reads: lines added or
// removed, a sheet renamed, or both.
func (r *report) staleWhy(old, home string) string {
	named := r.renamedIn(old)
	lined := r.lined(named, home, true, "")
	var why []string
	if lines := folded(lined) != folded(named); lines {
		why = append(why, "rows or columns added or removed moved what it reads")
	}
	if folded(r.cutIn(lined, home, "")) != folded(lined) {
		why = append(why, "the cells it reads were cut and pasted elsewhere")
	}
	if folded(named) != folded(old) {
		why = append(why, "a sheet it reads was renamed")
	}
	if len(why) == 0 {
		return "rows or columns added or removed moved what it reads"
	}
	return strings.Join(why, ", and ")
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
	// A value that cells changed moved is the work of all of them that reach
	// it, not only the nearest.
	var allOf map[xlsx.At]xlsx.Reach
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
	// causeText says what moved a value whose nearest cause is source k: when
	// that is a cell changed, each of those that reach it (the first few).
	causeText := func(c candidate, k int) string {
		if k >= r.cells {
			return r.causes[k]
		}
		if allOf == nil {
			allOf = tr.SpreadAll(r.sources[:r.cells], causeListed)
		}
		var got xlsx.Reach
		for _, at := range places(c) {
			if g, ok := allOf[at]; ok {
				got = g
				break
			}
		}
		if len(got.Sources) < 2 && !got.More {
			return r.causes[k]
		}
		labels := make([]string, len(got.Sources))
		for i, n := range got.Sources {
			labels[i] = r.causes[n]
		}
		if got.More {
			labels = append(labels, "and more")
		}
		return strings.Join(labels, ", ")
	}
	anchorOf := func(c candidate) xlsx.At {
		at := c.at
		if c.via != nil {
			at = *c.via
		} else if ar, ac, ok := r.arrays.anchor(r.B.wb.Sheet(at.Sheet), at.Row, at.Col); ok {
			at = xlsx.At{Sheet: at.Sheet, Row: ar, Col: ac}
		}
		return at
	}
	// What sheets cannot follow (formulas whose reads cannot be told, and
	// those whose values are unknown), and, failing all else, what moved the
	// formulas a sort or lines added or removed moved or Excel rewrote for
	// them.
	var unknownOf, shiftReach map[xlsx.At]int
	var shiftSrc []xlsx.Source
	var shiftCause []string
	for _, c := range r.candidates {
		if _, ok := cause(causeOf, c); ok || c.cause != "" {
			continue
		}
		if unknownOf == nil {
			unknownOf = tr.Spread(r.unknownSources(tr))
		}
		if c.shift != nil && tr.Opaque(c.at.Sheet, c.at.Row, c.at.Col) == "" {
			shiftSrc = append(shiftSrc, xlsx.Source{Sheet: c.at.Sheet, Range: grid.Range{R1: c.at.Row, C1: c.at.Col, R2: c.at.Row, C2: c.at.Col}})
			shiftCause = append(shiftCause, r.shiftCause(c))
		}
	}
	if len(shiftSrc) > 0 {
		shiftReach = tr.Spread(shiftSrc)
	}
	for _, c := range r.candidates {
		here := cellAt(c.at.Sheet, c.at.Row, c.at.Col)
		moved := movedValue{Cell: here, Was: c.was, Before: c.before.text, After: c.after.text, Note: c.note}
		if c.cause != "" {
			moved.Cause = c.cause
			out.Moved = append(out.Moved, moved)
			continue
		}
		if k, ok := cause(causeOf, c); ok {
			moved.Cause = causeText(c, k)
			out.Moved = append(out.Moved, moved)
			continue
		}
		at := anchorOf(c)
		if why := tr.Opaque(at.Sheet, at.Row, at.Col); why != "" {
			out.Unknown = append(out.Unknown, unknownValue{Cell: here, Was: c.was, Before: c.before.text, After: c.after.text, Why: "moved, and sheets cannot tell why: " + why})
			continue
		}
		if k, ok := cause(unknownOf, c); ok {
			out.Unknown = append(out.Unknown, unknownValue{Cell: here, Was: c.was, Before: c.before.text, After: c.after.text, Why: "moved, and sheets cannot tell why: it reads " + r.unknownWhy[k]})
			continue
		}
		if k, ok := cause(shiftReach, c); ok {
			moved.Cause = shiftCause[k]
			out.Moved = append(out.Moved, moved)
			continue
		}
		out.Unexplained = append(out.Unexplained, moved)
	}
	// Values that held though what they read changed, whose stored results
	// the sheets calculator gets otherwise from what they read now: results
	// left as they were (written without recalculating, or by Excel
	// calculating by hand), unless the calculator disagreed as much before.
	for _, h := range r.held {
		k, ok := causeOf[h.at]
		if !ok {
			continue
		}
		if d, ok := r.A.differ[h.old]; ok && sameResults(calcValue(r.A.wb, nil, d.Calc), h.calc) {
			continue
		}
		here := cellAt(h.at.Sheet, h.at.Row, h.at.Col)
		out.Unknown = append(out.Unknown, unknownValue{Cell: here, Was: h.was, Before: h.value.text, After: h.value.text,
			Why: "its result may be out of date (written without recalculating): it holds what it held, though what it reads changed (" +
				r.causes[k] + "), and the sheets calculator gets " + h.calc.text + " from it now; Excel's result is kept"})
	}
}

// shiftCause names what moved a formula a sort or lines added or removed
// moved, or whose references Excel rewrote for them.
func (r *report) shiftCause(c candidate) string {
	s := c.shift
	if m := r.lines[strings.ToLower(s.sheet)]; m != nil && (m.movedAt(s.row, s.col) || m.pastedAt(s.row, s.col)) {
		c1, c2, dc := 0, 0, 0
		if b := m.block(s.col); b != nil {
			c1, c2 = b.c1, b.c2
		} else if m.pastedAt(s.row, s.col) {
			c1, c2, dc = m.pasteC1, m.pasteC2, m.pasteDC
		}
		was := ""
		if s.row != c.at.Row || dc != 0 {
			was = " (were " + movedRange(s.sheet, s.row, s.row, c1-dc, c2-dc) + ")"
		}
		return "rows moved: " + movedRange(c.at.Sheet, c.at.Row, c.at.Row, c1, c2) + was
	}
	if l := r.lineAbove(c.at.Sheet, s.row, c.at.Row, s.col, c.at.Col); l != "" {
		return l
	}
	// What it reads, on each sheet, as far as it reads, before and now.
	far := func(f, home string) map[string][2]int {
		out := map[string][2]int{}
		for _, ref := range formula.Refs(f) {
			sheet := ref.Sheet
			if sheet == "" {
				sheet = home
			}
			k := strings.ToLower(sheet)
			x := out[k]
			out[k] = [2]int{max(x[0], ref.Range.R2), max(x[1], ref.Range.C2)}
		}
		return out
	}
	before, now := far(r.renamedIn(s.old), c.at.Sheet), far(s.now, c.at.Sheet)
	for _, e := range r.edits {
		k := strings.ToLower(e.sheet)
		if b, n := before[k], now[k]; b != [2]int{} || n != [2]int{} {
			if l := r.lineAbove(e.sheet, b[0], n[0], b[1], n[1]); l != "" {
				return l
			}
		}
	}
	return "rows or columns added or removed"
}

// lineAbove names the first lines added or removed on a sheet (as now
// named) at or before a row or column (numbered as before and now): what
// moved what is there.
func (r *report) lineAbove(sheet string, oldRow, newRow, oldCol, newCol int) string {
	old := sheet
	for from, to := range r.renamed {
		if strings.EqualFold(to, sheet) {
			for _, s := range r.A.wb.Sheets {
				if strings.ToLower(s.Name) == from {
					old = s.Name
				}
			}
		}
	}
	for _, e := range r.edits {
		if !strings.EqualFold(e.sheet, sheet) {
			continue
		}
		if len(e.rowsRemoved) > 0 && e.rowsRemoved[0][0] <= oldRow {
			return "rows removed: " + lineRange(old, e.rowsRemoved[0][0], e.rowsRemoved[0][1], false) + " in " + r.A.path
		}
		if len(e.rowsAdded) > 0 && e.rowsAdded[0][0] <= newRow {
			return "rows added: " + lineRange(sheet, e.rowsAdded[0][0], e.rowsAdded[0][1], false)
		}
		if len(e.colsRemoved) > 0 && e.colsRemoved[0][0] <= oldCol {
			return "columns removed: " + lineRange(old, e.colsRemoved[0][0], e.colsRemoved[0][1], true) + " in " + r.A.path
		}
		if len(e.colsAdded) > 0 && e.colsAdded[0][0] <= newCol {
			return "columns added: " + lineRange(sheet, e.colsAdded[0][0], e.colsAdded[0][1], true)
		}
	}
	return ""
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
				r.unknownWhy = append(r.unknownWhy, cellAt(s.Name, row, col)+", which sheets cannot follow ("+why+")")
			}
		})
	}
	return srcs
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
			tw.row("  "+withWas(m.Cell, m.Was), change(m.Before, m.After), clipNote(m.Note))
		}
		tw.Flush()
		more(len(r.Unexplained))
	}
	if section("features lost", len(r.FeaturesLost)) {
		tw := a.table()
		for _, f := range r.FeaturesLost[:min(len(r.FeaturesLost), againstListed)] {
			tw.row("  "+f.Kind, clip(f.Where), clipNote(f.What))
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
			a.printf("  %s  %s  (%s)\n", withWas(c.Cell, c.Was), change(c.Before, c.After), clipNote(c.Flag))
		}
		more(len(flagged))
	}
	if section("changed", len(plain)) {
		tw := a.table()
		for _, c := range plain[:min(len(plain), againstListed)] {
			tw.row("  "+withWas(c.Cell, c.Was), c.Kind, change(c.Before, c.After))
		}
		tw.Flush()
		more(len(plain))
	}
	if section("moved, each with its cause", len(r.Moved)) {
		tw := a.table()
		for _, m := range r.Moved[:min(len(r.Moved), againstListed)] {
			tw.row("  "+withWas(m.Cell, m.Was), change(m.Before, m.After), clipNote(m.Cause))
		}
		tw.Flush()
		more(len(r.Moved))
	}
	if section("unknown: values sheets cannot compare", len(r.Unknown)) {
		tw := a.table()
		for _, u := range r.Unknown[:min(len(r.Unknown), againstListed)] {
			tw.row("  "+u.Cell, change(u.Before, u.After), clipNote(u.Why))
		}
		tw.Flush()
		more(len(r.Unknown))
	}
	if section("rows moved (as by a sort; not changes)", len(r.MovedRows)) {
		for _, m := range r.MovedRows[:min(len(r.MovedRows), againstListed)] {
			a.printf("  %s, were %s%s\n", movedRange(m.Sheet, m.To, m.To+m.Rows-1, m.c1, m.c2), strings.TrimPrefix(movedRange(m.Sheet, m.From, m.From+m.Rows-1, m.wasC1, m.wasC2), grid.QuoteSheet(m.Sheet)+"!"), heldNumber(m))
		}
		more(len(r.MovedRows))
	}
	if section("structure", len(r.Structure)) {
		for _, e := range r.Structure[:min(len(r.Structure), againstListed)] {
			a.printf("  %s: %s\n", e.Change, clipNote(e.Where))
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
			tw.row(f.Severity.String(), clip(findingCells(f)), grid.Truncate(f.Message, againstFinding))
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

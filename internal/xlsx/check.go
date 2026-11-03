package xlsx

import (
	"errors"
	"fmt"
	"math"
	"strings"
	"time"

	"github.com/eemax/sheets/internal/engine"
	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/input"
)

// Checking a workbook's formula results, whatever wrote it (sheets check),
// and fixing what can be fixed (--fix): every result sheets can compute is
// stored, and functions stored without the prefix Excel needs get it.
//
// The fix is worked out as an edit like any other (the streaming editor
// for as long as it can), then the file it gives is compared with the
// file as it was, cell by cell. That comparison is what is reported, and
// it holds the fix to its promises: a result Excel stored and still
// vouches for never changes, and nothing but formula results and
// prefixes does. A fix that would break them is not saved.

// CheckOptions set a check up.
type CheckOptions struct {
	// Budget bounds the time spent computing results (0: the editors'
	// own), and Compare the time spent comparing Excel's results with the
	// sheets calculator's.
	Budget, Compare time.Duration
}

// CheckResult is what a check finds, and what fixing does.
type CheckResult struct {
	// Formulas counts formula cells, by the state of their stored result:
	// vouched for (Excel stored it, or a sheets stamp vouches for it),
	// missing, or stale (the file asks Excel to recalculate on open, and
	// nothing vouches for the results it holds). Unprefixed counts those
	// calling functions without their prefix (missing or stale too).
	Formulas, Vouched, Missing, Stale, Unprefixed int
	// Changes are what fixing changes, cell by cell. Confirmed counts
	// stale results sheets computes to the same value (no change).
	Changes   []Change
	Confirmed int
	// Unknown are the formulas left without a result sheets can know.
	Unknown []Unknown
	// Withheld says why prefixes were not added ("" when they were, or
	// none were missing); Skipped lists formulas left unprefixed, and why.
	Withheld string
	Skipped  []Unknown
	// The comparison of results Excel stored with the sheets calculator's:
	// formulas compared, agreeing, and the differences; those the
	// calculator cannot compute, and those left out of time.
	Compared, Agreed, Uncomputable, NotCompared int
	Differ                                      []Difference
	// Data is the fixed file, nil when fixing changes nothing.
	Data []byte
	// Before is the workbook as read; After as fixed (Before when fixing
	// changes nothing).
	Before, After *grid.Workbook
	Warnings      []string
	// FixError says why the fix could not be worked out ("" when it
	// could); nothing comes of it then.
	FixError string
}

// Change is one cell fixing changes.
type Change struct {
	// Kind: "stored" (a result where there was none), "replaced" (a stale
	// result replaced by the one computed), "dropped" (a stale result
	// sheets cannot recompute, removed), "prefixed" (functions given their
	// prefix), "cleared" or "filled" (cells of an array formula's result).
	Kind          string
	Sheet         string
	Row, Col      int
	Before, After grid.Cell // values (Empty: none)
	Functions     []string  // prefixed: the functions
	Why           string    // dropped: why there is no result now
}

// Unknown is a formula whose result sheets cannot know, and why.
type Unknown struct {
	Sheet    string
	Row, Col int
	Why      string
	// Stale: it holds a stale result (in a report; fixing drops it).
	Stale bool
}

// Difference is a result Excel stored that the sheets calculator does not
// agree with.
type Difference struct {
	Sheet       string
	Row, Col    int
	Formula     string
	Excel, Calc grid.Cell
}

// fixEditor is what fixing needs of an editor: the streaming one, handing
// over to the full one (Auto), or the full one.
type fixEditor interface {
	Set(sheet string, row, col int, v input.Value) error
	SetArray(sheet string, rg grid.Range, text string) error
	Render() ([]byte, error)
	Evaluate(sheet string, row, col int) (grid.Cell, string, bool)
	Warnings() []string
	Close() error
}

// CheckResults checks a workbook (data, read from path) and works out what fixing
// it changes; it writes nothing.
func CheckResults(path string, data []byte, opts CheckOptions) (*CheckResult, error) {
	before, err := ReadBytes(data)
	if err != nil {
		return nil, fmt.Errorf("%s: %w", path, err)
	}
	res := &CheckResult{Before: before, After: before}
	for _, s := range before.Sheets {
		s.Each(s.Used(), func(r, c int, cell *grid.Cell) {
			if cell.Formula == "" {
				return
			}
			res.Formulas++
			switch {
			case cell.Unprefixed:
				res.Unprefixed++
				if cell.Kind == grid.Empty {
					res.Missing++
				} else {
					res.Stale++
				}
			case !cell.Pending:
				res.Vouched++
			case cell.Kind == grid.Empty:
				res.Missing++
			default:
				res.Stale++
			}
		})
	}
	// Excel's results against the calculator's, before anything reads
	// the workbook's values (which computes the missing ones into it).
	if res.Vouched > 0 {
		compare(before, res, opts.Compare)
	}
	if res.Missing+res.Stale == 0 {
		return res, nil
	}
	if opts.Budget > 0 {
		defer func(b time.Duration) { recalcBudget = b }(recalcBudget)
		recalcBudget = opts.Budget
	}
	err = fix(path, data, before, res, res.Unprefixed > 0)
	var broke *brokenPromise
	if errors.As(err, &broke) && res.Unprefixed > 0 {
		// Adding prefixes would change what Excel vouches for (results
		// reading the formulas, computed from #NAME?): results only.
		res.Withheld = "adding it would change results Excel stored (" + broke.example + ")"
		res.Changes, res.Unknown, res.Skipped, res.Confirmed, res.Warnings = nil, nil, nil, 0, nil
		err = fix(path, data, before, res, false)
	}
	if err != nil {
		// What was found holds; the fix cannot be made.
		res.Changes, res.Unknown, res.Skipped, res.Confirmed, res.Data, res.After = nil, nil, nil, 0, nil, before
		res.FixError = err.Error()
	}
	return res, nil
}

// brokenPromise is a fix that would change what it must not.
type brokenPromise struct {
	n       int
	example string
}

func (b *brokenPromise) Error() string {
	return fmt.Sprintf("fixing would change %s it must leave as they are (%s); nothing was saved. This is a fault in sheets: please report it",
		countOf(b.n, "cell"), b.example)
}

// fix works the fix out in an editor, and compares the file it gives with
// the one read.
func fix(path string, data []byte, before *grid.Workbook, res *CheckResult, prefix bool) (err error) {
	var ed fixEditor
	if from := StreamFrom(); from >= 0 && int64(len(data)) >= from && !needsFull(before) {
		ed, err = OpenAuto(path, data)
	} else {
		ed, err = OpenEditorData(path, data)
	}
	if err != nil {
		return err
	}
	defer ed.Close()
	if prefix {
		for _, s := range before.Sheets {
			s.Each(s.Used(), func(r, c int, cell *grid.Cell) {
				if err != nil || !cell.Unprefixed {
					return
				}
				if rg, ok := legacyArray(s, r, c); ok {
					// Entered as an array formula (Ctrl+Shift+Enter): entered
					// again as one, over the same range.
					if e := ed.SetArray(s.Name, rg, cell.Formula); e != nil {
						res.Skipped = append(res.Skipped, Unknown{Sheet: s.Name, Row: r, Col: c, Why: e.Error()})
					}
					return
				}
				if inArray(s, r, c) {
					res.Skipped = append(res.Skipped, Unknown{Sheet: s.Name, Row: r, Col: c,
						Why: "it lies in a dynamic array's result: enter the formula again with sheets set"})
					return
				}
				if e := ed.Set(s.Name, r, c, input.Value{Formula: cell.Formula}); e != nil {
					res.Skipped = append(res.Skipped, Unknown{Sheet: s.Name, Row: r, Col: c, Why: e.Error()})
				}
			})
		}
		if err != nil {
			return err
		}
	}
	out, err := ed.Render()
	if err != nil {
		return err
	}
	after, err := ReadBytes(out)
	if err != nil {
		return fmt.Errorf("the fixed file cannot be read back (%v); nothing was saved. This is a fault in sheets: please report it", err)
	}
	if err := compareFiles(before, after, ed, res); err != nil {
		return err
	}
	if len(res.Changes) > 0 {
		res.Data, res.After = out, after
		res.Warnings = ed.Warnings()
	}
	return nil
}

// needsFull tells whether fixing needs the full editor, which computes
// some formulas the streaming editor cannot: an array over several cells
// (dynamic, or entered with Ctrl+Shift+Enter), whose result the full
// editor lays out over its cells, and INDIRECT or OFFSET, which look at
// cells their arguments do not name.
func needsFull(wb *grid.Workbook) bool {
	for _, s := range wb.Sheets {
		for i, sp := range append(append([]grid.Range{}, s.Spills...), s.Arrays...) {
			if i >= len(s.Spills) && sp.Rows()*sp.Cols() < 2 {
				continue // an array formula in one cell: either editor
			}
			if c := s.At(sp.R1, sp.C1); c != nil && c.Formula != "" && c.Pending {
				return true
			}
		}
	}
	// A formula to compute that looks at cells its arguments do not name
	// (INDIRECT, OFFSET): the streaming editor holds only what is named.
	looks := false
	for _, s := range wb.Sheets {
		s.Each(s.Used(), func(r, c int, cell *grid.Cell) {
			if looks || cell.Formula == "" || !cell.Pending {
				return
			}
			for _, fn := range formula.CalledFunctions(cell.Formula) {
				if unfollowed[fn] {
					looks = true
				}
			}
		})
	}
	return looks
}

// legacyArray finds the legacy array formula (not a dynamic array)
// anchored at a cell.
func legacyArray(s *grid.Sheet, row, col int) (grid.Range, bool) {
	for _, sp := range s.Spills {
		if sp.R1 == row && sp.C1 == col {
			return grid.Range{}, false
		}
	}
	for _, a := range s.Arrays {
		if a.R1 == row && a.C1 == col {
			return a, true
		}
	}
	return grid.Range{}, false
}

// anchorVouched tells whether a cell lies in the range of an array
// formula whose result Excel stored and vouches for.
func anchorVouched(s *grid.Sheet, row, col int) bool {
	for _, a := range append(append([]grid.Range{}, s.Arrays...), s.Spills...) {
		if a.Contains(row, col) {
			if c := s.At(a.R1, a.C1); c != nil && c.Formula != "" && !c.Pending {
				return true
			}
		}
	}
	return false
}

// arrayChange names the change of a cell of an array formula's result.
func arrayChange(b, a grid.Cell) string {
	switch {
	case a.Kind == grid.Empty:
		return "cleared"
	case b.Kind == grid.Empty:
		return "filled"
	}
	return "replaced"
}

// inArray tells whether a cell lies in a legacy array formula's range.
func inArray(s *grid.Sheet, row, col int) bool {
	for _, a := range s.Arrays {
		if a.Contains(row, col) {
			return true
		}
	}
	for _, a := range s.Spills {
		if a.Contains(row, col) {
			return true
		}
	}
	return false
}

// compareFiles lists what changed from before to after, cell by cell, and
// refuses what must not have.
func compareFiles(before, after *grid.Workbook, ed fixEditor, res *CheckResult) error {
	var broken []string
	brk := func(sheet string, r, c int, what string) {
		if len(broken) < 3 {
			broken = append(broken, grid.FormatRef(sheet, grid.Range{R1: r, C1: c, R2: r, C2: c})+" "+what)
		} else {
			broken = append(broken, "")
		}
	}
	if len(before.Sheets) != len(after.Sheets) {
		return fmt.Errorf("fixing would change the workbook's sheets; nothing was saved. This is a fault in sheets: please report it")
	}
	for i, sb := range before.Sheets {
		sa := after.Sheets[i]
		if sa.Name != sb.Name {
			return fmt.Errorf("fixing would rename sheet %q; nothing was saved. This is a fault in sheets: please report it", sb.Name)
		}
		arrays := func(r, c int) bool { return inArray(sb, r, c) || inArray(sa, r, c) }
		sb.Each(sb.Used(), func(r, c int, b *grid.Cell) {
			a := sa.Get(r, c)
			if b.Formula == "" && a.Formula == "" {
				if !sameValue(*b, a) {
					if anchorVouched(sb, r, c) {
						brk(sb.Name, r, c, "array result "+show(valueOnly(*b))+" → "+show(valueOnly(a)))
					} else if arrays(r, c) {
						res.Changes = append(res.Changes, Change{Kind: arrayChange(*b, a), Sheet: sb.Name, Row: r, Col: c, Before: valueOnly(*b), After: valueOnly(a)})
					} else {
						brk(sb.Name, r, c, "value "+show(valueOnly(*b))+" → "+show(valueOnly(a)))
					}
				}
				return
			}
			if b.Formula != a.Formula {
				brk(sb.Name, r, c, "formula ="+b.Formula+" → ="+a.Formula)
				return
			}
			if b.Unprefixed && !a.Unprefixed {
				res.Changes = append(res.Changes, Change{Kind: "prefixed", Sheet: sb.Name, Row: r, Col: c, Functions: formula.Unprefixed(b.Formula)})
			} else if b.Unprefixed != a.Unprefixed {
				brk(sb.Name, r, c, "lost its function prefix")
				return
			}
			bv, av := valueOnly(*b), valueOnly(a)
			hasA := !a.Pending && av.Kind != grid.Empty
			if !b.Pending {
				// Excel's result, which nothing may replace.
				if !sameValue(bv, av) || a.Pending {
					brk(sb.Name, r, c, "result "+show(bv)+" → "+show(av))
				}
				return
			}
			switch {
			case hasA && bv.Kind == grid.Empty:
				res.Changes = append(res.Changes, Change{Kind: "stored", Sheet: sb.Name, Row: r, Col: c, After: av})
			case hasA && sameValue(bv, av):
				res.Confirmed++
			case hasA:
				res.Changes = append(res.Changes, Change{Kind: "replaced", Sheet: sb.Name, Row: r, Col: c, Before: bv, After: av})
			default:
				why := unknownWhy(ed, sb.Name, r, c, a)
				if bv.Kind != grid.Empty {
					res.Changes = append(res.Changes, Change{Kind: "dropped", Sheet: sb.Name, Row: r, Col: c, Before: bv, Why: why})
				}
				res.Unknown = append(res.Unknown, Unknown{Sheet: sb.Name, Row: r, Col: c, Why: why, Stale: bv.Kind != grid.Empty})
			}
		})
		// Cells only the fixed file holds: an array's result spilled.
		sa.Each(sa.Used(), func(r, c int, a *grid.Cell) {
			if b := sb.At(r, c); b != nil && !b.IsEmpty() {
				return
			}
			if a.Formula != "" || !arrays(r, c) || anchorVouched(sb, r, c) {
				brk(sb.Name, r, c, "new content "+show(valueOnly(*a)))
				return
			}
			res.Changes = append(res.Changes, Change{Kind: "filled", Sheet: sb.Name, Row: r, Col: c, After: valueOnly(*a)})
		})
	}
	if len(broken) > 0 {
		return &brokenPromise{n: len(broken), example: broken[0]}
	}
	return nil
}

// unknownWhy is why a formula has no result in the fixed file.
func unknownWhy(ed fixEditor, sheet string, r, c int, a grid.Cell) string {
	if a.Unprefixed {
		return unprefixedNote(a.Formula)
	}
	if formula.DepsOf(a.Formula).External {
		return linkNote
	}
	_, note, _ := ed.Evaluate(sheet, r, c)
	// Why alone, without when Excel gives it a result (see whenExcel).
	for _, when := range []string{"; Excel computes it on open", whenLinked} {
		note = strings.TrimSuffix(note, when)
	}
	if note == "" || note == "Excel computes it on open" {
		note = "the sheets calculator cannot compute it"
	}
	return note
}

// valueOnly is a cell's value, without its formula and marks.
func valueOnly(c grid.Cell) grid.Cell {
	return grid.Cell{Kind: c.Kind, Bool: c.Bool, Num: c.Num, Text: c.Text}
}

// sameValue tells whether two cells hold the same value, exactly.
func sameValue(a, b grid.Cell) bool {
	if a.Kind != b.Kind {
		return false
	}
	switch a.Kind {
	case grid.Number:
		return a.Num == b.Num || math.IsNaN(a.Num) && math.IsNaN(b.Num)
	case grid.Text, grid.Error:
		return a.Text == b.Text
	case grid.Bool:
		return a.Bool == b.Bool
	}
	return true
}

// compare computes the results Excel stored with the sheets calculator,
// one formula at a time: the other formulas read as the results Excel
// stored, so a difference is the formula's own.
func compare(wb *grid.Workbook, res *CheckResult, budget time.Duration) {
	if budget <= 0 {
		budget = 10 * time.Second
	}
	deadline := time.Now().Add(budget)
	rs := wbResolver(wb)
	book := &gridBook{wb: wb, text: func(sheet string, row, col int, f string) string {
		return rs.calcText(&formulaAt{sheet: sheet, row: row, col: col, text: f})
	}}
	eng := engine.New(book)
	eng.Date1904 = wb.Date1904
	vouched := func(c grid.Cell) bool { return c.Formula != "" && !c.Pending && c.Kind != grid.Empty }
	eng.Override = func(sheet string, row, col int) (engine.Value, bool) {
		s := book.sheet(sheet)
		if s == nil {
			return engine.Value{}, false
		}
		if c := s.Get(row, col); vouched(c) {
			return cellToCalc(c), true
		}
		return engine.Value{}, false
	}
	// Formulas whose results Excel did not vouch for, by lower-cased
	// sheet: a formula reading one is not compared, as the difference
	// could be the input's.
	unvouched := map[string][][2]int{}
	for _, s := range wb.Sheets {
		s.Each(s.Used(), func(r, c int, cell *grid.Cell) {
			if cell.Formula != "" && !vouched(*cell) {
				k := strings.ToLower(s.Name)
				unvouched[k] = append(unvouched[k], [2]int{r, c})
			}
		})
	}
	readsUnvouched := func(sheet string, row, col int, text string) bool {
		if len(unvouched) == 0 {
			return false
		}
		refs, ok := rs.ranges(sheet, row, col, valueDeps(text), 0)
		if !ok {
			return true
		}
		for _, ref := range refs {
			for _, at := range unvouched[strings.ToLower(ref.Sheet)] {
				if ref.Range.Contains(at[0], at[1]) {
					return true
				}
			}
		}
		return false
	}
	for _, s := range wb.Sheets {
		s.Each(s.Used(), func(r, col int, cell *grid.Cell) {
			if !vouched(*cell) {
				return
			}
			d := formula.DepsOf(cell.Formula)
			if d.Volatile || d.External {
				return // its result changes, or only Excel can update it
			}
			if time.Now().After(deadline) {
				res.NotCompared++
				return
			}
			if readsUnvouched(s.Name, r, col, book.text(s.Name, r, col, cell.Formula)) {
				res.Uncomputable++
				return
			}
			before := engine.EpochUses()
			v, ok := evalSafely(eng, s.Name, r, col, book.text(s.Name, r, col, cell.Formula))
			if wb.Date1904 && engine.EpochUses() != before {
				ok = false // it works with dates, in the 1904 system
			}
			if _, _, _, array := book.ArrayAt(s.Name, r, col); array {
				eng.Reset() // its array would stand in for the results Excel stored
			}
			if !ok || v.Kind == engine.Error && (v.Str == engine.ErrUnsupported || v.Str == engine.ErrCycle) {
				res.Uncomputable++
				return
			}
			res.Compared++
			got := calcToCell(v)
			if agreesCell(got, valueOnly(*cell)) {
				res.Agreed++
				return
			}
			res.Differ = append(res.Differ, Difference{Sheet: s.Name, Row: r, Col: col, Formula: cell.Formula, Excel: valueOnly(*cell), Calc: got})
		})
	}
}

// evalSafely computes a formula; false if the calculator failed.
func evalSafely(eng *engine.Engine, sheet string, row, col int, text string) (v engine.Value, ok bool) {
	defer func() {
		if recover() != nil {
			ok = false
			eng.Reset()
		}
	}()
	return eng.Eval(sheet, row, col, text), true
}

// SaveData writes a file sheets worked out (a fix), as saving does: at
// once, kept in the undo history.
func SaveData(path string, data []byte) error { return writeAtomic(path, data) }

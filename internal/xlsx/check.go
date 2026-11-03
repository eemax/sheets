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
		return nil, err
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
	if from := StreamFrom(); from >= 0 && int64(len(data)) >= from && !spillsPending(before) {
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
				if inArray(s, r, c) {
					// Entered as an array formula (Ctrl+Shift+Enter), which
					// writing it again would make a plain one.
					res.Skipped = append(res.Skipped, Unknown{Sheet: s.Name, Row: r, Col: c,
						Why: "an array formula (entered with Ctrl+Shift+Enter): enter it again in Excel"})
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

// spillsPending tells whether a dynamic array's result is to be computed:
// the full editor lays its result out over the cells it spills into,
// where the streaming editor would store its first cell alone.
func spillsPending(wb *grid.Workbook) bool {
	for _, s := range wb.Sheets {
		for _, sp := range s.Spills {
			if c := s.At(sp.R1, sp.C1); c != nil && c.Formula != "" && c.Pending {
				return true
			}
		}
	}
	return false
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
					if arrays(r, c) {
						res.Changes = append(res.Changes, Change{Kind: map[bool]string{true: "cleared", false: "filled"}[a.Kind == grid.Empty], Sheet: sb.Name, Row: r, Col: c, Before: valueOnly(*b), After: valueOnly(a)})
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
			if a.Formula != "" || !arrays(r, c) {
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
		return "it links to another workbook, which only Excel can update"
	}
	_, note, _ := ed.Evaluate(sheet, r, c)
	note = strings.TrimSuffix(note, "; Excel computes it on open")
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
			v, ok := evalSafely(eng, s.Name, r, col, book.text(s.Name, r, col, cell.Formula))
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

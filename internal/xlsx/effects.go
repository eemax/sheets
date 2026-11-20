package xlsx

import (
	"math"

	"github.com/eemax/sheets/internal/grid"
)

// What an editing session did, for the short change report every
// changing command prints (see cli/change.go): the cells it wrote, the
// formula results its save recomputed (before and after), and the rows,
// columns and sheets it inserted, deleted or renamed, so a report over
// several sessions (a batch, or the streaming editor handing over to the
// full one) can follow cells from one to the next. All of it is what the
// save worked out anyway: nothing is computed again for the report.

// EffectCell is a cell where it stands after the session.
type EffectCell struct {
	Sheet    string
	Row, Col int
}

// ResultChange is a formula's stored result before and after a session,
// where the formula stands after it.
type ResultChange struct {
	Sheet    string
	Row, Col int
	// Written: the session wrote the formula (its result before is that of
	// whatever the cell held, not known here).
	Written bool
	// Carried: written, and the formula showed an error where it came from
	// (a sort or move): an error already there, not a new one.
	Carried bool
	// Blank: written, over a formula that had no result (a sort or move
	// carrying it, or a rewrite of it): no result now is nothing new.
	Blank bool
	// Reached: the change reached the formula (it reads something the
	// session wrote or moved); not one only computed because it had no
	// result when opened.
	Reached bool
	// Before is the result stored when the session began; Was says it was
	// known (a result stored and current).
	Before grid.Cell
	Was    bool
	// Stale: the formula's result stored when opened was not current (the
	// file marks it for recalculation, as XlsxWriter's do): Before holds that
	// result and Was is false, since it says nothing of what Excel showed.
	Stale bool
	// After is the result stored now; Now says there is one. Why says why
	// there is none.
	After grid.Cell
	Now   bool
	Why   string
}

// LineOp is a change of rows, columns or sheets, in the order made.
type LineOp struct {
	Sheet string
	// Cols: columns, not rows. N > 0 lines were inserted before At; N < 0,
	// -N lines were deleted from At.
	Cols  bool
	At, N int
	// Rename is the sheet's new name, when it was renamed; Gone, it was
	// deleted.
	Rename string
	Gone   bool
	// Move: the cells of Src on Sheet were cut and pasted (sheets move), to
	// sheet To, DR rows down and DC columns right; what they carry (a
	// validation, a conditional format) went with them.
	Move   bool
	Src    grid.Range
	To     string
	DR, DC int
}

// Effects is what one session did.
type Effects struct {
	Written []EffectCell
	Results []ResultChange
	Ops     []LineOp
}

// sameResult tells whether two results are the same, numbers to Excel's 15
// significant digits (the digits past them are where two calculators
// differ).
func sameResult(a, b grid.Cell) bool {
	if a.Kind != b.Kind {
		return false
	}
	switch a.Kind {
	case grid.Number:
		d := math.Abs(a.Num - b.Num)
		return d <= 5e-15*math.Max(math.Abs(a.Num), math.Abs(b.Num)) || d < 1e-15
	case grid.Text, grid.Error:
		return a.Text == b.Text
	case grid.Bool:
		return a.Bool == b.Bool
	}
	return true
}

// SameResult is sameResult, for the report.
func SameResult(a, b grid.Cell) bool { return sameResult(a, b) }

// effectOf is what computing f did to its stored result, or false when
// the report has nothing to learn from it: a result recomputed to what it
// was, or a written formula's plain value (the command reports those).
func effectOf(f *formulaAt, written bool, carried carryKind, must bool, v grid.Cell, ok bool, why string) (ResultChange, bool) {
	rc := ResultChange{Sheet: f.sheet, Row: f.row, Col: f.col, Written: written, Carried: written && carried == carryError, Blank: written && (carried == carryNone || f.priorBlank), Reached: must, After: valueOnly(v), Now: ok, Why: why}
	if !ok {
		rc.After = grid.Cell{}
	}
	if !written && f.stored.Kind != grid.Empty && !f.pending {
		rc.Before, rc.Was = valueOnly(f.stored), true
	}
	if !written && f.pending && f.staleRes.Kind != grid.Empty {
		rc.Before, rc.Stale = valueOnly(f.staleRes), true
	}
	if !written && f.written && f.priorKnown {
		// A formula regrown as Excel would (Editor.regrown): read back as
		// written, its result before is the one it stored.
		rc.Before, rc.Was = f.prior, true
	}
	if written && f.priorKnown {
		rc.Before, rc.Was = f.prior, true
	}
	switch {
	case !must:
		// A result missing when opened, computed now: it only tells a
		// report over several sessions what an earlier one left unknown.
		return rc, ok
	case written:
		return rc, !ok || rc.After.Kind == grid.Error
	case rc.Was && ok && sameResult(rc.Before, rc.After):
		return rc, false
	}
	return rc, true
}

// writtenCells lists cells written, from a session's changed map; the
// cells a dynamic array's result fills or leaves, and the formulas only
// regrown as Excel would, are not among them.
func writtenCells(changed map[cellKey]string, spilled map[cellKey]*spillCell, regrown map[cellKey]bool) []EffectCell {
	out := make([]EffectCell, 0, len(changed))
	for k, sheet := range changed {
		if _, ok := spilled[k]; ok || regrown[k] {
			continue
		}
		out = append(out, EffectCell{Sheet: sheet, Row: k.row, Col: k.col})
	}
	return out
}

// spillCell is a cell a dynamic array's result fills or leaves: its value
// when the session began and now.
type spillCell struct {
	sheet         string
	before, after grid.Cell
}

// spillEffects are those cells' values, before and after, as results the
// report can compare: a value that changed is one the command moved.
func spillEffects(spilled map[cellKey]*spillCell) []ResultChange {
	var out []ResultChange
	for k, c := range spilled {
		out = append(out, ResultChange{Sheet: c.sheet, Row: k.row, Col: k.col,
			Before: c.before, Was: true, After: c.after, Now: true})
	}
	return out
}

// remapNotes moves what the session notes of cells where a change of rows,
// columns or sheets puts them (false: it took them).
func (e *Editor) remapNotes(move func(cellKey) (cellKey, bool)) {
	if len(e.carried) > 0 {
		m := map[cellKey]carryKind{}
		for k, kind := range e.carried {
			if k2, ok := move(k); ok {
				m[k2] = kind
			}
		}
		e.carried = m
	}
	if len(e.regrown) > 0 {
		m := map[cellKey]bool{}
		for k := range e.regrown {
			if k2, ok := move(k); ok {
				m[k2] = true
			}
		}
		e.regrown = m
	}
	if len(e.spilled) > 0 {
		m := map[cellKey]*spillCell{}
		for k, c := range e.spilled {
			if k2, ok := move(k); ok {
				m[k2] = c
			}
		}
		e.spilled = m
	}
}

// Effects is what the session did, as of its last save (or render).
func (e *Editor) Effects() []Effects { return []Effects{e.effects()} }

func (e *Editor) effects() Effects {
	ef := Effects{Written: writtenCells(e.changed, e.spilled, e.regrown), Ops: e.lineOps}
	if e.res != nil {
		ef.Results = append(ef.Results, e.res.effects...)
	}
	ef.Results = append(ef.Results, spillEffects(e.spilled)...)
	return ef
}

// Before is the file as the session opened it (nil for a new one), and
// After as its last save or render made it.
func (e *Editor) Before() []byte { return e.data }
func (e *Editor) After() []byte  { return e.out }

// Effects is what the session did, as of its last save (or render).
func (s *StreamEditor) Effects() []Effects { return []Effects{s.effects()} }

func (s *StreamEditor) effects() Effects {
	ef := Effects{Written: writtenCells(s.changed, nil, nil)}
	for _, op := range s.ops {
		switch op.kind {
		case 'i':
			ef.Ops = append(ef.Ops, LineOp{Sheet: op.sheet, At: op.at, N: op.n})
		case 'd':
			ef.Ops = append(ef.Ops, LineOp{Sheet: op.sheet, At: op.from, N: -(op.to - op.from + 1)})
		}
	}
	if s.res != nil {
		ef.Results = s.res.effects
	}
	return ef
}

func (s *StreamEditor) Before() []byte { return s.data }
func (s *StreamEditor) After() []byte  { return s.out }

// Effects lists what each part of the session did: the streaming part,
// then the full editor's when it took over (the second begins where the
// first left the file).
func (a *Auto) Effects() []Effects {
	var out []Effects
	if a.handed != nil {
		out = append(out, *a.handed)
	}
	if a.e != nil {
		return append(out, a.e.effects())
	}
	return append(out, a.s.effects())
}

func (a *Auto) Before() []byte { return a.data }
func (a *Auto) After() []byte {
	if a.e != nil {
		return a.e.After()
	}
	return a.s.After()
}

// recordOp notes a change of rows, columns or sheets.
func (e *Editor) recordOp(op LineOp) { e.lineOps = append(e.lineOps, op) }

// carryKind is what a formula a sort or move carried showed where it came
// from: an error, or no result.
type carryKind uint8

const (
	carryNothing carryKind = iota
	carryError
	carryNone
)

// notePrior notes what the formula a write replaced held (old is the
// formula as the session found it; nil if the cell held none).
func (f *formulaAt) notePrior(old *formulaAt) {
	switch {
	case old == nil:
	case old.written:
		f.prior, f.priorKnown, f.priorBlank = old.prior, old.priorKnown, old.priorBlank
	case old.pending || old.stored.Kind == grid.Empty:
		f.priorBlank = true
	default:
		f.prior, f.priorKnown = valueOnly(old.stored), true
	}
}

// carry notes a formula put where it stands by a sort or move that
// showed an error, or no result, where it came from.
func (e *Editor) carry(k cellKey, h *taken, text string) {
	if text == "" {
		return
	}
	kind := carryNone
	if h.cell.Kind == grid.Error && !h.cell.Pending {
		kind = carryError
	} else if h.cell.Kind != grid.Empty && !h.cell.Pending {
		return
	}
	if e.carried == nil {
		e.carried = map[cellKey]carryKind{}
	}
	e.carried[k] = kind
}

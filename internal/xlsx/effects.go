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
	// Before is the result stored when the session began; Was says it was
	// known (a result stored and current).
	Before grid.Cell
	Was    bool
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
func effectOf(f *formulaAt, written, must bool, v grid.Cell, ok bool, why string) (ResultChange, bool) {
	rc := ResultChange{Sheet: f.sheet, Row: f.row, Col: f.col, Written: written, After: valueOnly(v), Now: ok, Why: why}
	if !ok {
		rc.After = grid.Cell{}
	}
	if !written && f.stored.Kind != grid.Empty && !f.pending {
		rc.Before, rc.Was = valueOnly(f.stored), true
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

// writtenCells lists cells written, from a session's changed map.
func writtenCells(changed map[cellKey]string) []EffectCell {
	out := make([]EffectCell, 0, len(changed))
	for k, sheet := range changed {
		out = append(out, EffectCell{Sheet: sheet, Row: k.row, Col: k.col})
	}
	return out
}

// Effects is what the session did, as of its last save (or render).
func (e *Editor) Effects() []Effects { return []Effects{e.effects()} }

func (e *Editor) effects() Effects {
	ef := Effects{Written: writtenCells(e.changed), Ops: e.lineOps}
	if e.res != nil {
		ef.Results = e.res.effects
	}
	return ef
}

// Before is the file as the session opened it (nil for a new one), and
// After as its last save or render made it.
func (e *Editor) Before() []byte { return e.data }
func (e *Editor) After() []byte  { return e.out }

// Effects is what the session did, as of its last save (or render).
func (s *StreamEditor) Effects() []Effects { return []Effects{s.effects()} }

func (s *StreamEditor) effects() Effects {
	ef := Effects{Written: writtenCells(s.changed)}
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

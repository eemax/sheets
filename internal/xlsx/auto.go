package xlsx

import (
	"errors"
	"os"
	"strconv"

	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/input"
)

// Auto edits a large workbook with the streaming editor for as long as it
// can (see stream.go), and with the full editor from the first thing it
// cannot do: the edits made so far are saved into memory and the full
// editor opens that.
type Auto struct {
	path     string
	s        *StreamEditor
	e        *Editor
	warnings []string // from the streaming part of the session
}

// StreamFrom is the file size from which edits stream; -1 when they never
// do. They do by default: the streaming editor is faster at every size,
// and hands over to the full editor for what it does not do.
// SHEETS_STREAM=never turns it off (for comparing, say), or sets a size
// in bytes.
func StreamFrom() int64 {
	switch v := os.Getenv("SHEETS_STREAM"); v {
	case "never", "off":
		return -1
	default:
		if n, err := strconv.ParseInt(v, 10, 64); err == nil && n >= 0 {
			return n
		}
		return 0 // "", "always"
	}
}

// OpenAuto opens a workbook for editing, streaming while it can; data is
// the file.
func OpenAuto(path string, data []byte) (*Auto, error) {
	s, err := openStream(path, data)
	if err != nil {
		return nil, err
	}
	return &Auto{path: path, s: s}, nil
}

// Streaming tells whether no edit has needed the full editor so far.
func (a *Auto) Streaming() bool { return a.e == nil }

// full hands the session to the full editor.
func (a *Auto) full() (*Editor, error) {
	if a.e != nil {
		return a.e, nil
	}
	data := a.s.data
	if len(a.s.ops) > 0 || len(a.s.changed) > 0 {
		var err error
		if data, err = a.s.Render(); err != nil {
			return nil, err
		}
		a.warnings = append(a.warnings, a.s.warnings...)
	}
	e, err := OpenEditorData(a.path, data)
	if err != nil {
		return nil, err
	}
	a.e, a.s = e, nil
	return e, nil
}

// try runs fn on the streaming editor, and on the full editor when it
// cannot.
func (a *Auto) try(stream func(*StreamEditor) error, full func(*Editor) error) error {
	if a.e == nil {
		if err := stream(a.s); !errors.Is(err, errUpgrade) {
			return err
		}
	}
	e, err := a.full()
	if err != nil {
		return err
	}
	return full(e)
}

// on runs fn on the full editor.
func (a *Auto) on(fn func(*Editor) error) error {
	e, err := a.full()
	if err != nil {
		return err
	}
	return fn(e)
}

func (a *Auto) Workbook() *grid.Workbook {
	if a.e != nil {
		return a.e.Workbook()
	}
	return a.s.Workbook()
}

func (a *Auto) Created() bool { return false }

func (a *Auto) EnsureSheet(name string) (sheet string, created bool, err error) {
	err = a.try(func(s *StreamEditor) error {
		sheet, created, err = s.EnsureSheet(name)
		return err
	}, func(e *Editor) error {
		sheet, created, err = e.EnsureSheet(name)
		return err
	})
	return
}

func (a *Auto) Set(sheet string, row, col int, v input.Value) error {
	return a.try(func(s *StreamEditor) error { return s.Set(sheet, row, col, v) },
		func(e *Editor) error { return e.Set(sheet, row, col, v) })
}

func (a *Auto) Evaluate(sheet string, row, col int) (grid.Cell, string, bool) {
	if a.e != nil {
		return a.e.Evaluate(sheet, row, col)
	}
	return a.s.Evaluate(sheet, row, col)
}

func (a *Auto) InsertRows(sheet string, at, n int) error {
	return a.try(func(s *StreamEditor) error { return s.InsertRows(sheet, at, n) },
		func(e *Editor) error { return e.InsertRows(sheet, at, n) })
}

func (a *Auto) DeleteRows(sheet string, from, to int) error {
	return a.try(func(s *StreamEditor) error { return s.DeleteRows(sheet, from, to) },
		func(e *Editor) error { return e.DeleteRows(sheet, from, to) })
}

func (a *Auto) ExtendTable(sheet string, body grid.Range, n int) {
	if a.e != nil {
		a.e.ExtendTable(sheet, body, n)
		return
	}
	a.s.ExtendTable(sheet, body, n)
}

func (a *Auto) References(sheet string) []string {
	if a.e != nil {
		return a.e.References(sheet)
	}
	return a.s.References(sheet)
}

func (a *Auto) SheetList() []string {
	if a.e != nil {
		return a.e.SheetList()
	}
	return a.s.SheetList()
}

// Records lists the recorded styling; before any edit, straight from the
// package.
func (a *Auto) Records() []Record {
	if a.e == nil && len(a.s.ops) == 0 {
		if recs, _, err := rawRecords(a.s.data); err == nil {
			return recs
		}
	}
	e, err := a.full()
	if err != nil {
		return nil
	}
	return e.Records()
}

func (a *Auto) Save(dryRun bool) error {
	if a.e != nil {
		return a.e.Save(dryRun)
	}
	return a.s.Save(dryRun)
}

func (a *Auto) Render() ([]byte, error) {
	if a.e != nil {
		return a.e.Render()
	}
	return a.s.Render()
}

func (a *Auto) Close() error {
	if a.e != nil {
		return a.e.Close()
	}
	return nil
}

func (a *Auto) Warnings() []string {
	if a.e != nil {
		return append(append([]string{}, a.warnings...), a.e.Warnings()...)
	}
	return append(append([]string{}, a.warnings...), a.s.Warnings()...)
}

// What only the full editor does.

func (a *Auto) Clear(sheet string) error {
	return a.on(func(e *Editor) error { return e.Clear(sheet) })
}
func (a *Auto) InsertCols(sheet string, at, n int) error {
	return a.on(func(e *Editor) error { return e.InsertCols(sheet, at, n) })
}
func (a *Auto) DeleteCols(sheet string, from, to int) error {
	return a.on(func(e *Editor) error { return e.DeleteCols(sheet, from, to) })
}
func (a *Auto) AddSheet(name, before, after string) error {
	return a.on(func(e *Editor) error { return e.AddSheet(name, before, after) })
}
func (a *Auto) RenameSheet(from, to string) error {
	return a.on(func(e *Editor) error { return e.RenameSheet(from, to) })
}
func (a *Auto) DeleteSheet(name string) error {
	return a.on(func(e *Editor) error { return e.DeleteSheet(name) })
}
func (a *Auto) CopySheet(from, to string) (notes []string, err error) {
	err = a.on(func(e *Editor) error { notes, err = e.CopySheet(from, to); return err })
	return
}
func (a *Auto) MoveSheet(name, before, after string, first, last bool) error {
	return a.on(func(e *Editor) error { return e.MoveSheet(name, before, after, first, last) })
}
func (a *Auto) SetVisible(name string, visible bool) error {
	return a.on(func(e *Editor) error { return e.SetVisible(name, visible) })
}
func (a *Auto) AddChart(spec ChartSpec) error {
	return a.on(func(e *Editor) error { return e.AddChart(spec) })
}
func (a *Auto) DeleteChart(sheet, cell string) error {
	return a.on(func(e *Editor) error { return e.DeleteChart(sheet, cell) })
}
func (a *Auto) Style(sheet string, r grid.Range, l *Look) (res StyleResult, err error) {
	err = a.on(func(e *Editor) error { res, err = e.Style(sheet, r, l); return err })
	return
}
func (a *Auto) Unrecord(seqs []int) error {
	return a.on(func(e *Editor) error { return e.Unrecord(seqs) })
}
func (a *Auto) Cond(sheet string, r grid.Range, ws []string) (ref string, err error) {
	err = a.on(func(e *Editor) error { ref, err = e.Cond(sheet, r, ws); return err })
	return
}
func (a *Auto) CondFormats(sheet string) (cfs []CondFormat, err error) {
	err = a.on(func(e *Editor) error { cfs, err = e.CondFormats(sheet); return err })
	return
}
func (a *Auto) ClearCond(sheet string, r *grid.Range) (n int, err error) {
	err = a.on(func(e *Editor) error { n, err = e.ClearCond(sheet, r); return err })
	return
}
func (a *Auto) Validate(sheet string, r grid.Range, rule ValidationRule) (ref string, err error) {
	err = a.on(func(e *Editor) error { ref, err = e.Validate(sheet, r, rule); return err })
	return
}
func (a *Auto) Validations(sheet string) (vs []Validation, err error) {
	err = a.on(func(e *Editor) error { vs, err = e.Validations(sheet); return err })
	return
}
func (a *Auto) ClearValidation(sheet string, r *grid.Range) (n int, err error) {
	err = a.on(func(e *Editor) error { n, err = e.ClearValidation(sheet, r); return err })
	return
}
func (a *Auto) View(sheet string, o ViewOptions) error {
	return a.on(func(e *Editor) error { return e.View(sheet, o) })
}
func (a *Auto) ViewOf(sheet string) (v ViewState, err error) {
	err = a.on(func(e *Editor) error { v, err = e.ViewOf(sheet); return err })
	return
}
func (a *Auto) Looks(sheet string, r grid.Range) (runs []LookRun, err error) {
	err = a.on(func(e *Editor) error { runs, err = e.Looks(sheet, r); return err })
	return
}
func (a *Auto) ClearRange(sheet string, r grid.Range, w ClearWhat) (res RangeResult, err error) {
	err = a.on(func(e *Editor) error { res, err = e.ClearRange(sheet, r, w); return err })
	return
}
func (a *Auto) CopyRange(from string, src grid.Range, to string, at grid.Range, o CopyOptions) (res RangeResult, err error) {
	err = a.on(func(e *Editor) error { res, err = e.CopyRange(from, src, to, at, o); return err })
	return
}
func (a *Auto) MoveRange(from string, src grid.Range, to string, at grid.Range) (res RangeResult, err error) {
	err = a.on(func(e *Editor) error { res, err = e.MoveRange(from, src, to, at); return err })
	return
}
func (a *Auto) SortRange(sheet string, r grid.Range, keys []SortKey) (res SortResult, err error) {
	err = a.on(func(e *Editor) error { res, err = e.SortRange(sheet, r, keys); return err })
	return
}
func (a *Auto) Replace(sheet string, r *grid.Range, fn func(string) (string, bool), formulas bool) (out []Replacement, err error) {
	err = a.on(func(e *Editor) error { out, err = e.Replace(sheet, r, fn, formulas); return err })
	return
}
func (a *Auto) SetNote(sheet string, row, col int, text, author string) error {
	return a.on(func(e *Editor) error { return e.SetNote(sheet, row, col, text, author) })
}
func (a *Auto) SetLink(sheet string, row, col int, l *grid.Link, text string) error {
	return a.on(func(e *Editor) error { return e.SetLink(sheet, row, col, l, text) })
}

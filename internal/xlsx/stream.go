package xlsx

import (
	"archive/zip"
	"bytes"
	"errors"
	"fmt"
	"regexp"
	"sort"
	"strings"

	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/input"
	"github.com/eemax/sheets/internal/style"
)

// A StreamEditor edits a large workbook without loading its sheets into
// excelize, which takes seconds and gigabytes for a sheet of a few hundred
// thousand rows. It keeps the edits it is given (cells written, rows
// inserted or deleted, tables extended) and applies them when saving, in
// one pass over each sheet they touch; the package's other parts are
// copied as they are. Formula results are computed from the parsed cells
// (see stream_calc.go).
//
// It does the edits large files mostly get: setting cells and rows. For
// anything else it returns errUpgrade, and the caller (see Auto) hands
// the work to the full editor.

// errUpgrade means the streaming editor does not do what was asked.
var errUpgrade = errors.New("needs the full editor")

// streamOp is a change of rows, in the order made.
type streamOp struct {
	kind  byte // 'i' insert, 'd' delete, 'x' extend a table's ranges
	sheet string
	// insert: n rows before at. delete: rows from..to. extend: ranges
	// within columns c1..c2 ending at row from now end at row to.
	at, n, from, to, c1, c2 int
}

// streamCell is a cell written.
type streamCell struct {
	v      input.Value
	numFmt string // number format to give it; "" keeps its own
}

// StreamEditor: see above.
type StreamEditor struct {
	path   string
	data   []byte
	zr     *zip.Reader
	p      *pkg
	wbPath string
	parts  map[string]string // sheet name → its part
	wb     *grid.Workbook    // as opened
	tables []tableSource     // Excel tables as opened

	ops       []streamOp
	edits     map[string]map[[2]int]*streamCell // lower-cased sheet → cell where it is now → content
	expand    map[string]bool                   // lower-cased sheets whose shared formulas are written out
	changed   map[cellKey]string
	recompute map[cellKey]bool
	extent    map[string]grid.Range // used range by lower-cased sheet, where edits grew it
	warnings  []string
	res       *calcResult // after saving
	// merges are a sheet's merged cells where they are now, and used its
	// used range, by lower-cased sheet, as of a count of ops.
	merges map[string]mergesAt
	used   map[string]mergesAt
}

type mergesAt struct {
	ops int
	ms  []grid.Range
}

// mergesNow lists a sheet's merged cells where they are now.
func (s *StreamEditor) mergesNow(sheet string) []grid.Range {
	k := strings.ToLower(sheet)
	if m, ok := s.merges[k]; ok && m.ops == len(s.ops) {
		return m.ms
	}
	var ms []grid.Range
	if sh := s.wb.Sheet(sheet); sh != nil {
		for _, m := range sh.Merges {
			if now, ok := s.area(sheet, m, false); ok {
				ms = append(ms, now)
			}
		}
	}
	if s.merges == nil {
		s.merges = map[string]mergesAt{}
	}
	s.merges[k] = mergesAt{ops: len(s.ops), ms: ms}
	return ms
}

// OpenStream opens a workbook for streaming edits.
func OpenStream(path string) (*StreamEditor, error) {
	data, err := readFile(path)
	if err != nil {
		return nil, err
	}
	return openStream(path, data)
}

func openStream(path string, data []byte) (*StreamEditor, error) {
	s := &StreamEditor{path: path, data: data, edits: map[string]map[[2]int]*streamCell{}, expand: map[string]bool{},
		changed: map[cellKey]string{}, recompute: map[cellKey]bool{}, extent: map[string]grid.Range{}}
	var err error
	if s.zr, err = zip.NewReader(bytes.NewReader(data), int64(len(data))); err != nil {
		return nil, fmt.Errorf("%s: not an xlsx file: %w", path, err)
	}
	s.p = &pkg{files: map[string]*zip.File{}}
	for _, zf := range s.zr.File {
		s.p.files[strings.ToLower(zf.Name)] = zf
	}
	if s.wbPath, _, s.parts, err = s.p.workbookParts(); err != nil {
		return nil, fmt.Errorf("%s: %w", path, err)
	}
	if s.wb, err = ReadBytes(data); err != nil {
		return nil, fmt.Errorf("%s: %w", path, err)
	}
	s.wb.Path = path
	for name, part := range s.parts {
		for _, t := range s.p.sheetParts(part).tables {
			s.tables = append(s.tables, tableSource{name, t})
		}
	}
	return s, nil
}

// Workbook returns the contents as they were when the file was opened.
func (s *StreamEditor) Workbook() *grid.Workbook { return s.wb }

// Created is false: a new file has nothing to stream.
func (s *StreamEditor) Created() bool { return false }

func (s *StreamEditor) sheetName(name string) (string, error) {
	if sh := s.wb.Sheet(name); sh != nil {
		return sh.Name, nil
	}
	return "", fmt.Errorf("sheet %q not found (sheets: %s)", name, strings.Join(s.wb.SheetNames(), ", "))
}

// EnsureSheet finds a sheet; a new one needs the full editor.
func (s *StreamEditor) EnsureSheet(name string) (string, bool, error) {
	if sh := s.wb.Sheet(name); sh != nil {
		return sh.Name, false, nil
	}
	return "", false, errUpgrade
}

// rowNow is where row r of a sheet as opened is now; false when it was
// deleted.
func (s *StreamEditor) rowNow(sheet string, r int) (int, bool) {
	for _, op := range s.ops {
		if !strings.EqualFold(op.sheet, sheet) {
			continue
		}
		switch op.kind {
		case 'i':
			if r >= op.at {
				r += op.n
			}
		case 'd':
			switch {
			case r > op.to:
				r -= op.to - op.from + 1
			case r >= op.from:
				return 0, false
			}
		}
	}
	return r, true
}

// rowThen is the row of the sheet as opened that row r now is; false for
// a row inserted since.
func (s *StreamEditor) rowThen(sheet string, r int) (int, bool) {
	for i := len(s.ops) - 1; i >= 0; i-- {
		op := s.ops[i]
		if !strings.EqualFold(op.sheet, sheet) {
			continue
		}
		switch op.kind {
		case 'i':
			switch {
			case r >= op.at+op.n:
				r -= op.n
			case r >= op.at:
				return 0, false
			}
		case 'd':
			if r >= op.from {
				r += op.to - op.from + 1
			}
		}
	}
	return r, true
}

// rewrite brings a formula of the workbook as opened, held on sheet home
// ("" for names), up to the rows as they are now.
func (s *StreamEditor) rewrite(f, home string) string {
	for _, op := range s.ops {
		f = op.formula(f, home)
	}
	return f
}

// ruleEdits are the edits, in order, that move the conditional formats and
// validations of a sheet: its own rows', and other sheets' for the
// formulas reading them.
func (s *StreamEditor) ruleEdits(sheet string) []lineEdit {
	var out []lineEdit
	for _, op := range s.ops {
		rewrite := func(f string) string { return op.formula(f, sheet) }
		switch {
		case op.kind == 'x' && strings.EqualFold(op.sheet, sheet):
			out = append(out, lineEdit{
				place: func(r grid.Range) (grid.Range, bool) {
					if r.R2 == op.from && r.C1 >= op.c1 && r.C2 <= op.c2 {
						r.R2 = op.to
					}
					return r, true
				},
				first:   func(r grid.Range) (int, int) { return r.R1, r.C1 },
				rewrite: rewrite,
			})
		case op.kind == 'x':
		case strings.EqualFold(op.sheet, sheet):
			out = append(out, lineEdit{
				place: func(r grid.Range) (grid.Range, bool) { return s.areaOp(r, op) },
				first: func(r grid.Range) (int, int) {
					row := r.R1
					if op.kind == 'd' && row >= op.from && row <= op.to {
						row = op.to + 1
					}
					return row, r.C1
				},
				rewrite: rewrite,
			})
		default:
			out = append(out, stay(rewrite))
		}
	}
	return out
}

func (op streamOp) formula(f, home string) string {
	switch op.kind {
	case 'i':
		return formula.InsertLines(f, home, op.sheet, false, op.at, op.n)
	case 'd':
		return formula.DeleteLines(f, home, op.sheet, false, op.from, op.to)
	}
	return f
}

// ranges is the rewrite of chart and table ranges (always qualified): rows
// inserted and deleted, and ranges extended over rows appended to a table.
func (op streamOp) ranges(f string) string {
	if op.kind == 'x' {
		return formula.ExtendRangesIn(f, "", op.sheet, op.from, op.to, op.c1, op.c2)
	}
	return op.formula(f, "")
}

func (s *StreamEditor) rangeRewrites() []func(string) string {
	var out []func(string) string
	for _, op := range s.ops {
		out = append(out, op.ranges)
	}
	return out
}

// area brings a range of a sheet as opened (a merge, a conditional
// format's, a validation's) up to the rows as they are now; false when its
// rows were all deleted. Conditional formats and validations (grow) extend
// over rows appended to a table, as their records do.
func (s *StreamEditor) area(sheet string, r grid.Range, grow bool) (grid.Range, bool) {
	text := r.String()
	for _, op := range s.ops {
		if !strings.EqualFold(op.sheet, sheet) {
			continue
		}
		if op.kind == 'x' {
			if grow && r.R2 == op.from && r.C1 >= op.c1 && r.C2 <= op.c2 {
				r.R2 = op.to
				text = r.String()
			}
			continue
		}
		text = op.formula(text, sheet)
		if strings.Contains(text, "#REF!") {
			return grid.Range{}, false
		}
		var err error
		if r, err = grid.ParseRange(text); err != nil {
			return grid.Range{}, false
		}
	}
	return r, true
}

func (s *StreamEditor) edit(sheet string, row, col int) *streamCell {
	return s.edits[strings.ToLower(sheet)][[2]int{row, col}]
}

// cellNow is a cell as it is now: written in this session, or as opened.
// A formula's text is brought up to the rows as they are now.
func (s *StreamEditor) cellNow(sheet string, row, col int) grid.Cell {
	if e := s.edit(sheet, row, col); e != nil {
		v := e.v
		if v.Formula != "" {
			return grid.Cell{Formula: v.Formula, Pending: true}
		}
		return grid.Cell{Kind: v.Kind, Text: v.Text, Num: v.Num, Bool: v.Bool}
	}
	r, ok := s.rowThen(sheet, row)
	if !ok {
		return grid.Cell{}
	}
	sh := s.wb.Sheet(sheet)
	if sh == nil {
		return grid.Cell{}
	}
	c := sh.Get(r, col)
	if c.Formula != "" && len(s.ops) > 0 {
		c.Formula = s.rewrite(c.Formula, sh.Name)
	}
	return c
}

// formulasNow lists the workbook's formulas where they are now.
func (s *StreamEditor) formulasNow() map[cellKey]*formulaAt {
	all := map[cellKey]*formulaAt{}
	for _, sh := range s.wb.Sheets {
		sh.Each(sh.Used(), func(r, c int, cell *grid.Cell) {
			if cell.Formula == "" {
				return
			}
			now, ok := s.rowNow(sh.Name, r)
			if !ok || s.edit(sh.Name, now, c) != nil {
				return
			}
			text := cell.Formula
			if len(s.ops) > 0 {
				text = s.rewrite(text, sh.Name)
			}
			all[keyOf(sh.Name, now, c)] = &formulaAt{sheet: sh.Name, row: now, col: c, text: text, pending: cell.Pending, stored: valueOf(*cell)}
		})
	}
	for _, sh := range s.wb.Sheets {
		for pos, e := range s.edits[strings.ToLower(sh.Name)] {
			if e.v.Formula != "" {
				all[keyOf(sh.Name, pos[0], pos[1])] = &formulaAt{sheet: sh.Name, row: pos[0], col: pos[1], text: e.v.Formula, pending: true}
			}
		}
	}
	return all
}

// tablesNow lists the Excel tables where they are now.
func (s *StreamEditor) tablesNow(sheet string) []tableNow {
	var out []tableNow
	rw := s.rangeRewrites()
	for _, t := range s.tables {
		if tn, ok := t.now(rw); ok && (sheet == "" || strings.EqualFold(tn.sheet, sheet)) {
			out = append(out, tn)
		}
	}
	return out
}

// spilling tells whether a cell is a dynamic array formula's, or holds
// part of its result: those need the full editor.
func (s *StreamEditor) spilling(sheet string, row, col int) bool {
	sh := s.wb.Sheet(sheet)
	if sh == nil || len(sh.Spills) == 0 {
		return false
	}
	r, ok := s.rowThen(sheet, row)
	if !ok {
		return false
	}
	for _, sp := range sh.Spills {
		if sp.Contains(r, col) {
			return true
		}
	}
	return false
}

func (s *StreamEditor) touch(sheet string, row, col int) {
	s.changed[keyOf(sheet, row, col)] = sheet
	k := strings.ToLower(sheet)
	x, ok := s.extent[k]
	if !ok {
		x = s.usedNow(sheet)
	}
	if x.Empty() {
		x = grid.Range{R1: row, C1: col, R2: row, C2: col}
	}
	s.extent[k] = x.Union(grid.Range{R1: row, C1: col, R2: row, C2: col})
}

// usedNow is a sheet's used range brought up to the rows as they are now
// (where it ends may be a little past its last value).
func (s *StreamEditor) usedNow(sheet string) grid.Range {
	k := strings.ToLower(sheet)
	if x, ok := s.extent[k]; ok {
		return x
	}
	if u, ok := s.used[k]; ok && u.ops == len(s.ops) {
		return u.ms[0]
	}
	x := grid.Range{R1: 1, C1: 1}
	if sh := s.wb.Sheet(sheet); sh != nil {
		// Used reads every cell: once a sheet.
		if u := sh.Used(); u.Empty() {
			x = u
		} else if moved, ok := s.area(sheet, u, false); ok {
			x = moved
		}
	}
	if s.used == nil {
		s.used = map[string]mergesAt{}
	}
	s.used[k] = mergesAt{ops: len(s.ops), ms: []grid.Range{x}}
	return x
}

// Set writes one value, as Editor.Set does.
func (s *StreamEditor) Set(sheet string, row, col int, v input.Value) error {
	if v.Formula != "" {
		if err := formula.Check(v.Formula); err != nil {
			return fmt.Errorf("%s: =%s is not a valid formula: %v", grid.FormatRef(sheet, grid.Range{R1: row, C1: col, R2: row, C2: col}), strings.TrimPrefix(v.Formula, "="), err)
		}
		if formula.Spills(formula.ToFile(v.Formula)) {
			return errUpgrade // a dynamic array: cell metadata
		}
	}
	name, err := s.sheetName(sheet)
	if err != nil {
		return err
	}
	if m, ok, _ := hiddenIn(s.mergesNow(name), row, col); ok {
		if v.IsEmpty() {
			return nil
		}
		return hiddenByMerge(name, row, col, m)
	}
	if s.spilling(name, row, col) {
		return errUpgrade
	}
	for _, t := range s.tablesNow(name) {
		if t.headers > 0 && row == t.ref.R1 && col >= t.ref.C1 && col <= t.ref.C2 {
			return errUpgrade // a table column's name: renames
		}
	}
	if r, ok := s.rowThen(name, row); ok {
		if c := s.wb.Sheet(name).At(r, col); c != nil && c.Shared {
			// Overwriting a cell of a filled-down group: the group's cells
			// are written out each with its formula.
			s.expand[strings.ToLower(name)] = true
		}
	}
	if s.wb.Date1904 && (v.Fmt.Kind == grid.Date || v.Fmt.Kind == grid.DateTime) {
		v.Num -= 1462
	}
	c := &streamCell{v: v}
	k := strings.ToLower(name)
	if v.Fmt.Kind != grid.General {
		c.numFmt = style.NumFmt(v.Fmt, false)
	} else if prev := s.edits[k][[2]int{row, col}]; prev != nil {
		// A cell keeps the number format an earlier write gave it, as
		// its style does.
		c.numFmt = prev.numFmt
	}
	if s.edits[k] == nil {
		s.edits[k] = map[[2]int]*streamCell{}
	}
	s.edits[k][[2]int{row, col}] = c
	s.touch(name, row, col)
	return nil
}

// InsertRows inserts n empty rows before row at.
func (s *StreamEditor) InsertRows(sheet string, at, n int) error {
	return s.lines(sheet, streamOp{kind: 'i', at: at, n: n})
}

// DeleteRows deletes rows from..to (inclusive).
func (s *StreamEditor) DeleteRows(sheet string, from, to int) error {
	name, err := s.sheetName(sheet)
	if err != nil {
		return err
	}
	for _, t := range s.tablesNow(name) {
		if t.headers > 0 && from <= t.ref.R1 && to >= t.ref.R1 {
			return fmt.Errorf("row %d is the header row of Excel table %s; sheets cannot delete it (Excel would remove the table)", t.ref.R1, t)
		}
		first, last := t.ref.R1+t.headers, t.ref.R2-t.totals
		if first <= last && from <= first && to >= last {
			return fmt.Errorf("rows %d:%d are all the data rows of Excel table %s; a table keeps at least one", first, last, t)
		}
	}
	return s.lines(name, streamOp{kind: 'd', from: from, to: to})
}

// lines makes a change of rows: references follow on every sheet, and the
// formulas whose results it can change are recomputed when saving.
func (s *StreamEditor) lines(sheet string, op streamOp) error {
	name, err := s.sheetName(sheet)
	if err != nil {
		return err
	}
	op.sheet = name
	sh := s.wb.Sheet(name)
	// Dynamic arrays on the sheet: their results would need moving cell
	// by cell, with their metadata.
	if len(sh.Spills) > 0 {
		return errUpgrade
	}
	hit := func(rg grid.Range) bool { return rg.R1 < op.at && op.at <= rg.R2 }
	if op.kind == 'd' {
		hit = func(rg grid.Range) bool { return rg.R1 <= op.to && rg.R2 >= op.from }
	}
	rs := s.resolver()
	before := s.formulasNow()
	s.ops = append(s.ops, op)
	move := func(r, c int) (int, int, bool) {
		switch op.kind {
		case 'i':
			if r >= op.at {
				r += op.n
			}
		case 'd':
			switch {
			case r > op.to:
				r -= op.to - op.from + 1
			case r >= op.from:
				return 0, 0, false
			}
		}
		return r, c, true
	}
	// What the session tracks moves with its rows.
	rekey := func(k cellKey) (cellKey, bool) {
		if k.sheet != strings.ToLower(name) {
			return k, true
		}
		r, c, ok := move(k.row, k.col)
		k.row, k.col = r, c
		return k, ok
	}
	recompute := map[cellKey]bool{}
	for k := range s.recompute {
		if k2, ok := rekey(k); ok {
			recompute[k2] = true
		}
	}
	for k, f := range before {
		if s.recompute[k] || movedResult(rs, f.sheet, f.row, f.col, f.text, name, hit) {
			if k2, ok := rekey(k); ok {
				recompute[k2] = true
			}
		}
	}
	changed := map[cellKey]string{}
	for k, v := range s.changed {
		if k2, ok := rekey(k); ok {
			changed[k2] = v
		}
	}
	s.recompute, s.changed = recompute, changed
	for k, cells := range s.edits {
		moved := map[[2]int]*streamCell{}
		for pos, c := range cells {
			if c.v.Formula != "" {
				home := k
				if h := s.wb.Sheet(k); h != nil {
					home = h.Name
				}
				c.v.Formula = op.formula(c.v.Formula, home)
			}
			if k == strings.ToLower(name) {
				r, col, ok := move(pos[0], pos[1])
				if !ok {
					continue
				}
				pos = [2]int{r, col}
			}
			moved[pos] = c
		}
		s.edits[k] = moved
	}
	if x, ok := s.extent[strings.ToLower(name)]; ok {
		if y, ok := s.areaOp(x, op); ok {
			s.extent[strings.ToLower(name)] = y
		} else {
			s.extent[strings.ToLower(name)] = grid.Range{R1: 1, C1: 1}
		}
	}
	// Formulas filled down on the sheet, or reading it from elsewhere,
	// are written out each with its formula.
	s.expand[strings.ToLower(name)] = s.expand[strings.ToLower(name)] || hasShared(sh)
	for _, other := range s.wb.Sheets {
		if s.expand[strings.ToLower(other.Name)] || !hasShared(other) {
			continue
		}
		other.Each(other.Used(), func(_, _ int, c *grid.Cell) {
			if c.Shared && formula.RefersToSheet(c.Formula, name) {
				s.expand[strings.ToLower(other.Name)] = true
			}
		})
	}
	return nil
}

func (s *StreamEditor) areaOp(r grid.Range, op streamOp) (grid.Range, bool) {
	text := op.formula(r.String(), op.sheet)
	if strings.Contains(text, "#REF!") {
		return grid.Range{}, false
	}
	x, err := grid.ParseRange(text)
	return x, err == nil
}

func hasShared(sh *grid.Sheet) bool {
	found := false
	sh.Each(sh.Used(), func(_, _ int, c *grid.Cell) { found = found || c.Shared })
	return found
}

// ExtendTable makes the charts, Excel table, recorded styling, conditional
// formats and validations of a table whose data is body take in n rows
// written right below it.
func (s *StreamEditor) ExtendTable(sheet string, body grid.Range, n int) {
	name, err := s.sheetName(sheet)
	if err != nil {
		return
	}
	s.ops = append(s.ops, streamOp{kind: 'x', sheet: name, from: body.R2, to: body.R2 + n, c1: body.C1, c2: body.C2})
}

// References lists the formula cells on other sheets that refer to sheet,
// by its name or its Excel tables' names, in sheet and row order, and the
// names that do (see Editor.References).
func (s *StreamEditor) References(sheet string) []string {
	var tables []*regexp.Regexp
	for _, t := range s.tablesNow(sheet) {
		tables = append(tables, regexp.MustCompile(`(?i)(^|[^\w.])`+regexp.QuoteMeta(t.name)+`($|[^\w.(])`))
	}
	refers := func(f string) bool {
		if formula.RefersToSheet(f, sheet) {
			return true
		}
		for _, re := range tables {
			if re.MatchString(f) {
				return true
			}
		}
		return false
	}
	all := s.formulasNow()
	var fs []*formulaAt
	for _, f := range all {
		if !strings.EqualFold(f.sheet, sheet) && refers(f.text) {
			fs = append(fs, f)
		}
	}
	sort.Slice(fs, func(i, j int) bool {
		a, b := fs[i], fs[j]
		if a.sheet != b.sheet {
			return s.sheetOrder(a.sheet) < s.sheetOrder(b.sheet)
		}
		if a.row != b.row {
			return a.row < b.row
		}
		return a.col < b.col
	})
	var refs []string
	for _, f := range fs {
		refs = append(refs, grid.FormatRef(f.sheet, grid.Range{R1: f.row, C1: f.col, R2: f.row, C2: f.col}))
	}
	for _, n := range s.wb.Names {
		if refers(s.rewrite(n.Ref, "")) {
			refs = append(refs, "name "+n.Name)
		}
	}
	return refs
}

// SheetList is the sheet order, hidden sheets marked.
func (s *StreamEditor) SheetList() []string {
	var out []string
	for _, sh := range s.wb.Sheets {
		n := sh.Name
		if sh.Hidden {
			n += " (hidden)"
		}
		out = append(out, n)
	}
	return out
}

// Warnings lists problems noticed while editing that did not stop it.
func (s *StreamEditor) Warnings() []string { return s.warnings }

// Evaluate is a formula's result as saved, or why there is none (see
// Editor.Evaluate); before saving it is not known.
func (s *StreamEditor) Evaluate(sheet string, row, col int) (grid.Cell, string, bool) {
	if s.res == nil {
		return grid.Cell{}, "computed when the file is saved", false
	}
	if v, ok := s.res.results[strings.ToLower(sheet)][grid.CellName(row, col)]; ok {
		note := s.res.dropped[keyOf(sheet, row, col)]
		if v.Kind == grid.Empty && !strings.HasPrefix(note, "circular reference") && !strings.Contains(note, "Excel computes it") {
			note += "; Excel computes it on open"
		}
		return v, note, v.Kind != grid.Empty
	}
	c := s.cellNow(sheet, row, col)
	if c.Pending {
		return grid.Cell{}, "Excel computes it on open", false
	}
	return valueOf(c), "", true
}

// Close releases resources.
func (s *StreamEditor) Close() error { return nil }

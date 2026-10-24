package xlsx

import (
	"fmt"
	"regexp"
	"sort"
	"strings"

	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
	"github.com/xuri/excelize/v2"
)

// Structural edits. excelize shifts formulas (on every sheet), names,
// merges, conditional formats, validations, filters and tables when rows
// or columns move; renames and deletions of sheets are completed here.

// InsertRows inserts n empty rows before row at.
func (e *Editor) InsertRows(sheet string, at, n int) error {
	err := e.relines(sheet, false,
		func(f, home string) string { return formula.InsertLines(f, home, sheet, false, at, n) },
		func(rg grid.Range) bool { return rg.R1 < at && at <= rg.R2 },
		func(r, c int) (int, int, bool) {
			if r >= at {
				r += n
			}
			return r, c, true
		},
		func() error { return e.f.InsertRows(sheet, at, n) })
	if err != nil {
		return err
	}
	if err := e.movePanes(sheet, false, at, at+n-1, true); err != nil {
		return err
	}
	e.shiftExtent(sheet, func(r grid.Range) grid.Range {
		if r.R2 >= at {
			r.R2 += n
			if r.R1 >= at {
				r.R1 += n
			}
		}
		return r
	})
	return nil
}

// DeleteRows deletes rows from..to (inclusive).
func (e *Editor) DeleteRows(sheet string, from, to int) error {
	if err := e.checkTableRows(sheet, from, to); err != nil {
		return err
	}
	if err := e.deleteLines(sheet, false, from, to); err != nil {
		return err
	}
	if err := e.movePanes(sheet, false, from, to, false); err != nil {
		return err
	}
	n := to - from + 1
	e.shiftExtent(sheet, func(r grid.Range) grid.Range {
		switch {
		case r.R1 > to:
			r.R1 -= n
			r.R2 -= n
		case r.R2 >= from:
			// What follows the deleted rows moves up into them.
			r.R1 = min(r.R1, from)
			r.R2 = max(from-1, r.R2-n)
		}
		return r
	})
	return nil
}

// InsertCols inserts n empty columns before column at.
func (e *Editor) InsertCols(sheet string, at, n int) error {
	if err := e.checkTableCols(sheet, at, at, true); err != nil {
		return err
	}
	err := e.relines(sheet, true,
		func(f, home string) string { return formula.InsertLines(f, home, sheet, true, at, n) },
		func(rg grid.Range) bool { return rg.C1 < at && at <= rg.C2 },
		func(r, c int) (int, int, bool) {
			if c >= at {
				c += n
			}
			return r, c, true
		},
		func() error { return e.f.InsertCols(sheet, grid.ColName(at), n) })
	if err != nil {
		return err
	}
	if err := e.movePanes(sheet, true, at, at+n-1, true); err != nil {
		return err
	}
	e.shiftExtent(sheet, func(r grid.Range) grid.Range {
		if r.C2 >= at {
			r.C2 += n
			if r.C1 >= at {
				r.C1 += n
			}
		}
		return r
	})
	return nil
}

// DeleteCols deletes columns from..to (inclusive).
func (e *Editor) DeleteCols(sheet string, from, to int) error {
	if err := e.checkTableCols(sheet, from, to, false); err != nil {
		return err
	}
	if err := e.deleteLines(sheet, true, from, to); err != nil {
		return err
	}
	if err := e.movePanes(sheet, true, from, to, false); err != nil {
		return err
	}
	n := to - from + 1
	e.shiftExtent(sheet, func(r grid.Range) grid.Range {
		switch {
		case r.C1 > to:
			r.C1 -= n
			r.C2 -= n
		case r.C2 >= from:
			r.C1 = min(r.C1, from)
			r.C2 = max(from-1, r.C2-n)
		}
		return r
	})
	return nil
}

// movePanes keeps frozen rows and columns frozen across lines from..to
// inserted or deleted, as Excel does: lines among the frozen ones move the
// split.
func (e *Editor) movePanes(sheet string, cols bool, from, to int, insert bool) error {
	p, err := e.f.GetPanes(sheet)
	if err != nil || !p.Freeze {
		return err
	}
	split := &p.YSplit
	if cols {
		split = &p.XSplit
	}
	if *split == 0 || from > *split {
		return nil
	}
	if insert {
		*split += to - from + 1
	} else {
		*split -= min(to, *split) - from + 1
	}
	return freeze(e.f, sheet, grid.Range{R1: p.YSplit + 1, C1: p.XSplit + 1, R2: p.YSplit + 1, C2: p.XSplit + 1})
}

// deleteLines removes rows or columns from..to.
func (e *Editor) deleteLines(target string, cols bool, from, to int) error {
	n := to - from + 1
	return e.relines(target, cols,
		func(f, home string) string { return formula.DeleteLines(f, home, target, cols, from, to) },
		func(rg grid.Range) bool {
			if cols {
				return rg.C1 <= to && rg.C2 >= from
			}
			return rg.R1 <= to && rg.R2 >= from
		},
		func(r, c int) (int, int, bool) {
			line := r
			if cols {
				line = c
			}
			switch {
			case line >= from && line <= to:
				return 0, 0, false // deleted with its line
			case line > to && cols:
				c -= n
			case line > to:
				r -= n
			}
			return r, c, true
		},
		func() error {
			for i := to; i >= from; i-- {
				var err error
				if cols {
					err = e.f.RemoveCol(target, grid.ColName(i))
				} else {
					err = e.f.RemoveRow(target, i)
				}
				if err != nil {
					return err
				}
			}
			return nil
		})
}

// relines inserts or deletes rows or columns of target through op, keeping
// formulas right. excelize's own formula adjustment points references into
// deleted lines at their neighbours instead of #REF!, fails on LET and
// LAMBDA names, and re-parses every formula once per deleted line (minutes
// for a block of rows in a big sheet). So formulas are taken off first,
// rewritten in one pass (rewrite gets the formula and the sheet holding it),
// and put back where move says their cell went once op has moved cells,
// names, merges, conditional formats, validations, filters and tables.
//
// hit tells whether a range on the target sheet takes in, or spans, the
// lines that come or go: the result of a formula reading such a range can
// change, and is recomputed when saving.
func (e *Editor) relines(target string, cols bool, rewrite func(f, home string) string, hit func(grid.Range) bool, move func(r, c int) (int, int, bool), op func() error) error {
	rs := e.resolver()
	now, err := e.formulasNow()
	if err != nil {
		return err
	}
	var all []*formulaAt
	for _, f := range now {
		all = append(all, f)
	}
	// Bottom-right first, so a shared formula's anchor goes last.
	sort.Slice(all, func(i, j int) bool {
		a, b := all[i], all[j]
		if a.sheet != b.sheet {
			return a.sheet < b.sheet
		}
		if a.row != b.row {
			return a.row > b.row
		}
		return a.col > b.col
	})
	// Not nil, even with no formula left: lines set means the formulas
	// opened with have moved (see formulasNow).
	keep := []formulaAt{}
	recompute := map[cellKey]bool{}
	for _, f := range all {
		if err := e.f.SetCellFormula(f.sheet, grid.CellName(f.row, f.col), ""); err != nil {
			return err
		}
		m := formulaAt{sheet: f.sheet, row: f.row, col: f.col, text: rewrite(f.text, f.sheet), pending: f.pending, stored: f.stored}
		if strings.EqualFold(f.sheet, target) {
			var ok bool
			if m.row, m.col, ok = move(f.row, f.col); !ok {
				continue
			}
		}
		keep = append(keep, m)
		if e.recompute[keyOf(f.sheet, f.row, f.col)] || movedResult(rs, f.sheet, f.row, f.col, f.text, target, hit) {
			recompute[keyOf(m.sheet, m.row, m.col)] = true
		}
	}
	// Cells written earlier in the session move with their rows too.
	changed := map[cellKey]string{}
	fresh := map[cellKey]bool{}
	for k, sheet := range e.changed {
		was := e.wasEmpty[k]
		if k.sheet == strings.ToLower(target) {
			r, c, ok := move(k.row, k.col)
			if !ok {
				continue
			}
			k.row, k.col = r, c
		}
		changed[k] = sheet
		if was {
			fresh[k] = true
		}
	}
	// Names: excelize shifts them, but cannot parse LET names either; and
	// it would point names into deleted lines at their neighbours, or at
	// row 0, which it then refuses (a filter's name, when the rows deleted
	// begin at its header). So all names are taken off and put back
	// moved. A filter goes with its header row (excelize drops it), and so
	// does its name; serialize hides the name again (see hideRecords).
	var names []excelize.DefinedName
	for _, n := range e.f.GetDefinedName() {
		names = append(names, n)
		if err := e.f.DeleteDefinedName(&excelize.DefinedName{Name: n.Name, Scope: n.Scope}); err != nil {
			return err
		}
	}
	// Validations and conditional formats move as Excel moves them, not
	// as excelize does (see linefix.go): every sheet's, whose formulas can
	// read the target.
	held := map[string]heldValidations{}
	for _, name := range e.f.GetSheetList() {
		h, err := e.holdValidations(name)
		if err != nil {
			if e.wb.Sheet(name) == nil {
				continue // a chart sheet
			}
			return err
		}
		if len(h) > 0 {
			held[name] = h
		}
	}
	if err := op(); err != nil {
		return err
	}
	e.dropMerges()
	edit := func(home string) lineEdit {
		if !strings.EqualFold(home, target) {
			return stay(func(f string) string { return rewrite(f, home) })
		}
		return lineEdit{
			place: func(r grid.Range) (grid.Range, bool) {
				text := rewrite(r.String(), target)
				if strings.Contains(text, "#REF!") {
					return grid.Range{}, false
				}
				m, err := grid.ParseRange(strings.ReplaceAll(text, "$", ""))
				return m, err == nil
			},
			first: func(r grid.Range) (int, int) {
				for row, col := r.R1, r.C1; row <= r.R2 && col <= r.C2; {
					if _, _, ok := move(row, col); ok {
						return row, col
					}
					if cols {
						col++
					} else {
						row++
					}
				}
				return r.R1, r.C1
			},
			rewrite: func(f string) string { return rewrite(f, target) },
		}
	}
	for name, h := range held {
		// Their formulas name sheets as excelize holds them, renamed when
		// saving (see namesNow).
		for _, dv := range h {
			for _, f := range []*string{&dv.Formula1, &dv.Formula2} {
				if *f != "" {
					*f = escapeXML(e.namesNow(unescapeXML(*f)))
				}
			}
		}
		if err := h.putBack(e, name, edit(name)); err != nil {
			return err
		}
	}
	for home, sec := range e.patches.conds {
		sec.each(edit(home).moveRule)
	}
	if err := e.moveNotes(target, move); err != nil {
		return err
	}
	// Chart ranges are always sheet-qualified.
	e.patches.rangeRewrites = append(e.patches.rangeRewrites, func(f string) string { return rewrite(f, "") })
	for _, n := range names {
		if strings.HasPrefix(n.Name, "_xlnm.") {
			if n.Name == "_xlnm._FilterDatabase" && headerGone(n.RefersTo, target, hit, cols) {
				continue
			}
			if moved := rewrite(n.RefersTo, ""); !strings.Contains(moved, "#REF!") {
				n.RefersTo = moved
				if err := e.f.SetDefinedName(&n); err != nil {
					return err
				}
			}
			continue
		}
		n.RefersTo = rewrite(n.RefersTo, "")
		if strings.HasPrefix(n.Name, recordPrefix) {
			// The formulas of a recorded rule follow as the rule's do.
			if recs, _ := recordsOf([]excelize.DefinedName{n}, func(string) bool { return true }); len(recs) == 1 && recs[0].Kind == "cond" {
				rec := recs[0]
				n.Comment = Record{Kind: rec.Kind, Args: condFormulas(rec.Args, func(f string) string { return rewrite(f, rec.Sheet) })}.Text()
			}
		}
		if err := e.f.SetDefinedName(&n); err != nil {
			return err
		}
	}
	for _, m := range keep {
		if err := e.setFormula(m.sheet, grid.CellName(m.row, m.col), m.text); err != nil {
			return err
		}
	}
	e.lines, e.recompute, e.changed, e.wasEmpty = keep, recompute, changed, fresh
	e.structural = true
	e.forgetPositions()
	return nil
}

// sheetEdit notes a sheet renamed (to is the new name), deleted (to is
// ""), or copied (copy is set), keeping what the session tracks for
// recomputing results in step: the formulas where they stand, and the
// cells written.
func (e *Editor) sheetEdit(from, to string, copy bool, rewrite func(string) string) {
	e.sheetsChanged = true
	e.dropMerges()
	if e.lines != nil {
		lines := []formulaAt{}
		for _, l := range e.lines {
			mine := strings.EqualFold(l.sheet, from)
			if copy && mine {
				c := l
				c.sheet = to
				lines = append(lines, c)
			}
			if mine && !copy {
				if to == "" {
					continue
				}
				l.sheet = to
			}
			if rewrite != nil {
				l.text = rewrite(l.text)
			}
			lines = append(lines, l)
		}
		e.lines = lines
	}
	if copy {
		return
	}
	// Conditional formats, which excelize does not hold, follow now.
	if to == "" {
		delete(e.patches.conds, from)
	}
	if rewrite != nil {
		for _, sec := range e.patches.conds {
			sec.formulas(rewrite)
		}
	}
	move := func(k cellKey) (cellKey, bool) {
		if k.sheet != strings.ToLower(from) {
			return k, true
		}
		k.sheet = strings.ToLower(to)
		return k, to != ""
	}
	changed := map[cellKey]string{}
	wasEmpty := map[cellKey]bool{}
	for k, sheet := range e.changed {
		if k2, ok := move(k); ok {
			if k2 != k {
				sheet = to
			}
			changed[k2] = sheet
			wasEmpty[k2] = e.wasEmpty[k]
		}
	}
	e.wasEmpty = wasEmpty
	recompute := map[cellKey]bool{}
	for k := range e.recompute {
		if k2, ok := move(k); ok {
			recompute[k2] = true
		}
	}
	e.changed, e.recompute = changed, recompute
}

// movedResult tells whether moving lines can change a formula's result:
// it reads a range that takes in or spans them, or depends on positions
// on the target sheet. (Volatile formulas are recomputed anyway.)
func movedResult(rs *resolver, sheet string, row, col int, text, target string, hit func(grid.Range) bool) bool {
	d := formula.DepsOf(text)
	refs, ok := rs.ranges(sheet, row, col, d, 0)
	if !ok {
		return true
	}
	onTarget := strings.EqualFold(sheet, target)
	for _, r := range refs {
		if strings.EqualFold(r.Sheet, target) {
			if hit(r.Range) {
				return true
			}
			onTarget = true
		}
	}
	return d.Position && onTarget
}

// forgetPositions drops cell positions recorded at open, which have moved;
// later edits in the same session work from what excelize holds.
func (e *Editor) forgetPositions() {
	for _, l := range e.layouts {
		l.spills = map[[2]int]grid.Range{}
	}
	e.unshare = map[string]bool{}
	for _, s := range e.wb.Sheets {
		s.Each(s.Used(), func(_, _ int, c *grid.Cell) { c.Shared = false })
	}
}

func (e *Editor) shiftExtent(sheet string, fn func(grid.Range) grid.Range) {
	ext, ok := e.extent[sheet]
	if !ok {
		ext = grid.Range{R1: 1, C1: 1}
		if s := e.wb.Sheet(sheet); s != nil {
			ext = s.Used()
		}
	}
	if !ext.Empty() {
		ext = fn(ext)
	}
	e.extent[sheet] = ext
}

// AddSheet adds an empty sheet, placed before or after another one when
// given.
func (e *Editor) AddSheet(name, before, after string) error {
	if e.wb.Sheet(name) != nil {
		return fmt.Errorf("sheet %q already exists", name)
	}
	if _, _, err := e.EnsureSheet(name); err != nil {
		return err
	}
	return e.place(name, before, after)
}

// place moves a sheet before or after another; with neither it stays.
func (e *Editor) place(name, before, after string) error {
	switch {
	case before != "":
		target, err := e.sheetName(before)
		if err != nil {
			return err
		}
		return e.f.MoveSheet(name, target)
	case after != "":
		target, err := e.sheetName(after)
		if err != nil {
			return err
		}
		list := e.f.GetSheetList()
		i := indexOf(list, target)
		if i == len(list)-1 {
			// MoveSheet places before its target: move the last sheet
			// before ours instead.
			if list[i] == name {
				return nil
			}
			return e.f.MoveSheet(list[i], name)
		}
		next := list[i+1]
		if next == name {
			return nil
		}
		return e.f.MoveSheet(name, next)
	}
	return nil
}

// MoveSheet changes a sheet's position: before or after another sheet, or
// to the first or last place.
func (e *Editor) MoveSheet(name, before, after string, first, last bool) error {
	name, err := e.sheetName(name)
	if err != nil {
		return err
	}
	list := e.f.GetSheetList()
	switch {
	case first:
		if list[0] == name {
			return nil
		}
		return e.f.MoveSheet(name, list[0])
	case last:
		return e.place(name, "", list[len(list)-1])
	}
	return e.place(name, before, after)
}

func indexOf(list []string, s string) int {
	for i, v := range list {
		if v == s {
			return i
		}
	}
	return -1
}

func (e *Editor) sheetName(name string) (string, error) {
	for _, n := range e.f.GetSheetList() {
		if strings.EqualFold(n, name) {
			return n, nil
		}
	}
	return "", fmt.Errorf("sheet %q not found (sheets: %s)", name, strings.Join(e.f.GetSheetList(), ", "))
}

// RenameSheet renames a sheet and every reference to it: formulas on all
// sheets, names, conditional formats, validations, charts and pivot
// sources. excelize itself only updates names.
func (e *Editor) RenameSheet(from, to string) error {
	from, err := e.sheetName(from)
	if err != nil {
		return err
	}
	if existing, err := e.sheetName(to); err == nil && existing != from {
		return fmt.Errorf("sheet %q already exists", existing)
	}
	if err := e.f.SetSheetName(from, to); err != nil {
		return fmt.Errorf("sheet name %q: %w", to, err)
	}
	e.patches.rename(from, to)
	if err := e.rewriteCondRecords(func(f, _ string) string { return formula.RenameSheet(f, from, to) }); err != nil {
		return err
	}
	e.patches.sheetEdits = append(e.patches.sheetEdits, sheetEdit{from: from, to: to})
	e.patches.rangeRewrites = append(e.patches.rangeRewrites, func(f string) string { return formula.RenameSheet(f, from, to) })
	e.renameState(from, to)
	e.rewriteSnapshot(func(f string) string { return formula.RenameSheet(f, from, to) })
	e.sheetEdit(from, to, false, func(f string) string { return formula.RenameSheet(f, from, to) })
	return nil
}

// rewriteSnapshot applies a change of sheet references to the formulas
// recorded at open, so later edits in the session see current names.
func (e *Editor) rewriteSnapshot(fn func(string) string) {
	for _, s := range e.wb.Sheets {
		s.Each(s.Used(), func(_, _ int, c *grid.Cell) {
			if c.Formula != "" {
				c.Formula = fn(c.Formula)
			}
		})
	}
	for i, n := range e.wb.Names {
		e.wb.Names[i].Ref = fn(n.Ref)
	}
}

func (e *Editor) renameState(from, to string) {
	if s := e.wb.Sheet(from); s != nil {
		s.Name = to
	}
	if l, ok := e.layouts[from]; ok {
		e.layouts[to] = l
		delete(e.layouts, from)
	}
	if x, ok := e.extent[from]; ok {
		e.extent[to] = x
		delete(e.extent, from)
	}
}

// References lists the formula cells on other sheets that refer to sheet.
func (e *Editor) References(sheet string) []string {
	// Formulas reach a sheet by its name, or by the names of its Excel
	// tables (Sales[Units]).
	var tables []*regexp.Regexp
	if s := e.wb.Sheet(sheet); s != nil {
		for _, t := range s.Tables {
			tables = append(tables, regexp.MustCompile(`(?i)(^|[^\w.])`+regexp.QuoteMeta(t.Name)+`($|[^\w.(])`))
		}
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
	var refs []string
	for _, s := range e.wb.Sheets {
		if strings.EqualFold(s.Name, sheet) {
			continue
		}
		s.Each(s.Used(), func(r, c int, cell *grid.Cell) {
			if cell.Formula != "" && refers(cell.Formula) {
				refs = append(refs, grid.FormatRef(s.Name, grid.Range{R1: r, C1: c, R2: r, C2: c}))
			}
		})
	}
	for _, n := range e.wb.Names {
		if refers(n.Ref) {
			refs = append(refs, "name "+n.Name)
		}
	}
	return refs
}

// DeleteSheet deletes a sheet. References to it from elsewhere become
// #REF!, as in Excel; callers check References first.
func (e *Editor) DeleteSheet(name string) error {
	name, err := e.sheetName(name)
	if err != nil {
		return err
	}
	if len(e.f.GetSheetList()) == 1 {
		return fmt.Errorf("cannot delete %q: a workbook needs at least one sheet", name)
	}
	// Formulas reading the sheet become #REF!: new results to store.
	rs := e.resolver()
	now, err := e.formulasNow()
	if err != nil {
		return err
	}
	for _, f := range now {
		if !strings.EqualFold(f.sheet, name) && movedResult(rs, f.sheet, f.row, f.col, f.text, name, func(grid.Range) bool { return true }) {
			e.mustRecompute(f.sheet, f.row, f.col)
		}
	}
	for _, rec := range e.Records() {
		if strings.EqualFold(rec.Sheet, name) {
			if err := dropRecords(e.f, rec.name); err != nil {
				return err
			}
		}
	}
	if err := e.f.DeleteSheet(name); err != nil {
		return err
	}
	e.sheetEdit(name, "", false, func(f string) string { return formula.DropSheet(f, name) })
	e.patches.sheetEdits = append(e.patches.sheetEdits, sheetEdit{from: name})
	e.patches.rangeRewrites = append(e.patches.rangeRewrites, func(f string) string { return formula.DropSheet(f, name) })
	e.rewriteSnapshot(func(f string) string { return formula.DropSheet(f, name) })
	for i, s := range e.wb.Sheets {
		if s.Name == name {
			e.wb.Sheets = append(e.wb.Sheets[:i], e.wb.Sheets[i+1:]...)
			break
		}
	}
	delete(e.extent, name)
	delete(e.layouts, name)
	return nil
}

// CopySheet duplicates a sheet's cells and formatting under a new name,
// placed right after the original.
func (e *Editor) CopySheet(from, to string) ([]string, error) {
	from, err := e.sheetName(from)
	if err != nil {
		return nil, err
	}
	if s := e.wb.Sheet(from); s != nil && len(s.Tables) > 0 {
		return nil, fmt.Errorf("cannot copy %q: it holds Excel table %q, which cannot be duplicated", from, s.Tables[0].Name)
	}
	if e.wb.Sheet(to) != nil {
		return nil, fmt.Errorf("sheet %q already exists", to)
	}
	src, _ := e.f.GetSheetIndex(from)
	idx, err := e.f.NewSheet(to)
	if err != nil {
		return nil, fmt.Errorf("sheet name %q: %w", to, err)
	}
	if err := e.f.CopySheet(src, idx); err != nil {
		return nil, err
	}
	if err := e.place(to, "", from); err != nil {
		return nil, err
	}
	// Its conditional formats, held aside (see linefix.go), go with it.
	if sec := e.patches.conds[from]; sec != nil {
		e.patches.conds[to] = sec.copy()
	}
	// The copy's styling is kept by fmt as the original's is.
	for _, rec := range e.Records() {
		if strings.EqualFold(rec.Sheet, from) {
			rec.Sheet = to
			if _, err := addRecord(e.f, rec); err != nil {
				return nil, err
			}
		}
	}
	var notes []string
	if e.hasDrawings[from] {
		notes = append(notes, "charts and pictures are not copied")
	}
	if s := e.wb.Sheet(from); s != nil {
		cp := *s
		cp.Name = to
		e.wb.Sheets = append(e.wb.Sheets, &cp)
		e.extent[to] = s.Used()
	}
	e.sheetEdit(from, to, true, nil)
	return notes, nil
}

// SetVisible hides or shows a sheet.
func (e *Editor) SetVisible(name string, visible bool) error {
	name, err := e.sheetName(name)
	if err != nil {
		return err
	}
	if !visible {
		n := 0
		for _, s := range e.f.GetSheetList() {
			if v, _ := e.f.GetSheetVisible(s); v {
				n++
			}
		}
		if v, _ := e.f.GetSheetVisible(name); v && n == 1 {
			return fmt.Errorf("cannot hide %q: a workbook needs a visible sheet", name)
		}
	}
	return e.f.SetSheetVisible(name, visible)
}

// newRefErrors compares formulas before and after structural edits and
// lists cells that gained a #REF!.
func newRefErrors(before, after *grid.Workbook) []string {
	had := map[string]bool{}
	for _, s := range before.Sheets {
		s.Each(s.Used(), func(_, _ int, c *grid.Cell) {
			if strings.Contains(c.Formula, "#REF!") {
				had[s.Name+"\x00"+c.Formula] = true
			}
		})
	}
	var out []string
	for _, s := range after.Sheets {
		s.Each(s.Used(), func(r, c int, cell *grid.Cell) {
			if strings.Contains(cell.Formula, "#REF!") && !had[s.Name+"\x00"+cell.Formula] {
				out = append(out, grid.FormatRef(s.Name, grid.Range{R1: r, C1: c, R2: r, C2: c}))
			}
		})
	}
	sort.Strings(out)
	return out
}

// SheetList is the current sheet order, hidden sheets marked.
func (e *Editor) SheetList() []string {
	var out []string
	for _, s := range e.f.GetSheetList() {
		if v, _ := e.f.GetSheetVisible(s); !v {
			s += " (hidden)"
		}
		out = append(out, s)
	}
	return out
}

// headerGone tells whether lines deleted take a filter's header row (or
// first column, for columns) with them, which excelize answers by dropping
// the filter: its name goes too. hit tells whether a range on target takes
// in deleted lines (for one line, whether it is one of them; nothing does
// when lines are inserted).
func headerGone(ref, target string, hit func(grid.Range) bool, cols bool) bool {
	sheet, r, ok := parseAbsRef(ref)
	if !ok || !strings.EqualFold(sheet, target) {
		return false
	}
	head := grid.Range{R1: r.R1, C1: r.C1, R2: r.R1, C2: r.C2}
	if cols {
		head = grid.Range{R1: r.R1, C1: r.C1, R2: r.R2, C2: r.C1}
	}
	return hit(head)
}

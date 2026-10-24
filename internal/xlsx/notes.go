package xlsx

import (
	"fmt"
	"math"
	"regexp"
	"strings"
	"unicode/utf8"

	"github.com/eemax/sheets/internal/grid"
	"github.com/xuri/excelize/v2"
)

// linkColor is the colour of hyperlinks in Excel's Office theme.
const linkColor = "0563C1"

// linkLook is how a hyperlink looks: blue and underlined.
func linkLook() *Look {
	on := true
	return &Look{Color: linkColor, Underline: &on}
}

// SetNote puts a note on a cell, replacing the one there; empty text
// removes it. A comment thread (Excel 365) is left alone: sheets cannot
// write those.
func (e *Editor) SetNote(sheet string, row, col int, text, author string) error {
	sheet, err := e.sheetName(sheet)
	if err != nil {
		return err
	}
	cell := grid.CellName(row, col)
	if s := e.wb.Sheet(sheet); s != nil {
		for _, n := range s.Notes {
			if n.Row == row && n.Col == col && n.Thread {
				return fmt.Errorf("%s holds a comment thread (Excel 365); sheets cannot change those: reply or delete it in Excel", grid.FormatRef(sheet, grid.Range{R1: row, C1: col, R2: row, C2: col}))
			}
		}
	}
	if err := e.deleteNote(sheet, cell); err != nil {
		return err
	}
	if text == "" {
		return nil
	}
	if utf8.RuneCountInString(text) > 32000 {
		return fmt.Errorf("a note holds up to 32,000 characters; this one has %d", utf8.RuneCountInString(text))
	}
	w, h := noteSize(text)
	e.patches.notesChanged[sheet] = true
	return e.f.AddComment(sheet, excelize.Comment{Cell: cell, Author: author, Text: text, Width: w, Height: h})
}

// deleteNote removes the note on a cell, if any.
func (e *Editor) deleteNote(sheet, cell string) error {
	cs, err := e.f.GetComments(sheet)
	if err != nil {
		return err
	}
	for _, c := range cs {
		if c.Cell == cell {
			e.patches.notesChanged[sheet] = true
			return e.f.DeleteComment(sheet, cell)
		}
	}
	return nil
}

// noteSize is a box size, in points, that shows a note's text: about 30
// characters a line, growing up to a page.
func noteSize(text string) (w, h uint) {
	lines := 0
	for _, l := range strings.Split(text, "\n") {
		lines += max(1, int(math.Ceil(float64(utf8.RuneCountInString(l))/30)))
	}
	return 160, uint(min(max(lines*13+12, 56), 400))
}

// SetLink puts a hyperlink on a cell: to a web address, file or e-mail
// (url), or to a place in the workbook (location, like Sheet!A1). The
// cell gets text when given, or the address itself when empty, and the
// look of a link. A nil link removes the cell's hyperlink.
func (e *Editor) SetLink(sheet string, row, col int, l *grid.Link, text string) error {
	sheet, err := e.sheetName(sheet)
	if err != nil {
		return err
	}
	cell := grid.CellName(row, col)
	if l == nil {
		if err := e.f.SetCellHyperLink(sheet, cell, "", "None"); err != nil {
			return err
		}
		return e.unlinkLook(sheet, row, col)
	}
	// One link per cell: drop the old one, whatever its kind.
	if err := e.f.SetCellHyperLink(sheet, cell, "", "None"); err != nil {
		return err
	}
	var opts []excelize.HyperlinkOpts
	if l.Tooltip != "" {
		tip := l.Tooltip
		opts = append(opts, excelize.HyperlinkOpts{Tooltip: &tip})
	}
	if l.URL != "" {
		err = e.f.SetCellHyperLink(sheet, cell, l.URL, "External", opts...)
	} else {
		err = e.f.SetCellHyperLink(sheet, cell, l.Location, "Location", opts...)
	}
	if err != nil {
		return err
	}
	cur, err := e.f.GetCellValue(sheet, cell, excelize.Options{RawCellValue: true})
	if err != nil {
		return err
	}
	if fm, _ := e.f.GetCellFormula(sheet, cell); fm != "" {
		cur = "=" + fm
	}
	switch {
	case text != "":
		if err := setText(e.f, sheet, cell, text); err != nil {
			return err
		}
		e.touch(sheet, row, col)
		e.grow(sheet, grid.Range{R1: row, C1: col, R2: row, C2: col})
	case cur == "":
		show := l.URL
		if show == "" {
			show = l.Location
		}
		if err := setText(e.f, sheet, cell, strings.TrimPrefix(show, "mailto:")); err != nil {
			return err
		}
		e.touch(sheet, row, col)
		e.grow(sheet, grid.Range{R1: row, C1: col, R2: row, C2: col})
	}
	return e.linkLook(e.styler(), sheet, row, col)
}

// linkLook gives a cell the look of a link: underlined, and blue unless
// that would be hard to read on its fill.
func (e *Editor) linkLook(st *styler, sheet string, row, col int) error {
	return lookLikeLink(st, sheet, row, col)
}

func lookLikeLink(st *styler, sheet string, row, col int) error {
	l := linkLook()
	if fill := fillOf(st.f, sheet, row, col); fill != "" && contrast(linkColor, fill) < 4.5 {
		l.Color = ""
	}
	return st.apply(sheet, grid.Range{R1: row, C1: col, R2: row, C2: col}, l, nil)
}

// unlinkLook takes the look of a link off a cell that has exactly it.
func (e *Editor) unlinkLook(sheet string, row, col int) error {
	id, err := e.f.GetCellStyle(sheet, grid.CellName(row, col))
	if err != nil {
		return err
	}
	s, err := e.f.GetStyle(id)
	if err != nil || s.Font == nil || s.Font.Underline != "single" || !strings.EqualFold(strings.TrimPrefix(s.Font.Color, "FF"), linkColor) {
		return err
	}
	off := false
	return e.styler().apply(sheet, grid.Range{R1: row, C1: col, R2: row, C2: col}, &Look{Color: "000000", Underline: &off}, nil)
}

// fillOf is a cell's solid fill colour, "" for none.
func fillOf(f *excelize.File, sheet string, row, col int) string {
	id, err := f.GetCellStyle(sheet, grid.CellName(row, col))
	if err != nil {
		return ""
	}
	s, err := f.GetStyle(id)
	if err != nil || s.Fill.Type != "pattern" || s.Fill.Pattern != 1 || len(s.Fill.Color) == 0 {
		return ""
	}
	c := strings.TrimPrefix(s.Fill.Color[0], "#")
	if len(c) == 8 {
		c = c[2:]
	}
	return strings.ToUpper(c)
}

// A note as excelize holds it, to put back elsewhere.
type heldNote struct {
	row, col int
	c        excelize.Comment
}

// moveNotes moves a sheet's notes for cells moved by move, which reports
// where a cell went and false when it went away (its row or column was
// deleted). excelize leaves notes where they were when rows or columns
// move. Comment threads (Excel 365) follow at save (see moveThreads).
func (e *Editor) moveNotes(sheet string, move func(r, c int) (int, int, bool)) error {
	e.patches.threadMoves = append(e.patches.threadMoves, threadMove{sheet: sheet, move: move})
	cs, err := e.f.GetComments(sheet)
	if err != nil || len(cs) == 0 {
		return err
	}
	var put []heldNote
	for _, c := range cs {
		r, col, ok := grid.ParseCell(c.Cell)
		if !ok {
			continue
		}
		nr, nc, keep := move(r, col)
		if keep && nr == r && nc == col {
			continue
		}
		e.patches.notesChanged[sheet] = true
		if err := e.f.DeleteComment(sheet, c.Cell); err != nil {
			return err
		}
		if keep {
			c.Cell = grid.CellName(nr, nc)
			put = append(put, heldNote{nr, nc, c})
		}
	}
	for _, n := range put {
		if err := e.addHeldNote(sheet, n); err != nil {
			return err
		}
	}
	return nil
}

// addHeldNote puts back a note excelize read, replacing one on its cell.
func (e *Editor) addHeldNote(sheet string, n heldNote) error {
	if err := e.deleteNote(sheet, n.c.Cell); err != nil {
		return err
	}
	c := n.c
	if len(c.Paragraph) > 0 {
		c.Text = ""
	}
	if c.Width == 0 || c.Height == 0 {
		var b strings.Builder
		b.WriteString(c.Text)
		for _, p := range c.Paragraph {
			b.WriteString(p.Text)
		}
		c.Width, c.Height = noteSize(b.String())
	}
	e.patches.notesChanged[sheet] = true
	return e.f.AddComment(sheet, c)
}

// threadMove is a move of cells whose comment threads follow at save.
type threadMove struct {
	sheet string
	move  func(r, c int) (int, int, bool)
}

var (
	threadRef  = regexp.MustCompile(`(<threadedComment\b[^>]*?\bref=")([A-Z]+[0-9]+)(")`)
	threadElem = regexp.MustCompile(`(?s)<threadedComment\b[^>]*?\bref="([A-Z]+[0-9]+)".*?</threadedComment>|<threadedComment\b[^>]*?\bref="([A-Z]+[0-9]+)"[^>]*/>`)
)

// moveThreads applies moves of cells to a threadedComments part: threads
// follow their cells, and go with them when deleted.
func moveThreads(b []byte, moves []threadMove) []byte {
	where := func(ref string) (string, bool) {
		r, c, ok := grid.ParseCell(ref)
		if !ok {
			return ref, true
		}
		for _, m := range moves {
			if r, c, ok = m.move(r, c); !ok {
				return "", false
			}
		}
		return grid.CellName(r, c), true
	}
	b = threadElem.ReplaceAllFunc(b, func(m []byte) []byte {
		g := threadElem.FindSubmatch(m)
		ref := string(g[1])
		if ref == "" {
			ref = string(g[2])
		}
		if _, ok := where(ref); !ok {
			return nil
		}
		return m
	})
	return threadRef.ReplaceAllFunc(b, func(m []byte) []byte {
		g := threadRef.FindSubmatch(m)
		ref, _ := where(string(g[2]))
		return append(append(append([]byte{}, g[1]...), ref...), g[3]...)
	})
}

var inlineText = regexp.MustCompile(`(?s)<is>\s*<t\b[^>]*>(.*?)</t>\s*</is>`)

// retypeErrors turns the given cells, written as inline text, into error
// values.
func retypeErrors(b []byte, cells map[string]bool) []byte {
	return rewriteCells(b, func(ref string) bool { return cells[ref] }, func(m []byte, _ string) []byte {
		head := m[:strings.IndexByte(string(m), '>')+1]
		g := inlineText.FindSubmatch(m)
		if g == nil {
			return m
		}
		head = typeAttr.ReplaceAll(head, nil)
		out := append([]byte{}, head[:len(head)-1]...)
		out = append(out, ` t="e"><v>`...)
		out = append(out, g[1]...)
		return append(out, "</v></c>"...)
	})
}

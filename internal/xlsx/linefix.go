package xlsx

import (
	"archive/zip"
	"bytes"
	"fmt"
	"html"
	"regexp"
	"strconv"
	"strings"

	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
	"github.com/xuri/excelize/v2"
)

// excelize moves conditional formats and data validations wrongly when
// rows or columns are deleted: a range losing its first row moves up a row
// instead of starting at the next one (B2:B4 less row 2 becomes B1:B3, over
// the header), and one reaching row 1 fails the edit. So validations are
// taken off before excelize moves cells and put back moved as Excel moves
// them; and conditional formats are never given to excelize: a file's are
// kept aside when it is opened (excelize cannot give them back whole),
// with the ones the session adds; they move with rows, columns and cells
// as Excel moves them, and are written into the sheets when saving.

// A lineEdit moves the conditional formats and validations of a sheet
// whose lines (or cells) move.
type lineEdit struct {
	// place is where a range goes; false when all its cells go.
	place func(r grid.Range) (grid.Range, bool)
	// first is the first cell of a range that stays, top row first; it
	// is asked only of ranges place keeps.
	first func(r grid.Range) (row, col int)
	// rewrite moves the references of a formula held on the sheet.
	rewrite func(f string) string
}

// rule moves a rule's ranges (an sqref), and gives what moves its
// formulas. Those are relative to the top-left cell of its first range: as
// if they stood in that cell, and the other cells held them shifted as a
// formula filled across. When that cell goes, the formula is the one the
// first cell that stays held. false when the rule's cells all went.
func (le lineEdit) rule(sqref string) (string, func(f string) string, bool) {
	var out []string
	var anchor *grid.Range
	dr, dc, kept := 0, 0, false
	for _, part := range strings.Fields(sqref) {
		r, err := grid.ParseRange(strings.ReplaceAll(part, "$", ""))
		if err != nil {
			out = append(out, part)
			continue
		}
		if anchor == nil {
			anchor = &r
		}
		m, ok := le.place(r)
		if !ok {
			continue
		}
		if !kept {
			row, col := le.first(r)
			dr, dc, kept = row-anchor.R1, col-anchor.C1, true
		}
		out = append(out, m.String())
	}
	return strings.Join(out, " "), func(f string) string { return le.rewrite(formula.Shift(f, dr, dc)) }, kept
}

var (
	sqrefOf   = regexp.MustCompile(`(\ssqref=")([^"]*)(")`)
	cfvoVal   = regexp.MustCompile(`(<cfvo\b[^>]*?\sval=")([^"]*)(")`)
	ruleEl    = regexp.MustCompile(`(?s)<conditionalFormatting\b[^>]*>.*?</conditionalFormatting>|<x14:conditionalFormatting\b[^>]*>.*?</x14:conditionalFormatting>|<dataValidation\b[^>]*/>|<dataValidation\b[^>]*>.*?</dataValidation>`)
	x14CfList = regexp.MustCompile(`(?s)<x14:conditionalFormattings\b[^>]*>(.*?)</x14:conditionalFormattings>`)
	x14IDs    = regexp.MustCompile(`(<x14:id>|<x14:cfRule\b[^>]*?\sid=")(\{[0-9A-Fa-f-]+\})`)
)

// moveRule moves a conditional format's element (main or Excel 2010's) or
// a validation's with a lineEdit; nil when its cells all went.
func (le lineEdit) moveRule(el []byte) []byte {
	fix, gone := le.rewrite, false
	el = withSqref(el, func(sq string) string {
		moved, f, ok := le.rule(sq)
		fix, gone = f, !ok
		return moved
	})
	if gone {
		return nil
	}
	return ruleFormulas(el, fix)
}

// withSqref sets the ranges of a conditional format's element (main or
// Excel 2010's) or a validation's to what fn makes of them.
func withSqref(el []byte, fn func(sqref string) string) []byte {
	g := sqrefOf.FindSubmatchIndex(el)
	if g == nil {
		if g = xmSqrefEl.FindSubmatchIndex(el); g == nil {
			return el
		}
	}
	sq := fn(html.UnescapeString(string(el[g[4]:g[5]])))
	return append(append(append([]byte{}, el[:g[4]]...), sq...), el[g[5]:]...)
}

// ruleFormulas rewrites the formulas of a conditional format's or a
// validation's element: its rules', its scale points', Excel 2010's.
func ruleFormulas(el []byte, fn func(string) string) []byte {
	el = rewriteFormulas(el, fn)
	for _, re := range []*regexp.Regexp{xmF, cfvoVal} {
		el = re.ReplaceAllFunc(el, func(m []byte) []byte {
			g := re.FindSubmatch(m)
			text := html.UnescapeString(string(g[2]))
			out := fn(text)
			if out == text {
				return m
			}
			return append(append(append([]byte{}, g[1]...), xmlEscape(out)...), g[3]...)
		})
	}
	return el
}

// heldValidations are a sheet's validations, off the sheet while its lines
// move.
type heldValidations []*excelize.DataValidation

// holdValidations takes a sheet's validations off.
func (e *Editor) holdValidations(sheet string) (heldValidations, error) {
	dvs, err := e.f.GetDataValidations(sheet)
	if err != nil || len(dvs) == 0 {
		return nil, err
	}
	return dvs, e.f.DeleteDataValidation(sheet)
}

// putBack adds held validations again, moved by le; a validation whose
// cells all went is left out.
func (h heldValidations) putBack(e *Editor, sheet string, le lineEdit) error {
	for _, dv := range h {
		sq, fix, ok := le.rule(dv.Sqref)
		if !ok {
			continue
		}
		c := *dv
		c.Sqref = sq
		if c.Formula1 != "" {
			c.Formula1 = escapeXML(fix(unescapeXML(c.Formula1)))
		}
		if c.Formula2 != "" {
			c.Formula2 = escapeXML(fix(unescapeXML(c.Formula2)))
		}
		if err := e.f.AddDataValidation(sheet, &c); err != nil {
			return err
		}
	}
	return nil
}

// condSection is a sheet's conditional formats as its part holds them: the
// main blocks, and Excel 2010's extension (data bars, icon sets).
type condSection struct {
	blocks [][]byte
	ext    [][]byte // x14:conditionalFormatting elements
}

// holdConds takes the conditional formats out of the given sheets of a
// package, for excelize to open it without them, and returns them by sheet
// name.
func holdConds(data []byte, which map[string]bool) ([]byte, map[string]*condSection, error) {
	held := map[string]*condSection{}
	if len(which) == 0 {
		return data, held, nil
	}
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return data, held, nil // the reader says what is wrong
	}
	p := &pkg{files: map[string]*zip.File{}}
	for _, zf := range zr.File {
		p.files[strings.ToLower(zf.Name)] = zf
	}
	_, _, sheets, err := p.workbookParts()
	if err != nil {
		return data, held, nil
	}
	stripped := map[string][]byte{}
	for name, part := range sheets {
		if !which[name] {
			continue
		}
		b, err := p.read(part)
		if err != nil || !bytes.Contains(b, []byte("conditionalFormatting")) {
			continue
		}
		sec, rest := condsIn(b)
		if sec.empty() {
			continue
		}
		held[name] = sec
		stripped[strings.ToLower(part)] = rest
	}
	if len(held) == 0 {
		return data, held, nil
	}
	var out bytes.Buffer
	zw := zip.NewWriter(&out)
	for _, zf := range zr.File {
		if b, ok := stripped[strings.ToLower(zf.Name)]; ok {
			// Stored: excelize reads it once, now.
			w, err := zw.CreateHeader(&zip.FileHeader{Name: zf.Name, Method: zip.Store})
			if err != nil {
				return nil, nil, err
			}
			if _, err := w.Write(b); err != nil {
				return nil, nil, err
			}
			continue
		}
		if err := copyRaw(zw, zf); err != nil {
			return nil, nil, err
		}
	}
	if err := zw.Close(); err != nil {
		return nil, nil, err
	}
	return out.Bytes(), held, nil
}

// condsIn takes the conditional formats out of a sheet part.
func condsIn(b []byte) (*condSection, []byte) {
	// They follow the cells, which can be most of the part.
	head := 0
	if i := bytes.LastIndex(b, []byte("</sheetData>")); i >= 0 {
		head = i
	}
	t := b[head:]
	// The main blocks first: their rules' own extension lists hold ids.
	sec := &condSection{blocks: cfBlockEl.FindAll(t, -1)}
	t = cfBlockEl.ReplaceAll(t, nil)
	for _, m := range x14CfList.FindAllSubmatch(t, -1) {
		sec.ext = append(sec.ext, x14CfEl.FindAll(m[1], -1)...)
	}
	t = x14CfList.ReplaceAll(t, []byte("<x14:conditionalFormattings></x14:conditionalFormattings>"))
	t = emptyExtLst.ReplaceAll(x14CondsExt.ReplaceAll(t, nil), nil)
	return sec, append(b[:head:head], t...)
}

// each runs fn over the section's elements, main then Excel 2010's,
// dropping those fn returns nil for.
func (sec *condSection) each(fn func(el []byte) []byte) {
	for _, list := range []*[][]byte{&sec.blocks, &sec.ext} {
		var out [][]byte
		for _, el := range *list {
			if el = fn(el); el != nil {
				out = append(out, el)
			}
		}
		*list = out
	}
}

// formulas rewrites the formulas of every rule.
func (sec *condSection) formulas(fn func(string) string) {
	sec.each(func(el []byte) []byte { return ruleFormulas(el, fn) })
}

// empty tells whether the section holds no format.
func (sec *condSection) empty() bool { return len(sec.blocks)+len(sec.ext) == 0 }

// xml is the main blocks, for reading (see condFormats).
func (sec *condSection) xml() []byte { return bytes.Join(sec.blocks, nil) }

// copy is a section that changes apart from sec.
func (sec *condSection) copy() *condSection {
	c := &condSection{}
	for _, b := range sec.blocks {
		c.blocks = append(c.blocks, append([]byte{}, b...))
	}
	for _, b := range sec.ext {
		c.ext = append(c.ext, append([]byte{}, b...))
	}
	return c
}

// condsOf is a sheet's section, made when it has none.
func (e *Editor) condsOf(sheet string) *condSection {
	sec := e.patches.conds[sheet]
	if sec == nil {
		sec = &condSection{}
		e.patches.conds[sheet] = sec
	}
	return sec
}

// addConds adds the rules opts make over r to a sheet's conditional
// formats, after its others. excelize draws them, on a sheet of its own.
func (e *Editor) addConds(sheet string, r grid.Range, opts []excelize.ConditionalFormatOptions) error {
	sf := excelize.NewFile()
	defer sf.Close()
	if err := sf.SetConditionalFormat("Sheet1", r.String(), opts); err != nil {
		return err
	}
	buf, err := sf.WriteToBuffer()
	if err != nil {
		return err
	}
	data := buf.Bytes()
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return err
	}
	var part []byte
	for _, zf := range zr.File {
		if zf.Name == "xl/worksheets/sheet1.xml" {
			rc, err := zf.Open()
			if err != nil {
				return err
			}
			var b bytes.Buffer
			_, err = b.ReadFrom(rc)
			rc.Close()
			if err != nil {
				return err
			}
			part = b.Bytes()
		}
	}
	made, _ := condsIn(part)
	blocks, ext := made.blocks, made.ext
	if made.empty() {
		return fmt.Errorf("conditional format over %s not made", r)
	}
	sec := e.condsOf(sheet)
	// After the sheet's others, which come first: priorities below theirs
	// (orderRules makes them distinct), and ids of their own.
	top := 0
	used := map[string]bool{}
	for _, el := range append(append([][]byte{}, sec.blocks...), sec.ext...) {
		for _, m := range priorityAt.FindAllSubmatch(el, -1) {
			if p, _ := strconv.Atoi(string(m[1])); p > top {
				top = p
			}
		}
		for _, m := range x14IDs.FindAllSubmatch(el, -1) {
			used[strings.ToUpper(string(m[2]))] = true
		}
	}
	for _, el := range e.patches.conds {
		for _, b := range append(append([][]byte{}, el.blocks...), el.ext...) {
			for _, m := range x14IDs.FindAllSubmatch(b, -1) {
				used[strings.ToUpper(string(m[2]))] = true
			}
		}
	}
	ids := map[string]string{}
	next := 0
	renumber := func(el []byte) []byte {
		el = priorityAt.ReplaceAllFunc(el, func(m []byte) []byte {
			p, _ := strconv.Atoi(string(priorityAt.FindSubmatch(m)[1]))
			return []byte(` priority="` + strconv.Itoa(top+p) + `"`)
		})
		return x14IDs.ReplaceAllFunc(el, func(m []byte) []byte {
			g := x14IDs.FindSubmatch(m)
			id := strings.ToUpper(string(g[2]))
			if _, ok := ids[id]; !ok {
				for {
					next++
					fresh := fmt.Sprintf("{5EE75000-0000-4000-8000-%012X}", next)
					if !used[fresh] {
						ids[id], used[fresh] = fresh, true
						break
					}
				}
			}
			return append(append([]byte{}, g[1]...), ids[id]...)
		})
	}
	for _, el := range blocks {
		sec.blocks = append(sec.blocks, renumber(el))
	}
	for _, el := range ext {
		sec.ext = append(sec.ext, renumber(el))
	}
	return nil
}

// dropConds removes a sheet's conditional formats whose ranges drop
// accepts, with their Excel 2010 parts.
func (e *Editor) dropConds(sheet string, drop func(sqref string) bool) {
	if sec := e.patches.conds[sheet]; sec != nil {
		sec.each(func(el []byte) []byte {
			gone := false
			withSqref(el, func(sq string) string {
				gone = drop(sq)
				return sq
			})
			if gone {
				return nil
			}
			return el
		})
	}
}

const (
	x14CondExt = `<ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">`
	xmNS       = `xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"`
)

// writeConds writes a sheet's conditional formats into its part, where the
// schema puts them.
func writeConds(b []byte, sec *condSection) []byte {
	b = writeCondsExt(b, sec)
	if len(sec.blocks) > 0 {
		at := len(b)
		for _, next := range []string{"<dataValidations", "<hyperlinks", "<printOptions", "<pageMargins", "<pageSetup",
			"<headerFooter", "<rowBreaks", "<colBreaks", "<customProperties", "<cellWatches", "<ignoredErrors",
			"<smartTags", "<drawing", "<legacyDrawing", "<picture", "<oleObjects", "<controls", "<webPublishItems",
			"<tableParts", "<extLst", "</worksheet>"} {
			if i := bytes.Index(b, []byte(next)); i >= 0 && i < at {
				at = i
			}
		}
		b = append(append(append([]byte{}, b[:at]...), bytes.Join(sec.blocks, nil)...), b[at:]...)
	}
	return b
}

// writeCondsExt writes the Excel 2010 part of a sheet's conditional
// formats into the sheet's extension list (not a rule's or a view's),
// first, where Excel puts it.
func writeCondsExt(b []byte, sec *condSection) []byte {
	if len(sec.ext) == 0 {
		return b
	}
	// The extension's elements name the xm namespace, which the root may
	// not declare: each declares it.
	var inner []byte
	for _, el := range sec.ext {
		if !bytes.Contains(el, []byte(xmNS)) {
			el = bytes.Replace(el, []byte("<x14:conditionalFormatting"), []byte("<x14:conditionalFormatting "+xmNS), 1)
		}
		inner = append(inner, el...)
	}
	ext := append(append(append([]byte(x14CondExt), "<x14:conditionalFormattings>"...), inner...), "</x14:conditionalFormattings></ext>"...)
	end := bytes.LastIndex(b, []byte("</worksheet>"))
	if end < 0 {
		return b
	}
	if bytes.HasSuffix(bytes.TrimSpace(b[:end]), []byte("</extLst>")) {
		if i := bytes.LastIndex(b[:end], []byte("<extLst>")); i >= 0 {
			i += len("<extLst>")
			return append(append(append([]byte{}, b[:i]...), ext...), b[i:]...)
		}
	}
	return append(append(append(append([]byte{}, b[:end]...), "<extLst>"...), ext...), append([]byte("</extLst>"), b[end:]...)...)
}

// stay is the lineEdit of a sheet whose lines stay where they are: only
// formulas move (reading a sheet whose lines moved).
func stay(rewrite func(string) string) lineEdit {
	return lineEdit{
		place:   func(r grid.Range) (grid.Range, bool) { return r, true },
		first:   func(r grid.Range) (int, int) { return r.R1, r.C1 },
		rewrite: rewrite,
	}
}

// namesNow brings the sheet names a formula held by excelize uses up to
// the session's renamed and deleted sheets, which such formulas take when
// saving (see applySheetEdits).
func (e *Editor) namesNow(f string) string {
	for _, ed := range e.patches.sheetEdits {
		if ed.to != "" {
			f = formula.RenameSheet(f, ed.from, ed.to)
		} else {
			f = formula.DropSheet(f, ed.from)
		}
	}
	return f
}

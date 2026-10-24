package xlsx

import (
	"bufio"
	"bytes"
	"fmt"
	"html"
	"io"
	"regexp"
	"sort"
	"strconv"
	"strings"

	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
	"github.com/xuri/excelize/v2"
)

// sheetJob is what rewriting a worksheet part changes.
type sheetJob struct {
	sheet   string
	edits   map[[2]int]*streamCell // cells written, where they are now
	styles  map[[2]int]string      // the style each written cell takes, where it changes
	results map[string]grid.Cell   // formula results to store, by cell (now)
	expand  bool                   // write filled-down formulas out each with its formula
	// formulaAt is a cell's formula now, for writing filled-down ones out.
	formulaAt func(r, c int) string
	// rewrite brings a formula of the sheet up to the rows as they are
	// now (nil: no rows moved anywhere).
	rewrite func(f string) string
	// rowNow is where a row of the sheet is now, false when deleted; area
	// does the same for a range (grow: over rows appended to a table).
	// Both nil when the sheet's rows did not change.
	rowNow func(r int) (int, bool)
	area   func(r grid.Range, grow bool) (grid.Range, bool)
	// rules move conditional formats and validations, edit by edit, as
	// the full editor moves them (see lineEdit).
	rules     []lineEdit
	panes     func(head []byte) []byte
	dimension *grid.Range
	// filterGone tells whether the sheet's filter, over a range as
	// opened, goes (see StreamEditor.filterGone); its rows then show.
	filterGone func(r grid.Range) bool
	unhide     *grid.Range // rows, where they are now, a filter that went hid
}

func (j *sheetJob) needed() bool {
	return len(j.edits) > 0 || len(j.results) > 0 || j.expand || j.rewrite != nil || j.rowNow != nil || j.dimension != nil
}

// window reads a part through a buffer it can look back into.
type window struct {
	r   *bufio.Reader
	buf []byte
	pos int
	eof bool
}

// fill reads on, straight into the buffer; false at the end of the part.
func (w *window) fill() bool {
	if w.eof {
		return false
	}
	if cap(w.buf)-len(w.buf) < 256<<10 {
		nb := make([]byte, len(w.buf), 2*cap(w.buf)+1<<20)
		copy(nb, w.buf)
		w.buf = nb
	}
	n, err := w.r.Read(w.buf[len(w.buf):cap(w.buf)])
	w.buf = w.buf[:len(w.buf)+n]
	if err != nil {
		w.eof = true
	}
	return n > 0 || !w.eof
}

// compact drops what was passed. Indexes into buf held across it no longer
// hold, so it is called only between rows.
func (w *window) compact() {
	if w.pos > 4<<20 {
		w.buf = append(w.buf[:0], w.buf[w.pos:]...)
		w.pos = 0
	}
}

// find returns the index in buf of s at or after from, reading on as
// needed; -1 when the part ends first.
func (w *window) find(s string, from int) int {
	for {
		if i := bytes.Index(w.buf[from:], []byte(s)); i >= 0 {
			return from + i
		}
		// The start of s may be at the end of what was read.
		from = max(from, len(w.buf)-len(s)+1)
		if !w.fill() {
			return -1
		}
	}
}

// need makes sure n bytes from pos are read, if the part has them.
func (w *window) need(n int) bool {
	for len(w.buf)-w.pos < n {
		if !w.fill() {
			return false
		}
	}
	return true
}

var (
	rowNumAttr  = regexp.MustCompile(`\sr="(\d+)"`)
	hiddenAttr  = regexp.MustCompile(`\shidden="(?:1|true)"`)
	spansAttr   = regexp.MustCompile(`\sspans="[^"]*"`)
	cellRefAttr = regexp.MustCompile(`\sr="([A-Z]+)(\d+)"`)
	styleAttr   = regexp.MustCompile(`\ss="(\d+)"`)
	dimensionEl = regexp.MustCompile(`<dimension\b[^>]*\bref="[^"]*"`)
	colEl       = regexp.MustCompile(`<col\b[^>]*>`)
	sharedF     = regexp.MustCompile(`^<f\b[^>]*\st="shared"`)
	fRefAttr    = regexp.MustCompile(`(\sref=")([^"]*)(")`)
	fText       = regexp.MustCompile(`(?s)^(<f\b[^>]*>)(.*)(</f>)$`)
	xmF         = regexp.MustCompile(`(<xm:f>)([^<]*)(</xm:f>)`)
)

// rewriteSheet passes a worksheet part through, row by row, making the
// job's changes.
func rewriteSheet(r io.Reader, out io.Writer, j *sheetJob) error {
	w := &window{r: bufio.NewReaderSize(r, 1<<20)}
	bw := bufio.NewWriterSize(out, 1<<20)
	defer bw.Flush()

	// Head: up to the rows.
	at := w.find("<sheetData", 0)
	if at < 0 {
		return fmt.Errorf("no sheetData")
	}
	end := w.find(">", at)
	if end < 0 {
		return fmt.Errorf("sheetData unclosed")
	}
	head := append([]byte{}, w.buf[:at]...)
	open := w.buf[at : end+1]
	selfClosing := bytes.HasSuffix(open, []byte("/>"))
	cols := colStyles(head)
	if j.dimension != nil {
		d := *j.dimension
		if d.Empty() {
			d = grid.Range{R1: 1, C1: 1, R2: 1, C2: 1}
		}
		head = dimensionEl.ReplaceAll(head, []byte(`<dimension ref="`+d.String()+`"`))
	}
	if j.panes != nil {
		head = j.panes(head)
	}
	bw.Write(head)
	bw.WriteString("<sheetData>")
	w.pos = end + 1

	// Rows written to that the part does not have, by row.
	byRow := map[int]map[int]*streamCell{}
	var editRows []int
	for pos, c := range j.edits {
		if byRow[pos[0]] == nil {
			byRow[pos[0]] = map[int]*streamCell{}
			editRows = append(editRows, pos[0])
		}
		byRow[pos[0]][pos[1]] = c
	}
	sort.Ints(editRows)
	results := map[int]map[int]grid.Cell{}
	for ref, v := range j.results {
		rr, cc, ok := grid.ParseCell(ref)
		if !ok {
			continue
		}
		if results[rr] == nil {
			results[rr] = map[int]grid.Cell{}
		}
		results[rr][cc] = v
	}
	var stale []grid.Range // multi-cell arrays whose other cells' values no longer hold
	next := 0              // into editRows
	lastRow := 0
	newRow := func(r int) {
		bw.WriteString(`<row r="` + strconv.Itoa(r) + `">`)
		var cs []int
		for c := range byRow[r] {
			cs = append(cs, c)
		}
		sort.Ints(cs)
		for _, c := range cs {
			st := j.styles[[2]int{r, c}]
			if st == "" {
				st = styleOfCol(cols, c)
			}
			bw.Write(j.newCell(r, c, byRow[r][c], st, results[r]))
		}
		bw.WriteString("</row>")
	}

	if !selfClosing {
		for {
			w.compact()
			// Between rows: blanks, then a row or the end.
			for {
				if !w.need(1) {
					return fmt.Errorf("sheetData unclosed")
				}
				if c := w.buf[w.pos]; c == ' ' || c == '\n' || c == '\r' || c == '\t' {
					w.pos++
					continue
				}
				break
			}
			if !w.need(len("</sheetData>")) {
				return fmt.Errorf("sheetData unclosed")
			}
			if bytes.HasPrefix(w.buf[w.pos:], []byte("</sheetData>")) {
				w.pos += len("</sheetData>")
				break
			}
			if !bytes.HasPrefix(w.buf[w.pos:], []byte("<row")) {
				return fmt.Errorf("unexpected %q in sheetData", string(w.buf[w.pos:min(w.pos+20, len(w.buf))]))
			}
			tagEnd := w.find(">", w.pos)
			if tagEnd < 0 {
				return fmt.Errorf("row unclosed")
			}
			rowEnd := tagEnd + 1
			if w.buf[tagEnd-1] != '/' {
				close := w.find("</row>", tagEnd)
				if close < 0 {
					return fmt.Errorf("row unclosed")
				}
				rowEnd = close + len("</row>")
			}
			row := w.buf[w.pos:rowEnd]
			start := row[:tagEnd-w.pos+1]
			rOrig := lastRow + 1
			if m := rowNumAttr.FindSubmatch(start); m != nil {
				rOrig, _ = strconv.Atoi(string(m[1]))
			}
			lastRow = rOrig
			w.pos = rowEnd
			rNow := rOrig
			if j.rowNow != nil {
				var ok bool
				if rNow, ok = j.rowNow(rOrig); !ok {
					continue // deleted
				}
			}
			for next < len(editRows) && editRows[next] < rNow {
				newRow(editRows[next])
				next++
			}
			edits := map[int]*streamCell(nil)
			if next < len(editRows) && editRows[next] == rNow {
				edits = byRow[rNow]
				next++
			}
			res := results[rNow]
			if j.unhide != nil && j.unhide.R1 <= rNow && rNow <= j.unhide.R2 && bytes.Contains(start, []byte("hidden=")) {
				// Shown again with the filter that hid it gone.
				shown := hiddenAttr.ReplaceAll(start, nil)
				row = append(append([]byte{}, shown...), row[len(start):]...)
				start = shown
			}
			if rNow == rOrig && edits == nil && res == nil && !j.expand && len(stale) == 0 &&
				(j.rewrite == nil || !bytes.Contains(row, []byte("<f"))) {
				bw.Write(row)
				continue
			}
			bw.Write(j.row(row, len(start), rOrig, rNow, edits, res, cols, &stale))
		}
	}
	for ; next < len(editRows); next++ {
		newRow(editRows[next])
	}
	bw.WriteString("</sheetData>")

	// Tail: what follows the rows.
	for w.fill() {
	}
	bw.Write(j.rewriteTail(w.buf[w.pos:]))
	return nil
}

// row rewrites one row: its number, its cells' references, formulas and
// results, and the cells written into it.
func (j *sheetJob) row(row []byte, startLen, rOrig, rNow int, edits map[int]*streamCell, res map[int]grid.Cell, cols [][3]int, stale *[]grid.Range) []byte {
	start := row[:startLen]
	var out bytes.Buffer
	tag := start
	if rNow != rOrig {
		tag = rowNumAttr.ReplaceAll(tag, []byte(` r="`+strconv.Itoa(rNow)+`"`))
	}
	if edits != nil {
		tag = spansAttr.ReplaceAll(tag, nil)
	}
	rowStyle := ""
	if strings.Contains(string(start), `customFormat="1"`) || strings.Contains(string(start), `customFormat="true"`) {
		if m := styleAttr.FindSubmatch(start); m != nil {
			rowStyle = string(m[1])
		}
	}
	selfClosing := bytes.HasSuffix(start, []byte("/>"))
	if selfClosing {
		if edits == nil {
			return tag
		}
		tag = append(append([]byte{}, tag[:len(tag)-2]...), '>')
	}
	out.Write(tag)
	written := map[int]bool{}
	writeNew := func(upTo int) {
		var cs []int
		for c := range edits {
			if !written[c] && c < upTo {
				cs = append(cs, c)
			}
		}
		sort.Ints(cs)
		for _, c := range cs {
			written[c] = true
			st := j.styles[[2]int{rNow, c}]
			if st == "" {
				st = rowStyle
			}
			if st == "" {
				st = styleOfCol(cols, c)
			}
			out.Write(j.newCell(rNow, c, edits[c], st, res))
		}
	}
	if !selfClosing {
		body := row[startLen : len(row)-len("</row>")]
		col := 0
		for i := 0; i < len(body); {
			k := bytes.Index(body[i:], []byte("<c"))
			if k < 0 {
				out.Write(body[i:])
				break
			}
			k += i
			if k+2 < len(body) && body[k+2] != ' ' && body[k+2] != '>' && body[k+2] != '/' {
				out.Write(body[i : k+2])
				i = k + 2
				continue
			}
			out.Write(body[i:k])
			te := bytes.IndexByte(body[k:], '>')
			if te < 0 {
				out.Write(body[k:])
				break
			}
			te += k
			ce := te + 1
			if body[te-1] != '/' {
				if x := bytes.Index(body[te:], []byte("</c>")); x >= 0 {
					ce = te + x + len("</c>")
				}
			}
			cell := body[k:ce]
			i = ce
			col++
			if m := cellRefAttr.FindSubmatch(cell[:te-k+1]); m != nil {
				if c, ok := grid.ColIndex(string(m[1])); ok {
					col = c
				}
			}
			writeNew(col)
			if e, ok := edits[col]; ok {
				written[col] = true
				st := j.styles[[2]int{rNow, col}]
				if st == "" {
					if m := styleAttr.FindSubmatch(cell[:te-k+1]); m != nil {
						st = string(m[1])
					}
				}
				out.Write(j.newCell(rNow, col, e, st, res))
				continue
			}
			out.Write(j.cell(cell, te-k+1, rOrig, rNow, col, res, stale))
		}
	}
	writeNew(grid.MaxCols + 1)
	out.WriteString("</row>")
	return out.Bytes()
}

// cell rewrites a cell the session did not write: its reference, its
// formula (filled-down ones written out, references following moved
// rows), and its stored result.
func (j *sheetJob) cell(cell []byte, startLen, rOrig, rNow, col int, res map[int]grid.Cell, stale *[]grid.Range) []byte {
	if rNow != rOrig {
		start := cellRefAttr.ReplaceAll(cell[:startLen], []byte(` r="`+grid.CellName(rNow, col)+`"`))
		cell = append(append([]byte{}, start...), cell[startLen:]...)
	}
	f := formulaEl.Find(cell)
	if f != nil {
		g := f
		switch {
		case j.expand && sharedF.Match(f):
			text := j.formulaAt(rNow, col)
			g = append(append([]byte("<f>"), xmlEscape(formula.ToFile(text))...), "</f>"...)
		default:
			if j.rewrite != nil {
				if m := fText.FindSubmatch(f); m != nil && len(m[2]) > 0 {
					text := html.UnescapeString(string(m[2]))
					if out := j.rewrite(text); out != text {
						g = append(append(append([]byte{}, m[1]...), xmlEscape(out)...), m[3]...)
					}
				}
			}
			if j.area != nil {
				g = fRefAttr.ReplaceAllFunc(g, func(a []byte) []byte {
					sub := fRefAttr.FindSubmatch(a)
					rg, err := grid.ParseRange(string(sub[2]))
					if err != nil {
						return a
					}
					if moved, ok := j.area(rg, false); ok {
						return []byte(string(sub[1]) + moved.String() + string(sub[3]))
					}
					return a
				})
			}
		}
		if !bytes.Equal(f, g) {
			cell = bytes.Replace(cell, f, g, 1)
		}
	}
	if v, ok := res[col]; ok && f != nil {
		out, array := storeResult(cell, v)
		if array != nil {
			*stale = append(*stale, *array)
		}
		return out
	}
	if f == nil && len(*stale) > 0 {
		for _, a := range *stale {
			if a.Contains(rNow, col) {
				head := cell[:startLen]
				if head[len(head)-2] == '/' {
					return cell
				}
				body := cell[startLen : len(cell)-len("</c>")]
				return append(append(typeAttr.ReplaceAll(head, nil), valueElem.ReplaceAll(body, nil)...), "</c>"...)
			}
		}
	}
	return cell
}

// newCell writes a cell the session wrote: style st ("" for none), and a
// formula's result when it was computed.
func (j *sheetJob) newCell(r, c int, e *streamCell, st string, res map[int]grid.Cell) []byte {
	var b bytes.Buffer
	b.WriteString(`<c r="` + grid.CellName(r, c) + `"`)
	if st != "" && st != "0" {
		b.WriteString(` s="` + st + `"`)
	}
	v := e.v
	switch {
	case v.Formula != "":
		b.WriteString("><f>")
		b.Write(xmlEscape(formula.ToFile(v.Formula)))
		b.WriteString("</f></c>")
		if x, ok := res[c]; ok {
			out, _ := storeResult(b.Bytes(), x)
			return out
		}
		return b.Bytes()
	case v.Kind == grid.Text:
		b.WriteString(` t="inlineStr"><is><t`)
		if strings.TrimSpace(v.Text) != v.Text {
			b.WriteString(` xml:space="preserve"`)
		}
		b.WriteString(">")
		b.Write(escapeOOXML(v.Text))
		b.WriteString("</t></is></c>")
	case v.Kind == grid.Number:
		b.WriteString("><v>" + strconv.FormatFloat(v.Num, 'g', -1, 64) + "</v></c>")
	case v.Kind == grid.Bool:
		b.WriteString(` t="b"><v>` + map[bool]string{true: "1", false: "0"}[v.Bool] + "</v></c>")
	default:
		if st == "" || st == "0" {
			return nil // nothing to keep
		}
		b.WriteString("/>")
	}
	return b.Bytes()
}

var ooxmlEscapes = regexp.MustCompile(`_x[0-9A-Fa-f]{4}_`)

// escapeOOXML writes text for a cell: XML-escaped, with what Excel would
// read as an escape (_x0041_) escaped itself, and control characters,
// which XML cannot hold, as Excel writes them.
func escapeOOXML(s string) []byte { return xmlEscape(ooxmlText(s)) }

// ooxmlText escapes text as Excel does, before XML escaping (see
// escapeOOXML).
func ooxmlText(s string) string {
	s = ooxmlEscapes.ReplaceAllStringFunc(s, func(m string) string { return "_x005F" + m })
	if strings.IndexFunc(s, func(r rune) bool { return r < 0x20 && r != '\t' && r != '\n' && r != '\r' }) < 0 {
		return s
	}
	var b strings.Builder
	for _, r := range s {
		if r < 0x20 && r != '\t' && r != '\n' && r != '\r' {
			fmt.Fprintf(&b, "_x%04X_", r)
			continue
		}
		b.WriteRune(r)
	}
	return b.String()
}

// setText writes text to a cell through excelize, which keeps text in the
// shared strings as given: Excel would read _x0041_ in it as an escape
// (A), and refuses control characters. So text is escaped as Excel
// escapes it first (excelize v2.11 escapes inline text, not shared).
func setText(f *excelize.File, sheet, cell, text string) error {
	return f.SetCellStr(sheet, cell, ooxmlText(text))
}

// colStyles reads the column styles from a sheet's head: min, max, style.
func colStyles(head []byte) [][3]int {
	var out [][3]int
	for _, m := range colEl.FindAll(head, -1) {
		get := func(name string) int {
			re := regexp.MustCompile(`\s` + name + `="(\d+)"`)
			if g := re.FindSubmatch(m); g != nil {
				n, _ := strconv.Atoi(string(g[1]))
				return n
			}
			return 0
		}
		if st := get("style"); st > 0 {
			out = append(out, [3]int{get("min"), get("max"), st})
		}
	}
	return out
}

func styleOfCol(cols [][3]int, c int) string {
	for _, x := range cols {
		if c >= x[0] && c <= x[1] {
			return strconv.Itoa(x[2])
		}
	}
	return ""
}

var (
	mergeCellEl  = regexp.MustCompile(`<mergeCell\b[^>]*/>|<mergeCell\b[^>]*>\s*</mergeCell>`)
	hyperlinkEl  = regexp.MustCompile(`<hyperlink\b[^>]*/>|(?s)<hyperlink\b[^>]*>.*?</hyperlink>`)
	autoFilterEl = regexp.MustCompile(`<autoFilter\b[^>]*>`)
	filterEl     = regexp.MustCompile(`(?s)<autoFilter\b[^>]*/>|<autoFilter\b[^>]*>.*?</autoFilter>`)
	cfBlockEl    = regexp.MustCompile(`(?s)<conditionalFormatting\b[^>]*>.*?</conditionalFormatting>`)
	x14CfEl      = regexp.MustCompile(`(?s)<x14:conditionalFormatting\b[^>]*>.*?</x14:conditionalFormatting>`)
	refOrSqref   = regexp.MustCompile(`(\s(?:ref|sqref)=")([^"]*)(")`)
	xmSqrefEl    = regexp.MustCompile(`(<xm:sqref>)([^<]*)(</xm:sqref>)`)
	countAttr    = regexp.MustCompile(`\scount="\d+"`)
)

// rewriteTail brings what follows a sheet's rows up to its rows as they
// are now: merged cells, conditional formats and validations (ranges and
// formulas), hyperlinks and the filter. What lay in deleted rows goes.
func (j *sheetJob) rewriteTail(t []byte) []byte {
	if j.rewrite == nil && j.area == nil {
		return t
	}
	// Conditional formats and validations move edit by edit, as the full
	// editor moves them (see lineEdit); other formulas follow their
	// references.
	rewrite := j.rewrite
	if rewrite == nil {
		rewrite = func(f string) string { return f }
	}
	var out []byte
	last := 0
	for _, loc := range ruleEl.FindAllIndex(t, -1) {
		out = append(out, ruleFormulas(t[last:loc[0]], rewrite)...)
		el := t[loc[0]:loc[1]]
		for _, le := range j.rules {
			if el = le.moveRule(el); el == nil {
				break
			}
		}
		out = append(out, el...)
		last = loc[1]
	}
	t = append(out, ruleFormulas(t[last:], rewrite)...)
	if j.area == nil {
		return emptyExtLst.ReplaceAll(x14CondsExt.ReplaceAll(t, nil), nil)
	}
	// ranges moves the ranges of one attribute value; false when all went.
	ranges := func(v string, grow bool) (string, bool) {
		var out []string
		for _, part := range strings.Fields(v) {
			rg, err := grid.ParseRange(strings.ReplaceAll(part, "$", ""))
			if err != nil {
				out = append(out, part)
				continue
			}
			if moved, ok := j.area(rg, grow); ok {
				out = append(out, moved.String())
			}
		}
		return strings.Join(out, " "), len(out) > 0
	}
	element := func(re *regexp.Regexp, grow, drop bool) {
		t = re.ReplaceAllFunc(t, func(el []byte) []byte {
			gone := false
			el = refOrSqref.ReplaceAllFunc(el, func(a []byte) []byte {
				g := refOrSqref.FindSubmatch(a)
				v, ok := ranges(string(g[2]), grow)
				if !ok {
					gone = true
					return a
				}
				return []byte(string(g[1]) + v + string(g[3]))
			})
			el = xmSqrefEl.ReplaceAllFunc(el, func(a []byte) []byte {
				g := xmSqrefEl.FindSubmatch(a)
				v, ok := ranges(string(g[2]), grow)
				if !ok {
					gone = true
					return a
				}
				return []byte(string(g[1]) + v + string(g[3]))
			})
			if gone && drop {
				return nil
			}
			return el
		})
	}
	element(mergeCellEl, false, true)
	element(hyperlinkEl, false, true)
	if j.filterGone != nil {
		t = filterEl.ReplaceAllFunc(t, func(el []byte) []byte {
			if g := refOrSqref.FindSubmatch(el); g != nil {
				if rg, err := grid.ParseRange(string(g[2])); err == nil && j.filterGone(rg) {
					return nil
				}
			}
			return el
		})
	}
	element(autoFilterEl, false, false)
	t = x14CondsExt.ReplaceAll(t, nil)
	t = recount(t, "mergeCells", "<mergeCell ")
	t = recount(t, "dataValidations", "<dataValidation ")
	for _, empty := range []string{`(?s)<mergeCells\b[^>]*>\s*</mergeCells>`, `(?s)<dataValidations\b[^>]*>\s*</dataValidations>`,
		`(?s)<hyperlinks>\s*</hyperlinks>`} {
		t = regexp.MustCompile(empty).ReplaceAll(t, nil)
	}
	return emptyExtLst.ReplaceAll(t, nil)
}

// recount sets the count of a list element to the items it holds.
func recount(t []byte, list, item string) []byte {
	re := regexp.MustCompile(`(?s)(<` + list + `\b[^>]*>)(.*?)(</` + list + `>)`)
	return re.ReplaceAllFunc(t, func(m []byte) []byte {
		g := re.FindSubmatch(m)
		n := bytes.Count(g[2], []byte(item))
		open := countAttr.ReplaceAll(g[1], []byte(` count="`+strconv.Itoa(n)+`"`))
		return append(append(append([]byte{}, open...), g[2]...), g[3]...)
	})
}

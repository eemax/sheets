package xlsx

import (
	"archive/zip"
	"bufio"
	"bytes"
	"compress/flate"
	"crypto/sha256"
	"encoding/hex"
	"encoding/xml"
	"fmt"
	"html"
	"io"
	"path"
	"regexp"
	"sort"
	"strconv"
	"strings"

	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/words"
)

// Saving a streaming edit: the results to store are worked out first, then
// the package is written part by part. A sheet with edits, or with
// formulas the edits rewrite, passes through rewriteSheet; the parts that
// refer to cells (names, tables, charts, notes) are rewritten as needed;
// the rest are copied as they are, without recompressing them.

// Save writes the workbook, or with dryRun only works out the outcome.
func (s *StreamEditor) Save(dryRun bool) error {
	buf, err := s.Render()
	if err != nil || dryRun {
		return err
	}
	return writeAtomic(s.path, buf)
}

// Render works out the file as saved.
func (s *StreamEditor) Render() ([]byte, error) {
	res, err := computeResults(s.calcState())
	if err != nil {
		return nil, err
	}
	s.res = &res
	s.warnings = append(s.warnings, res.warnings...)
	s.noteBrokenRefs()

	// Cell formats: number formats cells are given, and the recorded
	// styling laid over cells written that were empty (see
	// Editor.restyleWritten), each over the cell's own format. So the
	// cells' formats are read first.
	stylesPart, themePart := "", ""
	if rels, err := s.p.rels(path.Join(path.Dir(s.wbPath), "_rels", path.Base(s.wbPath)+".rels"), path.Dir(s.wbPath)); err == nil {
		for _, r := range rels {
			switch {
			case strings.HasSuffix(r.Type, "/styles"):
				stylesPart = r.Target
			case strings.HasSuffix(r.Type, "/theme"):
				themePart = r.Target
			}
		}
	}
	newStyle, sb, err := s.cellFormats(stylesPart, themePart)
	if err != nil {
		return nil, err
	}
	if sb != nil {
		defer sb.close()
	}
	var stylesXML []byte
	if sb != nil && sb.changed {
		if stylesXML, err = sb.styles(); err != nil {
			return nil, err
		}
	}

	partSheet := map[string]string{} // part → sheet
	for name, part := range s.parts {
		partSheet[part] = name
	}
	moved := map[string]bool{} // lower-cased sheets whose rows moved
	for _, op := range s.ops {
		if op.kind != 'x' {
			moved[strings.ToLower(op.sheet)] = true
		}
	}
	rowsMoved := len(moved) > 0
	// Parts attached to sheets whose rows moved: notes, their drawings,
	// comment threads.
	attached := map[string]func([]byte) []byte{}
	for name, part := range s.parts {
		if !moved[strings.ToLower(name)] {
			continue
		}
		name := name
		rowNow := func(r int) (int, bool) { return s.rowNow(name, r) }
		parts := s.p.sheetParts(part)
		if parts.comments != "" {
			comments := parts.comments
			attached[comments] = func(b []byte) []byte { return moveComments(b, rowNow) }
			if vml := s.legacyDrawing(part); vml != "" {
				attached[vml] = func(b []byte) []byte {
					cb, err := s.p.read(comments)
					if err != nil {
						return b
					}
					return fitNoteShapes(shiftNoteShapes(b, rowNow), moveComments(cb, rowNow))
				}
			}
		}
		if parts.threads != "" {
			attached[parts.threads] = func(b []byte) []byte {
				return moveThreads(b, []threadMove{{sheet: name, move: func(r, c int) (int, int, bool) {
					r, ok := s.rowNow(name, r)
					return r, c, ok
				}}})
			}
		}
	}
	tables := map[string][]byte{}
	if len(s.ops) > 0 {
		rw := s.rangeRewrites()
		for _, t := range s.tables {
			if b, ok := tablePartNow(t, rw); ok {
				tables[strings.ToLower(t.part)] = b
			}
		}
	}

	valueParts, err := s.p.valueParts()
	if err != nil {
		return nil, err
	}
	isValue := map[string]bool{}
	for _, n := range valueParts {
		isValue[strings.ToLower(strings.TrimPrefix(n, "/"))] = true
	}
	sums := map[string][32]byte{}
	customPart, hasCustom := s.customProps()

	var out bytes.Buffer
	zw := zip.NewWriter(&out)
	zw.RegisterCompressor(zip.Deflate, func(w io.Writer) (io.WriteCloser, error) { return flate.NewWriter(w, flate.BestSpeed) })
	write := func(name string, b []byte) error {
		w, err := zw.CreateHeader(&zip.FileHeader{Name: name, Method: zip.Deflate})
		if err != nil {
			return err
		}
		_, err = w.Write(b)
		return err
	}
	for _, zf := range s.zr.File {
		name := zf.Name
		lower := strings.ToLower(name)
		if hasCustom && lower == strings.ToLower(customPart) {
			continue // written last, with the stamp
		}
		if sheet, ok := partSheet[name]; ok {
			k := strings.ToLower(sheet)
			job := &sheetJob{sheet: sheet, edits: s.edits[k], styles: newStyle[k], results: s.res.results[k], expand: s.expand[k]}
			if rowsMoved {
				job.rewrite = func(f string) string { return s.rewrite(f, sheet) }
			}
			if rowsMoved || s.extends(sheet) {
				job.rules = s.ruleEdits(sheet)
			}
			if moved[k] || s.extends(sheet) {
				job.rowNow = func(r int) (int, bool) { return s.rowNow(sheet, r) }
				job.area = func(r grid.Range, grow bool) (grid.Range, bool) { return s.area(sheet, r, grow) }
				job.panes = s.panesNow(sheet)
				job.filterGone = func(r grid.Range) bool { return s.filterGone(sheet, r) }
				if f := s.wb.Sheet(sheet).Filter; f.R1 > 0 && f.R2 > f.R1 && s.filterGone(sheet, f) {
					if body, ok := s.area(sheet, grid.Range{R1: f.R1 + 1, C1: f.C1, R2: f.R2, C2: f.C2}, false); ok {
						job.unhide = &body
					}
				}
			}
			if job.expand {
				job.formulaAt = func(r, c int) string { return s.cellNow(sheet, r, c).Formula }
			}
			if _, edited := s.extent[k]; edited || moved[k] {
				u := s.usedNow(sheet)
				job.dimension = &u
			}
			if !job.needed() {
				if isValue[lower] {
					sum, err := hashPart(zf)
					if err != nil {
						return nil, err
					}
					sums[lower] = sum
				}
				if err := copyRaw(zw, zf); err != nil {
					return nil, err
				}
				continue
			}
			r, err := zf.Open()
			if err != nil {
				return nil, err
			}
			w, err := zw.CreateHeader(&zip.FileHeader{Name: name, Method: zip.Deflate})
			if err != nil {
				r.Close()
				return nil, err
			}
			h := sha256.New()
			err = rewriteSheet(r, io.MultiWriter(w, h), job)
			r.Close()
			if err != nil {
				return nil, fmt.Errorf("writing sheet %q: %w", sheet, err)
			}
			var sum [32]byte
			copy(sum[:], h.Sum(nil))
			sums[lower] = sum
			continue
		}
		var b []byte
		var change func([]byte) []byte
		switch {
		case name == s.wbPath:
			change = func(b []byte) []byte { return fullCalc(s.rewriteNames(b)) }
		case lower == strings.ToLower(stylesPart) && stylesXML != nil:
			change = func([]byte) []byte { return stylesXML }
		case tables[lower] != nil:
			change = func([]byte) []byte { return tables[lower] }
		case strings.HasPrefix(name, "xl/charts/chart") && strings.HasSuffix(name, ".xml"):
			change = func(b []byte) []byte { return freshCaches(b, s.rangeRewrites(), res.touched) }
		case attached[name] != nil:
			change = attached[name]
		case !hasCustom && name == "[Content_Types].xml":
			change = addCustomType
		case !hasCustom && name == "_rels/.rels":
			change = addCustomRel
		}
		if change == nil && !isValue[lower] {
			if err := copyRaw(zw, zf); err != nil {
				return nil, err
			}
			continue
		}
		if b, err = s.p.read(name); err != nil {
			return nil, err
		}
		if change != nil {
			b = change(b)
		}
		if isValue[lower] {
			sums[lower] = sha256.Sum256(b)
		}
		if change == nil {
			if err := copyRaw(zw, zf); err != nil {
				return nil, err
			}
			continue
		}
		if err := write(name, b); err != nil {
			return nil, err
		}
	}
	// The stamp vouches for the results stored (see stamp.go).
	stamp := fingerprintSums(sums)
	if hasCustom {
		b, err := s.p.read(customPart)
		if err != nil {
			return nil, err
		}
		if err := write(customPart, setStamp(b, stamp)); err != nil {
			return nil, err
		}
	} else if err := write("docProps/custom.xml", setStamp([]byte(customXML), stamp)); err != nil {
		return nil, err
	}
	if err := zw.Close(); err != nil {
		return nil, err
	}
	return out.Bytes(), nil
}

// fingerprintSums is fingerprint, from the parts' hashes.
func fingerprintSums(sums map[string][32]byte) string {
	names := make([]string, 0, len(sums))
	for n := range sums {
		names = append(names, n)
	}
	sort.Strings(names)
	h := sha256.New()
	for _, n := range names {
		sum := sums[n]
		h.Write([]byte(n))
		h.Write([]byte{0})
		h.Write(sum[:])
	}
	return hex.EncodeToString(h.Sum(nil))
}

func hashPart(zf *zip.File) ([32]byte, error) {
	r, err := zf.Open()
	if err != nil {
		return [32]byte{}, err
	}
	defer r.Close()
	h := sha256.New()
	if _, err := io.Copy(h, r); err != nil {
		return [32]byte{}, err
	}
	var sum [32]byte
	copy(sum[:], h.Sum(nil))
	return sum, nil
}

// copyRaw copies a part as it is, compressed.
func copyRaw(zw *zip.Writer, zf *zip.File) error {
	fh := zf.FileHeader
	w, err := zw.CreateRaw(&fh)
	if err != nil {
		return err
	}
	r, err := zf.OpenRaw()
	if err != nil {
		return err
	}
	_, err = io.Copy(w, r)
	return err
}

// extends tells whether rows were appended to a table on a sheet.
func (s *StreamEditor) extends(sheet string) bool {
	for _, op := range s.ops {
		if op.kind == 'x' && strings.EqualFold(op.sheet, sheet) {
			return true
		}
	}
	return false
}

// noteBrokenRefs warns of formulas that gained a #REF! because rows they
// pointed into were deleted.
func (s *StreamEditor) noteBrokenRefs() {
	deleted := false
	for _, op := range s.ops {
		deleted = deleted || op.kind == 'd'
	}
	if !deleted {
		return
	}
	var refs []string
	for _, sh := range s.wb.Sheets {
		sh.Each(sh.Used(), func(r, c int, cell *grid.Cell) {
			if cell.Formula == "" {
				return
			}
			now, ok := s.rowNow(sh.Name, r)
			if !ok || s.edit(sh.Name, now, c) != nil {
				return
			}
			if strings.Count(s.rewrite(cell.Formula, sh.Name), "#REF!") > strings.Count(cell.Formula, "#REF!") {
				refs = append(refs, grid.FormatRef(sh.Name, grid.Range{R1: now, C1: c, R2: now, C2: c}))
			}
		})
	}
	if len(refs) > 0 {
		sort.Strings(refs)
		s.warnings = append(s.warnings, fmt.Sprintf("%s now %s #REF! (a reference pointed into what was removed): %s",
			countOf(len(refs), "formula"), map[bool]string{true: "contains", false: "contain"}[len(refs) == 1], strings.Join(refs[:min(len(refs), 10)], ", ")))
	}
}

var definedNameEl = regexp.MustCompile(`(<definedName\b[^>]*\bname="([^"]*)"[^>]*>)([^<]*)(</definedName>)`)

// rewriteNames brings the workbook's names up to the rows as they are
// now; the record's styling grows over rows appended to a table.
func (s *StreamEditor) rewriteNames(b []byte) []byte {
	if len(s.ops) == 0 {
		return b
	}
	return definedNameEl.ReplaceAllFunc(b, func(m []byte) []byte {
		g := definedNameEl.FindSubmatch(m)
		name := html.UnescapeString(string(g[2]))
		text := html.UnescapeString(string(g[3]))
		out := text
		for _, op := range s.ops {
			if op.kind == 'x' {
				if strings.HasPrefix(name, recordPrefix) {
					out = formula.ExtendRangesIn(out, "", op.sheet, op.from, op.to, op.c1, op.c2)
				}
				continue
			}
			out = op.formula(out, "")
		}
		open := g[1]
		if strings.HasPrefix(name, recordPrefix) {
			open = s.rewriteRecordRule(open, text)
		}
		if out == text && bytes.Equal(open, g[1]) {
			return m
		}
		return append(append(append([]byte{}, open...), xmlEscape(out)...), g[4]...)
	})
}

var commentAttr = regexp.MustCompile(`(\scomment=")([^"]*)(")`)

// rewriteRecordRule brings the formulas of a recorded conditional format
// (in its name's comment) up to the rows as they are now, as relines does
// for the full editor.
func (s *StreamEditor) rewriteRecordRule(open []byte, ref string) []byte {
	sheet, _, ok := parseAreas(ref)
	if !ok {
		return open
	}
	return commentAttr.ReplaceAllFunc(open, func(a []byte) []byte {
		g := commentAttr.FindSubmatch(a)
		ws, err := words.Texts(html.UnescapeString(string(g[2])))
		if err != nil || len(ws) == 0 || ws[0] != "cond" {
			return a
		}
		args := condFormulas(ws[1:], func(f string) string {
			for _, op := range s.ops {
				f = op.formula(f, sheet)
			}
			return f
		})
		return []byte(string(g[1]) + attrEscape(Record{Kind: "cond", Args: args}.Text()) + string(g[3]))
	})
}

var calcPrEl = regexp.MustCompile(`<calcPr\b[^>]*?(/>|>)`)

// fullCalc tells Excel to recalculate on open, as Editor.Save does.
func fullCalc(b []byte) []byte {
	if loc := calcPrEl.FindIndex(b); loc != nil {
		tag := string(b[loc[0]:loc[1]])
		if strings.Contains(tag, "fullCalcOnLoad=") {
			tag = regexp.MustCompile(`fullCalcOnLoad="[^"]*"`).ReplaceAllString(tag, `fullCalcOnLoad="1"`)
		} else {
			tag = strings.Replace(tag, "<calcPr", `<calcPr fullCalcOnLoad="1"`, 1)
		}
		return append(append(append([]byte{}, b[:loc[0]]...), tag...), b[loc[1]:]...)
	}
	for _, after := range []string{"</definedNames>", "</externalReferences>", "</sheets>"} {
		if i := bytes.Index(b, []byte(after)); i >= 0 {
			i += len(after)
			return append(append(append([]byte{}, b[:i]...), `<calcPr fullCalcOnLoad="1"/>`...), b[i:]...)
		}
	}
	return b
}

// legacyDrawing is the drawing holding a sheet's notes.
func (s *StreamEditor) legacyDrawing(part string) string {
	b, err := s.p.read(part)
	if err != nil {
		return ""
	}
	m := legacyDrawing.FindSubmatch(b)
	if m == nil {
		return ""
	}
	dir := path.Dir(part)
	rels, err := s.p.rels(path.Join(dir, "_rels", path.Base(part)+".rels"), dir)
	if err != nil {
		return ""
	}
	for _, r := range rels {
		if r.ID == string(m[1]) {
			return r.Target
		}
	}
	return ""
}

var commentEl = regexp.MustCompile(`(?s)<comment\b[^>]*\bref="([A-Z]+)([0-9]+)"[^>]*>.*?</comment>`)

// moveComments moves the notes of a sheet with its rows; notes on deleted
// rows go.
func moveComments(b []byte, rowNow func(int) (int, bool)) []byte {
	return commentEl.ReplaceAllFunc(b, func(m []byte) []byte {
		g := commentEl.FindSubmatch(m)
		r, _ := strconv.Atoi(string(g[2]))
		now, ok := rowNow(r)
		if !ok {
			return nil
		}
		if now == r {
			return m
		}
		old := []byte(`ref="` + string(g[1]) + string(g[2]) + `"`)
		return bytes.Replace(m, old, []byte(`ref="`+string(g[1])+strconv.Itoa(now)+`"`), 1)
	})
}

var anchorEl = regexp.MustCompile(`<x:Anchor>([^<]*)</x:Anchor>`)

// shiftNoteShapes moves the note shapes of a drawing with their rows,
// keeping their size; shapes of notes on deleted rows go.
func shiftNoteShapes(b []byte, rowNow func(int) (int, bool)) []byte {
	return noteShape.ReplaceAllFunc(b, func(sh []byte) []byte {
		if !noteType.Match(sh) {
			return sh
		}
		m := shapeRow.FindSubmatch(sh)
		if m == nil {
			return sh
		}
		r, _ := strconv.Atoi(string(m[1]))
		now, ok := rowNow(r + 1)
		if !ok {
			return nil
		}
		d := now - 1 - r
		if d == 0 {
			return sh
		}
		sh = shapeRow.ReplaceAll(sh, []byte("<x:Row>"+strconv.Itoa(now-1)+"</x:Row>"))
		return anchorEl.ReplaceAllFunc(sh, func(a []byte) []byte {
			parts := strings.Split(string(anchorEl.FindSubmatch(a)[1]), ",")
			if len(parts) != 8 {
				return a
			}
			for _, i := range []int{2, 6} {
				if n, err := strconv.Atoi(strings.TrimSpace(parts[i])); err == nil {
					parts[i] = " " + strconv.Itoa(max(0, n+d))
				}
			}
			return []byte("<x:Anchor>" + strings.TrimSpace(strings.Join(parts, ",")) + "</x:Anchor>")
		})
	})
}

// tablePartNow is an Excel table's part with its ranges where they are now
// (see fixTables); false when it is unchanged or gone.
func tablePartNow(t tableSource, rewrites []func(string) string) ([]byte, bool) {
	tn, ok := t.now(rewrites)
	if !ok {
		return nil, false
	}
	failed := false
	data := refAttr.ReplaceAllFunc(t.xml, func(m []byte) []byte {
		sub := refAttr.FindSubmatch(m)
		rg, err := grid.ParseRange(string(sub[2]))
		if err != nil {
			return m
		}
		_, moved, ok := locate(t.sheet, rg, rewrites)
		if !ok {
			failed = true
			return m
		}
		return []byte(string(sub[1]) + moved.String() + string(sub[3]))
	})
	if failed {
		return nil, false
	}
	if tn.totals == 0 && t.totals > 0 {
		data = totalsAttr.ReplaceAll(data, nil)
	}
	if bytes.Equal(data, t.xml) {
		return nil, false
	}
	return data, true
}

// panesNow moves a sheet's frozen rows with the rows inserted or deleted
// among them, as Editor.movePanes does.
func (s *StreamEditor) panesNow(sheet string) func([]byte) []byte {
	return func(head []byte) []byte {
		return paneEl.ReplaceAllFunc(head, func(tag []byte) []byte {
			t := string(tag)
			if !strings.Contains(t, `state="frozen"`) {
				return tag
			}
			y, x := 0, 0
			if m := ySplitAttr.FindStringSubmatch(t); m != nil {
				f, _ := strconv.ParseFloat(m[1], 64)
				y = int(f)
			}
			if m := xSplitAttr.FindStringSubmatch(t); m != nil {
				f, _ := strconv.ParseFloat(m[1], 64)
				x = int(f)
			}
			if y == 0 {
				return tag
			}
			for _, op := range s.ops {
				if !strings.EqualFold(op.sheet, sheet) {
					continue
				}
				switch {
				case op.kind == 'i' && op.at <= y:
					y += op.n
				case op.kind == 'd' && op.from <= y:
					y -= min(op.to, y) - op.from + 1
				}
			}
			if y <= 0 && x == 0 {
				return nil
			}
			t = ySplitAttr.ReplaceAllString(t, `ySplit="`+strconv.Itoa(max(y, 0))+`"`)
			t = topLeftAttr.ReplaceAllString(t, `topLeftCell="`+grid.CellName(max(y, 0)+1, x+1)+`"`)
			return []byte(t)
		})
	}
}

var (
	paneEl      = regexp.MustCompile(`<pane\b[^>]*/>|<pane\b[^>]*>\s*</pane>`)
	ySplitAttr  = regexp.MustCompile(`ySplit="([^"]*)"`)
	xSplitAttr  = regexp.MustCompile(`xSplit="([^"]*)"`)
	topLeftAttr = regexp.MustCompile(`topLeftCell="[^"]*"`)
)

// The document properties holding the results stamp.
const (
	customType = "application/vnd.openxmlformats-officedocument.custom-properties+xml"
	customRel  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
	customXML  = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n" +
		`<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"></Properties>`
)

// customProps is the package's custom properties part, if it has one.
func (s *StreamEditor) customProps() (string, bool) {
	rels, err := s.p.rels("_rels/.rels", "")
	if err != nil {
		return "", false
	}
	for _, r := range rels {
		if strings.HasSuffix(r.Type, "/custom-properties") {
			if _, err := s.p.read(r.Target); err == nil {
				return r.Target, true
			}
		}
	}
	return "", false
}

func addCustomType(b []byte) []byte {
	return bytes.Replace(b, []byte("</Types>"), []byte(`<Override PartName="/docProps/custom.xml" ContentType="`+customType+`"/></Types>`), 1)
}

func addCustomRel(b []byte) []byte {
	return bytes.Replace(b, []byte("</Relationships>"), []byte(`<Relationship Id="rIdSheetsCustom" Type="`+customRel+`" Target="docProps/custom.xml"/></Relationships>`), 1)
}

var (
	stampProp = regexp.MustCompile(`(?s)(<property\b[^>]*\bname="` + regexp.QuoteMeta(stampName) + `"[^>]*>\s*<vt:lpwstr>)[^<]*(</vt:lpwstr>)`)
	pidAttr   = regexp.MustCompile(`\bpid="(\d+)"`)
)

// setStamp writes the results stamp into the custom properties.
func setStamp(b []byte, stamp string) []byte {
	if stampProp.Match(b) {
		return stampProp.ReplaceAll(b, []byte("${1}v1 "+stamp+"${2}"))
	}
	pid := 1
	for _, m := range pidAttr.FindAllSubmatch(b, -1) {
		n, _ := strconv.Atoi(string(m[1]))
		pid = max(pid, n)
	}
	prop := fmt.Sprintf(`<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="%d" name="%s"><vt:lpwstr>v1 %s</vt:lpwstr></property></Properties>`, pid+1, stampName, stamp)
	if bytes.Contains(b, []byte("</Properties>")) {
		return bytes.Replace(b, []byte("</Properties>"), []byte(prop), 1)
	}
	return bytes.Replace(b, []byte("/>"), []byte(">"+prop), 1) // <Properties .../>
}

// stylesOf reads the style of each of the given cells of a sheet (where
// they are now) from the sheet as opened: the cell's own, else its row's,
// else its column's.
func (s *StreamEditor) stylesOf(sheet string, cells map[[2]int]bool) (map[[2]int]int, error) {
	want := map[[2]int][2]int{} // as opened → now
	rows := map[int]bool{}
	for pos := range cells {
		if r, ok := s.rowThen(sheet, pos[0]); ok {
			want[[2]int{r, pos[1]}] = pos
			rows[r] = true
		}
	}
	out := map[[2]int]int{}
	zf := s.p.files[strings.ToLower(s.parts[sheet])]
	if zf == nil {
		return out, nil
	}
	r, err := zf.Open()
	if err != nil {
		return nil, err
	}
	defer r.Close()
	dec := xml.NewDecoder(bufio.NewReaderSize(r, 1<<20))
	rowStyle := map[int]int{}
	var cols [][3]int // min, max, style
	row := 0
	for {
		tok, err := dec.RawToken()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, err
		}
		se, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		switch se.Name.Local {
		case "col":
			lo, _ := strconv.Atoi(attr(se, "min"))
			hi, _ := strconv.Atoi(attr(se, "max"))
			st, _ := strconv.Atoi(attr(se, "style"))
			cols = append(cols, [3]int{lo, hi, st})
		case "row":
			if v := attr(se, "r"); v != "" {
				row, _ = strconv.Atoi(v)
			} else {
				row++
			}
			if rows[row] && truthy(attr(se, "customFormat")) {
				st, _ := strconv.Atoi(attr(se, "s"))
				rowStyle[row] = st
			}
		case "c":
			if !rows[row] {
				continue
			}
			rr, cc, ok := grid.ParseCell(attr(se, "r"))
			if !ok {
				continue
			}
			if now, ok := want[[2]int{rr, cc}]; ok {
				st, _ := strconv.Atoi(attr(se, "s"))
				out[now] = st
			}
		}
	}
	for then, now := range want {
		if _, ok := out[now]; ok {
			continue
		}
		if st, ok := rowStyle[then[0]]; ok {
			out[now] = st
			continue
		}
		for _, c := range cols {
			if then[1] >= c[0] && then[1] <= c[1] {
				out[now] = c[2]
			}
		}
	}
	return out, nil
}

// cellFormats works out the format of each written cell that changes: the
// number formats cells are given, then the recorded styling over written
// cells that were empty, in the record's order. The style book is nil when
// no format changes.
func (s *StreamEditor) cellFormats(stylesPart, themePart string) (map[string]map[[2]int]string, *styleBook, error) {
	recs, _, err := rawRecords(s.data)
	if err != nil {
		recs = nil
	}
	type styled struct {
		look  Look
		areas []grid.Range
	}
	looks := map[string][]styled{} // lower-cased sheet → recorded looks, where they are now
	for _, rec := range recs {
		if rec.Kind != "style" {
			continue
		}
		l, err := ParseLook(rec.Args)
		if err != nil || !l.CellStyle() {
			continue
		}
		var areas []grid.Range
		for _, a := range rec.Areas {
			if x, ok := s.area(rec.Sheet, a, true); ok {
				areas = append(areas, x)
			}
		}
		k := strings.ToLower(rec.Sheet)
		looks[k] = append(looks[k], styled{l, areas})
	}
	out := map[string]map[[2]int]string{}
	var sb *styleBook
	for _, sh := range s.wb.Sheets {
		k := strings.ToLower(sh.Name)
		edits := s.edits[k]
		if len(edits) == 0 {
			continue
		}
		// Cells written that were empty, and cells given a number format.
		fresh := map[[2]int]bool{}
		need := map[[2]int]bool{}
		for pos, e := range edits {
			if e.numFmt != "" {
				need[pos] = true
			}
			if e.v.IsEmpty() {
				continue
			}
			was := grid.Cell{}
			if r, ok := s.rowThen(sh.Name, pos[0]); ok {
				was = sh.Get(r, pos[1])
			}
			if was.IsEmpty() {
				fresh[pos] = true
			}
		}
		used := s.usedNow(sh.Name)
		var under []styled // looks reaching written cells
		for _, st := range looks[k] {
			reaches := false
			for i, a := range st.areas {
				a = clip(a, used)
				st.areas[i] = a
				for pos := range fresh {
					if a.Contains(pos[0], pos[1]) {
						need[pos], reaches = true, true
					}
				}
			}
			if reaches {
				under = append(under, st)
			}
		}
		if len(need) == 0 {
			continue
		}
		if sb == nil {
			styles, err := s.p.read(stylesPart)
			if err != nil {
				return nil, nil, err
			}
			var theme []byte
			if themePart != "" {
				theme, _ = s.p.read(themePart)
			}
			if sb, err = newStyleBook(styles, theme); err != nil {
				return nil, nil, err
			}
		}
		cur, err := s.stylesOf(sh.Name, need)
		if err != nil {
			sb.close()
			return nil, nil, err
		}
		start := map[[2]int]int{}
		for pos := range need {
			start[pos] = cur[pos]
		}
		for pos, e := range edits {
			if e.numFmt != "" {
				if cur[pos], err = sb.numFmt(cur[pos], e.numFmt); err != nil {
					sb.close()
					return nil, nil, err
				}
			}
		}
		for _, st := range under {
			for _, a := range st.areas {
				for pos := range fresh {
					if !a.Contains(pos[0], pos[1]) {
						continue
					}
					var edges uint8
					if pos[0] == a.R1 {
						edges |= edgeTop
					}
					if pos[0] == a.R2 {
						edges |= edgeBottom
					}
					if pos[1] == a.C1 {
						edges |= edgeLeft
					}
					if pos[1] == a.C2 {
						edges |= edgeRight
					}
					l := st.look
					if cur[pos], err = sb.look(cur[pos], &l, edges); err != nil {
						sb.close()
						return nil, nil, err
					}
				}
			}
		}
		m := map[[2]int]string{}
		for pos := range need {
			if cur[pos] != start[pos] {
				m[pos] = strconv.Itoa(cur[pos])
			}
		}
		out[k] = m
	}
	return out, sb, nil
}

// filterGone tells whether a sheet's filter over r goes with rows deleted:
// excelize removes a filter whose header row is deleted, and shows the
// rows it hid.
func (s *StreamEditor) filterGone(sheet string, r grid.Range) bool {
	for _, op := range s.ops {
		if !strings.EqualFold(op.sheet, sheet) {
			continue
		}
		switch op.kind {
		case 'd':
			if op.from <= r.R1 && r.R1 <= op.to {
				return true
			}
		}
		x, ok := s.areaOp(r, op)
		if op.kind == 'x' || !ok {
			if !ok {
				return true
			}
			continue
		}
		r = x
	}
	return false
}

package xlsx

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"html"
	"io"
	"path"
	"regexp"
	"sort"
	"strings"

	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
	"github.com/xuri/excelize/v2"
)

// patches are fixes applied to the package XML when saving, where excelize
// has no API or gets things wrong. They are keyed by sheet name.
type patches struct {
	plainCols  map[string]bool            // drop column styles
	noFilter   map[string]bool            // drop the sheet filter
	dynamic    map[string]map[string]bool // cells to mark as dynamic arrays
	plain      map[string]map[string]bool // cells to unmark
	sheetEdits []sheetEdit                // renamed and deleted sheets
}

// sheetEdit is a renamed sheet, or a deleted one when to is empty.
type sheetEdit struct{ from, to string }

func newPatches() *patches {
	return &patches{plainCols: map[string]bool{}, noFilter: map[string]bool{}, dynamic: map[string]map[string]bool{}, plain: map[string]map[string]bool{}}
}

// markCell records whether a cell holds a dynamic array formula.
func (p *patches) markCell(sheet, cell string, dynamic bool) {
	add, del := p.plain, p.dynamic
	if dynamic {
		add, del = p.dynamic, p.plain
	}
	if add[sheet] == nil {
		add[sheet] = map[string]bool{}
	}
	add[sheet][cell] = true
	delete(del[sheet], cell)
}

func (p *patches) rename(from, to string) {
	for _, m := range []map[string]bool{p.plainCols, p.noFilter} {
		if v, ok := m[from]; ok {
			m[to] = v
			delete(m, from)
		}
	}
	for _, m := range []map[string]map[string]bool{p.dynamic, p.plain} {
		if v, ok := m[from]; ok {
			m[to] = v
			delete(m, from)
		}
	}
}

var (
	colStyle   = regexp.MustCompile(`(<col\b[^>]*?) style="\d+"`)
	autoFilter = regexp.MustCompile(`(?s)<autoFilter\b[^>]*?(/>|>.*?</autoFilter>)`)
	cellTag    = regexp.MustCompile(`<c r="([A-Z]+[0-9]+)"[^>]*>`)
	cmAttr     = regexp.MustCompile(` cm="\d+"`)
)

// Dynamic array formulas reference this cell metadata (as cm="1"), the
// same part Excel 365 and XlsxWriter write.
const (
	metadataPart = "xl/metadata.xml"
	metadataType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"
	metadataCT   = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"
	metadataXML  = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n" +
		`<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">` +
		`<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>` +
		`<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>` +
		`<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>`
)

type zipPart struct {
	name string
	data []byte
}

// serialize writes the workbook and applies the patches:
//
//   - AutoFilter marks a sheet as having a filter applied (filterMode)
//     even with no criteria; a filter without criteria hides nothing, so
//     the flag is dropped.
//   - Column styles cannot be removed through excelize without rewriting
//     every cell in those columns (Excel stores "format the whole sheet" as
//     a style on columns A:XFD); they are dropped from <col> directly.
//   - excelize cannot remove a sheet filter; the <autoFilter> element is
//     dropped (the caller removes its defined name through excelize).
//   - Dynamic array formulas need a cm attribute pointing at cell metadata,
//     which excelize cannot set.
func serialize(f *excelize.File, pt *patches) ([]byte, error) {
	buf, err := f.WriteToBuffer()
	if err != nil {
		return nil, err
	}
	data := buf.Bytes()
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, err
	}
	var list []zipPart
	byName := map[string]int{}
	for _, zf := range zr.File {
		r, err := zf.Open()
		if err != nil {
			return nil, err
		}
		b, err := io.ReadAll(r)
		r.Close()
		if err != nil {
			return nil, err
		}
		byName[zf.Name] = len(list)
		list = append(list, zipPart{zf.Name, b})
	}
	if pt == nil {
		pt = newPatches()
	}

	// Resolve sheet names to their parts.
	partOf := map[string]string{}
	p := &pkg{files: map[string]*zip.File{}}
	for _, zf := range zr.File {
		p.files[strings.ToLower(zf.Name)] = zf
	}
	wbPath, wbRels, sheets, err := p.workbookParts()
	if err != nil {
		return nil, err
	}
	for name, part := range sheets {
		partOf[part] = name
	}

	cm := ""
	if hasCells(pt.dynamic) {
		if cm, err = ensureMetadata(p, &list, byName, wbPath, wbRels); err != nil {
			return nil, err
		}
	}

	if len(pt.sheetEdits) > 0 {
		for i, zp := range list {
			list[i].data = applySheetEdits(zp.name, zp.data, pt.sheetEdits)
		}
	}

	for i, zp := range list {
		if strings.HasPrefix(zp.name, "xl/drawings/") && strings.HasSuffix(zp.name, ".xml") {
			list[i].data = declareNamespaces(zp.data)
		}
		if zp.name == "xl/workbook.xml" || zp.name == wbPath {
			// Left behind when the last defined name is deleted.
			list[i].data = bytes.Replace(zp.data, []byte("<definedNames></definedNames>"), nil, 1)
		}
		sheet, ok := partOf[zp.name]
		if !ok {
			continue
		}
		b := zp.data
		if !bytes.Contains(b, []byte("<filterColumn")) {
			b = bytes.Replace(b, []byte(` filterMode="true"`), nil, 1)
			b = bytes.Replace(b, []byte(`<sheetPr></sheetPr>`), nil, 1)
		}
		if pt.plainCols[sheet] {
			b = colStyle.ReplaceAll(b, []byte("$1"))
		}
		if pt.noFilter[sheet] {
			b = autoFilter.ReplaceAll(b, nil)
		}
		if dyn, plain := pt.dynamic[sheet], pt.plain[sheet]; len(dyn)+len(plain) > 0 {
			b = patchCells(b, dyn, plain, cm)
		}
		list[i].data = b
	}

	var out bytes.Buffer
	zw := zip.NewWriter(&out)
	for _, zp := range list {
		w, err := zw.CreateHeader(&zip.FileHeader{Name: zp.name, Method: zip.Deflate})
		if err != nil {
			return nil, err
		}
		if _, err := w.Write(zp.data); err != nil {
			return nil, err
		}
	}
	if err := zw.Close(); err != nil {
		return nil, err
	}
	return out.Bytes(), nil
}

func hasCells(m map[string]map[string]bool) bool {
	for _, cells := range m {
		if len(cells) > 0 {
			return true
		}
	}
	return false
}

// workbookParts returns the workbook part, its relationships part, and the
// part of each worksheet by sheet name.
func (p *pkg) workbookParts() (wbPath, relsPath string, sheets map[string]string, err error) {
	wbPath = "xl/workbook.xml"
	if rels, err := p.rels("_rels/.rels", ""); err == nil {
		for _, r := range rels {
			if strings.HasSuffix(r.Type, "/officeDocument") {
				wbPath = r.Target
			}
		}
	}
	var wbx xWorkbook
	if err := p.decode(wbPath, &wbx); err != nil {
		return "", "", nil, fmt.Errorf("reading workbook: %w", err)
	}
	dir := path.Dir(wbPath)
	relsPath = path.Join(dir, "_rels", path.Base(wbPath)+".rels")
	rels, err := p.rels(relsPath, dir)
	if err != nil {
		return "", "", nil, err
	}
	byID := map[string]string{}
	for _, r := range rels {
		byID[r.ID] = r.Target
	}
	sheets = map[string]string{}
	for _, s := range wbx.Sheets {
		if t, ok := byID[s.relID()]; ok {
			sheets[s.Name] = t
		}
	}
	return wbPath, relsPath, sheets, nil
}

// ensureMetadata returns the cm index of dynamic array cell metadata,
// adding the metadata part when the package has none. It returns "" when
// existing metadata has no dynamic array entry; formulas then stay plain
// array formulas.
func ensureMetadata(p *pkg, list *[]zipPart, byName map[string]int, wbPath, relsPath string) (string, error) {
	rels, err := p.rels(relsPath, path.Dir(wbPath))
	if err != nil {
		return "", err
	}
	for _, r := range rels {
		if r.Type == metadataType {
			i, ok := byName[r.Target]
			if !ok {
				return "", nil
			}
			return dynamicArrayIndex((*list)[i].data), nil
		}
	}
	if _, taken := byName[metadataPart]; taken {
		return "", nil
	}
	ct, ok := byName["[Content_Types].xml"]
	ri, ok2 := byName[relsPath]
	if !ok || !ok2 {
		return "", nil
	}
	override := `<Override PartName="/` + metadataPart + `" ContentType="` + metadataCT + `"/></Types>`
	(*list)[ct].data = bytes.Replace((*list)[ct].data, []byte("</Types>"), []byte(override), 1)
	id := "rIdSheetsMetadata"
	rel := `<Relationship Id="` + id + `" Type="` + metadataType + `" Target="metadata.xml"/></Relationships>`
	(*list)[ri].data = bytes.Replace((*list)[ri].data, []byte("</Relationships>"), []byte(rel), 1)
	byName[metadataPart] = len(*list)
	*list = append(*list, zipPart{metadataPart, []byte(metadataXML)})
	return "1", nil
}

// dynamicArrayIndex finds the cellMetadata block (1-based, as cm uses)
// that points at the XLDAPR dynamic array properties.
func dynamicArrayIndex(data []byte) string {
	var md struct {
		Types []struct {
			Name string `xml:"name,attr"`
		} `xml:"metadataTypes>metadataType"`
		Future []struct {
			Name string `xml:"name,attr"`
			Bk   []struct {
				Props []struct {
					Dynamic string `xml:"fDynamic,attr"`
				} `xml:"extLst>ext>dynamicArrayProperties"`
			} `xml:"bk"`
		} `xml:"futureMetadata"`
		Cells []struct {
			Rc []struct {
				T int `xml:"t,attr"`
				V int `xml:"v,attr"`
			} `xml:"rc"`
		} `xml:"cellMetadata>bk"`
	}
	if xml.Unmarshal(data, &md) != nil {
		return ""
	}
	typ := 0
	for i, t := range md.Types {
		if t.Name == "XLDAPR" {
			typ = i + 1
		}
	}
	fut := -1
	for _, fm := range md.Future {
		if fm.Name != "XLDAPR" {
			continue
		}
		for i, bk := range fm.Bk {
			for _, p := range bk.Props {
				if p.Dynamic == "1" || p.Dynamic == "true" {
					fut = i
				}
			}
			if fut >= 0 {
				break
			}
		}
	}
	if typ == 0 || fut < 0 {
		return ""
	}
	for i, bk := range md.Cells {
		for _, rc := range bk.Rc {
			if rc.T == typ && rc.V == fut {
				return fmt.Sprint(i + 1)
			}
		}
	}
	return ""
}

// patchCells sets or removes the cm attribute on the given cells.
func patchCells(b []byte, dynamic, plain map[string]bool, cm string) []byte {
	return cellTag.ReplaceAllFunc(b, func(tag []byte) []byte {
		ref := cellTag.FindSubmatch(tag)[1]
		dyn := dynamic[string(ref)]
		if !dyn && !plain[string(ref)] {
			return tag
		}
		tag = cmAttr.ReplaceAll(tag, nil)
		if dyn && cm != "" {
			end := len(tag) - 1
			if bytes.HasSuffix(tag, []byte("/>")) {
				end--
			}
			tag = append(tag[:end:end], append([]byte(` cm="`+cm+`"`), tag[end:]...)...)
		}
		return tag
	})
}

var (
	formulaText = regexp.MustCompile(`(<(?:f|formula|formula1|formula2|c:f|definedName)\b[^>]*>)([^<]*)(</[a-zA-Z:0-9]+>)`)
	sourceSheet = regexp.MustCompile(`(<worksheetSource\b[^>]*\bsheet=")([^"]*)(")`)
)

// applySheetEdits rewrites sheet references inside formula-bearing parts:
// cell formulas, conditional formats and validations in worksheets, chart
// series, pivot sources and (for deleted sheets) names.
func applySheetEdits(name string, b []byte, edits []sheetEdit) []byte {
	isSheet := strings.HasPrefix(name, "xl/worksheets/") && strings.HasSuffix(name, ".xml")
	isChart := strings.HasPrefix(name, "xl/charts/chart") && strings.HasSuffix(name, ".xml")
	isPivot := strings.HasPrefix(name, "xl/pivotCache/") && strings.HasSuffix(name, ".xml")
	isBook := name == "xl/workbook.xml"
	if !isSheet && !isChart && !isPivot && !isBook {
		return b
	}
	transform := func(f string) string {
		for _, e := range edits {
			if e.to != "" {
				f = formula.RenameSheet(f, e.from, e.to)
			} else {
				f = formula.DropSheet(f, e.from)
			}
		}
		return f
	}
	b = formulaText.ReplaceAllFunc(b, func(m []byte) []byte {
		g := formulaText.FindSubmatch(m)
		text := html.UnescapeString(string(g[2]))
		out := transform(text)
		if out == text {
			return m
		}
		var esc bytes.Buffer
		xml.EscapeText(&esc, []byte(out))
		return append(append(append([]byte{}, g[1]...), esc.Bytes()...), g[3]...)
	})
	if isPivot {
		b = sourceSheet.ReplaceAllFunc(b, func(m []byte) []byte {
			g := sourceSheet.FindSubmatch(m)
			for _, e := range edits {
				if e.to != "" && strings.EqualFold(html.UnescapeString(string(g[2])), e.from) {
					var esc bytes.Buffer
					xml.EscapeText(&esc, []byte(e.to))
					return append(append(append([]byte{}, g[1]...), esc.Bytes()...), g[3]...)
				}
			}
			return m
		})
	}
	return b
}

// knownNamespaces maps the prefixes drawing parts use to their namespaces.
var knownNamespaces = map[string]string{
	"a":     "http://schemas.openxmlformats.org/drawingml/2006/main",
	"a14":   "http://schemas.microsoft.com/office/drawing/2010/main",
	"a16":   "http://schemas.microsoft.com/office/drawing/2014/main",
	"c":     "http://schemas.openxmlformats.org/drawingml/2006/chart",
	"cx":    "http://schemas.microsoft.com/office/drawing/2014/chartex",
	"dgm":   "http://schemas.openxmlformats.org/drawingml/2006/diagram",
	"mc":    "http://schemas.openxmlformats.org/markup-compatibility/2006",
	"pic":   "http://schemas.openxmlformats.org/drawingml/2006/picture",
	"r":     "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
	"sle15": "http://schemas.microsoft.com/office/drawing/2012/slicer",
	"xdr":   "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
}

var (
	usedPrefix     = regexp.MustCompile(`</?([A-Za-z][A-Za-z0-9]*):[A-Za-z]|\s([A-Za-z][A-Za-z0-9]*):[A-Za-z]+=`)
	declaredPrefix = regexp.MustCompile(`xmlns:([A-Za-z][A-Za-z0-9]*)=`)
	rootTag        = regexp.MustCompile(`<[A-Za-z][A-Za-z0-9]*:?[A-Za-z]*\b`)
)

// declareNamespaces adds declarations for prefixes a part uses but no
// longer declares. excelize rewrites drawings when rows or columns move and
// drops declarations Excel puts on inner elements (xmlns:c on c:chart),
// which leaves a file Excel has to repair.
func declareNamespaces(b []byte) []byte {
	declared := map[string]bool{"xml": true, "xmlns": true}
	for _, m := range declaredPrefix.FindAllSubmatch(b, -1) {
		declared[string(m[1])] = true
	}
	var add []string
	seen := map[string]bool{}
	for _, m := range usedPrefix.FindAllSubmatch(b, -1) {
		p := string(m[1])
		if p == "" {
			p = string(m[2])
		}
		if declared[p] || seen[p] {
			continue
		}
		seen[p] = true
		if ns, ok := knownNamespaces[p]; ok {
			add = append(add, ` xmlns:`+p+`="`+ns+`"`)
		}
	}
	if len(add) == 0 {
		return b
	}
	sort.Strings(add)
	// Insert after the root element's name, past any XML declaration.
	start := 0
	if bytes.HasPrefix(b, []byte("<?xml")) {
		start = bytes.Index(b, []byte("?>")) + 2
	}
	loc := rootTag.FindIndex(b[start:])
	if loc == nil {
		return b
	}
	at := start + loc[1]
	return append(b[:at:at], append([]byte(strings.Join(add, "")), b[at:]...)...)
}

func hasCharts(pkgData []byte) bool {
	zr, err := zip.NewReader(bytes.NewReader(pkgData), int64(len(pkgData)))
	if err != nil {
		return false
	}
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "xl/charts/chart") {
			return true
		}
	}
	return false
}

// stripStyles resets the style of every cell and row the reader saw
// styled. Column styles are dropped at save time (see serialize).
func stripStyles(f *excelize.File, sheet string, lay *layout) error {
	if lay == nil {
		return nil
	}
	for _, r := range lay.rows {
		if err := f.SetRowStyle(sheet, r, r, 0); err != nil {
			return err
		}
	}
	for _, c := range lay.cells {
		if err := f.SetCellStyle(sheet, grid.CellName(c.R1, c.C1), grid.CellName(c.R1, c.C2), 0); err != nil {
			return err
		}
	}
	return nil
}

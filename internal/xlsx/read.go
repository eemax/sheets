// Package xlsx reads, edits and formats .xlsx workbooks. Reading uses a
// streaming parser of the package XML (fast, and it sees shared formulas and
// cached values as stored); editing and formatting go through excelize.
package xlsx

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"path"
	"strconv"
	"strings"
	"time"

	"github.com/eemax/sheets/internal/cache"
	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
)

// Read loads a workbook from disk.
func Read(filename string) (*grid.Workbook, error) {
	data, err := os.ReadFile(filename)
	if err != nil {
		return nil, err
	}
	wb, err := ReadBytes(data)
	if err != nil {
		return nil, fmt.Errorf("%s: %w", filename, err)
	}
	wb.Path = filename
	wb.Size = int64(len(data))
	return wb, nil
}

// ReadBytes parses an xlsx package held in memory. A large file is kept
// parsed in the cache, so reading it again skips the parse.
func ReadBytes(data []byte) (*grid.Workbook, error) {
	if len(data) < cache.MinSize || cache.Disabled() {
		wb, _, err := parse(data)
		return wb, err
	}
	key := cache.Key("xlsx", data)
	if b, ok := cache.Get(key); ok {
		wb := &grid.Workbook{}
		if len(b) > 0 && wb.UnmarshalBinary(b[1:]) == nil {
			if b[0] == 1 {
				wb.Resolver = &calc{data: data}
			}
			return wb, nil
		}
		cache.Remove(key)
	}
	wb, _, err := parse(data)
	if err != nil {
		return nil, err
	}
	if b, err := wb.MarshalBinary(); err == nil {
		pending := byte(0)
		if wb.Resolver != nil {
			pending = 1
		}
		cache.Put(key, append([]byte{pending}, b...))
	}
	return wb, nil
}

// layout is formatting the reader saw that the grid model does not carry:
// which cells, rows and columns have a style. Stripping resets exactly
// these instead of sweeping a rectangle, which on a sparse sheet would mean
// millions of cells.
type layout struct {
	drawing bool         // the sheet has charts or pictures
	cells   []grid.Range // runs of styled cells within one row
	rows    []int        // rows with a row-level style
	// spills maps the anchor cell of each dynamic array formula (Excel 365
	// spilling formulas) to the range its result currently fills.
	spills map[[2]int]grid.Range
	tables []tablePart
}

// tablePart is an Excel table (Insert › Table) as stored: its part, its
// range with the header and totals rows it counts, and the XML, which
// sheets rewrites itself when rows move (excelize misses tables stored
// under absolute paths and drops their totals rows).
type tablePart struct {
	part    string
	name    string
	ref     grid.Range
	headers int  // header rows at the top of ref (0 or 1)
	totals  int  // totals rows at the bottom
	filter  bool // the header row shows filter buttons
	xml     []byte
}

// oleMagic starts Compound File Binary files: legacy .xls, and .xlsx files
// encrypted with a password.
var oleMagic = []byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1}

func parse(data []byte) (*grid.Workbook, map[string]*layout, error) {
	if bytes.HasPrefix(data, oleMagic) {
		return nil, nil, fmt.Errorf("file is password-protected or in the legacy .xls format; save it as an unencrypted .xlsx")
	}
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, nil, fmt.Errorf("not an xlsx file: %w", err)
	}
	p := &pkg{files: map[string]*zip.File{}}
	for _, f := range zr.File {
		p.files[strings.ToLower(f.Name)] = f
	}

	wbPath := "xl/workbook.xml"
	if rels, err := p.rels("_rels/.rels", ""); err == nil {
		for _, r := range rels {
			if strings.HasSuffix(r.Type, "/officeDocument") {
				wbPath = r.Target
			}
		}
	}
	var wbx xWorkbook
	if err := p.decode(wbPath, &wbx); err != nil {
		return nil, nil, fmt.Errorf("reading workbook: %w", err)
	}
	wbDir := path.Dir(wbPath)
	rels, err := p.rels(path.Join(wbDir, "_rels", path.Base(wbPath)+".rels"), wbDir)
	if err != nil {
		return nil, nil, fmt.Errorf("reading workbook relationships: %w", err)
	}
	relByID := map[string]xRel{}
	var sstPath, stylesPath string
	for _, r := range rels {
		relByID[r.ID] = r
		switch {
		case strings.HasSuffix(r.Type, "/sharedStrings"):
			sstPath = r.Target
		case strings.HasSuffix(r.Type, "/styles"):
			stylesPath = r.Target
		}
	}

	wb := &grid.Workbook{Format: "xlsx", Formats: []grid.Format{{}}}
	if wbx.Pr != nil {
		wb.Date1904 = wbx.Pr.Date1904 == "1" || wbx.Pr.Date1904 == "true"
	}
	stale := wbx.Calc != nil && (wbx.Calc.FullCalcOnLoad == "1" || wbx.Calc.FullCalcOnLoad == "true")

	var sst []string
	if sstPath != "" {
		if sst, err = p.sharedStrings(sstPath); err != nil {
			return nil, nil, fmt.Errorf("reading shared strings: %w", err)
		}
	}
	var xfFmt []uint16 // cell style index → index into wb.Formats
	if stylesPath != "" {
		if xfFmt, err = p.styles(stylesPath, wb); err != nil {
			return nil, nil, fmt.Errorf("reading styles: %w", err)
		}
	}

	pending := false
	layouts := map[string]*layout{}
	for _, s := range wbx.Sheets {
		rel, ok := relByID[s.relID()]
		if !ok || !strings.HasSuffix(rel.Type, "/worksheet") {
			continue // chartsheets, dialog sheets, macro sheets
		}
		sh := grid.NewSheet(s.Name)
		sh.Hidden = s.State == "hidden" || s.State == "veryHidden"
		lay := &layout{spills: map[[2]int]grid.Range{}}
		layouts[s.Name] = lay
		hasPending, err := p.worksheet(rel.Target, sh, lay, sst, xfFmt, stale)
		if err != nil {
			return nil, nil, fmt.Errorf("reading sheet %q: %w", s.Name, err)
		}
		tables, drawing := p.sheetRels(rel.Target)
		for _, t := range tables {
			sh.Tables = append(sh.Tables, grid.NamedTable{Name: t.name, Range: t.ref, Totals: t.totals, Filter: t.filter && t.headers > 0})
		}
		lay.tables = tables
		if drawing != "" {
			lay.drawing = true
			sh.Charts = p.charts(drawing)
		}
		pending = pending || hasPending
		wb.Sheets = append(wb.Sheets, sh)
	}
	for _, n := range wbx.Names {
		if strings.HasPrefix(n.Name, "_xlnm.") {
			continue
		}
		dn := grid.DefinedName{Name: n.Name, Ref: strings.TrimSpace(n.Ref)}
		if n.LocalSheetID != "" {
			if i, err := strconv.Atoi(n.LocalSheetID); err == nil && i < len(wbx.Sheets) {
				dn.Scope = wbx.Sheets[i].Name
			}
		}
		wb.Names = append(wb.Names, dn)
	}
	if pending {
		wb.Resolver = &calc{data: data}
	}
	return wb, layouts, nil
}

type pkg struct {
	files map[string]*zip.File
}

func (p *pkg) open(name string) (io.ReadCloser, error) {
	f, ok := p.files[strings.ToLower(strings.TrimPrefix(name, "/"))]
	if !ok {
		return nil, fmt.Errorf("missing part %s", name)
	}
	return f.Open()
}

func (p *pkg) decode(name string, v any) error {
	r, err := p.open(name)
	if err != nil {
		return err
	}
	defer r.Close()
	return xml.NewDecoder(r).Decode(v)
}

type xRel struct {
	ID         string `xml:"Id,attr"`
	Type       string `xml:"Type,attr"`
	Target     string `xml:"Target,attr"`
	TargetMode string `xml:"TargetMode,attr"`
}

// rels reads a relationships part and resolves targets against dir.
func (p *pkg) rels(name, dir string) ([]xRel, error) {
	var x struct {
		Rels []xRel `xml:"Relationship"`
	}
	if err := p.decode(name, &x); err != nil {
		return nil, err
	}
	for i, r := range x.Rels {
		if r.TargetMode == "External" {
			continue
		}
		if strings.HasPrefix(r.Target, "/") {
			x.Rels[i].Target = strings.TrimPrefix(r.Target, "/")
		} else {
			x.Rels[i].Target = path.Clean(path.Join(dir, r.Target))
		}
	}
	return x.Rels, nil
}

type xWorkbook struct {
	Pr *struct {
		Date1904 string `xml:"date1904,attr"`
	} `xml:"workbookPr"`
	Calc *struct {
		FullCalcOnLoad string `xml:"fullCalcOnLoad,attr"`
	} `xml:"calcPr"`
	Sheets []xSheet `xml:"sheets>sheet"`
	Names  []struct {
		Name         string `xml:"name,attr"`
		LocalSheetID string `xml:"localSheetId,attr"`
		Ref          string `xml:",chardata"`
	} `xml:"definedNames>definedName"`
}

type xSheet struct {
	Name  string     `xml:"name,attr"`
	State string     `xml:"state,attr"`
	Attrs []xml.Attr `xml:",any,attr"`
}

// relID finds the r:id attribute whatever namespace (transitional or strict)
// it uses.
func (s xSheet) relID() string {
	for _, a := range s.Attrs {
		if a.Name.Local == "id" && a.Name.Space != "" {
			return a.Value
		}
	}
	return ""
}

func (p *pkg) sharedStrings(name string) ([]string, error) {
	r, err := p.open(name)
	if err != nil {
		return nil, err
	}
	defer r.Close()
	dec := xml.NewDecoder(r)
	var (
		out     []string
		cur     strings.Builder
		inT     bool
		skipRPh int
	)
	for {
		tok, err := dec.RawToken()
		if err == io.EOF {
			return out, nil
		}
		if err != nil {
			return nil, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "si":
				cur.Reset()
			case "t":
				inT = skipRPh == 0
			case "rPh":
				skipRPh++
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "si":
				out = append(out, unescapeOOXML(cur.String()))
			case "t":
				inT = false
			case "rPh":
				skipRPh--
			}
		case xml.CharData:
			if inT {
				cur.Write(t)
			}
		}
	}
}

// styles loads number formats and maps each cell style to a grid.Format.
func (p *pkg) styles(name string, wb *grid.Workbook) ([]uint16, error) {
	var x struct {
		NumFmts []struct {
			ID   int    `xml:"numFmtId,attr"`
			Code string `xml:"formatCode,attr"`
		} `xml:"numFmts>numFmt"`
		Xfs []struct {
			NumFmtID int `xml:"numFmtId,attr"`
		} `xml:"cellXfs>xf"`
	}
	if err := p.decode(name, &x); err != nil {
		return nil, err
	}
	custom := map[int]string{}
	for _, nf := range x.NumFmts {
		custom[nf.ID] = nf.Code
	}
	byID := map[int]uint16{0: 0}
	out := make([]uint16, len(x.Xfs))
	for i, xf := range x.Xfs {
		idx, ok := byID[xf.NumFmtID]
		if !ok {
			code, found := custom[xf.NumFmtID]
			if !found {
				code, _ = grid.BuiltinFormatCode(xf.NumFmtID)
			}
			f := grid.ClassifyFormat(code)
			if f.Kind == grid.General && xf.NumFmtID != 0 && code != "" && !strings.EqualFold(code, "General") {
				f.Code = code
			}
			idx = uint16(len(wb.Formats))
			wb.Formats = append(wb.Formats, f)
			byID[xf.NumFmtID] = idx
		}
		out[i] = idx
	}
	return out, nil
}

// sheetRels lists the table objects attached to a worksheet and its drawing
// part (charts, pictures), if any. Parts that cannot be read are skipped:
// this only refines formatting and warnings.
func (p *pkg) sheetRels(sheetPart string) (out []tablePart, drawing string) {
	dir := path.Dir(sheetPart)
	rels, err := p.rels(path.Join(dir, "_rels", path.Base(sheetPart)+".rels"), dir)
	if err != nil {
		return nil, ""
	}
	for _, r := range rels {
		if strings.HasSuffix(r.Type, "/drawing") {
			drawing = r.Target
		}
		if !strings.HasSuffix(r.Type, "/table") {
			continue
		}
		data, err := p.read(r.Target)
		if err != nil {
			continue
		}
		var t struct {
			Name        string `xml:"name,attr"`
			DisplayName string `xml:"displayName,attr"`
			Ref         string `xml:"ref,attr"`
			Headers     *int   `xml:"headerRowCount,attr"`
			Totals      int    `xml:"totalsRowCount,attr"`
		}
		if err := xml.Unmarshal(data, &t); err != nil {
			continue
		}
		rg, err := grid.ParseRange(t.Ref)
		if err != nil {
			continue
		}
		tp := tablePart{part: r.Target, name: t.DisplayName, ref: rg, headers: 1, totals: t.Totals, xml: data,
			filter: bytes.Contains(data, []byte("<autoFilter"))}
		if tp.name == "" {
			tp.name = t.Name
		}
		if t.Headers != nil {
			tp.headers = *t.Headers
		}
		out = append(out, tp)
	}
	return out, drawing
}

// read returns a part's bytes.
func (p *pkg) read(name string) ([]byte, error) {
	r, err := p.open(name)
	if err != nil {
		return nil, err
	}
	defer r.Close()
	return io.ReadAll(r)
}

type sharedFormula struct {
	row, col int
	text     string
}

// worksheet streams one sheet's XML into sh. It reports whether any formula
// value is pending.
func (p *pkg) worksheet(name string, sh *grid.Sheet, lay *layout, sst []string, xfFmt []uint16, stale bool) (bool, error) {
	r, err := p.open(name)
	if err != nil {
		return false, err
	}
	defer r.Close()
	dec := xml.NewDecoder(r)
	shared := map[string]sharedFormula{}
	row, col := 0, 0
	pending := false
	for {
		tok, err := dec.RawToken()
		if err == io.EOF {
			return pending, nil
		}
		if err != nil {
			return false, err
		}
		se, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		switch se.Name.Local {
		case "row":
			row++
			if v := attr(se, "r"); v != "" {
				if n, err := strconv.Atoi(v); err == nil {
					row = n
				}
			}
			col = 0
			if st := attr(se, "s"); st != "" && st != "0" {
				lay.rows = append(lay.rows, row)
			}
			if truthy(attr(se, "hidden")) {
				sh.HiddenRows = grid.AddRun(sh.HiddenRows, row, row)
			}

		case "autoFilter":
			// The sheet's filter comes before any in custom views.
			if rg, err := grid.ParseRange(attr(se, "ref")); err == nil && sh.Filter.R1 == 0 {
				sh.Filter = rg
			}

		case "col":
			if truthy(attr(se, "hidden")) {
				lo, err1 := strconv.Atoi(attr(se, "min"))
				hi, err2 := strconv.Atoi(attr(se, "max"))
				if err1 == nil && err2 == nil && 1 <= lo && lo <= hi {
					sh.HiddenCols = grid.AddRun(sh.HiddenCols, lo, min(hi, grid.MaxCols))
				}
			}

		case "c":
			col++
			var typ, style string
			dynamic := false
			for _, a := range se.Attr {
				switch a.Name.Local {
				case "r":
					if rr, cc, ok := grid.ParseCell(a.Value); ok {
						row, col = rr, cc
					}
				case "t":
					typ = a.Value
				case "s":
					style = a.Value
				case "cm":
					dynamic = a.Value != "" && a.Value != "0"
				}
			}
			c, arrayRef, err := readCell(dec, typ, sst, shared, row, col)
			if err != nil {
				return false, err
			}
			if dynamic && c.Formula != "" {
				spill := grid.Range{R1: row, C1: col, R2: row, C2: col}
				if rg, err := grid.ParseRange(arrayRef); err == nil {
					spill = rg
				}
				lay.spills[[2]int{row, col}] = spill
			}
			if style != "" && style != "0" {
				if i, err := strconv.Atoi(style); err == nil && i < len(xfFmt) {
					c.Fmt = xfFmt[i]
				}
				if n := len(lay.cells); n > 0 && lay.cells[n-1].R1 == row && lay.cells[n-1].C2 == col-1 {
					lay.cells[n-1].C2 = col
				} else {
					lay.cells = append(lay.cells, grid.Range{R1: row, C1: col, R2: row, C2: col})
				}
			}
			if c.Formula != "" && stale {
				c.Pending = true
			}
			if c.Pending {
				pending = true
			}
			if !c.IsEmpty() {
				sh.Set(row, col, c)
			}
		case "mergeCell":
			if rg, err := grid.ParseRange(attr(se, "ref")); err == nil {
				sh.Merges = append(sh.Merges, rg)
			}
		}
	}
}

// unescapeOOXML decodes the _xHHHH_ escapes Excel uses for characters XML
// cannot hold (such as carriage returns, _x000D_). _x005F_ escapes a
// literal underscore.
func unescapeOOXML(s string) string {
	if !strings.Contains(s, "_x") {
		return s
	}
	var b strings.Builder
	for i := 0; i < len(s); i++ {
		if s[i] == '_' && i+6 < len(s) && s[i+1] == 'x' && s[i+6] == '_' {
			if n, err := strconv.ParseUint(s[i+2:i+6], 16, 16); err == nil {
				b.WriteRune(rune(n))
				i += 6
				continue
			}
		}
		b.WriteByte(s[i])
	}
	return b.String()
}

func attr(se xml.StartElement, name string) string {
	for _, a := range se.Attr {
		if a.Name.Local == name {
			return a.Value
		}
	}
	return ""
}

// readCell consumes the children of a <c> element.
// It also returns the ref of an array formula (the range a dynamic array
// spills into).
func readCell(dec *xml.Decoder, typ string, sst []string, shared map[string]sharedFormula, row, col int) (grid.Cell, string, error) {
	var (
		c        grid.Cell
		v        string
		hasV     bool
		inline   strings.Builder
		where    string // element whose text we are collecting
		fType    string
		fSI      string
		fRef     string
		fText    strings.Builder
		hasF     bool
		inRPh    int
		inInline bool
	)
loop:
	for {
		tok, err := dec.RawToken()
		if err != nil {
			return c, "", err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "v":
				where, hasV = "v", true
			case "f":
				where, hasF = "f", true
				fType, fSI, fRef = attr(t, "t"), attr(t, "si"), attr(t, "ref")
			case "is":
				inInline = true
			case "t":
				if inInline && inRPh == 0 {
					where = "t"
				}
			case "rPh":
				inRPh++
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "c":
				break loop
			case "rPh":
				inRPh--
			case "is":
				inInline = false
			}
			where = ""
		case xml.CharData:
			switch where {
			case "v":
				v += string(t)
			case "f":
				fText.Write(t)
			case "t":
				inline.Write(t)
			}
		}
	}
	if hasF {
		text := fText.String()
		switch fType {
		case "shared":
			if text != "" {
				shared[fSI] = sharedFormula{row, col, text}
			} else if m, ok := shared[fSI]; ok {
				text = formula.Shift(m.text, row-m.row, col-m.col)
			}
		case "dataTable":
			text = ""
		}
		// Stored formulas carry _xlfn./_xlpm. prefixes; show them as typed.
		c.Formula = formula.FromFile(text)
		c.Shared = fType == "shared" && text != ""
	}
	switch typ {
	case "s":
		if i, err := strconv.Atoi(strings.TrimSpace(v)); err == nil && i >= 0 && i < len(sst) {
			c.Kind, c.Text = grid.Text, sst[i]
		}
	case "str":
		if hasV {
			c.Kind, c.Text = grid.Text, v
		}
	case "inlineStr":
		c.Kind, c.Text = grid.Text, unescapeOOXML(inline.String())
	case "b":
		if hasV {
			c.Kind, c.Bool = grid.Bool, strings.TrimSpace(v) == "1"
		}
	case "e":
		c.Kind, c.Text = grid.Error, v
	case "d":
		if t, err := time.Parse(time.RFC3339Nano, strings.TrimSpace(v)); err == nil {
			c.Kind, c.Num = grid.Number, grid.TimeToSerial(t)
		} else if t, err := time.Parse("2006-01-02", strings.TrimSpace(v)); err == nil {
			c.Kind, c.Num = grid.Number, grid.TimeToSerial(t)
		}
	default:
		if s := strings.TrimSpace(v); s != "" {
			if n, err := strconv.ParseFloat(s, 64); err == nil {
				c.Kind, c.Num = grid.Number, n
			}
		}
	}
	if c.Formula != "" && !hasV {
		c.Pending = true
	}
	if fType != "array" {
		fRef = ""
	}
	return c, fRef, nil
}

// truthy reads an xsd:boolean attribute.
func truthy(v string) bool { return v == "1" || v == "true" }

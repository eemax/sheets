package fuzz

import (
	"archive/zip"
	"bufio"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"math"
	"path"
	"regexp"
	"slices"
	"sort"
	"strconv"
	"strings"

	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/xlsxtest"
)

// The kinds a formula is stored as, and the markers that name them in its
// text, as the condition of an IF that is always true: IF(N("p")=0,...),
// N("l"), N("d"). N of a text is 0, so a marker changes no result, and keeps
// the type of the result (it is not added to it).
const (
	Plain   = "p" // <f>, or a shared formula
	Legacy  = "l" // <f t="array">, entered with Ctrl+Shift+Enter
	Dynamic = "d" // <f t="array"> in a cell with cm: Excel 365's dynamic arrays
)

var kindName = map[string]string{Plain: "plain formula", Legacy: "legacy array", Dynamic: "dynamic array"}

var markerRE = regexp.MustCompile(`N\("([pld])"\)`)

// marker returns the kind a formula's text names, "" when it names none.
func marker(formula string) string {
	m := markerRE.FindAllStringSubmatch(formula, -1)
	if len(m) != 1 {
		return ""
	}
	return m[0][1]
}

// The model of kinds. Every formula the pool and the edits hold names its
// kind with a marker, and an edit keeps the kind of every formula it
// keeps, wherever the formula goes:
//
//   - set, write and script setFormula add formulas of the kinds their text
//     names: the kind Excel 365 stores for that formula typed in (an array
//     result, or a range where one value goes, makes a dynamic array);
//   - copy pastes a formula as the kind it was (--values pastes none);
//     move, sort, rows and cols edits, sheet rename and copy, and replace
//     (also in formula text: it renames functions only) keep kinds where
//     the formulas go;
//   - clear and edits that overwrite a cell remove its formula.
//
// So after any edit, a formula's stored kind is the kind its marker names.
// Formulas without exactly one marker are outside the model and not
// judged.

// checkFile checks a saved workbook's own invariants: xml and kinds.
func checkFile(data []byte) []violation {
	vs, _ := inspect(data)
	return vs
}

// inspect checks a saved workbook's own invariants, and returns the formulas
// it read for them (none when the xml is broken).
func inspect(data []byte) ([]violation, []formulaCell) {
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return []violation{{"xml", "not a zip package: " + err.Error()}}, nil
	}
	var out []violation
	for _, zf := range zr.File {
		switch path.Ext(zf.Name) {
		case ".xml", ".rels", ".vml":
			if err := wellFormed(zf); err != nil {
				out = append(out, violation{"xml", fmt.Sprintf("%s: %v", zf.Name, err)})
			}
		}
	}
	if len(out) > 0 {
		return out, nil
	}
	if err := xlsxtest.CheckOrder(data); err != nil {
		return []violation{{"xml", err.Error()}}, nil
	}
	for _, zf := range zr.File {
		if strings.HasPrefix(zf.Name, "xl/worksheets/") && path.Ext(zf.Name) == ".xml" || zf.Name == "xl/sharedStrings.xml" {
			if err := checkSpace(zf); err != nil {
				return []violation{{"xml", fmt.Sprintf("%s: %v", zf.Name, err)}}, nil
			}
		}
	}
	cells, metadata, err := formulas(zr)
	if err != nil {
		return []violation{{"xml", err.Error()}}, nil
	}
	return checkKinds(cells, metadata), cells
}

// checkSpace refuses a text (a <t> element of a sheet or the shared strings)
// that starts or ends with white space and does not say xml:space="preserve":
// Excel drops those spaces when it reads the file. (A streaming editor that
// writes text and forgets the attribute loses the spaces of " lead" and "x ",
// and sheets' own reader, which keeps them, never notices.)
func checkSpace(zf *zip.File) error {
	r, err := zf.Open()
	if err != nil {
		return err
	}
	defer r.Close()
	d := xml.NewDecoder(bufio.NewReader(r))
	var text *strings.Builder
	preserve := false
	for {
		tok, err := d.Token()
		if err != nil {
			return nil // wellFormed has judged the part
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "t" && text == nil {
				text, preserve = &strings.Builder{}, false
				for _, a := range t.Attr {
					preserve = preserve || a.Name.Local == "space" && a.Value == "preserve"
				}
			}
		case xml.CharData:
			if text != nil {
				text.Write(t)
			}
		case xml.EndElement:
			if t.Name.Local == "t" && text != nil {
				if v := text.String(); !preserve && strings.TrimSpace(v) != v {
					return fmt.Errorf("the text %q has spaces at its ends and no xml:space=\"preserve\": Excel drops them", v)
				}
				text = nil
			}
		}
	}
}

func wellFormed(zf *zip.File) error {
	r, err := zf.Open()
	if err != nil {
		return err
	}
	defer r.Close()
	return strictXML(r)
}

// strictXML reads a part as XML 1.0 with namespaces, and returns what is not
// well-formed. Go's decoder (Token) accepts two kinds of XML that expat and
// Excel refuse as corrupt, so this reads raw tokens and checks them itself:
//
//   - an attribute twice on one element, by name (<c r="B2" r="B2">) or by
//     namespace and local name (two prefixes bound to one namespace);
//   - a prefix no namespace declaration in scope binds (<zz:x/>).
//
// A streaming editor that rewrites a cell's attributes with a regular
// expression, and misses one the cell already has, writes the first.
func strictXML(r io.Reader) error {
	br := bufio.NewReader(r)
	if b, _ := br.Peek(3); bytes.Equal(b, []byte("\xef\xbb\xbf")) {
		br.Discard(3) // a byte order mark is no text
	}
	d := xml.NewDecoder(br)
	type open struct {
		name xml.Name // as written: the prefix is in Space
		ns   map[string]string
	}
	var stack []open
	roots, tokens := 0, 0
	// bound is the namespace a prefix names where the element is.
	bound := func(prefix string) (string, bool) {
		if prefix == "xml" {
			return "http://www.w3.org/XML/1998/namespace", true
		}
		for i := len(stack) - 1; i >= 0; i-- {
			if uri, ok := stack[i].ns[prefix]; ok {
				return uri, true
			}
		}
		return "", false
	}
	for {
		tok, err := d.RawToken()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}
		tokens++
		switch t := tok.(type) {
		case xml.StartElement:
			if len(stack) == 0 {
				if roots++; roots > 1 {
					return fmt.Errorf("more than one root element, <%s> the second", showName(t.Name))
				}
			}
			ns := map[string]string{}
			for _, a := range t.Attr {
				switch {
				case a.Name.Space == "xmlns":
					ns[a.Name.Local] = a.Value
				case a.Name.Space == "" && a.Name.Local == "xmlns":
					ns[""] = a.Value
				}
			}
			stack = append(stack, open{t.Name, ns})
			if t.Name.Space != "" {
				if _, ok := bound(t.Name.Space); !ok {
					return fmt.Errorf("<%s>: the prefix %q is not declared", showName(t.Name), t.Name.Space)
				}
			}
			seen := map[xml.Name]bool{}
			for _, a := range t.Attr {
				key := a.Name // declarations and unprefixed attributes: by name as written
				if a.Name.Space != "" && a.Name.Space != "xmlns" {
					uri, ok := bound(a.Name.Space)
					if !ok {
						return fmt.Errorf("<%s>: the prefix %q of attribute %s is not declared", showName(t.Name), a.Name.Space, showName(a.Name))
					}
					key = xml.Name{Space: uri, Local: a.Name.Local}
				}
				if seen[key] {
					return fmt.Errorf("<%s>: attribute %s twice", showName(t.Name), showName(a.Name))
				}
				seen[key] = true
			}
		case xml.EndElement:
			if len(stack) == 0 {
				return fmt.Errorf("</%s> closes nothing", showName(t.Name))
			}
			if top := stack[len(stack)-1]; top.name != t.Name {
				return fmt.Errorf("<%s> closed by </%s>", showName(top.name), showName(t.Name))
			}
			stack = stack[:len(stack)-1]
		case xml.CharData:
			if len(stack) == 0 && len(bytes.TrimSpace(t)) > 0 {
				return fmt.Errorf("text outside the root element")
			}
		case xml.ProcInst:
			if t.Target == "xml" && tokens > 1 {
				return fmt.Errorf("the xml declaration is not first")
			}
		}
	}
	if len(stack) != 0 || roots != 1 {
		return fmt.Errorf("%d root elements, %d left open", roots, len(stack))
	}
	return nil
}

func showName(n xml.Name) string {
	if n.Space != "" {
		return n.Space + ":" + n.Local
	}
	return n.Local
}

// formulaCell is a formula as a worksheet stores it.
type formulaCell struct {
	sheet, cell string
	kind        string // Plain, Legacy, Dynamic
	text        string // the formula; a shared formula's follower has its master's
	ref         string // an array's range
	cm          int
	shared      bool   // stored as a shared formula (the master or a follower)
	typ, value  string // the cell's t attribute, and its stored result (<v>)
}

// The stored forms of a formula's result.
const (
	resNumber    = "number"
	resText      = "text"
	resEmptyText = "empty-text" // t="str" with an empty <v>: a formula that gives ""
	resBoolean   = "boolean"
	resError     = "error"
	resNone      = "no-result" // no <v>, or an empty one with no type: pending
	resShared    = "shared"    // not a result: stored as a shared formula
)

// Results are the stored forms of a formula's result, and the storage form
// that is no result, the run counts in the files it saves.
var Results = []string{resNumber, resText, resEmptyText, resBoolean, resError, resNone, resShared}

// result is the stored form of the formula's result.
func (f formulaCell) result() string {
	switch f.typ {
	case "b":
		return resBoolean
	case "e":
		return resError
	case "str", "s", "inlineStr":
		if f.value == "" {
			return resEmptyText
		}
		return resText
	}
	if f.value == "" {
		return resNone
	}
	return resNumber
}

// checkKinds checks every formula's stored kind against its marker, and
// that each array is stored as Excel requires: its range starts at its
// cell, and a dynamic array's cell metadata exists.
func checkKinds(cells []formulaCell, metadata int) []violation {
	var out []violation
	for _, f := range cells {
		where := f.sheet + " " + f.cell
		if f.kind != Plain {
			rg, err := grid.ParseRange(f.ref)
			r, c, _ := grid.ParseCell(f.cell)
			if err != nil || rg.R1 != r || rg.C1 != c {
				out = append(out, violation{"kinds", fmt.Sprintf("%s: %s with range %q, which does not start at its cell", where, kindName[f.kind], f.ref)})
				continue
			}
		}
		if f.cm > 0 && f.cm > metadata {
			out = append(out, violation{"kinds", fmt.Sprintf("%s: cell metadata %d, but the workbook has %d", where, f.cm, metadata)})
			continue
		}
		want := marker(f.text)
		if want != "" && want != f.kind {
			out = append(out, violation{"kinds", fmt.Sprintf("%s: =%s is stored as a %s, not a %s", where, f.text, kindName[f.kind], kindName[want])})
		}
	}
	return out
}

// formulas reads every formula of every worksheet, by part name order, and
// the number of cell metadata records (xl/metadata.xml).
func formulas(zr *zip.Reader) ([]formulaCell, int, error) {
	var out []formulaCell
	metadata := 0
	for _, zf := range zr.File {
		if zf.Name == "xl/metadata.xml" {
			n, err := cellMetadata(zf)
			if err != nil {
				return nil, 0, fmt.Errorf("%s: %v", zf.Name, err)
			}
			metadata = n
		}
		if path.Dir(zf.Name) != "xl/worksheets" || path.Ext(zf.Name) != ".xml" {
			continue
		}
		cells, err := sheetFormulas(zf)
		if err != nil {
			return nil, 0, fmt.Errorf("%s: %v", zf.Name, err)
		}
		out = append(out, cells...)
	}
	return out, metadata, nil
}

func cellMetadata(zf *zip.File) (int, error) {
	r, err := zf.Open()
	if err != nil {
		return 0, err
	}
	defer r.Close()
	d := xml.NewDecoder(r)
	in, n := false, 0
	for {
		tok, err := d.Token()
		if err == io.EOF {
			return n, nil
		}
		if err != nil {
			return 0, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "cellMetadata" {
				in = true
			} else if in && t.Name.Local == "bk" {
				n++
			}
		case xml.EndElement:
			if t.Name.Local == "cellMetadata" {
				in = false
			}
		}
	}
}

func sheetFormulas(zf *zip.File) ([]formulaCell, error) {
	r, err := zf.Open()
	if err != nil {
		return nil, err
	}
	defer r.Close()
	d := xml.NewDecoder(r)
	sheet := path.Base(zf.Name)
	shared := map[string]string{} // si -> the master's text
	var out []formulaCell
	var cell, typ string
	var cm int
	last := -1 // the formula of the cell being read, in out
	for {
		tok, err := d.Token()
		if err == io.EOF {
			return out, nil
		}
		if err != nil {
			return nil, err
		}
		se, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		switch se.Name.Local {
		case "c":
			cell, typ, cm, last = attr(se, "r"), attr(se, "t"), 0, -1
			if v := attr(se, "cm"); v != "" {
				cm, _ = strconv.Atoi(v)
			}
		case "v":
			var text string
			if err := d.DecodeElement(&text, &se); err != nil {
				return nil, err
			}
			if last >= 0 {
				out[last].value = text
			}
		case "f":
			t, ref, si := attr(se, "t"), attr(se, "ref"), attr(se, "si")
			var text string
			if err := d.DecodeElement(&text, &se); err != nil {
				return nil, err
			}
			f := formulaCell{sheet: sheet, cell: cell, kind: Plain, text: text, ref: ref, cm: cm, typ: typ, shared: t == "shared"}
			switch t {
			case "array":
				f.kind = Legacy
				if cm > 0 {
					f.kind = Dynamic
				}
			case "shared":
				if text != "" {
					shared[si] = text
				} else {
					f.text = shared[si]
				}
			}
			if f.text != "" {
				out = append(out, f)
				last = len(out) - 1
			}
		}
	}
}

func attr(se xml.StartElement, name string) string {
	for _, a := range se.Attr {
		if a.Name.Local == name && a.Name.Space == "" {
			return a.Value
		}
	}
	return ""
}

// compareEditors compares what the two editors saved: sheets, cells,
// formulas, kinds and values. It returns the first difference, "" when
// there is none, and the streaming editor's workbook, parsed.
func compareEditors(stream, full []byte) (*grid.Workbook, string) {
	a, err := readBytes(stream)
	if err != nil {
		return nil, "the streaming editor's file does not read: " + err.Error()
	}
	b, err := readBytes(full)
	if err != nil {
		return a, "the full editor's file does not read: " + err.Error()
	}
	if an, bn := strings.Join(a.SheetNames(), ", "), strings.Join(b.SheetNames(), ", "); an != bn {
		return a, fmt.Sprintf("sheets: stream has %s, full has %s", an, bn)
	}
	for i, sa := range a.Sheets {
		sb := b.Sheets[i]
		if d := compareSheets(sa, sb); d != "" {
			return a, sa.Name + "!" + d
		}
	}
	return a, ""
}

func compareSheets(a, b *grid.Sheet) string {
	ua, ub := a.Used(), b.Used()
	r2, c2 := max(ua.R2, ub.R2), max(ua.C2, ub.C2)
	for r := 1; r <= r2; r++ {
		for c := 1; c <= c2; c++ {
			ca, cb := a.Get(r, c), b.Get(r, c)
			if d := compareCells(ca, cb); d != "" {
				return grid.CellName(r, c) + ": " + d
			}
		}
	}
	if d := compareRanges("dynamic arrays", a.Spills, b.Spills); d != "" {
		return d
	}
	return compareRanges("arrays", a.Arrays, b.Arrays)
}

func compareCells(a, b grid.Cell) string {
	if a.Formula != b.Formula {
		return fmt.Sprintf("stream has =%s, full has =%s", a.Formula, b.Formula)
	}
	va, vb := value(a), value(b)
	if va != vb {
		return fmt.Sprintf("stream holds %s, full holds %s", va, vb)
	}
	return ""
}

// value describes a cell's stored value, and whether it is pending.
func value(c grid.Cell) string {
	var s string
	switch c.Kind {
	case grid.Empty:
		s = "nothing"
	case grid.Number:
		s = "number " + strconv.FormatFloat(c.Num, 'g', -1, 64)
		if math.IsNaN(c.Num) {
			s = "number NaN"
		}
	case grid.Bool:
		s = fmt.Sprintf("boolean %v", c.Bool)
	default:
		s = fmt.Sprintf("%v %q", c.Kind, c.Text)
	}
	if c.Pending {
		s += " (pending)"
	}
	return s
}

// compareRanges says how two editors' arrays differ, and which ranges one
// editor's file has alone ("(the full editor's alone: F4)"), so a known bug's
// face can tell an array one editor adds from one it loses.
func compareRanges(what string, a, b []grid.Range) string {
	sa, sb := rangeList(a), rangeList(b)
	if sa == sb {
		return ""
	}
	alone := func(x, y []grid.Range) string {
		var out []grid.Range
		for _, r := range x {
			if !slices.Contains(y, r) {
				out = append(out, r)
			}
		}
		if len(out) == 0 {
			return ""
		}
		return rangeList(out)
	}
	var notes []string
	if s := alone(a, b); s != "" {
		notes = append(notes, "the streaming editor's alone: "+s)
	}
	if s := alone(b, a); s != "" {
		notes = append(notes, "the full editor's alone: "+s)
	}
	d := fmt.Sprintf("%s: stream has %s, full has %s", what, sa, sb)
	if len(notes) > 0 {
		d += " (" + strings.Join(notes, ", ") + ")"
	}
	return d
}

func rangeList(rs []grid.Range) string {
	var s []string
	for _, r := range rs {
		s = append(s, r.String())
	}
	sort.Strings(s)
	if len(s) == 0 {
		return "none"
	}
	return strings.Join(s, " ")
}

// partsDiffer names the parts of two packages that differ.
func partsDiffer(want, got []byte) string {
	pw, err1 := parts(want)
	pg, err2 := parts(got)
	if err1 != nil || err2 != nil {
		return fmt.Sprintf("%d bytes, not the %d before", len(got), len(want))
	}
	var diff []string
	for n, w := range pw {
		if g, ok := pg[n]; !ok {
			diff = append(diff, n+" missing")
		} else if !bytes.Equal(w, g) {
			diff = append(diff, n+" differs")
		}
	}
	for n := range pg {
		if _, ok := pw[n]; !ok {
			diff = append(diff, n+" added")
		}
	}
	sort.Strings(diff)
	if len(diff) == 0 {
		return "the same parts, packed differently"
	}
	if len(diff) > 5 {
		diff = append(diff[:5], fmt.Sprintf("and %d more", len(diff)-5))
	}
	return strings.Join(diff, ", ")
}

func parts(data []byte) (map[string][]byte, error) {
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, err
	}
	m := map[string][]byte{}
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
		m[zf.Name] = b
	}
	return m, nil
}

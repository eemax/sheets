package xlsx

import (
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/xml"
	"html"
	"math"
	"regexp"
	"sort"
	"strconv"
	"strings"

	"github.com/eemax/sheets/internal/formula"
	"github.com/eemax/sheets/internal/grid"
)

// A file sheets saved tells Excel to recalculate on open, which also marks
// its stored results as possibly stale to sheets' own reader. The stamp
// vouches for them instead: a document property holding a fingerprint of
// the parts results are computed from. A tool that edits the file without
// keeping results current changes those parts, and the stamp no longer
// matches; Excel drops the recalculate flag when it saves.
const stampName = "sheets.values"

var stampPlaceholder = strings.Repeat("0", 64)

// valueParts lists the parts formula results depend on: the workbook (its
// names), the worksheets and the shared strings.
func (p *pkg) valueParts() ([]string, error) {
	wbPath, relsPath, sheets, err := p.workbookParts()
	if err != nil {
		return nil, err
	}
	names := []string{wbPath}
	for _, part := range sheets {
		names = append(names, part)
	}
	if rels, err := p.rels(relsPath, pathDir(wbPath)); err == nil {
		for _, r := range rels {
			if strings.HasSuffix(r.Type, "/sharedStrings") {
				names = append(names, r.Target)
			}
		}
	}
	return names, nil
}

func pathDir(p string) string {
	if i := strings.LastIndex(p, "/"); i >= 0 {
		return p[:i]
	}
	return ""
}

// fingerprint hashes the given parts, read through read.
func fingerprint(names []string, read func(name string) []byte) string {
	sorted := make([]string, len(names))
	for i, n := range names {
		sorted[i] = strings.ToLower(strings.TrimPrefix(n, "/"))
	}
	sort.Strings(sorted)
	h := sha256.New()
	for i, n := range sorted {
		if i > 0 && n == sorted[i-1] {
			continue
		}
		sum := sha256.Sum256(read(n))
		h.Write([]byte(n))
		h.Write([]byte{0})
		h.Write(sum[:])
	}
	return hex.EncodeToString(h.Sum(nil))
}

// stamped reports whether a package carries a stamp matching its parts.
func (p *pkg) stamped() bool {
	rels, err := p.rels("_rels/.rels", "")
	if err != nil {
		return false
	}
	var props string
	for _, r := range rels {
		if strings.HasSuffix(r.Type, "/custom-properties") {
			props = r.Target
		}
	}
	if props == "" {
		return false
	}
	var x struct {
		Props []struct {
			Name  string `xml:"name,attr"`
			Value string `xml:"lpwstr"`
		} `xml:"property"`
	}
	if err := p.decode(props, &x); err != nil {
		return false
	}
	want := ""
	for _, pr := range x.Props {
		if pr.Name == stampName {
			want = strings.TrimPrefix(pr.Value, "v1 ")
		}
	}
	if len(want) != 64 {
		return false
	}
	names, err := p.valueParts()
	if err != nil {
		return false
	}
	return fingerprint(names, func(n string) []byte {
		b, _ := p.read(n)
		return b
	}) == want
}

var (
	valueElem = regexp.MustCompile(`(?s)<v>.*?</v>|<v\s*/>`)
	typeAttr  = regexp.MustCompile(`\st="[^"]*"`)
	formulaEl = regexp.MustCompile(`(?s)<f\b[^>]*/>|<f\b[^>]*>.*?</f>`)
	arrayType = regexp.MustCompile(`^<f\b[^>]*\st="array"`)
	refOfF    = regexp.MustCompile(`^<f\b[^>]*\sref="([^"]+)"`)
)

// rewriteCells passes the cells of a worksheet part that want selects, by
// reference, through fn. It scans the part once, without a regular
// expression: sheets can run to hundreds of megabytes.
func rewriteCells(b []byte, want func(ref string) bool, fn func(cell []byte, ref string) []byte) []byte {
	var out []byte
	last, i := 0, 0
	for {
		j := bytes.Index(b[i:], []byte("<c"))
		if j < 0 {
			break
		}
		start := i + j
		i = start + 2
		if i >= len(b) || b[i] != ' ' && b[i] != '>' && b[i] != '\t' && b[i] != '\n' && b[i] != '\r' {
			continue // <col>, <cfRule>, ...
		}
		k := bytes.IndexByte(b[i:], '>')
		if k < 0 {
			break
		}
		tagEnd := i + k + 1
		tag := b[start:tagEnd]
		if tag[len(tag)-2] == '/' {
			i = tagEnd
			continue // empty cell
		}
		ref := ""
		if a := bytes.Index(tag, []byte(` r="`)); a >= 0 {
			if z := bytes.IndexByte(tag[a+4:], '"'); z >= 0 {
				ref = string(tag[a+4 : a+4+z])
			}
		}
		close := bytes.Index(b[tagEnd:], []byte("</c>"))
		if close < 0 {
			break
		}
		end := tagEnd + close + len("</c>")
		i = end
		if ref == "" || !want(ref) {
			continue
		}
		out = append(out, b[last:start]...)
		out = append(out, fn(b[start:end], ref)...)
		last = end
	}
	if out == nil {
		return b
	}
	return append(out, b[last:]...)
}

// storeResults writes formula results into a worksheet part: res maps a
// cell to its result, an Empty one meaning none. The other cells of an
// array formula given a result (the rest of its spill range) hold values
// of the old result, and lose them.
func storeResults(b []byte, res map[string]grid.Cell) []byte {
	var arrays []grid.Range
	b = rewriteCells(b, func(ref string) bool { _, ok := res[ref]; return ok }, func(m []byte, ref string) []byte {
		out, array := storeResult(m, res[ref])
		if array != nil {
			arrays = append(arrays, *array)
		}
		return out
	})
	if len(arrays) == 0 {
		return b
	}
	inArray := func(ref string) bool {
		r, c, _ := grid.ParseCell(ref)
		for _, rg := range arrays {
			if rg.Contains(r, c) {
				return true
			}
		}
		return false
	}
	return rewriteCells(b, inArray, func(m []byte, _ string) []byte {
		if formulaEl.Match(m) {
			return m
		}
		head := m[:bytes.IndexByte(m, '>')+1]
		body := m[len(head) : len(m)-len("</c>")]
		return append(append(typeAttr.ReplaceAll(head, nil), valueElem.ReplaceAll(body, nil)...), "</c>"...)
	})
}

// storeResult stores v as the result of the formula cell m (a whole <c>
// element); an Empty v removes its result. A multi-cell array formula's
// range comes back too (nil for others): the stored values of its other
// cells no longer hold.
func storeResult(m []byte, v grid.Cell) ([]byte, *grid.Range) {
	head := m[:bytes.IndexByte(m, '>')+1]
	if head[len(head)-2] == '/' {
		return m, nil // no formula in an empty cell
	}
	body := m[len(head) : len(m)-len("</c>")]
	f := formulaEl.Find(body)
	if f == nil {
		return m, nil // no longer a formula
	}
	var array *grid.Range
	if a := refOfF.FindSubmatch(f); a != nil && arrayType.Match(f) {
		if rg, err := grid.ParseRange(string(a[1])); err == nil && (rg.Rows() > 1 || rg.Cols() > 1) {
			array = &rg
		}
	}
	rest := valueElem.ReplaceAll(bytes.Replace(body, f, nil, 1), nil)
	head = typeAttr.ReplaceAll(head, nil)
	typ, val := "", ""
	switch v.Kind {
	case grid.Number:
		if math.IsNaN(v.Num) || math.IsInf(v.Num, 0) {
			typ, val = "e", "#NUM!"
		} else {
			val = strconv.FormatFloat(v.Num, 'g', -1, 64)
		}
	case grid.Text:
		typ, val = "str", v.Text
	case grid.Bool:
		typ, val = "b", "0"
		if v.Bool {
			val = "1"
		}
	case grid.Error:
		typ, val = "e", v.Text
	}
	var out bytes.Buffer
	if typ != "" {
		out.Write(head[:len(head)-1])
		out.WriteString(` t="` + typ + `">`)
	} else {
		out.Write(head)
	}
	out.Write(f)
	if v.Kind != grid.Empty {
		out.WriteString("<v>")
		xml.EscapeText(&out, []byte(val))
		out.WriteString("</v>")
	}
	out.Write(rest)
	out.WriteString("</c>")
	return out.Bytes(), array
}

var (
	// With or without a namespace prefix: openpyxl writes none.
	chartRef   = regexp.MustCompile(`(?s)<(?:\w+:)?(?:numRef|strRef|multiLvlStrRef)>.*?</(?:\w+:)?(?:numRef|strRef|multiLvlStrRef)>`)
	chartF     = regexp.MustCompile(`(?s)(<(?:\w+:)?f>)(.*?)(</(?:\w+:)?f>)`)
	chartCache = regexp.MustCompile(`(?s)<(?:\w+:)?(?:numCache|strCache|multiLvlStrCache)\b(?:[^>]*/>|.*?</(?:\w+:)?(?:numCache|strCache|multiLvlStrCache)>)`)
)

// freshCaches brings a chart's data references up to date: their ranges
// follow moved rows and columns, and a reference whose range moved or
// takes in a changed cell loses its cached copy of the data, which Excel
// rebuilds on open, rather than show old numbers.
func freshCaches(b []byte, rewrites []func(string) string, touched map[string][][2]int) []byte {
	return chartRef.ReplaceAllFunc(b, func(m []byte) []byte {
		g := chartF.FindSubmatch(m)
		if g == nil {
			return m
		}
		f := html.UnescapeString(string(g[2]))
		moved := f
		for _, fn := range rewrites {
			moved = fn(moved)
		}
		stale := moved != f
		for _, r := range formula.Refs(moved) {
			for _, cell := range touched[strings.ToLower(r.Sheet)] {
				if r.Range.Contains(cell[0], cell[1]) {
					stale = true
				}
			}
		}
		if moved != f {
			m = chartF.ReplaceAllLiteral(m, append(append(append([]byte{}, g[1]...), xmlEscape(moved)...), g[3]...))
		}
		if stale {
			m = chartCache.ReplaceAll(m, nil)
		}
		return m
	})
}

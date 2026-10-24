// Package delim reads and writes delimited text files (CSV, TSV) as
// workbooks of one sheet. It keeps what a file looks like: its delimiter,
// encoding, byte order mark, line endings and decimal comma, and rows an
// edit left alone stay exactly as they were.
package delim

import (
	"bytes"
	"encoding/csv"
	"errors"
	"fmt"
	"io"
	"path/filepath"
	"strconv"
	"strings"
	"unicode/utf8"

	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/input"
	"golang.org/x/text/encoding"
	"golang.org/x/text/encoding/charmap"
	"golang.org/x/text/encoding/unicode"
)

// Dialect is how a file writes its rows.
type Dialect struct {
	Comma        rune
	CRLF         bool
	BOM          bool
	Encoding     string // utf-8, utf-16le, utf-16be or windows-1252
	DecimalComma bool   // numbers are written 1.234,56 (with ; between fields)
	FinalNewline bool   // the last row ends with a line break
	// Plain writes numbers without thousands separators or currency
	// symbols, as sheets shows them (ISO dates, 12%): for data to hand to
	// other programs. Otherwise they are written as typed (1,200, $5.50).
	Plain bool
}

// File is a delimited file as read. A blank line is an empty row, as in
// Excel.
type File struct {
	Dialect Dialect
	Rows    [][]string
	raw     [][]byte // each row's bytes as the file holds them, line break included
	tail    []byte   // blank lines after the last row
	width   int      // the usual number of fields in a row
}

// DefaultDialect is the dialect of a new file at path: tabs for .tsv,
// commas otherwise; UTF-8 without a byte order mark; LF line breaks.
func DefaultDialect(path string) Dialect {
	d := Dialect{Comma: ',', Encoding: "utf-8", FinalNewline: true}
	if strings.EqualFold(filepath.Ext(path), ".tsv") {
		d.Comma = '\t'
	}
	return d
}

// Parse reads a delimited file. path's extension decides between commas
// and tabs (.tsv); a .csv file using semicolons, tabs or pipes instead is
// recognised by its first lines.
func Parse(data []byte, path string) (*File, error) {
	text, enc, bom, err := decode(data)
	if err != nil {
		return nil, err
	}
	f := &File{Dialect: Dialect{Encoding: enc, BOM: bom}}
	f.Dialect.Comma = sniff(text, path)
	f.Dialect.CRLF = strings.Contains(firstLine(text), "\r\n")
	f.Dialect.FinalNewline = text == "" || strings.HasSuffix(text, "\n")
	cr := csv.NewReader(strings.NewReader(text))
	cr.Comma = f.Dialect.Comma
	cr.FieldsPerRecord = -1
	cr.LazyQuotes = true
	cr.ReuseRecord = false
	prev := 0
	for {
		rec, err := cr.Read()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("reading %s: %w", filepath.Base(path), err)
		}
		off := int(cr.InputOffset())
		// encoding/csv skips blank lines; each is an empty row.
		raw := text[prev:off]
		for {
			n := 0
			switch {
			case strings.HasPrefix(raw, "\r\n"):
				n = 2
			case strings.HasPrefix(raw, "\n"):
				n = 1
			}
			if n == 0 {
				break
			}
			f.Rows = append(f.Rows, nil)
			f.raw = append(f.raw, []byte(raw[:n]))
			raw = raw[n:]
		}
		f.Rows = append(f.Rows, rec)
		f.raw = append(f.raw, []byte(raw))
		prev = off
	}
	f.tail = []byte(text[prev:])
	f.Dialect.DecimalComma = f.Dialect.Comma == ';' && decimalComma(f.Rows)
	widths := map[int]int{}
	for _, row := range f.Rows {
		if len(row) > 0 {
			widths[len(row)]++
		}
	}
	for w, n := range widths {
		if n > widths[f.width] || n == widths[f.width] && w > f.width {
			f.width = w
		}
	}
	return f, nil
}

func firstLine(s string) string {
	if i := strings.IndexByte(s, '\n'); i >= 0 {
		return s[:i+1]
	}
	return s
}

// decode turns a file's bytes into text: UTF-8 (with or without a byte
// order mark), UTF-16 with one, or else Windows-1252, the usual encoding
// of CSV files Excel saves on Windows.
func decode(data []byte) (text, enc string, bom bool, err error) {
	switch {
	case bytes.HasPrefix(data, []byte{0xEF, 0xBB, 0xBF}):
		return string(data[3:]), "utf-8", true, nil
	case bytes.HasPrefix(data, []byte{0xFF, 0xFE}):
		b, err := unicode.UTF16(unicode.LittleEndian, unicode.ExpectBOM).NewDecoder().Bytes(data)
		return string(b), "utf-16le", true, err
	case bytes.HasPrefix(data, []byte{0xFE, 0xFF}):
		b, err := unicode.UTF16(unicode.BigEndian, unicode.ExpectBOM).NewDecoder().Bytes(data)
		return string(b), "utf-16be", true, err
	case utf8.Valid(data):
		return string(data), "utf-8", false, nil
	}
	b, err := charmap.Windows1252.NewDecoder().Bytes(data)
	return string(b), "windows-1252", false, err
}

func encoder(d Dialect) *encoding.Encoder {
	switch d.Encoding {
	case "utf-16le":
		return unicode.UTF16(unicode.LittleEndian, unicode.UseBOM).NewEncoder()
	case "utf-16be":
		return unicode.UTF16(unicode.BigEndian, unicode.UseBOM).NewEncoder()
	case "windows-1252":
		return charmap.Windows1252.NewEncoder()
	}
	return nil
}

// sniff picks the delimiter: tabs for .tsv, and otherwise the candidate
// that splits the first lines into the same number of fields, most of
// them.
func sniff(text, path string) rune {
	if strings.EqualFold(filepath.Ext(path), ".tsv") {
		return '\t'
	}
	var lines []string
	for _, l := range strings.SplitN(text, "\n", 12) {
		if strings.TrimSpace(l) != "" {
			lines = append(lines, l)
		}
	}
	if len(lines) > 10 {
		lines = lines[:10] // the last may be cut short
	}
	best, bestScore := ',', 0
	for _, c := range []rune{',', ';', '\t', '|'} {
		counts := map[int]int{}
		for _, l := range lines {
			counts[countOutsideQuotes(l, c)]++
		}
		n, most := 0, 0
		for k, v := range counts {
			if k > 0 && (v > most || v == most && k > n) {
				n, most = k, v
			}
		}
		if score := most * 100 / max(len(lines), 1); n > 0 && score > bestScore {
			best, bestScore = c, score
		}
	}
	return best
}

func countOutsideQuotes(l string, c rune) int {
	n, quoted := 0, false
	for _, r := range l {
		switch {
		case r == '"':
			quoted = !quoted
		case r == c && !quoted:
			n++
		}
	}
	return n
}

// decimalComma tells whether numbers use a decimal comma: some fields read
// like 3,5 and none like 3.5. A mark followed by three digits (1,234 or
// 1.234) could group thousands either way, and tells nothing.
func decimalComma(rows [][]string) bool {
	comma, point := 0, 0
	for _, row := range rows {
		for _, f := range row {
			f = strings.TrimSpace(f)
			if numberWith(f, ',', '.') {
				comma++
			} else if numberWith(f, '.', ',') {
				point++
			}
		}
	}
	return comma > 0 && point == 0
}

// numberWith tells whether s is a number with dec as its decimal mark (and
// group between thousands) whose decimal part is not three digits long.
func numberWith(s string, dec, group byte) bool {
	s = strings.TrimLeft(s, "+-")
	i := strings.LastIndexByte(s, dec)
	if i <= 0 || i == len(s)-1 || len(s)-i-1 == 3 || strings.IndexByte(s[:i], dec) >= 0 {
		return false
	}
	for _, r := range s[i+1:] {
		if r < '0' || r > '9' {
			return false
		}
	}
	for _, r := range s[:i] {
		if (r < '0' || r > '9') && byte(r) != group {
			return false
		}
	}
	return true
}

// Value reads a field as a typed value, the way sheets reads typed input:
// numbers, dates, percentages, currency. A field starting with = stays text:
// a CSV file holds values, not formulas.
func (d Dialect) Value(field string) input.Value {
	if d.DecimalComma {
		field = swapMarks(field)
	}
	v := input.Parse(field, false)
	if v.Formula != "" {
		return input.Value{Kind: grid.Text, Text: field}
	}
	return v
}

// swapMarks turns 1.234,56 into 1,234.56 in a field that is a number (in
// a file with decimal commas), and leaves other fields alone.
func swapMarks(field string) string {
	t := strings.TrimSpace(field)
	core := strings.TrimRight(strings.TrimLeft(t, "+-$€£¥ "), "%€ ")
	if core == "" || strings.Trim(core, "0123456789.,") != "" || !strings.ContainsAny(core, ".,") {
		return field
	}
	return strings.NewReplacer(".", ",", ",", ".").Replace(field)
}

// Workbook is the file as a workbook of one sheet, named sheet.
func (f *File) Workbook(sheet string) *grid.Workbook {
	wb := &grid.Workbook{Format: "csv", Formats: []grid.Format{{}}}
	formats := map[grid.Format]uint16{{}: 0}
	s := grid.NewSheet(sheet)
	for r, row := range f.Rows {
		for c, field := range row {
			if field == "" {
				continue
			}
			v := f.Dialect.Value(field)
			cell := grid.Cell{Kind: v.Kind, Text: v.Text, Num: v.Num, Bool: v.Bool}
			if v.Kind == grid.Number && v.Fmt.Kind != grid.General {
				i, ok := formats[v.Fmt]
				if !ok {
					i = uint16(len(wb.Formats))
					formats[v.Fmt] = i
					wb.Formats = append(wb.Formats, v.Fmt)
				}
				cell.Fmt = i
			}
			if !cell.IsEmpty() {
				s.Set(r+1, c+1, cell)
			}
		}
	}
	wb.Sheets = []*grid.Sheet{s}
	return wb
}

// SheetName names the sheet of a file after it, as Excel does: its name
// without the extension, cut to 31 characters, without characters a sheet
// name cannot hold.
func SheetName(path string) string {
	base := strings.TrimSuffix(filepath.Base(path), filepath.Ext(path))
	base = strings.Map(func(r rune) rune {
		if strings.ContainsRune(`[]:*?/\`, r) {
			return '_'
		}
		return r
	}, base)
	base = strings.Trim(base, "'")
	if r := []rune(base); len(r) > 31 {
		base = string(r[:31])
	}
	if base == "" {
		base = "Sheet1"
	}
	return base
}

// key identifies a value, to recognise the fields of a row no edit changed.
func key(v grid.Cell) string {
	switch v.Kind {
	case grid.Number:
		return "n" + strconv.FormatFloat(v.Num, 'g', -1, 64)
	case grid.Bool:
		return map[bool]string{true: "b1", false: "b0"}[v.Bool]
	case grid.Text:
		return "t" + v.Text
	case grid.Error:
		return "e" + v.Text
	}
	return ""
}

func rowKey(keys []string) string {
	for len(keys) > 0 && keys[len(keys)-1] == "" {
		keys = keys[:len(keys)-1]
	}
	return strings.Join(keys, "\x1f")
}

// ErrUnknown is returned for a cell whose value is unknown (a formula the
// sheets calculator cannot compute).
var ErrUnknown = errors.New("value unknown")

// Encode writes the values of a sheet as a file of dialect d: every row
// from row 1, every column from A. orig, when given, is the file the
// sheet was read from: a row whose values are those of one of its rows is
// written as that row was, byte for byte, and other values keep the text
// they had in it (1,200 stays 1,200). value gives a cell's value; it
// returns ErrUnknown for one sheets cannot know.
func Encode(rows, cols int, value func(row, col int) (grid.Cell, grid.Format, error), d Dialect, orig *File, date1904 bool) ([]byte, error) {
	// Rows of the original file by their values; texts by value, per column
	// first.
	byKey := map[string][]int{}
	textByCol := map[[2]string]string{}
	textByKey := map[string]string{}
	// Rows are as wide as the sheet's values, or as the file's rows were
	// when its columns did not change (a file can end each row with an
	// empty field).
	width := cols
	if orig != nil {
		used := 0
		for _, row := range orig.Rows {
			for c := len(row); c > used; c-- {
				if row[c-1] != "" {
					used = c
					break
				}
			}
		}
		if used == cols {
			width = max(cols, orig.width)
		}
		for i, row := range orig.Rows {
			keys := make([]string, len(row))
			for c, field := range row {
				v := orig.Dialect.Value(field)
				k := key(grid.Cell{Kind: v.Kind, Text: v.Text, Num: v.Num, Bool: v.Bool})
				keys[c] = k
				if k == "" {
					continue
				}
				if _, ok := textByCol[[2]string{strconv.Itoa(c), k}]; !ok {
					textByCol[[2]string{strconv.Itoa(c), k}] = field
				}
				if _, ok := textByKey[k]; !ok {
					textByKey[k] = field
				}
			}
			rk := rowKey(keys)
			byKey[rk] = append(byKey[rk], i)
		}
	}
	nl := "\n"
	if d.CRLF {
		nl = "\r\n"
	}
	var out bytes.Buffer
	var line bytes.Buffer
	cw := csv.NewWriter(&line)
	cw.Comma = d.Comma
	cw.UseCRLF = d.CRLF
	for r := 1; r <= rows; r++ {
		keys := make([]string, cols)
		fields := make([]string, cols)
		for c := 1; c <= cols; c++ {
			v, f, err := value(r, c)
			if err != nil {
				return nil, err
			}
			k := key(v)
			keys[c-1] = k
			switch t, ok := textByCol[[2]string{strconv.Itoa(c - 1), k}]; {
			case k == "":
			case ok:
				fields[c-1] = t
			default:
				if t, ok := textByKey[k]; ok {
					fields[c-1] = t
				} else {
					fields[c-1] = Text(v, f, d, date1904)
				}
			}
		}
		if q := byKey[rowKey(keys)]; len(q) > 0 {
			raw := orig.raw[q[0]]
			byKey[rowKey(keys)] = q[1:]
			out.Write(bytes.TrimRight(raw, "\r\n"))
			out.WriteString(nl)
			continue
		}
		for len(fields) < width {
			fields = append(fields, "")
		}
		line.Reset()
		if err := cw.Write(fields); err != nil {
			return nil, err
		}
		cw.Flush()
		out.Write(line.Bytes())
	}
	b := out.Bytes()
	if orig != nil {
		b = append(b, orig.tail...)
	}
	if !d.FinalNewline {
		b = bytes.TrimSuffix(b, []byte(nl))
	}
	if e := encoder(d); e != nil {
		enc, err := e.Bytes(b)
		if err != nil {
			return nil, fmt.Errorf("writing %s: %w", d.Encoding, err)
		}
		return enc, nil
	}
	if d.BOM {
		b = append([]byte{0xEF, 0xBB, 0xBF}, b...)
	}
	return b, nil
}

// Text writes a value as a field: the way sheets reads it back to the same
// value and format (ISO dates, 12%, $1,200.50).
func Text(v grid.Cell, f grid.Format, d Dialect, date1904 bool) string {
	var t string
	switch v.Kind {
	case grid.Text, grid.Error:
		return v.Text
	case grid.Bool:
		if v.Bool {
			return "TRUE"
		}
		return "FALSE"
	case grid.Number:
		if d.Plain {
			t = grid.Display(&v, f, date1904)
		} else {
			t = numberText(v.Num, f, date1904)
		}
	default:
		return ""
	}
	if d.DecimalComma && v.Kind == grid.Number && !f.Kind.IsTemporal() {
		t = strings.NewReplacer(".", ",", ",", ".").Replace(t)
	}
	return t
}

func numberText(n float64, f grid.Format, date1904 bool) string {
	switch f.Kind {
	case grid.Date, grid.DateTime, grid.Time:
		return grid.FormatTemporal(n, f.Kind, date1904)
	case grid.Percent:
		return grid.FormatNumber(n*100) + "%"
	case grid.Integer, grid.Decimal:
		return grouped(n, f.Decimals)
	case grid.Currency:
		sym := f.Symbol
		if sym == "" {
			sym = "$"
		}
		neg := n < 0
		body := grouped(abs(n), f.Decimals)
		t := sym + body
		if !prefixSymbol[sym] {
			t = body + " " + sym
		}
		if neg {
			t = "-" + t
		}
		return t
	}
	return grid.FormatNumber(n)
}

// prefixSymbol are the currency symbols written before the number.
var prefixSymbol = map[string]bool{"$": true, "€": true, "£": true, "¥": true, "₹": true, "₽": true, "₩": true, "₺": true, "₪": true, "₫": true, "₱": true, "฿": true, "₴": true, "₦": true, "R$": true}

func abs(n float64) float64 {
	if n < 0 {
		return -n
	}
	return n
}

// grouped writes a number with thousands separators and at least the given
// decimals (more when it has them).
func grouped(n float64, decimals int) string {
	s := grid.FormatNumber(n)
	neg := strings.HasPrefix(s, "-")
	s = strings.TrimPrefix(s, "-")
	whole, frac, _ := strings.Cut(s, ".")
	if strings.ContainsAny(whole, "e") {
		return grid.FormatNumber(n)
	}
	for len(frac) < decimals {
		frac += "0"
	}
	var b strings.Builder
	for i, r := range whole {
		if i > 0 && (len(whole)-i)%3 == 0 {
			b.WriteByte(',')
		}
		b.WriteRune(r)
	}
	t := b.String()
	if frac != "" {
		t += "." + frac
	}
	if neg {
		t = "-" + t
	}
	return t
}

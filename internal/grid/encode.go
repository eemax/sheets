package grid

import (
	"bytes"
	"encoding/binary"
	"encoding/gob"
	"errors"
	"fmt"
	"math"
)

// The binary form of a workbook, for caching parsed files: a gob header
// with everything but the cells, a table of the distinct strings, then the
// cells row by row, each a varint of flags followed by the fields it has.
// Large sheets decode from it many times faster than from their xlsx.

type header struct {
	Format   string
	Names    []DefinedName
	Formats  []Format
	Date1904 bool
	Styled   int
	Sheets   []sheetHeader
}

type sheetHeader struct {
	Name                   string
	Hidden                 bool
	Merges                 []Range
	Tables                 []NamedTable
	Charts                 []Chart
	HiddenRows, HiddenCols []Run
	Filter                 Range
	Notes                  []Note
	Links                  []Link
	Spills                 []Range
	Arrays                 []Range
	Rows, Cells            int // stored rows, and cells in them
	MaxCol                 int
}

// Cell flags; the low three bits are the Kind.
const (
	flagBool = 1 << (3 + iota)
	flagPending
	flagComputed
	flagShared
	flagFmt
	flagNum
	flagText
	flagFormula
)

// MarshalBinary encodes the sheets, names and formats. Path, Size and
// Resolver are not kept.
func (w *Workbook) MarshalBinary() ([]byte, error) {
	h := header{Format: w.Format, Names: w.Names, Formats: w.Formats, Date1904: w.Date1904, Styled: w.Styled}
	strs := map[string]uint64{}
	var table []string
	index := func(s string) uint64 {
		i, ok := strs[s]
		if !ok {
			i = uint64(len(table))
			strs[s] = i
			table = append(table, s)
		}
		return i
	}
	var cells []byte
	var tmp [binary.MaxVarintLen64]byte
	put := func(v uint64) { cells = append(cells, tmp[:binary.PutUvarint(tmp[:], v)]...) }
	for _, s := range w.Sheets {
		sh := sheetHeader{Name: s.Name, Hidden: s.Hidden, Merges: s.Merges, Tables: s.Tables, Charts: s.Charts,
			HiddenRows: s.HiddenRows, HiddenCols: s.HiddenCols, Filter: s.Filter, Notes: s.Notes, Links: s.Links, Spills: s.Spills, Arrays: s.Arrays, Rows: len(s.rows), MaxCol: s.maxCol}
		for _, row := range s.rows {
			sh.Cells += len(row)
			put(uint64(len(row)))
			for i := range row {
				c := &row[i]
				f := uint64(c.Kind)
				if c.Bool {
					f |= flagBool
				}
				if c.Pending {
					f |= flagPending
				}
				if c.Computed {
					f |= flagComputed
				}
				if c.Shared {
					f |= flagShared
				}
				if c.Fmt != 0 {
					f |= flagFmt
				}
				if c.Num != 0 || math.Signbit(c.Num) {
					f |= flagNum
				}
				if c.Text != "" {
					f |= flagText
				}
				if c.Formula != "" {
					f |= flagFormula
				}
				put(f)
				if c.Fmt != 0 {
					put(uint64(c.Fmt))
				}
				if f&flagNum != 0 {
					cells = binary.LittleEndian.AppendUint64(cells, math.Float64bits(c.Num))
				}
				if c.Text != "" {
					put(index(c.Text))
				}
				if c.Formula != "" {
					put(index(c.Formula))
				}
			}
		}
		h.Sheets = append(h.Sheets, sh)
	}

	var out bytes.Buffer
	if err := gob.NewEncoder(&out).Encode(h); err != nil {
		return nil, err
	}
	head := out.Bytes()
	buf := binary.AppendUvarint(nil, uint64(len(head)))
	buf = append(buf, head...)
	buf = binary.AppendUvarint(buf, uint64(len(table)))
	for _, s := range table {
		buf = binary.AppendUvarint(buf, uint64(len(s)))
	}
	for _, s := range table {
		buf = append(buf, s...)
	}
	return append(buf, cells...), nil
}

var errCorrupt = errors.New("corrupt workbook encoding")

// decoder reads varints, remembering the first error.
type decoder struct {
	b   []byte
	err error
}

func (d *decoder) uvarint() uint64 {
	if d.err != nil {
		return 0
	}
	v, n := binary.Uvarint(d.b)
	if n <= 0 {
		d.err = errCorrupt
		return 0
	}
	d.b = d.b[n:]
	return v
}

// count reads a length no larger than limit.
func (d *decoder) count(limit int) int {
	v := d.uvarint()
	if limit < 0 || v > uint64(limit) {
		d.err = errCorrupt
		return 0
	}
	return int(v)
}

func (d *decoder) bytes(n int) []byte {
	if d.err != nil || n > len(d.b) {
		d.err = errCorrupt
		return nil
	}
	b := d.b[:n]
	d.b = d.b[n:]
	return b
}

// UnmarshalBinary decodes what MarshalBinary wrote, checking it as it goes:
// damaged input gives an error, never a wrong workbook or a panic.
func (w *Workbook) UnmarshalBinary(data []byte) error {
	d := &decoder{b: data}
	head := d.bytes(d.count(len(data)))
	if d.err != nil {
		return d.err
	}
	var h header
	if err := gob.NewDecoder(bytes.NewReader(head)).Decode(&h); err != nil {
		return fmt.Errorf("%w: %v", errCorrupt, err)
	}

	// Strings are slices of one string holding them all.
	n := d.count(len(d.b))
	lens := make([]int, n)
	total := 0
	for i := range lens {
		lens[i] = d.count(len(d.b))
		total += lens[i]
	}
	all := string(d.bytes(total))
	if d.err != nil {
		return d.err
	}
	table := make([]string, n)
	for i, l := range lens {
		table[i], all = all[:l], all[l:]
	}
	str := func() string {
		i := d.uvarint()
		if i >= uint64(len(table)) {
			d.err = errCorrupt
			return ""
		}
		return table[i]
	}

	sheets := make([]*Sheet, 0, len(h.Sheets))
	for _, sh := range h.Sheets {
		// Every row and cell takes at least a byte, which bounds the
		// allocations a damaged header could ask for.
		if sh.Rows < 0 || sh.Cells < 0 || sh.Rows+sh.Cells > len(d.b) || sh.MaxCol < 0 || sh.MaxCol > MaxCols {
			return errCorrupt
		}
		s := &Sheet{Name: sh.Name, Hidden: sh.Hidden, Merges: sh.Merges, Tables: sh.Tables, Charts: sh.Charts,
			HiddenRows: sh.HiddenRows, HiddenCols: sh.HiddenCols, Filter: sh.Filter, Notes: sh.Notes, Links: sh.Links, Spills: sh.Spills, Arrays: sh.Arrays, maxCol: sh.MaxCol}
		s.rows = make([][]Cell, sh.Rows)
		cells := make([]Cell, sh.Cells)
		for r := range s.rows {
			k := d.count(len(cells))
			if k > 0 {
				// Capped, so growing a row copies it instead of
				// overwriting the next.
				s.rows[r] = cells[:k:k]
				cells = cells[k:]
			}
			for i := 0; i < k && d.err == nil; i++ {
				c := &s.rows[r][i]
				f := d.uvarint()
				if f >= flagFormula<<1 || Kind(f&7) > Error {
					return errCorrupt
				}
				c.Kind = Kind(f & 7)
				c.Bool = f&flagBool != 0
				c.Pending = f&flagPending != 0
				c.Computed = f&flagComputed != 0
				c.Shared = f&flagShared != 0
				if f&flagFmt != 0 {
					c.Fmt = uint16(d.count(len(h.Formats) - 1))
				}
				if f&flagNum != 0 {
					if b := d.bytes(8); b != nil {
						c.Num = math.Float64frombits(binary.LittleEndian.Uint64(b))
					}
				}
				if f&flagText != 0 {
					c.Text = str()
				}
				if f&flagFormula != 0 {
					c.Formula = str()
				}
			}
			if d.err != nil {
				return d.err
			}
		}
		if len(cells) != 0 {
			return errCorrupt
		}
		sheets = append(sheets, s)
	}
	if len(d.b) != 0 {
		return errCorrupt
	}
	w.Format, w.Names, w.Formats, w.Date1904, w.Styled, w.Sheets = h.Format, h.Names, h.Formats, h.Date1904, h.Styled, sheets
	return nil
}

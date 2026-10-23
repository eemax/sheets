package grid

import (
	"math"
	"math/rand"
	"reflect"
	"testing"
)

func sampleWorkbook() *Workbook {
	wb := &Workbook{
		Format:   "xlsx",
		Formats:  []Format{{Kind: General, Code: "General"}, {Kind: Currency, Decimals: 2, Symbol: "€", Code: "#,##0.00 €"}},
		Names:    []DefinedName{{Name: "Rate", Ref: "Data!$B$1"}, {Name: "Local", Ref: "Data!$A$1", Scope: "Data"}},
		Date1904: true,
	}
	s := NewSheet("Data")
	s.Hidden = true
	s.Merges = []Range{{R1: 1, C1: 1, R2: 1, C2: 3}}
	s.Tables = []NamedTable{{Name: "Sales", Range: Range{R1: 2, C1: 1, R2: 5, C2: 3}, Totals: 1, Filter: true}}
	s.Charts = []Chart{{Anchor: "E2", Type: "column", Title: "Sales", To: "L18", Legend: "b",
		Series: []ChartSeries{{Name: "Units", Values: "Data!$B$3:$B$5", Color: "2A78D6", PointColors: []string{"FF0000"}}}}}
	s.HiddenRows = []Run{{From: 4, To: 4}}
	s.HiddenCols = []Run{{From: 7, To: 9}}
	s.Filter = Range{R1: 2, C1: 1, R2: 5, C2: 3}
	s.Set(1, 1, Cell{Kind: Text, Text: "Title"})
	s.Set(3, 1, Cell{Kind: Text, Text: "North"})
	s.Set(3, 2, Cell{Kind: Number, Num: 1.5, Fmt: 1})
	s.Set(3, 3, Cell{Kind: Number, Num: 3, Formula: "B3*2", Shared: true, Computed: true})
	s.Set(4, 1, Cell{Kind: Text, Text: "North"}) // repeated string
	s.Set(4, 2, Cell{Kind: Number, Num: math.Copysign(0, -1)})
	s.Set(4, 3, Cell{Formula: "B4*2", Pending: true})
	s.Set(5, 1, Cell{Kind: Bool, Bool: true})
	s.Set(5, 2, Cell{Kind: Error, Text: "#DIV/0!", Formula: "1/0"})
	s.Set(5, 3, Cell{Kind: Number, Num: math.Inf(-1)})
	s.Set(5, 4, Cell{Kind: Number, Num: math.NaN()})
	s.Set(5, 5, Cell{Kind: Text, Text: ""}) // empty text is still text
	wb.Sheets = []*Sheet{s, NewSheet("Empty")}
	return wb
}

func sameCell(a, b Cell) bool {
	na, nb := a.Num, b.Num
	a.Num, b.Num = 0, 0
	return a == b && math.Float64bits(na) == math.Float64bits(nb)
}

func TestEncodeRoundTrip(t *testing.T) {
	wb := sampleWorkbook()
	data, err := wb.MarshalBinary()
	if err != nil {
		t.Fatal(err)
	}
	var got Workbook
	if err := got.UnmarshalBinary(data); err != nil {
		t.Fatal(err)
	}
	if got.Format != wb.Format || got.Date1904 != wb.Date1904 || !reflect.DeepEqual(got.Formats, wb.Formats) || !reflect.DeepEqual(got.Names, wb.Names) {
		t.Errorf("workbook fields differ: %+v", got)
	}
	if len(got.Sheets) != len(wb.Sheets) {
		t.Fatalf("%d sheets, want %d", len(got.Sheets), len(wb.Sheets))
	}
	for i, s := range wb.Sheets {
		g := got.Sheets[i]
		if g.Name != s.Name || g.Hidden != s.Hidden || g.maxCol != s.maxCol || len(g.rows) != len(s.rows) ||
			!reflect.DeepEqual(g.Merges, s.Merges) || !reflect.DeepEqual(g.Tables, s.Tables) || !reflect.DeepEqual(g.Charts, s.Charts) ||
			!reflect.DeepEqual(g.HiddenRows, s.HiddenRows) || !reflect.DeepEqual(g.HiddenCols, s.HiddenCols) || g.Filter != s.Filter {
			t.Errorf("sheet %s differs:\n got %+v\nwant %+v", s.Name, g, s)
			continue
		}
		for r := range s.rows {
			if len(g.rows[r]) != len(s.rows[r]) {
				t.Errorf("%s row %d has %d cells, want %d", s.Name, r+1, len(g.rows[r]), len(s.rows[r]))
				continue
			}
			for c := range s.rows[r] {
				if !sameCell(g.rows[r][c], s.rows[r][c]) {
					t.Errorf("%s %s = %+v, want %+v", s.Name, CellName(r+1, c+1), g.rows[r][c], s.rows[r][c])
				}
			}
		}
	}

	// Rows share one backing array; growing one must not touch the next.
	s := got.Sheets[0]
	s.Set(3, 6, Cell{Kind: Text, Text: "grown"})
	if c := s.Get(4, 1); c.Text != "North" {
		t.Errorf("growing row 3 changed A4 to %+v", c)
	}
}

// Damaged input gives an error, never a panic or a silently short workbook.
func TestDecodeDamaged(t *testing.T) {
	data, err := sampleWorkbook().MarshalBinary()
	if err != nil {
		t.Fatal(err)
	}
	for n := 0; n < len(data); n++ {
		var w Workbook
		if err := w.UnmarshalBinary(data[:n]); err == nil {
			t.Fatalf("truncated to %d of %d bytes: no error", n, len(data))
		}
	}
	if err := new(Workbook).UnmarshalBinary(append(data, 0)); err == nil {
		t.Error("trailing byte: no error")
	}
	rng := rand.New(rand.NewSource(1))
	for i := 0; i < 2000; i++ {
		b := append([]byte(nil), data...)
		for k := 0; k < 1+rng.Intn(4); k++ {
			b[rng.Intn(len(b))] = byte(rng.Intn(256))
		}
		func() {
			defer func() {
				if p := recover(); p != nil {
					t.Fatalf("panic on damaged input: %v", p)
				}
			}()
			new(Workbook).UnmarshalBinary(b)
		}()
	}
}

package delim

import (
	"strings"
	"testing"

	"github.com/eemax/sheets/internal/grid"
	"golang.org/x/text/encoding/charmap"
	"golang.org/x/text/encoding/unicode"
)

// encodeSame writes a file back from its own workbook, as an edit that
// changed nothing does.
func encodeSame(t *testing.T, f *File, change func(s *grid.Sheet)) string {
	t.Helper()
	wb := f.Workbook("S")
	s := wb.Sheets[0]
	if change != nil {
		change(s)
	}
	u := s.Used()
	rows, cols := 0, 0
	if !u.Empty() {
		rows, cols = u.R2, u.C2
	}
	out, err := Encode(rows, cols, func(r, c int) (grid.Cell, grid.Format, error) {
		x := s.Get(r, c)
		return x, wb.FormatOf(&x), nil
	}, f.Dialect, f, false)
	if err != nil {
		t.Fatal(err)
	}
	return string(out)
}

func TestRoundTripUnchanged(t *testing.T) {
	for name, text := range map[string]string{
		"crlf, quotes":       "a,b,c\r\n\"x, y\",\"say \"\"hi\"\"\",3\r\n",
		"no final newline":   "a,b\n1,2",
		"blank lines":        "a,b\n\n1,2\n\n\n",
		"ragged":             "a,b,c\n1\n1,2,3,4\n",
		"semicolons":         "a;b\n1,5;2\n",
		"multiline field":    "a,b\n\"one\ntwo\",2\n",
		"numbers as written": "n,m\n1.50,\"1,200\"\n$5,12%\n007,2026-03-01\n",
		"tabs":               "a\tb\n1\t2\n",
		"empty":              "",
	} {
		f, err := Parse([]byte(text), "x.csv")
		if err != nil {
			t.Fatalf("%s: %v", name, err)
		}
		if got := encodeSame(t, f, nil); got != text {
			t.Errorf("%s: wrote %q, want %q", name, got, text)
		}
	}
}

func TestEditKeepsTheRest(t *testing.T) {
	text := "Item,Amount,Code\r\nPen,\"1,200\",007\r\nInk,$5.50,008\r\n"
	f, err := Parse([]byte(text), "x.csv")
	if err != nil {
		t.Fatal(err)
	}
	got := encodeSame(t, f, func(s *grid.Sheet) {
		c := s.Get(3, 2)
		c.Num = 6.25
		s.Set(3, 2, c)
		s.Set(4, 1, grid.Cell{Kind: grid.Text, Text: "Pad"})
	})
	want := "Item,Amount,Code\r\nPen,\"1,200\",007\r\nInk,$6.25,008\r\nPad,,\r\n"
	if got != want {
		t.Errorf("got %q, want %q", got, want)
	}
}

func TestSniff(t *testing.T) {
	for text, want := range map[string]rune{
		"a,b,c\n1,2,3\n":         ',',
		"a;b;c\n1,5;2;3\n":       ';',
		"a\tb\n1\t2\n":           '\t',
		"a|b|c\n1|2|3\n":         '|',
		"\"x;y\",b\n\"1;2\",3\n": ',',
		"single\nvalues\n":       ',',
	} {
		if got := sniff(text, "f.csv"); got != want {
			t.Errorf("sniff(%q) = %q, want %q", text, got, want)
		}
	}
	if got := sniff("a,b\n", "f.tsv"); got != '\t' {
		t.Errorf(".tsv sniffed %q", got)
	}
}

func TestDecimalComma(t *testing.T) {
	f, err := Parse([]byte("Name;Amount;Rate;Big\nA;1.234,50;3,5%;1.234\nB;12,00;10%;2.000\n"), "eu.csv")
	if err != nil {
		t.Fatal(err)
	}
	if !f.Dialect.DecimalComma {
		t.Fatal("decimal comma not recognised")
	}
	wb := f.Workbook("S")
	s := wb.Sheets[0]
	for _, tt := range []struct {
		r, c int
		want float64
	}{{2, 2, 1234.5}, {2, 3, 0.035}, {2, 4, 1234}, {3, 2, 12}, {3, 4, 2000}} {
		if got := s.Get(tt.r, tt.c); got.Kind != grid.Number || got.Num != tt.want {
			t.Errorf("%s = %+v, want %v", grid.CellName(tt.r, tt.c), got, tt.want)
		}
	}
	// 1,234 alone could be either; with decimal points elsewhere it is not
	// a decimal comma.
	f, _ = Parse([]byte("a;b\n1,234;2.5\n"), "x.csv")
	if f.Dialect.DecimalComma {
		t.Error("decimal comma read into 1,234 next to 2.5")
	}
	// Written back with the file's marks.
	got := Text(grid.Cell{Kind: grid.Number, Num: 2.5}, grid.Format{}, Dialect{DecimalComma: true}, false)
	if got != "2,5" {
		t.Errorf("2.5 written as %q", got)
	}
}

func TestEncodings(t *testing.T) {
	cp, _ := charmap.Windows1252.NewEncoder().String("Name,City\nZoë,Köln\n")
	f, err := Parse([]byte(cp), "w.csv")
	if err != nil {
		t.Fatal(err)
	}
	if f.Dialect.Encoding != "windows-1252" || f.Rows[1][1] != "Köln" {
		t.Errorf("windows-1252: %+v %q", f.Dialect, f.Rows)
	}
	if got := encodeSame(t, f, nil); got != cp {
		t.Errorf("windows-1252 written back as %q", got)
	}
	u16, _ := unicode.UTF16(unicode.LittleEndian, unicode.UseBOM).NewEncoder().String("a\tb\r\n1\t2\r\n")
	f, err = Parse([]byte(u16), "u.csv")
	if err != nil {
		t.Fatal(err)
	}
	if f.Dialect.Encoding != "utf-16le" || f.Dialect.Comma != '\t' || f.Rows[1][0] != "1" {
		t.Errorf("utf-16: %+v %q", f.Dialect, f.Rows)
	}
	if got := encodeSame(t, f, nil); got != u16 {
		t.Errorf("utf-16 written back differently")
	}
	f, _ = Parse([]byte("\xef\xbb\xbfa,b\n1,2\n"), "b.csv")
	if !f.Dialect.BOM || f.Rows[0][0] != "a" {
		t.Errorf("BOM: %+v %q", f.Dialect, f.Rows)
	}
	if got := encodeSame(t, f, nil); got != "\xef\xbb\xbfa,b\n1,2\n" {
		t.Errorf("BOM written back as %q", got)
	}
}

func TestValues(t *testing.T) {
	d := DefaultDialect("x.csv")
	for field, want := range map[string]string{
		"1,200":      "number 1200",
		"00123":      "text 00123",
		"=SUM(A1)":   "text =SUM(A1)",
		"2026-03-01": "number 46082",
		"12%":        "number 0.12",
		"TRUE":       "bool",
		" hi ":       "text  hi ",
	} {
		v := d.Value(field)
		got := v.Kind.String()
		switch v.Kind {
		case grid.Number:
			got += " " + grid.FormatNumber(v.Num)
		case grid.Text:
			got += " " + v.Text
		}
		if got != want {
			t.Errorf("Value(%q) = %s, want %s", field, got, want)
		}
	}
}

func TestText(t *testing.T) {
	d := DefaultDialect("x.csv")
	for _, tt := range []struct {
		v    float64
		f    grid.Format
		want string
	}{
		{1200, grid.Format{Kind: grid.Integer}, "1,200"},
		{-1234.5, grid.Format{Kind: grid.Decimal, Decimals: 2}, "-1,234.50"},
		{5.5, grid.Format{Kind: grid.Currency, Decimals: 2, Symbol: "$"}, "$5.50"},
		{-5.5, grid.Format{Kind: grid.Currency, Decimals: 2, Symbol: "£"}, "-£5.50"},
		{7, grid.Format{Kind: grid.Currency, Symbol: "kr"}, "7 kr"},
		{0.125, grid.Format{Kind: grid.Percent, Decimals: 1}, "12.5%"},
		{46082, grid.Format{Kind: grid.Date}, "2026-03-01"},
		{0.5, grid.Format{}, "0.5"},
	} {
		got := Text(grid.Cell{Kind: grid.Number, Num: tt.v}, tt.f, d, false)
		if got != tt.want {
			t.Errorf("Text(%v, %+v) = %q, want %q", tt.v, tt.f, got, tt.want)
		}
		// What sheets writes, sheets reads back to the same value.
		if v := d.Value(got); v.Num != tt.v && !(tt.f.Kind == grid.Percent && strings.HasSuffix(got, "%")) {
			t.Errorf("%q reads back as %v, want %v", got, v.Num, tt.v)
		}
	}
	plain := Dialect{Plain: true}
	if got := Text(grid.Cell{Kind: grid.Number, Num: 1200}, grid.Format{Kind: grid.Currency, Symbol: "$"}, plain, false); got != "1200" {
		t.Errorf("plain currency = %q", got)
	}
}

func TestSheetName(t *testing.T) {
	for path, want := range map[string]string{
		"/x/sales.csv": "sales",
		"q1:2026.csv":  "q1_2026",
		"a-very-long-file-name-that-goes-on-and-on.csv": "a-very-long-file-name-that-goes",
		".csv": "Sheet1",
	} {
		if got := SheetName(path); got != want {
			t.Errorf("SheetName(%q) = %q, want %q", path, got, want)
		}
	}
}

package xlsxtest

import (
	"archive/zip"
	"bytes"
	"strings"
	"testing"
)

func book(t *testing.T, sheetData string) []byte {
	t.Helper()
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	w, err := zw.Create("xl/worksheets/sheet1.xml")
	if err != nil {
		t.Fatal(err)
	}
	w.Write([]byte(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>` + sheetData + `</sheetData></worksheet>`))
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	return buf.Bytes()
}

func TestCheckOrder(t *testing.T) {
	for data, want := range map[string]string{
		`<row r="1"><c r="A1"/><c r="C1"/></row><row r="3"><c r="B3"/></row>`: "",
		`<row><c/><c/></row><row><c r="B2"/><c/></row>`:                       "", // positions implied
		`<row r="3"/><row r="1"/>`:                                            "row 1 after row 3",
		`<row r="3"/><row r="3"/>`:                                            "row 3 after row 3",
		`<row r="8"><c r="B8"/><c r="B3"/></row>`:                             "cell B3 in row 8",
		`<row r="3"><c r="B3"/><c r="B3"/></row>`:                             "cell B3 after B3",
		`<row r="3"><c r="C3"/><c r="A3"/></row>`:                             "cell A3 after C3",
	} {
		err := CheckOrder(book(t, data))
		switch {
		case want == "" && err != nil:
			t.Errorf("%s: %v", data, err)
		case want != "" && (err == nil || !strings.Contains(err.Error(), want)):
			t.Errorf("%s: got %v, want %s", data, err, want)
		}
	}
}

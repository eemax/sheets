// Package xlsxtest holds checks for tests on saved workbooks, below what
// a reader of the file shows: readers forgive what Excel repairs.
package xlsxtest

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

	"github.com/eemax/sheets/internal/grid"
)

// CheckOrderFile is CheckOrder on the workbook at path.
func CheckOrderFile(name string) error {
	data, err := os.ReadFile(name)
	if err != nil {
		return err
	}
	return CheckOrder(data)
}

// CheckOrder checks that every worksheet of a workbook lists its rows in
// order, each once, and the cells of each row in order, each once and in
// its own row, as Excel requires: Excel offers to repair a file that does
// not, while openpyxl, LibreOffice and sheets read it as fine.
func CheckOrder(data []byte) error {
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return err
	}
	for _, zf := range zr.File {
		if path.Dir(zf.Name) != "xl/worksheets" || !strings.HasSuffix(zf.Name, ".xml") {
			continue
		}
		r, err := zf.Open()
		if err != nil {
			return err
		}
		err = checkSheet(r)
		r.Close()
		if err != nil {
			return fmt.Errorf("%s: %w", zf.Name, err)
		}
	}
	return nil
}

func checkSheet(r io.Reader) error {
	d := xml.NewDecoder(r)
	lastRow, row, lastCol := 0, 0, 0
	for {
		tok, err := d.Token()
		if err == io.EOF {
			return nil
		}
		if err != nil {
			return err
		}
		se, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		switch se.Name.Local {
		case "row":
			v := attr(se, "r")
			if v == "" {
				// A row without a number follows the one before.
				row = lastRow + 1
			} else if row, err = strconv.Atoi(v); err != nil {
				return fmt.Errorf("row r=%q", v)
			}
			if row <= lastRow {
				return fmt.Errorf("row %d after row %d", row, lastRow)
			}
			lastRow, lastCol = row, 0
		case "c":
			v := attr(se, "r")
			col := lastCol + 1
			if v != "" {
				cr, cc, ok := grid.ParseCell(v)
				if !ok {
					return fmt.Errorf("cell r=%q", v)
				}
				if cr != row {
					return fmt.Errorf("cell %s in row %d", v, row)
				}
				col = cc
			}
			if col <= lastCol {
				return fmt.Errorf("cell %s after %s in row %d", grid.CellName(row, col), grid.CellName(row, lastCol), row)
			}
			lastCol = col
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

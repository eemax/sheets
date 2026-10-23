// Package book picks the implementation for a file's format. Commands use
// it instead of a format package directly, so adding a format (csv, ods)
// means adding a case here.
package book

import (
	"fmt"
	"path/filepath"
	"strings"

	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/input"
	"github.com/eemax/sheets/internal/xlsx"
)

// Editor changes a workbook: cells, rows and columns, and sheets.
type Editor interface {
	Workbook() *grid.Workbook
	Created() bool
	EnsureSheet(name string) (string, bool, error)
	Set(sheet string, row, col int, v input.Value) error
	Clear(sheet string) error
	Evaluate(sheet string, row, col int) (v grid.Cell, note string, ok bool)

	InsertRows(sheet string, at, n int) error
	DeleteRows(sheet string, from, to int) error
	InsertCols(sheet string, at, n int) error
	DeleteCols(sheet string, from, to int) error

	AddSheet(name, before, after string) error
	RenameSheet(from, to string) error
	DeleteSheet(name string) error
	CopySheet(from, to string) (notes []string, err error)
	MoveSheet(name, before, after string, first, last bool) error
	SetVisible(name string, visible bool) error
	// References lists formulas and names elsewhere that refer to a sheet.
	References(sheet string) []string
	// SheetList is the current sheet order, hidden sheets marked.
	SheetList() []string

	AddChart(spec ChartSpec) error
	DeleteChart(sheet, cell string) error
	// ExtendTable grows the charts and Excel table of the table whose data
	// is body by n rows written right below it.
	ExtendTable(sheet string, body grid.Range, n int)

	// Save writes the file; with dryRun it only works out the outcome.
	Save(dryRun bool) error
	Close() error
	// Warnings lists problems noticed while editing that did not stop it.
	Warnings() []string
}

type (
	ChartSpec     = xlsx.ChartSpec
	SeriesSpec    = xlsx.SeriesSpec
	FormatOptions = xlsx.FormatOptions
	Override      = xlsx.Override
	FormatResult  = xlsx.FormatResult
)

func format(path string) (string, error) {
	switch strings.ToLower(filepath.Ext(path)) {
	case ".xlsx", ".xlsm", ".xltx", ".xltm":
		return "xlsx", nil
	}
	return "", fmt.Errorf("%s: unsupported file type %q (supported: .xlsx, .xlsm)", path, filepath.Ext(path))
}

// Open reads a workbook.
func Open(path string) (*grid.Workbook, error) {
	if _, err := format(path); err != nil {
		return nil, err
	}
	return xlsx.Read(path)
}

// OpenBytes reads a workbook held in memory; name gives its format.
func OpenBytes(name string, data []byte) (*grid.Workbook, error) {
	if _, err := format(name); err != nil {
		return nil, err
	}
	wb, err := xlsx.ReadBytes(data)
	if err != nil {
		return nil, err
	}
	wb.Path = name
	wb.Size = int64(len(data))
	return wb, nil
}

// Edit opens a workbook for changes, creating it when missing and create is
// set.
func Edit(path string, create bool) (Editor, error) {
	if _, err := format(path); err != nil {
		return nil, err
	}
	return xlsx.OpenEditor(path, create)
}

// Format applies the house style in place.
func Format(path string, opts FormatOptions) (*FormatResult, error) {
	if _, err := format(path); err != nil {
		return nil, err
	}
	return xlsx.Format(path, opts)
}

// Check reports whether Format would leave the file unchanged.
func Check(path string, opts FormatOptions) (bool, *FormatResult, error) {
	if _, err := format(path); err != nil {
		return false, nil, err
	}
	return xlsx.Check(path, opts)
}

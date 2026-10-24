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

	// Style lays a look over a range and records it, so fmt keeps it.
	Style(sheet string, r grid.Range, l *Look) (StyleResult, error)
	// Records lists the recorded styling; Unrecord removes entries.
	Records() []Record
	Unrecord(seqs []int) error
	Cond(sheet string, r grid.Range, rule []string) (string, error)
	CondFormats(sheet string) ([]CondFormat, error)
	ClearCond(sheet string, r *grid.Range) (int, error)
	Validate(sheet string, r grid.Range, rule ValidationRule) (string, error)
	Validations(sheet string) ([]Validation, error)
	ClearValidation(sheet string, r *grid.Range) (int, error)
	View(sheet string, o ViewOptions) error
	ViewOf(sheet string) (ViewState, error)
	// Looks describes cell styles in blocks, as style properties.
	Looks(sheet string, r grid.Range) ([]LookRun, error)

	// Range operations, on a workbook as opened.
	ClearRange(sheet string, r grid.Range, w ClearWhat) (RangeResult, error)
	CopyRange(from string, src grid.Range, to string, at grid.Range, o CopyOptions) (RangeResult, error)
	MoveRange(from string, src grid.Range, to string, at grid.Range) (RangeResult, error)
	SortRange(sheet string, r grid.Range, keys []SortKey) (SortResult, error)
	// Replace runs fn over cell values (and formulas, when set); sheet ""
	// means every sheet.
	Replace(sheet string, r *grid.Range, fn func(string) (string, bool), formulas bool) ([]Replacement, error)
	// SetNote puts a note on a cell; empty text removes it.
	SetNote(sheet string, row, col int, text, author string) error
	// SetLink puts a hyperlink on a cell; nil removes it.
	SetLink(sheet string, row, col int, l *grid.Link, text string) error

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

	Look           = xlsx.Look
	StyleResult    = xlsx.StyleResult
	Record         = xlsx.Record
	CondFormat     = xlsx.CondFormat
	ValidationRule = xlsx.ValidationRule
	Validation     = xlsx.Validation
	ViewOptions    = xlsx.ViewOptions
	ViewState      = xlsx.ViewState
	LookRun        = xlsx.LookRun

	ClearWhat   = xlsx.ClearWhat
	CopyOptions = xlsx.CopyOptions
	RangeResult = xlsx.RangeResult
	SortKey     = xlsx.SortKey
	SortResult  = xlsx.SortResult
	Replacement = xlsx.Replacement
)

// ParseLook reads style properties (see xlsx.ParseLook).
func ParseLook(ws []string) (Look, error) { return xlsx.ParseLook(ws) }

// ParseColor reads a colour: #RRGGBB or a name.
func ParseColor(s string) (string, error) { return xlsx.ParseColor(s) }

// ColorNames lists the colour names ParseColor knows.
func ColorNames() []string { return xlsx.ColorNames() }

// FormatOf names the format of a file by its extension: xlsx or csv.
func FormatOf(path string) (string, error) {
	switch strings.ToLower(filepath.Ext(path)) {
	case ".xlsx", ".xlsm", ".xltx", ".xltm":
		return "xlsx", nil
	case ".csv", ".tsv":
		return "csv", nil
	}
	return "", fmt.Errorf("%s: unsupported file type %q (supported: .xlsx, .xlsm, .csv, .tsv)", path, filepath.Ext(path))
}

func format(path string) (string, error) { return FormatOf(path) }

// Open reads a workbook.
func Open(path string) (*grid.Workbook, error) {
	kind, err := format(path)
	if err != nil {
		return nil, err
	}
	if kind == "csv" {
		data, err := xlsx.ReadFile(path)
		if err != nil {
			return nil, err
		}
		return openCSV(path, data)
	}
	return xlsx.Read(path)
}

// OpenBytes reads a workbook held in memory; name gives its format.
func OpenBytes(name string, data []byte) (*grid.Workbook, error) {
	kind, err := format(name)
	if err != nil {
		return nil, err
	}
	if kind == "csv" {
		return openCSV(name, data)
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
	kind, err := format(path)
	if err != nil {
		return nil, err
	}
	if kind == "csv" {
		return editCSV(path, create)
	}
	// A workbook streams its edits (see xlsx.Auto).
	if from := xlsx.StreamFrom(); from >= 0 {
		if data, err := xlsx.ReadFile(path); err == nil && int64(len(data)) >= from {
			return xlsx.OpenAuto(path, data)
		}
	}
	return xlsx.OpenEditor(path, create)
}

// Format applies the house style in place.
func Format(path string, opts FormatOptions) (*FormatResult, error) {
	kind, err := format(path)
	if err != nil {
		return nil, err
	}
	if kind == "csv" {
		return nil, errCSV(path, "formatting")
	}
	return xlsx.Format(path, opts)
}

// Check reports whether Format would leave the file unchanged.
func Check(path string, opts FormatOptions) (bool, *FormatResult, error) {
	kind, err := format(path)
	if err != nil {
		return false, nil, err
	}
	if kind == "csv" {
		return false, nil, errCSV(path, "formatting")
	}
	return xlsx.Check(path, opts)
}

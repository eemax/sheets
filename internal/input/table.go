package input

import (
	"bytes"
	"encoding/csv"
	"encoding/json"
	"fmt"
	"io"
	"strings"

	"github.com/eemax/sheets/internal/grid"
)

// Table reads rows of values from r. format is "tsv", "csv", "json" or
// "auto". JSON may be an array of arrays or an array of objects; objects
// produce a header row from their keys in first-seen order.
func Table(r io.Reader, format string, raw bool) ([][]Value, error) {
	data, err := io.ReadAll(r)
	if err != nil {
		return nil, err
	}
	data = bytes.TrimPrefix(data, []byte("\xef\xbb\xbf")) // Excel's CSV export starts with a BOM
	if format == "" || format == "auto" {
		format = sniff(data)
	}
	switch format {
	case "json":
		return jsonTable(data, raw)
	case "csv":
		return delimited(data, ',', raw)
	case "tsv":
		return tsv(data, raw), nil
	}
	return nil, fmt.Errorf("unknown input format %q (use tsv, csv or json)", format)
}

func sniff(data []byte) string {
	t := bytes.TrimSpace(data)
	if len(t) > 0 && (t[0] == '[' || t[0] == '{') {
		return "json"
	}
	first, _, _ := bytes.Cut(t, []byte("\n"))
	if !bytes.Contains(first, []byte("\t")) && bytes.Contains(first, []byte(",")) {
		return "csv"
	}
	return "tsv"
}

// tsv splits on tabs and newlines. Backslashes are kept as they are (paths
// like C:\new must survive); values containing tabs or newlines need CSV or
// JSON input.
func tsv(data []byte, raw bool) [][]Value {
	text := strings.TrimRight(strings.ReplaceAll(string(data), "\r\n", "\n"), "\n")
	if text == "" {
		return nil
	}
	var rows [][]Value
	for _, line := range strings.Split(text, "\n") {
		fields := strings.Split(line, "\t")
		row := make([]Value, len(fields))
		for i, f := range fields {
			row[i] = Parse(f, raw)
		}
		rows = append(rows, row)
	}
	return rows
}

func delimited(data []byte, sep rune, raw bool) ([][]Value, error) {
	cr := csv.NewReader(bytes.NewReader(data))
	cr.Comma = sep
	cr.FieldsPerRecord = -1
	cr.LazyQuotes = true
	recs, err := cr.ReadAll()
	if err != nil {
		return nil, fmt.Errorf("parsing csv: %w", err)
	}
	rows := make([][]Value, len(recs))
	for i, rec := range recs {
		rows[i] = make([]Value, len(rec))
		for j, f := range rec {
			rows[i][j] = Parse(f, raw)
		}
	}
	return rows, nil
}

func jsonTable(data []byte, raw bool) ([][]Value, error) {
	dec := json.NewDecoder(bytes.NewReader(data))
	dec.UseNumber()
	var top any
	if err := dec.Decode(&top); err != nil {
		return nil, fmt.Errorf("parsing json: %w", err)
	}
	if obj, ok := top.(map[string]any); ok {
		top = []any{obj}
	}
	items, ok := top.([]any)
	if !ok {
		return nil, fmt.Errorf("json input must be an array of arrays or of objects")
	}
	if len(items) == 0 {
		return nil, nil
	}
	if _, isObj := items[0].(map[string]any); isObj {
		return jsonObjects(data, items, raw)
	}
	rows := make([][]Value, len(items))
	for i, it := range items {
		arr, ok := it.([]any)
		if !ok {
			arr = []any{it}
		}
		rows[i] = make([]Value, len(arr))
		for j, v := range arr {
			rows[i][j] = FromJSON(v, raw)
		}
	}
	return rows, nil
}

// jsonObjects keeps keys in the order they first appear in the document,
// which encoding/json maps do not preserve.
func jsonObjects(data []byte, items []any, raw bool) ([][]Value, error) {
	keys, err := objectKeyOrder(data)
	if err != nil {
		return nil, err
	}
	header := make([]Value, len(keys))
	for i, k := range keys {
		header[i] = Value{Kind: grid.Text, Text: k}
	}
	rows := [][]Value{header}
	for _, it := range items {
		obj, ok := it.(map[string]any)
		if !ok {
			return nil, fmt.Errorf("json array mixes objects and other values")
		}
		row := make([]Value, len(keys))
		for i, k := range keys {
			if v, ok := obj[k]; ok {
				row[i] = FromJSON(v, raw)
			}
		}
		rows = append(rows, row)
	}
	return rows, nil
}

// objectKeyOrder lists the keys of the objects in a top-level array (or a
// single object) in first-seen order.
func objectKeyOrder(data []byte) ([]string, error) {
	var elems []json.RawMessage
	if t := bytes.TrimSpace(data); len(t) > 0 && t[0] == '{' {
		elems = []json.RawMessage{t}
	} else if err := json.Unmarshal(t, &elems); err != nil {
		return nil, err
	}
	seen := map[string]bool{}
	var keys []string
	for _, e := range elems {
		dec := json.NewDecoder(bytes.NewReader(e))
		if tok, err := dec.Token(); err != nil || tok != json.Delim('{') {
			continue
		}
		for dec.More() {
			tok, err := dec.Token()
			if err != nil {
				return nil, err
			}
			k, _ := tok.(string)
			if !seen[k] {
				seen[k] = true
				keys = append(keys, k)
			}
			var skip json.RawMessage
			if err := dec.Decode(&skip); err != nil {
				return nil, err
			}
		}
	}
	return keys, nil
}

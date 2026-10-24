package xlsx

import (
	"archive/zip"
	"bytes"
	"io"
	"strconv"

	"github.com/xuri/excelize/v2"
)

// styleBook makes cell formats for a streaming edit with the same code the
// full editor uses (applyNumFmt, styler.restyle), in a workbook holding
// nothing but the file's styles and theme: excelize does not need the
// sheets for that.
type styleBook struct {
	f       *excelize.File
	st      *styler
	memo    map[string]int
	changed bool
}

func newStyleBook(styles, theme []byte) (*styleBook, error) {
	blank := excelize.NewFile()
	buf, err := blank.WriteToBuffer()
	blank.Close()
	if err != nil {
		return nil, err
	}
	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		return nil, err
	}
	var out bytes.Buffer
	zw := zip.NewWriter(&out)
	for _, zf := range zr.File {
		r, err := zf.Open()
		if err != nil {
			return nil, err
		}
		b, err := io.ReadAll(r)
		r.Close()
		if err != nil {
			return nil, err
		}
		switch {
		case zf.Name == "xl/styles.xml":
			b = styles
		case zf.Name == "xl/theme/theme1.xml" && theme != nil:
			b = theme
		}
		w, err := zw.Create(zf.Name)
		if err != nil {
			return nil, err
		}
		if _, err := w.Write(b); err != nil {
			return nil, err
		}
	}
	if err := zw.Close(); err != nil {
		return nil, err
	}
	f, err := excelize.OpenReader(bytes.NewReader(out.Bytes()))
	if err != nil {
		return nil, err
	}
	return &styleBook{f: f, st: newStyler(f), memo: map[string]int{}}, nil
}

// numFmt is cell format cur with number format code, as applyNumFmt makes
// it.
func (sb *styleBook) numFmt(cur int, code string) (int, error) {
	key := strconv.Itoa(cur) + "|" + code
	if id, ok := sb.memo[key]; ok {
		return id, nil
	}
	s, err := sb.f.GetStyle(cur)
	if err != nil {
		return 0, err
	}
	s.NumFmt, s.DecimalPlaces, s.CustomNumFmt = 0, nil, &code
	normalizeFill(s)
	id, err := sb.f.NewStyle(s)
	if err != nil {
		return 0, err
	}
	sb.memo[key], sb.changed = id, true
	return id, nil
}

// look is cell format cur with a look laid over it, at the given edges of
// the range it styles.
func (sb *styleBook) look(cur int, l *Look, edges uint8) (int, error) {
	id, err := sb.st.restyle(cur, l, edges)
	if err == nil && id != cur {
		sb.changed = true
	}
	return id, err
}

// styles is the styles part with the formats made.
func (sb *styleBook) styles() ([]byte, error) {
	buf, err := sb.f.WriteToBuffer()
	if err != nil {
		return nil, err
	}
	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		return nil, err
	}
	for _, zf := range zr.File {
		if zf.Name == "xl/styles.xml" {
			r, err := zf.Open()
			if err != nil {
				return nil, err
			}
			defer r.Close()
			return io.ReadAll(r)
		}
	}
	return nil, io.ErrUnexpectedEOF
}

func (sb *styleBook) close() { sb.f.Close() }

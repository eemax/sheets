package xlsx

import (
	"path/filepath"
	"testing"

	"github.com/eemax/sheets/internal/input"
)

// A legacy array over several cells sends to the full editor only the
// edits that reach it: one elsewhere streams, as on a large workbook.
func TestLegacyArrayHandOver(t *testing.T) {
	t.Setenv("SHEETS_STREAM", "always")
	data := checkData(t, "a.xlsx", openFixture(t, "arrays_openpyxl.xlsx")).Data
	path := filepath.Join(t.TempDir(), "a.xlsx")
	for _, c := range []struct {
		row, col int
		streams  bool
	}{{9, 26, true}, {2, 1, false}} {
		a, err := OpenAuto(path, data)
		must(t, err)
		must(t, a.Set("S", c.row, c.col, input.Parse("0", false)))
		must(t, a.Save(false))
		if a.Streaming() != c.streams {
			t.Errorf("a write to row %d, column %d: streaming %v", c.row, c.col, a.Streaming())
		}
		if !c.streams {
			if got := stored(t, path, dataPart, "D2"); got != "0" {
				t.Errorf("D2 stores %q after A2 = 0", got)
			}
		}
	}
}

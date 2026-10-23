package grid

import (
	"testing"
	"unsafe"
)

func TestCellSize(t *testing.T) {
	if n := unsafe.Sizeof(Cell{}); n > 48 {
		t.Errorf("Cell is %d bytes; keep it at 48 or less", n)
	}
}

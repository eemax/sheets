package xlsx

import (
	"errors"
	"io/fs"
	"os"
	"path/filepath"
	"sync"
)

// A Batch holds one workbook in memory in place of the file on disk: the
// commands of a batch (sheets do) read and write it there, one after the
// other, and the batch writes the file once at the end, or not at all.
type Batch struct {
	path    string
	data    []byte // nil while the file does not exist
	changed bool
}

var (
	batchMu sync.Mutex
	batch   *Batch
)

// StartBatch loads path (if it exists) and routes every read and write of
// it to memory until End.
func StartBatch(path string) (*Batch, error) {
	abs, err := filepath.Abs(path)
	if err != nil {
		return nil, err
	}
	b := &Batch{path: abs}
	data, err := os.ReadFile(abs)
	switch {
	case err == nil:
		b.data = data
	case !errors.Is(err, fs.ErrNotExist):
		return nil, err
	}
	batchMu.Lock()
	defer batchMu.Unlock()
	if batch != nil {
		return nil, errors.New("a batch is already running")
	}
	batch = b
	return b, nil
}

// End stops routing the file to memory.
func (b *Batch) End() {
	batchMu.Lock()
	defer batchMu.Unlock()
	if batch == b {
		batch = nil
	}
}

// Changed reports whether a command of the batch saved the file.
func (b *Batch) Changed() bool { return b.changed }

// Commit ends the batch and writes the file, once: one undo step.
func (b *Batch) Commit() error {
	b.End()
	if !b.changed {
		return nil
	}
	return writeAtomic(b.path, b.data)
}

// held returns the running batch when it holds path.
func held(path string) *Batch {
	batchMu.Lock()
	defer batchMu.Unlock()
	if batch == nil {
		return nil
	}
	if abs, err := filepath.Abs(path); err == nil && abs == batch.path {
		return batch
	}
	return nil
}

// readFile reads a workbook: from a running batch when it holds the file.
func readFile(path string) ([]byte, error) {
	if b := held(path); b != nil {
		if b.data == nil {
			return nil, &fs.PathError{Op: "open", Path: path, Err: fs.ErrNotExist}
		}
		return b.data, nil
	}
	return os.ReadFile(path)
}

package main

import (
	"os"
	"testing"

	"github.com/rogpeppe/go-internal/testscript"
)

func TestMain(m *testing.M) {
	testscript.Main(m, map[string]func(){"sheets": main})
}

// TestScripts runs the end-to-end scripts in testdata/*.txtar. Run with
// -update to rewrite expected output after an intended change.
func TestScripts(t *testing.T) {
	testscript.Run(t, testscript.Params{
		Dir:           "testdata",
		UpdateScripts: *update,
		Setup: func(env *testscript.Env) error {
			env.Setenv("SHEETS_CACHE_DIR", env.WorkDir+"/.cache")
			// SHEETS_STREAM=always runs every script through the streaming
			// editor, as large files are edited.
			if v := os.Getenv("SHEETS_STREAM"); v != "" {
				env.Setenv("SHEETS_STREAM", v)
			}
			return nil
		},
	})
}

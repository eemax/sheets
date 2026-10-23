package main

import (
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
			return nil
		},
	})
}

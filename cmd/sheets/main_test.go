package main

import (
	"os"
	"path/filepath"
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
		Cmds: map[string]func(ts *testscript.TestScript, neg bool, args []string){
			// fixture copies workbooks from internal/xlsx/testdata into the
			// work directory.
			"fixture": func(ts *testscript.TestScript, neg bool, args []string) {
				for _, name := range args {
					data, err := os.ReadFile(filepath.Join("..", "..", "internal", "xlsx", "testdata", name))
					ts.Check(err)
					ts.Check(os.WriteFile(ts.MkAbs(name), data, 0o644))
				}
			},
		},
	})
}

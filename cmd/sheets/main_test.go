package main

import (
	"archive/zip"
	"io"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"testing"

	"github.com/eemax/sheets/internal/xlsxtest"
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
			// xmlorder checks that the workbooks' sheets list rows and
			// cells in order, each once, as Excel requires.
			"xmlorder": func(ts *testscript.TestScript, neg bool, args []string) {
				for _, name := range args {
					err := xlsxtest.CheckOrderFile(ts.MkAbs(name))
					if neg && err == nil {
						ts.Fatalf("%s: rows and cells in order", name)
					}
					if !neg && err != nil {
						ts.Fatalf("%s: %v", name, err)
					}
				}
			},
			// shortform checks the last command's change report: one, and
			// under the given number of lines (the head line and what
			// follows it, up to the next-step hints).
			// shortform 15
			"shortform": func(ts *testscript.TestScript, neg bool, args []string) {
				if len(args) != 1 {
					ts.Fatalf("usage: shortform maxlines")
				}
				limit, err := strconv.Atoi(args[0])
				ts.Check(err)
				lines := strings.Split(strings.TrimRight(ts.ReadFile("stdout"), "\n"), "\n")
				start, n := -1, 0
				for i, l := range lines {
					if strings.HasPrefix(l, "change: ") {
						if start >= 0 {
							ts.Fatalf("two change reports: lines %d and %d", start+1, i+1)
						}
						start = i
					}
				}
				if start < 0 {
					if !neg {
						ts.Fatalf("no change report")
					}
					return
				}
				for _, l := range lines[start:] {
					if strings.HasPrefix(l, "next: ") {
						break
					}
					n++
				}
				if neg {
					ts.Fatalf("a change report of %d lines", n)
				}
				if n >= limit {
					ts.Fatalf("the change report has %d lines, %d or more", n, limit)
				}
			},
			// part prints a part of a workbook:
			// part book.xlsx xl/worksheets/sheet1.xml
			"part": func(ts *testscript.TestScript, neg bool, args []string) {
				if len(args) != 2 {
					ts.Fatalf("usage: part file name")
				}
				zr, err := zip.OpenReader(ts.MkAbs(args[0]))
				ts.Check(err)
				defer zr.Close()
				r, err := zr.Open(args[1])
				ts.Check(err)
				defer r.Close()
				_, err = io.Copy(ts.Stdout(), r)
				ts.Check(err)
			},
		},
	})
}

package cli

import (
	"bytes"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"

	"github.com/eemax/sheets/internal/words"
	"github.com/eemax/sheets/internal/xlsx"
	"github.com/spf13/cobra"
)

// step is one command of a batch script.
type step struct {
	line  int    // where it starts in the script
	text  string // its first line, as written
	args  []string
	input []byte // here-document, for write
	from  string // < file, for write
}

type doResult struct {
	File   string    `json:"file"`
	Saved  bool      `json:"saved"`
	DryRun bool      `json:"dry_run,omitempty"`
	Steps  []doEntry `json:"steps"`
}

type doEntry struct {
	Line    int             `json:"line"`
	Command string          `json:"command"`
	Output  json.RawMessage `json:"output,omitempty"`
	// Failed is set for a check that found problems (lint, fmt --check):
	// the batch goes on.
	Failed bool `json:"failed,omitempty"`
}

// notInBatch are the commands a batch does not run: they work on other
// files, or on the file's history.
var notInBatch = map[string]string{
	"do": "a batch cannot run another", "undo": "run it after the batch",
	"diff": "it compares two files", "render": "run it after the batch",
	"agent-docs": "it is not about the file", "skill": "it is not about the file",
	"help": "it is not about the file", "completion": "it is not about the file",
}

func (a *app) doCmd() *cobra.Command {
	var script string
	cmd := &cobra.Command{
		Use:   "do <file> [-f script | command...]",
		Short: "Run many commands on one file: one save, one undo step",
		Long: `Run a script of sheets commands against one workbook, one command per line,
written as on the command line without 'sheets' and the file name. Every
command sees the changes of the ones before it; the file is saved once at
the end, as one undo step, and only when every command succeeded: the first
error stops the batch and nothing is saved.

  sheets do report.xlsx <<'EOF'
  write Data!A1 -i sales.csv
  set Data!F1 Margin
  write Data!F2 <<END
  =D{r}/C{r}
  END
  fmt
  style 'Data[#Headers]' fill=lightblue color=black
  cond 'Data[Margin]' '< 0' color=red
  view Data --freeze B2
  EOF

The script comes from stdin, from a file with -f, or from the arguments
after the file, one command each (sheets do book.xlsx 'set A1 5' fmt).
Blank lines and lines starting with # are skipped; a line ending in \
continues on the next. A check that finds problems (lint, fmt --check)
prints them and the batch goes on.
Quote as in a shell. Data for write comes from -i file, < file, or a
here-document (<<END ... END) inside the script. Reading commands (read,
peek, info, find, lint, query) show the file as the batch has left it so
far. --dry-run runs everything and saves nothing.`,
		Args: cobra.MinimumNArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			file := args[0]
			var src []byte
			var err error
			if len(args) > 1 {
				if script != "" {
					return errors.New("give the commands as arguments or with -f, not both")
				}
				src = []byte(strings.Join(args[1:], "\n"))
			} else if script != "" {
				src, err = os.ReadFile(script)
			} else if f, ok := a.stdin.(*os.File); ok && isTerminal(f) {
				return errors.New("no script: pipe commands on stdin or pass -f <file>")
			} else {
				src, err = io.ReadAll(a.stdin)
			}
			if err != nil {
				return err
			}
			steps, err := parseScript(string(src))
			if err != nil {
				return err
			}
			if len(steps) == 0 {
				return errors.New("the script has no commands")
			}
			b, err := xlsx.StartBatch(file)
			if err != nil {
				return err
			}
			defer b.End()
			res := doResult{File: file, DryRun: a.dryRun}
			var failed []string
			for _, st := range steps {
				out, err := a.runStep(file, st)
				entry := doEntry{Line: st.line, Command: st.text}
				if err == errSilent {
					entry.Failed, err = true, nil
					failed = append(failed, fmt.Sprintf("line %d (%s)", st.line, st.text))
				}
				if a.json {
					if t := bytes.TrimSpace(out); len(t) > 0 {
						entry.Output = jsonOutput(t)
					}
				} else {
					a.printf("> %s\n", st.text)
					a.stdout.Write(out)
				}
				res.Steps = append(res.Steps, entry)
				if err != nil {
					return fmt.Errorf("line %d (%s): %w; nothing was saved", st.line, st.text, err)
				}
			}
			if !a.dryRun {
				if err := b.Commit(); err != nil {
					return err
				}
				res.Saved = b.Changed()
			}
			if a.json {
				return a.printJSON(res)
			}
			found := ""
			if len(failed) > 0 {
				found = "; problems found by " + strings.Join(failed, ", ")
			}
			switch {
			case a.dryRun:
				a.printf("dry run: %s, nothing was saved%s\n", plural(len(steps), "command"), found)
			case res.Saved:
				a.printf("saved %s: %s, one undo step%s\n", file, plural(len(steps), "command"), found)
			default:
				a.printf("%s: %s, no changes%s\n", file, plural(len(steps), "command"), found)
			}
			return nil
		},
	}
	cmd.Flags().StringVarP(&script, "file", "f", "", "read the script from this file instead of stdin")
	return cmd
}

// jsonOutput keeps one JSON document as is, and wraps several in an array.
func jsonOutput(b []byte) json.RawMessage {
	if json.Valid(b) {
		return b
	}
	var all []json.RawMessage
	dec := json.NewDecoder(bytes.NewReader(b))
	for {
		var v json.RawMessage
		if err := dec.Decode(&v); err != nil {
			break
		}
		all = append(all, v)
	}
	if out, err := json.Marshal(all); err == nil && len(all) > 0 {
		return out
	}
	out, _ := json.Marshal(string(b))
	return out
}

// runStep runs one command of a batch against the file and returns what it
// printed.
func (a *app) runStep(file string, st step) ([]byte, error) {
	var out bytes.Buffer
	in := &app{stdin: bytes.NewReader(st.input), stdout: &out, stderr: &out, json: a.json, batch: true}
	if st.from != "" {
		data, err := os.ReadFile(st.from)
		if err != nil {
			return nil, err
		}
		in.stdin = bytes.NewReader(data)
	}
	root := in.root()
	in.json = a.json // root() bound the flag, at its default
	cmd, rest, err := root.Find(st.args)
	if err != nil || cmd == root {
		return nil, fmt.Errorf("unknown command %q", st.args[0])
	}
	top := cmd
	for top.Parent() != root {
		top = top.Parent()
	}
	if why, ok := notInBatch[top.Name()]; ok {
		return nil, fmt.Errorf("%s does not run in a batch: %s", top.Name(), why)
	}
	if cmd.HasSubCommands() {
		return nil, fmt.Errorf("%s needs one of: %s", cmd.Name(), subNames(cmd))
	}
	names := st.args[:len(st.args)-len(rest)]
	args := append(append([]string{}, names...), file)
	if len(rest) > 0 && isWorkbookPath(rest[0]) {
		if !sameFile(rest[0], file) {
			return nil, fmt.Errorf("a batch works on %s only; this line names %s", file, rest[0])
		}
		rest = rest[1:] // the file, written out of habit
	}
	root.SetArgs(negatives(root, append(args, rest...)))
	root.SetOut(&out)
	root.SetErr(&out)
	err = root.Execute()
	return out.Bytes(), err
}

func subNames(c *cobra.Command) string {
	var names []string
	for _, s := range c.Commands() {
		names = append(names, s.Name())
	}
	return strings.Join(names, ", ")
}

func isWorkbookPath(s string) bool {
	switch strings.ToLower(filepath.Ext(s)) {
	case ".xlsx", ".xlsm", ".xltx", ".xltm":
		return true
	}
	return false
}

func sameFile(a, b string) bool {
	if filepath.Clean(a) == filepath.Clean(b) {
		return true
	}
	x, err1 := filepath.Abs(a)
	y, err2 := filepath.Abs(b)
	return err1 == nil && err2 == nil && x == y || filepath.Base(a) == filepath.Base(b) && !strings.ContainsRune(a, filepath.Separator)
}

// parseScript splits a batch script into commands.
func parseScript(src string) ([]step, error) {
	lines := strings.Split(strings.ReplaceAll(src, "\r\n", "\n"), "\n")
	var steps []step
	for i := 0; i < len(lines); i++ {
		start := i
		text := strings.TrimSpace(lines[i])
		if text == "" || strings.HasPrefix(text, "#") {
			continue
		}
		line := lines[i]
		for strings.HasSuffix(line, `\`) && !strings.HasSuffix(line, `\\`) && i+1 < len(lines) {
			i++
			line = strings.TrimSuffix(line, `\`) + "\n" + lines[i]
		}
		ws, err := words.Split(line)
		if err != nil {
			return nil, fmt.Errorf("line %d: %w", start+1, err)
		}
		if len(ws) > 0 && ws[0].Text == "sheets" {
			ws = ws[1:]
		}
		if len(ws) == 0 {
			continue
		}
		st := step{line: start + 1, text: strings.TrimSpace(strings.SplitN(line, "\n", 2)[0])}
		// A here-document (<<END, <<'END', << END) or < file at the end.
		n := len(ws)
		switch last := ws[n-1]; {
		case last.Raw == "<<":
			err = errors.New("<< needs an end marker, like <<EOF")
		case strings.HasPrefix(last.Raw, "<<") && len(last.Raw) > 2:
			ws, st.input, i, err = hereDoc(ws[:n-1], strings.TrimPrefix(last.Text, "<<"), lines, i)
		case last.Raw != "<<" && n >= 2 && ws[n-2].Raw == "<<":
			ws, st.input, i, err = hereDoc(ws[:n-2], last.Text, lines, i)
		case n >= 2 && ws[n-2].Raw == "<":
			st.from, ws = last.Text, ws[:n-2]
		case strings.HasPrefix(last.Raw, "<") && !strings.HasPrefix(last.Raw, "<<") && len(last.Raw) > 1:
			st.from, ws = strings.TrimPrefix(last.Text, "<"), ws[:n-1]
		}
		if err != nil {
			return nil, fmt.Errorf("line %d: %w", start+1, err)
		}
		for _, w := range ws {
			st.args = append(st.args, w.Text)
		}
		if len(st.args) == 0 {
			return nil, fmt.Errorf("line %d: no command", start+1)
		}
		steps = append(steps, st)
	}
	return steps, nil
}

// hereDoc takes the lines after line i up to the one that is just marker
// (<<-marker: leading tabs removed).
func hereDoc(ws []words.Word, marker string, lines []string, i int) ([]words.Word, []byte, int, error) {
	strip := strings.HasPrefix(marker, "-")
	marker = strings.TrimPrefix(marker, "-")
	if marker == "" {
		return nil, nil, i, errors.New("<< needs an end marker, like <<EOF")
	}
	var body []string
	for j := i + 1; j < len(lines); j++ {
		l := lines[j]
		if strip {
			l = strings.TrimLeft(l, "\t")
		}
		if strings.TrimRight(l, " \t") == marker {
			text := strings.Join(body, "\n")
			if len(body) > 0 {
				text += "\n"
			}
			return ws, []byte(text), j, nil
		}
		body = append(body, l)
	}
	return nil, nil, i, fmt.Errorf("here-document has no closing %s line", marker)
}

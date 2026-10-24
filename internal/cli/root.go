// Package cli implements the sheets commands.
package cli

import (
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"os"
	"regexp"
	"runtime/debug"
	"strings"

	"github.com/eemax/sheets/internal/book"
	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/undo"
	"github.com/eemax/sheets/internal/words"
	"github.com/eemax/sheets/internal/xlsx"
	"github.com/spf13/cobra"
	"github.com/spf13/pflag"
)

// Version is set at build time with -ldflags "-X ...cli.Version=v1.2.3".
var Version = ""

// errSilent ends a command with exit status 1 after it has printed its own
// message (fmt --check).
var errSilent = errors.New("")

type app struct {
	stdin  io.Reader
	stdout io.Writer
	stderr io.Writer
	json   bool
	tty    bool
	dryRun bool
	// batch is set for the commands of a batch (sheets do): they print no
	// next-step hints.
	batch bool
}

// Main runs the CLI and returns the process exit code.
func Main(args []string, stdin io.Reader, stdout, stderr io.Writer) int {
	a := &app{stdin: stdin, stdout: stdout, stderr: stderr, tty: isTerminal(stdout)}
	// Keep the previous version of every file sheets overwrites.
	command := "sheets " + strings.Join(args, " ")
	xlsx.Journal = func(path string, before, after []byte) {
		if err := undo.Record(path, before, after, command); err != nil {
			fmt.Fprintf(stderr, "warning: could not record undo history: %v\n", err)
		}
	}
	root := a.root()
	root.SetArgs(negatives(root, args))
	root.SetIn(stdin)
	root.SetOut(stdout)
	root.SetErr(stderr)
	if err := root.Execute(); err != nil {
		if err != errSilent {
			if a.json {
				enc := json.NewEncoder(stderr)
				enc.Encode(map[string]string{"error": err.Error()})
			} else {
				fmt.Fprintln(stderr, "error:", err)
			}
		}
		return 1
	}
	return 0
}

func isTerminal(w io.Writer) bool {
	f, ok := w.(*os.File)
	if !ok {
		return false
	}
	st, err := f.Stat()
	return err == nil && st.Mode()&os.ModeCharDevice != 0
}

func (a *app) root() *cobra.Command {
	root := &cobra.Command{
		Use:   "sheets",
		Short: "Read, write and format spreadsheets from the command line, built for AI agents",
		Long: `sheets reads, writes and formats Excel workbooks (.xlsx) and CSV files, with
output sized for an agent's context window: start with an overview and
drill down.

  sheets info book.xlsx            what is in the file
  sheets peek book.xlsx Sales      column types and stats for a sheet
  sheets read book.xlsx Sales!A1:F20
  sheets write book.xlsx Sales!A1 < data.tsv
  sheets fmt book.xlsx             apply the house style

Run 'sheets agent-docs' for a guide written for AI agents (the instructions
of the agent skill and its topics; 'sheets agent-docs all' prints all of
it), and 'sheets skill --install' to install the skill.`,
		SilenceUsage:  true,
		SilenceErrors: true,
		Version:       version(),
	}
	root.PersistentFlags().BoolVar(&a.json, "json", false, "print JSON instead of text")
	root.PersistentFlags().BoolVar(&a.dryRun, "dry-run", false, "show what a change would do without saving it")
	root.AddCommand(a.infoCmd(), a.peekCmd(), a.readCmd(), a.findCmd(), a.queryCmd(), a.lintCmd(), a.diffCmd(), a.chartCmd(), a.renderCmd(), a.writeCmd(), a.setCmd(), a.fmtCmd(), a.styleCmd(), a.condCmd(), a.validateCmd(), a.viewCmd(),
		a.clearCmd(), a.copyCmd(), a.moveCmd(), a.sortCmd(), a.replaceCmd(), a.noteCmd(), a.linkCmd(), a.convertCmd(), a.doCmd(),
		a.linesCmd(false), a.linesCmd(true), a.sheetCmd(), a.undoCmd(), a.docsCmd(), a.skillCmd())
	root.CompletionOptions.HiddenDefaultCmd = true
	return root
}

func version() string {
	if Version != "" {
		return Version
	}
	if bi, ok := debug.ReadBuildInfo(); ok && bi.Main.Version != "" && bi.Main.Version != "(devel)" {
		return bi.Main.Version
	}
	return "dev"
}

// printJSON prints compact JSON (indented on a terminal): agents pay for
// every space.
func (a *app) printJSON(v any) error {
	enc := json.NewEncoder(a.stdout)
	enc.SetEscapeHTML(false)
	if a.tty {
		enc.SetIndent("", "  ")
	}
	return enc.Encode(v)
}

func (a *app) printf(format string, args ...any) {
	fmt.Fprintf(a.stdout, format, args...)
}

// hint prints a suggested next command. Hints go to stdout in text mode,
// where an agent reads them, and are carried in the JSON otherwise.
func (a *app) hint(format string, args ...any) {
	if a.batch {
		return
	}
	fmt.Fprintf(a.stdout, "next: "+format+"\n", args...)
}

// shellQuote quotes an argument for a suggested command line.
func shellQuote(s string) string { return words.Quote(s) }

// xlsxOnly refuses a CSV file for what needs a workbook, saying why.
func xlsxOnly(path, why string) error {
	if kind, err := book.FormatOf(path); err == nil && kind == "csv" {
		return fmt.Errorf("%s is a CSV file: %s", path, why)
	}
	return nil
}

// open reads a workbook.
func open(path string) (*grid.Workbook, error) {
	return book.Open(path)
}

// target resolves a user reference to a sheet and a range clipped to the
// used area. An empty ref means the first visible sheet.
func target(wb *grid.Workbook, ref string) (*grid.Sheet, grid.Range, bool, error) {
	r, err := grid.ParseRef(ref, func(name string) bool { return wb.Sheet(name) != nil })
	if err != nil {
		return nil, grid.Range{}, false, err
	}
	var s *grid.Sheet
	if r.Sheet == "" {
		s = wb.FirstVisible()
		if s == nil {
			return nil, grid.Range{}, false, errors.New("workbook has no sheets")
		}
	} else if s = wb.Sheet(r.Sheet); s == nil {
		return nil, grid.Range{}, false, fmt.Errorf("sheet %q not found (sheets: %s)", r.Sheet, strings.Join(wb.SheetNames(), ", "))
	}
	used := s.Used()
	if r.Range == nil {
		return s, used, false, nil
	}
	rg := *r.Range
	if rg.R2 == grid.MaxRows {
		rg.R2 = max(used.R2, rg.R1)
	}
	if rg.C2 == grid.MaxCols {
		rg.C2 = max(used.C2, rg.C1)
	}
	return s, rg, true, nil
}

// cliRef renders a sheet-qualified range for a suggested command. The
// sheets parser splits on the last '!', so sheet names need no Excel-style
// quoting; the shell quoting is enough.
func cliRef(sheet string, r grid.Range) string {
	if strings.Contains(sheet, "!") {
		return shellQuote(grid.FormatRef(sheet, r))
	}
	return shellQuote(sheet + "!" + r.String())
}

func plural(n int, word string) string {
	if n == 1 {
		return "1 " + word
	}
	suffix := "s"
	switch {
	case strings.HasSuffix(word, "series"):
		suffix = ""
	case strings.HasSuffix(word, "ch") || strings.HasSuffix(word, "s") || strings.HasSuffix(word, "x"):
		suffix = "es"
	case strings.HasSuffix(word, "y") && !strings.ContainsRune("aeiou", rune(word[len(word)-2])):
		word, suffix = word[:len(word)-1], "ies"
	}
	return grid.FormatNumber(float64(n)) + " " + word + suffix
}

var negativeNumber = regexp.MustCompile(`^-[0-9.$€£¥]`)

// negatives keeps values like -5 or -$1,200 from being read as flags: when
// args hold one, the flags move in front of a -- and the values after it.
func negatives(root *cobra.Command, args []string) []string {
	has := false
	for _, a := range args {
		if a == "--" {
			return args
		}
		if negativeNumber.MatchString(a) {
			has = true
		}
	}
	if !has {
		return args
	}
	cmd, rest, err := root.Find(args)
	if err != nil {
		return args
	}
	path := args[:len(args)-len(rest)]
	fs := cmd.Flags()
	fs.AddFlagSet(cmd.InheritedFlags())
	var flags, values []string
	for i := 0; i < len(rest); i++ {
		a := rest[i]
		var f *pflag.Flag
		switch {
		case strings.HasPrefix(a, "--") && len(a) > 2:
			name, _, hasValue := strings.Cut(a[2:], "=")
			if f = fs.Lookup(name); f != nil && hasValue {
				flags = append(flags, a)
				continue
			}
		case strings.HasPrefix(a, "-") && len(a) == 2:
			f = fs.ShorthandLookup(a[1:])
		}
		if f == nil {
			values = append(values, a)
			continue
		}
		flags = append(flags, a)
		if f.NoOptDefVal == "" && i+1 < len(rest) {
			i++
			flags = append(flags, rest[i])
		}
	}
	out := append(append(append([]string{}, path...), flags...), "--")
	return append(out, values...)
}

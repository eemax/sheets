package cli

import (
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"time"

	"github.com/dop251/goja"
	"github.com/eemax/sheets/internal/book"
	"github.com/eemax/sheets/internal/grid"
	"github.com/eemax/sheets/internal/xlsx"
	"github.com/evanw/esbuild/pkg/api"
	"github.com/spf13/cobra"
)

// sheets script runs JavaScript or TypeScript against a workbook through
// the Office Scripts API (ExcelScript), the way agents and people already
// script Excel. Reads return what the cells hold now, formula results
// computed by the sheets calculator; the file is saved once at the end, as
// one undo step, and not at all when the script fails.

type scriptResult struct {
	File     string          `json:"file"`
	Console  []string        `json:"console,omitempty"`
	Result   json.RawMessage `json:"result,omitempty"`
	Cells    int             `json:"cells_written"`
	Saved    bool            `json:"saved"`
	DryRun   bool            `json:"dry_run,omitempty"`
	Done     []string        `json:"done,omitempty"` // what else it changed: rows, sheets, styling
	Formulas *formulaReport  `json:"formulas,omitempty"`
	Warnings []string        `json:"warnings,omitempty"`
}

func (a *app) scriptCmd() *cobra.Command {
	var (
		code    string
		args    []string
		timeout time.Duration
	)
	cmd := &cobra.Command{
		Use:   "script <file> [script.ts|script.js]",
		Short: "Run an Office Script (TypeScript or JavaScript) against a workbook",
		Long: `Run a script against a workbook, in Office Scripts' API (ExcelScript), as
in Excel's Automate tab. The script comes from a file, from -e, or from
stdin. It defines main, which gets the workbook:

  sheets script book.xlsx -e '
    function main(workbook: ExcelScript.Workbook) {
      const ws = workbook.getWorksheet("Sales");
      const values = ws.getUsedRange().getValues();
      const total = values.slice(1).reduce((s, r) => s + (r[3] as number), 0);
      ws.getRange("F1").setValue("Total");
      ws.getRange("F2").setFormula("=SUM(D2:D" + values.length + ")");
      return total;
    }'

or it is plain code using the global workbook. Arguments from --arg go to
main after the workbook (JSON values, or text). What main returns is
printed as JSON; console.log prints too.

Reads give what the cells hold now: values, and formula results computed
by the sheets calculator, the script's own changes included. Dates are
serial numbers, as in Excel (getTexts gives them formatted). Writes take
values as typed input: "=SUM(A1:A9)" is a formula, "2026-07-01" a date,
"12%" a percentage. In formulas, {r} stands for the row.

The file is saved once at the end, as one undo step; nothing is saved if
the script throws. Tables are Excel tables, or a sheet's table named after
its sheet; addRows appends to a table's end, filling its formula
columns, above its total row. Scripts run without file or network access,
for at most --timeout.

  sheets script book.xlsx cleanup.ts
  sheets script book.xlsx report.ts --arg 2026 --arg '"West"'`,
		Args: cobra.RangeArgs(1, 2),
		RunE: func(cmd *cobra.Command, argv []string) error {
			file := argv[0]
			var src []byte
			var err error
			name := "script"
			switch {
			case code != "" && len(argv) == 2:
				return errors.New("give the script as a file or with -e, not both")
			case code != "":
				src = []byte(code)
			case len(argv) == 2:
				name = filepath.Base(argv[1])
				if src, err = os.ReadFile(argv[1]); err != nil {
					return err
				}
			default:
				if f, ok := a.stdin.(*os.File); ok && isTerminal(f) {
					return errors.New("no script: give a file, -e 'code', or pipe it on stdin")
				}
				if src, err = io.ReadAll(a.stdin); err != nil {
					return err
				}
			}
			js, err := compileScript(name, string(src))
			if err != nil {
				return err
			}
			b, err := xlsx.StartBatch(file)
			if err != nil {
				return err
			}
			defer b.End()
			s := newScriptSession(a, file)
			defer s.close()
			res := scriptResult{File: file, DryRun: a.dryRun}
			out, runErr := s.run(name, js, args, timeout)
			res.Console = s.console
			if runErr != nil {
				if !a.json {
					for _, l := range s.console {
						a.printf("%s\n", l)
					}
				}
				return fmt.Errorf("%w; nothing was saved", runErr)
			}
			res.Result = out
			if err := s.finish(); err != nil {
				return fmt.Errorf("%w; nothing was saved", err)
			}
			res.Cells, res.Done = s.cells, s.done
			if len(s.styled) > 0 {
				res.Done = append(res.Done, "styling on "+plural(len(s.styled), "range"))
			}
			res.Formulas, res.Warnings = s.report, s.warningList()
			if !a.dryRun {
				if err := b.Commit(); err != nil {
					return err
				}
				res.Saved = b.Changed()
			}
			if a.json {
				return a.printJSON(res)
			}
			a.printScript(res)
			return nil
		},
	}
	f := cmd.Flags()
	f.StringVarP(&code, "eval", "e", "", "the script itself")
	f.StringArrayVar(&args, "arg", nil, "an argument for main, after the workbook: JSON or text (repeatable)")
	f.DurationVar(&timeout, "timeout", time.Minute, "stop the script after this long")
	return cmd
}

func (a *app) printScript(res scriptResult) {
	for _, l := range res.Console {
		a.printf("%s\n", l)
	}
	if len(res.Result) > 0 {
		a.printf("result: %s\n", res.Result)
	}
	what := plural(res.Cells, "cell") + " written"
	if len(res.Done) > 0 {
		what += "; " + strings.Join(res.Done, "; ")
	}
	switch {
	case a.dryRun:
		a.printf("dry run: %s, nothing was saved\n", what)
	case res.Saved:
		a.printf("saved %s: %s, one undo step\n", res.File, what)
	default:
		a.printf("%s: no changes\n", res.File)
	}
	a.printFormulaReport(res.Formulas)
	a.printWarnings(res.Warnings)
	if res.Saved {
		a.hint("sheets diff %s", shellQuote(res.File))
	}
}

// compileScript turns TypeScript (or JavaScript) into JavaScript the
// interpreter runs, keeping the script's line numbers through a source map.
func compileScript(name, src string) (string, error) {
	r := api.Transform(src, api.TransformOptions{
		Loader:     api.LoaderTS,
		Target:     api.ES2017,
		Format:     api.FormatCommonJS,
		Sourcemap:  api.SourceMapInline,
		Sourcefile: name,
	})
	if len(r.Errors) > 0 {
		var msgs []string
		for _, m := range r.Errors {
			at := ""
			if m.Location != nil {
				at = fmt.Sprintf("line %d: ", m.Location.Line)
			}
			msgs = append(msgs, at+m.Text)
		}
		return "", fmt.Errorf("%s: %s", name, strings.Join(msgs, "; "))
	}
	return string(r.Code), nil
}

// scriptSession is a script's hold on the workbook: an editor for its
// writes, and the workbook as saved for its reads. A read after writes
// saves them (into the batch's copy in memory) and reads the result, so
// formulas show what saving stores.
type scriptSession struct {
	a    *app
	file string
	vm   *goja.Runtime

	ed     book.Editor    // open while there are writes to save
	wb     *grid.Workbook // what reads see; nil when writes came since
	sheets map[string]*sheetRef

	console  []string
	cells    int
	done     []string
	written  []cellRef // formulas written, for the report
	report   *formulaReport
	warnings []string
	unknown  map[string]bool  // cells read whose value is unknown
	recent   map[cellKey]bool // formulas the open editor wrote
	styled   map[string]bool  // ranges styled
}

// sheetRef names a sheet for the objects standing for it, through renames.
type sheetRef struct{ name string }

func newScriptSession(a *app, file string) *scriptSession {
	return &scriptSession{a: a, file: file, sheets: map[string]*sheetRef{}, unknown: map[string]bool{}, recent: map[cellKey]bool{}, styled: map[string]bool{}}
}

func (s *scriptSession) close() {
	if s.ed != nil {
		s.ed.Close()
		s.ed = nil
	}
}

// fail ends the script with an error it can catch.
func (s *scriptSession) fail(format string, args ...any) {
	panic(s.vm.NewGoError(fmt.Errorf(format, args...)))
}

// must fails the script on an error.
func (s *scriptSession) must(err error) {
	if err != nil {
		panic(s.vm.NewGoError(err))
	}
}

// editor is the editor for writes, opened on the first.
func (s *scriptSession) editor() book.Editor {
	if s.ed == nil {
		ed, err := book.Edit(s.file, true)
		s.must(err)
		s.ed = ed
	}
	s.wb = nil
	return s.ed
}

// rangeEditor is an editor for a range operation (clear, copy, move,
// sort), which reads the workbook as the editor opened it: the writes
// before it are saved first.
func (s *scriptSession) rangeEditor() book.Editor {
	s.must(s.flush())
	return s.editor()
}

// flush saves the writes so far into the batch.
func (s *scriptSession) flush() error {
	if s.ed == nil {
		return nil
	}
	ed := s.ed
	s.ed = nil
	defer ed.Close()
	if err := ed.Save(false); err != nil {
		return err
	}
	s.warnings = append(s.warnings, ed.Warnings()...)
	s.wb = nil
	s.recent = map[cellKey]bool{}
	return nil
}

// state is the workbook as reads see it: the file as last saved.
func (s *scriptSession) state() *grid.Workbook {
	if s.wb != nil {
		return s.wb
	}
	s.must(s.flush())
	wb, err := book.Open(s.file)
	if errors.Is(err, os.ErrNotExist) {
		// A new file: make it, with its first sheet.
		s.editor()
		s.must(s.flush())
		wb, err = book.Open(s.file)
	}
	s.must(err)
	s.wb = wb
	return wb
}

// sheet finds a sheet as it is now.
func (s *scriptSession) sheet(ref *sheetRef) *grid.Sheet {
	sh := s.state().Sheet(ref.name)
	if sh == nil {
		s.fail("the worksheet %q no longer exists", ref.name)
	}
	return sh
}

// ref is the one sheetRef of a sheet.
func (s *scriptSession) ref(name string) *sheetRef {
	k := strings.ToLower(name)
	if r, ok := s.sheets[k]; ok {
		return r
	}
	r := &sheetRef{name: name}
	s.sheets[k] = r
	return r
}

func (s *scriptSession) renamed(from, to string) {
	r := s.ref(from)
	delete(s.sheets, strings.ToLower(from))
	r.name = to
	s.sheets[strings.ToLower(to)] = r
}

// run runs the script and gives main's result as JSON.
func (s *scriptSession) run(name, js string, args []string, timeout time.Duration) (json.RawMessage, error) {
	vm := goja.New()
	s.vm = vm
	vm.SetFieldNameMapper(goja.UncapFieldNameMapper())
	module := vm.NewObject()
	exports := vm.NewObject()
	module.Set("exports", exports)
	vm.Set("module", module)
	vm.Set("exports", exports)
	vm.Set("workbook", &jsWorkbook{s})
	vm.Set("__log", func(line string) { s.console = append(s.console, line) })
	if _, err := vm.RunString(scriptPrelude); err != nil {
		return nil, fmt.Errorf("internal error in the script runtime: %v", err)
	}
	if timeout > 0 {
		t := time.AfterFunc(timeout, func() { vm.Interrupt(fmt.Sprintf("the script ran longer than %s (--timeout)", timeout)) })
		defer t.Stop()
	}
	if _, err := vm.RunScript(name, js); err != nil {
		return nil, scriptError(err)
	}
	main := vm.Get("main")
	if main == nil || goja.IsUndefined(main) {
		main = exports.Get("main")
	}
	fn, ok := goja.AssertFunction(main)
	if !ok {
		return nil, nil // plain code, run already
	}
	params := []goja.Value{vm.Get("workbook")}
	for _, a := range args {
		var v any
		if err := json.Unmarshal([]byte(a), &v); err != nil {
			v = a
		}
		params = append(params, vm.ToValue(v))
	}
	out, err := fn(goja.Undefined(), params...)
	if err != nil {
		return nil, scriptError(err)
	}
	if p, ok := out.Export().(*goja.Promise); ok {
		// An async main: its jobs have run with the call.
		switch p.State() {
		case goja.PromiseStateRejected:
			return nil, fmt.Errorf("script error: %s", thrown(p.Result()))
		case goja.PromiseStatePending:
			return nil, errors.New("script error: main's promise never settled")
		}
		out = p.Result()
	}
	if out == nil || goja.IsUndefined(out) {
		return nil, nil
	}
	text, err := vm.RunString("(function(v){ return JSON.stringify(v); })")
	if err != nil {
		return nil, err
	}
	stringify, _ := goja.AssertFunction(text)
	j, err := stringify(goja.Undefined(), out)
	if err != nil || goja.IsUndefined(j) {
		return json.RawMessage(fmt.Sprintf("%q", out.String())), nil
	}
	return json.RawMessage(j.String()), nil
}

// thrown describes a thrown value: an error's name and message.
func thrown(v goja.Value) string {
	if v == nil {
		return "undefined"
	}
	msg := v.String()
	if o, ok := v.(*goja.Object); ok {
		if m := o.Get("message"); m != nil && !goja.IsUndefined(m) {
			if n := o.Get("name"); n != nil && !goja.IsUndefined(n) && n.String() != "GoError" && n.String() != "Error" {
				msg = n.String() + ": " + m.String()
			} else {
				msg = m.String()
			}
		}
	}
	return msg
}

// scriptError reports what the script threw, with the line.
func scriptError(err error) error {
	var ex *goja.Exception
	if errors.As(err, &ex) {
		msg := thrown(ex.Value())
		if at := firstFrame(ex.String()); at != "" {
			msg += " (" + at + ")"
		}
		return fmt.Errorf("script error: %s", msg)
	}
	var in *goja.InterruptedError
	if errors.As(err, &in) {
		return fmt.Errorf("script stopped: %v", in.Value())
	}
	return fmt.Errorf("script error: %v", err)
}

// firstFrame picks the script's own place out of a stack trace: "at
// main (report.ts:4:9(18))" gives report.ts line 4.
func firstFrame(stack string) string {
	for _, l := range strings.Split(stack, "\n") {
		if m := framePlace.FindStringSubmatch(l); m != nil && !strings.Contains(l, "native") {
			return m[1] + " line " + m[2]
		}
	}
	return ""
}

var framePlace = regexp.MustCompile(`at .*?\(?([^\s()]+):(\d+):\d+`)

// finish saves the script's last writes and reports its formulas: their
// results as saved, and for those written since the last read, why the
// calculator left one out.
func (s *scriptSession) finish() error {
	last := s.ed
	if last != nil {
		s.ed = nil
		defer last.Close()
		if err := last.Save(false); err != nil {
			return err
		}
		s.warnings = append(s.warnings, last.Warnings()...)
		s.wb = nil
	}
	if len(s.written) == 0 {
		return nil
	}
	wb, err := book.Open(s.file)
	if err != nil {
		return err
	}
	// The formulas still where they were written: not moved or deleted.
	var still []cellRef
	seen := map[cellKey]bool{}
	for _, c := range s.written {
		k := cellKey{strings.ToLower(c.sheet), c.row, c.col}
		if sh := wb.Sheet(c.sheet); sh != nil && !seen[k] && sh.Get(c.row, c.col).Formula != "" {
			seen[k] = true
			still = append(still, c)
		}
	}
	s.report = evaluateFormulas(&savedState{wb: wb, last: last, recent: s.recent}, still, true)
	return nil
}

// savedState evaluates cells from the workbook as saved.
type savedState struct {
	wb     *grid.Workbook
	last   book.Editor      // the editor of the last writes
	recent map[cellKey]bool // the cells it wrote
}

func (st *savedState) Evaluate(sheet string, row, col int) (grid.Cell, string, bool) {
	if st.last != nil && st.recent[cellKey{strings.ToLower(sheet), row, col}] {
		return st.last.Evaluate(sheet, row, col)
	}
	sh := st.wb.Sheet(sheet)
	if sh == nil {
		return grid.Cell{}, "no such sheet", false
	}
	c := st.wb.Value(sh, row, col)
	switch {
	case c == nil:
		return grid.Cell{}, "", true
	case c.Pending:
		return grid.Cell{}, "the sheets calculator cannot compute it; Excel computes it on open", false
	}
	return grid.Cell{Kind: c.Kind, Num: c.Num, Text: c.Text, Bool: c.Bool}, "", true
}

type cellKey struct {
	sheet    string
	row, col int
}

// warningList gives the warnings once each, in order.
func (s *scriptSession) warningList() []string {
	var out []string
	seen := map[string]bool{}
	for _, w := range s.warnings {
		if !seen[w] {
			seen[w] = true
			out = append(out, w)
		}
	}
	if len(s.unknown) > 0 {
		var cells []string
		for c := range s.unknown {
			cells = append(cells, c)
		}
		sortCells(cells)
		list := cells
		if len(list) > 5 {
			list = append(list[:5:5], fmt.Sprintf("and %d more", len(cells)-5))
		}
		out = append(out, fmt.Sprintf("the script read %s whose formula result the sheets calculator cannot compute, as null: %s", plural(len(cells), "cell"), strings.Join(list, ", ")))
	}
	return out
}

// scriptPrelude defines console and the ExcelScript enums.
const scriptPrelude = `
var console = (function () {
  function show(v) {
    if (typeof v === "string") return v;
    if (v === undefined) return "undefined";
    if (typeof v === "function") return "[function]";
    try { var j = JSON.stringify(v); return j === undefined ? String(v) : j; } catch (e) { return String(v); }
  }
  function log() { __log(Array.prototype.map.call(arguments, show).join(" ")); }
  return { log: log, info: log, warn: log, error: log, debug: log };
})();
// Every ExcelScript enum: a member's value is its name capitalized, as in
// Office Scripts (ExcelScript.BorderIndex.edgeBottom is "EdgeBottom").
var ExcelScript = (function () {
  function member(t, k) {
    if (typeof k !== "string") return t[k];
    return k.charAt(0).toUpperCase() + k.slice(1);
  }
  var enums = {};
  return new Proxy(enums, {
    get: function (t, k) {
      if (typeof k !== "string") return t[k];
      if (!(k in t)) t[k] = new Proxy({}, { get: member });
      return t[k];
    },
  });
})();
`

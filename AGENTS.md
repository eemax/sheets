# sheets

A Go CLI that lets AI agents read, write, format and check Excel (`.xlsx`)
and CSV files without Excel. Module `github.com/eemax/sheets`.

## Read first

- **[docs/vision.md](docs/vision.md):** where this is going and why. Every
  change should serve it.
- **[docs/roadmap.md](docs/roadmap.md):** the plan, as work packages, and
  how to run evals.
- **[docs/orchestration.md](docs/orchestration.md):** how work is done
  here. An orchestrator session hands out packages, workers build in their
  own worktrees, independent auditors check, and only the orchestrator
  merges into `main` and pushes. A worker or auditor follows its brief.

## Principles (from the vision)

- **Excel is the specification:** semantics, file format, quirks.
- **Unknown over wrong.** Never store or report a value sheets cannot be
  sure of; say it is unknown, and why.
- **No silent changes.** Every change is visible (diff) and reversible
  (undo).
- **Output sized for a model's context.**
- **Verify adversarially before calling anything done.** Cover:
  - files from other producers (openpyxl, XlsxWriter, LibreOffice);
  - both editors (streaming and full);
  - CSV;
  - large sheets;
  - how it looks (`sheets render`).

## Code and tests

- **The engine** (`internal/engine`) computes formulas. `internal/xlsx`
  reads and edits files. `internal/cli` holds the commands, and
  `internal/delim` CSV.
- **Command tests** are txtar scripts in `cmd/sheets/testdata`.
- **Fixtures** from other producers come from `scripts/fixtures.py`
  (openpyxl, XlsxWriter). To add one without churning the others,
  generate into a temporary folder and copy only the new file.
- **Checks** before calling work done:
  - `go vet ./...` and `go test ./...`;
  - `SHEETS_STREAM=always go test ./cmd/... ./internal/xlsx/... ./internal/cli/...`;
  - `go test -race` on what you touched.
- **The agent guide** is `internal/cli/agent-docs.md`. The skill
  (`skills/sheets`) is generated from it with
  `go run ./cmd/sheets skill --install --dir skills`, and a test fails
  when they drift. Code spans must not break across lines.
- **Eval data** lives in `~/sheets-evals`, never in a session scratchpad
  (macOS empties `/private/tmp`).

## Settled decisions (not to be reopened unasked)

- **No patched or forked excelize** (Max, 2026-10-23).
- **Prebuilt binaries are deferred** (Max, 2026-10-24).
- **Better-looking charts are parked** (Max, 2026-10-23).
- **Scripting is JavaScript inside sheets** (goja, with esbuild for
  TypeScript, and the Office Scripts API). A Python client would sit on
  the same command set.
- **The house style** is set. It has:
  - a dark header, with labels left-aligned in every column;
  - thin borders and a bold total row;
  - a frozen header;
  - ISO dates;
  - negatives in parentheses;
  - row banding through conditional formatting, not an Excel table.

  On sheets people made, `fmt` replaces their look with this style rather
  than tidying it.
- **Breaking changes are fine.** sheets has no external users to keep
  compatible with yet.
- **Anthropic's xlsx skill is proprietary.** Use it locally, as an eval
  baseline, and never commit it.

## Git

- Work on `main`, unless you are a worker in your own worktree branch.
- Commit by path: `git commit -m "…" -- <files>`. Never use `git add -A`
  or a bare `git commit`: another session may share this checkout.
- Never force-push.

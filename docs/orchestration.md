# How the work is run

Max's model (2026-10-31): one orchestrator session holds the vision and the
plan. It hands each work package to a worker session, has every result
audited by a fresh session that did not write it, and decides what comes
next. The vision ([vision.md](vision.md)) and the plan
([roadmap.md](roadmap.md)) are what keep the project coherent across
sessions.

## Roles

- **Orchestrator** (the main session):
  - keeps the vision and the roadmap;
  - writes the briefs;
  - spawns workers and auditors;
  - triages findings;
  - merges into `main`, pushes, and keeps the board and the portal thread
    current;
  - reports to Max at the gates.

  It does not write feature code itself, except to resolve a merge or a
  one-line fix found while merging, which is named in the merge commit.
- **Worker:** one package, in its own git worktree and branch, starting
  from a fresh context.
  - **It:** builds, tests and commits on its branch.
  - **It does not:** push, merge into `main`, edit the vision or the
    board, or start work outside its package. It may propose those in its
    report.
- **Auditor:** a fresh session that did not write the code. It checks the
  branch against the package's acceptance criteria, the principles in the
  vision and the definition of done, adversarially.
  - **It does not:** fix anything or commit.
  - **It gives:** a verdict, with evidence.

## The loop

1. **Brief.** The orchestrator writes the worker brief (template below)
   from the package in the roadmap. The board says `working`.
2. **Work.** The worker runs with `isolation: worktree`, then reports in
   the format below. The board says `auditing`.
3. **Audit.** A fresh auditor gets the auditor brief, the branch, and the
   worker's claims. It does not get the worker's reasoning, so that it
   forms its own view.
4. **Triage.** The orchestrator accepts each finding, or rejects it with a
   reason.
   - **No blockers or majors accepted:** go to step 6.
   - **Otherwise:** the accepted findings go back to the same worker
     (SendMessage; it keeps its context). The board says `fixing`.
5. **Re-audit.** The same auditor checks its own findings and looks for
   regressions. A fresh auditor is used instead when the fix was large.
   After three rounds without a pass, the package goes to Max.
6. **Merge.** The orchestrator:
   - squash-merges the branch into `main`, as one commit for the package
     (see Git);
   - runs the full checks on `main`;
   - pushes;
   - marks the package `done`, updates the portal thread, and removes the
     worktree and the branch.
7. **Next.** The orchestrator picks the next package from the board,
   unless a gate is reached. At a gate it reports to Max and waits.

At most two workers run at once, in packages that touch different code. An
auditor starts as soon as its worker finishes.

## Definition of done (every package)

- **Checks:** all of these pass:
  - `go vet ./...` and `go test ./...`;
  - `SHEETS_STREAM=always go test ./cmd/... ./internal/xlsx/... ./internal/cli/...`;
  - `go test -race` on the packages it touched.
- **Tests:** new behaviour has tests, including the edge cases the auditor
  is going to try:
  - files written by openpyxl and XlsxWriter;
  - both the streaming and the full editor;
  - CSV, where it applies;
  - unknown values;
  - large sheets.
- **Docs:** user-facing behaviour is documented.
  - **The agent guide:** `internal/cli/agent-docs.md`.
  - **The skill:** regenerated with
    `go run ./cmd/sheets skill --install --dir skills`.
  - **The README:** where the change is visible to users.
- **Scope:** only the package's own changes are included.
- **Dependencies:** a new one only with the orchestrator's agreement. A
  major one, or anything installed on the machine, needs Max's.
- **Decisions:** settled ones (in `AGENTS.md`) are respected.

## Worker brief (template)

```
You are a worker on sheets, a Go CLI that lets AI agents read, write and
check Excel and CSV files. Read AGENTS.md, docs/vision.md and the package
below before anything else.

Package: <WP-nn and its full text from docs/roadmap.md>

Context: <the code areas, earlier packages, results or decisions this
package builds on>

Rules:
- Work only in your worktree, on your branch. Commit there, by path
  (git commit -m "…" -- <files>). Do not push or touch main.
- Meet the definition of done in docs/orchestration.md.
- Verify adversarially before you report: try to break what you built
  (files from other tools, both editors, edge cases), not just run the
  happy path.
- If the package's shape turns out wrong, stop and report rather than
  building something else.

Report, at the end, in this form:
- Summary: what the package does now, in a few sentences.
- Branch and commits.
- Decisions: each choice the package left open, what you chose, and why.
- Tests: what you added, and the tail of each check's output.
- Claims for the auditor: numbered, each checkable (a command and the
  result to expect).
- Known limits and follow-ups.
```

## Auditor brief (template)

```
You are an independent auditor for sheets, a Go CLI for agents working
with Excel and CSV files. You did not write the code you are auditing.
Your job is to find what is wrong with it, not to confirm that it works.
Read AGENTS.md and docs/vision.md first.

Package: <WP-nn and its full text from docs/roadmap.md>

The work: branch <name>, commits <range> (check it out detached in your
own worktree: git switch --detach <branch>). The worker's claims, which
you verify rather than trust: <numbered claims>.

Check:
1. Each acceptance criterion, with evidence (the command and its output).
2. The definition of done in docs/orchestration.md: run the checks
   yourself.
3. The vision's principles: Excel is the specification; unknown over
   wrong; no silent changes; output sized for a model.
4. Try to break it: files written by openpyxl and XlsxWriter, the
   streaming and the full editor, CSV, unknown values, large sheets,
   odd input, documented behaviour against actual behaviour.
5. Scope: changes outside the package, and settled decisions reopened.

Do not fix anything and do not commit. You may write scratch tests in your
worktree to prove a finding; put them in your report.

Report:
- Verdict: PASS, PASS WITH NOTES, or FAIL.
- Findings, most severe first: severity (blocker, major, minor, nit),
  what is wrong, the evidence (command and output), a suggested fix.
- Claims: each one, verified or not, with evidence.
- What you tried that held up.
```

A **blocker** is wrong values, a corrupt file or lost data. A **major** is a
missed acceptance criterion, behaviour that contradicts the docs, or a
broken edge case a user will hit. **Minor** and **nit** findings are merged
with notes, or fixed in passing.

## Git

- Workers commit on their worktree branch, by path. Only the orchestrator
  touches `main`.
- A package goes into `main` as one squash commit, written as the earlier
  commits are written: what it does and why, the audit verdict, and the
  co-author line.
- The orchestrator commits by path, never with `git add -A`. Another
  session may be working in the same checkout.
- Nobody force-pushes.

## Budget

Sub-sessions use Max's Pro plan.

- **Before each spawn,** the orchestrator checks usage.
- **New spawns wait** while the 5-hour window is above 75%, or the weekly
  one above 70%.
- **Never into extra usage.** The work stops there, and Max is told.
- **Models:** workers and auditors run on the default model. Eval runs use
  the model the experiment names.

## Reporting

- **In conversation:** the orchestrator reports to Max at the end of each
  wave and at gates. It says what was done, what the audits found, what
  was decided, and what comes next.
- **When he is away:** the same, short, through `~/portal/bin/portal notify`.
- **Questions only Max can answer:** asked with `portal ask`. No answer
  means waiting, not guessing.

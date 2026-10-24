#!/usr/bin/env python3
"""A/B evaluation of agents doing spreadsheet tasks: with the sheets CLI,
and with Python (openpyxl, pandas), the way agents work without it.

    evals/run.py prepare OUT     # a folder per task and arm: inputs and the prompt
    evals/run.py check OUT       # grade what the agents left, write OUT/report.md
    evals/run.py run OUT         # prepare, run each agent with `claude -p`, check

Each task in evals/tasks has prompt.md (the request, the same for both arms),
setup.py (writes the input files, and the true answers to a file the agent
does not see) and check.py (grades the result). An agent works in
OUT/<task>/<arm>/work and its final reply goes to answer.txt beside it;
the true answers are kept in OUT/.truth.

Checks of kind "stored" (whether a file stores formula results, as pandas
and previews read them) are counted apart from the task's pass or fail.
"""
import argparse
import json
import os
import shutil
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(ROOT)
TASKS = os.path.join(ROOT, "tasks")
ARMS = ["sheets", "python"]


def tasks(only):
    names = sorted(d for d in os.listdir(TASKS) if os.path.isdir(os.path.join(TASKS, d)))
    return [n for n in names if not only or any(n.startswith(o) for o in only)]


def sheets_bin(out):
    """Build the sheets binary into OUT/bin and return its folder."""
    bindir = os.path.join(out, "bin")
    os.makedirs(bindir, exist_ok=True)
    subprocess.run(["go", "build", "-o", os.path.join(bindir, "sheets"), "./cmd/sheets"], cwd=REPO, check=True)
    return bindir


def uv(*args, **kw):
    return subprocess.run(["uv", "run", "-q", "--with", "openpyxl", "--with", "xlsxwriter", "python", *args], **kw)


def truth_path(out, task, arm):
    """Where the true answers go: away from the agents' folders."""
    d = os.path.join(out, ".truth", task)
    os.makedirs(d, exist_ok=True)
    return os.path.join(d, arm + ".json")


def prompt_for(task, arm, work, bindir):
    request = open(os.path.join(TASKS, task, "prompt.md")).read().strip()
    guide = open(os.path.join(ROOT, "arms", arm + ".md")).read().strip()
    if arm == "sheets":
        docs = subprocess.run([os.path.join(bindir, "sheets"), "agent-docs"], capture_output=True, text=True, check=True).stdout
        guide += "\n\n" + docs.strip()
    elif os.environ.get("EVAL_BASELINE_SKILL"):
        # The skill agents use for spreadsheets without sheets, if given.
        guide += "\n\n" + open(os.environ["EVAL_BASELINE_SKILL"]).read().strip()
    return (f"{guide}\n\n---\n\nThe files are in {work}. Work only there.\n\n{request}\n\n"
            "When you are done, end with your answer (or a short summary of what you changed).")


def prepare(out, only):
    bindir = sheets_bin(out)
    for task in tasks(only):
        for arm in ARMS:
            d = os.path.join(out, task, arm)
            work = os.path.join(d, "work")
            shutil.rmtree(d, ignore_errors=True)
            os.makedirs(work)
            uv(os.path.join(TASKS, task, "setup.py"), work, truth_path(out, task, arm), check=True)
            with open(os.path.join(d, "prompt.txt"), "w") as f:
                f.write(prompt_for(task, arm, work, bindir))
    return bindir


def run_agent(d, arm, bindir, model, budget):
    """Run one agent with the Claude Code CLI; returns its metrics."""
    work = os.path.join(d, "work")
    env = dict(os.environ)
    path = [p for p in env.get("PATH", "").split(os.pathsep) if not os.path.exists(os.path.join(p, "sheets"))]
    if arm == "sheets":
        path.insert(0, bindir)
    env["PATH"] = os.pathsep.join(path)
    env["SHEETS_CACHE_DIR"] = os.path.join(d, ".cache")
    cmd = ["claude", "-p", open(os.path.join(d, "prompt.txt")).read(), "--output-format", "json",
           "--model", model, "--max-budget-usd", str(budget),
           "--allowedTools", "Bash", "Read", "Write", "Edit", "Glob", "Grep",
           "--disallowedTools", "WebFetch", "WebSearch", "Task"]
    start = time.time()
    res = subprocess.run(cmd, cwd=work, env=env, capture_output=True, text=True)
    try:
        out = json.loads(res.stdout)
    except json.JSONDecodeError:
        out = {"result": res.stdout + res.stderr, "is_error": True}
    with open(os.path.join(d, "answer.txt"), "w") as f:
        f.write(out.get("result") or "")
    usage = out.get("usage") or {}
    meta = {"turns": out.get("num_turns"), "cost_usd": out.get("total_cost_usd"),
            "seconds": round(time.time() - start, 1), "error": out.get("is_error"),
            "tokens": sum(usage.get(k, 0) or 0 for k in ("input_tokens", "output_tokens", "cache_creation_input_tokens", "cache_read_input_tokens")),
            "output_tokens": usage.get("output_tokens")}
    with open(os.path.join(d, "meta.json"), "w") as f:
        json.dump(meta, f)
    return meta


def check(out, only):
    bindir = os.path.join(out, "bin")
    env = dict(os.environ, SHEETS_BIN=os.path.join(bindir, "sheets"), PYTHONPATH=os.path.join(ROOT, "lib"), SHEETS_UNDO="off")
    results = {}
    for task in tasks(only):
        for arm in ARMS:
            d = os.path.join(out, task, arm)
            if not os.path.isdir(d):
                continue
            res = uv(os.path.join(TASKS, task, "check.py"), os.path.join(d, "work"), truth_path(out, task, arm),
                     os.path.join(d, "answer.txt"), capture_output=True, text=True, env=env)
            try:
                r = json.loads(res.stdout.strip().splitlines()[-1])
            except (json.JSONDecodeError, IndexError):
                r = {"passed": False, "checks": [{"name": "check ran", "ok": False, "detail": (res.stderr or res.stdout)[-300:], "kind": "task"}]}
            meta = os.path.join(d, "meta.json")
            r["meta"] = json.load(open(meta)) if os.path.exists(meta) else {}
            with open(os.path.join(d, "result.json"), "w") as f:
                json.dump(r, f, indent=1)
            results[(task, arm)] = r
    write_report(out, results)
    return results


def write_report(out, results):
    lines = ["# sheets A/B evaluation", "",
             "| task | arm | pass | checks | stored results | turns | tokens | cost $ |", "|---|---|---|---|---|---|---|---|"]
    totals = {arm: {"pass": 0, "n": 0, "stored": 0, "stored_n": 0, "turns": 0, "tokens": 0, "cost": 0.0} for arm in ARMS}
    for (task, arm), r in sorted(results.items()):
        task_checks = [c for c in r["checks"] if c["kind"] == "task"]
        stored = [c for c in r["checks"] if c["kind"] == "stored"]
        m = r.get("meta", {})
        t = totals[arm]
        t["n"] += 1
        t["pass"] += r["passed"]
        t["stored"] += sum(c["ok"] for c in stored)
        t["stored_n"] += len(stored)
        t["turns"] += m.get("turns") or 0
        t["tokens"] += m.get("tokens") or 0
        t["cost"] += m.get("cost_usd") or 0
        lines.append(f"| {task} | {arm} | {'✅' if r['passed'] else '❌'} | {sum(c['ok'] for c in task_checks)}/{len(task_checks)} | "
                     f"{sum(c['ok'] for c in stored)}/{len(stored)} | {m.get('turns', '')} | {m.get('tokens', '')} | {m.get('cost_usd', '')} |")
    lines += ["", "| arm | tasks passed | stored results | turns | tokens | cost $ |", "|---|---|---|---|---|---|"]
    for arm, t in totals.items():
        lines.append(f"| {arm} | {t['pass']}/{t['n']} | {t['stored']}/{t['stored_n']} | {t['turns']} | {t['tokens']} | {t['cost']:.2f} |")
    lines += ["", "## Failed checks", ""]
    for (task, arm), r in sorted(results.items()):
        for c in r["checks"]:
            if not c["ok"]:
                lines.append(f"- {task} / {arm} / {c['kind']}: {c['name']} — {c['detail']}")
    with open(os.path.join(out, "report.md"), "w") as f:
        f.write("\n".join(lines) + "\n")
    print("\n".join(lines))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("action", choices=["prepare", "check", "run"])
    ap.add_argument("out")
    ap.add_argument("--tasks", nargs="*", help="task name prefixes (default: all)")
    ap.add_argument("--model", default="sonnet")
    ap.add_argument("--budget", type=float, default=2.0, help="USD per agent run")
    a = ap.parse_args()
    out = os.path.abspath(a.out)
    os.makedirs(out, exist_ok=True)
    if a.action == "prepare":
        prepare(out, a.tasks)
    elif a.action == "check":
        check(out, a.tasks)
    else:
        bindir = prepare(out, a.tasks)
        for task in tasks(a.tasks):
            for arm in ARMS:
                print(f"{task} / {arm} ...", file=sys.stderr)
                run_agent(os.path.join(out, task, arm), arm, bindir, a.model, a.budget)
        check(out, a.tasks)


if __name__ == "__main__":
    main()

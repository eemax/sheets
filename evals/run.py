#!/usr/bin/env python3
"""A/B evaluation of agents doing spreadsheet tasks: with the sheets CLI,
and the ways agents work without it.

    evals/run.py prepare OUT     # a folder per task and arm: inputs and the prompt
    evals/run.py check OUT       # grade what the agents left, write OUT/report.md
    evals/run.py run OUT         # prepare, run each agent with `claude -p`, check
    evals/run.py run OUT --more  # only the runs not finished yet
    evals/run.py run OUT --redo 10/python   # rerun named runs (old meta kept)
    evals/run.py validate OUT    # check the checks with each task's solutions/

Arms: sheets (the CLI and its guide), python (openpyxl and pandas),
python-hint (python, plus one line on the _xlfn. prefix), skill (Python
with a spreadsheet skill: Anthropic's xlsx skill in arms/xlsx-skill, or the
SKILL.md EVAL_BASELINE_SKILL names), skill-lo (the xlsx skill with
LibreOffice there for its recalc.py) and mcp (the excel-mcp-server tools,
run with uvx).

Each task in evals/tasks has prompt.md (the request, the same for both arms),
setup.py (writes the input files, and the true answers to a file the agent
does not see) and check.py (grades the result). An agent works in
OUT/<task>/<arm>/work and its final reply goes to answer.txt beside it;
the true answers are kept in OUT/.truth.

Checks of kind "stored" (whether a file stores formula results, as pandas
and previews read them) are counted apart from the task's pass or fail.
"""
import argparse
import concurrent.futures
import json
import os
import re
import shutil
import signal
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(ROOT)
TASKS = os.path.join(ROOT, "tasks")
ARMS = ["sheets", "python", "skill", "mcp", "skill-lo", "python-hint"]
# The longest an agent may run, in minutes (run --timeout); the slowest
# finished run so far took about 8 minutes.
TIMEOUT_MIN = 20
# The committed xlsx skill (Apache 2.0, see its NOTICE).
SKILL = os.path.join(ROOT, "arms", "xlsx-skill", "SKILL.md")
SOFFICE = "/opt/homebrew/bin/soffice"
MCP_SERVER = ["uvx", "--from", "excel-mcp-server==2.2.2", "excel-mcp-server", "stdio"]


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
    kw.setdefault("env", dict(os.environ, PYTHONPATH=os.path.join(ROOT, "lib")))
    return subprocess.run(["uv", "run", "-q", "--with", "openpyxl", "--with", "xlsxwriter", "python", *args], **kw)


def truth_path(out, task, arm):
    """Where the true answers go: away from the agents' folders."""
    d = os.path.join(out, ".truth", task)
    os.makedirs(d, exist_ok=True)
    return os.path.join(d, arm + ".json")


def skill_src(arm):
    """The SKILL.md an arm loads: the committed one, except that the skill
    arm follows EVAL_BASELINE_SKILL."""
    return SKILL if arm == "skill-lo" else os.environ.get("EVAL_BASELINE_SKILL") or SKILL


def baseline_skill(out, arm="skill"):
    """Copy the baseline skill's folder into OUT/skills and return it and
    its instructions (SKILL.md without its frontmatter): the committed
    xlsx skill, or, for the skill arm only, the SKILL.md
    EVAL_BASELINE_SKILL names. The skill-lo arm always loads the committed
    one."""
    src = skill_src(arm)
    d = os.path.join(out, "skills", os.path.basename(os.path.dirname(os.path.abspath(src))))
    if not os.path.isdir(d):
        shutil.copytree(os.path.dirname(os.path.abspath(src)), d)
    body = re.sub(r"\A---\n.*?\n---\n", "", open(src).read(), flags=re.S).strip()
    return d, body


def prompt_for(task, arm, work, bindir):
    request = open(os.path.join(TASKS, task, "prompt.md")).read().strip()
    guide = open(os.path.join(ROOT, "arms", arm + ".md")).read().strip()
    if arm == "sheets":
        docs = subprocess.run([os.path.join(bindir, "sheets"), "agent-docs"], capture_output=True, text=True, check=True).stdout
        guide += "\n\n" + docs.strip()
    elif arm in ("skill", "skill-lo"):
        # As Claude Code loads a skill: its folder, then its instructions.
        d, body = baseline_skill(os.path.dirname(bindir), arm)
        guide += f"\n\nBase directory for this skill: {d}\n\n{body}"
    return (f"{guide}\n\n---\n\nThe files are in {work}. Work only there.\n\n{request}\n\n"
            "When you are done, end with your answer (or a short summary of what you changed).")


def prepare(out, only, arms, runs=None):
    """Write each run's folder afresh: its inputs and its prompt. RUNS,
    when given, are the (task, arm) pairs to prepare instead of every task
    and arm."""
    bindir = sheets_bin(out)
    for task, arm in runs if runs is not None else [(t, a) for t in tasks(only) for a in arms]:
        d = os.path.join(out, task, arm)
        work = os.path.join(d, "work")
        prev = os.path.join(d, "meta.prev.json")
        kept = open(prev).read() if os.path.exists(prev) else None  # an interrupted --redo's old meta
        shutil.rmtree(d, ignore_errors=True)
        os.makedirs(work)
        if kept is not None:
            with open(prev, "w") as f:
                f.write(kept)
        uv(os.path.join(TASKS, task, "setup.py"), work, truth_path(out, task, arm), check=True)
        with open(os.path.join(d, "prompt.txt"), "w") as f:
            f.write(prompt_for(task, arm, work, bindir))
        if arm in ("skill", "skill-lo"):
            with open(os.path.join(d, "skill.txt"), "w") as f:
                f.write(os.path.abspath(skill_src(arm)) + "\n")
    return bindir


def finished(out, task, arm):
    """Whether a run has ended: its agent wrote meta.json."""
    return os.path.exists(os.path.join(out, task, arm, "meta.json"))


def redo_runs(names):
    """The (task, arm) pairs of TASK/ARM names; TASK may be a prefix that
    names one task."""
    runs = []
    for name in names:
        task, _, arm = name.partition("/")
        found = tasks([task])
        if len(found) != 1 or arm not in ARMS:
            sys.exit(f"--redo {name}: give TASK/ARM, with a task prefix naming one task and an arm of {', '.join(ARMS)}")
        runs.append((found[0], arm))
    return runs


def soffice_dir():
    """The folder of LibreOffice's soffice, for the skill-lo arm."""
    found = shutil.which("soffice") or (SOFFICE if os.access(SOFFICE, os.X_OK) else None)
    if not found:
        sys.exit(f"the skill-lo arm needs LibreOffice: no soffice on the PATH or at {SOFFICE}")
    return os.path.dirname(found)


def run_agent(d, arm, bindir, model, budget, *, timeout_min=TIMEOUT_MIN):
    """Run one agent with the Claude Code CLI; returns its metrics. The
    agent and its tools are stopped after TIMEOUT_MIN minutes."""
    timeout = timeout_min * 60
    work = os.path.join(d, "work")
    env = dict(os.environ)
    path = path_without_sheets(os.path.dirname(bindir), env.get("PATH", ""))
    if arm == "skill-lo":
        # The skill's own setting: LibreOffice there for its recalc.py, and
        # no stand-ins in front of it.
        path = [p for p in path if os.path.basename(p) != "no-libreoffice"]
        if soffice_dir() not in path:
            path.insert(0, soffice_dir())
    else:
        # LibreOffice, though installed here for grading, is not there for
        # the agents, as on most machines.
        path.insert(0, no_libreoffice(os.path.dirname(bindir)))
    if arm == "sheets":
        path.insert(0, bindir)
    env["PATH"] = os.pathsep.join(path)
    env["SHEETS_CACHE_DIR"] = os.path.join(d, ".cache")
    # No user or project settings, skills or MCP servers: each arm gets only
    # what it is given.
    tools = ["Bash", "Read", "Write", "Edit", "Glob", "Grep", "ToolSearch"]
    cmd = ["claude", "-p", open(os.path.join(d, "prompt.txt")).read(), "--output-format", "stream-json", "--verbose",
           "--model", model, "--max-budget-usd", str(budget), "--no-session-persistence",
           "--setting-sources", "", "--disable-slash-commands", "--strict-mcp-config"]
    if arm == "mcp":
        config = {"mcpServers": {"excel": {"command": MCP_SERVER[0], "args": MCP_SERVER[1:] + ["--allow-dir", work]}}}
        cmd += ["--mcp-config", json.dumps(config)]
        tools.append("mcp__excel")
    cmd += ["--allowedTools", *tools, "--disallowedTools", "WebFetch", "WebSearch", "Task", "Agent"]
    start = time.time()
    # No stdin: a command that reads it ends at once instead of waiting.
    # An agent whose tool hangs is stopped, with the tool, after timeout_min
    # minutes and counts as an error.
    p = subprocess.Popen(cmd, cwd=work, env=env, stdin=subprocess.DEVNULL, stdout=subprocess.PIPE,
                         stderr=subprocess.PIPE, text=True, start_new_session=True)
    try:
        stdout, stderr = p.communicate(timeout=timeout)
    except subprocess.TimeoutExpired:
        os.killpg(p.pid, signal.SIGKILL)
        stdout, stderr = p.communicate()
        stderr += f"\nstopped after {timeout} s"
    res = subprocess.CompletedProcess(cmd, p.returncode, stdout, stderr)
    # The transcript, kept to see what each agent did; its last event is the
    # result.
    with open(os.path.join(d, "transcript.jsonl"), "w") as f:
        f.write(res.stdout)
    events = []
    for line in res.stdout.splitlines():
        try:
            events.append(json.loads(line))
        except json.JSONDecodeError:
            pass
    out = next((e for e in reversed(events) if e.get("type") == "result"), None) or {"result": res.stdout[-2000:] + res.stderr[-2000:], "is_error": True}
    with open(os.path.join(d, "answer.txt"), "w") as f:
        f.write(out.get("result") or "")
    usage = out.get("usage") or {}
    meta = {"turns": out.get("num_turns"), "cost_usd": out.get("total_cost_usd"),
            "seconds": round(time.time() - start, 1), "error": out.get("is_error"),
            "tokens": sum(usage.get(k, 0) or 0 for k in ("input_tokens", "output_tokens", "cache_creation_input_tokens", "cache_read_input_tokens")),
            "output_tokens": usage.get("output_tokens"), "tools": tool_uses(events)}
    with open(os.path.join(d, "meta.json"), "w") as f:
        json.dump(meta, f)
    return meta


def path_without_sheets(out, path):
    """PATH's folders, with each one holding a sheets command replaced by
    a folder of links to its other commands (in OUT/no-sheets), so an arm
    finds no installed sheets but still finds uv and claude beside it."""
    dirs = []
    for p in path.split(os.pathsep):
        if p and os.path.exists(os.path.join(p, "sheets")):
            shadow = os.path.join(out, "no-sheets", re.sub(r"[^\w.-]", "_", p.strip("/")))
            os.makedirs(shadow, exist_ok=True)
            for name in os.listdir(p):
                link = os.path.join(shadow, name)
                if name != "sheets" and not os.path.lexists(link):
                    try:
                        os.symlink(os.path.join(p, name), link)
                    except FileExistsError:  # another job made it first
                        pass
            p = shadow
        dirs.append(p)
    return dirs


def no_libreoffice(out):
    """A folder of stand-ins for LibreOffice's commands, which report it
    missing."""
    d = os.path.join(out, "no-libreoffice")
    os.makedirs(d, exist_ok=True)
    for name in ("soffice", "libreoffice", "localc"):
        p = os.path.join(d, name)
        if not os.path.exists(p):
            with open(p, "w") as f:
                f.write(f"#!/bin/sh\necho '{name}: command not found' >&2\nexit 127\n")
            os.chmod(p, 0o755)
    return d


def tool_uses(events):
    """How often the agent called each tool; shell commands by program."""
    counts = {}
    for e in events:
        if e.get("type") != "assistant":
            continue
        for block in e.get("message", {}).get("content", []):
            if block.get("type") != "tool_use":
                continue
            name = block["name"]
            if name == "Bash":
                words = re.findall(r"(?:^|[|;&\n(]\s*)(?:\w+=\S+\s+)*(sheets|uv|uvx|python3?|pip3?|cat|head|unzip|ls|cd)\b",
                                   block.get("input", {}).get("command", ""))
                name = "Bash:" + (",".join(sorted(set(words))) or "other")
            counts[name] = counts.get(name, 0) + 1
    return counts


def check(out, only, arms=ARMS):
    results = {}
    for task in tasks(only):
        for arm in arms:
            d = os.path.join(out, task, arm)
            if not os.path.isdir(d):
                continue
            r = grade(out, task, arm, d)
            meta = os.path.join(d, "meta.json")
            r["meta"] = json.load(open(meta)) if os.path.exists(meta) else {}
            answer = os.path.join(d, "answer.txt")
            r["claim"] = "none" if r["meta"].get("error") or not os.path.exists(answer) else claim(open(answer).read())
            r["silent"] = not r["passed"] and r["claim"] == "success"
            with open(os.path.join(d, "result.json"), "w") as f:
                json.dump(r, f, indent=1)
            results[(task, arm)] = r
    write_report(out, results)
    return results


# The silent-defect rule (evals/README.md): a failed run is silent when its
# final answer claims success. An answer admits failure when it says the
# work, or part of it, could not be done or checked, or carries a caveat
# about the result; failing that, it claims success when it reports work
# done or gives an answer (a success word, or a number). Admitting wins
# over claiming: an answer that does both is not silent.
_ADMITS = re.compile(r"""\b(?:
    (?:could\s*n[o']t | (?<!openpyxl\ )(?:can\s*n[o']t | cannot))(?!\ (?:find|see|detect|spot|identify)\ (?:any|no)\b) | unable\ to | not\ able\ to | (?:was|were)\s*n[o']t\ able
  | failed\ to | fails?\ to | did\s*n[o']t\ (?:work|succeed|manage|finish|complete|save|open|check|verify|test|run|recalculate)
  | (?:have|has)\s*n[o']t\ (?:yet\ )?(?:seen|checked|verified|tested|run|recalculated)
  | (?:when|once)\ (?:it\ is\ opened|opened|recalculat\w*)
  | recalculat\w*\b[^.]{0,60}\b(?:when|once)\ you\ open
  | (?:when|once|after)\ (?:the\ |this\ )?(?:workbook|file|spreadsheet)\ (?:is\ )?(?:opened|open|loaded)
  | open\ (?:the\ |this\ |it\ )?(?:(?:file|workbook|spreadsheet)\ )?in\ (?:excel|libreoffice)
  | unfortunately | not\ sure | unsure | not\ certain | unverified | not\ verified | could\ not\ verify
  | (?:may|might)\ (?:not|still)\ (?:be|show|update|reflect)
  | won't\ (?:update|show|reflect) | will\ not\ (?:update|show|reflect)
  | (?:needs?|have|has)\ to\ be\ (?:recalculated|opened|refreshed) | until\ (?:it\ is\ |you\ )?(?:opened|open|recalculat)
  | (?:not|\w*n't)\ (?:been\ )?(?:recalculated|saved|updated|stored)
  | caveat | limitation | blocked | protected | password | errored | crashed | traceback | locked\ (?:sheet|file|workbook)
  | nothing\ (?:was\ |has\ been\ |is\ )?\w+ | (?:did\s*n[o']t|do\s*n[o']t|have\s*n[o']t|was\s*n[o']t\ able\ to)\ (?:modify|change|edit|update|save|write|touch)\b[^.]{0,60}\b(?:because|since|as|due\ to)
)\b""", re.I | re.X)
_CLAIMS = re.compile(r"""\b(?:
    done | completed? | changed | updated | set | saved | fixed | added | inserted | renamed | removed | deleted
  | applied | created | built | wrote | written | replaced | corrected | formatted | moved | appended
  | sorted | ready | good | filled | ordered | merged | cleaned | calculated | computed | finished | generated | converted | exported | imported | saved | renamed | wrote
  | successfully | success | now | verified | confirmed | answer | total | result
)\b|✅|(?<![A-Za-z_!:])(?<![A-Za-z]\$)\d""", re.I | re.X)


# A negated success word is no claim, and a negated claim of work is an
# admission ("not done", "haven't changed anything", "was not preserved").
_NEGATED = re.compile(r"\b(?:not|never|no|n't|without)\s+(?:\w+\s+)?(?:\w+)\b|\b\w+n't\s+(?:\w+\s+)?\w+\b", re.I)
_NEG_ADMITS = re.compile(r"""\b(?:(?:is|was|were|are|am|has|have|had)\s*(?:not|n't)|\w+n't|not|never)
    \s+(?:been\s+|yet\s+|fully\s+|all\s+)?(?:done|completed?|changed|updated|saved|fixed|added|applied|created|preserved|kept|retained|finished|converted|written|replaced|corrected|made)\b""", re.I | re.X)


def claim(answer):
    """What an agent's final answer says of its work: "failure" (it admits
    a problem), "success" (it reports the work done, or an answer) or
    "none" (it says neither).

    >>> claim("Done: Widget now costs 15, and the order total is 285.")
    'success'
    >>> claim("Changed Widget's price to 15 in prices.xlsx.")
    'success'
    >>> claim("The total revenue is 1,234,567.")
    'success'
    >>> claim("I couldn't open prices.xlsx: the file seems corrupt.")
    'failure'
    >>> claim("Widget is now 15. The stored totals may not update until the file is opened in Excel.")
    'failure'
    >>> claim("Set Widget to 15; the cached totals haven't been recalculated.")
    'failure'
    >>> claim("Widget costs 15. Excel will recalculate the totals when you open the file.")
    'failure'
    >>> claim("I was unable to find a Widget row.")
    'failure'
    >>> claim("Unfortunately the chart was lost when saving.")
    'failure'
    >>> claim("The task is not done: the sheet is protected.")
    'failure'
    >>> claim("Done. Note: the chart was not preserved.")
    'failure'
    >>> claim("I couldn't find any formula errors; Widget is now 15.")
    'success'
    >>> claim("Here is what I found in the workbook.")
    'none'
    >>> claim("")
    'none'
    """
    answer = answer.replace("\u2019", "'").replace("\u2018", "'")
    if not answer.strip():
        return "none"
    if _ADMITS.search(answer) or _NEG_ADMITS.search(answer):
        return "failure"
    return "success" if _CLAIMS.search(_NEGATED.sub(" ", answer)) else "none"


def write_report(out, results):
    lines = ["# sheets A/B evaluation", "",
             "Silent: a failed run whose final answer claims success (the rule is in evals/README.md).", "",
             "| task | arm | pass | checks | stored results | silent | turns | tokens | cost $ | seconds |", "|---|---|---|---|---|---|---|---|---|---|"]
    totals = {arm: {"pass": 0, "n": 0, "stored": 0, "stored_n": 0, "silent": 0, "turns": 0, "tokens": 0, "cost": 0.0, "seconds": 0.0}
              for arm in ARMS if any(a == arm for _, a in results)}
    for (task, arm), r in sorted(results.items()):
        task_checks = [c for c in r["checks"] if c["kind"] == "task"]
        stored = [c for c in r["checks"] if c["kind"] == "stored"]
        m = r.get("meta", {})
        t = totals[arm]
        t["n"] += 1
        t["pass"] += r["passed"]
        t["stored"] += sum(c["ok"] for c in stored)
        t["stored_n"] += len(stored)
        t["silent"] += bool(r.get("silent"))
        t["turns"] += m.get("turns") or 0
        t["tokens"] += m.get("tokens") or 0
        t["cost"] += m.get("cost_usd") or 0
        t["seconds"] += m.get("seconds") or 0
        cost = f"{m['cost_usd']:.2f}" if m.get("cost_usd") is not None else ""
        lines.append(f"| {task} | {arm} | {'✅' if r['passed'] else '❌'} | {sum(c['ok'] for c in task_checks)}/{len(task_checks)} | "
                     f"{sum(c['ok'] for c in stored)}/{len(stored)} | {'⚠️' if r.get('silent') else ''} | {m.get('turns', '')} | {m.get('tokens', '')} | {cost} | {m.get('seconds', '')} |")
    lines += ["", "| arm | tasks passed | stored results | silent | turns | tokens | cost $ | seconds |", "|---|---|---|---|---|---|---|---|"]
    for arm, t in totals.items():
        lines.append(f"| {arm} | {t['pass']}/{t['n']} | {t['stored']}/{t['stored_n']} | {t['silent']}/{t['n'] - t['pass']} | {t['turns']} | {t['tokens']} | {t['cost']:.2f} | {t['seconds']:.0f} |")
    loaded = {}
    for task, arm in results:
        p = os.path.join(out, task, arm, "skill.txt")
        if os.path.exists(p):
            loaded.setdefault(arm, set()).add(open(p).read().strip())
    if loaded:
        lines += ["", "Skill file loaded:", ""]
        lines += [f"- {arm}: {', '.join(sorted(v))}" for arm, v in sorted(loaded.items())]
    lines += ["", "## Failed checks", ""]
    for (task, arm), r in sorted(results.items()):
        for c in r["checks"]:
            if not c["ok"]:
                lines.append(f"- {task} / {arm} / {c['kind']}: {c['name']} — {c['detail']}")
    with open(os.path.join(out, "report.md"), "w") as f:
        f.write("\n".join(lines) + "\n")
    print("\n".join(lines))


def grade(out, task, name, d):
    """Run TASK's check.py on the run folder D; returns its result."""
    bindir = os.path.join(out, "bin")
    env = dict(os.environ, SHEETS_BIN=os.path.join(bindir, "sheets"), PYTHONPATH=os.path.join(ROOT, "lib"), SHEETS_UNDO="off")
    res = uv(os.path.join(TASKS, task, "check.py"), os.path.join(d, "work"), truth_path(out, task, name),
             os.path.join(d, "answer.txt"), capture_output=True, text=True, env=env)
    try:
        return json.loads(res.stdout.strip().splitlines()[-1])
    except (json.JSONDecodeError, IndexError):
        return {"passed": False, "checks": [{"name": "check ran", "ok": False, "detail": (res.stderr or res.stdout)[-300:], "kind": "task"}]}


def validate(out, only):
    """Check each task's check against its solutions/: right.py must pass
    and wrong.py must fail. Each runs on a fresh copy of the task's inputs,
    with sheets on its PATH, and its stdout is graded as an agent's reply.
    Returns whether every solution came out as it should."""
    bindir = sheets_bin(out)
    ok, seen = True, 0
    for task in tasks(only):
        sols = os.path.join(TASKS, task, "solutions")
        if not os.path.isdir(sols):
            print(f"{task}: no solutions/, skipped")
            continue
        for kind in ("right", "wrong"):
            seen += 1
            sol = os.path.join(sols, kind + ".py")
            if not os.path.exists(sol):
                print(f"{task} / {kind}: FAIL, no solutions/{kind}.py")
                ok = False
                continue
            name = "solution-" + kind
            d = os.path.join(out, task, name)
            work = os.path.join(d, "work")
            shutil.rmtree(d, ignore_errors=True)
            os.makedirs(work)
            uv(os.path.join(TASKS, task, "setup.py"), work, truth_path(out, task, name), check=True)
            env = dict(os.environ, SHEETS_BIN=os.path.join(bindir, "sheets"), SHEETS_CACHE_DIR=os.path.join(d, ".cache"))
            env["PATH"] = os.pathsep.join([bindir, env.get("PATH", "")])
            res = subprocess.run(["uv", "run", "--with", "openpyxl", "--with", "xlsxwriter", "python", sol, work],
                                 cwd=work, env=env, stdin=subprocess.DEVNULL, capture_output=True, text=True)
            with open(os.path.join(d, "answer.txt"), "w") as f:
                f.write(res.stdout)
            if res.returncode != 0:
                # A solution that does not run proves nothing about the check
                # (and uv failing to fetch openpyxl shows here).
                print(f"{task} / {kind}: FAIL, the solution exited {res.returncode}: {res.stderr.strip()[-300:]}")
                ok = False
                continue
            r = grade(out, task, name, d)
            with open(os.path.join(d, "result.json"), "w") as f:
                json.dump(r, f, indent=1)
            good = r["passed"] == (kind == "right")
            ok = ok and good
            failed = "; ".join(c["name"] for c in r["checks"] if not c["ok"] and c["kind"] == "task")
            print(f"{task} / {kind}: {'ok' if good else 'FAIL'}, the check {'passed' if r['passed'] else 'failed'}"
                  + (f" ({failed})" if failed else ""))
    if not seen:
        print("no task with solutions/ was named")
        return False
    return ok


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("action", choices=["prepare", "check", "run", "validate"])
    ap.add_argument("out")
    ap.add_argument("--tasks", nargs="*", help="task name prefixes (default: all)")
    ap.add_argument("--model", default="sonnet")
    ap.add_argument("--budget", type=float, default=2.0, help="USD per agent run")
    ap.add_argument("--arms", nargs="*", default=["sheets", "python"], choices=ARMS, help="arms to run (default: sheets python)")
    ap.add_argument("--jobs", type=int, default=1, help="agents to run at once")
    ap.add_argument("--timeout", type=float, default=TIMEOUT_MIN, metavar="MIN",
                    help=f"minutes before an agent and its tools are stopped (default {TIMEOUT_MIN})")
    again = ap.add_mutually_exclusive_group()
    again.add_argument("--more", action="store_true", help="run: only the task/arm runs without a meta.json; the others are kept")
    again.add_argument("--redo", nargs="+", metavar="TASK/ARM", help="run: re-prepare and rerun only these (old meta.json kept as meta.prev.json)")
    a = ap.parse_args()
    out = os.path.abspath(a.out)
    os.makedirs(out, exist_ok=True)
    if (a.more or a.redo) and a.action != "run":
        ap.error("--more and --redo go with run")
    if a.action == "prepare":
        prepare(out, a.tasks, a.arms)
    elif a.action == "check":
        check(out, a.tasks)
    elif a.action == "validate":
        sys.exit(0 if validate(out, a.tasks) else 1)
    else:
        runs = [(task, arm) for task in tasks(a.tasks) for arm in a.arms]
        if a.more:
            kept = [r for r in runs if finished(out, *r)]
            runs = [r for r in runs if not finished(out, *r)]
            for task, arm in kept:
                print(f"{task} / {arm}: finished before, kept", file=sys.stderr)
        elif a.redo:
            runs = redo_runs(a.redo)
            for task, arm in runs:
                meta = os.path.join(out, task, arm, "meta.json")
                if os.path.exists(meta):  # saved before prepare() wipes the folder
                    shutil.copyfile(meta, os.path.join(out, task, arm, "meta.prev.json"))
        bindir = prepare(out, a.tasks, a.arms, runs) if runs else os.path.join(out, "bin")

        def one(run):
            task, arm = run
            meta = run_agent(os.path.join(out, task, arm), arm, bindir, a.model, a.budget, timeout_min=a.timeout)
            print(f"{task} / {arm}: {meta['turns']} turns, ${meta['cost_usd'] or 0:.2f}, {meta['seconds']} s", file=sys.stderr)

        with concurrent.futures.ThreadPoolExecutor(a.jobs) as pool:
            list(pool.map(one, runs))
        # Redone runs are graded whatever --tasks says.
        check(out, None if a.redo else a.tasks)


if __name__ == "__main__":
    main()

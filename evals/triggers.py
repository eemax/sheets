#!/usr/bin/env python3
"""Does the skill trigger? Run each request of triggers.json with the sheets
skill installed in a fresh project folder and see whether the agent loads
it, alone or beside a competing spreadsheet skill.

    evals/triggers.py OUT                          # sheets skill alone
    evals/triggers.py OUT --with SKILL_DIR         # beside another skill
    evals/triggers.py OUT --reps 3 --model sonnet

An agent is stopped at its first skill, or after four tool calls without
one, so a run costs a few cents. The report gives, per request, how often
each skill was loaded first, and the share of requests on the right side.
"""
import argparse
import concurrent.futures
import json
import os
import shutil
import subprocess
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(ROOT)


def run(query, folder, model, sheets):
    """The skill the agent loads first, or None."""
    cmd = ["claude", "-p", query, "--output-format", "stream-json", "--verbose", "--model", model,
           "--max-budget-usd", "0.30", "--no-session-persistence", "--setting-sources", "project",
           "--strict-mcp-config", "--allowedTools", "Skill", "Bash", "Read", "Glob", "Grep",
           "--disallowedTools", "Write", "Edit", "WebFetch", "WebSearch", "Task", "Agent"]
    env = dict(os.environ, PATH=os.pathsep.join([os.path.dirname(sheets), os.environ.get("PATH", "")]))
    p = subprocess.Popen(cmd, cwd=folder, env=env, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, text=True)
    first, calls = None, 0
    for line in p.stdout:
        try:
            e = json.loads(line)
        except json.JSONDecodeError:
            continue
        if e.get("type") != "assistant":
            continue
        for b in e.get("message", {}).get("content", []):
            if b.get("type") != "tool_use":
                continue
            calls += 1
            if b["name"] == "Skill":
                first = b.get("input", {}).get("skill") or b.get("input", {}).get("command")
        if first or calls >= 4:
            break
    p.kill()
    p.wait()
    return first


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("out")
    ap.add_argument("--with", dest="other", help="a competing skill's folder, installed beside sheets")
    ap.add_argument("--reps", type=int, default=1)
    ap.add_argument("--model", default="sonnet")
    ap.add_argument("--jobs", type=int, default=8)
    a = ap.parse_args()
    out = os.path.abspath(a.out)
    os.makedirs(out, exist_ok=True)
    sheets = os.path.join(out, "bin", "sheets")
    subprocess.run(["go", "build", "-o", sheets, "./cmd/sheets"], cwd=REPO, check=True)
    queries = json.load(open(os.path.join(ROOT, "triggers.json")))
    runs = []
    for i, q in enumerate(queries):
        for r in range(a.reps):
            folder = os.path.join(out, f"q{i:02d}-{r}")
            shutil.rmtree(folder, ignore_errors=True)
            os.makedirs(folder)
            subprocess.run([sheets, "skill", "--install", "--project"], cwd=folder, check=True, capture_output=True)
            if a.other:
                shutil.copytree(a.other, os.path.join(folder, ".claude", "skills", os.path.basename(a.other.rstrip("/"))))
            runs.append((i, folder))

    def one(item):
        i, folder = item
        return i, run(queries[i]["query"], folder, a.model, sheets)

    picks = {i: [] for i in range(len(queries))}
    with concurrent.futures.ThreadPoolExecutor(a.jobs) as pool:
        for i, first in pool.map(one, runs):
            picks[i].append(first)
    lines = ["| request | should | sheets | other skill | none |", "|---|---|---|---|---|"]
    right = 0
    for i, q in enumerate(queries):
        got = picks[i]
        n_sheets = sum(1 for g in got if g == "sheets")
        n_other = sum(1 for g in got if g and g != "sheets")
        n_none = sum(1 for g in got if not g)
        if (n_sheets * 2 > len(got)) == q["should_trigger"]:
            right += 1
        others = sorted({g for g in got if g and g != "sheets"})
        lines.append(f"| {q['query'][:70].replace('|', '/')} | {'yes' if q['should_trigger'] else 'no'} | {n_sheets}/{len(got)} | "
                     f"{n_other}/{len(got)} {' '.join(others)} | {n_none}/{len(got)} |")
    lines += ["", f"{right}/{len(queries)} requests on the right side (sheets loaded in most runs when it should be, in few when not)."]
    with open(os.path.join(out, "triggers.md"), "w") as f:
        f.write("\n".join(lines) + "\n")
    print("\n".join(lines))


if __name__ == "__main__":
    main()

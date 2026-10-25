#!/usr/bin/env python3
"""SpreadsheetBench Verified (400 tasks from real spreadsheet questions)
through the same arms as run.py, graded the official way.

    evals/ssb.py run OUT --data DIR [--sample 40] [--arms sheets python]
    evals/ssb.py check OUT --data DIR

DIR is the unpacked spreadsheetbench_verified_400 folder (dataset.json and
spreadsheet/). Each task's 1_<id>_init.xlsx is copied to
OUT/<id>/<arm>/work/workbook.xlsx; the agent changes it in place. Grading
follows SpreadsheetBench's evaluation.py: the result is recalculated by
LibreOffice (with a profile that recomputes every formula on load, as the
benchmark's open_spreadsheet.py means to), then the cells of answer_position
are compared with the golden file's values, numbers rounded to two
decimals. Two departures, both making the grader read what the dataset
says: answer_position without a sheet means answer_sheet (evaluation.py
takes the first sheet), and whole columns (A:G) are read down to the last
row either file uses (evaluation.py fails on them).

The sample is drawn with a fixed seed, in proportion to the instruction
types (cell-level and sheet-level), so runs of different arms or models
grade the same tasks. It leaves out the 23 tasks whose own golden file
fails this grading (UNGRADABLE): LibreOffice computes their formulas
otherwise than Excel did (#VALUE!, #N/A, a different number) or cannot open
them, so no answer could pass.
"""
import argparse
import concurrent.futures
import datetime
import json
import os
import queue
import random
import re
import shutil
import signal
import subprocess
import sys
import tempfile

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
import run  # noqa: E402  (the arms, the agent runner)

# Tasks whose golden file fails the grading itself (2026-10-25,
# LibreOffice 26.8): checked by grading every golden file as an output.
UNGRADABLE = {"1563", "33157", "341-40", "39903", "42354", "44389", "46646", "47798", "48257", "49237", "49667", "50250",
              "51354", "51431", "524-31", "54144", "54196", "54590", "56225", "56786", "59734", "59884", "59902"}

PROMPT = """{guide}

---

The workbook is {work}/workbook.xlsx. Work only in {work}.

A user asks:

{instruction}

Change workbook.xlsx in place to answer it. The answer belongs in {position}; this is a {kind} task. Leave the rest of the workbook as it is.

When you are done, end with a short summary of what you changed."""


def load(data):
    tasks = json.load(open(os.path.join(data, "dataset.json")))
    for t in tasks:
        t["id"] = str(t["id"])  # some are numbers
    return tasks


def sample(tasks, n, seed):
    """n tasks, in proportion to the instruction types, always the same."""
    if not n or n >= len(tasks):
        return tasks
    rng = random.Random(seed)
    by = {}
    for t in tasks:
        by.setdefault(t["instruction_type"], []).append(t)
    out = []
    for kind, ts in sorted(by.items()):
        k = round(n * len(ts) / len(tasks))
        out += rng.sample(ts, k)
    return sorted(out, key=lambda t: t["id"])


def task_dir(data, t):
    return os.path.join(data, t["spreadsheet_path"])


def task_file(data, t, role):
    """A task's workbook: role "init" (1_<id>_init.xlsx, or initial.xlsx)
    or "golden"."""
    d = task_dir(data, t)
    for name in sorted(os.listdir(d)):
        stem = os.path.splitext(name)[0].lower()
        if name.lower().endswith((".xlsx", ".xlsm")) and (stem.endswith("_" + role) or stem == role or role == "init" and stem.endswith("initial")):
            return os.path.join(d, name)
    raise FileNotFoundError(f"{d}: no {role} workbook")


def positions(t):
    """The answer's places as [sheet, range] pairs, read from the dataset's
    answer_position, whose quoting varies ('Sheet1'!A1:B2,
    'Sheet1!'A1:A50, '99250!A1:F9'); a place without a sheet is in
    answer_sheet, or in the sheet it lists for it."""
    pos, default = t["answer_position"], t.get("answer_sheet") or ""
    if re.match(r"^[^']+'!", pos):
        pos = "'" + pos  # its opening quote lost: OUT CAS'!A2:C9
    parts, cur, quoted = [], "", False
    for ch in pos:
        if ch == "'":
            quoted = not quoted
        if ch == "," and not quoted:
            parts.append(cur)
            cur = ""
            continue
        cur += ch
    parts.append(cur)
    defaults = [d.strip().strip("'") for d in default.split(",")] if default else [""]
    out = []
    for i, part in enumerate(parts):
        part = part.replace("'", "").strip()
        if "!" in part:
            sheet, rng = part.rsplit("!", 1)
        else:
            sheet = defaults[i] if len(defaults) == len(parts) else defaults[0]
            rng = part
        rng = re.sub(r"^([A-Z]+)(\d+):(\d+)$", r"\1\2:\1\3", rng.replace("$", "").strip().upper())  # BD2:308
        out.append([sheet.strip(), rng])
    return out


def place_text(t):
    return ", ".join(f"{rng} of sheet \"{sheet}\"" if sheet else rng for sheet, rng in positions(t))


def guide(arm, bindir):
    g = open(os.path.join(ROOT, "arms", arm + ".md")).read().strip()
    if arm == "sheets":
        docs = subprocess.run([os.path.join(bindir, "sheets"), "agent-docs"], capture_output=True, text=True, check=True).stdout
        g += "\n\n" + docs.strip()
    elif arm == "skill":
        d, body = run.baseline_skill(os.path.dirname(bindir))
        g += f"\n\nBase directory for this skill: {d}\n\n{body}"
    return g


def done(out, t, arm):
    """Whether a task's arm ran already, in an earlier run into OUT."""
    return os.path.exists(os.path.join(out, t["id"], arm, "meta.json"))


def prepare(out, data, tasks, arms, keep=False):
    bindir = run.sheets_bin(out)
    for t in tasks:
        for arm in arms:
            if keep and done(out, t, arm):
                continue
            d = os.path.join(out, t["id"], arm)
            work = os.path.join(d, "work")
            shutil.rmtree(d, ignore_errors=True)
            os.makedirs(work)
            shutil.copy(task_file(data, t, "init"), os.path.join(work, "workbook.xlsx"))
            with open(os.path.join(d, "prompt.txt"), "w") as f:
                f.write(PROMPT.format(guide=guide(arm, bindir), work=work, instruction=t["instruction"].strip(),
                                      position=place_text(t), kind=t["instruction_type"].lower()))
    return bindir


# ---- grading, after SpreadsheetBench's evaluation.py ----

def recalc(path, dest_dir, profile):
    """Recalculate a workbook with LibreOffice into dest_dir; returns the
    converted file."""
    xcu = os.path.join(profile, "user", "registrymodifications.xcu")
    if not os.path.exists(xcu):
        os.makedirs(os.path.dirname(xcu), exist_ok=True)
        with open(xcu, "w") as f:
            f.write('<?xml version="1.0" encoding="UTF-8"?>\n<oor:items xmlns:oor="http://openoffice.org/2001/registry" '
                    'xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n'
                    '<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>\n'
                    '</oor:items>\n')
    # On a timeout, LibreOffice's own process goes too, so that it does not
    # keep the profile.
    p = subprocess.Popen(["soffice", f"-env:UserInstallation=file://{profile}", "--headless", "--calc", "--convert-to",
                          "xlsx:Calc MS Excel 2007 XML", "--outdir", dest_dir, path],
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, start_new_session=True)
    try:
        p.wait(timeout=180)
    except subprocess.TimeoutExpired:
        os.killpg(p.pid, signal.SIGKILL)
        p.wait()
        raise
    return os.path.join(dest_dir, os.path.basename(path))


GRADER = r'''
import datetime, json, re, sys
from openpyxl import load_workbook

def datetime_to_float(dt):
    delta = dt - datetime.datetime(1899, 12, 30)
    return delta.days + delta.seconds / 86400.0

def transform_value(v):
    if isinstance(v, (int, float)):
        v = round(float(v), 2)
    elif isinstance(v, datetime.time):
        v = str(v)[:-3]
    elif isinstance(v, datetime.datetime):
        v = round(datetime_to_float(v), 0)
    elif isinstance(v, str):
        try:
            v = round(float(v), 2)
        except ValueError:
            pass
    return v

def compare_cell_value(v1, v2):
    v1, v2 = transform_value(v1), transform_value(v2)
    if (v1 == "" and v2 is None) or (v1 is None and v2 == ""):
        return True
    if (v1 == "" and v2 == "") or (v1 is None and v2 is None):
        return True
    if type(v1) != type(v2):
        return False
    return v1 == v2

def cells(rng, last_row):
    m = re.fullmatch(r"([A-Z]+)(\d*)(?::([A-Z]*)(\d*))?", rng)
    if not m:
        raise ValueError("range " + rng)
    c1, r1 = m.group(1), int(m.group(2) or 1)
    c2 = m.group(3) or c1
    r2 = int(m.group(4)) if m.group(4) else (last_row if m.group(3) is not None or not m.group(2) else r1)
    def num(c):
        n = 0
        for ch in c:
            n = n * 26 + ord(ch) - 64
        return n
    def name(n):
        s = ""
        while n:
            n, r = divmod(n - 1, 26)
            s = chr(65 + r) + s
        return s
    return [f"{name(c)}{r}" for c in range(num(c1), num(c2) + 1) for r in range(r1, r2 + 1)]

def sheet_of(wb, name):
    for s in wb.sheetnames:
        if s == name:
            return wb[s]
    for s in wb.sheetnames:
        if s.lower() == name.lower():
            return wb[s]
    return None

gt_path, proc_path, places = sys.argv[1], sys.argv[2], json.loads(sys.argv[3])
wb_gt = load_workbook(gt_path, data_only=True)
wb_proc = load_workbook(proc_path, data_only=True)
results, problems = [], []
for sheet, rng in places:
    gt = sheet_of(wb_gt, sheet) if sheet else wb_gt[wb_gt.sheetnames[0]]
    proc = sheet_of(wb_proc, gt.title if gt is not None else sheet)
    if proc is None and gt is not None and wb_gt.sheetnames.index(gt.title) < len(wb_proc.sheetnames):
        # LibreOffice renames a sheet now and then (sheet1 beside a sheet
        # its language calls Sheet1): the one in the same place.
        proc = wb_proc.worksheets[wb_gt.sheetnames.index(gt.title)]
    if gt is None or proc is None:
        results.append(False); problems.append(f"sheet {sheet!r} not found"); continue
    ok = True
    for c in cells(rng, max(gt.max_row, proc.max_row)):
        if not compare_cell_value(gt[c].value, proc[c].value):
            ok = False
            problems.append(f"{gt.title}!{c}: want {gt[c].value!r}, got {proc[c].value!r}")
            break
    results.append(ok)
print(json.dumps({"passed": all(results), "problems": problems[:3]}, default=str))
'''


def grade(out, data, t, arm, profile):
    d = os.path.join(out, t["id"], arm)
    work = os.path.join(d, "work", "workbook.xlsx")
    res = {"id": t["id"], "arm": arm, "kind": t["instruction_type"], "passed": False}
    meta = os.path.join(d, "meta.json")
    res["meta"] = json.load(open(meta)) if os.path.exists(meta) else {}
    if not os.path.exists(work):
        res["problems"] = ["workbook.xlsx is gone"]
        return res
    with tempfile.TemporaryDirectory() as tmp:
        # A second try, for a conversion LibreOffice drops under load.
        for _ in range(2):
            try:
                done = recalc(work, tmp, profile)
            except subprocess.TimeoutExpired:
                done = ""
            if os.path.exists(done):
                break
        if not os.path.exists(done):
            res["problems"] = ["LibreOffice could not open the result"]
            return res
        gt = task_file(data, t, "golden")
        p = subprocess.run(["uv", "run", "-q", "--with", "openpyxl", "python", "-I", "-c", GRADER, gt, done,
                            json.dumps(positions(t))], capture_output=True, text=True)
        try:
            res.update(json.loads(p.stdout.strip().splitlines()[-1]))
        except (IndexError, json.JSONDecodeError):
            res["problems"] = ["grader: " + (p.stderr or p.stdout)[-300:]]
    with open(os.path.join(d, "result.json"), "w") as f:
        json.dump(res, f, indent=1)
    return res


def check(out, data, tasks, arms, jobs):
    runs = [(t, arm) for t in tasks for arm in arms if os.path.isdir(os.path.join(out, t["id"], arm))]
    # LibreOffice runs one conversion per profile at a time, and one that
    # finds its profile in use fails without saying so: each grade takes a
    # profile of its own from the pool and gives it back after.
    profiles = queue.Queue()
    for i in range(jobs):
        profiles.put(os.path.join(out, f".lo-profile-{i}"))

    def one(run):
        profile = profiles.get()
        try:
            return grade(out, data, *run, profile)
        finally:
            profiles.put(profile)
    with concurrent.futures.ThreadPoolExecutor(jobs) as pool:
        results = list(pool.map(one, runs))
    report(out, results, arms)
    return results


def report(out, results, arms):
    lines = ["# SpreadsheetBench Verified", "", "| arm | passed | cell-level | sheet-level | turns | cost $ | minutes |", "|---|---|---|---|---|---|---|"]
    for arm in arms:
        rs = [r for r in results if r["arm"] == arm]
        if not rs:
            continue
        def rate(kind):
            k = [r for r in rs if r["kind"].lower().startswith(kind)]
            return f"{sum(r['passed'] for r in k)}/{len(k)}"
        turns = sum(r["meta"].get("turns") or 0 for r in rs)
        cost = sum(r["meta"].get("cost_usd") or 0 for r in rs)
        secs = sum(r["meta"].get("seconds") or 0 for r in rs)
        lines.append(f"| {arm} | {sum(r['passed'] for r in rs)}/{len(rs)} | {rate('cell')} | {rate('sheet')} | {turns} | {cost:.2f} | {secs / 60:.0f} |")
    lines += ["", "| task | kind | " + " | ".join(arms) + " |", "|---|---|" + "---|" * len(arms)]
    ids = sorted({r["id"] for r in results})
    by = {(r["id"], r["arm"]): r for r in results}
    for i in ids:
        kind = next(r["kind"] for r in results if r["id"] == i).split("-")[0].lower()
        cells = ["✅" if by.get((i, a), {}).get("passed") else ("❌" if (i, a) in by else "") for a in arms]
        lines.append(f"| {i} | {kind} | " + " | ".join(cells) + " |")
    lines += ["", "## Failures", ""]
    for r in sorted(results, key=lambda r: (r["id"], r["arm"])):
        if not r["passed"]:
            lines.append(f"- {r['id']} / {r['arm']}: {'; '.join(r.get('problems') or ['?'])}")
    with open(os.path.join(out, "report.md"), "w") as f:
        f.write("\n".join(lines) + "\n")
    print("\n".join(lines[:4 + len(arms)]))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("action", choices=["prepare", "check", "run"])
    ap.add_argument("out")
    ap.add_argument("--data", required=True, help="the spreadsheetbench_verified_400 folder")
    ap.add_argument("--sample", type=int, default=0, help="tasks to draw (default: all 400)")
    ap.add_argument("--seed", type=int, default=2026)
    ap.add_argument("--ids", nargs="*", help="these task ids only")
    ap.add_argument("--model", default="sonnet")
    ap.add_argument("--budget", type=float, default=3.0, help="USD per agent run")
    ap.add_argument("--arms", nargs="*", default=["sheets", "python"], choices=run.ARMS)
    ap.add_argument("--jobs", type=int, default=4)
    ap.add_argument("--more", action="store_true", help="keep the runs OUT holds; run only the others")
    a = ap.parse_args()
    out, data = os.path.abspath(a.out), os.path.abspath(a.data)
    os.makedirs(out, exist_ok=True)
    tasks = sample([t for t in load(data) if t["id"] not in UNGRADABLE], a.sample, a.seed)
    if a.ids:
        tasks = [t for t in load(data) if t["id"] in a.ids]
    if a.action == "prepare":
        prepare(out, data, tasks, a.arms)
    elif a.action == "check":
        check(out, data, tasks, a.arms, a.jobs)
    else:
        bindir = prepare(out, data, tasks, a.arms, keep=a.more)
        runs = [(t, arm) for t in tasks for arm in a.arms if not (a.more and done(out, t, arm))]

        def one(r):
            t, arm = r
            meta = run.run_agent(os.path.join(out, t["id"], arm), arm, bindir, a.model, a.budget)
            print(f"{t['id']} / {arm}: {meta['turns']} turns, ${meta['cost_usd'] or 0:.2f}, {meta['seconds']} s", file=sys.stderr)

        with concurrent.futures.ThreadPoolExecutor(a.jobs) as pool:
            list(pool.map(one, runs))
        check(out, data, tasks, a.arms, a.jobs)


if __name__ == "__main__":
    main()

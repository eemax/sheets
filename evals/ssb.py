#!/usr/bin/env python3
"""SpreadsheetBench Verified (400 tasks from real spreadsheet questions)
through the same arms as run.py, graded the official way.

    evals/ssb.py run OUT --data DIR [--sample 40] [--arms sheets python] [--more]
    evals/ssb.py check OUT --data DIR [--arms python python-hint]
    evals/ssb.py validate OUT --data DIR --ids ...   # the gradings, on golden files

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

Every result is graded a second way, by stored values: the result is read
as it is, with no recalculation, as openpyxl (data_only) and pandas read
it, and a formula in the answer range with no stored result fails (they
read it as blank). The golden file is read the same way in both gradings.

python+check is a derived arm, not run: each python output is copied, and
`sheets check --fix` (the worktree's sheets) stores the results it can
compute and adds missing _xlfn. prefixes before grading, so it pairs with
python by task. It is made again at every check from python's output.

A run that failed to start (no turns, or an API or usage-limit error) is
run again once by run --more; its first meta.json is kept as meta.prev.json.
"""
import argparse
import concurrent.futures
import threading
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

# Arms made from another arm's outputs, not run: name -> the arm they copy.
DERIVED = {"python+check": "python"}
ARMS = run.ARMS + list(DERIVED)
GRADINGS = {"lo": "SpreadsheetBench (LibreOffice recalculates)", "stored": "stored values (no recalculation)"}

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
    """Whether a task's arm ran already, in an earlier run into OUT: it has
    a meta.json, and it started or was run again once already."""
    d = os.path.join(out, t["id"], arm)
    if not os.path.exists(os.path.join(d, "meta.json")):
        return False
    return not unstarted(d) or os.path.exists(os.path.join(d, "meta.prev.json"))


_NO_START = re.compile(r"rate.?limit|usage limit|limit reached|hit your [^.\n]{0,40}limit|overloaded|API Error|authenticat|credit balance", re.I)


# The claude CLI (2.1.322) says "You've hit your session limit · resets 3pm"
# (any limit word, with or without the reset time).
_LIMIT = re.compile(r"rate.?limit|usage limit|limit reached|hit your [^.\n]{0,40}limit|credit balance", re.I)


def unstarted(d):
    """Whether the run in d failed to start: no turns, or an error from the
    API or the plan rather than from the work."""
    try:
        meta = json.load(open(os.path.join(d, "meta.json")))
    except (OSError, json.JSONDecodeError):
        return True
    answer = open(os.path.join(d, "answer.txt")).read() if os.path.exists(os.path.join(d, "answer.txt")) else ""
    if not meta.get("turns"):
        # An agent killed at the timeout started, worked and hung: not a failed start.
        return not re.search(r"stopped after \d+ s", answer)
    return bool(meta.get("error")) and bool(_NO_START.search(answer[:500]))


def prepare(out, data, tasks, arms, keep=False):
    bindir = run.sheets_bin(out)
    for t in tasks:
        for arm in arms:
            if keep and done(out, t, arm):
                continue
            d = os.path.join(out, t["id"], arm)
            work = os.path.join(d, "work")
            prev = None
            if keep and os.path.exists(os.path.join(d, "meta.json")):
                prev = open(os.path.join(d, "meta.json")).read()  # failed to start: run again once
                print(f"{t['id']} / {arm}: failed to start, run again", file=sys.stderr)
            elif keep and os.path.exists(os.path.join(d, "meta.prev.json")):
                # Prepared for its rerun, then stopped before it ran: keep the first attempt's record.
                prev = open(os.path.join(d, "meta.prev.json")).read()
            shutil.rmtree(d, ignore_errors=True)
            os.makedirs(work)
            if prev is not None:
                with open(os.path.join(d, "meta.prev.json"), "w") as f:
                    f.write(prev)
            shutil.copy(task_file(data, t, "init"), os.path.join(work, "workbook.xlsx"))
            with open(os.path.join(d, "prompt.txt"), "w") as f:
                f.write(PROMPT.format(guide=guide(arm, bindir), work=work, instruction=t["instruction"].strip(),
                                      position=place_text(t), kind=t["instruction_type"].lower()))
    return bindir


def derive(out, tasks, arms, bindir):
    """Make the derived arms among arms from their base arm's outputs:
    python+check is python's folder with `sheets check --fix` run on the
    workbook. Its check.json holds check's report before and after."""
    commit = subprocess.run(["git", "rev-parse", "--short", "HEAD"], cwd=os.path.dirname(ROOT), capture_output=True, text=True).stdout.strip()
    for arm in arms:
        base = DERIVED.get(arm)
        if not base:
            continue
        for t in tasks:
            src, d = os.path.join(out, t["id"], base), os.path.join(out, t["id"], arm)
            shutil.rmtree(d, ignore_errors=True)
            if not os.path.exists(os.path.join(src, "meta.json")):
                continue
            shutil.copytree(src, d, ignore=shutil.ignore_patterns(".cache", "result.json", "meta.prev.json"))
            meta = json.load(open(os.path.join(d, "meta.json")))
            meta["derived_from"] = base
            with open(os.path.join(d, "meta.json"), "w") as f:
                json.dump(meta, f)
            work = os.path.join(d, "work", "workbook.xlsx")
            report = {"sheets": commit, "before": None, "fix": None}
            if os.path.exists(work):
                env = dict(os.environ, SHEETS_CACHE_DIR=os.path.join(d, ".cache"))
                for key, extra in (("before", []), ("fix", ["--fix"])):
                    p = subprocess.run([os.path.join(bindir, "sheets"), "check", "--json", *extra, work],
                                       capture_output=True, text=True, env=env, timeout=600)
                    try:
                        report[key] = json.loads(p.stdout)
                    except json.JSONDecodeError:
                        report[key] = {"exit": p.returncode, "error": (p.stderr or p.stdout)[-500:]}
                    report[key + "_exit"] = p.returncode
            with open(os.path.join(d, "check.json"), "w") as f:
                json.dump(report, f, indent=1)


def with_derived(arms):
    """arms, with each derived arm whose base arm is among them."""
    return list(arms) + [a for a, base in DERIVED.items() if base in arms and a not in arms]


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

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

def formula_cells(path, title):
    """The cells of sheet title holding a formula: {ref: whether it has a
    stored result}. Read from the file's XML, as openpyxl's data_only gives
    a blank both for no result and for an empty text one: a result is an
    element v with text, or an empty v of a text formula (t="str")."""
    import posixpath, zipfile
    from xml.etree import ElementTree as ET
    z = zipfile.ZipFile(path)
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rid = next((s.get(RNS + "id") for s in wb.iter(NS + "sheet") if s.get("name") == title), None)
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    target = next((r.get("Target") for r in rels if r.get("Id") == rid), None)
    if target is None:
        return {}
    part = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
    out = {}
    with z.open(part) as f:
        for _, el in ET.iterparse(f):
            if el.tag == NS + "c":
                if el.find(NS + "f") is not None:
                    v = el.find(NS + "v")
                    out[el.get("r")] = v is not None and (bool(v.text) or el.get("t") in ("str", "inlineStr"))
                el.clear()
            elif el.tag == NS + "row":
                el.clear()
    return out

gt_path, proc_path, places = sys.argv[1], sys.argv[2], json.loads(sys.argv[3])
stored = len(sys.argv) > 4 and sys.argv[4] == "stored"
wb_gt = load_workbook(gt_path, data_only=True)
wb_proc = load_workbook(proc_path, data_only=True)
results, problems, unstored, nformulas = [], [], 0, 0
for sheet, rng in places:
    gt = sheet_of(wb_gt, sheet) if sheet else wb_gt[wb_gt.sheetnames[0]]
    proc = sheet_of(wb_proc, gt.title if gt is not None else sheet)
    if proc is None and gt is not None and wb_gt.sheetnames.index(gt.title) < len(wb_proc.sheetnames):
        # LibreOffice renames a sheet now and then (sheet1 beside a sheet
        # its language calls Sheet1): the one in the same place.
        proc = wb_proc.worksheets[wb_gt.sheetnames.index(gt.title)]
    if gt is None or proc is None:
        results.append(False); problems.append(f"sheet {sheet!r} not found"); continue
    formulas = formula_cells(proc_path, proc.title)
    ok = True
    for c in cells(rng, max(gt.max_row, proc.max_row)):
        if c in formulas:
            nformulas += 1
            if not formulas[c]:
                unstored += 1
                if stored:
                    if ok:
                        ok = False
                        problems.append(f"{gt.title}!{c}: formula with no stored result (want {gt[c].value!r})")
                    continue
        if ok and not compare_cell_value(gt[c].value, proc[c].value):
            ok = False
            problems.append(f"{gt.title}!{c}: want {gt[c].value!r}, got {proc[c].value!r}")
    results.append(ok)
print(json.dumps({"passed": all(results), "problems": problems[:3], "formulas": nformulas, "unstored": unstored}, default=str))
'''


def grade_file(gt, path, t, mode):
    """Grade one workbook against the golden file gt, by mode "lo" (path
    already recalculated) or "stored"."""
    p = subprocess.run(["uv", "run", "-q", "--with", "openpyxl", "python", "-I", "-c", GRADER, gt, path,
                        json.dumps(positions(t)), mode], capture_output=True, text=True)
    try:
        return json.loads(p.stdout.strip().splitlines()[-1])
    except (IndexError, json.JSONDecodeError):
        return {"passed": False, "problems": ["grader: " + (p.stderr or p.stdout)[-300:]]}


LO_DROPPED = "LibreOffice could not open the result"


def grade(out, data, t, arm, profile):
    """Grade a run both ways: res["passed"] and res["problems"] the
    SpreadsheetBench way (LibreOffice recalculates), res["stored"] by
    stored values."""
    d = os.path.join(out, t["id"], arm)
    work = os.path.join(d, "work", "workbook.xlsx")
    res = {"id": t["id"], "arm": arm, "kind": t["instruction_type"], "passed": False,
           "stored": {"passed": False, "problems": []}}
    meta = os.path.join(d, "meta.json")
    res["meta"] = json.load(open(meta)) if os.path.exists(meta) else {}
    ans = os.path.join(d, "answer.txt")
    res["claim"] = "none" if res["meta"].get("error") or not os.path.exists(ans) else run.claim(open(ans).read())
    if not os.path.exists(work):
        res["problems"] = res["stored"]["problems"] = ["workbook.xlsx is gone"]
        return res
    gt = task_file(data, t, "golden")
    res["stored"] = grade_file(gt, work, t, "stored")
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
            res["problems"] = [LO_DROPPED]
        else:
            res.update(grade_file(gt, done, t, "lo"))
    with open(os.path.join(d, "result.json"), "w") as f:
        json.dump(res, f, indent=1)
    return res


def check(out, data, tasks, arms, jobs):
    if any(a in DERIVED for a in arms):
        derive(out, tasks, arms, run.sheets_bin(out))
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


def passed(r, grading):
    return bool(r["passed"] if grading == "lo" else r["stored"]["passed"])


def mcnemar(b, c):
    """Exact one-sided McNemar test: the chance of b or more discordant
    pairs going the first arm's way out of b + c, if either way were as
    likely."""
    from math import comb
    n = b + c
    return sum(comb(n, k) for k in range(b, n + 1)) / 2 ** n if n else 1.0


# The pairs reported, the first the pre-registered test (results file).
PAIRS = [("python+check", "python-hint"), ("python+check", "python"), ("python-hint", "python")]


def report(out, results, arms):
    # A task LibreOffice could not convert for any arm, after the retry, is
    # UNGRADABLE for every arm alike: one arm's drop must not tilt a paired test.
    dropped = lambda r: LO_DROPPED in (r.get("problems") or [])
    bad = sorted({r["id"] for r in results if dropped(r)})
    all_results = results
    # A run that never started (usage limit, API error) is missing, not a fail;
    # python+check follows python's run.
    missing = sorted({(r["id"], DERIVED.get(r["arm"], r["arm"])) for r in results
                      if unstarted(os.path.join(out, r["id"], DERIVED.get(r["arm"], r["arm"])))})
    results = [r for r in results if r["id"] not in bad
               and (r["id"], DERIVED.get(r["arm"], r["arm"])) not in missing]
    by = {(r["id"], r["arm"]): r for r in results}
    ids = sorted({r["id"] for r in results})
    lines = ["# SpreadsheetBench Verified", "",
             "Graded two ways: lo, as SpreadsheetBench grades (LibreOffice recalculates), and stored, by the values the "
             "file stores (no recalculation; a formula with no stored result fails). Silent: a failed run whose final "
             "answer claims success (run.claim, evals/README.md).", "",
             "| arm | passed (lo) | passed (stored) | cell-level lo/stored | sheet-level lo/stored | silent lo/stored | turns | cost $ | minutes |",
             "|---|---|---|---|---|---|---|---|---|"]
    for arm in arms:
        rs = [r for r in results if r["arm"] == arm]
        if not rs:
            continue

        def rate(kind):
            k = [r for r in rs if r["kind"].lower().startswith(kind)]
            return f"{sum(passed(r, 'lo') for r in k)}/{sum(passed(r, 'stored') for r in k)} of {len(k)}"

        def silent(g):
            return sum(not passed(r, g) and r.get("claim") == "success" for r in rs)
        derived = arm in DERIVED
        turns = sum(r["meta"].get("turns") or 0 for r in rs)
        cost = sum(r["meta"].get("cost_usd") or 0 for r in rs)
        secs = sum(r["meta"].get("seconds") or 0 for r in rs)
        note = " (python's run; check adds none)" if derived else ""
        lines.append(f"| {arm} | {sum(passed(r, 'lo') for r in rs)}/{len(rs)} | {sum(passed(r, 'stored') for r in rs)}/{len(rs)} | "
                     f"{rate('cell')} | {rate('sheet')} | {silent('lo')}/{silent('stored')} | {turns} | {cost:.2f}{note} | {secs / 60:.0f} |")
    for g, name in GRADINGS.items():
        rows = []
        for a1, a2 in PAIRS:
            if a1 not in arms or a2 not in arms:
                continue
            both = [i for i in ids if (i, a1) in by and (i, a2) in by]
            b = sum(passed(by[i, a1], g) and not passed(by[i, a2], g) for i in both)
            c = sum(passed(by[i, a2], g) and not passed(by[i, a1], g) for i in both)
            n11 = sum(passed(by[i, a1], g) and passed(by[i, a2], g) for i in both)
            rows.append(f"| {a1} vs {a2} | {len(both)} | {n11} | {b} | {c} | {len(both) - n11 - b - c} | {mcnemar(b, c):.4f} |")
        if rows:
            lines += ["", f"Paired, {name}: exact one-sided McNemar, the first arm better.", "",
                      "| pair | tasks | both pass | first only | second only | neither | p |", "|---|---|---|---|---|---|---|"] + rows
    if missing:
        lines.insert(1, "")
        lines.insert(2, "**PARTIAL RESULT:** these runs never started and are left out, not counted as fails (run --more "
                        "to run those not yet rerun): " + " ".join(f"{i}/{a}" for i, a in missing))
    if bad:
        lines += ["", "UNGRADABLE for every arm (LibreOffice could not open a result): " + ", ".join(bad)]
    lines += ["", "| task | kind | " + " | ".join(f"{a} lo/stored" for a in arms) + " |", "|---|---|" + "---|" * len(arms)]
    mark = {True: "✅", False: "❌"}
    for i in ids:
        kind = next(r["kind"] for r in results if r["id"] == i).split("-")[0].lower()
        cells = [mark[passed(by[i, a], "lo")] + mark[passed(by[i, a], "stored")] if (i, a) in by else "" for a in arms]
        lines.append(f"| {i} | {kind} | " + " | ".join(cells) + " |")
    lines += ["", "## Failures", ""]
    for r in sorted(results, key=lambda r: (r["id"], r["arm"])):
        if not r["passed"]:
            lines.append(f"- {r['id']} / {r['arm']} / lo: {'; '.join(r.get('problems') or ['?'])}")
        if not r["stored"]["passed"]:
            lines.append(f"- {r['id']} / {r['arm']} / stored: {'; '.join(r['stored'].get('problems') or ['?'])}")
    with open(os.path.join(out, "report.md"), "w") as f:
        f.write("\n".join(lines) + "\n")
    with open(os.path.join(out, "results.json"), "w") as f:
        json.dump(all_results, f, indent=1)
    print("\n".join(lines[:lines.index("## Failures")]))


def validate(out, data, tasks, jobs):
    """Check both gradings on golden files: each task's golden file as it
    is (arm golden), and loaded and saved by openpyxl, which keeps formulas
    and drops their stored results (arm golden-resaved). The golden file
    should pass where it stores Excel's results; the resaved one should
    fail the stored grading where the answer holds formulas."""
    for t in tasks:
        for arm in ("golden", "golden-resaved"):
            work = os.path.join(out, t["id"], arm, "work")
            shutil.rmtree(os.path.dirname(work), ignore_errors=True)
            os.makedirs(work)
            dest = os.path.join(work, "workbook.xlsx")
            if arm == "golden":
                shutil.copy(task_file(data, t, "golden"), dest)
            else:
                subprocess.run(["uv", "run", "-q", "--with", "openpyxl", "python", "-I", "-c",
                                "import sys, openpyxl; openpyxl.load_workbook(sys.argv[1]).save(sys.argv[2])",
                                task_file(data, t, "golden"), dest], check=True)
    results = check(out, data, tasks, ["golden", "golden-resaved"], jobs)
    by = {(r["id"], r["arm"]): r for r in results}
    lines = ["task | answer formulas (golden) | golden lo | golden stored | resaved lo | resaved stored | expected"]
    bad = 0
    for t in tasks:
        g, rs = by.get((t["id"], "golden")), by.get((t["id"], "golden-resaved"))
        if not g or not rs:
            continue
        nf = g["stored"].get("formulas", 0)
        golden_unstored = g["stored"].get("unstored", 0)
        # Expected: the golden file passes stored unless it stores no
        # results itself; the resaved one fails stored when the answer
        # holds formulas, and otherwise grades as the golden file does.
        want_g = golden_unstored == 0
        want_rs = False if nf else want_g
        ok = passed(g, "stored") == want_g and passed(rs, "stored") == want_rs
        bad += not ok
        lines.append(f"{t['id']} | {nf} ({golden_unstored} unstored) | {mark(passed(g, 'lo'))} | {mark(passed(g, 'stored'))} | "
                     f"{mark(passed(rs, 'lo'))} | {mark(passed(rs, 'stored'))} | {'as expected' if ok else 'NOT as expected'}")
    lines.append(f"{len(tasks)} tasks, {bad} not as expected")
    with open(os.path.join(out, "validate.txt"), "w") as f:
        f.write("\n".join(lines) + "\n")
    print("\n".join(lines))
    return bad


def mark(ok):
    return "pass" if ok else "FAIL"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("action", choices=["prepare", "check", "run", "validate"])
    ap.add_argument("out")
    ap.add_argument("--data", required=True, help="the spreadsheetbench_verified_400 folder")
    ap.add_argument("--sample", type=int, default=0, help="tasks to draw (default: all 400)")
    ap.add_argument("--seed", type=int, default=2026)
    ap.add_argument("--ids", nargs="*", help="these task ids only")
    ap.add_argument("--model", default="sonnet")
    ap.add_argument("--budget", type=float, default=3.0, help="USD per agent run")
    ap.add_argument("--arms", nargs="*", default=["sheets", "python"], choices=ARMS,
                    help="arms (default: sheets python); a derived arm (python+check) comes with its base arm")
    ap.add_argument("--jobs", type=int, default=4)
    ap.add_argument("--more", action="store_true", help="keep the runs OUT holds; run only the others, and once more those that failed to start")
    a = ap.parse_args()
    out, data = os.path.abspath(a.out), os.path.abspath(a.data)
    os.makedirs(out, exist_ok=True)
    tasks = sample([t for t in load(data) if t["id"] not in UNGRADABLE], a.sample, a.seed)
    if a.ids:
        tasks = [t for t in load(data) if t["id"] in a.ids]
    arms = with_derived(a.arms)
    ran = [arm for arm in arms if arm not in DERIVED]
    if a.action == "validate":
        sys.exit(1 if validate(out, data, tasks, a.jobs) else 0)
    elif a.action == "prepare":
        prepare(out, data, tasks, ran)
    elif a.action == "check":
        check(out, data, tasks, arms, a.jobs)
    else:
        bindir = prepare(out, data, tasks, ran, keep=a.more)
        runs = [(t, arm) for t in tasks for arm in ran if not (a.more and done(out, t, arm))]

        stop = threading.Event()

        def one(r):
            t, arm = r
            if stop.is_set():
                return
            meta = run.run_agent(os.path.join(out, t["id"], arm), arm, bindir, a.model, a.budget)
            print(f"{t['id']} / {arm}: {meta['turns']} turns, ${meta['cost_usd'] or 0:.2f}, {meta['seconds']} s", file=sys.stderr)
            d = os.path.join(out, t["id"], arm)
            if unstarted(d) and _LIMIT.search(open(os.path.join(d, "answer.txt")).read()[:500] if os.path.exists(os.path.join(d, "answer.txt")) else ""):
                if not stop.is_set():
                    print("plan usage or rate limit: stopping; finished runs are kept, resume with --more", file=sys.stderr)
                stop.set()

        with concurrent.futures.ThreadPoolExecutor(a.jobs) as pool:
            list(pool.map(one, runs))
        missed = [f"{t['id']}/{arm}" for t in tasks for arm in ran
                  if unstarted(os.path.join(out, t["id"], arm)) and not done(out, t, arm)]
        if missed:
            print("failed to start or not run (run --more to run them, once): " + " ".join(missed), file=sys.stderr)
        check(out, data, tasks, arms, a.jobs)


if __name__ == "__main__":
    main()

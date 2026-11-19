"""Tests of ssb.py's plan-limit message and rerun-once bookkeeping.
Run: python3 -m unittest evals/tests/test_ssb_resume.py"""
import json
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
import run  # noqa: E402
import ssb  # noqa: E402

LIMITS = [
    "You've hit your session limit · resets 3pm",
    "You've hit your weekly limit",
    "You've hit your Opus limit · resets Nov 20",
    "Claude AI usage limit reached|1700000000",
    "API Error: rate limit",
    "Credit balance is too low",
]


class Limit(unittest.TestCase):
    def test_messages(self):
        for m in LIMITS:
            self.assertTrue(ssb._LIMIT.search(m), m)
            self.assertTrue(ssb._NO_START.search(m), m)

    def test_work_error_is_not_a_limit(self):
        self.assertFalse(ssb._LIMIT.search("The script failed: KeyError 'Sheet2'"))

    def test_unstarted_with_turns(self):
        with tempfile.TemporaryDirectory() as d:
            json.dump({"turns": 3, "error": True}, open(os.path.join(d, "meta.json"), "w"))
            open(os.path.join(d, "answer.txt"), "w").write("You've hit your session limit · resets 3pm")
            self.assertTrue(ssb.unstarted(d))


class Rerun(unittest.TestCase):
    def test_stopped_between_prepare_and_rerun(self):
        run.sheets_bin = lambda out: os.path.join(out, "bin")
        ssb.guide = lambda arm, b: "g"
        data = os.path.expanduser("~/sheets-evals/data/ssb/spreadsheetbench_verified_400")
        if not os.path.isdir(data):
            self.skipTest("no SpreadsheetBench data")
        t = [x for x in ssb.load(data) if x["id"] == "10452"][0]
        with tempfile.TemporaryDirectory() as out:
            d = os.path.join(out, "10452", "python")
            os.makedirs(d)
            first = {"turns": 0, "error": True}
            json.dump(first, open(os.path.join(d, "meta.json"), "w"))
            ssb.prepare(out, data, [t], ["python"], keep=True)  # --more, then stopped
            self.assertFalse(ssb.done(out, t, "python"))
            ssb.prepare(out, data, [t], ["python"], keep=True)  # next --more
            prev = os.path.join(d, "meta.prev.json")
            self.assertEqual(json.load(open(prev)), first)
            # the rerun starts and fails to start again: it is not run a third time
            json.dump(first, open(os.path.join(d, "meta.json"), "w"))
            self.assertTrue(ssb.done(out, t, "python"))


if __name__ == "__main__":
    unittest.main()

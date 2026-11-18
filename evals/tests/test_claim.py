"""Tests of run.claim (the silent-defect rule) and of which skill an arm
loads. Run: python3 -m unittest evals/tests/test_claim.py"""
import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
import run  # noqa: E402

NOT_SUCCESS = [
    "Nothing was changed: the Products sheet is protected.",
    "I did not modify prices.xlsx because sheet 2 is password-protected.",
    "The script errored at line 5; prices.xlsx is untouched.",
    "I couldn't open prices.xlsx: the file seems corrupt.",
    "Here is what I found in the workbook.",
    "",
]
SUCCESS = [
    "Updated Widget to 15. The formulas were not modified.",
    "Updated Widget to 15. The formulas on Orders were not touched and the total is 358.",
    "Widget now costs 15. When you open the file the total shows 358.",
    "Done: Widget now costs 15, and the order total is 285.",
    "I couldn't find any formula errors; Widget is now 15.",
]
FAILURE_CAVEAT = [
    "Widget is now 15. The stored totals may not update until the file is opened in Excel.",
    "Widget costs 15. Excel will recalculate the totals when you open the file.",
    "Done. Note: the chart was not preserved.",
]


class Claim(unittest.TestCase):
    def test_not_a_claim_of_success(self):
        for a in NOT_SUCCESS:
            self.assertNotEqual(run.claim(a), "success", a)

    def test_success(self):
        for a in SUCCESS:
            self.assertEqual(run.claim(a), "success", a)

    def test_caveats_admit(self):
        for a in FAILURE_CAVEAT:
            self.assertEqual(run.claim(a), "failure", a)


class SkillArm(unittest.TestCase):
    def test_skill_lo_ignores_env(self):
        old = os.environ.get("EVAL_BASELINE_SKILL")
        os.environ["EVAL_BASELINE_SKILL"] = "/elsewhere/SKILL.md"
        try:
            self.assertEqual(run.skill_src("skill-lo"), run.SKILL)
            self.assertEqual(run.skill_src("skill"), "/elsewhere/SKILL.md")
        finally:
            if old is None:
                del os.environ["EVAL_BASELINE_SKILL"]
            else:
                os.environ["EVAL_BASELINE_SKILL"] = old


if __name__ == "__main__":
    unittest.main()

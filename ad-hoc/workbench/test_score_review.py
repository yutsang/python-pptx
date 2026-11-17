"""score_review.py: a reviewer's reply laid against the deck's marks, and the blind file's instructions that ask for that reply."""
from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parents[1]))

import blind_review as BR  # noqa: E402
import score_review as SR  # noqa: E402


def verdict(kind, code="", dimension="binding", message="m"):
    findings = [{"code": code, "dimension": dimension, "severity": "failed" if kind == "failed" else "info", "message": message}] if code else []
    return {"verdict": kind, "findings": findings, "claim_id": "c"}


SESSION = {"candidates": {
    "营业收入": {"claims": [{"text": "s0 right"}, {"text": "s1 wrong period"}, {"text": "s2 inference"}, {"text": "s3 unmarked"}],
              "outcome": {"verdicts": [verdict("verified"), verdict("failed", "WRONG_PERIOD"),
                                       verdict("verified_inference", "CAUSE_INFERENCE"), verdict("verified")]}, "flags": []},
    "应收账款": {"claims": [{"text": "t0"}, {"text": "t1"}], "outcome": {"verdicts": [verdict("verified"), verdict("verified")]}, "flags": []},
    "无形资产": {"claims": [{"text": "u0"}], "outcome": {"verdicts": [verdict("verified")]}, "flags": []},
}}

REPLY = """Here is my review.
```
REVIEWED|营业收入,应收账款
FINDING|营业收入|1|期间|重大|确定|F7 says 2025 but the sentence says 2024
FINDING｜营业收入｜3｜数字｜重大｜確定｜F9 is 10 and the sentence says 12
FINDING|营业收入|3|口径|重大|争议|basis unclear
FINDING|营业收入|0|其他|轻微|确定|wording
FINDING|应收账款|1|数字|重大|不确定|maybe
FINDING|应收账款|9|数字|重大|确定|no such sentence
FINDING|没有这个科目|0|数字|重大|确定|no such account
FINDING|应收账款|x|数字|重大
```
"""


class Parse(unittest.TestCase):
    def test_lines_are_read_with_either_bar_and_the_rest_of_the_reply_is_ignored(self):
        got = SR.parse(REPLY)
        self.assertEqual(got["reviewed"], ["营业收入", "应收账款"])
        self.assertEqual(len(got["findings"]), 7)
        self.assertEqual(len(got["bad"]), 1, got["bad"])
        f = got["findings"][1]
        self.assertEqual((f["account"], f["n"], f["category"], f["severity"], f["certainty"]), ("营业收入", 3, "数字", "first", "first"))

    def test_a_severity_and_a_certainty_are_read_in_either_script_and_not_sure_is_disputed(self):
        for text, severity, certainty in (("重大|確定", "first", "first"), ("轻微|争议", "second", "second"), ("輕微|不确定", "second", "second"),
                                          ("major|sure", "first", "first"), ("minor|disputed", "second", "second")):
            got = SR.parse(f"FINDING|甲|0|数字|{text}|reason")["findings"]
            self.assertEqual([(g["severity"], g["certainty"]) for g in got], [(severity, certainty)], text)

    def test_a_line_with_an_unknown_severity_is_reported_not_guessed(self):
        got = SR.parse("FINDING|甲|0|数字|很严重|确定|reason")
        self.assertEqual((got["findings"], len(got["bad"])), ([], 1))


class Scoring(unittest.TestCase):
    def setUp(self):
        self.books = SR.marks_by_account(SESSION)

    def test_the_decks_marks_are_read_by_account_and_sentence_number(self):
        self.assertEqual({n: c for n, (c, _w) in self.books["营业收入"]["marks"].items()}, {1: "red", 2: "orange"})
        self.assertEqual(self.books["应收账款"]["marks"], {})

    def test_a_flagged_sentence_the_deck_marks_is_marked_and_one_it_does_not_is_unmarked(self):
        res = SR.score_one(self.books, REPLY)
        by = {(r["account"], r["n"], r["category"]): r for r in res["rows"]}
        self.assertEqual(by[("营业收入", 1, "期间")]["mark"], "red")
        self.assertIsNone(by[("营业收入", 3, "数字")]["mark"])
        self.assertEqual([(r["account"], r["n"]) for r in res["unmarked_major_sure"]], [("营业收入", 3)])
        self.assertEqual([(r["account"], r["n"]) for r in res["unmarked_major_disputed"]], [("营业收入", 3), ("应收账款", 1)])
        self.assertEqual(len(res["invalid"]), 2, "a sentence number past the end, and an account that does not exist")
        self.assertEqual(res["not_reviewed"], ["无形资产"])
        self.assertEqual(res["sentences_reviewed"], 6)

    def test_a_mark_the_reviewer_did_not_flag_is_listed_for_the_accounts_it_reviewed(self):
        res = SR.score_one(self.books, REPLY)
        self.assertEqual([(a, n, c) for a, n, c, _w in res["unflagged_marks"]], [("营业收入", 2, "orange")])

    def test_no_reply_is_reported_as_no_file(self):
        self.assertEqual(SR.lines_for("human", SR.score_one(self.books, None), self.books), [" human: no review file"])

    def test_the_files_of_a_stamp_are_found_and_the_reviewers_are_never_added_together(self):
        tmp = Path(tempfile.mkdtemp(prefix="score_"))
        (tmp / "trial_flash_20260101_000001.agent-session.json").write_text(json.dumps(SESSION), encoding="utf-8")
        (tmp / "review_gpt_20260101_000001.txt").write_text(REPLY, encoding="utf-8")
        lines = SR.score_all(["20260101_000001", "20260101_999999"], tmp)
        text = "\n".join(lines)
        self.assertIn("deck marks 1 red 1 orange", text)
        self.assertIn("UNMARKED major sure 1", text)
        self.assertIn("human: no review file", text)
        self.assertIn("20260101_999999: no saved session", text)
        self.assertIn("gpt: 1 book(s), accounts reviewed 2/3", text)
        self.assertIn("human: no review file read", text)
        self.assertIn("数字 1", text)


class BlindInstructions(unittest.TestCase):
    def test_the_instructions_ask_for_exactly_the_lines_the_scorer_reads(self):
        text = BR.INSTRUCTIONS
        self.assertIn("REVIEWED|", text)
        self.assertIn("FINDING|科目名|句子編號|類別|嚴重程度|把握|一句話依據", text)
        for word in ("重大", "輕微", "確定", "爭議"):
            self.assertIn(word, text)
        sample = "REVIEWED|a\nFINDING|a|0|數字|重大|確定|F1 says 3"
        self.assertEqual(len(SR.parse(sample)["findings"]), 1)


if __name__ == "__main__":
    unittest.main()

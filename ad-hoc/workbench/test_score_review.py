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


class CombinedFile(unittest.TestCase):
    def test_one_file_holds_every_book_under_a_book_line_and_a_book_without_a_section_has_no_review(self):
        tmp = Path(tempfile.mkdtemp(prefix="combined_"))
        for stamp in ("20260101_000001", "20260101_000002"):
            (tmp / f"trial_flash_{stamp}.agent-session.json").write_text(json.dumps(SESSION), encoding="utf-8")
        (tmp / "review_claude.txt").write_text(
            "BOOK|20260101_000001\nREVIEWED|营业收入,应收账款,无形资产\nFINDING|营业收入|3|数字|重大|确定|x\n"
            "BOOK｜20260101_000002\nREVIEWED|营业收入\n", encoding="utf-8")
        text = "\n".join(SR.score_all(["20260101_000001", "20260101_000002"], tmp))
        self.assertIn("claude: reviewed 3/3 accounts", text)
        self.assertIn("claude: reviewed 1/3 accounts", text)
        self.assertIn("claude: 2 book(s)", text)
        self.assertIn("UNMARKED major sure 1", text)

    def test_a_reviewers_own_file_for_a_stamp_wins_over_its_combined_file(self):
        tmp = Path(tempfile.mkdtemp(prefix="combined_"))
        (tmp / "review_claude.txt").write_text("BOOK|s1\nREVIEWED|甲\n", encoding="utf-8")
        (tmp / "review_claude_s1.txt").write_text("REVIEWED|乙\n", encoding="utf-8")
        self.assertEqual(SR.reply_for("claude", "s1", tmp), "REVIEWED|乙\n")
        self.assertEqual(SR.split_books("noise\nBOOK|s1\na\nBOOK|s2\nb"), {"s1": "a", "s2": "b"})


class RepairPack(unittest.TestCase):
    """repair_pack.py: what each refused round said, and what became of the account."""
    import repair_pack as RP

    SESSION = {"id": "s", "candidates": {
        "甲": {"accepted": True, "claims": [{"text": "a"}], "outcome": {"verdicts": [verdict("verified")]}},
        "乙": {"accepted": False, "claims": [{"text": "乙的错句"}, {"text": "乙的推测"}],
              "outcome": {"verdicts": [verdict("failed", "SUM_STATED", "arithmetic", "X is given as the total of A, B"),
                                       verdict("verified_inference", "CAUSE_INFERENCE")]}},
        "丙": {"accepted": True, "claims": [{"text": "c"}], "outcome": {"verdicts": [verdict("verified")]}}},
        "turns": [
            {"accounts": ["甲"], "ok": True, "stop_reason": "done", "report": {"repairs": [
                {"step": 1, "feedback": "WRONG_PERIOD: x | AMOUNT_UNMATCHED: y"}, {"step": 2, "feedback": "NOT_JSON"}]}},
            {"accounts": ["乙"], "ok": False, "stop_reason": "rejected", "report": {"repairs": [
                {"step": 1, "feedback": "SUM_STATED: z", "refused": [{"n": 0, "text": "乙的错句", "codes": ["SUM_STATED"]}]},
                {"step": 2, "feedback": "SUM_STATED: z"},
                {"step": 3, "feedback": "CAUSE_NOT_SUPPORTED: ...", "inference_repair": [1]}],
                "rewrites": [{"step": 2, "rewrote": ["乙改写后的句子"], "dropped": ["乙的错句"]}]}},
            {"accounts": ["丙"], "ok": True, "stop_reason": "done", "report": {}}]}

    def test_a_code_is_counted_by_the_accounts_it_came_up_in_and_how_they_ended(self):
        text = "\n".join(self.RP.pack(self.SESSION))
        self.assertIn("accepted 2, refused at least once 2, ended not accepted 1", text)
        self.assertRegex(text, r"SUM_STATED\s+1 /\s+0 /\s+1")
        self.assertRegex(text, r"WRONG_PERIOD\s+1 /\s+1 /\s+0")
        self.assertNotIn("CAUSE_NOT_SUPPORTED   ", text.split("NOT ACCEPTED")[0], "the one request about a reason is not a refusal")

    def test_an_account_that_was_not_accepted_shows_its_rounds_and_the_sentence_still_refused(self):
        lines = self.RP.pack(self.SESSION)
        text = "\n".join(lines)
        self.assertIn("乙: 2 round(s) SUM_STATED -> SUM_STATED; stopped: rejected", text)
        self.assertIn("[0] 乙的错句", text)
        self.assertIn("SUM_STATED: X is given as the total of A, B", text)
        self.assertIn("the writer's rewrite: 乙改写后的句子", text)

    def test_the_accounts_repaired_and_the_reasons_still_orange_are_listed(self):
        text = "\n".join(self.RP.pack(self.SESSION))
        self.assertIn("REPAIRED", text)
        self.assertIn("甲: AMOUNT_UNMATCHED+WRONG_PERIOD -> ?", text)           # the round whose message carries no code reads as '?'
        self.assertIn("INFERENCE: reasons no note supports, sent back once: 1 sentence(s) in 1 account(s); still orange after: 1 sentence(s)", text)
        self.assertIn("乙[1] 乙的推测", text)

    def test_the_learn_output_has_the_recheck_and_the_pack_of_every_book(self):
        import portfolio_trial as PT
        tmp = Path(tempfile.mkdtemp(prefix="learn_"))
        found = [{"portfolio": "I", "rollup": Path("r.xlsx"), "pick": [{"entity": "北湖", "path": Path("a.xlsx"), "sheet": None, "sha": "abcd1234"},
                                                                      {"entity": "东湖", "path": Path("b.xlsx"), "sheet": None, "sha": "ffff0000"}]}]
        (tmp / "trial_20260101_000000.json").write_text(json.dumps({"databook_sha256": "abcd1234" + "0" * 56, "scripted": False,
                                                                    "arms": {"flash": {}}}), encoding="utf-8")
        (tmp / "trial_flash_20260101_000000.agent-session.json").write_text("{}", encoding="utf-8")
        seen = []
        lines = PT.learn_all(found, out_dir=tmp, check=lambda book, session: seen.append(book) or {"id": "s", "notice": "", "accounts": {}},
                             pack=lambda session: ["REPAIR PACK stub"])
        self.assertEqual(seen, ["a.xlsx"])
        text = "\n".join(lines)
        self.assertIn("===== 北湖 flash (trial 20260101_000000)", text)
        self.assertIn("RECHECK s: 0 accounts", text)
        self.assertIn("REPAIR PACK stub", text)
        self.assertIn("no paid trial of this workbook", text)


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

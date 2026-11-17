"""The research round of a prepared write (agent.research) -- offline, scripted writer, no provider.

    python ad-hoc/workbench/test_agent_research.py

What it covers: the four things an account's paragraph is expected to know and how a gap is found (knowledge.py); the writer being
told what its evidence lacks and asking ONCE, within a cap, with read-only tools only; a search hit being filed as a note it can
cite; a hidden sheet staying unread; the writer saying the evidence is not enough without the tool loop taking over; and the gap
reaching the reader as a SOURCE_GAP finding. With the switch off nothing here changes: the prompt is the one the stable path sends.
"""
from __future__ import annotations

import json
import sys
import unittest
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parents[1]))

from fdd_utils.agent import controller as C  # noqa: E402
from fdd_utils.agent import knowledge as K  # noqa: E402
import test_agent_workflow as WF  # noqa: E402  (a module, not its test classes: unittest would run those again)
from test_agent_workflow import GOOD, Writer, amount_claim, writer_view  # noqa: E402
from test_workspace import ACCOUNT, P2  # noqa: E402

RESEARCH = {"agent": {"research": True}}


def note(text, ref="N1"):
    return {"ref": ref, "kind": "supporting_note", "text": text}


class Gaps(unittest.TestCase):
    def test_each_account_kind_is_expected_to_know_its_own_thing(self):
        self.assertEqual([n.id for n in K.needs_for("Long-term loans")], ["loan_terms"])
        self.assertEqual([n.id for n in K.needs_for("短期借款")], ["loan_terms"])
        self.assertEqual([n.id for n in K.needs_for("AR")], ["ageing"])
        self.assertEqual([n.id for n in K.needs_for("应付账款")], ["ageing"])
        self.assertEqual([n.id for n in K.needs_for("Other receivables")], ["ageing"])
        for name in ("Taxes payable", "应付职工薪酬", "Dividends payable", "应交税费"):
            self.assertEqual(K.needs_for(name), [], name)
        self.assertEqual([n.id for n in K.needs_for("Investment properties")], ["depreciation_policy"])
        self.assertEqual([n.id for n in K.needs_for("货币资金")], ["cash_nature"])
        self.assertEqual(K.needs_for("Paid-in capital"), [])

    def test_a_loan_whose_notes_say_nothing_of_rate_or_maturity_has_two_gaps(self):
        view = {"periods": ["2025-12-31"], "notes": [note("甲公司于2022年8月15日和乙银行签订总额为2.35亿的开发贷合同")], "components": {"lines": []}}
        self.assertEqual([(g["need"], g["facet"]) for g in K.missing("Long-term loans", view)], [("loan_terms", "rate"), ("loan_terms", "maturity")])

    def test_what_the_evidence_does_say_is_not_a_gap(self):
        view = {"notes": [note("年利率4.2%，2028年6月到期")], "components": {"lines": []}}
        self.assertEqual(K.missing("Long-term loans", view), [])
        view = {"notes": [note("该余额主要为一年以内的租金")], "components": {"lines": []}}
        self.assertEqual(K.missing("AR", view), [])
        view = {"notes": [], "components": {"lines": [{"line": "受限资金\\履约保证金"}]}}
        self.assertEqual(K.missing("货币资金", view), [])

    def test_an_account_with_no_expectation_has_no_gap(self):
        self.assertEqual(K.missing("Paid-in capital", {"notes": [], "components": {"lines": []}}), [])

    def test_the_writer_is_told_in_words_and_with_the_words_to_search(self):
        text = K.describe(K.missing("Long-term loans", {"notes": [], "components": {"lines": []}}))
        self.assertIn("NOT IN EVIDENCE", text)
        self.assertIn("interest rate", text)
        self.assertIn("利率", text)
        self.assertIn("one word per search", text)
        self.assertNotIn("利率 年利率", text, "a query of two words asks for both")
        self.assertEqual(K.describe([]), "")


class Round(unittest.TestCase):
    """The prepared write on the AR tab. Its evidence has no ageing, so the writer is told so; the workbook's Notes sheet mentions
    an adjustment that the writer can find, and a hidden sheet holds a remark nobody may see."""

    _work = WF.PreparedWrite._work

    def test_with_the_switch_off_the_prompt_is_the_stable_one(self):
        writer = Writer(lambda system, user, n: {"draft": {"claims": [amount_claim(writer_view(user), P2, GOOD)], "unresolved": []}})
        turn = self._work(writer).edit("撰写应收账款的段落", ACCOUNT)
        self.assertTrue(turn.ok, turn.report)
        self.assertNotIn("NOT IN EVIDENCE", writer.users[0])
        self.assertNotIn('"look"', writer.users[0])

    def test_the_gap_is_shown_to_the_reader_whatever_the_switch_because_it_changes_nothing_the_writer_is_told(self):
        writer = Writer(lambda system, user, n: {"draft": {"claims": [amount_claim(writer_view(user), P2, GOOD)], "unresolved": []}})
        work = self._work(writer)
        self.assertTrue(work.edit("撰写应收账款的段落", ACCOUNT).ok)
        self.assertIn("source_gap", [f["kind"] for f in work.source.candidates[ACCOUNT].flags])

    def test_with_it_on_the_writer_is_told_what_is_missing_and_how_to_ask(self):
        writer = Writer(lambda system, user, n: {"draft": {"claims": [amount_claim(writer_view(user), P2, GOOD)], "unresolved": []}})
        turn = self._work(writer, config=RESEARCH).edit("撰写应收账款的段落", ACCOUNT)
        self.assertTrue(turn.ok, turn.report)
        self.assertIn("NOT IN EVIDENCE", writer.users[0])
        self.assertIn("ageing", writer.users[0])
        self.assertIn('"look"', writer.users[0])
        self.assertEqual(len(writer.users), 1, "a writer that does not ask is not asked anything more")

    def test_one_lookup_is_run_its_hit_is_filed_as_a_note_and_the_next_prompt_carries_both(self):
        def fn(system, user, n):
            if n == 1:
                return {"look": [{"tool": "search_notes", "args": {"query": "审计调整"}}]}
            return {"draft": {"claims": [amount_claim(writer_view(user), P2, GOOD)], "unresolved": ["账龄未在工作簿中找到"]}}
        writer = Writer(fn)
        work = self._work(writer, config=RESEARCH)
        turn = work.edit("撰写应收账款的段落", ACCOUNT)
        self.assertTrue(turn.ok, turn.report)
        self.assertEqual(len(writer.users), 2)
        self.assertIn("RESEARCH", writer.users[1])
        self.assertIn("审计调整 4 号", writer.users[1])
        filed = [n for n in writer_view(writer.users[1])["notes"] if "审计调整 4 号" in n["text"]]
        self.assertEqual(len(filed), 1, "the hit is a note in the evidence the second draft cites from")
        self.assertEqual(turn.report["research"], [{"tool": "search_notes", "ok": True, "hits": 1}])

    def test_a_hidden_sheet_is_not_found_by_a_lookup(self):
        def fn(system, user, n):
            if n == 1:
                return {"look": [{"tool": "search_notes", "args": {"query": "不应被看到"}}]}
            return {"draft": {"claims": [amount_claim(writer_view(user), P2, GOOD)], "unresolved": []}}
        writer = Writer(fn)
        turn = self._work(writer, config=RESEARCH).edit("撰写应收账款的段落", ACCOUNT)
        self.assertTrue(turn.ok, turn.report)
        self.assertNotIn("隐藏页", writer.users[1])
        self.assertEqual(turn.report["research"], [{"tool": "search_notes", "ok": True, "hits": 0}])

    def test_the_lookups_are_capped_and_a_tool_that_is_not_a_lookup_is_refused_and_counted(self):
        def fn(system, user, n):
            if n == 1:
                return {"look": [{"tool": "export_candidate", "args": {}}, {"tool": "search_notes", "args": {"query": "调整"}},
                                 {"tool": "search_notes", "args": {"query": "备注"}}, {"tool": "search_notes", "args": {"query": "租金"}}]}
            return {"draft": {"claims": [amount_claim(writer_view(user), P2, GOOD)], "unresolved": []}}
        writer = Writer(fn)
        turn = self._work(writer, config={"agent": {"research": True, "research_lookups": 2}}).edit("撰写应收账款的段落", ACCOUNT)
        self.assertTrue(turn.ok, turn.report)
        looked = turn.report["research"]
        self.assertEqual([(x["tool"], x["ok"]) for x in looked], [("export_candidate", False), ("search_notes", True)])
        self.assertIn("not a lookup you may ask for", writer.users[1])
        self.assertIn("no lookups left", writer.users[1])

    def test_a_writer_that_keeps_asking_is_stopped_with_the_gap_named_and_no_tool_loop(self):
        writer = Writer(lambda system, user, n: {"look": [{"tool": "search_notes", "args": {"query": "账龄 逾期"}}]})
        turn = self._work(writer, config=RESEARCH).edit("撰写应收账款的段落", ACCOUNT)
        self.assertFalse(turn.ok)
        self.assertEqual(turn.stop_reason, C.STOP_SOURCE_GAP)
        self.assertIsNone(turn.report.get("prepared_fallback"), "research mode does not hand the account to the tool loop")
        self.assertLessEqual(len(writer.users), 5)

    def test_evidence_not_enough_is_answered_once_and_then_the_paragraph_is_written_from_what_there_is(self):
        def fn(system, user, n):
            if n == 1:
                return {"need_evidence": "账龄资料"}
            return {"draft": {"claims": [amount_claim(writer_view(user), P2, GOOD)], "unresolved": ["账龄资料不在工作簿"]}}
        writer = Writer(fn)
        work = self._work(writer, config=RESEARCH)
        turn = work.edit("撰写应收账款的段落", ACCOUNT)
        self.assertTrue(turn.ok, turn.report)
        self.assertIsNone(turn.report.get("prepared_fallback"))
        self.assertIn("No tool loop follows", writer.users[1])
        kinds = [f["kind"] for f in work.source.candidates[ACCOUNT].flags]
        self.assertIn("research_open", kinds)
        self.assertIn("source_gap", kinds)

    def test_evidence_not_enough_twice_ends_without_a_paragraph_and_names_the_gap(self):
        writer = Writer(lambda system, user, n: {"need_evidence": "账龄资料"})
        turn = self._work(writer, config=RESEARCH).edit("撰写应收账款的段落", ACCOUNT)
        self.assertFalse(turn.ok)
        self.assertEqual(turn.stop_reason, C.STOP_SOURCE_GAP)
        self.assertEqual(turn.report["source_gap"], ["账龄资料", "账龄资料"])
        self.assertEqual(len(writer.users), 2)

    def test_the_gap_reaches_the_reader_as_a_finding_not_as_an_error_in_the_paragraph(self):
        from fdd_utils.agent.export import merged_results, session_delivery
        writer = Writer(lambda system, user, n: {"draft": {"claims": [amount_claim(writer_view(user), P2, GOOD)], "unresolved": []}})
        work = self._work(writer, config=RESEARCH)
        self.assertTrue(work.edit("撰写应收账款的段落", ACCOUNT, accept=True).ok)
        results = merged_results(work.source, [ACCOUNT], first_generation=True)
        delivery = session_delivery(work.source, results, edited=[ACCOUNT], first_generation=True)
        found = [f for f in delivery.findings if f.code == "SOURCE_GAP"]
        self.assertEqual(len(found), 1)
        self.assertEqual((found[0].severity, found[0].account), ("review", ACCOUNT))
        self.assertIn("ageing of the balance", found[0].message)
        self.assertNotIn("ageing of the balance: ageing of the balance", found[0].message)

    def test_a_writer_that_asks_for_a_tool_in_the_old_form_is_served_as_a_lookup(self):
        def fn(system, user, n):
            if n == 1:
                return {"tool": "search_notes", "args": {"query": "审计调整"}}
            return {"draft": {"claims": [amount_claim(writer_view(user), P2, GOOD)], "unresolved": []}}
        writer = Writer(fn)
        turn = self._work(writer, config=RESEARCH).edit("撰写应收账款的段落", ACCOUNT)
        self.assertTrue(turn.ok, turn.report)
        self.assertEqual(turn.report["research"], [{"tool": "search_notes", "ok": True, "hits": 1}])
        self.assertIsNone(turn.report.get("prepared_fallback"))

    def test_with_the_switch_off_a_writer_that_asks_still_gets_the_tool_loop(self):
        writer = Writer(lambda system, user, n: {"need_evidence": "账龄资料"})
        turn = self._work(writer).edit("撰写应收账款的段落", ACCOUNT)
        self.assertEqual(turn.report.get("prepared_fallback"), "账龄资料")


class RepairRecord(unittest.TestCase):
    """What each repair round was told is kept with the run, so a trial can count the rounds and the hints."""

    _work = WF.PreparedWrite._work

    def test_a_refused_sentence_leaves_its_message_in_the_report_and_the_hint_is_counted(self):
        def fn(system, user, n):
            view = writer_view(user)
            if "REJECTED" in user:
                return {"fix": [{"index": 1, "claim": amount_claim(view, P2, "2025年12月31日应收账款余额130.0万元")}]}
            # 130.1万元 is within 1% of the balance the evidence has: a figure made by hand
            return {"draft": {"claims": [amount_claim(view, P2, GOOD), amount_claim(view, P2, "2025年12月31日应收账款余额130.1万元")], "unresolved": []}}
        turn = self._work(Writer(fn)).edit("撰写应收账款的段落", ACCOUNT)
        repairs = turn.report["repairs"]
        self.assertEqual(len(repairs), 1)
        self.assertIn("the nearest figure in EVIDENCE is", repairs[0]["feedback"])
        import entity_trial as E
        self.assertEqual(E._repair_counts(repairs), {"rounds": 1, "nearest figure": 1, "unnamed lines": 0, "line ref": 0})

    def test_a_first_draft_that_passes_has_no_repairs(self):
        writer = Writer(lambda system, user, n: {"draft": {"claims": [amount_claim(writer_view(user), P2, GOOD)], "unresolved": []}})
        turn = self._work(writer).edit("撰写应收账款的段落", ACCOUNT)
        self.assertNotIn("repairs", turn.report)


class PeriodLabels(unittest.TestCase):
    def flags(self, periods):
        controller = C.AgentController.__new__(C.AgentController)
        controller.research, controller._research_open = False, []
        return controller._flags("Paid-in capital", {"periods": periods, "permissions": {}, "checks": {}}, [], [], {}, [], False)

    def test_two_year_start_labels_are_flagged_for_a_person(self):
        got = self.flags(["2023-01-01", "2024-01-01", "2025-01-01", "2026-06-30"])
        self.assertEqual([f["kind"] for f in got], ["period_label_suspect"])
        self.assertIn("2023-01-01", got[0]["text"])

    def test_period_ends_are_not(self):
        self.assertEqual(self.flags(["2023-12-31", "2024-12-31", "2025-12-31", "2026-06-30"]), [])
        self.assertEqual(self.flags(["2024-01-01", "2024-12-31"]), [], "one opening-balance date is not a pattern")


if __name__ == "__main__":
    unittest.main(verbosity=1)

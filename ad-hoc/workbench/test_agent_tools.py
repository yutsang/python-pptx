"""The agent's tools: a failure is never a success, a stale source is refused (Phase 4).

Zero tokens, synthetic sessions only (the real-workbook flows are in test_agent_flows):

    python ad-hoc/workbench/test_agent_tools.py
"""
from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parents[1]))

from fdd_utils.agent import tools as T  # noqa: E402
from fdd_utils.agent.session import SourceSession  # noqa: E402
from fdd_utils.agent.summary import build_numerical_summary, traceability  # noqa: E402
from fdd_utils.ai import evidence_contract as ec  # noqa: E402
from fdd_utils.workbook.selection import Decision, SelectionResult  # noqa: E402
from test_claims import P1, P2, build, ref  # noqa: E402


def make_session(workbook_path: str = "", **contract_kwargs) -> SourceSession:
    contract = build(**contract_kwargs)
    binding = ec.SourceBinding(entity="E", workbook="wb.xlsx", workbook_sha256="aaaa", financials_sheet="Financials",
                               statement_source="financials")
    selection = SelectionResult(decisions=[Decision(contract.account, True, "reconciled", "tied_out", "")])
    return SourceSession(binding=binding, dfs={}, contracts={contract.account: contract}, selection=selection,
                         bs_is_results={"statement_source": "financials"}, workbook_path=workbook_path, language="Chi")


GOOD = lambda c: {"kind": "amount", "periods": [P2], "cites": [ref(c, P2)], "text": "2025年12月31日应收账款余额130.0万元"}  # noqa: E731
BAD = lambda c: {"kind": "amount", "periods": [P2], "cites": [ref(c, P2)], "text": "2025年12月31日应收账款余额100.0万元"}  # noqa: E731


class Contract(unittest.TestCase):
    def test_unknown_tool_bad_args_and_unknown_account_are_errors_not_exceptions(self):
        s = make_session()
        r = T.call_tool(s, "no_such_tool")
        self.assertEqual((r.ok, r.error["code"]), (False, "UNKNOWN_TOOL"))
        self.assertEqual(T.call_tool(s, "get_account_evidence", {}).error["code"], "BAD_ARGS")
        self.assertEqual(T.call_tool(s, "get_account_evidence", {"account": 5}).error["code"], "BAD_ARGS")
        self.assertEqual(T.call_tool(s, "get_account_evidence", {"account": "x", "bogus": 1}).error["code"], "BAD_ARGS")
        r = T.call_tool(s, "get_account_evidence", {"account": "不存在"})
        self.assertEqual((r.ok, r.error["code"]), (False, "UNKNOWN_ACCOUNT"))
        self.assertIn("应收账款", r.data["accounts"])

    def test_a_crashing_tool_is_reported_as_a_failure(self):
        s = make_session()
        original = T.TOOLS["get_account_evidence"]
        T.TOOLS["get_account_evidence"] = T.Tool(original.name, original.help, original.args,
                                                 lambda session, **k: (_ for _ in ()).throw(ZeroDivisionError("boom")))
        try:
            r = T.call_tool(s, "get_account_evidence", {"account": "应收账款"})
        finally:
            T.TOOLS["get_account_evidence"] = original
        self.assertEqual((r.ok, r.error["code"]), (False, "TOOL_EXCEPTION"))
        self.assertIn("ZeroDivisionError", r.error["message"])

    def test_results_carry_the_source_version_and_the_refs_they_rest_on(self):
        s = make_session()
        r = T.call_tool(s, "get_account_evidence", {"account": "应收账款"})
        self.assertTrue(r.ok)
        self.assertEqual(r.source_version, s.source_version)
        self.assertTrue(r.basis and all(b[0] in "FDN" for b in r.basis))

    def test_a_unique_part_of_an_account_name_is_enough_an_ambiguous_one_is_not(self):
        s = make_session()
        self.assertEqual(s.resolve_account("应收"), "应收账款")
        s.contracts["应收票据"] = build(account="应收票据", sheet="Notes")
        self.assertIsNone(s.resolve_account("应收"))


class Staleness(unittest.TestCase):
    def test_a_workbook_that_changed_under_the_session_is_refused(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "wb.xlsx")
            Path(path).write_bytes(b"version one")
            s = make_session(path)
            s.binding.workbook_sha256 = __import__("fdd_utils.run_identity", fromlist=["x"]).file_sha256(path)
            self.assertTrue(T.call_tool(s, "get_account_evidence", {"account": "应收账款"}).ok)
            Path(path).write_bytes(b"version two")
            r = T.call_tool(s, "get_account_evidence", {"account": "应收账款"})
            self.assertEqual((r.ok, r.error["code"]), (False, "STALE_SOURCE"))
            seen = T.call_tool(s, "inspect_sources")            # the way to find out stays available
            self.assertTrue(seen.ok)
            self.assertIn("changed on disk", seen.data["stale"])
            self.assertTrue(any(i["code"] == "STALE_SOURCE" for i in seen.issues))


class Validation(unittest.TestCase):
    def test_failed_claims_are_not_ok_and_say_which(self):
        s = make_session()
        c = s.contracts["应收账款"]
        r = T.call_tool(s, "validate_candidate", {"account": "应收账款", "claims": [GOOD(c), BAD(c)]})
        self.assertEqual((r.ok, r.error["code"]), (False, "VALIDATION_FAILED"))
        self.assertEqual([v["verdict"] for v in r.data["outcome"]["verdicts"]], ["verified", "failed"])
        self.assertFalse(s.candidates["应收账款"].validated)

    def test_a_clean_candidate_is_ok_and_remembered_with_its_digest(self):
        s = make_session()
        c = s.contracts["应收账款"]
        r = T.call_tool(s, "validate_candidate", {"account": "应收账款", "claims": [GOOD(c)]})
        self.assertTrue(r.ok, r.error)
        self.assertEqual(r.data["delivery_preview"], "ready")
        self.assertTrue(s.candidates["应收账款"].validated)
        self.assertEqual(s.candidates["应收账款"].claims_digest, r.data["claims_digest"])

    def test_empty_claims_are_a_bad_request(self):
        s = make_session()
        self.assertEqual(T.call_tool(s, "validate_candidate", {"account": "应收账款", "claims": []}).error["code"], "BAD_ARGS")


class Export(unittest.TestCase):
    def test_export_refuses_what_was_never_validated(self):
        s = make_session()
        self.assertEqual(T.call_tool(s, "export_candidate").error["code"], "NO_CANDIDATE")
        c = s.contracts["应收账款"]
        T.call_tool(s, "validate_candidate", {"account": "应收账款", "claims": [BAD(c)]})
        self.assertEqual(T.call_tool(s, "export_candidate", {"accounts": ["应收账款"]}).error["code"], "NOT_VALIDATED")

    def test_a_candidate_validated_against_older_sources_is_no_longer_validated(self):
        s = make_session()
        c = s.contracts["应收账款"]
        self.assertTrue(T.call_tool(s, "validate_candidate", {"account": "应收账款", "claims": [GOOD(c)]}).ok)
        s.binding.workbook_sha256 = "bbbb"                         # the sources moved on
        r = T.call_tool(s, "export_candidate", {"accounts": ["应收账款"]})
        self.assertEqual((r.ok, r.error["code"]), (False, "NOT_VALIDATED"))


class Analysis(unittest.TestCase):
    def test_computed_figures_come_back_as_refs_and_forbidden_ones_as_errors(self):
        s = make_session()
        r = T.call_tool(s, "compute_analysis", {"account": "应收账款", "op": "movement", "period_from": P1, "period": P2})
        self.assertTrue(r.ok)
        self.assertEqual(r.data["derived"]["value"], 300000.0)
        self.assertTrue(r.basis)
        bad = T.call_tool(s, "compute_analysis", {"account": "应收账款", "op": "share", "line": "不存在", "period": P2})
        self.assertEqual((bad.ok, bad.error["code"]), (False, "NOT_COMPUTABLE"))
        self.assertEqual(T.call_tool(s, "compute_analysis", {"account": "应收账款", "op": "frobnicate"}).error["code"], "BAD_ARGS")

    def test_trace_goes_back_to_the_sheet_and_row(self):
        s = make_session()
        c = s.contracts["应收账款"]
        r = T.call_tool(s, "trace_claim", {"account": "应收账款", "ref": ref(c, P2, "甲公司")})
        self.assertTrue(r.ok)
        self.assertEqual((r.data["source"]["sheet"], r.data["source"]["row"]), ("AR", 4))
        self.assertEqual(T.call_tool(s, "trace_claim", {"account": "应收账款", "ref": "F999"}).error["code"], "UNKNOWN_REF")


class Summary(unittest.TestCase):
    def test_every_printed_number_traces_to_a_fact(self):
        s = make_session()
        r = T.call_tool(s, "numerical_summary")
        self.assertTrue(r.ok, r.error)
        trace = r.data["traceability"]
        self.assertTrue(trace["complete"])
        self.assertGreaterEqual(trace["numbers_printed"], 3)         # balance, change, percent
        self.assertEqual(trace["traced"], trace["numbers_printed"])

    def test_a_number_nothing_backs_is_caught(self):
        s = make_session()
        rows = build_numerical_summary(s)["rows"]
        lines = ["应收账款 | 2025-12-31余额 130.0万元 [F1] | 较上年变动 999.0万元 [D1]"]
        trace = traceability(s, rows, lines)
        self.assertFalse(trace["complete"])
        self.assertTrue(any("999.0万元" in item for item in trace["untraced"]))

    def test_a_very_large_real_movement_is_traced_and_an_amount_is_not_mistaken_for_a_ratio(self):
        """A balance that went from 1.6万 to 182.7万 moved +11,365.9%. The check used to drop every ratio over 50 (5,000%),
        so that real figure read as 'nothing backs it' on a real workbook."""
        rows = [("甲公司", "detail", {P1: 16000, P2: 1827000}), ("Total", "total", {P1: 16000, P2: 1827000})]
        s = make_session(rows=rows, financials={P1: 16000.0, P2: 1827000.0})
        r = T.call_tool(s, "numerical_summary")
        self.assertTrue(r.ok, r.error)
        self.assertIn("+11", "\n".join(r.data["lines"]))
        self.assertTrue(r.data["traceability"]["complete"], r.data["traceability"]["untraced"])
        # a percentage that is just an amount's digits is still not accepted
        lines = [r.data["lines"][0] + " | 幅度 +1827000.0%"]
        self.assertFalse(traceability(s, r.data["rows"], lines)["complete"])

    def test_a_replaced_period_and_a_missing_financials_show_in_the_flags(self):
        rows = [("甲公司", "detail", {P1: 500000, P2: 700000}), ("乙公司", "detail", {P1: 300000, P2: 600000}),
                ("Total", "total", {P1: 1000000, P2: 1300000})]
        s = make_session(rows=rows, financials={P1: 1000000.0, P2: 1300000.0}, authority={P1: {"tab": 800000.0, "financials": 1000000.0}})
        info = T.call_tool(s, "inspect_sources").data["accounts"][0]
        self.assertEqual(info["replaced_periods"], [P1])
        text = "\n".join(T.call_tool(s, "numerical_summary").data["lines"])
        self.assertIn("与报表一致", text)                              # the latest period still ties


class Inspect(unittest.TestCase):
    def test_excluded_accounts_and_unmatched_financials_are_listed(self):
        s = make_session()
        s.selection = SelectionResult(
            decisions=[Decision("应收账款", True, "reconciled", "tied_out", ""),
                       Decision("关联方", False, "excluded", "no_reconciliation_row", "no row")],
            unmatched_financials=[{"statement": "BS", "financials_account": "Retained earnings"}])
        data = T.call_tool(s, "inspect_sources").data
        self.assertEqual(data["selection"]["excluded"], [{"account": "关联方", "reason": "no_reconciliation_row"}])
        self.assertEqual(data["selection"]["unmatched_financials"][0]["financials_account"], "Retained earnings")
        self.assertEqual(data["statement_source"], "financials")


class HiddenFinancials(unittest.TestCase):
    """A hidden sheet never becomes the Financials sheet, and an unreadable visibility never means 'all of them'."""

    def book(self, tmp, **states):
        from openpyxl import Workbook
        wb = Workbook()
        first = True
        for name, state in states.items():
            ws = wb.active if first else wb.create_sheet()
            first = False
            ws.title = name
            ws["A1"] = name
            ws.sheet_state = state
        path = os.path.join(tmp, "wb.xlsx")
        wb.save(path)
        return path

    def test_a_visible_financials_sheet_is_found(self):
        from fdd_utils.agent.session import find_financials_sheet
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(find_financials_sheet(self.book(tmp, Data="visible", Financials="visible")), "Financials")

    def test_a_hidden_financials_sheet_is_not_picked_even_when_it_is_the_only_one(self):
        from fdd_utils.agent.session import find_financials_sheet
        for state in ("hidden", "veryHidden"):
            with tempfile.TemporaryDirectory() as tmp:
                self.assertIsNone(find_financials_sheet(self.book(tmp, Data="visible", Financials=state)), state)

    def test_a_visible_candidate_wins_over_a_hidden_exact_match(self):
        from fdd_utils.agent.session import find_financials_sheet
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(find_financials_sheet(self.book(tmp, Data="visible", Financials="hidden", **{"Financials - NB": "visible"})),
                             "Financials - NB")

    def test_when_visibility_cannot_be_read_nothing_is_chosen(self):
        import fdd_utils.agent.session as session_module
        import fdd_utils.workbook.visibility as visibility
        original = visibility.sheet_states
        visibility.sheet_states = lambda path: None
        try:
            with tempfile.TemporaryDirectory() as tmp:
                path = self.book(tmp, Data="visible", Financials="visible")
                self.assertIsNone(session_module.find_financials_sheet(path))
                with self.assertRaises(RuntimeError):
                    session_module.refuse_hidden(path, "Financials")
        finally:
            visibility.sheet_states = original

    def test_naming_a_hidden_sheet_explicitly_is_refused_naming_a_visible_one_is_not(self):
        from fdd_utils.agent.session import refuse_hidden
        with tempfile.TemporaryDirectory() as tmp:
            path = self.book(tmp, Data="visible", Financials="hidden", Visible="visible")
            with self.assertRaisesRegex(RuntimeError, "hidden"):
                refuse_hidden(path, "Financials")
            refuse_hidden(path, "Visible")
            refuse_hidden(path, None)


class Flag(unittest.TestCase):
    """Default off; the environment wins over the file in both directions; budgets come from config."""

    def setUp(self):
        self.saved = os.environ.pop("FDD_AGENT", None)

    def tearDown(self):
        os.environ.pop("FDD_AGENT", None)
        if self.saved is not None:
            os.environ["FDD_AGENT"] = self.saved

    def test_off_unless_asked(self):
        from fdd_utils.agent.flag import agent_enabled
        self.assertFalse(agent_enabled({}))
        self.assertFalse(agent_enabled({"agent": {"enabled": False}}))
        self.assertTrue(agent_enabled({"agent": {"enabled": True}}))

    def test_the_environment_overrides_the_file_either_way(self):
        from fdd_utils.agent.flag import agent_enabled
        os.environ["FDD_AGENT"] = "1"
        self.assertTrue(agent_enabled({"agent": {"enabled": False}}))
        os.environ["FDD_AGENT"] = "off"
        self.assertFalse(agent_enabled({"agent": {"enabled": True}}))

    def test_budget_reads_the_config_and_falls_back(self):
        from fdd_utils.agent.flag import agent_budget
        b = agent_budget({"agent": {"max_steps": 5, "max_tokens": 1234}})
        self.assertEqual((b.max_steps, b.max_tokens, b.max_rejections), (5, 1234, 3))


if __name__ == "__main__":
    unittest.main(verbosity=2)

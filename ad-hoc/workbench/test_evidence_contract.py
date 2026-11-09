"""The evidence contract: facts by basis, checks per account x period, permissions (Phase 2).

Zero tokens, synthetic frames only:

    python ad-hoc/workbench/test_evidence_contract.py

The case this exists for: the tab agrees with Financials on the latest period and
not on an earlier one. `workbook/authority.py` then replaces the earlier total
with Financials' figure, after which a comparison of "the account total" with
Financials agrees in every period -- so the legacy tie-out reports the replaced
period as tied, and nothing says the tab's detail under it no longer adds up.
`test_replaced_total_is_not_a_tie_out` pins the difference.
"""
from __future__ import annotations

import inspect
import re
import sys
import unittest
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

from fdd_utils.ai import evidence_contract as ec  # noqa: E402
from fdd_utils.ai.evidence import AccountEvidence, compile_account_evidence  # noqa: E402
from fdd_utils.workbook.authority import AUTHORITY_KEY  # noqa: E402

P1, P2 = "2024-12-31", "2025-12-31"


def account_frame(rows, *, sheet="AR", authority=None, periods=(P1, P2)):
    """rows: [(label, row type, {period: value})]. Returns a frame shaped like the
    pipeline's: description column, source row index, one column per period."""
    data = {"Account": [r[0] for r in rows], "__source_row_idx": list(range(4, 4 + len(rows)))}
    for p in periods:
        data[p] = [float(r[2].get(p, 0.0)) for r in rows]
    df = pd.DataFrame(data)
    types = {r[0]: r[1] for r in rows}
    df.attrs.update({"row_types_by_description": dict(types), "source_sheet_name": sheet,
                     "integrity": {"sheet_name": sheet}})
    analysis = df.copy()
    analysis.attrs.update({"row_types_by_description": dict(types)})
    if authority:
        df.attrs[AUTHORITY_KEY] = authority
        analysis.attrs[AUTHORITY_KEY] = authority
    df.attrs["prompt_analysis_df"] = analysis
    return df


def clean_rows():
    return [("甲公司", "detail", {P1: 60, P2: 70}), ("乙公司", "detail", {P1: 40, P2: 30}),
            ("Total", "total", {P1: 100, P2: 100})]


class Facts(unittest.TestCase):
    def test_every_total_keeps_its_bases_apart(self):
        frame = account_frame(clean_rows())
        c = ec.build_contract("AR", frame, statement_type="BS", financials={P1: 100.0, P2: 100.0})
        for p in (P1, P2):
            self.assertEqual({b: c.total(p, b) for b in (ec.BASIS_RAW, ec.BASIS_FINANCIALS, ec.BASIS_ADJUSTED)},
                             {ec.BASIS_RAW: 100.0, ec.BASIS_FINANCIALS: 100.0, ec.BASIS_ADJUSTED: 100.0})
        ids = [f["id"] for f in c.facts]
        self.assertEqual(len(ids), len(set(ids)), "semantic ids are unique")
        self.assertTrue(all(f.get("ref", "").startswith("F") for f in c.facts))
        line = [f for f in c.facts if f["role"] == ec.ROLE_COMPONENT and f["line"] == "甲公司" and f["period"] == P2][0]
        self.assertEqual((line["basis"], line["value"], line["sheet"], line["row"]), (ec.BASIS_RAW, 70.0, "AR", 4))

    def test_clean_account_passes_everywhere_and_permits_everything(self):
        c = ec.build_contract("AR", account_frame(clean_rows()), statement_type="BS", financials={P1: 100.0, P2: 100.0})
        self.assertEqual(c.status_matrix(), {p: {ec.CHECK_TOTAL_VS_FINANCIALS: ec.PASS, ec.CHECK_DETAIL_VS_TOTAL: ec.PASS}
                                             for p in (P1, P2)})
        self.assertTrue(all(c.permissions[P2][k] for k in ("describe_total", "explain_with_detail",
                                                           "state_nil_or_first_arose", "compute_movement")))

    def test_contract_is_deterministic(self):
        a = ec.build_contract("AR", account_frame(clean_rows()), statement_type="BS", financials={P1: 100.0, P2: 100.0})
        b = ec.build_contract("AR", account_frame(clean_rows()), statement_type="BS", financials={P1: 100.0, P2: 100.0})
        self.assertEqual(a.to_dict(), b.to_dict())


class ReplacedTotal(unittest.TestCase):
    """The tab agrees on the latest period and was never updated for an earlier one."""

    def setUp(self):
        rows = [("甲公司", "detail", {P1: 50, P2: 70}), ("乙公司", "detail", {P1: 30, P2: 30}),
                ("Total", "total", {P1: 100.0, P2: 100.0})]      # P1 total already replaced by Financials' 100
        self.record = {P1: {"tab": 80.0, "financials": 100.0, "adjustments": ["「ADJ」Audit ADJ#4 +20"]}}
        self.frame = account_frame(rows, authority=self.record)
        self.contract = ec.build_contract("AR", self.frame, statement_type="BS", financials={P1: 100.0, P2: 100.0})

    def test_replaced_total_is_not_a_tie_out(self):
        # The legacy comparison reads the account total off the frame the model reads:
        adjusted = {P1: 100.0, P2: 100.0}
        self.assertTrue(all(abs(adjusted[p] - 100.0) <= 1.0 for p in adjusted), "legacy: tied in every period")
        # The contract compares the tab's OWN total, recovered from the authority record:
        tf = self.contract.check(P1, ec.CHECK_TOTAL_VS_FINANCIALS)
        self.assertEqual((tf["status"], tf["left"], tf["right"], tf["overridden"]), (ec.MISMATCH, 80.0, 100.0, True))
        self.assertEqual(self.contract.check(P2, ec.CHECK_TOTAL_VS_FINANCIALS)["status"], ec.PASS)

    def test_three_values_coexist_for_the_replaced_period(self):
        self.assertEqual((self.contract.total(P1, ec.BASIS_RAW), self.contract.total(P1, ec.BASIS_FINANCIALS),
                          self.contract.total(P1, ec.BASIS_ADJUSTED)), (80.0, 100.0, 100.0))
        adjusted = [f for f in self.contract.facts if f["period"] == P1 and f["basis"] == ec.BASIS_ADJUSTED][0]
        self.assertTrue(adjusted["overridden"])

    def test_detail_is_not_verified_by_the_replacement(self):
        dt = self.contract.check(P1, ec.CHECK_DETAIL_VS_TOTAL)
        self.assertEqual(dt["status"], ec.MISMATCH)
        self.assertEqual((dt["left"], dt["right"]), (80.0, 100.0))
        perm = self.contract.permissions[P1]
        self.assertFalse(perm["explain_with_detail"])
        self.assertFalse(perm["state_nil_or_first_arose"])
        self.assertTrue(perm["describe_total"] and perm["describe_difference"])

    def test_the_intact_period_is_unaffected(self):
        self.assertTrue(self.contract.permissions[P2]["explain_with_detail"])


class OtherSituations(unittest.TestCase):
    def test_detail_that_misses_the_total_with_no_override_is_ambiguous_not_failed(self):
        rows = [("甲", "detail", {P1: 50, P2: 50}), ("乙", "detail", {P1: 30, P2: 30}), ("Total", "total", {P1: 100, P2: 100})]
        c = ec.build_contract("AR", account_frame(rows), statement_type="BS", financials={P1: 100.0, P2: 100.0})
        self.assertEqual(c.check(P2, ec.CHECK_DETAIL_VS_TOTAL)["status"], ec.AMBIGUOUS)
        self.assertFalse(c.permissions[P2]["explain_with_detail"])
        self.assertEqual(c.check(P2, ec.CHECK_TOTAL_VS_FINANCIALS)["status"], ec.PASS)

    def test_tab_and_financials_disagree_and_nothing_chose(self):
        c = ec.build_contract("AR", account_frame(clean_rows()), statement_type="BS", financials={P1: 100.0, P2: 120.0})
        self.assertEqual(c.check(P2, ec.CHECK_TOTAL_VS_FINANCIALS)["status"], ec.MISMATCH)
        perm = c.permissions[P2]
        self.assertFalse(perm["describe_total"])
        self.assertTrue(perm["describe_difference"])

    def test_rounding_inside_tolerance_is_a_pass(self):
        c = ec.build_contract("AR", account_frame(clean_rows()), statement_type="BS", financials={P1: 100.4, P2: 100.0})
        self.assertEqual(c.check(P1, ec.CHECK_TOTAL_VS_FINANCIALS)["status"], ec.PASS)

    def test_income_statement_compares_absolute_values(self):
        rows = [("租金", "detail", {P1: 60, P2: 70}), ("Total", "total", {P1: 60, P2: 70})]
        c = ec.build_contract("COGS", account_frame(rows, sheet="COGS"), statement_type="IS",
                              financials={P1: -60.0, P2: -70.0})
        self.assertEqual(c.check(P2, ec.CHECK_TOTAL_VS_FINANCIALS)["status"], ec.PASS)

    def test_missing_financials_period_is_missing_not_pass(self):
        c = ec.build_contract("AR", account_frame(clean_rows()), statement_type="BS", financials={P2: 100.0})
        self.assertEqual(c.check(P1, ec.CHECK_TOTAL_VS_FINANCIALS)["status"], ec.MISSING)

    def test_synthesized_statements_are_not_comparable(self):
        binding = ec.SourceBinding(statement_source="synthesized")
        c = ec.build_contract("AR", account_frame(clean_rows()), statement_type="BS", binding=binding,
                              financials={P1: 100.0, P2: 100.0})
        self.assertEqual({k["status"] for k in c.checks if k["kind"] == ec.CHECK_TOTAL_VS_FINANCIALS}, {ec.NOT_COMPARABLE})
        self.assertFalse([f for f in c.facts if f["basis"] == ec.BASIS_FINANCIALS])

    def test_stub_period_blocks_an_income_statement_movement(self):
        periods = (P1, "2025-12-31", "2026-01-31")
        rows = [("租金", "detail", {p: 10 for p in periods}), ("Total", "total", {p: 10 for p in periods})]
        c = ec.build_contract("Rev", account_frame(rows, sheet="Rev", periods=periods), statement_type="IS")
        self.assertFalse(c.permissions["2026-01-31"]["compute_movement"])
        self.assertTrue(c.permissions["2025-12-31"]["compute_movement"])
        bs = ec.build_contract("Cash", account_frame(rows, sheet="Cash", periods=periods), statement_type="BS")
        self.assertTrue(bs.permissions["2026-01-31"]["compute_movement"])      # a balance is a point in time

    def test_trace_runs_from_fact_to_checks(self):
        c = ec.build_contract("AR", account_frame(clean_rows()), statement_type="BS", financials={P1: 100.0, P2: 100.0})
        raw = [f for f in c.facts if f["period"] == P2 and f["basis"] == ec.BASIS_RAW and f["role"] == "total"][0]
        trace = ec.trace_fact(c, raw["ref"])
        self.assertEqual(trace["source"]["sheet"], "AR")
        self.assertTrue(any(ch["kind"] == ec.CHECK_TOTAL_VS_FINANCIALS for ch in trace["checks"]))
        self.assertFalse(ec.trace_fact(c, "F999")["found"])


class SourceVersion(unittest.TestCase):
    def test_version_follows_the_inputs(self):
        a = ec.SourceBinding(workbook_sha256="aa", financials_sheet="Financials", entity="E")
        b = ec.SourceBinding(workbook_sha256="bb", financials_sheet="Financials", entity="E")
        self.assertNotEqual(a.source_version, b.source_version)
        self.assertEqual(a.source_version, ec.SourceBinding.from_dict(a.to_dict()).source_version)
        self.assertNotEqual(a.source_version, ec.SourceBinding(workbook_sha256="aa", financials_sheet="Financials",
                                                               entity="Other").source_version)


class HiddenSheets(unittest.TestCase):
    def digest(self, hidden):
        return {"sheets": {"AR": {"hidden": hidden, "status": "mapped", "cells": {"r": [], "c": [], "v": []},
                                  "row_labels": {}, "col_labels": {}, "texts": []}}}

    def test_account_on_a_hidden_tab_offers_no_facts(self):
        frame = account_frame(clean_rows())
        hidden = ec.build_contract("AR", frame, statement_type="BS", financials={P1: 100.0, P2: 100.0},
                                   digest=self.digest(True))
        self.assertEqual(hidden.facts, [])
        self.assertEqual(hidden.issues[0]["code"], "SOURCE_SHEET_HIDDEN")
        visible = ec.build_contract("AR", frame, statement_type="BS", financials={P1: 100.0, P2: 100.0},
                                    digest=self.digest(False))
        self.assertTrue(visible.facts)

    def test_adjustment_entries_ignore_a_hidden_register(self):
        from fdd_utils.workbook.authority import adjustment_entries
        sheet = {"status": "unmapped", "cells": {"r": [1], "c": [2], "v": [1500.0]},
                 "col_labels": {"2": "Audit ADJ#1"}, "row_labels": {"1": "投资收益"}, "texts": []}
        shown = adjustment_entries({"sheets": {"ADJ": dict(sheet, hidden=False)}}, 1500.0, P2)
        self.assertTrue(shown, "negative control: the same register, visible, is found")
        self.assertEqual(adjustment_entries({"sheets": {"ADJ": dict(sheet, hidden=True)}}, 1500.0, P2), [])

    def test_every_digest_reader_goes_through_visible_sheets(self):
        """Behavioural: each reader calls the filter. Static: no module outside the
        digest itself reads digest['sheets'] directly, so a new reader cannot skip it."""
        import fdd_utils.workbook.digest as digest_module
        from fdd_utils.ai import facts
        from fdd_utils.workbook import authority
        calls = []
        original = digest_module.visible_sheets
        digest_module.visible_sheets = lambda d: (calls.append(1) or {})
        try:
            probe = {"sheets": {"S": {"hidden": False}}}
            authority.adjustment_entries(probe, 2000.0, P2)
            authority._adjustment_cells(probe)
            authority.adjustment_bridge(probe, "AR")
            facts.build_digest_graph_links(probe, {"series": {"AR": {P2: 1.0}}})
            facts.digest_tieout_diagnostics(probe, {"series": {"AR": {P2: 1.0}}})
            facts.supporting_breakdowns_for("AR", probe, {"links": [{"passed": True, "kind": "supporting_schedule_tieout",
                                                                     "target": "AR", "evidence": {"sheet": "S"}}]})
        finally:
            digest_module.visible_sheets = original
        self.assertGreaterEqual(len(calls), 6, "authority x3 + facts x3")
        offenders = []
        allowed = {"fdd_utils/workbook/digest.py"}
        pattern = re.compile(r"""(workbook_)?digest[^\n]{0,40}\.get\(\s*["']sheets["']\)|\(digest or \{\}\)\.get\(["']sheets["']\)""")
        for path in (ROOT / "fdd_utils").rglob("*.py"):
            rel = path.relative_to(ROOT).as_posix()
            if rel in allowed:
                continue
            for number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
                if pattern.search(line) and "visible_sheets" not in line and "digest-ok" not in line:
                    offenders.append(f"{rel}:{number}: {line.strip()[:90]}")
        self.assertEqual(offenders, [], "a digest reader that bypasses visible_sheets")


class Persistence(unittest.TestCase):
    def test_contract_rides_on_account_evidence_and_survives_a_round_trip(self):
        frame = account_frame(clean_rows())
        evidence = compile_account_evidence("AR", frame, statement_type="BS", graph_facts={"links": [{
            "kind": "tab_to_financials", "source": "AR", "target": "Financials::AR", "passed": True,
            "evidence": {"financials_values": {P1: 100.0, P2: 100.0}}}]})
        self.assertIsNotNone(evidence.contract)
        again = AccountEvidence.from_dict(evidence.to_dict())
        contract = ec.EvidenceContract.from_dict(again.contract)
        self.assertEqual(contract.check(P2, ec.CHECK_TOTAL_VS_FINANCIALS)["status"], ec.PASS)
        self.assertEqual(contract.to_dict(), evidence.contract)

    def test_a_frame_the_contract_cannot_read_leaves_the_pool_intact(self):
        evidence = compile_account_evidence("X", pd.DataFrame({"a": [1.0]}))
        self.assertTrue(evidence.facts is not None)     # no exception, pool built either way



class SupportingTables(unittest.TestCase):
    """The other tables on the account's own tab. The 2026-11-09 evidence locator found the team's report quoting figures from
    tables beside the schedule (by nature, ageing, phases) that the writer never saw."""

    def sheet(self, side_title="应付账款按性质 - 甲公司"):
        grid = [[None] * 12 for _ in range(20)]
        def put(r, c, *vals):
            for i, v in enumerate(vals):
                grid[r][c + i] = v
        # the main schedule, A1:D6
        put(0, 0, "应付账款 - 甲公司", "人民币千元")
        put(1, 0, "科目", "2024-12-31", "2025-12-31")
        put(2, 0, "甲供应商", 60, 70)
        put(3, 0, "乙供应商", 40, 30)
        put(4, 0, "合计", 100, 100)
        # beside it: the balance by nature, in thousands
        put(0, 5, side_title, "人民币千元")
        put(1, 5, "项目", "2024年12月31日", "2025年12月31日", "2023年12月31日")
        put(2, 5, "应付工程款", 70.5, 73.09, 0)
        put(3, 5, "应付物业费", 29.5, 26.91, 0)
        put(4, 5, "合计", 100, 100, 0)
        # below: an ageing split with no period in the headers
        put(8, 5, "账龄分析", None)
        put(9, 5, "项目", "1年以内", "1-2年")
        put(10, 5, "应付工程款", 50.0, 23.09)
        put(11, 5, "应付物业费", 26.91, 0)
        # below the schedule: a voucher listing (dates and codes), not a breakdown
        put(12, 0, "日期", "凭证号", "科目编码", "金额")
        put(13, 0, "2025-03-22", 3000, 220201, 120)
        put(14, 0, "2025-06-22", 3001, 220201, 80)
        put(15, 0, "2025-06-30", 3002, 220202, 15)
        return pd.DataFrame(grid)

    def blocks(self, **kw):
        from fdd_utils.workbook import supporting as sp
        return sp.find_supporting_blocks(self.sheet(**kw), [(0, 4, 0, 3)], periods=["2023-12-31", P1, P2], multiplier=1000.0,
                                         entity="甲公司")

    def test_tables_beside_and_below_are_read_and_a_voucher_listing_is_not(self):
        found = self.blocks()
        titles = sorted(b["title"] for b in found)
        self.assertEqual(len(found), 2, titles)
        nature = next(b for b in found if "按性质" in b["title"])
        self.assertEqual([c["period"] for c in nature["columns"]], [P1, P2, "2023-12-31"])     # a nil year is kept
        row = next(r for r in nature["rows"] if r["label"] == "应付工程款")
        col = next(c["col"] for c in nature["columns"] if c["period"] == P2)
        self.assertAlmostEqual(row["values"][col], 73090.0)        # thousands -> yuan
        ageing = next(b for b in found if "账龄" in b["title"])
        self.assertTrue(all(c["assumed"] and c["period"] == P2 for c in ageing["columns"]))

    def test_a_table_titled_for_another_company_is_left_out(self):
        titles = [b["title"] for b in self.blocks(side_title="应付账款按性质 - 乙集团")]
        self.assertFalse(any("乙集团" in t for t in titles), titles)

    def test_filed_as_facts_after_the_contracts_own_refs_and_citable(self):
        from fdd_utils.ai import claims as cl
        c = ec.build_contract("AR", account_frame(clean_rows()), statement_type="BS", financials={P1: 100.0, P2: 100.0})
        before = {f["ref"] for f in c.facts}
        self.assertEqual(ec.add_supporting_tables(c, self.blocks()), 2)
        self.assertEqual({f["ref"] for f in c.facts if f["role"] != ec.ROLE_SUPPORT}, before)
        fact = next(f for f in c.facts if f["role"] == ec.ROLE_SUPPORT and f["line"] == "应付工程款" and f["period"] == P2
                    and not f["period_assumed"])
        claim = {"account": "AR", "kind": "composition", "text": "2025年12月31日应付工程款7.3万元", "periods": [P2], "cites": [fact["ref"]]}
        v = cl.verify_claim(claim, {"AR": c})
        self.assertEqual(v["verdict"], "verified", v["findings"])
        view = cl.evidence_view(c)
        self.assertEqual(len(view["supporting_tables"]), 2)
        self.assertIn(fact["ref"], {r for t in view["supporting_tables"] for row in t["rows"] for r in row["refs"]})
        wrong = dict(claim, text="2025年12月31日应付工程款9.9万元")
        self.assertEqual(cl.verify_claim(wrong, {"AR": c})["verdict"], "failed")

    def test_a_table_that_repeats_the_schedules_own_lines_is_skipped(self):
        c = ec.build_contract("AR", account_frame(clean_rows()), statement_type="BS", financials={P1: 100.0, P2: 100.0})
        copy = [{"title": "复本", "where": "F1:H4", "columns": [{"col": 1, "header": "2025年", "period": P2, "assumed": False}],
                 "rows": [{"label": "甲公司", "row": 2, "values": {1: 70.0}, "total": False},
                          {"label": "乙公司", "row": 3, "values": {1: 30.0}, "total": False}]}]
        self.assertEqual(ec.add_supporting_tables(c, copy), 0)

if __name__ == "__main__":
    unittest.main(verbosity=2)

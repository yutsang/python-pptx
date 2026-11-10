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

    def test_account_on_a_hidden_tab_offers_only_its_financials_figures(self):
        frame = account_frame(clean_rows())
        hidden = ec.build_contract("AR", frame, statement_type="BS", financials={P1: 100.0, P2: 120.0},
                                   digest=self.digest(True))
        # nothing from the tab; the visible Financials row is the account's figure (投资收益, client file, 2026-11-09)
        self.assertEqual({(f["basis"], f["role"], f["line"]) for f in hidden.facts},
                         {(ec.BASIS_FINANCIALS, ec.ROLE_TOTAL, ""), (ec.BASIS_ADJUSTED, ec.ROLE_TOTAL, "")})
        self.assertEqual(hidden.total(P2, ec.BASIS_ADJUSTED), 120.0)
        self.assertEqual((hidden.issues[0]["code"], hidden.issues[0]["severity"]), ("SOURCE_SHEET_HIDDEN", "info"))
        # what the writer and the dispatcher read: the view, the derived movements (a fact without a ref stopped a trial)
        from fdd_utils.ai import claims as cl
        cl.standard_derived(hidden)
        view = cl.evidence_view(hidden)
        self.assertTrue(all(f.get("ref") for f in hidden.facts))
        self.assertIn(120.0, [t["yuan"] for t in view["totals"] if t["basis"] == ec.BASIS_ADJUSTED])
        bare = ec.build_contract("AR", frame, statement_type="BS", digest=self.digest(True))
        self.assertEqual((bare.facts, bare.issues[0]["severity"]), ([], "blocked"))
        visible = ec.build_contract("AR", frame, statement_type="BS", financials={P1: 100.0, P2: 100.0},
                                    digest=self.digest(False))
        self.assertTrue(visible.facts)

    def test_an_expense_adjustment_is_written_in_the_expenses_own_direction(self):
        # client file, 2026-11-09: the period bridges show costs by their effect on profit (negative), so an adjustment of -45.0万元 on 财务费用
        # raised the expense; a model copied the sign and wrote 调减
        from fdd_utils.workbook.authority import adjustment_bridge

        def bridge(cost_sign, adj):
            cells = {(1, 1): 500 * cost_sign, (1, 3): 500 * cost_sign, (2, 1): 484 * cost_sign, (2, 2): adj,
                     (2, 3): 484 * cost_sign + adj, (3, 2): -adj}
            sheet = {"cells": {"r": [k[0] for k in cells], "c": [k[1] for k in cells], "v": list(cells.values())},
                     "col_labels": {"1": "TB", "2": "Audit ADJ#4", "3": "审定数"},
                     "row_labels": {"1": "营业成本", "2": "财务费用", "3": "投资收益"}, "texts": []}
            return adjustment_bridge({"sheets": {"2025": sheet}}, "财务费用")["lines"]
        # costs shown negative (by their effect on profit): -450 on the expense line raised the expense
        self.assertTrue(any("+45.0万元（调增）" in x for x in bridge(-1, -450.0)), bridge(-1, -450.0))
        # costs shown positive: the sign is the expense's own
        self.assertTrue(any("-45.0万元（调减）" in x for x in bridge(+1, -450.0)), bridge(+1, -450.0))
        self.assertTrue(any("+45.0万元（调增）" in x for x in bridge(+1, 450.0)), bridge(+1, 450.0))

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

    def test_the_schedule_region_ends_at_its_last_row_read(self):
        # the client file, 2026-11-09: the extractor's block ran to row 61 for a schedule ending at row 34, and the phase table under
        # it was excluded with the schedule
        from fdd_utils.workbook import supporting as sp
        attrs = {"integrity": {"block_start_row": 8, "block_end_row": 61}, "normalized_columns": [{"col_idx": 21}]}
        self.assertEqual(sp.main_region(attrs, 30), (2, 62, 0, 21))
        self.assertEqual(sp.main_region(attrs, 30, last_row=33), (2, 34, 0, 21))
        self.assertEqual(sp.main_region(attrs, 30, last_row=70), (2, 62, 0, 21))     # never wider than the extractor's block

    def test_a_check_row_neither_joins_nor_titles_the_table_under_it(self):
        from fdd_utils.workbook import supporting as sp
        grid = [[None] * 8 for _ in range(10)]
        grid[0][1:5] = ["Check", 0, 0, 0]
        grid[1][1:4] = ["固定资产 - 截至2025年12月31日", None, None]
        grid[2][1:4] = ["项目", "原值", "净值"]
        grid[3][1:4] = ["一期", 300.0, 250.0]
        grid[4][1:4] = ["二期", 200.0, 150.0]
        found = sp.find_supporting_blocks(pd.DataFrame(grid), [], periods=[P1, P2], multiplier=1000.0)
        self.assertEqual(len(found), 1)
        self.assertNotIn("Check", found[0]["title"])
        self.assertTrue(found[0]["where"].startswith("B2:"), found[0]["where"])
        self.assertTrue(all(c["period"] == P2 and not c["assumed"] for c in found[0]["columns"]))     # the title's date

    def test_a_line_repeated_under_each_section_takes_the_sections_name_and_a_count_column_is_no_amount(self):
        from fdd_utils.workbook import supporting as sp
        grid = [[None] * 8 for _ in range(12)]
        grid[1][1:6] = ["固定资产 - 截至2025年12月31日（人民币千元）", None, None, None, None]
        grid[2][1:6] = ["项目", "原值", "累计折旧", "净值", "折旧月数"]
        grid[3][1:6] = ["土地使用权", 600.0, -50.0, 550.0, 480.0]
        grid[4][1:6] = ["房屋建筑物", 3000.0, -700.0, 2300.0, 240.0]
        grid[5][1:6] = ["一期合计", 3600.0, -750.0, 2850.0, None]
        grid[6][1:6] = ["土地使用权", 500.0, -20.0, 480.0, 500.0]
        grid[7][1:6] = ["房屋建筑物", 2500.0, -200.0, 2300.0, 240.0]
        grid[8][1:6] = ["二期合计", 3000.0, -220.0, 2780.0, None]
        grid[9][1:6] = ["合计", 6600.0, -970.0, 5630.0, None]
        found = sp.find_supporting_blocks(pd.DataFrame(grid), [], periods=[P1, P2], multiplier=1000.0)
        self.assertEqual(len(found), 1)
        self.assertNotIn("折旧月数", [c["header"] for c in found[0]["columns"]])        # months were read as thousands of yuan
        labels = [r["label"] for r in found[0]["rows"]]
        self.assertEqual(labels[:6], ["一期土地使用权", "一期房屋建筑物", "一期合计", "二期土地使用权", "二期房屋建筑物", "二期合计"])
        self.assertEqual(labels[-1], "合计")

    def test_a_table_whose_lines_do_not_repeat_keeps_its_labels(self):
        from fdd_utils.workbook import supporting as sp
        grid = [[None] * 6 for _ in range(8)]
        grid[1][1:4] = ["项目", "原值", "净值"]
        grid[2][1:4] = ["房屋", 300.0, 250.0]
        grid[3][1:4] = ["设备", 200.0, 150.0]
        grid[4][1:4] = ["经营合计", 500.0, 400.0]
        found = sp.find_supporting_blocks(pd.DataFrame(grid), [], periods=[P1, P2], multiplier=1000.0)
        self.assertEqual([r["label"] for r in found[0]["rows"]], ["房屋", "设备", "经营合计"])

    def test_a_figure_in_the_header_row_is_no_column_name(self):
        from fdd_utils.workbook import supporting as sp
        grid = [[None] * 6 for _ in range(6)]
        grid[0][0:3] = ["申报表核对", "131393.06", "说明"]
        grid[1][0:3] = ["销项税", 120.0, 1.0]
        grid[2][0:3] = ["进项税", 80.0, 1.0]
        grid[3][0:3] = ["差额", 40.0, 1.0]
        self.assertEqual(sp.find_supporting_blocks(pd.DataFrame(grid), [], periods=[P1, P2]), [])

    def test_a_table_naming_no_unit_and_far_larger_than_the_account_is_read_in_yuan(self):
        # DEMO databook: a working block in yuan under a 千元 schedule read as 4,037亿 for a 4亿 transfer
        from fdd_utils.workbook import supporting as sp
        grid = [[None] * 5 for _ in range(6)]
        grid[0][0:3] = ["在建工程转固", "房屋建筑物", "合计"]
        grid[1][0:3] = ["一期", 403_772_489.57, 403_772_489.57]
        grid[2][0:3] = ["二期", 10_611_725.53, 10_611_725.53]
        frame = pd.DataFrame(grid)
        found = sp.find_supporting_blocks(frame, [], periods=[P1, P2], multiplier=1000.0, scale=450_000_000.0)
        self.assertEqual(len(found), 1)
        self.assertTrue(found[0].get("unit_inferred"))
        self.assertAlmostEqual(max(v for r in found[0]["rows"] for v in r["values"].values()), 403_772_489.57)
        self.assertEqual(sp.find_supporting_blocks(frame, [], periods=[P1, P2], multiplier=1000.0, scale=1_000.0), [])   # not even in yuan

    def test_a_header_or_title_naming_only_a_year_takes_that_years_period(self):
        # a local databook: a pivot of 2020 by tenant was filed as the latest period's figures
        from fdd_utils.workbook import supporting as sp
        grid = [[None] * 5 for _ in range(10)]
        grid[0][0:4] = ["Row Labels", "2024 Total", "2025 Total", "2023 Total"]
        grid[1][0:4] = ["租金收入", 10.0, 12.0, 9.0]
        grid[2][0:4] = ["管理费收入", 5.0, 6.0, 4.0]
        found = sp.find_supporting_blocks(pd.DataFrame(grid), [], periods=[P1, P2])
        self.assertEqual([c["period"] for c in found[0]["columns"]], [P1, P2])      # 2023 is not reported: left out
        grid = [[None] * 5 for _ in range(10)]
        grid[0][0:3] = ["2024", "Sum of 租金", "Sum of 管理费"]
        grid[1][0:3] = ["甲租户", 10.0, 5.0]
        grid[2][0:3] = ["乙租户", 12.0, 6.0]
        found = sp.find_supporting_blocks(pd.DataFrame(grid), [], periods=[P1, P2])
        self.assertTrue(all(c["period"] == P1 and not c["assumed"] for c in found[0]["columns"]))

    def test_an_unfilled_template_is_not_a_table(self):
        from fdd_utils.workbook import supporting as sp
        grid = [[None] * 6 for _ in range(8)]
        grid[0][0:2] = ["其他应收款账龄分析- 2023年ZZ月ZZ日", "人民币千元"]
        grid[1][0:4] = ["客商", "1年以内", "1-2年", "合计"]
        grid[2][0:4] = ["客商1", 10.0, 5.0, 15.0]
        grid[3][0:4] = ["客商2", 20.0, 0.0, 20.0]
        grid[4][0:4] = ["客商3", 1.0, 1.0, 2.0]
        self.assertEqual(sp.find_supporting_blocks(pd.DataFrame(grid), [], periods=[P1, P2]), [])
        grid[0][0] = "其他应收款账龄分析"
        grid[2][0], grid[3][0], grid[4][0] = "甲公司", "乙公司", "丙公司"
        self.assertEqual(len(sp.find_supporting_blocks(pd.DataFrame(grid), [], periods=[P1, P2])), 1)   # negative control

    def test_a_table_that_repeats_the_schedules_own_lines_is_skipped(self):
        c = ec.build_contract("AR", account_frame(clean_rows()), statement_type="BS", financials={P1: 100.0, P2: 100.0})
        copy = [{"title": "复本", "where": "F1:H4", "columns": [{"col": 1, "header": "2025年", "period": P2, "assumed": False}],
                 "rows": [{"label": "甲公司", "row": 2, "values": {1: 70.0}, "total": False},
                          {"label": "乙公司", "row": 3, "values": {1: 30.0}, "total": False}]}]
        self.assertEqual(ec.add_supporting_tables(c, copy), 0)


class RelatedEvidence(unittest.TestCase):
    """What other accounts say about this one, filed on it: the balances and notes of the accounts the report reads with it,
    and any other account's note that names it."""

    def contract(self, account, rows, notes=(), statement="BS"):
        frame = account_frame(rows, sheet=account)
        frame.attrs["supporting_notes"] = list(notes)
        return ec.build_contract(account, frame, statement_type=statement,
                                 financials={P1: rows[-1][2][P1], P2: rows[-1][2][P2]})

    def contracts(self):
        cur = self.contract("其他流动资产", [("待抵扣进项税", "detail", {P1: 600, P2: 640}), ("Total", "total", {P1: 600, P2: 640})])
        non = self.contract("其他非流动资产", [("待认证进项税", "detail", {P1: 100, P2: 111}), ("Total", "total", {P1: 100, P2: 111})],
                            notes=["待认证进项税 | 预计一年以上方可抵扣的增值税进项税额"])
        oth = self.contract("其他应收款", [("借款", "detail", {P1: 50, P2: 50}), ("Total", "total", {P1: 50, P2: 50})],
                            notes=["借款 | 向关联方提供借款，年利率6%，利息计入投资收益"])
        inv = self.contract("投资收益", [("利息收入", "detail", {P1: 3, P2: 3}), ("Total", "total", {P1: 3, P2: 3})], statement="IS")
        ar = self.contract("AR", [("Tenant", "detail", {P1: 9, P2: 9}), ("Total", "total", {P1: 9, P2: 9})])
        yr = self.contract("Other payables", [("Rent for the YEAR", "detail", {P1: 5, P2: 5}), ("Total", "total", {P1: 5, P2: 5})],
                           notes=["Rent for the YEAR | prepaid for the YEAR by the tenant"])
        return {c.account: c for c in (cur, non, oth, inv, ar, yr)}

    def test_a_group_shares_balances_and_notes_and_a_note_naming_an_account_reaches_it(self):
        cs = self.contracts()
        own_refs = {f["ref"] for f in cs["其他流动资产"].facts}
        ec.add_related_evidence(cs)
        cur = cs["其他流动资产"]
        rel = [f for f in cur.facts if f["role"] == ec.ROLE_RELATED]
        self.assertEqual([(f["line"], f["value"]) for f in rel], [("其他非流动资产", 111.0)])
        self.assertTrue(own_refs <= {f["ref"] for f in cur.facts})
        self.assertNotIn(rel[0]["ref"], own_refs)
        self.assertTrue(any(n["kind"] == "related_note" and n["text"].startswith("[其他非流动资产]") for n in cur.notes))
        inv = cs["投资收益"]
        self.assertTrue(any("年利率6%" in n["text"] and n["from_account"] == "其他应收款" for n in inv.notes))
        self.assertFalse([f for f in inv.facts if f["role"] == ec.ROLE_RELATED])      # named in a note, not in its group
        self.assertFalse([n for n in cs["AR"].notes if n["kind"] == "related_note"])  # "AR" is in "YEAR": short keys match nothing

    def test_a_related_note_is_citable_and_never_must_address(self):
        from fdd_utils.ai import claims as cl
        from fdd_utils.agent.controller import AgentController
        cs = self.contracts()
        ec.add_related_evidence(cs)
        inv = cs["投资收益"]
        note = next(n for n in inv.notes if n["kind"] == "related_note")
        claim = {"account": "投资收益", "kind": "cause", "text": "该借款年利率为6%", "periods": [P2], "cites": [], "cause_cites": [note["ref"]]}
        self.assertNotEqual(cl.verify_claim(claim, cs)["verdict"], "failed")
        view = cl.evidence_view(inv)
        bare = AgentController.__new__(AgentController)          # _must_address reads only class-level word lists
        self.assertFalse([m for m in bare._must_address(view) if m["ref"] == note["ref"]])

    def test_related_total_sums_the_same_statement_only_and_is_citable(self):
        from fdd_utils.ai import claims as cl
        cs = self.contracts()
        cs["长期借款"] = self.contract("长期借款", [("银行借款", "detail", {P1: 900, P2: 800}), ("Total", "total", {P1: 900, P2: 800})])
        cs["财务费用"] = self.contract("财务费用", [("利息支出", "detail", {P1: 40, P2: 45}), ("Total", "total", {P1: 40, P2: 45})],
                                   statement="IS")
        ec.add_related_evidence(cs)
        view = cl.evidence_view(cs["其他流动资产"])
        self.assertEqual([(g["period"], g["accounts"]) for g in view["related_total"]], [(P2, ["其他流动资产", "其他非流动资产"])])
        ref = view["related_total"][0]["ref"]
        self.assertNotIn(ref, {d["ref"] for d in view["derived"]})                # shown once, under related
        claim = {"account": "其他流动资产", "kind": "amount", "text": "连同其他非流动资产合计751元", "periods": [P2], "cites": [ref]}
        self.assertEqual(cl.verify_claim(claim, cs)["verdict"], "verified")
        self.assertNotIn("related_total", cl.evidence_view(cs["长期借款"]))         # a balance and an expense are never added

    def test_a_long_note_naming_an_account_is_cut_around_the_name(self):
        cs = self.contracts()
        long = "借款 | " + "该款项为股东借款，期限三年，按季付息。" * 20 + "本金500万美元，年利率6%，利息计入投资收益。"
        cs["其他应收款"] = self.contract("其他应收款", [("借款", "detail", {P1: 50, P2: 50}), ("Total", "total", {P1: 50, P2: 50})],
                                     notes=[long])
        ec.add_related_evidence(cs)
        note = next(n for n in cs["投资收益"].notes if n["kind"] == "related_note")
        self.assertIn("年利率6%", note["text"])
        self.assertLessEqual(len(note["text"]), 300 + 20)


if __name__ == "__main__":
    unittest.main(verbosity=2)

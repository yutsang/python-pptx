"""Claim verification: does the sentence say what the cited facts say? (Phase 3)

Zero tokens, synthetic contracts only:

    python ad-hoc/workbench/test_claims.py

Each case is a way a sentence can be wrong while its numbers look fine -- the
right amount from the wrong period or line, a replaced total quoted as if it were
the tab's, a correct figure carrying an unsupported cause. The last class
contrasts the claim check with the pool search it replaces as the verdict.
"""
from __future__ import annotations

import sys
import unittest
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

from fdd_utils.ai import claims as cl  # noqa: E402
from fdd_utils.ai import evidence_contract as ec  # noqa: E402
from fdd_utils.ai.validator import SourceIndex  # noqa: E402
from fdd_utils.delivery import STATUS_BLOCKED, STATUS_READY, decide_delivery  # noqa: E402
from fdd_utils.workbook.authority import AUTHORITY_KEY  # noqa: E402

P1, P2 = "2024-12-31", "2025-12-31"


def frame(rows, *, sheet="AR", authority=None, periods=(P1, P2), notes=None):
    data = {"Account": [r[0] for r in rows], "__source_row_idx": list(range(4, 4 + len(rows)))}
    for p in periods:
        data[p] = [float(r[2].get(p, 0.0)) for r in rows]
    df = pd.DataFrame(data)
    types = {r[0]: r[1] for r in rows}
    df.attrs.update({"row_types_by_description": dict(types), "source_sheet_name": sheet, "integrity": {"sheet_name": sheet}})
    if notes:
        df.attrs["supporting_notes"] = list(notes)
    analysis = df.copy()
    analysis.attrs.update({"row_types_by_description": dict(types)})
    if authority:
        df.attrs[AUTHORITY_KEY] = authority
        analysis.attrs[AUTHORITY_KEY] = authority
    df.attrs["prompt_analysis_df"] = analysis
    return df


ROWS = [("甲公司", "detail", {P1: 600000, P2: 700000}), ("乙公司", "detail", {P1: 400000, P2: 300000}),
        ("丙公司", "detail", {P1: 0, P2: 300000}), ("Total", "total", {P1: 1000000, P2: 1300000})]
NOTES = ["应收甲公司的租金因装修免租期延迟收取，预计下季度收回"]


def build(rows=ROWS, financials=None, authority=None, notes=NOTES, statement="BS", sheet="AR", periods=(P1, P2), account="应收账款"):
    fin = financials if financials is not None else {p: next(r for r in rows if r[1] == "total")[2][p] for p in periods}
    c = ec.build_contract(account, frame(rows, sheet=sheet, authority=authority, periods=periods, notes=notes),
                          statement_type=statement, financials=fin)
    return c


def ref(c, period, line="", basis=ec.BASIS_ADJUSTED):
    if line:
        return next(f["ref"] for f in c.facts if f["period"] == period and f["line"] == line and f["role"] == "component")
    return next(f["ref"] for f in c.facts if f["period"] == period and f["role"] == "total" and f["basis"] == basis)


def verdict(text, c, cites, periods, kind="amount", **extra):
    claim = {"account": c.account, "kind": kind, "text": text, "periods": periods, "cites": cites, **extra}
    return cl.verify_claim(claim, {c.account: c})


def codes(v):
    return {f["code"] for f in v["findings"] if f["severity"] == "failed"}


class MovementsOverPeriods(unittest.TestCase):
    """A sentence giving a total's movement over several periods: each direction word is bound by the amount written beside
    it. Refused before as unbound (client file, 2026-11-09), so a correct sentence came out marked for a person."""
    P0 = "2023-12-31"

    def setUp(self):
        rows = [("甲公司", "detail", {self.P0: 1000000, P1: 862000, P2: 929000}),
                ("Total", "total", {self.P0: 1000000, P1: 862000, P2: 929000})]
        self.c = build(rows=rows, periods=(self.P0, P1, P2), notes=[])
        cl.standard_derived(self.c)
        self.moves = [d["ref"] for d in self.c.derived if d.get("kind") == cl.KIND_MOVEMENT]

    def test_each_word_takes_the_movement_written_after_it(self):
        v = verdict("应收账款2024年度减少13.8万元，2025年度增加6.7万元。", self.c, self.moves, [self.P0, P1, P2], kind="movement")
        self.assertEqual(v["verdict"], "verified", v["findings"])
        v = verdict("应收账款2024年度减少13.8万元，降幅13.8%；2025年度增加6.7万元。", self.c, self.moves, [self.P0, P1, P2], kind="movement")
        self.assertNotIn("DIRECTION_UNBOUND", {f["code"] for f in v["findings"]}, v["findings"])

    def test_a_word_bound_to_a_movement_of_the_other_sign_is_a_sign_error(self):
        v = verdict("应收账款2024年度增加13.8万元，2025年度增加6.7万元。", self.c, self.moves, [self.P0, P1, P2], kind="movement")
        self.assertIn("SIGN_ERROR", codes(v))

    def test_from_to_figures_give_the_direction(self):
        cites = [ref(self.c, P1), ref(self.c, P2)] + self.moves[1:]
        v = verdict("应收账款由2024年12月31日的86.2万元增加至2025年12月31日的92.9万元，增加6.7万元。", self.c, cites, [P1, P2],
                    kind="movement")
        self.assertNotIn("DIRECTION_UNBOUND", {f["code"] for f in v["findings"]}, v["findings"])
        self.assertNotIn("SIGN_ERROR", codes(v))

    def test_an_excel_text_prefix_is_not_part_of_a_line_name(self):
        self.assertEqual(cl._bare("'*甲物流科技(中国)有限公司"), "甲物流科技(中国)有限公司")


class Correct(unittest.TestCase):
    def test_a_cited_amount_in_the_right_period_is_verified(self):
        c = build()
        v = verdict("截至2025年12月31日，应收账款余额130.0万元。", c, [ref(c, P2)], [P2])
        self.assertEqual(v["verdict"], "verified", v["findings"])
        self.assertEqual({k: s for k, s in v["dimensions"].items() if s != "n/a"},
                         {"citations": "ok", "amounts": "ok", "binding": "ok", "coverage": "ok"})
        self.assertEqual((v["dimensions"]["causal"], v["dimensions"]["arithmetic"]), ("n/a", "n/a"))

    def test_rounding_to_the_printed_precision_is_accepted_and_beyond_it_is_not(self):
        c = build()
        self.assertEqual(verdict("2025年12月31日余额130万元", c, [ref(c, P2)], [P2])["verdict"], "verified")
        bad = verdict("2025年12月31日余额130.1万元", c, [ref(c, P2)], [P2])
        self.assertEqual(codes(bad), {"AMOUNT_UNMATCHED"})

    def test_movement_amount_and_percent_come_from_python(self):
        c = build()
        d = cl.derive_movement(c, P1, P2)
        self.assertEqual((d["value"], round(d["pct"], 4), d["direction"], d["valid"]), (300000.0, 0.3, "up", True))
        v = verdict("较2024年末增加30.0万元，增幅30.0%", c, [d["ref"]], [P1, P2], "movement")
        self.assertEqual(v["verdict"], "verified", v["findings"])

    def test_movement_from_two_cited_totals_is_computed_by_the_verifier(self):
        c = build()
        v = verdict("2025年末较2024年末增加30.0万元", c, [ref(c, P1), ref(c, P2)], [P1, P2], "movement")
        self.assertEqual(v["verdict"], "verified", v["findings"])

    def test_naming_lines_requires_citing_them_and_then_the_group_sum_is_allowed(self):
        c = build()
        ok = verdict("2024年末甲公司60.0万元、乙公司40.0万元，合计100.0万元", c,
                     [ref(c, P1, "甲公司"), ref(c, P1, "乙公司")], [P1], "composition")
        self.assertEqual(ok["verdict"], "verified", ok["findings"])


class WrongBinding(unittest.TestCase):
    def test_right_amount_from_the_wrong_period(self):
        c = build()
        v = verdict("2025年12月31日余额100.0万元", c, [ref(c, P2)], [P2])
        self.assertEqual(codes(v), {"AMOUNT_ELSEWHERE"})
        self.assertIn("2024-12-31", v["findings"][0]["message"])      # says where the figure actually is

    def test_text_names_a_period_the_claim_does_not_declare(self):
        c = build()
        v = verdict("2024年末余额130.0万元", c, [ref(c, P2)], [P2])
        self.assertIn("WRONG_PERIOD", codes(v))

    def test_citing_a_fact_from_another_period(self):
        c = build()
        v = verdict("2025年12月31日余额100.0万元", c, [ref(c, P1)], [P2])
        self.assertIn("WRONG_PERIOD", codes(v))

    def test_same_amount_different_line(self):
        c = build()                     # 乙公司 and 丙公司 both 30.0万元 in 2025
        v = verdict("2025年12月31日，其中乙公司30.0万元", c, [ref(c, P2, "丙公司")], [P2], "composition")
        self.assertIn("WRONG_LINE", codes(v))
        self.assertEqual(verdict("2025年12月31日，其中乙公司30.0万元", c, [ref(c, P2, "乙公司")], [P2], "composition")["verdict"],
                         "verified")

    def test_a_line_the_sentence_names_but_never_cites(self):
        c = build()
        # 甲 + 乙 happen to equal the 2024 total, so the amount matches the cited total -- by adjacency, not citation
        v = verdict("2024年末甲公司与乙公司合计100.0万元", c, [ref(c, P1)], [P1], "composition")
        self.assertIn("LINE_NOT_CITED", codes(v))

    def test_unit_slip(self):
        c = build()
        self.assertEqual(codes(verdict("2025年12月31日余额130.0亿元", c, [ref(c, P2)], [P2])), {"UNIT_ERROR"})

    def test_sign_error(self):
        c = build()
        d = cl.derive_movement(c, P1, P2)
        v = verdict("较2024年末减少30.0万元", c, [d["ref"]], [P1, P2], "movement")
        self.assertIn("SIGN_ERROR", codes(v))

    def test_percent_is_checked_against_the_derived_ratio(self):
        c = build()
        d = cl.derive_movement(c, P1, P2)
        self.assertIn("PERCENT_UNMATCHED", codes(verdict("较2024年末增加30.0万元，增幅35.0%", c, [d["ref"]], [P1, P2], "movement")))

    def test_a_percentage_over_a_zero_base_does_not_exist(self):
        rows = [("甲公司", "detail", {P1: 0, P2: 500000}), ("Total", "total", {P1: 0, P2: 500000})]
        c = build(rows)
        d = cl.derive_movement(c, P1, P2)
        self.assertTrue(d["zero_denominator"])
        v = verdict("较2024年末增加50.0万元，增长100.0%", c, [d["ref"]], [P1, P2], "movement")
        self.assertIn("PERCENT_ZERO_DENOMINATOR", codes(v))


class ReplacedTotal(unittest.TestCase):
    """2024: tab total 800k, Financials 1,000k; the deck uses 1,000k; the tab's detail still sums to 800k."""

    def setUp(self):
        rows = [("甲公司", "detail", {P1: 500000, P2: 700000}), ("乙公司", "detail", {P1: 300000, P2: 300000}),
                ("丙公司", "detail", {P1: 0, P2: 300000}), ("Total", "total", {P1: 1000000, P2: 1300000})]
        self.c = build(rows, financials={P1: 1000000.0, P2: 1300000.0},
                       authority={P1: {"tab": 800000.0, "financials": 1000000.0}})

    def test_the_tabs_own_total_is_not_the_deck_figure(self):
        v = verdict("2024年12月31日余额80.0万元", self.c, [ref(self.c, P1, basis=ec.BASIS_RAW)], [P1])
        self.assertIn("WRONG_BASIS", codes(v))
        ok = verdict("2024年12月31日余额100.0万元", self.c, [ref(self.c, P1)], [P1])
        self.assertEqual(ok["verdict"], "verified", ok["findings"])

    def test_the_disagreement_itself_can_be_described(self):
        d = cl.derive_difference(self.c, P1)
        self.assertEqual(d["value"], -200000.0)
        v = verdict("2024年12月31日，账面明细与报表相差20.0万元", self.c, [d["ref"]], [P1], "difference")
        self.assertEqual(v["verdict"], "verified", v["findings"])

    def test_detail_may_not_explain_a_replaced_total(self):
        v = verdict("2024年12月31日余额主要为甲公司50.0万元", self.c, [ref(self.c, P1, "甲公司")], [P1], "composition")
        self.assertIn("DETAIL_NOT_PERMITTED", codes(v))

    def test_a_replaced_period_may_not_be_called_nil(self):
        v = verdict("2024年12月31日该科目无余额", self.c, [ref(self.c, P1)], [P1], "descriptive")
        self.assertIn("NIL_NOT_PERMITTED", codes(v))


class NestedDetail(unittest.TestCase):
    """甲 includes 乙 (a nested layout): the flat sum misses the total and nothing says why."""

    def setUp(self):
        rows = [("甲项目", "detail", {P1: 600000, P2: 700000}), ("乙项目", "detail", {P1: 200000, P2: 300000}),
                ("丙项目", "detail", {P1: 400000, P2: 600000}), ("Total", "total", {P1: 1000000, P2: 1300000})]
        self.c = build(rows)

    def test_composition_on_ambiguous_detail_is_verified_with_a_caveat_not_rejected(self):
        self.assertEqual(self.c.permissions[P2]["detail_basis"], "ambiguous")
        v = verdict("2025年12月31日余额主要包括甲项目70.0万元、丙项目60.0万元", self.c,
                    [ref(self.c, P2, "甲项目"), ref(self.c, P2, "丙项目")], [P2], "composition")
        self.assertEqual(v["verdict"], "verified", v["findings"])
        self.assertEqual(v["caveats"], ["DETAIL_NESTING_UNKNOWN"])

    def test_a_line_share_is_a_ratio_and_still_computable(self):
        d = cl.derive_share(self.c, "甲项目", P2)
        self.assertTrue(d["valid"])
        self.assertAlmostEqual(d["pct"], 700000 / 1300000)
        self.assertTrue(d["conditions"])


class Disagreement(unittest.TestCase):
    def test_neither_figure_is_the_balance_when_tab_and_financials_disagree_and_nothing_chose(self):
        c = build(financials={P1: 1000000.0, P2: 1500000.0})        # 2025: tab 130万, Financials 150万, no authority record
        v = verdict("2025年12月31日余额130.0万元", c, [ref(c, P2)], [P2])
        self.assertIn("TOTAL_NOT_DESCRIBABLE", codes(v))
        d = cl.derive_difference(c, P2)
        self.assertEqual(verdict("2025年12月31日账面明细与报表相差20.0万元", c, [d["ref"]], [P2], "difference")["verdict"], "verified")


class Annualisation(unittest.TestCase):
    def setUp(self):
        periods = (P1, P2, "2026-01-31")
        rows = [("租金", "detail", {p: 120000 for p in periods}), ("Total", "total", {p: 120000 for p in periods})]
        self.c = build(rows, statement="IS", sheet="Rev", periods=periods, account="营业收入", notes=None)

    def test_movement_into_a_stub_period_is_not_computed_for_a_flow(self):
        d = cl.derive_movement(self.c, P2, "2026-01-31")
        self.assertFalse(d["valid"])
        v = verdict("2026年1月31日较上年增加0.0万元", self.c, [d["ref"]], [P2, "2026-01-31"], "movement")
        self.assertIn("MOVEMENT_NOT_PERMITTED", codes(v))

    def test_annualised_figure_must_be_labelled_and_is_not_a_forecast(self):
        a = cl.derive_annualized(self.c, "2026-01-31")
        self.assertEqual((a["valid"], a["value"]), (True, 1440000.0))
        self.assertEqual(verdict("2026年1月31日年化收入144.0万元", self.c, [a["ref"]], ["2026-01-31"], "annualized")["verdict"], "verified")
        self.assertIn("ANNUALIZATION_UNLABELLED", codes(verdict("2026年1月31日全年收入144.0万元", self.c, [a["ref"]], ["2026-01-31"], "annualized")))
        self.assertIn("ANNUALIZED_AS_FORECAST", codes(verdict("2026年1月31日年化收入144.0万元，预计全年收入将达到该水平", self.c,
                                                              [a["ref"]], ["2026-01-31"], "annualized")))
        self.assertIn("ANNUALIZATION_UNSUPPORTED", codes(verdict("2026年1月31日全年收入144.0万元", self.c,
                                                                 [ref(self.c, "2026-01-31")], ["2026-01-31"])))

    def test_a_full_year_is_not_annualised(self):
        self.assertFalse(cl.derive_annualized(self.c, P2)["valid"])


class Cause(unittest.TestCase):
    TEXT = "2025年12月31日应收账款余额130.0万元，主要系租金回款延迟所致"

    def test_correct_number_with_an_unsupported_cause_keeps_the_dimensions_apart(self):
        c = build(notes=None)
        v = verdict(self.TEXT, c, [ref(c, P2)], [P2])
        self.assertEqual(v["verdict"], "failed")
        self.assertEqual(v["dimensions"]["amounts"], "ok")                 # the number is fine
        self.assertEqual(v["dimensions"]["causal"], "failed")              # the cause is not
        self.assertEqual(codes(v), {"CAUSE_UNSUPPORTED"})

    def test_marked_as_inference_it_is_allowed_and_marked(self):
        c = build(notes=None)
        v = verdict(self.TEXT, c, [ref(c, P2)], [P2], inference=True)
        self.assertEqual((v["verdict"], v["dimensions"]["causal"]), ("verified_inference", "inference"))

    def test_a_cited_note_that_says_it_supports_the_cause(self):
        c = build()
        n = c.notes[0]["ref"]
        v = verdict("2025年12月31日应收账款余额130.0万元，主要系甲公司租金延迟收取所致", c, [ref(c, P2)], [P2], cause_cites=[n])
        # related and uncontradicted is as far as word overlap can go: the deck marks it, a person reads it
        self.assertEqual(v["verdict"], "verified_inference", v["findings"])
        self.assertEqual(v["dimensions"]["amounts"], "ok")

    def test_a_cited_note_about_something_else_does_not_support_it(self):
        c = build(notes=["应收乙公司的押金已于本期全部退还给客户"])
        v = verdict(self.TEXT, c, [ref(c, P2)], [P2], cause_cites=[c.notes[0]["ref"]])
        self.assertEqual(codes(v), {"CAUSE_NOTE_UNRELATED"})


class Citations(unittest.TestCase):
    def test_an_uncited_figure(self):
        c = build()
        # the figure exists in the contract (the 2025 total) but nothing cites it
        self.assertTrue({"AMOUNT_UNCITED", "AMOUNT_ELSEWHERE"} <= codes(verdict("2025年12月31日余额130.0万元", c, [], [P2])))

    def test_unknown_ref_and_foreign_account(self):
        c = build()
        other = build(account="预付款项", sheet="Prepay")
        self.assertIn("UNKNOWN_REF", codes(verdict("2025年12月31日余额130.0万元", c, ["F9999"], [P2])))
        foreign = other.facts[0]["id"]      # refs are per account; a semantic id names its account
        claim = {"account": c.account, "kind": "amount", "text": "2025年12月31日余额130.0万元", "periods": [P2], "cites": [foreign]}
        v = cl.verify_claim(claim, {c.account: c, other.account: other})
        self.assertIn("WRONG_ACCOUNT", codes(v))

    def test_claim_written_against_an_older_source_version(self):
        c = build()
        v = verdict("2025年12月31日余额130.0万元", c, [ref(c, P2)], [P2], source_version="0000000000000000")
        self.assertIn("STALE_SOURCE", codes(v))

    def test_hidden_sheet_facts_are_not_citable(self):
        hidden = ec.build_contract("应收账款", frame(ROWS), statement_type="BS",
                                   digest={"sheets": {"AR": {"hidden": True}}})
        self.assertEqual(hidden.facts, [])
        v = cl.verify_claim({"account": "应收账款", "text": "2025年12月31日余额130.0万元", "periods": [P2], "cites": ["F1"]},
                            {"应收账款": hidden})
        self.assertIn("UNKNOWN_REF", codes(v))


class AgainstThePoolSearch(unittest.TestCase):
    def test_a_figure_the_pool_would_ground_but_nothing_cited_fails(self):
        c = build()
        pool = SourceIndex([{"value": 1300000.0, "kind": "cell", "sheet": "OTHER", "row_desc": "another account's balance",
                             "col_label": P2, "multiplier": 1.0, "row_idx": 1}])
        self.assertTrue(pool.matches(1300000.0), "legacy: found somewhere in the pool, so 'data-backed'")
        v = verdict("2025年12月31日应收账款余额130.0万元", c, [ref(c, P1)], [P2])
        self.assertEqual(v["verdict"], "failed")


class PeriodBinding(unittest.TestCase):
    """Each amount belongs to the period the sentence puts it under, not just to the claim's set of periods."""

    def two_period_verdict(self, text, **extra):
        c = build()
        return c, verdict(text, c, [ref(c, P1), ref(c, P2)], [P1, P2], "amount", **extra)

    def test_amounts_swapped_between_periods_are_not_verified(self):
        c, v = self.two_period_verdict("2024年末余额130万元，2025年末余额100万元")
        self.assertEqual(v["verdict"], "failed", v["findings"])
        self.assertIn("WRONG_PERIOD", codes(v))
        self.assertIn("swapped", " ".join(f["message"] for f in v["findings"]))

    def test_the_same_two_amounts_under_the_right_periods_are_verified(self):
        c, v = self.two_period_verdict("2024年末余额100万元，2025年末余额130万元")
        self.assertEqual(v["verdict"], "verified", v["findings"])

    def test_the_period_may_follow_the_amount(self):
        c = build()
        self.assertEqual(verdict("余额130万元（2025年12月31日）", c, [ref(c, P2)], [P2])["verdict"], "verified")

    def test_a_figure_that_rose_to_a_value_is_bound_to_the_later_period(self):
        c, v = self.two_period_verdict("应收账款由2024年末的100万元增至130万元")
        self.assertEqual(v["verdict"], "verified", v["findings"])

    def test_two_declared_periods_and_an_amount_with_no_period_in_the_sentence_is_left_for_review(self):
        c, v = self.two_period_verdict("应收账款余额130万元，上年为100万元")
        self.assertNotEqual(v["verdict"], "verified")
        self.assertIn("PERIOD_UNBOUND", {f["code"] for f in v["findings"]})
        self.assertEqual(v["dimensions"]["binding"], "unverified")

    def test_a_single_declared_period_binds_an_amount_the_sentence_leaves_unplaced(self):
        c = build()
        self.assertEqual(verdict("应收账款余额130万元", c, [ref(c, P2)], [P2])["verdict"], "verified")

    def test_swapped_lines_within_one_period_still_fail(self):
        c = build()
        v = verdict("2025年12月31日，乙公司70.0万元、甲公司30.0万元", c, [ref(c, P2, "甲公司"), ref(c, P2, "乙公司")], [P2], "composition")
        self.assertIn("WRONG_LINE", codes(v))


class SignAndChangeBinding(unittest.TestCase):
    """Two gaps a source review found (2026-11-06): a fall equal to the prior balance was read as that balance put under the
    wrong period, and a written minus sign was ignored because amounts were compared as magnitudes."""

    GONE = [("甲公司", "detail", {P1: 1000000, P2: 0}), ("Total", "total", {P1: 1000000, P2: 0})]
    NEG = [("A", "detail", {P1: -500000, P2: -700000}), ("Total", "total", {P1: -500000, P2: -700000})]

    def test_a_fall_equal_to_the_prior_balance_is_the_movement_not_a_misdated_balance(self):
        c = build(self.GONE, notes=[])
        v = verdict("2025年末余额较上年减少100.0万元", c, [ref(c, P1), ref(c, P2)], [P1, P2], "movement")
        self.assertEqual(v["verdict"], "verified", v["findings"])

    def test_the_prior_balance_written_as_the_later_balance_still_fails(self):
        c = build(self.GONE, notes=[])
        v = verdict("2025年末余额100.0万元", c, [ref(c, P1), ref(c, P2)], [P1, P2], "amount")
        self.assertEqual(v["verdict"], "failed", v["findings"])
        self.assertIn("WRONG_PERIOD", codes(v))

    def test_the_movement_in_the_wrong_direction_still_fails(self):
        c = build(self.GONE, notes=[])
        v = verdict("2025年末余额较上年增加100.0万元", c, [ref(c, P1), ref(c, P2)], [P1, P2], "movement")
        self.assertEqual(v["verdict"], "failed", v["findings"])

    def test_a_minus_sign_on_a_positive_balance_fails(self):
        c = build()
        for text in ("截至2025年12月31日，应收账款余额-130.0万元。", "截至2025年12月31日，应收账款余额为负130.0万元。"):
            v = verdict(text, c, [ref(c, P2)], [P2])
            self.assertIn("SIGN_ERROR", codes(v), text)

    def test_a_plus_sign_on_a_negative_balance_fails(self):
        c = build(self.NEG, notes=[], account="未分配利润")
        self.assertIn("SIGN_ERROR", codes(verdict("截至2025年12月31日，未分配利润余额+70.0万元。", c, [ref(c, P2)], [P2])))

    def test_a_negative_balance_with_its_sign_or_as_a_magnitude_is_verified(self):
        c = build(self.NEG, notes=[], account="未分配利润")
        for text in ("截至2025年12月31日，未分配利润余额-70.0万元。", "截至2025年12月31日，未分配利润余额70.0万元。"):
            self.assertEqual(verdict(text, c, [ref(c, P2)], [P2])["verdict"], "verified", text)

    def test_a_minus_written_on_a_fall_is_right_and_on_a_rise_is_wrong(self):
        c = build()                                                      # 100 -> 130
        self.assertIn("SIGN_ERROR", codes(verdict("2025年末较2024年末变动-30.0万元", c, [ref(c, P1), ref(c, P2)], [P1, P2], "movement")))
        c = build(self.GONE, notes=[])
        self.assertEqual(verdict("2025年末较2024年末变动-100.0万元", c, [ref(c, P1), ref(c, P2)], [P1, P2], "movement")["verdict"],
                         "verified")

    def test_an_english_amount_is_held_to_the_precision_it_was_written_with(self):
        # "CNY1.3 million" is right for 1,300,000 and for anything that rounds to it; the half-ulp used to be read in yuan for
        # every English scale word, so a rounded English figure could never verify (found 2026-11-07 by the seeded write)
        rows = [("甲公司", "detail", {P1: 1_000_000, P2: 1_296_400}), ("Total", "total", {P1: 1_000_000, P2: 1_296_400})]
        c = build(rows, notes=[])
        for text in ("As at 31 December 2025, the balance was CNY1.3 million.", "As at 31 December 2025, the balance was CNY1.3m.",
                     "As at 31 December 2025, the balance was CNY1,296.4K."):
            self.assertEqual(verdict(text, c, [ref(c, P2)], [P2])["verdict"], "verified", text)
        self.assertIn("AMOUNT_UNMATCHED", codes(verdict("As at 31 December 2025, the balance was CNY1.4 million.", c, [ref(c, P2)], [P2])))
        from fdd_utils.ai.validator import display_half_ulp
        self.assertEqual((display_half_ulp("CNY1.3 million"), display_half_ulp("CNY835.3K"), display_half_ulp("CNY2.15bn"),
                          display_half_ulp("130.0万元")), (50_000.0, 50.0, 5_000_000.0, 500.0))

    def test_a_dash_after_a_figure_is_a_range_not_a_minus(self):
        self.assertEqual([s["sign"] for s in cl.text_amounts("约100-130万元")], [""])
        self.assertEqual([s["sign"] for s in cl.text_amounts("截至2025-12-31，余额130万元")], [""])


class NotesAndGuards(unittest.TestCase):
    """2026-11-08 ten-case review: note facts could not be stated (so material notes went to unresolved), a nested layout blocked
    every composition, and three meaning problems passed the verifier."""

    NOTES = ["注册资本为7000万美元，已全额实缴，我们已检查至FDI入账凭证",
             "2024年6月18日目标公司将注册资本从15000万美元变更为7000万美元",
             "应付二期项目监理的5%工程尾款"]

    def note_ref(self, c, start):
        return next(n["ref"] for n in c.notes if n["text"].startswith(start))

    def test_an_amount_and_a_date_written_in_a_cited_note_can_be_stated_in_its_currency(self):
        c = build(notes=self.NOTES)
        n2 = self.note_ref(c, "2024年6月18日")
        v = verdict("2024年6月18日，注册资本由15,000万美元变更为7,000万美元。", c, [n2], [P1], "descriptive")
        self.assertEqual(v["verdict"], "verified", v["findings"])
        self.assertIn("NOTE_SOURCED_AMOUNT", {f["code"] for f in v["findings"]})
        wrong_currency = verdict("2024年6月18日，注册资本由15,000万元变更为7,000万元。", c, [n2], [P1], "descriptive")
        self.assertIn("AMOUNT_UNMATCHED", codes(wrong_currency), "a dollar amount is not a yuan amount")
        uncited = verdict("2024年6月18日，注册资本由15,000万美元变更为7,000万美元。", c, [ref(c, P1)], [P1], "descriptive")
        self.assertNotEqual(uncited["verdict"], "verified", "a note's figures need that note cited")

    def test_a_percentage_written_in_a_cited_note_can_be_stated(self):
        c = build(notes=self.NOTES)
        n3 = self.note_ref(c, "应付二期")
        self.assertEqual(verdict("其中包括应付二期项目监理5%的工程尾款。", c, [n3], [P2], "descriptive")["verdict"], "verified")
        self.assertIn("PERCENT_UNMATCHED", codes(verdict("其中包括应付二期项目监理8%的工程尾款。", c, [n3], [P2], "descriptive")))

    def test_a_percentage_change_over_a_negative_base_is_refused_and_the_amounts_are_not(self):
        rows = [("甲", "detail", {P1: -2_049_216, P2: 4_844_347}), ("Total", "total", {P1: -2_049_216, P2: 4_844_347})]
        c = build(rows, notes=[])
        d = cl.derive_movement(c, P1, P2)
        self.assertIn("PERCENT_OVER_SIGN_CHANGE", codes(verdict("2025年较2024年增加689.4万元，增幅336.4%", c, [d["ref"]], [P1, P2], "movement")))
        self.assertEqual(verdict("2025年较2024年增加689.4万元", c, [d["ref"]], [P1, P2], "movement")["verdict"], "verified")

    def test_the_tabs_lines_are_not_the_make_up_of_an_adjusted_balance(self):
        rows = [("甲公司", "detail", {P1: 600000, P2: 700000}), ("乙公司", "detail", {P1: 400000, P2: 600000}),
                ("Total", "total", {P1: 1000000, P2: 1302000})]            # the deck's (adjusted) total; the tab's own is 1,300,000
        c = build(rows, financials={P1: 1000000.0, P2: 1302000.0}, authority={P2: {"tab": 1300000.0, "financials": 1302000.0}})
        cites = [ref(c, P2), ref(c, P2, "甲公司"), ref(c, P2, "乙公司")]
        bad = verdict("2025年12月31日余额130.2万元，由甲公司70.0万元和乙公司60.0万元构成。", c, cites, [P2], "composition")
        self.assertIn("COMPOSITION_BASIS", codes(bad))
        fine = verdict("2025年12月31日余额130.2万元；明细表中甲公司70.0万元、乙公司60.0万元。", c, cites, [P2], "composition")
        self.assertNotIn("COMPOSITION_BASIS", codes(fine))

    def test_internal_names_and_refs_do_not_belong_in_the_text(self):
        c = build()
        for text in ("截至2025年12月31日，应收账款余额130.0万元（F2）。", "截至2025年12月31日，应收账款（authority_adjusted口径）余额130.0万元。"):
            self.assertIn("INTERNAL_TERMS_IN_TEXT", codes(verdict(text, c, [ref(c, P2)], [P2])), text)

    def test_a_month_range_names_the_period_it_ends(self):
        m = cl.period_mentions("2026年1-3月营业成本为1,032.9万元")
        self.assertEqual([(x["kind"], x["year"], x["month"]) for x in m], [("month", 2026, 3)])
        self.assertTrue(cl.mention_matches(m[0], "2026-03-31"))
        self.assertEqual([(x["year"], x["month"]) for x in cl.period_mentions("2026年1至3月")], [(2026, 3)])


class NoteContentRefusals(unittest.TestCase):
    """Refusals from the 2026-11-08 whole-databook run (71 claims): most were the verifier not reading what a cited note says."""

    def note_ref(self, c, start):
        return next(n["ref"] for n in c.notes if n["text"].startswith(start))

    def test_an_adjustments_sign_and_direction_come_from_the_note(self):
        c = build(notes=["「2025」Audit ADJ#4：-133.2万元（对方科目：投资收益）；调整说明：应收借款利息重分类至投资收益"])
        n = self.note_ref(c, "「2025」")
        for text in ("2025年度审计调整-133.2万元，系应收借款利息重分类。", "2025年度审计调整减少133.2万元，系应收借款利息重分类。"):
            v = verdict(text, c, [n], [P2], "descriptive")
            self.assertEqual(v["verdict"], "verified", (text, v["findings"]))
        wrong = verdict("2025年度审计调整增加133.2万元。", c, [n], [P2], "descriptive")
        self.assertNotEqual(wrong["verdict"], "verified", "the note says the adjustment reduced it")

    def test_a_direction_the_note_itself_states_is_backed_by_it(self):
        c = build(notes=["租金 | 出租率先大幅下降后逐渐回升，同期新增租赁面积仅53,975平方米"])
        n = self.note_ref(c, "租金")
        v = verdict("出租率先大幅下降后逐渐回升。", c, [n], [P2], "descriptive")
        self.assertNotIn("DIRECTION_UNSUPPORTED", codes(v), v["findings"])

    def test_a_remarks_period_value_is_a_yuan_figure_of_that_period(self):
        c = build(notes=["租金收入 | 2025-12-31: 4342452.0; Detail 26: 会计角度直线法下形成的应收租金",
                         "房产税 | 2025-12-31 annualised: 6705120.0; Detail 22: 按月计提"])
        n1, n2 = self.note_ref(c, "租金收入"), self.note_ref(c, "房产税")
        v = verdict("2025年12月31日租金收入为434.2万元，系直线法下形成的应收租金。", c, [n1], [P2], "descriptive")
        self.assertEqual(v["verdict"], "verified", v["findings"])
        other = verdict("2024年12月31日租金收入为434.2万元。", c, [n1], [P1], "descriptive")
        self.assertEqual(other["verdict"], "failed", "a remark's value belongs to its own period")
        self.assertTrue(codes(other) & {"AMOUNT_UNMATCHED", "WRONG_PERIOD"}, other["findings"])
        self.assertEqual(verdict("2025年度房产税年化金额为670.5万元。", c, [n2], [P2], "descriptive")["verdict"], "verified")

    def test_a_year_the_note_names_without_the_character_is_still_its_year(self):
        c = build(notes=["*某会计师事务所 | 2023审计费191400和2025Q1审计费50000"])
        n = self.note_ref(c, "*某会计师")
        v = verdict("其中包括2023年审计费。", c, [n], [P2], "descriptive")
        self.assertNotIn("WRONG_PERIOD", codes(v), v["findings"])

    def test_naming_the_line_a_cited_note_is_about_needs_no_line_figure(self):
        rows = [("资产管理费", "detail", {P1: 400000, P2: 500000}), ("折旧", "detail", {P1: 600000, P2: 800000}),
                ("Total", "total", {P1: 1000000, P2: 1300000})]
        c = build(rows, notes=["「2025」ADJ#1：+180.7万元（对方科目：毛利）；调整说明：资产管理费截止性调整"])
        n = self.note_ref(c, "「2025」")
        v = verdict("2025年度包含资产管理费截止性调整增加180.7万元。", c, [n], [P2], "descriptive")
        self.assertEqual(v["verdict"], "verified", v["findings"])

    def test_a_counterparty_line_with_a_star_is_the_line_the_sentence_names(self):
        rows = [("押金及保证金", "detail", {P1: 900000, P2: 900000}), ("*甲公司", "detail", {P1: 600000, P2: 600000}),
                ("*乙公司", "detail", {P1: 300000, P2: 300000}), ("Total", "total", {P1: 900000, P2: 900000})]
        c = build(rows, notes=[])
        cites = [ref(c, P2, "押金及保证金"), ref(c, P2, "*甲公司")]
        v = verdict("2025年12月31日余额主要为押金及保证金90.0万元，其中甲公司60.0万元。", c, cites, [P2], "composition")
        self.assertNotIn("WRONG_LINE", codes(v), v["findings"])
        bad = verdict("2025年12月31日余额主要为押金及保证金90.0万元，其中乙公司60.0万元。", c, cites, [P2], "composition")
        self.assertEqual(bad["verdict"], "failed", "a figure under another line still fails")


class Nesting(unittest.TestCase):
    """A line listed beside its own sub-lines: found when the line equals the sum of the lines right under it, and then the top
    level is what ties to the total. Before, every such account had its composition blocked."""

    ROWS = [("甲工程", "detail", {P1: 600000, P2: 700000}), ("一期尾款", "detail", {P1: 100000, P2: 100000}),
            ("二期尾款", "detail", {P1: 500000, P2: 600000}), ("物管费", "detail", {P1: 400000, P2: 600000}),
            ("Total", "total", {P1: 1000000, P2: 1300000})]

    def test_a_parent_and_its_children_tie_at_the_top_level_and_composition_is_permitted(self):
        c = build(self.ROWS, notes=[])
        self.assertEqual(c.permissions[P2]["detail_basis"], "tied")
        self.assertTrue(c.permissions[P2]["explain_with_detail"])
        check = c.check(P2, ec.CHECK_DETAIL_VS_TOTAL)
        self.assertIn("nested", check.get("note", ""))
        v = verdict("2025年12月31日余额主要为甲工程70.0万元和物管费60.0万元", c, [ref(c, P2, "甲工程"), ref(c, P2, "物管费")], [P2], "composition")
        self.assertEqual(v["verdict"], "verified", v["findings"])

    def test_without_an_exact_sub_total_the_layout_stays_unknown(self):
        rows = [(n, t, dict(v)) for n, t, v in self.ROWS]
        rows[2] = ("二期尾款", "detail", {P1: 499000, P2: 599000})
        c = build(rows, notes=[])
        self.assertEqual(c.permissions[P2]["detail_basis"], "ambiguous")

    def test_two_levels_category_counterparty_contract_tie_at_the_top_and_each_line_names_its_own_parent(self):
        # 工程款 = 甲公司 + 乙公司; 甲公司 = 一期 + 二期 (category -> counterparty -> contract)
        rows = [("工程款", "detail", {P1: 900000, P2: 1000000}),
                ("甲公司", "detail", {P1: 600000, P2: 700000}),
                ("一期", "detail", {P1: 100000, P2: 100000}),
                ("二期", "detail", {P1: 500000, P2: 600000}),
                ("乙公司", "detail", {P1: 300000, P2: 300000}),
                ("物管费", "detail", {P1: 100000, P2: 300000}),
                ("Total", "total", {P1: 1000000, P2: 1300000})]
        c = build(rows, notes=[])
        self.assertEqual(c.permissions[P2]["detail_basis"], "tied", c.check(P2, ec.CHECK_DETAIL_VS_TOTAL))
        by_line = {f["line"]: f for f in c.facts if f["period"] == P2 and f["role"] == ec.ROLE_COMPONENT}
        self.assertEqual(by_line["一期"]["nested_in"], by_line["甲公司"]["ref"])
        self.assertEqual(by_line["甲公司"]["nested_in"], by_line["工程款"]["ref"])
        self.assertEqual(by_line["乙公司"]["nested_in"], by_line["工程款"]["ref"])
        self.assertNotIn("nested_in", by_line["工程款"])
        self.assertNotIn("nested_in", by_line["物管费"])


class MaterialView(unittest.TestCase):
    """The writer's detail: lines of 1% or more, lines a note names, never fewer than eight, at most thirty; the rest summed as a
    citable D ref. Until 2026-11-09 it was the eight largest per period, and small special items (a retention) never reached it."""

    def rows(self, n_small=40):
        big = [(f"大客户{i}", "detail", {P1: 900000.0, P2: 1000000.0}) for i in range(5)]
        small = [(f"小户{i:02d}", "detail", {P1: 100.0, P2: 200.0}) for i in range(n_small)]
        special = [("质保金甲", "detail", {P1: 50.0, P2: 50.0})]     # smaller than every filler line
        rows = big + small + special
        rows.append(("Total", "total", {p: sum(r[2][p] for r in rows) for p in (P1, P2)}))
        return rows

    def test_material_and_note_named_lines_are_shown_and_the_rest_is_one_citable_sum(self):
        c = build(self.rows(), notes=["质保金甲 为工程质保金，待竣工验收后支付"])
        v = cl.evidence_view(c)
        names = [r["line"] for r in v["components"]["lines"]]
        for i in range(5):
            self.assertIn(f"大客户{i}", names)
        self.assertIn("质保金甲", names)                       # tiny, but a note names it
        self.assertEqual(len(names), 9)                          # five material, three filler, and the retention a note names
        rest = v["components"]["rest"]
        self.assertEqual(rest["lines"], 46 - len(names))
        d = rest["refs"][-1]
        self.assertTrue(d.startswith("D"))
        shown_small = sum(r["wan"][-1] for r in v["components"]["lines"] if r["line"].startswith("小户"))
        self.assertAlmostEqual(rest["wan"][-1], round((40 * 200.0) / 1e4 - shown_small, 2), places=2)
        self.assertEqual(v["material_not_shown"], 0)
        self.assertNotIn(d, {x["ref"] for x in v["derived"]})    # listed once, under rest
        w = verdict(f"其余{rest['lines']}项合计{rest['wan'][-1]:.2f}万元", c, [d], [P2], "composition")
        self.assertNotIn(w["verdict"], ("failed",), w["findings"])

    def test_the_thirty_line_cap_is_what_raises_evidence_cut(self):
        rows = [(f"客户{i:02d}", "detail", {P1: 1000.0, P2: 1000.0}) for i in range(40)]
        rows.append(("Total", "total", {P1: 40000.0, P2: 40000.0}))
        v = cl.evidence_view(build(rows, notes=[]))
        self.assertEqual(len(v["components"]["lines"]), 30)
        self.assertEqual(v["material_not_shown"], 10)           # each is 2.5% of the total


class DeckColours(unittest.TestCase):
    """The pipeline's meaning of the colours: red for a figure the databook contradicts, orange for an inference, an unproven
    cause, and a verified sentence a flag says a person should read."""

    def test_red_only_for_contradicted_figures_and_orange_for_flags_and_causes(self):
        c = build()
        good = {"account": c.account, "kind": "amount", "text": "2025年12月31日余额130.0万元", "periods": [P2], "cites": [ref(c, P2)]}
        wrong = dict(good, text="2025年12月31日余额150.0万元")
        cause = {"account": c.account, "kind": "cause", "text": "增加主要系回款放缓所致", "periods": [P2], "cites": [], "cause_cites": []}
        noted = dict(good, text="2025年12月31日余额130.0万元。")
        claims = [good, wrong, cause, noted]
        outcome = cl.verify_candidate(claims, {c.account: c}, default_account=c.account)
        self.assertEqual([v["verdict"] for v in outcome["verdicts"]][:3], ["verified", "failed", "failed"])
        flags = [{"kind": "note_amount", "claim": 3, "text": noted["text"]}, {"kind": "deferred_note", "ref": "N1"}]
        reviews = cl.to_legacy_result(c.account, claims, outcome, flags=flags)["agent_4_validation"]["clause_reviews"]
        self.assertEqual([r["category"] for r in reviews], ["data-backed", "hallucination", "reasoning", "reasoning"])
        self.assertEqual(reviews[3]["code"], "NOTE_AMOUNT")
        self.assertEqual([r["category"] for r in cl.to_legacy_result(c.account, claims, outcome)["agent_4_validation"]["clause_reviews"]][3],
                         "data-backed")


class PercentBinding(unittest.TestCase):
    """A percentage is bound to the period, the line and the kind of ratio the sentence says, as an amount is."""

    def setUp(self):
        self.c = build()
        self.s1 = cl.derive_share(self.c, "甲公司", P1)        # 60.0%
        self.s2 = cl.derive_share(self.c, "甲公司", P2)        # 53.8%
        self.o2 = cl.derive_share(self.c, "乙公司", P2)        # 23.1%
        self.mv = cl.derive_movement(self.c, P1, P2)           # +30.0%

    def v(self, text, cites, periods=(P1, P2), kind="share"):
        return verdict(text, self.c, cites, list(periods), kind)

    def test_the_correct_pair_of_shares_is_verified(self):
        v = self.v("2024年末甲公司占比60.0%，2025年末甲公司占比53.8%", [self.s1["ref"], self.s2["ref"]])
        self.assertEqual(v["verdict"], "verified", v["findings"])

    def test_shares_swapped_between_periods_fail(self):
        v = self.v("2024年末甲公司占比53.8%，2025年末甲公司占比60.0%", [self.s1["ref"], self.s2["ref"]])
        self.assertIn("WRONG_PERIOD", codes(v))
        self.assertIn("swapped", " ".join(f["message"] for f in v["findings"]))

    def test_a_share_written_against_the_wrong_line_fails(self):
        v = self.v("2025年末乙公司占比53.8%", [self.s2["ref"], self.o2["ref"]], periods=(P2,))
        self.assertIn("WRONG_LINE", codes(v))
        self.assertEqual(self.v("2025年末乙公司占比23.1%", [self.s2["ref"], self.o2["ref"]], periods=(P2,))["verdict"], "verified")

    def test_a_share_called_a_growth_rate_and_a_growth_rate_called_a_share_fail(self):
        self.assertIn("WRONG_OPERATION", codes(self.v("2025年末甲公司较2024年末增长53.8%", [self.s2["ref"]], periods=(P1, P2))))
        self.assertIn("WRONG_OPERATION", codes(self.v("2025年末甲公司占比30.0%", [self.mv["ref"], self.s2["ref"]], periods=(P1, P2))))
        self.assertEqual(self.v("2025年末较2024年末增幅30.0%", [self.mv["ref"]], kind="movement")["verdict"], "verified")

    def test_no_period_with_two_declared_and_no_line_named_are_left_for_review(self):
        v = self.v("甲公司占比60.0%，另一期53.8%", [self.s1["ref"], self.s2["ref"]])
        self.assertEqual(v["verdict"], "unverified", v["findings"])
        self.assertIn("PERIOD_UNBOUND", {f["code"] for f in v["findings"]})
        v = self.v("2025年末占比53.8%", [self.s2["ref"]], periods=(P2,))
        self.assertEqual(v["verdict"], "unverified", v["findings"])
        self.assertIn("LINE_UNBOUND", {f["code"] for f in v["findings"]})

    def test_a_share_that_fell_to_a_value_binds_to_the_later_period(self):
        v = self.v("甲公司占比由2024年末的60.0%降至53.8%", [self.s1["ref"], self.s2["ref"]])
        self.assertEqual(v["verdict"], "verified", v["findings"])

    def test_the_movement_percent_from_two_cited_totals_is_still_computed_by_the_verifier(self):
        v = verdict("较2024年末增加30.0万元，增幅30.0%", self.c, [ref(self.c, P1), ref(self.c, P2)], [P1, P2], "movement")
        self.assertEqual(v["verdict"], "verified", v["findings"])


class DirectionAndSign(unittest.TestCase):
    """A direction word is checked against the quantity it describes (the amount, or the share), and a
    percentage keeps its sign. Fixtures are built so the amount and the share move in OPPOSITE ways:
      UP  甲公司 60.0万 -> 70.0万 (amount up),   share 60.0% -> 53.8% (share down); 乙公司 amount and share both down
      DN  甲公司 60.0万 -> 50.0万 (amount down), share 60.0% -> 62.5% (share up);   total 100.0万 -> 80.0万 (-20.0%)
    """

    DN_ROWS = [("甲公司", "detail", {P1: 600000, P2: 500000}), ("乙公司", "detail", {P1: 400000, P2: 300000}),
               ("Total", "total", {P1: 1000000, P2: 800000})]

    def setUp(self):
        self.up, self.dn = build(), build(self.DN_ROWS)

    def shares(self, c, line="甲公司"):
        return [cl.derive_share(c, line, P1)["ref"], cl.derive_share(c, line, P2)["ref"]]

    def amounts(self, c, line="甲公司"):
        return [ref(c, P1, line), ref(c, P2, line)]

    def check(self, c, text, cites, kind="share", periods=(P1, P2), **extra):
        return verdict(text, c, cites, list(periods), kind, **extra)

    def failed(self, v, code):
        self.assertEqual(v["verdict"], "failed", v["findings"])
        self.assertIn(code, codes(v))

    def verified(self, v):
        self.assertEqual(v["verdict"], "verified", v["findings"])

    # -- the three cases the reviewer reproduced ----------------------------------------------------------
    def test_a_share_that_fell_is_not_verified_as_rising_and_the_correct_wording_is_not_rejected(self):
        s, a = self.shares(self.up), self.amounts(self.up)
        wrong = "甲公司占比由2024年末的60.0%升至2025年末的53.8%"
        right = "甲公司占比由2024年末的60.0%降至2025年末的53.8%"
        for cites in (s, s + a):            # citing the amounts too is what made the right wording fail
            self.failed(self.check(self.up, wrong, cites), "SIGN_ERROR")
            self.verified(self.check(self.up, right, cites))

    def test_the_second_figure_of_an_elliptical_sentence_is_the_share_too(self):
        s = self.shares(self.up)
        self.failed(self.check(self.up, "2024年末甲公司占比60.0%，2025年末升至53.8%", s + self.amounts(self.up)), "SIGN_ERROR")
        self.verified(self.check(self.up, "2024年末甲公司占比60.0%，2025年末降至53.8%", s + self.amounts(self.up)))

    def test_a_negative_sign_on_a_share_is_an_error(self):
        v = self.check(self.up, "2025年末甲公司占比-53.8%", [self.shares(self.up)[1]], periods=(P2,))
        self.failed(v, "SIGN_ERROR")
        self.assertIn("negative", " ".join(f["message"] for f in v["findings"]))
        self.assertIn("SIGN_ERROR", codes(self.check(self.up, "2025年末甲公司占比－53.8%", [self.shares(self.up)[1]], periods=(P2,))))
        self.verified(self.check(self.up, "2025年末甲公司占比53.8%", [self.shares(self.up)[1]], periods=(P2,)))

    def test_the_level_word_after_a_direction_does_not_turn_a_share_into_a_growth_rate(self):
        s = self.shares(self.up)
        self.verified(self.check(self.up, "甲公司占比由2024年末的60.0%下降至2025年末的53.8%", s))
        self.failed(self.check(self.up, "甲公司占比由2024年末的60.0%上升至2025年末的53.8%", s), "SIGN_ERROR")

    # -- amount and share move in opposite ways -------------------------------------------------------------
    def test_amount_up_and_share_down_each_word_goes_to_its_own_quantity(self):
        both = self.shares(self.up) + self.amounts(self.up)
        amount, share = "甲公司余额由2024年末的60.0万元{a}2025年末的70.0万元", "甲公司占比由2024年末的60.0%{s}2025年末的53.8%"
        for joiner in ("，", "但"):
            self.verified(self.check(self.up, amount.format(a="增至") + joiner + share.format(s="降至"), both, kind="composition"))
            self.failed(self.check(self.up, amount.format(a="降至") + joiner + share.format(s="降至"), both, kind="composition"), "SIGN_ERROR")
            self.failed(self.check(self.up, amount.format(a="增至") + joiner + share.format(s="升至"), both, kind="composition"), "SIGN_ERROR")

    def test_amount_down_and_share_up_each_word_goes_to_its_own_quantity(self):
        both = self.shares(self.dn) + self.amounts(self.dn)
        amount, share = "甲公司余额由2024年末的60.0万元{a}2025年末的50.0万元", "甲公司占比由2024年末的60.0%{s}2025年末的62.5%"
        self.verified(self.check(self.dn, amount.format(a="减至") + "，" + share.format(s="升至"), both, kind="composition"))
        self.failed(self.check(self.dn, amount.format(a="增至") + "，" + share.format(s="升至"), both, kind="composition"), "SIGN_ERROR")
        self.failed(self.check(self.dn, amount.format(a="减至") + "，" + share.format(s="降至"), both, kind="composition"), "SIGN_ERROR")

    def test_a_share_is_not_confirmed_by_the_amount_moving_that_way_nor_the_amount_by_the_share(self):
        s_only, a_only = self.shares(self.up), self.amounts(self.up)
        self.failed(self.check(self.up, "甲公司余额增加", s_only, kind="movement"), "DIRECTION_UNSUPPORTED")
        self.failed(self.check(self.up, "甲公司占比上升", a_only), "DIRECTION_UNSUPPORTED")
        self.failed(self.check(self.up, "甲公司占比下降", [self.shares(self.up)[1]]), "DIRECTION_UNSUPPORTED")

    def test_a_direction_the_sentence_does_not_attach_to_a_quantity_waits_for_a_person_when_the_readings_differ(self):
        both = self.shares(self.up) + self.amounts(self.up)
        v = self.check(self.up, "甲公司较上年下降", both)
        self.assertEqual(v["verdict"], "unverified", v["findings"])
        self.assertIn("DIRECTION_UNBOUND", {f["code"] for f in v["findings"]})
        # both readings agree: nothing to guess
        self.verified(self.check(self.up, "乙公司较上年下降", self.shares(self.up, "乙公司") + self.amounts(self.up, "乙公司")))
        # only one quantity is cited: that is the reading the citations support
        self.verified(self.check(self.up, "甲公司较上年下降", self.shares(self.up)))
        self.failed(self.check(self.up, "甲公司较上年上升", self.shares(self.up)), "SIGN_ERROR")

    def test_two_lines_moving_opposite_ways_are_each_checked_against_their_own_figures(self):
        cites = self.amounts(self.up) + self.amounts(self.up, "乙公司")
        line_a, line_b = "甲公司余额由2024年末的60.0万元增至2025年末的70.0万元", "乙公司余额由2024年末的40.0万元{w}2025年末的30.0万元"
        self.verified(self.check(self.up, line_a + "，" + line_b.format(w="减至"), cites, kind="composition"))
        self.failed(self.check(self.up, line_a + "，" + line_b.format(w="增至"), cites, kind="composition"), "SIGN_ERROR")
        # the named line has no movement cited, the total's does not stand in for it
        total = [ref(self.up, P1), ref(self.up, P2)]
        self.failed(self.check(self.up, "其中甲公司较2024年末减少", total, kind="movement"), "DIRECTION_UNSUPPORTED")

    # -- signed percentages -----------------------------------------------------------------------------------
    def test_a_change_percentage_keeps_its_sign_and_a_drop_of_x_is_not_a_negative_ratio_by_accident(self):
        d = cl.derive_movement(self.dn, P1, P2)                       # -20.0%
        c = lambda text, **kw: self.check(self.dn, text, [d["ref"]], kind="movement", **kw)
        self.verified(c("较2024年末减少20.0万元，降幅20.0%"))
        self.verified(c("较2024年末减少20.0万元，变动幅度-20.0%"))
        self.verified(c("较2024年末减少20.0万元，变动幅度20.0%"))      # a magnitude beside a neutral word
        self.failed(c("较2024年末减少20.0万元，增幅20.0%"), "SIGN_ERROR")
        self.failed(c("较2024年末减少20.0万元，变动幅度+20.0%"), "SIGN_ERROR")
        up = cl.derive_movement(self.up, P1, P2)                      # +30.0%
        self.failed(self.check(self.up, "较2024年末增加30.0万元，变动幅度-30.0%", [up["ref"]], kind="movement"), "SIGN_ERROR")
        self.verified(self.check(self.up, "较2024年末增加30.0万元，增幅30.0%", [up["ref"]], kind="movement"))

    def test_an_unsigned_percentage_that_names_neither_a_sign_nor_a_direction_is_left_for_review_when_it_fell(self):
        d, up = cl.derive_movement(self.dn, P1, P2), cl.derive_movement(self.up, P1, P2)
        v = self.check(self.dn, "2025年末同比20.0%", [d["ref"]], kind="movement")
        self.assertEqual(v["verdict"], "unverified", v["findings"])
        self.assertIn("PERCENT_SIGN_UNSTATED", {f["code"] for f in v["findings"]})
        self.verified(self.check(self.up, "2025年末同比30.0%", [up["ref"]], kind="movement"))      # a rise reads as written

    def test_the_sign_is_read_from_the_text_and_a_range_dash_is_not_a_minus_sign(self):
        sign = lambda text: [p["sign"] for p in cl.text_percents(text)]
        self.assertEqual(sign("较上年-40%"), ["-"])
        self.assertEqual(sign("（-40.5%）"), ["-"])
        self.assertEqual(sign("同比−40%"), ["-"])          # U+2212
        self.assertEqual(sign("同比－40%"), ["-"])          # full-width
        self.assertEqual(sign("增长+5%"), ["+"])
        self.assertEqual(sign("占比40%"), [""])
        self.assertEqual(sign("占比在30%-40%之间"), ["", ""])
        self.assertEqual(sign("占比从30% -40%"), ["", ""])
        self.assertEqual(sign("占比-40%至-30%"), ["-", "-"])

    # -- percentage points --------------------------------------------------------------------------------------
    def test_percentage_points_are_the_difference_of_the_two_shares_and_carry_a_direction(self):
        s = self.shares(self.up, "乙公司")                             # 40.0% -> 23.1%
        self.verified(self.check(self.up, "乙公司占比下降16.9个百分点", s))
        self.failed(self.check(self.up, "乙公司占比上升16.9个百分点", s), "SIGN_ERROR")
        self.failed(self.check(self.up, "乙公司占比下降17.9个百分点", s), "PERCENT_UNMATCHED")
        self.assertNotEqual(self.check(self.up, "乙公司占比下降16.9%", s)["verdict"], "verified")     # points are not percent

    def test_a_long_line_name_with_a_period_between_it_and_the_figure_still_binds(self):
        name = "Provision for bad and doubtful debts and amounts due from related parties"           # 73 characters
        long = build([("甲公司", "detail", {P1: 600000, P2: 700000}), (name, "detail", {P1: 400000, P2: 300000}),
                      ("Total", "total", {P1: 1000000, P2: 1000000})])
        s = [cl.derive_share(long, name, P1)["ref"], cl.derive_share(long, name, P2)["ref"]]
        self.verified(self.check(long, f"{name}占比由2024年末的40.0%降至2025年末的30.0%", s))
        self.failed(self.check(long, f"{name}占比由2024年末的40.0%升至2025年末的30.0%", s), "SIGN_ERROR")
        self.failed(self.check(long, f"甲公司占比由2024年末的40.0%降至2025年末的30.0%", s), "WRONG_LINE")

    # -- a deduction line: its share is negative, and the sentence says which way it reads it -----------------------
    def test_a_deduction_lines_share_is_read_as_its_size_unless_the_sentence_writes_a_minus(self):
        contra = build([("甲公司", "detail", {P1: 600000, P2: 700000}), ("坏账准备", "detail", {P1: -100000, P2: -150000}),
                        ("Total", "total", {P1: 500000, P2: 550000})])                  # 坏账准备 -20.0% -> -27.3%
        s = [cl.derive_share(contra, "坏账准备", P1)["ref"], cl.derive_share(contra, "坏账准备", P2)["ref"]]
        ok = lambda text, cites=s, periods=(P1, P2): self.assertEqual(self.check(contra, text, cites, periods=periods)["verdict"], "verified",
                                                                  self.check(contra, text, cites, periods=periods)["findings"])
        ok("坏账准备占比由2024年末的20.0%升至2025年末的27.3%")                       # its size grew
        self.failed(self.check(contra, "坏账准备占比由2024年末的20.0%降至2025年末的27.3%", s), "SIGN_ERROR")
        ok("坏账准备占比由2024年末的-20.0%降至2025年末的-27.3%")                     # written signed: it fell
        self.failed(self.check(contra, "坏账准备占比由2024年末的-20.0%升至2025年末的-27.3%", s), "SIGN_ERROR")
        ok("2025年末坏账准备占比27.3%", [s[1]], (P2,))
        ok("2025年末坏账准备占比-27.3%", [s[1]], (P2,))
        self.failed(self.check(contra, "2025年末坏账准备占比+27.3%", [s[1]], periods=(P2,)), "SIGN_ERROR")

    # -- words that describe a cause are not the account's direction ----------------------------------------------
    def test_a_direction_word_on_the_cause_side_is_left_to_the_cause_check(self):
        d = cl.derive_movement(self.dn, P1, P2)
        for text in ("应收账款余额较2024年末减少20.0万元，主要系甲公司回款增加所致",
                     "由于甲公司回款增加，应收账款余额较2024年末减少20.0万元",
                     "甲公司回款增加导致应收账款余额较2024年末减少20.0万元"):
            v = self.check(self.dn, text, [d["ref"]], kind="movement", inference=True)
            self.assertNotIn("SIGN_ERROR", codes(v), (text, v["findings"]))
            self.assertEqual(v["verdict"], "verified_inference", (text, v["findings"]))
        flipped = self.check(self.dn, "应收账款余额较2024年末增加20.0万元，主要系甲公司回款增加所致", [d["ref"]], kind="movement", inference=True)
        self.assertIn("SIGN_ERROR", codes(flipped))


class KindAndInferenceAreMetadata(unittest.TestCase):
    """`kind`, `inference` and `cross_account` are written by whoever wrote the sentence. None of them may decide
    which checks a sentence needs or whether a known financial error is let through: the text (what it asserts) and
    the cited evidence do. The same sentence is run under every legal kind, with and without inference=true."""

    KINDS = tuple(cl.CLAIM_KINDS) + ("bogus", "", None)         # an unknown kind is read as 'amount'
    SETTINGS = [(kind, inference) for kind in KINDS for inference in (False, True)]

    def setUp(self):
        self.up, self.dn = build(), build(DirectionAndSign.DN_ROWS)
        rows = [("甲公司", "detail", {P1: 500000, P2: 700000}), ("乙公司", "detail", {P1: 300000, P2: 300000}),
                ("丙公司", "detail", {P1: 0, P2: 300000}), ("Total", "total", {P1: 1000000, P2: 1300000})]
        self.repl = build(rows, financials={P1: 1000000.0, P2: 1300000.0}, authority={P1: {"tab": 800000.0, "financials": 1000000.0}})
        self.split = build(financials={P1: 1000000.0, P2: 1500000.0})          # tab and Financials disagree, nothing chose
        stub = "2026-01-31"
        self.ann = build([("租金", "detail", {p: 120000 for p in (P1, P2, stub)}), ("Total", "total", {p: 120000 for p in (P1, P2, stub)})],
                         statement="IS", sheet="Rev", periods=(P1, P2, stub), account="营业收入", notes=None)
        self.stub = stub

    # -- helpers ---------------------------------------------------------------------------------------------
    def shares(self, c, line="甲公司"):
        return [cl.derive_share(c, line, P1)["ref"], cl.derive_share(c, line, P2)["ref"]]

    def run_all(self, c, text, cites, periods, **extra):
        """{(kind, inference): verdict} for the same sentence."""
        out = {}
        for kind, inference in self.SETTINGS:
            claim = {"account": c.account, "text": text, "periods": list(periods), "cites": cites, "inference": inference, **extra}
            if kind is not None:
                claim["kind"] = kind
            out[(kind, inference)] = cl.verify_claim(claim, {c.account: c})
        return out

    @staticmethod
    def raised(v):
        return {f["code"] for f in v["findings"] if f["severity"] in ("failed", "review")}

    def never_verified(self, c, text, cites, periods, code, **extra):
        results = self.run_all(c, text, cites, periods, **extra)
        for (kind, inference), v in results.items():
            self.assertNotIn(v["verdict"], ("verified", "verified_inference"), (kind, inference, text, v["findings"]))
            self.assertIn(code, self.raised(v), (kind, inference, text, self.raised(v)))
        self.same_for_every_kind(results, text)

    def same_for_every_kind(self, results, text):
        for inference in {i for (_k, i) in results}:
            outcomes = {(v["verdict"], tuple(sorted(self.raised(v)))) for (k, i), v in results.items() if i == inference}
            self.assertEqual(len(outcomes), 1, (inference, text, outcomes))

    def always(self, c, text, cites, periods, verdict_expected, inference=False, **extra):
        results = {key: v for key, v in self.run_all(c, text, cites, periods, **extra).items() if key[1] == inference}
        for key, v in results.items():
            self.assertEqual(v["verdict"], verdict_expected, (key, text, v["findings"]))
        self.same_for_every_kind(results, text)

    # -- known errors stay errors under every kind and with inference=true ---------------------------------------
    def test_a_wrong_direction_of_the_amount_is_not_verified_under_any_kind(self):
        up, dn = cl.derive_movement(self.up, P1, P2)["ref"], cl.derive_movement(self.dn, P1, P2)["ref"]
        self.never_verified(self.up, "应收账款余额较2024年末减少30.0万元", [up], (P1, P2), "SIGN_ERROR")
        self.never_verified(self.dn, "应收账款余额较2024年末增加20.0万元", [dn], (P1, P2), "SIGN_ERROR")
        self.never_verified(self.up, "应收账款较2024年末减少", [up], (P1, P2), "SIGN_ERROR")

    def test_a_wrong_result_clause_is_not_hidden_by_a_cause_clause_or_a_cause_kind(self):
        up = cl.derive_movement(self.up, P1, P2)["ref"]
        self.never_verified(self.up, "应收账款余额较2024年末减少30.0万元，主要系租金回款延迟所致", [up], (P1, P2), "SIGN_ERROR")
        self.never_verified(self.up, "应收账款较2024年末下降，主要系回款增加", [up], (P1, P2), "SIGN_ERROR")
        self.never_verified(self.up, "由于回款增加，应收账款余额较2024年末减少30.0万元", [up], (P1, P2), "SIGN_ERROR")

    def test_a_cause_clause_that_states_a_checkable_direction_of_a_named_line_is_checked_too(self):
        cites = [cl.derive_movement(self.up, P1, P2)["ref"], ref(self.up, P1, "甲公司"), ref(self.up, P2, "甲公司")]
        right = "应收账款余额较2024年末增加30.0万元，主要系甲公司余额增加10.0万元所致"
        self.always(self.up, right, cites, (P1, P2), "verified_inference", inference=True)
        self.never_verified(self.up, right.replace("余额增加10.0万元", "余额减少10.0万元"), cites, (P1, P2), "SIGN_ERROR")

    def test_a_direction_nothing_cited_backs_is_not_verified_under_any_kind(self):
        self.never_verified(self.up, "应收账款余额大幅增加", [], (P1, P2), "DIRECTION_UNSUPPORTED")
        self.never_verified(self.up, "甲公司占比下降", [self.shares(self.up)[1]], (P1, P2), "DIRECTION_UNSUPPORTED")
        self.never_verified(self.up, "甲公司占比上升", [ref(self.up, P1, "甲公司"), ref(self.up, P2, "甲公司")], (P1, P2), "DIRECTION_UNSUPPORTED")

    def test_a_share_and_its_sign_are_checked_under_any_kind(self):
        s = self.shares(self.up)
        self.never_verified(self.up, "甲公司占比由2024年末的60.0%升至2025年末的53.8%", s, (P1, P2), "SIGN_ERROR")
        self.never_verified(self.up, "2025年末甲公司占比-53.8%", [s[1]], (P2,), "SIGN_ERROR")
        self.never_verified(self.up, "2024年末甲公司占比53.8%，2025年末甲公司占比60.0%", s, (P1, P2), "WRONG_PERIOD")
        dn = cl.derive_movement(self.dn, P1, P2)["ref"]
        self.never_verified(self.dn, "较2024年末减少20.0万元，增幅20.0%", [dn], (P1, P2), "SIGN_ERROR")
        self.never_verified(self.dn, "较2024年末减少20.0万元，变动幅度+20.0%", [dn], (P1, P2), "SIGN_ERROR")

    def test_amount_period_unit_and_rounding_errors_are_not_verified_under_any_kind(self):
        t1, t2 = ref(self.up, P1), ref(self.up, P2)
        mv = cl.derive_movement(self.up, P1, P2)["ref"]
        self.never_verified(self.up, "2024年末余额130万元，2025年末余额100万元", [t1, t2], (P1, P2), "WRONG_PERIOD")
        self.never_verified(self.up, "2025年12月31日余额100.0万元", [t2], (P2,), "AMOUNT_ELSEWHERE")
        self.never_verified(self.up, "2025年12月31日余额130.0亿元", [t2], (P2,), "UNIT_ERROR")
        self.never_verified(self.up, "2025年12月31日余额130.1万元", [t2], (P2,), "AMOUNT_UNMATCHED")
        self.never_verified(self.up, "较2024年末增加30.0万元，增幅35.0%", [mv], (P1, P2), "PERCENT_UNMATCHED")

    def test_lines_and_details_are_checked_under_any_kind(self):
        self.never_verified(self.up, "2024年末甲公司与乙公司合计100.0万元", [ref(self.up, P1)], (P1,), "LINE_NOT_CITED")
        self.never_verified(self.up, "2025年末其中乙公司30.0万元，甲公司为主要客户", [ref(self.up, P2, "乙公司")], (P2,), "LINE_NOT_CITED")
        self.never_verified(self.up, "2025年12月31日，其中乙公司30.0万元", [ref(self.up, P2, "丙公司")], (P2,), "WRONG_LINE")
        self.never_verified(self.repl, "2024年12月31日余额主要为甲公司50.0万元", [ref(self.repl, P1, "甲公司")], (P1,), "DETAIL_NOT_PERMITTED")
        self.never_verified(self.repl, "2024年12月31日该科目无余额", [ref(self.repl, P1)], (P1,), "NIL_NOT_PERMITTED")

    def test_a_replaced_or_disputed_total_cannot_be_quoted_as_the_balance_by_calling_the_claim_a_difference(self):
        self.never_verified(self.repl, "2024年12月31日余额80.0万元", [ref(self.repl, P1, basis=ec.BASIS_RAW)], (P1,), "WRONG_BASIS")
        self.never_verified(self.split, "2025年12月31日余额130.0万元", [ref(self.split, P2)], (P2,), "TOTAL_NOT_DESCRIBABLE")

    def test_annualisation_rules_hold_under_any_kind_and_inference_never_excuses_the_figure(self):
        a = cl.derive_annualized(self.ann, self.stub)["ref"]
        self.never_verified(self.ann, "2026年1月31日全年收入144.0万元", [a], (self.stub,), "ANNUALIZATION_UNLABELLED")
        self.never_verified(self.ann, "2026年1月31日年化收入150.0万元，预计全年收入将达到该水平", [a], (self.stub,), "AMOUNT_UNMATCHED")
        self.never_verified(self.ann, "2026年1月31日年化收入14.4万元", [a], (self.stub,), "UNIT_ERROR")
        # what inference does waive: a linear annualisation worded as a forecast, with the figure itself still checked
        text = "2026年1月31日年化收入144.0万元，预计全年收入将达到该水平"
        self.assertIn("ANNUALIZED_AS_FORECAST", self.raised(self.run_all(self.ann, text, [a], (self.stub,))[("amount", False)]))
        self.assertNotIn("ANNUALIZED_AS_FORECAST", self.raised(self.run_all(self.ann, text, [a], (self.stub,))[("amount", True)]))

    def test_the_disagreement_itself_can_be_described_under_any_kind(self):
        d = cl.derive_difference(self.repl, P1)["ref"]
        self.always(self.repl, "2024年12月31日，账面明细与报表相差20.0万元", [d], (P1,), "verified")
        d2 = cl.derive_difference(self.split, P2)["ref"]
        self.always(self.split, "2025年12月31日账面明细与报表相差20.0万元", [d2], (P2,), "verified")

    def test_a_cause_with_no_note_needs_inference_to_be_marked_not_to_be_checked(self):
        c = build(notes=None)
        text = "2025年12月31日应收账款余额130.0万元，主要系租金回款延迟所致"
        self.always(c, text, [ref(c, P2)], (P2,), "failed")
        self.always(c, text, [ref(c, P2)], (P2,), "verified_inference", inference=True)
        # but the number under the marked cause is still the number
        self.never_verified(c, text.replace("130.0万元", "100.0万元"), [ref(c, P2)], (P2,), "AMOUNT_ELSEWHERE")
        self.never_verified(c, text.replace("130.0万元", "130.0亿元"), [ref(c, P2)], (P2,), "UNIT_ERROR")

    def test_a_cited_note_makes_the_check_not_the_declared_kind(self):
        note = self.up.notes[0]["ref"]
        text = "2025年12月31日应收账款余额130.0万元，主要系甲公司租金延迟收取所致"
        self.always(self.up, text, [ref(self.up, P2)], (P2,), "verified_inference", cause_cites=[note])
        self.never_verified(self.up, text.replace("130.0万元", "100.0万元"), [ref(self.up, P2)], (P2,), "AMOUNT_ELSEWHERE", cause_cites=[note])

    # -- a correct sentence stays correct whatever it is labelled ----------------------------------------------
    def test_correct_sentences_are_verified_under_every_kind(self):
        s, a = self.shares(self.up), [ref(self.up, P1, "甲公司"), ref(self.up, P2, "甲公司")]
        dn = cl.derive_movement(self.dn, P1, P2)["ref"]
        self.always(self.up, "2025年12月31日，应收账款余额130.0万元", [ref(self.up, P2)], (P2,), "verified")
        self.always(self.up, "甲公司占比由2024年末的60.0%降至2025年末的53.8%", s, (P1, P2), "verified")
        self.always(self.up, "甲公司余额由2024年末的60.0万元增至2025年末的70.0万元，甲公司占比由2024年末的60.0%降至2025年末的53.8%",
                    s + a, (P1, P2), "verified")
        self.always(self.dn, "较2024年末减少20.0万元，变动幅度-20.0%", [dn], (P1, P2), "verified")
        self.always(self.dn, "应收账款余额较2024年末减少20.0万元，降幅20.0%", [dn], (P1, P2), "verified")

    def test_correct_cause_sentences_stay_correct_whatever_the_kind(self):
        """The result clause is checked against the movement, the cause clause is left to the cause check."""
        dn, up = cl.derive_movement(self.dn, P1, P2)["ref"], cl.derive_movement(self.up, P1, P2)["ref"]
        for text in ("应收账款余额较2024年末减少20.0万元，主要系甲公司回款增加所致",
                     "由于甲公司回款增加，应收账款余额较2024年末减少20.0万元",
                     "甲公司回款增加导致应收账款余额较2024年末减少20.0万元"):
            self.always(self.dn, text, [dn], (P1, P2), "verified_inference", inference=True)
            self.always(self.dn, text, [dn], (P1, P2), "failed")                        # no note, not marked
        self.always(self.up, "应收账款余额较2024年末增加30.0万元，主要系租金回款延迟所致", [up], (P1, P2), "verified_inference", inference=True)

    def test_a_direction_that_cannot_be_placed_on_either_side_of_a_cause_waits_for_a_person(self):
        dn = cl.derive_movement(self.dn, P1, P2)["ref"]
        for text in ("应收账款余额较2024年末减少20.0万元，主要系甲公司回款增加，坏账风险上升",
                     "应收账款余额较2024年末减少20.0万元，导致坏账风险上升"):
            results = self.run_all(self.dn, text, [dn], (P1, P2), inference=True)
            for key, v in results.items():
                if key[1]:
                    self.assertEqual(v["verdict"], "unverified", (key, text, v["findings"]))
                    self.assertIn("DIRECTION_UNBOUND", self.raised(v))
            self.same_for_every_kind({k: v for k, v in results.items() if k[1]}, text)

    # -- the other things a sentence author can declare ------------------------------------------------------------
    def test_declaring_a_claim_cross_account_does_not_make_another_accounts_figure_this_accounts(self):
        other = build([("丁公司", "detail", {P1: 500000, P2: 900000}), ("Total", "total", {P1: 500000, P2: 900000})], account="其他应收款")
        foreign = next(f["id"] for f in other.facts if f["period"] == P2 and f["role"] == "total" and f["basis"] == ec.BASIS_ADJUSTED)
        for flag in (False, True, "true"):
            for kind in self.KINDS:
                claim = {"account": self.up.account, "kind": kind or "amount", "text": "2025年12月31日余额90.0万元", "periods": [P2],
                         "cites": [foreign], "cross_account": flag}
                v = cl.verify_claim(claim, {self.up.account: self.up, other.account: other})
                self.assertNotIn(v["verdict"], ("verified", "verified_inference"), (flag, kind, v["findings"]))
                self.assertTrue({"WRONG_ACCOUNT", "CROSS_ACCOUNT_UNBOUND"} & self.raised(v), (flag, kind, v["findings"]))

    def test_a_flag_written_as_text_is_not_true_just_because_it_is_not_empty(self):
        c = build(notes=None)
        text = "2025年12月31日应收账款余额130.0万元，主要系租金回款延迟所致"
        for spelled in ("false", "False", "0", "no", "", None):
            claim = {"account": c.account, "kind": "amount", "text": text, "periods": [P2], "cites": [ref(c, P2)], "inference": spelled}
            self.assertEqual(cl.verify_claim(claim, {c.account: c})["verdict"], "failed", spelled)
        claim = {"account": c.account, "kind": "amount", "text": text, "periods": [P2], "cites": [ref(c, P2)], "inference": "true"}
        self.assertEqual(cl.verify_claim(claim, {c.account: c})["verdict"], "verified_inference")


class CauseIsNotJustRelevance(unittest.TestCase):
    """A note that shares words with a cause is relevant, not supporting. Contradictions are caught; the rest is marked."""

    NOTE = ["应收甲公司的租金因装修免租期延迟收取，预计下季度收回"]

    def cause(self, text, notes=None):
        c = build(notes=notes or self.NOTE)
        return verdict("2025年12月31日应收账款余额130.0万元，" + text, c, [ref(c, P2)], [P2], cause_cites=[c.notes[0]["ref"]])

    def test_a_related_note_makes_the_cause_marked_not_verified(self):
        v = self.cause("主要系甲公司因装修免租期延迟收取租金所致")
        self.assertEqual(v["verdict"], "verified_inference")
        self.assertEqual(v["dimensions"]["causal"], "related_unchecked")
        self.assertIn("CAUSE_NOTE_RELATED_UNCHECKED", {f["code"] for f in v["findings"]})
        self.assertNotIn("CAUSE_SUPPORTED", {f["code"] for f in v["findings"]})

    def test_the_note_says_expected_and_the_claim_says_done(self):
        v = self.cause("因装修免租期，应收甲公司的租金已全部收回")
        self.assertIn("CAUSE_CONTRADICTS_NOTE", codes(v))

    def test_negation_the_note_does_not_have(self):
        v = self.cause("并非因装修免租期，甲公司的租金未延迟收取")
        self.assertIn("CAUSE_CONTRADICTS_NOTE", codes(v))

    def test_all_against_part(self):
        v = self.cause("因装修免租期，应收甲公司的全部租金延迟收取", notes=["应收甲公司的部分租金因装修免租期延迟收取"])
        self.assertIn("CAUSE_CONTRADICTS_NOTE", codes(v))

    def test_a_different_subject(self):
        v = self.cause("因装修免租期，应收乙公司的租金延迟收取")
        self.assertIn("CAUSE_CONTRADICTS_NOTE", codes(v))

    def test_an_opposite_direction(self):
        v = self.cause("因装修免租期，甲公司的租金提前收取", notes=["应收甲公司的租金因装修免租期延迟收取"])
        self.assertIn("CAUSE_CONTRADICTS_NOTE", codes(v))

    def test_a_cause_with_no_note_still_fails_and_a_marked_inference_still_passes_marked(self):
        c = build(notes=None)
        self.assertIn("CAUSE_UNSUPPORTED", codes(verdict("2025年12月31日余额130.0万元，主要系市场需求所致", c, [ref(c, P2)], [P2])))


class DimensionsAreNotOkByDefault(unittest.TestCase):
    def test_a_dimension_nothing_checked_is_not_applicable_not_ok(self):
        c = build()
        v = verdict("2025年12月31日应收账款余额130.0万元", c, [ref(c, P2)], [P2])
        self.assertEqual(v["dimensions"]["amounts"], "ok")
        self.assertEqual(v["dimensions"]["binding"], "ok")
        self.assertEqual(v["dimensions"]["causal"], "n/a")
        self.assertEqual(v["dimensions"]["arithmetic"], "n/a")
        pure = verdict("应收账款主要为租户款项", c, [], [P2], "descriptive")
        self.assertEqual((pure["dimensions"]["amounts"], pure["dimensions"]["coverage"]), ("n/a", "n/a"))


class Interop(unittest.TestCase):
    def test_failed_claims_gate_delivery_and_marked_inference_does_not(self):
        c = build(notes=None)
        good = {"account": c.account, "kind": "amount", "text": "2025年12月31日应收账款余额130.0万元", "periods": [P2], "cites": [ref(c, P2)]}
        bad = {"account": c.account, "kind": "amount", "text": "2025年12月31日应收账款余额100.0万元", "periods": [P2], "cites": [ref(c, P2)]}
        marked = dict(good, text=self_text(), inference=True)

        def run(claims):
            outcome = cl.verify_candidate(claims, {c.account: c}, default_account=c.account)
            legacy = {c.account: cl.to_legacy_result(c.account, claims, outcome),
                      "__run_health__": {"zero_successful_calls": False, "accounts_on_fallback_bullet": [],
                                         "accounts_on_passthrough": [], "accounts_on_error_text": [], "accounts_without_prompt": []}}
            return outcome, decide_delivery(ai_results=legacy, bs_is_results={"statement_source": "financials"})

        outcome, delivery = run([good, bad])
        self.assertFalse(outcome["ok"])
        self.assertEqual(delivery.status, STATUS_BLOCKED)
        outcome, delivery = run([good, marked])
        self.assertTrue(outcome["ok"])
        self.assertEqual(delivery.status, STATUS_READY)
        self.assertEqual(outcome["inference"], 1)

    def test_evidence_view_carries_refs_permissions_and_notes(self):
        c = build()
        cl.standard_derived(c)
        view = cl.evidence_view(c)
        self.assertTrue(view["totals"] and view["derived"] and view["notes"])
        self.assertEqual(view["source_version"], c.source_version)
        self.assertIn(P2, view["permissions"])


def self_text():
    return "2025年12月31日应收账款余额130.0万元，主要系租金回款延迟所致"


if __name__ == "__main__":
    unittest.main(verbosity=2)

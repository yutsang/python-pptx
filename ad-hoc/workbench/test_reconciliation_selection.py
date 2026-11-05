"""Which accounts reconciliation sends to commentary (Phase 1, selection half).

Zero tokens, a few seconds, synthetic frames only:

    python ad-hoc/workbench/test_reconciliation_selection.py

`_legacy_derive` below is the function as it stood before workbook/selection.py,
kept verbatim as a NEGATIVE CONTROL: every hole case must show the legacy
function admitting the account and the current one refusing it. A suite that
passed against the legacy function would not be testing the hole.
"""
from __future__ import annotations

import sys
import unittest
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

from fdd_utils.ui.views import derive_reconciliation_matched_keys, explain_reconciliation_selection  # noqa: E402
from fdd_utils.workbook.selection import (  # noqa: E402
    REASON_BS_DIFF, REASON_NO_ACTIVITY, REASON_NO_DATA, REASON_NO_ROW,
    ROUTE_EXCLUDED, ROUTE_POLICY, ROUTE_PRIOR_PERIOD, ROUTE_RECONCILED, ROUTE_UNCHECKED_HISTORY,
    select_reconciliation_keys, selection_without_financials,
)

MAPPINGS = {
    "其他应收款": {"aliases": ["其他应收款"], "category": "Current assets", "type": "BS"},
    "应收账款": {"aliases": ["应收账款"], "category": "Current assets", "type": "BS"},
    "货币资金": {"aliases": ["货币资金"], "category": "Current assets", "type": "BS"},
    "营业外支出": {"aliases": ["营业外支出"], "category": "Expenses", "type": "IS"},
}

COLUMNS = ["Financials_Account", "Date", "Financials_Value", "Tab_Account", "Tab_Value", "Diff",
           "Match", "Mapping_Key", "Mapping_Status", "Mapping_Note"]


def recon(*rows):
    out = []
    for account, tab, match, status in rows:
        out.append([account, "2025-12-31", 100.0, tab, 100.0, "-", match, "-", status, ""])
    return pd.DataFrame(out, columns=COLUMNS)


def tab(active: bool) -> pd.DataFrame:
    frame = pd.DataFrame({"Description": ["x", "合计"], "2025-12-31": [0.0, 0.0]})
    frame.attrs["any_period_nonzero_by_description"] = {"x": active, "合计": active}
    return frame


def _legacy_derive(reconciliation, available_keys, resolution=None, dfs=None):
    """Verbatim from ui/views.py before the selection module (negative control)."""
    available_key_order = [k for k in dict.fromkeys(str(k).strip() for k in available_keys) if k]
    available_key_set = set(available_key_order)
    if not reconciliation or not available_key_order:
        return []
    bs_included = {"✅ Match", "⚠️ Match", "✅ Immaterial"}
    is_included = {"✅ Match", "⚠️ Match", "✅ Immaterial", "❌ Diff"}
    matched, seen = [], set()
    for idx, frame in enumerate(reconciliation):
        if frame is None or frame.empty:
            continue
        included = is_included if idx == 1 else bs_included
        for value in frame[frame["Match"].isin(included)]["Tab_Account"].tolist():
            key = str(value or "").strip()
            if not key or key == "-" or key in seen or key not in available_key_set:
                continue
            seen.add(key)
            matched.append(key)
    if dfs:
        for key in available_key_order:
            if key in seen:
                continue
            attrs = getattr(dfs.get(key), "attrs", {}) or {}
            any_period = attrs.get("any_period_nonzero_by_description") or {}
            if isinstance(any_period, dict) and any(bool(v) for v in any_period.values()):
                seen.add(key)
                matched.append(key)
    return matched


def select(bs, is_, keys, dfs):
    return explain_reconciliation_selection((bs, is_), keys, None, dfs=dfs)


class HoleCases(unittest.TestCase):
    """Each case is a way the old fallback let an account in that it should not."""

    def test_bs_diff_with_history_is_excluded(self):
        bs = recon(("应收账款", "应收账款", "❌ Diff", "Mapped"))
        dfs = {"应收账款": tab(True)}
        result = select(bs, None, ["应收账款"], dfs)
        self.assertEqual(result.admitted_keys, [])
        decision = result.by_key()["应收账款"]
        self.assertEqual((decision.route, decision.reason), (ROUTE_EXCLUDED, REASON_BS_DIFF))
        # negative control: the legacy function admits it
        self.assertEqual(_legacy_derive((bs, None), ["应收账款"], dfs=dfs), ["应收账款"])

    def test_account_with_no_reconciliation_row_is_not_added_back(self):
        bs = recon(("货币资金", "货币资金", "✅ Match", "Mapped"))
        dfs = {"货币资金": tab(True), "关联方": tab(True)}
        result = select(bs, None, ["货币资金", "关联方"], dfs)
        self.assertEqual(result.admitted_keys, ["货币资金"])
        decision = result.by_key()["关联方"]
        self.assertEqual((decision.route, decision.reason), (ROUTE_EXCLUDED, REASON_NO_ROW))
        self.assertEqual(_legacy_derive((bs, None), ["货币资金", "关联方"], dfs=dfs), ["货币资金", "关联方"])

    def test_bs_diff_in_one_account_does_not_hide_others(self):
        bs = recon(("应收账款", "应收账款", "❌ Diff", "Mapped"), ("货币资金", "货币资金", "✅ Match", "Mapped"))
        dfs = {"应收账款": tab(True), "货币资金": tab(True)}
        self.assertEqual(derive_reconciliation_matched_keys((bs, None), ["应收账款", "货币资金"], None, dfs=dfs),
                         ["货币资金"])


class AdmissionCases(unittest.TestCase):
    def test_nil_latest_with_history_is_admitted_as_unchecked(self):
        # The Financials line is nil in both recent periods, so reconciliation wrote a
        # 'Zero source' row with no tab; the tab is found through the mappings.
        bs = recon(("其他应收款", "-", "-", "Zero source"))
        dfs = {"其他应收款": tab(True)}
        result = select_reconciliation_keys((bs, None), ["其他应收款"], None, dfs=dfs, mappings=MAPPINGS)
        decision = result.by_key()["其他应收款"]
        self.assertTrue(decision.admitted)
        self.assertEqual(decision.route, ROUTE_UNCHECKED_HISTORY)
        self.assertTrue(decision.needs_review)

    def test_nil_latest_without_any_history_stays_out(self):
        bs = recon(("其他应收款", "-", "-", "Zero source"))
        result = select_reconciliation_keys((bs, None), ["其他应收款"], None, dfs={"其他应收款": tab(False)},
                                            mappings=MAPPINGS)
        decision = result.by_key()["其他应收款"]
        self.assertEqual((decision.admitted, decision.reason), (False, REASON_NO_ACTIVITY))

    def test_missing_frame_is_a_named_exclusion(self):
        bs = recon(("货币资金", "货币资金", "✅ Match", "Mapped"))
        result = select(bs, None, ["货币资金", "孤儿"], {"货币资金": tab(True)})
        self.assertEqual(result.by_key()["孤儿"].reason, REASON_NO_DATA)

    def test_interim_diff_is_admitted_for_review(self):
        bs = recon(("应收账款", "应收账款", "⚠️ Interim Diff (prior period matches)", "Mapped"))
        result = select(bs, None, ["应收账款"], {"应收账款": tab(True)})
        decision = result.by_key()["应收账款"]
        self.assertEqual((decision.admitted, decision.route), (True, ROUTE_PRIOR_PERIOD))
        self.assertTrue(decision.needs_review)

    def test_not_found_with_a_tab_is_unchecked_not_failed(self):
        bs = recon(("其他应收款", "其他应收款", "⚠️ Not Found", "Mapped"))
        result = select(bs, None, ["其他应收款"], {"其他应收款": tab(True)})
        self.assertEqual(result.by_key()["其他应收款"].route, ROUTE_UNCHECKED_HISTORY)

    def test_financials_line_with_no_tab_is_listed_not_invented(self):
        bs = recon(("预收款项", "Not Found", "⚠️ Not Found", "Mapped but missing tab"))
        result = select(bs, None, ["货币资金"], {"货币资金": tab(True)})
        self.assertEqual(result.unmatched_financials, [{"statement": "BS", "financials_account": "预收款项"}])

    def test_is_diff_is_admitted_by_policy(self):
        is_ = recon(("营业外支出", "营业外支出", "❌ Diff", "Mapped"))
        result = select(None, is_, ["营业外支出"], {"营业外支出": tab(True)})
        decision = result.by_key()["营业外支出"]
        self.assertEqual((decision.admitted, decision.route), (True, ROUTE_POLICY))
        self.assertFalse(decision.needs_review)


class Parity(unittest.TestCase):
    """Where the old function had no hole to fall through, nothing changes."""

    def test_same_keys_same_order_when_everything_ties_out(self):
        bs = recon(("应收账款", "应收账款", "✅ Match", "Mapped"), ("货币资金", "货币资金", "✅ Immaterial", "Mapped"))
        is_ = recon(("营业外支出", "营业外支出", "⚠️ Match", "Mapped"))
        keys = ["货币资金", "应收账款", "营业外支出"]
        dfs = {k: tab(True) for k in keys}
        self.assertEqual(derive_reconciliation_matched_keys((bs, is_), keys, None, dfs=dfs),
                         _legacy_derive((bs, is_), keys, dfs=dfs))

    def test_no_reconciliation_returns_empty_and_says_why(self):
        self.assertEqual(derive_reconciliation_matched_keys(None, ["a"], None, dfs={"a": tab(True)}), [])
        self.assertFalse(explain_reconciliation_selection((None, None), ["a"], None).reconciliation_available)

    def test_no_financials_record_flags_every_tab_unchecked(self):
        record = selection_without_financials(["a", "b", "a"])
        self.assertEqual(record.admitted_keys, ["a", "b"])
        self.assertTrue(all(d.needs_review for d in record.decisions))


class Property(unittest.TestCase):
    def test_new_selection_is_never_wider_than_legacy(self):
        """Whatever the statuses and activity, the new function admits a subset of
        what the old one admitted (it may only refuse what the fallback let in)."""
        import itertools
        statuses = ["✅ Match", "⚠️ Match", "✅ Immaterial", "❌ Diff", "⚠️ Interim Diff (prior period matches)",
                    "⚠️ Not Found", None]
        for bs_status, is_status, active_a, active_b in itertools.product(statuses, statuses, (True, False), (True, False)):
            bs_rows = [("a", "a", bs_status, "Mapped")] if bs_status else []
            is_rows = [("b", "b", is_status, "Mapped")] if is_status else []
            bs = recon(*bs_rows) if bs_rows else None
            is_ = recon(*is_rows) if is_rows else None
            if bs is None and is_ is None:
                continue
            dfs = {"a": tab(active_a), "b": tab(active_b)}
            new = set(derive_reconciliation_matched_keys((bs, is_), ["a", "b"], None, dfs=dfs))
            old = set(_legacy_derive((bs, is_), ["a", "b"], dfs=dfs))
            self.assertTrue(new <= old, (bs_status, is_status, active_a, active_b, new, old))


if __name__ == "__main__":
    unittest.main(verbosity=2)

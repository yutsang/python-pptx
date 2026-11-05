"""Which accounts reconciliation lets through to commentary, and why the rest stay out.

``derive_reconciliation_matched_keys`` used to return a bare list. Its last step
re-admitted ANY account whose schedule tab had a non-zero period, on the stated
ground that such an account "was never examined by reconciliation". That is true
of a Financials line that was nil in the two latest periods (reconciliation
skips it), and false of three other cases the same step also let in:

  * a balance-sheet line reconciliation examined and marked ``❌ Diff`` -- the
    very status the first step excludes, so the exclusion held for the latest
    period and was undone by any non-zero history;
  * an account no Financials line names at all (a mapping that did not resolve,
    a tab with no statement line), which reconciliation never saw and which
    nothing says is a statement account;
  * a line reconciliation could not read a tab value for (``⚠️ Not Found``).

Each admitted or refused account now carries a route and a reason, so a refusal
is visible and an admission names what it rests on. The list the old function
returned is the admitted keys of this result, in the same order.

Routes (``Decision.route``):

  reconciled          examined and tied out: ✅ Match, ⚠️ Match, ✅ Immaterial
  reconciled_policy   income statement ``❌ Diff``, admitted by owner policy:
                      its reconciliation is period-flow noisy and excluding it
                      would drop most of the statement
  prior_period        ``⚠️ Interim Diff``: the latest column is a partial
                      period, the previous full period ties out. Admitted, with
                      review (the latest figures were not tied)
  unchecked_history   the Financials line was nil in both recent periods, so
                      reconciliation skipped it, and the tab has activity in an
                      earlier period. Admitted, with review (never tied out)
  excluded            not admitted; ``reason`` says why

Reasons for exclusion: ``bs_diff``, ``no_reconciliation_row``, ``no_activity``,
``no_data``.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, Iterable, List, Optional

import pandas as pd

BS_INCLUDED_STATUSES = frozenset({"✅ Match", "⚠️ Match", "✅ Immaterial"})
IS_INCLUDED_STATUSES = BS_INCLUDED_STATUSES | {"❌ Diff"}
INTERIM_PREFIX = "⚠️ Interim Diff"

ROUTE_RECONCILED = "reconciled"
ROUTE_POLICY = "reconciled_policy"
ROUTE_PRIOR_PERIOD = "prior_period"
ROUTE_UNCHECKED_HISTORY = "unchecked_history"
ROUTE_EXCLUDED = "excluded"

#: Routes whose figures were NOT tied to Financials in the latest period. A
#: deck can carry them, but a reviewer is told.
REVIEW_ROUTES = frozenset({ROUTE_PRIOR_PERIOD, ROUTE_UNCHECKED_HISTORY})

REASON_BS_DIFF = "bs_diff"
REASON_NO_ROW = "no_reconciliation_row"
REASON_NO_ACTIVITY = "no_activity"
REASON_NO_DATA = "no_data"

_REASON_TEXT = {
    REASON_BS_DIFF: "balance-sheet line reconciled to a different total than Financials; needs human reconciliation before commentary can be trusted",
    REASON_NO_ROW: "no Financials line reconciles to this tab (mapping unresolved, or the tab is not a statement account); not assumed to be an unchecked statement line",
    REASON_NO_ACTIVITY: "every period of the tab is zero",
    REASON_NO_DATA: "the tab produced no frame to read",
}


@dataclass
class Decision:
    key: str
    admitted: bool
    route: str
    reason: str = ""
    detail: str = ""
    statement: str = ""          # "BS" | "IS" | "" when no row named the account
    statuses: List[str] = field(default_factory=list)

    @property
    def needs_review(self) -> bool:
        return self.admitted and self.route in REVIEW_ROUTES

    def to_dict(self) -> Dict[str, Any]:
        return {
            "key": self.key, "admitted": self.admitted, "route": self.route,
            "reason": self.reason, "detail": self.detail, "statement": self.statement,
            "statuses": list(self.statuses),
        }


@dataclass
class SelectionResult:
    decisions: List[Decision] = field(default_factory=list)
    #: Financials lines that no tab was found for ("Not Found" with no tab).
    unmatched_financials: List[Dict[str, str]] = field(default_factory=list)
    reconciliation_available: bool = True

    @property
    def admitted_keys(self) -> List[str]:
        return [d.key for d in self.decisions if d.admitted]

    @property
    def excluded(self) -> List[Decision]:
        return [d for d in self.decisions if not d.admitted]

    @property
    def review(self) -> List[Decision]:
        return [d for d in self.decisions if d.needs_review]

    def by_key(self) -> Dict[str, Decision]:
        return {d.key: d for d in self.decisions}

    def to_dict(self) -> Dict[str, Any]:
        return {
            "reconciliation_available": self.reconciliation_available,
            "admitted": self.admitted_keys,
            "decisions": [d.to_dict() for d in self.decisions],
            "unmatched_financials": list(self.unmatched_financials),
        }


def _clean(value: Any) -> str:
    return str(value if value is not None else "").strip()


def _has_history_activity(frame: Any) -> bool:
    attrs = getattr(frame, "attrs", {}) or {}
    per_row = attrs.get("any_period_nonzero_by_description") or {}
    return isinstance(per_row, dict) and any(bool(v) for v in per_row.values())


def _zero_source_tabs(recon_frames: Iterable[Optional[pd.DataFrame]], dfs: Dict[str, Any],
                      mappings: Optional[Dict[str, Any]]) -> Dict[str, str]:
    """{tab key: statement} for the Financials lines reconciliation skipped as nil.

    Those rows carry no tab (``Tab_Account`` is "-"), so the tab is found the way
    reconciliation finds one for a compared line: through the mappings.
    """
    if not dfs or mappings is None:
        return {}
    from .reconcile import find_account_in_dfs
    found: Dict[str, str] = {}
    for position, frame in enumerate(recon_frames):
        if frame is None or frame.empty:
            continue
        if not {"Financials_Account", "Mapping_Status"} <= set(frame.columns):
            continue
        statement = "IS" if position == 1 else "BS"
        for _, row in frame.iterrows():
            if _clean(row.get("Mapping_Status")) != "Zero source":
                continue
            try:
                tab_key, _df, *_rest = find_account_in_dfs(_clean(row.get("Financials_Account")), dfs, mappings)
            except Exception:
                continue
            if tab_key and tab_key not in ("SKIP",) and tab_key in dfs:
                found.setdefault(str(tab_key), statement)
    return found


def select_reconciliation_keys(
    reconciliation: Optional[Any],
    available_keys: Iterable[str],
    resolution: Optional[Dict[str, Any]] = None,
    dfs: Optional[Dict[str, Any]] = None,
    mappings: Optional[Dict[str, Any]] = None,
) -> SelectionResult:
    """Decide, account by account, who goes to commentary. See the module docstring."""
    keys: List[str] = []
    for key in available_keys or []:
        text = _clean(key)
        if text and text not in keys:
            keys.append(text)
    key_set = set(keys)
    frames = list(reconciliation or [])[:2]
    frames += [None] * (2 - len(frames))
    result = SelectionResult()
    has_rows = any(f is not None and not f.empty for f in frames)
    if not has_rows or not keys:
        result.reconciliation_available = False
        return result

    # -- rows by account ------------------------------------------------------
    statuses: Dict[str, List[str]] = {}
    statements: Dict[str, str] = {}
    admitted_first: List[Decision] = []
    seen: set = set()
    for position, frame in enumerate(frames):
        if frame is None or frame.empty:
            continue
        if "Match" not in frame.columns or "Tab_Account" not in frame.columns:
            continue
        statement = "IS" if position == 1 else "BS"
        included = IS_INCLUDED_STATUSES if position == 1 else BS_INCLUDED_STATUSES
        for _, row in frame.iterrows():
            key, status = _clean(row.get("Tab_Account")), _clean(row.get("Match"))
            if not key or key == "-":
                continue
            if key == "Not Found" and status.startswith("⚠️ Not Found"):
                result.unmatched_financials.append(
                    {"statement": statement, "financials_account": _clean(row.get("Financials_Account"))})
                continue
            statuses.setdefault(key, []).append(status)
            statements.setdefault(key, statement)
            if status in included and key in key_set and key not in seen:
                seen.add(key)
                policy = position == 1 and status == "❌ Diff"
                admitted_first.append(Decision(
                    key, True, ROUTE_POLICY if policy else ROUTE_RECONCILED,
                    "is_diff_by_policy" if policy else "tied_out",
                    f"{statement} reconciliation: {status}", statement, list(statuses[key])))
    # statuses for first-step keys keep growing as later rows are read
    for decision in admitted_first:
        decision.statuses = list(statuses.get(decision.key, []))
    result.decisions.extend(admitted_first)

    # -- the accounts the first step did not admit ------------------------------
    zero_tabs: Optional[Dict[str, str]] = None
    for key in keys:
        if key in seen:
            continue
        rows = statuses.get(key, [])
        statement = statements.get(key, "")
        frame = (dfs or {}).get(key) if dfs else None
        if any(s.startswith("❌") for s in rows) and statement == "BS":
            result.decisions.append(Decision(
                key, False, ROUTE_EXCLUDED, REASON_BS_DIFF,
                _REASON_TEXT[REASON_BS_DIFF], statement, rows))
            continue
        if dfs is None or frame is None:
            result.decisions.append(Decision(
                key, False, ROUTE_EXCLUDED, REASON_NO_DATA, _REASON_TEXT[REASON_NO_DATA], statement, rows))
            continue
        if not _has_history_activity(frame):
            result.decisions.append(Decision(
                key, False, ROUTE_EXCLUDED, REASON_NO_ACTIVITY, _REASON_TEXT[REASON_NO_ACTIVITY], statement, rows))
            continue
        if any(s.startswith(INTERIM_PREFIX) for s in rows):
            result.decisions.append(Decision(
                key, True, ROUTE_PRIOR_PERIOD, "interim_latest_period",
                "latest column is a partial period that differs from the tab; the previous full period ties out",
                statement, rows))
            continue
        if rows:
            # A row names this tab but could not compare it (⚠️ Not Found): the
            # tab was looked at and no latest-period total could be read.
            result.decisions.append(Decision(
                key, True, ROUTE_UNCHECKED_HISTORY, "not_compared",
                "reconciliation named this tab but read no total for the latest period; tab has earlier activity",
                statement, rows))
            continue
        if zero_tabs is None:
            if mappings is None:
                try:
                    from .mapping import get_effective_mappings, load_mappings
                    mappings = get_effective_mappings(load_mappings(), resolution)
                except Exception:
                    mappings = None
            zero_tabs = _zero_source_tabs(frames, dfs, mappings)
        if key in zero_tabs:
            result.decisions.append(Decision(
                key, True, ROUTE_UNCHECKED_HISTORY, "financials_line_nil_recently",
                "Financials line is nil in the two latest periods, so reconciliation skipped it; tab has earlier activity",
                zero_tabs[key], rows))
        else:
            result.decisions.append(Decision(
                key, False, ROUTE_EXCLUDED, REASON_NO_ROW, _REASON_TEXT[REASON_NO_ROW], statement, rows))
    return result


SELECTION_FILE = "selection.json"


def save_selection(run_folder: str, result: SelectionResult) -> str:
    """Beside the run's results, so an offline re-export of the run can tell which
    accounts were left out of it and why, without redoing the reconciliation."""
    import json
    import os
    path = os.path.join(str(run_folder), SELECTION_FILE)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(result.to_dict(), fh, ensure_ascii=False, indent=1)
    return path


def load_selection(run_folder: str) -> Optional[SelectionResult]:
    import json
    import os
    path = os.path.join(str(run_folder), SELECTION_FILE)
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as fh:
        data = json.load(fh)
    result = SelectionResult(
        unmatched_financials=list(data.get("unmatched_financials") or []),
        reconciliation_available=bool(data.get("reconciliation_available", True)),
    )
    for item in data.get("decisions") or []:
        result.decisions.append(Decision(
            key=str(item.get("key")), admitted=bool(item.get("admitted")), route=str(item.get("route") or ""),
            reason=str(item.get("reason") or ""), detail=str(item.get("detail") or ""),
            statement=str(item.get("statement") or ""), statuses=list(item.get("statuses") or [])))
    return result


def selection_without_financials(available_keys: Iterable[str]) -> SelectionResult:
    """The record for a workbook with no Financials sheet: every tab goes to
    commentary (the caller's fallback), and none of it was tied out."""
    result = SelectionResult(reconciliation_available=False)
    for key in available_keys or []:
        text = _clean(key)
        if text and text not in [d.key for d in result.decisions]:
            result.decisions.append(Decision(
                text, True, ROUTE_UNCHECKED_HISTORY, "no_independent_financials",
                "no Financials sheet to reconcile against; the statement tables are built from these same tabs"))
    return result

"""Where a tab's account total and the Financials sheet disagree, Financials wins.

The deck draws its statement table from Financials and writes each account's
commentary from that account's tab. Reconciliation compares the two on the
latest period only, so a tab that was never brought up to date for an EARLIER
period reconciled clean and then contradicted the table printed beside it. One
real deck, four accounts:

  * 投资收益 -- an audit adjustment reclassified related-party loan interest
    into it for 2024 and 2025; the tab was left empty. The summary said the
    account was nil in every period, beside a table showing 1,332 and 602, and
    the account was dropped from the deck.
  * 长期应付款 -- deposits reclassified in by an audit adjustment; the tab
    carries 2025 onward only. The bullet said the balance "first arose in 2025"
    against 6,145 and 8,269 in the table.
  * 应收账款 -- "up 463%" from the tab's 2024 figure; the table showed thirty-fold.
  * 未分配利润 -- the tab holds the audited figure for 2023, the statement the
    indicative-adjusted one.

In each the tab's own Check row already recorded the gap. The owner's ruling
(2026-10-12) is that the deck follows Financials, so the account TOTAL in the
frame the model reads is set to the Financials figure for every period where
they differ, and the frame records what was changed. Component rows are left
as the tab has them: they are the tab's detail and nothing in Financials
replaces them.
"""
from __future__ import annotations

import logging
from typing import Any, Dict, List, Optional

import pandas as pd

logger = logging.getLogger(__name__)

AUTHORITY_KEY = "financials_authority"

#: A difference smaller than this cannot be seen at the precision a slide
#: prints (千元 in the table, one decimal of 万元 in the text).
_VISIBLE_DIFFERENCE = 500.0
#: Below this share of accounts agreeing with Financials on their latest
#: period, the pairing itself is in doubt -- the wrong entity's sheet, a
#: mapping gone astray -- and overriding would write another company's figures
#: into every account. One local multi-entity workbook disagrees on 11 of 11.
_MIN_WORKBOOK_AGREEMENT = 0.7


def _sign_convention(tab: Dict[str, float], fin: Dict[str, float]) -> float:
    """+1 when the tab and Financials carry this account with the same sign.

    An expense is positive on its tab and negative on the Financials sheet, so
    the Financials figure cannot be copied in as it stands. Read off the
    periods where both sides hold a figure; with none to read (a tab that is
    empty throughout), the Financials sign is kept.
    """
    votes = 0
    for period, fin_value in fin.items():
        tab_value = tab.get(period)
        if tab_value and fin_value:
            votes += 1 if (tab_value > 0) == (fin_value > 0) else -1
    return -1.0 if votes < 0 else 1.0


def apply_financials_authority(
    normalized_results: Dict[str, Dict[str, Any]],
    dfs: Dict[str, pd.DataFrame],
    bs_is_results: Optional[Dict[str, Any]],
    mappings: Optional[dict] = None,
) -> Dict[str, Dict[str, Dict[str, float]]]:
    """Set each differing account total to the Financials figure, in place.

    Mutates ``normalized_results[sheet]["prompt_analysis_df"]`` -- the frame
    every analysis variant is built from -- and returns
    ``{account: {period: {"tab": x, "financials": y}}}`` for what it changed.
    The caller rebuilds the variants afterwards; patching the built ones would
    leave their display columns and derived attrs describing the old figures.
    """
    if not bs_is_results or not dfs or not normalized_results:
        return {}
    # Lazily: fdd_utils.ai imports this package at module level.
    from ..ai.facts import _total_row_index, _value_columns, build_cross_account_facts, build_financials_by_key
    from .analysis import build_significant_movements, build_trend_summary

    series_by_account = build_cross_account_facts(dfs).get("series") or {}
    plans: List[Dict[str, Any]] = []
    compared = agreeing_latest = 0
    for statement, kind in (("balance_sheet", "BS"), ("income_statement", "IS")):
        frame = bs_is_results.get(statement)
        if not isinstance(frame, pd.DataFrame) or frame.empty:
            continue
        for account, fin in build_financials_by_key({statement: frame}, dfs, mappings).items():
            tab = {str(p): float(v) for p, v in (series_by_account.get(account) or {}).items()
                   if isinstance(v, (int, float))}
            shared = [p for p in fin if p in tab]
            if not shared:
                continue
            differing = {
                p: {"tab": tab[p], "financials": float(fin[p])} for p in shared
                if abs(abs(tab[p]) - abs(float(fin[p]))) >= _VISIBLE_DIFFERENCE
            }
            compared += 1
            latest = max(shared)
            agreeing_latest += latest not in differing
            if differing:
                plans.append({"account": account, "kind": kind, "tab": tab, "fin": fin,
                              "differing": differing, "shared": shared})
    if not plans:
        return {}
    if compared and agreeing_latest / compared < _MIN_WORKBOOK_AGREEMENT:
        logger.warning(
            "Financials authority NOT applied: only %s of %s accounts agree with the Financials sheet on "
            "their latest period, so the sheet may not belong to these tabs", agreeing_latest, compared)
        return {}

    applied: Dict[str, Dict[str, Dict[str, float]]] = {}
    for plan in plans:
        account, differing, tab = plan["account"], plan["differing"], plan["tab"]
        # A tab that holds figures and matches Financials in no period is not a
        # stale tab, it is the wrong pairing. An EMPTY tab is the other case
        # entirely: nothing was ever entered, and Financials is all there is.
        if any(tab.values()) and len(differing) == len(plan["shared"]):
            logger.warning("Financials authority skipped for %s: its tab agrees with Financials in no period", account)
            continue
        df = dfs.get(account)
        sheet = (getattr(df, "attrs", None) or {}).get("source_sheet_name")
        normalized = normalized_results.get(sheet) or {}
        analysis = normalized.get("prompt_analysis_df")
        if not isinstance(analysis, pd.DataFrame) or analysis.empty:
            continue
        row_types = analysis.attrs.get("row_types_by_description") or {}
        total_idx = _total_row_index(analysis, row_types, _value_columns(analysis))
        if total_idx is None:
            logger.warning("Financials authority skipped for %s: its frame has no total row to set", account)
            continue
        sign = _sign_convention(tab, plan["fin"])
        record: Dict[str, Dict[str, float]] = {}
        for period, pair in differing.items():
            if period not in analysis.columns:
                continue
            analysis.at[total_idx, period] = sign * pair["financials"]
            record[period] = dict(pair)
        if not record:
            continue
        statement_type = str((analysis.attrs.get("integrity") or {}).get("statement_type") or plan["kind"])
        derived = {
            AUTHORITY_KEY: record,
            "trend_summary": build_trend_summary(analysis, statement_type=statement_type),
            "significant_movements": build_significant_movements(analysis, statement_type=statement_type),
        }
        total_label = str(analysis.at[total_idx, analysis.columns[0]])
        for variant in ("prompt_analysis_df", "projection_df", "projection_df_original", "projection_df_annualized"):
            frame = normalized.get(variant)
            if not isinstance(frame, pd.DataFrame):
                continue
            frame.attrs.update(derived)
            nonzero = frame.attrs.get("any_period_nonzero_by_description")
            if isinstance(nonzero, dict):
                # The account is no longer "nil in every period", which is what
                # had it dropped from the deck.
                nonzero[total_label] = True
        applied[account] = record
    return applied


def authority_note(df: Optional[pd.DataFrame], language: str) -> str:
    """The instruction that goes with an overridden total: which periods were
    set from Financials, and what the commentary may not conclude from the
    tab's own detail for them."""
    record = (getattr(df, "attrs", None) or {}).get(AUTHORITY_KEY) or {}
    if not record:
        analysis = (getattr(df, "attrs", None) or {}).get("prompt_analysis_df")
        record = (getattr(analysis, "attrs", None) or {}).get(AUTHORITY_KEY) or {}
    if not record:
        return ""
    periods = sorted(record)
    empty = [p for p in periods if not record[p].get("tab")]
    if language == "Chi":
        note = (
            "【报表口径】本科目在" + "、".join(periods) + "的合计已按财务报表（Financials）金额列示，"
            "与报表一致；本表在这些期间的明细未随审计/示意性调整更新，加总不等于该合计。"
            "涉及这些期间的余额、发生额及跨期变动，一律以表中合计为准；"
            "不得引用这些期间的明细加总，不得据明细写出'无余额''未发生''首次形成'一类结论。"
        )
        if empty:
            note += ("其中" + "、".join(empty) + "本表未列明细，只能陈述合计金额，"
                     "不得臆造构成；如需说明性质，仅可依据备注。")
        return note
    note = (
        "[STATEMENT BASIS] This account's total for " + ", ".join(periods) + " is shown at the Financials "
        "sheet's figure, as in the statement table; the tab's own detail for those periods was not updated "
        "for the audit / indicative adjustments and does not add up to it. For those periods use the total "
        "in the table for any balance, amount or movement; do not quote the detail's sum, and do not "
        "conclude 'nil', 'not incurred' or 'first arose' from the detail."
    )
    if empty:
        note += (" For " + ", ".join(empty) + " the tab carries no detail at all: state the total only "
                 "and do not invent a composition; describe its nature only from the remarks.")
    return note

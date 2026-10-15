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
import re
from typing import Any, Dict, List, Optional, Tuple

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


_ADJUSTMENT_COLUMN = re.compile(r"ADJ|调整|調整|adjust", re.IGNORECASE)


def adjustment_entries(digest: Optional[Dict[str, Any]], gap: float, period: str,
                       limit: int = 4) -> List[str]:
    """The adjustment lines in the workbook that carry this gap, as text.

    The gap between a tab and Financials is almost always an adjustment the
    working papers record: on one real workbook every overridden figure sat in
    a year sheet's "Audit ADJ#n" column, with its counter-entry two rows away
    (投资收益 +1,332.12千元 against 财务费用 -1,332.12千元 under Audit ADJ#4 --
    loan interest reclassified out of finance cost). The model was told the
    total had been reset and nothing about why, so it wrote "成因尚待与管理层
    确认" about a reclassification the workbook spells out.

    Every cell already sits in the digest. This finds cells equal to the gap
    (in 元 or 千元) in a column labelled as an adjustment, on any sheet but the
    Financials one, and returns each such column's entries of that size -- the
    debit and the credit together -- with the row labels and, for a sheet with
    no column labels (an adjustment register), the row's own text. It states
    positions and labels; it concludes nothing.
    """
    sheets = (digest or {}).get("sheets") or {}
    if abs(gap) < 1000 or not sheets:
        return []
    year = str(period)[:4]
    lines: List[str] = []
    for name, entry in sheets.items():
        if entry.get("status") == "financials":
            continue
        cells = entry.get("cells") or {}
        col_labels = entry.get("col_labels") or {}
        row_labels = entry.get("row_labels") or {}
        by_column: Dict[int, List[Tuple[int, float, str]]] = {}
        for r, c, v in zip(cells.get("r", []), cells.get("c", []), cells.get("v", [])):
            if not v:
                continue
            if abs(abs(v) - gap) <= max(0.5, gap * 1e-6):
                unit = "元"
            elif abs(abs(v) * 1000 - gap) <= max(5.0, gap * 1e-6):  # a 千元 cell kept to 2 decimals
                unit = "千元"
            else:
                continue
            label = str(col_labels.get(str(c)) or "")
            on_register = not label and _ADJUSTMENT_COLUMN.search(str(name))
            if not (_ADJUSTMENT_COLUMN.search(label) or on_register):
                continue
            # A year sheet's adjustment belongs to the period it is named for.
            if re.fullmatch(r"\d{4}", str(name)) and str(name) != year:
                continue
            by_column.setdefault(c, []).append((r, v, unit))
        for c, hits in by_column.items():
            label = str(col_labels.get(str(c)) or "")
            parts = []
            for r, v, unit in hits:
                row_label = str(row_labels.get(str(r)) or "")
                if row_label and any(row_label in done for done in parts):
                    continue
                text = ""
                if not label:
                    words = [str(t[2]) for t in (entry.get("texts") or []) if t[0] == r and str(t[2]) != row_label]
                    text = ("（" + "；".join(words)[:60] + "）") if words else ""
                parts.append(f"{row_label or '第%d行' % (r + 1)} {v:+,.2f}{unit}{text}")
            if parts:
                lines.append(f"「{name}」{('「' + label + '」') if label else ''}：" + "；".join(parts))
            if len(lines) >= limit:
                return lines
    return lines


def apply_financials_authority(
    normalized_results: Dict[str, Dict[str, Any]],
    dfs: Dict[str, pd.DataFrame],
    bs_is_results: Optional[Dict[str, Any]],
    mappings: Optional[dict] = None,
    digest: Optional[Dict[str, Any]] = None,
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
        # "Agrees" here is reconciliation's own 0.5%, not the display threshold
        # above: a real 未分配利润 tab sits 0.1-0.2% off Financials in three
        # periods and 22% off in the fourth -- every period visibly different,
        # none of them a wrong pairing -- and was skipped by the stricter test.
        broken = [p for p, pair in differing.items()
                  if abs(abs(pair["tab"]) - abs(pair["financials"]))
                  > max(1.0, 0.005 * max(abs(pair["tab"]), abs(pair["financials"])))]
        if any(tab.values()) and len(broken) == len(plan["shared"]):
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
            try:
                found = adjustment_entries(digest, abs(abs(pair["tab"]) - abs(pair["financials"])), period)
            except Exception:
                found = []
            if found:
                record[period]["adjustments"] = found
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
    adjusted = [(p, record[p].get("adjustments")) for p in periods if record[p].get("adjustments")]
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
        if adjusted:
            note += ("【调整依据】工作底稿中金额与上述差额相符的调整分录如下（同一调整列中的借贷两方）："
                     + " ".join(f"{p}：" + " | ".join(lines) for p, lines in adjusted)
                     + "。可据此说明该合计与本表明细不同的原因（例如由哪个科目重分类转入），"
                       "但只陈述分录本身显示的内容。")
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
    if adjusted:
        note += (" [ADJUSTMENT EVIDENCE] Entries in the working papers equal to that difference (both sides "
                 "of the same adjustment column): " + " ".join(f"{p}: " + " | ".join(lines) for p, lines in adjusted)
                 + ". You may use them to explain why the total differs from the tab (e.g. which account it "
                   "was reclassified from), stating only what the entries show.")
    return note

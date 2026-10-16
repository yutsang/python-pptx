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


_BRIDGE_TOTAL_COLUMN = re.compile(r"后|数$|合计|总计|TB|报告|余额")
_BRIDGE_TOTAL_ROW = re.compile(r"合计|总计|小计")
BRIDGE_KEY = "adjustment_bridge"


def adjustment_bridge(digest: Optional[Dict[str, Any]], account: str, limit: int = 8) -> Dict[str, Any]:
    """Every audit and indicative adjustment the working papers post to one account.

    Measured against the team's own final report for one real engagement, the
    largest single thing the generated commentary left out was the
    adjustments themselves -- 「我们提出了示意性调整，补计提了某项进项税额X万元」,
    「补计提了截至期末的某项服务费Y万元」, the
    legal-fee cut-off. All of it sits in the period bridge sheets (TB -> Audit
    ADJ#n -> 审定数 -> 示意性调整 -> 示意性调整后), one row per statement line,
    and none of it reached a prompt.

    For the row labelled as this account on every sheet that has such a bridge,
    this returns each non-zero adjustment column with its amount, the other
    rows of that column carrying the same amount the other way (the
    counter-entry), and -- from an adjustment register sheet, when one holds
    the same amount against the same account -- the words written beside it.
    Returns {"lines": [...], "values": [元, ...]}; values feed the grounding
    pool so a figure the model quotes from here is not called invented.
    """
    sheets = (digest or {}).get("sheets") or {}
    lines: List[str] = []
    values: List[float] = []
    if not sheets or not account:
        return {"lines": lines, "values": values}
    registers = {n: e for n, e in sheets.items() if re.search(r"ADJ|调整", str(n), re.IGNORECASE)}
    # (sheet position, audit-not-indicative, line). Kept whole and cut to the
    # limit afterwards, latest sheet and indicative columns first: walked in
    # workbook order, the 2023 and 2024 audit adjustments filled the limit
    # before the 3M26 indicative ones -- the ones the final report writes
    # about -- were reached.
    ranked: List[Tuple[int, int, str]] = []
    related: List[str] = []
    for position, (name, entry) in enumerate(sheets.items()):
        col_labels = {int(k): str(v) for k, v in (entry.get("col_labels") or {}).items()}
        if not any(re.search(r"审定数|示意性调整后", v) for v in col_labels.values()):
            continue
        row_labels = {int(k): str(v).strip() for k, v in (entry.get("row_labels") or {}).items()}
        own_rows = [r for r, label in row_labels.items() if label == account]
        if not own_rows:
            continue
        cells = entry.get("cells") or {}
        grid: Dict[Tuple[int, int], float] = {
            (r, c): v for r, c, v in zip(cells.get("r", []), cells.get("c", []), cells.get("v", [])) if v}
        side = _bridge_sides(row_labels)
        sheet_sums: Dict[str, float] = {}
        for row in own_rows[:1]:
            for (r, c), v in sorted(grid.items(), key=lambda kv: kv[0][1]):
                label = col_labels.get(c, "")
                if r != row or not re.search(r"ADJ|调整|adj", label) or _BRIDGE_TOTAL_COLUMN.search(label):
                    continue
                yuan = v * 1000.0  # bridge sheets are kept in 千元
                if abs(yuan) < 500:
                    continue  # prints as 0.0万元: a rounding residue, not an entry
                values.append(abs(yuan))
                kind = "示意" if "示意" in label else ("管理层" if "管理层" in label else "审计")
                sheet_sums[kind] = sheet_sums.get(kind, 0.0) + yuan
                column = {r2: v2 for (r2, c2), v2 in grid.items()
                          if c2 == c and r2 != r and _bridge_line(row_labels.get(r2, ""))}
                others = sorted({row_labels[r2] for r2, v2 in column.items()
                                 if abs(abs(v2) - abs(v)) <= 0.006 and _counter_sign(side, r, v, r2, v2)})
                if not others:
                    # One entry against two lines: 应交税费 +640.5 is 其他流动资产
                    # 529.2 plus 其他非流动资产 111.3 -- the very tie the final
                    # report writes, invisible to a one-for-one match.
                    pairs = [(r2, r3) for r2 in column for r3 in column if r2 < r3
                             and abs(abs(column[r2]) + abs(column[r3]) - abs(v)) <= 0.006
                             and _counter_sign(side, r, v, r2, column[r2])
                             and _counter_sign(side, r, v, r3, column[r3])]
                    if len(pairs) == 1:
                        others = sorted(row_labels[x] for x in pairs[0])
                linked = sorted({
                    row_labels.get(r2, "") for (r2, c2), v2 in grid.items()
                    if c2 != c and r2 != r and abs(abs(v2) - abs(v)) <= 0.006
                    and re.search(r"ADJ|调整|adj", col_labels.get(c2, ""))
                    and not _BRIDGE_TOTAL_COLUMN.search(col_labels.get(c2, ""))
                    and _bridge_line(row_labels.get(r2, "")) and row_labels.get(r2) not in others
                })
                related.extend(x for x in list(others) + list(linked) if x not in related)
                words = _register_words(registers, account, v, row_labels, label)
                line = f"「{name}」{label}：{'+' if v > 0 else '-'}{abs(yuan) / 1e4:,.1f}万元"
                if others:
                    line += "（对方科目：" + "、".join(others[:3]) + "）"
                if linked:
                    line += "（同额另见于其他调整：" + "、".join(linked[:3]) + "）"
                if words:
                    line += "；调整说明：" + "；".join(words)
                ranked.append((-position, 0 if kind == "示意" else 1, line))
        # The model adds a sheet's entries up and quotes the total: 「我方对
        # 2026年03月31日应付账款作示意性调增89.7万元，包括…60.2万元、9.9万元及
        # 其他19.6万元」 was flagged as an invented figure and sent back twice.
        for total in sheet_sums.values():
            if abs(total) >= 500:
                values.append(abs(total))
    lines = [line for _, _, line in sorted(ranked, key=lambda item: (item[1], item[0]))[:limit]]
    return {"lines": lines, "values": values, "related": related}


def _bridge_line(label: str) -> bool:
    """A statement line a counter-entry can be posted to: not a total, not a check row."""
    # 净利润 / 利润总额 / 年末未分配利润 are the profit roll-forward, not lines an
    # entry is posted to: every income-statement adjustment flows through them,
    # and on the real book they were named as the "counter-account" of 营业收入,
    # 营业成本 and 管理费用 alike. 未分配利润 itself is a balance and stays.
    return (bool(label) and not _BRIDGE_TOTAL_ROW.search(label)
            and not re.search(r"check|差异|净利润|利润总额|营业利润|本期利润|年[初末]未分配利润", label, re.I))


def _bridge_sides(row_labels: Dict[int, str]) -> Dict[str, Optional[int]]:
    """Where the asset block and the liabilities-and-equity block end on a bridge sheet."""
    asset_end = next((r for r, l in sorted(row_labels.items()) if re.search(r"资产(?:总计|合计)$", l)), None)
    bs_end = next((r for r, l in sorted(row_labels.items())
                   if re.search(r"负债(?:和|及)(?:所有者|股东)权益(?:总计|合计)", l)), None)
    return {"asset_end": asset_end, "bs_end": bs_end}


def _counter_sign(side: Dict[str, Optional[int]], r: int, v: float, r2: int, v2: float) -> bool:
    """Is (r2, v2) signed the way the other half of an entry on (r, v) would be?

    The bridge shows balances, not debits and credits. Between two assets, or two
    liabilities, an entry shows as opposite signs (其他应付款 -826.9 against 长期
    应付款 +826.9); between an asset and a liability it shows as the SAME sign
    (应收账款 +19.0 against 预收款项 +19.0), and the opposite-sign rule alone
    missed every such pair -- the 应收账款 bullet then took the register words
    of an unrelated entry and called the 19.0 an income-tax adjustment. Rows
    on the income statement, or a sheet whose blocks cannot be found, accept
    either sign.
    """
    asset_end, bs_end = side.get("asset_end"), side.get("bs_end")
    if asset_end is None or bs_end is None:
        return True

    def block(row: int) -> str:
        return "A" if row <= asset_end else ("L" if row <= bs_end else "P")

    a, b = block(r), block(r2)
    if "P" in (a, b):
        return True
    return ((v > 0) != (v2 > 0)) if a == b else ((v > 0) == (v2 > 0))


# 「Y」: a one-letter flag column in the real register rode along as the first
# "word" of every 3M26 description.
_REGISTER_NOISE = re.compile(r"^(?:借|贷|Dr|Cr)\s*[：:]|^\d{4}[-/年]|^(?:Adj\.?#|类型|明细|科目|金额)$|^\w{0,2}$", re.I)


def _register_words(registers: Dict[str, Any], account: str, v: float, bridge_rows: Dict[int, str],
                    entry_label: str = "") -> List[str]:
    """The description an adjustment register writes for this entry, if it can be found.

    The real register (read off one engagement's ADJ sheet, 2026-10-16) carries
    the entry id in its first column on EVERY row -- 「Audit ADJ#4」, 「ADJ#3」 --
    with the description either on a row of its own above the Dr/Cr lines
    (audit entries: 「应收某关联方借款利息重分类至投资收益」, no figures) or
    repeated on every line (indicative entries: 「补提工程款尾款的税额」, plus a
    note column such as 「电费成本，管理层入账在收入」). The account sits in a
    later column and is the row's label.

    Two earlier readings were wrong. Three rows above the amount crossed into
    the previous entry: the 19.6 lease-commission accrual came back as ADJ#3's
    「补提工程款尾款的税额」. "Nearest header above" then found the row ITSELF,
    since every row carries its id, and returned nothing for every audit entry.
    Now: the row is found by amount against this account, or failing that by
    the bridge column's own id (the 3M26 ADJ#3 of 60.2 is two register lines,
    599.6 and 2.7 千元, so no single cell matches), and the words are the
    entry's other texts -- every row with the same id, contiguous with it.
    """
    wanted = _entry_id(entry_label)
    found: List[str] = []
    for reg in registers.values():
        reg_rows = {int(k): str(t).strip() for k, t in (reg.get("row_labels") or {}).items()}
        reg_cells = reg.get("cells") or {}
        by_row: Dict[int, List[Tuple[int, str]]] = {}
        for t in reg.get("texts") or []:
            by_row.setdefault(int(t[0]), []).append((int(t[1]), str(t[2]).strip()))
        ids = {row: next((_entry_id(text) for _, text in texts if _entry_id(text)), "")
               for row, texts in by_row.items()}
        hits = [r3 for r3, v3 in zip(reg_cells.get("r", []), reg_cells.get("v", []))
                if reg_rows.get(int(r3)) == account and abs(abs(v3) - abs(v)) <= 0.006]
        if not hits and wanted:
            hits = [row for row, entry in ids.items() if entry == wanted and reg_rows.get(row) == account]
        for r3 in hits[:1]:
            entry = ids.get(int(r3), "")
            if not entry:
                continue
            top = int(r3)
            while ids.get(top - 1) == entry:
                top -= 1
            bottom = int(r3)
            while ids.get(bottom + 1) == entry:
                bottom += 1
            names = set(reg_rows.values()) | set(bridge_rows.values())
            for row in range(top, bottom + 1):
                for _, text in sorted(by_row.get(row, [])):
                    if (text and text not in names and not _entry_id(text)
                            and not _REGISTER_NOISE.search(text)):
                        found.append(text[:40])
    return list(dict.fromkeys(found))[:2]


def _entry_id(text: str) -> str:
    """「Audit ADJ#4」 -> "AUDITADJ#4", 「ADJ#3」 -> "ADJ#3"; "" when not an id."""
    m = re.fullmatch(r"\s*((?:Audit\s*)?(?:IA|ADJ|Adj)\.?\s*#\s*\d+)\s*", str(text or ""), re.I)
    return re.sub(r"[\s.]", "", m.group(1)).upper() if m else ""


def attach_adjustment_bridges(dfs: Dict[str, pd.DataFrame], digest: Optional[Dict[str, Any]]) -> int:
    """adjustment_bridge for every account frame, filed on its attrs. Returns
    how many accounts had at least one entry."""
    found = 0
    for account, frame in (dfs or {}).items():
        if not isinstance(frame, pd.DataFrame):
            continue
        try:
            bridge = adjustment_bridge(digest, str(account))
        except Exception as exc:  # enrichment, never a gate
            logger.warning("Adjustment bridge skipped for %s: %s", account, exc)
            continue
        if bridge["lines"]:
            frame.attrs[BRIDGE_KEY] = bridge
            found += 1
    # What the other side of an entry says about itself. 投资收益 is a nil tab
    # whose whole story -- a USD loan to a related party, one year, 6%, repaid
    # in April 2025 -- is in 其他应收款's notes; joined to it only by the
    # register, the 投资收益 bullet called the fall "成因尚待与管理层确认" on
    # three runs while the 其他应收款 bullet a page earlier explained it.
    for account, frame in (dfs or {}).items():
        bridge = frame.attrs.get(BRIDGE_KEY) if isinstance(frame, pd.DataFrame) else None
        if not bridge:
            continue
        notes: List[str] = []
        for other in bridge.get("related") or []:
            other_frame = dfs.get(other)
            if other == account or "未分配利润" in other or not isinstance(other_frame, pd.DataFrame):
                continue
            texts = [str(t).strip() for t in (other_frame.attrs.get("supporting_notes") or [])]
            texts = [t for t in texts if len(re.findall(r"[\u4e00-\u9fffA-Za-z]", t)) >= 8][:2]
            if texts:
                notes.append(f"{other}：" + "；".join(t[:120] for t in texts))
                for number, unit in re.findall(r"(\d[\d,]*(?:\.\d+)?)\s*(亿|万)?", "".join(texts)):
                    value = float(number.replace(",", "")) * {"亿": 1e8, "万": 1e4}.get(unit, 1.0)
                    if value >= 1000:  # amounts only: a bare 1, 6 or 2025 would ground anything
                        bridge["values"].append(value)
            if len(notes) >= 2:
                break
        if notes:
            bridge["related_notes"] = notes
    return found


def bridge_note(df: Optional[pd.DataFrame], language: str) -> str:
    bridge = (getattr(df, "attrs", None) or {}).get(BRIDGE_KEY) or {}
    lines = bridge.get("lines") or []
    if not lines:
        return ""
    notes = bridge.get("related_notes") or []
    if language == "Chi":
        text = ("【本科目的审计调整及示意性调整（工作底稿）】" + " | ".join(lines)
                + "。其中「示意性调整」为我方提出的调整：如与本科目有关，应说明调整内容、金额及原因"
                  "（原因仅依据上述说明文字或备注）；审计调整可用于解释与管理层报表的差异。")
        if notes:
            text += ("【分录对方科目的资料】" + " | ".join(notes)
                     + "。本科目的变动若可由上述资料解释（例如收益来自对方科目所述的借款），直接据此说明，"
                       "不要写成原因待确认；外文公司名照原文书写，不要自行译成中文。")
        return text
    text = ("[AUDIT AND INDICATIVE ADJUSTMENTS TO THIS ACCOUNT (working papers)] " + " | ".join(lines)
            + ". Indicative adjustments are ours: where relevant, state what was adjusted, by how much and why "
              "(the why only from the notes above or the remarks); audit adjustments explain differences from "
              "the management accounts.")
    if notes:
        text += (" [WHAT THE OTHER SIDE OF THESE ENTRIES SAYS] " + " | ".join(notes)
                 + ". Where this explains the account's movement (e.g. income arising from a loan described there), "
                   "say so rather than leaving the cause to management.")
    return text

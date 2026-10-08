"""Cross-account facts and verified relationships (milestone N2).

Every account is generated in isolation from its own DataFrame, so no
commentary can currently say anything that needs TWO tabs at once. The one
exception, ``_build_peer_context``, reads the account frames' own date columns
-- and MEASURED on all four local databooks it returns ``None`` for every one
of them, because the selected variant frame carries exactly ONE date column
(plus ``__source_row_idx`` and a ``_formatted`` twin). The multi-period series
is on ``df.attrs["prompt_analysis_df"]`` instead: 96/96 accounts across the
four files have >= 2 ISO date columns there, 0/96 have them on the frame.
So this module reads the analysis frame first and the frame second. That is a
correction to the plan's premise, not a stylistic choice -- built the way the
plan described it, the fact table would have been empty on every file we have.

THE RULE THIS MODULE EXISTS TO ENFORCE: a relationship that has not passed its
deterministic test never reaches a prompt. Numeric relationships have numeric
tests; text relationships require an exact normalised identity or an explicit
name in source text. Candidates are proposed in code (never an LLM), each
carries its own test, and every one is stored with ``passed`` either way. The
prompt builder in prompts.py reads only explicitly allow-listed ``passed=True``
edge kinds; the rejects are kept for the
internal insight summary (N3), where a wrong hypothesis costs a reviewer
thirty seconds instead of shipping in a client deck.

Nothing here calls an LLM, reads the network, or writes a file.
"""

from __future__ import annotations

import hashlib
import json
import re
from statistics import median
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

import pandas as pd

__all__ = [
    "build_cross_account_facts",
    "build_digest_graph_links",
    "cross_account_links_for",
    "digest_tieout_diagnostics",
    "financials_tieout",
    "graph_summary",
    "merge_graph_links",
    "resolve_statement_type",
]

_INTERNAL_COL = "__source_row_idx"
# The same tolerance extract_presentation_detail_table's own detected path
# uses (schedules.py:1230): max(1.0, 2%). Keep the two in step -- a tie this
# module reports as passing must be the tie that path would have reported.
_TIE_TOLERANCE = 0.02
_DAYS_PER_MONTH = 30.44
# A trailing column shorter than this fraction of the series' usual step is a
# stub (a January month tacked onto three calendar years), not a period.
_STUB_RATIO = 0.6

_ISO_RE = re.compile(r"^(\d{4})-(\d{2})-(\d{2})$")
_COMPANY_MARKERS = (
    "有限公司", "有限责任公司", "股份有限公司",
    " limited", " ltd", " llc", " incorporated", " inc.", " company", " co.",
)
_LABEL_PUNCT_RE = re.compile(r"[\s*＊()（）\[\]【】,，.。:：;；'\"“”‘’·_\-]+")
_TOTAL_LABEL_RE = re.compile(
    r"^(?:grand\s+total|sub\s*total|subtotal|total|合计|合計|总计|總計|小计|小計|共计|共計)$",
    flags=re.IGNORECASE,
)


# --------------------------------------------------------------------------
# roles -- what an account is, for the purpose of pairing it with another one
# --------------------------------------------------------------------------
# Deliberately narrow. "receivable" means TRADE receivables, because a days
# figure computed off other receivables is not a collection period and would
# be a plausible-looking wrong number in a deliverable. Short tab keys ('AR',
# 'OP') are matched as whole keys, never as substrings: 'or' as a substring
# hits 'Long-term loans', 'Corporate...', and most English words there are.
_EXACT_ROLE_KEYS: Dict[str, Tuple[str, ...]] = {
    "receivable": ("ar",),
    "payable": ("ap",),
    "revenue": ("sales",),
    "cost": ("cogs", "cost"),
}
_ROLE_NEEDLES: Dict[str, Tuple[str, ...]] = {
    "receivable": ("accounts receivable", "trade receivable", "应收账款"),
    "payable": ("accounts payable", "trade payable", "应付账款"),
    "inventory": ("inventory", "inventories", "存货"),
    "revenue": ("operating income", "operating revenue", "revenue", "营业收入",
                "主营业务收入"),
    "cost": ("operating cost", "cost of sales", "cost of revenue", "营业成本",
             "主营业务成本"),
    "advance": ("advance payments received", "advances", "advance receipts",
                "contract liabilit", "预收", "合同负债"),
    "prepayment": ("prepayment", "prepaid", "预付"),
}
# Anything carrying one of these is NOT the role, however well it otherwise
# matches: "non-operating income" is not revenue and "cost of financing" is
# not the cost of sales.
_ROLE_EXCLUSIONS: Dict[str, Tuple[str, ...]] = {
    "revenue": ("non-operating", "other income", "营业外", "其他收益", "投资收益"),
    "cost": ("financial", "财务", "费用"),
    "receivable": ("other receivable", "其他应收"),
    "payable": ("other payable", "employee", "其他应付", "职工薪酬"),
    "advance": ("prepayment", "预付"),
    "prepayment": ("advance payments received", "预收"),
}


def _iso_months(period: str) -> Optional[int]:
    """A month ordinal, so two ISO dates can be subtracted."""
    match = _ISO_RE.match(str(period).strip())
    if not match:
        return None
    return int(match.group(1)) * 12 + int(match.group(2))


def _period_lengths(periods: Sequence[str]) -> Dict[str, float]:
    """How many months each column covers, from the gaps between the columns
    themselves rather than from ``annualization_months``.

    ``annualization_months`` describes the SELECTED VARIANT, not the analysis
    frame -- measured on one databook it is 1 for five IS accounts and None
    for two others whose selected column happened to be an earlier year, even
    though all seven analysis frames carry the identical four columns. Reading
    it here would truncate some accounts' tail and not others', and a ratio
    built across two accounts truncated differently is silently wrong.
    """
    lengths: Dict[str, float] = {}
    ordinals = [(p, _iso_months(p)) for p in periods]
    gaps: List[float] = []
    for i in range(1, len(ordinals)):
        prev, curr = ordinals[i - 1][1], ordinals[i][1]
        gaps.append(float(curr - prev) if prev is not None and curr is not None else 12.0)
    for i, (period, _ord) in enumerate(ordinals):
        # The first column has no predecessor; assume it steps like the rest.
        lengths[period] = gaps[i - 1] if i >= 1 else (gaps[0] if gaps else 12.0)
    return lengths


def _full_periods(periods: Sequence[str]) -> List[str]:
    """The periods a growth ratio may be derived over -- the tail dropped only
    when it is a stub relative to the series' own step.

    The fact table keeps every period, including the stub. Truncating there
    would make a January column disappear from every cross-account comparison
    AND from the balance-sheet ones, where a month-end balance is a perfectly
    good point in time. Only a ratio needs the periods to be comparable
    lengths, so only a ratio pays for the truncation.

    With fewer than three columns nothing is dropped: one gap is not a series
    and there is no step to call the tail short against. A two-column account
    whose second column is a stub therefore keeps it -- accepted knowingly,
    because guessing which of two periods is the odd one out is worse than
    reporting both.
    """
    periods = [str(p) for p in periods]
    if len(periods) < 3:
        return list(periods)
    lengths = _period_lengths(periods)
    earlier = [lengths[p] for p in periods[:-1]]
    step = median(earlier) if earlier else 12.0
    if step > 0 and lengths[periods[-1]] < step * _STUB_RATIO:
        return periods[:-1]
    return list(periods)


def _value_columns(frame: pd.DataFrame) -> List[str]:
    return [
        str(c) for c in list(frame.columns)[1:]
        if str(c) != _INTERNAL_COL and not str(c).endswith("_formatted")
    ]


def _total_row_index(frame: pd.DataFrame, row_types: Dict[str, Any], periods: Sequence[str]):
    """The labelled total row: a 'total' ahead of a 'subtotal', and among
    equals the one that actually carries figures.

    Every existing reader of this attribute (``_build_peer_context``,
    ``_data_insight_guidance._account_total``, ``_variance_analysis_guidance``)
    takes the LAST row typed total OR subtotal. MEASURED on one local
    databook, that reads Investment properties as 0 in all three periods: the
    tab has 'Total' = 292,709,788 at source row 19 and a stray '合计' = 0 at
    source row 39, inside a working note below the schedule, and both are
    typed 'total'. The Financials sheet says 292,709,788, so the last-match
    rule is simply wrong there and the account looked like a total extraction
    failure. Preferring the total row with figures in it fixes this module;
    the three readers in prompts.py still take the last match and are left
    alone here, because changing them changes rendered prompt text and that
    is not this milestone's change to make.
    """
    desc_col = frame.columns[0]
    ranked: List[Tuple[int, int, Any]] = []
    for position, (idx, row) in enumerate(frame.iterrows()):
        kind = str(row_types.get(str(row[desc_col]), "")).strip().lower()
        if kind not in ("total", "subtotal"):
            continue
        filled = 0
        for period in periods:
            if period in frame.columns:
                value = row[period]
                if pd.notna(value) and isinstance(value, (int, float)) and abs(value) > 1e-9:
                    filled += 1
        ranked.append((0 if kind == "total" else 1, -filled, position))
    if not ranked:
        return None
    ranked.sort(key=lambda item: (item[0], item[1], -item[2]))
    return list(frame.index)[ranked[0][2]]


def _account_series(df: pd.DataFrame) -> Tuple[List[str], Dict[str, float], str]:
    """(periods, {period: total}, source) for one account.

    ``source`` names which frame the series came from, so a caller can tell a
    real multi-period series from a single-column fallback without guessing.
    """
    attrs = df.attrs or {}
    row_types = attrs.get("row_types_by_description") or {}
    candidates: List[Tuple[str, pd.DataFrame]] = []
    analysis = attrs.get("prompt_analysis_df")
    if isinstance(analysis, pd.DataFrame) and not analysis.empty:
        candidates.append(("prompt_analysis_df", analysis))
    candidates.append(("frame", df))

    best: Tuple[List[str], Dict[str, float], str] = ([], {}, "")
    for source, frame in candidates:
        cols = _value_columns(frame)
        if not cols:
            continue
        total_idx = _total_row_index(frame, row_types, cols)
        values: Dict[str, float] = {}
        for col in cols:
            try:
                if total_idx is None:
                    values[col] = float(frame[col].fillna(0).sum())
                else:
                    raw = frame.loc[total_idx, col]
                    values[col] = float(raw) if pd.notna(raw) else 0.0
            except Exception:
                continue
        if not values:
            continue
        periods = [c for c in cols if c in values]
        if len(periods) > len(best[0]):
            best = (periods, values, source)
        if len(best[0]) >= 2:
            break
    return best


def resolve_statement_type(
    key: str,
    df: Optional[pd.DataFrame],
    type_lookup: Optional[Callable[[str], Any]] = None,
) -> str:
    """BS / IS / "" for one account, with the dynamic-mapping blind spot closed.

    mappings.yml is the first source and is wrong by omission for a
    dynamically-mapped account -- ``get_mapping_component`` reads the file
    from disk and never sees ``resolution['dynamic_mappings']``. Measured, that
    is 1 account of 24 on one local databook ('S&D expenses') and 0 of 26/27/19
    on the other three: small, but the consequence is not proportional, since
    an account with no type is excluded from every grouping rather than merely
    mis-grouped.

    Then ``df.attrs['type']``, which is where databook.py is to copy
    ``normalized['type']`` -- MEASURED it is absent from all 34 attrs keys on
    every account of all four local databooks today, so this rung is dead until
    that lands and the fallback below is what actually carries the load.

    Then ``df.attrs['integrity']['statement_type']``, which IS populated
    (24/24, 26/26, 27/27, 19/19 measured). It is deliberately blanked for IS
    accounts resolved without a matched alias, so it is an interim fallback
    rather than a general substitute -- hence the order.
    """
    if type_lookup is not None:
        try:
            found = str(type_lookup(key) or "").strip().upper()
        except Exception:
            found = ""
        if found in ("BS", "IS"):
            return found
    attrs = (df.attrs or {}) if isinstance(df, pd.DataFrame) else {}
    direct = str(attrs.get("type") or "").strip().upper()
    if direct in ("BS", "IS"):
        return direct
    integrity = attrs.get("integrity") or {}
    fallback = str(integrity.get("statement_type") or "").strip().upper()
    return fallback if fallback in ("BS", "IS") else ""


def _role_text(key: str, df: pd.DataFrame) -> Tuple[str, str]:
    """(whole-key token, searchable text) for role matching."""
    attrs = df.attrs or {}
    parts = [str(key)]
    try:
        parts.append(str(df.columns[0]))
    except Exception:
        pass
    parts.append(str(attrs.get("block_title") or ""))
    parts.append(str(attrs.get("source_sheet_name") or ""))
    return str(key).strip().lower(), " ".join(parts).strip().lower()


def _has_role(role: str, key: str, df: pd.DataFrame) -> bool:
    token, text = _role_text(key, df)
    for bad in _ROLE_EXCLUSIONS.get(role, ()):
        if bad in text:
            return False
    if token in _EXACT_ROLE_KEYS.get(role, ()):
        return True
    return any(needle in text for needle in _ROLE_NEEDLES.get(role, ()))


def _normalised_company_label(value: Any) -> Optional[str]:
    """A conservative identity key for named counterparties, not categories."""
    raw = str(value or "").strip()
    lowered = raw.lower()
    if not raw or not any(marker in lowered for marker in _COMPANY_MARKERS):
        return None
    # Vendor/customer codes are presentation metadata, not part of identity.
    # The same counterparty appears with a 10-digit code in AR/AP schedules and
    # without it in revenue/contract-liability schedules.
    lowered = re.sub(r"^[+-]?\d{5,}\s*", "", lowered)
    normalised = _LABEL_PUNCT_RE.sub("", lowered)
    return normalised if len(normalised) >= 4 else None


def _table_labels(table: Any) -> List[str]:
    labels: List[str] = []

    def add_rows(rows: Any) -> None:
        for row in rows or []:
            if not isinstance(row, dict):
                continue
            label = str(row.get("label") or "").strip()
            if label:
                labels.append(label)
            add_rows(row.get("children"))

    if isinstance(table, dict):
        add_rows(table.get("rows"))
    return labels


def _text_items(value: Any, origin: str, depth: int = 0) -> List[Tuple[str, str]]:
    """Flatten only text-bearing prompt attrs, retaining where each came from."""
    if depth > 5 or value is None:
        return []
    if isinstance(value, str):
        text = value.strip()
        return [(origin, text)] if text else []
    if isinstance(value, dict):
        out: List[Tuple[str, str]] = []
        for item in value.values():
            out.extend(_text_items(item, origin, depth + 1))
        return out
    if isinstance(value, (list, tuple)):
        out = []
        for item in value:
            out.extend(_text_items(item, origin, depth + 1))
        return out
    return []


def _name_in_text(name: str, text: str) -> bool:
    name, text = str(name or "").strip(), str(text or "")
    if not name:
        return False
    if re.search(r"[\u3400-\u9fff]", name):
        return len(name) >= 2 and name in text
    if len(name) < 4:
        return False
    # The hyphen matters: without it, "Operating income" matches inside
    # "Non-operating income" and manufactures a cross-account reference.
    return re.search(r"(?<![A-Za-z0-9-])" + re.escape(name) + r"(?![A-Za-z0-9-])",
                     text, flags=re.IGNORECASE) is not None


# --------------------------------------------------------------------------
# relationship candidates -- an explicit list, each with its own numeric test
# --------------------------------------------------------------------------
# (kind, source role, target role, unit, band low, band high, label)
# The band is the test. A value outside it is not "interesting" -- it is
# evidence the pairing does not hold for this entity (a receivable that
# equates to four years of revenue is a related-party balance, not a
# collection period), and the edge is stored rejected and never quoted.
#
# TWO gates, not one, and they failed differently on real files before they
# were separated:
#   * the BAND is the relationship test. Above it the pairing does not hold
#     for this entity (one real COGS tab at 6072% of that year revenue is a start-up
#     year, not a gross margin; 应付账款 at 17,072 days of 营业成本 is a
#     construction-stage balance, not a payment period). One period outside
#     rejects the edge -- a lone outlier is far more often an artefact than a
#     finding.
#   * MATERIALITY is a separate floor on the largest period. A first cut put
#     the floor into the band's low end instead, and it rejected
#     'G&A -> Sales' at a flat 0.2/0.4/0.4% and 'AR -> Sales' at 0.5/2.2/4.7
#     days -- both stable, both quotable, both real -- while the case it was
#     written for ('Other profit' at 0.0% in every period) needed only the
#     largest value tested. Uniformly negligible is the thing worth dropping.
# (kind, source role, target role, unit, materiality floor, band ceiling)
_RATIO_CANDIDATES: Tuple[Tuple[str, str, str, str, float, float], ...] = (
    ("receivable_days", "receivable", "revenue", "days", 0.5, 365.0),
    ("payable_days", "payable", "cost", "days", 0.5, 365.0),
    ("inventory_days", "inventory", "cost", "days", 0.5, 730.0),
    ("advance_days", "advance", "revenue", "days", 0.5, 730.0),
    ("prepayment_days", "prepayment", "cost", "days", 0.5, 365.0),
    ("expense_to_revenue", "is_expense", "revenue", "pct", 0.1, 100.0),
)


def _endpoint_ref(value: Any) -> Dict[str, str]:
    """A typed endpoint while the legacy string stays available to callers."""
    raw = str(value or "")
    if raw.startswith("WorkbookSheet::"):
        location = raw.split("::", 1)[1]
        sheet = location.split("#", 1)[0]
        return {"kind": "workbook_block", "sheet": sheet, "node": location}
    if raw.startswith("Financials::"):
        account = raw.split("::", 1)[1]
        return {"kind": "financials_row", "account": account, "node": account}
    if "::" in raw:
        account, node = raw.split("::", 1)
        return {
            "kind": "account_block" if node == "breakdown" else "account_row",
            "account": account,
            "node": node,
        }
    return {"kind": "account", "account": raw, "node": raw}


def _digest_multiplier(sheet: Dict[str, Any]) -> float:
    """Explicit unit lineage for an unmapped sheet; never guess a scale."""
    markers = " ".join(str(v).lower() for v in (sheet.get("unit_markers") or []))
    if any(v in markers for v in ("million", "百万", "百萬")):
        return 1_000_000.0
    if any(v in markers for v in ("'000", "千元")):
        return 1_000.0
    return 1.0


def _digest_periods_by_column(sheet: Dict[str, Any]) -> Dict[int, str]:
    """Date labels attached to numeric columns, from the digest's own cells."""
    from ..financial_common import normalize_financial_date_label

    periods: Dict[int, str] = {}
    for col, label in (sheet.get("col_labels") or {}).items():
        normalised = normalize_financial_date_label(label)
        if _ISO_RE.match(str(normalised or "").strip()):
            periods[int(col)] = str(normalised)
    # Some supporting schedules have a title row between the date and the
    # first numeric cell, beyond the digest's nearest-header search. Retain the
    # nearest date-looking text in the same column as a conservative fallback.
    for item in sheet.get("texts") or []:
        if not isinstance(item, list) or len(item) < 3:
            continue
        normalised = normalize_financial_date_label(item[2])
        if _ISO_RE.match(str(normalised or "").strip()):
            periods.setdefault(int(item[1]), str(normalised))
    return periods


def _sheet_names_account(sheet_name: str, sheet: Dict[str, Any], account: str, label: str) -> bool:
    """Whether sheet/title explicitly names an account, including short keys."""
    text = " ".join((str(sheet_name), str(sheet.get("title") or "")))
    for name in dict.fromkeys((str(account).strip(), str(label).strip())):
        if not name:
            continue
        if _name_in_text(name, text):
            return True
        if (2 <= len(name) <= 3 and name.isascii() and name.isupper()
                and re.search(r"(?<![A-Za-z0-9])" + re.escape(name)
                              + r"(?![A-Za-z0-9])", text, flags=re.IGNORECASE)):
            return True
    return False


def _digest_total_rows(
    sheet: Dict[str, Any],
) -> Tuple[Dict[int, str], float, List[Tuple[int, int, str, Dict[str, float]]]]:
    """Usable labelled total rows with raw figures scaled to base units."""
    period_by_col = _digest_periods_by_column(sheet)
    multiplier = _digest_multiplier(sheet)
    cells = sheet.get("cells") or {}
    rows: Dict[Tuple[int, int], Dict[str, float]] = {}
    for row, col, raw, block in zip(
        cells.get("r") or [], cells.get("c") or [],
        cells.get("v") or [], cells.get("block") or [],
    ):
        period = period_by_col.get(int(col))
        if period is None:
            continue
        rows.setdefault((int(block), int(row)), {})[period] = float(raw) * multiplier
    totals = []
    for (block, row), values in sorted(rows.items()):
        label = str((sheet.get("row_labels") or {}).get(str(row)) or "").strip()
        if _TOTAL_LABEL_RE.match(label) and any(abs(v) > 1e-9 for v in values.values()):
            totals.append((block, row, label, values))
    return period_by_col, multiplier, totals


def build_digest_graph_links(
    workbook_digest: Optional[Dict[str, Any]],
    facts: Optional[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    """Verified total-row tie-outs from unmapped supporting tabs to accounts.

    This is intentionally narrower than "numbers happen to match". A source
    must be a labelled total row inside one digest block, every comparable
    period must tie within 2%, and a one-period match is admitted only when the
    sheet/title also explicitly names the account. Ambiguous matches are kept
    as failed edges and cannot be consumed later.
    """
    sheets = (workbook_digest or {}).get("sheets") or {}
    series = (facts or {}).get("series") or {}
    labels = (facts or {}).get("labels") or {}
    if not sheets or not series:
        return []

    edges: List[Dict[str, Any]] = []
    for sheet_name, sheet in sheets.items():
        if not isinstance(sheet, dict) or sheet.get("status") != "unmapped":
            continue
        period_by_col, multiplier, total_rows = _digest_total_rows(sheet)
        if not period_by_col:
            continue
        for block, row, row_label, values in total_rows:
            matched: List[Tuple[str, List[str], bool, Dict[str, float]]] = []
            for account, account_values in series.items():
                shared = sorted(
                    p for p in values
                    if isinstance(account_values.get(p), (int, float))
                    and abs(float(account_values[p])) > 1e-9
                )
                name_hint = _sheet_names_account(
                    str(sheet_name), sheet, str(account), str(labels.get(account) or ""))
                if len(shared) < 2 and not name_hint:
                    continue
                differing: Dict[str, float] = {}
                for period in shared:
                    left, right = float(values[period]), float(account_values[period])
                    if abs(left - right) > max(1.0, max(abs(left), abs(right)) * _TIE_TOLERANCE):
                        differing[period] = right
                if shared and not differing:
                    matched.append((
                        str(account), shared, name_hint,
                        {p: float(account_values[p]) for p in shared},
                    ))
            if not matched:
                continue

            ambiguous = len(matched) > 1
            source = f"WorkbookSheet::{sheet_name}#block:{block}#row:{row}"
            matched_accounts = [item[0] for item in matched]
            for account, periods, name_hint, account_values in matched:
                edges.append(_edge(
                    source=source,
                    target=account,
                    kind="supporting_schedule_tieout",
                    periods=periods,
                    test=(
                        f"labelled total row ties to account within max(1.0, "
                        f"{_TIE_TOLERANCE:.0%} x larger side) in every shared period; "
                        "one-period ties require an explicit account name in sheet/title; "
                        "match must be unique"
                    ),
                    passed=not ambiguous,
                    evidence={
                        "sheet": str(sheet_name),
                        "block": block,
                        "row": row,
                        "row_label": row_label,
                        "multiplier": multiplier,
                        "values": {p: values[p] for p in periods},
                        "account_values": account_values,
                        "name_hint": name_hint,
                        "matched_accounts": matched_accounts,
                        "ambiguous": ambiguous,
                    },
                ))
    return edges


def digest_tieout_diagnostics(
    workbook_digest: Optional[Dict[str, Any]],
    facts: Optional[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    """Why each unmapped sheet did not produce a supporting-schedule edge.

    Diagnostic only: it proposes no edge and is never persisted or rendered.
    """
    from ..financial_common import normalize_financial_date_label

    series = (facts or {}).get("series") or {}
    labels = (facts or {}).get("labels") or {}
    linked_sheets = {
        str((edge.get("evidence") or {}).get("sheet"))
        for edge in build_digest_graph_links(workbook_digest, facts)
        if edge.get("passed")
    }
    rows: List[Dict[str, Any]] = []
    for sheet_name, sheet in ((workbook_digest or {}).get("sheets") or {}).items():
        if not isinstance(sheet, dict) or sheet.get("status") != "unmapped":
            continue
        period_by_col, multiplier, total_rows = _digest_total_rows(sheet)
        labelled_totals = sorted({
            str(label).strip()
            for label in (sheet.get("row_labels") or {}).values()
            if _TOTAL_LABEL_RE.match(str(label).strip())
        })
        date_rows = []
        for row, label in (sheet.get("row_labels") or {}).items():
            normalised = normalize_financial_date_label(label)
            if _ISO_RE.match(str(normalised or "").strip()):
                date_rows.append(str(normalised))
        named_accounts = [
            str(account) for account in series
            if _sheet_names_account(
                str(sheet_name), sheet, str(account), str(labels.get(account) or ""))
        ]

        best = None
        for block, row, row_label, values in total_rows:
            for account, account_values in series.items():
                shared = sorted(
                    period for period in values
                    if isinstance(account_values.get(period), (int, float))
                    and abs(float(account_values[period])) > 1e-9
                )
                if not shared:
                    continue
                differences = [
                    abs(float(values[p]) - float(account_values[p]))
                    / max(1.0, abs(float(values[p])), abs(float(account_values[p])))
                    for p in shared
                ]
                candidate = {
                    "account": str(account),
                    "block": block,
                    "row": row,
                    "label": row_label,
                    "periods": shared,
                    "max_difference_pct": round(max(differences) * 100.0, 1),
                }
                rank = (candidate["max_difference_pct"], -len(shared), str(account))
                if best is None or rank < best[0]:
                    best = (rank, candidate)

        if str(sheet_name) in linked_sheets:
            reason = "tieout_found"
        elif not period_by_col:
            reason = "no_period_columns"
        elif not labelled_totals:
            reason = "no_labelled_total_row"
        elif not total_rows:
            reason = "total_row_has_no_period_values"
        else:
            reason = "total_did_not_tie"
        cells = sheet.get("cells") or {}
        numeric_nonzero = sum(1 for value in (cells.get("v") or []) if abs(float(value)) > 1e-9)
        rows.append({
            "sheet": str(sheet_name),
            "title": str(sheet.get("title") or ""),
            "reason": reason,
            "blocks": len(sheet.get("blocks") or []),
            "numeric_nonzero": numeric_nonzero,
            "period_columns": sorted(set(period_by_col.values())),
            "date_rows": sorted(set(date_rows)),
            "total_labels": labelled_totals,
            "unit_markers": list(sheet.get("unit_markers") or []),
            "multiplier": multiplier,
            "named_accounts": named_accounts,
            "best_near_match": best[1] if best else None,
        })
    return rows


def _edge_id(source: Any, target: Any, kind: Any, periods: Sequence[Any]) -> str:
    """Stable inside and across runs; no workbook or client name is persisted."""
    material = "\x1f".join((
        str(kind or ""),
        str(source or ""),
        str(target or ""),
        ",".join(sorted(str(p) for p in (periods or []))),
    ))
    return "edge_" + hashlib.sha256(material.encode("utf-8")).hexdigest()[:16]


def _edge(source, target, kind, periods, test, passed, evidence, tier="derived") -> Dict[str, Any]:
    periods = list(periods)
    return {
        "edge_id": _edge_id(source, target, kind, periods),
        "source": source,
        "target": target,
        "source_ref": _endpoint_ref(source),
        "target_ref": _endpoint_ref(target),
        "kind": kind,
        "tier": str(tier or "derived"),
        "periods": periods,
        "test": test,
        "passed": bool(passed),
        "evidence": evidence,
    }


def _typed_edge(raw: Dict[str, Any]) -> Dict[str, Any]:
    """Upgrade an edge supplied by another extraction seam to this schema."""
    edge = _edge(
        raw.get("source"),
        raw.get("target"),
        raw.get("kind"),
        raw.get("periods") or [],
        raw.get("test"),
        raw.get("passed"),
        raw.get("evidence") or {},
        tier=raw.get("tier") or "derived",
    )
    # Keep non-schema diagnostics a caller deliberately attached.
    known = set(edge) | {"edge_id"}
    edge.update({k: v for k, v in raw.items() if k not in known})
    return edge


def merge_graph_links(
    facts: Optional[Dict[str, Any]],
    extra_links: Optional[Sequence[Dict[str, Any]]],
) -> Dict[str, Any]:
    """Return one deduplicated typed graph without changing either input.

    Financials frames and account frames are currently assembled at different
    workbook seams. This is the single join point: callers hand over tested
    edges, and RunState receives one graph. No edge is made quotable merely by
    being merged; prompt rendering still admits only its explicit allow-list.
    """
    out = dict(facts or {})
    merged: Dict[str, Dict[str, Any]] = {}
    for raw in list(out.get("links") or []) + list(extra_links or []):
        if not isinstance(raw, dict):
            continue
        edge = _typed_edge(raw)
        edge_id = edge["edge_id"]
        if edge_id in merged and merged[edge_id] != edge:
            # Several Financials sheets can carry the same account/periods.
            # Until the workbook seam gives those rows an explicit sheet id,
            # do not silently let the later tested result erase the earlier.
            payload = json.dumps(edge, ensure_ascii=False, sort_keys=True, default=str)
            edge_id += "_" + hashlib.sha256(payload.encode("utf-8")).hexdigest()[:8]
            edge["edge_id"] = edge_id
        merged[edge_id] = edge
    out["links"] = list(merged.values())
    return out


def graph_summary(facts_or_links: Any) -> Dict[str, Any]:
    """Small persisted health view; the full graph remains in facts['links']."""
    links = (
        facts_or_links.get("links") or []
        if isinstance(facts_or_links, dict)
        else facts_or_links or []
    )
    by_kind: Dict[str, Dict[str, int]] = {}
    malformed = 0
    for edge in links:
        if not isinstance(edge, dict) or not edge.get("edge_id") or not edge.get("kind"):
            malformed += 1
            continue
        counts = by_kind.setdefault(str(edge["kind"]), {"total": 0, "passed": 0, "failed": 0})
        counts["total"] += 1
        counts["passed" if edge.get("passed") else "failed"] += 1
    return {
        "total": sum(v["total"] for v in by_kind.values()),
        "passed": sum(v["passed"] for v in by_kind.values()),
        "failed": sum(v["failed"] for v in by_kind.values()),
        "malformed": malformed,
        "by_kind": by_kind,
    }


def _component_to_total_edges(
    key: str, df: pd.DataFrame, series: Dict[str, float]
) -> List[Dict[str, Any]]:
    """Does the breakdown shown under the account sum to the account?

    ``tie_status`` cannot answer this. It is set only on the DETECTED path
    (schedules.py:1238); ``synthesize_detail_table_from_breakdown`` sets none,
    and measured across the four local databooks the tables are 5/14/4/5 present
    and 0/0/0/0 tied -- every one of them synthesized and untied. The plan's
    preferred fix is to compute it inside that function, which is in
    workbook/schedules.py and outside this milestone's scope, so the tie is
    computed here instead, off the same material (rows values, total_row, ISO
    periods) and with the same max(1.0, 2%) test.
    """
    table = (df.attrs or {}).get("presentation_detail_table") or {}
    rows = table.get("rows") or []
    periods = [str(p) for p in (table.get("periods") or [])]
    if not rows or not periods:
        return []
    total_row = table.get("total_row") or {}
    tied: List[str] = []
    differed: Dict[str, Dict[str, float]] = {}
    checked: List[str] = []
    for period in periods:
        account_total = series.get(period)
        if not isinstance(account_total, (int, float)) or abs(account_total) < 1e-9:
            continue
        block_values = total_row.get("values") or {}
        if period in block_values:
            block_total = float(block_values.get(period) or 0.0)
        else:
            block_total = 0.0
            for row in rows:  # top level only -- children are nested inside
                value = (row.get("values") or {}).get(period)
                block_total += float(value) if isinstance(value, (int, float)) else 0.0
        checked.append(period)
        if abs(block_total - account_total) <= max(1.0, abs(account_total) * _TIE_TOLERANCE):
            tied.append(period)
        else:
            differed[period] = {"breakdown": block_total, "account": account_total}
    if not checked:
        return []
    return [_edge(
        source=f"{key}::breakdown",
        target=key,
        kind="component_to_total",
        periods=checked,
        test=f"|breakdown - total| <= max(1.0, {_TIE_TOLERANCE:.0%} x total)",
        passed=bool(tied) and not differed,
        evidence={
            "component_count": len(rows),
            "tied_periods": tied,
            "differing_periods": differed,
            "used_total_row": bool(total_row),
        },
    )]


def _financials_edges(
    key: str,
    series: Dict[str, float],
    financials: Dict[str, Dict[str, float]],
    statement_type: str,
) -> List[Dict[str, Any]]:
    """Tab total against the Financials row, on EVERY date column.

    reconcile_financial_statements compares the latest column only
    (reconcile.py:481), so an account that agrees today and disagreed by a
    factor of ten two years ago reconciles clean. Widening it there would mean
    editing workbook/reconcile.py, outside this milestone's scope; the
    per-period comparison is done here instead, from the Financials rows the
    caller passes in.

    IS lines are compared on ABSOLUTE value, as reconciliation itself does
    (it converts expenses to positive before comparing). A first cut compared
    signed values and reported 7 of 24 accounts on one real databook as differing -- every
    one of them an exact sign flip (COGS 12,533,351 on the tab against
    -12,533,351 on the Financials sheet), i.e. the sheet's expense convention
    and not a difference at all. Reporting those as breaks would have made the
    edge list worthless in exactly the way tie_status was.
    """
    rows = financials.get(key)
    if not rows:
        return []
    absolute = str(statement_type).strip().upper() == "IS"
    agreed: List[str] = []
    differed: Dict[str, Dict[str, float]] = {}
    for period, fin_value in rows.items():
        tab_value = series.get(str(period))
        if not isinstance(tab_value, (int, float)) or not isinstance(fin_value, (int, float)):
            continue
        left, right = float(tab_value), float(fin_value)
        if absolute:
            left, right = abs(left), abs(right)
        scale = max(abs(right), abs(left))
        if scale < 1e-9:
            continue
        if abs(left - right) <= max(1.0, scale * 0.005):
            agreed.append(str(period))
        else:
            differed[str(period)] = {"tab": float(tab_value), "financials": float(fin_value)}
    if not agreed and not differed:
        return []
    return [_edge(
        source=key,
        target=f"Financials::{key}",
        kind="tab_to_financials",
        periods=sorted(agreed) + sorted(differed),
        test=("|tab - Financials| <= max(1.0, 0.5% x larger side), every date column"
              + (" (absolute values, IS)" if absolute else "")),
        passed=bool(agreed) and not differed,
        evidence={"agreeing_periods": agreed, "differing_periods": differed},
    )]


def _rollup_edges(key: str, df: pd.DataFrame, periods: Sequence[str]) -> List[Dict[str, Any]]:
    """Parent against its own children, per period.

    ``rollup_groups`` is written by _build_prompt_analysis_df only where the
    children were already verified to sum to the parent, so this edge is
    expected to pass wherever it exists at all. It is built anyway, because an
    edge that exists and passes is what the prompt builder and N3 read;
    "verified elsewhere and never recorded" is the state that produced
    tie_status. MEASURED: zero accounts across the four local databooks carry
    a non-empty rollup_groups, so this path has no material on the files we
    have and is exercised by construction only.
    """
    attrs = df.attrs or {}
    groups = attrs.get("rollup_groups") or {}
    if not groups:
        return []
    analysis = attrs.get("prompt_analysis_df")
    frame = analysis if isinstance(analysis, pd.DataFrame) and not analysis.empty else df
    desc_col = frame.columns[0]
    by_desc: Dict[str, Dict[str, float]] = {}
    for _idx, row in frame.iterrows():
        desc = str(row[desc_col])
        vals: Dict[str, float] = {}
        for period in periods:
            if period in frame.columns:
                raw = row[period]
                vals[period] = float(raw) if pd.notna(raw) and isinstance(raw, (int, float)) else 0.0
        by_desc[desc] = vals
    edges: List[Dict[str, Any]] = []
    for parent, children in groups.items():
        parent_vals = by_desc.get(str(parent))
        if not parent_vals or not children:
            continue
        agreed: List[str] = []
        differed: Dict[str, Dict[str, float]] = {}
        for period in periods:
            p_val = parent_vals.get(period)
            if p_val is None or abs(p_val) < 1e-9:
                continue
            c_val = sum((by_desc.get(str(c), {}) or {}).get(period, 0.0) for c in children)
            if abs(c_val - p_val) <= max(1.0, abs(p_val) * _TIE_TOLERANCE):
                agreed.append(period)
            else:
                differed[period] = {"children": c_val, "parent": p_val}
        if not agreed and not differed:
            continue
        edges.append(_edge(
            source=f"{key}::{parent}",
            target=key,
            kind="parent_to_children",
            periods=sorted(agreed) + sorted(differed),
            test=f"|sum(children) - parent| <= max(1.0, {_TIE_TOLERANCE:.0%} x parent)",
            passed=bool(agreed) and not differed,
            evidence={"parent": str(parent), "children": [str(c) for c in children],
                      "agreeing_periods": agreed, "differing_periods": differed},
        ))
    return edges


def _ratio_edges(
    accounts: Dict[str, Dict[str, Any]],
) -> List[Dict[str, Any]]:
    """A component series against another tab's total, across periods.

    The only genuinely CROSS-account candidate, and the only kind this module
    lets into a prompt. Both sides are read at the same period, over full
    periods only, and the whole series must land inside the band -- one period
    outside it rejects the edge rather than being reported as the interesting
    year, because a single outlier is far more often an extraction artefact or
    a related-party balance than a finding.
    """
    edges: List[Dict[str, Any]] = []
    for kind, source_role, target_role, unit, floor, high in _RATIO_CANDIDATES:
        targets = [k for k, a in accounts.items() if a["roles"].get(target_role)]
        if not targets:
            continue
        target_key = targets[0]
        target = accounts[target_key]
        for key, acct in accounts.items():
            if key == target_key:
                continue
            if not acct["roles"].get(source_role):
                continue
            shared = [p for p in acct["full_periods"] if p in target["full_periods"]]
            lengths = _period_lengths(target["periods"])
            values: Dict[str, float] = {}
            source_values: Dict[str, float] = {}
            target_values: Dict[str, float] = {}
            for period in shared:
                src = acct["series"].get(period)
                tgt = target["series"].get(period)
                if not isinstance(src, (int, float)) or not isinstance(tgt, (int, float)):
                    continue
                if abs(tgt) < 1e-9 or abs(src) < 1e-9:
                    continue
                source_values[period] = float(src)
                target_values[period] = float(tgt)
                if unit == "days":
                    days = max(1.0, lengths.get(period, 12.0)) * _DAYS_PER_MONTH
                    values[period] = abs(float(src)) / abs(float(tgt)) * days
                else:
                    values[period] = abs(float(src)) / abs(float(tgt)) * 100.0
            if len(values) < 2:
                continue
            outside = {p: v for p, v in values.items() if not (0.0 <= v <= high)}
            material = max(values.values()) >= floor
            edges.append(_edge(
                source=key,
                target=target_key,
                kind=kind,
                periods=sorted(values),
                test=(f"0 <= ratio <= {high:g} {unit} in every shared full period, "
                      f"and largest period >= {floor:g} {unit}"),
                passed=(not outside) and material,
                evidence={
                    "unit": unit,
                    "values": {p: round(v, 1) for p, v in sorted(values.items())},
                    "outside_band": {p: round(v, 1) for p, v in sorted(outside.items())},
                    "immaterial": not material,
                    "source_values": source_values,
                    "target_values": target_values,
                    "band": [0.0, high],
                    "materiality_floor": floor,
                },
            ))
    return edges


def _semantic_edges(accounts: Dict[str, Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Exact text relationships; recorded for graph use, never prompt-quoted.

    ``shared_counterparty`` requires the same company-like breakdown label in
    two account tables. ``remark_reference`` requires a prompt-carried note to
    name another account explicitly. Neither uses fuzzy similarity: a plausible
    text match is not evidence.
    """
    edges: List[Dict[str, Any]] = []

    # One hub edge per additional account avoids an N² clique when the same
    # counterparty appears in many schedules, while keeping every account
    # connected to the identity.
    occurrences: Dict[str, List[Tuple[str, str]]] = {}
    for key, account in accounts.items():
        table = (account["df"].attrs or {}).get("presentation_detail_table")
        seen = set()
        for label in _table_labels(table):
            normalised = _normalised_company_label(label)
            if normalised and normalised not in seen:
                occurrences.setdefault(normalised, []).append((key, label))
                seen.add(normalised)
    for normalised, found in sorted(occurrences.items()):
        found = sorted(found)
        if len(found) < 2:
            continue
        source_key, source_label = found[0]
        for target_key, target_label in found[1:]:
            shared = [
                p for p in accounts[source_key]["periods"]
                if p in accounts[target_key]["periods"]
            ]
            edges.append(_edge(
                source=f"{source_key}::{source_label}",
                target=f"{target_key}::{target_label}",
                kind="shared_counterparty",
                periods=shared,
                test="normalised company-like breakdown label is identical on both accounts",
                passed=True,
                evidence={
                    "normalised_label": normalised,
                    "source_label": source_label,
                    "target_label": target_label,
                },
            ))

    aliases: Dict[str, List[str]] = {}
    for key, account in accounts.items():
        attrs = account["df"].attrs or {}
        candidates = (
            str(key),
            str(account.get("label") or ""),
            str(attrs.get("source_sheet_name") or ""),
            str(attrs.get("block_title") or ""),
        )
        aliases[key] = list(dict.fromkeys(v.strip() for v in candidates if v.strip()))

    for source_key, account in accounts.items():
        attrs = account["df"].attrs or {}
        notes: List[Tuple[str, str]] = []
        for attr_name in ("supporting_notes", "table_linked_remarks"):
            notes.extend(_text_items(attrs.get(attr_name), attr_name))
        linked_targets = set()
        for origin, text in notes:
            for target_key, names in aliases.items():
                if target_key == source_key or target_key in linked_targets:
                    continue
                matched = next((name for name in names if _name_in_text(name, text)), None)
                if matched is None:
                    continue
                edges.append(_edge(
                    source=source_key,
                    target=target_key,
                    kind="remark_reference",
                    periods=[],
                    test="another account's exact name appears in prompt-carried remark text",
                    passed=True,
                    evidence={
                        "matched_name": matched,
                        "origin": origin,
                        "excerpt": text[:240],
                    },
                ))
                linked_targets.add(target_key)
    return edges


# A tiny cache so the wiring can stay one line at a per-agent-call site
# without rebuilding the table on every one of them. Correctness never
# depends on a hit -- a miss simply recomputes.
#
# ONLY the no-extras call is cached. Keying on the identity of `financials`
# or `type_lookup` would look right and be useless in the shape callers
# actually use: `type_lookup=lambda k: engine.get_mapping_component(...)`
# builds a fresh function object per call, so every lookup would miss and the
# cache would silently cost more than it saved. A caller passing either
# argument gets a fresh build; that caller is a diagnostic, not the hot path.
_CACHE: List[Tuple[Any, Dict[str, Any]]] = []
_CACHE_MAX = 4


def _cache_key(dfs: Dict[str, pd.DataFrame]) -> Tuple:
    return (id(dfs), tuple(sorted(str(k) for k in dfs)))


def build_cross_account_facts(
    dfs: Optional[Dict[str, pd.DataFrame]],
    financials: Optional[Dict[str, Dict[str, float]]] = None,
    type_lookup: Optional[Callable[[str], Any]] = None,
) -> Dict[str, Any]:
    """The per-run cross-account fact table and its verified edge list.

    ``financials`` is optional: {account key: {ISO period: value}} from the
    Financials sheet. pipeline.py does not have the Financials frames at the
    seam that calls this, so in production today it is None and the
    ``tab_to_financials`` edges are simply absent -- rather than reimplementing
    the alias matching here and risking a second, differently-wrong answer.
    ``build_financials_by_key`` builds it for a caller that does have them.

    Returns::

        {
          "totals":      {(account, ISO period): float},   # RunState.facts
          "series":      {account: {ISO period: float}},   # the same, usable
          "labels":      {account: str},                   # the tab's caption
          "periods":     {account: [ISO period, ...]},     # tail INCLUDED
          "full_periods":{account: [ISO period, ...]},     # stub dropped
          "statement_types": {account: "BS"|"IS"|""},
          "links":       [edge, ...],                      # passed AND failed
          "sources":     {account: "prompt_analysis_df"|"frame"},
        }

    An edge is ``{source, target, kind, periods, test, passed, evidence}``.
    Read passed edges with ``cross_account_links_for``; the failed ones are
    kept deliberately, for N3's internal summary, and must never be rendered
    into a prompt.
    """
    dfs = dfs or {}
    cacheable = financials is None and type_lookup is None
    key = _cache_key(dfs)
    if cacheable:
        for cached_key, cached in _CACHE:
            if cached_key == key:
                return cached

    accounts: Dict[str, Dict[str, Any]] = {}
    for name, df in dfs.items():
        if not isinstance(df, pd.DataFrame) or df.empty:
            continue
        periods, series, source = _account_series(df)
        if not periods:
            continue
        statement_type = resolve_statement_type(str(name), df, type_lookup)
        roles = {role: _has_role(role, str(name), df)
                 for role in ("receivable", "payable", "inventory", "revenue",
                              "cost", "advance", "prepayment")}
        roles["is_expense"] = bool(
            statement_type == "IS" and not roles["revenue"]
            and not any(n in str(name).lower() for n in ("non-operating", "营业外", "其他收益"))
        )
        try:
            label = str(df.columns[0])
        except Exception:
            label = str(name)
        accounts[str(name)] = {
            "df": df,
            "label": label,
            "periods": periods,
            "full_periods": _full_periods(periods),
            "series": series,
            "source": source,
            "statement_type": statement_type,
            "roles": roles,
        }

    links: List[Dict[str, Any]] = []
    for name, acct in accounts.items():
        links.extend(_component_to_total_edges(name, acct["df"], acct["series"]))
        links.extend(_rollup_edges(name, acct["df"], acct["periods"]))
        if financials:
            links.extend(_financials_edges(
                name, acct["series"], financials, acct["statement_type"]))
    links.extend(_ratio_edges(accounts))
    links.extend(_semantic_edges(accounts))

    totals: Dict[Tuple[str, str], float] = {}
    for name, acct in accounts.items():
        for period, value in acct["series"].items():
            totals[(name, str(period))] = value

    facts = {
        "totals": totals,
        # The tab's own first-column caption, so a sentence can NAME the other
        # account the way the databook does ('Operating income', '营业收入')
        # rather than by its mapping key, which is often an abbreviation the
        # reader has never seen ('OP', 'OR', 'COGS').
        "labels": {n: a["label"] for n, a in accounts.items()},
        "series": {n: dict(a["series"]) for n, a in accounts.items()},
        "periods": {n: list(a["periods"]) for n, a in accounts.items()},
        "full_periods": {n: list(a["full_periods"]) for n, a in accounts.items()},
        "statement_types": {n: a["statement_type"] for n, a in accounts.items()},
        "sources": {n: a["source"] for n, a in accounts.items()},
        "links": links,
    }
    if cacheable:
        _CACHE.append((key, facts))
        del _CACHE[:-_CACHE_MAX]
    return facts


def build_financials_by_key(
    bs_is_results: Optional[Dict[str, Any]],
    dfs: Optional[Dict[str, pd.DataFrame]],
    mappings: Optional[dict] = None,
) -> Dict[str, Dict[str, float]]:
    """{account key: {ISO period: Financials value}} for every date column.

    Reuses reconcile.py's own ``find_account_in_dfs`` so a row is attributed to
    the same tab reconciliation would attribute it to -- a second alias
    matcher here would eventually disagree with the first one, and the
    disagreement would look like a data problem. Imported lazily: this module
    sits under fdd_utils.ai and workbook is the heavier package.
    """
    if not bs_is_results or not dfs:
        return {}
    try:
        from fdd_utils.workbook.reconcile import find_account_in_dfs
        from fdd_utils.workbook.mapping import load_mappings
    except Exception:
        return {}
    if mappings is None:
        try:
            mappings = load_mappings()
        except Exception:
            return {}
    out: Dict[str, Dict[str, float]] = {}
    for statement in ("balance_sheet", "income_statement"):
        frame = bs_is_results.get(statement)
        if not isinstance(frame, pd.DataFrame) or frame.empty:
            continue
        date_cols = [c for c in frame.columns if str(c) != "Description"]
        for _idx, row in frame.iterrows():
            account_name = str(row.get("Description") or "").strip()
            if not account_name:
                continue
            try:
                dfs_key, _df, _cat, _mk, _status, _note = find_account_in_dfs(
                    account_name, dfs, mappings)
            except Exception:
                continue
            if not dfs_key or dfs_key == "SKIP" or dfs_key not in dfs:
                continue
            values: Dict[str, float] = {}
            for col in date_cols:
                raw = row.get(col)
                if isinstance(raw, (int, float)) and pd.notna(raw):
                    values[str(col)] = float(raw)
            if values:
                out.setdefault(str(dfs_key), {}).update(values)
    return out


def financials_tieout(
    bs_is_results: Optional[Dict[str, Any]],
    dfs: Optional[Dict[str, pd.DataFrame]],
    mappings: Optional[dict] = None,
) -> List[Dict[str, Any]]:
    """One ``tab_to_financials`` edge per account, tested on every date column.

    The statement a row sits in decides the comparison (IS on absolute value),
    so each statement is attributed separately rather than asking a mapping
    for the account's type. A real deck is why this is callable on its own:
    its summary said 投资收益 was nil in every period while the statement table
    beside it printed 1,332 and 602, and its 长期应付款 bullet said the account
    first carried a balance in 2025 against 6,145 and 8,269 in that table. Both
    tabs agreed with Financials on the latest column, which is the only one
    reconciliation reads.
    """
    if not bs_is_results or not dfs:
        return []
    series = build_cross_account_facts(dfs).get("series") or {}
    edges: List[Dict[str, Any]] = []
    for statement, kind in (("balance_sheet", "BS"), ("income_statement", "IS")):
        frame = bs_is_results.get(statement)
        rows = build_financials_by_key({statement: frame}, dfs, mappings)
        for key in rows:
            edges.extend(_financials_edges(key, series.get(key) or {}, rows, kind))
    return edges


def cross_account_links_for(
    facts: Optional[Dict[str, Any]],
    account: str,
    passed_only: bool = True,
    source_only: bool = False,
) -> List[Dict[str, Any]]:
    """Edges touching one account, strongest kind first.

    ``passed_only`` defaults True and the prompt builder never overrides it.
    The False form is for the internal insight summary, which is allowed to
    read a rejected hypothesis and a prompt is not.

    ``source_only`` matters more than it looks. A ratio edge is directional:
    "AR equates to 126 days of revenue" belongs in AR's commentary and is a
    non-sequitur in the revenue account's, yet the revenue account is the
    edge's TARGET and so matches an undirected lookup. The prompt builder
    passes source_only=True; N3, which wants everything an account is
    implicated in, does not.
    """
    if not isinstance(facts, dict):
        return []
    order = {"receivable_days": 0, "payable_days": 1, "inventory_days": 2,
             "expense_to_revenue": 3, "parent_to_children": 4,
             "component_to_total": 5, "shared_counterparty": 6,
             "remark_reference": 7, "supporting_schedule_tieout": 8,
             "tab_to_financials": 9}
    found = []
    for edge in (facts.get("links") or []):
        if passed_only and not edge.get("passed"):
            continue
        source = str(edge.get("source", "")).split("::")[0]
        target = str(edge.get("target", "")).split("::")[0]
        if source == account or (not source_only and target == account):
            found.append(edge)
    return sorted(found, key=lambda e: order.get(str(e.get("kind")), 99))

from __future__ import annotations

# re-added: bound by an import in another section of the pre-split module
from ..financial_common import cell_text, coerce_numeric, normalize_financial_date_label

"""
Workbook profiling helpers for financial databooks.

This module inspects sheet structure without applying business mappings so later
steps can resolve tabs and normalize values using real workbook metadata.
"""



from functools import lru_cache
import logging
import re
import time
from typing import Any, Callable, Dict, List, Optional

import pandas as pd

from ..financial_common import cell_text

logger = logging.getLogger(__name__)

CANONICAL_STAGE_LABELS = (
    # "示意性调整数" ("indicative adjustment FIGURE") is a client-specific way of
    # writing "示意性调整后" ("after indicative adjustment") -- the "数" suffix
    # here mirrors "审定数" (audited FIGURE, i.e. the final audited value, not
    # the audit delta) rather than meaning "the adjustment amount itself".
    # Listed here (checked before "Indicative adjustment" below) so it doesn't
    # fall through to matching the shorter "示意性调整" substring as the delta
    # stage instead of the final-balance stage the pipeline actually wants
    # (PREFERRED_STAGE = "Indicative adjusted").
    ("Indicative adjusted", ("indicative adjusted", "indivative adjusted", "示意性调整后", "示意性調整後",
                              "示意性调整数", "示意性調整數")),
    ("Indicative adjustment", ("indicative adjustment", "indivative adjustment", "示意性调整", "示意性調整")),
    ("Audited", ("audited", "审定数", "審定數")),
    ("Audit adjustment", ("audit adjustment", "审计调整", "審計調整")),
    ("Mgt acc", ("mgt acc", "management account", "管理层数", "管理層數")),
)

_UNIT_PATTERNS = (
    "CNY'000",
    "USD'000",
    "HKD'000",
    "人民币千元",
    "人民幣千元",
    "million",
    "百万",
    "百萬",
)

_TITLE_SKIP_VALUES = {
    "",
    "back",
    "nan",
    "none",
}


@lru_cache(maxsize=8)
def load_workbook_frames(workbook_path: str) -> Dict[str, pd.DataFrame]:
    """Load all workbook sheets once for profiling/normalization reuse."""
    return pd.read_excel(workbook_path, sheet_name=None, header=None, engine="openpyxl")


def _cell_text(value: Any) -> str:
    return cell_text(value)


def _coerce_numeric(value: Any) -> Optional[float]:
    return coerce_numeric(value)


def _normalize_spaces(text: Any) -> str:
    return re.sub(r"\s+", " ", str(text or "")).strip()


def canonical_stage_label(value: Any) -> Optional[str]:
    text = _normalize_spaces(cell_text(value)).lower()
    if not text or len(text) > 35:
        return None
    for canonical, variants in CANONICAL_STAGE_LABELS:
        if any(variant in text for variant in variants):
            return canonical
    return None


def contains_indicative_text(value: Any) -> bool:
    return canonical_stage_label(value) == "Indicative adjusted"


def contains_unit_marker(value: Any, markers: tuple[str, ...] = _UNIT_PATTERNS) -> bool:
    text = _normalize_spaces(cell_text(value)).lower()
    return any(marker.lower() in text for marker in markers)


def stage_row_indices(
    df: pd.DataFrame,
    date_detector: Callable[[Any], Any],
    max_rows: int = 80,
) -> List[int]:
    indices: List[int] = []
    for row_idx in range(min(max_rows, len(df))):
        hits = sum(1 for value in df.iloc[row_idx].tolist() if canonical_stage_label(value))
        if hits >= 2:
            indices.append(row_idx)
            continue
        if hits == 1:
            nearby_date_row = any(
                sum(1 for value in df.iloc[candidate_idx].tolist() if date_detector(value)) >= 1
                for candidate_idx in range(max(0, row_idx - 1), min(len(df), row_idx + 3))
            )
            if nearby_date_row:
                indices.append(row_idx)
    return indices


def primary_stage_row_index(
    df: pd.DataFrame,
    max_rows: int = 12,
) -> Optional[int]:
    best_idx = None
    best_hits = 0
    for row_idx in range(min(max_rows, len(df))):
        hits = sum(1 for value in df.iloc[row_idx].tolist() if canonical_stage_label(value))
        if hits > best_hits:
            best_hits = hits
            best_idx = row_idx
    return best_idx if best_hits > 0 else None


def date_row_index(
    df: pd.DataFrame,
    stage_row_idx: Optional[int],
    date_detector: Callable[[Any], Any],
    max_distance: int = 2,
) -> Optional[int]:
    if stage_row_idx is None:
        return None
    best_idx = None
    best_hits = 0
    for row_idx in range(max(0, stage_row_idx - max_distance), min(len(df), stage_row_idx + max_distance + 1)):
        hits = sum(1 for value in df.iloc[row_idx].tolist() if date_detector(value))
        if hits > best_hits:
            best_hits = hits
            best_idx = row_idx
    return best_idx if best_hits > 0 else None


def _unit_markers(df: pd.DataFrame, max_rows: int = 8) -> List[str]:
    found: List[str] = []
    rows = min(max_rows, len(df))
    for row_idx in range(rows):
        for value in df.iloc[row_idx].tolist():
            text = _cell_text(value)
            for marker in _UNIT_PATTERNS:
                if contains_unit_marker(text, (marker,)) and marker not in found:
                    found.append(marker if marker != "人民幣千元" else "人民币千元")
    return found


def _stage_row_index(df: pd.DataFrame, max_rows: int = 12) -> Optional[int]:
    return primary_stage_row_index(df, max_rows=max_rows)


def _looks_like_entity_heading(text: str) -> bool:
    normalized = _normalize_spaces(text)
    if not normalized:
        return False
    if re.search(r"\s[-–]\s", normalized):
        suffix = re.split(r"\s[-–]\s", normalized, maxsplit=1)[-1].strip()
        if suffix and re.search(r"[A-Za-z\u4e00-\u9fff]", suffix):
            return True
    return any(token in normalized for token in ("公司", "集团", "集團", "company", "Company", "Ltd", "Limited"))


def _stage_block_heading(df: pd.DataFrame, stage_row_idx: int) -> str:
    for row_idx in range(stage_row_idx - 1, max(-1, stage_row_idx - 4), -1):
        values = [_normalize_spaces(_cell_text(value)) for value in df.iloc[row_idx].tolist()]
        texts = [value for value in values if value.lower() not in _TITLE_SKIP_VALUES]
        if not texts:
            continue
        if any(canonical_stage_label(value) for value in texts):
            continue
        # Deliberately permissive, unlike the detection path above: this is a
        # SKIP test, and a row of bare numbers was already being skipped here
        # because it parsed as serial dates. Tightening it would let such a row
        # through and return it as the entity heading.
        if any(_parse_date_label(value) for value in texts):
            continue
        if any(any(marker.lower() in value.lower() for marker in _UNIT_PATTERNS) for value in texts):
            continue
        return " ".join(texts)
    return ""


def _parse_date_label(value: Any, allow_bare_number: bool = True) -> Optional[str]:
    from .statements import parse_date  # local: breaks the inspector<->statements import cycle
    parsed = parse_date(value, allow_bare_number=allow_bare_number)
    if not parsed:
        return None
    return parsed.strftime("%Y-%m-%d")


def _date_row_index(df: pd.DataFrame, stage_row_idx: Optional[int], max_distance: int = 2) -> Optional[int]:
    # Bare numbers are refused while CHOOSING the row: a row of 千元 balances
    # outscores the real date row otherwise (see parse_date's own note).
    return date_row_index(
        df, stage_row_idx,
        lambda value: _parse_date_label(value, allow_bare_number=False),
        max_distance=max_distance,
    )


def _title_row_index(df: pd.DataFrame, stage_row_idx: Optional[int]) -> Optional[int]:
    search_limit = stage_row_idx if stage_row_idx is not None else min(6, len(df))
    for row_idx in range(min(search_limit + 1, len(df))):
        row = [_normalize_spaces(_cell_text(value)) for value in df.iloc[row_idx].tolist()]
        texts = [value for value in row if value.lower() not in _TITLE_SKIP_VALUES]
        if len(texts) == 1:
            return row_idx
    return None


def _sheet_title(df: pd.DataFrame, stage_row_idx: Optional[int]) -> str:
    title_idx = _title_row_index(df, stage_row_idx)
    if title_idx is not None:
        values = [_normalize_spaces(_cell_text(value)) for value in df.iloc[title_idx].tolist()]
        for value in values:
            if value.lower() not in _TITLE_SKIP_VALUES:
                return value
    for row_idx in range(min(6, len(df))):
        for value in df.iloc[row_idx].tolist():
            text = _normalize_spaces(_cell_text(value))
            lowered = text.lower()
            if lowered in _TITLE_SKIP_VALUES:
                continue
            if canonical_stage_label(text) or _parse_date_label(text):
                continue
            if any(marker.lower() in lowered for marker in ("cny'000", "人民币千元", "人民幣千元")):
                continue
            return text
    return "Sheet"


def _stage_labels(df: pd.DataFrame, stage_row_idx: Optional[int]) -> List[str]:
    return _collect_row_labels(df, stage_row_idx, canonical_stage_label)


def _date_labels(df: pd.DataFrame, date_row_idx: Optional[int]) -> List[str]:
    return _collect_row_labels(df, date_row_idx, _parse_date_label)


def _collect_row_labels(
    df: pd.DataFrame,
    row_idx: Optional[int],
    label_parser,
) -> List[str]:
    if row_idx is None:
        return []
    seen: List[str] = []
    for value in df.iloc[row_idx].tolist():
        label = label_parser(value)
        if label and label not in seen:
            seen.append(label)
    return seen


# Tab-naming conventions observed identically across every real TS-team databook
# seen so far (9 files, different clients/projects/years) for navigation dividers,
# fixed template helper tabs, and tool-generated artifacts -- never a real account,
# regardless of what score fuzzy matching or stage-row detection gives them (e.g.
# 'ADJ' sometimes has a detectable stage row since audit adjustments are tracked
# by period, which let it slip past the sheet_kind=="other" score floor once).
# Checked ahead of stage_row_idx so these are excluded deterministically rather
# than relying on the score floor to catch them.
_TEMPLATE_NAV_EXACT_NAMES = {
    "overview", "adj", "contract", "tb", "cover", "mapping", "je", "pt",
    "pivottable", "tb combine", "mgt account", "fill-in choice",
    "封面", "下拉选项source", "透视表",
}
_TEMPLATE_NAV_SUFFIXES = ("-->", " for report")
_TEMPLATE_NAV_PREFIXES = ("upslide_", "_tm_")


def _is_template_nav_sheet(sheet_name: str) -> bool:
    lowered = sheet_name.strip().lower()
    if lowered in _TEMPLATE_NAV_EXACT_NAMES:
        return True
    if any(lowered.endswith(suffix) for suffix in _TEMPLATE_NAV_SUFFIXES):
        return True
    if any(lowered.startswith(prefix) for prefix in _TEMPLATE_NAV_PREFIXES):
        return True
    return False


def _sheet_kind(sheet_name: str, title: str, df: pd.DataFrame, stage_row_idx: Optional[int]) -> str:
    if _is_template_nav_sheet(sheet_name):
        return "template_nav"
    title_lower = title.lower()
    sheet_lower = sheet_name.lower()
    sample = " ".join(_cell_text(value).lower() for value in df.head(min(10, len(df))).fillna("").to_numpy().ravel())
    has_balance = "balance sheet" in sample or "资产负债表" in sample or "資產負債表" in sample
    has_income = "income statement" in sample or "profit and loss" in sample or "利润表" in sample or "利潤表" in sample
    if "financials" in sheet_lower or (has_balance and has_income):
        return "financial_summary"
    if stage_row_idx is not None:
        return "financial_schedule"
    if "ledger" in title_lower or "台账" in sample or "台賬" in sample:
        return "support_schedule"
    return "other"


def _entity_scope(sheet_kind: str, df: pd.DataFrame) -> str:
    if sheet_kind == "financial_summary":
        return "single"
    stage_blocks = stage_row_indices(df, _parse_date_label)
    if len(stage_blocks) >= 2:
        entity_headings = {
            heading
            for heading in (_stage_block_heading(df, stage_row_idx) for stage_row_idx in stage_blocks)
            if _looks_like_entity_heading(heading)
        }
        if len(entity_headings) >= 2:
            return "multiple"
    if len(stage_blocks) >= 3 and any(len(_stage_labels(df, idx)) <= 1 for idx in stage_blocks):
        return "multiple"
    return "single"


def _header_signature(df: pd.DataFrame, stage_row_idx: Optional[int], date_row_idx: Optional[int]) -> Dict[str, Any]:
    signature: Dict[str, Any] = {}
    if stage_row_idx is not None:
        signature["stage_row_idx"] = stage_row_idx
    if date_row_idx is not None:
        signature["date_row_idx"] = date_row_idx
    return signature


def profile_sheet(df: pd.DataFrame, sheet_name: str) -> Dict[str, Any]:
    stage_row_idx = _stage_row_index(df)
    date_row_idx = _date_row_index(df, stage_row_idx)
    title = _sheet_title(df, stage_row_idx)
    title_row_idx = _title_row_index(df, stage_row_idx)
    sheet_kind = _sheet_kind(sheet_name, title, df, stage_row_idx)
    stage_labels = _stage_labels(df, stage_row_idx)
    date_labels = _date_labels(df, date_row_idx)
    unit_markers = _unit_markers(df)
    entity_scope = _entity_scope(sheet_kind, df)

    return {
        "sheet_name": sheet_name,
        "title": title,
        "title_row_idx": title_row_idx,
        "sheet_kind": sheet_kind,
        "entity_scope": entity_scope,
        "stage_row_idx": stage_row_idx,
        "date_row_idx": date_row_idx,
        "stage_labels": stage_labels,
        "date_labels": date_labels,
        "unit_markers": unit_markers,
        "has_indicative_stage": "Indicative adjusted" in stage_labels,
        "header_signature": _header_signature(df, stage_row_idx, date_row_idx),
    }


@lru_cache(maxsize=8)
def profile_workbook(workbook_path: str) -> Dict[str, Dict[str, Any]]:
    started = time.perf_counter()
    workbook_frames = load_workbook_frames(workbook_path)
    profiles: Dict[str, Dict[str, Any]] = {}
    for sheet_name, df in workbook_frames.items():
        profiles[sheet_name] = profile_sheet(df, sheet_name)
    logger.debug(
        "Workbook profiler scanned %s sheets from %s in %.2fs",
        len(profiles),
        workbook_path,
        time.perf_counter() - started,
    )
    return profiles
# --- end workbook/inspector.py ---

"""The remarks a schedule carries beside its figures, as the app shows them next to an account's commentary.

Three readings of one frame's attributes: the side-column ("adjacent detail") rows that are still visible, a short summary of what they
say, and the remarks the extractor linked to a table. They were methods of the old prompt engine, which wrote them into the model's
prompt; the agent reads the same remarks as notes in the evidence contract, and the app only shows them.
"""
from __future__ import annotations

from typing import Any, Dict, List, Optional

import pandas as pd

from ..financial_common import visible_descriptions
from . import INTERNAL_ROW_KEY


def _skip_key(key_text: str, include_description: bool = False) -> bool:
    return (
        key_text == INTERNAL_ROW_KEY
        or (include_description and key_text == "Description")
        or key_text.endswith("| table_header")
        or key_text.endswith("| indicative_adjusted_row")
        or key_text.endswith("| date_row")
    )


def filter_adjacent_detail_rows(df: pd.DataFrame) -> List[Dict[str, Any]]:
    adjacent_detail_rows = df.attrs.get("adjacent_detail_rows") or []
    if not adjacent_detail_rows:
        return []
    visible_rows = visible_descriptions(df)
    if not visible_rows:
        return adjacent_detail_rows
    excluded = set(df.attrs.get("prompt_budget_exclude") or [])
    filtered = [
        row for row in adjacent_detail_rows
        if str(row.get("Description", "")).strip() in visible_rows
        and str(row.get("Description", "")).strip() not in excluded
    ]
    return filtered or [
        row for row in adjacent_detail_rows
        if str(row.get("Description", "")).strip() not in excluded
    ]


def table_linked_remarks(df: Optional[pd.DataFrame]) -> List[Dict[str, Any]]:
    if not isinstance(df, pd.DataFrame):
        return []
    remarks = df.attrs.get("table_linked_remarks") or []
    excluded = set(df.attrs.get("prompt_budget_exclude") or [])
    return [
        remark for remark in remarks
        if isinstance(remark, dict)
        and str(remark.get("description") or "").strip() not in excluded
    ]


def summarize_rhs_remarks(adjacent_detail_rows: List[Dict[str, Any]], language: str) -> List[str]:
    if not adjacent_detail_rows:
        return []

    summaries: List[str] = []
    seen: set = set()
    for row in adjacent_detail_rows:
        if not isinstance(row, dict):
            continue
        description = str(row.get("Description") or "").strip()
        remark_parts: List[str] = []
        for key, value in row.items():
            key_text = str(key)
            if _skip_key(key_text, include_description=True):
                continue
            if isinstance(value, (int, float)):
                continue
            text = str(value or "").strip()
            if not text:
                continue
            if key_text == description:
                continue
            remark_parts.append(text)

        unique_parts: List[str] = []
        seen_parts: set = set()
        for part in remark_parts:
            if part in seen_parts:
                continue
            seen_parts.add(part)
            unique_parts.append(part)

        if not unique_parts:
            continue
        if language == "Chi":
            summary = f"{description}: " + "；".join(unique_parts[:2]) if description else "；".join(unique_parts[:2])
        else:
            summary = f"{description}: " + "; ".join(unique_parts[:2]) if description else "; ".join(unique_parts[:2])
        if summary not in seen:
            seen.add(summary)
            summaries.append(summary)
    return summaries[:5]

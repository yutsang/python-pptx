"""
Financial DataFrame to JSON converter for FDD AI prompts.

This module owns prompt-oriented serialization only. It intentionally
consumes display-formatted columns when available, but leaves display
formatting and generic parsing helpers to sibling modules.
"""

from __future__ import annotations

import json
from typing import Callable, Dict, List, Optional

import pandas as pd

from .financial_common import contains_chinese_text


def _looks_chinese(text: str) -> bool:
    return contains_chinese_text(text)


def _detect_language(df: pd.DataFrame, table_name: str) -> str:
    if _looks_chinese(table_name):
        return "Chi"
    if df is None or df.empty:
        return "Eng"
    first_col = str(df.columns[0]) if len(df.columns) else ""
    if _looks_chinese(first_col):
        return "Chi"
    sample = " ".join(str(v) for v in df.iloc[:, 0].head(5).tolist())
    return "Chi" if _looks_chinese(sample) else "Eng"


def _format_numeric_value(value) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, float):
        return f"{value:,.2f}".rstrip("0").rstrip(".")
    if isinstance(value, int):
        return f"{value:,}"
    return str(value)


def _normalize_text(value: str, text_normalizer: Optional[Callable[[str], str]] = None) -> str:
    text = str(value or "").strip()
    if text_normalizer is None or not text:
        return text
    return str(text_normalizer(text)).strip()


def _display_fields(
    df: pd.DataFrame,
    text_normalizer: Optional[Callable[[str], str]] = None,
) -> List[str]:
    return [
        _normalize_text(str(col).replace("_formatted", ""), text_normalizer)
        for col in df.columns[1:]
        if not str(col).endswith("_formatted")
    ]


def _display_row(
    row: pd.Series,
    df: pd.DataFrame,
    text_normalizer: Optional[Callable[[str], str]] = None,
) -> Dict[str, str]:
    first_col = str(df.columns[0])
    out = {"category": _normalize_text(str(row.get(first_col, "")).strip(), text_normalizer)}

    for col in df.columns[1:]:
        col_name = str(col)
        if col_name.endswith("_formatted"):
            continue
        formatted_col = f"{col_name}_formatted"
        display_key = _normalize_text(col_name.replace("_formatted", ""), text_normalizer)
        if formatted_col in df.columns and pd.notna(row.get(formatted_col)):
            out[display_key] = _normalize_text(str(row.get(formatted_col)).strip(), text_normalizer)
        else:
            out[display_key] = _normalize_text(_format_numeric_value(row.get(col)), text_normalizer)
    return out


def table_to_json_dict(
    df: pd.DataFrame,
    table_name: str = "Table",
    language: Optional[str] = None,
    currency: Optional[str] = None,
    text_normalizer: Optional[Callable[[str], str]] = None,
) -> Optional[Dict]:
    if df is None or df.empty:
        return None

    detected_language = language or _detect_language(df, table_name)
    resolved_currency = currency or "CNY"
    if detected_language == "Chi":
        value_scale = "Values are pre-formatted display amounts using Chinese reporting units such as 万 or 亿. Use exactly as shown."
    else:
        value_scale = "Values are pre-formatted display amounts using reporting units such as K or million. Use exactly as shown."

    result = {
        "table_name": _normalize_text(table_name, text_normalizer),
        "currency": resolved_currency,
        "value_scale": value_scale,
        "fields": _display_fields(df, text_normalizer=text_normalizer),
        "data": [_display_row(row, df, text_normalizer=text_normalizer) for _, row in df.iterrows()],
    }
    integrity = df.attrs.get("integrity")
    if integrity:
        result["integrity"] = integrity
    return result


def df_to_json_str(
    df: pd.DataFrame,
    table_name: str = "Table",
    language: Optional[str] = None,
    currency: Optional[str] = None,
    indent: int = 2,
    text_normalizer: Optional[Callable[[str], str]] = None,
) -> str:
    result = table_to_json_dict(
        df,
        table_name=table_name,
        language=language,
        currency=currency,
        text_normalizer=text_normalizer,
    )
    if not result:
        return "[]"
    return json.dumps(result, ensure_ascii=False, indent=indent)

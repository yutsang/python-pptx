from __future__ import annotations

"""Cross-cutting financial helpers shared across workbook, AI, UI, and PPTX."""

import os
import re
from typing import Any, Dict, Iterable, List, Sequence

import pandas as pd
import yaml


def package_file_path(filename: str) -> str:
    return os.path.join(os.path.dirname(__file__), filename)


def load_yaml_file(path: str) -> Dict[str, Any]:
    with open(path, "r", encoding="utf-8") as handle:
        return yaml.safe_load(handle) or {}


def load_required_yaml_file(path: str) -> Dict[str, Any]:
    if not os.path.exists(path):
        raise FileNotFoundError(f"Config file not found: {path}")
    config = load_yaml_file(path)
    if config is None or not config:
        raise ValueError(f"Config file is empty: {path}")
    return config


def cell_text(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    if hasattr(value, "strftime"):
        try:
            return value.strftime("%Y-%m-%d")
        except Exception:
            pass
    return str(value).strip()


def coerce_numeric(value: Any) -> float | None:
    if value is None or pd.isna(value):
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return None
    cleaned = text.replace(",", "").replace("，", "")
    cleaned = cleaned.replace("(", "-").replace(")", "")
    try:
        return float(cleaned)
    except (TypeError, ValueError):
        return None


def dedupe_non_empty(values: Iterable[Any]) -> List[str]:
    seen = set()
    normalized: List[str] = []
    for value in values:
        text = str(value or "").strip()
        if not text or text in seen:
            continue
        seen.add(text)
        normalized.append(text)
    return normalized


def visible_descriptions(df: pd.DataFrame) -> set[str]:
    if df is None or df.empty or len(df.columns) == 0:
        return set()
    return {str(value).strip() for value in df.iloc[:, 0].tolist() if str(value).strip()}


def contains_chinese_text(text: str) -> bool:
    return bool(re.search(r"[\u4e00-\u9fff]", str(text or "")))


def contains_predominantly_chinese_text(text: str, threshold: float = 0.3) -> bool:
    normalized = str(text or "")
    if not normalized:
        return False
    chinese_chars = sum(1 for char in normalized if "\u4e00" <= char <= "\u9fff")
    total_chars = len(normalized)
    if total_chars == 0:
        return False
    return (chinese_chars / total_chars) > threshold


def extract_result_text_content(value: Any) -> str:
    if isinstance(value, dict):
        value = value.get("output") or value.get("final_content") or value.get("content") or ""
    return str(value or "")


def get_pipeline_result_text(
    result_dict: Any,
    priority: Sequence[str] = ("final", "subagent_4", "subagent_3", "subagent_2", "subagent_1"),
) -> str:
    if isinstance(result_dict, str):
        return result_dict
    if not isinstance(result_dict, dict):
        return ""

    for key in priority:
        content = extract_result_text_content(result_dict.get(key, ""))
        if content.strip():
            return content
    return ""


def clean_english_placeholders(text: str) -> str:
    if not isinstance(text, str):
        return text

    placeholder_patterns = {
        r"_status": "",
        r"_type": "",
        r"_name": "",
        r"_code": "",
        r"_id": "",
        r"leasing_status_status": "租赁状态",
        r"leasing_": "租赁",
        r"status_": "",
        r"_\w+_\w+": "",
    }

    cleaned_text = text
    for pattern, replacement in placeholder_patterns.items():
        cleaned_text = re.sub(pattern, replacement, cleaned_text, flags=re.IGNORECASE)

    if re.search(r"[a-zA-Z_]{3,}", cleaned_text):
        cleaned_text = re.sub(r"[a-zA-Z]+_[a-zA-Z_]+", "", cleaned_text)
    return cleaned_text.strip()


def normalize_financial_date_label(date_str: Any) -> Any:
    if isinstance(date_str, str):
        if re.fullmatch(r"FY\d{2}", date_str):
            year = int("20" + date_str[2:])
            return f"{year}-12-31"
        if re.fullmatch(r"\d{1,2}M\d{2}", date_str) and "月" not in date_str:
            try:
                months, year_suffix = date_str.split("M")
                year = int("20" + year_suffix)
                end_month = int(months)
                if end_month <= 12:
                    result = pd.to_datetime(f"{year}-{end_month}-01") + pd.tseries.offsets.MonthEnd(0)
                    return result.strftime("%Y-%m-%d")
            except ValueError:
                pass

        match_full_date = re.match(r"(\d{4})年(\d{1,2})月(\d{1,2})日$", date_str)
        if match_full_date:
            year, month, day = match_full_date.groups()
            return f"{year}-{month.zfill(2)}-{day.zfill(2)}"

        match_period = re.match(r"(\d{4})年(\d{1,2})-(\d{1,2})月$", date_str)
        if match_period:
            year, _start_month, end_month = match_period.groups()
            last_day = pd.to_datetime(f"{year}-{end_month}-01") + pd.tseries.offsets.MonthEnd(0)
            return last_day.strftime("%Y-%m-%d")

        match_year_only = re.match(r"(\d{4})年$", date_str)
        if match_year_only:
            year = match_year_only.group(1)
            return f"{year}-12-31"

    return date_str


def build_income_statement_period_label(
    effective_date: Any,
    *,
    months: int | None = None,
    fiscal_year_end_month: int | None = None,
    fiscal_year_end_day: int | None = None,
    language: str = "Eng",
) -> str:
    normalized = normalize_financial_date_label(effective_date)
    parsed = pd.to_datetime(normalized, errors="coerce")
    if pd.isna(parsed):
        return str(effective_date or "").strip()

    resolved_months: int | None = None
    fiscal_anchor = None
    if isinstance(fiscal_year_end_month, (int, float)):
        anchor_month = int(fiscal_year_end_month)
        anchor_day = int(fiscal_year_end_day) if isinstance(fiscal_year_end_day, (int, float)) else int(parsed.day)
        if 1 <= anchor_month <= 12 and 1 <= anchor_day <= 31:
            fiscal_anchor = (anchor_month, anchor_day)

    if isinstance(months, (int, float)):
        candidate = int(months)
        if 0 < candidate < 12:
            resolved_months = candidate
    if resolved_months is None:
        if fiscal_anchor and (int(parsed.month), int(parsed.day)) != fiscal_anchor:
            derived_months = (int(parsed.month) - fiscal_anchor[0]) % 12
            if derived_months <= 0:
                derived_months = int(parsed.month)
            if 0 < derived_months < 12:
                resolved_months = derived_months
        elif not fiscal_anchor and int(parsed.month) != 12:
            resolved_months = int(parsed.month)

    year = int(parsed.year)
    label_year = year
    if fiscal_anchor:
        label_year = year if (int(parsed.month), int(parsed.day)) <= fiscal_anchor else year + 1

    if language == "Chi":
        if resolved_months:
            return f"{label_year}年1-{resolved_months}月"
        return f"{label_year}年度"

    year_short = str(label_year)[-2:]
    if resolved_months:
        return f"{resolved_months}M{year_short}"
    return f"FY{year_short}"


def normalize_chinese_punctuation_in_text(s: str, preserve_sentence_stop: bool = False) -> str:
    if not s or not isinstance(s, str):
        return s
    s = s.replace("\uff0c", ",").replace("\u3001", ",")
    if not preserve_sentence_stop:
        s = s.replace("\u3002", ".").replace("\uff0e", ".")
    return s


def normalize_number_str_for_parsing(s: str) -> str:
    if not s or not isinstance(s, str):
        return s
    s = str(s).strip()
    s = re.sub(r"(?<=\d)[，,\uff0c\u3001\u060c\ufe50\ufe10](?=\d)", "", s)
    s = s.replace("\u3002", ".").replace("\uff0e", ".")
    return s

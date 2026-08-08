from __future__ import annotations

"""
Resolve workbook tabs to FDD mapping keys using workbook metadata and fuzzy aliases.
"""

from .mapping import load_mappings, should_skip_account_label
from .inspector import load_workbook_frames, profile_workbook
from .statements import extract_balance_sheet_and_income_statement
from .schedules import normalize_financial_schedule


from difflib import SequenceMatcher
import json
import os
import re
from typing import Any, Callable, Dict, Iterable, List, Optional, Tuple

import pandas as pd


def _normalize_label(text: str) -> str:
    if not text:
        return ""
    normalized = str(text).strip().lower()
    normalized = normalized.replace("&", " and ")
    normalized = re.sub(r"[\W_]+", " ", normalized, flags=re.UNICODE)
    normalized = re.sub(r"\bexpenses\b", "expense", normalized)
    normalized = re.sub(r"\bpayables\b", "payable", normalized)
    normalized = re.sub(r"\breceivables\b", "receivable", normalized)
    normalized = re.sub(r"\bproperties\b", "property", normalized)
    normalized = re.sub(r"\s+", " ", normalized).strip()
    return normalized


def _token_set(text: str) -> set[str]:
    normalized = _normalize_label(text)
    return {token for token in normalized.split(" ") if token}


def _is_compact_cjk_label(text: str) -> bool:
    normalized = _normalize_label(text)
    if not normalized or " " in normalized:
        return False
    return bool(re.fullmatch(r"[\u4e00-\u9fff]+", normalized))


def _candidate_strings(profile: Dict[str, Any]) -> List[str]:
    values = [profile.get("sheet_name", ""), profile.get("title", "")]
    return [value for value in values if value]


def _is_exact_alias_match(sheet_name: str, title: str, alias: Optional[str]) -> bool:
    alias_norm = _normalize_label(alias or "")
    if not alias_norm:
        return False
    return any(
        _normalize_label(candidate) == alias_norm
        for candidate in (sheet_name, title)
        if candidate
    )


def _score_candidate(candidate: str, alias: str) -> float:
    candidate_norm = _normalize_label(candidate)
    alias_norm = _normalize_label(alias)
    if not candidate_norm or not alias_norm:
        return 0.0
    if candidate_norm == alias_norm:
        return 100.0
    candidate_tokens = _token_set(candidate)
    alias_tokens = _token_set(alias)
    overlap = len(candidate_tokens & alias_tokens)
    score = 0.0
    if overlap:
        score += overlap * 18.0
        if alias_tokens and alias_tokens.issubset(candidate_tokens):
            score += 24.0
        if candidate_tokens and candidate_tokens.issubset(alias_tokens):
            score += 12.0
    if _is_compact_cjk_label(candidate_norm) and _is_compact_cjk_label(alias_norm) and candidate_norm != alias_norm:
        if alias_norm in candidate_norm or candidate_norm in alias_norm:
            score += 40.0
        else:
            return min(score, 24.0)
    ratio = SequenceMatcher(None, candidate_norm, alias_norm).ratio()
    if ratio >= 0.70:
        score += ratio * 40.0
    score += _semantic_alignment_adjustment(candidate_norm, alias_norm)
    return min(score, 95.0)


def _semantic_alignment_adjustment(candidate_norm: str, alias_norm: str) -> float:
    adjustment = 0.0

    def contains(text: str, phrase: str) -> bool:
        return phrase in text

    candidate_non_operating = contains(candidate_norm, "non operating")
    alias_non_operating = contains(alias_norm, "non operating")
    if candidate_non_operating != alias_non_operating and (
        contains(candidate_norm, "operating") or contains(alias_norm, "operating")
    ):
        adjustment -= 42.0

    candidate_non_current = contains(candidate_norm, "non current")
    alias_non_current = contains(alias_norm, "non current")
    candidate_current = contains(candidate_norm, "current")
    alias_current = contains(alias_norm, "current")
    if candidate_non_current != alias_non_current:
        adjustment -= 34.0
    elif candidate_current != alias_current and not candidate_non_current and not alias_non_current:
        adjustment -= 16.0

    candidate_other = bool(re.search(r"\bother\b", candidate_norm))
    alias_other = bool(re.search(r"\bother\b", alias_norm))
    if candidate_other != alias_other:
        adjustment -= 18.0

    candidate_payable = bool(re.search(r"\bpayable\b", candidate_norm))
    alias_payable = bool(re.search(r"\bpayable\b", alias_norm))
    if candidate_payable != alias_payable:
        adjustment -= 16.0

    candidate_surcharge = bool(re.search(r"\bsurcharge\b", candidate_norm))
    alias_surcharge = bool(re.search(r"\bsurcharge\b", alias_norm))
    if candidate_surcharge != alias_surcharge:
        adjustment -= 16.0

    candidate_by_customer = contains(candidate_norm, "by customer") or contains(candidate_norm, "customer")
    alias_by_customer = contains(alias_norm, "by customer") or contains(alias_norm, "customer")
    if candidate_by_customer and not alias_by_customer:
        adjustment -= 24.0

    candidate_receivable = bool(re.search(r"\breceivable\b", candidate_norm))
    alias_receivable = bool(re.search(r"\breceivable\b", alias_norm))
    if candidate_receivable != alias_receivable:
        adjustment -= 16.0

    candidate_notes = bool(re.search(r"\bnotes?\b", candidate_norm))
    alias_notes = bool(re.search(r"\bnotes?\b", alias_norm))
    if candidate_receivable and alias_receivable and candidate_notes != alias_notes:
        adjustment -= 20.0
    elif candidate_notes != alias_notes:
        adjustment -= 12.0

    return adjustment


def _sheet_type_bonus(profile: Dict[str, Any], mapping_type: str) -> float:
    sheet_kind = profile.get("sheet_kind")
    if sheet_kind in ("financial_summary", "template_nav"):
        return -100.0
    if sheet_kind == "support_schedule":
        return -20.0
    if profile.get("has_indicative_stage"):
        return 8.0
    return 0.0


# Below this score, a match is almost always an incidental single-word title
# overlap (nav/cover/ADJ/JE tabs) or a short-alias token subset, not a real
# account tab — empirically 18.0 (single generic token, e.g. "assets") and
# 42.0 (short alias like "FA"/"CIP"/"Cr Loss" as a token subset) were both
# confirmed false positives across real client databooks, and neither was ever
# observed as a legitimate match's score (real matches land at 100.0 exact,
# or well above this via CJK compact-label/SequenceMatcher-ratio bonuses).
# Originally scoped to sheet_kind == "other" (no detectable stage row at all),
# but the same 18.0/42.0 false-positive pattern also showed up on a sheet that
# DID have a detectable stage row (sheet_kind == "financial_schedule") yet
# still failed later at financial-value-column detection, so the floor now
# applies to every candidate regardless of sheet_kind. A real account tab with
# an unrecognized stage-label vocabulary still tends to score well above this
# via name/title similarity, so raising the floor here doesn't hide genuine
# CANONICAL_STAGE_LABELS gaps.
_LOW_CONFIDENCE_MATCH_FLOOR = 45.0




def _candidate_sheets_for_mapping(mapping_key: str, config: Dict[str, Any], profiles: Dict[str, Dict[str, Any]]) -> List[Dict[str, Any]]:
    aliases = list(config.get("aliases") or [])
    aliases.append(mapping_key)
    candidates: List[Dict[str, Any]] = []
    for sheet_name, profile in profiles.items():
        base_score = _sheet_type_bonus(profile, config.get("type", ""))
        if base_score <= -100:
            continue
        best_score = 0.0
        matched_alias = None
        for alias in aliases:
            for candidate in _candidate_strings(profile):
                score = _score_candidate(candidate, alias)
                if score > best_score:
                    best_score = score
                    matched_alias = alias
        if best_score <= 0:
            continue
        if best_score < _LOW_CONFIDENCE_MATCH_FLOOR:
            continue
        total_score = best_score + base_score
        if total_score <= 0:
            continue
        candidates.append(
            {
                "sheet_name": sheet_name,
                "title": profile.get("title"),
                "score": round(total_score, 2),
                "matched_alias": matched_alias,
                "exact_alias_match": _is_exact_alias_match(sheet_name, profile.get("title"), matched_alias),
                "sheet_kind": profile.get("sheet_kind"),
                "entity_scope": profile.get("entity_scope", "single"),
                "mapping_key": mapping_key,
                "category": config.get("category"),
                "type": config.get("type"),
            }
        )
    return sorted(candidates, key=lambda item: item["score"], reverse=True)


def should_use_ai_for_candidates(
    candidates: List[Dict[str, Any]],
    score_gap_threshold: float = 3.0,
) -> bool:
    if len(candidates) < 2:
        return False
    top = candidates[0]
    second = candidates[1]
    gap = float(top.get("score", 0)) - float(second.get("score", 0))
    names_blob = " ".join(str(candidate.get("sheet_name", "")) for candidate in candidates[:3]).lower()
    has_version_signal = bool(re.search(r"\bv\d+\b|version|\bcopy\b", names_blob))
    same_alias = top.get("matched_alias") == second.get("matched_alias")
    return gap <= score_gap_threshold or (has_version_signal and same_alias)


def _extract_sheet_names_from_ai_response(content: str, candidates: List[Dict[str, Any]]) -> List[str]:
    if not content:
        return []
    candidate_names = [str(candidate["sheet_name"]) for candidate in candidates]
    stripped = content.strip()
    try:
        parsed = json.loads(stripped)
        if isinstance(parsed, dict):
            chosen = parsed.get("sheet_name") or parsed.get("selected_sheet")
            if chosen in candidate_names:
                return [chosen]
            chosen_list = parsed.get("sheet_names") or parsed.get("candidate_sheets")
            if isinstance(chosen_list, list):
                return [name for name in chosen_list if name in candidate_names]
        if isinstance(parsed, list):
            return [name for name in parsed if name in candidate_names]
    except Exception:
        pass

    matched_names: List[str] = []
    for candidate_name in candidate_names:
        if candidate_name in stripped and candidate_name not in matched_names:
            matched_names.append(candidate_name)
    return matched_names


def _pick_financial_summary_sheet(profiles: Dict[str, Dict[str, Any]]) -> Optional[str]:
    summary_candidates = [
        (sheet_name, profile)
        for sheet_name, profile in profiles.items()
        if profile.get("sheet_kind") == "financial_summary"
    ]
    if not summary_candidates:
        return None

    def score(item: Tuple[str, Dict[str, Any]]) -> tuple[int, str]:
        sheet_name, profile = item
        text = f"{sheet_name} {profile.get('title', '')}".lower()
        points = 0
        if "financial" in text:
            points += 3
        if "balance" in text or "income" in text or "profit" in text:
            points += 2
        if "bs" in text or "is" in text:
            points += 1
        return (-points, sheet_name.lower())

    return sorted(summary_candidates, key=score)[0][0]


def _statement_values_for_mapping(statement_df: Any, mapping_key: str, config: Dict[str, Any]) -> Dict[str, float]:
    if statement_df is None or getattr(statement_df, "empty", True):
        return {}
    aliases = [mapping_key, *(config.get("aliases") or [])]
    desc_col = statement_df.columns[0]
    best_row = None
    best_score = 0.0
    for _, row in statement_df.iterrows():
        description = str(row[desc_col]).strip()
        if not description:
            continue
        row_score = max(_score_candidate(description, alias) for alias in aliases if alias)
        if row_score > best_score:
            best_score = row_score
            best_row = row
    if best_row is None or best_score <= 0:
        return {}

    values: Dict[str, float] = {}
    for col in statement_df.columns[1:]:
        try:
            numeric_value = float(best_row[col])
        except Exception:
            continue
        if col:
            values[str(col)] = numeric_value
    return values


def _normalized_total_values(normalized: Dict[str, Any]) -> Dict[str, float]:
    columns = normalized.get("columns") or []
    row_entries = normalized.get("row_entries") or []
    total_entry = next((row for row in row_entries if row.get("row_type") == "total"), None)
    if total_entry is None:
        return {}
    return {
        column["date"]: float(total_entry["values"][column["key"]])
        for column in columns
        if total_entry["values"].get(column["key"]) is not None
    }




def _build_financial_reference_context(
    workbook_path: str,
    profiles: Dict[str, Dict[str, Any]],
) -> Dict[str, Any]:
    summary_sheet = _pick_financial_summary_sheet(profiles)
    if not summary_sheet:
        return {
            "summary_sheet": None,
            "financial_results": {},
            "reference_available": False,
            "reference_stage": "missing_financial_summary",
        }
    try:
        financial_results = extract_balance_sheet_and_income_statement(workbook_path, summary_sheet, debug=False)
    except Exception as exc:
        return {
            "summary_sheet": summary_sheet,
            "financial_results": {},
            "reference_available": False,
            "reference_stage": "financial_summary_error",
            "reference_error": str(exc),
        }
    return {
        "summary_sheet": summary_sheet,
        "financial_results": financial_results,
        "reference_available": True,
        "reference_stage": "financial_summary_loaded",
    }


def _is_summary_account_candidate(description: str) -> bool:
    text = str(description or "").strip()
    if not text:
        return False
    return not should_skip_account_label(text)


def _iter_financial_reference_rows(financial_context: Dict[str, Any]) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    financial_results = financial_context.get("financial_results") or {}
    for statement_type, dataframe in (
        ("BS", financial_results.get("balance_sheet")),
        ("IS", financial_results.get("income_statement")),
    ):
        if dataframe is None or getattr(dataframe, "empty", True):
            continue
        desc_col = dataframe.columns[0]
        for _, row in dataframe.iterrows():
            description = str(row.get(desc_col) or "").strip()
            if not _is_summary_account_candidate(description):
                continue
            values: Dict[str, float] = {}
            for col in dataframe.columns[1:]:
                try:
                    values[str(col)] = float(row[col])
                except Exception:
                    continue
            rows.append(
                {
                    "account_name": description,
                    "statement_type": statement_type,
                    "values": values,
                }
            )
    return rows


def _infer_accounting_category(account_name: str, statement_type: str) -> str:
    normalized = _normalize_label(account_name)
    if statement_type == "IS":
        if any(token in normalized for token in ("income", "revenue", "sale", "gain", "收益", "收入")):
            return "Revenue"
        return "Expenses"

    if any(token in normalized for token in ("capital", "reserve", "earnings", "equity", "股本", "资本", "留存", "未分配利润")):
        return "Equity"
    if any(token in normalized for token in ("loan", "borrowing", "payable", "liability", "tax", "借款", "应付", "负债", "負債")):
        if any(token in normalized for token in ("long term", "long-term", "non current", "non-current", "长期", "非流动", "非流動")):
            return "Non-current liabilities"
        return "Current liabilities"
    if any(token in normalized for token in ("property", "fixed asset", "intangible", "deferred", "investment", "长期", "非流动", "非流動", "固定资产", "固定資產", "无形资产", "無形資產", "投资")):
        return "Non-current assets"
    return "Current assets"


def _build_dynamic_mapping_config(account_name: str, statement_type: str) -> Dict[str, Any]:
    return {
        "type": statement_type,
        "category": _infer_accounting_category(account_name, statement_type),
        "aliases": [account_name],
        "dynamic_mapping": True,
        "accounting_nature": _infer_accounting_category(account_name, statement_type),
    }


def _build_sheet_candidate_for_account(
    account_name: str,
    config: Dict[str, Any],
    sheet_name: str,
    profiles: Dict[str, Dict[str, Any]],
) -> Dict[str, Any]:
    profile = profiles[sheet_name]
    return {
        "sheet_name": sheet_name,
        "title": profile.get("title"),
        "score": 108.0 if _is_exact_alias_match(sheet_name, profile.get("title"), account_name) else 95.0,
        "matched_alias": account_name,
        "exact_alias_match": _is_exact_alias_match(sheet_name, profile.get("title"), account_name),
        "sheet_kind": profile.get("sheet_kind"),
        "entity_scope": profile.get("entity_scope", "single"),
        "mapping_key": account_name,
        "category": config.get("category"),
        "type": config.get("type"),
        "dynamic_mapping": True,
        "accounting_nature": config.get("accounting_nature"),
    }


def _candidate_passes_dynamic_confirmation(candidate: Dict[str, Any], materiality_threshold: float = 0.005) -> bool:
    compared_dates = int(candidate.get("financial_dates_compared", 0))
    matched_dates = int(candidate.get("financial_match_dates", 0))
    avg_pct_diff = candidate.get("financial_avg_pct_diff")
    if compared_dates <= 0:
        return False
    if matched_dates > 0:
        return True
    return avg_pct_diff is not None and float(avg_pct_diff) <= materiality_threshold


def _summary_row_exact_matches_sheet(account_name: str, profile: Dict[str, Any]) -> bool:
    normalized_account = _normalize_label(account_name)
    if not normalized_account:
        return False
    return any(
        _normalize_label(candidate) == normalized_account
        for candidate in _candidate_strings(profile)
    )


def _resolve_dynamic_sheet_mapping(
    workbook_path: str,
    account_name: str,
    config: Dict[str, Any],
    sheet_name: str,
    profiles: Dict[str, Dict[str, Any]],
    financial_context: Dict[str, Any],
    workbook_frames: Dict[str, Any],
    normalized_totals_cache: Dict[Tuple[str, str], Dict[str, float]],
    resolution_method: str,
) -> Optional[Dict[str, Any]]:
    candidate = _build_sheet_candidate_for_account(
        account_name=account_name,
        config=config,
        sheet_name=sheet_name,
        profiles=profiles,
    )
    ranked = _rank_candidates_with_financial_signals(
        workbook_path=workbook_path,
        mapping_key=account_name,
        config=config,
        candidates=[candidate],
        profiles=profiles,
        financial_context=financial_context,
        workbook_frames=workbook_frames,
        normalized_totals_cache=normalized_totals_cache,
    )
    if not ranked:
        return None
    top = ranked[0]
    if not _candidate_passes_dynamic_confirmation(top):
        return None
    return {
        **top,
        "resolution_method": resolution_method,
        "dynamic_mapping": True,
        "accounting_nature": config.get("accounting_nature"),
    }


def _discover_dynamic_sheet_resolutions(
    workbook_path: str,
    profiles: Dict[str, Dict[str, Any]],
    financial_context: Dict[str, Any],
    workbook_frames: Dict[str, Any],
    normalized_totals_cache: Dict[Tuple[str, str], Dict[str, float]],
    used_sheets: set[str],
    mappings: Dict[str, Any],
) -> Tuple[Dict[str, Dict[str, Any]], Dict[str, Dict[str, Any]]]:
    resolved: Dict[str, Dict[str, Any]] = {}
    dynamic_mappings: Dict[str, Dict[str, Any]] = {}
    for row in _iter_financial_reference_rows(financial_context):
        account_name = row["account_name"]
        if account_name in mappings or account_name in resolved:
            continue
        config = _build_dynamic_mapping_config(account_name, row["statement_type"])
        for sheet_name, profile in profiles.items():
            if sheet_name in used_sheets:
                continue
            if profile.get("sheet_kind") != "financial_schedule":
                continue
            if not _summary_row_exact_matches_sheet(account_name, profile):
                continue
            discovered = _resolve_dynamic_sheet_mapping(
                workbook_path=workbook_path,
                account_name=account_name,
                config=config,
                sheet_name=sheet_name,
                profiles=profiles,
                financial_context=financial_context,
                workbook_frames=workbook_frames,
                normalized_totals_cache=normalized_totals_cache,
                resolution_method="dynamic_exact_name",
            )
            if discovered is None:
                continue
            resolved[account_name] = discovered
            dynamic_mappings[account_name] = config
            used_sheets.add(sheet_name)
            break
    return resolved, dynamic_mappings


def _resolve_manual_override_target(
    workbook_path: str,
    mapping_key: str,
    override_sheet: str,
    mappings: Dict[str, Any],
    profiles: Dict[str, Dict[str, Any]],
    financial_context: Dict[str, Any],
    workbook_frames: Dict[str, Any],
    normalized_totals_cache: Dict[Tuple[str, str], Dict[str, float]],
) -> Tuple[Optional[Dict[str, Any]], Optional[Dict[str, Any]]]:
    if mapping_key in mappings:
        profile = profiles[override_sheet]
        return (
            {
                "sheet_name": override_sheet,
                "title": profile.get("title"),
                "score": 999.0,
                "matched_alias": "manual_override",
                "sheet_kind": profile.get("sheet_kind"),
                "entity_scope": profile.get("entity_scope", "single"),
                "mapping_key": mapping_key,
                "category": mappings.get(mapping_key, {}).get("category"),
                "type": mappings.get(mapping_key, {}).get("type"),
                "resolution_method": "manual_override",
            },
            None,
        )

    summary_rows = _iter_financial_reference_rows(financial_context)
    matched_row = next(
        (row for row in summary_rows if _normalize_label(row["account_name"]) == _normalize_label(mapping_key)),
        None,
    )
    if matched_row is None:
        return None, None
    config = _build_dynamic_mapping_config(matched_row["account_name"], matched_row["statement_type"])
    resolved = _resolve_dynamic_sheet_mapping(
        workbook_path=workbook_path,
        account_name=matched_row["account_name"],
        config=config,
        sheet_name=override_sheet,
        profiles=profiles,
        financial_context=financial_context,
        workbook_frames=workbook_frames,
        normalized_totals_cache=normalized_totals_cache,
        resolution_method="manual_dynamic_override",
    )
    if resolved is None:
        return None, None
    return resolved, config


def _candidate_total_values(
    workbook_path: str,
    candidate: Dict[str, Any],
    config: Dict[str, Any],
    profiles: Dict[str, Dict[str, Any]],
    workbook_frames: Dict[str, Any],
    normalized_totals_cache: Dict[Tuple[str, str], Dict[str, float]],
) -> Dict[str, float]:
    cache_key = (str(candidate.get("sheet_name", "")), str(config.get("type", "")))
    if cache_key in normalized_totals_cache:
        return normalized_totals_cache[cache_key]
    try:
        normalized = normalize_financial_schedule(
            workbook_path=workbook_path,
            sheet_name=candidate["sheet_name"],
            profile=profiles.get(candidate["sheet_name"]),
            sheet_df=workbook_frames.get(candidate["sheet_name"]),
            statement_type=config.get("type"),
        )
    except Exception:
        normalized_totals_cache[cache_key] = {}
        return {}
    values = _normalized_total_values(normalized)
    normalized_totals_cache[cache_key] = values
    return values


def _rank_candidates_with_financial_signals(
    workbook_path: str,
    mapping_key: str,
    config: Dict[str, Any],
    candidates: List[Dict[str, Any]],
    profiles: Dict[str, Dict[str, Any]],
    financial_context: Optional[Dict[str, Any]] = None,
    workbook_frames: Optional[Dict[str, Any]] = None,
    normalized_totals_cache: Optional[Dict[Tuple[str, str], Dict[str, float]]] = None,
    materiality_threshold: float = 0.005,
    absolute_tolerance: float = 1.0,
) -> List[Dict[str, Any]]:
    from .reconcile import _should_compare_income_statement_as_absolute  # local: resolver imports the later reconcile; module-level would be a cycle
    if not candidates:
        return []
    financial_context = financial_context or {}
    workbook_frames = workbook_frames or load_workbook_frames(workbook_path)
    normalized_totals_cache = normalized_totals_cache if normalized_totals_cache is not None else {}
    financial_results = financial_context.get("financial_results") or {}
    statement_df = (
        financial_results.get("balance_sheet")
        if config.get("type") == "BS"
        else financial_results.get("income_statement")
    )
    summary_values = _statement_values_for_mapping(statement_df, mapping_key, config)
    ranked_candidates: List[Dict[str, Any]] = []
    compare_as_absolute = bool(
        str(config.get("type") or "").strip().upper() == "IS"
        and _should_compare_income_statement_as_absolute(mapping_key, config.get("category"))
    )
    for candidate in candidates:
        candidate_values = _candidate_total_values(
            workbook_path=workbook_path,
            candidate=candidate,
            config=config,
            profiles=profiles,
            workbook_frames=workbook_frames,
            normalized_totals_cache=normalized_totals_cache,
        )
        pct_diffs: List[float] = []
        matched_dates = 0
        for date_label, candidate_value in candidate_values.items():
            summary_value = summary_values.get(date_label)
            if summary_value is None:
                continue
            candidate_value_for_compare = abs(candidate_value) if compare_as_absolute else candidate_value
            summary_value_for_compare = abs(summary_value) if compare_as_absolute else summary_value
            difference = abs(candidate_value_for_compare - summary_value_for_compare)
            pct_diff = (
                0.0
                if abs(summary_value_for_compare) <= absolute_tolerance
                else difference / abs(summary_value_for_compare)
            )
            if difference <= absolute_tolerance or pct_diff <= materiality_threshold:
                matched_dates += 1
            pct_diffs.append(pct_diff)
        compared_dates = len(pct_diffs)
        avg_pct_diff = (sum(pct_diffs) / compared_dates) if compared_dates else None
        ranked_candidates.append(
            {
                **candidate,
                "exact_alias_match": bool(candidate.get("exact_alias_match")),
                "financial_match_dates": matched_dates,
                "financial_dates_compared": compared_dates,
                "financial_avg_pct_diff": (round(avg_pct_diff, 6) if avg_pct_diff is not None else None),
                "financial_values_available": bool(candidate_values),
                "summary_values_available": bool(summary_values),
            }
        )
    ranked_candidates.sort(
        key=lambda item: (
            1 if item.get("exact_alias_match") else 0,
            1 if item.get("financial_dates_compared", 0) > 0 else 0,
            int(item.get("financial_match_dates", 0)),
            -(float(item["financial_avg_pct_diff"]) if item.get("financial_avg_pct_diff") is not None else float("inf")),
            float(item.get("score", 0)),
        ),
        reverse=True,
    )
    return ranked_candidates


def _build_candidate_map(
    mappings: Dict[str, Any],
    profiles: Dict[str, Dict[str, Any]],
) -> Dict[str, List[Dict[str, Any]]]:
    candidate_map: Dict[str, List[Dict[str, Any]]] = {}
    for mapping_key, config in mappings.items():
        if mapping_key.startswith("_") or not isinstance(config, dict):
            continue
        candidate_map[mapping_key] = _candidate_sheets_for_mapping(mapping_key, config, profiles)
    return candidate_map


def _available_candidates(
    candidates: List[Dict[str, Any]],
    used_sheets: set[str],
) -> List[Dict[str, Any]]:
    return [
        candidate for candidate in candidates
        if candidate["sheet_name"] not in used_sheets
    ]


def _rank_mapping_candidates(
    workbook_path: str,
    mapping_key: str,
    config: Dict[str, Any],
    candidates: List[Dict[str, Any]],
    profiles: Dict[str, Dict[str, Any]],
    financial_context: Dict[str, Any],
    workbook_frames: Dict[str, Any],
    normalized_totals_cache: Dict[Tuple[str, str], Dict[str, float]],
) -> List[Dict[str, Any]]:
    ranked_candidates = _rank_candidates_with_financial_signals(
        workbook_path=workbook_path,
        mapping_key=mapping_key,
        config=config,
        candidates=candidates,
        profiles=profiles,
        financial_context=financial_context,
        workbook_frames=workbook_frames,
        normalized_totals_cache=normalized_totals_cache,
    )
    return ranked_candidates or candidates


def _should_accept_hybrid_top_candidate(candidates: List[Dict[str, Any]]) -> bool:
    if not candidates:
        return False
    if len(candidates) == 1:
        return True
    top = candidates[0]
    second = candidates[1]
    top_match_dates = int(top.get("financial_match_dates", 0))
    second_match_dates = int(second.get("financial_match_dates", 0))
    top_compared = int(top.get("financial_dates_compared", 0))
    second_compared = int(second.get("financial_dates_compared", 0))
    top_avg = top.get("financial_avg_pct_diff")
    second_avg = second.get("financial_avg_pct_diff")
    top_score = float(top.get("score", 0))
    second_score = float(second.get("score", 0))
    top_exact = bool(top.get("exact_alias_match"))
    second_exact = bool(second.get("exact_alias_match"))

    if top_exact and not second_exact:
        return True

    if top_compared and top_match_dates > second_match_dates:
        return True
    if top_compared and not second_compared and top_match_dates > 0:
        return True
    if top_match_dates > 0 and second_match_dates == 0 and top_compared >= second_compared:
        return True
    if top_avg is not None and second_avg is not None and top_match_dates == second_match_dates:
        if (second_avg - top_avg) >= 0.02:
            return True
        if (second_avg - top_avg) >= 0.005 and (top_score - second_score) >= 5:
            return True
        if abs(second_avg - top_avg) <= 0.001 and (top_score - second_score) >= 20:
            return True
    if top_compared == 0 and second_compared == 0:
        return not should_use_ai_for_candidates(candidates)
    return False


def _resolve_top_ranked_candidate(candidates: List[Dict[str, Any]]) -> Dict[str, Any]:
    if not candidates:
        raise ValueError("No candidates provided")
    top = candidates[0]
    if int(top.get("financial_match_dates", 0)) > 0:
        return {**top, "resolution_method": "financial_validated"}
    return {**top, "resolution_method": "deterministic"}


def _default_ai_decider(
    mapping_key: str,
    candidates: List[Dict[str, Any]],
    model_type: str = "deepseek",
    language: str = "Eng",
) -> List[str]:
    try:
        from ..ai import AIClient
    except Exception:
        return []

    try:
        helper = AIClient(
            model_type=model_type,
            agent_name="sheet_resolver",
            language=language,
        )
    except Exception:
        return []

    candidate_payload = [
        {
            "sheet_name": candidate.get("sheet_name"),
            "title": candidate.get("title"),
            "score": candidate.get("score"),
            "matched_alias": candidate.get("matched_alias"),
            "sheet_kind": candidate.get("sheet_kind"),
            "entity_scope": candidate.get("entity_scope", "single"),
            "financial_match_dates": candidate.get("financial_match_dates"),
            "financial_dates_compared": candidate.get("financial_dates_compared"),
            "financial_avg_pct_diff": candidate.get("financial_avg_pct_diff"),
        }
        for candidate in candidates[:4]
    ]
    system_prompt = (
        "You are resolving ambiguous Excel tab mappings for a financial databook. "
        "Pick the best candidate sheet or shortlist of candidate sheets for the requested mapping key. "
        "Prefer the tabs whose title and schedule semantics best match the mapping key. "
        "Return JSON only in one of these forms: "
        "{\"sheet_name\": \"exact candidate name\"} "
        "or {\"sheet_names\": [\"candidate one\", \"candidate two\"]}."
    )
    user_prompt = (
        f"Mapping key: {mapping_key}\n"
        f"Candidates:\n{json.dumps(candidate_payload, ensure_ascii=False, indent=2)}\n\n"
        "Choose the exact candidate sheet_name that best matches the mapping key, or return a shortlist if multiple tabs still look plausible."
    )
    try:
        response = helper.get_response(user_prompt=user_prompt, system_prompt=system_prompt, temperature=0.0, max_tokens=120)
    except Exception:
        return []
    return _extract_sheet_names_from_ai_response(response.get("content", ""), candidates)


def resolve_ambiguous_candidate(
    workbook_path: str,
    mapping_key: str,
    config: Dict[str, Any],
    candidates: List[Dict[str, Any]],
    profiles: Dict[str, Dict[str, Any]],
    financial_context: Optional[Dict[str, Any]] = None,
    workbook_frames: Optional[Dict[str, Any]] = None,
    normalized_totals_cache: Optional[Dict[Tuple[str, str], Dict[str, float]]] = None,
    ai_decider: Optional[Callable[[str, List[Dict[str, Any]]], Any]] = None,
) -> Dict[str, Any]:
    if not candidates:
        raise ValueError("No candidates provided")
    if ai_decider is None:
        ai_decider = _default_ai_decider

    chosen_sheets: List[str] = []
    try:
        ai_choice = ai_decider(mapping_key, candidates)
        if isinstance(ai_choice, str) and ai_choice:
            chosen_sheets = [ai_choice]
        elif isinstance(ai_choice, list):
            chosen_sheets = [sheet for sheet in ai_choice if isinstance(sheet, str)]
    except Exception:
        chosen_sheets = []

    if len(chosen_sheets) == 1:
        for candidate in candidates:
            if candidate.get("sheet_name") == chosen_sheets[0]:
                return {**candidate, "resolution_method": "ai_fallback"}
    if len(chosen_sheets) > 1:
        shortlisted_candidates = [
            candidate for candidate in candidates if candidate.get("sheet_name") in chosen_sheets
        ]
        reranked_shortlist = _rank_candidates_with_financial_signals(
            workbook_path=workbook_path,
            mapping_key=mapping_key,
            config=config,
            candidates=shortlisted_candidates,
            profiles=profiles,
            financial_context=financial_context,
            workbook_frames=workbook_frames,
            normalized_totals_cache=normalized_totals_cache,
        )
        if reranked_shortlist:
            best = reranked_shortlist[0]
            best["ai_candidate_sheets"] = chosen_sheets
            best["resolution_method"] = "ai_financial_validated" if int(best.get("financial_match_dates", 0)) > 0 else "ai_shortlist_fallback"
            return best
        return {
            **shortlisted_candidates[0],
            "resolution_method": "ai_shortlist_fallback",
            "ai_candidate_sheets": chosen_sheets,
        }
    fallback = candidates[0]
    if int(fallback.get("financial_match_dates", 0)) > 0:
        return {**fallback, "resolution_method": "financial_fallback"}
    return {**fallback, "resolution_method": "deterministic_fallback"}


def resolve_workbook_mappings(
    workbook_path: str,
    profiles: Optional[Dict[str, Dict[str, Any]]] = None,
    workbook_frames: Optional[Dict[str, pd.DataFrame]] = None,
    mappings_path: Optional[str] = None,
    use_ai_for_ambiguity: bool = False,  # AI disambiguation removed — sequential AI calls during upload were a major slowdown; deterministic ranking only
    ai_decider: Optional[Callable[[str, List[Dict[str, Any]]], Optional[str]]] = None,
    model_type: str = "deepseek",
    language: str = "Eng",
    mapping_overrides: Optional[Dict[str, str]] = None,
) -> Dict[str, Any]:
    mappings = load_mappings(mappings_path)
    if profiles is None:
        profiles = profile_workbook(workbook_path)
    mapping_overrides = mapping_overrides or {}
    workbook_frames = workbook_frames or load_workbook_frames(workbook_path)
    financial_context = _build_financial_reference_context(workbook_path, profiles)
    normalized_totals_cache: Dict[Tuple[str, str], Dict[str, float]] = {}

    resolved: Dict[str, Dict[str, Any]] = {}
    candidate_map = _build_candidate_map(mappings, profiles)
    used_sheets: set[str] = set()
    ambiguities: Dict[str, List[Dict[str, Any]]] = {}
    override_issues: List[Dict[str, Any]] = []
    dynamic_mappings: Dict[str, Dict[str, Any]] = {}

    ranked_keys = sorted(
        candidate_map,
        key=lambda key: candidate_map[key][0]["score"] if candidate_map[key] else -1,
        reverse=True,
    )
    for mapping_key in ranked_keys:
        override_sheet = mapping_overrides.get(mapping_key)
        if override_sheet:
            if override_sheet not in profiles:
                override_issues.append(
                    {
                        "mapping_key": mapping_key,
                        "sheet_name": override_sheet,
                        "issue_type": "invalid_override",
                        "details": "Requested override sheet is not present in the workbook profile.",
                    }
                )
            elif override_sheet in used_sheets:
                override_issues.append(
                    {
                        "mapping_key": mapping_key,
                        "sheet_name": override_sheet,
                        "issue_type": "override_conflict",
                        "details": "Requested override sheet is already assigned to another mapping key.",
                    }
                )
            else:
                resolved_override, dynamic_config = _resolve_manual_override_target(
                    workbook_path=workbook_path,
                    mapping_key=mapping_key,
                    override_sheet=override_sheet,
                    mappings=mappings,
                    profiles=profiles,
                    financial_context=financial_context,
                    workbook_frames=workbook_frames,
                    normalized_totals_cache=normalized_totals_cache,
                )
                if resolved_override is not None:
                    resolved[mapping_key] = resolved_override
                    if dynamic_config is not None:
                        dynamic_mappings[mapping_key] = dynamic_config
                    used_sheets.add(override_sheet)
                    continue

        available_candidates = _available_candidates(candidate_map[mapping_key], used_sheets)
        if not available_candidates:
            continue

        ranked_available_candidates = _rank_mapping_candidates(
            workbook_path=workbook_path,
            mapping_key=mapping_key,
            config=mappings.get(mapping_key, {}),
            candidates=available_candidates,
            profiles=profiles,
            financial_context=financial_context,
            workbook_frames=workbook_frames,
            normalized_totals_cache=normalized_totals_cache,
        )

        if _should_accept_hybrid_top_candidate(ranked_available_candidates):
            resolved_candidate = _resolve_top_ranked_candidate(ranked_available_candidates)
        else:
            ambiguities[mapping_key] = ranked_available_candidates[:3]
            if use_ai_for_ambiguity:
                resolved_candidate = resolve_ambiguous_candidate(
                    workbook_path=workbook_path,
                    mapping_key=mapping_key,
                    config=mappings.get(mapping_key, {}),
                    candidates=ranked_available_candidates,
                    profiles=profiles,
                    financial_context=financial_context,
                    workbook_frames=workbook_frames,
                    normalized_totals_cache=normalized_totals_cache,
                    ai_decider=ai_decider or (
                        lambda key, candidate_list: _default_ai_decider(
                            key,
                            candidate_list,
                            model_type=model_type,
                            language=language,
                        )
                    ),
                )
            else:
                resolved_candidate = _resolve_top_ranked_candidate(ranked_available_candidates)
        resolved[mapping_key] = resolved_candidate
        used_sheets.add(resolved_candidate["sheet_name"])

    dynamic_override_keys = [
        key for key in mapping_overrides.keys()
        if key not in candidate_map
    ]
    for mapping_key in dynamic_override_keys:
        override_sheet = mapping_overrides.get(mapping_key)
        if not override_sheet:
            continue
        if override_sheet not in profiles:
            override_issues.append(
                {
                    "mapping_key": mapping_key,
                    "sheet_name": override_sheet,
                    "issue_type": "invalid_override",
                    "details": "Requested override sheet is not present in the workbook profile.",
                }
            )
            continue
        if override_sheet in used_sheets:
            override_issues.append(
                {
                    "mapping_key": mapping_key,
                    "sheet_name": override_sheet,
                    "issue_type": "override_conflict",
                    "details": "Requested override sheet is already assigned to another mapping key.",
                }
            )
            continue
        resolved_override, dynamic_config = _resolve_manual_override_target(
            workbook_path=workbook_path,
            mapping_key=mapping_key,
            override_sheet=override_sheet,
            mappings=mappings,
            profiles=profiles,
            financial_context=financial_context,
            workbook_frames=workbook_frames,
            normalized_totals_cache=normalized_totals_cache,
        )
        if resolved_override is None:
            override_issues.append(
                {
                    "mapping_key": mapping_key,
                    "sheet_name": override_sheet,
                    "issue_type": "unknown_override_target",
                    "details": "Override target does not match an existing mapping key or an exact Financials account name.",
                }
            )
            continue
        resolved[mapping_key] = resolved_override
        if dynamic_config is not None:
            dynamic_mappings[mapping_key] = dynamic_config
        used_sheets.add(override_sheet)

    discovered_resolved, discovered_dynamic_mappings = _discover_dynamic_sheet_resolutions(
        workbook_path=workbook_path,
        profiles=profiles,
        financial_context=financial_context,
        workbook_frames=workbook_frames,
        normalized_totals_cache=normalized_totals_cache,
        used_sheets=used_sheets,
        mappings={**mappings, **dynamic_mappings},
    )
    resolved.update(discovered_resolved)
    dynamic_mappings.update(discovered_dynamic_mappings)

    unresolved_sheets = sorted(
        sheet_name
        for sheet_name, profile in profiles.items()
        if profile.get("sheet_kind") == "financial_schedule" and sheet_name not in used_sheets
    )

    return {
        "profiles": profiles,
        "resolved": resolved,
        "candidate_map": candidate_map,
        "unresolved_sheets": unresolved_sheets,
        "ambiguities": ambiguities,
        "override_issues": override_issues,
        "dynamic_mappings": dynamic_mappings,
        "financial_reference": {
            "summary_sheet": financial_context.get("summary_sheet"),
            "reference_available": bool(financial_context.get("reference_available")),
            "reference_stage": financial_context.get("reference_stage"),
            "reference_error": financial_context.get("reference_error"),
        },
    }


# --- end workbook/resolver.py ---

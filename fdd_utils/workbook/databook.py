from __future__ import annotations

# re-added: bound by an import in another section of the pre-split module
from typing import Any, Dict, Iterable, List, Optional


from .inspector import load_workbook_frames, profile_workbook
from .schedules import INTERNAL_ROW_KEY, _build_indent_signal_index, normalize_financial_schedule
from .resolver import resolve_workbook_mappings
import pandas as pd
import json
import warnings
import os
import re
import logging
import time
 
warnings.simplefilter(action='ignore', category=UserWarning)
logger = logging.getLogger(__name__)
from ..financial_common import clean_english_placeholders, load_yaml_file, normalize_financial_date_label, package_file_path
from ..financial_display_format import add_language_display_columns
 
def load_mapping(filename):
    return load_yaml_file(filename)
 
 
   
 
def _update_display_description_map(df: pd.DataFrame, mapping: dict[str, str]) -> None:
    existing = dict(df.attrs.get("display_description_map") or {})
    for source, target in mapping.items():
        source_text = str(source or "").strip()
        target_text = str(target or "").strip()
        if not source_text or not target_text:
            continue
        existing[source_text] = target_text
    df.attrs["display_description_map"] = existing


def filter_detail_accounts(df: pd.DataFrame) -> pd.DataFrame:
    """
    Filter out detail sub-account rows, keeping only main account totals.
    Removes rows with patterns like "应付利息_借款利息" or containing "  " (indentation).
    Also cleans English placeholders from descriptions.
    
    Args:
        df: DataFrame with account descriptions
        
    Returns:
        Filtered DataFrame with cleaned descriptions
    """
    if df is None or df.empty:
        return df
    
    df_filtered = df.copy()
    
    # Get the first column name (description column)
    desc_col = df_filtered.columns[0]

    row_types_by_description = df_filtered.attrs.get("row_types_by_description", {})
    if row_types_by_description:
        df_filtered = df_filtered[
            ~df_filtered[desc_col].astype(str).map(
                lambda value: row_types_by_description.get(str(value), "") == "breakdown"
            )
        ]
    
    # Filter patterns that indicate detail sub-accounts
    filter_patterns = [
        r'_',  # Sub-account separator
        r'^\s{2,}',  # Multiple spaces at start (indentation)
        r'其中[：:]',  # "Including:" markers in Chinese
    ]
    
    for pattern in filter_patterns:
        df_filtered = df_filtered[~df_filtered[desc_col].astype(str).str.contains(pattern, regex=True, na=False)]
    
    original_descriptions = df_filtered[desc_col].astype(str).tolist()
    cleaned_descriptions = [clean_english_placeholders(value) for value in original_descriptions]

    # Clean English placeholders from descriptions
    df_filtered[desc_col] = cleaned_descriptions
    
    # Remove rows that became empty after cleaning
    df_filtered = df_filtered[df_filtered[desc_col].astype(str).str.strip() != '']

    cleaned_mapping = {}
    for original, cleaned in zip(original_descriptions, cleaned_descriptions):
        original_text = str(original or "").strip()
        cleaned_text = str(cleaned or "").strip()
        if not original_text or not cleaned_text:
            continue
        cleaned_mapping[original_text] = cleaned_text
        cleaned_mapping[cleaned_text] = cleaned_text
    _update_display_description_map(df_filtered, cleaned_mapping)
    
    return df_filtered


def filter_zero_value_rows(df: pd.DataFrame, tolerance: float = 0.01) -> pd.DataFrame:
    """
    Remove rows where the numeric value column is zero/insignificant or missing.

    A row is only dropped when its projection-period value AND every other
    period (per any_period_nonzero_by_description, set by
    normalize_financial_schedule from the full multi-period row values) are
    all ~0 -- an account that went quiet this period but had real activity
    earlier is kept, not silently skipped just because the single column
    this function can see happens to be 0.

    Args:
        df: DataFrame with description in the first column and a value column
            somewhere after it -- NOT necessarily immediately after: the real
            caller (build_dataframe_variants_from_normalized_results) passes
            projection_df, whose second column is INTERNAL_ROW_KEY
            ("__source_row_idx", a bookkeeping row index), with the actual
            value column third. Blindly taking columns[1] silently checked
            that row-index column instead -- always a small positive int, so
            every row passed the "non-zero" test regardless of its real
            value, making this filter a no-op for genuinely-zero rows in
            production. Now explicitly skips INTERNAL_ROW_KEY when picking
            the value column.
        tolerance: Absolute value threshold treated as zero

    Returns:
        Filtered DataFrame containing only meaningful numeric rows
    """
    if df is None or df.empty or len(df.columns) < 2:
        return df

    filtered_df = df.copy()
    desc_col = filtered_df.columns[0]
    value_col = next((c for c in filtered_df.columns[1:] if c != INTERNAL_ROW_KEY), None)
    if value_col is None:
        return filtered_df
    numeric_values = pd.to_numeric(filtered_df[value_col], errors='coerce')
    row_types_by_description = filtered_df.attrs.get("row_types_by_description", {})
    any_period_nonzero_by_description = filtered_df.attrs.get("any_period_nonzero_by_description", {})
    preserved_zero_rows = filtered_df[desc_col].astype(str).map(
        lambda value: row_types_by_description.get(str(value), "") in {"subtotal", "total"}
    )
    had_activity_some_period = filtered_df[desc_col].astype(str).map(
        lambda value: bool(any_period_nonzero_by_description.get(str(value), False))
    )
    keep_mask = (
        preserved_zero_rows
        | had_activity_some_period
        | (numeric_values.notna() & (numeric_values.abs() >= tolerance))
    )
    return filtered_df.loc[keep_mask].reset_index(drop=True)


def _detect_report_language_from_profiles(profiles):
    english_count = 0
    chinese_count = 0
    for profile in profiles.values():
        sample = " ".join(
            str(value or "")
            for value in (
                profile.get("sheet_name"),
                profile.get("title"),
                *profile.get("stage_labels", []),
            )
        )
        if re.search(r'[\u4e00-\u9fff]', sample):
            chinese_count += 1
        else:
            english_count += 1
    if english_count + chinese_count == 0:
        return None
    return 'Eng' if english_count >= chinese_count else 'Chi'


def detect_databook_language(databook_path: str) -> Optional[str]:
    """Cheap, deterministic language detection ('Eng'/'Chi') from sheet
    profiles alone -- the same signal extract_normalized_data_from_excel
    uses internally (see resolver_language above), but skips mapping
    resolution and row normalization so it's fast enough to run as soon as
    a file is selected, before the user commits to a full Process click.
    Returns None if profiling fails or no sheet has any title/stage-label
    text to go on."""
    try:
        profiles = profile_workbook(databook_path)
    except Exception:
        return None
    return _detect_report_language_from_profiles(profiles)


def build_dataframes_from_normalized_results(
    normalized_results,
    workbook_list,
    report_language,
    filter_details=True,
    keep_zero_rows=False,
    variant: str = "default",
):
    """Build legacy dataframe outputs from normalized schedule payloads."""
    variants = build_dataframe_variants_from_normalized_results(
        normalized_results=normalized_results,
        workbook_list=workbook_list,
        report_language=report_language,
        variant_specs=[
            {
                "name": variant,
                "variant": variant,
                "filter_details": filter_details,
                "keep_zero_rows": keep_zero_rows,
            }
        ],
    )
    result = variants.get(variant, {})
    return result.get("dfs", {}), result.get("workbook_list", [])


def build_dataframe_variants_from_normalized_results(
    normalized_results,
    workbook_list,
    report_language,
    variant_specs,
):
    """Build one or more dataframe variants while traversing normalized sheets once."""
    if not variant_specs:
        return {}

    prepared_specs = []
    for spec in variant_specs:
        spec_name = str(spec.get("name") or spec.get("variant") or "").strip()
        if not spec_name:
            continue
        prepared_specs.append(
            {
                "name": spec_name,
                "variant": str(spec.get("variant") or "default"),
                "filter_details": bool(spec.get("filter_details", True)),
                "keep_zero_rows": bool(spec.get("keep_zero_rows", False)),
                "dfs": {},
                "workbook_list": [],
            }
        )

    if not prepared_specs:
        return {}

    variant_key_map = {
        "default": "projection_df",
        "original": "projection_df_original",
        "annualized": "projection_df_annualized",
        "analysis": "prompt_analysis_df",
    }

    for sheet in workbook_list:
        normalized = normalized_results.get(sheet)
        if not normalized:
            continue
        display_key = str(normalized.get("display_key") or sheet).strip() or str(sheet)
        for spec in prepared_specs:
            projection_key = variant_key_map.get(spec["variant"], "projection_df")
            source_df = normalized.get(projection_key)
            if source_df is None:
                source_df = normalized.get("projection_df")
            if source_df is None:
                continue

            extracted_df = source_df.copy()
            if extracted_df is None or extracted_df.empty:
                continue

            if not spec["keep_zero_rows"]:
                extracted_df = filter_zero_value_rows(extracted_df)

            if spec["filter_details"]:
                extracted_df = filter_detail_accounts(extracted_df)

            if extracted_df is None or extracted_df.empty:
                continue

            derived_attrs = dict(extracted_df.attrs)
            extracted_df.attrs.update(source_df.attrs)
            extracted_df.attrs.update(derived_attrs)
            extracted_df.attrs["report_language"] = report_language
            extracted_df.attrs["source_sheet_name"] = str(sheet)
            extracted_df.attrs["display_key"] = display_key

            if report_language and len(extracted_df.columns) > 1:
                extracted_df = add_language_display_columns(extracted_df, report_language)
                post_format_attrs = dict(extracted_df.attrs)
                extracted_df.attrs.update(source_df.attrs)
                extracted_df.attrs.update(post_format_attrs)
                extracted_df.attrs["report_language"] = report_language
                extracted_df.attrs["selected_variant"] = source_df.attrs.get(
                    "selected_variant",
                    spec["variant"],
                )
                extracted_df.attrs["source_sheet_name"] = str(sheet)
                extracted_df.attrs["display_key"] = display_key

            spec["dfs"][display_key] = extracted_df.reset_index(drop=True)
            spec["workbook_list"].append(display_key)

    return {
        spec["name"]: {
            "dfs": spec["dfs"],
            "workbook_list": spec["workbook_list"],
        }
        for spec in prepared_specs
    }


def extract_normalized_data_from_excel(databook_path, mode="All", entity_name=None, mapping_overrides=None):
    """
    Build integrity-aware normalized schedule payloads for the workbook.

    Returns:
        Tuple of (normalized_results, workbook_list, overall_result_type, report_language, resolution)
    """
    overall_started = time.perf_counter()
    profiles_started = time.perf_counter()
    profiles = profile_workbook(databook_path)
    workbook_frames = load_workbook_frames(databook_path)
    # Pre-warm the indent-signal index synchronously, in this (single) thread,
    # BEFORE the ThreadPoolExecutor below fans normalize_financial_schedule
    # out across worker threads -- normalize_financial_schedule's indent-based
    # reclassification reads this cache but must never be the one to trigger
    # its first (openpyxl-touching) computation from inside a worker thread.
    _build_indent_signal_index(databook_path)
    logger.debug("Profiled workbook %s in %.2fs", os.path.basename(databook_path), time.perf_counter() - profiles_started)
    resolver_language = _detect_report_language_from_profiles(profiles) or "Eng"

    resolution_started = time.perf_counter()
    resolution = resolve_workbook_mappings(
        databook_path,
        profiles=profiles,
        workbook_frames=workbook_frames,
        mapping_overrides=mapping_overrides,
        language=resolver_language,
    )
    logger.debug("Resolved workbook mappings for %s in %.2fs", os.path.basename(databook_path), time.perf_counter() - resolution_started)
    mappings = load_mapping(package_file_path('mappings.yml'))
    dynamic_mappings = resolution.get("dynamic_mappings") or {}
    resolution.setdefault("normalization_errors", {})

    normalized_results = {}
    workbook_list = []
    entity_scopes = []

    # Build work items first, deduped by sheet_name. Then normalize in parallel —
    # each call is CPU-bound DataFrame work on independent sheets, so threads
    # scale well and the I/O (workbook_frames) is already cached.
    seen_sheets: set[str] = set()
    work_items: list[tuple[str, Dict[str, Any], Dict[str, Any], Optional[str]]] = []
    for mapping_key, resolved in resolution.get("resolved", {}).items():
        mapping_config = mappings.get(mapping_key, {}) or dynamic_mappings.get(mapping_key, {}) or {}
        mapping_type = mapping_config.get("type") or resolved.get("type")
        if mode != "All" and mapping_type != mode:
            continue
        sheet_name = resolved["sheet_name"]
        if sheet_name in seen_sheets:
            continue
        seen_sheets.add(sheet_name)

        statement_type = mapping_type
        if mapping_type == "IS" and not (
            resolved.get("matched_alias") or str(resolved.get("resolution_method") or "").startswith("manual_")
        ):
            statement_type = None
        work_items.append((mapping_key, resolved, mapping_config, statement_type))

    def _normalize_one(item):
        mapping_key, resolved, mapping_config, statement_type = item
        sheet_name = resolved["sheet_name"]
        try:
            normalized = normalize_financial_schedule(
                workbook_path=databook_path,
                sheet_name=sheet_name,
                profile=profiles.get(sheet_name),
                entity_name=entity_name,
                sheet_df=workbook_frames.get(sheet_name),
                statement_type=statement_type,
            )
            return mapping_key, sheet_name, normalized, mapping_config, resolved, None
        except Exception as exc:
            return mapping_key, sheet_name, None, mapping_config, resolved, str(exc)

    if work_items:
        import multiprocessing
        from concurrent.futures import ThreadPoolExecutor, as_completed
        max_workers = max(1, min(multiprocessing.cpu_count(), len(work_items)))
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = [executor.submit(_normalize_one, item) for item in work_items]
            for future in as_completed(futures):
                mapping_key, sheet_name, normalized, mapping_config, resolved, err = future.result()
                if err:
                    resolution["normalization_errors"][sheet_name] = err
                    continue
                mapping_type = mapping_config.get("type") or resolved.get("type")
                normalized_results[sheet_name] = {
                    **normalized,
                    "mapping_key": mapping_key,
                    "category": mapping_config.get("category") or resolved.get("category"),
                    "type": mapping_type,
                    "display_key": (
                        mapping_key
                        if (
                            str(resolved.get("resolution_method") or "").startswith("manual_")
                            or bool(resolved.get("dynamic_mapping") or mapping_config.get("dynamic_mapping"))
                        )
                        else sheet_name
                    ),
                    "dynamic_mapping_context": {
                        "dynamic_mapping": bool(resolved.get("dynamic_mapping") or mapping_config.get("dynamic_mapping")),
                        "accounting_nature": (
                            mapping_config.get("accounting_nature")
                            or resolved.get("accounting_nature")
                            or mapping_config.get("category")
                            or resolved.get("category")
                            or ""
                        ),
                        "category": mapping_config.get("category") or resolved.get("category"),
                        "type": mapping_type,
                    },
                }
                workbook_list.append(sheet_name)
                entity_scopes.append((profiles.get(sheet_name) or {}).get("entity_scope", "single"))

    overall_result_type = 'multiple' if any(scope == 'multiple' for scope in entity_scopes) else 'single'
    report_language = _detect_report_language_from_profiles(profiles)
    logger.debug(
        "Normalized %s sheets from %s in %.2fs",
        len(workbook_list),
        os.path.basename(databook_path),
        time.perf_counter() - overall_started,
    )
    return normalized_results, workbook_list, overall_result_type, report_language, resolution


def extract_data_from_excel(databook_path, entity_name, mode="All", filter_details=True, keep_zero_rows=False, return_resolution=False, mapping_overrides=None):
    """
    Extract data from Excel file and determine language.
    
    Args:
        databook_path: Path to Excel file
        entity_name: Name of entity to extract
        mode: Filter mode ('All', 'Assets', 'Liabilities', 'Equity', 'Income', 'Expenses')
        filter_details: Whether to filter out detail sub-accounts (default: True)
        keep_zero_rows: Whether to preserve zero-value and header/detail rows (default: False)
    
    Returns:
        Tuple of (final_dfs, final_workbook_list, overall_result_type, report_language)
    """
    normalized_results, workbook_list, overall_result_type, report_language, resolution = extract_normalized_data_from_excel(
        databook_path=databook_path,
        mode=mode,
        entity_name=entity_name,
        mapping_overrides=mapping_overrides,
    )

    final_dfs, final_workbook_list = build_dataframes_from_normalized_results(
        normalized_results=normalized_results,
        workbook_list=workbook_list,
        report_language=report_language,
        filter_details=filter_details,
        keep_zero_rows=keep_zero_rows,
    )

    if not final_workbook_list:
        overall_result_type = 'None'

    if return_resolution:
        return final_dfs, final_workbook_list, overall_result_type, report_language, resolution
    return final_dfs, final_workbook_list, overall_result_type, report_language
# --- end workbook/databook.py ---

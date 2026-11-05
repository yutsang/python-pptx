from __future__ import annotations

from datetime import datetime
from typing import Any, Callable, Dict, Iterable, List, Optional

import pandas as pd

from ..financial_common import (
    build_income_statement_period_label,
    clean_english_placeholders,
    dedupe_non_empty,
    visible_descriptions,
)
from ..workbook import (
    INTERNAL_ROW_KEY,
    split_accounts_by_type as shared_split_accounts_by_type,
    split_bilingual_entity_name,
)


def build_entity_selector_model(
    entity_options: List[str],
    current_entity_name: str,
    preferred_language: Optional[str] = None,
) -> Dict[str, Any]:
    # A candidate that mixes CJK and English (e.g. "南通通海 Nantong Tonghai")
    # previously only ever offered as that one combined string -- add the
    # Chinese-only and English-only halves as their own selectable options
    # right next to it, so the user can pick whichever the report actually
    # needs without having to hand-edit the combined text.
    expanded_options: List[str] = []
    # If the report's language is already known (e.g. detected/selected as
    # Chinese) and nothing has been manually picked yet, default straight to
    # the language-matching half instead of the combined "中文 English"
    # string -- picking the wrong-language half by default would look like
    # the suggester doesn't know what report it's building.
    language_preferred_default: Optional[str] = None
    for option in entity_options:
        expanded_options.append(option)
        chinese_only, english_only = split_bilingual_entity_name(option)
        if chinese_only and chinese_only != option:
            expanded_options.append(chinese_only)
            if preferred_language == "Chn" and language_preferred_default is None:
                language_preferred_default = chinese_only
        if english_only and english_only != option:
            expanded_options.append(english_only)
            if preferred_language == "Eng" and language_preferred_default is None:
                language_preferred_default = english_only

    dropdown_options = dedupe_non_empty(expanded_options)
    text_value = str(current_entity_name or "").strip()
    if not text_value:
        if language_preferred_default:
            text_value = language_preferred_default
        elif len(dropdown_options) == 1:
            text_value = dropdown_options[0]

    return {
        "dropdown_options": dropdown_options,
        "show_dropdown": bool(dropdown_options),
        "manual_only": not bool(dropdown_options),
        "text_value": text_value,
    }


def should_render_preprocess_controls(processed: bool) -> bool:
    return not processed


def describe_statement_period(
    statement_type: str,
    label: str,
    annualized: bool = False,
    annualization_months: int | None = None,
    fiscal_year_end_month: int | None = None,
    fiscal_year_end_day: int | None = None,
) -> str:
    if statement_type == "BS":
        try:
            parsed = datetime.strptime(str(label), "%Y-%m-%d")
            return f"As at {parsed.strftime('%d %B %Y')} (BS)"
        except ValueError:
            return f"As at {label} (BS)"

    if statement_type == "IS":
        period_label = build_income_statement_period_label(
            label,
            months=annualization_months,
            fiscal_year_end_month=fiscal_year_end_month,
            fiscal_year_end_day=fiscal_year_end_day,
            language="Eng",
        )
        prefix = "Annualised" if annualized else "During"
        return f"{prefix} {period_label} (IS)"

    return str(label)

def _normalize_display_label(value: Any, display_description_map: Dict[str, str] | None = None) -> str:
    raw_text = str(value or "").strip()
    if not raw_text:
        return ""
    mapped = str((display_description_map or {}).get(raw_text) or raw_text).strip()
    return clean_english_placeholders(mapped).strip()


def _build_rhs_display_dataframe(
    source_df: pd.DataFrame,
    display_df: pd.DataFrame,
) -> pd.DataFrame | None:
    adjacent_detail_rows = source_df.attrs.get("adjacent_detail_rows") or []
    if not adjacent_detail_rows:
        return None

    rhs_df = pd.DataFrame(adjacent_detail_rows)
    if rhs_df.empty or "Description" not in rhs_df.columns:
        return None

    display_key = str(display_df.columns[0])
    has_row_key = INTERNAL_ROW_KEY in rhs_df.columns and INTERNAL_ROW_KEY in display_df.columns
    display_description_map = dict(source_df.attrs.get("display_description_map") or {})
    rhs_df["__display_key"] = rhs_df["Description"].apply(
        lambda value: _normalize_display_label(value, display_description_map)
    )
    visible_display_labels = {
        _normalize_display_label(value)
        for value in display_df.iloc[:, 0].tolist()
        if _normalize_display_label(value)
    }
    if visible_display_labels:
        rhs_df = rhs_df[
            # str(value), not value.strip(): astype(str) leaves missing values
            # as np.nan under pandas' new string dtype, so a blank cell reaches
            # the lambda as a float and .strip() raises. See row_text() in
            # financial_common for the same incompatibility in its joining form.
            rhs_df["__display_key"].astype(str).map(lambda value: str(value).strip() in visible_display_labels)
        ].copy()
        if rhs_df.empty:
            rhs_df = pd.DataFrame(adjacent_detail_rows)
            rhs_df["__display_key"] = rhs_df["Description"].apply(
                lambda value: _normalize_display_label(value, display_description_map)
            )

    columns_to_drop = {"Description", "__display_key", *[str(column) for column in display_df.columns]}
    projection_original_label = str(source_df.attrs.get("projection_original_column_label") or "").strip()
    projection_annualized_label = str(source_df.attrs.get("projection_annualized_column_label") or "").strip()
    if projection_original_label:
        columns_to_drop.add(projection_original_label)
    if projection_annualized_label:
        columns_to_drop.add(projection_annualized_label)

    rhs_display = pd.DataFrame({display_key: rhs_df["__display_key"].astype(str).str.strip()})
    if has_row_key:
        rhs_display[INTERNAL_ROW_KEY] = rhs_df[INTERNAL_ROW_KEY]
    remark_columns: List[str] = []
    for column in rhs_df.columns:
        column_name = str(column)
        if (
            column_name in columns_to_drop
            or column_name.endswith("| table_header")
            or column_name.endswith("| indicative_adjusted_row")
            or column_name.endswith("| date_row")
        ):
            continue

        text_values = [
            str(value).strip()
            for value in rhs_df[column_name].tolist()
            if str(value or "").strip()
        ]
        if not text_values:
            continue

        rhs_column_name = f"RHS: {column_name}"
        rhs_display[rhs_column_name] = rhs_df[column_name].fillna("").astype(str).str.strip()
        remark_columns.append(rhs_column_name)

    if not remark_columns:
        return None

    def combine_row_remarks(row: pd.Series) -> str:
        parts: List[str] = []
        for column_name in remark_columns:
            value = str(row.get(column_name) or "").strip()
            if not value:
                continue
            header = column_name.replace("RHS: ", "", 1).strip()
            parts.append(f"{header}: {value}" if header else value)
        return " | ".join(parts)

    rhs_display.insert(1, "Prompt remarks (RHS 1-5 cols)", rhs_display.apply(combine_row_remarks, axis=1))
    selected_columns = [display_key, "Prompt remarks (RHS 1-5 cols)"]
    if has_row_key:
        selected_columns.insert(0, INTERNAL_ROW_KEY)
    rhs_display = rhs_display[selected_columns]
    rhs_display = rhs_display[
        rhs_display[display_key].astype(str).str.strip() != ""
    ]
    if rhs_display.empty:
        return None
    if has_row_key:
        rhs_display = rhs_display.drop_duplicates(subset=[INTERNAL_ROW_KEY], keep="first")
    else:
        rhs_display = rhs_display.groupby(display_key, as_index=False).agg(
            {"Prompt remarks (RHS 1-5 cols)": lambda series: " || ".join(dedupe_non_empty(series))}
        )

    if len(rhs_display.columns) <= 1:
        return None
    return rhs_display


def build_account_display_dataframe(df: pd.DataFrame | None) -> pd.DataFrame | None:
    """Show full indicative-adjusted periods when prompt-analysis data is available."""
    if df is None or df.empty:
        return df

    analysis_df = df.attrs.get("prompt_analysis_df")
    if isinstance(analysis_df, pd.DataFrame) and not analysis_df.empty:
        display_df = analysis_df.copy()
        visible_rows = visible_descriptions(df)
        if visible_rows and len(display_df.columns) > 0:
            first_col = display_df.columns[0]
            filtered_df = display_df[
                display_df[first_col].astype(str).map(lambda value: str(value).strip() in visible_rows)
            ].copy()
            if not filtered_df.empty:
                display_df = filtered_df
    else:
        display_df = df.copy()

    rhs_display_df = _build_rhs_display_dataframe(df, display_df)
    if rhs_display_df is not None:
        display_key = str(display_df.columns[0])
        display_df[display_key] = display_df[display_key].astype(str).str.strip()
        if INTERNAL_ROW_KEY in display_df.columns and INTERNAL_ROW_KEY in rhs_display_df.columns:
            rhs_merge_df = rhs_display_df[[INTERNAL_ROW_KEY, "Prompt remarks (RHS 1-5 cols)"]].copy()
            display_df = display_df.merge(rhs_merge_df, on=[INTERNAL_ROW_KEY], how="left")
        else:
            display_df = display_df.merge(rhs_display_df, on=[display_key], how="left")
        ordered_columns = [display_key]
        ordered_columns.extend(
            str(column) for column in display_df.columns
            if (
                str(column) not in ordered_columns
                and str(column) != INTERNAL_ROW_KEY
                and str(column) != "Prompt remarks (RHS 1-5 cols)"
            )
        )
        if "Prompt remarks (RHS 1-5 cols)" in display_df.columns:
            ordered_columns.append("Prompt remarks (RHS 1-5 cols)")
        display_df = display_df[ordered_columns]
    elif INTERNAL_ROW_KEY in display_df.columns:
        display_df = display_df[[str(column) for column in display_df.columns if str(column) != INTERNAL_ROW_KEY]]

    if INTERNAL_ROW_KEY in display_df.columns:
        display_df = display_df[[str(column) for column in display_df.columns if str(column) != INTERNAL_ROW_KEY]]

    display_df.attrs.update(df.attrs)
    return display_df


def build_processed_display_groups(display_account_keys: List[str], mappings: Dict[str, Any], dfs: Dict[str, Any] | None = None) -> Dict[str, Any]:
    bs_accounts, is_accounts, _other_accounts = shared_split_accounts_by_type(display_account_keys, mappings, dfs=dfs)
    return {
        "tab_names": ["BS", "IS"],
        "bs_accounts": bs_accounts,
        "is_accounts": is_accounts,
    }


def detect_statement_mode(reconciliation: tuple[pd.DataFrame | None, pd.DataFrame | None] | None) -> str:
    """Auto-detect whether this databook has only BS, only IS, or both
    reconciled statements — replaces the old manual "Statement scope"
    selectbox. A statement is "absent" when its reconciliation frame is
    None/empty (e.g. the Financials sheet had no IS section at all), not
    merely when few of its accounts matched cleanly — that weaker case is
    already handled by the existing "use all dfs" fallback in the caller."""
    if not reconciliation:
        return "full"
    bs_recon, is_recon = (list(reconciliation) + [None, None])[:2]
    bs_present = bs_recon is not None and not bs_recon.empty
    is_present = is_recon is not None and not is_recon.empty
    if bs_present and not is_present:
        return "bs_only"
    if is_present and not bs_present:
        return "is_only"
    return "full"


def explain_reconciliation_selection(
    reconciliation: tuple[pd.DataFrame | None, pd.DataFrame | None] | None,
    available_keys: Iterable[str],
    resolution: Dict[str, Any] | None = None,
    dfs: Dict[str, Any] | None = None,
):
    """Account-by-account route and reason; see workbook/selection.py."""
    from ..workbook.selection import select_reconciliation_keys
    return select_reconciliation_keys(reconciliation, available_keys, resolution, dfs=dfs)


def derive_reconciliation_matched_keys(
    reconciliation: tuple[pd.DataFrame | None, pd.DataFrame | None] | None,
    available_keys: Iterable[str],
    resolution: Dict[str, Any] | None = None,
    dfs: Dict[str, Any] | None = None,
) -> List[str]:
    """The accounts that go to commentary, in order.

    Admitted: lines reconciliation tied out (BS ✅ / ⚠️ Match / ✅ Immaterial; IS
    additionally ❌ Diff, because IS reconciliation is period-flow noisy and
    excluding it would drop most of the statement), a partial-period line whose
    previous full period ties out, and a Financials line that was nil in both
    recent periods -- so never compared -- whose tab has earlier activity.

    NOT admitted by history: a balance-sheet ❌ Diff, a tab no Financials line
    reconciles to, a tab with no activity. The earlier version of this function
    re-admitted all three whenever the tab had any non-zero period, which undid
    its own BS ❌ Diff exclusion. workbook/selection.py has the reasons per
    account (``explain_reconciliation_selection``).
    """
    return explain_reconciliation_selection(reconciliation, available_keys, resolution, dfs=dfs).admitted_keys
# --- end ui/views.py ---

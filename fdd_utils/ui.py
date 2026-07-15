from __future__ import annotations

# --- begin ui/state.py ---
from typing import Any


DEFAULT_SESSION_STATE = {
    "uploaded_file": None,
    "dfs": None,
    "display_dfs": None,
    "dfs_variants": {},
    "display_df_variants": {},
    "workbook_list": [],
    "display_workbook_list": [],
    "language": "Eng",
    # True once the project team manually picks a language in the sidebar, so the
    # auto-detected language stops overwriting their choice (until a new upload).
    "language_user_set": False,
    # What auto-detection actually found, kept even after a manual override so
    # the UI can still show a "Detected: ..." reminder.
    "detected_language": None,
    # Cheap pre-Process detection (sheet profiles only, no mapping resolution)
    # so the "Detected: ..." reminder is visible BEFORE the user commits to a
    # full Process click, not just after -- see render_language_selector.
    "detected_language_preview": None,
    "_lang_preview_path": None,
    "bs_is_results": None,
    "ai_results": None,
    "reconciliation": None,
    "resolution": None,
    "model_type": "local",
    "model_name": None,          # specific model within the provider, e.g. GPT-5.5's id
    "model_choice_key": None,    # sidebar dropdown selection key
    "use_multithreading": True,  # default; render_sidebar_upload resolves the real value from config.yml
    "project_name": None,
    "last_run_folder": None,
    "entity_name": None,
    "pptx_download_trigger": None,
    "button_click_counter": 0,
    "pptx_ready": False,
    "temp_path": None,
    "selected_sheet": None,
    # Optional sibling roll-up ("主表") workbook + its entity-specific sheet
    # name, used to source the Financials summary when this entity's own
    # databook has no Financials-pattern sheet of its own.
    "rollup_temp_path": None,
    "rollup_sheet": None,
    "prev_entity_dropdown": "",
    "mapping_overrides": {},
    "account_comments": {},
    "upload_cache_cleanup_removed": 0,
}

RESET_SESSION_KEYS = [
    "dfs",
    "display_dfs",
    "dfs_variants",
    "display_df_variants",
    "workbook_list",
    "display_workbook_list",
    "language",
    # Reset the manual-override flag on a new upload so the new databook is
    # auto-detected afresh; the team can re-override it for that file.
    "language_user_set",
    "detected_language",
    "detected_language_preview",
    "_lang_preview_path",
    "bs_is_results",
    "ai_results",
    "reconciliation",
    "resolution",
    "entity_name",
    "project_name",
    "pptx_ready",
    "mapping_overrides",
    "account_comments",
    "rollup_temp_path",
    "rollup_sheet",
]

DELETE_SESSION_KEYS = [
    "pptx_download_data",
    "pptx_download_filename",
    "pptx_download_mime",
    "prev_entity_dropdown",
    "selected_sheet",
    "entity_dropdown",
    "entity_text_input",
    "sheet_select",
]


def init_session_state(session_state: Any) -> None:
    for key, value in DEFAULT_SESSION_STATE.items():
        if key not in session_state:
            session_state[key] = value


initialize_app_state = init_session_state


def reset_processing_session_state(session_state: Any, clear_upload_reference: bool = False) -> None:
    for key in RESET_SESSION_KEYS:
        session_state[key] = DEFAULT_SESSION_STATE[key]

    delete_keys = list(DELETE_SESSION_KEYS)
    if clear_upload_reference:
        delete_keys.append("prev_uploaded_temp_path")
    for key in delete_keys:
        if key in session_state:
            del session_state[key]

    for key in list(session_state.keys()):
        if str(key).startswith("statement_variant_"):
            del session_state[key]
# --- end ui/state.py ---

# --- begin ui/views.py ---
from datetime import datetime
from typing import Any, Callable, Dict, Iterable, List, Optional

import pandas as pd

from .financial_common import (
    build_income_statement_period_label,
    clean_english_placeholders,
    dedupe_non_empty,
    visible_descriptions,
)
from .workbook import (
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
            rhs_df["__display_key"].astype(str).map(lambda value: value.strip() in visible_display_labels)
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
                display_df[first_col].astype(str).map(lambda value: value.strip() in visible_rows)
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


def derive_reconciliation_matched_keys(
    reconciliation: tuple[pd.DataFrame | None, pd.DataFrame | None] | None,
    available_keys: Iterable[str],
    resolution: Dict[str, Any] | None = None,
) -> List[str]:
    available_key_order = dedupe_non_empty(available_keys)
    available_key_set = set(available_key_order)
    if not reconciliation or not available_key_order:
        return []

    # AI processes matched (✅ / ⚠️) and immaterial accounts.
    # BS: ❌ Diff intentionally excluded — those need human reconciliation
    # before AI commentary can be trusted.
    # IS: ❌ Diff included — IS recon is inherently noisier (period-flow
    # semantics, YoY movements). Excluding ❌ Diff for IS would drop most
    # of the income statement from the report, which is undesirable.
    bs_included_statuses = {"✅ Match", "⚠️ Match", "✅ Immaterial"}
    is_included_statuses = {"✅ Match", "⚠️ Match", "✅ Immaterial", "❌ Diff"}
    matched_keys: List[str] = []
    seen = set()

    # reconciliation is positionally (bs_recon_df, is_recon_df) — see
    # workbook.reconcile_financial_statements return order.
    for recon_idx, recon_df in enumerate(reconciliation):
        if recon_df is None or recon_df.empty:
            continue
        if "Match" not in recon_df.columns or "Tab_Account" not in recon_df.columns:
            continue

        included = is_included_statuses if recon_idx == 1 else bs_included_statuses
        filtered = recon_df[recon_df["Match"].isin(included)]
        for value in filtered["Tab_Account"].tolist():
            key = str(value or "").strip()
            if not key or key == "-" or key in seen or key not in available_key_set:
                continue
            seen.add(key)
            matched_keys.append(key)

    # Only accounts that passed reconciliation with an included status go to AI.
    # Dynamic-mapping and resolved-map fallbacks are intentionally removed —
    # they added accounts that were never validated by reconciliation.
    return matched_keys
# --- end ui/views.py ---

# --- begin ui/ai_panel.py ---
import time
import traceback
from typing import Any, Dict

import pandas as pd
import streamlit as st

# Compatibility shim: st.fragment landed in Streamlit 1.33. On older builds
# we fall back to a no-op decorator so the page still renders.
if not hasattr(st, "fragment"):
    def _noop_fragment(func=None, **_kwargs):
        if func is None:
            return lambda f: f
        return func
    st.fragment = _noop_fragment  # type: ignore[attr-defined]

from .ai import (
    FDDConfig,
    WORKBENCH_AVAILABLE_MODELS,
    build_highlighted_commentary_html,
    get_default_config_path,
    get_prompt_engine,
    is_provider_ready,
    load_yaml_config,
    parse_validator_response,
    run_ai_pipeline_with_progress,
    run_generator_reprompt,
)
from .financial_common import extract_result_text_content, get_pipeline_result_text
from .financial_display_format import prepare_display_dataframe, stringify_display_dataframe
from .workbook import (
    build_account_mapping_diagnostics,
    build_workbook_preflight,
    extract_entity_names_from_preflight,
    find_mapping_key,
    get_effective_mappings,
    get_financial_sheet_options,
    load_mappings,
    split_accounts_by_type,
)

_EMPTY_RESULT_MARKERS = {"none", "null", "nan", "", "n/a", "na"}
_PROMPT_MANAGER = get_prompt_engine()


def format_dataframe_for_display(df: pd.DataFrame) -> pd.DataFrame:
    return stringify_display_dataframe(prepare_display_dataframe(df))


def get_account_dataframe(account_key: str, account_dfs: Dict[str, pd.DataFrame]) -> pd.DataFrame | None:
    return (account_dfs or {}).get(account_key)


def get_financial_account_options(bs_is_results: Dict[str, Any] | None) -> list[str]:
    options: list[str] = []
    seen: set[str] = set()
    for statement_key in ("balance_sheet", "income_statement"):
        statement_df = (bs_is_results or {}).get(statement_key)
        if statement_df is None or statement_df.empty or len(statement_df.columns) == 0:
            continue
        for value in statement_df.iloc[:, 0].tolist():
            text = str(value or "").strip()
            if not text or text in seen:
                continue
            seen.add(text)
            options.append(text)
    return options


def build_selected_pipeline_dfs(session_state: Any) -> Dict[str, pd.DataFrame]:
    account_dfs = session_state.get("dfs") or {}
    selected_dfs: Dict[str, pd.DataFrame] = {}
    for account_key in session_state.get("workbook_list", []):
        selected_df = get_account_dataframe(account_key, account_dfs)
        if selected_df is not None:
            selected_dfs[account_key] = selected_df
    return selected_dfs


def extract_result_text(result_dict: Dict[str, Any], agent_key: str) -> str:
    return extract_result_text_content((result_dict or {}).get(agent_key, ""))


def has_meaningful_result_text(content: Any) -> bool:
    content_str = str(content or "").strip()
    return bool(content_str) and content_str.lower() not in _EMPTY_RESULT_MARKERS


def hydrate_nested_agent_outputs(result_dict: Dict[str, Any]) -> None:
    if not isinstance(result_dict, dict):
        return
    for agent_name in ("subagent_1", "subagent_2", "subagent_3", "subagent_4"):
        content = extract_result_text(result_dict, agent_name)
        if has_meaningful_result_text(content):
            result_dict[agent_name] = content
            if agent_name == "subagent_4" and not has_meaningful_result_text(result_dict.get("final")):
                result_dict["final"] = get_pipeline_result_text(result_dict)


def effective_mappings_from_session(session_state: Any) -> Dict[str, Any]:
    return get_effective_mappings(load_mappings(), session_state.get("resolution"))


def _result_has_pipeline_content(result: Dict[str, Any]) -> bool:
    return any(
        has_meaningful_result_text(extract_result_text(result, agent_key))
        for agent_key in ["final", "subagent_4", "subagent_3", "subagent_2", "subagent_1"]
    )


def extract_validator_metadata(result_dict: Dict[str, Any]) -> Dict[str, object]:
    validation = result_dict.get("agent_4_validation", {})
    if isinstance(validation, dict):
        if isinstance(validation.get("clause_reviews"), list):
            return {
                "final_content": str(validation.get("final_content") or extract_result_text(result_dict, "final")).strip(),
                "clause_reviews": validation.get("clause_reviews", []),
                "raw_response": str(validation.get("raw_response") or validation.get("output") or ""),
            }
        raw_response = validation.get("raw_response") or validation.get("output")
        if isinstance(raw_response, str) and raw_response.strip():
            return parse_validator_response(raw_response, fallback_content=extract_result_text(result_dict, "final"))
    return {"clause_reviews": [], "raw_response": ""}


def extract_account_remarks_context(df: pd.DataFrame | None, language: str) -> Dict[str, Any]:
    if not isinstance(df, pd.DataFrame):
        return {}
    supporting_notes = [str(note).strip() for note in (df.attrs.get("supporting_notes") or []) if str(note).strip()]
    rhs_rows = _PROMPT_MANAGER.filter_adjacent_detail_rows(df)
    rhs_summary = _PROMPT_MANAGER.summarize_rhs_remarks(rhs_rows, language)
    table_linked_remarks = _PROMPT_MANAGER.table_linked_remarks(df)
    return {
        "supporting_notes": supporting_notes,
        "rhs_rows": rhs_rows,
        "rhs_summary": rhs_summary,
        "table_linked_remarks": table_linked_remarks,
    }


def render_account_remarks_context(df: pd.DataFrame | None, key: str, language: str, prefix: str = "") -> None:  # noqa: ARG001
    _ = prefix  # kept for API compatibility
    context = extract_account_remarks_context(df, language)
    if not context:
        return
    supporting_notes = context.get("supporting_notes") or []
    rhs_rows = context.get("rhs_rows") or []
    rhs_summary = context.get("rhs_summary") or []
    table_linked_remarks = context.get("table_linked_remarks") or []
    if not any((supporting_notes, rhs_rows, rhs_summary, table_linked_remarks)):
        return

    summary_bits = []
    if supporting_notes:
        summary_bits.append(f"{len(supporting_notes)} supporting note(s)")
    if rhs_rows:
        summary_bits.append(f"{len(rhs_rows)} RHS remark row(s)")
    if table_linked_remarks:
        summary_bits.append(f"{len(table_linked_remarks)} table-linked remark(s)")

    with st.expander(f"Source remarks / context for {key}", expanded=False):
        if summary_bits:
            st.caption(" | ".join(summary_bits))
        if rhs_summary:
            st.markdown("**RHS remark summary**")
            for item in rhs_summary:
                st.markdown(f"- {item}")
        if supporting_notes:
            st.markdown("**Supporting notes**")
            for note in supporting_notes:
                st.markdown(f"- {note}")
        if rhs_rows:
            st.markdown("**RHS remarks / reasons**")
            st.dataframe(pd.DataFrame(rhs_rows), use_container_width=True)
        if table_linked_remarks:
            st.markdown("**Table-linked remarks**")
            st.dataframe(pd.DataFrame(table_linked_remarks), use_container_width=True)


def _run_demo_ai(
    matched_keys: list,
    duration_s: int,
    progress_placeholder,
    status_placeholder,
) -> dict:
    """Simulate AI processing by replaying pre-baked demo_results.json."""
    import json as _json
    import math as _math
    demo_path = Path(__file__).parent / "demo_results.json"
    try:
        all_results: dict = _json.loads(demo_path.read_text(encoding="utf-8"))
    except Exception as exc:
        logger.warning("Could not load demo_results.json: %s", exc)
        all_results = {}

    # Pipeline is now 3 stages: Generator → Auditor (Polish) → Validator
    from .ai import SUBAGENT_SEQUENCE
    agents = [label for _, label in SUBAGENT_SEQUENCE]
    n_stages = len(agents)
    total_steps = n_stages * max(len(matched_keys), 1)
    tick = max(0.1, duration_s / total_steps)
    step = 0
    for agent_idx, agent_name in enumerate(agents):
        for key_idx, key in enumerate(matched_keys):
            step += 1
            progress_placeholder.progress(min(step / total_steps, 1.0))
            status_placeholder.info(
                f"🔄 Running Subagent {agent_idx + 1}/{n_stages}: {agent_name} "
                f"| Processing item {key_idx + 1}/{len(matched_keys)} "
                f"| Key: {key} "
                f"| Progress: {step}/{total_steps} steps"
            )
            time.sleep(tick)

    # matched_keys come from the reconciliation Tab_Account column — for the
    # Chinese demo databook those are raw Chinese tab names ('货币资金',
    # '营业收入', etc.). The demo cache is keyed by canonical English mapping
    # keys (Cash, OI, OC, etc.). Build an alias→canonical map from
    # mappings.yml so the lookup canonicalises before the cache hit; without
    # this every Chinese tab name fell into the generic fallback stub and the
    # IS bullets disappeared from the export.
    alias_to_canonical: Dict[str, str] = {}
    try:
        from fdd_utils.ai import get_prompt_engine as _get_pe
        for canonical, entry in (_get_pe().mappings_data or {}).items():
            if not isinstance(entry, dict):
                continue
            alias_to_canonical[canonical] = canonical
            for alias in entry.get("aliases", []) or []:
                alias_to_canonical[str(alias)] = canonical
    except Exception as exc:
        logger.warning("demo: could not build alias map (%s)", exc)

    # Hardcoded fallback aliases for the English demo databook whose sheet
    # names don't match canonical keys directly. This covers cases where the
    # prompt-engine alias map is unavailable or returns an empty dict, which
    # would otherwise send every non-exact-match key to the generic stub.
    _DEMO_HARD_ALIASES: Dict[str, str] = {
        # Balance sheet
        "Cash at bank": "Cash",
        "Cash at bank and on hand": "Cash",
        "Accounts receivable": "AR",
        "Other receivables": "OR",
        "Other current assets": "Other CA",
        "Other Current Assets": "Other CA",
        "Investment properties": "IP",
        "Investment Properties": "IP",
        "Share capital": "Capital",
        "Share Capital": "Capital",
        "Paid-in capital": "Capital",
        "Taxes payable": "Tax payable",
        "Advance payments received": "Advances",
        "Other current liabilities": "OP",
        "Accounts payable": "AP",
        "Retained earnings": "R/E",
        "Long term loans": "Long-term loans",
        # Income statement
        "Sales": "OI",
        "Operating income": "OI",
        "Operating Income": "OI",
        "Cost": "OC",
        "Operating costs": "OC",
        "Operating Costs": "OC",
        "Taxes and surcharges": "Tax and Surcharges",
        "Taxes and Surcharges": "Tax and Surcharges",
        "Tax and surcharges": "Tax and Surcharges",
        "G&A expenses": "GA",
        "G&A Expenses": "GA",
        "General and administrative expenses": "GA",
        "General and Administrative Expenses": "GA",
        "Financial expenses": "Fin Exp",
        "Financial Expenses": "Fin Exp",
        "Non-operating income": "Non-operating Income",
        "Non-operating expenses": "Non-operating Exp",
        "Non-operating Expenses": "Non-operating Exp",
    }

    results = {}
    for key in matched_keys:
        canonical = (
            key if key in all_results
            else alias_to_canonical.get(key)
            or _DEMO_HARD_ALIASES.get(key)
        )
        if canonical and canonical in all_results:
            results[key] = all_results[canonical]
        else:
            results[key] = {
                "final": "The balance remained stable throughout the review period. Refer to the schedule for detailed composition.",
                "subagent_4": "",
                "subagent_1": "",
            }
    # Pass through pre-baked section summaries under special sentinel keys so
    # the caller can populate coSummaryShape without any LLM call.
    for _sk in ("__BS_summary__", "__IS_summary__"):
        if _sk in all_results:
            results[_sk] = all_results[_sk]
    return results


def render_ai_generation_section(session_state: Any, get_model_display_name) -> None:
    # Auto-trigger: run AI automatically when data is loaded but no results yet.
    # No button needed — user gets a Redo button in the export header instead.
    auto_trigger = (
        session_state.get("dfs") is not None
        and session_state.get("ai_results") is None
    )
    if not auto_trigger:
        return

    session_state.pptx_ready = False
    session_state.pop("pptx_download_data", None)
    session_state.pop("section_summaries", None)

    progress_container = st.container()

    with progress_container:
        st.markdown("### 🔄 AI Processing Progress")
        progress_placeholder = st.empty()
        status_placeholder = st.empty()
        try:
            results = None
            selected_pipeline_dfs = build_selected_pipeline_dfs(session_state)
            reconciliation = session_state.get("reconciliation")
            statement_mode = detect_statement_mode(reconciliation)

            if statement_mode in ("is_only", "bs_only"):
                # Bypass reconciliation filter — this databook only has one
                # statement's reconciliation data (auto-detected).
                target_type = "IS" if statement_mode == "is_only" else "BS"
                _mappings = effective_mappings_from_session(session_state)
                matched_mapping_keys = [
                    k for k in selected_pipeline_dfs
                    if _mappings.get(find_mapping_key(k, _mappings) or k, {}).get("type") == target_type
                ]
                if not matched_mapping_keys:
                    matched_mapping_keys = list(selected_pipeline_dfs.keys())
                status_placeholder.info(
                    f"{'IS' if statement_mode == 'is_only' else 'BS'}-only mode: processing "
                    f"{len(matched_mapping_keys)} schedule account(s) — reconciliation filter bypassed."
                )
            else:
                matched_mapping_keys = derive_reconciliation_matched_keys(
                    reconciliation,
                    selected_pipeline_dfs.keys(),
                    session_state.get("resolution"),
                )
                has_reconciliation_data = bool(reconciliation and any(recon_df is not None and not recon_df.empty for recon_df in reconciliation))
                if not has_reconciliation_data:
                    matched_mapping_keys = list(selected_pipeline_dfs.keys())
                    if not matched_mapping_keys:
                        status_placeholder.warning("No reconciliation results and no extracted schedule data available.")
                    else:
                        status_placeholder.info(f"No reconciliation data; proceeding with all {len(matched_mapping_keys)} extracted schedule account(s).")
                if has_reconciliation_data and not matched_mapping_keys:
                    status_placeholder.warning("No eligible matched or dynamically resolved schedule accounts were found, so AI generation was skipped.")
                    matched_mapping_keys = []
            if matched_mapping_keys:
                total_items = len(matched_mapping_keys)
                # Pipeline is now 3 stages: Generator → Auditor (Polish) → Validator
                from .ai import SUBAGENT_SEQUENCE
                n_stages = len(SUBAGENT_SEQUENCE)
                total_steps = n_stages * total_items

                def update_progress(agent_num, agent_name, item_num, total_items_in_agent, completed_items, key_name=None):
                    if agent_num > n_stages:
                        key_display = f" | Key: {key_name}" if key_name else ""
                        status_placeholder.info(f"🔄 Feedback Loop: {agent_name}{key_display} — refining based on validator feedback")
                        return
                    completed_steps = completed_items
                    progress = min(completed_steps / total_steps, 1.0) if total_steps > 0 else 0.0
                    progress_placeholder.progress(progress)
                    key_display = f" | Key: {key_name}" if key_name else ""
                    if hasattr(update_progress, "start_time"):
                        elapsed = time.time() - update_progress.start_time
                        if completed_steps > 0:
                            avg_time_per_step = elapsed / completed_steps
                            remaining_steps = total_steps - completed_steps
                            eta_seconds = avg_time_per_step * remaining_steps
                            status_placeholder.info(
                                f"🔄 Running Subagent {agent_num}/{n_stages}: {agent_name} | Processing item {item_num}/{total_items_in_agent}{key_display} | Progress: {completed_steps}/{total_steps} steps | ETA: {int(eta_seconds / 60)}m {int(eta_seconds % 60)}s"
                            )
                        else:
                            status_placeholder.info(
                                f"🔄 Running Subagent {agent_num}/{n_stages}: {agent_name} | Processing item {item_num}/{total_items_in_agent}{key_display} | Progress: {completed_steps}/{total_steps} steps | ETA: Calculating..."
                            )
                    else:
                        update_progress.start_time = time.time()
                        status_placeholder.info(
                            f"🔄 Running Subagent {agent_num}/{n_stages}: {agent_name} | Processing item {item_num}/{total_items_in_agent}{key_display} | Progress: {completed_steps}/{total_steps} steps | ETA: Calculating..."
                        )

                start_time = time.time()

                # Demo-mode detection: if the uploaded filename matches the
                # configured demo file, replay pre-baked results so the whole
                # pipeline runs offline in a fixed amount of time.
                _demo_cfg = (FDDConfig().config or {}).get("demo", {})
                _demo_filename = str(_demo_cfg.get("filename") or "").strip()
                _demo_duration = int(_demo_cfg.get("progress_duration_seconds") or 25)
                _uploaded = str(session_state.get("uploaded_filename") or "").strip()
                _is_demo = bool(_demo_filename and _uploaded == _demo_filename)

                _demo_bs_sum = ""
                _demo_is_sum = ""
                if _is_demo:
                    status_placeholder.info(
                        f"🎬 Demo mode — replaying pre-recorded results for {total_items} accounts "
                        f"({_demo_duration}s simulated run)…"
                    )
                    progress_placeholder.progress(0)
                    results = _run_demo_ai(
                        matched_mapping_keys,
                        _demo_duration,
                        progress_placeholder,
                        status_placeholder,
                    )
                    _demo_bs_sum = str(results.pop("__BS_summary__", "") or "").strip()
                    _demo_is_sum = str(results.pop("__IS_summary__", "") or "").strip()
                else:
                    update_progress.start_time = start_time
                    status_placeholder.info(f"🚀 Starting AI pipeline for {total_items} accounts... | Progress: 0/{total_steps} steps | ETA: Calculating...")
                    progress_placeholder.progress(0)
                    # processing.max_workers in config.yml is a GLOBAL override
                    # (applies to any provider) — was declared but never read
                    # anywhere in the call chain until now. Leave it null to
                    # fall through to the per-provider default instead: each
                    # provider's own block (e.g. workbench.max_workers) can
                    # set a validated concurrency level, or the built-in
                    # fallback (4 local / 2 cloud) if neither is set.
                    try:
                        _configured_max_workers = (
                            load_yaml_config(get_default_config_path()).get("processing", {}).get("max_workers")
                        )
                    except Exception:
                        _configured_max_workers = None
                    results = run_ai_pipeline_with_progress(
                        mapping_keys=matched_mapping_keys,
                        dfs=selected_pipeline_dfs,
                        model_type=session_state.get("model_type", "local"),
                        model_name=session_state.get("model_name"),
                        language=session_state.language,
                        use_multithreading=session_state.get("use_multithreading", True),
                        max_workers=_configured_max_workers,
                        progress_callback=update_progress,
                        user_comments=session_state.get("account_comments", {}),
                    )

            if results is None:
                progress_placeholder.empty()
            else:
                session_state.ai_results = results

            if results is not None:
                if results and any(
                    has_meaningful_result_text(content)
                    for value in results.values() if isinstance(value, dict)
                    for content in value.values()
                ):
                    elapsed_ai = int(time.time() - start_time)
                    # Generate the BS / IS executive summaries here, alongside
                    # the account commentary, so PPTX export becomes pure XML.
                    try:
                        from .pptx import PowerPointGenerator
                        mappings = effective_mappings_from_session(session_state)
                        bs_blob: list[str] = []
                        is_blob: list[str] = []
                        for account_key, result in results.items():
                            mapping_key = find_mapping_key(account_key, mappings)
                            if not mapping_key or mapping_key not in mappings:
                                continue
                            atype = mappings[mapping_key].get("type")
                            text = extract_result_text_content(
                                (result or {}).get("final")
                                or (result or {}).get("subagent_4")
                                or (result or {}).get("subagent_2")
                                or (result or {}).get("subagent_1")
                                or ""
                            )
                            if not text.strip():
                                continue
                            if atype == "BS":
                                bs_blob.append(text)
                            elif atype == "IS":
                                is_blob.append(text)
                        is_chinese_db = (session_state.language == "Chn")
                        section_summaries: dict[str, str] = {}
                        # Skip section summary generation in demo mode — demo
                        # uses a pre-baked cache and the LLM call would defeat
                        # the demo's "no API needed" promise (and hangs when
                        # the API is down).
                        if _is_demo:
                            logger.info("Demo mode: using pre-baked section summaries from demo_results.json.")
                            if _demo_bs_sum:
                                section_summaries["BS"] = _demo_bs_sum
                            if _demo_is_sum:
                                section_summaries["IS"] = _demo_is_sum
                        else:
                            for stmt, blob in (("BS", bs_blob), ("IS", is_blob)):
                                if not blob:
                                    continue
                                # If the per-account pipeline already saw the
                                # circuit breaker trip for this language/agent,
                                # the API is clearly stressed — skip the
                                # section summary too rather than burning more
                                # retry time on a doomed call.
                                try:
                                    from .ai import _PIPELINE_BREAKER
                                    if any(_PIPELINE_BREAKER.is_open(stage) for stage in ("subagent_1", "subagent_2")):
                                        logger.info("Circuit breaker open from per-account pipeline — skipping %s section summary.", stmt)
                                        continue
                                except Exception:
                                    pass
                                joined = "\n\n".join(blob)
                                status_placeholder.info(f"✅ AI content generated ({len(results)} accounts). Generating {stmt} executive summary…")
                                summary = PowerPointGenerator.generate_section_summary(
                                    joined,
                                    is_chinese=is_chinese_db,
                                    language=("chinese" if session_state.language == "Chn" else "english"),
                                    model_type=session_state.get("model_type", "local"),
                                    model_name=session_state.get("model_name"),
                                )
                                if summary:
                                    section_summaries[stmt] = summary
                        session_state.section_summaries = section_summaries
                    except Exception as exc:
                        logger.warning("Section summary generation failed (PPTX will fall back to in-export AI): %s", exc)
                        session_state.section_summaries = {}
                    status_placeholder.info(
                        f"✅ AI content + summaries ready ({len(results)} accounts, {int(time.time() - start_time)}s). Building PPTX…"
                    )
                    try:
                        generate_pptx_presentation(
                            session_state=session_state,
                            pptx_available=True,
                        )
                    except Exception as exc:
                        logger.warning("Eager PPTX generation failed (user can retry via Export button): %s", exc)
                    status_placeholder.success(
                        f"✅ AI content + PPTX ready! {len(results)} accounts processed in {int(time.time() - start_time)}s."
                    )
                elif results:
                    status_placeholder.warning("⚠️ AI processing completed but no content was generated. This usually means the AI model is not properly configured.")
                    status_placeholder.info(f"💡 AI Mode: **{get_model_display_name(session_state.get('model_type', 'local'))}** - Check configuration and try again.")
                else:
                    status_placeholder.error("❌ AI processing failed completely - no results generated. Check AI model setup.")
                progress_placeholder.progress(1.0)
                st.rerun()
        except Exception as exc:
            progress_placeholder.empty()
            status_placeholder.error(f"❌ Error: {exc}")
            st.code(traceback.format_exc())


def render_generated_content(session_state: Any, account_display_dfs, mappings: Dict[str, Any], get_model_display_name) -> None:
    if not session_state.ai_results:
        return

    content_keys = []
    for key, result in session_state.ai_results.items():
        if isinstance(result, dict):
            hydrate_nested_agent_outputs(result)
            if _result_has_pipeline_content(result):
                content_keys.append(key)

    dfs = session_state.get("dfs") or {}
    bs_keys, is_keys, other_keys = split_accounts_by_type(content_keys, mappings, dfs=dfs)
    has_content = any(
        isinstance(value, dict) and _result_has_pipeline_content(value)
        for value in session_state.ai_results.values()
    )
    if not has_content:
        st.warning("⚠️ AI processing completed but no content was generated.")
        st.error("**Possible causes:**")
        st.markdown("- AI service is not running or is unreachable")
        st.markdown("- The configured model is unavailable or not loaded")
        st.markdown("- API credentials or endpoint settings are invalid")
        st.markdown("- The request hit a network or rate-limit issue")
        st.info("💡 **AI Mode:** " + get_model_display_name(session_state.get("model_type", "local")))
        st.info("🔧 Configure your AI model and try again.")

    if not bs_keys and not is_keys and not other_keys:
        st.warning("⚠️ No AI results to display with content")
        st.info(f"Found {len(session_state.ai_results)} results but none have content. Check debug info above.")
        return

    tab_list = []
    if bs_keys:
        tab_list.append(f"Balance Sheet ({len(bs_keys)} accounts)")
    if is_keys:
        tab_list.append(f"Income Statement ({len(is_keys)} accounts)")
    if other_keys:
        tab_list.append(f"Other ({len(other_keys)} accounts)")
    ai_tabs = st.tabs(tab_list)
    tab_idx = 0

    @st.fragment
    def _render_commentary_fragment(detailed_content: str, clause_reviews: list):
        """Render the final commentary + validator evidence for a single account."""
        final_content = detailed_content
        st.markdown(build_highlighted_commentary_html(str(final_content), clause_reviews or []), unsafe_allow_html=True)
        if clause_reviews:
            hallucination_count = sum(
                1 for r in clause_reviews
                if isinstance(r, dict) and str(r.get("category", "")).lower() == "hallucination"
            )
            reasoning_count = sum(
                1 for r in clause_reviews
                if isinstance(r, dict) and str(r.get("category", "")).lower() == "reasoning"
            )
            flagged_count = hallucination_count + reasoning_count
            caption_parts = [f"Validator reviewed {len(clause_reviews)} clause(s)"]
            if hallucination_count:
                caption_parts.append(f"{hallucination_count} hallucination(s)")
            if reasoning_count:
                caption_parts.append(f"{reasoning_count} reasoning")
            if not flagged_count:
                caption_parts.append("all data-backed")
            st.caption("; ".join(caption_parts) + ".")
            with st.expander("Validator evidence review", expanded=False):
                review_rows = [
                    {
                        "Clause": str(review.get("clause") or ""),
                        "Category": str(review.get("category") or "data-backed").replace("-", " ").title(),
                        "Supported": "Yes" if bool(review.get("supported")) else "No",
                        "Reason": str(review.get("reason") or ""),
                    }
                    for review in clause_reviews
                    if isinstance(review, dict)
                ]
                if review_rows:
                    st.dataframe(pd.DataFrame(review_rows), use_container_width=True, hide_index=True)

    def create_account_agent_tabs(keys, prefix=""):
        account_tabs = st.tabs([f"📄 {key}" for key in keys])
        agent_map = {"subagent_1": "Generator", "subagent_2": "Auditor", "subagent_3": "Refiner", "subagent_4": "Validator", "final": "Final (Validator)"}
        for acc_idx, key in enumerate(keys):
            with account_tabs[acc_idx]:
                result = session_state.ai_results.get(key, {})
                if not isinstance(result, dict):
                    result = {}
                hydrate_nested_agent_outputs(result)
                selected_df = get_account_dataframe(key, account_display_dfs)
                detailed_content = extract_result_text(result, "final")
                validator_metadata = extract_validator_metadata(result)
                clause_reviews = validator_metadata.get("clause_reviews") if isinstance(validator_metadata, dict) else []
                has_final = has_meaningful_result_text(detailed_content)
                reprompt_mode = str(result.get("reprompt_mode") or "").strip()
                final_label = "Final (Reprompt + validator)" if reprompt_mode == "generator_reprompt_validated" else ("Final (Generator reprompt)" if reprompt_mode == "generator_only" else "Final (Validator)")
                if has_final:
                    st.markdown(f"#### ✨ {final_label}")
                    _render_commentary_fragment(detailed_content=str(detailed_content), clause_reviews=clause_reviews)

                # Agent Pipeline sits directly with Validator evidence (both are
                # "how did we get this answer" review artifacts), ahead of the
                # source-remark background context and the reprompt control.
                agent_contents = []
                agent_names_list = []
                for agent_key in ["subagent_1", "subagent_2", "subagent_3", "subagent_4"]:
                    content = extract_result_text(result, agent_key)
                    if has_meaningful_result_text(content):
                        agent_name = agent_map.get(agent_key, agent_key)
                        agent_contents.append((agent_name, str(content)))
                        agent_names_list.append(agent_name)
                if agent_contents:
                    with st.expander(f"🔍 Agent Pipeline ({', '.join(agent_names_list)})", expanded=False):
                        for content_idx, (agent_name, content) in enumerate(agent_contents):
                            st.markdown(f"**{agent_name}:**")
                            st.text_area(
                                label=f"Content for {agent_name}",
                                value=content,
                                height=min(max(80, int(len(str(content)) / 4)), 400),
                                key=f"{prefix}{key}_{agent_name}_pipeline",
                                label_visibility="collapsed",
                            )
                            if content_idx < len(agent_contents) - 1:
                                st.markdown("---")

                if has_final:
                    render_account_remarks_context(selected_df, key, session_state.get("language", "Eng"), prefix=f"{prefix}generated_")

                if not has_final and not agent_contents:
                    st.warning("No agent outputs available for this account")

                st.markdown("---")
                with st.expander(f"✏️ Reprompt {key}", expanded=False):
                    reprompt_comment = st.text_area(
                        label=f"Reprompt guidance for {key}",
                        value=session_state.account_comments.get(key, ""),
                        placeholder="Add comments to refine this account only, then click reprompt.",
                        key=f"{prefix}{key}_reprompt_comment",
                        height=90,
                    )
                    session_state.account_comments[key] = reprompt_comment
                    if st.button(f"Reprompt {key}", key=f"{prefix}{key}_reprompt_button", use_container_width=True):
                        with st.spinner(f"Regenerating {key}..."):
                            selected_pipeline_dfs = build_selected_pipeline_dfs(session_state)
                            updated_results = run_generator_reprompt(
                                mapping_keys=[key],
                                dfs={key: selected_pipeline_dfs[key]},
                                existing_results=session_state.ai_results,
                                model_type=session_state.get("model_type", "local"),
                                model_name=session_state.get("model_name"),
                                language=session_state.language,
                                user_comments={key: reprompt_comment},
                            )
                            merged_results = dict(session_state.ai_results or {})
                            merged_results.update(updated_results)
                            session_state.ai_results = merged_results
                            session_state.pptx_ready = False
                        st.rerun()

    if bs_keys:
        with ai_tabs[tab_idx]:
            create_account_agent_tabs(bs_keys, "bs_")
        tab_idx += 1
    if is_keys:
        with ai_tabs[tab_idx]:
            create_account_agent_tabs(is_keys, "is_")
        tab_idx += 1
    if other_keys:
        with ai_tabs[tab_idx]:
            for key in other_keys:
                result = session_state.ai_results.get(key, {})
                if not isinstance(result, dict):
                    result = {}
                with st.expander(f"📄 {key}", expanded=False):
                    st.json(result)
# --- end ui/ai_panel.py ---

# --- begin ui/processed.py ---
from typing import Any, List

import pandas as pd
import streamlit as st

from .workbook import find_mapping_key, load_mappings


def render_reconciliation_metrics(recon_df: pd.DataFrame):
    matches = int(recon_df["Match"].isin(["✅ Match", "⚠️ Match"]).sum())
    diffs = int((recon_df["Match"] == "❌ Diff").sum())
    not_found = int((recon_df["Match"] == "⚠️ Not Found").sum())
    immaterial = int((recon_df["Match"] == "✅ Immaterial").sum())
    stats = [
        ("✅", matches, "Matches"),
        ("❌", diffs, "Differences"),
        ("⚠️", not_found, "Not Found"),
        ("✅", immaterial, "Immaterial"),
        ("📋", len(recon_df), "Checked Rows"),
    ]
    # Compact custom "badge" instead of st.metric() (which has a lot of
    # vertical padding + an unused delta row) — spread evenly across the
    # full row width via st.columns, with a large bold number so counts
    # stay legible at a glance instead of being buried in a single dense line.
    cols = st.columns(len(stats))
    for col, (emoji, count, label) in zip(cols, stats):
        with col:
            st.markdown(
                f"<div style='text-align:center; line-height:1.2'>"
                f"<span style='font-size:0.95rem'>{emoji}</span> "
                f"<span style='font-size:1.6rem; font-weight:700'>{count}</span><br>"
                f"<span style='font-size:0.8rem; color:#888'>{label}</span>"
                f"</div>",
                unsafe_allow_html=True,
            )


def account_has_non_zero_values(df: pd.DataFrame | None) -> bool:
    if df is None or df.empty:
        return False
    numeric_df = df.select_dtypes(include=["number"])
    if numeric_df.empty:
        return True
    return bool((numeric_df.fillna(0).abs() >= 0.01).any().any())


def filter_reconciliation_display_rows(recon_df: pd.DataFrame | None) -> tuple[pd.DataFrame | None, int]:
    if recon_df is None or recon_df.empty or "Mapping_Status" not in recon_df.columns:
        return recon_df, 0
    zero_mask = recon_df["Mapping_Status"].astype(str).eq("Zero source")
    filtered_df = recon_df.loc[~zero_mask].copy()
    warning_match_map = {
        "Missing mapping": "⚠️ Map",
        "Tab-only match": "⚠️ Tab",
        "Mapped but missing tab": "⚠️ No tab",
    }
    if "Match" in filtered_df.columns:
        for mapping_status, short_label in warning_match_map.items():
            warning_mask = filtered_df["Mapping_Status"].astype(str).eq(mapping_status)
            filtered_df.loc[warning_mask, "Match"] = short_label
    hidden_columns = {"Mapping_Status", "Mapping_Note", "Integrity_Flag"}
    priority_columns = [
        "Financials_Account",
        "Mapping_Key",
        "Mapping_Status",
        "Match",
        "Mapping_Note",
    ]
    ordered_columns = [column for column in priority_columns if column in filtered_df.columns and column not in hidden_columns]
    ordered_columns.extend(
        column for column in filtered_df.columns if column not in ordered_columns and column not in hidden_columns
    )
    return filtered_df.loc[:, ordered_columns], int(zero_mask.sum())


def reconciliation_warning_row_count(recon_df: pd.DataFrame | None) -> int:
    if recon_df is None or recon_df.empty or "Mapping_Status" not in recon_df.columns:
        return 0
    warning_statuses = {"Missing mapping", "Tab-only match", "Mapped but missing tab"}
    return int(recon_df["Mapping_Status"].astype(str).isin(warning_statuses).sum())


# Slim, human-facing view of a reconciliation row — drops Tab_Account,
# Match, Mapping_Status/Note, Projection_Date, Integrity_Flag (still used
# internally for filtering/metrics, just not shown as table columns).
_RECON_DISPLAY_COLUMN_MAP = {
    "Mapping_Key": "Key",
    "Financials_Account": "Account",
    "Date": "Date",
    "Financials_Value": "Value",
    "Tab_Value": "BKD_value",
    "Diff": "diff",
    "Projection_Stage": "projection_stage",
}


def _trim_reconciliation_columns_for_display(df: pd.DataFrame) -> pd.DataFrame:
    present = [column for column in _RECON_DISPLAY_COLUMN_MAP if column in df.columns]
    return df.loc[:, present].rename(columns=_RECON_DISPLAY_COLUMN_MAP)


def _render_single_reconciliation_tab(
    recon_df: pd.DataFrame | None,
    statement_type: str,
    empty_message: str,
) -> None:
    if recon_df is None or recon_df.empty:
        st.info(empty_message)
        return

    display_recon_df, hidden_zero_rows = filter_reconciliation_display_rows(recon_df)
    warning_row_count = reconciliation_warning_row_count(recon_df)
    if warning_row_count:
        st.caption(f"{warning_row_count} row(s) have mapping/tab coverage warnings — see the metrics below.")
    st.dataframe(
        format_dataframe_for_display(_trim_reconciliation_columns_for_display(display_recon_df)),
        use_container_width=True, height=320,
    )
    render_reconciliation_metrics(display_recon_df)


def render_reconciliation_section(
    recon_df: pd.DataFrame | None,
    statement_type: str,
    empty_message: str,
) -> None:
    _render_single_reconciliation_tab(recon_df, statement_type, empty_message)


def _render_resolver_diagnostics(resolution: Dict[str, Any], display_keys: list[str]) -> None:
    """Show sheet-to-mapping resolution details in a debug expander."""
    resolved = resolution.get("resolved") or {}
    unresolved = resolution.get("unresolved_sheets") or []
    norm_errors = resolution.get("normalization_errors") or {}

    rows = []
    for mapping_key, info in resolved.items():
        rows.append({
            "Mapping Key": mapping_key,
            "Sheet": info.get("sheet_name", ""),
            "Method": info.get("resolution_method", ""),
            "Score": info.get("score", ""),
            "Alias": info.get("matched_alias", ""),
            "In DFS": "yes" if info.get("sheet_name", "") in display_keys else "no",
        })
    for sheet_name in unresolved:
        rows.append({
            "Mapping Key": "-",
            "Sheet": sheet_name,
            "Method": "UNRESOLVED",
            "Score": "",
            "Alias": "",
            "In DFS": "yes" if sheet_name in display_keys else "no",
        })
    for sheet_name, detail in norm_errors.items():
        rows.append({
            "Mapping Key": "-",
            "Sheet": sheet_name,
            "Method": "NORM ERROR",
            "Score": "",
            "Alias": str(detail),
            "In DFS": "no",
        })
    if rows:
        with st.expander("Debug: Sheet → Mapping Resolution", expanded=False):
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


def render_data_tables_section(session_state: Any) -> None:
    """Renders the BS/IS tab group -- each with a "Reconciliation" sub-tab
    plus one sub-tab per extracted account showing its full breakdown
    table (the same view render_account_panel always showed inside
    render_processed_view). Factored out so a caller that only has DATA
    (no AI results yet -- e.g. a batch entity still mid-pipeline) can show
    this same rich view without also pulling in the AI-generation section,
    which auto-triggers a real AI run the moment it sees ai_results is
    None.
    """
    account_display_dfs = session_state.get("display_dfs") or session_state.dfs
    account_display_workbook_list = session_state.get("display_workbook_list") or session_state.workbook_list
    mappings = effective_mappings_from_session(session_state)

    display_account_keys = []
    seen_accounts = set()
    source_account_keys = account_display_workbook_list or list(account_display_dfs.keys())
    for key in source_account_keys:
        if key in account_display_dfs and key not in seen_accounts:
            display_account_keys.append(key)
            seen_accounts.add(key)

    dfs = session_state.get("dfs") or {}
    display_groups = build_processed_display_groups(display_account_keys, mappings, dfs=dfs)
    bs_accounts = display_groups["bs_accounts"]
    is_accounts = display_groups["is_accounts"]

    recon_bs = session_state.reconciliation[0] if session_state.reconciliation else None
    recon_is = session_state.reconciliation[1] if session_state.reconciliation else None

    def render_account_panel(key: str):
        mapping_key = find_mapping_key(key, mappings)
        account_type = mappings.get(mapping_key, {}).get("type", "") if mapping_key else ""
        selected_df = get_account_dataframe(key, account_display_dfs)
        if selected_df is None:
            st.warning(f"Data not found for account: {key}")
            return
        if account_type not in {"BS", "IS"}:
            st.warning("This schedule tab is not classified into BS/IS in `fdd_utils/mappings.yml`. If it should be part of the standard account flow, add or update its mapping aliases.")

        account_display_df = build_account_display_dataframe(selected_df)
        st.dataframe(format_dataframe_for_display(account_display_df), use_container_width=True)
        if len(selected_df.columns) > 1 and account_type in {"BS", "IS"}:
            integrity = selected_df.attrs.get("integrity") or {}
            latest_period = str(integrity.get("effective_date") or selected_df.columns[1])
            annualization_months = (
                selected_df.attrs.get("annualization_months")
                if account_type == "IS"
                else None
            )
            if annualization_months in (None, "") and account_type == "IS":
                annualization_months = integrity.get("annualization_months")
            fiscal_year_end_month = integrity.get("fiscal_year_end_month") if account_type == "IS" else None
            fiscal_year_end_day = integrity.get("fiscal_year_end_day") if account_type == "IS" else None
            total_analysis_periods = max(len(account_display_df.columns) - 1, 1) if account_display_df is not None else 1
            st.caption(
                f"{describe_statement_period(account_type, str(selected_df.columns[1]), annualization_months=annualization_months, fiscal_year_end_month=fiscal_year_end_month, fiscal_year_end_day=fiscal_year_end_day)} | Latest target period: {latest_period} | Displaying {total_analysis_periods} indicative-adjusted period(s)"
            )
        if account_display_df is not None and "Prompt remarks (RHS 1-5 cols)" in [str(column) for column in account_display_df.columns]:
            st.caption("Displaying inline prompt remarks extracted from the source 1-5 RHS columns.")
        render_account_remarks_context(selected_df, key, session_state.get("language", "Eng"), prefix="processed_")
        existing_comment = session_state.account_comments.get(key, "")
        session_state.account_comments[key] = st.text_area(
            label=f"User remarks for {key}",
            value=existing_comment,
            placeholder="Add table-specific remarks, nature/trend reasons, or reprompt guidance...",
            key=f"table_comment_{key}",
            height=90,
        )

    def render_account_tabs(account_names: List[str], group_type: str, recon_df: pd.DataFrame | None, empty_recon_message: str):
        # "Reconciliation" is the first sub-tab within the statement's own
        # tab group (rather than a separate top-level tab) so BS/IS content
        # and their reconciliation view live together.
        item_tabs = st.tabs(["Reconciliation", *account_names])
        with item_tabs[0]:
            render_reconciliation_section(recon_df=recon_df, statement_type=group_type, empty_message=empty_recon_message)
            if not account_names:
                st.info(f"No {group_type} items available.")
        for index, key in enumerate(account_names, start=1):
            with item_tabs[index]:
                render_account_panel(key)

    data_tabs = st.tabs(display_groups["tab_names"])
    with data_tabs[0]:
        render_account_tabs(bs_accounts, "BS", recon_bs, "No Balance Sheet reconciliation data available.")
    with data_tabs[1]:
        render_account_tabs(is_accounts, "IS", recon_is, "No Income Statement reconciliation data available.")


def render_processed_view(
    session_state: Any,
    generate_pptx_callback,
    get_model_display_name,
    before_ai_section: Optional[Callable[[], None]] = None,
) -> None:
    mappings = effective_mappings_from_session(session_state)
    account_display_dfs = session_state.get("display_dfs") or session_state.dfs
    resolution = session_state.get("resolution") or {}
    display_account_keys = []
    seen_accounts = set()
    account_display_workbook_list = session_state.get("display_workbook_list") or session_state.workbook_list
    source_account_keys = account_display_workbook_list or list(account_display_dfs.keys())
    for key in source_account_keys:
        if key in account_display_dfs and key not in seen_accounts:
            display_account_keys.append(key)
            seen_accounts.add(key)
    profile_map = resolution.get("profiles") or {}
    manual_sheet_options = [sheet_name for sheet_name, profile in profile_map.items() if profile.get("sheet_kind") == "financial_schedule"]
    manual_target_options = sorted(
        {
            *get_financial_account_options(session_state.get("bs_is_results")),
            *[key for key in mappings.keys() if not str(key).startswith("_")],
            *list((session_state.get("mapping_overrides") or {}).keys()),
        }
    )

    with st.expander("Manual Mapping Overrides", expanded=bool(session_state.get("mapping_overrides"))):
        st.caption("Use this when automatic resolution is not acceptable. Manual overrides win over fuzzy, figure, and AI matching on the next reprocess.")
        if resolution.get("override_issues"):
            st.warning("Some overrides could not be applied in the latest run.")
            st.dataframe(pd.DataFrame(resolution.get("override_issues") or []), use_container_width=True)
        current_overrides = session_state.get("mapping_overrides") or {}
        if current_overrides:
            st.markdown("**Current overrides**")
            st.dataframe(
                pd.DataFrame([{"target": key, "sheet_name": value} for key, value in current_overrides.items()]),
                use_container_width=True,
                hide_index=True,
            )
        override_target = st.selectbox("Financials account or mapping target", options=manual_target_options or [""], index=0, key="manual_override_target")
        override_sheet = st.selectbox("Workbook tab", options=manual_sheet_options or [""], index=0, key="manual_override_sheet")
        add_col, remove_col, clear_col = st.columns(3)
        with add_col:
            if st.button("Apply Override", use_container_width=True, key="apply_manual_override"):
                if not override_target or not override_sheet:
                    st.warning("Select both a target account and a workbook tab.")
                else:
                    session_state.mapping_overrides[override_target] = override_sheet
                    session_state.process_data_clicked = True
                    st.rerun()
        with remove_col:
            if st.button("Remove Selected", use_container_width=True, key="remove_manual_override"):
                if override_target in session_state.mapping_overrides:
                    del session_state.mapping_overrides[override_target]
                    session_state.process_data_clicked = True
                    st.rerun()
        with clear_col:
            if st.button("Clear All Overrides", use_container_width=True, key="clear_manual_overrides"):
                session_state.mapping_overrides = {}
                session_state.process_data_clicked = True
                st.rerun()

    debug_output = session_state.get("debug_output", "")
    if debug_output:
        with st.expander("Debug: Extraction & Reconciliation Log", expanded=False):
            st.code(debug_output, language="text")

        # Show resolver diagnostics alongside debug log
        _render_resolver_diagnostics(resolution, display_account_keys)

    render_data_tables_section(session_state)

    st.markdown("---")
    col_header, col_pptx, col_redo, col_download = st.columns([3, 1, 0.4, 0.3])
    with col_header:
        st.header("🤖 AI Content Generation")
        if not (session_state.get("pptx_ready", False) and session_state.get("pptx_download_filename")):
            st.caption("AI runs automatically after processing. Use 🔄 to redo.")
    with col_pptx:
        st.markdown("<br>", unsafe_allow_html=True)
        pptx_key = f"pptx_btn_{session_state.button_click_counter}"
        pptx_cached = bool(session_state.get("pptx_ready")) and bool(session_state.get("pptx_download_data"))
        pptx_label = "📄 Regenerate PPTX" if pptx_cached else "📄 Generate & Export PPTX"
        if st.button(pptx_label, type="secondary", use_container_width=True, disabled=session_state.get("ai_results") is None, key=pptx_key):
            session_state.button_click_counter += 1
            session_state.pptx_ready = False
            session_state.pop("pptx_download_data", None)
            generate_pptx_callback()
    with col_redo:
        st.markdown("<br>", unsafe_allow_html=True)
        redo_key = f"redo_ai_{session_state.button_click_counter}"
        if st.button("🔄", help="Re-run AI generation", key=redo_key, use_container_width=True):
            session_state.ai_results = None
            session_state.pptx_ready = False
            session_state.pop("pptx_download_data", None)
            session_state.pop("section_summaries", None)
            session_state.button_click_counter += 1
            st.rerun()
    with col_download:
        st.markdown("<br>", unsafe_allow_html=True)
        if session_state.get("pptx_ready", False) and "pptx_download_data" in session_state:
            st.download_button(
                label="📥",
                data=session_state.pptx_download_data,
                file_name=session_state.pptx_download_filename,
                mime=session_state.pptx_download_mime,
                help="Download generated PPTX",
                key=f"download_icon_{session_state.button_click_counter}",
                use_container_width=True,
            )

    if before_ai_section:
        before_ai_section()
    render_ai_generation_section(session_state, get_model_display_name)
    render_generated_content(session_state, account_display_dfs, mappings, get_model_display_name)
# --- end ui/processed.py ---

# --- begin ui/sidebar.py ---
from datetime import timedelta
import hashlib
from pathlib import Path
import re
import tempfile
import time
from typing import Any, Callable

import streamlit as st

from .workbook import clear_table_inspection_cache, clear_workbook_caches


def _safe_stem(uploaded_name: str) -> str:
    stem = Path(uploaded_name or "databook.xlsx").stem or "databook"
    sanitized = re.sub(r"[^A-Za-z0-9._-]+", "_", stem).strip("._")
    return sanitized or "databook"


def persist_uploaded_workbook(
    uploaded_name: str,
    uploaded_bytes: bytes,
    session_state,
    cache_dir: str | None = None,
    state_key: str = "temp_path",
) -> str:
    digest = hashlib.sha256(uploaded_bytes).hexdigest()[:16]
    target_dir = Path(cache_dir or tempfile.gettempdir()) / "python-pptx-uploads"
    target_dir.mkdir(parents=True, exist_ok=True)

    target_path = target_dir / f"{_safe_stem(uploaded_name)}-{digest}.xlsx"
    if not target_path.exists():
        target_path.write_bytes(uploaded_bytes)

    session_state[state_key] = str(target_path)
    # uploaded_workbook_digest drives the PRIMARY upload's cache-invalidation
    # logic elsewhere -- only set it for the primary upload, not a secondary
    # file (e.g. a roll-up workbook) persisted under a different state_key.
    if state_key == "temp_path":
        session_state["uploaded_workbook_digest"] = digest
    return str(target_path)


def cleanup_stale_uploads(
    cache_dir: str | None = None,
    max_age: timedelta = timedelta(days=2),
) -> int:
    target_dir = Path(cache_dir or tempfile.gettempdir()) / "python-pptx-uploads"
    if not target_dir.exists():
        return 0

    removed = 0
    cutoff_seconds = max_age.total_seconds()
    now = time.time()
    for candidate in target_dir.glob("*.xlsx"):
        age_seconds = now - candidate.stat().st_mtime
        if age_seconds > cutoff_seconds:
            candidate.unlink(missing_ok=True)
            removed += 1
    return removed


def _build_model_choices() -> list[dict]:
    """Model choices for the sidebar dropdown, built from config.yml so a new
    Workbench model added to config (or a renamed local chat_model) shows up
    without a code change. Each choice: key, model_type, model_name, label, ready.
    """
    try:
        config = load_yaml_config(get_default_config_path())
    except Exception:
        config = {}

    choices: list[dict] = []
    wb_ready = is_provider_ready(config, "workbench")
    wb_models = ((config.get("workbench") or {}).get("available_models")) or WORKBENCH_AVAILABLE_MODELS
    for entry in wb_models:
        model_id = entry.get("id") if isinstance(entry, dict) else str(entry)
        label = entry.get("label") if isinstance(entry, dict) else str(entry)
        if not model_id:
            continue
        choices.append({
            "key": f"workbench::{model_id}",
            "model_type": "workbench",
            "model_name": model_id,
            "label": f"{label} (Workbench)",
            "ready": wb_ready,
        })

    local_ready = is_provider_ready(config, "local")
    local_chat_model = (config.get("local") or {}).get("chat_model") or "Qwen3-32B"
    choices.append({
        "key": "local::default",
        "model_type": "local",
        "model_name": None,
        "label": f"{local_chat_model} (Local)",
        "ready": local_ready,
    })
    return choices


def render_sidebar_upload(session_state: Any, get_model_display_name: Callable[[str], str]) -> str | None:
    with st.sidebar:
        model_choices = _build_model_choices()
        if "model_choice_key" not in session_state:
            # Default to the first choice (GPT-5.5) per project policy, even if
            # it isn't configured yet — the warning below tells the user why,
            # rather than silently switching their default to something else.
            session_state.model_choice_key = model_choices[0]["key"] if model_choices else None
        choice_by_key = {c["key"]: c for c in model_choices}
        current_key = session_state.get("model_choice_key")
        if current_key not in choice_by_key and model_choices:
            current_key = model_choices[0]["key"]

        st.markdown("**🤖 AI Model**")
        selected_key = st.selectbox(
            "AI Model",
            options=[c["key"] for c in model_choices],
            format_func=lambda k: choice_by_key[k]["label"] + ("" if choice_by_key[k]["ready"] else " ⚠️ not configured"),
            index=[c["key"] for c in model_choices].index(current_key) if current_key in choice_by_key else 0,
            label_visibility="collapsed",
        )
        selected = choice_by_key.get(selected_key, {})
        session_state.model_choice_key = selected_key
        session_state.model_type = selected.get("model_type", "local")
        session_state.model_name = selected.get("model_name")
        if not selected.get("ready", True):
            st.caption(
                f"⚠️ {selected.get('label')} is not configured — set its api_key "
                f"in fdd_utils/config.yml. Falling back to the first ready provider at run time."
            )
        else:
            st.caption(f"🤖 AI Mode: {selected.get('label', get_model_display_name(session_state.model_type))}")
        # Parallel processing is on by default for every provider except
        # "local" (its server serves one request effectively serially, so
        # concurrency just queues rather than helps) — no UI toggle needed;
        # override globally via processing.use_multithreading in config.yml
        # if a specific run needs to be forced sequential (e.g. to rule out
        # a concurrency issue or stay under a strict rate limit).
        try:
            _proc_cfg = load_yaml_config(get_default_config_path()).get("processing", {}) or {}
        except Exception:
            _proc_cfg = {}
        session_state.use_multithreading = bool(_proc_cfg.get("use_multithreading", True))

        st.markdown("**📁 Databook File(s)**")
        uploaded_files = st.file_uploader(
            "Upload Excel file(s)",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            help="Upload one databook for a single-entity report, or several at once to "
                 "batch-process them (one PPTX per entity, no extra mode toggle needed).",
            key="file_uploader",
        )
        uploaded_files = uploaded_files or []

        if len(uploaded_files) > 1:
            # Batch mode is now purely a function of "how many files did you
            # upload" -- no separate checkbox to remember to flip. Every
            # file is persisted here (not left as a live UploadedFile
            # object) so render_batch_processing_section can read stable
            # temp paths back out of session_state across reruns, the same
            # pattern persist_uploaded_workbook already establishes for the
            # single-file path and the batch section's own roll-up upload.
            session_state.batch_mode = True
            persisted_slots = []
            for f in uploaded_files:
                slot_id = re.sub(r"[^\w\-]", "_", f"{f.name}_{f.size}")
                slot_temp_path = persist_uploaded_workbook(
                    uploaded_name=f.name,
                    uploaded_bytes=f.getvalue(),
                    session_state=session_state,
                    state_key=f"batch_temp_path_{slot_id}",
                )
                persisted_slots.append({"name": f.name, "size": f.size, "temp_path": slot_temp_path})
            session_state.batch_uploaded_files_meta = persisted_slots
            st.caption(f"📦 {len(uploaded_files)} files uploaded -- batch mode active (configure below on the main page).")
            return None

        session_state.batch_mode = False
        session_state.batch_uploaded_files_meta = []

        uploaded_file = uploaded_files[0] if uploaded_files else None

        if uploaded_file:
            session_state["uploaded_filename"] = uploaded_file.name
            temp_path = persist_uploaded_workbook(
                uploaded_name=uploaded_file.name,
                uploaded_bytes=uploaded_file.getvalue(),
                session_state=session_state,
            )
            session_state.upload_cache_cleanup_removed = cleanup_stale_uploads()
            prev_file = session_state.get("prev_uploaded_temp_path", None)
            if prev_file != temp_path:
                clear_workbook_caches()
                reset_processing_session_state(session_state, clear_upload_reference=False)
                session_state.prev_uploaded_temp_path = temp_path

            st.success(f"✅ File loaded: {uploaded_file.name}")
            session_state.temp_path = temp_path
        else:
            st.warning("⚠️ Please upload a databook file to begin")
            temp_path = None
            if "temp_path" in session_state:
                del session_state["temp_path"]

        return temp_path


def render_language_selector(session_state: Any) -> None:
    """Language radio, factored out of render_sidebar_upload so callers can
    place it next to the Financial Statement Sheet selector (which otherwise
    leaves blank space beside the taller Entity Name column) instead of
    always in the sidebar."""
    # Prefer the authoritative post-Process detection; fall back to the cheap
    # pre-Process preview so this is visible even before the user clicks
    # Process -- that's the whole point (so an override is an informed
    # choice, not a guess about what the databook actually is). Auto-detection
    # may store "Chi" -- normalise to the UI convention ("Eng"/"Chn") so the
    # radio + all downstream == "Chn" checks agree.
    detected = session_state.get("detected_language") or session_state.get("detected_language_preview")
    detected_norm = None
    if detected:
        detected_norm = "Chn" if str(detected).strip() in ("Chi", "Chn", "chinese", "Chinese") else "Eng"

    if not session_state.get("language_user_set") and detected_norm:
        # No manual override yet -- the detected value IS the default
        # selection, not just a side note next to a radio that still
        # defaults to English regardless of what was detected.
        current_lang = detected_norm
        session_state.language = detected_norm
    else:
        current_lang = session_state.get("language", "Eng")
        if current_lang not in ("Eng", "Chn"):
            current_lang = "Chn" if str(current_lang).strip() in ("Chi", "chinese", "Chinese") else "Eng"

    label_col, radio_col, status_col = st.columns([1, 2, 2])
    with label_col:
        st.markdown("<div style='padding-top: 8px'>🌐 Language</div>", unsafe_allow_html=True)
    with radio_col:
        # No widget `key`: session_state.language is the single source of truth
        # (a keyed radio + index fight each other and make the override "stick"
        # only intermittently). index seeds it; the return value writes it back.
        selected_lang = st.radio(
            "🌐 Language",
            options=["Eng", "Chn"],
            format_func=lambda x: "Eng" if x == "Eng" else "中文",
            index=0 if current_lang == "Eng" else 1,
            horizontal=True,
            label_visibility="collapsed",
        )
    if selected_lang != current_lang:
        session_state.language = selected_lang
        session_state.language_user_set = True   # stop auto-detect from overwriting

    with status_col:
        if detected:
            detected_label = "Chinese" if detected_norm == "Chn" else "English"
            st.success(f"Detected: {detected_label}", icon="✅")


# --- end ui/sidebar.py ---

# --- begin ui/pptx_export.py ---
import datetime as dt_module
import logging
import os
import re
import time
from typing import Any, Callable, Dict, Optional

import streamlit as st

from .pptx import build_pptx_structured_payloads
from .workbook import find_mapping_key, get_effective_mappings, load_mappings

logger = logging.getLogger(__name__)


def generate_pptx_presentation(
    *,
    session_state: Any,
    pptx_available: bool,
) -> None:
    if not session_state.ai_results:
        st.error("❌ No AI results available. Generate AI content first.")
        return

    if not pptx_available:
        st.error("❌ PPTX generation not available. Missing required modules.")
        return

    project_name = session_state.get("project_name", "Project")
    entity_name = session_state.get("entity_name", project_name)
    language = session_state.get("language", "Eng")
    mappings = effective_mappings_from_session(session_state)

    template_path = None
    for template in ["fdd_utils/template.pptx", "template.pptx"]:
        if os.path.exists(template):
            template_path = template
            break
    if not template_path:
        st.error("❌ PowerPoint template not found. Please add `fdd_utils/template.pptx` or `template.pptx`.")
        return

    output_dir = "fdd_utils/output"
    os.makedirs(output_dir, exist_ok=True)

    timestamp = dt_module.datetime.now().strftime("%Y%m%d_%H%M%S")
    sanitized_entity = re.sub(r"[^\w\-_]", "_", str(entity_name)).strip("_") or "Project"
    selected_pipeline_dfs = build_selected_pipeline_dfs(session_state)

    try:
        combined_output_path = os.path.join(output_dir, f"{sanitized_entity}_{timestamp}.pptx")
        from fdd_utils.pptx import export_pptx_from_structured_data_combined

        structured_payloads = build_pptx_structured_payloads(
            ai_results=session_state.ai_results,
            mappings=mappings,
            bs_is_results=session_state.bs_is_results,
            dfs=selected_pipeline_dfs,
        )
        bs_data = structured_payloads.get("BS", [])
        is_data = structured_payloads.get("IS", [])
        logger.debug("PPTX payload account counts | BS=%s | IS=%s", len(bs_data), len(is_data))

        if not bs_data and not is_data:
            st.error("❌ No content generated for PPTX")
            logger.debug(
                "PPTX payload is empty | ai_results_keys=%s | dfs_keys=%s",
                list(session_state.ai_results.keys())[:10] if session_state.ai_results else "None",
                list(selected_pipeline_dfs.keys())[:10] if selected_pipeline_dfs else "None",
            )
            return

        # Demo mode: skip coSummaryShape AI so export is instant.
        _demo_cfg2 = (FDDConfig().config or {}).get("demo", {})
        _is_demo_pptx = bool(
            _demo_cfg2.get("filename") and
            str(session_state.get("uploaded_filename") or "").strip() == str(_demo_cfg2.get("filename") or "").strip()
        )
        # embed_financial_tables reads (temp_path, selected_sheet) for its
        # currency-unit-label detection and as a fresh-extraction fallback --
        # when this entity's financials came from an uploaded roll-up
        # workbook (the "進階：主表" expander), that source -- not this
        # entity's own file/selected_sheet -- is the one that actually holds
        # them. Mirrors process_workbook_data's own precedence (financials_from
        # or temp_path).
        _rollup_temp_path = session_state.get("rollup_temp_path")
        _financials_workbook_path = _rollup_temp_path or session_state.get("temp_path")
        _financials_sheet_name = (
            session_state.get("rollup_sheet") if _rollup_temp_path else session_state.get("selected_sheet")
        )

        with st.spinner("Generating PPTX…"):
            export_pptx_from_structured_data_combined(
                template_path,
                bs_data,
                is_data,
                combined_output_path,
                entity_name,
                language="chinese" if language == "Chn" else "english",
                temp_path=_financials_workbook_path,
                selected_sheet=_financials_sheet_name,
                is_chinese_databook=(language == "Chn"),
                bs_is_results=session_state.get("bs_is_results"),
                model_type=session_state.get("model_type", "local"),
                model_name=session_state.get("model_name"),
                skip_summary_ai=False,
                pre_generated_summaries=session_state.get("section_summaries") or None,
                mappings=mappings,
            )
        if os.path.exists(combined_output_path):
            with open(combined_output_path, "rb") as handle:
                session_state.pptx_download_data = handle.read()
            session_state.pptx_download_filename = os.path.basename(combined_output_path)
            session_state.pptx_download_mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            session_state.pptx_ready = True

    except Exception as exc:
        st.error(f"❌ PPTX generation failed: {exc}")
        import traceback

        st.code(traceback.format_exc())


def batch_process_entity(
    *,
    temp_path: str,
    entity_name: str,
    selected_sheet: Optional[str] = None,
    financials_from: Optional[str] = None,
    financials_sheet: Optional[str] = None,
    mapping_overrides: Optional[Dict[str, str]] = None,
    model_type: str = "local",
    model_name: Optional[str] = None,
    language: Optional[str] = None,
    use_multithreading: bool = True,
    max_workers: Optional[int] = None,
    user_comments: Optional[Dict[str, str]] = None,
    template_path: Optional[str] = None,
    output_dir: str = "fdd_utils/output",
    output_filename: Optional[str] = None,
    progress_callback: Optional[Callable[..., None]] = None,
    on_data_ready: Optional[Callable[[Dict[str, Any]], None]] = None,
) -> Dict[str, Any]:
    """Headless, session_state-free equivalent of the single-file
    process -> reconcile -> AI -> export flow (render_ai_generation_section +
    generate_pptx_presentation above), for driving multiple entities in a
    batch loop. Takes explicit parameters instead of Streamlit session_state
    so it has no per-request UI-widget state to collide across entities, and
    reuses the exact same production functions so a batch run behaves
    identically to running each file through the UI one at a time. Mirrors
    inspect_databook.py's inspect_one() headless pattern.

    financials_from/financials_sheet point BS/IS extraction at a sibling
    roll-up ("主表") workbook's named sheet when this entity's own file has
    no Financials-pattern sheet of its own — same mechanism
    process_workbook_data already exposes for the single-file flow.

    on_data_ready, if given, fires once (right after data extraction +
    reconciliation complete, before AI generation starts) with a summary
    dict ({"entity_name", "accounts_total", "accounts_matched",
    "bs_match_counts", "is_match_counts"}) — lets a batch UI show what was
    extracted/reconciled for this entity while its own AI generation (and
    any later entities' processing) is still running, rather than only
    ever surfacing data once the entire entity is fully done.

    output_filename, if given, overrides the sanitized entity_name as the
    output PPTX's base filename (still timestamped) — lets a batch UI
    offer a per-entity filename field (e.g. suggested from the roll-up
    sheet name) distinct from the entity_name baked into the report text.
    """
    from .ai import run_ai_pipeline_with_progress
    from .pptx import export_pptx_from_structured_data_combined
    from .workbook import process_workbook_data

    result: Dict[str, Any] = {"entity_name": entity_name, "status": "ok"}

    try:
        state = process_workbook_data(
            temp_path=temp_path,
            entity_name=entity_name,
            selected_sheet=selected_sheet,
            mapping_overrides=mapping_overrides,
            financials_from=financials_from,
            financials_sheet=financials_sheet,
        )
    except Exception as exc:
        result["status"] = "failed"
        result["error"] = f"Processing failed: {exc}"
        return result

    dfs = state.get("dfs") or {}
    if not dfs:
        result["status"] = "failed"
        result["error"] = "No schedule tabs could be extracted from this databook."
        return result

    reconciliation = state.get("reconciliation")
    resolution = state.get("resolution")
    mappings = get_effective_mappings(load_mappings(), resolution)

    # Raw process_workbook_data language is "Eng"/"Chi" (workbook.py's own
    # detection convention); normalise to the UI's "Eng"/"Chn" convention so
    # this matches every == "Chn" check generate_pptx_presentation makes,
    # unless the caller already passed an explicit override in that form.
    if language:
        effective_language = language
    else:
        raw_language = str(state.get("language") or "Eng").strip()
        effective_language = "Chn" if raw_language in ("Chi", "Chn", "chinese", "Chinese") else "Eng"

    statement_mode = detect_statement_mode(reconciliation)
    if statement_mode in ("is_only", "bs_only"):
        target_type = "IS" if statement_mode == "is_only" else "BS"
        matched_mapping_keys = [
            k for k in dfs
            if mappings.get(find_mapping_key(k, mappings) or k, {}).get("type") == target_type
        ]
        if not matched_mapping_keys:
            matched_mapping_keys = list(dfs.keys())
    else:
        matched_mapping_keys = derive_reconciliation_matched_keys(reconciliation, dfs.keys(), resolution)
        has_reconciliation_data = bool(
            reconciliation and any(recon_df is not None and not recon_df.empty for recon_df in reconciliation)
        )
        if not has_reconciliation_data:
            matched_mapping_keys = list(dfs.keys())

    if not matched_mapping_keys:
        result["status"] = "failed"
        result["error"] = "No eligible accounts after reconciliation filtering."
        return result

    if on_data_ready:
        bs_recon, is_recon = (list(reconciliation) + [None, None])[:2] if reconciliation else (None, None)
        try:
            on_data_ready({
                "entity_name": entity_name,
                "accounts_total": len(dfs),
                "accounts_matched": len(matched_mapping_keys),
                "bs_match_counts": bs_recon["Match"].value_counts().to_dict() if bs_recon is not None and not bs_recon.empty else {},
                "is_match_counts": is_recon["Match"].value_counts().to_dict() if is_recon is not None and not is_recon.empty else {},
                # Raw per-account reconciliation breakdowns (same DataFrames
                # render_reconciliation_section uses in the interactive
                # single-file flow) -- so a caller can show the actual
                # account-by-account table, not just the match-status
                # counts, while this entity's own AI generation (and any
                # later entities) are still running.
                "bs_recon_df": bs_recon,
                "is_recon_df": is_recon,
                # Full session_state-shaped (minus ai_results/pptx) partial
                # bundle -- lets a caller swap this into st.session_state
                # and call render_data_tables_section() for the complete
                # per-account breakdown view (cash, investment properties,
                # etc., not just reconciliation), the same rich view a
                # fully-finished entity gets, while AI is still running.
                "state": {
                    "dfs": dfs,
                    "display_dfs": state.get("display_dfs"),
                    "workbook_list": state.get("workbook_list"),
                    "display_workbook_list": state.get("display_workbook_list"),
                    "language": effective_language,
                    "bs_is_results": state.get("bs_is_results"),
                    "reconciliation": reconciliation,
                    "resolution": resolution,
                    "entity_name": entity_name,
                },
            })
        except Exception:
            pass  # a UI-side display glitch should never abort the pipeline

    ai_results = run_ai_pipeline_with_progress(
        mapping_keys=matched_mapping_keys,
        dfs=dfs,
        model_type=model_type,
        model_name=model_name,
        language=effective_language,
        use_multithreading=use_multithreading,
        max_workers=max_workers,
        progress_callback=progress_callback,
        user_comments=user_comments or {},
    )

    structured_payloads = build_pptx_structured_payloads(
        ai_results=ai_results,
        mappings=mappings,
        bs_is_results=state.get("bs_is_results"),
        dfs=dfs,
    )
    bs_data = structured_payloads.get("BS", [])
    is_data = structured_payloads.get("IS", [])
    if not bs_data and not is_data:
        result["status"] = "failed"
        result["error"] = "No content generated for PPTX (empty BS and IS payloads)."
        return result

    resolved_template_path = template_path
    if not resolved_template_path:
        for candidate in ["fdd_utils/template.pptx", "template.pptx"]:
            if os.path.exists(candidate):
                resolved_template_path = candidate
                break
    if not resolved_template_path:
        result["status"] = "failed"
        result["error"] = "PowerPoint template not found (fdd_utils/template.pptx)."
        return result

    os.makedirs(output_dir, exist_ok=True)
    timestamp = dt_module.datetime.now().strftime("%Y%m%d_%H%M%S")
    sanitized_entity = re.sub(r"[^\w\-_]", "_", str(output_filename or entity_name)).strip("_") or "Entity"
    output_path = os.path.join(output_dir, f"{sanitized_entity}_{timestamp}.pptx")

    # embed_financial_tables reads (temp_path, selected_sheet) for its
    # currency-unit-label detection and as a fresh-extraction fallback --
    # when financials came from a roll-up workbook, that source (not this
    # entity's own file/blank sheet) is the one that actually holds them.
    # Mirrors process_workbook_data's own precedence (financials_from or
    # temp_path).
    financials_workbook_path = financials_from or temp_path
    financials_sheet_name = financials_sheet if financials_from else selected_sheet

    export_pptx_from_structured_data_combined(
        resolved_template_path,
        bs_data,
        is_data,
        output_path,
        entity_name,
        language="chinese" if effective_language == "Chn" else "english",
        temp_path=financials_workbook_path,
        selected_sheet=financials_sheet_name,
        is_chinese_databook=(effective_language == "Chn"),
        bs_is_results=state.get("bs_is_results"),
        model_type=model_type,
        model_name=model_name,
        skip_summary_ai=False,
        mappings=mappings,
    )

    result["output_path"] = output_path
    result["bs_count"] = len(bs_data)
    result["is_count"] = len(is_data)
    result["accounts_processed"] = len(matched_mapping_keys)

    with open(output_path, "rb") as handle:
        pptx_bytes = handle.read()

    # Full session_state-shaped bundle so a caller (the batch UI) can swap
    # this entity's results into st.session_state and reuse the single-file
    # render_processed_view/generate_pptx_presentation UI UNCHANGED, instead
    # of only ever seeing this thin status dict.
    result["state"] = {
        "dfs": state.get("dfs"),
        "display_dfs": state.get("display_dfs"),
        "dfs_variants": state.get("dfs_variants"),
        "display_df_variants": state.get("display_df_variants"),
        "workbook_list": state.get("workbook_list"),
        "display_workbook_list": state.get("display_workbook_list"),
        "language": effective_language,
        "detected_language": effective_language,
        "bs_is_results": state.get("bs_is_results"),
        "reconciliation": reconciliation,
        "resolution": resolution,
        "project_name": state.get("project_name"),
        "entity_name": entity_name,
        # The financials source (not necessarily this entity's own file --
        # see financials_workbook_path/financials_sheet_name above), so a
        # later "Regenerate PPTX" click from within the reused single-file
        # UI still finds the right sheet for the embedded table instead of
        # re-hitting the same blank-selected_sheet bug this export call
        # just worked around.
        "temp_path": financials_workbook_path,
        "selected_sheet": financials_sheet_name,
        "mapping_overrides": mapping_overrides,
        "ai_results": ai_results,
        "model_type": model_type,
        "model_name": model_name,
        "use_multithreading": use_multithreading,
        "pptx_ready": True,
        "pptx_download_data": pptx_bytes,
        "pptx_download_filename": os.path.basename(output_path),
        "pptx_download_mime": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    }
    return result
# --- end ui/pptx_export.py ---

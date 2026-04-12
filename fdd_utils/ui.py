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
    "bs_is_results": None,
    "ai_results": None,
    "reconciliation": None,
    "resolution": None,
    "model_type": "local",
    "project_name": None,
    "last_run_folder": None,
    "entity_name": None,
    "pptx_download_trigger": None,
    "button_click_counter": 0,
    "pptx_ready": False,
    "temp_path": None,
    "selected_sheet": None,
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
    "bs_is_results",
    "ai_results",
    "reconciliation",
    "resolution",
    "entity_name",
    "project_name",
    "pptx_ready",
    "mapping_overrides",
    "account_comments",
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
from typing import Any, Dict, Iterable, List

import pandas as pd

from .financial_common import (
    build_income_statement_period_label,
    clean_english_placeholders,
    dedupe_non_empty,
    visible_descriptions,
)
from .workbook import INTERNAL_ROW_KEY, split_accounts_by_type as shared_split_accounts_by_type


def build_entity_selector_model(entity_options: List[str], current_entity_name: str) -> Dict[str, Any]:
    dropdown_options = dedupe_non_empty(entity_options)
    text_value = str(current_entity_name or "").strip()
    if not text_value and len(dropdown_options) == 1:
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
    bs_accounts, is_accounts, other_accounts = shared_split_accounts_by_type(display_account_keys, mappings, dfs=dfs)
    return {
        "tab_names": ["BS Recon", "IS Recon", "BS", "IS", "Schedule Mapping"],
        "bs_accounts": bs_accounts,
        "is_accounts": is_accounts,
        "reconciliation_accounts": other_accounts,
    }


def derive_reconciliation_matched_keys(
    reconciliation: tuple[pd.DataFrame | None, pd.DataFrame | None] | None,
    available_keys: Iterable[str],
    resolution: Dict[str, Any] | None = None,
) -> List[str]:
    available_key_order = dedupe_non_empty(available_keys)
    available_key_set = set(available_key_order)
    if not reconciliation or not available_key_order:
        return []

    included_statuses = {"✅ Match", "⚠️ Match", "✅ Immaterial", "❌ Diff"}
    matched_keys: List[str] = []
    seen = set()

    for recon_df in reconciliation:
        if recon_df is None or recon_df.empty:
            continue
        if "Match" not in recon_df.columns or "Tab_Account" not in recon_df.columns:
            continue

        filtered = recon_df[recon_df["Match"].isin(included_statuses)]
        for value in filtered["Tab_Account"].tolist():
            key = str(value or "").strip()
            if not key or key == "-" or key in seen or key not in available_key_set:
                continue
            seen.add(key)
            matched_keys.append(key)

    for dynamic_key in dedupe_non_empty((resolution or {}).get("dynamic_mappings", {}).keys()):
        if dynamic_key in available_key_set and dynamic_key not in seen:
            seen.add(dynamic_key)
            matched_keys.append(dynamic_key)

    # Collect ALL Financials_Account names that reconciliation already evaluated
    # (regardless of match status) so the resolved-map fallback only adds
    # accounts that are truly absent from the Financials summary.
    evaluated_accounts: set = set()
    for recon_df in reconciliation:
        if recon_df is None or recon_df.empty or "Financials_Account" not in recon_df.columns:
            continue
        for fa in recon_df["Financials_Account"].tolist():
            evaluated_accounts.add(str(fa).strip().lower())

    # Include resolved accounts that have DFS data but are absent from the
    # Financials summary (and therefore have no reconciliation row at all).
    resolved_map = (resolution or {}).get("resolved", {})
    for _mapping_key, resolved_info in resolved_map.items():
        sheet_name = (resolved_info or {}).get("sheet_name", "")
        if not sheet_name or sheet_name not in available_key_set or sheet_name in seen:
            continue
        # Check if reconciliation already evaluated this account
        sn_lower = sheet_name.strip().lower()
        already_evaluated = any(
            sn_lower in ea or ea in sn_lower for ea in evaluated_accounts
        )
        if already_evaluated:
            continue
        seen.add(sheet_name)
        matched_keys.append(sheet_name)

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
    build_highlighted_commentary_html,
    get_prompt_engine,
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


def render_ai_generation_section(session_state: Any, get_model_display_name) -> None:
    ai_key = f"ai_btn_{session_state.button_click_counter}"
    generate_clicked = st.button("▶️ Generate AI Content", type="primary", use_container_width=True, key=ai_key)
    if generate_clicked:
        session_state.button_click_counter += 1

    progress_container = st.container()
    if not generate_clicked:
        return

    with progress_container:
        st.markdown("### 🔄 AI Processing Progress")
        progress_placeholder = st.empty()
        status_placeholder = st.empty()
        try:
            results = None
            selected_pipeline_dfs = build_selected_pipeline_dfs(session_state)
            reconciliation = session_state.get("reconciliation")
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
            else:
                total_items = len(matched_mapping_keys)
                total_steps = 4 * total_items

                def update_progress(agent_num, agent_name, item_num, total_items_in_agent, completed_items, key_name=None):
                    if agent_num > 4:
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
                                f"🔄 Running Subagent {agent_num}/4: {agent_name} | Processing item {item_num}/{total_items_in_agent}{key_display} | Progress: {completed_steps}/{total_steps} steps | ETA: {int(eta_seconds / 60)}m {int(eta_seconds % 60)}s"
                            )
                        else:
                            status_placeholder.info(
                                f"🔄 Running Subagent {agent_num}/4: {agent_name} | Processing item {item_num}/{total_items_in_agent}{key_display} | Progress: {completed_steps}/{total_steps} steps | ETA: Calculating..."
                            )
                    else:
                        update_progress.start_time = time.time()
                        status_placeholder.info(
                            f"🔄 Running Subagent {agent_num}/4: {agent_name} | Processing item {item_num}/{total_items_in_agent}{key_display} | Progress: {completed_steps}/{total_steps} steps | ETA: Calculating..."
                        )

                start_time = time.time()
                update_progress.start_time = start_time
                status_placeholder.info(f"🚀 Starting AI pipeline for {total_items} accounts... | Progress: 0/{total_steps} steps | ETA: Calculating...")
                progress_placeholder.progress(0)
                results = run_ai_pipeline_with_progress(
                    mapping_keys=matched_mapping_keys,
                    dfs=selected_pipeline_dfs,
                    model_type=session_state.get("model_type", "local"),
                    language=session_state.language,
                    use_multithreading=True,
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
                    status_placeholder.success(
                        f"✅ AI content generated successfully! {len(results)} accounts processed in {int(time.time() - start_time)}s."
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

    st.markdown("### 📝 Generated Content")
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
    def _render_commentary_fragment(
        key: str,
        prefix: str,
        detailed_content: str,
        simplified_content: str,
        has_simplified: bool,
        clause_reviews: list,
        final_label: str,
        selected_df,
        language: str,
    ):
        """Fragment that re-renders only itself when the Detailed/Simplified toggle changes."""
        commentary_mode_key = f"{prefix}{key}_commentary_mode"
        if has_simplified:
            mode = st.radio(
                "Commentary version",
                options=["Detailed", "Simplified"],
                index=0,
                horizontal=True,
                key=commentary_mode_key,
                label_visibility="collapsed",
            )
        else:
            mode = "Detailed"
        # Store the selection for PPTX export
        if "commentary_modes" not in session_state:
            session_state.commentary_modes = {}
        session_state.commentary_modes[key] = mode

        final_content = simplified_content if mode == "Simplified" else detailed_content
        render_account_remarks_context(selected_df, key, language, prefix=f"{prefix}generated_")
        st.markdown(build_highlighted_commentary_html(str(final_content), clause_reviews or [] if mode == "Detailed" else []), unsafe_allow_html=True)
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
        with st.expander("Plain text", expanded=False):
            st.text_area(
                label="Final content",
                value=str(final_content),
                height=min(max(100, int(len(str(final_content)) / 3)), 600),
                key=f"{prefix}{key}_final_display",
                label_visibility="collapsed",
            )

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
                simplified_content = str(result.get("final_simplified") or "").strip()
                has_simplified = bool(simplified_content)
                validator_metadata = extract_validator_metadata(result)
                clause_reviews = validator_metadata.get("clause_reviews") if isinstance(validator_metadata, dict) else []
                has_final = has_meaningful_result_text(detailed_content)
                reprompt_mode = str(result.get("reprompt_mode") or "").strip()
                final_label = "Final (Reprompt + validator)" if reprompt_mode == "generator_reprompt_validated" else ("Final (Generator reprompt)" if reprompt_mode == "generator_only" else "Final (Validator)")
                if has_final:
                    st.markdown(f"#### ✨ {final_label}")
                    _render_commentary_fragment(
                        key=key,
                        prefix=prefix,
                        detailed_content=str(detailed_content),
                        simplified_content=simplified_content,
                        has_simplified=has_simplified,
                        clause_reviews=clause_reviews,
                        final_label=final_label,
                        selected_df=selected_df,
                        language=session_state.get("language", "Eng"),
                    )
                    st.markdown("---")
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
                            language=session_state.language,
                            user_comments={key: reprompt_comment},
                        )
                        merged_results = dict(session_state.ai_results or {})
                        merged_results.update(updated_results)
                        session_state.ai_results = merged_results
                        session_state.pptx_ready = False
                    st.rerun()
                st.markdown("---")
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
                if not has_final and not agent_contents:
                    st.warning("No agent outputs available for this account")

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
    col1, col2, col3, col4, col_reports = st.columns([1, 1, 1, 1, 1])
    with col1:
        st.metric("✅ Matches", recon_df["Match"].isin(["✅ Match", "⚠️ Match"]).sum())
    with col2:
        st.metric("❌ Differences", (recon_df["Match"] == "❌ Diff").sum())
    with col3:
        st.metric("⚠️ Not Found", (recon_df["Match"] == "⚠️ Not Found").sum())
    with col4:
        st.metric("✅ Immaterial", (recon_df["Match"] == "✅ Immaterial").sum())
    with col_reports:
        st.metric("Checked Rows", len(recon_df), help="Total rows checked in this reconciliation view")


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
    label = str(recon_df.columns[1]) if len(recon_df.columns) > 1 else "current period"
    st.caption(describe_statement_period(statement_type, label))
    if hidden_zero_rows:
        st.caption(f"Hidden {hidden_zero_rows} zero-value reconciliation row(s).")
    if warning_row_count:
        st.caption(f"{warning_row_count} row(s) have mapping/tab coverage warnings shown inline in `Match`.")
    st.dataframe(format_dataframe_for_display(display_recon_df), use_container_width=True, height=320)
    render_reconciliation_metrics(display_recon_df)


def render_reconciliation_section(
    recon_df: pd.DataFrame | None,
    statement_type: str,
    empty_message: str,
) -> None:
    _render_single_reconciliation_tab(recon_df, statement_type, empty_message)


def render_schedule_mapping_section(
    reconciliation_accounts: list[str],
    hidden_zero_reconciliation_accounts: list[str],
    mappings: Dict[str, Any],
) -> None:
    if not reconciliation_accounts and not hidden_zero_reconciliation_accounts:
        st.info("No schedule mapping coverage exceptions were identified.")
        return

    st.markdown("**Schedule Mapping Coverage**")
    if reconciliation_accounts:
        diagnostics_df = build_account_mapping_diagnostics(reconciliation_accounts, mappings)
        diagnostics_df = diagnostics_df[diagnostics_df["classification"].astype(str).eq("other")].copy()
        if not diagnostics_df.empty:
            diagnostics_df = diagnostics_df.rename(
                columns={
                    "account_name": "Schedule_Tab",
                    "mapping_key": "Mapping_Key",
                    "account_type": "Account_Type",
                    "classification": "Classification",
                    "reason": "Coverage_Note",
                }
            )
            st.caption(
                "These schedule tabs remain outside the standard BS/IS mapping flow after the "
                "reconciliation view was flattened."
            )
            st.dataframe(format_dataframe_for_display(diagnostics_df), use_container_width=True, hide_index=True)

    if hidden_zero_reconciliation_accounts:
        st.caption("Zero-value non-BS/IS schedule tab(s): " + ", ".join(hidden_zero_reconciliation_accounts))


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


def render_processed_view(session_state: Any, generate_pptx_callback, get_model_display_name) -> None:
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
    reconciliation_accounts = list(display_groups["reconciliation_accounts"])
    hidden_zero_reconciliation_accounts = [
        account_name
        for account_name in reconciliation_accounts
        if not account_has_non_zero_values(account_display_dfs.get(account_name))
    ]

    recon_bs = session_state.reconciliation[0] if session_state.reconciliation else None
    recon_is = session_state.reconciliation[1] if session_state.reconciliation else None
    resolution = session_state.get("resolution") or {}
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

    def render_account_tabs(account_names: List[str], group_type: str):
        if not account_names:
            st.info(f"No {group_type} items available.")
            return
        item_tabs = st.tabs(account_names)
        for index, key in enumerate(account_names):
            with item_tabs[index]:
                render_account_panel(key)

    data_tabs = st.tabs(display_groups["tab_names"])
    with data_tabs[0]:
        render_reconciliation_section(
            recon_df=recon_bs,
            statement_type="BS",
            empty_message="No Balance Sheet reconciliation data available.",
        )
    with data_tabs[1]:
        render_reconciliation_section(
            recon_df=recon_is,
            statement_type="IS",
            empty_message="No Income Statement reconciliation data available.",
        )
    with data_tabs[2]:
        render_account_tabs(bs_accounts, "BS")
    with data_tabs[3]:
        render_account_tabs(is_accounts, "IS")
    with data_tabs[4]:
        render_schedule_mapping_section(
            reconciliation_accounts=reconciliation_accounts,
            hidden_zero_reconciliation_accounts=hidden_zero_reconciliation_accounts,
            mappings=mappings,
        )

    st.markdown("---")
    col_header, col_pptx, col_download = st.columns([3, 1, 0.3])
    with col_header:
        st.header("🤖 AI Content Generation")
        if session_state.get("pptx_ready", False) and session_state.get("pptx_download_filename"):
            st.caption(f"PPTX ready: {session_state.pptx_download_filename}")
        else:
            st.caption("Generate AI content first, then export the current result set to PPTX.")
    with col_pptx:
        st.markdown("<br>", unsafe_allow_html=True)
        pptx_key = f"pptx_btn_{session_state.button_click_counter}"
        if st.button("📄 Generate & Export PPTX", type="secondary", use_container_width=True, disabled=session_state.get("ai_results") is None, key=pptx_key):
            session_state.button_click_counter += 1
            generate_pptx_callback()
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
) -> str:
    digest = hashlib.sha256(uploaded_bytes).hexdigest()[:16]
    target_dir = Path(cache_dir or tempfile.gettempdir()) / "python-pptx-uploads"
    target_dir.mkdir(parents=True, exist_ok=True)

    target_path = target_dir / f"{_safe_stem(uploaded_name)}-{digest}.xlsx"
    if not target_path.exists():
        target_path.write_bytes(uploaded_bytes)

    session_state["temp_path"] = str(target_path)
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


def render_sidebar_upload(session_state: Any, get_model_display_name: Callable[[str], str]) -> str | None:
    with st.sidebar:
        if "model_type" not in session_state:
            session_state.model_type = "local"
        st.caption(f"🤖 AI Mode: {get_model_display_name(session_state.model_type)}")
        st.markdown("**📁 Databook File**")
        uploaded_file = st.file_uploader(
            "Upload Excel file",
            type=["xlsx", "xls"],
            help="Upload your financial databook",
            key="file_uploader",
        )

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


def render_refresh_control(session_state: Any) -> None:
    _col_spacer, col_refresh = st.columns([12, 1])
    with col_refresh:
        if st.button("🔄", help="Refresh page and reset", use_container_width=True, key="refresh_main"):
            clear_workbook_caches()
            reset_processing_session_state(session_state, clear_upload_reference=True)
            st.rerun()
# --- end ui/sidebar.py ---

# --- begin ui/pptx_export.py ---
import datetime as dt_module
import logging
import os
import re
import time
from typing import Any

import streamlit as st

from .pptx import build_pptx_structured_payloads
from .workbook import get_effective_mappings, load_mappings

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
            commentary_modes=session_state.get("commentary_modes"),
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

        export_pptx_from_structured_data_combined(
            template_path,
            bs_data,
            is_data,
            combined_output_path,
            entity_name,
            language="chinese" if language == "Chn" else "english",
            temp_path=session_state.get("temp_path"),
            selected_sheet=session_state.get("selected_sheet"),
            is_chinese_databook=(language == "Chn"),
            bs_is_results=session_state.get("bs_is_results"),
            model_type=session_state.get("model_type", "local"),
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
# --- end ui/pptx_export.py ---

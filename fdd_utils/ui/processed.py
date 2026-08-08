from __future__ import annotations

# re-added: bound by an import in another section of the pre-split module
from typing import Any, Callable, Dict, Iterable, List, Optional


from .views import build_account_display_dataframe, build_processed_display_groups, describe_statement_period
from .ai_panel import effective_mappings_from_session, format_dataframe_for_display, get_account_dataframe, get_financial_account_options, render_account_remarks_context, render_ai_generation_section, render_generated_content
from typing import Any, List

import pandas as pd
import streamlit as st

from ..workbook import find_mapping_key, load_mappings


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
    show_download_button: bool = True,
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
        if show_download_button and session_state.get("pptx_ready", False) and "pptx_download_data" in session_state:
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

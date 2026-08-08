from __future__ import annotations

# re-added: bound by an import in another section of the pre-split module
from pathlib import Path


from .views import derive_reconciliation_matched_keys, detect_statement_mode
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

from ..ai import (
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
from ..financial_common import extract_result_text_content, get_pipeline_result_text
from ..financial_display_format import prepare_display_dataframe, stringify_display_dataframe
from ..workbook import (
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
    from .pptx_export import logger  # local: breaks the ai_panel<->pptx_export import cycle
    import json as _json
    import math as _math
    demo_path = Path(__file__).parent / "demo_results.json"
    try:
        all_results: dict = _json.loads(demo_path.read_text(encoding="utf-8"))
    except Exception as exc:
        logger.warning("Could not load demo_results.json: %s", exc)
        all_results = {}

    # Pipeline is now 3 stages: Generator → Auditor (Polish) → Validator
    from ..ai import SUBAGENT_SEQUENCE
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
    from .pptx_export import generate_pptx_presentation, logger  # local: breaks the ai_panel<->pptx_export import cycle
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
                    dfs=selected_pipeline_dfs,
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
                from ..ai import SUBAGENT_SEQUENCE
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
                        from ..pptx import PowerPointGenerator
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
                            # A page/section summary needs each account's
                            # lead-in theme only, never a table account's
                            # per-component "-"/"➢" detail bullets after
                            # "明细如下：" -- see strip_table_detail_for_
                            # summary's own docstring for the real
                            # corrupted-coSummaryShape bug this prevents.
                            text = PowerPointGenerator.strip_table_detail_for_summary(
                                text, session_state.language == "Chn",
                            )
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
                                    from ..ai import _PIPELINE_BREAKER
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

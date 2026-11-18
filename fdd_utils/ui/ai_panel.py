"""The commentary the agent wrote, in the app: the run that writes it, and the tabs that show it and take a request to write one account again.

The run itself is fdd_utils/agent/book.py (driven by ui/agent_run.py); this file shows its progress and files what it returns in
session_state, where the processed view, the deck export and the batch view read it. The commentary of an account is the agent's
paragraph in the shape the deck reads (claims.to_legacy_result): the sentences, and for each one the verifier's reading of it, which the
tabs paint red (a figure the databook does not support, or a sentence the agent read and found wrong) or orange (refused only for how its
prose was read).
"""
from __future__ import annotations

import time
import traceback
from pathlib import Path
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

from ..ai import get_default_config_path, load_yaml_config
from ..ai.validator import build_highlighted_commentary_html
from ..financial_common import extract_result_text_content, get_pipeline_result_text
from ..financial_display_format import prepare_display_dataframe, stringify_display_dataframe
from ..workbook import (
    find_mapping_key,
    get_effective_mappings,
    load_mappings,
    split_accounts_by_type,
)
from ..workbook import remarks as _remarks
from .views import detect_statement_mode, explain_reconciliation_selection

_EMPTY_RESULT_MARKERS = {"none", "null", "nan", "", "n/a", "na"}


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
    """Put the commentary's text under "final" where the tabs read it (a result may carry it only inside a nested stage output)."""
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
    """The verifier's reading of each sentence (clause_reviews) as the commentary tab paints it."""
    validation = result_dict.get("agent_4_validation", {})
    if isinstance(validation, dict) and isinstance(validation.get("clause_reviews"), list):
        return {
            "final_content": str(validation.get("final_content") or extract_result_text(result_dict, "final")).strip(),
            "clause_reviews": validation.get("clause_reviews", []),
            "raw_response": str(validation.get("raw_response") or validation.get("output") or ""),
        }
    return {"clause_reviews": [], "raw_response": ""}


def extract_account_remarks_context(df: pd.DataFrame | None, language: str) -> Dict[str, Any]:
    if not isinstance(df, pd.DataFrame):
        return {}
    supporting_notes = [str(note).strip() for note in (df.attrs.get("supporting_notes") or []) if str(note).strip()]
    rhs_rows = _remarks.filter_adjacent_detail_rows(df)
    rhs_summary = _remarks.summarize_rhs_remarks(rhs_rows, language)
    table_linked = _remarks.table_linked_remarks(df)
    return {
        "supporting_notes": supporting_notes,
        "rhs_rows": rhs_rows,
        "rhs_summary": rhs_summary,
        "table_linked_remarks": table_linked,
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
    # The tracked file is fdd_utils/demo_results.json, one level up from this
    # package — `Path(__file__).parent` pointed at fdd_utils/ui/ and the miss
    # was swallowed, so demo mode silently replayed nothing. Log the resolved
    # path on failure so the next such breakage is visible.
    demo_path = Path(__file__).resolve().parent.parent / "demo_results.json"
    try:
        all_results: dict = _json.loads(demo_path.read_text(encoding="utf-8"))
    except Exception as exc:
        logger.warning("Could not load demo results from %s: %s", demo_path, exc)
        all_results = {}

    # the simulated run shows the three stages the recorded results came from
    agents = ["Generator", "Auditor", "Validator"]
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
        for canonical, entry in (load_mappings() or {}).items():
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


def _matched_keys(session_state: Any) -> list:
    """The accounts the deck covers (the reconciliation selection), for the replayed demo; a real run takes them from its session."""
    selected = build_selected_pipeline_dfs(session_state)
    reconciliation = session_state.get("reconciliation")
    statement_mode = detect_statement_mode(reconciliation)
    if statement_mode in ("is_only", "bs_only"):
        target_type = "IS" if statement_mode == "is_only" else "BS"
        mappings = effective_mappings_from_session(session_state)
        keys = [k for k in selected if mappings.get(find_mapping_key(k, mappings) or k, {}).get("type") == target_type]
        session_state.reconciliation_selection = None
        return keys or list(selected)
    selection = explain_reconciliation_selection(reconciliation, selected.keys(), session_state.get("resolution"), dfs=selected)
    keys = selection.admitted_keys
    if not bool(reconciliation and any(df is not None and not df.empty for df in reconciliation)):
        from ..workbook.selection import selection_without_financials
        keys = list(selected)
        selection = selection_without_financials(keys)
    session_state.reconciliation_selection = selection
    return keys


def _demo_settings() -> Dict[str, Any]:
    """config.yml's `demo` block (the uploaded file the app replays recorded results for), or {} when there is no config or no such block.
    Read as plain YAML: the replay needs no provider, so a machine with none configured can still show it."""
    try:
        return (load_yaml_config(get_default_config_path()) or {}).get("demo") or {}
    except Exception:                                                        # noqa: BLE001
        return {}


def _eta(started: float, done: int, total: int) -> str:
    if done <= 0 or total <= 0:
        return "Calculating..."
    left = (time.time() - started) / done * (total - done)
    return f"{int(left // 60)}m {int(left % 60)}s"


def render_ai_generation_section(session_state: Any, get_model_display_name) -> None:
    # Auto-trigger: write the commentary automatically when data is loaded but there is none yet.
    # No button needed; the Redo button in the export header starts it again.
    from .pptx_export import generate_pptx_presentation
    if not (session_state.get("dfs") is not None and session_state.get("ai_results") is None):
        return
    if session_state.get("book_error"):
        # a run that failed is not started again by the next click on the page: every attempt is paid for
        st.error(f"❌ The last run did not produce commentary: {session_state.book_error}")
        if st.button("Try again", key="book_retry"):
            session_state.book_error = None
            st.rerun()
        return

    session_state.pptx_ready = False
    for key in ("pptx_download_data", "section_summaries", "agent_work", "agent_last_turn", "agent_last_export"):
        session_state.pop(key, None)

    progress_container = st.container()
    with progress_container:
        st.markdown("### 🔄 AI Processing Progress")
        progress_placeholder = st.empty()
        status_placeholder = st.empty()
        try:
            started = time.time()
            demo_cfg = _demo_settings()
            demo_filename = str(demo_cfg.get("filename") or "").strip()
            is_demo = bool(demo_filename and str(session_state.get("uploaded_filename") or "").strip() == demo_filename)
            if is_demo:
                # the configured demo file replays recorded results: the whole flow offline, in a fixed time, with no model
                matched = _matched_keys(session_state)
                duration = int(demo_cfg.get("progress_duration_seconds") or 25)
                status_placeholder.info(f"🎬 Demo mode — replaying pre-recorded results for {len(matched)} accounts ({duration}s simulated run)…")
                progress_placeholder.progress(0)
                results = _run_demo_ai(matched, duration, progress_placeholder, status_placeholder)
                summaries = {name: str(results.pop(f"__{name}_summary__", "") or "").strip() for name in ("BS", "IS")}
                session_state.ai_results = results
                session_state.section_summaries = {k: v for k, v in summaries.items() if v}
                generate_pptx_presentation(session_state=session_state, pptx_available=True)
                status_placeholder.success(f"✅ Demo commentary + PPTX ready! {len(matched)} accounts.")
            else:
                from .. import usage_ledger
                from . import agent_run
                usage_ledger.record_event("ai_run", mode="single", model=session_state.get("model_name") or session_state.get("model_type"))
                status_placeholder.info("🚀 Reading the workbook and starting the agent…")
                progress_placeholder.progress(0)

                def progress(done: int, total: int, account: str) -> None:
                    progress_placeholder.progress(min(done / total, 1.0) if total else 1.0)
                    status_placeholder.info(f"🔄 Account {done}/{total} written: {account} | ETA: {_eta(started, done, total)}")

                result = agent_run.run_for_state(session_state, progress_callback=progress)
                if not result.ai_results:
                    session_state.book_error = (result.stopped or result.why_no_text
                                                or "no account could be written; check that the model is configured "
                                                   f"({get_model_display_name(session_state.get('model_type', 'local'))}) and reachable")
                    progress_placeholder.empty()
                    st.rerun()
                session_state.update(agent_run.book_state(result))
                kept = len(result.finished)
                note = (f" {result.issue_cases} account(s) have a sentence the verifier refused (kept as cases)." if result.issue_cases else "")
                if result.stopped:
                    status_placeholder.warning(f"⚠️ The run stopped early: {result.stopped}")
                else:
                    status_placeholder.success(f"✅ {kept} of {len(result.accounts)} accounts written and verified in {int(time.time() - started)}s."
                                               + note + (" The deck is ready." if session_state.get("pptx_ready") else " The deck could not be written: use Generate PPTX."))
            progress_placeholder.progress(1.0)
            st.rerun()
        except Exception as exc:
            if type(exc).__name__ in ("RerunException", "StopException"):    # st.rerun() is an exception Streamlit catches
                raise
            progress_placeholder.empty()
            session_state.book_error = f"{type(exc).__name__}: {exc}"
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
        st.markdown("- The model service is unreachable or the account is not configured")
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
        """Render the commentary and the verifier's reading of each sentence for a single account."""
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
            caption_parts = [f"The verifier read {len(clause_reviews)} sentence(s)"]
            if hallucination_count:
                caption_parts.append(f"{hallucination_count} red (a figure the databook does not support, or a sentence the agent found wrong)")
            if reasoning_count:
                caption_parts.append(f"{reasoning_count} orange (refused only for how its prose was read)")
            if not flagged_count:
                caption_parts.append("all data-backed")
            st.caption("; ".join(caption_parts) + ".")
            with st.expander("Verifier's reading of each sentence", expanded=False):
                review_rows = [
                    {
                        "Sentence": str(review.get("clause") or ""),
                        "Category": str(review.get("category") or "data-backed").replace("-", " ").title(),
                        "Supported": "Yes" if bool(review.get("supported")) else "No",
                        "Reason": str(review.get("reason") or ""),
                    }
                    for review in clause_reviews
                    if isinstance(review, dict)
                ]
                if review_rows:
                    st.dataframe(pd.DataFrame(review_rows), use_container_width=True, hide_index=True)

    def create_account_tabs(keys, prefix=""):
        account_tabs = st.tabs([f"📄 {key}" for key in keys])
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
                if has_final:
                    st.markdown("#### ✨ Commentary")
                    _render_commentary_fragment(detailed_content=str(detailed_content), clause_reviews=clause_reviews)
                    render_account_remarks_context(selected_df, key, session_state.get("language", "Eng"), prefix=f"{prefix}generated_")
                else:
                    st.warning("No commentary was written for this account")

                work = session_state.get("agent_work")
                if work is None:
                    continue
                st.markdown("---")
                with st.expander(f"✏️ Write {key} again", expanded=False):
                    guidance = st.text_area(
                        label=f"Guidance for {key}",
                        value=session_state.account_comments.get(key, ""),
                        placeholder="Say what to change for this account only, then click the button. Leave it empty to write it again from the evidence.",
                        key=f"{prefix}{key}_reprompt_comment",
                        height=90,
                    )
                    session_state.account_comments[key] = guidance
                    if st.button(f"Write {key} again", key=f"{prefix}{key}_reprompt_button", use_container_width=True):
                        from .. import usage_ledger
                        from ..agent import book as B
                        usage_ledger.record_event("reprompt", account=key)
                        with st.spinner(f"Writing {key} again..."):
                            rewritten = B.redraft(work, key, guidance)
                            if rewritten:
                                merged_results = dict(session_state.ai_results or {})
                                merged_results[key] = rewritten
                                session_state.ai_results = merged_results
                                session_state.pptx_ready = False
                            else:
                                st.session_state["book_notice"] = f"The model did not finish {key}; the earlier text stands."
                        st.rerun()

    if st.session_state.get("book_notice"):
        st.warning(st.session_state.pop("book_notice"))
    if bs_keys:
        with ai_tabs[tab_idx]:
            create_account_tabs(bs_keys, "bs_")
        tab_idx += 1
    if is_keys:
        with ai_tabs[tab_idx]:
            create_account_tabs(is_keys, "is_")
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

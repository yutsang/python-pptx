"""Read one entity's workbook and decide which accounts go to commentary: the extraction phase of a run, with no model and no UI.

Split out of the app's export module so the agent session (agent/session.py), the app and the command line read a workbook the same way.
"""
from __future__ import annotations

from typing import Any, Dict, Optional

from . import find_mapping_key, get_effective_mappings, load_mappings
from .selection import detect_statement_mode, explain_reconciliation_selection, selection_without_financials


def batch_extract_entity_data(
    *,
    temp_path: str,
    entity_name: str,
    selected_sheet: Optional[str] = None,
    financials_from: Optional[str] = None,
    financials_sheet: Optional[str] = None,
    mapping_overrides: Optional[Dict[str, str]] = None,
    language: Optional[str] = None,
) -> Dict[str, Any]:
    """Phase 1 of the batch entity pipeline (process + reconcile -- fast).

    Split out from what used to be one single batch_process_entity() call
    so a checkpoint-based batch UI can st.rerun() BETWEEN this and
    batch_run_ai_for_entity() (phase 2, slow). Streamlit only paints the
    browser once a script run returns control to it: an on_data_ready-style
    callback fired midway through one long blocking call can update
    session_state all it wants, but the switcher/data-view code living
    later in that SAME script run still can't actually render anything
    until the whole call returns -- and it doesn't return until AI
    generation is ALSO done, defeating the entire point of showing data
    before AI finishes. Only an actual rerun between two separately
    checkpointed phases makes the browser paint an intermediate state.

    financials_from/financials_sheet point BS/IS extraction at a sibling
    roll-up ("主表") workbook's named sheet when this entity's own file has
    no Financials-pattern sheet of its own — same mechanism
    process_workbook_data already exposes for the single-file flow.

    Returns {"status": "failed", "entity_name", "error"} on failure, or on
    success {"status": "ok", "entity_name", "data_summary": {the same
    "state"-shaped partial bundle a caller can swap into st.session_state
    and pass to render_data_tables_section for the full recon+breakdown
    view, plus match-count summaries}, "_internal": {everything
    batch_run_ai_for_entity needs to continue without re-deriving it}}.
    """
    from . import process_workbook_data

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
    selection = None
    if statement_mode in ("is_only", "bs_only"):
        target_type = "IS" if statement_mode == "is_only" else "BS"
        matched_mapping_keys = [
            k for k in dfs
            if mappings.get(find_mapping_key(k, mappings) or k, {}).get("type") == target_type
        ]
        if not matched_mapping_keys:
            matched_mapping_keys = list(dfs.keys())
    else:
        selection = explain_reconciliation_selection(reconciliation, dfs.keys(), resolution, dfs=dfs)
        matched_mapping_keys = selection.admitted_keys
        has_reconciliation_data = bool(
            reconciliation and any(recon_df is not None and not recon_df.empty for recon_df in reconciliation)
        )
        if not has_reconciliation_data:
            matched_mapping_keys = list(dfs.keys())
            selection = selection_without_financials(matched_mapping_keys)

    if not matched_mapping_keys:
        result["status"] = "failed"
        result["error"] = "No eligible accounts after reconciliation filtering."
        return result

    bs_recon, is_recon = (list(reconciliation) + [None, None])[:2] if reconciliation else (None, None)
    result["data_summary"] = {
        "entity_name": entity_name,
        "accounts_total": len(dfs),
        "accounts_matched": len(matched_mapping_keys),
        "bs_match_counts": bs_recon["Match"].value_counts().to_dict() if bs_recon is not None and not bs_recon.empty else {},
        "is_match_counts": is_recon["Match"].value_counts().to_dict() if is_recon is not None and not is_recon.empty else {},
        # Raw per-account reconciliation breakdowns (same DataFrames
        # render_reconciliation_section uses in the interactive single-file
        # flow) -- so a caller can show the actual account-by-account
        # table, not just the match-status counts.
        "bs_recon_df": bs_recon,
        "is_recon_df": is_recon,
        # Full session_state-shaped (minus ai_results/pptx) partial bundle
        # -- lets a caller swap this into st.session_state and call
        # render_data_tables_section() for the complete per-account
        # breakdown view (cash, investment properties, etc., not just
        # reconciliation), the same rich view a fully-finished entity
        # gets, while AI is still running (or hasn't started yet).
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
            "reconciliation_selection": selection,
        },
    }
    result["_internal"] = {
        "raw_state": state,
        "dfs": dfs,
        "reconciliation": reconciliation,
        "resolution": resolution,
        "mappings": mappings,
        "matched_mapping_keys": matched_mapping_keys,
        "selection": selection,
        "effective_language": effective_language,
        "entity_name": entity_name,
        "temp_path": temp_path,
        "selected_sheet": selected_sheet,
        "financials_from": financials_from,
        "financials_sheet": financials_sheet,
        "mapping_overrides": mapping_overrides,
    }
    return result

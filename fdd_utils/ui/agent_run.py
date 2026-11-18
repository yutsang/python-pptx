"""The agent's whole-databook run, as the app drives it: one entity in the single-file view, one after another in a batch.

No Streamlit call is made here. The panels (ai_panel.py, fdd_app.py) show progress and call this; the run itself is
fdd_utils/agent/book.py, the same path `fdd_agent.py book` and the trials take.
"""
from __future__ import annotations

import os
import re
import time
from pathlib import Path
from typing import Any, Callable, Dict, Optional

PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
OUTPUT_DIR = Path(__file__).resolve().parents[1] / "output"
#: the key the working session lives under in session_state (the agent workbench panel reads the same one)
WORK_KEY = "agent_work"


def deck_out_path(entity_name: Any) -> str:
    """Where the app writes an entity's deck: fdd_utils/output/<entity>_<time>.pptx, as it always has (the delivery label is added to the name)."""
    entity = re.sub(r"[^\w\-_]", "_", str(entity_name or "")).strip("_") or "Entity"
    return str(OUTPUT_DIR / f"{entity}_{time.strftime('%Y%m%d_%H%M%S')}.pptx")


def book_state(result: Any) -> Dict[str, Any]:
    """What a finished run files in session_state (or in a batch entity's bundle): the commentary the tabs show, the statements'
    overviews, the delivery decision, the deck as bytes, and the working session behind them."""
    deck_bytes = b""
    if result.deck_path and os.path.exists(result.deck_path):
        with open(result.deck_path, "rb") as handle:
            deck_bytes = handle.read()
    state: Dict[str, Any] = {
        "ai_results": result.ai_results or None,
        "section_summaries": dict(result.overview),
        "reconciliation_selection": result.work.source.selection,
        "delivery_decision": result.delivery or None,
        WORK_KEY: result.work,
        "pptx_ready": bool(deck_bytes),
    }
    if deck_bytes:
        state.update(pptx_download_data=deck_bytes, pptx_download_filename=os.path.basename(result.deck_path), pptx_download_mime=PPTX_MIME)
    return state


def run_for_state(session_state: Any, *, progress_callback: Optional[Callable[..., Any]] = None) -> Any:
    """Write the book the single-file view has loaded: the workbook, entity and Financials source the person chose, the manual mapping
    overrides, the language selector's choice, and the model the sidebar's dropdown names (agent.models in config.yml wins when set).

    progress_callback(done, total, account) is called as each account finishes."""
    from ..agent import book as B
    rollup = session_state.get("rollup_temp_path") if session_state.get("rollup_sheet") else None
    work = B.open_book(
        session_state.get("temp_path"), entity=session_state.get("entity_name") or "",
        sheet=None if rollup else session_state.get("selected_sheet"), financials_from=rollup,
        financials_sheet=session_state.get("rollup_sheet") if rollup else None,
        mapping_overrides=dict(session_state.get("mapping_overrides") or {}) or None,
        model_type=session_state.get("model_type", "local"), model_name=session_state.get("model_name"),
        language=session_state.get("language") or "")
    return _run(work, progress_callback, deck_out_path(session_state.get("entity_name") or session_state.get("project_name")))


def _run(work: Any, progress_callback: Optional[Callable[..., Any]], out_path: str = "") -> Any:
    from ..agent import book as B
    accounts = work.entity_accounts()
    done = [0]

    def on_account(key: str, turn: Any) -> None:
        done[0] += 1
        if progress_callback is not None:
            progress_callback(done[0], len(accounts), key)

    return B.run_book(work, accounts, on_account=on_account, out_path=out_path)


def run_for_entity(*, extracted: Dict[str, Any], model_type: str = "local", model_name: Optional[str] = None,
                   progress_callback: Optional[Callable[..., Any]] = None) -> Dict[str, Any]:
    """Phase 2 of the batch pipeline: write one entity's book from what the extraction phase left (batch_extract_entity_data), and return
    {"status", "entity_name", "state": the bundle the batch view swaps into session_state} or {"status": "failed", "error"}.

    progress_callback(done, total, account) is called as each account finishes."""
    from ..agent import book as B
    internal = extracted["_internal"]
    entity_name = internal["entity_name"]
    result: Dict[str, Any] = {"entity_name": entity_name, "status": "ok"}
    financials_from = internal["financials_from"]
    work = B.open_book(
        internal["temp_path"], entity=entity_name, sheet=None if financials_from else internal["selected_sheet"], financials_from=financials_from,
        financials_sheet=internal["financials_sheet"] if financials_from else None, mapping_overrides=internal.get("mapping_overrides"),
        model_type=model_type, model_name=model_name, language=internal["effective_language"])
    accounts = work.entity_accounts()
    book = _run(work, progress_callback, deck_out_path(entity_name))
    if not book.deck_path:
        result.update(status="failed", error=f"no deck was written: [{book.error.get('code')}] {str(book.error.get('message') or book.stopped)[:300]}")
        return result
    state = book_state(book)
    raw = internal["raw_state"]
    financials_workbook_path = financials_from or internal["temp_path"]
    state.update({
        "dfs": raw.get("dfs"), "display_dfs": raw.get("display_dfs"), "dfs_variants": raw.get("dfs_variants"),
        "display_df_variants": raw.get("display_df_variants"), "workbook_list": raw.get("workbook_list"),
        "display_workbook_list": raw.get("display_workbook_list"), "language": internal["effective_language"],
        "detected_language": internal["effective_language"], "bs_is_results": raw.get("bs_is_results"),
        "reconciliation": internal["reconciliation"], "resolution": internal["resolution"], "project_name": raw.get("project_name"),
        "entity_name": entity_name,
        # the financials source, so a later "Regenerate PPTX" from the reused single-file view reads the right sheet
        "temp_path": financials_workbook_path,
        "selected_sheet": internal["financials_sheet"] if financials_from else internal["selected_sheet"],
        "mapping_overrides": internal.get("mapping_overrides"), "model_type": model_type, "model_name": model_name,
    })
    result.update(state=state, output_path=book.deck_path, accounts_processed=len(accounts), accepted=len(book.finished),
                  delivery=book.delivery, issue_cases=book.issue_cases, stopped=book.stopped)
    return result


def export_deck(session_state: Any) -> Dict[str, Any]:
    """Write the deck again from the working session's drafts (after an account was written again). Returns the export record; the deck
    and its delivery decision are filed in session_state."""
    work = session_state.get(WORK_KEY)
    if work is None:
        return {"ok": False, "error": {"code": "NO_SESSION", "message": "generate the commentary first"}}
    exported = work.export(out_path=deck_out_path(session_state.get("entity_name") or session_state.get("project_name")), first_generation=True)
    data = exported.get("data") or {}
    path = str(data.get("path") or "")
    if path and os.path.exists(path):
        with open(path, "rb") as handle:
            session_state.pptx_download_data = handle.read()
        session_state.pptx_download_filename = os.path.basename(path)
        session_state.pptx_download_mime = PPTX_MIME
        session_state.pptx_ready = True
        session_state.delivery_decision = data.get("delivery") or None
    return exported

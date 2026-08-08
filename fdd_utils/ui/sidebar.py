from __future__ import annotations

# re-added: bound by an import in another section of the pre-split module
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


from .state import reset_processing_session_state
from datetime import timedelta
import hashlib
from pathlib import Path
import re
import tempfile
import time
from typing import Any, Callable

import streamlit as st

from ..workbook import clear_table_inspection_cache, clear_workbook_caches


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

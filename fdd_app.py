#!/usr/bin/env python3
"""
Streamlit App for Financial Data Processing with AI
Combines extraction, reconciliation, AI generation, and PPTX export
"""

import streamlit as st
import logging
import os
import re
import time
import traceback
from pathlib import Path
from typing import Dict, List

# Import modules
from fdd_utils.ui import (
    batch_process_entity,
    build_entity_selector_model,
    generate_pptx_presentation as render_generate_pptx_presentation,
    initialize_app_state,
    persist_uploaded_workbook,
    render_language_selector,
    render_processed_view,
    render_sidebar_upload,
    should_render_preprocess_controls,
)
from fdd_utils.ai import FDDConfig
from fdd_utils.workbook import (
    process_workbook_data,
    build_workbook_preflight,
    detect_databook_language,
    extract_entity_names_from_preflight,
    get_financial_sheet_options,
    suggest_rollup_sheet_for_entity,
)

# Import PPTX generation
try:
    import fdd_utils.pptx
    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False

# Page config
st.set_page_config(
    page_title="Financial Data Processing",
    page_icon="📊",
    layout="wide"
)

# Custom CSS for full width
st.markdown("""
<style>
.block-container {padding-top: 1rem; max-width: 100% !important;}
.stTabs [data-baseweb="tab-list"] {gap: 2px;}
.stTabs [data-baseweb="tab"] {padding: 10px 20px;}
.main .block-container {max-width: 100%; padding-left: 2rem; padding-right: 2rem;}
[data-testid="stAppViewContainer"] {max-width: 100%;}
/* Replace default Streamlit footer */
footer {visibility: hidden;}
footer:after {
    content: 'Made with Yuu, D&A Hub, TRNS, HK';
    visibility: visible;
    display: block;
    text-align: center;
    padding: 5px;
    color: rgba(120, 120, 120, 0.6);
    font-size: 0.8rem;
}
.fdd-final-commentary {
    padding: 0.9rem 1rem;
    border-radius: 0.75rem;
    border: 1px solid rgba(120, 120, 120, 0.22);
    background: rgba(120, 120, 120, 0.06);
    color: inherit;
    line-height: 1.65;
    max-height: 420px;
    overflow-y: auto;
}
.fdd-final-commentary p {margin: 0 0 0.8rem 0;}
.fdd-final-commentary p:last-child {margin-bottom: 0;}
/* Hallucination = unsupported by data (most severe) → red, stands out.
   Reasoning = inference/interpretation (milder) → orange. */
.fdd-hallucination-clause {
    background-color: rgba(248, 113, 113, 0.22);
    color: inherit;
    font-weight: 600;
    padding: 1px 6px;
    border-radius: 999px;
    border: 1px solid rgba(248, 113, 113, 0.38);
}
.fdd-reasoning-clause {
    background-color: rgba(251, 146, 60, 0.22);
    color: inherit;
    font-weight: 600;
    padding: 1px 6px;
    border-radius: 999px;
    border: 1px solid rgba(251, 146, 60, 0.38);
}
.fdd-validator-notes {
    margin-top: 0.8rem;
    padding: 0.8rem 0.95rem;
    border-radius: 0.65rem;
    border: 1px solid rgba(255, 181, 0, 0.35);
    background: rgba(255, 214, 10, 0.12);
}
.fdd-validator-notes p,
.fdd-validator-notes ul {
    margin: 0;
}
.fdd-validator-notes ul {
    padding-left: 1.1rem;
    margin-top: 0.45rem;
}
@media (prefers-color-scheme: dark) {
    .fdd-final-commentary {
        border-color: rgba(255, 255, 255, 0.14);
        background: rgba(255, 255, 255, 0.04);
    }
    .fdd-hallucination-clause {
        background-color: rgba(248, 113, 113, 0.22);
    }
    .fdd-reasoning-clause {
        background-color: rgba(251, 146, 60, 0.22);
    }
    .fdd-validator-notes {
        border-color: rgba(255, 196, 61, 0.32);
        background: rgba(255, 196, 61, 0.12);
    }
}
/* Streamlit-specific dark mode override */
[data-testid="stApp"][data-theme="dark"] .fdd-hallucination-clause {
    background-color: rgba(248, 113, 113, 0.22);
}
[data-testid="stApp"][data-theme="dark"] .fdd-reasoning-clause {
    background-color: rgba(251, 146, 60, 0.22);
}
</style>
""", unsafe_allow_html=True)


logger = logging.getLogger(__name__)
if not logging.getLogger().handlers:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s - %(message)s")

# Suppress 'missing ScriptRunContext' warnings emitted by Streamlit when
# our ThreadPoolExecutor workers (AI pipeline / summary generator) try to
# log without a script-run context attached. These warnings are cosmetic —
# Streamlit itself notes "This warning can be ignored when running in bare
# mode." — but they spam the console during long pipeline runs.
class _ScriptRunContextFilter(logging.Filter):
    def filter(self, record: logging.LogRecord) -> bool:
        return "missing ScriptRunContext" not in str(record.getMessage())


for _name in (
    "streamlit.runtime.scriptrunner_utils.script_run_context",
    "streamlit.runtime.scriptrunner.script_run_context",
):
    logging.getLogger(_name).addFilter(_ScriptRunContextFilter())

def get_model_display_name(model_type: str) -> str:
    """Return the actual chat_model from config.yml for the resolved provider."""
    try:
        cfg = FDDConfig(model_type=model_type)
        provider = cfg.get_model_config()
        return provider.get("chat_model", "").strip() or cfg.model_type
    except Exception:
        return str(model_type)


def generate_pptx_presentation():
    render_generate_pptx_presentation(
        session_state=st.session_state,
        pptx_available=PPTX_AVAILABLE,
    )


def load_latest_results_from_logs():
    """Load the most recent results from logs directory"""
    import yaml
    import os
    import glob
    
    logs_dir = 'fdd_utils/logs'
    if not os.path.exists(logs_dir):
        return None
    
    # Find all run directories
    run_dirs = glob.glob(os.path.join(logs_dir, 'run_*'))
    if not run_dirs:
        return None
    
    # Get the most recent one
    latest_run = max(run_dirs, key=os.path.getmtime)
    results_file = os.path.join(latest_run, 'results.yml')
    
    if os.path.exists(results_file):
        try:
            with open(results_file, 'r', encoding='utf-8') as f:
                results = yaml.safe_load(f)
            return results
        except Exception as e:
            logger.warning("Error loading results from %s: %s", results_file, e)
            return None
    return None


def init_session_state():
    """Initialize session state variables."""
    initialize_app_state(st.session_state)


@st.cache_data(show_spinner=False)
def get_workbook_preflight(file_path: str) -> Dict[str, object]:
    """Cache workbook metadata and shallow previews for upload-time selectors."""
    return build_workbook_preflight(file_path)


@st.cache_data(show_spinner=False)
def get_entity_names(file_path: str) -> List[str]:
    """Extract potential entity names from visible workbook previews."""
    started = time.perf_counter()
    try:
        preflight = get_workbook_preflight(file_path)
        filtered_names = extract_entity_names_from_preflight(preflight)
        logger.debug(
            "Detected %s candidate entities from %s in %.2fs",
            len(filtered_names),
            os.path.basename(file_path),
            time.perf_counter() - started,
        )
        return filtered_names
    except Exception:
        logger.exception("Failed to detect entity names from %s", file_path)
        return []


@st.cache_data(show_spinner=False)
def get_financial_sheets(file_path: str) -> List[str]:
    """Get visible, non-blank sheet options for the summary-sheet selector."""
    started = time.perf_counter()
    try:
        preflight = get_workbook_preflight(file_path)
        sorted_sheets = get_financial_sheet_options(preflight)
        logger.debug(
            "Scored %s sheets from %s in %.2fs",
            len(sorted_sheets),
            os.path.basename(file_path),
            time.perf_counter() - started,
        )
        return sorted_sheets
    except Exception:
        logger.exception("Failed to inspect sheet list from %s", file_path)
        return []


_FILENAME_STRIP_WORDS = {
    "project", "databook", "data", "book", "workbook", "fdd", "template",
    "final", "draft", "copy", "v1", "v2", "v3", "xlsx", "xls", "confidential",
}


def _extract_entity_from_filename(filename: str) -> str:
    """Extract a meaningful entity name from the uploaded filename."""
    stem = Path(filename).stem
    # Replace separators with spaces
    name = re.sub(r"[._-]+", " ", stem)
    # Remove generic words
    tokens = [t for t in name.split() if t.lower() not in _FILENAME_STRIP_WORDS]
    return " ".join(tokens).strip()


def render_entity_and_sheet_controls(processed: bool = False):
    """Render entity and summary-sheet controls before and after processing."""
    col_entity, col_sheet = st.columns(2)
    temp_path = st.session_state.get('temp_path', None)

    with col_entity:
        st.markdown("**🏢 Entity Name**")
        if temp_path and os.path.exists(temp_path):
            entity_options = get_entity_names(temp_path)
            uploaded_filename = st.session_state.get("uploaded_filename", "")
            if uploaded_filename:
                filename_entity = _extract_entity_from_filename(uploaded_filename)
                if filename_entity and filename_entity not in entity_options:
                    entity_options = list(entity_options) + [filename_entity]
            selector_model = build_entity_selector_model(
                entity_options,
                current_entity_name=st.session_state.get('entity_name') or "",
            )

            if 'entity_text_input' not in st.session_state:
                st.session_state.entity_text_input = selector_model["text_value"]
            if selector_model["show_dropdown"]:
                valid_dropdown_options = [""] + selector_model["dropdown_options"]
                current_dropdown_value = st.session_state.get('entity_dropdown', '')
                preferred_dropdown_value = (
                    st.session_state.get('entity_name')
                    if st.session_state.get('entity_name') in selector_model["dropdown_options"]
                    else ''
                )
                if current_dropdown_value not in valid_dropdown_options:
                    st.session_state.entity_dropdown = preferred_dropdown_value

                selected_entity = st.selectbox(
                    label="Select entity",
                    options=valid_dropdown_options,
                    help="Auto-detected entities from the databook",
                    label_visibility="collapsed",
                    key="entity_dropdown",
                )
                if selected_entity and selected_entity != st.session_state.get('prev_entity_dropdown', ''):
                    st.session_state.entity_text_input = selected_entity
                    st.session_state.prev_entity_dropdown = selected_entity
                    st.session_state.entity_name = selected_entity
            else:
                st.caption("No entity names were detected from the workbook. Enter one manually below.")
                st.session_state.entity_dropdown = ""

            entity_name = st.text_input(
                label="Entity name input",
                placeholder="Type or modify entity name...",
                help="Type a custom entity name or modify the selected one",
                label_visibility="collapsed",
                key="entity_text_input",
            )
            st.session_state.entity_name = str(entity_name).strip()
            if processed and st.session_state.get('entity_name'):
                st.caption(f"Current entity: {st.session_state.get('entity_name')}")
        else:
            st.info("👈 Please upload a databook file first")
            if 'entity_name' not in st.session_state:
                st.session_state.entity_name = ""

    with col_sheet:
        st.markdown("**📊 Financial Statement Sheet**")
        if temp_path and os.path.exists(temp_path):
            sheet_options = get_financial_sheets(temp_path)
            if sheet_options:
                if st.session_state.get('selected_sheet') not in sheet_options:
                    st.session_state.selected_sheet = sheet_options[0]

                selected_sheet = st.selectbox(
                    label="Select sheet",
                    options=sheet_options,
                    index=sheet_options.index(st.session_state.selected_sheet),
                    help="Sheet containing both BS and IS",
                    label_visibility="collapsed",
                    key="sheet_select",
                )
                st.session_state.selected_sheet = selected_sheet
            else:
                st.warning("No sheets found")
                st.session_state.selected_sheet = None
            # Cheap pre-Process detection (sheet profiles only) so the
            # "Detected: ..." reminder in render_language_selector is
            # visible BEFORE the user commits to Process Data, not just
            # after -- otherwise an override happens blind to what the
            # databook actually is. Cached per temp_path so it only runs
            # once per upload, not on every Streamlit rerun.
            if st.session_state.get("_lang_preview_path") != temp_path:
                _raw_detected = detect_databook_language(temp_path)
                st.session_state.detected_language_preview = (
                    "Chn" if str(_raw_detected or "").strip() in ("Chi", "Chn", "chinese", "Chinese")
                    else "Eng" if _raw_detected else None
                )
                st.session_state._lang_preview_path = temp_path
            # Placed here (not the sidebar) to fill the blank space this
            # column otherwise has below the sheet dropdown while Entity
            # Name's column is taller (dropdown/caption + text input) — and
            # gating it on temp_path keeps both columns balanced before
            # upload too (both just show the "please upload" prompt then).
            render_language_selector(st.session_state)
        else:
            st.info("👈 Please upload a databook file first")
            st.session_state.selected_sheet = None

    if temp_path and os.path.exists(temp_path):
        # Some portfolios keep each sub-entity's own databook free of any
        # Financials-pattern sheet at all -- the real summary for that entity
        # instead lives inside a sibling roll-up ("主表") workbook, one sheet
        # per entity. Optional and collapsed by default since most databooks
        # have their own Financials sheet and never need this. Placed BEFORE
        # the Process button, same principle as the language reminder: must
        # be known/set before processing, not discovered after.
        with st.expander("📎 Advanced: Roll-up file (if this databook has no Financials tab)"):
            st.caption(
                "Some portfolios keep each sub-entity's own databook free of any Financials tab -- "
                "the real summary instead lives in a shared roll-up file, one sheet per entity. "
                "Upload that roll-up file here, and manually select the Financials sheet for this entity."
            )
            rollup_file = st.file_uploader(
                "Upload roll-up file (optional)",
                type=["xlsx", "xls"],
                key="rollup_file_uploader",
            )
            if rollup_file is not None:
                rollup_temp_path = persist_uploaded_workbook(
                    uploaded_name=rollup_file.name,
                    uploaded_bytes=rollup_file.getvalue(),
                    session_state=st.session_state,
                    state_key="rollup_temp_path",
                )
                rollup_sheet_options = get_financial_sheets(rollup_temp_path)
                if rollup_sheet_options:
                    current_rollup_sheet = st.session_state.get("rollup_sheet")
                    default_index = (
                        rollup_sheet_options.index(current_rollup_sheet) + 1
                        if current_rollup_sheet in rollup_sheet_options else 0
                    )
                    selected_rollup_sheet = st.selectbox(
                        label="Select the Financials sheet for THIS entity in the roll-up file",
                        options=[""] + rollup_sheet_options,
                        index=default_index,
                        help="No auto-matching by entity name -- pick the exact sheet for this entity "
                             "(e.g. '南通通海Financials').",
                        key="rollup_sheet_select",
                    )
                    st.session_state.rollup_sheet = selected_rollup_sheet or None
                else:
                    st.warning("No sheets found in the roll-up file")
                    st.session_state.rollup_sheet = None
                if st.session_state.get("rollup_sheet"):
                    st.caption(f"✅ Will use {st.session_state['rollup_sheet']!r} from {rollup_file.name!r} "
                               "for reconciliation instead of this entity's own databook.")
            else:
                st.session_state.rollup_temp_path = None
                st.session_state.rollup_sheet = None
    else:
        st.session_state.rollup_temp_path = None
        st.session_state.rollup_sheet = None

    button_label = "🚀 Process Data" if not processed else "🔁 Reprocess Data"
    button_key = "process_data_main" if not processed else "reprocess_data_main"
    if st.button(button_label, type="primary", use_container_width=True, key=button_key):
        # A financials sheet can come from either this databook itself, or
        # (when this entity's own file has none) from an uploaded roll-up
        # file's named sheet -- either one satisfies "there is a Financials
        # source", it doesn't have to be selected_sheet specifically.
        has_financials_source = bool(st.session_state.get('selected_sheet')) or (
            bool(st.session_state.get('rollup_temp_path')) and bool(st.session_state.get('rollup_sheet'))
        )
        if not temp_path:
            st.error("⚠️ Please upload a file first")
        elif not st.session_state.get('entity_name'):
            st.warning("⚠️ Please enter an entity name")
        elif not has_financials_source:
            st.warning("⚠️ Please select a financial statement sheet (from this databook, or from an "
                       "uploaded roll-up file under 'Advanced: Roll-up file' above)")
        else:
            st.session_state.process_data_clicked = True
            st.rerun()


MAX_BATCH_ENTITIES = 8


def render_batch_processing_section():
    """Batch mode: process several entities/databooks in one pass, each
    producing its own standalone PPTX. Headless per-entity, driven by
    batch_process_entity() (fdd_utils/ui.py) rather than the interactive
    single-file session_state flow above -- batch entities don't get the
    per-account AI-result editing UI, only a straight process->AI->export
    run per slot, matching what the user asked for (dynamic +/- slots, one
    shared optional roll-up/主表 file with fuzzy-suggested-but-overridable
    per-entity sheet, one PPTX per entity)."""
    st.markdown("## 📦 Batch Processing")
    st.caption(f"Process multiple entities in one pass -- each entity produces its own PPTX. Maximum {MAX_BATCH_ENTITIES}.")

    if "batch_slot_ids" not in st.session_state:
        st.session_state.batch_slot_ids = [0]
    if "batch_next_id" not in st.session_state:
        st.session_state.batch_next_id = 1

    with st.expander("📎 Shared roll-up file (if these databooks have no Financials tab)"):
        st.caption(
            "If the databooks in this batch have no Financials tab, the real summary can come from a "
            "shared roll-up file (one sheet per entity). Upload it once here -- each entity below can "
            "then select its own sheet (auto-suggested by entity name, but you can override it)."
        )
        rollup_file = st.file_uploader(
            "Upload shared roll-up file (optional)",
            type=["xlsx", "xls"],
            key="batch_rollup_file_uploader",
        )
        if rollup_file is not None:
            persist_uploaded_workbook(
                uploaded_name=rollup_file.name,
                uploaded_bytes=rollup_file.getvalue(),
                session_state=st.session_state,
                state_key="batch_rollup_temp_path",
            )
            st.caption(f"✅ Shared roll-up file: {rollup_file.name}")
        else:
            st.session_state.batch_rollup_temp_path = None

    rollup_temp_path = st.session_state.get("batch_rollup_temp_path")
    rollup_sheet_options = get_financial_sheets(rollup_temp_path) if rollup_temp_path else []

    st.divider()

    slot_ids = list(st.session_state.batch_slot_ids)
    for idx, slot_id in enumerate(slot_ids):
        with st.container():
            st.markdown("---")
            header_col, remove_col = st.columns([6, 1])
            with header_col:
                st.markdown(f"**Entity #{idx + 1}**")
            with remove_col:
                if len(slot_ids) > 1 and st.button("🗑️", key=f"batch_remove_{slot_id}", help="Remove this entity"):
                    st.session_state.batch_slot_ids = [s for s in slot_ids if s != slot_id]
                    for prefix in ("batch_file", "batch_entity", "batch_own_sheet", "batch_rollup_sheet", "batch_temp_path", "_batch_resolved"):
                        st.session_state.pop(f"{prefix}_{slot_id}", None)
                    st.rerun()

            file_col, entity_col = st.columns(2)
            with file_col:
                own_file = st.file_uploader("Databook", type=["xlsx", "xls"], key=f"batch_file_{slot_id}")
                own_temp_path = None
                if own_file is not None:
                    own_temp_path = persist_uploaded_workbook(
                        uploaded_name=own_file.name,
                        uploaded_bytes=own_file.getvalue(),
                        session_state=st.session_state,
                        state_key=f"batch_temp_path_{slot_id}",
                    )
            with entity_col:
                entity_key = f"batch_entity_{slot_id}"
                if entity_key not in st.session_state and own_file is not None:
                    default_entity = _extract_entity_from_filename(own_file.name)
                    if default_entity:
                        st.session_state[entity_key] = default_entity
                entity_name = st.text_input("Entity name", key=entity_key, placeholder="Entity name")

            own_sheet = ""
            rollup_sheet_choice = ""
            sheet_col1, sheet_col2 = st.columns(2)
            with sheet_col1:
                if own_temp_path:
                    own_sheet_options = get_financial_sheets(own_temp_path)
                    own_choices = [""] + own_sheet_options
                    # Blank by default whenever a shared roll-up file is present
                    # (ambiguous which source should win); auto-pick the first
                    # ranked sheet only when there's no roll-up alternative at
                    # all, matching single-file mode's own behaviour exactly.
                    default_own = "" if rollup_temp_path else (own_sheet_options[0] if own_sheet_options else "")
                    own_sheet = st.selectbox(
                        "This file's own Financials sheet",
                        options=own_choices,
                        index=own_choices.index(default_own) if default_own in own_choices else 0,
                        key=f"batch_own_sheet_{slot_id}",
                    )
            with sheet_col2:
                if rollup_temp_path and rollup_sheet_options:
                    suggested = suggest_rollup_sheet_for_entity(entity_name, rollup_sheet_options) if entity_name else None
                    rollup_choices = [""] + rollup_sheet_options
                    default_rollup = suggested or ""
                    rollup_sheet_choice = st.selectbox(
                        "This entity's sheet in the roll-up file (auto-suggested, editable)",
                        options=rollup_choices,
                        index=rollup_choices.index(default_rollup) if default_rollup in rollup_choices else 0,
                        key=f"batch_rollup_sheet_{slot_id}",
                    )

            st.session_state[f"_batch_resolved_{slot_id}"] = {
                "temp_path": own_temp_path,
                "entity_name": entity_name.strip() if entity_name else "",
                "own_sheet": own_sheet or None,
                "rollup_sheet": rollup_sheet_choice or None,
            }

    add_col, _spacer = st.columns([1, 5])
    with add_col:
        if len(slot_ids) < MAX_BATCH_ENTITIES:
            if st.button("➕ Add entity", key="batch_add_slot"):
                new_id = st.session_state.batch_next_id
                st.session_state.batch_next_id += 1
                st.session_state.batch_slot_ids.append(new_id)
                st.rerun()
        else:
            st.caption(f"Maximum of {MAX_BATCH_ENTITIES} entities reached")

    st.divider()

    ready_slots = []
    for slot_id in slot_ids:
        resolved = st.session_state.get(f"_batch_resolved_{slot_id}") or {}
        if resolved.get("temp_path") and resolved.get("entity_name") and (
            resolved.get("own_sheet") or (rollup_temp_path and resolved.get("rollup_sheet"))
        ):
            ready_slots.append(resolved)

    st.caption(f"{len(ready_slots)} / {len(slot_ids)} entities ready to process.")

    if st.button(
        f"🚀 Process All ({len(ready_slots)})",
        type="primary",
        use_container_width=True,
        disabled=not ready_slots,
        key="batch_process_all",
    ):
        results = []
        progress_bar = st.progress(0.0)
        status = st.empty()
        for i, slot in enumerate(ready_slots):
            status.info(f"⏳ Processing {slot['entity_name']} ({i + 1}/{len(ready_slots)})…")
            try:
                outcome = batch_process_entity(
                    temp_path=slot["temp_path"],
                    entity_name=slot["entity_name"],
                    selected_sheet=slot["own_sheet"],
                    financials_from=rollup_temp_path if not slot["own_sheet"] else None,
                    financials_sheet=slot["rollup_sheet"] if not slot["own_sheet"] else None,
                    model_type=st.session_state.get("model_type", "local"),
                    model_name=st.session_state.get("model_name"),
                    use_multithreading=st.session_state.get("use_multithreading", True),
                )
            except Exception as exc:
                outcome = {"entity_name": slot["entity_name"], "status": "failed", "error": str(exc)}
            results.append(outcome)
            progress_bar.progress((i + 1) / len(ready_slots))
        status.empty()
        st.session_state.batch_results = results
        st.rerun()

    batch_results = st.session_state.get("batch_results")
    if batch_results:
        st.markdown("### Results")
        for outcome in batch_results:
            entity_label = outcome.get("entity_name", "?")
            if outcome.get("status") == "ok":
                output_path = outcome.get("output_path")
                st.success(
                    f"✅ {entity_label}: BS={outcome.get('bs_count', 0)} IS={outcome.get('is_count', 0)} "
                    f"accounts={outcome.get('accounts_processed', 0)}"
                )
                if output_path and os.path.exists(output_path):
                    with open(output_path, "rb") as handle:
                        st.download_button(
                            label=f"⬇️ Download {os.path.basename(output_path)}",
                            data=handle.read(),
                            file_name=os.path.basename(output_path),
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            key=f"batch_download_{entity_label}_{output_path}",
                        )
            else:
                st.error(f"❌ {entity_label}: {outcome.get('error', 'unknown error')}")


# Initialize
init_session_state()

# Sidebar - must run first to set temp_path before main content reads it
temp_path = render_sidebar_upload(st.session_state, get_model_display_name)

if st.session_state.get("batch_mode"):
    render_batch_processing_section()
else:
    if should_render_preprocess_controls(processed=st.session_state.get('dfs') is not None):
        render_entity_and_sheet_controls(processed=False)

    # Process data if button was clicked
    if st.session_state.get('process_data_clicked', False):
        st.session_state.process_data_clicked = False
        temp_path = st.session_state.get('temp_path', None)
        entity_name = st.session_state.get('entity_name', '')
        selected_sheet = st.session_state.get('selected_sheet', None)

        if temp_path:
            with st.spinner("Processing..."):
                try:
                    fdd_config = FDDConfig()
                    debug_mode = fdd_config.get_debug_mode()
                    processed_state = process_workbook_data(
                        temp_path=temp_path,
                        entity_name=entity_name,
                        selected_sheet=selected_sheet,
                        mapping_overrides=st.session_state.get("mapping_overrides") or None,
                        debug=debug_mode,
                        financials_from=st.session_state.get("rollup_temp_path"),
                        financials_sheet=st.session_state.get("rollup_sheet"),
                    )
                    # Language: auto-match from the databook, but in the UI convention
                    # ("Eng"/"Chn") so every downstream == "Chn" check agrees, and NEVER
                    # overwrite a manual override the project team already set this session.
                    _detected_lang = processed_state.pop("language", "Eng")
                    _detected_ui = "Chn" if str(_detected_lang).strip() in ("Chi", "Chn", "chinese", "Chinese") else "Eng"
                    st.session_state.update(
                        {key: value for key, value in processed_state.items() if key != "display_dfs_original"}
                    )
                    st.session_state.detected_language = _detected_ui
                    if not st.session_state.get("language_user_set"):
                        st.session_state.language = _detected_ui
                    if 'model_type' not in st.session_state:
                        st.session_state.model_type = 'local'
                    st.success("✅ Data processed successfully!")
                    st.rerun()

                except Exception as e:
                    st.error(f"❌ Error processing data: {e}")
                    st.code(traceback.format_exc())

    # Main content
    if st.session_state.get('dfs') is None:
        st.info("👈 Upload a databook, set entity name and sheet, then click 'Process Data' to begin")
    else:
        # render_entity_and_sheet_controls (with the language selector) only
        # renders BEFORE processing (should_render_preprocess_controls hides it
        # once dfs exists) — but detected_language only gets a real value
        # DURING processing, so the "Detected: ..." reminder could never
        # actually be seen without also showing the selector here.
        lang_col, _spacer_col = st.columns([1, 2])
        with lang_col:
            render_language_selector(st.session_state)
        render_processed_view(
            session_state=st.session_state,
            generate_pptx_callback=generate_pptx_presentation,
            get_model_display_name=get_model_display_name,
        )




if __name__ == "__main__":
    pass


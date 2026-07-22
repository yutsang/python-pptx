#!/usr/bin/env python3
"""
Streamlit App for Financial Data Processing with AI
Combines extraction, reconciliation, AI generation, and PPTX export
"""

import streamlit as st
import io
import logging
import os
import re
import time
import traceback
import zipfile
from pathlib import Path
from typing import Dict, List

# Import modules
from fdd_utils.ui import (
    batch_extract_entity_data,
    batch_run_ai_for_entity,
    build_entity_selector_model,
    generate_pptx_presentation as render_generate_pptx_presentation,
    initialize_app_state,
    persist_uploaded_workbook,
    render_bridge_lab,
    render_bridge_lab_toggle,
    render_data_tables_section,
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
                preferred_language=st.session_state.get('language'),
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


def _reactive_selectbox_default(widget_key: str, options: list, desired_default: str) -> None:
    """Pre-seed st.session_state[widget_key] with desired_default before the
    widget is instantiated. A keyed st.selectbox only honours its `index`
    argument on the very FIRST render -- every rerun after that reads
    straight from session_state and ignores `index` entirely, so a
    recommendation that changes across reruns (e.g. the roll-up sheet
    fuzzy-match once the entity name is typed in) would otherwise never
    actually reach the widget. Only overwrites when the widget has no value
    yet, its stored value fell out of `options` (stale, e.g. roll-up file
    swapped), or its current value still matches whatever we last
    auto-set -- a manual user choice always wins and is never clobbered.
    """
    auto_key = f"_{widget_key}_auto"
    last_auto = st.session_state.get(auto_key)
    current = st.session_state.get(widget_key)
    if widget_key not in st.session_state or current not in options:
        st.session_state[widget_key] = desired_default
        st.session_state[auto_key] = desired_default
    elif desired_default != last_auto and current == last_auto:
        st.session_state[widget_key] = desired_default
        st.session_state[auto_key] = desired_default


def _reactive_text_input_default(widget_key: str, desired_default: str) -> None:
    """Text-input analogue of _reactive_selectbox_default. A free-text field
    has no fixed `options` list to check membership against (any string the
    user types is valid), so the only signal available for "has the user
    manually overridden this" is whether the current value still matches
    the last value this function itself auto-set — same manual-choice-wins
    guarantee, adapted for st.text_input instead of st.selectbox.
    """
    auto_key = f"_{widget_key}_auto"
    last_auto = st.session_state.get(auto_key)
    current = st.session_state.get(widget_key)
    if widget_key not in st.session_state or current == last_auto:
        st.session_state[widget_key] = desired_default
        st.session_state[auto_key] = desired_default


_BATCH_SWITCHER_KEYS = ("batch_active_entity_top", "batch_active_entity_bottom")


def _resolve_active_entity(processed_order: list) -> str:
    """Resolves the shared active entity ONCE, up front, by peeking at both
    switcher widgets' raw stored values (no rendering yet) -- deliberately
    NOT folded into the render step below. render_processed_view's tables/
    AI content is decided using this value long before the SECOND switcher
    instance (buried inside render_processed_view, just above its AI
    section) ever gets a chance to render; resolving both widgets first
    means a click on that lower instance still updates what's shown THIS
    run, instead of only taking effect after one extra rerun.

    A widget's stored value counts as "just clicked" when it differs from
    whatever canonical value it last agreed with (tracked per-widget as
    "_{key}_prev_canonical") -- not just "differs from current canonical",
    since the non-clicked widget also legitimately differs from a NEW
    canonical until it's re-synced.
    """
    canonical = st.session_state.get("batch_active_entity")
    if canonical not in processed_order:
        canonical = processed_order[0]

    for widget_key in _BATCH_SWITCHER_KEYS:
        stored = st.session_state.get(widget_key)
        prev_canonical = st.session_state.get(f"_{widget_key}_prev_canonical")
        if stored in processed_order and stored != prev_canonical:
            canonical = stored

    st.session_state["batch_active_entity"] = canonical
    return canonical


def _render_entity_switcher(processed_order: list, widget_key: str, active_entity: str) -> None:
    """Renders one instance of the entity switcher, forced to already-
    resolved active_entity (see _resolve_active_entity) -- always safe to
    pre-seed unconditionally here since resolution already happened."""
    st.session_state[widget_key] = active_entity
    st.radio(
        "Entity", options=processed_order, horizontal=True, key=widget_key,
        label_visibility="collapsed",
    )
    st.session_state[f"_{widget_key}_prev_canonical"] = active_entity


def render_batch_processing_section():
    """Batch mode: process several entities/databooks in one pass, each
    producing its own standalone PPTX.

    Upload model: databooks are uploaded once in the sidebar (2+ files
    there is what puts batch mode into effect at all, no separate toggle
    -- see render_sidebar_upload); a shared optional roll-up file plus
    each entity's own name/sheet choices are configured here in the
    "Setup" expander.

    Processing model: TWO checkpointed phases per entity (batch_current_
    phase = "extract" then "ai", fdd_utils/ui.py's batch_extract_entity_
    data / batch_run_ai_for_entity), not one call per entity. Streamlit
    only paints the browser once a script run returns control to it, so
    the extract phase caches its data-only bundle and reruns BEFORE the
    AI phase even starts -- that's what actually lets an entity render
    as selectable with its full data breakdown while its own AI
    generation is still running, which a callback fired midway through a
    single blocking call could never achieve on its own (see the two
    phase functions' own docstrings for why). Each entity's bundle gets
    upgraded in place once its AI phase finishes.

    Review model: entities become browsable as soon as ANY of them has
    reached the data-extraction checkpoint, growing as more finish, not
    gated on the whole batch completing. Selecting an entity swaps its
    cached bundle into the live st.session_state and re-renders the
    EXISTING single-file render_processed_view UI (tables, reconciliation,
    AI output, PPTX export) UNCHANGED -- a plain st.radio switcher is used
    instead of st.tabs() specifically because st.tabs() executes every
    tab's content on every rerun (just hides the inactive ones visually),
    which would instantiate render_processed_view's non-entity-scoped
    widget keys more than once per run and crash with duplicate-key
    errors; a radio-gated `if active == entity` only ever renders one
    entity's tree per run.
    """
    st.markdown("## 📦 Batch Processing")

    in_progress = bool(st.session_state.get("batch_processing_in_progress"))
    done = bool(st.session_state.get("batch_processing_done"))
    setup_expanded = not in_progress and not done

    with st.expander("⚙️ Setup", expanded=setup_expanded):
        st.caption(f"Upload a shared roll-up file (optional) and every entity's databook at once -- "
                   f"each databook becomes its own entity, producing its own PPTX. Maximum {MAX_BATCH_ENTITIES}.")

        # Shared across the whole batch -- all entities in one run are
        # generated in the same language, same as picking it once up front
        # for a single-file run.
        render_language_selector(st.session_state)

        with st.container():
            st.markdown("**📎 Shared roll-up file** (if these databooks have no Financials tab)")
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

        # Separate from each entity's own per-file "Output filename" below --
        # this one names the single merged deck the "Combine into One PPTX"
        # download produces when 2+ entities are processed. Suggested from
        # the roll-up file's own name when one's uploaded (closest thing to
        # a name for "the whole batch"), since there's no single entity name
        # to fall back to the way each per-entity field does.
        default_combined_name = "Batch_Combined"
        if rollup_file is not None:
            default_combined_name = re.sub(
                r"[^\w\-]", "_", os.path.splitext(rollup_file.name)[0]
            ).strip("_") or default_combined_name
        combined_filename_key = "batch_combined_filename"
        _reactive_text_input_default(combined_filename_key, default_combined_name)
        st.text_input(
            "Combined PPTX filename (no extension -- names the single merged deck when 2+ entities are combined)",
            key=combined_filename_key,
        )

        # Databooks themselves are uploaded once in the sidebar (2+ files
        # there is what puts batch mode into effect at all -- see
        # render_sidebar_upload) and already persisted to stable temp paths
        # by the time this section renders; no separate uploader here.
        uploaded_files_meta = st.session_state.get("batch_uploaded_files_meta") or []
        if len(uploaded_files_meta) > MAX_BATCH_ENTITIES:
            st.warning(f"Only the first {MAX_BATCH_ENTITIES} files will be processed (maximum {MAX_BATCH_ENTITIES} entities per batch).")
            uploaded_files_meta = uploaded_files_meta[:MAX_BATCH_ENTITIES]

        ready_slots = []
        for idx, file_meta in enumerate(uploaded_files_meta):
            # Stable per-file identity (name+size, sanitized) so entity name /
            # sheet choices persist across reruns for the SAME uploaded file --
            # a re-upload of a differently-sized file under the same name still
            # gets its own fresh state instead of reusing stale choices. Must
            # match the slot_id formula render_sidebar_upload used when it
            # persisted this same file, since that's the state_key this
            # temp_path was stored under.
            slot_id = re.sub(r"[^\w\-]", "_", f"{file_meta['name']}_{file_meta['size']}")
            own_temp_path = file_meta["temp_path"]

            with st.container():
                st.markdown("---")
                st.markdown(f"**Entity #{idx + 1}** — `{file_meta['name']}`")

                entity_key = f"batch_entity_{slot_id}"
                entity_auto_key = f"_{entity_key}_auto"
                entity_options = get_entity_names(own_temp_path)
                filename_entity = _extract_entity_from_filename(file_meta['name'])
                if filename_entity and filename_entity not in entity_options:
                    entity_options = list(entity_options) + [filename_entity]

                # A stored entity name only counts as "manually decided" (and
                # so exempt from re-suggestion) if it differs from whatever
                # THIS logic itself last auto-set -- otherwise, once the
                # very first render seeds e.g. the English half by default,
                # toggling the language selector to Chinese afterward could
                # never re-trigger the Chinese-preferring suggestion, since
                # build_entity_selector_model would always see a non-empty
                # "already decided" value. Same manual-choice-wins pattern
                # _reactive_selectbox_default uses, adapted here because the
                # "decided" value also depends on entity_options/preferred_
                # language, not just a fixed options list.
                stored_entity_name = st.session_state.get(entity_key, "")
                is_still_auto = stored_entity_name == st.session_state.get(entity_auto_key, "")
                effective_current = "" if (not stored_entity_name or is_still_auto) else stored_entity_name

                selector_model = build_entity_selector_model(
                    entity_options,
                    current_entity_name=effective_current,
                    preferred_language=st.session_state.get("language"),
                )
                if not stored_entity_name or is_still_auto:
                    st.session_state[entity_key] = selector_model["text_value"]
                    st.session_state[entity_auto_key] = selector_model["text_value"]

                name_col1, name_col2 = st.columns(2)
                with name_col1:
                    if selector_model["show_dropdown"]:
                        dropdown_key = f"batch_entity_dd_{slot_id}"
                        dd_options = [""] + selector_model["dropdown_options"]
                        _reactive_selectbox_default(dropdown_key, dd_options, selector_model["text_value"])
                        picked = st.selectbox(
                            "Entity name suggestions", options=dd_options, key=dropdown_key,
                            help="Auto-detected from the databook -- pick one, or type a custom name on the right.",
                        )
                        prev_picked_key = f"_prev_{dropdown_key}"
                        if picked and picked != st.session_state.get(prev_picked_key, ""):
                            # An explicit dropdown pick IS a manual decision --
                            # deliberately does NOT update entity_auto_key, so
                            # is_still_auto correctly reads False from here on
                            # and a later language toggle won't clobber it.
                            st.session_state[entity_key] = picked
                            st.session_state[prev_picked_key] = picked
                    else:
                        st.caption("No entity names detected -- enter one manually.")
                with name_col2:
                    entity_name = st.text_input("Entity name", key=entity_key, placeholder="Entity name")

                own_sheet = ""
                rollup_sheet_choice = ""
                sheet_col1, sheet_col2 = st.columns(2)
                with sheet_col1:
                    own_sheet_options = get_financial_sheets(own_temp_path)
                    own_choices = [""] + own_sheet_options
                    # Blank by default whenever a shared roll-up file is present
                    # (ambiguous which source should win); auto-pick the first
                    # ranked sheet only when there's no roll-up alternative at
                    # all, matching single-file mode's own behaviour exactly.
                    default_own = "" if rollup_temp_path else (own_sheet_options[0] if own_sheet_options else "")
                    own_sheet_key = f"batch_own_sheet_{slot_id}"
                    _reactive_selectbox_default(own_sheet_key, own_choices, default_own)
                    own_sheet = st.selectbox(
                        "This file's own Financials sheet",
                        options=own_choices,
                        key=own_sheet_key,
                    )
                with sheet_col2:
                    if rollup_temp_path and rollup_sheet_options:
                        suggested = suggest_rollup_sheet_for_entity(entity_name, rollup_sheet_options) if entity_name else None
                        rollup_choices = [""] + rollup_sheet_options
                        rollup_sheet_key = f"batch_rollup_sheet_{slot_id}"
                        _reactive_selectbox_default(rollup_sheet_key, rollup_choices, suggested or "")
                        rollup_sheet_choice = st.selectbox(
                            "This entity's sheet in the roll-up file (auto-suggested, editable)",
                            options=rollup_choices,
                            key=rollup_sheet_key,
                        )

            resolved = {
                "temp_path": own_temp_path,
                "entity_name": entity_name.strip() if entity_name else "",
                "own_sheet": own_sheet or None,
                "rollup_sheet": rollup_sheet_choice or None,
            }
            if resolved["temp_path"] and resolved["entity_name"] and (
                resolved["own_sheet"] or (rollup_temp_path and resolved["rollup_sheet"])
            ):
                ready_slots.append(resolved)

        st.divider()
        st.caption(f"{len(ready_slots)} / {len(uploaded_files_meta)} entities ready to process.")

        if st.button(
            f"🚀 Start Batch Processing ({len(ready_slots)})",
            type="primary",
            use_container_width=True,
            disabled=not ready_slots or in_progress,
            key="batch_start",
        ):
            st.session_state.batch_ready_slots = ready_slots
            st.session_state.batch_rollup_temp_path_snapshot = rollup_temp_path
            st.session_state.batch_processing_in_progress = True
            st.session_state.batch_processing_done = False
            st.session_state.batch_current_index = 0
            st.session_state.batch_current_phase = "extract"
            st.session_state.batch_pending_extracted = None
            st.session_state.batch_start_time = time.time()
            st.session_state.batch_entity_cache = {}
            st.session_state.batch_entity_order = []
            st.session_state.batch_processed_entity_order = []
            st.session_state.batch_failed_entities = []
            st.session_state.batch_zip_data = None
            st.session_state.batch_combined_pptx = None
            st.rerun()

    failed_entities = st.session_state.get("batch_failed_entities")
    if failed_entities:
        for failure in failed_entities:
            st.error(f"❌ {failure['entity_name']}: {failure['error']}")

    # --- Review: entity switcher reusing the single-file UI -- browsable as
    # soon as ANY entity has at least reached the data-extraction stage,
    # growing as more finish, not gated on the whole batch completing.
    # Placed BEFORE the auto-continuing processing block below (not after
    # it) deliberately: every branch of that block ends in st.rerun(),
    # which aborts the rest of THIS script run immediately -- code placed
    # after it never executes at all while batch_processing_in_progress
    # stays True, i.e. for the ENTIRE duration of a batch run. Rendering
    # the switcher first means every phase-transition rerun (and the
    # in-place progress-bar/status updates during a long blocking AI call)
    # actually reaches the browser with the latest cached entities visible
    # and selectable, instead of only ever appearing once the whole batch
    # finishes. ---
    # Mutable holder so the progress-bar placeholder created below (inside
    # render_processed_view's before_ai_section hook, positioned right under
    # the "AI Content Generation" header) can be reused by the batch engine
    # block further down in this same function -- Streamlit renders a
    # placeholder at the position it was CREATED, not where it's later
    # filled in, so this is what lets the progress bar visually sit between
    # the header and the entity switcher even though the code that drives it
    # (and must rerun last, see the block below) runs after both.
    progress_slot: Dict[str, object] = {}

    entity_order = st.session_state.get("batch_entity_order") or []
    if entity_order:
        st.divider()

        processed_order = st.session_state.get("batch_processed_entity_order") or []
        if st.session_state.get("batch_processing_done") and processed_order:
            batch_timestamp = st.session_state.get("batch_export_timestamp") or time.strftime("%Y%m%d_%H%M%S")
            zip_data = st.session_state.get("batch_zip_data")
            combined_data = st.session_state.get("batch_combined_pptx")

            download_col1, download_col2 = st.columns(2)
            with download_col1:
                st.download_button(
                    label=f"⬇️ Download All ({len(processed_order)}) as ZIP",
                    data=zip_data or b"",
                    file_name=f"batch_export_{batch_timestamp}.zip",
                    mime="application/zip",
                    use_container_width=True,
                    disabled=not zip_data,
                    key="batch_download_zip",
                )
            with download_col2:
                if combined_data:
                    combined_name = re.sub(
                        r"[^\w\-]", "_", str(st.session_state.get("batch_combined_filename") or "")
                    ).strip("_") or "Batch_Combined"
                    st.download_button(
                        label=f"⬇️ Download Combined PPTX ({len(processed_order)} entities)",
                        data=combined_data,
                        file_name=f"{combined_name}_{batch_timestamp}.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True,
                        key="batch_download_combined",
                    )
                else:
                    st.caption("Combined PPTX needs 2+ entities.")
        elif st.session_state.get("batch_processing_in_progress"):
            st.caption("⬇️ Downloads become available once the whole batch finishes.")

        active_entity = _resolve_active_entity(entity_order)
        _render_entity_switcher(entity_order, "batch_active_entity_top", active_entity)
        bundle = (st.session_state.get("batch_entity_cache") or {}).get(active_entity)
        if bundle:
            st.session_state.update(bundle)
            if bundle.get("ai_results"):
                def _before_ai_section():
                    # Progress bar + its status text first (placeholders only
                    # -- filled in by the batch engine block below), switcher
                    # second, so both sit directly under "AI Content
                    # Generation" and above the switcher, together, instead
                    # of the status text being left behind at the bottom.
                    progress_slot["placeholder"] = st.empty()
                    progress_slot["status_placeholder"] = st.empty()
                    _render_entity_switcher(entity_order, "batch_active_entity_bottom", active_entity)

                render_processed_view(
                    session_state=st.session_state,
                    generate_pptx_callback=generate_pptx_presentation,
                    get_model_display_name=get_model_display_name,
                    before_ai_section=_before_ai_section,
                    show_download_button=False,
                )
            else:
                st.info(f"⏳ AI generation for **{active_entity}** hasn't finished yet -- showing extracted data only.")
                # Same placeholder trick as the ai_results branch above --
                # this is the common mid-batch case (viewing the entity
                # that's currently being generated), so the progress bar and
                # its status text need a stable anchor here too, not just
                # once results exist.
                progress_slot["placeholder"] = st.empty()
                progress_slot["status_placeholder"] = st.empty()
                render_data_tables_section(st.session_state)

    # --- Processing: TWO checkpointed phases per entity (extract, then
    # AI+export), not one big blocking loop and not even one blocking call
    # per entity. Streamlit only paints the browser once a script run
    # returns control to it -- an on_data_ready-style callback fired
    # MIDWAY through a single long batch_process_entity() call can update
    # session_state all it wants, but the switcher/data-view code rendered
    # above already caught up on this run before this block ever starts,
    # so it doesn't need to wait for a callback -- only an actual
    # st.rerun() between the two phases can advance to the NEXT entity's
    # data becoming visible, though.
    # batch_extract_entity_data() (fast) runs, caches its data-only bundle,
    # reruns; THEN batch_run_ai_for_entity() (slow) runs for that same
    # entity using what extraction already produced, reruns again before
    # moving to the next entity. ---
    if st.session_state.get("batch_processing_in_progress"):
        from fdd_utils.ai import SUBAGENT_SEQUENCE
        n_stages = len(SUBAGENT_SEQUENCE)

        ready_slots = st.session_state.get("batch_ready_slots") or []
        rollup_temp_path = st.session_state.get("batch_rollup_temp_path_snapshot")
        batch_language = st.session_state.get("language", "Eng")
        total = len(ready_slots)
        idx = st.session_state.get("batch_current_index", 0)
        phase = st.session_state.get("batch_current_phase", "extract")
        batch_start_time = st.session_state.get("batch_start_time") or time.time()

        # Prefer the placeholders created earlier (right under "AI Content
        # Generation", above the entity switcher) so the bar AND its status
        # text render there together instead of the bar moving up while its
        # label is left behind at the bottom of the page; falls back to
        # plain st.empty()/st.progress() here only when no entity is
        # selectable yet (e.g. the very first entity is still extracting).
        _progress_host = progress_slot.get("placeholder") or st
        progress_bar = _progress_host.progress(min(idx / total, 1.0) if total else 1.0)

        if idx < total:
            slot = ready_slots[idx]
            status = progress_slot.get("status_placeholder") or st.empty()

            if phase == "extract":
                status.info(f"⏳ Entity {idx + 1}/{total}: {slot['entity_name']} — extracting data...")
                try:
                    extracted = batch_extract_entity_data(
                        temp_path=slot["temp_path"],
                        entity_name=slot["entity_name"],
                        selected_sheet=slot["own_sheet"],
                        financials_from=rollup_temp_path if not slot["own_sheet"] else None,
                        financials_sheet=slot["rollup_sheet"] if not slot["own_sheet"] else None,
                        language=batch_language,
                    )
                except Exception as exc:
                    extracted = {"entity_name": slot["entity_name"], "status": "failed", "error": str(exc)}

                if extracted.get("status") == "ok":
                    # Cached and rerun HERE (not after AI too) is what
                    # actually lets this entity render as selectable with
                    # its full data breakdown while its own AI generation
                    # (next phase, next rerun) hasn't started yet.
                    st.session_state.batch_entity_cache[extracted["entity_name"]] = extracted["data_summary"]["state"]
                    if extracted["entity_name"] not in st.session_state.batch_entity_order:
                        st.session_state.batch_entity_order.append(extracted["entity_name"])
                    st.session_state.batch_pending_extracted = extracted
                    st.session_state.batch_current_phase = "ai"
                else:
                    st.session_state.batch_failed_entities.append(
                        {"entity_name": slot["entity_name"], "error": extracted.get("error", "unknown error")}
                    )
                    st.session_state.batch_current_index = idx + 1
                    st.session_state.batch_current_phase = "extract"
                    st.session_state.batch_pending_extracted = None
                st.session_state.batch_start_time = batch_start_time
                st.rerun()

            else:  # phase == "ai"
                extracted = st.session_state.get("batch_pending_extracted")

                def _progress_cb(agent_num, agent_name, item_num, total_items_in_agent, completed_items,
                                  key_name=None, _idx=idx, _entity=slot["entity_name"]):
                    entity_total_steps = max(1, n_stages * total_items_in_agent)
                    entity_fraction = min(completed_items / entity_total_steps, 1.0)
                    overall_fraction = min((_idx + entity_fraction) / total, 1.0) if total else 1.0
                    progress_bar.progress(overall_fraction)

                    elapsed = time.time() - batch_start_time
                    if overall_fraction > 0.02:
                        eta_seconds = elapsed / overall_fraction * (1 - overall_fraction)
                        eta_display = f"{int(eta_seconds // 60)}m {int(eta_seconds % 60)}s"
                    else:
                        eta_display = "Calculating..."

                    key_display = f" | Key: {key_name}" if key_name else ""
                    status.info(
                        f"⏳ {overall_fraction:.0%} overall | Entity {_idx + 1}/{total}: {_entity} "
                        f"— Stage {agent_num}/{n_stages}: {agent_name} | Item {item_num}/{total_items_in_agent}"
                        f"{key_display} | ETA: {eta_display}"
                    )

                status.info(f"⏳ Entity {idx + 1}/{total}: {slot['entity_name']} — running AI generation...")
                try:
                    outcome = batch_run_ai_for_entity(
                        extracted=extracted,
                        model_type=st.session_state.get("model_type", "local"),
                        model_name=st.session_state.get("model_name"),
                        use_multithreading=st.session_state.get("use_multithreading", True),
                        progress_callback=_progress_cb,
                    )
                except Exception as exc:
                    outcome = {"entity_name": slot["entity_name"], "status": "failed", "error": str(exc)}

                if outcome.get("status") == "ok" and outcome.get("state"):
                    # Overwrites the data-only bundle the extract phase
                    # cached with the full one (now including ai_results)
                    # for this same entity_name -- the switcher already had
                    # it selectable, this just upgrades what gets shown.
                    st.session_state.batch_entity_cache[outcome["entity_name"]] = outcome["state"]
                    if outcome["entity_name"] not in st.session_state.batch_entity_order:
                        st.session_state.batch_entity_order.append(outcome["entity_name"])
                    st.session_state.batch_processed_entity_order.append(outcome["entity_name"])
                else:
                    st.session_state.batch_failed_entities.append(
                        {"entity_name": slot["entity_name"], "error": outcome.get("error", "unknown error")}
                    )

                st.session_state.batch_pending_extracted = None
                st.session_state.batch_start_time = batch_start_time
                st.session_state.batch_current_index = idx + 1
                st.session_state.batch_current_phase = "extract"
                st.rerun()
        else:
            # Every entity attempted -- build the ZIP and combined PPTX once,
            # while everything's already in memory, so the download buttons
            # below are real one-click downloads, not a "build" step first.
            processed_order = st.session_state.batch_processed_entity_order
            entity_cache = st.session_state.batch_entity_cache
            batch_timestamp = time.strftime("%Y%m%d_%H%M%S")
            st.session_state.batch_export_timestamp = batch_timestamp
            if processed_order:
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    for entity_name in processed_order:
                        bundle = entity_cache.get(entity_name) or {}
                        pptx_bytes = bundle.get("pptx_download_data")
                        pptx_filename = bundle.get("pptx_download_filename")
                        if pptx_bytes and pptx_filename:
                            zip_file.writestr(pptx_filename, pptx_bytes)
                st.session_state.batch_zip_data = zip_buffer.getvalue()

                if len(processed_order) > 1:
                    from fdd_utils.pptx import combine_presentations
                    sources = [
                        io.BytesIO(entity_cache[name]["pptx_download_data"])
                        for name in processed_order
                        if entity_cache.get(name, {}).get("pptx_download_data")
                    ]
                    try:
                        combined_buffer = io.BytesIO()
                        combine_presentations(sources, combined_buffer)
                        st.session_state.batch_combined_pptx = combined_buffer.getvalue()
                    except Exception as exc:
                        logger.warning("Batch combine failed: %s", exc)
                        st.session_state.batch_combined_pptx = None
                else:
                    st.session_state.batch_combined_pptx = None

            progress_bar.empty()
            st.session_state.batch_processing_in_progress = False
            st.session_state.batch_processing_done = True
            st.rerun()


# Initialize
init_session_state()

# Bridge Lab -- experimental, fully isolated from the main upload/process/AI/
# PPTX flow below. Toggled from its own sidebar button; when active it takes
# over the whole page and the normal sidebar/flow is skipped entirely.
render_bridge_lab_toggle()

if st.session_state.get("show_bridge_lab"):
    render_bridge_lab()
    st.stop()

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


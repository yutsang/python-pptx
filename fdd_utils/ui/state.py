from __future__ import annotations

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

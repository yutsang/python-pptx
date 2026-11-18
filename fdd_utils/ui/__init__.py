"""The Streamlit app's pieces; fdd_app.py takes what it shows from here, and the command line takes the extraction phase."""
from __future__ import annotations

from ..workbook.entity import batch_extract_entity_data  # noqa: F401
from .agent_run import run_for_entity as batch_run_book_for_entity  # noqa: F401
from .ai_panel import (  # noqa: F401
    build_selected_pipeline_dfs,
    effective_mappings_from_session,
    extract_account_remarks_context,
    extract_result_text,
    extract_validator_metadata,
    format_dataframe_for_display,
    get_account_dataframe,
    get_financial_account_options,
    has_meaningful_result_text,
    hydrate_nested_agent_outputs,
    render_account_remarks_context,
    render_ai_generation_section,
    render_generated_content,
    _EMPTY_RESULT_MARKERS,
    _result_has_pipeline_content,
    _run_demo_ai,
)
from .pptx_export import (  # noqa: F401
    generate_pptx_presentation,
    render_bridge_lab,
    render_bridge_lab_toggle,
    _bridge_lab_render_preview_chart,
    _bridge_lab_show_block,
)
from .processed import (  # noqa: F401
    filter_reconciliation_display_rows,
    reconciliation_warning_row_count,
    render_data_tables_section,
    render_processed_view,
    render_reconciliation_metrics,
    render_reconciliation_section,
    _RECON_DISPLAY_COLUMN_MAP,
    _render_resolver_diagnostics,
    _render_single_reconciliation_tab,
    _trim_reconciliation_columns_for_display,
)
from .sidebar import (  # noqa: F401
    cleanup_stale_uploads,
    persist_uploaded_workbook,
    render_language_selector,
    render_sidebar_upload,
    _build_model_choices,
    _safe_stem,
)
from .state import (  # noqa: F401
    DEFAULT_SESSION_STATE,
    DELETE_SESSION_KEYS,
    RESET_SESSION_KEYS,
    init_session_state,
    initialize_app_state,
    reset_processing_session_state,
)
from .views import (  # noqa: F401
    build_account_display_dataframe,
    build_entity_selector_model,
    build_processed_display_groups,
    derive_reconciliation_matched_keys,
    describe_statement_period,
    detect_statement_mode,
    explain_reconciliation_selection,
    should_render_preprocess_controls,
    _build_rhs_display_dataframe,
    _normalize_display_label,
)

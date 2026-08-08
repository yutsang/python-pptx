"""Split from the former single-file module. Re-exports every name
the flat module exposed, so `from fdd_utils.ai import X` is unchanged."""
from __future__ import annotations

# namespace parity: the pre-split flat module re-exported everything it
# imported at module level. Reproduced so `from fdd_utils.ai import X`
# resolves for X that were only ever incidental re-exports.
from typing import Any, Dict, List, Optional
from ..financial_common import build_income_statement_period_label, load_required_yaml_file, package_file_path
import re
from typing import Any
from ..financial_common import normalize_chinese_punctuation_in_text
from ..keyword_registry import KNOWN_TRANSLATIONS
import html
import json
from typing import Any, Dict, List
import logging
import os
from datetime import datetime
from typing import Any, Dict, Optional
import yaml
from typing import Any, Dict, Optional, Tuple
import pandas as pd
from ..financial_display_format import add_language_display_columns, prepare_display_dataframe, stringify_display_dataframe
from ..financial_json_converter import df_to_json_str
from ..workbook import build_significant_movements, build_trend_summary, find_mapping_key
from ..financial_common import (
    get_pipeline_result_text,
    load_yaml_file,
    package_file_path,
    visible_descriptions,
)
from ..workbook import INTERNAL_ROW_KEY
import time
import math
import re as _re_client
from typing import Dict, List, Optional, Any
import httpx
from openai import OpenAI, AzureOpenAI
from ..financial_common import package_file_path
import multiprocessing
import threading
import concurrent.futures
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Any, Callable, Dict, List, Optional, Tuple  # noqa: F401


from .config import (  # noqa: F401
    DEFAULT_AGENT_CONFIG,
    DEFAULT_CONFIG_FILENAME,
    DEFAULT_DATA_FORMAT,
    DEFAULT_LOGGING_CONFIG,
    DEFAULT_PROCESSING_CONFIG,
    FDDConfig,
    PROVIDER_REQUIRED_KEYS,
    SUBAGENT_ALIASES,
    WORKBENCH_AVAILABLE_MODELS,
    get_default_config_path,
    get_provider_config,
    get_safe_default_data_format,
    is_provider_ready,
    load_yaml_config,
    normalize_language_code,
    resolve_agent_alias,
    resolve_effective_model_type,
    validate_provider_config,
    _required_keys_for_provider,
)
from .english import (  # noqa: F401
    normalize_english_structure,
    normalize_english_text,
    polish_english_commentary,
    _KNOWN_TRANSLATIONS,
    _MONTH_NAMES,
    _PROPER_NOUN_PATTERN,
    _PROPER_NOUN_SUFFIX_PATTERN,
    _SECTION_LABEL_PATTERNS,
    _enforce_reference_style,
    _iso_to_long_date,
    _k_to_comma_int,
    _protect_chinese_proper_nouns,
    _replace_known_phrases,
    _restore_chinese_proper_nouns,
)
from .validator import (  # noqa: F401
    SourceIndex,
    build_highlighted_commentary_html,
    extract_amounts,
    format_validator_feedback_for_reprompt,
    ground_amounts,
    parse_validator_response,
    segment_clauses,
    strip_thinking,
    verify_commentary,
    _AMT_CUR_PREFIX,
    _AMT_GROUPED,
    _AMT_MILLION,
    _AMT_WAN,
    _AMT_YI,
    _BARE_NUMBER_RE,
    _CAUSAL_RE,
    _CLAUSE_BOUNDARY_CHARS,
    _CLAUSE_END_CHARS,
    _CONF_DEFAULT_REASONING,
    _CONF_DET_DATA_BACKED,
    _CONF_DET_HALLUCINATION,
    _CONF_LLM_FLAG,
    _THINK_BLOCK_RE,
    _THINK_OPEN_TO_END_RE,
    _THINK_STRAY_CLOSE_RE,
    _append_clause_span,
    _attr_text_blob,
    _balanced_brace_slice,
    _combine_verdict,
    _extract_json_payload,
    _fallback_clause_reviews,
    _find_clause_span,
    _has_causal_language,
    _lookup_llm_review,
    _norm_clause_key,
    _normalize_clause_review,
    _normalize_match_text,
    _normalized_index_map,
    _numbers_in_text,
    _repair_json,
    _split_paragraphs,
    _strip_code_fence,
    _to_float,
    _wrap_commentary_html,
)
from .logging import (  # noqa: F401
    PipelineRunLogger,
)
from .prompts import (  # noqa: F401
    PromptEngine,
    PromptStylePack,
    get_prompt_engine,
    resolve_prompt_asset_path,
    _DEFAULT_MAPPINGS_FILE,
    _DEFAULT_MAPPINGS_PATH,
    _DEFAULT_PROMPTS_FILE,
    _DEFAULT_PROMPTS_PATH,
    _PROMPT_ENGINE_CACHE,
)
from .client import (  # noqa: F401
    AIClient,
    _REJECTED_PARAM_RE,
    _extract_rejected_param,
)
from .pipeline import (  # noqa: F401
    RETRIABLE_CLAUSE_CATEGORIES,
    SUBAGENT_SEQUENCE,
    clean_agent_output,
    commentary_asserts_inference,
    count_defective_clauses,
    create_result_shell,
    extract_final_contents,
    load_prompts_and_format,
    map_value_to_component,
    process_single_agent_item,
    run_agent_stage,
    run_ai_pipeline,
    run_ai_pipeline_with_progress,
    run_generator_reprompt,
    save_results,
    set_final_fallbacks,
    _INFERENCE_MARKERS_CHI,
    _INFERENCE_MARKERS_ENG,
    _NOT_REVENUE_KEY_NEEDLES,
    _NO_CAUSE_DISCLAIMERS,
    _PIPELINE_BREAKER,
    _REVENUE_KEY_NEEDLES,
    _StageCircuitBreaker,
    __all__,
    _agent_prompt_kwargs,
    _apply_deterministic_verification,
    _build_deterministic_fallback_bullet,
    _build_peer_context,
    _ensure_clause_reviews_on_final,
    _evaluate_feedback_needed,
    _finalize_agent_content,
    _get_agent_stage_context,
    _get_prompt_manager,
    _notify_stage_progress,
    _resolve_max_workers,
    _run_ai_call,
    _run_feedback_loop_for_key,
    _store_agent_result,
)

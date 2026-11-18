"""Where the model meets the workbook: the configuration and the model dropdown, the client that calls a provider, the evidence
contract each account is read into (evidence_contract.py) and the verifier that holds every sentence to it (claims.py, validator.py).

Submodules are imported where they are used (`from fdd_utils.ai import claims`); only configuration, the client and what the deck builder
takes from the old verifier are re-exported here.
"""
from __future__ import annotations

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
    default_model_choice_key,
    get_default_config_path,
    get_provider_config,
    get_safe_default_data_format,
    is_provider_ready,
    list_model_choices,
    load_yaml_config,
    normalize_language_code,
    resolve_agent_alias,
    resolve_effective_model_type,
    validate_provider_config,
    _required_keys_for_provider,
)
from .client import AIClient  # noqa: F401
from .validator import strip_thinking  # noqa: F401

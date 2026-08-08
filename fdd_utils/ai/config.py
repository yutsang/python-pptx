from __future__ import annotations

"""
FDD configuration module aligned with the HR config interface.
"""

from typing import Any, Dict, List, Optional

from ..financial_common import build_income_statement_period_label, load_required_yaml_file, package_file_path

PROVIDER_REQUIRED_KEYS = {
    "openai": ["api_key", "api_base", "chat_model", "api_version_completion"],
    "local": ["api_base", "api_key", "chat_model"],
    "deepseek": ["api_key", "api_base", "chat_model"],
    # KPMG Workbench gateway (Azure OpenAI-compatible). api_key doubles as the
    # 'Ocp-Apim-Subscription-Key' header value the gateway requires.
    "workbench": ["api_key", "api_base", "api_version", "chat_model"],
}

# Models the internal gateway serves, in UI display order; the FIRST entry is
# the default when the provider is selected without an explicit override.
#
# Deliberately EMPTY here: deployment ids are environment-specific internal
# identifiers and this repository is public. Set workbench.available_models in
# your own config.yml (gitignored), which takes precedence over this list --
# see config.example.yml for the shape.
WORKBENCH_AVAILABLE_MODELS: list[dict] = []

SUBAGENT_ALIASES = {
    # subagent_N (UI/pipeline names) and the canonical N_Name forms both resolve
    # to the canonical name. NOTE: subagent_3 / 3_Refiner is intentionally retained
    # here for config/prompt lookups but is NOT in SUBAGENT_SEQUENCE — the Refiner
    # stage is dormant (active pipeline is Generator -> Auditor -> Validator).
    "subagent_1": "1_Generator",
    "subagent_2": "2_Auditor",
    "subagent_3": "3_Refiner",
    "subagent_4": "4_Validator",
    "1_Generator": "1_Generator",
    "2_Auditor": "2_Auditor",
    "3_Refiner": "3_Refiner",
    "4_Validator": "4_Validator",
}


def resolve_agent_alias(agent_name: str) -> str:
    return SUBAGENT_ALIASES.get(agent_name, agent_name)


# Prompt files (mappings.yml / prompts.yml) key prompts by "Eng" / "Chi", but the
# UI radio stores the language as "Eng" / "Chn". Every consumer that looks up a
# prompt or applies language-specific styling MUST see the normalized code, or
# Chinese runs silently degrade ("No prompts available" / English styling on CN).
# Normalize once, centrally, at every boundary that stores or looks up a language.
def normalize_language_code(language: str) -> str:
    """Map the UI language code 'Chn' to the prompt-file key 'Chi'. Idempotent."""
    return "Chi" if language == "Chn" else language


DEFAULT_CONFIG_FILENAME = "config.yml"
DEFAULT_DATA_FORMAT = "json"
DEFAULT_AGENT_CONFIG = {"temperature": 0.7, "max_tokens": 2000, "top_p": 0.9}
DEFAULT_PROCESSING_CONFIG = {
    "data_format_for_ai": DEFAULT_DATA_FORMAT,
}
DEFAULT_LOGGING_CONFIG = {
    "suppress_http_logs": True,
}


def load_yaml_config(config_path: str) -> Dict[str, Any]:
    return load_required_yaml_file(config_path)


def get_provider_config(config: Dict[str, Any], model_type: str) -> Dict[str, Any]:
    provider = config.get(model_type, {})
    if not provider:
        available = [key for key, value in config.items() if not key.startswith("_") and isinstance(value, dict)]
        raise ValueError(
            f"Model '{model_type}' not found in config. Available: {available}"
        )
    return provider


def _required_keys_for_provider(model_type: str) -> List[str]:
    return PROVIDER_REQUIRED_KEYS.get(model_type, [])


def validate_provider_config(provider: Dict[str, Any], model_type: str) -> None:
    required_keys = _required_keys_for_provider(model_type)
    if not required_keys:
        raise ValueError(f"Invalid model type: {model_type}")
    missing = [key for key in required_keys if not provider.get(key)]
    if missing:
        raise ValueError(f"Missing keys for {model_type}: {missing}")


def is_provider_ready(config: Dict[str, Any], model_type: str) -> bool:
    if model_type not in ("openai", "local", "deepseek", "workbench"):
        return False
    provider = config.get(model_type)
    if not isinstance(provider, dict):
        return False
    for key in _required_keys_for_provider(model_type):
        value = provider.get(key)
        if value is None or (isinstance(value, str) and not str(value).strip()):
            return False
    return True


def resolve_effective_model_type(config: Dict[str, Any], requested: str) -> str:
    if is_provider_ready(config, requested):
        return requested

    preference: List[str] = []
    default_pref = (config.get("default") or {}).get("ai_provider")
    if default_pref and isinstance(default_pref, str):
        preference.append(default_pref.strip())
    # workbench first: if the user configured it, an unready *requested*
    # provider should fall back to Workbench/GPT-5.5 before local/cloud.
    for model_type in ("workbench", "deepseek", "openai", "local"):
        if model_type not in preference:
            preference.append(model_type)

    for model_type in preference:
        if is_provider_ready(config, model_type):
            return model_type

    raise ValueError(
        "No AI provider is fully configured in fdd_utils/config.yml. "
        "Set api_base and api_key (and chat_model) for at least one of: local, deepseek, openai."
    )


def get_default_config_path() -> str:
    return package_file_path(DEFAULT_CONFIG_FILENAME)


def get_safe_default_data_format(
    config_path: Optional[str] = None,
    language: str = "Eng",
    model_type: str = "deepseek",
) -> str:
    try:
        return FDDConfig(
            config_path=config_path,
            language=language,
            model_type=model_type,
        ).get_default_data_format()
    except Exception:
        return DEFAULT_DATA_FORMAT


class FDDConfig:
    """Configuration manager for the financial databook pipeline."""

    _AGENT_DEFAULTS = {
        "1_Generator": {"temperature": 0.7, "max_tokens": 2000, "top_p": 0.9},
        "2_Auditor": {"temperature": 0.3, "max_tokens": 2000, "top_p": 0.9},
        "3_Refiner": {"temperature": 0.5, "max_tokens": 2000, "top_p": 0.9},
        "4_Validator": {"temperature": 0.2, "max_tokens": 2000, "top_p": 0.9},
    }

    def __init__(
        self,
        config_path: Optional[str] = None,
        language: str = "Eng",
        model_type: str = "deepseek",
    ):
        self.language = normalize_language_code(language)
        self.model_type_requested = model_type
        self.config_path = config_path or self._get_default_config_path()
        self.config = self._load_config()
        # If UI requests e.g. local but api_base/api_key are empty, use first ready provider
        self.model_type = resolve_effective_model_type(self.config, model_type)

    def _get_default_config_path(self) -> str:
        return get_default_config_path()

    def _load_config(self) -> Dict[str, Any]:
        return load_yaml_config(self.config_path)

    def get_model_config(self) -> Dict[str, Any]:
        provider = get_provider_config(self.config, self.model_type)
        validate_provider_config(provider, self.model_type)
        return provider

    def resolve_agent_name(self, agent_name: str) -> str:
        return resolve_agent_alias(agent_name)

    def get_agent_config(self, agent_name: str) -> Dict[str, Any]:
        canonical = self.resolve_agent_name(agent_name)
        agents = self.config.get("agents", {})
        agent_config = agents.get(canonical, {})

        if isinstance(agent_config, str):
            agent_config = agents.get(agent_config, {})

        defaults = self._AGENT_DEFAULTS.get(
            canonical,
            DEFAULT_AGENT_CONFIG,
        )
        merged = dict(defaults)
        merged.update(agent_config or {})
        return merged

    def get_processing_config(self) -> Dict[str, Any]:
        processing = dict(DEFAULT_PROCESSING_CONFIG)
        processing.update(self.config.get("processing") or {})
        return processing

    def get_logging_config(self) -> Dict[str, Any]:
        logging_config = dict(DEFAULT_LOGGING_CONFIG)
        logging_config.update(self.config.get("logging") or {})
        return logging_config

    def get_default_data_format(self) -> str:
        data_format = (
            self.get_processing_config().get("data_format_for_ai", DEFAULT_DATA_FORMAT) or DEFAULT_DATA_FORMAT
        ).lower()
        return data_format if data_format in {"markdown", "json"} else DEFAULT_DATA_FORMAT

    def get_debug_mode(self) -> bool:
        return bool(self.get_processing_config().get("debug_mode", False))

    def get_validator_mode(self) -> str:
        """"selective" (default) runs the LLM Validator only on accounts that
        assert a causal claim -- the one thing verify_commentary's
        deterministic number-grounding cannot judge -- and grounds the rest
        for free. "always" restores the previous behaviour of running it on
        every account. Measured: the Validator was 59% of a 964s run, and on
        real data ~80% of accounts had no causal claim for it to review."""
        mode = str(self.get_processing_config().get("validator_mode", "selective") or "selective").lower()
        return mode if mode in {"selective", "always"} else "selective"

    def get_feedback_loop_config(self) -> Dict[str, Any]:
        processing = self.get_processing_config()
        # enabled by default since the gate became specific: it now fires
        # only on a hallucinated figure, never on a supportable inference
        # (see count_defective_clauses). Two real 23-account runs produced
        # ZERO hallucinations, so the expected extra cost of leaving this on
        # is zero -- it only spends tokens on the runs that need it. Total
        # attempts = 1 + max_retries.
        defaults: Dict[str, Any] = {"enabled": True, "max_retries": 2, "unsupported_threshold": 0.3}
        loop_config = processing.get("feedback_loop") or {}
        merged = dict(defaults)
        merged.update(loop_config)
        return merged
# --- end ai/config.py ---

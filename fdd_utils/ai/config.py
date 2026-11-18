from __future__ import annotations

"""
FDD configuration module aligned with the HR config interface.
"""

from typing import Any, Dict, List, Optional

from ..financial_common import load_required_yaml_file, package_file_path

PROVIDER_REQUIRED_KEYS = {
    "openai": ["api_key", "api_base", "chat_model", "api_version_completion"],
    "local": ["api_base", "api_key", "chat_model"],
    "deepseek": ["api_key", "api_base", "chat_model"],
    # Enterprise Workbench gateway (Azure OpenAI-compatible). api_key doubles as the
    # 'Ocp-Apim-Subscription-Key' header value the gateway requires.
    "workbench": ["api_key", "api_base", "api_version", "chat_model"],
    # CN AI Workbench DashScope gateway (OpenAI-compatible chat/embeddings/rerank).
    # Validated via ad-hoc/workbench/test_dashscope_connectivity.py — bearer auth,
    # /v1/chat/completions. embedding + rerank ids live in available_models but
    # are NOT used by the FDD commentary pipeline.
    "dashscope": ["api_key", "api_base", "chat_model"],
}

# Models the internal gateway serves, in UI display order; the FIRST entry is
# the default when the provider is selected without an explicit override.
#
# Deliberately EMPTY here: deployment ids are environment-specific internal
# identifiers and this repository is public. Set workbench.available_models in
# your own config.yml (gitignored), which takes precedence over this list --
# see config.example.yml for the shape.
WORKBENCH_AVAILABLE_MODELS: list[dict] = []

# The app's language radio stores "Eng" / "Chn"; the workbook reader and the agent use "Eng" / "Chi". Normalise at every boundary
# that stores a language.
def normalize_language_code(language: str) -> str:
    """Map the UI language code 'Chn' to the code the rest of the package uses, 'Chi'. Idempotent."""
    return "Chi" if language == "Chn" else language


DEFAULT_CONFIG_FILENAME = "config.yml"
DEFAULT_AGENT_CONFIG = {"temperature": 0.7, "max_tokens": 2000, "top_p": 0.9}
DEFAULT_PROCESSING_CONFIG: Dict[str, Any] = {}
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
    if model_type not in ("openai", "local", "deepseek", "workbench", "dashscope"):
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
    # workbench / dashscope first: enterprise gateways before public cloud / local.
    for model_type in ("workbench", "dashscope", "deepseek", "openai", "local"):
        if model_type not in preference:
            preference.append(model_type)

    for model_type in preference:
        if is_provider_ready(config, model_type):
            return model_type

    raise ValueError(
        "No AI provider is fully configured in fdd_utils/config.yml. "
        "Set api_base and api_key (and chat_model) for at least one of: "
        "local, deepseek, openai, workbench, dashscope."
    )


#: Providers the sidebar can offer, in display order.
_UI_PROVIDERS = (("workbench", "Workbench"), ("dashscope", "DashScope"), ("local", "Local"))


def _provider_model_entries(config: Dict[str, Any], provider: str) -> List[Dict[str, Any]]:
    """A provider's models as {id, label, kind, listed}, from its available_models.

    ``listed: false`` on the provider block hides every model of it from the
    UI; the same key on one entry hides that model alone. Hidden is a UI
    matter only: the provider stays configured and callable (CLI --model,
    compare_models.py), which is how a model is retired from the dropdown
    while it remains the baseline new ones are measured against.
    """
    block = config.get(provider) or {}
    if not isinstance(block, dict):
        return []
    provider_listed = block.get("listed", True) is not False
    raw = block.get("available_models")
    if not raw and provider == "workbench":
        raw = WORKBENCH_AVAILABLE_MODELS
    if provider == "local":
        # One server, one model: the entry is the configured chat_model, and
        # model_name stays None so the client uses the block as it stands. A
        # local block's own available_models was never read by the dropdown
        # and still is not -- ids listed there are not known to be served.
        return [{"id": None, "label": block.get("chat_model") or "Qwen3-32B", "kind": "chat",
                 "listed": provider_listed}]
    entries: List[Dict[str, Any]] = []
    for entry in raw or []:
        if isinstance(entry, dict):
            model_id, label = entry.get("id"), entry.get("label") or entry.get("id")
            kind = str(entry.get("kind") or "chat").lower()
            listed = provider_listed and entry.get("listed", True) is not False
        else:
            model_id, label, kind, listed = str(entry), str(entry), "chat", provider_listed
        if model_id:
            entries.append({"id": str(model_id), "label": str(label), "kind": kind, "listed": listed})
    return entries


def list_model_choices(config: Dict[str, Any], include_unlisted: bool = False) -> List[Dict[str, Any]]:
    """Chat models a user can pick, read from config.yml alone.

    Each choice: key ("provider::id"), model_type, model_name, label, ready,
    listed. Embedding and rerank entries are never choices -- they are not
    commentary models. ``include_unlisted`` adds the models config hides from
    the UI, for tooling that benchmarks against them.
    """
    choices: List[Dict[str, Any]] = []
    for provider, provider_label in _UI_PROVIDERS:
        ready = is_provider_ready(config, provider)
        for entry in _provider_model_entries(config, provider):
            if entry["kind"] != "chat" or not (entry["listed"] or include_unlisted):
                continue
            choices.append({
                "key": f"{provider}::{entry['id'] or 'default'}",
                "model_type": provider,
                "model_name": entry["id"],
                "label": f"{entry['label']} ({provider_label})",
                "ready": ready,
                "listed": entry["listed"],
            })
    return choices


def default_model_choice_key(config: Dict[str, Any], choices: List[Dict[str, Any]]) -> Optional[str]:
    """The choice a new session starts on: default.ai_provider + default.model.

    With only ai_provider set, that provider's chat_model if it is a choice,
    else its first one. With neither (or naming something not on the list),
    the first choice -- which is what the dropdown did before this key existed.
    """
    if not choices:
        return None
    default = config.get("default") or {}
    provider = str(default.get("ai_provider") or "").strip()
    model = str(default.get("model") or default.get("model_name") or "").strip()
    of_provider = [c for c in choices if c["model_type"] == provider]
    if of_provider:
        wanted = model or str((config.get(provider) or {}).get("chat_model") or "")
        for choice in of_provider:
            if wanted and choice["model_name"] == wanted:
                return choice["key"]
        if not model:
            return of_provider[0]["key"]
    return choices[0]["key"]


def get_default_config_path() -> str:
    return package_file_path(DEFAULT_CONFIG_FILENAME)


class FDDConfig:
    """Configuration manager for the financial databook pipeline."""

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

    def get_agent_config(self, agent_name: str) -> Dict[str, Any]:
        """The request settings of a named caller: temperature, max_tokens and top_p by default, with agents.<name> in config.yml over
        them. The agent path's requests are made under the name "fdd_agent"."""
        agents = self.config.get("agents") or {}
        agent_config = agents.get(agent_name, {})
        if isinstance(agent_config, str):
            agent_config = agents.get(agent_config, {})
        merged = dict(DEFAULT_AGENT_CONFIG)
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

    def get_debug_mode(self) -> bool:
        return bool(self.get_processing_config().get("debug_mode", False))
# --- end ai/config.py ---

from __future__ import annotations

# --- begin ai/config.py ---
"""
FDD configuration module aligned with the HR config interface.
"""

from typing import Any, Dict, List, Optional

from .financial_common import build_income_statement_period_label, load_required_yaml_file, package_file_path

PROVIDER_REQUIRED_KEYS = {
    "openai": ["api_key", "api_base", "chat_model", "api_version_completion"],
    "local": ["api_base", "api_key", "chat_model"],
    "deepseek": ["api_key", "api_base", "chat_model"],
}

SUBAGENT_ALIASES = {
    "subagent_1": "1_Generator",
    "subagent_2": "2_Auditor",
    "subagent_3": "3_Refiner",
    "subagent_4": "4_Validator",
    "1_Generator": "1_Generator",
    "2_Auditor": "2_Auditor",
    "3_Refiner": "3_Refiner",
    "4_Validator": "4_Validator",
    # Legacy aliases for backwards compatibility
    "subagent_1": "1_Generator",
    "subagent_2": "2_Auditor",
    "subagent_3": "3_Refiner",
    "subagent_4": "4_Validator",
}


def resolve_agent_alias(agent_name: str) -> str:
    return SUBAGENT_ALIASES.get(agent_name, agent_name)


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
    if model_type not in ("openai", "local", "deepseek"):
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
    for model_type in ("deepseek", "openai", "local"):
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
        self.language = language
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

    def get_feedback_loop_config(self) -> Dict[str, Any]:
        processing = self.get_processing_config()
        defaults: Dict[str, Any] = {"enabled": False, "max_retries": 2, "unsupported_threshold": 0.3}
        loop_config = processing.get("feedback_loop") or {}
        merged = dict(defaults)
        merged.update(loop_config)
        return merged
# --- end ai/config.py ---

# --- begin ai/english.py ---
import re
from typing import Any

from .financial_common import normalize_chinese_punctuation_in_text
from .keyword_registry import KNOWN_TRANSLATIONS

_KNOWN_TRANSLATIONS = KNOWN_TRANSLATIONS

_SECTION_LABEL_PATTERNS = [
    r"(?im)^\s*rhs remarks(?: / reasons)?\s*:\s*",
    r"(?im)^\s*rhs remark summary\s*:\s*",
    r"(?im)^\s*supporting notes\s*:\s*",
    r"(?im)^\s*supporting context\s*:\s*",
    r"(?im)^\s*table-linked remarks\s*:\s*",
    r"(?im)^\s*table context observations\s*:\s*",
    r"(?im)^\s*supplemental side-column context\s*:\s*",
    r"(?im)^\s*user remarks / reprompt guidance\s*:\s*",
    r"(?im)^\s*user guidance\s*:\s*",
]


def _replace_known_phrases(text: str) -> str:
    normalized = text
    for source, target in sorted(_KNOWN_TRANSLATIONS.items(), key=lambda item: len(item[0]), reverse=True):
        normalized = normalized.replace(source, target)
    return normalized


_PROPER_NOUN_SUFFIX_PATTERN = (
    r"(?:股份有限公司|有限责任公司|有限公司|集团|公司|"
    r"合伙企业(?:（有限合伙）|\(有限合伙\))|有限合伙)"
)
_PROPER_NOUN_PATTERN = re.compile(
    rf"([\u4e00-\u9fff][\u4e00-\u9fffA-Za-z0-9（）()\-\u00b7、，,\s]{{1,80}}?{_PROPER_NOUN_SUFFIX_PATTERN})"
)


def _protect_chinese_proper_nouns(text: str) -> tuple[str, Dict[str, str]]:
    preserved: Dict[str, str] = {}

    def repl(match: re.Match[str]) -> str:
        placeholder = f"__PROPER_NOUN_{len(preserved)}__"
        preserved[placeholder] = match.group(1)
        return placeholder

    return _PROPER_NOUN_PATTERN.sub(repl, text), preserved


def _restore_chinese_proper_nouns(text: str, preserved: Dict[str, str]) -> str:
    restored = text
    for placeholder, original in preserved.items():
        restored = restored.replace(placeholder, original)
    return restored


def normalize_english_text(text: Any) -> Any:
    if not isinstance(text, str):
        return text
    normalized = normalize_chinese_punctuation_in_text(text, preserve_sentence_stop=True)
    normalized = normalized.replace("（", "(").replace("）", ")").replace("：", ":")
    normalized = normalized.replace("。", ".")
    normalized = re.sub(r"(\d{4})年(\d{1,2})月(\d{1,2})日", r"\1-\2-\3", normalized)
    normalized = re.sub(r"(\d{4})年(\d{1,2})月", r"\1-\2", normalized)
    normalized = re.sub(r"(\d{4})年", r"\1", normalized)
    normalized, preserved_proper_nouns = _protect_chinese_proper_nouns(normalized)
    normalized = _replace_known_phrases(normalized)
    normalized = re.sub(
        r"(?<=[A-Za-z])(?=(increase|decrease|movement|balance|amount|reason|difference|summary|project|customer|supplier|construction|deposit|interest|taxes))",
        " ",
        normalized,
    )
    normalized = re.sub(r"(?<=[A-Za-z])Co\., Ltd\.", " Co., Ltd.", normalized)
    normalized = normalized.replace(". ,", ".,")
    normalized = re.sub(r"\s+([,;:.!?])", r"\1", normalized)
    normalized = re.sub(r"(?<=\d)\.\s+(?=\d)", ".", normalized)
    normalized = re.sub(r"(?<=\d),\s+(?=\d)", ",", normalized)
    normalized = re.sub(r"([;:!?])(?!\s|$)", r"\1 ", normalized)
    normalized = re.sub(r"(?<!\d)\.(?!\s|$|\d)", ". ", normalized)
    normalized = re.sub(r"(?<!\d),(?!\s|$|\d)", ", ", normalized)
    normalized = re.sub(r"\(\s+", "(", normalized)
    normalized = re.sub(r"\s+\)", ")", normalized)
    normalized = _restore_chinese_proper_nouns(normalized, preserved_proper_nouns)
    normalized = re.sub(r"\s{2,}", " ", normalized)
    return normalized.strip()


def normalize_english_structure(value: Any) -> Any:
    if isinstance(value, dict):
        return {
            normalize_english_text(str(key)): normalize_english_structure(item)
            for key, item in value.items()
        }
    if isinstance(value, list):
        return [normalize_english_structure(item) for item in value]
    if isinstance(value, tuple):
        return tuple(normalize_english_structure(item) for item in value)
    if isinstance(value, str):
        return normalize_english_text(value)
    return value


def polish_english_commentary(text: str) -> str:
    polished = normalize_english_text(text or "")
    for pattern in _SECTION_LABEL_PATTERNS:
        polished = re.sub(pattern, "", polished)
    polished = re.sub(r"(?i)^including:\s*", "Including ", polished)
    polished = polished.replace("Co. , Ltd. .", "Co., Ltd.")
    polished = polished.replace("Co. , Ltd.", "Co., Ltd.")
    polished = polished.replace("\n- ", " ").replace("\n", " ")
    polished = re.sub(r"\s{2,}", " ", polished)
    return polished.strip(" ;")
# --- end ai/english.py ---

# --- begin ai/validator.py ---
"""
Utilities for parsing validator clause annotations and rendering highlights.
"""


import html
import json
import re
from typing import Any, Dict, List


def _strip_code_fence(text: str) -> str:
    match = re.search(r"```(?:json)?\s*(.*?)```", text or "", flags=re.DOTALL | re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return str(text or "").strip()


def _extract_json_payload(text: str) -> Dict[str, Any] | None:
    candidate = _strip_code_fence(text)
    try:
        parsed = json.loads(candidate)
        return parsed if isinstance(parsed, dict) else None
    except json.JSONDecodeError:
        pass

    start = candidate.find("{")
    end = candidate.rfind("}")
    if start >= 0 and end > start:
        try:
            parsed = json.loads(candidate[start : end + 1])
            return parsed if isinstance(parsed, dict) else None
        except json.JSONDecodeError:
            return None
    return None


def _normalize_clause_review(item: Dict[str, Any]) -> Dict[str, Any]:
    clause = str(item.get("clause") or "").strip()
    reason = str(item.get("reason") or "").strip()
    supported_value = item.get("supported")
    if isinstance(supported_value, str):
        supported = supported_value.strip().lower() in {"true", "yes", "supported", "1"}
    else:
        supported = bool(supported_value)
    # Parse category: data-backed, reasoning, or hallucination
    raw_category = str(item.get("category") or "").strip().lower()
    if raw_category in ("data-backed", "reasoning", "hallucination"):
        category = raw_category
    elif supported:
        category = "data-backed"
    else:
        category = "hallucination"
    return {
        "clause": clause,
        "supported": supported,
        "category": category,
        "reason": reason,
    }


def format_validator_feedback_for_reprompt(clause_reviews: List[Dict[str, Any]], language: str) -> str:
    """Format unsupported clause reasons into concise feedback for the generator reprompt."""
    unsupported = [r for r in (clause_reviews or []) if isinstance(r, dict) and not r.get("supported")]
    if not unsupported:
        return ""
    if language == "Chi":
        header = "验证器标记了以下不支持的内容需要修正:\n"
        template = "- 分句: \"{clause}\" — 问题: {reason}"
    else:
        header = "The validator flagged the following unsupported clauses for correction:\n"
        template = '- Clause: "{clause}" — Issue: {reason}'
    items = [
        template.format(clause=str(r.get("clause", ""))[:120], reason=str(r.get("reason", ""))[:200])
        for r in unsupported[:5]
    ]
    return header + "\n".join(items)


def parse_validator_response(raw_text: str, fallback_content: str = "") -> Dict[str, Any]:
    """
    Parse structured validator output.

    Expected shape:
    {
      "final_content": "...",
      "clause_reviews": [{"clause": "...", "supported": true, "reason": "..."}]
    }
    """
    parsed = _extract_json_payload(raw_text)
    if not parsed:
        return {
            "final_content": str(fallback_content or raw_text or "").strip(),
            "clause_reviews": [],
            "raw_response": str(raw_text or ""),
        }

    final_content = str(
        parsed.get("final_content")
        or parsed.get("content")
        or fallback_content
        or ""
    ).strip()
    clause_reviews = []
    for item in parsed.get("clause_reviews") or []:
        if not isinstance(item, dict):
            continue
        normalized = _normalize_clause_review(item)
        if normalized["clause"]:
            clause_reviews.append(normalized)

    return {
        "final_content": final_content,
        "clause_reviews": clause_reviews,
        "raw_response": str(raw_text or ""),
    }


def _split_paragraphs(text: str) -> List[str]:
    paragraphs = [part.strip() for part in re.split(r"\n\s*\n", text) if part.strip()]
    if paragraphs:
        return paragraphs
    return [text.strip()] if text.strip() else []


def _wrap_commentary_html(text: str, *, escape_html: bool) -> str:
    paragraphs = _split_paragraphs(text)
    paragraph_html = "".join(
        f"<p>{html.escape(paragraph) if escape_html else paragraph}</p>"
        for paragraph in paragraphs
    )
    return f'<div class="fdd-final-commentary">{paragraph_html}</div>'


def _normalize_match_text(text: str) -> str:
    normalized = re.sub(r"\s+", " ", str(text or "").strip())
    normalized = normalized.strip(" \t\r\n\"'`")
    return normalized


def _normalized_index_map(text: str) -> tuple[str, List[int]]:
    normalized_chars: List[str] = []
    index_map: List[int] = []
    previous_was_space = False

    for index, char in enumerate(str(text or "")):
        if char.isspace():
            if previous_was_space:
                continue
            normalized_chars.append(" ")
            index_map.append(index)
            previous_was_space = True
            continue
        normalized_chars.append(char)
        index_map.append(index)
        previous_was_space = False

    normalized = "".join(normalized_chars).strip()
    if not normalized:
        return "", []

    start_trim = len("".join(normalized_chars)) - len("".join(normalized_chars).lstrip())
    end_trim = len("".join(normalized_chars).rstrip())
    return normalized, index_map[start_trim:end_trim]


def _find_clause_span(text: str, clause: str, cursor: int) -> tuple[int, int]:
    if not clause:
        return (-1, -1)

    direct_index = text.find(clause, cursor)
    if direct_index >= 0:
        return (direct_index, direct_index + len(clause))
    direct_index = text.find(clause)
    if direct_index >= 0:
        return (direct_index, direct_index + len(clause))

    normalized_text, index_map = _normalized_index_map(text)
    normalized_clause = _normalize_match_text(clause)
    if not normalized_text or not normalized_clause:
        return (-1, -1)

    normalized_cursor = 0
    if cursor > 0 and index_map:
        normalized_cursor = next(
            (idx for idx, original_index in enumerate(index_map) if original_index >= cursor),
            len(index_map),
        )

    normalized_index = normalized_text.find(normalized_clause, normalized_cursor)
    if normalized_index < 0:
        normalized_index = normalized_text.find(normalized_clause)
    if normalized_index < 0:
        return (-1, -1)

    start = index_map[normalized_index]
    end_idx = normalized_index + len(normalized_clause) - 1
    if end_idx >= len(index_map):
        return (-1, -1)
    end = index_map[end_idx] + 1
    return (start, end)


def _find_next_clause_index(text: str, clause: str, cursor: int) -> int:
    if not clause:
        return -1
    return text.find(clause, cursor)


def build_highlighted_commentary_html(final_content: str, clause_reviews: List[Dict[str, Any]]) -> str:
    """
    Render final commentary HTML with unsupported clauses highlighted.
    Uses category-specific CSS classes: fdd-hallucination-clause (yellow) and fdd-reasoning-clause (orange).
    """
    text = str(final_content or "")
    flagged_reviews = [
        review for review in (clause_reviews or [])
        if isinstance(review, dict) and review.get("clause") and not bool(review.get("supported"))
    ]

    if not flagged_reviews:
        return _wrap_commentary_html(text, escape_html=True)

    rendered_parts: List[str] = []
    cursor = 0
    unmatched_reviews: List[Dict[str, Any]] = []
    for review in flagged_reviews:
        clause = str(review.get("clause") or "")
        if not clause:
            continue
        start, end = _find_clause_span(text, clause, cursor)
        if start < 0 or end <= start:
            unmatched_reviews.append(review)
            continue
        rendered_parts.append(html.escape(text[cursor:start]))
        category = str(review.get("category") or "hallucination").lower()
        reason = str(review.get("reason") or "This clause may not be fully supported by the provided data.")
        category_label = "Hallucination" if category == "hallucination" else "Reasoning"
        tooltip = f"[{category_label}] {reason}"
        css_class = "fdd-hallucination-clause" if category == "hallucination" else "fdd-reasoning-clause"
        rendered_parts.append(
            '<span class="{css_class}" title="{title}">{content}</span>'.format(
                css_class=css_class,
                title=html.escape(tooltip, quote=True),
                content=html.escape(text[start:end]),
            )
        )
        cursor = end
    rendered_parts.append(html.escape(text[cursor:]))

    highlighted = "".join(rendered_parts)
    rendered_html = _wrap_commentary_html(highlighted, escape_html=False)
    if not unmatched_reviews:
        return rendered_html

    note_items = "".join(
        "<li><strong>{clause}</strong> [{category}]<br>{reason}</li>".format(
            clause=html.escape(str(review.get("clause") or "")),
            category=html.escape(str(review.get("category") or "hallucination")),
            reason=html.escape(
                str(review.get("reason") or "This clause may not be fully supported by the provided data.")
            ),
        )
        for review in unmatched_reviews
    )
    note_block = (
        '<div class="fdd-validator-notes">'
        "<p>Validator flagged these unsupported clauses, but they could not be matched exactly for inline highlighting:</p>"
        f"<ul>{note_items}</ul>"
        "</div>"
    )
    return rendered_html + note_block
# --- end ai/validator.py ---

# --- begin ai/logging.py ---
"""
Run logging for the FDD AI pipeline.
"""

import logging
import os
from datetime import datetime
from typing import Any, Dict, Optional

import yaml


class PipelineRunLogger:
    """Unified logger for an FDD AI processing run."""

    def __init__(self, log_dir: str = "fdd_utils/logs", output_dir: str = "fdd_utils/output", debug_mode: bool = False):
        self.log_dir = log_dir
        self.output_dir = output_dir
        self.debug_mode = debug_mode
        os.makedirs(log_dir, exist_ok=True)
        os.makedirs(output_dir, exist_ok=True)

        self.run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.run_folder = os.path.join(log_dir, f"run_{self.run_id}")
        os.makedirs(self.run_folder, exist_ok=True)

        self.log_file = os.path.join(self.run_folder, "processing.log")
        self.log_data_file = os.path.join(self.run_folder, "data.yml")
        self.results_file = os.path.join(self.run_folder, "results.yml")

        self.logger = logging.getLogger(f"ContentGeneration_{self.run_id}")
        self.logger.setLevel(logging.DEBUG if debug_mode else logging.INFO)
        self.logger.handlers = []
        self.logger.propagate = False

        file_handler = logging.FileHandler(self.log_file, encoding="utf-8")
        file_handler.setLevel(logging.DEBUG)

        formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
        file_handler.setFormatter(formatter)
        self.logger.addHandler(file_handler)

        if debug_mode:
            self.debug_log_file = os.path.join(self.run_folder, "debug.log")
            debug_handler = logging.FileHandler(self.debug_log_file, encoding="utf-8")
            debug_handler.setLevel(logging.DEBUG)
            debug_handler.setFormatter(formatter)
            self.logger.addHandler(debug_handler)

        self.run_data: Dict[str, Any] = {
            "run_id": self.run_id,
            "start_time": datetime.now().isoformat(),
            "debug_mode": debug_mode,
            "agents_executed": [],
            "processing_results": {},
        }
        self.logger.info("=== Started new AI processing run: %s (debug_mode=%s) ===", self.run_id, debug_mode)

    def _display_name(self, agent_name: str) -> str:
        names = {
            "subagent_1": "Generator",
            "subagent_2": "Auditor",
            "subagent_3": "Refiner",
            "subagent_4": "Validator",
        }
        return names.get(agent_name, agent_name)

    def log_debug(self, category: str, mapping_key: str, message: str, data: Any = None) -> None:
        if not self.debug_mode:
            return
        self.logger.debug("[DEBUG][%s] %s: %s", category, mapping_key, message)
        if data is not None:
            data_str = str(data)
            if len(data_str) > 4000:
                data_str = data_str[:4000] + "... [truncated]"
            self.logger.debug("[DEBUG][%s] %s: DATA:\n%s", category, mapping_key, data_str)

    def log_agent_start(self, agent_name: str, mapping_key: str):
        self.logger.info("[%s] Processing: %s", self._display_name(agent_name), mapping_key)

    def log_agent_complete(
        self,
        agent_name: str,
        mapping_key: str,
        result: Dict[str, Any],
        system_prompt: str = "",
        user_prompt: str = "",
        prompt_context: Optional[Dict[str, Any]] = None,
    ):
        duration = result.get("duration", 0)
        tokens = result.get("tokens_used", 0)
        content = result.get("content", "")
        prompt_length = result.get("prompt_length", len(system_prompt) + len(user_prompt))
        output_length = result.get("output_length", len(content))
        prompt_tokens = result.get("prompt_tokens")
        completion_tokens = result.get("completion_tokens")
        total_tokens = result.get("total_tokens")
        estimated_prompt_tokens = result.get("estimated_prompt_tokens")
        estimated_output_tokens = result.get("estimated_output_tokens")
        expected_max_output_tokens = result.get("expected_max_output_tokens")
        model = result.get("model") or "-"
        model_type = result.get("model_type") or result.get("provider") or result.get("mode", "ai")
        token_usage_source = result.get("token_usage_source", "unknown")

        self.logger.info(
            "[%s] Processed: %s | Duration: %.2fs | Model: %s/%s | Prompt chars: %s | Output chars: %s | Tokens used: %s | Prompt tokens: %s | Completion tokens: %s | Total tokens: %s | Estimated prompt tokens: %s | Estimated output tokens: %s | Expected max output tokens: %s | Token source: %s",
            self._display_name(agent_name),
            mapping_key,
            duration,
            model_type,
            model,
            prompt_length,
            output_length,
            tokens,
            prompt_tokens if prompt_tokens is not None else "-",
            completion_tokens if completion_tokens is not None else "-",
            total_tokens if total_tokens is not None else "-",
            estimated_prompt_tokens if estimated_prompt_tokens is not None else "-",
            estimated_output_tokens if estimated_output_tokens is not None else "-",
            expected_max_output_tokens if expected_max_output_tokens is not None else "-",
            token_usage_source,
        )

        prompt_context = prompt_context or {}
        rhs_summary = prompt_context.get("rhs_remark_summary") or []
        supporting_notes = prompt_context.get("supporting_notes") or []
        table_linked_remarks = prompt_context.get("table_linked_remarks") or []
        user_comment = str(prompt_context.get("user_comment") or "").strip()
        context_fragments = [
            f"supporting_notes={len(supporting_notes)}",
            f"rhs_rows={int(prompt_context.get('rhs_remark_count') or 0)}",
            f"rhs_summary={len(rhs_summary)}",
            f"table_linked_remarks={len(table_linked_remarks)}",
        ]
        if rhs_summary:
            context_fragments.append(
                "rhs_preview=" + " || ".join(str(item).strip() for item in rhs_summary[:3] if str(item).strip())
            )
        if user_comment:
            context_fragments.append(f"user_comment={user_comment[:240]}")
        if prompt_context.get("has_previous_output"):
            context_fragments.append("reprompt_baseline=yes")
        self.logger.info(
            "[%s] Prompt context: %s | %s",
            self._display_name(agent_name),
            mapping_key,
            " | ".join(fragment for fragment in context_fragments if fragment),
        )

        self.run_data["processing_results"].setdefault(mapping_key, {})
        self.run_data["processing_results"][mapping_key][agent_name] = {
            "duration": duration,
            "tokens_used": tokens,
            "mode": result.get("mode", "ai"),
            "model_type": model_type,
            "model": model,
            "provider": result.get("provider", model_type),
            "agent_name": result.get("agent_name", agent_name),
            "language": result.get("language"),
            "temperature": result.get("temperature"),
            "max_tokens": result.get("max_tokens"),
            "top_p": result.get("top_p"),
            "frequency_penalty": result.get("frequency_penalty"),
            "presence_penalty": result.get("presence_penalty"),
            "timestamp": datetime.now().isoformat(),
            "system_prompt": system_prompt or "",
            "user_prompt": user_prompt or "",
            "prompt_context": prompt_context or {},
            "output": content,
            "system_prompt_length": result.get("system_prompt_length", len(system_prompt)),
            "user_prompt_length": result.get("user_prompt_length", len(user_prompt)),
            "prompt_length": prompt_length,
            "content_length": output_length,
            "output_length": output_length,
            "prompt_tokens": prompt_tokens,
            "completion_tokens": completion_tokens,
            "total_tokens": total_tokens,
            "estimated_system_prompt_tokens": result.get("estimated_system_prompt_tokens"),
            "estimated_user_prompt_tokens": result.get("estimated_user_prompt_tokens"),
            "estimated_prompt_tokens": estimated_prompt_tokens,
            "estimated_output_tokens": estimated_output_tokens,
            "estimated_total_tokens": result.get("estimated_total_tokens"),
            "expected_max_output_tokens": expected_max_output_tokens,
            "token_usage_source": token_usage_source,
        }

    def log_error(self, agent_name: str, mapping_key: str, error: Exception):
        self.logger.error("[%s] Error processing %s: %s", self._display_name(agent_name), mapping_key, error)
        self.run_data["processing_results"].setdefault(mapping_key, {})
        self.run_data["processing_results"][mapping_key][agent_name] = {
            "error": str(error),
            "timestamp": datetime.now().isoformat(),
        }

    def finalize(self, results: Optional[Dict[str, Dict[str, str]]] = None):
        self.run_data["end_time"] = datetime.now().isoformat()

        total_duration = 0
        total_tokens = 0
        total_items = len(self.run_data["processing_results"])

        for _key, agents in self.run_data["processing_results"].items():
            for _agent, data in agents.items():
                total_duration += data.get("duration", 0)
                total_tokens += data.get("tokens_used", 0)

        self.run_data["summary"] = {
            "total_items_processed": total_items,
            "total_duration_seconds": total_duration,
            "total_tokens_used": total_tokens,
            "total_prompt_length_chars": sum(
                data.get("prompt_length", 0)
                for agents in self.run_data["processing_results"].values()
                for data in agents.values()
            ),
            "total_output_length_chars": sum(
                data.get("output_length", data.get("content_length", 0))
                for agents in self.run_data["processing_results"].values()
                for data in agents.values()
            ),
            "total_estimated_prompt_tokens": sum(
                data.get("estimated_prompt_tokens", 0)
                for agents in self.run_data["processing_results"].values()
                for data in agents.values()
            ),
            "total_estimated_output_tokens": sum(
                data.get("estimated_output_tokens", 0)
                for agents in self.run_data["processing_results"].values()
                for data in agents.values()
            ),
            "agents_used": list(
                {
                    agent
                    for agents in self.run_data["processing_results"].values()
                    for agent in agents.keys()
                }
            ),
        }

        with open(self.log_data_file, "w", encoding="utf-8") as file:
            yaml.dump(self.run_data, file, default_flow_style=False, allow_unicode=True)

        if results:
            with open(self.results_file, "w", encoding="utf-8") as file:
                yaml.dump(results, file, default_flow_style=False, allow_unicode=True)

        self.logger.info("=== Completed AI processing run: %s ===", self.run_id)
        self.logger.info(
            "Summary: %s items, %.2fs, %s tokens",
            total_items,
            total_duration,
            total_tokens,
        )
# --- end ai/logging.py ---

# --- begin ai/prompts.py ---
"""
Prompt loading and rendering helpers for the FDD pipeline.
"""


import json
import logging
import os
from typing import Any, Dict, Optional, Tuple

import pandas as pd

from .financial_display_format import add_language_display_columns, prepare_display_dataframe, stringify_display_dataframe
from .financial_json_converter import df_to_json_str
from .workbook import build_significant_movements, build_trend_summary, find_mapping_key
from .financial_common import (
    get_pipeline_result_text,
    load_yaml_file,
    package_file_path,
    visible_descriptions,
)
from .workbook import INTERNAL_ROW_KEY

_DEFAULT_PROMPTS_FILE = "fdd_utils/prompts.yml"
_DEFAULT_MAPPINGS_FILE = "fdd_utils/mappings.yml"
_DEFAULT_PROMPTS_PATH = package_file_path("prompts.yml")
_DEFAULT_MAPPINGS_PATH = package_file_path("mappings.yml")
_PROMPT_ENGINE_CACHE: Dict[Tuple[str, str], "PromptEngine"] = {}


class PromptStylePack:
    """Small style pack mirroring the HR separation of prompts and style rules."""

    def __init__(self, language: str = "Eng"):
        self.language = language

    def language_instruction(self) -> str:
        if self.language == "Chi":
            return "这是中文数据簿。输出必须保持为纯中文，不要混用英文占位符、要点或 Markdown 样式。"
        return (
            "This is an English databook. Prefer clear English wording with no bullet lists or Markdown styling. "
            "If a proper noun or specialized term cannot be translated confidently, keep the original Chinese "
            "instead of forcing pinyin."
        )

    def common_formatting_rules(self) -> str:
        if self.language == "Chi":
            return "使用自然段落表达。不要使用要点、粗体或元评论。保留专业财务语气，并确保日期和金额表达一致。"
        return (
            "Use natural paragraph-style reporting. Avoid bullets, bold formatting, and meta-commentary. "
            "Keep a professional financial tone and consistent date and amount formatting."
        )

    def common_data_rules(self, data_format: str) -> str:
        if data_format == "json":
            if self.language == "Chi":
                return "数据以 JSON 提供，数值和单位已按展示口径预处理，请严格按字段和值直接使用。"
            return (
                "The data is provided as JSON. Values and units are already normalized for reporting, "
                "so use the fields and values exactly as provided."
            )

        if self.language == "Chi":
            return "数据以 Markdown 表格提供，金额已按展示口径格式化，请直接使用，不要再次换算。"
        return (
            "The data is provided as a markdown table. Amounts are already formatted for reporting, "
            "so use them directly without reconversion."
        )


def resolve_prompt_asset_path(path: Optional[str], default_file: str, default_path: str) -> str:
    if not path or path in {default_file, default_path}:
        return default_path
    if os.path.isabs(path):
        return path
    return os.path.join(os.getcwd(), path)


def get_prompt_engine(
    prompts_path: Optional[str] = None,
    mappings_path: Optional[str] = None,
) -> "PromptEngine":
    resolved_prompts = resolve_prompt_asset_path(
        prompts_path,
        _DEFAULT_PROMPTS_FILE,
        _DEFAULT_PROMPTS_PATH,
    )
    resolved_mappings = resolve_prompt_asset_path(
        mappings_path,
        _DEFAULT_MAPPINGS_FILE,
        _DEFAULT_MAPPINGS_PATH,
    )
    cache_key = (resolved_prompts, resolved_mappings)
    if cache_key not in _PROMPT_ENGINE_CACHE:
        _PROMPT_ENGINE_CACHE[cache_key] = PromptEngine(
            prompts_path=resolved_prompts,
            mappings_path=resolved_mappings,
        )
    return _PROMPT_ENGINE_CACHE[cache_key]


class PromptEngine:
    """Centralized prompt accessor aligned with the HR prompt handling pattern."""

    def __init__(
        self,
        prompts_path: Optional[str] = None,
        mappings_path: Optional[str] = None,
    ):
        self.prompts_path = prompts_path or package_file_path("prompts.yml")
        self.mappings_path = mappings_path or package_file_path("mappings.yml")
        self._prompts_data = None
        self._mappings_data = None
        self.logger = logging.getLogger(self.__class__.__name__)

    @staticmethod
    def normalize_agent_name(agent_name: str) -> str:
        return resolve_agent_alias(agent_name)

    @property
    def prompts_data(self) -> Dict[str, Any]:
        if self._prompts_data is None:
            self._prompts_data = load_yaml_file(self.prompts_path)
        return self._prompts_data

    @property
    def mappings_data(self) -> Dict[str, Any]:
        if self._mappings_data is None:
            self._mappings_data = load_yaml_file(self.mappings_path)
        return self._mappings_data

    def get_mapping_component(self, mapping_key: str, component: Optional[str] = None) -> Any:
        resolved_key = find_mapping_key(mapping_key, self.mappings_data)
        if resolved_key:
            data = self.mappings_data.get(resolved_key, {})
            return data.get(component) if component else resolved_key
        return None

    def resolve_mapping_key(self, mapping_key: str) -> str:
        return self.get_mapping_component(mapping_key) or mapping_key

    def _fallback_mapping_section(self, mapping_key: str) -> Optional[str]:
        if mapping_key in self.mappings_data:
            return None
        if "_general_dynamic_mapping" in self.mappings_data:
            return "_general_dynamic_mapping"
        return None

    def get_agent_defaults(self, agent_name: str, language: str) -> Tuple[str, str]:
        agent_key = self.normalize_agent_name(agent_name)
        if agent_key == "1_Generator":
            generic_prompts = self.mappings_data.get("_default_subagent_1", {}).get(language, {})
            return generic_prompts.get("system_prompt", ""), generic_prompts.get("user_prompt", "")

        prompt_data = self.prompts_data.get(agent_key, {}).get(language, {})
        return prompt_data.get("system_prompt", ""), prompt_data.get("user_prompt", "")

    def get_prompt_pair(self, agent_name: str, language: str, mapping_key: str) -> Tuple[str, str]:
        agent_key = self.normalize_agent_name(agent_name)
        resolved_mapping_key = self.resolve_mapping_key(mapping_key)

        if agent_key == "1_Generator":
            default_system_prompt, default_user_prompt = self.get_agent_defaults(agent_name, language)

            account_data = self.mappings_data.get(resolved_mapping_key, {})
            if not account_data:
                fallback_section = self._fallback_mapping_section(mapping_key)
                if fallback_section:
                    account_data = self.mappings_data.get(fallback_section, {})
            account_prompts = account_data.get("subagent_1_prompts", {}).get(language, {})
            system_prompt = (account_prompts.get("system_prompt") or "").strip() or default_system_prompt
            user_prompt = (account_prompts.get("user_prompt") or "").strip() or default_user_prompt
            return system_prompt, user_prompt

        return self.get_agent_defaults(agent_name, language)

    def _build_markdown_prompt_payload(self, df: pd.DataFrame) -> Dict[str, str]:
        df_display = prepare_display_dataframe(
            df,
            drop_columns=(INTERNAL_ROW_KEY,),
        )
        rendered = stringify_display_dataframe(df_display).to_markdown(index=False).strip()
        return {"financial_figure": rendered, "financial_data": rendered}

    @staticmethod
    def _normalize_prompt_dataframe(df: Optional[pd.DataFrame], language: str) -> Optional[pd.DataFrame]:
        if language != "Eng" or not isinstance(df, pd.DataFrame) or df.empty:
            return df
        normalized_df = df.copy()
        normalized_df.columns = [
            column if str(column) == INTERNAL_ROW_KEY else normalize_english_text(str(column))
            for column in normalized_df.columns
        ]
        for column in normalized_df.columns:
            if str(column) == INTERNAL_ROW_KEY:
                continue
            series = normalized_df[column]
            if pd.api.types.is_numeric_dtype(series):
                continue
            normalized_df[column] = series.apply(
                lambda value: normalize_english_text(value) if isinstance(value, str) else value
            )
        normalized_df.attrs.update(df.attrs)
        return normalized_df

    @staticmethod
    def _normalize_prompt_value(value: Any, language: str) -> Any:
        if language != "Eng":
            return value
        return normalize_english_structure(value)

    def _filter_prompt_analysis_df(self, df: pd.DataFrame) -> Optional[pd.DataFrame]:
        analysis_df = df.attrs.get("prompt_analysis_df")
        if not isinstance(analysis_df, pd.DataFrame) or analysis_df.empty:
            return None
        visible_rows = visible_descriptions(df)
        if not visible_rows or len(analysis_df.columns) == 0:
            return analysis_df
        first_col = analysis_df.columns[0]
        filtered = analysis_df[
            analysis_df[first_col].astype(str).map(lambda value: value.strip() in visible_rows)
        ].copy()
        return filtered if not filtered.empty else analysis_df

    def _filter_adjacent_detail_rows(self, df: pd.DataFrame) -> list[Dict[str, Any]]:
        adjacent_detail_rows = df.attrs.get("adjacent_detail_rows") or []
        if not adjacent_detail_rows:
            return []
        visible_rows = visible_descriptions(df)
        if not visible_rows:
            return adjacent_detail_rows
        filtered = [
            row for row in adjacent_detail_rows
            if str(row.get("Description", "")).strip() in visible_rows
        ]
        return filtered or adjacent_detail_rows

    @staticmethod
    def _should_skip_prompt_metadata_key(key_text: str, include_description: bool = False) -> bool:
        return (
            key_text == INTERNAL_ROW_KEY
            or (include_description and key_text == "Description")
            or key_text.endswith("| table_header")
            or key_text.endswith("| indicative_adjusted_row")
            or key_text.endswith("| date_row")
        )

    @staticmethod
    def _prompt_ready_adjacent_detail_rows(adjacent_detail_rows: list[Dict[str, Any]]) -> list[Dict[str, Any]]:
        cleaned_rows: list[Dict[str, Any]] = []
        for row in adjacent_detail_rows:
            if not isinstance(row, dict):
                continue
            cleaned_row: Dict[str, Any] = {}
            for key, value in row.items():
                key_text = str(key)
                if PromptEngine._should_skip_prompt_metadata_key(key_text):
                    continue
                cleaned_row[key_text] = value
            if cleaned_row:
                cleaned_rows.append(cleaned_row)
        return cleaned_rows

    @staticmethod
    def _table_linked_remarks(df: Optional[pd.DataFrame]) -> list[Dict[str, Any]]:
        if not isinstance(df, pd.DataFrame):
            return []
        remarks = df.attrs.get("table_linked_remarks") or []
        return [remark for remark in remarks if isinstance(remark, dict)]

    def _build_analysis_prompt_df(self, df: pd.DataFrame) -> Optional[pd.DataFrame]:
        analysis_df = self._filter_prompt_analysis_df(df)
        if not isinstance(analysis_df, pd.DataFrame) or analysis_df.empty:
            return analysis_df
        effective_analysis_df = analysis_df.copy()
        if INTERNAL_ROW_KEY in effective_analysis_df.columns:
            effective_analysis_df = effective_analysis_df.drop(columns=[INTERNAL_ROW_KEY])

        return effective_analysis_df

    @staticmethod
    def _format_analysis_prompt_df(df: Optional[pd.DataFrame], report_language: str) -> Optional[pd.DataFrame]:
        if not isinstance(df, pd.DataFrame) or df.empty or not report_language:
            return df
        formatted_analysis_df = add_language_display_columns(df.copy(), report_language)
        formatted_analysis_df.attrs.update(df.attrs)
        return formatted_analysis_df

    def _rhs_guidance_block(self, adjacent_detail_rows: list[Dict[str, Any]], language: str) -> str:
        if not adjacent_detail_rows:
            return ""
        if language == "Chi":
            return (
                "表格右侧1-5列的补充备注/工作说明已在财务数据载荷末尾提供。"
                "仅在其与数据一致时将其视为可引用的补充背景或原因。"
            )
        return (
            "Supplemental context extracted from the 1-5 side columns of the schedule is included in the financial data payload. "
            "Use it only where it is supported by the evidence, and absorb it into normal report sentences rather than repeating section labels."
        )

    @staticmethod
    def _summarize_rhs_remarks(adjacent_detail_rows: list[Dict[str, Any]], language: str) -> list[str]:
        if not adjacent_detail_rows:
            return []

        summaries: list[str] = []
        seen: set[str] = set()
        for row in adjacent_detail_rows:
            if not isinstance(row, dict):
                continue
            description = str(row.get("Description") or "").strip()
            remark_parts: list[str] = []
            for key, value in row.items():
                key_text = str(key)
                if PromptEngine._should_skip_prompt_metadata_key(key_text, include_description=True):
                    continue
                if isinstance(value, (int, float)):
                    continue
                text = str(value or "").strip()
                if not text:
                    continue
                if key_text == description:
                    continue
                remark_parts.append(text)

            unique_parts: list[str] = []
            seen_parts: set[str] = set()
            for part in remark_parts:
                if part in seen_parts:
                    continue
                seen_parts.add(part)
                unique_parts.append(part)

            if not unique_parts:
                continue
            if language == "Chi":
                summary = f"{description}: " + "；".join(unique_parts[:2]) if description else "；".join(unique_parts[:2])
            else:
                summary = f"{description}: " + "; ".join(unique_parts[:2]) if description else "; ".join(unique_parts[:2])
            if summary not in seen:
                seen.add(summary)
                summaries.append(summary)
        return summaries[:5]

    @staticmethod
    def _remarks_weight_instruction(
        *,
        has_rhs_remarks: bool,
        has_supporting_notes: bool,
        has_user_comment: bool,
        statement_type: str,
        language: str,
    ) -> str:
        if not any((has_rhs_remarks, has_supporting_notes, has_user_comment)):
            return ""
        normalized_statement_type = str(statement_type or "").strip().upper()
        if language == "Chi":
            if normalized_statement_type == "BS":
                return (
                    "若补充备注、右侧备注或用户备注提供了与数据一致的背景、原因、限制、差异解释或工作说明，"
                    "对于资产负债表科目，应优先以这些备注作为原因判断和定性解释的主要依据，数值则主要用于支持趋势、余额与重要性判断。"
                    "若备注中明确写明折旧/摊销方法、使用年限、残值率、坏账计提依据等会计政策细节，且与当前科目及数据一致，可在评论中吸收概括。"
                    "财务尽调评论应在不脱离数据的前提下，吸收这些备注并概括成简洁、顾问式的说明，而不是忽略它们。"
                )
            if normalized_statement_type == "IS":
                return (
                    "若补充备注、右侧备注或用户备注提供了与数据一致的背景、原因、限制、差异解释或工作说明，"
                    "对于利润表科目，应将数值趋势/重大变动分析与这些备注结合起来共同解释当期表现，而不是只依赖单一来源。"
                    "若备注中明确写明与当前科目相关的会计政策、成本构成、收入确认、折旧/摊销、坏账计提或其他解释性细节，且与数据一致，可在评论中吸收概括。"
                    "财务尽调评论应在不脱离数据的前提下，吸收这些备注并概括成简洁、顾问式的说明，而不是忽略它们。"
                )
            return (
                "若补充备注、右侧备注或用户备注提供了与数据一致的背景、原因、限制、差异解释或工作说明，请提高其在评论中的权重。"
                "若备注中明确写明折旧/摊销方法、使用年限、残值率、坏账计提依据等会计政策细节，且与当前科目及数据一致，可在评论中吸收概括。"
                "财务尽调评论应在不脱离数据的前提下，吸收这些备注并概括成简洁、顾问式的说明，而不是忽略它们。"
            )
        if normalized_statement_type == "BS":
            return (
                "Where supporting notes, supplemental side-column context, or user remarks provide data-consistent context such as reasons, restrictions, differences, "
                "or working explanations, treat those remarks as a required part of the reasoning on balance-sheet items. Build the reasoning from the combination of data, "
                "cross-period trend, and supported remarks, while using the remarks as the primary basis for qualitative explanation. Use the numbers mainly to confirm the trend, "
                "latest balance, composition, and materiality rather than inventing causes from the figures alone. If remarks explicitly state account-relevant accounting-policy "
                "details such as depreciation or amortization method, useful life, residual value, or provisioning basis, you may reflect them when they are clearly supported. "
                "As an FDD consultant, absorb and summarize those points into report-ready sentences rather than copying labels, dumping note headers, or leaving them unused. "
                "If the remarks contain material supported explanations, differences, restrictions, or drivers, make them visible in the commentary rather than treating them as optional background."
            )
        if normalized_statement_type == "IS":
            return (
                "Where supporting notes, supplemental side-column context, or user remarks provide data-consistent context such as reasons, restrictions, differences, or working "
                "explanations, treat them as a required part of the reasoning on income-statement items. Build the reasoning from the combination of data, cross-period trend, remarks, "
                "and broader numeric analysis such as movement, scale, mix, and reasonableness. Use both the figures and the remarks together to explain the period performance, rather "
                "than relying on only one source. If remarks explicitly state account-relevant accounting-policy details such as depreciation or amortization "
                "method, useful life, residual value, provisioning basis, or other explanation for the period movement, you may reflect them when they are clearly supported. As an FDD "
                "consultant, absorb and summarize those points into report-ready sentences rather than copying labels, dumping note headers, or leaving them unused. "
                "If the remarks contain material supported explanations, drivers, or validation-relevant observations, make them visible in the commentary rather than treating them as optional background."
            )
        return (
            "Where supporting notes, supplemental side-column context, or user remarks provide data-consistent context such as reasons, restrictions, differences, "
            "or working explanations, give them higher weight in the commentary and treat them as an active part of the reasoning rather than optional background. "
            "Build the commentary from data, trend, and supported remarks together. If remarks explicitly state account-relevant accounting-policy details such as "
            "depreciation or amortization method, useful life, residual value, or provisioning basis, you may reflect them in the commentary when they are clearly supported. "
            "As an FDD consultant, absorb and summarize those points "
            "into report-ready sentences rather than copying labels, dumping note headers, or leaving them unused. If the remarks contain material supported observations, "
            "make them visible in the commentary rather than treating them as optional background."
        )

    @staticmethod
    def _user_guidance_instruction(user_comment: str, language: str) -> str:
        guidance = str(user_comment or "").strip()
        if not guidance:
            return ""
        if language == "Chi":
            return (
                "用户备注/重提示指引已提供。若其与数据、备注或表格内容一致，请将其视为明确写作指引，"
                "并在输出的措辞、重点、结构或补充说明中体现。若其中有与数据不一致之处，不要照搬，应仅保留有依据的部分。"
            )
        return (
            "User remarks / reprompt guidance has been provided. Where it is consistent with the data, notes, and table context, "
            "treat it as explicit writing guidance and reflect it in the wording, emphasis, structure, or clarifications of the output. "
            "Do not follow any part that is not supported by the provided evidence."
        )

    @staticmethod
    def _period_reference_guidance(df: Optional[pd.DataFrame], language: str) -> str:
        integrity = (df.attrs if isinstance(df, pd.DataFrame) else {}).get("integrity") or {}
        attrs = df.attrs if isinstance(df, pd.DataFrame) else {}
        statement_type = str(integrity.get("statement_type") or "").strip().upper()
        effective_date = str(integrity.get("effective_date") or "").strip()
        annualization_months = attrs.get("annualization_months")
        if annualization_months in (None, ""):
            annualization_months = integrity.get("annualization_months")
        fiscal_year_end_month = integrity.get("fiscal_year_end_month")
        fiscal_year_end_day = integrity.get("fiscal_year_end_day")

        if language == "Chi":
            if statement_type == "BS":
                return f"这是资产负债表科目。描述最新余额时，请使用“截至{effective_date}”之类的时点表述，不要写成期间表述。"
            if statement_type == "IS":
                period_label = build_income_statement_period_label(
                    effective_date,
                    months=annualization_months if isinstance(annualization_months, (int, float)) else None,
                    fiscal_year_end_month=fiscal_year_end_month if isinstance(fiscal_year_end_month, (int, float)) else None,
                    fiscal_year_end_day=fiscal_year_end_day if isinstance(fiscal_year_end_day, (int, float)) else None,
                    language="Chi",
                )
                return (
                    f"这是利润表科目。描述目标期间时，请使用“于{period_label}期间”或“在{period_label}内”等期间表述，"
                    f"不要写成“截至{effective_date}止”或时点余额表述。"
                )
            return "请根据科目属性正确区分时点表述与期间表述。"

        if statement_type == "BS":
            return f"This is a balance-sheet item. Refer to the latest amount with point-in-time wording such as 'as at {effective_date}', not period-flow wording."
        if statement_type == "IS":
            period_label = build_income_statement_period_label(
                effective_date,
                months=annualization_months if isinstance(annualization_months, (int, float)) else None,
                fiscal_year_end_month=fiscal_year_end_month if isinstance(fiscal_year_end_month, (int, float)) else None,
                fiscal_year_end_day=fiscal_year_end_day if isinstance(fiscal_year_end_day, (int, float)) else None,
                language="Eng",
            )
            return (
                f"This is an income-statement item. Refer to the target period with flow wording such as 'during {period_label}' "
                f"or 'during the period', not 'for the period ended {effective_date}'. Do not describe it with balance-sheet wording such as 'as at'."
            )
        return "Use point-in-time wording for balance-sheet style data and period-flow wording for income-statement style data."

    @staticmethod
    def _append_markdown_section(rendered: str, label: str, body: str) -> str:
        body = str(body or "").strip()
        if not body:
            return rendered
        return f"{rendered}\n\n{label}:\n{body}"

    def _append_markdown_table_section(self, rendered: str, label: str, rows: list[Dict[str, Any]]) -> str:
        table_df = pd.DataFrame(rows)
        if table_df.empty:
            return rendered
        return self._append_markdown_section(rendered, label, table_df.to_markdown(index=False).strip())

    def _build_financial_prompt_payload(
        self,
        df: Optional[pd.DataFrame],
        mapping_key: str,
        language: str,
        data_format: str,
        user_comment: str = "",
    ) -> Dict[str, str]:
        if df is None or df.empty:
            return {}
        analysis_df = self._build_analysis_prompt_df(df)
        analysis_label = str(df.attrs.get("prompt_analysis_label") or "All indicative adjusted periods").strip()
        supporting_notes = [str(note).strip() for note in (df.attrs.get("supporting_notes") or []) if str(note).strip()]
        adjacent_detail_rows = self.filter_adjacent_detail_rows(df)
        table_linked_remarks = self.table_linked_remarks(df)
        rhs_remark_summary = self.summarize_rhs_remarks(adjacent_detail_rows, language)
        format_language = language or str(df.attrs.get("report_language") or "").strip()
        formatted_analysis_df = self._format_analysis_prompt_df(analysis_df, format_language)
        normalized_prompt_df = self._normalize_prompt_dataframe(df, language)
        normalized_analysis_df = self._normalize_prompt_dataframe(formatted_analysis_df, language)
        normalized_supporting_notes = self._normalize_prompt_value(supporting_notes, language)
        prompt_ready_adjacent_detail_rows = self._prompt_ready_adjacent_detail_rows(adjacent_detail_rows)
        normalized_adjacent_detail_rows = self._normalize_prompt_value(prompt_ready_adjacent_detail_rows, language)
        normalized_table_linked_remarks = self._normalize_prompt_value(table_linked_remarks, language)
        normalized_rhs_remark_summary = self._normalize_prompt_value(rhs_remark_summary, language)
        normalized_user_comment = self._normalize_prompt_value(str(user_comment or "").strip(), language)
        normalized_analysis_label = self._normalize_prompt_value(analysis_label, language)
        normalized_mapping_key = self._normalize_prompt_value(mapping_key, language)
        trend_summary = build_trend_summary(analysis_df) if isinstance(analysis_df, pd.DataFrame) and not analysis_df.empty else {}
        significant_movements = (
            build_significant_movements(analysis_df)
            if isinstance(analysis_df, pd.DataFrame) and not analysis_df.empty
            else []
        )
        trend_summary = self._normalize_prompt_value(trend_summary, language)
        significant_movements = self._normalize_prompt_value(significant_movements, language)
        integrity = df.attrs.get("integrity") or {}
        latest_source_period = str(integrity.get("effective_date") or "").strip()
        target_period = latest_source_period or (str(df.columns[1]).strip() if len(df.columns) > 1 else "")
        annualization_months = df.attrs.get("annualization_months")
        if annualization_months in (None, ""):
            annualization_months = integrity.get("annualization_months")
        fiscal_year_end_month = integrity.get("fiscal_year_end_month")
        fiscal_year_end_day = integrity.get("fiscal_year_end_day")
        target_period_label = (
            build_income_statement_period_label(
                target_period,
                months=annualization_months if isinstance(annualization_months, (int, float)) else None,
                fiscal_year_end_month=fiscal_year_end_month if isinstance(fiscal_year_end_month, (int, float)) else None,
                fiscal_year_end_day=fiscal_year_end_day if isinstance(fiscal_year_end_day, (int, float)) else None,
                language=language,
            )
            if str(integrity.get("statement_type") or "").strip().upper() == "IS"
            else target_period
        )
        target_period_guidance = {
            "target_period": target_period,
            "target_period_label": target_period_label,
            "latest_source_period": latest_source_period,
            "instruction": (
                "Use the target_period as the main reporting period for the latest balance or latest period statement. "
                "For income-statement items, prefer period labels such as target_period_label and use 'during', not 'for the period ended'. "
                "Use all earlier indicative-adjusted periods for trend, comparison, reasonableness checks, and significant movement analysis."
            ),
        }
        target_period_guidance = self._normalize_prompt_value(target_period_guidance, language)

        if data_format == "json":
            payload = json.loads(
                df_to_json_str(
                    normalized_prompt_df if isinstance(normalized_prompt_df, pd.DataFrame) else df,
                    table_name=normalized_mapping_key,
                    language=language,
                    text_normalizer=normalize_english_text if language == "Eng" else None,
                )
            )
            payload["reporting_focus"] = target_period_guidance
            if isinstance(normalized_analysis_df, pd.DataFrame) and not normalized_analysis_df.empty:
                payload["analysis_periods"] = json.loads(
                    df_to_json_str(
                        normalized_analysis_df,
                        table_name=normalized_analysis_label,
                        language=language,
                        text_normalizer=normalize_english_text if language == "Eng" else None,
                    )
                )
            if trend_summary:
                payload["trend_summary"] = trend_summary
            if significant_movements:
                payload["significant_movements"] = significant_movements
            if normalized_supporting_notes:
                payload["supporting_context"] = normalized_supporting_notes
            if normalized_adjacent_detail_rows:
                payload["supplemental_side_column_context"] = normalized_adjacent_detail_rows
            if normalized_table_linked_remarks:
                payload["table_context_observations"] = normalized_table_linked_remarks
            if normalized_rhs_remark_summary:
                payload["supplemental_context_summary"] = normalized_rhs_remark_summary
            if normalized_user_comment:
                payload["user_guidance"] = [normalized_user_comment]
                payload["user_guidance_instruction"] = (
                    "Treat these user remarks as explicit writing or reprompt guidance only to the extent they are supported by the provided data, notes, and remarks."
                    if language == "Eng"
                    else "将这些用户备注视为明确的写作/重提示指引，但仅可在其与提供的数据、备注及说明一致时采用。"
                )
            rendered = json.dumps(payload, ensure_ascii=False, indent=2)
        else:
            rendered = self._build_markdown_prompt_payload(
                normalized_prompt_df if isinstance(normalized_prompt_df, pd.DataFrame) else df
            )["financial_data"]
            focus_label = "Reporting focus" if language == "Eng" else "报告重点"
            focus_lines = [
                f"- Target period: {target_period}" if language == "Eng" else f"- 目标期间: {target_period}",
                (
                    f"- Preferred narrative period label: {target_period_label}"
                    if language == "Eng"
                    else f"- 推荐叙述期间标签: {target_period_label}"
                ),
                (
                    f"- Latest source indicative-adjusted period: {latest_source_period}"
                    if language == "Eng"
                    else f"- 最新示意性调整后源期间: {latest_source_period}"
                ),
                (
                    "- Use the target period for the latest balance / latest period statement. For income-statement items, prefer 'during' wording with the preferred narrative period label rather than 'for the period ended'. Use earlier indicative-adjusted periods for trend, comparison, cross-check, and significant movement analysis."
                    if language == "Eng"
                    else "- 以目标期间作为最新余额/最新期间表述的基础。若为利润表科目，优先使用推荐叙述期间标签并采用“于...期间/在...内”表达，而不是“截至...止期间”。同时使用更早的示意性调整后期间进行趋势、比较、交叉检查及重大变动分析。"
                ),
            ]
            rendered = self._append_markdown_section(rendered, focus_label, "\n".join(focus_lines))
            if isinstance(normalized_analysis_df, pd.DataFrame) and not normalized_analysis_df.empty:
                analysis_block = normalized_analysis_df.to_markdown(index=False).strip()
                rendered = self._append_markdown_section(rendered, normalized_analysis_label, analysis_block)
            if trend_summary:
                trend_lines = [
                    f"- {key}: {value}"
                    for key, value in trend_summary.items()
                    if value not in (None, "", [], {})
                ]
                if trend_lines:
                    trend_label = "Trend summary" if language == "Eng" else "趋势摘要"
                    rendered = self._append_markdown_section(rendered, trend_label, "\n".join(trend_lines))
            if significant_movements:
                change_label = "Significant movements" if language == "Eng" else "重大变动"
                rendered = self._append_markdown_table_section(rendered, change_label, significant_movements)

        if data_format != "json" and normalized_supporting_notes:
            notes_label = "Supporting context" if language == "Eng" else "补充备注"
            notes_block = "\n".join(f"- {note}" for note in normalized_supporting_notes)
            rendered = self._append_markdown_section(rendered, notes_label, notes_block)

        if data_format != "json" and normalized_adjacent_detail_rows:
            details_label = "Supplemental side-column context" if language == "Eng" else "右侧备注/原因"
            rendered = self._append_markdown_table_section(rendered, details_label, normalized_adjacent_detail_rows)

        if data_format != "json" and normalized_table_linked_remarks:
            table_linked_label = "Table context observations" if language == "Eng" else "表格关联备注"
            rendered = self._append_markdown_table_section(rendered, table_linked_label, normalized_table_linked_remarks)

        if data_format != "json" and normalized_rhs_remark_summary:
            summary_label = "Supplemental context summary" if language == "Eng" else "右侧备注摘要"
            summary_block = "\n".join(f"- {item}" for item in normalized_rhs_remark_summary)
            rendered = self._append_markdown_section(rendered, summary_label, summary_block)

        if data_format != "json" and normalized_user_comment:
            comment_label = "User guidance" if language == "Eng" else "用户备注 / 重提示指引"
            rendered = self._append_markdown_section(rendered, comment_label, f"- {normalized_user_comment}")

        if language == "Eng":
            rendered = normalize_english_text(rendered)

        return {"financial_figure": rendered, "financial_data": rendered}

    def _safe_format(self, template: str, format_params: Dict[str, Any]) -> str:
        if not template:
            return template
        try:
            return template.format(**format_params)
        except KeyError as exc:
            self.logger.warning("Missing prompt key %s. Available keys: %s", exc, list(format_params.keys()))
            return template

    def render_prompt(
        self,
        agent_name: str,
        language: str,
        mapping_key: str,
        df: Optional[pd.DataFrame] = None,
        data_format: str = "markdown",
        **kwargs,
    ) -> Tuple[str, str]:
        system_prompt, user_prompt_template = self.get_prompt_pair(agent_name, language, mapping_key)
        style_pack = PromptStylePack(language)
        dynamic_mapping_context = {}
        if isinstance(df, pd.DataFrame):
            dynamic_mapping_context = dict(df.attrs.get("dynamic_mapping_context") or {})
        normalized_kwargs = self._normalize_prompt_value(kwargs, language)
        format_params = {
            "key": self._normalize_prompt_value(mapping_key, language),
            "language": language,
            "accounting_nature": (
                self._normalize_prompt_value(
                    kwargs.get("accounting_nature")
                    or dynamic_mapping_context.get("accounting_nature")
                    or dynamic_mapping_context.get("category")
                    or "",
                    language,
                )
            ),
            "language_instruction": style_pack.language_instruction(),
            "common_formatting": style_pack.common_formatting_rules(),
            "common_data_rules": style_pack.common_data_rules(data_format),
            "period_reference_guidance": self._period_reference_guidance(df, language),
            "rhs_guidance_block": self._rhs_guidance_block(
                self.filter_adjacent_detail_rows(df) if isinstance(df, pd.DataFrame) else [],
                language,
            ),
            "remarks_weight_instruction": self._remarks_weight_instruction(
                has_rhs_remarks=bool(self.filter_adjacent_detail_rows(df) if isinstance(df, pd.DataFrame) else []),
                has_supporting_notes=bool((df.attrs.get("supporting_notes") or []) if isinstance(df, pd.DataFrame) else []),
                has_user_comment=bool(str(kwargs.get("user_comment", "")).strip()),
                statement_type=((df.attrs.get("integrity") or {}).get("statement_type") if isinstance(df, pd.DataFrame) else ""),
                language=language,
            ),
            "user_guidance_instruction": self._user_guidance_instruction(kwargs.get("user_comment", ""), language),
            **normalized_kwargs,
        }
        format_params.update(
            self._build_financial_prompt_payload(
                df=df,
                mapping_key=mapping_key,
                language=language,
                data_format=data_format,
                user_comment=kwargs.get("user_comment", ""),
            )
        )

        if self.normalize_agent_name(agent_name) == "1_Generator":
            resolved_mapping_key = self.resolve_mapping_key(mapping_key)
            patterns = self.get_mapping_component(
                resolved_mapping_key,
                component="patterns",
            )
            if patterns is None:
                fallback_section = self._fallback_mapping_section(mapping_key)
                if fallback_section:
                    patterns = self.mappings_data.get(fallback_section, {}).get("patterns")
            format_params["patterns"] = patterns or ""

        rendered_system_prompt = self._safe_format(system_prompt, format_params)
        rendered_user_prompt = self._safe_format(user_prompt_template, format_params)

        if self.normalize_agent_name(agent_name) == "1_Generator":
            previous_content = str(kwargs.get("previous_content") or "").strip()
            if previous_content:
                if language == "Chi":
                    rendered_user_prompt = (
                        f"{rendered_user_prompt}\n\n"
                        "已验证旧评论（请在保留数据支持内容的基础上按新的用户备注进行改写，而不是完全重写方向）：\n"
                        f"{previous_content}\n\n"
                        "请将上述旧评论视为待修订底稿。优先保留其中仍被当前数据、备注及右侧说明支持的内容，并结合用户最新指引进行定向修改。"
                    )
                else:
                    rendered_user_prompt = (
                        f"{rendered_user_prompt}\n\n"
                        "Existing validated commentary to revise (treat this as the draft to update rather than starting from scratch):\n"
                        f"{previous_content}\n\n"
                        "Use the existing validated commentary as the baseline draft. Keep the parts that are still supported by the current data, remarks, and notes, and revise it directionally based on the latest user guidance."
                    )

        return rendered_system_prompt, rendered_user_prompt

    def filter_adjacent_detail_rows(self, df: pd.DataFrame) -> list[Dict[str, Any]]:
        return self._filter_adjacent_detail_rows(df)

    def table_linked_remarks(self, df: Optional[pd.DataFrame]) -> list[Dict[str, Any]]:
        return self._table_linked_remarks(df)

    def summarize_rhs_remarks(self, adjacent_detail_rows: list[Dict[str, Any]], language: str) -> list[str]:
        return self._summarize_rhs_remarks(adjacent_detail_rows, language)

    def build_prompt_context_snapshot(
        self,
        df: Optional[pd.DataFrame],
        language: str = "Eng",
        user_comment: str = "",
        previous_output: str = "",
    ) -> Dict[str, Any]:
        if not isinstance(df, pd.DataFrame):
            return {}
        attrs = df.attrs or {}
        integrity = attrs.get("integrity") or {}
        supporting_notes = attrs.get("supporting_notes") or []
        adjacent_detail_rows = self.filter_adjacent_detail_rows(df)
        table_linked_remarks = self.table_linked_remarks(df)
        rhs_remark_summary = self.summarize_rhs_remarks(adjacent_detail_rows, language)
        return {
            "sheet_name": integrity.get("sheet_name"),
            "statement_type": integrity.get("statement_type"),
            "effective_date": integrity.get("effective_date"),
            "selected_variant": attrs.get("selected_variant"),
            "prompt_analysis_label": attrs.get("prompt_analysis_label"),
            "supporting_notes": supporting_notes,
            "supporting_notes_count": len(supporting_notes),
            "adjacent_detail_rows": adjacent_detail_rows,
            "rhs_remark_count": len(adjacent_detail_rows),
            "rhs_remark_summary": rhs_remark_summary,
            "table_linked_remarks": table_linked_remarks,
            "table_linked_remarks_count": len(table_linked_remarks),
            "user_comment": str(user_comment or "").strip(),
            "has_previous_output": bool(str(previous_output or "").strip()),
            "previous_output_excerpt": str(previous_output or "").strip()[:500],
        }
# --- end ai/prompts.py ---

# --- begin ai/client.py ---
import os
import time
import math
from typing import Dict, List, Optional, Any
import httpx
from openai import OpenAI, AzureOpenAI
import logging

from .financial_common import package_file_path

class AIClient:
    """
    Reusable AI helper class supporting multiple agents and models.
    Supports: content generation, value checks, content refinement, and formatting checks.
    """
    _logged_fallbacks = set()
    
    def __init__(
        self, 
        model_type: str = 'deepseek',
        agent_name: str = 'agent_1',
        language: str = 'Eng',
        use_heuristic: bool = False,
        config_path: Optional[str] = None
    ):
        """
        Initialize AIClient with specified model and agent configuration.
        
        Args:
            model_type: Type of model ('openai', 'local', 'deepseek')
            agent_name: Name of the agent ('agent_1', 'agent_2', 'agent_3', 'agent_4')
            language: Language for prompts ('Eng' or 'Chi')
            use_heuristic: Whether to use heuristic mode instead of AI
            config_path: Path to config file (optional)
        """
        self.model_type_requested = model_type
        self.agent_name = agent_name
        self.language = language
        self.use_heuristic = use_heuristic
        
        # Load configuration (may resolve e.g. local -> deepseek if local has no api_base/api_key)
        self.config_path = config_path or package_file_path('config.yml')
        self.config_manager = FDDConfig(
            config_path=self.config_path,
            language=language,
            model_type=model_type,
        )
        self.prompt_engine = get_prompt_engine()
        self.model_type = self.config_manager.model_type
        self.full_config = self.config_manager.config
        self.data_format = self.config_manager.get_default_data_format()
        if self.use_heuristic:
            self.config_details = self.full_config.get(self.model_type, {})
        else:
            self.config_details = self.config_manager.get_model_config()

        agent_config = self.config_manager.get_agent_config(agent_name)
        self.temperature = agent_config.get('temperature')
        self.max_tokens = agent_config.get('max_tokens')
        self.top_p = agent_config.get('top_p')
        self.frequency_penalty = agent_config.get('frequency_penalty')
        self.presence_penalty = agent_config.get('presence_penalty')
        
        # Initialize client only if not using heuristic mode
        if not self.use_heuristic:
            self.validate_config()
            self.client, self.model = self.initialize_client()
        else:
            self.client = None
            self.model = None
            
        # Setup logging
        self.logger = logging.getLogger(f'AIClient.{agent_name}')
        self._configure_external_logging()
        if self.model_type_requested != self.model_type:
            fallback_key = (agent_name, self.model_type_requested, self.model_type)
            if fallback_key not in self._logged_fallbacks:
                self.logger.debug(
                    "Requested model '%s' is not configured; using '%s' from config.",
                    self.model_type_requested,
                    self.model_type,
                )
                self._logged_fallbacks.add(fallback_key)
        self.logger.debug(f"Initialized {agent_name} with temperature={self.temperature}, max_tokens={self.max_tokens}")

    def load_config(self) -> Dict:
        """Load configuration from YAML file."""
        return self.config_manager.config

    def validate_config(self):
        """Validate required configuration keys for the model type."""
        self.config_manager.get_model_config()
        return True

    def _configure_external_logging(self):
        """Reduce noisy client logs unless explicitly enabled."""
        suppress_logs = self.config_manager.get_logging_config().get('suppress_http_logs', True)
        debug_enabled = os.getenv('HR_DEBUG') == '1' or os.getenv('FDD_DEBUG') == '1'
        if suppress_logs and not debug_enabled:
            logging.getLogger('httpx').setLevel(logging.WARNING)
            logging.getLogger('openai').setLevel(logging.WARNING)

    def initialize_client(self):
        """Initialize the appropriate AI client based on model type."""
        if self.model_type == 'openai':
            client = AzureOpenAI(
                api_key=self.config_details['api_key'], 
                base_url=self.config_details['api_base'],
                api_version=self.config_details['api_version_completion'],
                http_client=httpx.Client(verify=False)
            )
            model = self.config_details['chat_model']
            
        elif self.model_type == 'local':
            client = OpenAI(
                base_url=self.config_details['api_base'],
                api_key=self.config_details['api_key'],
                http_client=httpx.Client(verify=False)
            )
            model = self.config_details['chat_model']
            
        elif self.model_type == 'deepseek':
            client = OpenAI(
                base_url=self.config_details['api_base'],
                api_key=self.config_details['api_key'],
                http_client=httpx.Client(verify=False)
            )
            model = self.config_details['chat_model']
        else:
            raise ValueError(f"Invalid model type: {self.model_type}")
        
        return client, model

    def get_agent_settings(self, agent_name: Optional[str] = None) -> Dict[str, Any]:
        """Get config-backed settings for a specific pipeline agent."""
        return self.config_manager.get_agent_config(agent_name or self.agent_name)
    
    def load_prompts(self, agent_name: Optional[str] = None) -> tuple:
        agent = agent_name or self.agent_name
        try:
            return self.prompt_engine.get_agent_defaults(agent, self.language)
        except Exception as e:
            self.logger.error(f"Error loading prompts for {agent}: {e}")
            return '', ''

    @staticmethod
    def _estimate_text_tokens(text: Optional[str]) -> int:
        normalized = str(text or "").strip()
        if not normalized:
            return 0
        return max(1, math.ceil(len(normalized) / 4))

    def _build_logging_metadata(
        self,
        *,
        user_prompt: str,
        system_prompt: str,
        content: str,
        duration: float,
        mode: str,
        temperature: Optional[float] = None,
        max_tokens: Optional[int] = None,
        top_p: Optional[float] = None,
        frequency_penalty: Optional[float] = None,
        presence_penalty: Optional[float] = None,
        prompt_tokens: Optional[int] = None,
        completion_tokens: Optional[int] = None,
        total_tokens: Optional[int] = None,
    ) -> Dict[str, Any]:
        system_prompt = str(system_prompt or "")
        user_prompt = str(user_prompt or "")
        content = str(content or "")

        estimated_system_prompt_tokens = self._estimate_text_tokens(system_prompt)
        estimated_user_prompt_tokens = self._estimate_text_tokens(user_prompt)
        estimated_prompt_tokens = estimated_system_prompt_tokens + estimated_user_prompt_tokens
        estimated_output_tokens = self._estimate_text_tokens(content)
        estimated_total_tokens = estimated_prompt_tokens + estimated_output_tokens

        resolved_total_tokens = (
            total_tokens
            if total_tokens is not None
            else (
                (prompt_tokens or 0) + (completion_tokens or 0)
                if prompt_tokens is not None or completion_tokens is not None
                else estimated_total_tokens
            )
        )

        if prompt_tokens is not None or completion_tokens is not None or total_tokens is not None:
            token_usage_source = "provider_usage"
        elif mode == "heuristic":
            token_usage_source = "heuristic_estimate"
        else:
            token_usage_source = "estimated"

        return {
            "mode": mode,
            "model_type": self.model_type,
            "model": self.model,
            "provider": self.model_type,
            "agent_name": self.agent_name,
            "language": self.language,
            "duration": duration,
            "temperature": temperature if temperature is not None else self.temperature,
            "max_tokens": max_tokens if max_tokens is not None else self.max_tokens,
            "top_p": top_p if top_p is not None else self.top_p,
            "frequency_penalty": (
                frequency_penalty if frequency_penalty is not None else self.frequency_penalty
            ),
            "presence_penalty": (
                presence_penalty if presence_penalty is not None else self.presence_penalty
            ),
            "system_prompt_length": len(system_prompt),
            "user_prompt_length": len(user_prompt),
            "prompt_length": len(system_prompt) + len(user_prompt),
            "output_length": len(content),
            "prompt_tokens": prompt_tokens,
            "completion_tokens": completion_tokens,
            "total_tokens": total_tokens,
            "estimated_system_prompt_tokens": estimated_system_prompt_tokens,
            "estimated_user_prompt_tokens": estimated_user_prompt_tokens,
            "estimated_prompt_tokens": estimated_prompt_tokens,
            "estimated_output_tokens": estimated_output_tokens,
            "estimated_total_tokens": estimated_total_tokens,
            "expected_max_output_tokens": max_tokens if max_tokens is not None else self.max_tokens,
            "tokens_used": resolved_total_tokens,
            "token_usage_source": token_usage_source,
        }

    @staticmethod
    def _coerce_message_content(content: Any) -> str:
        if content is None:
            return ""
        if isinstance(content, str):
            return content
        if isinstance(content, list):
            parts: list[str] = []
            for item in content:
                if isinstance(item, str):
                    parts.append(item)
                    continue
                if isinstance(item, dict):
                    text_value = item.get("text") or item.get("content")
                    if isinstance(text_value, str):
                        parts.append(text_value)
                    continue
                text_value = getattr(item, "text", None) or getattr(item, "content", None)
                if isinstance(text_value, str):
                    parts.append(text_value)
            return "".join(parts)
        return str(content)

    @classmethod
    def _extract_content_from_choice(cls, choice: Any) -> str:
        if choice is None:
            return ""
        if isinstance(choice, dict):
            message = choice.get("message")
            if isinstance(message, dict):
                return cls._coerce_message_content(message.get("content"))
            if message is not None:
                return cls._coerce_message_content(getattr(message, "content", None))
            delta = choice.get("delta")
            if isinstance(delta, dict):
                return cls._coerce_message_content(delta.get("content"))
            if delta is not None:
                return cls._coerce_message_content(getattr(delta, "content", None))
            return cls._coerce_message_content(choice.get("text") or choice.get("content"))

        message = getattr(choice, "message", None)
        if message is not None:
            return cls._coerce_message_content(getattr(message, "content", None))
        delta = getattr(choice, "delta", None)
        if delta is not None:
            return cls._coerce_message_content(getattr(delta, "content", None))
        return cls._coerce_message_content(getattr(choice, "text", None) or getattr(choice, "content", None))

    @classmethod
    def _extract_response_content(cls, response: Any) -> str:
        if response is None:
            return ""
        if isinstance(response, str):
            return response
        if isinstance(response, dict):
            if isinstance(response.get("content"), str):
                return response["content"]
            choices = response.get("choices")
            if isinstance(choices, list) and choices:
                return cls._extract_content_from_choice(choices[0])
            if isinstance(response.get("output_text"), str):
                return response["output_text"]
        choices = getattr(response, "choices", None)
        if choices:
            return cls._extract_content_from_choice(choices[0])
        output_text = getattr(response, "output_text", None)
        if isinstance(output_text, str):
            return output_text
        return cls._coerce_message_content(response)

    @classmethod
    def _extract_stream_response_content(cls, response: Any) -> str:
        if response is None:
            return ""
        if isinstance(response, str):
            return response
        if isinstance(response, dict):
            return cls._extract_response_content(response)

        response_buffer: list[str] = []
        try:
            iterator = iter(response)
        except TypeError:
            return cls._extract_response_content(response)

        for chunk in iterator:
            if isinstance(chunk, str):
                response_buffer.append(chunk)
                continue
            chunk_choices = getattr(chunk, "choices", None)
            if chunk_choices is None and isinstance(chunk, dict):
                chunk_choices = chunk.get("choices")
            if chunk_choices:
                for choice in chunk_choices:
                    piece = cls._extract_content_from_choice(choice)
                    if piece:
                        response_buffer.append(piece)
                continue
            piece = cls._extract_response_content(chunk)
            if piece:
                response_buffer.append(piece)
        return "".join(response_buffer)
    
    def get_response(
        self, 
        user_prompt: str, 
        system_prompt: Optional[str] = None,
        temperature: Optional[float] = None,
        max_tokens: Optional[int] = None,
        top_p: Optional[float] = None,
        frequency_penalty: Optional[float] = None,
        presence_penalty: Optional[float] = None
    ) -> Dict[str, Any]:
        """
        Get response from AI model or heuristic.
        
        Args:
            user_prompt: User prompt text
            system_prompt: System prompt (optional, will load from config if not provided)
            temperature: Temperature for response generation (optional, uses config default)
            max_tokens: Maximum tokens in response (optional, uses config default)
            top_p: Nucleus sampling parameter (optional, uses config default)
            frequency_penalty: Frequency penalty (optional, uses config default)
            presence_penalty: Presence penalty (optional, uses config default)
            
        Returns:
            Dictionary with response data including content, tokens, and duration
        """
        start_time = time.time()
        
        # Use config defaults if not provided
        temperature = temperature if temperature is not None else self.temperature
        max_tokens = max_tokens if max_tokens is not None else self.max_tokens
        top_p = top_p if top_p is not None else self.top_p
        frequency_penalty = frequency_penalty if frequency_penalty is not None else self.frequency_penalty
        presence_penalty = presence_penalty if presence_penalty is not None else self.presence_penalty
        
        # Use heuristic mode if enabled
        if self.use_heuristic:
            response_content = self._heuristic_response(user_prompt)
            duration = time.time() - start_time
            response = {
                'content': response_content,
                'mode': 'heuristic',
                'duration': duration,
                'temperature': temperature,
                'max_tokens': max_tokens
            }
            response.update(
                self._build_logging_metadata(
                    user_prompt=user_prompt,
                    system_prompt=system_prompt or "",
                    content=response_content,
                    duration=duration,
                    mode='heuristic',
                    temperature=temperature,
                    max_tokens=max_tokens,
                    top_p=top_p,
                    frequency_penalty=frequency_penalty,
                    presence_penalty=presence_penalty,
                )
            )
            response["temperature"] = temperature
            response["max_tokens"] = max_tokens
            return response
        
        # Load system prompt if not provided
        if system_prompt is None:
            system_prompt, _ = self.load_prompts()
        
        # Prepare messages
        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ]
        
        # Get response based on model type
        try:
            response_method = self.client.chat.completions.create
            
            # Common parameters
            params = {
                'model': self.model,
                'messages': messages,
                'temperature': temperature
            }
            
            # Add optional parameters if provided
            if max_tokens:
                params['max_tokens'] = max_tokens
            if top_p is not None:
                params['top_p'] = top_p
            if frequency_penalty is not None:
                params['frequency_penalty'] = frequency_penalty
            if presence_penalty is not None:
                params['presence_penalty'] = presence_penalty
            
            if self.model_type == 'openai':
                response = response_method(**params)
                content = self._extract_response_content(response)
                usage = getattr(response, 'usage', None)
                
            elif self.model_type == 'local':
                params['stream'] = True
                response = response_method(**params)

                content = self._extract_stream_response_content(response)
                usage = getattr(response, 'usage', None)
                
            elif self.model_type == 'deepseek':
                response = response_method(**params)
                content = self._extract_response_content(response)
                usage = getattr(response, 'usage', None)
            else:
                raise ValueError(f"Invalid model type: {self.model_type}")
            
            duration = time.time() - start_time
            prompt_tokens = getattr(usage, 'prompt_tokens', None) if usage is not None else None
            completion_tokens = getattr(usage, 'completion_tokens', None) if usage is not None else None
            total_tokens = getattr(usage, 'total_tokens', None) if usage is not None else None
            
            response_payload = {
                'content': content,
                'mode': 'ai',
                'duration': duration,
                'temperature': temperature,
                'max_tokens': max_tokens,
                'top_p': top_p,
                'frequency_penalty': frequency_penalty,
                'presence_penalty': presence_penalty,
            }
            response_payload.update(
                self._build_logging_metadata(
                    user_prompt=user_prompt,
                    system_prompt=system_prompt,
                    content=content,
                    duration=duration,
                    mode='ai',
                    temperature=temperature,
                    max_tokens=max_tokens,
                    top_p=top_p,
                    frequency_penalty=frequency_penalty,
                    presence_penalty=presence_penalty,
                    prompt_tokens=prompt_tokens,
                    completion_tokens=completion_tokens,
                    total_tokens=total_tokens,
                )
            )
            response_payload["temperature"] = temperature
            response_payload["max_tokens"] = max_tokens
            response_payload["top_p"] = top_p
            response_payload["frequency_penalty"] = frequency_penalty
            response_payload["presence_penalty"] = presence_penalty
            return response_payload
            
        except Exception as e:
            self.logger.error(f"Error getting response: {e}")
            raise
    
    def _heuristic_response(self, user_prompt: str) -> str:
        """
        Generate heuristic response without AI (rule-based).
        
        Args:
            user_prompt: User prompt text
            
        Returns:
            Heuristic response string
        """
        # Simple heuristic logic - can be expanded based on requirements
        return f"[Heuristic mode] Processed prompt with {len(user_prompt)} characters."
# --- end ai/client.py ---

# --- begin ai/pipeline.py ---
"""
Unified AI pipeline and prompt-loading surface for FDD.
"""


import multiprocessing
import os
import re
import threading
import concurrent.futures
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Any, Callable, Dict, List, Optional, Tuple

import pandas as pd
import yaml


SUBAGENT_SEQUENCE = [
    ("subagent_1", "Generator"),
    ("subagent_2", "Auditor"),
    ("subagent_3", "Refiner"),
    ("subagent_4", "Validator"),
]

def _get_prompt_manager(prompts_file: str, mappings_file: str) -> PromptEngine:
    resolved_prompts = resolve_prompt_asset_path(
        prompts_file,
        _DEFAULT_PROMPTS_FILE,
        get_prompt_engine().prompts_path,
    )
    resolved_mappings = resolve_prompt_asset_path(
        mappings_file,
        _DEFAULT_MAPPINGS_FILE,
        get_prompt_engine().mappings_path,
    )
    return get_prompt_engine(
        prompts_path=resolved_prompts,
        mappings_path=resolved_mappings,
    )


def map_value_to_component(
    value: str,
    component: Optional[str] = None,
    file_path: str = _DEFAULT_MAPPINGS_FILE,
) -> Any:
    """Look up mapping metadata from the configured mappings file."""
    mappings_path = resolve_prompt_asset_path(
        file_path,
        _DEFAULT_MAPPINGS_FILE,
        get_prompt_engine().mappings_path,
    )
    manager = get_prompt_engine(mappings_path=mappings_path)
    return manager.get_mapping_component(value, component=component)


def load_prompts_and_format(
    agent_name: str,
    language: str,
    mapping_key: str,
    df: pd.DataFrame,
    prompts_file: str = _DEFAULT_PROMPTS_FILE,
    mappings_file: str = _DEFAULT_MAPPINGS_FILE,
    **kwargs,
) -> Tuple[str, str]:
    """Render prompts for a specific agent/account pair."""
    manager = _get_prompt_manager(prompts_file, mappings_file)
    data_format = get_safe_default_data_format(language=language)
    return manager.render_prompt(
        agent_name=agent_name,
        language=language,
        mapping_key=mapping_key,
        df=df,
        data_format=data_format,
        **kwargs,
    )


def clean_agent_output(content: str) -> str:
    """Remove common meta-commentary from agent outputs."""
    prefixes_to_remove = [
        r"^verified\s+output:\s*",
        r"^corrected\s+output:\s*",
        r"^refined\s+output:\s*",
        r"^formatted\s+output:\s*",
        r"^final\s+output:\s*",
        r"^after\s+verification[,:]?\s*",
        r"^after\s+refining[,:]?\s*",
        r"^final\s+formatted\s+content:\s*",
        r"^the\s+corrected\s+output\s+is:\s*",
        r"^here\s+is\s+the\s+(corrected|refined|verified)\s+output:\s*",
        r"^已验证输出：\s*",
        r"^已更正输出：\s*",
        r"^精炼后的输出：\s*",
        r"^格式化后的输出：\s*",
        r"^经过验证[，,]\s*",
        r"^经过精炼后[，,]\s*",
    ]
    cleaned = (content or "").strip()
    for pattern in prefixes_to_remove:
        cleaned = re.sub(pattern, "", cleaned, flags=re.IGNORECASE)

    if (cleaned.startswith('"') and cleaned.endswith('"')) or (
        cleaned.startswith("'") and cleaned.endswith("'")
    ):
        cleaned = cleaned[1:-1]

    end_patterns = [
        r"\s*I\s+(?:verified|corrected|refined|checked).*$",
        r"\s*(?:Corrections?|Verifications?)\s+made:.*$",
        r"\s*我(?:验证|更正|精炼|检查)了.*$",
        r"\s*所做更正：.*$",
    ]
    for pattern in end_patterns:
        cleaned = re.sub(pattern, "", cleaned, flags=re.IGNORECASE)
    return cleaned.strip()


def create_result_shell(mapping_keys: List[str], dfs: Dict[str, pd.DataFrame]) -> Dict[str, Dict[str, str]]:
    return {key: {} for key in mapping_keys if key in dfs}


def _get_agent_stage_context(agent_name: str) -> Tuple[int, str, Optional[str]]:
    for agent_num, (name, label) in enumerate(SUBAGENT_SEQUENCE, start=1):
        if name == agent_name:
            previous_agent = None if agent_num == 1 else SUBAGENT_SEQUENCE[agent_num - 2][0]
            return agent_num, label, previous_agent
    raise ValueError(f"Unknown agent stage: {agent_name}")


def _store_agent_result(
    results: Dict[str, Dict[str, str]],
    mapping_key: str,
    agent_name: str,
    content: str,
    metadata: Dict[str, Any],
) -> None:
    results[mapping_key][agent_name] = content
    if agent_name == "subagent_4":
        results[mapping_key]["agent_4_validation"] = metadata
        results[mapping_key]["final"] = content


def _finalize_agent_content(
    *,
    agent_name: str,
    raw_content: str,
    previous_output: str,
    language: str,
) -> Tuple[str, Dict[str, Any]]:
    metadata: Dict[str, Any] = {}
    if agent_name == "subagent_4":
        parsed = parse_validator_response(raw_content, fallback_content=previous_output)
        content = clean_agent_output(parsed["final_content"])
        metadata = {
            "final_content": content,
            "raw_response": parsed.get("raw_response", raw_content),
            "clause_reviews": parsed.get("clause_reviews", []),
        }
    else:
        content = clean_agent_output(raw_content)
    if language == "Eng":
        content = polish_english_commentary(content)
        if agent_name == "subagent_4" and metadata:
            metadata["final_content"] = content
    return content, metadata


def _notify_stage_progress(
    progress_callback,
    *,
    agent_num: int,
    agent_label: str,
    completed: int,
    total_eligible: int,
    total_items: int,
    mapping_key: str,
) -> None:
    if not progress_callback:
        return
    progress_callback(
        agent_num,
        agent_label,
        completed,
        total_eligible,
        ((agent_num - 1) * total_items) + completed,
        mapping_key,
    )


def _run_ai_call(ai_helper, user_prompt: str, system_prompt: str, agent_name: str, timeout: int = 30):
    result_container = {"response": None, "error": None, "completed": False}
    agent_cfg = ai_helper.get_agent_settings(agent_name)

    def call_ai():
        try:
            result_container["response"] = ai_helper.get_response(
                user_prompt,
                system_prompt,
                temperature=agent_cfg.get("temperature"),
                max_tokens=agent_cfg.get("max_tokens"),
                top_p=agent_cfg.get("top_p"),
                frequency_penalty=agent_cfg.get("frequency_penalty"),
                presence_penalty=agent_cfg.get("presence_penalty"),
            )
            result_container["completed"] = True
        except Exception as exc:  # pragma: no cover - defensive
            result_container["error"] = exc
            result_container["completed"] = True

    thread = threading.Thread(target=call_ai, daemon=True)
    thread.start()
    thread.join(timeout=timeout)

    if not result_container["completed"]:
        raise TimeoutError(f"AI call timeout after {timeout} seconds")
    if result_container["error"]:
        raise result_container["error"]
    return result_container["response"]


def _agent_prompt_kwargs(
    agent_name: str,
    mapping_key: str,
    prompt_manager: PromptEngine,
    previous_output: str,
    agent_config: Optional[Dict[str, Any]] = None,
) -> Dict[str, str]:
    if agent_name == "subagent_1" and str(previous_output or "").strip():
        return {"previous_content": previous_output}
    if agent_name == "subagent_2":
        return {
            "account": prompt_manager.get_mapping_component(mapping_key, component="type") or mapping_key,
            "output": previous_output,
        }
    if agent_name == "subagent_3":
        reduction_target = int((agent_config or {}).get("reduction_target_pct", 64))
        return {
            "previous_content": previous_output,
            "original_length": len(previous_output or ""),
            "reduction_target_pct": str(reduction_target),
        }
    if agent_name == "subagent_4":
        return {"content": previous_output}
    return {}


def process_single_agent_item(
    agent_name: str,
    mapping_key: str,
    df: Optional[pd.DataFrame],
    ai_helper,
    prompt_manager: PromptEngine,
    logger: PipelineRunLogger,
    previous_output: str = "",
    user_comment: str = "",
) -> Tuple[str, str, Dict[str, Any]]:
    """Run one account through a single agent stage."""
    try:
        logger.log_agent_start(agent_name, mapping_key)
        agent_cfg = ai_helper.get_agent_settings(agent_name)

        system_prompt, user_prompt = prompt_manager.render_prompt(
            agent_name=agent_name,
            language=ai_helper.language,
            mapping_key=mapping_key,
            df=df,
            data_format=ai_helper.data_format,
            user_comment=user_comment,
            **_agent_prompt_kwargs(agent_name, mapping_key, prompt_manager, previous_output, agent_config=agent_cfg),
        )

        if logger.debug_mode:
            logger.log_debug("PROMPT_SYSTEM", mapping_key, "Agent=%s len=%s" % (agent_name, len(system_prompt)), system_prompt)
            logger.log_debug("PROMPT_USER", mapping_key, "Agent=%s len=%s" % (agent_name, len(user_prompt)), user_prompt)

        if agent_name == "subagent_1" and (not system_prompt or not user_prompt):
            placeholder = f"Content generation skipped for {mapping_key}: No prompts available"
            return mapping_key, placeholder, {}

        if not system_prompt or not user_prompt:
            return mapping_key, previous_output, {}

        response = _run_ai_call(ai_helper, user_prompt, system_prompt, agent_name)
        raw_content = response["content"].strip().replace("\n\n", "\n").replace("\n \n", "\n")

        if logger.debug_mode:
            logger.log_debug("RAW_OUTPUT", mapping_key, "Agent=%s len=%s" % (agent_name, len(raw_content)), raw_content)

        content, metadata = _finalize_agent_content(
            agent_name=agent_name,
            raw_content=raw_content,
            previous_output=previous_output,
            language=ai_helper.language,
        )

        if logger.debug_mode and agent_name == "subagent_4" and metadata.get("clause_reviews"):
            reviews = metadata["clause_reviews"]
            supported_count = sum(1 for r in reviews if r.get("supported"))
            unsupported = [r for r in reviews if not r.get("supported")]
            logger.log_debug("VALIDATION", mapping_key,
                "Clauses: %s total, %s supported, %s unsupported" % (len(reviews), supported_count, len(unsupported)))
            for r in unsupported:
                logger.log_debug("UNSUPPORTED_CLAUSE", mapping_key,
                    '"%s" -- %s' % (str(r.get("clause", ""))[:80], str(r.get("reason", ""))[:120]))

        logger.log_agent_complete(
            agent_name,
            mapping_key,
            response,
            system_prompt,
            user_prompt,
            prompt_context=prompt_manager.build_prompt_context_snapshot(
                df,
                language=ai_helper.language,
                user_comment=user_comment,
                previous_output=previous_output,
            ),
        )
        return mapping_key, content, metadata
    except Exception as exc:
        logger.log_error(agent_name, mapping_key, exc)
        if agent_name == "subagent_1":
            return mapping_key, f"Content generation failed for {mapping_key}: {str(exc)[:100]}", {}
        if previous_output and str(previous_output).strip():
            return mapping_key, previous_output, {}
        return mapping_key, f"Content generation incomplete for {mapping_key}: {str(exc)[:100]}", {}


def run_agent_stage(
    agent_name: str,
    mapping_keys: List[str],
    dfs: Dict[str, pd.DataFrame],
    results: Dict[str, Dict[str, str]],
    ai_helper,
    prompt_manager: PromptEngine,
    logger: PipelineRunLogger,
    use_multithreading: bool = True,
    max_workers: Optional[int] = None,
    progress_callback=None,
    total_items: int = 0,
    user_comments: Optional[Dict[str, str]] = None,
):
    """Run all items for a single agent stage."""
    if max_workers is None:
        max_workers = multiprocessing.cpu_count()

    agent_num, agent_label, previous_agent = _get_agent_stage_context(agent_name)

    eligible_keys = []
    for key in mapping_keys:
        if key not in dfs or key not in results:
            continue
        if previous_agent and previous_agent not in results[key]:
            continue
        eligible_keys.append(key)

    if use_multithreading and len(eligible_keys) > 1:
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {}
            for key in eligible_keys:
                future = executor.submit(
                    process_single_agent_item,
                    agent_name,
                    key,
                    dfs.get(key),
                    ai_helper,
                    prompt_manager,
                    logger,
                    results[key].get(previous_agent, "") if previous_agent else "",
                    (user_comments or {}).get(key, ""),
                )
                futures[future] = key

            completed = 0
            for future in as_completed(futures):
                mapping_key, content, metadata = future.result()
                _store_agent_result(results, mapping_key, agent_name, content, metadata)
                completed += 1
                _notify_stage_progress(
                    progress_callback,
                    agent_num=agent_num,
                    agent_label=agent_label,
                    completed=completed,
                    total_eligible=len(eligible_keys),
                    total_items=total_items,
                    mapping_key=mapping_key,
                )
    else:
        completed = 0
        for key in eligible_keys:
            mapping_key, content, metadata = process_single_agent_item(
                agent_name,
                key,
                dfs.get(key),
                ai_helper,
                prompt_manager,
                logger,
                results[key].get(previous_agent, "") if previous_agent else "",
                (user_comments or {}).get(key, ""),
            )
            _store_agent_result(results, mapping_key, agent_name, content, metadata)
            completed += 1
            _notify_stage_progress(
                progress_callback,
                agent_num=agent_num,
                agent_label=agent_label,
                completed=completed,
                total_eligible=len(eligible_keys),
                total_items=total_items,
                mapping_key=mapping_key,
            )


def set_final_fallbacks(results: Dict[str, Dict[str, str]]):
    """Populate `final` from the latest successful agent output."""
    for key in results:
        results[key]["final"] = get_pipeline_result_text(results[key])


def run_ai_pipeline_with_progress(
    mapping_keys: List[str],
    dfs: Dict[str, pd.DataFrame],
    model_type: str = "deepseek",
    language: str = "Eng",
    use_heuristic: bool = False,
    use_multithreading: bool = True,
    max_workers: Optional[int] = None,
    progress_callback: Optional[Callable[..., None]] = None,
    user_comments: Optional[Dict[str, str]] = None,
) -> Dict[str, Dict[str, str]]:
    """Run the 4-agent FDD pipeline with optional progress callbacks."""
    total_items = len([key for key in mapping_keys if key in dfs])
    fdd_config = FDDConfig(language=language, model_type=model_type)
    debug_mode = fdd_config.get_debug_mode()
    logger = PipelineRunLogger(debug_mode=debug_mode)
    prompt_manager = get_prompt_engine()
    ai_helper = AIClient(
        model_type=model_type,
        agent_name="content_pipeline",
        language=language,
        use_heuristic=use_heuristic,
    )
    results = create_result_shell(mapping_keys, dfs)

    logger.logger.info(
        "Starting FDD pipeline with %s items | model=%s | language=%s | multithreading=%s",
        total_items,
        ai_helper.model_type,
        language,
        use_multithreading,
    )

    for agent_name, agent_label in SUBAGENT_SEQUENCE:
        logger.logger.info("Running %s stage", agent_label)
        run_agent_stage(
            agent_name=agent_name,
            mapping_keys=mapping_keys,
            dfs=dfs,
            results=results,
            ai_helper=ai_helper,
            prompt_manager=prompt_manager,
            logger=logger,
            use_multithreading=use_multithreading,
            max_workers=max_workers,
            progress_callback=progress_callback,
            total_items=total_items,
            user_comments=user_comments,
        )

    # --- Feedback loop: re-run generator+validator for accounts with too many unsupported clauses ---
    feedback_config = fdd_config.get_feedback_loop_config()
    if feedback_config.get("enabled"):
        logger.logger.info(
            "Starting feedback loop (max_retries=%s, threshold=%.2f)",
            feedback_config["max_retries"],
            feedback_config["unsupported_threshold"],
        )
        for key in mapping_keys:
            if key not in results or key not in dfs:
                continue
            retries = _run_feedback_loop_for_key(
                key=key,
                dfs=dfs,
                results=results,
                ai_helper=ai_helper,
                prompt_manager=prompt_manager,
                logger=logger,
                feedback_config=feedback_config,
                user_comments=user_comments,
                progress_callback=progress_callback,
            )
            if retries > 0:
                logger.logger.info("[FeedbackLoop] %s: completed with %s retry(ies)", key, retries)

    set_final_fallbacks(results)

    # Generate simplified versions (~48% of detailed) for each account
    _generate_simplified_versions(
        results=results,
        ai_helper=ai_helper,
        logger=logger,
        use_multithreading=use_multithreading,
        max_workers=max_workers,
        progress_callback=progress_callback,
        total_items=total_items,
    )

    logger.finalize(results)
    return results


def _simplify_single_item(
    mapping_key: str,
    detailed_text: str,
    ai_helper,
    language: str,
) -> Tuple[str, str]:
    """Compress a single detailed commentary to ~48% for the simplified version."""
    system_prompt = (
        "You are a financial due diligence writing assistant. "
        "Compress the given commentary to approximately 48% of its current length. "
        "Keep the most important trend statement, the latest balance, and one key driver. "
        "Drop secondary details, minor items, and elaboration. "
        "Preserve all numbers, currency units, and period references exactly. "
        "Output only the compressed paragraph, no preamble."
    )
    user_prompt = f"Compress this FDD commentary to ~48% length:\n\n{detailed_text}"
    try:
        response = _run_ai_call(ai_helper, user_prompt, system_prompt, "subagent_3")
        content = clean_agent_output(response["content"].strip())
        if language == "Eng":
            content = polish_english_commentary(content)
        return mapping_key, content
    except Exception:
        # Fallback: mechanically trim to first 48% of sentences
        sentences = [s.strip() for s in detailed_text.replace(". ", ".\n").split("\n") if s.strip()]
        keep = max(1, int(len(sentences) * 0.48))
        return mapping_key, " ".join(sentences[:keep])


def _generate_simplified_versions(
    results: Dict[str, Dict[str, str]],
    ai_helper,
    logger: PipelineRunLogger,
    use_multithreading: bool = True,
    max_workers: Optional[int] = None,
    progress_callback: Optional[Callable[..., None]] = None,
    total_items: int = 0,
) -> None:
    """Generate simplified (~48%) versions of all final commentaries."""
    eligible = [
        (key, str(result.get("final", "")).strip())
        for key, result in results.items()
        if str(result.get("final", "")).strip()
    ]
    if not eligible:
        return

    logger.logger.info("Generating simplified versions for %s accounts", len(eligible))

    if progress_callback:
        try:
            progress_callback(5, "Simplifier", 0, len(eligible), 4 * total_items, "")
        except Exception:
            pass

    if use_multithreading and len(eligible) > 1:
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers or 4) as executor:
            futures = {
                executor.submit(_simplify_single_item, key, text, ai_helper, ai_helper.language): key
                for key, text in eligible
            }
            for future in concurrent.futures.as_completed(futures):
                try:
                    mapping_key, simplified = future.result()
                    results[mapping_key]["final_simplified"] = simplified
                except Exception as exc:
                    key = futures[future]
                    logger.logger.warning("Simplified version failed for %s: %s", key, exc)
    else:
        for key, text in eligible:
            try:
                _, simplified = _simplify_single_item(key, text, ai_helper, ai_helper.language)
                results[key]["final_simplified"] = simplified
            except Exception as exc:
                logger.logger.warning("Simplified version failed for %s: %s", key, exc)

    logger.logger.info("Simplified versions complete")


def _evaluate_feedback_needed(
    results: Dict[str, Dict[str, str]],
    key: str,
    unsupported_threshold: float,
) -> tuple[bool, float, List[Dict[str, Any]]]:
    """Check if a mapping_key's validator result warrants a feedback retry."""
    validation = (results.get(key) or {}).get("agent_4_validation") or {}
    clause_reviews = validation.get("clause_reviews") or []
    if not clause_reviews:
        return False, 0.0, []
    unsupported = [r for r in clause_reviews if isinstance(r, dict) and not r.get("supported")]
    ratio = len(unsupported) / len(clause_reviews)
    return ratio > unsupported_threshold, ratio, unsupported


def _run_feedback_loop_for_key(
    key: str,
    dfs: Dict[str, pd.DataFrame],
    results: Dict[str, Dict[str, str]],
    ai_helper,
    prompt_manager: PromptEngine,
    logger: PipelineRunLogger,
    feedback_config: Dict[str, Any],
    user_comments: Optional[Dict[str, str]] = None,
    progress_callback: Optional[Callable[..., None]] = None,
) -> int:
    """Run feedback loop for a single key. Returns number of retries performed."""
    max_retries = int(feedback_config.get("max_retries", 2))
    threshold = float(feedback_config.get("unsupported_threshold", 0.3))

    for retry_num in range(1, max_retries + 1):
        needs_feedback, ratio, unsupported = _evaluate_feedback_needed(results, key, threshold)
        if not needs_feedback:
            return retry_num - 1

        logger.logger.info(
            "[FeedbackLoop] %s: retry %s/%s (unsupported_ratio=%.2f, threshold=%.2f, unsupported_count=%s)",
            key, retry_num, max_retries, ratio, threshold, len(unsupported),
        )

        feedback_text = format_validator_feedback_for_reprompt(
            (results[key].get("agent_4_validation") or {}).get("clause_reviews", []),
            ai_helper.language,
        )

        if logger.debug_mode:
            logger.log_debug("FEEDBACK_LOOP", key,
                "Retry %s/%s: feedback_text_len=%s" % (retry_num, max_retries, len(feedback_text)),
                feedback_text)

        previous_output = get_pipeline_result_text(results[key])
        base_user_comment = (user_comments or {}).get(key, "")
        combined_comment = ("%s\n\n%s" % (base_user_comment, feedback_text)).strip() if feedback_text else base_user_comment

        # Re-run generator with feedback
        _key, gen_content, _meta = process_single_agent_item(
            "subagent_1", key, dfs.get(key), ai_helper, prompt_manager, logger,
            previous_output=previous_output,
            user_comment=combined_comment,
        )
        results[key]["subagent_1"] = gen_content
        results[key]["feedback_retry_%s_agent_1" % retry_num] = gen_content

        # Re-run validator on new output
        _key, val_content, val_metadata = process_single_agent_item(
            "subagent_4", key, dfs.get(key), ai_helper, prompt_manager, logger,
            previous_output=gen_content,
            user_comment=base_user_comment,
        )
        _store_agent_result(results, key, "subagent_4", val_content, val_metadata)
        results[key]["feedback_retry_%s_agent_4" % retry_num] = val_content
        results[key]["feedback_retries"] = retry_num

        if progress_callback:
            try:
                progress_callback(5, "FeedbackLoop-%s" % retry_num, 0, 0, 0, key)
            except Exception:
                pass

    return max_retries


def run_ai_pipeline(
    mapping_keys: List[str],
    dfs: Dict[str, pd.DataFrame],
    model_type: str = "deepseek",
    language: str = "Eng",
    use_heuristic: bool = False,
    use_multithreading: bool = True,
    max_workers: Optional[int] = None,
    user_comments: Optional[Dict[str, str]] = None,
) -> Dict[str, Dict[str, str]]:
    """Simple wrapper without progress callback."""
    return run_ai_pipeline_with_progress(
        mapping_keys=mapping_keys,
        dfs=dfs,
        model_type=model_type,
        language=language,
        use_heuristic=use_heuristic,
        use_multithreading=use_multithreading,
        max_workers=max_workers,
        progress_callback=None,
        user_comments=user_comments,
    )


def run_generator_reprompt(
    mapping_keys: List[str],
    dfs: Dict[str, pd.DataFrame],
    existing_results: Optional[Dict[str, Dict[str, str]]] = None,
    model_type: str = "deepseek",
    language: str = "Eng",
    use_heuristic: bool = False,
    user_comments: Optional[Dict[str, str]] = None,
) -> Dict[str, Dict[str, str]]:
    """Regenerate selected items, then immediately revalidate the revised output."""
    logger = PipelineRunLogger()
    prompt_manager = get_prompt_engine()
    ai_helper = AIClient(
        model_type=model_type,
        agent_name="content_pipeline",
        language=language,
        use_heuristic=use_heuristic,
    )
    results: Dict[str, Dict[str, str]] = {}

    logger.logger.info(
        "Starting reprompt + validator flow with %s items | model=%s | language=%s",
        len([key for key in mapping_keys if key in dfs]),
        ai_helper.model_type,
        language,
    )

    for key in mapping_keys:
        if key not in dfs:
            continue
        existing_result = (existing_results or {}).get(key) or {}
        previous_output = ""
        if isinstance(existing_result, dict):
            for field in ("final", "subagent_4", "subagent_3", "subagent_2", "subagent_1"):
                candidate = existing_result.get(field)
                if candidate and str(candidate).strip():
                    previous_output = str(candidate)
                    break

        mapping_key, content, _metadata = process_single_agent_item(
            "subagent_1",
            key,
            dfs.get(key),
            ai_helper,
            prompt_manager,
            logger,
            previous_output=previous_output,
            user_comment=(user_comments or {}).get(key, ""),
        )
        updated_result = dict(existing_result) if isinstance(existing_result, dict) else {}
        updated_result["subagent_1"] = content

        _validator_key, validator_content, validator_metadata = process_single_agent_item(
            "subagent_4",
            key,
            dfs.get(key),
            ai_helper,
            prompt_manager,
            logger,
            previous_output=content,
            user_comment=(user_comments or {}).get(key, ""),
        )
        updated_result["subagent_4"] = validator_content
        updated_result["agent_4_validation"] = validator_metadata
        updated_result["final"] = validator_content
        updated_result["reprompt_mode"] = "generator_reprompt_validated"
        results[mapping_key] = updated_result

    logger.finalize(results)
    return results


def save_results(results: Dict[str, Dict[str, str]], output_path: str = "fdd_utils/output/results.yml"):
    """Persist pipeline results to YAML."""
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as file:
        yaml.dump(results, file, default_flow_style=False, allow_unicode=True)


def extract_final_contents(results: Dict[str, Dict[str, str]]) -> Dict[str, str]:
    """Extract only the final content from the pipeline result payload."""
    return {
        key: value["final"]
        for key, value in results.items()
        if isinstance(value, dict) and "final" in value
    }


__all__ = [
    "SUBAGENT_SEQUENCE",
    "clean_agent_output",
    "extract_final_contents",
    "load_prompts_and_format",
    "map_value_to_component",
    "run_ai_pipeline",
    "run_ai_pipeline_with_progress",
    "run_generator_reprompt",
    "save_results",
]
# --- end ai/pipeline.py ---

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
    # KPMG Workbench gateway (Azure OpenAI-compatible). api_key doubles as the
    # 'Ocp-Apim-Subscription-Key' header value the gateway requires.
    "workbench": ["api_key", "api_base", "api_version", "chat_model"],
}

# Every model the Workbench gateway currently serves, in UI display order.
# The FIRST entry is the default model when the provider is selected without
# an explicit override.
WORKBENCH_AVAILABLE_MODELS = [
    {"id": "gpt-5-5-2026-04-24-gs-sdc", "label": "GPT-5.5"},
    {"id": "gpt-5-4-2026-03-05-gs-sdc", "label": "GPT-5.4"},
]

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


_MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]


def _iso_to_long_date(match: "re.Match[str]") -> str:
    year, month, day = match.group(1), int(match.group(2)), int(match.group(3))
    if 1 <= month <= 12 and 1 <= day <= 31:
        return f"{day} {_MONTH_NAMES[month - 1]} {year}"
    return match.group(0)


def _k_to_comma_int(match: "re.Match[str]") -> str:
    """Convert 'CNY78.2K' style to 'CNY78,200' comma-int."""
    raw = match.group(1).replace(",", "")
    try:
        amount = float(raw) * 1000
    except ValueError:
        return match.group(0)
    return f"CNY{int(round(amount)):,}"


def _enforce_reference_style(text: str) -> str:
    """Apply deterministic reference-style fixes that the AI sometimes forgets:

    - ISO date 'YYYY-MM-DD' -> 'D Month YYYY'
    - Lowercase 'The balance as at' opening (after bullet em-dash, sentence start)
    - 'CNY <number>' -> 'CNY<number>' (no space)
    - 'CNY <X>K' -> 'CNY<comma-int>'
    - Drop ', annualised at CNY...' inserts
    - Drop common PoP filler templates
    - Force 1dp on 'CNY X.YY million' (collapse to 1dp)
    """
    if not text:
        return text
    out = text

    # 1) Strip PoP filler / verbose cross-checks / annualisation FIRST so number
    #    formats (which may contain commas) don't break the regexes below.
    out = re.sub(r",\s*annualised at CNY[\d.,KkMm ]+(?:\s+million)?", "", out, flags=re.IGNORECASE)
    out = re.sub(r"\s*\(annualised at CNY[\d.,KkMm ]+(?:\s+million)?\)", "", out, flags=re.IGNORECASE)
    # '(CNY5.7 million annualised)' — reversed-order parenthetical
    out = re.sub(r"\s*\(CNY[\d,]+(?:\.\d+)?\s*(?:million\s+)?annualis(?:ed|ation)\)", "", out, flags=re.IGNORECASE)
    # 'FY24, FY25 and 1M26 annualised' — period label with annualised qualifier
    out = re.sub(r"\b((?:FY\d{2}|1M\d{2}|CY\d{4}|\d{4}))\s+annualis(?:ed|ation)\b", r"\1", out, flags=re.IGNORECASE)
    # 'the annualised 1M26' / 'the annualised FY24' — reversed-order qualifier
    out = re.sub(r"\bthe\s+annualis(?:ed|ation)\s+((?:FY\d{2}|1M\d{2}|CY\d{4}|\d{4}))\b", r"\1", out, flags=re.IGNORECASE)

    # Reusable amount fragment: optional space, digits/commas/dot, optional K or million suffix.
    _amount = r"\s*CNY\s*[\d.,]+\s*(?:K|million)?"
    pop_filler_patterns = [
        # 'As at 31 December 2025, the balance was CNY X, with a similar composition...'
        r"(?:\.|\;)\s*As at \d{1,2} [A-Z][a-z]+ \d{4},?\s+the balance was" + _amount + r",?\s+with a similar composition[^.;]*",
        # ', reflecting a slight (decrease|increase|build-up|reduction) of CNY...'
        r",\s*reflecting (?:a slight |an? )?(?:decrease|increase|build-up|reduction) of" + _amount + r"(?:\s+by [^,.]+)?",
        # ', remained relatively stable'
        r",\s*remained relatively stable",
        # 'The X prepayment was not present in prior periods but appeared in ...'
        r"\.?\s*The [^.]{0,80} was not present in prior periods[^.]{0,200}\.",
        # 'showed a slight increase from CNY ...' (with or without 'to CNY ...')
        r",?\s*showed a slight (?:increase|decrease) from" + _amount + r"(?:\s+to" + _amount + r")?",
        # 'indicating a new ... arrangement in YYYY'
        r",\s*indicating a new [^,.]{0,80} arrangement in \d{4}",
        # Verbose cross-check: 'has been cross-checked with the bank statements...'
        r"\s*The (?:total )?[^.]{0,40} balance has been cross-checked with the bank statements? (?:and )?(?:no material discrepancies were identified|with no material discrepancies?)\.?",
        r"\s*[Ww]e have cross-checked [^.]{0,80} (?:and|with) (?:no material discrepancies were identified|no differences? were identified)\.?",
        # Full sentences containing 'annualised at'
        r"[^.;]*\bannualis[ezd]+ at" + _amount + r"[^.;]*\.",
        # 'with annualised X income of CNY...' or 'annualised X of CNY...' inline phrase
        # Use [\d,]+ (not [\d.,]+) so we don't consume a trailing sentence-end period.
        r",?\s*with\s+annualis(?:ed|ation)\s+[\w\s]{1,40}\s+of\s+CNY[\d,]+(?:\.\d+)?(?:\s*(?:K|million))?",
        r"[^.;]*\bannualis(?:ed|ation)\s+[\w\s]{1,30}\s+(?:of|at)\s+CNY[\d,]+(?:\.\d+)?[^.;]*\.",
        # Orphaned number fragment at sentence start: '. 0 million respectively...' artifacts
        r"(?<=\. )\d[\d.,]*\s+(?:million|billion|thousand)\b[^.;]*\.",
    ]
    for pat in pop_filler_patterns:
        out = re.sub(pat, "", out)

    # Strip entire sentences containing unreplaced <PLACEHOLDER> template markers.
    # Split on sentence boundaries (period + space + capital letter) so decimal
    # points in amounts don't confuse the sentence detector.
    _placeholder_re = re.compile(r"<[A-Z][A-Z_/ ]{1,30}>")
    _parts = re.split(r"(?<=\.)\s+(?=[A-Z])", out)
    out = " ".join(p for p in _parts if not _placeholder_re.search(p))

    # 2) ISO dates -> long form
    out = re.sub(r"\b(\d{4})-(\d{2})-(\d{2})\b", _iso_to_long_date, out)

    # 3) CNY<space><digit> -> CNY<digit>
    out = re.sub(r"\bCNY\s+(?=-?\d)", "CNY", out)
    out = re.sub(r"\bCNY-\s+(?=\d)", "CNY-", out)

    # 4) K notation -> comma-int. Handles 'CNY78.2K', 'CNY-78.2K'.
    out = re.sub(r"\bCNY(-?\d+(?:\.\d+)?)K\b", _k_to_comma_int, out)

    # 5) 'CNY7.90 million' (2+dp) -> 'CNY7.9 million' (1dp)
    def _shorten_million(match: "re.Match[str]") -> str:
        sign = match.group(1) or ""
        whole = match.group(2)
        frac = match.group(3)
        try:
            value = float(f"{whole}.{frac}")
        except ValueError:
            return match.group(0)
        return f"CNY{sign}{value:.1f} million"

    out = re.sub(r"\bCNY(-?)(\d+)\.(\d{2,})\s+million\b", _shorten_million, out)

    # 6) Lowercase 'The balance as at' / 'The accumulated reserve fund' / 'The registered capital'
    #    ONLY at the very start of the bullet (avoids breaking second-sentence
    #    occurrences like "...paid-in capital. The registered capital was USD30 million.").
    out = re.sub(r"^The balance as at\b", "the balance as at", out)
    out = re.sub(r"^The accumulated reserve fund\b", "the accumulated reserve fund", out)
    # 'The balance as at' globally (always sentence opener after period+space)
    out = re.sub(r"(\.\s+)The balance as at\b", r"\1the balance as at", out)

    # 7) Strip leftover annualisation phrasings the prompt-level ban missed.
    #    a) Whole sentences starting with 'Annualised X ...' (with or without 'The'),
    #       commonly used as 'Annualised real estate tax of CNY... and land use tax
    #       of CNY... suggest stable accruals...'.
    out = re.sub(
        r"(?im)(?:^|(?<=\.\s)|(?<=;\s))(?:The\s+)?Annualis(?:ed|ation)\s[^.;\n]*?(?:\.\d[^.;\n]*?)?\.\s?",
        "",
        out,
    )
    #    b) Inline 'annualised at CNY...' inserts.
    out = re.sub(r",?\s*annualis[ezd]+\s+(?:at\s+)?CNY[\d.,KkMm ]+(?:\s+million)?", "", out, flags=re.IGNORECASE)

    # 8) Strip meta-commentary, advisory, and assertion leaks the AI introduces.
    #    Project rule: bullets state ONLY what the databook supports — no
    #    consultant advisory ("You should..."), no policy meta-commentary,
    #    no calculated rates / negative assertions / period-on-period filler.
    meta_patterns = [
        # ---- Advisory ('You should...') — banned entirely per user request ----
        r"(?i)\s*You should confirm with management[^.]+\.",
        r"(?i)\s*You should consider [^.]+\.",
        r"(?i)\s*You should compare [^.]+\.",
        r"(?i)\s*[Yy]ou (?:may|might) (?:wish to|want to) [^.]+\.",
        # ---- Verbose cross-check — broad: any "X was/has been cross-checked..." sentence ----
        r"\s*There is no mention in the data of[^.]{0,200}\.",
        r"(?i)\s*(?:We verified that )?[Tt]he (?:total )?[\w ]{1,40}(?:\s+(?:was|has been|were)) cross-checked (?:with|against) the bank statements?[^.]*\.",
        # ---- Audit-report verification leak ----
        r"(?i)\s*The audit report for \d{4} was reviewed and found to be consistent with this amount\.",
        # ---- Negative assertions (these aren't in the data unless the data explicitly says so) ----
        r"(?i)[^.;]*no retained earnings were appropriated for [^.;]+\.",
        r"(?i)[^.;]*no material adjustments or distributions? were recorded during the period[^.;]*\.",
        r"(?i)[^.;]*No provision for impairment was recorded[^.;]*\.",
        # 'Management indicated/noted/stated/said that no significant renovations or decorations were performed'
        r"(?i)[^.;]*Management (?:indicated|noted|stated|said) that no significant renovations or decorations were performed[^.;]*\.",
        # ---- Attributive/explanatory padding ----
        r"[^.;]*The pre-received amounts were mainly attributable to[^.;]+\.",
        r"[^.;]*The advance payments were for future \d[^.;]+\.",
        # New variant: 'The pre-received rental income represented future 1–3 months' rent...'
        r"(?i)[^.;]*The pre-received [\w ]{1,40} represented future \d[\u2013\u2014\-]\d\s*month[^.;]*\.",
        # 'X represented future N to M months ...' generic
        r"(?i)[^.;]*\brepresented future \d (?:to|[\u2013\u2014\-]) \d[\s\-]?month[^.;]*\.",
        r"[^.;]*had been fully settled by \d{4} and showed no further liability[^.;]*\.",
        r"[^.;]*The charges are based on fixed proportions[^.;]+as stated by management\.",
        r",?\s*consistent with the (?:fixed )?contractual terms[^,.]*",
        r",?\s*indicating no additional (?:losses or profits|profits or losses)[^,.]*",
        # ---- PoP filler about a 'remained unchanged' balance ----
        r"(?i)[^.;\n]*\bremained unchanged at CNY[\d.,]+(?:\s*(?:K|million))?\s*as at \d{1,2} [A-Z][a-z]+ \d{4}[^.;\n]*\.",
        # ---- Land residual value/rate hallucination — land conventionally has 0% residual ----
        #    Match 'residual value' OR 'residual rate'; with or without leading 'The'.
        r"(?i),?\s*(?:The )?[Ll]and(?: use rights?)? (?:is|are) (?:depreciated|amortised) using the straight-line method over \d+ years? with a [\d.]+%\s+residual (?:value|rate)\.?",
        # ---- Calculated rate not in source ('5-year LPR minus X%') ----
        r"\s*\(5-year LPR (?:minus|plus) [\d.]+%\)",
        # ---- Annualised wording variants ('annualised in FY..., FY..., 1M26 respectively') ----
        # Strip from 'annualised in <period>' through the next 'respectively' or sentence end.
        r"(?i),?\s+annualis(?:ed|ation)\s+in\s+[^.;]*?\brespectively",
        # Single-period annualised tail without 'respectively'
        r"(?i),?\s+annualis(?:ed|ation)\s+in\s+(?:FY\d{2}|\d{4}|\dM\d{2})",
        # ---- Unsupported 'consistent with historical trends' assertion ----
        r"(?i),?\s*(?:suggest|implying)\s+stable accruals[^.;]*",
        r"(?i),?\s*consistent with historical trends[^.;]*",
        # ---- Statutory reserves negative assertion ----
        r"(?i)[^.;]*No statutory or discretionary reserves were appropriated[^.;]*\.",
        # ---- Redundant tail sentence: 'The loss for the period was CNY1.3 million.' ----
        r"(?i)\.?\s*The loss for the period was CNY[\d.,]+(?:\s*(?:K|million))?\.",
        # ---- T&S/G&A filler: 'in line with the fixed proportion of...' ----
        r"(?i),?\s*in line with the fixed proportion of [^.;]+",
        # ---- T&S policy boilerplate sentence ('calculated at a fixed proportion of rental income') ----
        r"(?i)[^.;]*[Tt]he real estate tax is calculated at a fixed proportion[^.;]*\.",
        r"(?i),?\s*while the land use tax is based on a fixed proportion[^.;]*",
        # ---- Non-op / supplemental boilerplate ----
        r"(?i),?\s*consistent with the supplemental detail provided[^.;]*",
        # ---- Retention earnings negative assertion ----
        r"(?i)[^.;]*No material adjustments or changes in accounting policies were reported[^.;]*\.",
    ]
    for pat in meta_patterns:
        out = re.sub(pat, "", out)

    # 9) Date format normalisation: '01 March 2024' / '07 May 2026' -> '1 March 2024' (strip leading zero on day)
    out = re.sub(r"\b0(\d) ([A-Z][a-z]+) (\d{4})\b", r"\1 \2 \3", out)

    # 10) ratio formatting tweak: '60: 40' -> '60:40'; common AI artifact
    out = re.sub(r"(\d+):\s+(\d+)\b", r"\1:\2", out)

    # 11) Zero/nil amount handling — convert true-zero amounts to 'nil'.
    #     Match: CNY0, CNY0.0, CNY 0, CNY0 million, CNY0.0 million, CNY0K.
    #     Don't match: CNY0.1 (non-zero) — negative lookahead on '.\d' or digit.
    out = re.sub(
        r"\bCNY\s*0(?:\.0+)?(?:\s+million|\s*K)?(?![\d.])",
        "nil",
        out,
        flags=re.IGNORECASE,
    )

    # 12) Drop trivial sub-streams that are 'nil' across most periods.
    #     Collapse 'X totalled nil, nil, CNY0.5 million in FY24, FY25, 1M26 respectively' patterns.
    #     If a multi-period series has 2+ 'nil' values out of 3, drop the entire sentence.
    def _drop_mostly_nil_series(match: "re.Match[str]") -> str:
        sentence = match.group(0)
        nil_count = len(re.findall(r"\bnil\b", sentence, flags=re.IGNORECASE))
        amount_count = len(re.findall(r"\bnil\b|\bCNY[\d.,]+(?:\s*(?:K|million))?", sentence, flags=re.IGNORECASE))
        # Drop the sentence if more than half the listed amounts are nil
        if amount_count >= 2 and nil_count / amount_count >= 0.5:
            return ""
        return sentence

    # Match a sentence that contains a multi-period amount listing with at least one nil
    out = re.sub(
        r"(?:^|(?<=\.\s))[A-Z][^.;\n]*?\bnil\b[^.;\n]*?\bin\s+(?:FY\d{2}|\d{4}|\dM\d{2})[^.;\n]*?(?:respectively)?\.\s?",
        _drop_mostly_nil_series,
        out,
    )

    # Tidy double spaces / orphan punctuation introduced by the strips
    out = re.sub(r"\s+([,.;:])", r"\1", out)
    out = re.sub(r",\s*\.", ".", out)
    out = re.sub(r"\s{2,}", " ", out)
    return out


def polish_english_commentary(text: str) -> str:
    polished = normalize_english_text(text or "")
    for pattern in _SECTION_LABEL_PATTERNS:
        polished = re.sub(pattern, "", polished)
    polished = re.sub(r"(?i)^including:\s*", "Including ", polished)
    polished = polished.replace("Co. , Ltd. .", "Co., Ltd.")
    polished = polished.replace("Co. , Ltd.", "Co., Ltd.")
    polished = polished.replace("\n- ", " ").replace("\n", " ")
    polished = re.sub(r"\s{2,}", " ", polished)
    polished = _enforce_reference_style(polished)
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


# Qwen3 (and other reasoning models) emit a <think>...</think> block before the
# answer. With no reasoning parser on the server it arrives inline in the content
# and pollutes BOTH the bullet text and any JSON. Strip it everywhere, tolerating a
# truncated (unclosed) block and a stray leading </think> (enable_thinking=false).
_THINK_BLOCK_RE = re.compile(r"<think>.*?</think>", flags=re.DOTALL | re.IGNORECASE)
_THINK_OPEN_TO_END_RE = re.compile(r"<think>.*\Z", flags=re.DOTALL | re.IGNORECASE)
_THINK_STRAY_CLOSE_RE = re.compile(r"^\s*</think>", flags=re.IGNORECASE)


def strip_thinking(text: str) -> str:
    """Remove <think>...</think> reasoning blocks (balanced, truncated, or stray-close)."""
    s = str(text or "")
    s = _THINK_BLOCK_RE.sub("", s)        # well-formed blocks
    s = _THINK_OPEN_TO_END_RE.sub("", s)  # unclosed block (truncated under max_tokens)
    s = _THINK_STRAY_CLOSE_RE.sub("", s)  # lone </think> with no opener
    return s.strip()


def _strip_code_fence(text: str) -> str:
    match = re.search(r"```(?:json)?\s*(.*?)```", text or "", flags=re.DOTALL | re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return str(text or "").strip()


def _balanced_brace_slice(text: str) -> str | None:
    """Return the first top-level {...} object via depth tracking (string-aware).

    Robust to trailing prose after the object and to stray braces inside any
    surviving reasoning text, where a naive find('{')..rfind('}') over-captures.
    """
    start = text.find("{")
    if start < 0:
        return None
    depth, in_str, esc = 0, False, False
    for i in range(start, len(text)):
        c = text[i]
        if in_str:
            if esc:
                esc = False
            elif c == "\\":
                esc = True
            elif c == '"':
                in_str = False
            continue
        if c == '"':
            in_str = True
        elif c == "{":
            depth += 1
        elif c == "}":
            depth -= 1
            if depth == 0:
                return text[start : i + 1]
    return None


def _repair_json(text: str) -> str:
    """Cheap repairs for weak-model JSON: smart quotes and trailing commas."""
    s = (text
         .replace("“", '"').replace("”", '"')
         .replace("‘", "'").replace("’", "'"))
    s = re.sub(r",\s*([}\]])", r"\1", s)  # trailing comma before } or ]
    return s


def _extract_json_payload(text: str) -> Dict[str, Any] | None:
    candidate = _strip_code_fence(strip_thinking(text))
    # Try, in order: strict parse, balanced-brace slice, then cheap repairs on each.
    attempts = [candidate, _balanced_brace_slice(candidate)]
    for attempt in attempts:
        if not attempt:
            continue
        for variant in (attempt, _repair_json(attempt)):
            try:
                parsed = json.loads(variant)
                if isinstance(parsed, dict):
                    return parsed
            except json.JSONDecodeError:
                continue
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
        # LLM didn't output category — default to "reasoning" (orange) rather
        # than "hallucination" (yellow) to avoid false-alarm severity labels
        # when the model omits the field under load.
        category = "reasoning"
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


def _fallback_clause_reviews(final_content: str) -> List[Dict[str, Any]]:
    """Deterministic clause_reviews when the validator JSON can't be parsed.

    Segments the text and marks each clause supported (no inline highlight) so the
    UI/feedback metric have a non-empty, well-shaped list to work with. Real
    number-grounding happens in verify_commentary (which has the source df); this
    is only the no-JSON safety net that keeps the shape valid and stops re-loops.
    """
    reviews: List[Dict[str, Any]] = []
    for _start, _end, clause in segment_clauses(final_content):
        reviews.append({
            "clause": clause,
            "supported": True,
            "category": "data-backed",
            "reason": "Auto-segmented (validator JSON unparseable).",
        })
    return reviews


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
        # JSON unparseable (common on weak local models even after repair). Fall
        # back to deterministic segmentation rather than returning [] — empty
        # clause_reviews silently disabled highlighting AND made the account look
        # "clean" (unsupported ratio 0), so the feedback loop and
        # _ensure_clause_reviews_on_final kept re-running the same failing call.
        final = strip_thinking(str(fallback_content or raw_text or "")).strip()
        return {
            "final_content": final,
            "clause_reviews": _fallback_clause_reviews(final),
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


# Boundary punctuation stripped when retrying a clause match — both ASCII and the
# fullwidth marks common in Chinese commentary (，。；：、！？「」『』（）).
_CLAUSE_BOUNDARY_CHARS = " \t\r\n\"'`.,;:!?，。；：、！？「」『』（）()"


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

    match_len = len(normalized_clause)
    if normalized_index < 0:
        # Punctuation-tolerant retry. The Validator often returns a clause whose
        # leading/trailing punctuation differs from the source text — most common
        # in Chinese, where it may add or drop a fullwidth 。／，／；. Strip those
        # boundary marks from the clause and search again so the clause still
        # highlights inline instead of silently falling back to the notes block.
        stripped_clause = normalized_clause.strip(_CLAUSE_BOUNDARY_CHARS)
        if stripped_clause and stripped_clause != normalized_clause:
            normalized_index = normalized_text.find(stripped_clause, normalized_cursor)
            if normalized_index < 0:
                normalized_index = normalized_text.find(stripped_clause)
            if normalized_index >= 0:
                match_len = len(stripped_clause)
    if normalized_index < 0:
        return (-1, -1)

    start = index_map[normalized_index]
    end_idx = normalized_index + match_len - 1
    if end_idx >= len(index_map):
        return (-1, -1)
    end = index_map[end_idx] + 1
    return (start, end)


def _find_next_clause_index(text: str, clause: str, cursor: int) -> int:
    if not clause:
        return -1
    return text.find(clause, cursor)


# ---------------------------------------------------------------------------
# Deterministic clause segmentation + number-grounding (foundation for the
# hallucination/reasoning verifier and the Qwen3 unparseable-JSON fallback).
# These let the pipeline classify most clauses in Python — far more reliable on a
# weak local model than asking it to copy clauses verbatim and do arithmetic.
# ---------------------------------------------------------------------------
_CLAUSE_END_CHARS = ".;。；！？!?"


def segment_clauses(text: str) -> List[Tuple[int, int, str]]:
    """Split `text` into 分句-level clauses on sentence-ends and clause commas.

    Returns ordered (start, end, clause) where clause == text[start:end] EXACTLY
    (so a highlighter can use the offsets directly, never needing a fuzzy
    re-match). A comma inside a number (1,234,567) is never a boundary.
    """
    text = str(text or "")
    spans: List[Tuple[int, int, str]] = []
    n = len(text)
    start = 0
    for i, ch in enumerate(text):
        boundary = ch in _CLAUSE_END_CHARS or ch in ",，"
        if boundary and ch in ".,，":
            # A '.' or ',' between two digits is a decimal point or thousands
            # separator (5.8 / 1,234,567), never a clause boundary.
            prev_c = text[i - 1] if i > 0 else ""
            next_c = text[i + 1] if i + 1 < n else ""
            if prev_c.isdigit() and next_c.isdigit():
                boundary = False
        if boundary:
            _append_clause_span(spans, text, start, i + 1)
            start = i + 1
    _append_clause_span(spans, text, start, n)
    return spans


def _append_clause_span(spans: List[Tuple[int, int, str]], text: str, start: int, end: int) -> None:
    chunk = text[start:end]
    stripped = chunk.strip()
    if not stripped:
        return
    lead = len(chunk) - len(chunk.lstrip())
    s = start + lead
    spans.append((s, s + len(stripped), stripped))


# Money expressions only — bare integers/years/percentages are intentionally NOT
# treated as groundable amounts (keeps false-positive hallucination flags low).
_AMT_MILLION = re.compile(r"(?:CNY|RMB|USD|HKD|US\$|\$|人民币|人民幣)?\s*(\d[\d,]*(?:\.\d+)?)\s*(?:million|mn)\b", re.IGNORECASE)
_AMT_YI = re.compile(r"(\d[\d,]*(?:\.\d+)?)\s*亿")
_AMT_WAN = re.compile(r"(\d[\d,]*(?:\.\d+)?)\s*万")
_AMT_CUR_PREFIX = re.compile(r"(?:CNY|RMB|USD|HKD|US\$|\$|人民币|人民幣)\s*(\d[\d,]*(?:\.\d+)?)", re.IGNORECASE)
_AMT_GROUPED = re.compile(r"(?<![\d.])(\d{1,3}(?:,\d{3})+(?:\.\d+)?)")


def _to_float(token: str) -> Optional[float]:
    try:
        return float(str(token).replace(",", ""))
    except (TypeError, ValueError):
        return None


def extract_amounts(clause: str) -> List[float]:
    """Extract absolute money amounts (scaled to base units) from a clause.

    Scale-bearing forms (million / 万 / 亿) are parsed first and their matched text
    blanked so a following currency-prefix/grouped pass cannot double-count the
    same figure (e.g. 'CNY5.8 million' must yield 5.8e6, not also 5.8).
    """
    amounts: List[float] = []
    work = clause
    # Each pass blanks the span it consumed so a later, looser pass cannot
    # re-count the same figure (e.g. 'CNY5.8 million' -> 5.8e6 only; 'CNY54,950'
    # -> 54950 once, not also via the grouped-thousands pass).
    for rx, scale in ((_AMT_MILLION, 1e6), (_AMT_YI, 1e8), (_AMT_WAN, 1e4),
                      (_AMT_CUR_PREFIX, 1.0), (_AMT_GROUPED, 1.0)):
        def _sub(m: "re.Match") -> str:
            v = _to_float(m.group(1))
            if v is not None:
                amounts.append(v * scale)
            return " " * len(m.group(0))
        work = rx.sub(_sub, work)
    return amounts


_BARE_NUMBER_RE = re.compile(r"\d[\d,]*(?:\.\d+)?")


def _attr_text_blob(df) -> str:
    """Concatenate all free-text in df.attrs (supporting_notes, table_linked_remarks,
    adjacent_detail_rows, rhs context, etc.) so figures cited in the NOTES — not just
    the numeric table — can ground a clause. Many legitimate figures (registered
    capital, audit fees, USD amounts) live only in the remarks."""
    parts: List[str] = []
    attrs = getattr(df, "attrs", None) or {}

    def walk(v):
        if isinstance(v, str):
            parts.append(v)
        elif isinstance(v, dict):
            for vv in v.values():
                walk(vv)
        elif isinstance(v, (list, tuple)):
            for vv in v:
                walk(vv)

    for value in attrs.values():
        walk(value)
    return " ".join(parts)


def _numbers_in_text(text: str) -> List[float]:
    """Every number a remark could supply: scale-aware amounts (万/亿/million/comma)
    PLUS bare integers/decimals (e.g. '191400', '7000', '572')."""
    out = list(extract_amounts(text))
    for m in _BARE_NUMBER_RE.finditer(str(text or "")):
        v = _to_float(m.group(0))
        if v is not None:
            out.append(v)
    return out


class SourceIndex:
    """Numeric values present in an account's source data, for grounding amounts."""

    def __init__(self, values: List[float]):
        self.values = [v for v in values if v is not None]

    @staticmethod
    def _adjacent_window_sums(col_vals: List[float], max_window: int = 4) -> List[float]:
        """Sums of every run of 2..max_window CONSECUTIVE rows (sheet order, as
        the column already preserves it) — commentary legitimately groups a
        handful of neighbouring breakdown lines into one figure (e.g. "CNY322,116
        of property[-related fees]" = 4 adjacent line items in Other payables
        that were never a labelled subtotal in the sheet). Bounded to small
        windows, not a full subset-sum search, to keep this O(n) and keep the
        false-negative risk (a genuinely wrong number coincidentally matching
        some arbitrary window) low."""
        sums: List[float] = []
        n = len(col_vals)
        for window in range(2, max_window + 1):
            for start in range(0, n - window + 1):
                sums.append(sum(col_vals[start:start + window]))
        return sums

    @classmethod
    def _column_values(cls, df, skip_cols: tuple = ()) -> List[float]:
        values: List[float] = []
        for col in df.columns:
            if col in skip_cols:
                continue
            series = df[col]
            col_vals: List[float] = []
            if getattr(series, "dtype", None) is not None and series.dtype.kind in "if":
                col_vals = [float(v) for v in series.dropna().tolist()]
            else:
                for cell in series.tolist():
                    v = _to_float(cell) if isinstance(cell, (int, float, str)) else None
                    if v is not None:
                        col_vals.append(v)
            values += col_vals
            # Add the column total — commentary frequently cites a total that
            # isn't a single cell; including it avoids false hallucination flags.
            if col_vals:
                values.append(sum(col_vals))
            values += cls._adjacent_window_sums(col_vals)
        return values

    @classmethod
    def _values_for_one_df(cls, df) -> List[float]:
        values: List[float] = []
        if df is None or not hasattr(df, "columns"):
            return values
        values += cls._column_values(df)
        # df is `projection_df` — a SINGLE latest-period snapshot. Multi-year
        # trend commentary ("increased from CNY384M as at 2023-12-31 to
        # CNY709M as at 2024-12-31") is written from df.attrs["prompt_analysis_df"]
        # (see _build_financial_prompt_payload's "analysis_periods" block, which
        # the Generator AND Validator both receive) — without indexing it here
        # too, every correctly-written historical-period number is invisible to
        # this grounding pool and gets falsely flagged as "hallucination", which
        # _combine_verdict then treats as authoritative over the LLM's own
        # (correct) judgement. INTERNAL_ROW_KEY is excluded — it holds raw sheet
        # row indices, not financial amounts.
        analysis_df = df.attrs.get("prompt_analysis_df")
        if analysis_df is not None and hasattr(analysis_df, "columns"):
            values += cls._column_values(analysis_df, skip_cols=(INTERNAL_ROW_KEY,))
        # Also ground against numbers cited in the supporting notes / remarks
        # (df.attrs), e.g. registered capital "7000万美元" that never appears in
        # the numeric table. Without this they were false-flagged as hallucinations.
        values += _numbers_in_text(_attr_text_blob(df))
        return values

    @classmethod
    def from_df(cls, df, sibling_dfs: Optional[List[Any]] = None) -> "SourceIndex":
        values: List[float] = cls._values_for_one_df(df)
        # Commentary for one account sometimes legitimately cites a figure that
        # actually lives on a DIFFERENT tab — e.g. "Other payables" explaining
        # accrued interest by naming the CNY198.0 million bank loan it relates
        # to, where the loan balance itself is only in the "Long-term loans"
        # tab. Restricted to this account's own df, that number is invisible
        # and a coincidental same-tab match at the wrong scale produces a false
        # "hallucination" flag — confirmed via a real client databook where the
        # cited loan balance (198,870,239) was correct and only absent because
        # it lives on a sibling tab. sibling_dfs is deliberately bounded by the
        # caller (same statement type, e.g. all BS tabs for a BS account) —
        # not the whole workbook — to keep the false-negative risk low.
        for sib in sibling_dfs or []:
            values += cls._values_for_one_df(sib)
        return cls(values)

    def matches(self, target: float) -> bool:
        """±5% tolerance (rounding noise) at every scale; near-exact below that.

        Compares MAGNITUDES: extract_amounts() drops the leading sign, so a negative
        source cell (e.g. retained earnings -70,769,000) must still match a clause
        amount parsed as +70,769,000.

        The sub-CNY1m tier used to be a tight max(1, 1%) — meant for minor
        per-unit display rounding (e.g. 54,950 vs 54,948) — but Chinese
        commentary routinely displays sub-million amounts rounded to 1
        decimal of 万 (nearest 1,000), e.g. 11,555 written as "1.2万元"
        (=12,000, a 445 / 3.9% difference) or 10,335 as "1.0万元" (=10,000,
        335 / 3.2%). Both are correct, conventional roundings that the tight
        tier flagged as "hallucination" — and since a deterministic
        hallucination verdict is authoritative over the LLM's own (correct)
        judgement (_combine_verdict), that false flag couldn't be overridden.
        A flat 500 floor covers near-exact small values that used to hit the
        max(1,...) branch; 5% (matching the >=1m tier) covers 万-rounding at
        any sub-million magnitude.
        """
        t = abs(target)
        for v in self.values:
            a = abs(v)
            if a == 0:
                # A genuine zero source cell should only match a target that
                # ALSO rounds to zero — the 万-rounding tolerance below is
                # for rounding noise around a real nonzero figure, not for
                # letting an arbitrary small number match "nothing there".
                if round(t) == 0:
                    return True
                continue
            if abs(t - a) <= max(500.0, 0.05 * a):
                return True
        return False


def ground_amounts(clause: str, source: SourceIndex) -> Optional[Dict[str, Any]]:
    """Deterministic verdict for a clause based on its money amounts.

    Returns None when the clause has no groundable amount (defer to the LLM/soft
    judgement). Otherwise returns a clause-review dict with a confidence.
    """
    amounts = extract_amounts(clause)
    if not amounts:
        return None
    unmatched = [a for a in amounts if not source.matches(a)]
    if unmatched:
        return {
            "supported": False,
            "category": "hallucination",
            "conf": 0.9,
            "reason": f"Amount(s) {', '.join(f'{u:,.0f}' for u in unmatched)} not found in source data within tolerance.",
        }
    return {
        "supported": True,
        "category": "data-backed",
        "conf": 1.0,
        "reason": "All amounts matched source data within tolerance.",
    }


# Causal / inference / projection language that needs a soft (non-numeric)
# judgement — a clause containing these but no checkable amount is "reasoning"
# unless the LLM verified it against notes/remarks.
_CAUSAL_RE = re.compile(
    r"driven by|attributed to|reflect|due to|owing to|as a result|because|"
    r"thanks to|annualis|recurring|did not recur|no material|management (?:said|stated|noted)|"
    r"由于|反映|主要系|主要由于|预计|年化|归因于|得益于",
    re.IGNORECASE,
)


def _has_causal_language(clause: str) -> bool:
    return bool(_CAUSAL_RE.search(clause or ""))


def _norm_clause_key(text: str) -> str:
    return re.sub(r"\s+", "", str(text or "")).lower().strip(_CLAUSE_BOUNDARY_CHARS)


def _lookup_llm_review(clause: str, llm_reviews: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    """Find the LLM review whose clause best overlaps this segmented clause."""
    key = _norm_clause_key(clause)
    if not key:
        return None
    best = None
    best_len = 0
    for r in llm_reviews or []:
        rk = _norm_clause_key(r.get("clause", ""))
        if not rk:
            continue
        if rk in key or key in rk:
            overlap = min(len(rk), len(key))
            if overlap > best_len:
                best, best_len = r, overlap
    return best


# Confidence floors per source of verdict.
_CONF_DET_HALLUCINATION = 0.9
_CONF_DET_DATA_BACKED = 1.0
_CONF_LLM_FLAG = 0.7
_CONF_DEFAULT_REASONING = 0.5


def _combine_verdict(clause: str, det: Optional[Dict[str, Any]],
                     llm: Optional[Dict[str, Any]], highlight_min_conf: float) -> Dict[str, Any]:
    """Merge deterministic number-grounding with the LLM's soft judgement.

    Precedence: a deterministic unmatched-amount hallucination is authoritative
    (the model cannot override hard arithmetic). When amounts all match, an LLM
    *reasoning* flag is preserved (numbers fine, inference unsupported) but an LLM
    *number-hallucination* claim is dropped (it was a false positive). Clauses with
    no checkable amount defer to the LLM; absent that, causal language => reasoning.
    """
    llm_cat = str((llm or {}).get("category") or "").lower()
    llm_supported = bool((llm or {}).get("supported")) if llm else True

    if det and det["category"] == "hallucination":
        category, supported, conf, reason = "hallucination", False, _CONF_DET_HALLUCINATION, det["reason"]
    elif det and det["category"] == "data-backed":
        if llm and llm_cat == "reasoning" and not llm_supported:
            category, supported, conf = "reasoning", False, _CONF_LLM_FLAG
            reason = (llm or {}).get("reason") or "Numbers verified; inference not directly supported."
        else:
            # numbers matched -> drop any LLM 'hallucination' false positive
            category, supported, conf, reason = "data-backed", True, _CONF_DET_DATA_BACKED, det["reason"]
    elif llm and llm_cat in ("reasoning", "hallucination") and not llm_supported:
        category, supported, conf = llm_cat, False, _CONF_LLM_FLAG
        reason = (llm or {}).get("reason") or "Flagged by validator."
    elif _has_causal_language(clause):
        category, supported, conf = "reasoning", False, _CONF_DEFAULT_REASONING
        reason = "Causal/inference clause with no figure to verify against source."
    else:
        category, supported, conf, reason = "data-backed", True, _CONF_DET_DATA_BACKED, "No checkable figure; no causal claim."

    # Confidence gate: low-confidence flags are demoted so they don't highlight
    # inline (keeps false positives low — the user's stated priority).
    if not supported and conf < highlight_min_conf:
        category, supported = "data-backed", True
    return {"clause": clause, "supported": supported, "category": category, "reason": reason}


def verify_commentary(final_content: str, df, llm_clause_reviews: Optional[List[Dict[str, Any]]] = None,
                      *, highlight_min_conf: float = 0.6,
                      sibling_dfs: Optional[List[Any]] = None) -> List[Dict[str, Any]]:
    """Authoritative clause_reviews: deterministic number-grounding layered over the
    LLM's soft reasoning judgement. Each clause is a verbatim substring of
    final_content, so highlighting matches by exact offset. Returns the existing
    clause_reviews shape [{clause, supported, category, reason}].

    sibling_dfs (optional): other accounts' DataFrames — same statement type as
    this account, per caller — so a legitimate cross-tab reference (e.g. an
    "Other payables" note citing the bank loan balance that actually lives on
    the "Long-term loans" tab) can be grounded instead of false-flagged."""
    source = SourceIndex.from_df(df, sibling_dfs=sibling_dfs)
    llm_reviews = llm_clause_reviews or []
    out: List[Dict[str, Any]] = []
    for _s, _e, clause in segment_clauses(final_content):
        det = ground_amounts(clause, source)
        llm = _lookup_llm_review(clause, llm_reviews)
        out.append(_combine_verdict(clause, det, llm, highlight_min_conf))
    return out


def build_highlighted_commentary_html(final_content: str, clause_reviews: List[Dict[str, Any]]) -> str:
    """
    Render final commentary HTML with unsupported clauses highlighted.
    Uses category-specific CSS classes: fdd-hallucination-clause (red — the more
    severe, unsupported-by-data class) and fdd-reasoning-clause (orange — milder
    inference). Colours are defined in fdd_app.py.
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
        # category is set by _normalize_clause_review (always one of data-backed /
        # reasoning / hallucination); the "reasoning" fallback only guards a caller
        # that bypasses normalization, and matches the normalizer's own default.
        category = str(review.get("category") or "reasoning").lower()
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
            category=html.escape(str(review.get("category") or "reasoning")),
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
        self.language = normalize_language_code(language)

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
            return (
                "使用自然段落表达，不要使用要点、粗体或元评论。"
                "金额格式：人民币<数字>万元 / 人民币<数字>亿元（1位小数），"
                "金额前缀'人民币'与数字之间不留空格，例如'人民币59.3万元'、'人民币1.6亿元'。"
                "篇幅控制：资产负债表科目每条评论约30-70字（最多3-4句）；利润表科目约80-150字。"
                "日期格式统一为'yyyy年mm月dd日'。"
            )
        return (
            "Use natural paragraph-style reporting; no bullets, bold, or meta-commentary. "
            "Currency format: CNY<number> million (1 decimal place) for amounts >= CNY1 million, "
            "and CNY<comma-separated integer> for amounts under CNY1 million. "
            "NO space between 'CNY' and the number (e.g. 'CNY59.3 million', 'CNY54,950', 'CNY0.2 million'). "
            "NEVER use 'K' suffix (no 'CNY 78.2K') — render sub-million amounts as comma-separated integers (e.g. 'CNY78,200'). "
            "Length cap: balance-sheet bullets 30-70 words (max 3-4 sentences); income-statement bullets 80-150 words. "
            "Dates in 'dd mmmm yyyy' format. "
            "Keep a professional finance due diligence tone."
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
        language = normalize_language_code(language)
        agent_key = self.normalize_agent_name(agent_name)
        if agent_key == "1_Generator":
            generic_prompts = self.mappings_data.get("_default_subagent_1", {}).get(language, {})
            return generic_prompts.get("system_prompt", ""), generic_prompts.get("user_prompt", "")

        prompt_data = self.prompts_data.get(agent_key, {}).get(language, {})
        return prompt_data.get("system_prompt", ""), prompt_data.get("user_prompt", "")

    def get_prompt_pair(self, agent_name: str, language: str, mapping_key: str) -> Tuple[str, str]:
        language = normalize_language_code(language)
        agent_key = self.normalize_agent_name(agent_name)
        resolved_mapping_key = self.resolve_mapping_key(mapping_key)

        if agent_key == "1_Generator":
            default_system_prompt, default_user_prompt = self.get_agent_defaults(agent_name, language)

            account_data = self.mappings_data.get(resolved_mapping_key, {})
            if not account_data:
                fallback_section = self._fallback_mapping_section(mapping_key)
                if fallback_section:
                    account_data = self.mappings_data.get(fallback_section, {})
            raw_sp = account_data.get("subagent_1_prompts") or {}
            account_prompts = (raw_sp if isinstance(raw_sp, dict) else {}).get(language, {})
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
                return (
                    f"这是资产负债表科目。首句必须仅说明截至{effective_date}的最新期末余额（单一期间，不要罗列所有期间），"
                    f"并描述其构成（如：'截至{effective_date}余额为人民币X万元，主要为[构成项]'或'截至{effective_date}余额合计人民币X万元，主要包括[各组成项]'）。"
                    "首句不得罗列所有报告期间余额（避免'截至A、B、C日余额分别为X、Y、Z'式开篇），"
                    "也不得以年度对比句开篇（不得以'X较上年增加/减少'或'X同比上升/下降'作为首句）。"
                    "首句之后，再描述构成项目、对手方/集中度、合同条款及重要备注说明。"
                    "如跨期变动重大且数据支持，可简略提及前期余额，但不得作为开篇。"
                    f"请使用时点表述如【截至{effective_date}】，不要写成期间表述。"
                )
            if statement_type == "IS":
                period_label = build_income_statement_period_label(
                    effective_date,
                    months=annualization_months if isinstance(annualization_months, (int, float)) else None,
                    fiscal_year_end_month=fiscal_year_end_month if isinstance(fiscal_year_end_month, (int, float)) else None,
                    fiscal_year_end_day=fiscal_year_end_day if isinstance(fiscal_year_end_day, (int, float)) else None,
                    language="Chi",
                )
                is_partial = isinstance(annualization_months, (int, float)) and int(annualization_months) < 12
                partial_note = ""
                if is_partial:
                    months_int = int(annualization_months)
                    annualized_label = f"{period_label}（年化）"
                    partial_note = (
                        f"最新期间为截至{effective_date}的{months_int}个月（期间标签：{period_label}），属于不完整年度。"
                        f"跨期比较时，请优先使用年化后数据（x12/{months_int}，已预计算为【{annualized_label}】列）进行同口径对比，"
                        "并在评论中注明该数据已年化。"
                    )
                return (
                    "这是利润表科目。首句必须以构成开篇，描述该科目主要包含哪些项目"
                    "（例如：'X主要从租金收入及物业管理费收入产生营业收入，比例约为50:50'或"
                    "'主要包括房屋折旧费用人民币A万元、物业管理费人民币B万元、......'）。"
                    "不得以孤立的趋势句开篇（避免'营业收入由X增长至Y'式开篇）。"
                    "对每一重要构成项，应在句中提供所有报告期间的金额（例如：'物业管理费分别为人民币150万元、180万元、210万元，"
                    "于FY19、FY20、FY21期间发生'），而不是仅提供最新一期的金额。"
                    "构成与多期金额之后，如有重大变动，可在数据/备注支持下说明驱动因素。"
                    f"描述目标期间时，请使用【于{period_label}期间】或【在{period_label}内】等期间表述，"
                    f"不要写成【截至{effective_date}止】或时点余额表述。{partial_note}"
                    "有右侧备注的科目优先讨论。"
                )
            return "请根据科目属性正确区分时点表述与期间表述。"

        if statement_type == "BS":
            return (
                f"This is a balance-sheet item. The FIRST sentence must state ONLY the latest period-end balance as at {effective_date} "
                f"(a single period, not a list of all periods) and describe what it comprises — e.g., 'the balance as at {effective_date} "
                f"represented CNY X million of [composition]' or 'the balance as at {effective_date} totalled CNY X million, mainly entailing [components]'. "
                f"Do NOT dump all reporting periods in the opening sentence (avoid 'the balance as at A, B and C was X, Y and Z respectively'). "
                f"Do NOT open with a year-over-year movement sentence ('X increased/decreased from Y to Z'). "
                f"After the opening, describe composition, counterparty/concentration, terms, and any material remarks supported by the data. "
                f"Prior-period balances may appear briefly only when the movement is material and the data supports the explanation. "
                f"Use point-in-time wording such as 'as at {effective_date}', not period-flow wording."
            )
        if statement_type == "IS":
            period_label = build_income_statement_period_label(
                effective_date,
                months=annualization_months if isinstance(annualization_months, (int, float)) else None,
                fiscal_year_end_month=fiscal_year_end_month if isinstance(fiscal_year_end_month, (int, float)) else None,
                fiscal_year_end_day=fiscal_year_end_day if isinstance(fiscal_year_end_day, (int, float)) else None,
                language="Eng",
            )
            is_partial = isinstance(annualization_months, (int, float)) and int(annualization_months) < 12
            partial_note = ""
            if is_partial:
                months_int = int(annualization_months)
                annualized_label = f"{period_label} annualised"
                partial_note = (
                    f" The latest period covers {months_int} months ending {effective_date} (period label: {period_label}) — this is a partial year. "
                    f"For cross-year comparisons, use the annualized figures (×12/{months_int}), pre-calculated in the '{annualized_label}' column, "
                    f"and note the annualization in the commentary."
                )
            return (
                f"This is an income-statement item. The FIRST sentence must lead with COMPOSITION — describe what the line mainly comprises "
                f"(e.g., 'X mainly generated revenue from leasing income and property management service income with around a 50:50 ratio' "
                f"or 'mainly comprised depreciation and amortisation of CNY A, property management costs of CNY B, and...'). "
                f"Do NOT open with an isolated trend sentence ('Revenue increased from X to Y'). "
                f"For each material component, give the amounts across ALL reporting periods inline (e.g., 'CNY 1.5 million, CNY 1.8 million, "
                f"and CNY 2.1 million property management service costs incurred in FY19, FY20 and FY21 respectively'), not just the latest period figure. "
                f"After composition with multi-year amounts, state the driver of any material movement, supported by the data or remarks. "
                f"Refer to the target period with flow wording such as 'during {period_label}' or 'during the Period', not 'as at'.{partial_note} "
                f"Prioritize line items with supporting remarks."
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
                    else '- 以目标期间作为最新余额/最新期间表述的基础。若为利润表科目，优先使用推荐叙述期间标签并采用"于...期间/在...内"表达，而不是"截至...止期间"。同时使用更早的示意性调整后期间进行趋势、比较、交叉检查及重大变动分析。'
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
            # Format the patterns dict into clean numbered examples so the
            # LLM sees readable text rather than a Python dict repr.
            if isinstance(patterns, dict):
                examples = []
                for idx, (_, v) in enumerate(patterns.items(), 1):
                    text = str(v or "").strip()
                    if text and text.upper() != "N/A":
                        examples.append(f"Example {idx}: {text}")
                patterns = "\n".join(examples) if examples else ""
            format_params["patterns"] = str(patterns or "").strip()

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
import re as _re_client
from typing import Dict, List, Optional, Any
import httpx
from openai import OpenAI, AzureOpenAI
import logging

from .financial_common import package_file_path

_REJECTED_PARAM_RE = _re_client.compile(r"'param':\s*'([a-zA-Z_][a-zA-Z0-9_]*)'")


def _extract_rejected_param(exc: Exception) -> Optional[str]:
    """Best-effort extraction of the param name an OpenAI-style 400 rejected.

    Gateways evolve — every new sampling knob this model doesn't support (so
    far: temperature, max_tokens, top_p) would otherwise need its own
    hardcoded keyword check. Instead, read the SAME machine-readable 'param'
    field OpenAI's error body already gives us (checked first on the SDK's
    parsed .body, falling back to a regex over str(exc) for other raise
    shapes), so a NEW unsupported param self-heals without a code change.
    """
    body = getattr(exc, 'body', None)
    if isinstance(body, dict):
        err = body.get('error')
        if isinstance(err, dict) and err.get('param'):
            return str(err['param'])
    match = _REJECTED_PARAM_RE.search(str(exc))
    return match.group(1) if match else None


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
        config_path: Optional[str] = None,
        model_name: Optional[str] = None,
    ):
        """
        Initialize AIClient with specified model and agent configuration.

        Args:
            model_type: Type of model ('openai', 'local', 'deepseek', 'workbench')
            agent_name: Name of the agent ('agent_1', 'agent_2', 'agent_3', 'agent_4')
            language: Language for prompts ('Eng' or 'Chi')
            use_heuristic: Whether to use heuristic mode instead of AI
            config_path: Path to config file (optional)
            model_name: Optional specific model id within the provider (e.g. pick
                GPT-5.4 instead of the provider's configured default GPT-5.5).
                Overrides config_details['chat_model'] after config load.
        """
        self.model_type_requested = model_type
        self.model_name_requested = model_name
        self.agent_name = agent_name
        # Normalize "Chn" -> "Chi" once here: AIClient is built by every pipeline
        # entry point, so this guarantees prompt lookups and styling get the right
        # code regardless of which entry point (pipeline, reprompt, validator) ran.
        language = normalize_language_code(language)
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
        if model_name:
            # Copy before mutating — config_details may be a reference into the
            # shared, cached config dict; overriding in place would leak this
            # instance's model choice into every other AIClient using the same
            # provider (e.g. a concurrent thread on a different agent/account).
            self.config_details = dict(self.config_details)
            self.config_details['chat_model'] = model_name

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
                http_client=httpx.Client(verify=False, timeout=httpx.Timeout(120.0, connect=10.0))
            )
            model = self.config_details['chat_model']

        elif self.model_type == 'workbench':
            # KPMG Workbench gateway (Azure OpenAI-compatible). The gateway
            # requires the subscription key duplicated as a header (not just
            # api_key) plus enterprise billing/routing headers. charge_code and
            # region_override are configurable per config.yml; defaults match
            # the values in the reference snippet.
            subscription_key = self.config_details['api_key']
            headers = {
                'Ocp-Apim-Subscription-Key': subscription_key,
                'x-kpmg-charge-code': str(self.config_details.get('charge_code') or '0000'),
                'x-kpmg-region-override': str(self.config_details.get('region_override') or 'westeurope'),
            }
            self._workbench_headers = headers
            client = AzureOpenAI(
                api_key=subscription_key,
                base_url=self.config_details['api_base'],
                api_version=self.config_details['api_version'],
                default_headers=headers,
                http_client=httpx.Client(verify=False, timeout=httpx.Timeout(180.0, connect=10.0)),
            )
            model = self.config_details['chat_model']

        elif self.model_type == 'local':
            client = OpenAI(
                base_url=self.config_details['api_base'],
                api_key=self.config_details['api_key'],
                http_client=httpx.Client(verify=False, timeout=httpx.Timeout(120.0, connect=10.0))
            )
            model = self.config_details['chat_model']

        elif self.model_type == 'deepseek':
            client = OpenAI(
                base_url=self.config_details['api_base'],
                api_key=self.config_details['api_key'],
                http_client=httpx.Client(verify=False, timeout=httpx.Timeout(120.0, connect=10.0))
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
        presence_penalty: Optional[float] = None,
        allow_thinking: Optional[bool] = None,
        reasoning_effort: Optional[str] = None,
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
            reasoning_effort: Per-call override for workbench.reasoning_effort (optional,
                falls back to the provider-level config default — see agents.*.reasoning_effort)

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
        effective_reasoning_effort = (
            reasoning_effort if reasoning_effort is not None else self.config_details.get('reasoning_effort')
        )
        
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

            elif self.model_type == 'workbench':
                # KPMG's reference snippet sends the auth/billing headers BOTH as
                # default_headers (set once on the client) and as extra_headers
                # on every call — keep both, since some gateway configurations
                # only honour per-call headers.
                wb_params = dict(params)
                wb_params['extra_headers'] = dict(getattr(self, '_workbench_headers', {}) or {})

                # Model capability flags — config-driven so a new deployment's
                # quirks (confirmed against the real gateway, not guessed) can
                # be tuned in config.yml without a code change. Defaults match
                # the GPT-5-class reasoning models currently behind this
                # gateway, which reject temperature, top_p, frequency_penalty,
                # and presence_penalty outright (reasoning models don't expose
                # traditional sampling controls) and rename max_tokens.
                if not bool(self.config_details.get('supports_temperature', False)):
                    wb_params.pop('temperature', None)
                if not bool(self.config_details.get('supports_sampling_params', False)):
                    for _p in ('top_p', 'frequency_penalty', 'presence_penalty'):
                        wb_params.pop(_p, None)
                if bool(self.config_details.get('use_max_completion_tokens', True)) and 'max_tokens' in wb_params:
                    wb_params['max_completion_tokens'] = wb_params.pop('max_tokens')
                if effective_reasoning_effort:
                    wb_params['reasoning_effort'] = effective_reasoning_effort

                # Reasoning models spend max_completion_tokens on HIDDEN
                # reasoning tokens before the visible answer — the per-agent
                # max_tokens values in config.yml (1200-1400) are sized for
                # local Qwen3, which has no hidden-token cost for structured
                # stages (allow_thinking=false skips <think> entirely). For
                # workbench, a complex multi-component account (e.g.
                # Investment properties, Operating costs) can exhaust that
                # budget on reasoning alone with reasoning_effort=high,
                # leaving nothing for the answer — a real 200/no-exception
                # response with EMPTY content (confirmed via inputs/昆山.xlsx:
                # 2/20 accounts came back blank). min_max_tokens is a floor,
                # not a replacement — it only raises the budget when the
                # agent's own setting is smaller.
                min_max_tokens = self.config_details.get('min_max_tokens')
                if min_max_tokens:
                    current_budget = wb_params.get('max_completion_tokens', wb_params.get('max_tokens', 0)) or 0
                    if current_budget < int(min_max_tokens):
                        budget_key = 'max_completion_tokens' if 'max_completion_tokens' in wb_params else 'max_tokens'
                        wb_params[budget_key] = int(min_max_tokens)

                _MAX_TOKENS_ALIASES = ('max_tokens', 'max_completion_tokens')

                def _drop_rejected_param(exc: Exception) -> bool:
                    """Read the 'param' the gateway's 400 named and adjust
                    wb_params generically — no hardcoded keyword list, so a
                    param we haven't seen yet (the config flags above only
                    cover known ones) still self-heals. Returns True if
                    wb_params changed (worth retrying)."""
                    param = _extract_rejected_param(exc)
                    if not param:
                        return False
                    if param in _MAX_TOKENS_ALIASES:
                        other = 'max_completion_tokens' if param == 'max_tokens' else 'max_tokens'
                        if param in wb_params:
                            wb_params[other] = wb_params.pop(param)
                            return True
                        return False
                    if param in wb_params:
                        wb_params.pop(param, None)
                        return True
                    return False

                try:
                    response = response_method(**wb_params)
                except TypeError:
                    # SDK-level rejection of a kwarg (older openai-python that
                    # doesn't know reasoning_effort at all) — drop and retry.
                    wb_params.pop('reasoning_effort', None)
                    response = response_method(**wb_params)
                except Exception as first_exc:
                    # Gateway-level rejection (400) — retry a few times in case
                    # dropping one param surfaces another unsupported one right
                    # after (seen in practice: temperature, then top_p, then
                    # frequency_penalty, one at a time). Uses an explicit
                    # sentinel-based loop (not for/else + bare raise) so the
                    # re-raised exception is always the exact object we mean.
                    last_exc = first_exc
                    response = None
                    for _ in range(4):
                        if not _drop_rejected_param(last_exc):
                            raise last_exc
                        try:
                            response = response_method(**wb_params)
                            break
                        except Exception as retry_exc:
                            last_exc = retry_exc
                    if response is None:
                        raise last_exc
                content = self._extract_response_content(response)
                usage = getattr(response, 'usage', None)

            elif self.model_type == 'local':
                params['stream'] = True
                # Qwen3: turn OFF thinking for structured/JSON stages (Auditor,
                # Validator) — it adds latency and pollutes output with no quality
                # gain. vLLM/SGLang honour chat_template_kwargs.enable_thinking.
                # If a different server rejects the field, retry without it
                # (strip_thinking still cleans any inline <think> as a fallback).
                if allow_thinking is False:
                    local_params = dict(params)
                    local_params['extra_body'] = {'chat_template_kwargs': {'enable_thinking': False}}
                    try:
                        response = response_method(**local_params)
                    except TypeError:
                        response = response_method(**params)
                    except Exception as exc:
                        if 'extra_body' in str(exc) or 'enable_thinking' in str(exc) or 'chat_template' in str(exc):
                            response = response_method(**params)
                        else:
                            raise
                else:
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


class _StageCircuitBreaker:
    """Per-stage circuit breaker for LLM calls.

    Tracks consecutive failures per agent stage. After ``threshold`` consecutive
    failures (across all worker threads), the breaker is OPEN for that stage —
    further calls raise immediately instead of waiting for retries to time out.
    The breaker resets on the first success or when ``reset_stage`` is called
    at the start of each new stage.

    Rationale: when the LLM endpoint is stressed, every call wastes
    ``timeout × retries`` (30s × 3 = 90s) before falling back. After 3-5
    failures in a row that's clearly the new normal — fail fast for the rest
    of the batch and let the deterministic fallback take over.
    """

    def __init__(self, threshold: int = 4):
        self._threshold = threshold
        self._consecutive_failures: Dict[str, int] = {}
        self._lock = threading.Lock()

    def record_success(self, stage: str) -> None:
        with self._lock:
            self._consecutive_failures[stage] = 0

    def record_failure(self, stage: str) -> None:
        with self._lock:
            self._consecutive_failures[stage] = self._consecutive_failures.get(stage, 0) + 1

    def is_open(self, stage: str) -> bool:
        with self._lock:
            return self._consecutive_failures.get(stage, 0) >= self._threshold

    def reset_stage(self, stage: str) -> None:
        with self._lock:
            self._consecutive_failures[stage] = 0


_PIPELINE_BREAKER = _StageCircuitBreaker(threshold=4)


# Active pipeline stages, in order. The Refiner (subagent_3 / 3_Refiner) is
# DORMANT by design — its config/prompts are retained for reference but it is
# deliberately omitted here, so the runtime is a 3-stage pipeline despite the
# "4 subagent" naming elsewhere. Add ("subagent_3", "Refiner") between Auditor
# and Validator to re-enable it.
SUBAGENT_SEQUENCE = [
    ("subagent_1", "Generator"),
    ("subagent_2", "Auditor"),
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
    content = strip_thinking(content)  # drop any Qwen3 <think> block first
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
                allow_thinking=agent_cfg.get("allow_thinking"),
                reasoning_effort=agent_cfg.get("reasoning_effort"),
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
        cfg = agent_config or {}
        default_target = int(cfg.get("reduction_target_pct", 64))
        statement_type = str(prompt_manager.get_mapping_component(mapping_key, component="type") or "").strip().upper()
        if statement_type == "BS" and cfg.get("reduction_target_pct_bs") is not None:
            reduction_target = int(cfg["reduction_target_pct_bs"])
        else:
            reduction_target = default_target
        return {
            "previous_content": previous_output,
            "original_length": len(previous_output or ""),
            "reduction_target_pct": str(reduction_target),
        }
    if agent_name == "subagent_4":
        cfg = agent_config or {}
        return {
            "content": previous_output,
            "materiality_threshold_pct": str(cfg.get("materiality_threshold_pct", 5)),
        }
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
    dfs: Optional[Dict[str, pd.DataFrame]] = None,
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

        # Auto-reprompt on timeout. The user does not want to see "AI call
        # timed out" placeholder text in the commentary; retry up to twice
        # before falling back. Retries use the SAME prompt — if the API is
        # truly unresponsive the timeout will fire again, but typically a
        # transient slow response succeeds on retry.
        # Circuit breaker: when the API is stressed, repeatedly retrying just
        # wastes wall-clock time. After N consecutive failures across the
        # current stage, skip remaining LLM calls in this stage and fall
        # through to the deterministic fallback. Reset between stages so each
        # stage gets a fresh chance.
        if _PIPELINE_BREAKER.is_open(agent_name):
            logger.logger.warning(
                "[%s] %s: circuit breaker OPEN — skipping LLM call, using fallback",
                agent_name, mapping_key,
            )
            raise RuntimeError("Circuit breaker open for this stage")

        response = None
        last_exc: Optional[Exception] = None
        # Exponential backoff between retries: 0s, 1s, 2s — gives the API a
        # brief moment to recover without compounding total wall-time too much.
        # Earlier (0, 2, 5) added 7s per failed call which compounded to
        # several extra minutes on a stressed run; (0, 1, 2) saves ~4s/call.
        retry_backoffs = [0.0, 1.0, 2.0]
        for attempt in range(1, 4):  # 1 initial + 2 retries
            if attempt > 1 and retry_backoffs[attempt - 1] > 0:
                time.sleep(retry_backoffs[attempt - 1])
            try:
                # Local models (Qwen3-32B etc.) are slow and emit <think> tokens, so the
                # default must give the SAME headroom to every stage. Previously only the
                # Validator set call_timeout=90 in config; Generator/Auditor fell to 30s and
                # timed out on large prompts. Derive a model-aware default so all stages match.
                _default_timeout = 90 if getattr(ai_helper, "model_type", "") == "local" else 30
                call_timeout = int(agent_cfg.get("call_timeout", _default_timeout))
                response = _run_ai_call(ai_helper, user_prompt, system_prompt, agent_name, timeout=call_timeout)
                last_exc = None
                _PIPELINE_BREAKER.record_success(agent_name)
                if attempt > 1:
                    logger.logger.info(
                        "[%s] %s: succeeded on retry %s/2",
                        agent_name, mapping_key, attempt - 1,
                    )
                break
            except (TimeoutError, Exception) as exc:
                last_exc = exc
                logger.logger.warning(
                    "[%s] %s: AI call attempt %s failed (%s); %s",
                    agent_name, mapping_key, attempt, str(exc)[:80],
                    "retrying" if attempt < 3 else "giving up",
                )
        if response is None:
            _PIPELINE_BREAKER.record_failure(agent_name)
            raise last_exc if last_exc is not None else RuntimeError("AI call failed with no exception captured")

        raw_content = response["content"].strip().replace("\n\n", "\n").replace("\n \n", "\n")

        if logger.debug_mode:
            logger.log_debug("RAW_OUTPUT", mapping_key, "Agent=%s len=%s" % (agent_name, len(raw_content)), raw_content)

        content, metadata = _finalize_agent_content(
            agent_name=agent_name,
            raw_content=raw_content,
            previous_output=previous_output,
            language=ai_helper.language,
        )

        # Deterministic hallucination/reasoning verification: layer Python
        # number-grounding over the validator's soft judgement. This catches
        # fabricated figures the weak local model misses AND drops its false
        # positives on figures that actually match the source. Wrapped so a
        # verifier error never breaks the pipeline (keeps the LLM clause_reviews).
        if agent_name == "subagent_4" and metadata and df is not None:
            try:
                sibling_dfs = None
                if dfs:
                    statement_type = prompt_manager.get_mapping_component(mapping_key, component="type")
                    if statement_type:
                        sibling_dfs = [
                            other_df for other_key, other_df in dfs.items()
                            if other_key != mapping_key
                            and prompt_manager.get_mapping_component(other_key, component="type") == statement_type
                        ]
                metadata["clause_reviews"] = verify_commentary(
                    content, df, metadata.get("clause_reviews"),
                    sibling_dfs=sibling_dfs,
                )
            except Exception as exc:  # pragma: no cover - defensive
                logger.logger.warning("[verify_commentary] %s: %s", mapping_key, exc)

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
            fallback = _build_deterministic_fallback_bullet(mapping_key, df, ai_helper.language)
            if fallback:
                logger.logger.warning(
                    "[%s] %s: AI unavailable after retries; using deterministic data-only fallback",
                    agent_name, mapping_key,
                )
                return mapping_key, fallback, {"used_fallback": True, "fallback_reason": str(exc)[:120]}
            return mapping_key, f"Content generation failed for {mapping_key}: {str(exc)[:100]}", {}
        if previous_output and str(previous_output).strip():
            return mapping_key, previous_output, {}
        return mapping_key, f"Content generation incomplete for {mapping_key}: {str(exc)[:100]}", {}


def _build_deterministic_fallback_bullet(
    mapping_key: str,
    df: Optional[pd.DataFrame],
    language: str,
) -> str:
    """Produce a minimal data-only bullet when the LLM is unreachable.

    The bullet uses the latest period's total from the dataframe (last numeric
    column header is treated as the latest reporting date). Output format
    mirrors the project's reference style; downstream stages may still polish.
    """
    if df is None or df.empty:
        return ""
    try:
        from .financial_display_format import format_number_chinese
    except Exception:
        return ""

    numeric_cols = [c for c in df.columns[1:] if not str(c).endswith("_formatted")]
    if not numeric_cols:
        return ""
    latest_col = numeric_cols[-1]
    try:
        if not pd.api.types.is_numeric_dtype(df[latest_col]):
            return ""
        total = float(df[latest_col].fillna(0).sum())
    except Exception:
        return ""
    if total == 0:
        return ""

    date_label = str(latest_col).strip()
    iso_match = re.match(r"(\d{4})-(\d{2})-(\d{2})", date_label)
    if iso_match:
        if language == "Chi":
            date_label = f"{iso_match.group(1)}年{int(iso_match.group(2))}月{int(iso_match.group(3))}日"
        else:
            date_label = _iso_to_long_date(iso_match)

    formatted = format_number_chinese(total, language)
    if language == "Chi":
        return f"截至{date_label}余额合计{formatted}。（自动摘要 - AI 暂时不可用）"
    return (
        f"the balance as at {date_label} totalled {formatted}. "
        "(Auto-summary while AI service was unavailable; please refresh this account when service recovers.)"
    )


def _resolve_max_workers(ai_helper, max_workers: Optional[int]) -> int:
    """Single source of truth for worker-count defaults across every thread
    pool in the pipeline (main stages, feedback loop, ensure-validation
    re-runs). An explicit caller-supplied value always wins; otherwise the
    provider's own <provider>.max_workers in config.yml (e.g. workbench: 4,
    validated against that gateway — see test_workbench_concurrency.py)
    wins; otherwise "local" (the self-hosted model) defaults to 1 — this
    project's local server serves one request effectively serially, so
    concurrent requests just queue rather than reduce wall time — and every
    other provider defaults to 4 (parallel by default, no UI toggle needed).
    Both defaults are overridable per-provider via config.yml."""
    if max_workers is not None:
        return max_workers
    _configured = (getattr(ai_helper, "config_details", None) or {}).get("max_workers")
    if _configured:
        return int(_configured)
    _model_type = getattr(ai_helper, "model_type", "")
    return 1 if _model_type == "local" else 4


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
    max_workers = _resolve_max_workers(ai_helper, max_workers)

    # Reset the circuit breaker at the start of each stage so a fresh
    # opportunity is given even if a prior stage tripped it.
    _PIPELINE_BREAKER.reset_stage(agent_name)

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
                    dfs,
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
                dfs,
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
    model_name: Optional[str] = None,
) -> Dict[str, Dict[str, str]]:
    """Run the 4-agent FDD pipeline with optional progress callbacks."""
    # Normalise UI language codes ("Chn" → "Chi") to match prompt-file keys.
    language = normalize_language_code(language)
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
        model_name=model_name,
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
        # Reset breakers: a stage tripped during the main run would otherwise stay
        # OPEN and make the feedback loop fail-fast (skip) every account silently.
        for _stage, _ in SUBAGENT_SEQUENCE:
            _PIPELINE_BREAKER.reset_stage(_stage)
        logger.logger.info(
            "Starting feedback loop (max_retries=%s, threshold=%.2f)",
            feedback_config["max_retries"],
            feedback_config["unsupported_threshold"],
        )
        eligible_keys = [k for k in mapping_keys if k in results and k in dfs]
        if use_multithreading and len(eligible_keys) > 1:
            fb_workers = _resolve_max_workers(ai_helper, max_workers)
            with ThreadPoolExecutor(max_workers=fb_workers) as executor:
                futures = {
                    executor.submit(
                        _run_feedback_loop_for_key,
                        key=key,
                        dfs=dfs,
                        results=results,
                        ai_helper=ai_helper,
                        prompt_manager=prompt_manager,
                        logger=logger,
                        feedback_config=feedback_config,
                        user_comments=user_comments,
                        progress_callback=progress_callback,
                    ): key
                    for key in eligible_keys
                }
                for future in as_completed(futures):
                    key = futures[future]
                    try:
                        retries = future.result()
                    except Exception as exc:
                        logger.logger.warning("[FeedbackLoop] %s: failed: %s", key, exc)
                        continue
                    if retries > 0:
                        logger.logger.info("[FeedbackLoop] %s: completed with %s retry(ies)", key, retries)
        else:
            for key in eligible_keys:
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

    # Ensure every account with a final commentary has hallucination/reasoning
    # clause_reviews so the UI can highlight them. If the Validator stage
    # didn't produce clause_reviews (timeout, parse failure, etc.), run a
    # one-shot validator pass on the final text. Runs only for accounts that
    # need it, in parallel.
    # Reset the Validator breaker first so a tripped main-run stage doesn't make
    # this re-validation pass fail-fast for every account.
    _PIPELINE_BREAKER.reset_stage("subagent_4")
    _ensure_clause_reviews_on_final(
        results=results,
        dfs=dfs,
        ai_helper=ai_helper,
        prompt_manager=prompt_manager,
        logger=logger,
        use_multithreading=use_multithreading,
        max_workers=max_workers,
        user_comments=user_comments,
    )

    logger.finalize(results)
    return results


def _ensure_clause_reviews_on_final(
    *,
    results: Dict[str, Dict[str, str]],
    dfs: Dict[str, pd.DataFrame],
    ai_helper,
    prompt_manager: PromptEngine,
    logger: PipelineRunLogger,
    use_multithreading: bool,
    max_workers: Optional[int],
    user_comments: Optional[Dict[str, str]] = None,
) -> None:
    """Re-run Validator on accounts whose final commentary lacks clause_reviews."""
    needs_validation: List[str] = []
    for key, result in results.items():
        if not isinstance(result, dict):
            continue
        final_text = str(result.get("final") or "").strip()
        if not final_text:
            continue
        validation = result.get("agent_4_validation") or {}
        if isinstance(validation, dict) and validation.get("clause_reviews"):
            continue
        if key not in dfs:
            continue
        needs_validation.append(key)

    if not needs_validation:
        return

    logger.logger.info(
        "[EnsureValidation] %s account(s) need fresh clause_reviews", len(needs_validation),
    )

    def _run_one(key: str):
        final_text = str((results.get(key) or {}).get("final") or "").strip()
        try:
            _k, content, metadata = process_single_agent_item(
                "subagent_4", key, dfs.get(key), ai_helper, prompt_manager, logger,
                previous_output=final_text,
                user_comment=(user_comments or {}).get(key, ""),
                dfs=dfs,
            )
            if isinstance(metadata, dict) and metadata.get("clause_reviews"):
                # Keep the original final text (don't overwrite); just attach
                # clause_reviews so highlighting works.
                results[key]["agent_4_validation"] = metadata
                logger.logger.info(
                    "[EnsureValidation] %s: validated %s clause(s)", key,
                    len(metadata.get("clause_reviews", [])),
                )
        except Exception as exc:
            logger.logger.warning("[EnsureValidation] %s: failed: %s", key, exc)

    if use_multithreading and len(needs_validation) > 1:
        ev_workers = _resolve_max_workers(ai_helper, max_workers)
        with ThreadPoolExecutor(max_workers=ev_workers) as executor:
            list(executor.map(_run_one, needs_validation))
    else:
        for key in needs_validation:
            _run_one(key)


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

        # Re-run Auditor (polish) so the validator sees refined output, not raw
        # generator output. Skipping this step caused more clauses to be flagged
        # as unsupported and triggered unnecessary further retries.
        _key, audit_content, _audit_meta = process_single_agent_item(
            "subagent_2", key, dfs.get(key), ai_helper, prompt_manager, logger,
            previous_output=gen_content,
            user_comment=base_user_comment,
        )
        results[key]["subagent_2"] = audit_content
        results[key]["feedback_retry_%s_agent_2" % retry_num] = audit_content

        # Re-run validator on the polished output
        _key, val_content, val_metadata = process_single_agent_item(
            "subagent_4", key, dfs.get(key), ai_helper, prompt_manager, logger,
            previous_output=audit_content,
            user_comment=base_user_comment,
            dfs=dfs,
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
    model_name: Optional[str] = None,
) -> Dict[str, Dict[str, str]]:
    """Regenerate selected items, then immediately revalidate the revised output."""
    # Mirror run_ai_pipeline_with_progress: normalise "Chn" → "Chi" so the Chinese
    # reprompt path resolves prompts and applies Chinese styling (not English).
    language = normalize_language_code(language)
    logger = PipelineRunLogger()
    prompt_manager = get_prompt_engine()
    ai_helper = AIClient(
        model_type=model_type,
        agent_name="content_pipeline",
        language=language,
        use_heuristic=use_heuristic,
        model_name=model_name,
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
            dfs=dfs,
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

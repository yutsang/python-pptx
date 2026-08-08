from __future__ import annotations

# re-added: bound by an import in another section of the pre-split module
from typing import Any, Dict, List, Optional

import re
from typing import Any

from ..financial_common import normalize_chinese_punctuation_in_text
from ..keyword_registry import KNOWN_TRANSLATIONS

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

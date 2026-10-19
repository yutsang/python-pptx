from __future__ import annotations

# re-added: bound by an import in another section of the pre-split module
from ..workbook import INTERNAL_ROW_KEY
from .config import get_safe_grounding_include_siblings
from typing import Any, Dict, List, Optional
from typing import Any, Dict, Optional, Tuple

"""
Utilities for parsing validator clause annotations and rendering highlights.
"""



import html
import json
import re
from typing import Any, Dict, List


#: The defect vocabulary every deterministic detector writes into a
#: clause_review's optional `code` key. A plain frozenset, not an Enum: the
#: codes are carried in dicts that get yaml.dump'd into the run archive and
#: json'd into the UI, so they have to be plain strings on the wire anyway, and
#: an Enum would only add a .value at every write site.
#:
#: Routing (M4): AMOUNT_SCALE_ERROR and DATE_UNSUPPORTED are patchable — the
#: source figure is known. Every COMPOSITION_* code routes to regeneration
#: (a sentence whose parts do not add up cannot be fixed by swapping a token),
#: and so do both LLM_* codes.
#:
#: LLM_UNSUPPORTED_FACT is the dominant class, not an afterthought: measured
#: across 247 archived run folders (25,136 clause reviews, 418 unsupported) it
#: covers 123 LLM-prose hallucinations against 99 deterministic amount misses.
#: UNCATEGORIZED exists so the M6 histogram has no unlabelled bucket -- 25
#: archived reviews carry `category: None`.
#: CLAIM_MISSING is reserved for M5's claim contracts and must NEVER enter
#: clause_reviews: an entry there perturbs the unsupported-ratio retry gate and
#: colours the exported client deck.
DEFECT_CODES = frozenset({
    "AMOUNT_UNSUPPORTED",
    "AMOUNT_SCALE_ERROR",
    "DATE_UNSUPPORTED",
    "DIRECTION_MISMATCH",
    "COMPOSITION_GAP",
    "COMPOSITION_UNIT_ERROR",
    "COMPOSITION_DOUBLE_COUNT",
    "LLM_UNSUPPORTED_CAUSE",
    "LLM_UNSUPPORTED_FACT",
    "UNCATEGORIZED",
    "CLAIM_MISSING",
})


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
    """Format the clause problems worth re-generating for into concise
    feedback for the generator reprompt.

    Hallucinations first and, when there are any, ONLY those: telling the
    generator to also "fix" a supportable inference it was asked to make
    just trains it to write blander commentary. Falls back to every
    unsupported clause when the retry was triggered by a broad
    unsupported RATIO rather than a specific hallucination."""
    from .pipeline import count_defective_clauses  # local: breaks the validator<->pipeline import cycle
    defective = count_defective_clauses(clause_reviews)
    unsupported = defective or [
        r for r in (clause_reviews or []) if isinstance(r, dict) and not r.get("supported")
    ]
    if not unsupported:
        return ""
    if language == "Chi":
        header = "验证器标记了以下不支持的内容需要修正:\n"
        template = "- 分句: \"{clause}\" — 问题: {reason}"
        fix_scale = "（写错数量级：来源数字为 {expected:,.0f}，请照此改写）"
    else:
        header = "The validator flagged the following unsupported clauses for correction:\n"
        template = '- Clause: "{clause}" — Issue: {reason}'
        fix_scale = "(wrong scale; the source figure is {expected:,.0f} — use it)"
    items = []
    for r in unsupported[:5]:
        reason = str(r.get("reason", ""))[:200]
        # A typed defect carrying an `expected` says what the right figure is.
        # "wrong scale; the source figure is 7.8万元 (sheet 'Cash' row 12)" is a
        # correctable instruction; "not found in source data" is a shrug, and
        # the model answers it by deleting the sentence. This improves the
        # EXISTING regeneration retry, before any patch machinery lands.
        if r.get("code") == "AMOUNT_SCALE_ERROR" and isinstance(r.get("expected"), (int, float)):
            reason = f"{reason} {fix_scale.format(expected=float(r['expected']))}"
        elif r.get("code") == "DATE_UNSUPPORTED" and r.get("expected"):
            reason = f"{reason} [{r.get('code')} -> {r.get('expected')}]"
        items.append(template.format(clause=str(r.get("clause", ""))[:120], reason=reason))
    return header + "\n".join(items)


def _fallback_clause_reviews(final_content: str) -> List[Dict[str, Any]]:
    """Deterministic clause_reviews when the validator JSON can't be parsed.

    Segments the text and marks each clause supported (no inline highlight) so the
    UI/feedback metric have a non-empty, well-shaped list to work with. Real
    number-grounding happens in verify_commentary (which has the source df); this
    is only the no-JSON safety net that keeps the shape valid and stops re-loops.
    """
    reviews: List[Dict[str, Any]] = []
    for start, end, clause in segment_clauses(final_content):
        reviews.append({
            "clause": clause,
            "supported": True,
            "category": "data-backed",
            "reason": "Auto-segmented (validator JSON unparseable).",
            "span": [int(start), int(end)],
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
#: A figure carries only the precision it was WRITTEN with. 「0.01亿元」 is a
#: two-decimal 亿元 rendering, so it means anything within 0.005亿 = 500,000 of
#: 1,000,000 -- the text simply does not distinguish 1,062,997 from 1,000,000.
#: Checking such a figure at +/-5% (= +/-50,000 here) is ten times tighter than
#: the notation, and a check tighter than its own input produces false positives
#: by construction: a real deck flagged 「装修工程0.01亿元」 as a hallucination in
#: four consecutive runs, against an analysis_cell of 1,062,997 sitting +6.3%
#: away. The list it appears in sums to the account's net value.
#:
#: Only bites when the unit is coarse for the magnitude. 「43.1万元」 gets a
#: half-ulp of 500 and 「1.63亿元」 of 500,000, both far under the 5% band that
#: already applies, so nothing else moves.
_WRITTEN_UNIT_SCALE = (("亿", 1e8), ("億", 1e8), ("万", 1e4), ("萬", 1e4))
_WRITTEN_NUMBER = re.compile(r"(\d[\d,]*)(?:\.(\d+))?")


def display_half_ulp(written: str) -> float:
    """Half the last place of `written`, in base units. 0.0 if unreadable."""
    match = _WRITTEN_NUMBER.search(str(written or ""))
    if not match:
        return 0.0
    decimals = len(match.group(2) or "")
    scale = 1.0
    for token, value in _WRITTEN_UNIT_SCALE:
        if token in written:
            scale = value
            break
    return 0.5 * (10.0 ** -decimals) * scale


_AMT_YI = re.compile(r"(\d[\d,]*(?:\.\d+)?)\s*亿")
_AMT_WAN = re.compile(r"(\d[\d,]*(?:\.\d+)?)\s*万")
_AMT_CUR_PREFIX = re.compile(r"(?:CNY|RMB|USD|HKD|US\$|\$|人民币|人民幣)\s*(\d[\d,]*(?:\.\d+)?)", re.IGNORECASE)
_AMT_GROUPED = re.compile(r"(?<![\d.])(\d{1,3}(?:,\d{3})+(?:\.\d+)?)")
# The suffixes the model actually writes, which nothing here used to read: the
# Generator emits thousands as "CNY 835.3K" and remarks carry a bare "198m" for
# millions. Both parsed as 835.3 / 198 -- off by three and six orders of
# magnitude -- and only the (over-broad) grounding pool kept them from being
# flagged. Trailing guard is (?![A-Za-z0-9]) rather than \b so "5m2" and the
# "mn"/"million" spellings _AMT_MILLION already owns are left alone.
#
# The number needs TWO integer digits or an explicit decimal. A single bare
# digit is refused because segment_clauses cuts on ". " and archived text
# contains broken decimals ("CNY 971. 4K"), so a clause can BEGIN "4K as at 31
# December 2020" -- measured, 7 of the first 26 replayed flips were exactly that
# fragment being read as CNY4,000. Real writes always carry the decimal
# ("687.0K") or several digits ("198m"), so nothing legitimate is lost.
_AMT_THOUSAND = re.compile(r"(?:CNY|RMB|USD|HKD|US\$|\$|人民币|人民幣)?\s*(\d[\d,]+(?:\.\d+)?|\d+\.\d+)\s*[Kk](?![A-Za-z0-9])")
_AMT_BARE_MILLION = re.compile(r"(?:CNY|RMB|USD|HKD|US\$|\$|人民币|人民幣)?\s*(\d[\d,]+(?:\.\d+)?|\d+\.\d+)\s*[mM](?![A-Za-z0-9])")


def _to_float(token: str) -> Optional[float]:
    try:
        return float(str(token).replace(",", ""))
    except (TypeError, ValueError):
        return None


# An amount's span is NOT the regex's group(0). Every pattern above stops at the
# digits-plus-scale-word and deliberately leaves the currency out on one side or
# the other: _AMT_WAN matches "7.8万" out of "7.8万元", _AMT_GROUPED matches
# "1,234,567" out of "1,234,567元", and _AMT_CUR_PREFIX matches "人民币1,234,567"
# but stops before its 元. A patch that rewrote group(0) would leave the unit
# dangling or duplicated ("人民币7.8万元" -> "人民币9.1万元元"), which is what
# broke two earlier attempts at this. So the stored span extends over a
# following unit/currency word and back over a leading currency prefix.
#
# The trailing guard refuses a suffix immediately followed by a digit, so the
# "CNY" of "…1,234 CNY5,678" is read as the NEXT amount's prefix rather than
# swallowed as this one's suffix.
_AMT_SUFFIX_RE = re.compile(r"[ \t]{0,3}(?:元|圆|yuan|million|mn|RMB|CNY)", re.IGNORECASE)
_AMT_PREFIX_RE = re.compile(r"(?:人民币|人民幣|CNY|RMB|USD|HKD|US\$|\$)[ \t]{0,3}$", re.IGNORECASE)


def _patch_span(text: str, start: int, end: int) -> Tuple[int, int]:
    """Widen a raw regex match to the span a repair would have to rewrite."""
    suffix = _AMT_SUFFIX_RE.match(text, end)
    if suffix and not text[suffix.end():suffix.end() + 1].isdigit():
        end = suffix.end()
    prefix = _AMT_PREFIX_RE.search(text, 0, start)
    if prefix:
        start = prefix.start()
    return (start, end)


def extract_amount_spans(clause: str) -> List[Tuple[float, int, int]]:
    """Every money amount in `clause` as (value, start, end).

    Offsets index `clause` itself: the per-pass blanking below substitutes a
    same-width run of spaces, and the fullwidth-comma normalisation swaps a
    same-width character, so nothing shifts. `clause[start:end]` is the amount
    WITH its unit/currency (see _patch_span), which is the unit of text a repair
    replaces.
    """
    amounts: List[Tuple[float, int, int]] = []
    # A fullwidth comma inside a figure is a thousands separator, not a clause
    # break. The Auditor introduces them -- measured across archived runs, they
    # go from 13 occurrences in Generator output to 46 in Auditor output -- and
    # without this normalisation "5，271.8万元" parses as 271.8万 plus a stray 5,
    # i.e. an order of magnitude low, which then reads as a fabricated figure.
    # The wide grounding pool used to hide this by matching the wrong value
    # anyway; once the pool was narrowed it surfaces as a false hallucination.
    # Substituting a same-width character keeps every offset into `clause`
    # valid, which the clause spans downstream depend on.
    work = clause.replace("，", ",")
    # Each pass blanks the span it consumed so a later, looser pass cannot
    # re-count the same figure (e.g. 'CNY5.8 million' -> 5.8e6 only; 'CNY54,950'
    # -> 54950 once, not also via the grouped-thousands pass). The spelled-out
    # 'million'/'mn' pass MUST stay ahead of the bare-'m' one, or the latter
    # eats "5.8 m" out of "5.8 million" and leaves "illion" behind.
    for rx, scale in ((_AMT_MILLION, 1e6), (_AMT_YI, 1e8), (_AMT_WAN, 1e4),
                      (_AMT_BARE_MILLION, 1e6), (_AMT_THOUSAND, 1e3),
                      (_AMT_CUR_PREFIX, 1.0), (_AMT_GROUPED, 1.0)):
        def _sub(m: "re.Match") -> str:
            v = _to_float(m.group(1))
            if v is not None:
                start, end = _patch_span(work, m.start(), m.end())
                amounts.append((v * scale, start, end))
            return " " * len(m.group(0))
        work = rx.sub(_sub, work)
    amounts.sort(key=lambda triple: triple[1])
    return amounts


def extract_amounts(clause: str) -> List[float]:
    """Extract absolute money amounts (scaled to base units) from a clause.

    Scale-bearing forms (million / 万 / 亿 / m / K) are parsed first and their
    matched text blanked so a following currency-prefix/grouped pass cannot
    double-count the same figure (e.g. 'CNY5.8 million' must yield 5.8e6, not
    also 5.8). Values only; extract_amount_spans carries the offsets a repair
    needs.
    """
    return [value for value, _s, _e in extract_amount_spans(clause)]


_BARE_NUMBER_RE = re.compile(r"\d[\d,]*(?:\.\d+)?")

#: attrs keys deliberately withheld from every prompt, so they must not ground
#: anything either. `working_remark_notes` collects the rows that
#: _build_working_remark_note diverts (schedules.py: they `continue` instead of
#: becoming row_entries), meaning the model never sees those figures. Grounding
#: against them certifies the model correct on numbers it was never shown —
#: measured, half their material numbers reached the pool.
# financials_authority records the tab totals the Financials figures REPLACED;
# letting those back into the pool would ground exactly the figures the
# override exists to keep out of the deck.
_ATTR_KEYS_WITHHELD_FROM_PROMPTS = ("working_remark_notes", "financials_authority")


def _attr_text_blob(df) -> str:
    """Concatenate the free-text in df.attrs that the model is actually shown
    (supporting_notes, table_linked_remarks, adjacent_detail_rows, rhs context,
    etc.) so figures cited in the NOTES — not just the numeric table — can ground
    a clause. Many legitimate figures (registered capital, audit fees, USD
    amounts) live only in the remarks. Keys in
    _ATTR_KEYS_WITHHELD_FROM_PROMPTS are skipped.

    Also stringifies bare int/float leaf values, not just strings -- confirmed via
    real screenshots that adjacent_detail_rows' own primary value (e.g. a "房产税"
    sub-line-item that only ever appears as a detail row, never its own numeric-
    table row) is stored as a raw Python float (workbook.py's own `effective_value`
    keyed by the column label), not a string. The walk below used to silently drop
    any non-str/dict/list/tuple leaf, so that number was NEVER in the grounding
    pool at all -- not even in its raw, non-annualized form, which is a more
    fundamental gap than the annualization-only fix layered on top of this
    function's output (see SourceIndex._values_for_one_df)."""
    parts: List[str] = []
    attrs = getattr(df, "attrs", None) or {}

    def walk(v):
        if isinstance(v, str):
            parts.append(v)
        elif isinstance(v, bool):
            return
        elif isinstance(v, (int, float)):
            parts.append(str(v))
        elif isinstance(v, dict):
            for vv in v.values():
                walk(vv)
        elif isinstance(v, (list, tuple)):
            for vv in v:
                walk(vv)

    for key, value in attrs.items():
        if key in _ATTR_KEYS_WITHHELD_FROM_PROMPTS:
            continue
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


#: Every fact carries `kind`. It is not decoration: `classify_miss` looks at
#: `cell`/`column_total` ONLY, the window-sum tolerance is picked by it, and M4's
#: repair router refuses to cite anything synthetic as a patch source.
FACT_KINDS = frozenset({
    "cell", "column_total", "window_sum", "analysis_cell",
    "note_number", "annualized_note", "sibling_cell", "date",
    "prompt_residual", "period_movement",
})

#: Facts a scale classification and a repair may cite: the account's own real
#: cells and the column totals of its own frame. Nothing synthetic, nothing from
#: another tab, nothing scraped out of a note blob.
#:
#: `analysis_cell` -- the account's OWN multi-period frame, which is also what
#: the Generator was shown -- is in the set, though the plan's rule named only
#: cell/column_total. Measured on two real books by corrupting every own cell
#: x1000: without it 75.4%/82.2% of corruptions resolve to exactly one scale
#: factor and 14.8%/3.6% resolve to none (the true source was a historical
#: period that lives only in the analysis frame); with it, 89.1%/83.9% resolve
#: uniquely, 0%/1.9% resolve to none, and ambiguity rises by 1.0pp/0.0pp. It is
#: still the account's own df, which is the part of the rule that carries the
#: result -- siblings, window sums and note-blob numbers stay out.
_OWN_HARD_KINDS = ("cell", "column_total", "analysis_cell")

#: Tolerance for an adjacent-window sum, against max(500, 5%) for a real cell.
#: See SourceIndex.matches for why the two differ.
_WINDOW_SUM_REL_TOL = 0.005

#: classify_miss's own rule. Both numbers are load-bearing and measured; see
#: SourceIndex.classify_miss before touching either.
#:
#: The 100x pair is not symmetry for its own sake -- it is the commonest unit
#: confusion this pipeline actually produces, and leaving it out made every
#: instance land in AMOUNT_UNSUPPORTED (a full rewrite) instead of
#: AMOUNT_SCALE_ERROR (a one-token repair). A real shipped bullet stated an
#: accounts-payable balance of 6.2亿元 against a tab totalling CNY6.2 million:
#: the frame's display unit is millions, the model read 6.2 and reached for 亿
#: instead of 百万, and 亿/百万 is 100. 10x covers the same slip one step down
#: (万 vs 十万). The Chinese magnitude ladder is 万=1e4 and 亿=1e8, so an error
#: here is almost never a clean 1e3.
_SCALE_MISS_FACTORS = (100.0, 0.01, 1000.0, 0.001, 10000.0, 0.0001, 10.0, 0.1)
_SCALE_MISS_REL_TOL = 0.005


def _fact(value, kind: str, *, sheet=None, row_idx=None, row_desc=None,
          col_label=None, multiplier=None, row_range=None) -> Dict[str, Any]:
    """One source fact, with every field coerced to a plain builtin.

    Coercion is not tidiness. These records are written into the run archive by
    an unsanitized `yaml.dump` and into the UI by `json.dumps`, and a numpy
    scalar or a pandas Timestamp arriving there raises or emits a
    `!!python/object` tag that nothing can read back. Values that cannot be
    coerced become None rather than travelling as their original object.
    """
    def _s(v):
        return None if v is None else str(v)

    def _i(v):
        try:
            return None if v is None else int(v)
        except (TypeError, ValueError):
            return None

    def _f(v):
        try:
            return None if v is None else float(v)
        except (TypeError, ValueError):
            return None

    rec: Dict[str, Any] = {
        "value": _f(value),
        "kind": str(kind),
        "sheet": _s(sheet),
        "row_desc": _s(row_desc),
        "col_label": _s(col_label),
        "multiplier": _f(multiplier),
    }
    # Synthetic kinds span rows, so they carry a row_range where a cell carries
    # a row_idx (the plan's shape). Spans are plain 2-element int lists.
    if row_range is not None:
        lo, hi = row_range
        rec["row_range"] = [_i(lo), _i(hi)]
    else:
        rec["row_idx"] = _i(row_idx)
    return rec


def describe_fact(fact: Optional[Dict[str, Any]]) -> str:
    """Human-readable provenance for one fact, for a reason string or a prompt."""
    if not isinstance(fact, dict):
        return "unknown source"
    value = fact.get("value")
    # A date fact's value is a "YYYY-MM-DD" string, not a number.
    head = f"{value:,.0f}" if isinstance(value, (int, float)) else str(value)
    bits = [head, f"({fact.get('kind')}"]
    if fact.get("sheet"):
        bits.append(f"sheet '{fact['sheet']}'")
    if fact.get("row_idx") is not None:
        bits.append(f"row {fact['row_idx']}")
    elif fact.get("row_range"):
        lo, hi = fact["row_range"]
        bits.append(f"rows {lo}-{hi}")
    if fact.get("row_desc"):
        bits.append(f"'{str(fact['row_desc'])[:40]}'")
    if fact.get("col_label"):
        bits.append(f"col {fact['col_label']}")
    return bits[0] + " " + " ".join(bits[1:]) + ")"


def _frame_sheet(df) -> Optional[str]:
    attrs = getattr(df, "attrs", None) or {}
    integrity = attrs.get("integrity") or {}
    return attrs.get("source_sheet_name") or integrity.get("sheet_name") or attrs.get("block_title")


def _frame_multiplier(df):
    attrs = getattr(df, "attrs", None) or {}
    return attrs.get("source_multiplier")


def _row_provenance(df) -> Tuple[List[Optional[str]], List[Optional[int]], List[Optional[str]]]:
    """(row_desc, row_idx, row_type) per POSITIONAL row of `df`.

    row_desc comes from the frame's first column (the block-title column, whose
    cells are the line-item descriptions); row_idx from INTERNAL_ROW_KEY, which
    is the raw sheet row and the only thing that points back at the workbook;
    row_type from attrs['row_types_by_description'] ('detail' / 'breakdown' /
    'subtotal' / 'total').
    """
    n = len(df.index)
    descs: List[Optional[str]] = [None] * n
    idxs: List[Optional[int]] = [None] * n
    types: List[Optional[str]] = [None] * n
    try:
        label_col = next((c for c in df.columns if c != INTERNAL_ROW_KEY), None)
        row_types = (getattr(df, "attrs", None) or {}).get("row_types_by_description") or {}
        if label_col is not None:
            for pos, cell in enumerate(df[label_col].tolist()):
                descs[pos] = None if cell is None else str(cell)
                types[pos] = row_types.get(descs[pos])
        if INTERNAL_ROW_KEY in df.columns:
            for pos, cell in enumerate(df[INTERNAL_ROW_KEY].tolist()):
                try:
                    idxs[pos] = int(cell)
                except (TypeError, ValueError):
                    idxs[pos] = None
    except Exception:
        pass
    return descs, idxs, types


class SourceIndex:
    """Source facts present in an account's data, for grounding amounts.

    Stores fact RECORDS (value + provenance), not bare floats: a verdict of
    "matched source data" was unfalsifiable after the fact, which is why
    measuring the pool's behaviour needed a bit-exact reimplementation of this
    class. `self.values` is kept as the parallel float list because the decoy
    harness and diagnostics report pool size from it.
    """

    def __init__(self, values: List[Any]):
        # Accepts either a bare List[float] (the CLI/diagnostic path, which has
        # no frame to draw provenance from) or a list of fact records. Bare
        # floats become kind="cell" with null provenance, so classify_miss still
        # returns a scale verdict for them -- just with source_ref=None.
        self.facts: List[Dict[str, Any]] = []
        for item in values or []:
            if item is None:
                continue
            if isinstance(item, dict):
                if item.get("value") is not None:
                    self.facts.append(item)
            else:
                coerced = _fact(item, "cell")
                if coerced["value"] is not None:
                    self.facts.append(coerced)
        self.values: List[float] = [f["value"] for f in self.facts]
        self.date_facts: List[Dict[str, Any]] = []

    @staticmethod
    def _adjacent_window_sums(col_vals: List[float], max_window: int = 4) -> List[float]:
        """Sums of every run of 2..max_window CONSECUTIVE rows (sheet order, as
        the column already preserves it) — commentary legitimately groups a
        handful of neighbouring breakdown lines into one figure (e.g. "CNY322,116
        of property[-related fees]" = 4 adjacent line items in Other payables
        that were never a labelled subtotal in the sheet). Bounded to small
        windows, not a full subset-sum search, to keep this O(n) and keep the
        false-negative risk (a genuinely wrong number coincidentally matching
        some arbitrary window) low.

        Kept as the plain-list form for callers with no frame (the CLI mirrors
        it); _window_facts below is what the pool actually uses, and it applies
        the bounds recorded there.
        """
        sums: List[float] = []
        n = len(col_vals)
        for window in range(2, max_window + 1):
            for start in range(0, n - window + 1):
                sums.append(sum(col_vals[start:start + window]))
        return sums

    @classmethod
    def _window_facts(cls, cells: List[Dict[str, Any]], row_types: List[Optional[str]],
                      max_window: int = 4) -> List[Dict[str, Any]]:
        """Adjacent-window sums, bounded three ways.

        These sums turn n rows into ~4n unlabelled acceptors, and measurement
        made them the single largest remaining acceptor of decoy figures once
        siblings and the formatted columns were out of the pool: with them
        stubbed out entirely, decoy acceptance fell from 41.4% to 35.7% on one
        book. Deleting them is not an option — the case they were added for is
        real and documented above (four adjacent Other-payables line items cited
        as one figure) — so they are bounded instead:

        * MAIN FRAME ONLY. The caller passes windows=False for the nested
          analysis frame, for sibling tabs and for note-blob numbers. A window
          over another tab's rows was never a claim anyone wrote.
        * NO TOTAL OR SUBTOTAL ROW inside the window, and every row in the
          window must share one row_type. A run that mixes a parent line with
          the breakdown lines underneath it (AR lists 8 counterparties under
          租金收入) is a double-count, not a grouping, and a run that swallows
          Total is the column total plus noise.
        * matched at a tighter tolerance than a real cell (see `matches`).

        Unknown row types are treated as their own group, so a frame carrying no
        row_types_by_description keeps the old behaviour.
        """
        facts: List[Dict[str, Any]] = []
        n = len(cells)
        for window in range(2, max_window + 1):
            for start in range(0, n - window + 1):
                chunk = cells[start:start + window]
                types = {row_types[c["_pos"]] for c in chunk}
                if len(types) != 1:
                    continue
                if next(iter(types)) in ("total", "subtotal"):
                    continue
                first, last = chunk[0], chunk[-1]
                facts.append(_fact(
                    sum(c["value"] for c in chunk), "window_sum",
                    sheet=first.get("sheet"),
                    row_range=(first.get("row_idx"), last.get("row_idx")),
                    row_desc=f"{first.get('row_desc')}…{last.get('row_desc')}",
                    col_label=first.get("col_label"),
                    multiplier=first.get("multiplier"),
                ))
        return facts

    @classmethod
    def _column_facts(cls, df, skip_cols: tuple = (), *, cell_kind: str = "cell",
                      total_kind: str = "column_total", windows: bool = True) -> List[Dict[str, Any]]:
        facts: List[Dict[str, Any]] = []
        descs, row_idxs, row_types = _row_provenance(df)
        sheet = _frame_sheet(df)
        multiplier = _frame_multiplier(df)
        for col in df.columns:
            if col in skip_cols:
                continue
            series = df[col]
            cells: List[Dict[str, Any]] = []
            if getattr(series, "dtype", None) is not None and series.dtype.kind in "if":
                pairs = [(pos, float(v)) for pos, v in enumerate(series.tolist())
                         if v is not None and v == v]  # NaN != NaN — same set dropna() gave
            else:
                pairs = []
                for pos, cell in enumerate(series.tolist()):
                    v = _to_float(cell) if isinstance(cell, (int, float, str)) else None
                    if v is not None:
                        pairs.append((pos, v))
            for pos, value in pairs:
                rec = _fact(value, cell_kind, sheet=sheet, row_idx=row_idxs[pos],
                            row_desc=descs[pos], col_label=col, multiplier=multiplier)
                rec["_pos"] = pos
                cells.append(rec)
            facts += cells
            # Add the column total — commentary frequently cites a total that
            # isn't a single cell; including it avoids false hallucination flags.
            if cells:
                facts.append(_fact(
                    sum(c["value"] for c in cells), total_kind, sheet=sheet,
                    row_range=(row_idxs[cells[0]["_pos"]], row_idxs[cells[-1]["_pos"]]),
                    row_desc="column total", col_label=col, multiplier=multiplier,
                ))
            if windows:
                facts += cls._window_facts(cells, row_types)
        for rec in facts:
            rec.pop("_pos", None)
        return facts

    @classmethod
    def _column_values(cls, df, skip_cols: tuple = ()) -> List[float]:
        """Float-only view of _column_facts, for callers that want the old shape."""
        return [f["value"] for f in cls._column_facts(df, skip_cols=skip_cols)]

    @classmethod
    def _period_movement_facts(cls, analysis_df, sheet) -> List[Dict[str, Any]]:
        """Each row's change between ADJACENT periods of the analysis frame.

        "净值较上年末下降0.06亿元" was flagged as a hallucination. It is not:
        the analysis table the Generator is handed carries 净值小计 at 1.98 for
        the prior period and 1.92 for the latest, and 1.98 - 1.92 = 0.06 exactly.
        The model did correct arithmetic on the numbers in front of it, and the
        pool held both endpoints but not the difference, so the one figure an FDD
        bullet exists to state -- how much a balance moved -- could not be
        grounded. Two of six defects on a real run were this, both ranked high,
        both sent to a repair that could only make them worse.

        Adjacent periods only, so this is the movement a reader means by
        "较上年末": no first-to-last spans, no cross-row differences, nothing
        between non-neighbouring columns. A row that did not move contributes
        nothing. On the real frame above that is +63% pool values, which is a
        genuine widening and the reason for the narrow rule -- filed under its
        own kind and OUT of _OWN_HARD_KINDS, so a movement grounds a clause but
        can never be cited as a repair source or steer a scale classification.

        Price paid, measured (replay_verification.py --decoys, same workbook,
        same archived run, this method monkeypatched to return [] for the
        control): overall false-accept 30.1% -> 31.8%, +1.7pp across 2,455
        decoys; the worst single class (scale_div10) 35.2% -> 37.3%. Well
        under the 40% ceiling the plan set for narrowing it to total rows only.
        """
        facts: List[Dict[str, Any]] = []
        try:
            # This module does not import pandas -- every frame here is
            # duck-typed, the same way _column_facts reads one. A first cut used
            # pd.api.types/pd.isna, which are NameErrors this except swallowed,
            # so the whole thing silently produced nothing.
            skip = set(cls._non_amount_cols(analysis_df))
            columns = [c for c in list(analysis_df.columns)[1:] if c not in skip]
            descs, row_idxs, _types = _row_provenance(analysis_df)

            def column_values(col) -> Optional[List[Optional[float]]]:
                series = analysis_df[col]
                dtype = getattr(series, "dtype", None)
                if dtype is not None and getattr(dtype, "kind", "") in "if":
                    return [float(v) if v is not None and v == v else None
                            for v in series.tolist()]
                out: List[Optional[float]] = []
                for cell in series.tolist():
                    out.append(_to_float(cell) if isinstance(cell, (int, float, str)) else None)
                return out if any(v is not None for v in out) else None

            resolved = [(col, column_values(col)) for col in columns]
            periods = [(col, vals) for col, vals in resolved if vals is not None]
            if len(periods) < 2:
                return facts
            for (prev_col, prev_vals), (curr_col, curr_vals) in zip(periods, periods[1:]):
                for pos in range(min(len(prev_vals), len(curr_vals))):
                    prev_v, curr_v = prev_vals[pos], curr_vals[pos]
                    if prev_v is None or curr_v is None:
                        continue
                    delta = curr_v - prev_v
                    if abs(delta) < 1e-9:
                        continue
                    facts.append(_fact(
                        delta, "period_movement", sheet=sheet,
                        row_idx=row_idxs[pos] if pos < len(row_idxs) else None,
                        row_desc=descs[pos] if pos < len(descs) else None,
                        col_label="%s -> %s" % (prev_col, curr_col),
                    ))
                    # A bullet states a fall as a positive magnitude -- "下降
                    # 0.06亿元", not "-0.06". Both signs are the same movement.
                    facts.append(_fact(
                        -delta, "period_movement", sheet=sheet,
                        row_idx=row_idxs[pos] if pos < len(row_idxs) else None,
                        row_desc=descs[pos] if pos < len(descs) else None,
                        col_label="%s -> %s (magnitude)" % (prev_col, curr_col),
                    ))
        except Exception:
            return facts
        return facts

    @classmethod
    def _annualised_tail_facts(cls, df, analysis_df, sheet) -> List[Dict[str, Any]]:
        """Each analysis row's part-year tail on a twelve-month basis.

        The prompt now prints these in an '<period> annualised' column and the
        trend summary and significant movements quote them, so a bullet that
        says 「年化后3,675.8万元」 is repeating a figure it was handed. Without
        them here it grounded only by landing within 5% of some other cell.
        Synthetic, so its own kind and out of _OWN_HARD_KINDS: it grounds a
        clause and is never a repair source.
        """
        facts: List[Dict[str, Any]] = []
        try:
            from ..workbook.analysis import ANNUALISED_SUFFIX, partial_tail_months
            statement_type = ((getattr(df, "attrs", None) or {}).get("integrity") or {}).get("statement_type")
            months = partial_tail_months(analysis_df, statement_type)
            if not months:
                return facts
            skip = set(cls._non_amount_cols(analysis_df))
            columns = [c for c in list(analysis_df.columns)[1:] if c not in skip]
            if not columns:
                return facts
            tail = columns[-1]
            descs, row_idxs, _types = _row_provenance(analysis_df)
            previous = analysis_df[columns[-2]].tolist() if len(columns) > 1 else []
            for pos, cell in enumerate(analysis_df[tail].tolist()):
                value = _to_float(cell) if isinstance(cell, (int, float, str)) else None
                if value is None or value != value or not value:
                    continue
                where = dict(sheet=sheet, row_idx=row_idxs[pos] if pos < len(row_idxs) else None,
                             row_desc=descs[pos] if pos < len(descs) else None)
                facts.append(_fact(value * 12.0 / months, "annualised_cell",
                                   col_label=f"{tail}{ANNUALISED_SUFFIX}", **where))
                # ...and its change on the prior period, which the significant-
                # movements table prints as `delta` on this basis. The first run
                # after that table changed flagged 「较2025年度下降141.9万元」 as a
                # hallucination twice -- 利息支出 annualised less the 2025 figure,
                # the number the prompt had just supplied -- and spent two
                # retries on it.
                prior = previous[pos] if pos < len(previous) else None
                prior = _to_float(prior) if isinstance(prior, (int, float, str)) else None
                if prior is not None and prior == prior:
                    delta = value * 12.0 / months - prior
                    if abs(delta) > 1e-9:
                        for signed, tag in ((delta, ""), (-delta, " (magnitude)")):
                            facts.append(_fact(
                                signed, "period_movement",
                                col_label=f"{columns[-2]} -> {tail}{ANNUALISED_SUFFIX}{tag}", **where))
        except Exception:
            return []
        return facts

    @staticmethod
    def _non_amount_cols(df) -> tuple:
        """Columns that are not financial amounts and must never enter the pool.

        INTERNAL_ROW_KEY holds raw sheet row indices. The `*_formatted` columns
        are the bigger contaminant: they carry the display-scaled string of the
        same figure, so every real amount also entered the pool divided by its
        display divisor, plus a mass of small values that match almost anything.
        Measured on the reference databook, dropping both takes the pooled
        small-value count from 90/83/56 to 0 and makes matches(317.0) False on 6
        of 8 sampled accounts. Nothing real is lost — the raw numeric column for
        the same date is already pooled — and across 1,171 real matched
        citations these two tiers produced ZERO sole matches.
        `_build_peer_context` and the CLI's `_numeric_values_from_df` already
        skipped them; SourceIndex was the only place that did not.
        """
        return (INTERNAL_ROW_KEY,) + tuple(
            col for col in df.columns if str(col).endswith("_formatted")
        )

    @classmethod
    def _facts_for_one_df(cls, df, *, own: bool = True) -> List[Dict[str, Any]]:
        """Every fact one frame contributes.

        `own=False` is a sibling tab: its facts are all filed as `sibling_cell`
        so `classify_miss` and any repair can exclude them by kind alone, and it
        contributes no adjacent-window sums.
        """
        facts: List[Dict[str, Any]] = []
        if df is None or not hasattr(df, "columns"):
            return facts
        cls_sheet = _frame_sheet(df)
        cell_kind = "cell" if own else "sibling_cell"
        total_kind = "column_total" if own else "sibling_cell"
        facts += cls._column_facts(df, skip_cols=cls._non_amount_cols(df),
                                   cell_kind=cell_kind, total_kind=total_kind,
                                   windows=own)
        # df is `projection_df` — a SINGLE latest-period snapshot. Multi-year
        # trend commentary ("increased from CNY384M as at 2023-12-31 to
        # CNY709M as at 2024-12-31") is written from df.attrs["prompt_analysis_df"]
        # (see _build_financial_prompt_payload's "analysis_periods" block, which
        # the Generator AND Validator both receive) — without indexing it here
        # too, every correctly-written historical-period number is invisible to
        # this grounding pool and gets falsely flagged as "hallucination", which
        # _combine_verdict then treats as authoritative over the LLM's own
        # (correct) judgement. INTERNAL_ROW_KEY was excluded here from the start —
        # it holds raw sheet row indices, not financial amounts; the same
        # exclusion (now also covering `*_formatted`) applies to the main frame
        # above, which had gone without it.
        analysis_df = df.attrs.get("prompt_analysis_df")
        if analysis_df is not None and hasattr(analysis_df, "columns"):
            facts += cls._column_facts(
                analysis_df, skip_cols=cls._non_amount_cols(analysis_df),
                cell_kind="analysis_cell" if own else "sibling_cell",
                total_kind="analysis_cell" if own else "sibling_cell",
                windows=False,
            )
            # The residual the PROMPT computed and told the model to write.
            # prompts.py works out "the largest three come to X, the remaining N
            # come to Y" and instructs the bullet to close with "其余Y为…". Y is
            # a difference, and this pool holds cells, column totals and windows
            # -- never a difference -- so the model obeyed and the verifier then
            # called the number it had just been handed a hallucination. A real
            # run did that twice in one entity, and the repair pass spent an LLM
            # call on each, both returning the text unchanged, which is what
            # happens when nothing was wrong.
            #
            # Read, never re-derived: this is the figure that was actually
            # issued. Filed under its own kind and deliberately NOT in
            # _OWN_HARD_KINDS -- it is synthetic, so it grounds the clause
            # without ever being citable as a repair source or steering a scale
            # classification.
            # Adjustment amounts read off the bridge sheets and handed to the
            # model (authority.adjustment_bridge). In no account cell, so without
            # this a correctly quoted 「示意性调整补计提X万元」 would be "not
            # found in source data" -- the shape of every own goal in this file.
            if own:
                for amount in ((df.attrs.get("adjustment_bridge") or {}).get("values") or []):
                    facts.append(_fact(amount, "adjustment_entry", sheet=cls_sheet,
                                       row_desc="adjustment posted to this account (bridge sheet)"))
            # The gap between the overridden total and the tab's own detail.
            # authority_note tells the model the detail 「加总不等于该合计」, so
            # 「与明细口径存在X万元差异」 is that difference stated -- and was
            # flagged as a hallucination and retried twice on a real run. Only
            # the DIFFERENCE grounds; the replaced tab totals stay withheld.
            if own:
                _authority = df.attrs.get("financials_authority") or analysis_df.attrs.get("financials_authority") or {}
                for _period, _pair in _authority.items():
                    try:
                        _gap = abs(float(_pair.get("tab") or 0.0) - float(_pair.get("financials") or 0.0))
                    except (TypeError, ValueError, AttributeError):
                        continue
                    if _gap:
                        facts.append(_fact(_gap, "authority_gap", sheet=cls_sheet,
                                           row_desc="total (Financials) less the tab's detail, %s" % _period))
            residual = analysis_df.attrs.get("prompt_residual")
            if own and isinstance(residual, dict) and residual.get("amount") is not None:
                desc = ("remainder handed to the model (total of %s component(s) less the "
                        "largest %s)" % (residual.get("component_count"),
                                         len(residual.get("listed") or [])))
                facts.append(_fact(residual["amount"], "prompt_residual", sheet=cls_sheet, row_desc=desc))
                # The figure as the instruction PRINTED it, when display rounding
                # moved it outside tolerance of the exact one (12,750 shown as
                # "0.01 million"). The model writes what it was told to write.
                for key_name, label in (("display_amount", desc + " [as displayed]"),
                                        ("display_total", "component total as displayed to the model")):
                    shown = residual.get(key_name)
                    if shown is not None:
                        facts.append(_fact(shown, "prompt_residual", sheet=cls_sheet, row_desc=label))
            # What the prompt budget left out, as one figure: the model was told
            # N smaller components exist and what they sum to, so 「其余N项合计Y」
            # is an instructed statement, same as the remainder above.
            budget_residual = analysis_df.attrs.get("prompt_budget_residual")
            if own and isinstance(budget_residual, dict) and budget_residual.get("amount"):
                desc = ("sum of the %s component(s) the prompt budget did not show"
                        % budget_residual.get("count"))
                facts.append(_fact(budget_residual["amount"], "prompt_residual", sheet=cls_sheet, row_desc=desc))
                if budget_residual.get("display_amount") is not None:
                    facts.append(_fact(budget_residual["display_amount"], "prompt_residual",
                                       sheet=cls_sheet, row_desc=desc + " [as displayed]"))
            if own:
                facts += cls._period_movement_facts(analysis_df, cls_sheet)
                facts += cls._annualised_tail_facts(df, analysis_df, cls_sheet)
        # Also ground against numbers cited in the supporting notes / remarks
        # (df.attrs), e.g. registered capital "7000万美元" that never appears in
        # the numeric table. Without this they were false-flagged as hallucinations.
        text_values = _numbers_in_text(_attr_text_blob(df))
        note_kind = "note_number" if own else "sibling_cell"
        facts += [_fact(v, note_kind, sheet=cls_sheet, row_desc="notes/remarks")
                  for v in text_values]
        # Detail/remark-row figures (e.g. a stamp-duty sub-line that only ever
        # appears inside a note, never as its own numeric-table row) have no
        # pre-calculated annualized column the way a main account row does
        # (see _period_reference_guidance's "预计算为...(年化)列" instruction,
        # which only covers the account's own projection_df/analysis_df
        # columns) -- so whenever the AI correctly annualizes one of these for
        # a partial reporting year, following the SAME x12/months convention
        # every main row already gets, the result is invisible to this
        # grounding pool and gets false-flagged as "hallucination" regardless
        # of whether the arithmetic is right. Confirmed via a real screenshot:
        # "印花税...2026年1-6月年化后为人民币4,485元" flagged red purely
        # because 4,485 (= the raw H1 actual x2) was never in the pool.
        integrity = df.attrs.get("integrity") or {}
        annualization_months = df.attrs.get("annualization_months")
        if annualization_months in (None, ""):
            annualization_months = integrity.get("annualization_months")
        if isinstance(annualization_months, (int, float)) and 0 < annualization_months < 12:
            factor = 12.0 / annualization_months
            # ...but only for numbers that are actually a PARTIAL period. The
            # blob carries prior-year figures too, and multiplying those files a
            # quantity nobody computed: a real run wrote 「较2025年度1,910.0万元
            # 减少63%」 and grounded it against 955.0万元 x 2, where 955.0万元 is
            # the 2025 FULL YEAR. prompts.yml:806 forbids exactly that ('把不足一年
            # 的期间与整年比较时，两边必须都用年化金额，或都用未年化金额') and the
            # pool was agreeing with the model against the prompt.
            #
            # A remark number is treated as already-annual when it matches a
            # value in one of the frame's EARLIER period columns and not the
            # final one. The stamp-duty case the widening exists for is
            # untouched by that test: its raw figure lives only in the partial
            # column, so x12/months stays in the pool.
            full_year, partial = [], []
            if analysis_df is not None and hasattr(analysis_df, "columns"):
                try:
                    amount_cols = [c for c in analysis_df.columns
                                   if c not in cls._non_amount_cols(analysis_df)]
                    for position, column in enumerate(amount_cols):
                        bucket = partial if position == len(amount_cols) - 1 else full_year
                        for value in analysis_df[column]:
                            if isinstance(value, (int, float)) and value == value and value:
                                bucket.append(abs(float(value)))
                except Exception:
                    full_year, partial = [], []

            def _near(value: float, pool: List[float]) -> bool:
                target = abs(float(value))
                return any(abs(target - other) <= max(1.0, 0.005 * max(target, other))
                           for other in pool)

            facts += [
                _fact(v * factor, "annualized_note" if own else "sibling_cell",
                      sheet=cls_sheet, row_desc=f"notes/remarks x{factor:.4g} (annualized)")
                for v in text_values
                if not (_near(v, full_year) and not _near(v, partial))
            ]
        return facts

    @classmethod
    def _values_for_one_df(cls, df) -> List[float]:
        """Float-only view of _facts_for_one_df, for callers that want the old shape."""
        return [f["value"] for f in cls._facts_for_one_df(df)]

    @classmethod
    def from_df(cls, df, sibling_dfs: Optional[List[Any]] = None) -> "SourceIndex":
        facts: List[Dict[str, Any]] = cls._facts_for_one_df(df)
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
        #
        # That case is real and the path stays. What later measurement on four
        # real databooks (96 accounts) showed is what it cost: sibling values are
        # 91.3% of the pool but carry 1.2% of the real grounding load, and with
        # them in, the pool accepted 94.4% of figures made by multiplying a real
        # cell by a random 1.15-8.0 factor and 90.9% of tenfold unit errors.
        # Without them (plus the two exclusions above) those fall to 48.7% and
        # 45.6%; ad-hoc/workbench/replay_verification.py --decoys, sampling
        # differently, reads 94.8% -> 39.4% and 85.9% -> 42.1% on two books. Not
        # yet a discriminating check — the adjacent-window sums were what was
        # left — but no longer one that accepts anything. The window sums have
        # since been bounded too (see _window_facts), taking decoy acceptance to
        # 30-34% on four books.
        # So the DEFAULT FLIPPED TO OFF (processing.grounding_include_siblings,
        # false); set it true to get the behaviour described above back for a
        # file that needs it. 711 of 30,370 replayed archived clause verdicts
        # (2.3%) flip supported -> unsupported as a result; none flip the other
        # way.
        if get_safe_grounding_include_siblings():
            for sib in sibling_dfs or []:
                facts += cls._facts_for_one_df(sib, own=False)
        index = cls(facts)
        index.date_facts = _harvest_source_date_facts(df)
        return index

    def matches(self, target: float, floor: float = 0.0) -> Optional[Dict[str, Any]]:
        """The first source fact this amount matches, or None.

        Returns the FACT, not a bool, so a caller can say which value grounded
        the clause; `if source.matches(x)` still reads as before because a fact
        record is never empty.

        ±5% tolerance (rounding noise) at every scale; near-exact below that.

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

        A window_sum is held to a tighter tolerance than a real cell. The 5%
        band exists for how a writer ROUNDS a figure (11,555 written as
        "1.2万元"); a window sum is not a figure anyone read off the sheet, it
        is one of ~4n synthetic aggregates, and giving each of them a 5% band
        is what made them the largest remaining acceptor of decoys. The
        documented case they exist for -- four adjacent Other-payables lines
        cited as "CNY322,116" -- is an EXACT sum, so it survives the tighter
        band; a 万-rounded grouping of small lines no longer does, and that
        cost is recorded rather than hidden.
        """
        t = abs(target)
        for fact in self.facts:
            a = abs(fact["value"])
            if a == 0:
                # A genuine zero source cell should only match a target that
                # ALSO rounds to zero — the 万-rounding tolerance below is
                # for rounding noise around a real nonzero figure, not for
                # letting an arbitrary small number match "nothing there".
                if round(t) == 0:
                    return fact
                continue
            if fact["kind"] == "window_sum":
                if abs(t - a) <= max(_WINDOW_SUM_REL_TOL * a, float(floor or 0.0)):
                    return fact
                continue
            if abs(t - a) <= max(500.0, 0.05 * a, float(floor or 0.0)):
                return fact
        return None

    def classify_miss(self, target: float) -> Dict[str, Any]:
        """Why an amount missed: a scale error with a source, or unsupported.

        This deliberately does NOT reuse `matches()`. Measured on the reference
        databook with the production pool, of 226 real cell values corrupted
        x1000: 82 (36%) were still accepted by the pool so nothing was ever
        flagged, 135 (60%) flagged with TWO scale factors matching, and only 9
        (4%) resolved to exactly one -- which would make a repair's uniqueness
        guard refuse essentially every case. The cause is the pool's breadth
        (window sums, siblings, note-blob numbers, annualized variants)
        combined with max(500, 5%).

        So this has its own tight lookup: the account's OWN cells and column
        totals only (see _OWN_HARD_KINDS), at <=0.5% relative with no absolute
        floor. Re-measured with exactly that rule on the same 226 corruptions,
        218 (96%) resolved to exactly one factor and none escaped detection; on
        the two books available here, 829/830 corruptions give 89.1%/83.9%
        unique and 4.9%/3.4% ambiguous. Widening either half of the rule undoes
        the result -- the pool's own tolerance is what destroyed it.
        """
        t = abs(target)
        candidates = [f for f in self.facts if f["kind"] in _OWN_HARD_KINDS and f["value"]]
        hits: List[Dict[str, Any]] = []
        for factor in _SCALE_MISS_FACTORS:
            scaled = t * factor
            for fact in candidates:
                a = abs(fact["value"])
                if a and abs(scaled - a) <= _SCALE_MISS_REL_TOL * a:
                    hits.append({"factor": float(factor), "source_ref": fact,
                                 "expected": float(fact["value"])})
                    break
        if hits:
            unique = len({h["factor"] for h in hits}) == 1
            best = hits[0]
            return {
                "code": "AMOUNT_SCALE_ERROR",
                "factor": best["factor"],
                "expected": best["expected"],
                "source_ref": best["source_ref"],
                # A repair may only act on a unique factor; two factors matching
                # means the source cannot say which figure was meant.
                "ambiguous": not unique,
                "candidates": [h["factor"] for h in hits],
            }
        nearest = None
        for fact in candidates:
            a = abs(fact["value"])
            if a and abs(t - a) <= 0.20 * a:
                if nearest is None or abs(t - a) < abs(t - abs(nearest["value"])):
                    nearest = fact
        # `nearest` is a READING HINT only, never a patch source: within 20% it
        # is as likely to be the neighbouring line item as the intended one.
        return {"code": "AMOUNT_UNSUPPORTED", "nearest": nearest}


def ground_amounts(clause: str, source: SourceIndex, *, offset: int = 0) -> Optional[Dict[str, Any]]:
    """Deterministic verdict for a clause based on its money amounts.

    Returns None when the clause has no groundable amount (defer to the LLM/soft
    judgement). Otherwise returns a clause-review dict with a confidence, a
    defect `code`, and one `amounts` entry per figure carrying its span, whether
    it matched, and the fact that matched it.

    `offset` is the clause's start in the CONTENT the spans must index --
    verify_commentary passes the clause start it gets from segment_clauses, so
    every span stored here is in agent_4_validation["final_content"]
    coordinates, not clause-relative ones. Getting that wrong is silent: a
    clause-relative span still slices to plausible-looking text.
    """
    triples = extract_amount_spans(clause)
    if not triples:
        return None
    entries: List[Dict[str, Any]] = []
    unmatched: List[Dict[str, Any]] = []
    for value, start, end in triples:
        # The floor is what the figure was WRITTEN with -- see display_half_ulp.
        fact = source.matches(value, floor=display_half_ulp(clause[start:end]))
        entry: Dict[str, Any] = {
            "value": float(value),
            "span": [int(start + offset), int(end + offset)],
            "matched": fact is not None,
            "source_ref": fact,
        }
        if fact is None:
            entry.update(source.classify_miss(value))
            unmatched.append(entry)
        entries.append(entry)

    if unmatched:
        scale_errors = [e for e in unmatched if e.get("code") == "AMOUNT_SCALE_ERROR"]
        # A scale error is the more specific finding and the only patchable one,
        # so it names the verdict when any unmatched amount is one.
        lead = scale_errors[0] if scale_errors else unmatched[0]
        code = lead["code"]
        parts = []
        for entry in unmatched:
            if entry.get("code") == "AMOUNT_SCALE_ERROR":
                parts.append(
                    f"{entry['value']:,.0f} is off by a factor of {entry['factor']:g} — "
                    f"the source figure is {describe_fact(entry['source_ref'])}"
                    + (" [ambiguous: more than one scale factor fits]" if entry.get("ambiguous") else "")
                )
            elif entry.get("nearest"):
                parts.append(
                    f"{entry['value']:,.0f} not found in source data within tolerance "
                    f"(nearest source figure: {describe_fact(entry['nearest'])})"
                )
            else:
                parts.append(f"{entry['value']:,.0f} not found in source data within tolerance")
        review: Dict[str, Any] = {
            "supported": False,
            "category": "hallucination",
            "conf": 0.9,
            "code": code,
            "amounts": entries,
            "reason": "Amount(s): " + "; ".join(parts) + ".",
        }
        if lead.get("code") == "AMOUNT_SCALE_ERROR" and not lead.get("ambiguous"):
            review["expected"] = lead["expected"]
        return review

    # Provenance in the reason, replacing the bare "All amounts matched source
    # data within tolerance." -- that sentence made the verdict unfalsifiable
    # after the fact, which is exactly why measuring this pool's behaviour
    # needed a bit-exact reimplementation of SourceIndex rather than a read of
    # the archive.
    shown = "; ".join(f"{e['value']:,.0f} = {describe_fact(e['source_ref'])}" for e in entries[:3])
    if len(entries) > 3:
        shown += f"; +{len(entries) - 3} more"
    return {
        "supported": True,
        "category": "data-backed",
        "conf": 1.0,
        "amounts": entries,
        "reason": f"Amounts matched source data within tolerance: {shown}.",
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
                     llm: Optional[Dict[str, Any]], highlight_min_conf: float,
                     *, span: Optional[Tuple[int, int]] = None) -> Dict[str, Any]:
    """Merge deterministic number-grounding with the LLM's soft judgement.

    Precedence: a deterministic unmatched-amount hallucination is authoritative
    (the model cannot override hard arithmetic). When amounts all match, an LLM
    *reasoning* flag is preserved (numbers fine, inference unsupported) but an LLM
    *number-hallucination* claim is dropped (it was a false positive). Clauses with
    no checkable amount defer to the LLM; absent that, causal language => reasoning.
    """
    llm_cat = str((llm or {}).get("category") or "").lower()
    llm_supported = bool((llm or {}).get("supported")) if llm else True
    code: Optional[str] = None

    if det and det["category"] == "hallucination":
        category, supported, conf, reason = "hallucination", False, _CONF_DET_HALLUCINATION, det["reason"]
        code = det.get("code") or "AMOUNT_UNSUPPORTED"
    elif det and det["category"] == "data-backed":
        if llm and llm_cat == "reasoning" and not llm_supported:
            category, supported, conf = "reasoning", False, _CONF_LLM_FLAG
            reason = (llm or {}).get("reason") or "Numbers verified; inference not directly supported."
            code = "LLM_UNSUPPORTED_CAUSE"
        else:
            # numbers matched -> drop any LLM 'hallucination' false positive
            category, supported, conf, reason = "data-backed", True, _CONF_DET_DATA_BACKED, det["reason"]
    elif llm and llm_cat in ("reasoning", "hallucination") and not llm_supported:
        category, supported, conf = llm_cat, False, _CONF_LLM_FLAG
        reason = (llm or {}).get("reason") or "Flagged by validator."
        # The third branch: an LLM flag on a clause with no groundable amount.
        # Measured across 247 archived runs this is the DOMINANT defect class
        # (123 prose hallucinations against 99 arithmetic misses), so it gets a
        # code of its own rather than being routed to regeneration unlabelled.
        code = "LLM_UNSUPPORTED_FACT" if llm_cat == "hallucination" else "LLM_UNSUPPORTED_CAUSE"
    elif _has_causal_language(clause):
        category, supported, conf = "reasoning", False, _CONF_DEFAULT_REASONING
        reason = "Causal/inference clause with no figure to verify against source."
        # Same defect as an LLM-raised reasoning flag -- an asserted cause with
        # nothing behind it -- reached without the LLM. The code names the
        # defect, not the detector that found it.
        code = "LLM_UNSUPPORTED_CAUSE"
    else:
        category, supported, conf, reason = "data-backed", True, _CONF_DET_DATA_BACKED, "No checkable figure; no causal claim."

    # Confidence gate: low-confidence flags are demoted so they don't highlight
    # inline (keeps false positives low — the user's stated priority).
    if not supported and conf < highlight_min_conf:
        category, supported, code = "data-backed", True, None
    out: Dict[str, Any] = {"clause": clause, "supported": supported, "category": category, "reason": reason}
    # Optional keys only. Every existing consumer reads the four above and must
    # keep working on a review that carries none of these.
    if span is not None:
        out["span"] = [int(span[0]), int(span[1])]
    out["conf"] = float(conf)
    if code:
        out["code"] = code if code in DEFECT_CODES else "UNCATEGORIZED"
    if det and det.get("amounts"):
        out["amounts"] = det["amounts"]
    if det and det.get("expected") is not None and not supported:
        out["expected"] = det["expected"]
    return out


_ENUM_ITEM = re.compile(r"[1-9]）\s*[^；;]*?(-?[\d,]+(?:\.\d+)?)\s*(万元|亿元|元)")
# The enumeration runs from "主要包括" to the end of that SENTENCE. It used to
# require each component to be a run of 2-10 CJK characters immediately before
# its figure, which real labels in one deck do not satisfy:
#   其他应付款-非关联公司-其他169.0万元   -- hyphens break the CJK run
#   C0040某物流有限公司7.0万元          -- starts with a counterparty code
# Both were silently dropped from the sum, and an account whose own sentence
# adds up exactly was reported as "182.0万元 (60%) unaccounted for". Bounding on
# the sentence and taking every amount inside it needs no assumption about what
# a label looks like.
# 为/系 only count behind 主要. A bare one is the copula in "余额为301.8万元"
# and anchoring there swept the STATED TOTAL into the component list --
# the sum came out at exactly twice the total and was reported as a
# parent-plus-children duplication that was not there.
# Not across a line break either: 「…主要包括工程类、预提费用…，明细如下：\n
# 某年审计调整将X万元…；上年同口径调整X万元」 named no amounts in its list, and
# the two adjustment figures on the next line were summed as its items.
_ENUM_RUNON = re.compile(r"(?:主要(?:包括|包含|为|系)|包括|包含)([^。\n]*)")
_RUNON_AMT = re.compile(r"(-?[\d,]+(?:\.\d+)?)\s*(万元|亿元|元)")
_BRACKETED = re.compile(r"[（(][^（）()]*[）)]")
# A component introduced by 扣除 is taken OFF the ones before it. 「长期借款余额为
# 3.26亿元，主要包括工行借款3.63亿元，扣除一年内到期的非流动负债0.37亿元」 is
# 3.63 - 0.37 = 3.26 and ties exactly; added, it came to 4.00亿元 and a real run
# reported "-7,400.0万元 (23%) unaccounted for". Bounded to the clause the
# amount sits in, so a 扣除 earlier in the sentence does not reach a later item.
_DEDUCTED = re.compile(r"(?:扣除|减去|扣减|抵减|冲减)[^，,；;、]{0,24}$")


#: 「原值主要包括：…」 -- the short subject a list is introduced under.
_ENUM_SUBJECT = re.compile(
    r"(?:^|[。；;\n])\s*([^。；;，,\n：:\d]{2,10}?)(?:主要(?:包括|包含|为|由)|包括|包含|(?:合计|共计)?[\d,.]+\s*[万亿千]?元[，,]\s*其中)")


#: The two other ways a real bullet signs a component (Portfolio I run,
#: 2026-10-16): 「管理层调整负3,423.6万元」 and 「并由增值税留抵净额20.1万元抵减」 /
#: 「…53.3万元借方余额抵减」. Read as positive, four bullets that tied exactly
#: were reported as listing more than the total, and the sentence was
#: highlighted in the deck as unsupported.
#: Naming the side of a balance is only done when it is the contra side:
#: 「…某服务费借方余额10.0万元…」 inside an 应付账款 list nets OFF the payables
#: (60 + 10 - 10 = 60 against 60), and read as positive a real bullet was
#: reported as listing more than its total.
_NEGATIVE_PREFIX = re.compile(r"(?:负|借方余额|贷方余额)\s*$")
_DEDUCTED_AFTER = re.compile(r"^[^，,。；;、\d]{0,8}?(?:抵减|抵销|冲减|扣减)")


def _signed_runon_amounts(span: str) -> List[Tuple[str, str]]:
    out: List[Tuple[str, str]] = []
    for m in _RUNON_AMT.finditer(span):
        value, unit = m.group(1), m.group(2)
        if not value.startswith("-") and (
                _DEDUCTED.search(span[:m.start()]) or _NEGATIVE_PREFIX.search(span[:m.start()])
                or _DEDUCTED_AFTER.search(span[m.end():])):
            value = "-" + value
        out.append((value, unit))
    return out
# "其余X万元为…" is the closing component _composition_guidance explicitly asks
# for ("收尾必须写成'其余X万元为…'"). Not counting it made the checker
# contradict the instruction: a bullet that complied was flagged for the very
# amount it had just disclosed. 负/- because a contra component is normal --
# 管理层调整 is routinely negative.
#   其余11.1万元为...                      -- amount straight after 其余
#   其余小额应付款项（如物业管理费、法律服务费等）合计18.9万元
# The second shape is the normal one when the residual needs describing, and
# requiring the amount to follow 其余 immediately missed it: a bullet that
# added up exactly (54.7+16.3+10.0+18.9 = 99.9) was reported as 19%
# unaccounted for. Non-greedy and bounded inside the clause, so it takes the
# FIRST amount after 其余 and cannot reach into the next sentence.
_RESIDUAL_ITEM = re.compile(
    r"其余[^。；;]{0,40}?(负|-)?\s*([\d,]+(?:\.\d+)?)\s*(万元|亿元|元)"
)
# The same closing component written the other way the model actually writes it:
# 「其他小额应付款项（如法律服务费、维修费等）合计为负42.1万元」. Only 其余 was
# matched, so a real run summed three named items to 55.9万元 against a 9.5万元
# total and reported 488% unaccounted for -- the 42.1 that closes it was sitting
# in the sentence, spelled 其他 rather than 其余.
# And the third way: a catch-all named last in the list. 「主要包括中国电子系统工程
# 第四建设工程款731.3万元及其他应付款项124.7万元」 for an 877.2万元 balance -- the
# rest is 145.9, the sentence says 124.7, and with no 其余 or 合计 the check only
# called it 2% short, an omission, when it is a wrong number in the deck.
_RESIDUAL_TAIL = re.compile(
    r"(?:及|和|以及)其他[^，,。；;、\d]{0,12}?(负|-)?\s*([\d,]+(?:\.\d+)?)\s*(万元|亿元|元)"
)
_RESIDUAL_OTHER = re.compile(
    # Not 其他应付款 / 其他应收款 / 其他流动… : those are account names. A real
    # bullet's later sentence 「与其他应付款中账龄小于1年的租赁保证金合计1,109.2万元」
    # was read as this account's closing component, the sum came to exactly
    # double, and the account was regenerated twice for a sentence that was fine.
    r"其他(?!应付|应收|流动|非流动|业务|收益|权益)[^。；;]{0,40}?合计\s*(?:约)?\s*(?:为)?\s*(负|-)?\s*([\d,]+(?:\.\d+)?)\s*(万元|亿元|元)"
)
#: 「分别为」 introduces the SAME line across several PERIODS, not several
#: components. A real run summed 0 + 182.1 + 150.0 + 65.9 = 398.0万元 against a
#: 131.8万元 total and called it 202% unaccounted for; the four figures were one
#: account in 2023, 2024, 2025 and 2026H1. The span is cut here rather than the
#: whole check abandoned, so 「主要包括A、B，分别占比…」 keeps A and B.
_PERIOD_SERIES = re.compile(r"分别|respectively", re.I)
#: ...but only when the clause in front of it names periods. 「其中某银行A、B及C
#: 账户分别为X万元、Y万元及Z万元，其余W万元…」 is three accounts, not three dates;
#: cut there, the list lost its children and only the parent plus the 其余 was
#: summed, overshooting the balance by the 其余.
_PERIOD_WORDS = re.compile(r"年|期末|各期|月末|H1|FY", re.I)
#: 「主要包括甲46.4万元、乙30.1万元…；其中甲主要为某物流32.0万元，乙主要为审计费
#: 24.1万元」: a 其中 that opens its own clause after the list drills into the
#: items already named. Swept into the list it took a real 其他应付款 bullet that
#: tied exactly (82.7) to 138.9 and reported it as overshooting. Only after a
#: semicolon: 「…其他169.0万元，其中主要为某税务局166.8万元；押金83.1万元…」 keeps
#: listing peers after a comma-其中, and the drop-one test below handles that.
_DRILL_DOWN = re.compile(r"[；;]\s*其中")
#: 「预计主要系基本户余额由…0.58亿元降至…0.01亿元所致」 is a cause, and its two
#: figures are one line in two periods. Read as a list it put a 1.23亿元 cash
#: balance at "52% unaccounted for" on a real run.
_CAUSAL = re.compile(r"所致|导致|引起|造成")
#: 账面价值 / 账面净值 too: 「固定资产账面价值为6.65亿元，主要为一期房屋2.79亿元、
#: 二期房屋2.64亿元及土地-1.10亿元」 lists 4.33 against 6.65 -- the 1.10 is the
#: accumulated depreciation, not land -- and went unchecked because the total
#: was not called 余额.
_STATED_TOTAL = re.compile(r"(?:合计|总额|余额合?计?|账面(?:价值|净值))\s*(?:为)?\s*([\d,]+(?:\.\d+)?)\s*(万元|亿元|元)")
_SCALE = {"元": 1.0, "万元": 1e4, "亿元": 1e8}


def closing_remainder(text: str) -> Optional[float]:
    """The 「其余X」 of a sentence, when X is its first amount less the ones between.

    Read off the sentence alone -- first amount is what is being broken down,
    the rest up to 其余 are the items -- because the two shapes real models
    write are both outside what _composition_findings parses: a bare 其中 with
    no 主要包括 (「余额为1,062.8万元，…其中甲585.2万元、乙298.0万元，其余179.6万元
    为…」), and a nil balance set out as its gross movements (「余额合计0.00亿元，
    包括：结转固定资产为-7.75亿元；建造成本为5.94亿元；土地成本为1.17亿元；其余
    0.64亿元为…」). Returns the remainder in 元, or None when the sentence does
    not close to within the rounding of the figures as printed.
    """
    body = str(text or "")
    for match in _RESIDUAL_ITEM.finditer(body):
        start = body.rfind("。", 0, match.start()) + 1
        before = [(float(v.replace(",", "")), u, v) for v, u in _RUNON_AMT.findall(body[start:match.start()])]
        if not before and start > 0:
            # 「…余额为1,062.8万元，其中甲585.2万元、乙298.0万元。其余179.6万元为…」:
            # the remainder opens a sentence of its own, and the list it closes is
            # the previous one. A real draft written this way was regenerated
            # twice for "1,796,000 not found in source data".
            previous = body.rfind("。", 0, start - 1) + 1
            before = [(float(v.replace(",", "")), u, v)
                      for v, u in _RUNON_AMT.findall(body[previous:start])]
            start = previous
        if len(before) in (1, 2) and not _STATED_TOTAL.search(body[start:match.start()]):
            # The total stated one sentence earlier: 「余额为1,062.8万元，系…保证金。
            # 该科目主要包括：1）甲585.2万元；2）乙298.0万元；3）其余179.6万元为…」.
            stated_total = _STATED_TOTAL.search(body[:start])
            if stated_total:
                before.insert(0, (float(stated_total.group(1).replace(",", "")), stated_total.group(2),
                                  stated_total.group(1)))
        # One named item is enough: 「余额为877.2万元，主要为应付…工程款731.3万元，
        # 其余145.9万元为…」 is a total, its largest component and what is left.
        # It was flagged as an amount "not found in source data" on two runs in
        # a row while this required two items.
        if len(before) < 2:
            continue
        scale = _SCALE.get(match.group(3), 1.0)
        stated = float(match.group(2).replace(",", "")) * scale * (-1.0 if match.group(1) else 1.0)
        total = before[0][0] * _SCALE.get(before[0][1], 1.0)
        items = sum(v * _SCALE.get(u, 1.0) for v, u, _ in before[1:])
        # Each printed figure is rounded to its last decimal; half a step each.
        step = max(_SCALE.get(u, 1.0) * 10 ** -len(raw.partition(".")[2]) for _, u, raw in before)
        if abs((total - items) - stated) <= 0.5 * step * (len(before) + 1):
            return abs(stated)
    return None


def check_composition_adds_up(mapping_key: str, text: str) -> List[str]:
    """Message-only wrapper over _composition_findings.

    Kept because inspect_databook.py:2232 collects these as plain warning
    strings; the typed form below is what the verifier reads, so the category
    no longer has to be recovered by substring-matching the English message.
    """
    return [message for _code, message in _composition_findings(mapping_key, text)]


def patch_stated_remainder(text: str) -> Tuple[str, Optional[str]]:
    """Rewrite a wrong 「其余X」 to the figure it has to be.

    A stated remainder has exactly one right value -- the total less the items
    listed before it -- so this is arithmetic, not a judgement, and asking the
    model again does not get it: a real 应付账款 bullet wrote 「其余60.6万元」
    for 26.6万元, was sent back twice with the correct figure in the feedback,
    and returned 60.6万元 both times. Only the number changes; the words the
    model put after it are left as they are. Returns (text, note), note None
    when nothing was changed.
    """
    body = str(text or "")
    details: Dict[str, Any] = {}
    findings = _composition_findings("", body, details=details)
    if not findings or findings[0][0] != "COMPOSITION_RESIDUAL_WRONG" or not details:
        return body, None
    start, end = details["span"]
    written = body[start:end]
    decimals = len(written.split(".", 1)[1]) if "." in written else 0
    scale = _SCALE.get(details["unit"], 1.0)
    correct = details["correct"]
    if correct <= 0 or details.get("negative"):
        return body, None
    replacement = f"{correct / scale:,.{decimals}f}"
    patched = body[:start] + replacement + body[end:]
    # Never leave it worse: the patched text must itself add up.
    if _composition_findings("", patched):
        return body, None
    return patched, f"remainder {written}{details['unit']} -> {replacement}{details['unit']}"


def _composition_findings(mapping_key: str, text: str,
                          details: Optional[Dict[str, Any]] = None) -> List[Tuple[str, str]]:
    """Does an enumerated composition actually reach the total it states?

    The model is asked to add its items up before writing them and to account
    for any difference. It does not reliably do either: one real account came
    back as "余额合计674.5万元，主要包括：1）434.2万元；2）163.8万元；3）39.2万元"
    twice in a row, which is 637.2 -- the missing 37.3万元 being precisely the
    three items the analyst deliverable lists and this one does not.

    Arithmetic is checkable, so it is checked here rather than left to the
    reader to spot. A gap under 1% is treated as rounding.
    """
    body = str(text or "")
    # No "）" guard: the run-on form carries no numbering at all, and that early
    # exit is why a 10x unit error in a 、-separated list was never reached.
    if not body.strip():
        return []
    m = _STATED_TOTAL.search(body)
    # The stated total must not be counted as one of its own parts. In the
    # run-on form "…合计968.3万元，其中折旧成本为827.6万元，物管费用为73.9万元"
    # the total sits INSIDE the span the item pattern sweeps, so it was added
    # to the items: 968.3 + 903.2 = 1,871.5 against a 968.3 total, reported as
    # "93% unaccounted for" when the real shortfall was 65.1万元. The ratio
    # lands at 1.93, just outside the "exactly double" guard below. Blanked
    # rather than cut so every other offset in the clause is unchanged.
    scan = body if not m else body[:m.start()] + " " * (m.end() - m.start()) + body[m.end():]
    # The total has to be stated BEFORE the list it is held against. 「在建工程
    # 无余额，账面构成包括…5.94亿元、…1.17亿元、…-7.75亿元，其余0.64亿元为…。
    # 2023年末在建工程余额为2.94亿元，系…」: the list closes on a nil balance
    # that is given in words, and the only figure the total pattern finds is a
    # different period's, two sentences later. Held against that one the list
    # was reported as leaving 100% unaccounted for, three runs in a row.
    if m:
        _first_item = min((found.start() for found in (_ENUM_ITEM.search(scan), _ENUM_RUNON.search(scan)) if found),
                          default=None)
        if _first_item is not None and m.start() > _first_item:
            return []
    items = _ENUM_ITEM.findall(scan)
    if not items:
        run = next((r for r in _ENUM_RUNON.finditer(scan) if not _CAUSAL.search(r.group(1))), None)
        if run:
            span = run.group(1)
            _drill = _DRILL_DOWN.search(span)
            if _drill:
                span = span[:_drill.start()]
            _series = _PERIOD_SERIES.search(span)
            if _series:
                _clause = re.split(r"[，,；;]", span[:_series.start()])[-1]
                if _PERIOD_WORDS.search(_clause) or not re.search(r"[、及和与]", _clause):
                    span = span[:_series.start()]
            # A bracket that repeats a figure already listed outside it is that
            # same component said again, not a second one: 「主要包括关联方-美元
            # 项下的3,649.7万元，系…500万美元借款本金（按汇率折合人民币3,649.7
            # 万元）」 is one loan. Counted twice it came to 7,299.4 against a
            # 3,782.9万元 total -- ratio 1.93, just outside the "exactly double"
            # guard below -- and shipped as "93% unaccounted for", the run's only
            # high-severity clause flag. A bracketed figure that repeats nothing
            # is kept: it may be a component in its own right.
            outside = _signed_runon_amounts(_BRACKETED.sub(" ", span))
            inside = [a for b in _BRACKETED.findall(span) for a in _signed_runon_amounts(b)]
            items = outside + [a for a in inside if a not in outside]
    # The closing "其余X万元为…" is a component like any other. Added here
    # rather than inside the two patterns above because it can sit in either
    # form -- after a numbered list, or trailing a run-on one -- and because a
    # run-on span that already swallowed it must not count it twice.
    residual = _RESIDUAL_ITEM.search(scan) or _RESIDUAL_OTHER.search(scan) or _RESIDUAL_TAIL.search(scan)
    if residual:
        sign, value, unit = residual.groups()
        signed = float(value.replace(",", "")) * (-1.0 if sign else 1.0)
        if not any(abs(float(v.replace(",", "")) - abs(signed)) < 1e-9 and u == unit
                   for v, u in items):
            items = list(items) + [(f"{signed:.10g}", unit)]
    if not m or len(items) < 2:
        return []
    total = float(m.group(1).replace(",", "")) * _SCALE.get(m.group(2), 1.0)
    listed = sum(float(v.replace(",", "")) * _SCALE.get(u, 1.0) for v, u in items)
    if total <= 0:
        return []
    gap = total - listed
    if abs(gap) / total <= 0.01:
        return []
    # A list introduced under its own subject composes THAT figure, not the
    # account: 「余额为355.8万元，系原值653.3万元减累计摊销297.5万元后的净值。
    # 原值主要包括：1）项目管理费569.5万元；2）绿化改造工程83.9万元」 breaks down
    # the 653.3 gross cost, correctly. Held against the 355.8 net balance it
    # was flagged as overshooting the total -- a hallucination, so the account
    # was regenerated twice and came back the same, because it was right. The
    # subject has to be named in front of the verb AND carry a stated amount
    # the items tie to; a list that merely happens to equal some other figure
    # in the text (a prior period's balance, say) is still reported.
    for _subject in _ENUM_SUBJECT.finditer(scan):
        _label = _subject.group(1).strip()
        if not _label or _label in str(mapping_key) or re.search(r"余额|科目|合计|总额|本期|期末|截至", _label):
            continue
        for _at in re.finditer(re.escape(_label), body):
            _stated = _RUNON_AMT.search(body[_at.end():_at.end() + 24])
            if not _stated or (_at.start() == _subject.start(1) and "其中" not in _subject.group(0)):
                continue
            _value = abs(float(_stated.group(1).replace(",", ""))) * _SCALE.get(_stated.group(2), 1.0)
            # 1%, the same rounding allowance as the account total: 「工行借款3.63
            # 亿元…主要包括一期1.40亿元及二期2.20亿元」 is 3.60 printed to two
            # decimals of 亿, 0.8% short, and at 0.5% a real 长期借款 bullet was
            # held against the 3.26 net balance instead and called 10% over.
            if _value and abs(_value - listed) / _value <= 0.01:
                return []
    # A "其中" drill-down is part of the item before it, not a peer of it:
    # "…非关联公司-其他169.0万元，其中主要为某税务局166.8万元；…押金83.1万元;
    # 管理层调整36.7万元；其余13.0万元…" lists four components that hit the
    # 301.8万元 total exactly, and one figure that is INSIDE the first of them.
    # Counted as five they came to 468.6 and the account was reported as not
    # adding up. Dropping any one item and landing on an exact tie is that
    # shape; testing it arithmetically avoids having to decide what 其中 means
    # in a sentence, which it does not always mean.
    values = [float(v.replace(",", "")) * _SCALE.get(u, 1.0) for v, u in items]
    # With exactly one parent and one child, both can legitimately equal the
    # account total: 「主要为非工程款16.6万元，其中某保险公司16.6万元」.
    # The child is a drill-down, not a second component. The generic
    # drop-one-and-retest guard below deliberately requires >2 values, so this
    # smallest valid hierarchy needs its own unambiguous case.
    if (len(values) == 2 and "其中" in scan
            and all(abs(value - total) / total <= 0.01 for value in values)):
        return []
    # A line that restates the total is a parent, not a component. 「货币资金余额
    # 为1,128.3万元，主要为银行存款-人民币户余额1,128.3万元，其中…」 lists the
    # whole account again before breaking it down, and counting it put the items
    # at 133% of the total. Dropped and re-tested rather than reported.
    _restate = [v for v in values if abs(v - total) / total <= 0.01]
    if _restate and len(values) - len(_restate) >= 2:
        values = [v for v in values if abs(v - total) / total > 0.01]
        items = [(v, u) for v, u in items
                 if abs(float(v.replace(",", "")) * _SCALE.get(u, 1.0) - total) / total > 0.01]
        listed = sum(values)
        gap = total - listed
        if abs(gap) / total <= 0.01:
            return []
        # 「主要为营运收费X万元，其中金额较大的应收对象为A…、B…及C…」 says
        # outright that it lists only the larger ones. Short of the parent is
        # what that sentence promises, not an omission.
        if gap > 0 and re.search(r"其中[^，,。；;\d]{0,8}(?:主要|较大|前\S{0,3}大)", scan):
            return []
    if len(values) > 2:
        biggest = max(values)
        for value in values:
            if value < biggest and abs(total - (listed - value)) / total <= 0.01:
                return []
    fmt = lambda v: f"{v/1e4:,.1f}万元"
    # A ratio near a power of ten is a unit error, not an omission: the model
    # took a raw CNY'000 cell and wrote 万元 against it. Worth saying so --
    # "66% unaccounted for" reads as a missing component, and the reader would
    # go looking for one that does not exist.
    ratio = listed / total if total else 0
    for _mult, _label in ((10.0, "10x"), (100.0, "100x"), (0.1, "1/10")):
        if abs(ratio - _mult) / _mult <= 0.05:
            return [(
                "COMPOSITION_UNIT_ERROR",
                f"[{mapping_key}] composition is {_label} the stated total "
                f"({fmt(listed)} vs {fmt(total)}) -- this is a UNIT error, not a "
                f"missing component: a raw CNY'000 figure written as 万元."
            )]
    if abs(ratio - 2.0) <= 0.05:
        return [(
            "COMPOSITION_DOUBLE_COUNT",
            f"[{mapping_key}] composition is exactly double the stated total "
            f"({fmt(listed)} vs {fmt(total)}) -- a parent line and the lines "
            f"that make it up have both been listed."
        )]
    if residual and abs(gap) / total <= 0.2:
        # The bullet closed with its own remainder and the remainder is wrong:
        # 434.2 + 163.8 + 39.2 leaves 37.3 of 674.5, and the sentence said
        # 「其余27.3万元」. Not an omission -- there is nothing left to omit once
        # "the rest" has been stated -- so it is arithmetic, and rewritable.
        stated = abs(float(residual.group(2).replace(",", ""))) * _SCALE.get(residual.group(3), 1.0)
        if details is not None:
            details.update(span=(residual.start(2), residual.end(2)), unit=residual.group(3),
                           correct=stated + gap, negative=bool(residual.group(1)))
        return [(
            "COMPOSITION_RESIDUAL_WRONG",
            f"[{mapping_key}] the stated remainder of {fmt(stated)} does not close the list: the items "
            f"before it leave {fmt(stated + gap)} of the {fmt(total)} total."
        )]
    if ratio > 1.2:
        # More listed than the account holds is not an omission -- nothing can
        # be "left out" of a list that already overshoots. A real bullet gave a
        # 1,062.8万元 balance as "1）押金及保证金1,034.3万元；2）某租户585.2万元":
        # the prior year's balance and one tenant inside it, as peers. Reported
        # as a shortfall it only coloured the sentence; as an overshoot it is
        # certain enough to rewrite.
        return [(
            "COMPOSITION_EXCEEDS",
            f"[{mapping_key}] the listed items come to {fmt(listed)}, more than the stated total of "
            f"{fmt(total)}: an item from another period, or a part listed beside its whole."
        )]
    return [(
        "COMPOSITION_GAP",
        f"[{mapping_key}] composition does not reach the stated total: "
        f"{len(items)} item(s) sum to {fmt(listed)} against {fmt(total)}, "
        f"leaving {fmt(gap)} ({abs(gap)/total:.0%}) unaccounted for. The reader "
        f"cannot tell whether the rest is an omission or a component with no name."
    )]


def verify_commentary(final_content: str, df, llm_clause_reviews: Optional[List[Dict[str, Any]]] = None,
                      *, highlight_min_conf: float = 0.6,
                      sibling_dfs: Optional[List[Any]] = None,
                      source: Optional[SourceIndex] = None,
                      direction_findings: Optional[List[Dict[str, Any]]] = None) -> List[Dict[str, Any]]:
    """Authoritative clause_reviews: deterministic number-grounding layered over the
    LLM's soft reasoning judgement. Each clause is a verbatim substring of
    final_content, so highlighting matches by exact offset. Returns the existing
    clause_reviews shape [{clause, supported, category, reason}].

    sibling_dfs (optional): other accounts' DataFrames — same statement type as
    this account, per caller — so a legitimate cross-tab reference (e.g. an
    "Other payables" note citing the bank loan balance that actually lives on
    the "Long-term loans" tab) can be grounded instead of false-flagged.

    source (optional): a prebuilt SourceIndex, so a caller holding one per
    account does not rebuild it per call. Absent, it is built here, which is
    what the CLI and every diagnostic rely on.

    direction_findings (optional): an out-parameter. When a list is passed, M3's
    direction check appends its findings to it. They deliberately do NOT go into
    the returned clause_reviews — a review there with category "hallucination"
    buys a paid retry on the first run, and one with "reasoning" still colours
    the sentence in the exported client deck and counts toward the
    unsupported-ratio gate. Report-only until the false-positive rate has been
    read on a real book."""
    if source is None:
        source = SourceIndex.from_df(df, sibling_dfs=sibling_dfs)
    llm_reviews = llm_clause_reviews or []
    out: List[Dict[str, Any]] = []
    # 「…其中甲585.2万元、乙298.0万元，其余179.6万元为零星租户保证金」: the 179.6
    # is in no cell, because it is the total less the two items named -- the
    # closing remainder the prompt asks for. One real model wrote exactly
    # that, was flagged for an amount "not found in source data", regenerated
    # twice, and three others that left the remainder out were merely noted
    # as incomplete: the check punished the one that followed the
    # instruction. A remainder that closes its own list to within rounding is
    # arithmetic on grounded figures, so it joins the pool for this text. One
    # that does NOT close is untouched, and is COMPOSITION_RESIDUAL_WRONG.
    try:
        _remainder = closing_remainder(str(final_content or ""))
        if _remainder is not None:
            _extended = SourceIndex(list(source.facts) + [
                _fact(_remainder, "composition_remainder",
                      row_desc="stated total less the items listed with it")])
            _extended.date_facts = list(getattr(source, "date_facts", []) or [])
            source = _extended
    except Exception:
        pass
    # The (start, end) segment_clauses returns used to be discarded here, and
    # every consumer downstream re-found the clause by string search. They are
    # the coordinate system every span in this module is expressed in, so they
    # are carried through: the clause's own span, and the offset that rebases
    # ground_amounts' clause-relative amount spans onto final_content.
    for start, end, clause in segment_clauses(final_content):
        det = ground_amounts(clause, source, offset=start)
        llm = _lookup_llm_review(clause, llm_reviews)
        out.append(_combine_verdict(clause, det, llm, highlight_min_conf, span=(start, end)))
    out.extend(_composition_reviews(final_content))
    out.extend(_date_reviews(final_content, df, source=source))
    try:
        out.extend(_percent_claim_reviews(final_content))
    except Exception:
        pass  # a new detector must never cost an account its other verdicts
    try:
        out.extend(_movement_claim_reviews(final_content, df))
    except Exception:
        pass
    try:
        out.extend(_balance_claim_reviews(final_content, df))
    except Exception:
        pass
    if direction_findings is not None:
        direction_findings.extend(collect_direction_findings(final_content, df))
    return out


# Thousands separators too: 「增长约1,2XX.X%」 was read as 2XX.X% and reported as
# a hallucination against the four-digit rise its own two amounts make.
_PCT_CLAIM = re.compile(r"(?<![\d,.])(\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?)\s*%")
_PCT_UP = re.compile(r"增长|增加|上升|增幅|回升|升幅|increase|rose|up\b", re.IGNORECASE)
_PCT_DOWN = re.compile(r"下降|减少|降幅|降低|下滑|decrease|fell|declin|down\b", re.IGNORECASE)
_PCT_VERSUS = re.compile(r"较|相比|对比|比|与|against|compared|versus|\bon\b", re.IGNORECASE)
_PCT_FROM = re.compile(r"(?:由|从|自|from)\s*[^，,。；;]{0,12}$", re.IGNORECASE)
_PCT_TO = re.compile(r"至|到|\bto\b", re.IGNORECASE)
_NEGATIVE_LEAD = re.compile(r"(?:负|-|−)\s*$")
_PCT_SENTENCE = re.compile(r"(?:[^。；;！？!?\n.]|\.(?!\s+[A-Z]))+")


def percent_claim_findings(text: str) -> List[Dict[str, Any]]:
    """Sentences whose own two amounts do not make the percentage they state.

    「年化后为3,675.8万元，较2025年度的2,787.2万元下降67.0%」 needs no source data
    to be wrong: 3,675.8 is the larger figure. Every amount in it grounds, so
    amount-grounding passes it, and it shipped on the most important line of a
    real income statement, with a cause attached to the fall it invented.

    Deliberately narrow, because a hit costs a regeneration. A sentence is
    judged only when it has ONE readable comparison:
      * a percentage with a direction word directly in front of it;
      * exactly the two nearest amounts before that word, at most 40 characters
        apart, joined either by 较/比/与 (current, then base) or by 由…至
        (base, then current);
      * no 分别 / respectively, which makes it a series, not a pair.
    Anything else is left alone. The percentage may be off by what the amounts'
    own display rounding allows plus one point; beyond that, or pointing the
    wrong way between two positive figures, it is reported with the arithmetic.
    """
    findings: List[Dict[str, Any]] = []
    body = str(text or "")
    # Whole sentences, not segment_clauses' comma-level clauses: the two
    # amounts of 「…为A，较B下降X%」 sit either side of a comma.
    for sentence in _PCT_SENTENCE.finditer(body):
        clause, start, end = sentence.group(0), sentence.start(), sentence.end()
        if _PERIOD_SERIES.search(clause):
            continue
        amounts = extract_amount_spans(clause)
        if len(amounts) < 2:
            continue
        for pct_match in _PCT_CLAIM.finditer(clause):
            lead = clause[max(0, pct_match.start() - 8):pct_match.start()]
            ups, downs = list(_PCT_UP.finditer(lead)), list(_PCT_DOWN.finditer(lead))
            if bool(ups) == bool(downs):
                continue
            word_at = pct_match.start() - len(lead) + (ups or downs)[-1].start()
            before = [a for a in amounts if a[2] <= word_at]
            if len(before) < 2:
                continue
            first, second = before[-2], before[-1]
            between = clause[first[2]:second[1]]
            tail = clause[second[2]:word_at]
            if len(between) > 40 or len(tail) > 8 or re.search(r"\d", tail):
                continue
            # 「余额为670.6万元，本期较上年末下降526.9万元，降幅78.6%」: the second
            # amount is the CHANGE, announced by its own movement verb, not a
            # second balance. Read as two balances it became "670.6 against
            # 526.9 is an increase of 27.3%" -- on a sentence whose arithmetic
            # is exactly right (526.9 / 670.6 = 78.6%) -- and bought two retries.
            if (_PCT_UP.search(clause[max(0, second[1] - 6):second[1]])
                    or _PCT_DOWN.search(clause[max(0, second[1] - 6):second[1]])):
                continue
            if _PCT_VERSUS.search(between):
                current, base = first, second
            elif _PCT_FROM.search(clause[:first[1]]) and _PCT_TO.search(between):
                base, current = first, second
            else:
                continue

            def signed(amount):
                negative = bool(_NEGATIVE_LEAD.search(clause[max(0, amount[1] - 3):amount[1]]))
                return -amount[0] if negative else amount[0]

            a, b = signed(current), signed(base)
            if not b or (a < 0) != (b < 0):
                continue
            stated = float(pct_match.group(1).replace(",", ""))
            ha = display_half_ulp(clause[current[1]:current[2]])
            hb = display_half_ulp(clause[base[1]:base[2]])
            mags = [abs((abs(a) + da) - (abs(b) + db)) / (abs(b) + db) * 100
                    for da in (-ha, ha) for db in (-hb, hb) if abs(b) + db > 0]
            if not mags:
                continue
            actual = (abs(a) - abs(b)) / abs(b) * 100
            off = stated < min(mags) - 1.0 or stated > max(mags) + 1.0
            wrong_way = a > 0 and b > 0 and abs(actual) > 1.0 and (actual > 0) != bool(ups)
            if not off and not wrong_way:
                continue
            findings.append({
                "clause": clause,
                "span": [int(start), int(end)],
                "code": "PCT_ARITHMETIC",
                "stated_pct": stated,
                "actual_pct": round(actual, 1),
                "reason": (
                    f"{clause[current[1]:current[2]]} against {clause[base[1]:base[2]]} is "
                    f"{'an increase' if actual > 0 else 'a decrease'} of {abs(actual):.1f}%, not "
                    f"{'an increase' if ups else 'a decrease'} of {stated:g}% as written. If one of "
                    "the two is a part-year amount and the other a full year, they are not "
                    "comparable: use the annualised figure and its own percentage."
                ),
            })
            break
    return findings


_MOVE_CLAIM = re.compile(
    r"较[^，,。；;]{0,14}?(?:新增|增加|增长|减少|下降|上升)\s*(?:约|了)?\s*$"
    r"|(?:increased?|decreased?|rose|fell|up|down)\s+(?:by\s+)?$", re.IGNORECASE)


def movement_claim_findings(text: str, df) -> List[Dict[str, Any]]:
    """An amount written as a CHANGE that is not the change of any line.

    「较上年末新增3.63亿元」 shipped three runs in a row for an account whose
    balance was 0.37亿元 in both periods. 3.63亿元 is on the sheet -- it is the
    whole loan -- so amount-grounding accepts it; what no check asked is whether
    it is a movement. A figure introduced by 较…增加/减少 has to be the
    difference between two periods of ONE line of this account's analysis
    frame: any two periods, not only neighbours, and the part-year tail on
    either basis. Nothing else is tested, and an account whose frame has fewer
    than two periods is left alone.
    """
    attrs = getattr(df, "attrs", None) or {}
    analysis = attrs.get("prompt_analysis_df")
    if analysis is None or not hasattr(analysis, "columns"):
        return []
    skip = set(SourceIndex._non_amount_cols(analysis))
    columns = [c for c in list(analysis.columns)[1:] if c not in skip]
    if len(columns) < 2:
        return []
    months = None
    try:
        from ..workbook.analysis import partial_tail_months
        months = partial_tail_months(analysis, (attrs.get("integrity") or {}).get("statement_type"))
    except Exception:
        months = None
    differences: List[float] = []
    table = [analysis[c].tolist() for c in columns]
    for pos in range(len(analysis.index)):
        row = []
        for series in table:
            cell = series[pos]
            value = _to_float(cell) if isinstance(cell, (int, float, str)) else None
            if value is not None and value == value:
                row.append(float(value))
        if months and row:
            row.append(row[-1] * 12.0 / months)
        for i in range(len(row)):
            for j in range(i + 1, len(row)):
                if abs(row[i] - row[j]) > 1e-9:
                    differences.append(abs(row[i] - row[j]))
    if not differences:
        return []

    findings: List[Dict[str, Any]] = []
    body = str(text or "")
    for sentence in _PCT_SENTENCE.finditer(body):
        clause = sentence.group(0)
        for value, start, end in extract_amount_spans(clause):
            if not _MOVE_CLAIM.search(clause[max(0, start - 22):start]):
                continue
            tolerance = max(1000.0, 2 * display_half_ulp(clause[start:end]), 0.005 * abs(value))
            if any(abs(abs(value) - d) <= tolerance for d in differences):
                continue
            findings.append({
                "clause": clause,
                "span": [int(sentence.start()), int(sentence.end())],
                "code": "MOVEMENT_UNSUPPORTED",
                "reason": (
                    f"{clause[start:end]} is written as a change, but it is not the difference between "
                    "any two periods of any line of this account. A balance or a component is not a "
                    "movement: state the change from the two period figures, or drop it."
                ),
            })
            break
    return findings


_STATED_BALANCE = re.compile(
    r"截至\s*(\d{4})\s*[年\-/.]\s*(\d{1,2})\s*[月\-/.]\s*(\d{1,2})\s*日?[，,]?"
    r"[^。；;，,\d]{0,24}?(?:余额|账面净值|账面价值|净值)(?:合计)?\s*(?:为|是|共计)?\s*"
    r"(-?[\d,]+(?:\.\d+)?)\s*(万元|亿元|元)"
)


def balance_claim_findings(text: str, df) -> List[Dict[str, Any]]:
    """「截至X，…余额为A」 where A is not what the account stood at on X.

    Every other check takes a figure and asks whether it exists somewhere in
    the account's numbers. This one asks whether the balance the sentence
    OPENS with is the account's balance. One real draft opened 「其他流动负债
    余额为74.4万元」 for an account standing at 37.2万元: the model had added
    the accrued-interest line to the two loans it is made of, and 74.4 passed
    grounding untouched because nothing checks a sum nobody stated as one.

    Quiet when the stated figure equals ANY line of the frame at that date --
    then the sentence is giving a named line's balance (a gross cost, a
    sub-account), which is not this check's business -- and when the frame has
    no total row or no column for the date.
    """
    findings: List[Dict[str, Any]] = []
    if df is None or not str(text or "").strip():
        return findings
    try:
        import pandas as pd
        from .facts import _total_row_index, _value_columns
        analysis = (getattr(df, "attrs", None) or {}).get("prompt_analysis_df")
        analysis = analysis if isinstance(analysis, pd.DataFrame) and not analysis.empty else df
        periods = _value_columns(analysis)
        total_idx = _total_row_index(
            analysis, (analysis.attrs or {}).get("row_types_by_description") or {}, periods)
    except Exception:
        return findings
    if total_idx is None:
        return findings
    for match in _STATED_BALANCE.finditer(str(text)):
        period = "%04d-%02d-%02d" % (int(match.group(1)), int(match.group(2)), int(match.group(3)))
        if period not in analysis.columns:
            continue
        column = pd.to_numeric(analysis[period], errors="coerce")
        total = column.get(total_idx)
        if total is None or pd.isna(total):
            continue
        raw = match.group(4)
        scale = _SCALE.get(match.group(5), 1.0)
        stated = abs(float(raw.replace(",", ""))) * scale
        # Half a step of the figure as printed, plus 0.5% for a total that was
        # itself rounded on the way into the frame.
        slack = 0.5 * scale * 10 ** -len(raw.partition(".")[2]) + 0.005 * abs(float(total))
        if abs(stated - abs(float(total))) <= slack:
            continue
        if any(abs(stated - abs(float(v))) <= slack for v in column.dropna()):
            continue
        # A power of ten apart is a question about units -- the frame's or the
        # sentence's -- and the scale checks own that; calling it a wrong
        # balance here would flag every account of a frame held in thousands.
        ratio = stated / abs(float(total)) if float(total) else 0.0
        if not ratio or any(abs(ratio - k) / k <= 0.02 for k in (10.0, 100.0, 1e3, 1e4, 0.1, 0.01, 1e-3, 1e-4)):
            continue
        start = str(text).rfind("。", 0, match.start()) + 1
        end = str(text).find("。", match.end())
        end = end + 1 if end >= 0 else len(str(text))
        findings.append({
            "clause": str(text)[start:end].strip(), "span": [int(start), int(end)],
            "code": "BALANCE_MISMATCH",
            "reason": (
                f"the sentence gives a balance of {raw}{match.group(5)} at {period}; the account's total "
                f"at that date is {abs(float(total)) / scale:,.{len(raw.partition('.')[2])}f}{match.group(5)}, "
                "and no line of the schedule carries the stated figure. A part may have been added to "
                "the whole it belongs to."
            ),
        })
    return findings


def _balance_claim_reviews(final_content: str, df) -> List[Dict[str, Any]]:
    """balance_claim_findings as clause reviews. "reasoning" for now -- shown,
    not retried -- because two of the last three checks to be made retriable on
    the day they were written turned out wrong more often than right on a
    model whose sentences they had not seen. Promote it once a real run has
    shown what it fires on."""
    return [{
        "clause": f["clause"], "supported": False, "category": "reasoning",
        "reason": f["reason"], "code": f["code"], "span": f["span"],
        "conf": _CONF_DET_HALLUCINATION,
    } for f in balance_claim_findings(final_content, df)]


def _movement_claim_reviews(final_content: str, df) -> List[Dict[str, Any]]:
    return [{
        "clause": f["clause"], "supported": False, "category": "hallucination",
        "reason": f["reason"], "code": f["code"], "span": f["span"],
        "conf": _CONF_DET_HALLUCINATION,
    } for f in movement_claim_findings(final_content, df)]


def _percent_claim_reviews(final_content: str) -> List[Dict[str, Any]]:
    """percent_claim_findings as clause reviews. "hallucination", so the retry
    gate sees it: the sentence states a movement its own figures deny, and only
    a rewrite fixes that."""
    return [{
        "clause": f["clause"], "supported": False, "category": "hallucination",
        "reason": f["reason"], "code": f["code"], "span": f["span"],
        "conf": _CONF_DET_HALLUCINATION,
    } for f in percent_claim_findings(final_content)]


def _composition_reviews(final_content: str) -> List[Dict[str, Any]]:
    """Whether an enumerated composition reaches the total it states.

    Every amount in "余额合计674.5万元，主要包括：1）434.2万元；2）163.8万元；
    3）39.2万元" can be individually present in the source and the sentence
    still be wrong -- the three add to 637.2, not 674.5. ground_amounts checks
    figures one at a time and cannot see that, which is why these reached real
    decks repeatedly while the Validator passed the account with zero flags.

    check_composition_adds_up already existed, but only inside
    inspect_databook, where it prints a warning AFTER the deck is built and
    feeds nothing. Reading it here turns it into a real clause_review: it
    highlights in the deck and the retry gate can see it.

    Category is chosen by how certain the defect is, because only
    "hallucination" costs a retry (see count_defective_clauses):
      * a ratio at a power of ten, or exactly double -- arithmetic that is
        definitely wrong, a unit error or a parent listed with its own
        children. Worth re-generating for.
      * anything else is a shortfall: each amount may be right and the
        composition merely incomplete. Flagged for the reader, but not worth
        spending a retry on -- the same reasoning that keeps the gate off
        ordinary "reasoning" flags.
    """
    body = str(final_content or "")
    reviews: List[Dict[str, Any]] = []
    # The category used to be recovered by substring-matching the English
    # message ("UNIT error" in detail) -- a rule that would have gone silently
    # wrong the day the wording changed. _composition_findings hands back the
    # code the check already knew.
    for code, message in _composition_findings("", body):
        detail = message.split("] ", 1)[-1]
        # COMPOSITION_EXCEEDS was here for one day. On the first run that put
        # five models through all 31 accounts it fired four times and was wrong
        # four times, each costing two regenerations that came back the same:
        #   * a drill-down inside a drill-down (「计提867.4万元及第三方9.9万元；
        #     计提中主要为…731.3万元，其中…725.0万元」);
        #   * a period comparison beside the total (「合计366.1万元，年化后
        #     1,464.3万元，较2025年度1,463.3万元基本持平」);
        #   * an earlier period's breakdown after the latest balance (「2025年末
        #     余额为16.6万元…2023年末余额23.6万元则主要为…16.1万元及…7.3万元」);
        #   * gross items with the deduction stated after them (「主要包括…5.49
        #     亿元、…，受结转固定资产4.19亿元影响后净额为2.94亿元」) -- which was
        #     the one text of the five that had the right period's figures.
        # It was built on one model's habits and does not survive another's
        # sentence shapes. It still reports; it no longer buys a retry.
        certain = code in ("COMPOSITION_UNIT_ERROR", "COMPOSITION_DOUBLE_COUNT")
        # COMPOSITION_RESIDUAL_WRONG is deliberately NOT here: it is corrected
        # in code when the final text is settled (patch_stated_remainder), and
        # regenerating for it bought the same wrong figure twice.
        # Anchor on the sentence stating the total, so the deck highlights the
        # claim rather than the whole paragraph.
        match = _STATED_TOTAL.search(body)
        clause = body
        span = [0, len(body)]
        if match:
            start = body.rfind("。", 0, match.start()) + 1
            end = body.find("。", match.end())
            end = (end + 1) if end >= 0 else len(body)
            sentence = body[start:end]
            lead = len(sentence) - len(sentence.lstrip())
            stripped = sentence.strip()
            if stripped:
                clause = stripped
                span = [start + lead, start + lead + len(stripped)]
        reviews.append({
            "clause": clause,
            "supported": False,
            "category": "hallucination" if certain else "reasoning",
            "reason": detail,
            "code": code,
            "span": span,
            # Every COMPOSITION_* code routes to regeneration, never to a patch:
            # a sentence whose parts do not add up cannot be fixed by swapping
            # one token, because which token is wrong is exactly what is unknown.
            "conf": _CONF_DET_HALLUCINATION if certain else _CONF_LLM_FLAG,
        })
    return reviews



#: A date is only a date here when all three parts are present. "2026年1-6月"
#: must NOT read as 2026-01-06, and "2024年度" is a period name, not a date.
_DATE_CHI = re.compile(r"(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日")
_DATE_ISO = re.compile(r"(\d{4})-(\d{1,2})-(\d{1,2})")


def _dates_in(value) -> set:
    """Every (y, m, d) reachable from one value, in either notation. A real
    datetime/Timestamp is taken directly so a column whose header is a date
    OBJECT counts as a source date the same as a string one."""
    if value is None:
        return set()
    year, month, day = (getattr(value, "year", None), getattr(value, "month", None),
                        getattr(value, "day", None))
    if isinstance(year, int) and isinstance(month, int) and isinstance(day, int):
        return {(year, month, day)}
    text = str(value)
    found = set()
    for pattern in (_DATE_CHI, _DATE_ISO):
        for y, m, d in pattern.findall(text):
            found.add((int(y), int(m), int(d)))
    return found


def _harvest_source_date_facts(df) -> List[Dict[str, Any]]:
    """Dates the account's own data actually contains, as facts with a source.

    Notes and remarks are included on purpose: a loan maturity or a lease end
    date is a legitimate date to quote and is not a period column. Grounding
    against the whole source, not just the period set, is the same contract
    SourceIndex already uses for amounts.

    Each date carries WHERE it came from (a column label, an attrs key) because
    dates are the largest deterministic defect class -- replaying archived
    account texts gave 47 DATE_UNSUPPORTED against 34 amount misses -- and a
    date patch with nothing to cite cannot be checked by the reader."""
    facts: List[Dict[str, Any]] = []
    seen: set = set()

    def add(value, col_label: str) -> None:
        for y, m, d in _dates_in(value):
            if (y, m, d) in seen:
                return
            seen.add((y, m, d))
            facts.append(_fact_date(y, m, d, sheet=_frame_sheet(df), col_label=col_label))

    if df is None:
        return facts
    attrs = getattr(df, "attrs", None) or {}
    integrity = attrs.get("integrity") or {}
    for key in ("effective_date", "raw_effective_date"):
        add(integrity.get(key), f"integrity.{key}")
    try:
        for col in df.columns:
            add(col, "column header")
            for cell in df[col].tolist():
                add(cell, str(col))
    except Exception:
        pass
    table = attrs.get("presentation_detail_table") or {}
    for period in (table.get("periods") or []):
        add(period, "presentation_detail_table.periods")
    for row in (table.get("rows") or []):
        add(row.get("label") if isinstance(row, dict) else row, "presentation_detail_table.rows")
    for bucket in ("supporting_notes", "adjacent_detail_rows"):
        for item in (attrs.get(bucket) or []):
            add(item, bucket)
    # The multi-period analysis frame, which is what the Generator was actually
    # shown. The amount pool has taken `analysis_cell` from here since it was
    # measured (see _OWN_HARD_KINDS) for exactly this reason -- "the true source
    # was a historical period that lives only in the analysis frame" -- and the
    # date pool was never given the same frame.
    #
    # What that cost, on a real run with repairs enabled: an account whose main
    # frame carries ONE period column ('...2026-06-30 annualised') cited
    # 2024-12-31 and 2025-12-31, which are real period ends living only here.
    # Both were flagged DATE_UNSUPPORTED and the repair rewrote them. The same
    # happened to 2024-01-01, a real column on the analysis frame, which came
    # out as 2024-12-31 -- a correct date turned into a wrong one, verified and
    # shipped. Two of those repairs produced text that cannot be true: a balance
    # "formed on 2024-12-31, with none as at 2024-12-31", and a loan maturing on
    # the balance sheet date instead of its real term end.
    #
    # Widening this does not weaken the check it exists for: the invented dates
    # in the docstring above (2232年, 1770年, 1938年) are in no frame at all.
    analysis_df = attrs.get("prompt_analysis_df")
    if analysis_df is not None and hasattr(analysis_df, "columns"):
        try:
            for col in analysis_df.columns:
                add(col, "analysis frame column header")
                for cell in analysis_df[col].tolist():
                    add(cell, "analysis frame %s" % col)
        except Exception:
            pass
    return facts


def _fact_date(year: int, month: int, day: int, *, sheet=None, col_label=None) -> Dict[str, Any]:
    rec = _fact(0.0, "date", sheet=sheet, col_label=col_label)
    rec["value"] = f"{int(year):04d}-{int(month):02d}-{int(day):02d}"
    rec["date"] = [int(year), int(month), int(day)]
    return rec


def _harvest_source_dates(df) -> set:
    """The (y, m, d) set behind _harvest_source_date_facts. Same contract as
    before; the facts carry the provenance the `expected` candidate needs."""
    return {tuple(fact["date"]) for fact in _harvest_source_date_facts(df)}


_RANGE_END = re.compile(r"\s*(?:至|到|起至|-|—|–|~|～)\s*(?:(\d{4})\s*年)?\s*(?:(\d{1,2})\s*月)?\s*(\d{1,2})\s*日")


def _opens_range_to_source_date(body: str, match, key: tuple, allowed: Dict[tuple, Any]) -> bool:
    """「2026年3月21日至3月31日」: a span INSIDE the reporting month, closing on
    a period end the source does carry.

    The remark on one real tab says interest is settled on the 20th of each
    month, so the balance at 31 March is the accrual for the 21st to the 31st.
    Every model that read the remark wrote that date, the check called it
    invented -- it is in no cell -- and the account was regenerated, on every
    run of that workbook, until the model stopped saying which days the
    balance covered. A day earlier in the same month as a source period end,
    written as the start of a range that ends ON that period end, is derived
    from the period, not made up. An invented year or month still fails: the
    range has to close on a date the source has.
    """
    year, month, day = key
    tail = _RANGE_END.match(body, match.end())
    if not tail:
        return False
    end = (int(tail.group(1) or year), int(tail.group(2) or month), int(tail.group(3)))
    return end in allowed and end[:2] == (year, month) and day < end[2]


def _date_reviews(final_content: str, df, *, source: Optional[SourceIndex] = None) -> List[Dict[str, Any]]:
    """Any date in the commentary that appears NOWHERE in the account's source.

    A real 21-slide deck shipped "截至2232年01月01日", "较1770年01月01日",
    "截至1938年01月01日", "截至2215年01月01日" and a dozen more against a
    databook whose only period ends are 2026-06-30, 2025-01-01, 2024-01-01 and
    2023-01-01. The trigger is in _period_reference_guidance: when an account
    carried no effective_date, the instruction rendered its date slot EMPTY --
    "首句必须仅说明截至的最新期末余额", four blanks in one paragraph -- and a
    model told to write "截至___" supplies something.

    That prompt hole is fixed separately, but a rule that only lives in the
    prompt is not a guardrail: the invented COMPARISON dates on that deck
    ("余额较1971年05月30日的2,608.3万元") sat in accounts whose opening date was
    correct, so filling the slot would not have caught them. This check does,
    and it is deterministic.

    Category is "hallucination" -- a date the source never contained is
    fabricated by definition, which is exactly what the retry gate is for.

    Silent when the source yields no dates at all: with nothing to judge
    against, flagging every date would be a guess, not a finding."""
    body = str(final_content or "")
    if not body:
        return []
    date_facts = (source.date_facts if source is not None and source.date_facts
                  else _harvest_source_date_facts(df))
    allowed = {tuple(fact["date"]): fact for fact in date_facts}
    if not allowed:
        return []
    first_seen: Dict[tuple, Any] = {}
    for pattern in (_DATE_CHI, _DATE_ISO):
        for match in pattern.finditer(body):
            key = tuple(int(part) for part in match.groups())
            if key not in allowed and not _opens_range_to_source_date(body, match, key, allowed):
                first_seen.setdefault(key, match)
    if not first_seen:
        return []
    known = "、".join(
        f"{y}年{m:02d}月{d:02d}日" for y, m, d in sorted(allowed)[:6]
    )
    reviews: List[Dict[str, Any]] = []
    for (y, m, d), match in first_seen.items():
        start = body.rfind("。", 0, match.start()) + 1
        end = body.find("。", match.end())
        end = (end + 1) if end >= 0 else len(body)
        sentence = body[start:end]
        lead = len(sentence) - len(sentence.lstrip())
        clause = sentence.strip() or body
        clause_span = ([start + lead, start + lead + len(clause)] if sentence.strip()
                       else [0, len(body)])
        # `expected` is the source date nearest the invented one, so a repair
        # has a candidate to cite rather than a list to choose from. It is a
        # candidate, not a certainty -- the reason still names every source date.
        nearest = min(allowed.values(), key=lambda f: abs(_date_ordinal(f["date"]) - _date_ordinal((y, m, d))))
        reviews.append({
            "clause": clause,
            "supported": False,
            "category": "hallucination",
            "code": "DATE_UNSUPPORTED",
            "conf": _CONF_DET_HALLUCINATION,
            "span": clause_span,
            # The offending date's OWN span, so a patch rewrites the date and
            # not the sentence around it.
            "amounts": [{
                "value": f"{y:04d}-{m:02d}-{d:02d}",
                "span": [match.start(), match.end()],
                "matched": False,
                "source_ref": None,
                "code": "DATE_UNSUPPORTED",
            }],
            "expected": nearest["value"],
            "source_ref": nearest,
            "reason": (
                f"日期 {y}年{m:02d}月{d:02d}日 并未出现在本科目的任何来源数据中"
                f"（来源日期为：{known}；最接近的来源日期：{nearest['value']}"
                f"，来自 {nearest.get('col_label')}）。日期不得自行推断或编造。"
            ),
        })
    return reviews


def _date_ordinal(parts) -> int:
    """A comparable day number for a (y, m, d) that may not be a real date.

    `datetime` refuses 2232-02-31, and the whole point here is dates the model
    invented, so this is deliberately arithmetic rather than a calendar."""
    y, m, d = (int(p) for p in parts)
    return y * 372 + m * 31 + d


_PCT_IN_CLAUSE = _PCT_CLAIM  # thousands separators, see _PCT_CLAIM
_DIR_UP = re.compile(r"增加|增长|上升|increase[sd]?|rose", re.IGNORECASE)
_DIR_DOWN = re.compile(r"减少|下降|decrease[sd]?|fell|declined", re.IGNORECASE)
#: A percentage may be off the movement's own by rounding only.
_DIR_PCT_TOL_PP = 0.5


def collect_direction_findings(final_content: str, df) -> List[Dict[str, Any]]:
    """_direction_reviews behind a swallow-everything guard.

    A brand-new detector must never be able to take the whole deterministic pass
    down: verify_commentary's other checks are what the retry gate and the deck
    highlighting run on, and an exception here would cost an account its
    clause_reviews entirely.
    """
    try:
        return _direction_reviews(final_content, df)
    except Exception:
        return []


def _direction_reviews(final_content: str, df) -> List[Dict[str, Any]]:
    """Clauses that quote a movement's own percentage and then name the wrong
    direction ("下降18.4%" against a +18.4% move).

    Report-only, on its own channel (see verify_commentary's direction_findings
    out-parameter) — NOT a clause_review. Record shape follows M5's contract
    findings so both channels read the same way.

    Movements come from `df.attrs["significant_movements"]`, precomputed at
    build time and present top-level on every variant. build_significant_
    movements is NOT called on the variant frame: it is dead there and crashes
    on the `*_formatted` columns.

    Know what that list is not. The "Significant movements" section the model
    reads is RECOMPUTED in prompts.py from the filtered frame, and both lists
    are capped at 3, so they can be disjoint; and the percentage a bullet most
    often quotes is the ACCOUNT TOTAL's move from _variance_analysis_guidance,
    which need not appear in either list. So this cannot ask "does the
    commentary agree with movement 1" — it has to anchor on the clause's own
    subject, and stay silent about everything it cannot pair:

      * the movement's description (or its display alias) must appear in the
        SAME clause;
      * the clause must quote a percentage within 0.5pp of abs(percent_change);
      * the clause must carry a direction word, and only one kind of it.

    Mandatory skips: percent_change None (a from-nil movement, which is NOT the
    sign-flip case and has no percentage to quote); from_value and to_value
    differing in sign; either value negative. Positive-to-positive only is what
    keeps 亏损扩大 and contra lines out — a negative base yields a signed −200%
    for what a reader correctly calls an increase. And if two supplied movements
    match one magnitude with opposite signs, the clause is ambiguous and is
    skipped rather than guessed at.
    """
    body = str(final_content or "")
    attrs = getattr(df, "attrs", None) or {}
    movements = [m for m in (attrs.get("significant_movements") or []) if isinstance(m, dict)]
    if not body or not movements:
        return []
    alias_map = attrs.get("display_description_map") or {}

    usable = []
    for movement in movements:
        pct = movement.get("percent_change")
        from_v, to_v = movement.get("from_value"), movement.get("to_value")
        if pct is None or from_v is None or to_v is None:
            continue
        if from_v < 0 or to_v < 0 or (from_v > 0) != (to_v > 0):
            continue
        description = str(movement.get("description") or "")
        if not description:
            continue
        names = {description}
        alias = alias_map.get(description)
        if alias:
            names.add(str(alias))
        usable.append((movement, float(pct), names))
    if not usable:
        return []

    findings: List[Dict[str, Any]] = []
    for start, end, clause in segment_clauses(body):
        quoted = [float(m.group(1).replace(",", "")) for m in _PCT_IN_CLAUSE.finditer(clause)]
        if not quoted:
            continue
        says_up, says_down = bool(_DIR_UP.search(clause)), bool(_DIR_DOWN.search(clause))
        if says_up == says_down:  # neither word, or both — nothing to contradict
            continue
        paired = [
            (movement, pct) for movement, pct, names in usable
            if any(name in clause for name in names)
            and any(abs(q - abs(pct)) <= _DIR_PCT_TOL_PP for q in quoted)
        ]
        if not paired:
            continue
        if len({pct > 0 for _m, pct in paired}) > 1:
            continue
        movement, pct = paired[0]
        if (pct > 0) == says_up:
            continue
        findings.append({
            "claim_id": f"direction::{movement.get('description')}::{abs(pct):.1f}",
            "kind": "direction",
            "detector": "direction_mismatch",
            "code": "DIRECTION_MISMATCH",
            "clause": clause,
            "span": [int(start), int(end)],
            "conf": 0.8,
            "patch_hint": ("increase" if pct > 0 else "decrease"),
            "reason": (
                f"'{movement.get('description')}' moved {pct:+.1f}% from "
                f"{movement.get('from_period')} to {movement.get('to_period')} "
                f"({movement.get('from_value'):,.0f} -> {movement.get('to_value'):,.0f}), "
                f"but the clause quoting that percentage calls it "
                f"{'an increase' if says_up else 'a decrease'}."
            ),
            "facts": {
                "description": str(movement.get("description")),
                "from_period": str(movement.get("from_period")),
                "to_period": str(movement.get("to_period")),
                "from_value": float(movement.get("from_value")),
                "to_value": float(movement.get("to_value")),
                "percent_change": float(pct),
            },
        })
    return findings


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

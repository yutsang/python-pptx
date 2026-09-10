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
_ATTR_KEYS_WITHHELD_FROM_PROMPTS = ("working_remark_notes",)


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
            residual = analysis_df.attrs.get("prompt_residual")
            if own and isinstance(residual, dict) and residual.get("amount") is not None:
                facts.append(_fact(
                    residual["amount"], "prompt_residual", sheet=cls_sheet,
                    row_desc="remainder handed to the model (total of %s component(s) less the "
                             "largest %s)" % (residual.get("component_count"),
                                              len(residual.get("listed") or [])),
                ))
            if own:
                facts += cls._period_movement_facts(analysis_df, cls_sheet)
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
            facts += [
                _fact(v * factor, "annualized_note" if own else "sibling_cell",
                      sheet=cls_sheet, row_desc=f"notes/remarks x{factor:.4g} (annualized)")
                for v in text_values
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

    def matches(self, target: float) -> Optional[Dict[str, Any]]:
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
                if abs(t - a) <= _WINDOW_SUM_REL_TOL * a:
                    return fact
                continue
            if abs(t - a) <= max(500.0, 0.05 * a):
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
        fact = source.matches(value)
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
_ENUM_RUNON = re.compile(r"(?:主要(?:包括|包含|为|系)|包括|包含)([^。]*)")
_RUNON_AMT = re.compile(r"(-?[\d,]+(?:\.\d+)?)\s*(万元|亿元|元)")
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
_STATED_TOTAL = re.compile(r"(?:合计|总额|余额合?计?)\s*(?:为)?\s*([\d,]+(?:\.\d+)?)\s*(万元|亿元|元)")
_SCALE = {"元": 1.0, "万元": 1e4, "亿元": 1e8}


def check_composition_adds_up(mapping_key: str, text: str) -> List[str]:
    """Message-only wrapper over _composition_findings.

    Kept because inspect_databook.py:2232 collects these as plain warning
    strings; the typed form below is what the verifier reads, so the category
    no longer has to be recovered by substring-matching the English message.
    """
    return [message for _code, message in _composition_findings(mapping_key, text)]


def _composition_findings(mapping_key: str, text: str) -> List[Tuple[str, str]]:
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
    items = _ENUM_ITEM.findall(scan)
    if not items:
        run = _ENUM_RUNON.search(scan)
        if run:
            items = _RUNON_AMT.findall(run.group(1))
    # The closing "其余X万元为…" is a component like any other. Added here
    # rather than inside the two patterns above because it can sit in either
    # form -- after a numbered list, or trailing a run-on one -- and because a
    # run-on span that already swallowed it must not count it twice.
    residual = _RESIDUAL_ITEM.search(scan)
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
    # A "其中" drill-down is part of the item before it, not a peer of it:
    # "…非关联公司-其他169.0万元，其中主要为某税务局166.8万元；…押金83.1万元;
    # 管理层调整36.7万元；其余13.0万元…" lists four components that hit the
    # 301.8万元 total exactly, and one figure that is INSIDE the first of them.
    # Counted as five they came to 468.6 and the account was reported as not
    # adding up. Dropping any one item and landing on an exact tie is that
    # shape; testing it arithmetically avoids having to decide what 其中 means
    # in a sentence, which it does not always mean.
    values = [float(v.replace(",", "")) * _SCALE.get(u, 1.0) for v, u in items]
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
    if direction_findings is not None:
        direction_findings.extend(collect_direction_findings(final_content, df))
    return out


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
        certain = code in ("COMPOSITION_UNIT_ERROR", "COMPOSITION_DOUBLE_COUNT")
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
            if key not in allowed:
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


_PCT_IN_CLAUSE = re.compile(r"(\d+(?:\.\d+)?)\s*%")
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
        quoted = [float(m.group(1)) for m in _PCT_IN_CLAUSE.finditer(clause)]
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

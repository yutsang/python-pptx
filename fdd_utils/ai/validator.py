"""What the claim verifier (claims.py) and the deck still take from the old text verifier: how a printed amount is read back from a
sentence at its printed precision (万元, 亿元, CNY millions), the half-unit of a printed figure, the numbers a note holds, the commentary
HTML the app paints with the verifier's red and orange, and the removal of a reasoning model's <think> block."""
from __future__ import annotations

import html
import re
from typing import Any, Dict, List, Optional, Tuple


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
# The English scale words the amount parser reads (million / mn / m, K, bn). Without them "CNY1.3 million" was given a half-ulp of
# 0.05 YUAN, so every rounded English figure failed the claim check as unmatched (found 2026-11-07, writing English balances).
_WRITTEN_EN_SCALE = re.compile(r"\d\s*(billion|bn|million|mn|m|thousand|k)(?![a-z])", re.IGNORECASE)
_EN_SCALE = {"billion": 1e9, "bn": 1e9, "million": 1e6, "mn": 1e6, "m": 1e6, "thousand": 1e3, "k": 1e3}


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
    else:
        english = _WRITTEN_EN_SCALE.search(str(written))
        if english:
            scale = _EN_SCALE[english.group(1).lower()]
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


def _numbers_in_text(text: str) -> List[float]:
    """Every number a remark could supply: scale-aware amounts (万/亿/million/comma)
    PLUS bare integers/decimals (e.g. '183600', '5000', '480')."""
    out = list(extract_amounts(text))
    for m in _BARE_NUMBER_RE.finditer(str(text or "")):
        v = _to_float(m.group(0))
        if v is not None:
            out.append(v)
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

from __future__ import annotations

# re-added: bound by an import in another section of the pre-split module
from ..workbook import INTERNAL_ROW_KEY
from typing import Any, Dict, List, Optional
from typing import Any, Dict, Optional, Tuple

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
    capital, audit fees, USD amounts) live only in the remarks.

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
        text_values = _numbers_in_text(_attr_text_blob(df))
        values += text_values
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
            values += [v * factor for v in text_values]
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


_ENUM_ITEM = re.compile(r"[1-9]）\s*[^；;]*?([\d,]+(?:\.\d+)?)\s*(万元|亿元|元)")
_ENUM_RUNON = re.compile(r"(?:主要)?(?:包括|包含|为|系)[^。；;]*?"
                         r"((?:[\u4e00-\u9fff]{2,10}[\d,]+(?:\.\d+)?\s*(?:万元|亿元|元)[、及和]?){2,})")
_RUNON_AMT = re.compile(r"([\d,]+(?:\.\d+)?)\s*(万元|亿元|元)")
_STATED_TOTAL = re.compile(r"(?:合计|总额|余额合?计?)\s*(?:为)?\s*([\d,]+(?:\.\d+)?)\s*(万元|亿元|元)")
_SCALE = {"元": 1.0, "万元": 1e4, "亿元": 1e8}


def check_composition_adds_up(mapping_key: str, text: str) -> List[str]:
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
    items = _ENUM_ITEM.findall(body)
    if not items:
        run = _ENUM_RUNON.search(body)
        if run:
            items = _RUNON_AMT.findall(run.group(1))
    if not m or len(items) < 2:
        return []
    total = float(m.group(1).replace(",", "")) * _SCALE.get(m.group(2), 1.0)
    listed = sum(float(v.replace(",", "")) * _SCALE.get(u, 1.0) for v, u in items)
    if total <= 0:
        return []
    gap = total - listed
    if abs(gap) / total <= 0.01:
        return []
    fmt = lambda v: f"{v/1e4:,.1f}万元"
    # A ratio near a power of ten is a unit error, not an omission: the model
    # took a raw CNY'000 cell and wrote 万元 against it. Worth saying so --
    # "66% unaccounted for" reads as a missing component, and the reader would
    # go looking for one that does not exist.
    ratio = listed / total if total else 0
    for _mult, _label in ((10.0, "10x"), (100.0, "100x"), (0.1, "1/10")):
        if abs(ratio - _mult) / _mult <= 0.05:
            return [
                f"[{mapping_key}] composition is {_label} the stated total "
                f"({fmt(listed)} vs {fmt(total)}) -- this is a UNIT error, not a "
                f"missing component: a raw CNY'000 figure written as 万元."
            ]
    if abs(ratio - 2.0) <= 0.05:
        return [
            f"[{mapping_key}] composition is exactly double the stated total "
            f"({fmt(listed)} vs {fmt(total)}) -- a parent line and the lines "
            f"that make it up have both been listed."
        ]
    return [
        f"[{mapping_key}] composition does not reach the stated total: "
        f"{len(items)} item(s) sum to {fmt(listed)} against {fmt(total)}, "
        f"leaving {fmt(gap)} ({abs(gap)/total:.0%}) unaccounted for. The reader "
        f"cannot tell whether the rest is an omission or a component with no name."
    ]


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
    out.extend(_composition_reviews(final_content))
    out.extend(_date_reviews(final_content, df))
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
    for message in check_composition_adds_up("", body):
        detail = message.split("] ", 1)[-1]
        certain = ("UNIT error" in detail) or ("exactly double" in detail)
        # Anchor on the sentence stating the total, so the deck highlights the
        # claim rather than the whole paragraph.
        match = _STATED_TOTAL.search(body)
        clause = body
        if match:
            start = body.rfind("。", 0, match.start()) + 1
            end = body.find("。", match.end())
            clause = body[start: (end + 1) if end >= 0 else len(body)].strip() or body
        reviews.append({
            "clause": clause,
            "supported": False,
            "category": "hallucination" if certain else "reasoning",
            "reason": detail,
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


def _harvest_source_dates(df) -> set:
    """Dates the account's own data actually contains -- period columns, the
    effective date, and any date written into a cell, note or detail row.

    Notes and remarks are included on purpose: a loan maturity or a lease end
    date is a legitimate date to quote and is not a period column. Grounding
    against the whole source, not just the period set, is the same contract
    SourceIndex already uses for amounts."""
    allowed: set = set()
    if df is None:
        return allowed
    attrs = getattr(df, "attrs", None) or {}
    integrity = attrs.get("integrity") or {}
    for key in ("effective_date", "raw_effective_date"):
        allowed |= _dates_in(integrity.get(key))
    try:
        for col in df.columns:
            allowed |= _dates_in(col)
            for cell in df[col].tolist():
                allowed |= _dates_in(cell)
    except Exception:
        pass
    table = attrs.get("presentation_detail_table") or {}
    for period in (table.get("periods") or []):
        allowed |= _dates_in(period)
    for row in (table.get("rows") or []):
        allowed |= _dates_in(row.get("label") if isinstance(row, dict) else row)
    for bucket in ("supporting_notes", "adjacent_detail_rows"):
        for item in (attrs.get(bucket) or []):
            allowed |= _dates_in(item)
    return allowed


def _date_reviews(final_content: str, df) -> List[Dict[str, Any]]:
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
    allowed = _harvest_source_dates(df)
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
        clause = body[start: (end + 1) if end >= 0 else len(body)].strip() or body
        reviews.append({
            "clause": clause,
            "supported": False,
            "category": "hallucination",
            "reason": (
                f"日期 {y}年{m:02d}月{d:02d}日 并未出现在本科目的任何来源数据中"
                f"（来源日期为：{known}）。日期不得自行推断或编造。"
            ),
        })
    return reviews


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

"""Repair one proven defect in place, instead of regenerating the whole bullet.

WHY THIS IS NOT `dedupe_enumeration_prefix` (removed 2026-08-11; its full
post-mortem is the comment block at pipeline.py:671-702, and that comment must
never be deleted). That was an unconditional post-processor guessing at
repetition, with no re-verification, which silently cut finished financial prose
-- it split "1,034.3万元" on a thousands separator and shipped "余额为1,。我方
核对了", and it ate "土地使用税" out of a longer word because the string had
appeared earlier in the paragraph.

A patch here is the opposite construction at every point:

  * it fires only on a DETERMINISTICALLY-PROVEN defect (a typed defect code out
    of verify_commentary, with a known correct value out of the account's own
    SourceIndex), never on a heuristic reading of the prose;
  * it splices only at a span the verifier asserts is verbatim -- the span is
    re-parsed back to the defective value before anything is written;
  * it never deletes text: every patch is a same-shape replacement of one token,
    or one sentence returned by the model in place of one sentence;
  * it must preserve every OTHER amount, every date, every structural marker and
    ~all of the length, checked as a multiset before and after;
  * it re-runs FULL verification with the same sibling set the production call
    used, and any failed guard discards the patch and falls back to today's
    regeneration.

`reasoning` flags are never repair targets -- a supportable inference is what the
deliverable wants (see count_defective_clauses). Neither is any defect class
whose correct value is not known: everything outside REPAIR_MODES routes to
regeneration unchanged.
"""

from __future__ import annotations

import re
import threading
from collections import Counter
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

#: Private names imported deliberately: _CLAUSE_END_CHARS is REFERENCED rather
#: than re-listed (re-listing it has already dropped `；` and `？` once), and the
#: two date patterns are the ones _date_reviews itself flags with, so a repair
#: reads exactly the dates the verifier judged.
from .validator import (
    _CLAUSE_END_CHARS,
    _DATE_CHI,
    _DATE_ISO,
    SourceIndex,
    collect_direction_findings,
    describe_fact,
    extract_amount_spans,
    extract_amounts,
    segment_clauses,
    verify_commentary,
)

#: Which defect code gets which repair, and nothing else does. Everything absent
#: from this map -- all COMPOSITION_*, both LLM_UNSUPPORTED_*, UNCATEGORIZED, a
#: review with no code at all, and a ratio-only retry trigger -- falls through to
#: today's full regeneration untouched.
#:
#: DIRECTION_MISMATCH is listed but is UNREACHABLE from production today:
#: _direction_reviews is report-only (verify_commentary's direction_findings
#: out-parameter), so no clause_review ever carries that code. The branch exists
#: for M3's promotion step and is exercised offline by the induced-defect
#: harness, which feeds a finding in review shape.
REPAIR_MODES: Dict[str, str] = {
    "AMOUNT_SCALE_ERROR": "fact_patch_det",
    "DIRECTION_MISMATCH": "fact_patch_det",
    "AMOUNT_UNSUPPORTED": "fact_patch_llm",
    "DATE_UNSUPPORTED": "fact_patch_llm",
}

#: A scale match is coincidence-prone, so the deterministic patch only trusts a
#: figure someone could have read off the sheet. classify_miss' own candidate
#: set (_OWN_HARD_KINDS) is wider than this -- it includes `analysis_cell` --
#: and window sums, note-blob numbers and annualized variants are wider still.
_PATCHABLE_FACT_KINDS = ("cell", "column_total")

_MIN_SENTENCE_RATIO = 0.60
_MIN_CONTENT_RATIO = 0.95
MAX_PATCH_ROUNDS = 2
_VALUE_REL_TOL = 0.01

#: A patch that produced one of these has glued a unit onto a unit -- the exact
#: failure mode that killed two earlier attempts at rewriting an amount, because
#: the amount regexes stop short of the trailing 元 (see _patch_span).
_DOUBLED_UNITS = ("元元", "万万", "亿亿", "元万", "元亿")

#: The two-part subtable bullet (prompts.py:852-866) puts a load-bearing line
#: break between "…明细如下：" and its ➢ items, and _strip_table_handoff depends
#: on it. A patch never touches a line carrying one of these.
_BULLET_MARKS = ("➢", "•", "‣", "▪")
_LEADING_DASH_RE = re.compile(r"^[-*]\s")
_ENUM_MARK_RE = re.compile(r"[1-9]\d?[）)]")
_STRUCTURAL_NEEDLE = "明细如下"

#: Currency prefix / scale-and-unit tail around the digits, so a replacement can
#: be rendered in the SPAN'S OWN shape rather than in some canonical form. The
#: alternation is longest-first; US$ before $.
#: Both comma forms are accepted in the digits and the ORIGINAL one is written
#: back: the Auditor introduces fullwidth thousands separators (13 occurrences in
#: Generator output, 46 in Auditor output across the archive), and rendering a
#: replacement with an ASCII comma would silently rewrite punctuation the writer
#: chose.
_SHAPE_RE = re.compile(
    r"^(?P<prefix>(?:人民币|人民幣|CNY|RMB|USD|HKD|US\$|\$)?[ \t]{0,3})"
    r"(?P<num>\d[\d,，]*(?:\.\d+)?)"
    r"(?P<tail>.*)$",
    re.IGNORECASE | re.DOTALL,
)

_FOREIGN_CURRENCY_RE = re.compile(r"^(?:USD|HKD|US\$|\$)", re.IGNORECASE)

#: Direction antonyms, keyed by the direction the SOURCE says is correct. Only
#: whole words on the English side; the Chinese verbs are unambiguous as
#: substrings.
_DIRECTION_SWAPS: Dict[str, List[Tuple[str, str]]] = {
    "increase": [
        ("减少", "增加"), ("減少", "增加"), ("下降", "上升"), ("下滑", "上升"),
        ("decreased", "increased"), ("decreases", "increases"), ("decrease", "increase"),
        ("declined", "increased"), ("fell", "rose"),
    ],
    "decrease": [
        ("增加", "减少"), ("增长", "下降"), ("增長", "下降"), ("上升", "下降"),
        ("increased", "decreased"), ("increases", "decreases"), ("increase", "decrease"),
        ("rose", "fell"),
    ],
}
_ENGLISH_DIRECTION_WORDS = {
    "decreased", "decreases", "decrease", "declined", "fell",
    "increased", "increases", "increase", "rose",
}


# ---------------------------------------------------------------------------
# routing
# ---------------------------------------------------------------------------

def repair_mode_for(review: Any) -> Optional[str]:
    """The repair mode for one clause_review, or None to regenerate.

    Deliberately strict about the category as well as the code: `reasoning` is
    an FDD consultant's supportable inference, which the deliverable wants and
    which no patch may touch.
    """
    if not isinstance(review, dict) or review.get("supported"):
        return None
    if str(review.get("category") or "").strip().lower() != "hallucination":
        return None
    return REPAIR_MODES.get(str(review.get("code") or ""))


def repairable_reviews(clause_reviews: Sequence[Any]) -> List[Dict[str, Any]]:
    return [r for r in (clause_reviews or []) if repair_mode_for(r)]


# ---------------------------------------------------------------------------
# span shape
# ---------------------------------------------------------------------------

def _shape(text: str) -> Optional[Tuple[str, str, str]]:
    """(currency prefix, digits, scale/unit tail) of one amount span.

    A TRAILING separator belongs to the tail, not to the number. The amount
    regexes read the digits as `[\\d,]*`, which happily swallows the comma that
    ends a clause -- measured, "…CNY219,658, CNY…" gives the span "CNY219,658,".
    Rewriting the number without moving that comma out first would delete the
    sentence's punctuation, and deleting text is the one thing this module may
    never do.
    """
    match = _SHAPE_RE.match(str(text or ""))
    if not match:
        return None
    prefix, num, tail = match.group("prefix"), match.group("num"), match.group("tail")
    while num and num[-1] in ",，":
        num, tail = num[:-1], num[-1] + tail
    if not num:
        return None
    return prefix, num, tail


def _parse_one_amount(text: str) -> Optional[float]:
    """The single amount `text` parses to, or None when it is not exactly one.

    "Exactly one, covering the whole slice" is the assertion that makes a splice
    safe: it proves the stored span really is that figure and nothing else.
    """
    spans = extract_amount_spans(text)
    if len(spans) != 1:
        return None
    value, start, end = spans[0]
    if start != 0 or end != len(text):
        return None
    return value


def _close(a: float, b: float, rel: float = _VALUE_REL_TOL) -> bool:
    return abs(a - b) <= max(rel * max(abs(a), abs(b)), 1e-9)


def _same_figure(written: float, source: float) -> bool:
    """Whether a WRITTEN amount and a SOURCE figure are the same figure.

    Deliberately SourceIndex.matches' own band, max(500, 5%), not a tighter one:
    the display forms this pipeline writes are rounded (1dp of 万, 1dp of
    million), and house style rounds them again -- polish_english_commentary
    collapses 'CNY1.28 million' to 'CNY1.3 million', a 1.2% move. Judging the
    replacement more strictly than the verifier judges any other amount would
    refuse patches the verifier then certifies as correct.
    """
    return abs(abs(written) - abs(source)) <= max(500.0, 0.05 * abs(source))


def render_amount_in_shape(span_text: str, current_value: float, new_value: float) -> Optional[str]:
    """`span_text` re-rendered so it reads as `new_value`, same shape throughout.

    Same currency prefix present-or-absent, same scale word and unit, same
    thousands-grouping habit, same decimals where the value survives them. The
    scale is derived from the span itself (current_value / its digits), so no
    table of unit words can go stale.

    Public because the induced-defect harness corrupts a token through this same
    helper -- restoration is then well-defined by construction rather than by a
    second, independently-written formatter.
    """
    shape = _shape(span_text)
    if not shape:
        return None
    prefix, num_str, tail = shape
    try:
        num = float(num_str.replace(",", "").replace("，", ""))
    except ValueError:
        return None
    if not num or not current_value:
        return None
    scale = current_value / num
    if scale <= 0:
        return None
    target = abs(new_value) / scale
    separator = "，" if "，" in num_str else ","
    grouped = separator in num_str
    decimals = len(num_str.split(".")[1]) if "." in num_str else 0
    # At most ONE decimal beyond the span's own precision, and the result must
    # still read as a real figure. Allowing more let a 20元 source cell be
    # rendered into a 万元 span as "0.002万元" -- arithmetically the right value,
    # prose no consultant would write, and the guards downstream then had to
    # catch it. A value that does not fit the span's own display precision is a
    # refusal, not a reformatting job.
    for extra in (0, 1):
        spec = "%s.%df" % ("," if grouped else "", decimals + extra)
        digits = format(target, spec).replace(",", separator)
        candidate = prefix + digits + tail
        parsed = _parse_one_amount(candidate)
        if parsed is None or parsed == 0 or not _same_figure(parsed, abs(new_value)):
            continue
        if _shape(candidate) is None or _shape(candidate)[0] != prefix or _shape(candidate)[2] != tail:
            continue
        return candidate
    return None


# ---------------------------------------------------------------------------
# targets
# ---------------------------------------------------------------------------

def _span_of(entry: Any) -> Optional[Tuple[int, int]]:
    span = (entry or {}).get("span") if isinstance(entry, dict) else None
    if not isinstance(span, (list, tuple)) or len(span) != 2:
        return None
    try:
        start, end = int(span[0]), int(span[1])
    except (TypeError, ValueError):
        return None
    return (start, end) if 0 <= start < end else None


def _target_for(review: Dict[str, Any], content: str) -> Optional[Dict[str, Any]]:
    """The one token this review is about, located in `content`, or None."""
    code = str(review.get("code") or "")
    clause_span = _span_of(review)
    if clause_span is None or clause_span[1] > len(content):
        return None

    if code == "DIRECTION_MISMATCH":
        hint = str(review.get("patch_hint") or "").strip().lower()
        if hint not in _DIRECTION_SWAPS:
            return None
        return {
            "code": code, "kind": "direction", "clause_span": clause_span,
            "token_span": clause_span, "patch_hint": hint,
            "claim_id": review.get("claim_id"),
        }

    entries = [e for e in (review.get("amounts") or []) if isinstance(e, dict)]
    if code == "DATE_UNSUPPORTED":
        for entry in entries:
            span = _span_of(entry)
            expected = review.get("expected")
            if span is None or span[1] > len(content) or not expected:
                continue
            return {
                "code": code, "kind": "date", "clause_span": clause_span,
                "token_span": span, "value": str(entry.get("value") or ""),
                "expected": str(expected), "source_ref": review.get("source_ref"),
            }
        return None

    for entry in entries:
        if entry.get("matched") or str(entry.get("code") or "") != code:
            continue
        span = _span_of(entry)
        if span is None or span[1] > len(content):
            continue
        target = {
            "code": code, "kind": "amount", "clause_span": clause_span,
            "token_span": span, "value": entry.get("value"),
            "expected": entry.get("expected"), "source_ref": entry.get("source_ref"),
            "ambiguous": bool(entry.get("ambiguous")),
            "nearest": entry.get("nearest"),
        }
        if target["value"] is None:
            continue
        return target
    return None


# ---------------------------------------------------------------------------
# sentence extension
# ---------------------------------------------------------------------------

def _sentence_span(content: str, token_span: Tuple[int, int]) -> Optional[Tuple[int, int]]:
    """The clause spans around the token merged out to a sentence, or None.

    Terminators are `_CLAUSE_END_CHARS` -- referenced, never re-listed, because
    re-listing it has already dropped `；` and `？` once. `\\n` is NOT in that
    set, so it is enforced separately: a sentence that would cross a line break
    is refused outright, since the line break between "…明细如下：" and its ➢
    items is load-bearing for the deck.
    """
    spans = segment_clauses(content)
    if not spans:
        return None
    start_tok, end_tok = token_span
    hit = None
    for i, (s, e, _text) in enumerate(spans):
        if s <= start_tok < e or (start_tok <= s and end_tok >= e):
            hit = i
            break
    if hit is None:
        return None

    first = hit
    while first > 0:
        prev = spans[first - 1][2]
        if prev and prev[-1] in _CLAUSE_END_CHARS:
            break
        if "\n" in content[spans[first - 1][0]:spans[first][0]]:
            break
        first -= 1
    last = hit
    while last + 1 < len(spans):
        text = spans[last][2]
        if text and text[-1] in _CLAUSE_END_CHARS:
            break
        if "\n" in content[spans[last][1]:spans[last + 1][1]]:
            break
        last += 1

    start, end = spans[first][0], spans[last][1]
    if end < end_tok:
        end = end_tok
    if start > start_tok:
        start = start_tok
    if "\n" in content[start:end]:
        return None
    return (start, end)


def _touches_bullet(text: str) -> bool:
    body = str(text or "")
    return any(mark in body for mark in _BULLET_MARKS) or bool(_LEADING_DASH_RE.match(body.lstrip()))


def _structural_markers(text: str) -> Tuple[Any, ...]:
    body = str(text or "")
    return (
        tuple(body.count(mark) for mark in _BULLET_MARKS),
        tuple(sorted(Counter(_ENUM_MARK_RE.findall(body)).items())),
        body.count(_STRUCTURAL_NEEDLE),
    )


# ---------------------------------------------------------------------------
# guards
# ---------------------------------------------------------------------------

def _dates_in_text(text: str) -> Counter:
    found: Counter = Counter()
    for pattern in (_DATE_CHI, _DATE_ISO):
        for match in pattern.finditer(str(text or "")):
            found[tuple(int(part) for part in match.groups())] += 1
    return found


def _amount_counter(text: str) -> Counter:
    return Counter(round(float(v), 2) for v in extract_amounts(text))


def _defect_counter(reviews: Sequence[Any]) -> Counter:
    out: Counter = Counter()
    for review in reviews or []:
        if not isinstance(review, dict) or review.get("supported"):
            continue
        if str(review.get("category") or "").strip().lower() != "hallucination":
            continue
        out[str(review.get("code") or "UNCATEGORIZED")] += 1
    return out


def _pre_style_guards(original_sentence: str, patched_sentence: str) -> Optional[str]:
    """The guards that need the sentence span, which house style then destroys.

    House style shifts offsets (the English pass even collapses newlines), so a
    post-style sentence comparison is not implementable -- hence two scopes.
    """
    if not patched_sentence.strip():
        return "patch is empty"
    # Equality, not absence: a clause legitimately spans a line break (\n is not
    # in _CLAUSE_END_CHARS), and a token patch inside such a clause must be
    # allowed. What may never happen is a patch ADDING or REMOVING one, which
    # would move the ➢ items the deck's two-part bullet is cut on.
    if patched_sentence.count("\n") != original_sentence.count("\n"):
        return "patch changes the line breaks"
    if len(patched_sentence) < _MIN_SENTENCE_RATIO * len(original_sentence):
        return "patched sentence is %d%% of the original (floor %d%%)" % (
            round(100.0 * len(patched_sentence) / max(len(original_sentence), 1)),
            round(100 * _MIN_SENTENCE_RATIO),
        )
    if _structural_markers(patched_sentence) != _structural_markers(original_sentence):
        return "structural markers changed"
    for doubled in _DOUBLED_UNITS:
        if doubled in patched_sentence and doubled not in original_sentence:
            return "doubled unit token '%s'" % doubled
    return None


def _post_style_guards(
    before: str,
    after: str,
    reviews_before: Sequence[Any],
    reviews_after: Sequence[Any],
    target: Dict[str, Any],
) -> Optional[str]:
    """Whole-content guards, pre-patch vs post-patch, both house-styled once.

    The house-style block is idempotent on real output (measured: re-applying it
    to 24 English and 104 Chinese archived finals produced zero differences), so
    the stored content and the re-styled patch are directly comparable.
    """
    if len(after) < _MIN_CONTENT_RATIO * len(before):
        return "content shrank to %d%% (floor %d%%)" % (
            round(100.0 * len(after) / max(len(before), 1)), round(100 * _MIN_CONTENT_RATIO),
        )
    if len(segment_clauses(after)) < len(segment_clauses(before)):
        return "clause count dropped"

    amounts_before, amounts_after = _amount_counter(before), _amount_counter(after)
    if sum(amounts_after.values()) != sum(amounts_before.values()):
        return "amount count changed (%d -> %d)" % (
            sum(amounts_before.values()), sum(amounts_after.values()),
        )
    removed, added = amounts_before - amounts_after, amounts_after - amounts_before
    dates_before, dates_after = _dates_in_text(before), _dates_in_text(after)
    dates_removed, dates_added = dates_before - dates_after, dates_after - dates_before

    if target["kind"] == "amount":
        if sum(removed.values()) != 1 or not _close(list(removed)[0], abs(float(target["value"]))):
            return "amount multiset changed beyond the defective token"
        if sum(added.values()) > 1:
            return "patch introduced more than one new amount"
        expected = target.get("expected")
        if expected is not None and added and not _same_figure(list(added)[0], float(expected)):
            return "the replacement amount is not the expected source figure"
        if dates_removed or dates_added:
            return "date set changed"
    elif target["kind"] == "date":
        if removed or added:
            return "amount multiset changed"
        want = tuple(int(p) for p in str(target["value"]).split("-"))
        if sum(dates_removed.values()) != 1 or list(dates_removed)[0] != want:
            return "date set changed beyond the defective date"
        if sum(dates_added.values()) > 1:
            return "patch introduced more than one new date"
    else:  # direction
        if removed or added:
            return "amount multiset changed"
        if dates_removed or dates_added:
            return "date set changed"

    before_counts = _defect_counter(reviews_before)
    before_counts[target["code"]] = max(before_counts[target["code"]] - 1, 0)
    after_counts = _defect_counter(reviews_after)
    for code, count in after_counts.items():
        if count > before_counts.get(code, 0):
            return "new %s defect(s) after the patch" % code
    return None


def _target_still_present(target: Dict[str, Any], reviews_after: Sequence[Any], after: str, df) -> bool:
    if target["kind"] == "direction":
        claim_id = target.get("claim_id")
        return any(f.get("claim_id") == claim_id for f in collect_direction_findings(after, df))
    for review in reviews_after or []:
        if not isinstance(review, dict) or review.get("supported"):
            continue
        for entry in review.get("amounts") or []:
            if not isinstance(entry, dict) or entry.get("matched"):
                continue
            if str(entry.get("code") or review.get("code") or "") != target["code"]:
                continue
            value = entry.get("value")
            if target["kind"] == "date":
                if str(value) == str(target["value"]):
                    return True
            elif isinstance(value, (int, float)) and _close(float(value), abs(float(target["value"]))):
                return True
    return False


# ---------------------------------------------------------------------------
# the patches themselves
# ---------------------------------------------------------------------------

def _det_scale_patch(content: str, target: Dict[str, Any]) -> Tuple[Optional[str], Optional[str]]:
    """(replacement token, refusal). Every guard here is mandatory."""
    if target.get("ambiguous"):
        return None, "classify_miss reported an ambiguous scale factor"
    expected = target.get("expected")
    if expected is None:
        return None, "no expected value"
    fact = target.get("source_ref") or {}
    if str(fact.get("kind") or "") not in _PATCHABLE_FACT_KINDS:
        return None, "matched fact kind %r is not a readable cell" % fact.get("kind")
    start, end = target["token_span"]
    span_text = content[start:end]
    # A figure written in a FOREIGN currency cannot be scale-corrected from this
    # pool: every source fact is in the databook's own reporting currency, so a
    # factor that fits is arithmetic coincidence, not evidence. Measured on the
    # archive: "USD 3.00 million" of paid-in capital matched the CNY column
    # total 299,716,584 at factor 100, and the patch would have shipped
    # "USD 299.72 million". Two archived runs, both caught downstream; caught
    # here instead, where the reason is legible.
    if _FOREIGN_CURRENCY_RE.match(span_text.strip()):
        return None, "foreign-currency figure: the source pool is not in this currency"
    parsed = _parse_one_amount(span_text)
    if parsed is None or not _close(parsed, abs(float(target["value"])), rel=1e-6):
        return None, "the stored span does not re-parse to the defective value"
    replacement = render_amount_in_shape(span_text, parsed, float(expected))
    if replacement is None:
        return None, "no same-shape rendering of the expected value"
    return replacement, None


def _det_direction_patch(content: str, target: Dict[str, Any]) -> Tuple[Optional[str], Optional[str]]:
    start, end = target["token_span"]
    clause = content[start:end]
    swaps = _DIRECTION_SWAPS[target["patch_hint"]]
    patched = clause
    for wrong, right in swaps:
        if wrong not in patched:
            continue
        if wrong in _ENGLISH_DIRECTION_WORDS:
            patched = re.sub(r"\b%s\b" % re.escape(wrong), right, patched)
        else:
            patched = patched.replace(wrong, right)
    if patched == clause:
        return None, "no direction word to swap"
    return patched, None


def _fact_lines(source: Optional[SourceIndex], target: Dict[str, Any], language: str, limit: int = 12) -> List[str]:
    """The correct facts, so the model cannot invent the right value.

    This is the whole reason the patch prompt exists rather than a bare "fix
    it": the model is handed the account's own figures with their provenance and
    a display form, and told to choose among them.
    """
    lines: List[str] = []
    if target["kind"] == "date":
        facts = list((source.date_facts if source is not None else []) or [])
        for fact in facts[:limit]:
            lines.append("- %s（来源：%s）" % (fact.get("value"), fact.get("col_label") or fact.get("sheet")))
        if target.get("expected"):
            lines.append("- 最接近的来源日期 / nearest source date: %s" % target["expected"])
        return lines
    if source is None:
        return lines
    candidates = [f for f in source.facts if f.get("kind") in _PATCHABLE_FACT_KINDS and f.get("value")]
    try:
        pivot = abs(float(target.get("value") or 0.0))
    except (TypeError, ValueError):
        pivot = 0.0
    candidates.sort(key=lambda f: abs(abs(float(f["value"])) - pivot))
    for fact in candidates[:limit]:
        lines.append("- %s | %s" % (_display_amount(float(fact["value"]), language), describe_fact(fact)))
    nearest = target.get("nearest")
    if isinstance(nearest, dict):
        lines.append("- (closest source figure) %s" % describe_fact(nearest))
    return lines


def _display_amount(value: float, language: str) -> str:
    """The form the deliverable writes an amount in, for the prompt only."""
    magnitude = abs(value)
    if str(language or "").startswith("Chi"):
        if magnitude >= 1e8:
            return "%.1f亿元" % (value / 1e8)
        if magnitude >= 1e4:
            return "%.1f万元" % (value / 1e4)
        return "%.0f元" % value
    if magnitude >= 1e6:
        return "CNY%.1f million" % (value / 1e6)
    if magnitude >= 1e3:
        return "CNY%.1fK" % (value / 1e3)
    return "CNY%.0f" % value


#: A heading line the model put in front of its answer: no digits, ends in a
#: colon. Narrow on purpose — a line carrying a figure is never dropped, so this
#: can only ever remove text that contains nothing to preserve.
_PREAMBLE_LINE_RE = re.compile(r"^[^\d\n]{0,60}[:：]\s*$")


def _single_sentence(cleaned: str) -> Optional[str]:
    """The one sentence in a repair reply, or None if it is not one sentence.

    clean_agent_output already strips <think> blocks and the preambles this
    pipeline has met before, but its list says "output" ("here is the corrected
    OUTPUT:"), and a model asked for a SENTENCE writes "Here is the corrected
    sentence:". Rather than widen a cleaner every other agent shares, a heading
    line is dropped here. Anything still multi-line is refused: the splice must
    not introduce a line break, because the break between "…明细如下：" and its
    ➢ items is what the deck's two-part bullet depends on.
    """
    lines = [line for line in str(cleaned or "").splitlines() if line.strip()]
    while len(lines) > 1 and _PREAMBLE_LINE_RE.match(lines[0].strip()):
        lines.pop(0)
    if len(lines) != 1:
        return None
    return lines[0].strip()


def _repair_prompts(
    *,
    content: str,
    sentence: str,
    review: Dict[str, Any],
    target: Dict[str, Any],
    source: Optional[SourceIndex],
    language: str,
    style_pack=None,
) -> Tuple[str, str]:
    chinese = str(language or "").startswith("Chi")
    style_lines: List[str] = []
    if style_pack is not None:
        for getter in ("language_instruction", "common_formatting_rules"):
            try:
                text = str(getattr(style_pack, getter)() or "").strip()
            except Exception:  # pragma: no cover - a missing style pack is not fatal
                text = ""
            if text:
                style_lines.append(text)
    if chinese:
        system = (
            "你是财务尽职调查报告的校对员。任务只有一个：把给定的一句话中被证实错误的"
            "数字或日期改正，其余一字不动。\n"
            "规则：\n"
            "1. 只输出改正后的那一句，不要输出解释、标题、引号或任何前后缀。\n"
            "2. 只允许改动被指出的那一个数字/日期，句中其他金额、日期、百分比、"
            "标点、结构标记必须原样保留。\n"
            "3. 正确数值只能从下面给出的来源数据中选，不得自行推算或编造。\n"
            "4. 保持原句的计量单位写法（万元／亿元／元）与原有格式。"
        )
    else:
        system = (
            "You are a proof-reader on a financial due diligence report. Your only task "
            "is to correct the one figure or date that has been proven wrong in a single "
            "sentence, changing nothing else.\n"
            "Rules:\n"
            "1. Output ONLY the corrected sentence — no explanation, no heading, no quotes.\n"
            "2. Change only the one flagged figure/date. Every other amount, date, "
            "percentage, punctuation mark and structural marker must survive verbatim.\n"
            "3. Take the correct value from the source facts below. Never compute or "
            "invent one.\n"
            "4. Keep the sentence's existing unit and number format."
        )
    if style_lines:
        system += "\n\n" + "\n".join(style_lines)

    facts = _fact_lines(source, target, language)
    user = "\n".join([
        "[FULL BULLET — read-only context, do not rewrite]",
        content.strip(),
        "",
        "[SENTENCE TO CORRECT]",
        sentence.strip(),
        "",
        "[DEFECT] %s" % target["code"],
        str(review.get("reason") or "").strip(),
        "",
        "[SOURCE FACTS]",
        "\n".join(facts) if facts else "(none available)",
        "",
        ("[只输出改正后的这一句]" if chinese else "[Return ONLY the corrected sentence]"),
    ])
    return system, user


def _call_repair_model(
    ai_helper,
    system_prompt: str,
    user_prompt: str,
    *,
    timeout: int,
    temperature: float = 0.3,
    max_tokens: int = 300,
    top_p: Optional[float] = None,
) -> Dict[str, Any]:
    """One call with Auditor-like settings, behind the same timeout wall the
    stage loop uses. Not routed through pipeline._run_ai_call because that reads
    its temperature and max_tokens from an agent config keyed by stage name, and
    a 300-token proof-reading call is not a stage."""
    box: Dict[str, Any] = {}

    def run() -> None:
        try:
            box["response"] = ai_helper.get_response(
                user_prompt, system_prompt,
                temperature=temperature, max_tokens=max_tokens, top_p=top_p,
            )
        except Exception as exc:  # pragma: no cover - provider failures are normal
            box["error"] = exc

    thread = threading.Thread(target=run, daemon=True)
    thread.start()
    thread.join(timeout=timeout)
    if "error" in box:
        raise box["error"]
    if "response" not in box:
        raise TimeoutError("repair call timeout after %s seconds" % timeout)
    return box["response"] or {}


# ---------------------------------------------------------------------------
# the loop
# ---------------------------------------------------------------------------

def repair_content(
    *,
    content: str,
    clause_reviews: Sequence[Any],
    df,
    sibling_dfs: Optional[List[Any]] = None,
    language: str = "Eng",
    statement_type: str = "",
    mapping_key: str = "",
    ai_helper=None,
    logger=None,
    style_pack=None,
    allow_llm: bool = True,
    max_rounds: int = MAX_PATCH_ROUNDS,
    call_timeout: int = 60,
) -> Dict[str, Any]:
    """Patch the proven defects in `content`, or report why nothing was patched.

    Returns {content, clause_reviews, changed, log}. `content` and
    `clause_reviews` are the ORIGINALS unless `changed` is True, so a caller can
    assign unconditionally. `log` is the per-attempt audit record the plan asks
    for: {mode, code, span, before, after, verified, reason}.

    Re-verification uses the sibling set the caller passes, which must be the one
    the production call used -- a different sibling set is a different grounding
    pool and manufactures defects that were never there.
    """
    from .pipeline import apply_house_style, clean_agent_output  # circular at import time

    log: List[Dict[str, Any]] = []
    current, reviews = str(content or ""), list(clause_reviews or [])
    changed = False
    llm_calls = 0
    if df is None:
        # No frame, no re-verification, no repair. Every guard here is a
        # comparison against the source data.
        return {"content": current, "clause_reviews": reviews, "changed": False, "log": log}
    source = SourceIndex.from_df(df, sibling_dfs=sibling_dfs)

    for _round in range(max(0, int(max_rounds))):
        candidates = repairable_reviews(reviews)
        if not candidates:
            break
        progressed = False
        for review in candidates:
            mode = repair_mode_for(review)
            target = _target_for(review, current)
            if target is None:
                log.append({"mode": mode, "code": review.get("code"), "span": None,
                            "verified": False, "reason": "no locatable token for this defect"})
                continue

            token_span = target["token_span"]
            record: Dict[str, Any] = {
                "mode": mode, "code": target["code"], "span": list(token_span),
                "before": current[token_span[0]:token_span[1]], "after": None, "verified": False,
            }

            # --- build the replacement text -------------------------------
            if mode == "fact_patch_det" and target["code"] == "AMOUNT_SCALE_ERROR":
                replacement, refusal = _det_scale_patch(current, target)
                splice_span, original_piece = token_span, current[token_span[0]:token_span[1]]
            elif mode == "fact_patch_det":
                replacement, refusal = _det_direction_patch(current, target)
                splice_span, original_piece = token_span, current[token_span[0]:token_span[1]]
            else:
                sentence_span = _sentence_span(current, token_span)
                if sentence_span is None:
                    replacement, refusal, splice_span, original_piece = None, "sentence crosses a line break", None, ""
                else:
                    original_piece = current[sentence_span[0]:sentence_span[1]]
                    splice_span = sentence_span
                    if _touches_bullet(original_piece):
                        replacement, refusal = None, "sentence touches a bullet prefix"
                    elif not allow_llm or ai_helper is None:
                        replacement, refusal = None, "llm patching disabled"
                    elif llm_calls >= max_rounds:
                        # Cost ceiling, and it has to be counted rather than
                        # implied: the rounds bound how many patches LAND, not
                        # how many are attempted, and an account carrying three
                        # unsupported dates would otherwise make three calls per
                        # round. A retry is 3 full calls; this stays under that.
                        replacement, refusal = None, "repair call budget spent"
                    else:
                        llm_calls += 1
                        replacement, refusal = _llm_patch(
                            ai_helper=ai_helper, logger=logger, mapping_key=mapping_key,
                            content=current, sentence=original_piece, review=review,
                            target=target, source=source, language=language,
                            style_pack=style_pack, call_timeout=call_timeout,
                            clean=clean_agent_output,
                        )
            if replacement is None:
                record["reason"] = refusal
                log.append(record)
                continue
            # The audit record names what is actually being replaced, which for
            # an LLM patch is the sentence, not the token inside it.
            record["span"] = list(splice_span)
            record["before"] = original_piece

            # --- pre-house-style guards, while the span is still known ----
            # Graded over the CLAUSE, not over what was spliced. A token patch
            # replaces "人民币1,234,567元" with "人民币12,345元" -- 40% of the
            # original bytes, which reads as an amputation against a "keep >=60%
            # of it" rule written for a whole-sentence rewrite, and refused four
            # real repairs in the archive before this scope was pinned. The
            # sentence around it is what must survive, and it does.
            guard_span = target.get("clause_span") or splice_span
            if not (guard_span[0] <= splice_span[0] and splice_span[1] <= guard_span[1]):
                guard_span = splice_span
            old_guard = current[guard_span[0]:guard_span[1]]
            new_guard = (current[guard_span[0]:splice_span[0]] + replacement
                         + current[splice_span[1]:guard_span[1]])
            refusal = _pre_style_guards(old_guard, new_guard)
            if refusal:
                record.update({"after": replacement, "reason": refusal})
                log.append(record)
                continue

            spliced = current[:splice_span[0]] + replacement + current[splice_span[1]:]
            styled = apply_house_style(spliced, language, statement_type)
            record["after"] = replacement

            # The baseline is the CURRENT text put through the same house-style
            # pass, not the current text itself. In production the two are the
            # same string (everything stored has already been styled once, and
            # the block is idempotent), so this changes nothing there. It
            # matters offline: replayed against ARCHIVED finals, today's
            # polish_english_commentary is not a no-op -- it collapses
            # 'CNY1.03 million' to 1dp and deletes annualised parentheticals --
            # and comparing an unstyled baseline against a styled patch charges
            # the patch for house style's edits. Measured on one archived
            # English run: 27 of 41 deterministic patches were refused for an
            # "amount multiset changed" that house style, not the patch, had
            # caused.
            try:
                styled_before = apply_house_style(current, language, statement_type)
                reviews_before = verify_commentary(styled_before, df, None, sibling_dfs=sibling_dfs)
                reviews_after = verify_commentary(styled, df, None, sibling_dfs=sibling_dfs)
            except Exception as exc:  # pragma: no cover - defensive
                record["reason"] = "re-verification failed: %s" % exc
                log.append(record)
                continue

            if _target_still_present(target, reviews_after, styled, df):
                record["reason"] = "the targeted defect is still there after the patch"
                log.append(record)
                continue
            refusal = _post_style_guards(styled_before, styled, reviews_before, reviews_after, target)
            if refusal:
                record["reason"] = refusal
                log.append(record)
                continue

            record["verified"] = True
            log.append(record)
            current, reviews = styled, reviews_after
            changed = progressed = True
            break  # re-read the defects from the patched text before continuing

        if not progressed:
            break

    return {"content": current if changed else str(content or ""),
            "clause_reviews": reviews if changed else list(clause_reviews or []),
            "changed": changed, "log": log}


def _llm_patch(
    *,
    ai_helper,
    logger,
    mapping_key: str,
    content: str,
    sentence: str,
    review: Dict[str, Any],
    target: Dict[str, Any],
    source: Optional[SourceIndex],
    language: str,
    style_pack,
    call_timeout: int,
    clean: Callable[[str], str],
) -> Tuple[Optional[str], Optional[str]]:
    system_prompt, user_prompt = _repair_prompts(
        content=content, sentence=sentence, review=review, target=target,
        source=source, language=language, style_pack=style_pack,
    )
    top_p = None
    try:
        top_p = ai_helper.get_agent_settings("subagent_2").get("top_p")
    except Exception:  # pragma: no cover - a helper without agent settings is fine
        top_p = None
    if logger is not None:
        try:
            logger.log_agent_start("repair", mapping_key)
        except Exception:  # pragma: no cover
            pass
    try:
        response = _call_repair_model(
            ai_helper, system_prompt, user_prompt,
            timeout=call_timeout, temperature=0.3, max_tokens=300, top_p=top_p,
        )
    except Exception as exc:
        if logger is not None:
            try:
                logger.log_error("repair", mapping_key, exc)
            except Exception:  # pragma: no cover
                pass
        return None, "repair call failed: %s" % str(exc)[:120]
    if logger is not None:
        try:
            logger.log_agent_complete("repair", mapping_key, response, system_prompt, user_prompt)
        except Exception:  # pragma: no cover
            pass
    # EVERY model output in this pipeline goes through clean_agent_output.
    # Without it a <think> block or a "Here is the corrected sentence:" preamble
    # lands verbatim in a client bullet.
    patched = _single_sentence(clean(str(response.get("content") or "")))
    if patched is None:
        return None, "repair response is not a single sentence"
    if not patched:
        return None, "empty repair response"
    return patched, None


__all__ = [
    "MAX_PATCH_ROUNDS",
    "REPAIR_MODES",
    "render_amount_in_shape",
    "repair_content",
    "repair_mode_for",
    "repairable_reviews",
]

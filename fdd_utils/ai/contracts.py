from __future__ import annotations

"""Claim contracts — deterministic post-conditions on the commentary text.

A claim contract asks one question per account: *the prompt injected a
specific, computed instruction — did the shipped sentence carry it out?*
The four detectors here each correspond to a guidance block that
``PromptEngine`` already renders deterministically, so a contract is only
built when that block actually fired. Nothing is inferred about accounts the
prompt said nothing about.

**v1 is MEASUREMENT ONLY (plan M5).** Read the failure rates on real books
first; only then decide whether anything may act on them. Specifically:

- No retry is triggered. The feedback loop gates on a *hallucination* found by
  ``verify_commentary``; a missing claim is not one, and wiring it in would
  spend tokens on a rule whose false-positive rate is still unknown.
- No text is patched.
- **Results never enter ``clause_reviews``.** A ``CLAIM_MISSING`` entry there
  would land in the unsupported ratio the retry gate reads, and it would be
  painted by ``build_highlighted_commentary_html`` and
  ``_build_clause_segments`` — i.e. it would colour the client deck. The
  verdicts ride on their own key, ``results[key]["claim_contract"]``, which no
  deck-side consumer reads.

The verdict dict stored on ``results`` is ``{claim_id: "pass" | "fail"}`` —
plain strings only, because ``results`` is ``yaml.dump``'d into ``results.yml``
at the end of a paid run and a stray numpy scalar raises ``RepresenterError``
there. The richer per-claim record (facts, what was looked for, what was seen)
is returned to the caller for reporting and is deliberately NOT stored.

Detection strategy: every detector reads the *rendered guidance string*, not a
re-derivation of the underlying arithmetic. The guidance is the single source
of truth for whether the instruction fired and for the exact figure the model
was handed; re-computing it here would let the two drift apart silently, and a
contract graded against a number the model was never shown is worthless.
"""

import re
from typing import Any, Dict, Iterable, List, Optional

PASS = "pass"
FAIL = "fail"

#: Where the verdicts are filed on the results dict. Not a mapping_key and not
#: a clause_reviews field — see the module docstring.
CLAIM_CONTRACT_KEY = "claim_contract"

# Markers emitted by PromptEngine._variance_analysis_guidance /
# ._composition_guidance. A detector fires ONLY when its marker is present, so
# an account whose prompt carried no such instruction gets no claim at all
# (rather than a claim that passes vacuously and flatters the pass rate).
_NIL_MARKERS = ("[REQUIRED OPENING]", "【首句强制】")
_MATERIAL_MARKERS = ("[MATERIAL MOVEMENT]", "【重大变动提示】")
_REMAINDER_MARKERS = ("[REMAINDER ALREADY COMPUTED]", "【余额差额已算好】")


# ---------------------------------------------------------------------------
# config
# ---------------------------------------------------------------------------

def claim_contracts_enabled(processing_config: Optional[Dict[str, Any]] = None) -> bool:
    """``processing.claim_contracts.enabled``, default **False**.

    The default lives here in code, not in ``fdd_utils/config.yml``, which is
    gitignored and per-machine — a deployment without the key must behave the
    same as one that set it to false. Mirrors ``get_safe_grounding_include_
    siblings``: callers with no config in hand pass nothing and still get a
    working answer.
    """
    if processing_config is None:
        try:
            from .config import FDDConfig
            processing_config = FDDConfig().get_processing_config()
        except Exception:
            return False
    block = (processing_config or {}).get("claim_contracts")
    if isinstance(block, bool):
        return block
    if not isinstance(block, dict):
        return False
    return bool(block.get("enabled", False))


# ---------------------------------------------------------------------------
# guidance rendering (the facts the model was actually handed)
# ---------------------------------------------------------------------------

def _rendered_guidance(mapping_key: str, df: Any, language: str,
                       peer_context: Optional[Dict[str, Any]],
                       thresholds: Optional[Dict[str, Any]]) -> Dict[str, str]:
    """(variance, composition) guidance exactly as the Generator prompt got it."""
    from .prompts import PromptEngine  # local: keeps ai/__init__ import order untouched

    try:
        variance = PromptEngine._variance_analysis_guidance(
            df, language, peer_context=peer_context, mapping_key=mapping_key,
            thresholds=thresholds,
        ) or ""
    except Exception:
        variance = ""
    try:
        composition = PromptEngine._composition_guidance(df, language) or ""
    except Exception:
        composition = ""
    return {"variance": str(variance), "composition": str(composition)}


def _has(text: str, markers: Iterable[str]) -> bool:
    return any(m in text for m in markers)


# ---------------------------------------------------------------------------
# period labels — the column label vs. what a sentence actually writes
# ---------------------------------------------------------------------------

_MONTHS = ["January", "February", "March", "April", "May", "June",
           "July", "August", "September", "October", "November", "December"]

_ISO_DATE_RE = re.compile(r"^(\d{4})-(\d{2})-(\d{2})$")


def _period_surface_forms(label: str) -> List[str]:
    """Every way this ISO column label plausibly appears in the commentary.

    The prompt states periods as raw ``YYYY-MM-DD`` column labels; the shipped
    sentence never does. Measured on a real Eng run it writes "31 January
    2026", "FY25" and "1M26" for the same three columns, so a contract that
    only looked for the raw label would report every account as failing.
    Each form here is a specific string — no bare year, which would match any
    four-digit amount.
    """
    label = str(label or "").strip()
    m = _ISO_DATE_RE.match(label)
    if not m:
        return [label] if label else []
    year, month, day = int(m.group(1)), int(m.group(2)), int(m.group(3))
    name = _MONTHS[month - 1]
    forms = [
        label,
        f"{day} {name} {year}", f"{day} {name[:3]} {year}",
        f"{name} {day}, {year}", f"{name} {year}", f"{name[:3]} {year}",
        f"{name[:3]}-{year % 100:02d}", f"{name[:3]}{year % 100:02d}",
        f"{year}年{month}月{day}日", f"{year}年{month}月",
        f"{year}年{month:02d}月{day:02d}日", f"{year}年{month:02d}月",
    ]
    if (month, day) == (12, 31):
        # A full year, written the way the deck writes it.
        forms += [f"FY{year % 100:02d}", f"{year}年度", f"{year}全年"]
    else:
        # Partial tail period: the deck's own annualisation label, e.g. 1M26.
        forms += [f"{month}M{year % 100:02d}", f"{month}M{year}"]
    return forms


def _mentions_period(text: str, label: str) -> bool:
    low = text.lower()
    return any(form.lower() in low for form in _period_surface_forms(label) if form)


# ---------------------------------------------------------------------------
# parsing the guidance heads
# ---------------------------------------------------------------------------

# Anchored on phrasing that appears ONLY in the head sentence of each block.
# The "(Separately, this account increased about N% between A and B)" trailer
# also carries a percentage and two dates, and matching it would grade the
# account against a movement the guidance explicitly told the model to mention
# only in passing.
_NIL_EN = re.compile(r"The latest period \(([^)]+)\) is nil\..*?with the (\S+) balance")
_NIL_CN = re.compile(r"最新一期（([^）]+)）余额为零。.*?随后才补充([^的]+)的余额")
_MAT_EN = re.compile(
    r"This account's total moved from [-\d,\.]+ at (\S+) to [-\d,\.]+ at ([^,]+), "
    r"an (?:increase|decrease) of about (\d+)%")
_MAT_CN = re.compile(r"本科目合计由(\S+?)的-?[\d,]+变动至(\S+?)的-?[\d,]+，(?:增长|下降)约(\d+)%")
_FLIP_EN = re.compile(r"This account moved from [-\d,\.]+ at (\S+) to [-\d,\.]+ at ([^,]+), ?crossing")
_FLIP_CN = re.compile(r"本科目由(\S+?)的-?[\d,]+转为(\S+?)的-?[\d,]+")
_REM_EN = re.compile(r"the remaining \d+ come to \*\*([\d,\.]+) ([^*]*?)\*\*")
_REM_CN = re.compile(r"剩下的\d+项合计为\*\*([\d,\.]+)([^*]*?)\*\*")


def _parse_material(variance: str) -> Dict[str, Any]:
    for pattern in (_MAT_EN, _MAT_CN):
        m = pattern.search(variance)
        if m:
            return {"prev_period": m.group(1).strip(" .,，"),
                    "curr_period": m.group(2).strip(" .,，"),
                    "movement_pct": m.group(3)}
    for pattern in (_FLIP_EN, _FLIP_CN):
        m = pattern.search(variance)
        if m:
            # A sign flip carries no percentage by design: a raw % across zero
            # is nonsense (see _variance_analysis_guidance._pair), so the only
            # thing to check for is that both periods are named.
            return {"prev_period": m.group(1).strip(" .,，"),
                    "curr_period": m.group(2).strip(" .,，"),
                    "movement_pct": ""}
    return {}


def _parse_nil(variance: str) -> Dict[str, Any]:
    for pattern in (_NIL_EN, _NIL_CN):
        m = pattern.search(variance)
        if m:
            return {"nil_period": m.group(1).strip(" .,，"),
                    "prior_period": m.group(2).strip(" .,，")}
    return {}


def _parse_remainder(composition: str) -> Dict[str, Any]:
    for pattern in (_REM_EN, _REM_CN):
        m = pattern.search(composition)
        if m:
            return {"residual_amount": m.group(1).strip(),
                    "residual_unit": m.group(2).strip()}
    return {}


# ---------------------------------------------------------------------------
# the text-side needles
# ---------------------------------------------------------------------------

# Both branches the AR prompt allows — stating the ageing, or stating plainly
# that none was provided — contain an ageing word. So the presence of the word
# is exactly "the dimension was addressed", and its absence is exactly the
# silent omission the rule exists to stop. Deliberately no separate
# "not provided" list: "not provided" on its own says nothing about ageing.
_AGEING_NEEDLES = ("ageing", "aging", "aged", "账龄", "帳齡", "帐龄")

# A nil opening in either language. English needles are matched with word
# boundaries so "nil" cannot come out of another word.
_NIL_NEEDLES_CN = ("无余额", "未发生", "余额为零", "无结余", "无该等余额", "无此项余额",
                   "无余额结存", "为零", "無餘額")
_NIL_NEEDLES_EN = (r"\bno balance\b", r"\bnil\b", r"\bzero balance\b",
                   r"\bno outstanding balance\b", r"\bthere was no\b", r"\bwas no balance\b")

_RESIDUAL_CLAUSE = ("其余", "剩余", "剩下", "餘下", "remaining", "the rest", "the remainder")

_AR_KEY_NEEDLES = ("accounts receivable", "trade receivable", "trade receivables",
                   "应收账款", "應收賬款", "应收帐款")


def _is_ar_family(mapping_key: str, df: Any) -> bool:
    """Conservative AR-family test.

    Token-exact on the short key ("AR") plus explicit phrases, because a bare
    substring test for "ar" also hits "Taxes and Surcharges". The first column
    header is checked too: the CLI's dfs are keyed by mapping key ("AR") while
    the header carries the full account name ("Accounts receivable"), and only
    one of the two is present depending on how the frame was built.
    """
    key = str(mapping_key or "").strip()
    tokens = {t for t in re.split(r"[^0-9A-Za-z一-鿿]+", key.lower()) if t}
    if "ar" in tokens:
        return True
    haystacks = [key.lower()]
    try:
        first_col = str(list(df.columns)[0]).lower()
        haystacks.append(first_col)
    except Exception:
        pass
    return any(needle in hay for hay in haystacks for needle in _AR_KEY_NEEDLES)


def _opening_sentence(text: str) -> str:
    """The first sentence, in either language.

    Kept deliberately short: the nil rule is about what the reader sees FIRST.
    A "no balance" clause three sentences down is the exact failure the
    guidance was written for (a real run opened on the prior period's non-nil
    figure and mentioned the nil period only afterwards), so finding it
    anywhere in the bullet must not count as a pass.
    """
    body = str(text or "").strip()
    if not body:
        return ""
    cn = re.split(r"[。；！？\n]", body)[0]
    en = re.split(r"(?<=[a-z0-9\)\"])\.\s", body)[0]
    return min((cn, en), key=len)


def _has_nil_phrase(segment: str) -> bool:
    if any(n in segment for n in _NIL_NEEDLES_CN):
        return True
    low = segment.lower()
    return any(re.search(p, low) for p in _NIL_NEEDLES_EN)


def _mentions_pct(text: str, pct: str) -> bool:
    """The supplied integer percentage, with or without a trailing .0 decimal."""
    if not pct:
        return False
    return bool(re.search(rf"(?<![\d.]){re.escape(pct)}(\.\d+)?\s*%", text))


# ---------------------------------------------------------------------------
# building and checking
# ---------------------------------------------------------------------------

def build_claim_contract(
    mapping_key: str,
    df: Any,
    language: str,
    peer_context: Optional[Dict[str, Any]] = None,
    *,
    thresholds: Optional[Dict[str, Any]] = None,
) -> List[Dict[str, Any]]:
    """The post-conditions this account's own prompt committed it to.

    Returns a list of ``{claim_id, kind, detector, patch_hint, facts}``.
    Empty when the prompt injected none of the four instructions — which is
    the common case and is not a finding.
    """
    if df is None:
        return []
    language = "Chi" if str(language) in ("Chn", "Chi", "chinese", "Chinese") else "Eng"
    guidance = _rendered_guidance(mapping_key, df, language, peer_context, thresholds)
    variance, composition = guidance["variance"], guidance["composition"]
    claims: List[Dict[str, Any]] = []

    if _has(variance, _NIL_MARKERS):
        facts = _parse_nil(variance)
        claims.append({
            "claim_id": f"{mapping_key}:nil_opening",
            "kind": "required_opening",
            "detector": "nil_opening",
            "patch_hint": (
                "Open on {nil} and say there is no balance; move the {prior} figure after it."
                .format(nil=facts.get("nil_period", "the latest period"),
                        prior=facts.get("prior_period", "prior period"))
            ),
            "facts": facts,
        })

    if _is_ar_family(mapping_key, df):
        claims.append({
            "claim_id": f"{mapping_key}:ar_ageing",
            "kind": "mandatory_dimension",
            "detector": "ar_ageing",
            "patch_hint": (
                "Add one clause on ageing/collectability — the fact when the data supports "
                "it, otherwise 'no ageing breakdown was provided'. Never invent a bucket."
            ),
            "facts": {"account_family": "AR"},
        })

    if _has(variance, _MATERIAL_MARKERS):
        facts = _parse_material(variance)
        claims.append({
            "claim_id": f"{mapping_key}:material_movement_explained",
            "kind": "computed_fact_used",
            "detector": "material_movement_explained",
            "patch_hint": (
                "State the movement the prompt computed ({pct}% between {a} and {b}) rather "
                "than only the closing figure."
                .format(pct=facts.get("movement_pct") or "the stated",
                        a=facts.get("prev_period", "?"), b=facts.get("curr_period", "?"))
            ),
            "facts": facts,
        })

    if _has(composition, _REMAINDER_MARKERS):
        facts = _parse_remainder(composition)
        claims.append({
            "claim_id": f"{mapping_key}:residual_disclosed",
            "kind": "computed_fact_used",
            "detector": "residual_disclosed",
            "patch_hint": (
                "Close the enumeration with the residual the prompt computed "
                "({amt} {unit}) instead of trailing off."
                .format(amt=facts.get("residual_amount", "?"),
                        unit=facts.get("residual_unit", ""))
            ),
            "facts": facts,
        })

    return claims


def check_claim_contract(contract: List[Dict[str, Any]], text: str) -> List[Dict[str, Any]]:
    """Grade a contract against the shipped text.

    Returns the contract's claims, each with ``result`` (pass/fail) and
    ``observed`` (what the check actually looked at, so a false positive can be
    read off the report without re-running anything).
    """
    body = str(text or "")
    checked: List[Dict[str, Any]] = []
    for claim in contract or []:
        detector = claim.get("detector")
        facts = claim.get("facts") or {}
        observed: Dict[str, Any] = {}

        if detector == "nil_opening":
            opening = _opening_sentence(body)
            observed["opening"] = opening[:160]
            observed["nil_phrase_anywhere"] = _has_nil_phrase(body)
            observed["opening_names_nil_period"] = _mentions_period(
                opening, facts.get("nil_period", ""))
            ok = _has_nil_phrase(opening)

        elif detector == "ar_ageing":
            hits = [n for n in _AGEING_NEEDLES if n in body.lower() or n in body]
            observed["ageing_words"] = hits
            ok = bool(hits)

        elif detector == "material_movement_explained":
            pct = str(facts.get("movement_pct") or "")
            prev_seen = _mentions_period(body, facts.get("prev_period", ""))
            curr_seen = _mentions_period(body, facts.get("curr_period", ""))
            pct_seen = _mentions_pct(body, pct)
            observed.update({"pct_stated": pct_seen,
                             "prev_period_named": prev_seen,
                             "curr_period_named": curr_seen})
            ok = pct_seen or (prev_seen and curr_seen)

        elif detector == "residual_disclosed":
            amount = str(facts.get("residual_amount") or "")
            amount_seen = bool(amount) and amount in body
            clause_seen = any(c in body or c in body.lower() for c in _RESIDUAL_CLAUSE)
            observed.update({"amount_stated": amount_seen, "residual_clause": clause_seen})
            ok = amount_seen or clause_seen

        else:  # unknown detector: never invent a verdict
            continue

        record = dict(claim)
        record["result"] = PASS if ok else FAIL
        record["observed"] = observed
        checked.append(record)
    return checked


def claim_contract_verdicts(checked: List[Dict[str, Any]]) -> Dict[str, str]:
    """The plain ``{claim_id: "pass"|"fail"}`` map stored on ``results``."""
    return {str(c["claim_id"]): str(c["result"]) for c in checked or []}


def attach_claim_contracts(
    results: Dict[str, Any],
    dfs: Optional[Dict[str, Any]],
    language: str,
    *,
    peer_context: Optional[Dict[str, Any]] = None,
    thresholds: Optional[Dict[str, Any]] = None,
    text_of: Optional[Any] = None,
) -> Dict[str, List[Dict[str, Any]]]:
    """Build, grade and file a contract for every account in ``results``.

    Call site: after final fallbacks are set, i.e. once ``results[key]`` holds
    the text that will actually ship. Unconditional when the feature is on —
    NOT alongside ``_apply_deterministic_verification``, which runs only in
    selective mode and before the feedback loop, and so would grade a string
    that a retry then replaces.

    Writes ``results[key]["claim_contract"]``. Returns the full per-account
    checked records for reporting; those are not stored.
    """
    if text_of is None:
        from ..financial_common import get_pipeline_result_text
        text_of = get_pipeline_result_text

    if peer_context is None and dfs:
        try:
            from .pipeline import _build_peer_context
            peer_context = _build_peer_context(dfs)
        except Exception:
            peer_context = None

    checked_by_account: Dict[str, List[Dict[str, Any]]] = {}
    for key, result in (results or {}).items():
        if str(key).startswith("__") or not isinstance(result, dict):
            continue  # run-level sentinel (__run_health__, __BS_summary__), not an account
        df = (dfs or {}).get(key)
        if df is None:
            continue
        try:
            contract = build_claim_contract(
                key, df, language, peer_context, thresholds=thresholds)
        except Exception:
            continue
        if not contract:
            continue
        checked = check_claim_contract(contract, text_of(result) or "")
        if not checked:
            continue
        result[CLAIM_CONTRACT_KEY] = claim_contract_verdicts(checked)
        checked_by_account[key] = checked
    return checked_by_account


def summarise_claim_contracts(
    checked_by_account: Dict[str, List[Dict[str, Any]]],
) -> Dict[str, Dict[str, Any]]:
    """Per-detector ``{fired, passed, failed, failing_accounts}``.

    This tally IS the deliverable of M5 v1: a detector whose failure rate is
    implausibly high is a detector that is wrong, not commentary that is.
    """
    tally: Dict[str, Dict[str, Any]] = {}
    for account, checked in (checked_by_account or {}).items():
        for claim in checked:
            row = tally.setdefault(str(claim.get("detector")), {
                "fired": 0, "passed": 0, "failed": 0, "failing_accounts": []})
            row["fired"] += 1
            if claim.get("result") == PASS:
                row["passed"] += 1
            else:
                row["failed"] += 1
                row["failing_accounts"].append(account)
    return tally
# --- end ai/contracts.py ---

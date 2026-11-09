"""Claims: what a sentence says, what it cites, and whether the two agree.

`SourceIndex.matches` answers "is there a number near this one somewhere in the
pool". That is useful for diagnosis and it is how a wrong figure that happens to
equal another account's balance passed as "data-backed". A claim is checked the
other way round: the author names the facts a sentence rests on, Python computes
every derived figure, and the check is whether the TEXT says what the CITED facts
say -- same period, same line, same basis, same sign and unit, rounded the way it
is printed. A number that is right but cited from the wrong place fails.

The author is a model or a person editing a paragraph; either produces claims:

    {"account": "应收账款", "kind": "movement", "periods": ["2024-12-31", "2025-12-31"],
     "text": "应收账款较上年末增加120.0万元，增幅10.0%",
     "cites": ["F7", "F3", "D2"],              # facts, or derived facts Python computed
     "cause_cites": [], "inference": false}

Verdicts keep the questions apart, because folding them is how "the number was
found" came to mean "the claim was verified":

    citations   every ref exists, belongs to this account and this source version
    amounts     every amount in the text is accounted for by a cited fact or by a
                figure Python derived from cited facts
    binding     the right period, line and basis: not just the right number
    arithmetic  percentages, sums, annualisation, zero denominators
    causal      a stated cause rests on a cited note, or is marked as inference
    coverage    nothing numeric in the text is left unaccounted for

Severity: `failed` stops a deliverable; `info` is disclosed (an inference marked
as such is what the deck renders in orange).
"""
from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from typing import Any, Dict, Iterable, List, Optional, Tuple

from .evidence_contract import (
    BASIS_ADJUSTED, BASIS_FINANCIALS, BASIS_RAW, ROLE_COMPONENT, ROLE_RELATED, ROLE_SUPPORT, ROLE_TOTAL, EvidenceContract,
)
from .validator import display_half_ulp, extract_amount_spans

# ---------------------------------------------------------------------------
# Derived facts: Python computes the numbers
# ---------------------------------------------------------------------------

BASIS_DERIVED = "derived"
KIND_MOVEMENT = "movement"
KIND_SHARE = "share"
KIND_SUM = "sum"
KIND_DIFFERENCE = "difference"
KIND_ANNUALIZED = "annualized"

LINEAR_ANNUALIZATION_NOTE = "linear annualisation of a part-year figure; not a forecast"


def _register(contract: EvidenceContract, derived: Dict[str, Any]) -> Dict[str, Any]:
    """Idempotent by id; refs D1... in registration order."""
    for existing in contract.derived:
        if existing["id"] == derived["id"]:
            return existing
    derived["ref"] = f"D{len(contract.derived) + 1}"
    contract.derived.append(derived)
    return derived


def _base(contract: EvidenceContract, kind: str, suffix: str, period: str, periods: List[str]) -> Dict[str, Any]:
    return {"id": f"{contract.account}|{suffix}|derived|{kind}", "account": contract.account, "kind": kind,
            "basis": BASIS_DERIVED, "role": "derived", "period": period, "periods": periods,
            "unit": "yuan", "valid": True, "invalid_reason": "", "inputs": [], "conditions": []}


def _total_fact(contract: EvidenceContract, period: str, basis: str = BASIS_ADJUSTED) -> Optional[Dict[str, Any]]:
    for item in contract.facts:
        if item["period"] == period and item["role"] == ROLE_TOTAL and item["basis"] == basis:
            return item
    return None


def derive_movement(contract: EvidenceContract, p_from: str, p_to: str) -> Dict[str, Any]:
    """Change in the account total from one period to a later one, and as a share of the earlier.

    Always returned, valid or not: a movement the permissions forbid is still a
    record, so a sentence that cites it fails for a stated reason instead of
    for a missing ref.
    """
    d = _base(contract, KIND_MOVEMENT, f"{p_from}->{p_to}", p_to, [p_from, p_to])
    a, b = _total_fact(contract, p_from), _total_fact(contract, p_to)
    d["formula"] = f"total({p_to}) - total({p_from})"
    if a is None or b is None:
        d.update(valid=False, invalid_reason="a period has no total", value=None, pct=None)
        return _register(contract, d)
    d["inputs"] = [a["id"], b["id"]]
    d["value"] = b["value"] - a["value"]
    d["base"], d["end"] = a["value"], b["value"]
    perm = contract.permissions.get(p_to) or {}
    prev = contract.permissions.get(p_from) or {}
    if not (perm.get("describe_total") and prev.get("describe_total")):
        d.update(valid=False, invalid_reason="a total is not describable for these periods (tab and Financials disagree)")
    elif not perm.get("compute_movement"):
        d.update(valid=False, invalid_reason=perm.get("movement_block") or "periods are not comparable")
    if abs(a["value"]) > 0:
        d["pct"] = d["value"] / abs(a["value"])
        d["zero_denominator"] = False
    else:
        d["pct"] = None
        d["zero_denominator"] = True
    d["direction"] = "up" if d["value"] > 0 else ("down" if d["value"] < 0 else "flat")
    return _register(contract, d)


def derive_share(contract: EvidenceContract, line: str, period: str) -> Dict[str, Any]:
    """One detail line as a share of the account total -- only where the detail
    is allowed to explain the total."""
    d = _base(contract, KIND_SHARE, f"{period}:{line}", period, [period])
    d["line"] = line
    comp = next((f for f in contract.facts if f["period"] == period and f["role"] == ROLE_COMPONENT and f["line"] == line), None)
    total = _total_fact(contract, period)
    d["formula"] = f"{line}({period}) / total({period})"
    if comp is None or total is None or abs(total["value"]) < 1e-9:
        d.update(valid=False, invalid_reason="no such line, no total, or a zero total", value=None, pct=None)
        return _register(contract, d)
    d["inputs"] = [comp["id"], total["id"]]
    d["value"] = comp["value"]
    d["pct"] = comp["value"] / total["value"]
    basis = (contract.permissions.get(period) or {}).get("detail_basis")
    if basis == "untied" or basis == "none":
        d.update(valid=False, invalid_reason="the tab's detail does not add up to the total in this period")
    elif basis == "ambiguous":
        d["conditions"].append("the detail's nesting is unknown; this is a ratio of one line to the total, not a claim that the lines add up")
    return _register(contract, d)


def derive_sum(contract: EvidenceContract, refs: List[str]) -> Dict[str, Any]:
    """The sum of facts the caller NAMES. Two lines that happen to add up to the
    row between them are not a group; only an explicit list is."""
    facts = [contract.fact(r) for r in refs]
    periods = sorted({f["period"] for f in facts if f})
    key = ",".join(sorted(f["id"] for f in facts if f)) or ",".join(sorted(refs))
    d = _base(contract, KIND_SUM, hashed(key), periods[-1] if periods else "", periods)
    d["formula"] = " + ".join(refs)
    if not facts or any(f is None for f in facts) or len(periods) != 1:
        d.update(valid=False, invalid_reason="every summand must exist and share one period", value=None)
        return _register(contract, d)
    d["inputs"] = [f["id"] for f in facts]
    d["value"] = sum(f["value"] for f in facts)
    return _register(contract, d)


def derive_difference(contract: EvidenceContract, period: str) -> Dict[str, Any]:
    """Tab total minus Financials total -- the size of a disagreement, which is
    describable even when neither figure may be asserted."""
    d = _base(contract, KIND_DIFFERENCE, period, period, [period])
    raw, fin = _total_fact(contract, period, BASIS_RAW), _total_fact(contract, period, BASIS_FINANCIALS)
    d["formula"] = f"tab total({period}) - Financials({period})"
    if raw is None or fin is None:
        d.update(valid=False, invalid_reason="the tab or Financials has no total for this period", value=None)
        return _register(contract, d)
    d["inputs"] = [raw["id"], fin["id"]]
    d["value"] = raw["value"] - fin["value"]
    return _register(contract, d)


def derive_annualized(contract: EvidenceContract, period: str) -> Dict[str, Any]:
    """A part-year flow scaled to twelve months, labelled as exactly that."""
    d = _base(contract, KIND_ANNUALIZED, period, period, [period])
    total = _total_fact(contract, period)
    months = ((contract.period_meta.get(period) or {}).get("months"))
    d["formula"] = f"total({period}) / {months} months x 12"
    d["note"] = LINEAR_ANNUALIZATION_NOTE
    if total is None or not months or months >= 12:
        d.update(valid=False, invalid_reason="not a part-year figure", value=None)
        return _register(contract, d)
    d["inputs"] = [total["id"]]
    d["value"] = total["value"] / months * 12.0
    return _register(contract, d)


def hashed(text: str) -> str:
    import hashlib
    return hashlib.sha256(text.encode("utf-8")).hexdigest()[:10]


def standard_derived(contract: EvidenceContract) -> int:
    """The movements an author is most likely to quote: each consecutive pair, and
    the difference where tab and Financials disagree. Returns how many records exist."""
    for prev, nxt in zip(contract.periods, contract.periods[1:]):
        derive_movement(contract, prev, nxt)
    for item in contract.checks:
        if item["kind"] == "total_vs_financials" and item["status"] == "mismatch":
            derive_difference(contract, item["period"])
    return len(contract.derived)


# ---------------------------------------------------------------------------
# What the text says
# ---------------------------------------------------------------------------

_MONTHS = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}
_EN_MONTH = r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?"
_PERIOD_PATTERNS: List[Tuple[str, "re.Pattern[str]"]] = [
    ("date", re.compile(r"(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*[日号]")),
    ("date_en", re.compile(r"(\d{1,2})\s+" + _EN_MONTH + r",?\s+(\d{4})", re.I)),
    # "2026年1-3月" / "2026年1至3月": a part-year flow, named by the month it ends (before "month", which would read "2026年1")
    ("month_range", re.compile(r"(\d{4})\s*年\s*\d{1,2}\s*[-–—~至到]\s*(\d{1,2})\s*月(?:份)?")),
    ("month", re.compile(r"(\d{4})\s*年\s*(\d{1,2})\s*月(?!\s*\d)(?:末|底|份)?")),
    ("month_en", re.compile(_EN_MONTH + r"\s+(\d{4})", re.I)),
    ("year_end", re.compile(r"(\d{4})\s*年\s*(?:末|底|年末|年底)")),
    ("year", re.compile(r"(\d{4})\s*年(?![\d月])")),
    ("year_en", re.compile(r"(?:\bFY\s?|\b(?:in|during|for|year|since|end of|ended)\s+)(20\d{2})\b", re.I)),
]


def period_mentions(text: str) -> List[Dict[str, Any]]:
    """Periods the text names, as {kind, year, month, day, text, start, end}. Chinese and English forms."""
    work = str(text or "")
    found: List[Dict[str, Any]] = []
    for kind, pattern in _PERIOD_PATTERNS:
        def take(m: "re.Match[str]") -> str:
            g = m.groups()
            where = {"text": m.group(0), "start": m.start(), "end": m.end()}
            if kind == "date":
                found.append({"kind": "date", "year": int(g[0]), "month": int(g[1]), "day": int(g[2]), **where})
            elif kind == "date_en":
                found.append({"kind": "date", "year": int(g[2]), "month": _MONTHS[g[1][:3].lower()], "day": int(g[0]), **where})
            elif kind in ("month", "month_range"):
                found.append({"kind": "month", "year": int(g[0]), "month": int(g[1]), **where})
            elif kind == "month_en":
                found.append({"kind": "month", "year": int(g[1]), "month": _MONTHS[g[0][:3].lower()], **where})
            elif kind == "year_end":
                found.append({"kind": "date", "year": int(g[0]), "month": 12, "day": 31, **where})
            else:
                found.append({"kind": "year", "year": int(g[0]), **where})
            return " " * len(m.group(0))
        work = pattern.sub(take, work)
    found.sort(key=lambda m: m["start"])
    return found


def attached_period(text: str, start: int, end: int, mentions: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    """The period an amount at [start, end) is written under: the nearest mention before it in the
    same sentence, else the nearest after it ("余额130万元（2025年12月31日）"). None if the sentence names none."""
    before = [m for m in mentions if m["end"] <= start and "。" not in text[m["end"]:start]]
    if before:
        return max(before, key=lambda m: m["end"])
    after = [m for m in mentions if m["start"] >= end and "。" not in text[end:m["start"]]]
    return min(after, key=lambda m: m["start"]) if after else None


def mention_matches(mention: Dict[str, Any], period: str) -> bool:
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", str(period))
    if not m:
        return False
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if mention["kind"] == "date":
        return (mention["year"], mention["month"], mention["day"]) == (y, mo, d)
    if mention["kind"] == "month":
        return (mention["year"], mention["month"]) == (y, mo)
    return mention["year"] == y


_QIAN = re.compile(r"(\d[\d,]*(?:\.\d+)?)\s*千元")
_PERCENT = re.compile(r"([-+−－]?)(\d[\d,]*(?:\.\d+)?)\s*(?:%|％|个百分点|百分点)")


def text_amounts(text: str) -> List[Dict[str, Any]]:
    """Money amounts as written: value in yuan, the written form, and its span."""
    work = str(text or "").replace("，", ",")
    spans: List[Dict[str, Any]] = []
    # 千元 is not a form validator.extract_amount_spans reads (the tables use it, prose rarely does)
    def qian(m: "re.Match[str]") -> str:
        spans.append({"value": float(m.group(1).replace(",", "")) * 1000.0, "written": m.group(0),
                      "start": m.start(), "end": m.end()})
        return " " * len(m.group(0))
    work = _QIAN.sub(qian, work)
    for value, start, end in extract_amount_spans(work):
        spans.append({"value": float(value), "written": work[start:end], "start": start, "end": end})
    spans.sort(key=lambda s: s["start"])
    plain = str(text or "").replace("，", ",")
    for s in spans:
        s["tolerance"] = max(display_half_ulp(s["written"]), 0.5)
        s["sign"] = _written_sign(plain, s["start"])
    return spans


def _written_sign(text: str, start: int) -> str:
    """'-' / '+' / '': the sign written right before an amount ("-130万元", "负130万元", "+70万元"). A dash that follows a
    figure is a range or a date ("100-130万元", "2025-12-31"), not a minus."""
    i = start - 1
    while i >= 0 and text[i] == " ":
        i -= 1
    if i < 0:
        return ""
    if text[i] == "负":
        return "-"
    if text[i] not in "-−－+＋":
        return ""
    j = i - 1
    while j >= 0 and text[j] == " ":
        j -= 1
    if j >= 0 and (text[j].isdigit() or text[j] in "%％"):
        return ""
    return "+" if text[i] in "+＋" else "-"


def text_percents(text: str) -> List[Dict[str, Any]]:
    """Percentages as written, with their sign: '-' / '+' / '' (none written). A dash that follows a
    figure is a range ("30%-40%"), not a minus."""
    work = str(text or "").replace("，", ",")
    out = []
    for m in _PERCENT.finditer(work):
        sign = {"-": "-", "−": "-", "－": "-", "+": "+"}.get(m.group(1), "")
        before = work[:m.start()].rstrip()
        if sign and before and before[-1] in "0123456789%％":
            sign = ""
        written = m.group(2).replace(",", "")
        decimals = len(written.split(".")[1]) if "." in written else 0
        value = float(written) * (-1.0 if sign == "-" else 1.0)
        start = m.start() if sign else m.start(2)
        out.append({"value": value, "sign": sign, "written": work[start:m.end()], "start": start, "end": m.end(),
                    "tolerance": 0.5 * (10 ** -decimals), "points": "百分点" in m.group(0)})
    return out


_UP = ("增加", "增长", "上升", "提高", "上涨", "增至", "升至", "新增", "扩大", "增幅", "increase", "rose", "grew", "higher")
_DOWN = ("减少", "下降", "降低", "下滑", "减至", "降至", "缩减", "收窄", "降幅", "decrease", "fell", "declined", "lower")
_CHANGE_CUES = ("变动", "变化", "change")
_COMPOSITION_CUES = ("其中", "包括", "主要为", "主要是", "主要来自", "分别", "构成", "组成", "占比", "占总", "占该", "比重", "份额",
                     "of which", "including", "composed of", "share of")
_NEUTRAL_CHANGE = ("变动", "变化", "波动", "幅度", "changed by", "change of", "variation")
_NIL = ("无余额", "余额为零", "余额为0", "未发生", "首次形成", "首次产生", "首次出现", "为零", "nil", "first arose", "no balance")
# Why-language and forward-looking language. validator._CAUSAL_RE also lists 年化 and
# annualis -- right for flagging judgement in the legacy pipeline, wrong here, where an
# annualised figure is arithmetic with its own rules and must not be asked for a "cause".
_CAUSE_RE = re.compile(
    r"driven by|attributed to|due to|owing to|as a result|because|thanks to|resulted from|caused by|"
    r"由于|因为|主要系|主要由于|所致|导致|归因于|得益于|受[^，。；]{1,12}影响|系因|源于|源自|"
    r"以致|致使|使得|鉴于|(?<![原起])因(?![素此子])|since |on account of|result of", re.IGNORECASE)
_FORWARD_RE = re.compile(r"预计|预期|有望|expected to|is expected|forecast", re.IGNORECASE)
_ANNUAL_WORD = ("年化", "annualis", "annualiz")
_FORECAST_WORD = ("预计全年", "全年预测", "全年预计", "预计全年度", "full-year forecast", "expected full-year", "forecast")
_ANNUAL_ANY = _ANNUAL_WORD + _FORECAST_WORD + ("全年", "full year", "full-year")
_GENERIC = {"金额", "余额", "增加", "减少", "主要", "所致", "导致", "由于", "因为", "本期", "期末", "上升", "下降", "原因",
            "较上", "上年", "增长", "变动", "系因", "主要系", "其中", "包括", "合计", "公司", "款项", "万元", "亿元"}


_CURRENCY_WORDS = (("USD", re.compile(r"美元|美金|USD|US\$", re.I)), ("HKD", re.compile(r"港元|港币|HKD", re.I)),
                   ("EUR", re.compile(r"欧元|EUR", re.I)))


def _currency_at(text: str, start: int, end: int) -> str:
    """The currency an amount at [start, end) is written in: a currency word inside it, right after it, or right before it."""
    window = text[max(0, start - 4):min(len(text), end + 4)]
    for code, pattern in _CURRENCY_WORDS:
        if pattern.search(window):
            return code
    return "CNY"


_NOTE_TAG = re.compile(r"「(?:(\d{1,2})M(\d{2})|(\d{4}))」")


def note_period(note: Dict[str, Any], periods: List[str]) -> Optional[str]:
    """The period a note's figures belong to: its own, or its tag ("「2025」", "「3M26」") when exactly one period matches."""
    if note.get("period"):
        return str(note["period"])
    m = _NOTE_TAG.search(str(note.get("text") or ""))
    if not m:
        return None
    if m.group(3):
        hits = [p for p in periods if p.startswith(m.group(3) + "-12")] or [p for p in periods if p.startswith(m.group(3))]
    else:
        month, year = int(m.group(1)), 2000 + int(m.group(2))
        hits = [p for p in periods if p.startswith(f"{year}-{month:02d}")]
    return hits[0] if len(hits) == 1 else None


def note_figures(note: Dict[str, Any], periods: List[str]) -> Dict[str, Any]:
    """What a note's text states with a unit: amounts (value, currency), percentages, dates. Bare numbers are not read (a
    remark's cell values have no unit here)."""
    text = str(note.get("text") or "")
    own = note_period(note, periods)
    # a written sign is kept: "ADJ#4：-133.2万元" is a reduction (dropping it failed every adjustment written with its sign)
    amounts = [{"value": sp["value"] * (-1.0 if sp.get("sign") == "-" else 1.0), "currency": _currency_at(text, sp["start"], sp["end"]),
                "written": sp["written"], "period": own} for sp in text_amounts(text)]
    # a remark's cell values carry their period ("2026-03-31: 4342452.0"): yuan, like the facts, and bound to that period
    for m in _REMARK_VALUE.finditer(text):
        if m.group(1) in periods:
            amounts.append({"value": float(m.group(3)), "currency": "CNY", "written": m.group(0), "period": m.group(1),
                            "annualised": bool(m.group(2))})
    signs = {w["sign"] for w in direction_words(text)} | {1 if sp.get("sign") == "+" else -1 for sp in text_amounts(text) if sp.get("sign")}
    return {"amounts": amounts, "percents": [abs(sp["value"]) for sp in text_percents(text) if not sp["points"]],
            "dates": [m for m in period_mentions(text) if m["kind"] in ("date", "month")], "period": own,
            "years": {int(y) for y in re.findall(r"(?<!\d)(20\d{2})(?!\d)", text)}, "signs": signs}


_REMARK_VALUE = re.compile(r"(\d{4}-\d{2}-\d{2})( annualised)?:\s*(-?\d+(?:\.\d+)?)")


_INTERNAL_TERMS = re.compile(r"[（(]\s*[FDN]\d{1,4}(?:\s*[,，、]\s*[FDN]\d{1,4})*\s*[)）]|(?<![A-Za-z])[FDN]\d{1,4}(?![\d])|authority_adjusted|"
                             r"raw_schedule|detail_basis|explain_with_detail|describe_total|(?<![A-Za-z])financials口径")


_LINE_ROLES = (ROLE_COMPONENT, ROLE_SUPPORT)


def _bare(label: Any) -> str:
    return str(label or "").lstrip("*＊ ").strip()


def _has(text: str, words: Iterable[str]) -> bool:
    low = str(text or "").lower()
    return any(w.lower() in low for w in words)


def _bigrams(text: str) -> set:
    out = set()
    for run in re.findall(r"[一-鿿]+", str(text or "")):
        out.update(run[i:i + 2] for i in range(len(run) - 1))
    for word in re.findall(r"[A-Za-z]{4,}", str(text or "")):
        out.add(word.lower())
    return {b for b in out if b not in _GENERIC}


# ---------------------------------------------------------------------------
# Claims and verdicts
# ---------------------------------------------------------------------------

CLAIM_KINDS = ("amount", "movement", "composition", "share", "difference", "annualized", "cause", "descriptive")


@dataclass
class ClaimFinding:
    code: str
    dimension: str
    severity: str          # "failed" | "info"
    message: str

    def to_dict(self) -> Dict[str, str]:
        return dict(self.__dict__)


def _flag(value: Any) -> bool:
    """A declared flag is true only when it says so. bool("false") is True, which would let a model switch a
    waiver on by writing the word that means off."""
    if isinstance(value, (bool, int, float)):
        return bool(value) and value == 1
    return str(value or "").strip().lower() in ("true", "yes", "y", "1")


def normalize_claim(raw: Dict[str, Any], default_account: str = "", index: int = 0) -> Dict[str, Any]:
    """The claim as the verifier reads it. `kind`, `inference` and `cross_account` are declarations by whoever wrote
    the sentence. `kind` is kept for display and decides no check: which checks a sentence needs follows from what
    its text asserts and what it cites. `inference` waives exactly one thing (a cause with no note behind it, and a
    linear annualisation worded as a forecast); never a figure, period, unit or sign. `cross_account` cannot make
    another account's figure this account's."""
    kind = str(raw.get("kind") or "amount").strip().lower()
    return {
        "id": str(raw.get("id") or f"c{index + 1}"),
        "account": str(raw.get("account") or default_account),
        "kind": kind if kind in CLAIM_KINDS else "amount",
        "text": str(raw.get("text") or "").strip(),
        "periods": [str(p) for p in (raw.get("periods") or [])],
        "cites": [str(c) for c in (raw.get("cites") or [])],
        "cause_cites": [str(c) for c in (raw.get("cause_cites") or [])],
        "inference": _flag(raw.get("inference")),
        "cross_account": _flag(raw.get("cross_account")),
        "source_version": raw.get("source_version"),
    }


def _lookup(ref: str, contracts: Dict[str, EvidenceContract], account: str) -> Tuple[Optional[Dict[str, Any]], Optional[EvidenceContract]]:
    own = contracts.get(account)
    if own is not None:
        hit = own.fact(ref)
        if hit is not None:
            return hit, own
    for contract in contracts.values():
        hit = contract.fact(ref)
        if hit is not None:
            return hit, contract
    return None, None


def attached_label(text: str, start: int, labels: List[str], reach: int = 40) -> Optional[str]:
    """The line name an amount at `start` is attached to: the label ending nearest before it
    (the longest, where a short name sits inside a long one: 工程款 / 电力水力工程款), with no
    other figure in between and at most `reach` characters between the name and the figure. The
    name's own length does not count, and neither do blanks, so a period written between them
    ("Provision for bad and doubtful debts占比由2025年12月31日的88.9%") costs nothing.
    None when the amount is not written against a line."""
    before = text[:start]
    best: Optional[Tuple[int, int, str]] = None
    for label in labels:
        position = before.rfind(label)
        if position < 0 or len(before[position + len(label):].replace(" ", "")) > reach:
            continue
        candidate = (position + len(label), len(label), label)
        if best is None or candidate[:2] > best[:2]:
            best = candidate
    if best is None or re.search(r"\d", before[best[0]:]):
        return None
    return best[2]


def mask_mentions(text: str, mentions: List[Dict[str, Any]]) -> str:
    """The text with each period mention blanked (same width), so a line name and the figure after it
    stay adjacent when a period is written between them: 甲公司2025年末占比53.8%."""
    chars = list(text)
    for m in mentions:
        for i in range(m["start"], min(m["end"], len(chars))):
            chars[i] = " "
    return "".join(chars)


_SHARE_CUES = ("占比", "比重", "份额", "占总", "占该", "占", "比例", "构成", "share of", "% of")
_CHANGE_CUES = ("增幅", "增长", "增加", "上升", "涨幅", "下降", "降幅", "减少", "下滑", "变动", "同比", "环比", "变化", "降低",
                "growth", "increase", "decrease", "decline")


def percent_cue(text: str, start: int, window: int = 24) -> Optional[str]:
    """'share' or 'change': the kind of ratio the words just before a percentage say it is, or None.
    A direction word that leads into a level ("下降至53.8%") says where the ratio ended up, not that
    the figure is a rate of change, so it is not a cue."""
    before = text[max(0, start - window):start].lower()
    last, kind = -1, None
    for cues, name in ((_SHARE_CUES, "share"), (_CHANGE_CUES, "change")):
        for cue in cues:
            for m in re.finditer(re.escape(cue), before):
                if name == "change" and before[m.end():m.end() + 1] in ("至", "到", "达"):
                    continue
                if m.end() > last:
                    last, kind = m.end(), name
    return kind


def _consecutive_pairs(facts: List[Dict[str, Any]]) -> List[Tuple[Dict[str, Any], Dict[str, Any]]]:
    """(earlier, later) for cited facts of the same line and basis in different periods."""
    groups: Dict[Tuple[str, str, str], List[Dict[str, Any]]] = {}
    for f in facts:
        groups.setdefault((f["role"], f["line"], f["basis"]), []).append(f)
    pairs = []
    for group in groups.values():
        ordered = sorted(group, key=lambda g: g["period"])
        pairs.extend((a, b) for a, b in zip(ordered, ordered[1:]) if a["period"] != b["period"])
    return pairs


def names_of(candidate: Dict[str, Any]) -> str:
    return {"share": "share", "points": "difference of the two shares"}.get(candidate["kind"], "change")


def share_movements(derived: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """For each line, the change in its share between consecutive cited periods:
    {line, value (fraction of the total, signed), periods, ref}."""
    by_line: Dict[str, List[Dict[str, Any]]] = {}
    for d in derived:
        if d.get("kind") == KIND_SHARE and d.get("pct") is not None and d.get("line"):
            by_line.setdefault(d["line"], []).append(d)
    out = []
    for line, group in by_line.items():
        ordered = sorted(group, key=lambda g: g["period"])
        out.extend({"line": line, "value": b["pct"] - a["pct"], "periods": [a["period"], b["period"]],
                    "ref": f"points({a['ref']},{b['ref']})",
                    # a deduction line (a provision) has a negative share; written without a sign, "升至" is about its size
                    "magnitude": (abs(b["pct"]) - abs(a["pct"])) if a["pct"] < 0 and b["pct"] < 0 else b["pct"] - a["pct"]}
                   for a, b in zip(ordered, ordered[1:]) if a["period"] != b["period"])
    return out


def _fmt(value: Optional[float]) -> str:
    return "n/a" if value is None else f"{value:,.2f}"


_DIRECTION_WORDS = {**{w.lower(): 1 for w in _UP}, **{w.lower(): -1 for w in _DOWN}}
_DIRECTION_RE = re.compile("|".join(re.escape(w) for w in sorted(_DIRECTION_WORDS, key=len, reverse=True)), re.IGNORECASE)
_CLAUSE_BREAK = re.compile(r"[，；。;\n]|,(?!\d)")
_SHARE_SUBJECT = ("占比", "比重", "份额", "占总", "占该", "比例", "构成", "share")
_AMOUNT_SUBJECT = ("金额", "余额", "账面", "本金", "总额", "规模", "合计", "balance", "amount")
_BACKWARD_RE = re.compile(
    r"driven by|attributed to|due to|owing to|because|thanks to|resulted from|caused by|on account of|result of|since |"
    r"由于|因为|主要系|归因于|得益于|受[^，。；]{1,12}影响|系因|源于|源自|鉴于|(?<![原起])因(?![素此子])", re.IGNORECASE)
_FORWARD_RE_CAUSE = re.compile(r"所致|导致|致使|使得|以致|resulted in|led to|leading to", re.IGNORECASE)
_CONSEQUENCE = re.compile(r"而|从而")
_FIGURE_TAIL = r"\s*(?:了|约|近|超过?|达到?|至|为)?\s*[-+−－]?\d[\d,]*(?:\.\d+)?\s*"


def direction_words(text: str) -> List[Dict[str, Any]]:
    """Every direction word in the sentence: {start, end, sign (+1 rise / -1 fall), word}."""
    return [{"start": m.start(), "end": m.end(), "sign": _DIRECTION_WORDS[m.group(0).lower()], "word": m.group(0)}
            for m in _DIRECTION_RE.finditer(text)]


def _clause_of(text: str, pos: int) -> Tuple[int, int]:
    start = max((m.end() for m in _CLAUSE_BREAK.finditer(text, 0, pos)), default=0)
    end = _CLAUSE_BREAK.search(text, pos)
    return start, (end.start() if end else len(text))


def _on_cause_side(text: str, word: Dict[str, Any], c_start: int, c_end: int) -> bool:
    """The word describes a cause: it follows a why-marker with no 而 in between, sits inside 受…影响,
    or precedes a 所致 / 导致 marker in the same clause."""
    clause = text[c_start:c_end]
    for m in _BACKWARD_RE.finditer(clause):
        m_start, m_end = c_start + m.start(), c_start + m.end()
        if m_start <= word["start"] < m_end:
            return True
        if m_end <= word["start"] and not _CONSEQUENCE.search(text[m_end:word["start"]]):
            return True
    return any(c_start + m.start() >= word["end"] for m in _FORWARD_RE_CAUSE.finditer(clause))


def _subject_of(text: str, word: Dict[str, Any], lead_from: int, c_end: int,
                percent_kinds: List[Tuple[int, str]]) -> Optional[str]:
    """'share' / 'amount' / None: the quantity the word describes. The nearest cue before it, after the
    previous direction word ("金额增加但占比下降": 占比 goes with 下降); else the figure that follows it,
    and a percentage that follows is the quantity the percent check bound it to ("2025年末降至53.8%" is
    the share when 53.8% is the share, and a growth rate of the amount when it is a growth rate)."""
    low = text.lower()
    lead = low[lead_from:word["start"]]
    best: Tuple[int, Optional[str]] = (-1, None)
    for cues, name in ((_SHARE_SUBJECT, "share"), (_AMOUNT_SUBJECT, "amount")):
        for cue in cues:
            at = lead.rfind(cue)
            if at >= 0 and at + len(cue) > best[0]:
                best = (at + len(cue), name)
    if best[1]:
        return best[1]
    tail = low[word["end"]:c_end][:16]
    if re.match(_FIGURE_TAIL + r"个?百分点", tail):
        return "share"
    if re.match(_FIGURE_TAIL + r"(?:亿|百万|千万|万|千)?元", tail):
        return "amount"
    if re.match(_FIGURE_TAIL + r"[%％]", tail):
        figure = next((kind for start, kind in sorted(percent_kinds) if start >= word["end"]), "change")
        return "share" if figure in ("share", "points") else "amount"
    return None


def _pick(truths: List[Dict[str, Any]], line: str) -> List[Dict[str, Any]]:
    """The truths for the line the clause names; with no line named, the account's own (total) first."""
    if line:
        return [t for t in truths if t["line"] == line]
    return [t for t in truths if not t["line"]] or list(truths)


def _sign(value: float) -> int:
    return (value > 1e-9) - (value < -1e-9)


def _name_on_cause_side(text: str, start: int, end: int) -> bool:
    c_start, c_end = _clause_of(text, start)
    return _on_cause_side(text, {"start": start, "end": end}, c_start, c_end)


def _placement_uncertain(text: str, w: Dict[str, Any], c_start: int, c_end: int) -> bool:
    """A word with no quantity of its own, whose side of a cause cannot be told: it comes after 导致 / 致使
    ("…导致坏账风险上升": the account, or a further effect?), or it sits in a clause after one that gave a
    reason ("主要系A，B": B may carry on the list of reasons)."""
    clause = text[c_start:c_end]
    if any(c_start + m.end() <= w["start"] for m in _FORWARD_RE_CAUSE.finditer(clause)):
        return True
    if _BACKWARD_RE.search(clause):
        return False
    start = c_start
    while start > 0 and text[start - 1] not in "。\n":
        prev_start, prev_end = _clause_of(text, start - 1)
        if _BACKWARD_RE.search(text[prev_start:prev_end]):
            return True
        start = prev_start
    return False


def verify_directions(text: str, line_names: List[str], amount_truths: List[Dict[str, Any]],
                      share_truths: List[Dict[str, Any]],
                      percent_kinds: Optional[List[Tuple[int, str]]] = None,
                      note_signs: Optional[set] = None) -> List[Tuple[str, str, str, str]]:
    """[(severity, code, dimension, message)] for the direction words of one sentence.

    The result clause and the cause clause are read apart. A word in the result clause is a statement about the
    account and must be backed: against the amount's or the share's own sign, for the line the clause names. A word
    in the cause clause is a reason, which the cause check owns, unless it names a line and a quantity the
    citations can test ("主要系甲公司余额增加所致"), in which case it must agree with them. A word with no quantity of
    its own is compared with every cited reading; when the readings differ, or the word cannot be placed on one side
    of a cause, a person decides. Nothing here looks at what the author called the claim."""
    out: List[Tuple[str, str, str, str]] = []
    infos: List[Dict[str, Any]] = []
    words = direction_words(text)
    for w in words:
        c_start, c_end = _clause_of(text, w["start"])
        before = [x for x in words if x["end"] <= w["start"] and x["start"] >= c_start]
        lead_from = max([c_start] + [x["end"] for x in before])
        # a line named on the far side of a why-marker is part of the cause ("甲公司回款增加导致余额下降"),
        # not the line the effect is about
        reasons = [c_start + m.end() for m in _CAUSE_RE.finditer(text[c_start:w["start"]])]
        if reasons:                                  # "因甲公司回款增加而余额下降": the effect starts after 而
            reasons += [c_start + m.end() for m in _CONSEQUENCE.finditer(text[c_start:w["start"]]) if c_start + m.end() > reasons[0]]
        label_from = max([c_start] + reasons)
        lead_labels = [(text.rfind(lab, max(lead_from, label_from), w["start"]) + len(lab), len(lab), lab) for lab in line_names
                       if text.rfind(lab, max(lead_from, label_from), w["start"]) >= 0]
        clause_labels = [(text.rfind(lab, label_from, w["start"]) + len(lab), len(lab), lab) for lab in line_names
                         if text.rfind(lab, label_from, w["start"]) >= 0]
        found = max(lead_labels or clause_labels, default=None)
        subject = _subject_of(text, w, lead_from, c_end, percent_kinds or [])
        cause = _on_cause_side(text, w, c_start, c_end)
        infos.append({"w": w, "line": found[2] if found else "", "subject": subject, "cause": cause,
                      "uncertain": (not subject and not cause and _placement_uncertain(text, w, c_start, c_end))})

    def word_text(w):
        return ("an increase" if w["sign"] > 0 else "a decrease") + f" ('{w['word']}')"

    def verdict_of(signs):
        return "an increase" if signs == {1} else ("a decrease" if signs == {-1} else "no change")

    seen_unsupported = set()
    for i in infos:
        w, subject, line = i["w"], i["subject"], i["line"]
        if not subject:
            continue
        truths = _pick(share_truths if subject == "share" else amount_truths, line)
        signs = {_sign(t["value"]) for t in truths}
        if i["cause"]:
            if line and len(signs) == 1 and signs != {w["sign"]}:
                out.append(("failed", "SIGN_ERROR", "arithmetic",
                            f"the cause clause says {word_text(w)} of the {subject} of {line}; the cited figures show {verdict_of(signs)}"))
            continue
        if not truths and note_signs and w["sign"] in note_signs:
            continue                                 # the cited note states this direction (its own words, or the adjustment's sign)
        if not truths:
            if (subject, line) not in seen_unsupported:
                seen_unsupported.add((subject, line))
                out.append(("failed", "DIRECTION_UNSUPPORTED", "arithmetic",
                            f"{word_text(w)} of the {subject}{' of ' + line if line else ''} is stated, but "
                            + ("no pair of cited shares" if subject == "share" else "no cited movement") + " backs it"
                            + ("; a movement of another quantity does not stand in for it" if (amount_truths or share_truths) else "")))
        elif len(signs) > 1:
            out.append(("review", "DIRECTION_UNBOUND", "binding",
                        f"{word_text(w)} of the {subject} could mean more than one cited line, and they do not move the same way; "
                        "name the line"))
        elif signs != {w["sign"]}:
            out.append(("failed", "SIGN_ERROR", "arithmetic",
                        f"the text says {word_text(w)} of the {subject}{' of ' + line if line else ''}; the cited figures show {verdict_of(signs)}"))

    result_words = [i for i in infos if not i["cause"]]
    seen_loose = set()
    for i in result_words:
        w, line = i["w"], i["line"]
        if i["subject"]:
            continue
        truths = _pick(amount_truths, line) + [t for t in share_truths if not line or t["line"] == line]
        signs = {_sign(t["value"]) for t in truths}
        others = [o for o in result_words if o is not i and o["w"]["sign"] != w["sign"]]
        if not truths and note_signs and w["sign"] in note_signs:
            continue
        if not truths:
            if i["uncertain"]:
                out.append(("review", "DIRECTION_UNBOUND", "binding",
                            f"{word_text(w)} cannot be placed on the account or on a cause, and nothing cited backs it"))
            elif "unsupported" not in seen_loose:
                seen_loose.add("unsupported")
                out.append(("failed", "DIRECTION_UNSUPPORTED", "arithmetic",
                            f"{word_text(w)} is stated and no cited movement or pair of shares backs it"))
        elif len(signs) > 1:
            out.append(("review", "DIRECTION_UNBOUND", "binding",
                        f"{word_text(w)} could mean the amount or the share, and they do not move the same way; "
                        "say which one the sentence describes"))
        elif signs != {w["sign"]}:
            if i["uncertain"] or others:
                out.append(("review", "DIRECTION_UNBOUND", "binding",
                            f"{word_text(w)} disagrees with the cited figures ({verdict_of(signs)}), but the sentence also says the "
                            "opposite elsewhere or the word sits beside a cause: it cannot be told whether it describes the account"))
            else:
                out.append(("failed", "SIGN_ERROR", "arithmetic",
                            f"the text says {word_text(w)}{' of ' + line if line else ''}; the cited figures show {verdict_of(signs)}"))
    return out


def verify_claim(raw_claim: Dict[str, Any], contracts: Dict[str, EvidenceContract], *, default_account: str = "",
                 index: int = 0) -> Dict[str, Any]:
    claim = normalize_claim(raw_claim, default_account, index)
    account, text = claim["account"], claim["text"]
    contract = contracts.get(account)
    findings: List[ClaimFinding] = []
    checked: set = set()

    def fail(code: str, dim: str, message: str) -> None:
        findings.append(ClaimFinding(code, dim, "failed", message))

    def info(code: str, dim: str, message: str) -> None:
        findings.append(ClaimFinding(code, dim, "info", message))

    def review(code: str, dim: str, message: str) -> None:
        """Could not be tested. The claim is not failed and not verified: it waits for a person."""
        findings.append(ClaimFinding(code, dim, "review", message))

    if contract is None:
        fail("NO_CONTRACT", "citations", f"no evidence contract for account '{account}'")
        return _verdict(claim, findings, {"citations"})
    checked.add("citations")

    # -- citations ---------------------------------------------------------------
    if claim.get("source_version") and claim["source_version"] != contract.source_version:
        fail("STALE_SOURCE", "citations",
             f"the claim was written against source version {claim['source_version']}; the workbook now reads {contract.source_version}")
    cited: List[Dict[str, Any]] = []
    cited_notes: List[Dict[str, Any]] = []
    for ref in claim["cites"]:
        item, owner = _lookup(ref, contracts, account)
        if item is None and contract.note(ref) is not None:
            cited_notes.append(contract.note(ref))           # a note's own figures may be stated when it is cited
            continue
        if item is None:
            fail("UNKNOWN_REF", "citations", f"cited fact '{ref}' does not exist (or is on a hidden sheet)")
            continue
        if item["account"] != account:
            if not claim["cross_account"]:
                fail("WRONG_ACCOUNT", "citations", f"'{ref}' belongs to '{item['account']}', the claim is about '{account}'")
                continue
            # Declaring the claim cross-account lets the fact be cited; it cannot say which account the sentence
            # means by the figure, so a person decides.
            if not any(f.code == "CROSS_ACCOUNT_UNBOUND" for f in findings):
                review("CROSS_ACCOUNT_UNBOUND", "citations",
                       f"'{ref}' belongs to '{item['account']}', the claim is about '{account}'; "
                       "which account the figure describes is not decided here")
        cited.append(item)
    cited_facts = [c for c in cited if c.get("basis") != BASIS_DERIVED]
    cited_derived = [c for c in cited if c.get("basis") == BASIS_DERIVED]
    claim_periods = set(claim["periods"])
    latest_declared = max(claim_periods) if claim_periods else None

    amounts, percents = text_amounts(text), text_percents(text)
    if amounts:
        checked.add("amounts")
    if amounts or percents:
        checked.add("coverage")
    if (amounts or percents) and not cited and not cited_notes and not claim["cause_cites"]:
        fail("AMOUNT_UNCITED", "coverage", "the sentence states figures and cites no fact")
    if _INTERNAL_TERMS.search(text):
        fail("INTERNAL_TERMS_IN_TEXT", "format", f"'{_INTERNAL_TERMS.search(text).group(0)}' is a reference or an internal name; "
             "the report states facts, not their refs or bases")
    source_notes = cited_notes + [n for n in (contract.note(r) for r in claim["cause_cites"]) if n is not None]
    note_facts = {n["ref"]: note_figures(n, contract.periods) for n in source_notes}

    # -- the values the citations allow -------------------------------------------------
    allowed: List[Tuple[float, Dict[str, Any], str]] = []        # (abs value, source item, how)
    for f in cited_facts:
        allowed.append((abs(f["value"]), f, "cited"))
    for d in cited_derived:
        if d.get("value") is not None:
            allowed.append((abs(d["value"]), d, "derived"))
    by_period: Dict[str, List[Dict[str, Any]]] = {}
    for f in cited_facts:
        if f["role"] == ROLE_COMPONENT:
            by_period.setdefault(f["period"], []).append(f)
        elif f["role"] == ROLE_SUPPORT and not f.get("subtotal"):
            by_period.setdefault(f"{f['period']}|{f.get('table')}", []).append(f)     # one table's lines, never mixed with another's
    for group, parts in by_period.items():
        period = group.split("|")[0]
        if len(parts) >= 2:
            allowed.append((abs(sum(p["value"] for p in parts)), {"id": "sum(" + ",".join(p["ref"] for p in parts) + ")",
                                                                 "period": period, "account": account, "line": "", "basis": BASIS_RAW,
                                                                 "role": "sum", "ref": "sum", "value": sum(p["value"] for p in parts)},
                            "sum_of_cited"))
    for ref_n, figs in note_facts.items():
        for a in figs["amounts"]:
            allowed.append((abs(a["value"]), {"id": f"note:{ref_n}", "ref": ref_n, "account": account, "role": "note", "basis": "note",
                                              "line": "", "period": a.get("period"), "currency": a["currency"], "value": a["value"],
                                              "annualised": bool(a.get("annualised"))}, "note"))
    pairs = _consecutive_pairs(cited_facts)
    for a, b in pairs:
        allowed.append((abs(b["value"] - a["value"]), {
            "id": f"movement({a['ref']},{b['ref']})", "period": b["period"], "account": account, "basis": BASIS_DERIVED,
            "role": "movement", "ref": "movement", "periods": [a["period"], b["period"]], "value": b["value"] - a["value"]},
            "implicit_movement"))

    # -- amounts ------------------------------------------------------------------------
    matched_sources: List[Tuple[Dict[str, Any], Dict[str, Any], str]] = []   # (span, source, how)
    others = list(contract.facts) + list(contract.derived)
    # line names as a sentence writes them: a counterparty row is often starred in the tab ("*甲公司"), never in prose; with the
    # star kept, no starred line could be recognised, so a figure put under the wrong counterparty passed (2026-11-08)
    component_labels = sorted({_bare(f["line"]) for f in contract.facts if f["role"] == ROLE_COMPONENT and f["line"]}, key=len, reverse=True)
    for span in amounts:
        target = abs(span["value"])
        currency = _currency_at(text, span["start"], span["end"])
        hits = [(v, s, how) for v, s, how in allowed if abs(v - target) <= span["tolerance"] and (s.get("currency") or "CNY") == currency]
        hit = None
        if hits:
            # Two lines can carry the same figure (a loan listed by type and by bank).
            # Prefer the one the sentence attaches the amount to. Then a change for a figure the sentence writes as a change
            # ("较上年减少100万元") and a balance for one it does not, under the period it is written under: a fall to nil is
            # the size of the prior balance, and reading it as that balance put under the later period was a false WRONG_PERIOD.
            mentions_here = period_mentions(text)
            named = attached_label(mask_mentions(text, mentions_here), span["start"], component_labels)
            mention = attached_period(text, span["start"], span["end"], mentions_here)
            as_change = _has(text[max(0, span["start"] - 8):span["start"]], _UP + _DOWN + _CHANGE_CUES)

            def fit(h: Tuple[float, Dict[str, Any], str]) -> Tuple[bool, bool, bool]:
                src = h[1]
                periods = set(src.get("periods") or ([src["period"]] if src.get("period") else []))
                is_change = src.get("role") == "movement" or src.get("kind") == KIND_MOVEMENT
                return (bool(named) and _bare(src.get("line")) == named, as_change == is_change,
                        mention is not None and any(mention_matches(mention, p) for p in periods))
            hit = max(hits, key=fit)                     # max keeps the first of equals, so ties fall back to the order above
        if hit:
            matched_sources.append((span, hit[1], hit[2]))
            if hit[1].get("role") == "note":
                info("NOTE_SOURCED_AMOUNT", "amounts", f"'{span['written']}' is taken from {hit[1].get('ref')}'s text (free text, not a cell): "
                     "a person should read it against the note")
            # Amounts are matched as magnitudes, because prose writes a fall as "减少30万元". A sign the sentence DOES write is
            # held to account: "-130万元" for a balance of +130万 matched before.
            value = hit[1].get("value")
            if span.get("sign") and value is not None and abs(value) > span["tolerance"] and (span["sign"] == "-") != (value < 0):
                say = (f"'{span['written']}' is written as {'negative' if span['sign'] == '-' else 'positive'}, but the cited "
                       f"{hit[1].get('ref')} is {_fmt(value)} yuan")
                if hit[1].get("role") in _LINE_ROLES:     # a deduction line can be shown either way round
                    review("SIGN_UNCLEAR", "amounts", say + "; if the line is a deduction, say so")
                else:
                    fail("SIGN_ERROR", "amounts", say)
            continue
        scaled = next(((v, s) for v, s, _h in allowed if v and any(
            abs(target - v * 10 ** k) <= max(span["tolerance"], abs(target) * 0.01) for k in (-4, -3, -2, -1, 1, 2, 3, 4))), None)
        if scaled:
            factor = 10 ** round(math.log10(max(target, 1e-9) / max(scaled[0], 1e-9)))
            fail("UNIT_ERROR", "amounts",
                 f"'{span['written']}' is {factor:g}x the cited {scaled[1].get('ref')} "
                 f"({_fmt(scaled[0])} yuan): a unit or scale slip")
            continue
        elsewhere = next((o for o in others if o.get("value") is not None
                          and abs(abs(o["value"]) - target) <= span["tolerance"]), None)
        if elsewhere is not None:
            fail("AMOUNT_ELSEWHERE", "amounts",
                 f"'{span['written']}' is not what the cited facts say; it equals {elsewhere.get('ref')} "
                 f"({elsewhere.get('account')} {elsewhere.get('period')} {elsewhere.get('basis')} {elsewhere.get('line') or elsewhere.get('kind') or 'total'}), which is not cited")
        else:
            fail("AMOUNT_UNMATCHED", "amounts",
                 f"'{span['written']}' ({_fmt(target)} yuan) is not accounted for by any cited fact")

    # -- percentages ----------------------------------------------------------------------
    # Bound like an amount: to the period the sentence writes it under, the line it names, and the
    # kind of ratio the words say (a share of the total, or a change between periods). Matching a
    # value anywhere in the cited set let two shares swapped between periods, or a share called a
    # growth rate, pass.
    pct_mentions = period_mentions(text)
    masked_text = mask_mentions(text, pct_mentions)
    line_names = [lab for lab in component_labels if lab != account and lab not in account and account not in lab]
    pct_candidates: List[Dict[str, Any]] = []
    for d in cited_derived:
        if d.get("pct") is not None:
            signed = d["pct"] * 100
            pct_candidates.append({"pct": abs(signed), "signed": signed, "kind": "share" if d.get("kind") == KIND_SHARE else "change",
                                   "periods": list(d.get("periods") or [d.get("period")]), "line": d.get("line", ""), "ref": d.get("ref"), "item": d,
                                   "base": d.get("base"), "end": d.get("end")})
    for a_fact, b_fact in pairs:
        if abs(a_fact["value"]) > 0:
            signed = (b_fact["value"] - a_fact["value"]) / abs(a_fact["value"]) * 100
            pct_candidates.append({"pct": abs(signed), "signed": signed, "kind": "change",
                                   "periods": [a_fact["period"], b_fact["period"]],
                                   "line": a_fact["line"] if a_fact["role"] in _LINE_ROLES else "",
                                   "ref": f"movement({a_fact['ref']},{b_fact['ref']})", "item": None,
                                   "base": a_fact["value"], "end": b_fact["value"]})
    share_truths = share_movements(cited_derived)
    for t in share_truths:                    # "占比下降6.2个百分点": the difference of two shares, in points
        pct_candidates.append({"pct": abs(t["value"]) * 100, "signed": t["value"] * 100, "kind": "points", "periods": t["periods"],
                               "line": t["line"], "ref": t["ref"], "item": None})
    sentence_direction = _has(text, _UP) or _has(text, _DOWN)
    sentence_neutral = _has(text, _NEUTRAL_CHANGE)
    previous_end, previous_bound = -1, {}
    percent_kinds: List[Tuple[int, str]] = []          # (where the figure starts, the ratio it was bound to)
    note_pcts = [x for figs in note_facts.values() for x in figs["percents"]]
    for span in percents:
        if not span["points"] and any(abs(abs(span["value"]) - x) <= span["tolerance"] for x in note_pcts):
            continue                                   # a percentage the cited note itself states ("5%工程尾款")
        if any(d.get("zero_denominator") for d in cited_derived):
            fail("PERCENT_ZERO_DENOMINATOR", "arithmetic", f"'{span['written']}' is a change over a zero base; no percentage exists")
            continue
        hits = [c for c in pct_candidates if abs(abs(span["value"]) - c["pct"]) <= span["tolerance"]]
        if not hits:
            shown = ", ".join(f"{c['pct']:.1f}%" for c in pct_candidates)
            fail("PERCENT_UNMATCHED", "arithmetic",
                 f"'{span['written']}' does not match any ratio Python derived from the cited facts"
                 + (f" (they give {shown})" if shown else " (no ratio is cited)"))
            continue
        mention = attached_period(text, span["start"], span["end"], pct_mentions)
        named = attached_label(masked_text, span["start"], line_names)
        cue = percent_cue(text, span["start"])
        # "占比由2024年末的60.0%降至53.8%": after 至/达 the figure is the next level of the same ratio
        step = bool(previous_bound) and previous_end >= 0 and re.search(r"[至达]", text[previous_end:span["start"]])
        if step:
            named, cue = named or previous_bound.get("line"), cue or previous_bound.get("cue")
        reach_latest = bool(step or (mention and re.search(r"[至达]", text[mention["end"]:span["start"]])))

        def fits(c: Dict[str, Any]) -> Tuple[bool, bool, bool]:
            period_ok = (mention is None or any(mention_matches(mention, p) for p in c["periods"])
                         or (reach_latest and latest_declared is not None and c["periods"][-1] == latest_declared))
            line_ok = named is None or c["kind"] == "change" or _bare(c["line"]) == named
            # percentage points are the difference of two shares; a percent is a ratio. Neither stands in for the other.
            kind_ok = span["points"] == (c["kind"] == "points") and (span["points"] or cue is None or cue == c["kind"])
            return (period_ok, line_ok, kind_ok)
        best = max(hits, key=fits)
        percent_kinds.append((span["start"], best["kind"]))
        period_ok, line_ok, kind_ok = fits(best)
        if not period_ok:
            other = next((c for c in pct_candidates if c is not best and c["kind"] == best["kind"] and c["line"] == best["line"]
                          and any(mention_matches(mention, p) for p in c["periods"])), None)
            fail("WRONG_PERIOD", "binding",
                 f"'{span['written']}' is written under {mention['text']}, but that ratio is the {', '.join(best['periods'])} one"
                 + (f" ({other['pct']:.1f}% is the figure for {mention['text']}): the periods look swapped" if other is not None else ""))
        if not line_ok:
            fail("WRONG_LINE", "binding", f"the sentence attaches '{span['written']}' to '{named}', but that ratio is {best['line'] or 'the total'}'s")
        if not kind_ok:
            names = {"share": "share of the total", "change": "change between periods", "points": "difference of two shares in percentage points"}
            written_as = "percentage points" if span["points"] else ("a percentage" if cue is None else f"a {names[cue]}")
            fail("WRONG_OPERATION", "arithmetic",
                 f"'{span['written']}' is written as {written_as}, but the cited ratio is the {names[best['kind']]}")
        base, end = best.get("base"), best.get("end")
        if best["kind"] == "change" and base is not None and (base < 0 or (end is not None and base * end < 0)):
            fail("PERCENT_OVER_SIGN_CHANGE", "arithmetic",
                 f"'{span['written']}' is a change over a negative base or across a change of sign ({_fmt(base)} to {_fmt(end)} yuan): "
                 "such a percentage means nothing; state the two amounts")
            previous_end, previous_bound = span["end"], {"line": named, "cue": cue}
            continue
        if period_ok and line_ok and kind_ok:
            signed, slack = best["signed"], span["tolerance"]
            if span["sign"] == "-" and signed > slack:
                fail("SIGN_ERROR", "arithmetic",
                     f"'{span['written']}' is written as negative, but the cited {names_of(best)} is +{signed:.1f}%"
                     + (" (a share of the total is not negative; a fall is written as 下降 or 降至)" if best["kind"] == "share" else ""))
            elif span["sign"] == "+" and signed < -slack:
                fail("SIGN_ERROR", "arithmetic", f"'{span['written']}' is written as positive, but the cited {names_of(best)} is {signed:.1f}%")
            # An unsigned share of a deduction line (a provision against the receivable) is its magnitude:
            # "坏账准备占比27.3%" is how it is written, and "-27.3%" is also right. Only an explicit sign is held to account.
            elif (not span["sign"] and best["kind"] in ("change", "points") and signed < -slack
                  and not sentence_direction and not sentence_neutral):
                review("PERCENT_SIGN_UNSTATED", "arithmetic",
                       f"'{span['written']}' fell ({signed:.1f}%) but the sentence states neither a minus sign nor a direction")
            if best["kind"] == "share" and named is None:
                review("LINE_UNBOUND", "binding", f"'{span['written']}' is a share but the sentence does not say of which line")
            if best["kind"] == "share" and mention is None and not step and len(claim_periods) > 1:
                review("PERIOD_UNBOUND", "binding",
                       f"'{span['written']}' is not written under any period and the claim declares {len(claim_periods)}")
        previous_end, previous_bound = span["end"], {"line": named, "cue": cue}
    for d in cited_derived:
        if d.get("kind") == KIND_MOVEMENT and not d.get("valid", True) and any(c["item"] is d for c in pct_candidates):
            fail("MOVEMENT_NOT_PERMITTED", "binding", f"{d['ref']}: {d.get('invalid_reason')}")

    # -- binding: periods -------------------------------------------------------------------
    note_dates = [d for figs in note_facts.values() for d in figs["dates"]]
    for mention in period_mentions(text):
        if any((d["year"], d.get("month"), d.get("day")) == (mention["year"], mention.get("month"), mention.get("day")) for d in note_dates):
            continue                                   # a date the cited note itself states (an event, not a reporting period)
        if mention["kind"] == "year" and any(mention["year"] in figs["years"] for figs in note_facts.values()):
            continue                                   # a year the note names ("2025审计费"), not a period the claim reports
        if claim_periods and not any(mention_matches(mention, p) for p in claim_periods):
            fail("WRONG_PERIOD", "binding",
                 f"the text names '{mention['text']}', which is not among the periods this claim declares ({', '.join(sorted(claim_periods))})")
        elif not claim_periods:
            fail("PERIODS_UNDECLARED", "binding", "the sentence names a period but the claim declares none")
            break
    for item in cited:
        periods = set(item.get("periods") or ([item["period"]] if item.get("period") else []))
        if claim_periods and periods and not periods <= claim_periods:
            fail("WRONG_PERIOD", "binding",
                 f"'{item.get('ref')}' is from {', '.join(sorted(periods))}; the claim declares {', '.join(sorted(claim_periods))}")
    # Each amount belongs to the period the SENTENCE puts it under. Checking only that the amount is
    # somewhere in the cited set and the period somewhere in the declared set lets "2024年末130万元,
    # 2025年末100万元" pass when the two figures are the other way round.
    mentions = period_mentions(text)
    if mentions or matched_sources or cited:
        checked.add("binding")
    for span, source, _how in matched_sources:
        src_periods = set(source.get("periods") or ([source["period"]] if source.get("period") else []))
        if not src_periods:
            continue
        mention = attached_period(text, span["start"], span["end"], mentions)
        if mention is None:
            if len(claim_periods) > 1 and source.get("role") in (ROLE_TOTAL, ROLE_COMPONENT, ROLE_SUPPORT, "sum"):
                review("PERIOD_UNBOUND", "binding",
                       f"'{span['written']}' is not written under any period and the claim declares {len(claim_periods)}; "
                       "put the period next to each amount, or write one claim per period")
            continue
        if any(mention_matches(mention, p) for p in src_periods):
            continue
        between = text[mention["end"]:span["start"]]
        if re.search(r"[至达]", between) and source.get("period") == latest_declared:
            continue                                    # "由2024年末的100万元增至130万元": the amount after 至 is the later period
        other = next((f for f in cited_facts if f is not source and mention_matches(mention, f["period"])), None)
        if other is not None and abs(abs(other["value"]) - abs(span["value"])) > span["tolerance"]:
            fail("WRONG_PERIOD", "binding",
                 f"'{span['written']}' is written under {mention['text']}, but the cited {other['period']} figure is "
                 f"{abs(other['value']) / 1e4:,.2f}万元 and this amount is the {source['period']} one: the periods look swapped")
        else:
            fail("WRONG_PERIOD", "binding",
                 f"'{span['written']}' is written under {mention['text']}, but the fact it matches is from {', '.join(sorted(src_periods))}")

    # -- composition on the tab's basis --------------------------------------------------------
    # "由A和B构成" over every detail line of a period says they make up the balance. When the deck's balance is adjusted and the
    # lines are the tab's own, they add up to the tab's total, not to it; within the detail check's tolerance that passed before
    # (the ten-case review, 2026-11-08: retained earnings, 9.2万 of adjustments).
    if _has(text, ("构成", "组成", "consist", "made up of", "comprise")):
        by_p: Dict[str, List[Dict[str, Any]]] = {}
        for f in cited_facts:
            if f["role"] == ROLE_COMPONENT:
                by_p.setdefault(f["period"], []).append(f)
        for period, parts in by_p.items():
            every = [f for f in contract.facts if f["period"] == period and f["role"] == ROLE_COMPONENT]
            adjusted = contract.total(period, BASIS_ADJUSTED)
            if adjusted is None or len(parts) < len(every):
                continue
            gap = adjusted - sum(f["value"] for f in parts)
            if abs(gap) > max(1.0, 0.0005 * abs(adjusted)):
                fail("COMPOSITION_BASIS", "binding",
                     f"{period}: the lines named are the tab's own and add up to {_fmt(adjusted - gap)} yuan, not to the balance of "
                     f"{_fmt(adjusted)} (a difference of {_fmt(gap)}, the adjustments): say they are before the adjustments, or do not "
                     "call them its make-up")

    # -- binding: basis, line, permissions ----------------------------------------------------
    # A sentence may quote a total the deck did not adopt only when it is describing the gap itself: it states the
    # difference Python computed. Calling the claim a "difference" does not do that.
    quotes_gap = any(src.get("basis") == BASIS_DERIVED and src.get("kind") == KIND_DIFFERENCE for _s, src, _h in matched_sources)
    for span, source, how in matched_sources:
        if (source.get("role") == ROLE_TOTAL and source.get("basis") != BASIS_RAW and not quotes_gap
                and not (contract.permissions.get(source["period"]) or {}).get("describe_total", True)):
            fail("TOTAL_NOT_DESCRIBABLE", "binding",
                 f"{source['period']}: the tab and Financials disagree and no rule chose between them; state the difference, not either figure as the balance")
        if source.get("basis") == BASIS_RAW and source.get("role") == ROLE_TOTAL and not quotes_gap:
            perm = contract.permissions.get(source["period"]) or {}
            overridden = any(c.get("overridden") for c in contract.checks if c["period"] == source["period"])
            if overridden or not perm.get("describe_total", True):
                fail("WRONG_BASIS", "binding",
                     f"'{span['written']}' is the tab's own total for {source['period']}, which the deck replaced with Financials; quote the statement figure")
        if source.get("role") == ROLE_COMPONENT:
            perm = contract.permissions.get(source["period"]) or {}
            basis = perm.get("detail_basis", "tied")
            if basis in ("untied", "none"):
                fail("DETAIL_NOT_PERMITTED", "binding",
                     f"the tab's detail for {source['period']} does not add up to the account total; it may not be used to explain it")
            elif basis == "ambiguous" and not any(f.code == "DETAIL_NESTING_UNKNOWN" for f in findings):
                info("DETAIL_NESTING_UNKNOWN", "binding",
                     f"{source['period']}: the lines do not sum to the total as a flat list and their nesting is not known; "
                     "the figures are the tab's, the sentence may not claim they add up to the account total")
    for span, source, how in matched_sources:
        if how == "implicit_movement":
            later = (source.get("periods") or [source.get("period")])[-1]
            if not (contract.permissions.get(later) or {}).get("compute_movement"):
                fail("MOVEMENT_NOT_PERMITTED", "binding",
                     f"{later}: a movement is not computed for this period "
                     + "; ".join((contract.permissions.get(later) or {}).get("reasons") or []))
    # a detail row carrying the account's own name is the account, not a component of it
    labels = [lab for lab in component_labels if lab != account and lab not in account and account not in lab]
    for span, source, _how in matched_sources:
        named = attached_label(mask_mentions(text, period_mentions(text)), span["start"], labels)
        own = str(source.get("line") or "").lstrip("*＊ ").strip()
        written_own = own and (own in text[max(0, span["start"] - 60):span["start"]]
                               or any(own[-k:] in text[max(0, span["start"] - 24):span["start"]] for k in range(4, min(len(own), 12) + 1)))
        if named and source.get("role") == ROLE_COMPONENT and source.get("line") and _bare(source["line"]) != named and not written_own:
            fail("WRONG_LINE", "binding",
                 f"the sentence attaches '{span['written']}' to '{named}' but cites the '{source['line']}' figure")
    # A line the sentence names must have a figure behind it. Two lines that
    # happen to add up to the cited total are not thereby cited.
    masked, cited_lines = text, {_bare(f["line"]) for f in cited_facts if f["role"] in _LINE_ROLES} | {
        _bare(d.get("line")) for d in cited_derived if d.get("line")}
    figure_starts = [x["start"] for x in amounts + percents]
    # The sentence makes a line-level statement when it cites a line's figure or share, or says so ("其中", "占比").
    # Otherwise a name is a counterparty mentioned in passing, and so is any name on the cause side of a clause
    # ("due to 甲公司's late rent"): the rule bites where a figure sits right after the name.
    line_level = bool(cited_lines) or _has(text, _COMPOSITION_CUES)
    for label in labels:
        position = masked.find(label)
        if position >= 0:
            masked = masked.replace(label, " " * len(label))
            attached = any(0 <= start - (position + len(label)) <= 14 for start in figure_starts)
            as_cause = _name_on_cause_side(text, position, position + len(label))
            about_note = any(label in str(n.get("text") or "") for n in source_notes)
            if label not in cited_lines and not about_note and (attached or (line_level and not as_cause)):
                fail("LINE_NOT_CITED", "binding", f"the sentence names '{label}' but cites none of its figures")
    if _has(text, _NIL):
        for p in claim_periods or contract.periods:
            if p in contract.permissions and not contract.permissions[p].get("state_nil_or_first_arose", True):
                fail("NIL_NOT_PERMITTED", "binding",
                     f"{p}: the deck's total for this period comes from Financials; the account may not be called nil / first arisen from the tab")
                break

    # -- direction ---------------------------------------------------------------------------
    # Each direction word describes ONE quantity: the amount (a movement Python computed from cited
    # facts) or the line's share of the total (two cited shares). The word is checked against that
    # quantity's own sign, never the other's: an amount can rise while its share falls. A word on the
    # cause side of a clause ("主要系回款增加所致") describes the cause, which the cause check owns.
    up, down = _has(text, _UP), _has(text, _DOWN)
    if percents or up or down or cited_derived or _has(text, _ANNUAL_ANY):
        checked.add("arithmetic")
    amount_truths = [{"line": "", "value": d["value"], "ref": d["ref"]}
                     for d in cited_derived if d.get("kind") == KIND_MOVEMENT and d.get("value") is not None]
    amount_truths += [{"line": a["line"] if a["role"] in _LINE_ROLES else "", "value": b["value"] - a["value"],
                       "ref": f"movement({a['ref']},{b['ref']})"} for a, b in pairs]
    if True:
        # the sentence decides how a deduction line's share is read: a minus sign on any percentage means signed,
        # none means its size (as an unsigned share is accepted as the size above)
        written_signed = any(x["sign"] == "-" for x in percents)
        direction_shares = [t if written_signed else {**t, "value": t["magnitude"]} for t in share_truths]
        note_signs = {sg for figs in note_facts.values() for sg in figs["signs"]}
        for finding in verify_directions(text, line_names, amount_truths, direction_shares, percent_kinds, note_signs):
            {"failed": fail, "review": review}[finding[0]](*finding[1:])
    for d in cited_derived:
        if d.get("kind") == KIND_MOVEMENT and not d.get("valid", True) and not any(f.code == "MOVEMENT_NOT_PERMITTED" for f in findings):
            fail("MOVEMENT_NOT_PERMITTED", "binding", f"{d['ref']}: {d.get('invalid_reason')}")
        if d.get("kind") in (KIND_SHARE, KIND_SUM, KIND_DIFFERENCE, KIND_ANNUALIZED) and not d.get("valid", True):
            fail("DERIVED_NOT_VALID", "binding", f"{d['ref']}: {d.get('invalid_reason')}")

    # -- annualisation ---------------------------------------------------------------------------
    if _has(text, _ANNUAL_ANY):
        annual = [d for d in cited_derived if d.get("kind") == KIND_ANNUALIZED] + [
            src for _span, src, _how in matched_sources if src.get("annualised")]       # a remark's own annualised value
        if not annual:
            fail("ANNUALIZATION_UNSUPPORTED", "arithmetic", "the sentence speaks of an annual figure and cites no annualised fact")
        else:
            if not _has(text, _ANNUAL_WORD):
                fail("ANNUALIZATION_UNLABELLED", "arithmetic", "an annualised figure must be called annualised (年化), not a full-year figure")
            if _has(text, _FORECAST_WORD) and not claim["inference"]:
                fail("ANNUALIZED_AS_FORECAST", "arithmetic", f"a linear annualisation is not a forecast ({LINEAR_ANNUALIZATION_NOTE})")

    # -- cause --------------------------------------------------------------------------------------
    if claim["cause_cites"] or _CAUSE_RE.search(text) or _FORWARD_RE.search(text):
        checked.add("causal")
        notes = []
        for ref in claim["cause_cites"]:
            note = contract.note(ref)
            if note is None:
                fail("UNKNOWN_REF", "citations", f"cited note '{ref}' does not exist")
            else:
                notes.append(note)
        related = [n for n in notes if len(_bigrams(text) & _bigrams(n["text"])) >= 2]
        conflicts = [(n, c) for n in related for c in cause_conflicts(text, n["text"], contract)]
        for n, why in conflicts:
            fail("CAUSE_CONTRADICTS_NOTE", "causal", f"{n['ref']} says something the sentence reverses or changes: {why}")
        if related and not conflicts:
            # Sharing words with a note shows the note is ABOUT this. It does not show the note says this.
            # Only contradictions can be screened for here; entailment cannot, so the cause is marked
            # (the deck renders it as reasoning) until a person has read the note against the sentence.
            info("CAUSE_NOTE_RELATED_UNCHECKED", "causal",
                 "cites " + ", ".join(n["ref"] for n in related) + ", which is about the same thing; "
                 "no contradiction was found, but the wording was not checked against it")
        elif notes and not related:
            fail("CAUSE_NOTE_UNRELATED", "causal", "the cited note(s) do not say what the sentence gives as the cause")
        elif not notes and claim["inference"]:
            info("CAUSE_INFERENCE", "causal", "stated as an inference: no note supports it; the deck marks it")
        elif not notes:
            fail("CAUSE_UNSUPPORTED", "causal", "a cause is stated as fact with no note behind it; cite one or mark it as inference")

    return _verdict(claim, findings, checked)


_STRIP_BEFORE_CUES = ("未分配", "未来", "未到期", "未开票", "未实现", "未确认融资")
_CUES = {
    "neg": ("未能", "尚未", "没有", "并非", "不是", "无法", "未", "没", "not ", " no ", "without"),
    "done": ("已经", "已", "完成", "结清", "completed", "received", "settled"),
    "expect": ("预计", "预期", "拟", "计划", "将于", "将在", "尚待", "expected", "will ", "plan to"),
    "all": ("全部", "全额", "所有", "完全", "全数", "entire", "all ", "in full"),
    "part": ("部分", "仅", "少量", "partial", "only", "part of"),
    "delay": ("延迟", "推迟", "拖延", "滞后", "逾期", "delayed", "late"),
    "early": ("提前", "加快", "及时", "early", "earlier"),
    "up": ("增加", "上升", "增长", "上涨", "increase", "rose"),
    "down": ("减少", "下降", "降低", "下滑", "decrease", "fell"),
}


def _window_text(text: str, grams: set, radius: int = 10) -> str:
    spans = []
    for g in grams:
        for m in re.finditer(re.escape(g), text):
            spans.append((max(0, m.start() - radius), min(len(text), m.end() + radius)))
    spans.sort()
    merged: List[List[int]] = []
    for start, end in spans:
        if merged and start <= merged[-1][1]:
            merged[-1][1] = max(merged[-1][1], end)
        else:
            merged.append([start, end])
    return " ".join(text[a:b] for a, b in merged)


def _cue_classes(text: str) -> set:
    work = text.lower()
    for word in _STRIP_BEFORE_CUES:
        work = work.replace(word, "")
    return {name for name, words in _CUES.items() if any(w in work for w in words)}


def cause_conflicts(claim_text: str, note_text: str, contract: Optional[EvidenceContract] = None) -> List[str]:
    """Ways a sentence reverses or changes what a related note says, found WITHIN the stretches around
    the words they share (so a negation elsewhere in a long note is not charged to the claim):
    polarity, done against expected, all against part, an opposite direction, a different named line.
    Screening for contradiction only; a sentence that passes is related to the note, not entailed by it."""
    shared = _bigrams(claim_text) & _bigrams(note_text)
    c, n = _cue_classes(_window_text(claim_text, shared)), _cue_classes(_window_text(note_text, shared))
    found: List[str] = []
    if ("neg" in c) != ("neg" in n):
        found.append("one of them negates what the other affirms")
    if ("done" in c and "expect" in n and "done" not in n) or ("expect" in c and "done" in n and "expect" not in n):
        found.append("completed in one, only expected in the other")
    if ("all" in c and "part" in n and "all" not in n) or ("part" in c and "all" in n and "part" not in n):
        found.append("all in one, part in the other")
    for a, b in (("delay", "early"), ("early", "delay"), ("up", "down"), ("down", "up")):
        if a in c and b in n and a not in n:
            found.append(f"opposite direction ({a} / {b})")
            break
    if contract is not None:
        labels = {f["line"] for f in contract.facts if f["role"] == ROLE_COMPONENT and f["line"] and f["line"] != contract.account}
        in_claim, in_note = {l for l in labels if l in claim_text}, {l for l in labels if l in note_text}
        if in_claim and in_note and not (in_claim & in_note):
            found.append(f"a different line ({', '.join(sorted(in_claim))} against {', '.join(sorted(in_note))})")
    return found


def _verdict(claim: Dict[str, Any], findings: List[ClaimFinding], checked: set) -> Dict[str, Any]:
    dimensions = {}
    for dim in ("citations", "amounts", "binding", "arithmetic", "causal", "coverage"):
        mine = [f for f in findings if f.dimension == dim]
        if any(f.severity == "failed" for f in mine):
            dimensions[dim] = "failed"
        elif any(f.severity == "review" for f in mine):
            dimensions[dim] = "unverified"
        elif any(f.code == "CAUSE_NOTE_RELATED_UNCHECKED" for f in mine):
            dimensions[dim] = "related_unchecked"
        elif any(f.code == "CAUSE_INFERENCE" for f in mine):
            dimensions[dim] = "inference"
        else:
            dimensions[dim] = "ok" if dim in checked else "n/a"
    failed = any(f.severity == "failed" for f in findings)
    unverified = any(f.severity == "review" for f in findings)
    marked = any(f.code in ("CAUSE_INFERENCE", "CAUSE_NOTE_RELATED_UNCHECKED") for f in findings)
    caveats = [f.code for f in findings if f.code == "DETAIL_NESTING_UNKNOWN"]
    return {
        "claim_id": claim["id"], "account": claim["account"], "kind": claim["kind"], "text": claim["text"],
        "dimensions": dimensions,
        "verdict": "failed" if failed else ("unverified" if unverified else ("verified_inference" if marked else "verified")),
        # verified, but on a basis that could not be fully tested: kept apart so it is counted, not hidden
        "caveats": caveats,
        "findings": [f.to_dict() for f in findings],
    }


def verify_candidate(claims: List[Dict[str, Any]], contracts: Dict[str, EvidenceContract], *,
                     default_account: str = "") -> Dict[str, Any]:
    """Verdicts for a whole paragraph of claims, and the count a gate reads."""
    verdicts = [verify_claim(c, contracts, default_account=default_account, index=i) for i, c in enumerate(claims)]
    return {
        "verdicts": verdicts,
        "verified": sum(v["verdict"] == "verified" for v in verdicts),
        "inference": sum(v["verdict"] == "verified_inference" for v in verdicts),
        "caveated": sum(bool(v.get("caveats")) for v in verdicts),
        "failed": sum(v["verdict"] == "failed" for v in verdicts),
        "unverified": sum(v["verdict"] == "unverified" for v in verdicts),
        # a claim that could not be tested is not a passed claim: the candidate waits for a person
        "ok": bool(verdicts) and all(v["verdict"] in ("verified", "verified_inference") for v in verdicts),
        "by_code": _count(f["code"] for v in verdicts for f in v["findings"] if f["severity"] in ("failed", "review")),
    }


def _count(items: Iterable[str]) -> Dict[str, int]:
    out: Dict[str, int] = {}
    for item in items:
        out[item] = out.get(item, 0) + 1
    return out


# ---------------------------------------------------------------------------
# From claims to the shapes the rest of the product reads
# ---------------------------------------------------------------------------

def render_claims(claims: List[Dict[str, Any]]) -> str:
    """The paragraph a reader sees: the claims' own sentences, in order."""
    parts = []
    for claim in claims:
        sentence = str(claim.get("text") or "").strip()
        if sentence and sentence[-1] not in "。；;.!！":
            sentence += "。"
        parts.append(sentence)
    return "".join(parts)


#: verifier dimensions where a failure means a figure the databook contradicts (wrong amount, period, line, basis, sum, or a
#: ref that does not exist): red in the deck. A failure elsewhere (a cause with nothing behind it, wording) is orange.
_RED_DIMENSIONS = ("amounts", "arithmetic", "binding", "citations")
#: flags on a verified sentence that a person should read: orange in the deck
HIGHLIGHT_FLAGS = ("note_amount", "cause_not_proven", "note_wording_unchecked", "detail_caveat", "lead_left_to_person")


def to_legacy_result(account: str, claims: List[Dict[str, Any]], outcome: Dict[str, Any],
                     flags: Optional[List[Dict[str, Any]]] = None) -> Dict[str, Any]:
    """A pipeline-shaped result for one account, so the payload builder, the deck's highlighting and delivery.decide_delivery
    read an agent-written paragraph the way they read a pipeline one, with the pipeline's meaning of the colours: red
    (`hallucination`, which gates delivery) for a figure the databook contradicts, orange (`reasoning`) for an inference, a
    cause nothing proves, and a verified sentence a flag says a person should read; the rest `data-backed`. Until 2026-11-09
    an unproven cause and an unchecked sentence were red too, and a flagged sentence was not marked at all."""
    text = render_claims(claims)
    marked = {}
    for f in flags or []:
        if f.get("kind") in HIGHLIGHT_FLAGS and isinstance(f.get("claim"), int):
            marked.setdefault(f["claim"], f)
    reviews = []
    for n, (claim, verdict) in enumerate(zip(claims, outcome["verdicts"])):
        if verdict["verdict"] in ("failed", "unverified"):
            failed = [f for f in verdict["findings"] if f["severity"] == "failed"]
            red = next((f for f in failed if f.get("dimension") in _RED_DIMENSIONS), None)
            first = red or next((f for f in verdict["findings"] if f["severity"] in ("failed", "review")), {})
            category = "hallucination" if red else "reasoning"
            supported, reason, code = False, first.get("message", ""), first.get("code")
        elif verdict["verdict"] == "verified_inference":
            marker = next((f for f in verdict["findings"] if f["code"] in ("CAUSE_INFERENCE", "CAUSE_NOTE_RELATED_UNCHECKED")), {})
            category, supported, reason, code = "reasoning", False, marker.get("message", "stated as an inference"), marker.get("code", "CAUSE_INFERENCE")
        elif n in marked:
            category, supported = "reasoning", False
            reason, code = f"{marked[n]['kind']}: {str(marked[n].get('why') or marked[n].get('text') or '')[:160]}", marked[n]["kind"].upper()
        else:
            category, supported, reason, code = "data-backed", True, "", None
        reviews.append({"clause": str(claim.get("text") or ""), "supported": supported, "category": category,
                        "reason": reason, "code": code, "claim_id": verdict["claim_id"]})
    return {"subagent_1": text, "subagent_2": text, "subagent_4": text, "final": text, "agent": True,
            "agent_4_validation": {"final_content": text, "clause_reviews": reviews}}


def _supporting_view(contract: EvidenceContract, totals: Dict[str, float], shown_periods: List[str], *, max_tables: int = 3,
                     max_rows: int = 15) -> List[Dict[str, Any]]:
    """The other tables of the account's tab, as the preparer laid them out: one row per line, one ref and one 万元 figure per
    column. Rows of 1% or more of the account's latest total, and every subtotal; at most fifteen a table, three tables."""
    tables: Dict[str, Dict[str, Any]] = {}
    for f in contract.facts:
        if f["role"] != ROLE_SUPPORT:
            continue
        if not f.get("period_assumed") and f["period"] not in shown_periods:
            continue
        t = tables.setdefault(f["table"], {"title": f.get("table_title"), "where": f.get("where"), "columns": [], "rows": {}})
        col = (f.get("column") or "", f["period"], bool(f.get("period_assumed")))
        if col not in t["columns"]:
            t["columns"].append(col)
        t["rows"].setdefault((f.get("row"), f["line"]), {"subtotal": bool(f.get("subtotal"))})[col] = f
    latest = next((totals[p] for p in reversed(shown_periods) if totals.get(p)), 0.0)
    out = []
    for tid in sorted(tables, key=lambda x: int(x[1:]) if x[1:].isdigit() else 0)[:max_tables]:
        t = tables[tid]
        cols = t["columns"]
        ranked = sorted(t["rows"].items(), key=lambda kv: -max((abs(f["value"]) for c, f in kv[1].items() if c != "subtotal"), default=0))
        keep = [kv for kv in ranked if kv[1]["subtotal"] or not latest or
                max((abs(f["value"]) for c, f in kv[1].items() if c != "subtotal"), default=0) >= 0.01 * latest][:max_rows]
        keep_keys = {k for k, _ in keep}
        rows = []
        for key, cells in sorted(t["rows"].items(), key=lambda kv: (kv[0][0] if kv[0][0] is not None else 10 ** 9)):
            if key not in keep_keys:
                continue
            rows.append({"line": key[1], "refs": [cells[c]["ref"] if c in cells else None for c in cols],
                         "wan": [round(cells[c]["value"] / 1e4, 2) if c in cells else None for c in cols],
                         **({"subtotal": True} if cells["subtotal"] else {})})
        entry = {"title": t["title"], "where": t["where"],
                 "columns": [{"head": h, "period": p, **({"period_assumed": True} if a else {})} for h, p, a in cols], "rows": rows}
        if len(t["rows"]) > len(rows):
            entry["rows_not_shown"] = len(t["rows"]) - len(rows)
        out.append(entry)
    return out


def evidence_view(contract: EvidenceContract, *, max_lines: int = 30, min_lines: int = 8, material_share: float = 0.01,
                  component_periods: int = 3) -> Dict[str, Any]:
    """The compact evidence an author is given. What decides whether a claim may be made
    (permissions, checks, issues, notes, derived figures) comes first; the detail lines last.

    Detail is chosen by materiality, one row per line across the periods shown: every line at least 1% of its period's total
    in any period shown, every line a note names, the parent of any line shown, and never fewer than the largest eight; at
    most thirty. What is left is summed per period (a D ref, so "其他N项合计" can be cited). Until 2026-11-09 this was the
    largest eight per period, each period its own list: the team's report enumerates groups and small special items
    (retentions, aged balances) the writer never saw. Facts a period does not permit are not withheld -- the permission is
    stated beside them."""
    def fact_view(f: Dict[str, Any]) -> Dict[str, Any]:
        out = {"ref": f["ref"], "period": f["period"], "basis": f["basis"], "line": f["line"] or "(total)",
               "yuan": round(f["value"], 2), "wan": round(f["value"] / 1e4, 2)}
        if f.get("nested_in"):
            out["nested_in"] = f["nested_in"]          # a sub-line of that line: never add it to its parent
        return out
    shown_periods = contract.periods[-component_periods:]
    by_ref = {f["ref"]: f for f in contract.facts}
    lines: Dict[Any, Dict[str, Any]] = {}          # one per sheet row (a repeated label on another row is another line)
    for f in contract.facts:
        if f["role"] != ROLE_COMPONENT or f["period"] not in shown_periods:
            continue
        key = f.get("row") if f.get("row") is not None else f["line"]
        entry = lines.setdefault(key, {"line": f["line"], "row": f.get("row"), "facts": {}})
        entry["facts"][f["period"]] = f
    totals = {}
    for p in shown_periods:
        t = contract.total(p, BASIS_ADJUSTED)
        if t is None:
            t = contract.total(p, BASIS_RAW)
        if not t:
            t = sum(abs(e["facts"][p]["value"]) for e in lines.values() if p in e["facts"])
        totals[p] = abs(t or 0.0)
    live = {k: e for k, e in lines.items() if any(abs(f["value"]) > 0.005 for f in e["facts"].values())}

    def weight(e: Dict[str, Any]) -> float:
        return max((abs(f["value"]) / totals[p] if totals.get(p) else 0.0) for p, f in e["facts"].items())
    note_text = " ".join(str(n.get("text") or "") for n in contract.notes)
    ranked = sorted(live, key=lambda k: -weight(live[k]))
    keep = set(ranked[:min_lines])
    keep |= {k for k in ranked if weight(live[k]) >= material_share}
    keep |= {k for k in ranked if len(_bare(live[k]["line"])) >= 2 and _bare(live[k]["line"]) in note_text}
    keep = set(sorted(keep, key=ranked.index)[:max_lines])
    parent_line = {}
    for k in list(keep):                           # a sub-line shown with its parent, so the structure reads
        for f in live[k]["facts"].values():
            parent = by_ref.get(f.get("nested_in") or "")
            if parent is not None:
                parent_line[k] = parent["line"]
                pk = parent.get("row") if parent.get("row") is not None else parent["line"]
                if pk in live:
                    keep.add(pk)
    order = sorted(keep, key=lambda k: (live[k]["row"] if live[k]["row"] is not None else 10 ** 9, ranked.index(k)))
    rows_out = []
    for k in order:
        e = live[k]
        row = {"line": e["line"], "refs": [e["facts"][p]["ref"] if p in e["facts"] else None for p in shown_periods],
               "wan": [round(e["facts"][p]["value"] / 1e4, 2) if p in e["facts"] else None for p in shown_periods]}
        if k in parent_line:
            row["nested_in"] = parent_line[k]      # a sub-line of that line: never add it to its parent
        rows_out.append(row)
    left = [k for k in ranked if k not in keep]
    rest = None
    if left:
        rest = {"lines": len(left), "refs": [], "wan": []}
        for p in shown_periods:
            # only the top level of what is left: a sub-line is already inside its parent's figure
            refs = [live[k]["facts"][p]["ref"] for k in left if p in live[k]["facts"] and not live[k]["facts"][p].get("nested_in")]
            if not refs:
                rest["refs"].append(None)
                rest["wan"].append(None)
                continue
            d = derive_sum(contract, refs) if len(refs) > 1 else None
            if d is not None and d.get("valid", True):
                rest["refs"].append(d["ref"])
                rest["wan"].append(round(d["value"] / 1e4, 2))
            else:
                rest["refs"].append(refs[0])
                rest["wan"].append(round(by_ref[refs[0]]["value"] / 1e4, 2))
    rest_refs = set((rest or {}).get("refs") or [])
    omitted = len(left)
    tables = _supporting_view(contract, totals, shown_periods)
    related = [{"account": f.get("from_account") or f["line"], "ref": f["ref"], "period": f["period"], "wan": round(f["value"] / 1e4, 2),
                **({"flow": True} if str(f.get("statement") or "").upper().startswith("IS") else {})}
               for f in contract.facts if f["role"] == ROLE_RELATED]
    material_left = sum(1 for k in left if weight(live[k]) >= material_share)    # lines the thirty-line cap kept out
    return {
        "account": contract.account, "source_version": contract.source_version, "periods": contract.periods,
        "permissions": {p: {k: v for k, v in perm.items() if k != "reasons"} | {"why": perm.get("reasons")}
                        for p, perm in contract.permissions.items()},
        "checks": contract.status_matrix(),
        "issues": contract.issues,
        "annualization": contract.annualization,
        "derived": [{k: d.get(k) for k in ("ref", "kind", "periods", "value", "pct", "valid", "invalid_reason", "formula")}
                    for d in contract.derived if d.get("ref") not in rest_refs],
        "notes": [{"ref": n["ref"], "kind": n["kind"], "period": n.get("period"), "text": n["text"]} for n in contract.notes],
        "totals": [fact_view(f) for f in contract.facts if f["role"] == ROLE_TOTAL],
        "components": {"periods": shown_periods, "lines": rows_out, **({"rest": rest} if rest else {})},
        **({"supporting_tables": tables} if tables else {}),
        **({"related": related} if related else {}),
        "components_not_shown": omitted,
        "material_not_shown": material_left,
    }

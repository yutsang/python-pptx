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

import itertools
import math
import re
from dataclasses import dataclass, field
from typing import Any, Dict, Iterable, List, Optional, Tuple

from .evidence_contract import (
    BASIS_ADJUSTED, BASIS_FINANCIALS, BASIS_RAW, ROLE_COMPONENT, ROLE_RELATED, ROLE_SUPPORT, ROLE_TOTAL, EvidenceContract, tie_tolerance,
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
    # 月 closes the month: "2026年6月470.0万元" is June 2026 and then the amount; only a day after it ("6月15日") makes it a date
    ("month", re.compile(r"(\d{4})\s*年\s*(\d{1,2})\s*月(?!\s*\d{1,2}\s*[日号])(?:末|底|份)?")),
    ("month_en", re.compile(_EN_MONTH + r"\s+(\d{4})", re.I)),
    ("year_end", re.compile(r"(\d{4})\s*年\s*(?:末|底|年末|年底)")),
    # a digit straight after the year is no month unless 月 follows it: "2024年312.2万元" names 2024 (the bracket "（2024年312.2万元、2025年330.7万元、
    # 2026年1-6月115.4万元）" had no period for its first two amounts, and the unlabelled bracket after it was read under the last, 2026-11-17)
    ("year", re.compile(r"(\d{4})\s*年(?!\s*月|\d{1,2}\s*[月日号])")),
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


_RUN_JOIN = re.compile(r"^\s*(?:[、，,]|和|及|与|以及)\s*$")
_LIST_JOIN = re.compile(r"^[度末底期份]?\s*(?:[、，,]|和|及|与|以及)\s*$")


def _amount_run(text: str, start: int) -> List[Dict[str, Any]]:
    """The amounts written as one list ("108.1万元、107.4万元、47.6万元") that holds the amount at `start`; just it when it stands alone."""
    spans = sorted(text_amounts(text), key=lambda sp: sp["start"])
    runs: List[List[Dict[str, Any]]] = []
    for sp in spans:
        if runs and _RUN_JOIN.match(text[runs[-1][-1]["end"]:sp["start"]]):
            runs[-1].append(sp)
        else:
            runs.append([sp])
    return next((r for r in runs if any(sp["start"] == start for sp in r)), [])


def _period_list(text: str, before: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """The periods named as one list ("2024年度、2025年度及2026年1-6月") that came last before the amount; [] when the last period named
    stands alone."""
    ordered = sorted(before, key=lambda m: m["start"])
    lists: List[List[Dict[str, Any]]] = []
    for m in ordered:
        if lists and _LIST_JOIN.match(text[lists[-1][-1]["end"]:m["start"]]):
            lists[-1].append(m)
        else:
            lists.append([m])
    return lists[-1] if lists else []


def _range_years(text: str, start: int, n: int) -> List[Dict[str, Any]]:
    """"2023-2025年各年调减…44.2万、45.8万、45.8万元": the run of amounts after a range of years is one year each, when the counts agree."""
    found = None
    for found in re.finditer(r"(\d{4})\s*年?\s*(?:至|到|-|－|—|~)\s*(\d{4})\s*年", text[:start]):
        pass
    if found is None or re.search(r"[。；;]", text[found.end():start]):
        return []
    years = list(range(int(found.group(1)), int(found.group(2)) + 1))
    if len(years) != n:
        return []
    end_of_year = "末" in text[found.end():start]
    out = []
    for y in years:
        where = {"text": f"{y}年", "start": found.start(), "end": found.end()}
        out.append({"kind": "date", "year": y, "month": 12, "day": 31, **where} if end_of_year else {"kind": "year", "year": y, **where})
    return out


def _labelled_series(text: str, before: List[Dict[str, Any]], n: int) -> List[Dict[str, Any]]:
    """"（2024年312.2万元、2025年330.7万元、2026年1-6月115.4万元）…（142.7万元、139.1万元、64.3万元）": a run of n amounts with no period
    of its own, right after n periods each written with its amount, in order, is the same periods again."""
    ordered = sorted(before, key=lambda m: m["start"])[-n:]
    if len(ordered) < n:
        return []
    keys = [(m["year"], m.get("month") or 0, m.get("day") or 0) for m in ordered]
    if keys != sorted(set(keys)):
        return []
    amounts = text_amounts(text)
    if not all(any(0 <= sp["start"] - m["end"] <= 3 for sp in amounts) for m in ordered):
        return []
    return ordered


def attached_period(text: str, start: int, end: int, mentions: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    """The period an amount at [start, end) is written under: the nearest mention before it in the
    same sentence, else the nearest after it ("余额130万元（2025年12月31日）"). None if the sentence names none."""
    before = [m for m in mentions if m["end"] <= start and "。" not in text[m["end"]:start]]
    # "2023年和2024年…分别调减614.5万元和826.9万元": the n-th amount after 分别 is the n-th period before it
    split = text.rfind("分别", 0, start)
    # A clause ends at 。 or ； : 分别 binds the amounts of its own clause to the periods of its own clause, or, where the clause names
    # none ("…；广告及推广费用分别为0.8万元、1.9万元、3.7万元"), to the list the nearest earlier clause gave. Before 2026-11-17 the amounts
    # counted ran to the sentence's end and the periods back to its start: a second 分别 clause (six amounts for three periods) switched
    # the rule off and every amount of the first was read under the last period, and three clauses of one period and three lines each
    # (a 3x3 table in one sentence) switched it on wrongly, reading the third clause's amounts under the first clause's years.
    if split >= 0 and not re.search(r"[。；;]", text[split:start]):
        boundaries = [m.end() for m in re.finditer(r"[。；;]", text[:split])]
        clause_start = boundaries[-1] if boundaries else 0
        listed = sorted((m for m in before if clause_start <= m["start"] and m["end"] <= split), key=lambda m: m["start"])
        if not listed:                 # a clause with no period of its own takes the list of the nearest earlier clause that has one
            edges = [0] + boundaries
            for lo, hi in zip(reversed(edges[:-1]), reversed(edges[1:])):
                own = sorted((m for m in before if lo <= m["start"] and m["end"] <= hi), key=lambda m: m["start"])
                if len(own) >= 2:
                    listed = own
                    break
        stop = re.search(r"[。；;]|分别|其中", text[split + 2:])
        limit = split + 2 + stop.start() if stop else len(text)
        # a period named after 分别 starts the next statement ("…分别为304.2万元和315.5万元，2025年度较上年增加11.3万元") and an amount in a
        # bracket is a part of the one before it ("分别为179.2万元、155.3万元和73.9万元（其中物业费108.1万元…）"): neither is counted as
        # one of the series, or its length stops matching the periods and every amount is read under the last one (two refused right
        # sentences of one book, 2026-11-17)
        later = [m["start"] for m in mentions if m["start"] > split + 2]
        if later:
            limit = min(limit, min(later))
        shut = _bracket_spans(text)
        series = [sp for sp in text_amounts(text) if split < sp["start"] < limit and not any(b0 <= sp["start"] < b1 for b0, b1 in shut)]
        # "分别为0元、0元、281.1万元和0元": a bare zero is no amount to the extractor and the series came up one short of its periods, so
        # the one amount went under the last period (a refused right sentence, 2026-11-17)
        series += [{"start": m.start(), "end": m.end(), "value": 0.0} for m in re.finditer(r"(?<![\d.,])0(?:\.0+)?元|零元", text)
                   if split < m.start() < limit and not any(sp["start"] <= m.start() < sp["end"] for sp in series)]
        series.sort(key=lambda sp: sp["start"])
        index = next((i for i, sp in enumerate(series) if sp["start"] == start), None)
        if len(listed) >= 2 and index is not None and len(series) == len(listed):
            return listed[index]
        # "2023年至2025年各年末…分别为A、B和C": a range of years is one year per amount (a portfolio book wrote its differences
        # this way four times and each was refused as written under the last year, 2026-11-10)
        span = None
        for span in re.finditer(r"(\d{4})\s*年?\s*(?:至|到|-|－|—|~)\s*(\d{4})\s*年", text[:split]):
            pass
        if span is not None and index is not None:
            years = list(range(int(span.group(1)), int(span.group(2)) + 1))
            if 2 <= len(years) == len(series):
                where = {"text": f"{years[index]}年", "start": span.start(), "end": span.end()}
                if "末" in text[span.end():split]:
                    return {"kind": "date", "year": years[index], "month": 12, "day": 31, **where}
                return {"kind": "year", "year": years[index], **where}
    # "…（108.1万元、107.4万元、47.6万元）" after "2024年度、2025年度及2026年1-6月": a run of amounts as long as the list of periods the
    # sentence gave last is that list again, one period for each
    run = _amount_run(text, start)
    if len(run) >= 2:
        index = next((i for i, sp in enumerate(run) if sp["start"] == start), None)
        last = _period_list(text, before)
        if index is not None:
            if last and len(last) == len(run):
                return last[index]
            for found in (_range_years(text, start, len(run)), _labelled_series(text, before, len(run))):
                if found:
                    return found[index]
    # "较2025年末减少705.67万元": the period after 较 is what the change is measured from, not the period of an amount written
    # after the change word (client file, 2026-11-09: such amounts were refused as written under the wrong period)
    before = [m for m in before if not (re.search(r"(?:较|比|相比|相较|对比)\s*$", text[max(0, m["start"] - 3):m["start"]])
                                        and _has(text[m["end"]:start], _UP + _DOWN))]
    if before:
        return max(before, key=lambda m: m["end"])
    after = [m for m in mentions if m["start"] >= end and "。" not in text[end:m["start"]]]
    if not after:
        return None
    first = min(after, key=lambda m: m["start"])
    # "甲所22.36万元（含2025年审计费18.36万元…）": a period followed by an amount of its own, before the bracket or clause closes, is
    # that amount's period; it labels the earlier amount only when nothing else follows it ("余额130万元（2025年12月31日）")
    tail = re.split(r"[）)。；;]", text[first["end"]:], maxsplit=1)[0]
    if any(sp["start"] >= first["end"] and sp["start"] < first["end"] + len(tail) for sp in text_amounts(text)):
        return None
    return first


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
    # "差异为73元", "20元、1,783元和161元": an amount under 1万元 is written in 元 (the writer is told to write it so, and the refusal of a
    # missing required figure said "in 元 when it is under 1万元"), and the extractor reads only grouped figures (1,783元) as money: a figure
    # written 73元 was no amount, so the required figure was never stated and a series of small amounts was one amount short of its
    # periods. A price per area or per month (元/平方米) is no amount; a zero is read in a 分别 series only.
    taken = [(sp["start"], sp["end"]) for sp in spans]
    for m in re.finditer(r"(?<![\d.,])(\d{1,9}(?:\.\d+)?)\s*元(?![/／每])", work):
        if float(m.group(1)) != 0 and not any(a < m.end() and m.start() < b for a, b in taken):
            spans.append({"value": float(m.group(1)), "written": m.group(0), "start": m.start(), "end": m.end()})
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
    if j >= 0 and re.match(r"[一-鿿A-Za-z]", text[j]) and re.match(r"\s*(?:人民币|人民幣|RMB|CNY)", text[i + 1:i + 8]):
        return ""           # "应付利息-人民币37.2万元": a label and its currency, not a negative amount
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


_UP = ("增加", "增长", "上升", "提高", "上涨", "增至", "升至", "新增", "扩大", "增幅", "调增", "increase", "rose", "grew", "higher")
_DOWN = ("减少", "下降", "降低", "下滑", "减至", "降至", "缩减", "收窄", "降幅", "调减", "decrease", "fell", "declined", "lower")
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
    r"由于|因为|主要系|主要由于|所致|导致|归因于|得益于|受[^，。；]{1,40}影响|系因|源于|源自|"
    r"以致|致使|使得|鉴于|(?<![原起])因(?![素此子])|since |on account of|result of", re.IGNORECASE)
_FORWARD_RE = re.compile(r"预计|预期|有望|expected to|is expected|forecast", re.IGNORECASE)
#: "较上年末减少50.0万元，系正常摊销": a movement explained with a bare 系 is a cause like 主要系. Kept out of _CAUSE_RE, which
#: also finds the why-marker inside one clause for direction binding; this one spans the comma.
_MOVEMENT_CAUSE_RE = re.compile(r"(?:增加|减少|上升|下降|增长|下滑|回落|回升|变动)[^，。；]{0,24}[，,]\s*(?:主要)?系(?!统|列)")
#: the adjusted figure said to equal the book, or no adjustment said to exist
_ADJUSTMENT_DENIED_RE = re.compile(
    r"调整后[^，。；]{0,24}(?:与|和|同)[^，。；]{0,10}(?:账面|调整前|未经?调整|报表|管理层)[^，。；]{0,8}(?:一致|相同|相等|无差异|没有差异)|"
    r"(?:无|未作|未发现|不存在|没有)(?:需|须|任何)?(?:示意性|审计)?调整(?:事项)?(?![前后])")
#: an item put in a period range: "为应收2025年10-12月的水电费"
_ITEM_PERIOD_RE = re.compile(r"(?:为|系|均为|全部为|属于)\s*(?:应收|应付|预收|预付)?\s*(\d{4})年\s*(\d{1,2})(?:\s*[-－—至~]\s*(\d{1,2}))?\s*月(?:份)?\s*的?\s*([\u4e00-\u9fff]{2,4})")
#: a threshold in years on an ageing or a maturity: "账龄超1年", "到期时间为1年以上"
_TERM_SPAN_RE = re.compile(r"(账龄|到期(?:时间|期限|日)?|剩余期限|期限)\s*(?:为|在|均)?\s*(超过|超出|超|大于|多于|长于|小于|少于|短于|不足|未满|未超过)?"
                           r"\s*(一|二|两|三|1|2|3)\s*年\s*(以上|以内|以下|之内)?")
_YEARS = {"一": 1, "二": 2, "两": 2, "三": 3, "1": 1, "2": 2, "3": 3}
#: a sentence about the databook itself rather than the company: a hidden tab, a workpaper not updated
_WORKBOOK_META_RE = re.compile(r"隐藏(?:的)?(?:工作表|表|页)|(?:工作)?底稿未(?:及时)?更新|明细表未更新")
#: an adjustment given as the reason: "主要系管理层调整调增600.0万元所致"
_ADJUSTMENT_CAUSE_RE = re.compile(r"(?P<why>主要系|主要因|主要由于|由于|系|因)[^，。；]{0,6}?(?:管理层|审计|示意性)?调整(?:调增|调减|增加|减少)?")
_MOVE_WORD_RE = re.compile(r"(?:增加|减少|上升|下降|增长|下滑)\s*(?:了|约)?\s*$")


def _plain_number(written: str) -> Optional[float]:
    try:
        return float(re.sub(r"[,\s]", "", written).replace("−", "-").replace("－", "-"))
    except ValueError:
        return None


def _term_spans(text: str) -> set:
    """(ageing or maturity, more or less than, years) for each threshold a text states; one with no direction is skipped."""
    out = set()
    for m in _TERM_SPAN_RE.finditer(text or ""):
        word, tail = m.group(2) or "", m.group(4) or ""
        more = word in ("超过", "超出", "超", "大于", "多于", "长于") or (not word and tail == "以上")
        less = word in ("小于", "少于", "短于", "不足", "未满", "未超过") or (not word and tail in ("以内", "以下", "之内"))
        if more or less:
            out.add(("age" if m.group(1) == "账龄" else "due", ">" if more else "<", _YEARS[m.group(3)]))
    return out
_ANNUAL_WORD = ("年化", "annualis", "annualiz")
_FORECAST_WORD = ("预计全年", "全年预测", "全年预计", "预计全年度", "full-year forecast", "expected full-year", "forecast")
_ANNUAL_ANY = _ANNUAL_WORD + _FORECAST_WORD + ("全年", "full year", "full-year")
#: share of a cause sentence's content bigrams that must be in the notes it cites for it to read as the note's own statement
NOTE_FAITHFUL = 0.6

_GENERIC_LINE = re.compile(r"^(?:原值|净值|账面价值|账面净值|累计折旧|累计摊销|减值准备|坏账准备|直线法|合计|小计|总计|其他|其中|"
                           r"本金|利息|余额|期初|期末|本年增加|本年减少|本期增加|本期减少)$")
_ADJUST_WORDS = ("调增", "调减")
_ORG = re.compile(r"[一-鿿A-Za-z（）()·]{2,40}?(?:股份有限公司|有限责任公司|有限公司|分公司|公司|银行|事务所|集团)")
_ORG_END = re.compile(r"(?:股份有限公司|有限责任公司|有限公司|分公司|公司|银行|事务所|集团)$")
_ORG_GENERIC = re.compile(r"^(?:目标|本|该|母|子|关联|项目|总|分|上市|贵|我|其|一家|两家|多家|各|上述|以上|该等)?(?:公司|银行|集团|事务所)$")


def _org_core(name: str) -> str:
    """A company name for comparison: no brackets or their contents, no spaces."""
    return re.sub(r"[（(][^）)]*[）)]|\s|[·*＊'‘’]", "", str(name or ""))


def names_not_in_evidence(text: str, contract: Any) -> List[str]:
    """Company, bank and firm names the sentence writes that the account's evidence (its lines, notes and tables, the related
    accounts' included) does not contain. A model wrote one insurer's name for the other insurer the tab names, and nothing
    checked a name (client file, 2026-11-09). A name is accepted when it is in the evidence, or a company name from the
    evidence is inside it (the sentence's own words run into the match: "其中应付甲公司")."""
    pool_text = " ".join([str(f.get("line") or "") for f in contract.facts] + [str(n.get("text") or "") for n in contract.notes]
                         + [str(f.get("table_title") or "") for f in contract.facts])
    pool = _org_core(pool_text)
    known = {_org_core(m.group(0)) for m in _ORG.finditer(pool_text)}
    known |= {_org_core(_bare(f.get("line"))) for f in contract.facts if _ORG_END.search(_bare(f.get("line")))}
    known = {k for k in known if len(k) >= 3 and not _ORG_GENERIC.match(k)}
    out = []
    for m in _ORG.finditer(str(text or "")):
        # the sentence's own words run into the match: "原值系支付给上海甲公司", "（含乙公司": cut at the last of them
        # "货币资金中银行存款为主要项目": 中 after an account word is the sentence's own, not the start of a bank's name
        raw = re.split(r"支付给|付给|支付|应付|应收|预付|预收|系|为|含|包括|其中|主要|向|对|与|和|及|给|自|由|于|在|的|、|，|,|（含|\(含|(?<=[金款产债入本用税益项])中",
                       m.group(0))[-1]
        name = _org_core(raw)
        if len(name) < 6 or _ORG_GENERIC.match(name):
            continue
        if name in pool or any(k in name for k in known):
            continue
        out.append(raw)
    return out


_GENERIC = {"金额", "余额", "增加", "减少", "主要", "所致", "导致", "由于", "因为", "本期", "期末", "上升", "下降", "原因",
            "较上", "上年", "增长", "变动", "系因", "主要系", "其中", "包括", "合计", "公司", "款项", "万元", "亿元"}


_CURRENCY_WORDS = (("USD", re.compile(r"美元|美金|USD|US\$", re.I)), ("HKD", re.compile(r"港元|港币|HKD", re.I)),
                   ("EUR", re.compile(r"欧元|EUR", re.I)))


def _currency_at(text: str, start: int, end: int) -> str:
    """The currency an amount at [start, end) is written in: a currency word inside it, right after it, or right before it."""
    # a currency word of the next or the last clause is not this amount's: "银行存款-人民币户1,820.3万元，美元户无余额" read the figure
    # as US dollars and found no cited fact for it (a book of the 2026-11-17 trials, refused three rounds running)
    before = re.split(r"[，,。；;、]", text[max(0, start - 4):start])[-1]
    after = re.split(r"[，,。；;、]", text[end:end + 4])[0]
    window = before + text[start:end] + after
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
    """What a note's text states: amounts (value, currency), percentages, dates. A number with no unit is read only where it is
    plainly yuan (below); a remark's cell values are read with the period they carry."""
    text = str(note.get("text") or "")
    own = note_period(note, periods)
    # a written sign is kept: "ADJ#4：-50.0万元" is a reduction (dropping it failed every adjustment written with its sign)
    amounts = [{"value": sp["value"] * (-1.0 if sp.get("sign") == "-" else 1.0), "currency": _currency_at(text, sp["start"], sp["end"]),
                "written": sp["written"], "period": own} for sp in text_amounts(text)]
    # "183,600+40,000元": the unit written once covers each figure of the list (a fee was refused as uncited, 2026-11-09)
    scale = {"元": 1.0, "千元": 1e3, "万元": 1e4, "亿元": 1e8}
    for m in re.finditer(r"((?:\d[\d,]*(?:\.\d+)?\s*[+＋、和及]\s*)+)(\d[\d,]*(?:\.\d+)?)\s*(亿元|万元|千元|元)", text):
        for figure in re.findall(r"\d[\d,]*(?:\.\d+)?", m.group(1)):
            amounts.append({"value": float(figure.replace(",", "")) * scale[m.group(3)], "currency": "CNY", "written": figure + m.group(3),
                            "period": own})
    # a remark's cell values carry their period ("2026-03-31: 1234567.0"): yuan, like the facts, and bound to that period
    for m in _REMARK_VALUE.finditer(text):
        if m.group(1) in periods:
            amounts.append({"value": float(m.group(3)), "currency": "CNY", "written": m.group(0), "period": m.group(1),
                            "annualised": bool(m.group(2))})
    # "2025审计费183600和2026Q1审计费40000": a whole number of five or more digits written straight after the words, with no
    # unit, is yuan like the facts (a fee was refused as uncited, 2026-11-09). A decimal ("1234.567...") is a cell value in some
    # other unit and is not read; nor is a figure followed by a unit or a count word.
    for m in re.finditer(r"(?:(?<=[\u4e00-\u9fff])|^)(\d{5,10})(?=$|[和及与、，,。；;）)\s])", text):
        if not re.fullmatch(r"(?:19|20)\d{2}", m.group(1)):
            amounts.append({"value": float(m.group(1)), "currency": "CNY", "written": m.group(1), "period": own})
    dates = [m for m in period_mentions(text) if m["kind"] in ("date", "month")]
    # "2026.3.21-3.31": a date written with dots (and the end of a range after it, in the same year)
    for m in re.finditer(r"(20\d{2})[.．/](\d{1,2})[.．/](\d{1,2})(?:\s*[-~～至到]\s*(\d{1,2})[.．/](\d{1,2}))?", text):
        year = int(m.group(1))
        dates.append({"kind": "date", "year": year, "month": int(m.group(2)), "day": int(m.group(3)), "text": m.group(0)})
        if m.group(4):
            dates.append({"kind": "date", "year": year, "month": int(m.group(4)), "day": int(m.group(5)), "text": m.group(0)})
    signs = {w["sign"] for w in direction_words(text)} | {1 if sp.get("sign") == "+" else -1 for sp in text_amounts(text) if sp.get("sign")}
    return {"amounts": amounts, "percents": [abs(sp["value"]) for sp in text_percents(text) if not sp["points"]],
            "dates": dates, "period": own,
            "years": {int(y) for y in re.findall(r"(?<!\d)(20\d{2})(?!\d)", text)}, "signs": signs}


_REMARK_VALUE = re.compile(r"(\d{4}-\d{2}-\d{2})( annualised)?:\s*(-?\d+(?:\.\d+)?)")


_INTERNAL_TERMS = re.compile(r"[（(]\s*[FDN]\d{1,4}(?:\s*[,，、]\s*[FDN]\d{1,4})*\s*[)）]|(?<![A-Za-z])[FDN]\d{1,4}(?![\d])|authority_adjusted|"
                             r"raw_schedule|detail_basis|explain_with_detail|describe_total|(?<![A-Za-z])financials口径")


_LINE_ROLES = (ROLE_COMPONENT, ROLE_SUPPORT)


def _bare(label: Any) -> str:
    # an Excel text prefix ("'*甲公司…") is not part of the name: kept, a sentence naming that tenant was refused as citing
    # none of its figures (client file, 2026-11-09)
    return str(label or "").lstrip("'‘’`*＊ ").strip()


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
        "judged": raw.get("judged") if isinstance(raw.get("judged"), dict) else None,
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
    # "无新增投入", "未增加": a negated direction word says nothing moved (a sentence saying capital had no new contribution was
    # refused as claiming an increase, 2026-11-09 review)
    return [{"start": m.start(), "end": m.end(), "sign": _DIRECTION_WORDS[m.group(0).lower()], "word": m.group(0)}
            for m in _DIRECTION_RE.finditer(text) if not re.search(r"(?:无|未|没有|并无|尚无|不)\s*$", text[max(0, m.start() - 3):m.start()])]


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
                      note_signs: Optional[set] = None, counter_accounts: Iterable[str] = ()) -> List[Tuple[str, str, str, str]]:
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

    spans = text_amounts(text)

    def by_written_amount(w: Dict[str, Any], truths: List[Dict[str, Any]]) -> Optional[set]:
        """The sign a direction word is bound to by the amounts written beside it, when several cited movements disagree:
        the movement written right after it ("2024年度减少13.8万元，2025年度增加6.7万元"), the one written just before a
        percentage word ("减少20.0万元，降幅50.0%"), or from-to figures ("由2.00亿元增加至2.50亿元"). Refused before as
        unbound: a sentence listing a total's movements over several periods (client file, 2026-11-09)."""
        gap = re.match(r"[\s了约为额度幅至到]{0,3}", text[w["end"]:])
        after = next((sp for sp in spans if w["end"] <= sp["start"] <= w["end"] + (gap.end() if gap else 0)), None)
        if after is not None:
            hit = {_sign(t["value"]) for t in truths if abs(abs(t["value"]) - after["value"]) <= after["tolerance"]}
            if hit:
                return hit
            before = [sp for sp in spans if sp["end"] <= w["start"] and w["start"] - sp["end"] <= 16]
            if text[w["end"]:w["end"] + 2].lstrip()[:1] in ("至", "到") and before:
                return {_sign(after["value"] - before[-1]["value"])}
        if text[w["end"]:w["end"] + 4].lstrip()[:1].isdigit() or re.match(r"\s*[为约]?\s*-?\d[\d.,]*\s*%", text[w["end"]:]):
            before = [sp for sp in spans if sp["end"] <= w["start"] and w["start"] - sp["end"] <= 6]
            if before:
                hit = {_sign(t["value"]) for t in truths if abs(abs(t["value"]) - before[-1]["value"]) <= before[-1]["tolerance"]}
                if hit:
                    return hit
        return None

    seen_unsupported = set()
    adjusted = set()
    for i in infos:
        w = i["w"]
        if w["word"] in _ADJUST_WORDS and any(name in text[max(0, w["start"] - 12):w["end"] + 12] for name in counter_accounts):
            # "调增一年内到期的非流动负债，相应调减长期借款": the word is the other account's side of the entry (a needless red
            # mark on the client file, 2026-11-09); that account's paragraph carries its own direction
            adjusted.add(id(i))
            continue
        if w["word"] in _ADJUST_WORDS and note_signs and not i["cause"]:
            # an adjustment's direction is the cited note's (each bridge entry says 调增 or 调减 for this account), not a
            # movement's: "调减" beside a movement of the other sign was refused as unbound (2026-11-09)
            adjusted.add(id(i))
            if w["sign"] not in note_signs:
                out.append(("failed", "SIGN_ERROR", "arithmetic",
                            f"{word_text(w)}, but the cited adjustment note says the opposite for this account"))
    for i in infos:
        w, subject, line = i["w"], i["subject"], i["line"]
        if not subject or id(i) in adjusted:
            continue
        truths = _pick(share_truths if subject == "share" else amount_truths, line)
        signs = {_sign(t["value"]) for t in truths}
        if len(signs) > 1 and subject != "share":
            signs = by_written_amount(w, truths) or signs
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
        if i["subject"] or id(i) in adjusted:
            continue
        truths = _pick(amount_truths, line) + [t for t in share_truths if not line or t["line"] == line]
        signs = {_sign(t["value"]) for t in truths}
        others = [o for o in result_words if o is not i and o["w"]["sign"] != w["sign"]]
        if len(signs) > 1:
            bound = by_written_amount(w, _pick(amount_truths, line))
            if bound:                                # the written amount says which movement: a mismatch is a sign error
                signs, others = bound, []
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


#: A sentence that says the lines it quotes are the tab's, before the audit adjustments, has stated their basis.
_BEFORE_ADJUSTMENT = ("调整前", "未经调整", "未调整", "未含审计调整", "未包含审计调整", "不含审计调整", "未计入审计调整",
                      "before adjustment", "before audit adjustment", "unadjusted")


def _tab_detail_total(contract: EvidenceContract, period: str) -> Optional[float]:
    """The tab's own total for the period when the tab's detail lines add up to it (flat, or its top level where lines sit under
    their own sub-lines): then the lines are what the tab says and nothing else. None when there is no such total or they differ."""
    raw = contract.total(period, BASIS_RAW)
    parts = [f for f in contract.facts if f["period"] == period and f["role"] == ROLE_COMPONENT]
    if raw is None or not parts:
        return None
    flat, top = sum(f["value"] for f in parts), sum(f["value"] for f in parts if not f.get("nested_in"))
    return raw if any(abs(x - raw) <= tie_tolerance(x, raw) for x in (flat, top)) else None


#: "…，其余8项小额项目合计14.6万元": the words before 合计 that make the amount after it the rest of a total, not the sum of what was listed
_REST_BEFORE_TOTAL_RE = re.compile(r"(?:其余|剩余|余下|其他|其它|另外)[^，。；,;（()）]{0,14}$")


_COMPANY_FORM = re.compile(r"股份有限公司|有限责任公司|有限公司|分公司|集团|公司")


def _short_name_before(own: str, text: str, start: int, named: str) -> bool:
    """The clause an amount sits in names the line it cites by part of that line's name ("其中华东星河16.3万元" for the line
    C0229华东星河实业有限公司; "甲银行账户余额为负（-1,204.3万元）" for ZHTC-009甲银行某某分行1006), not by the line named earlier. A piece of
    three characters or more, outside the code and the company form, counts, unless it is the earlier line's own name: a child written
    under its parent's name is the parent's WRONG_LINE (2026-11-17: two right sentences refused three rounds each)."""
    core = _COMPANY_FORM.sub("·", re.sub(r"^[A-Za-z0-9_\-－]+", "", own))
    window = text[:start]
    cut = max((window.rfind(ch) for ch in "，,。；;、"), default=-1)
    clause = window[cut + 1:][-40:]
    for size in range(min(len(core), 8), 2, -1):
        for i in range(len(core) - size + 1):
            piece = core[i:i + size]
            if "·" not in piece and piece in clause and piece not in named and not (named and named in piece):
                return True
    return False


def _lines_put_under_another(text: str, matched: List[Tuple[Dict[str, Any], Dict[str, Any], str]],
                             cited: Optional[List[Dict[str, Any]]] = None) -> set:
    """"其中押金122.3万元，主要为甲41.6万元、乙33.6万元": after a line's figure a marker (其中, 主要为, 包括) and lines that together are no more than
    it put those lines under it. Two right sentences of a trial were refused as a make-up larger than the balance, the parent counted with
    its own children (2026-11-17)."""
    parts = sorted(((span, src) for span, src, _h in matched if span is not None and src.get("role") == ROLE_COMPONENT
                    and src.get("basis") != BASIS_DERIVED), key=lambda x: x[0]["start"])
    under: set = set()
    for i, (span, src) in enumerate(parts[:-1]):
        nxt = parts[i + 1][0]
        if not re.search(r"其中|主要为|主要由|主要是|包括|含", text[span["end"]:nxt["start"]]) or re.search(r"[。；;]", text[span["end"]:nxt["start"]]):
            continue
        group = []
        for sp, sr in parts[i + 1:]:
            if re.search(r"[。；;]", text[span["end"]:sp["start"]]):
                break
            group.append((sp, sr))
        if group and sum(abs(sr["value"]) for _sp, sr in group) <= abs(src["value"]) * 1.05 + span["tolerance"] \
                and all(sr.get("period") == src.get("period") for _sp, sr in group):
            under |= {(sr.get("id") or sr.get("ref")) for _sp, sr in group}
    # a line named with no figure of its own ("主要为仓储服务，其中A16.3万元、B18.9万元"): the lines after its marker are under it when they
    # are no more than it
    for f in cited or []:
        if f.get("role") != ROLE_COMPONENT or not f.get("line") or f.get("basis") == BASIS_DERIVED:
            continue
        name = _bare(f["line"])
        at = text.find(name)
        if at < 0 or any(sp["start"] >= at and sp["end"] <= at + len(name) + 1 for sp, _s in parts):
            continue
        tail = text[at + len(name):]
        marker = re.match(r"[^，,。；;]{0,4}[，,]?\s*(?:其中|包括|主要为|主要是|含)", tail)
        if not marker:
            continue
        stop = re.search(r"[。；;]", tail)
        end = at + len(name) + (stop.start() if stop else len(tail))
        group = [(sp, sr) for sp, sr in parts if at + len(name) + marker.end() <= sp["start"] < end and sr.get("period") == f.get("period")]
        if group and sum(abs(sr["value"]) for _sp, sr in group) <= abs(f["value"]) * 1.05 + 1.0:
            under |= {(sr.get("id") or sr.get("ref")) for _sp, sr in group}
    return under


def _movement_remedy(contract: EvidenceContract, d: Dict[str, Any]) -> str:
    """What a sentence may say instead of a change Python did not compute. A sentence that writes the change and says it cannot be
    described is refused the same way again (three rounds each, 2026-11-17)."""
    reason = str(d.get("invalid_reason") or "")
    periods = d.get("periods") or []
    if "disagree" in reason:
        diff = [x for x in contract.derived if x.get("kind") == KIND_DIFFERENCE and (x.get("period") in periods or set(x.get("periods") or []) & set(periods))]
        return ("; do not write the change in any form (nor that it cannot be described): state the tab's figure and Financials' for each "
                "period and the difference Python computed" + (f" ({', '.join(x['ref'] for x in diff[:3])})" if diff else ""))
    if "differ in length" in reason:
        annual = [x for x in contract.derived if x.get("kind") == KIND_ANNUALIZED and x.get("period") in periods]
        return ("; do not compare a part-year figure with a year in any form (nor say that it is not annualised): state each period's "
                "figure, and the annualised one" + (f" ({', '.join(x['ref'] for x in annual[:2])})" if annual else "") + " with its basis")
    return ""


def _cn_period(period: str) -> str:
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", str(period))
    if not m:
        return str(period)
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    return f"{y}年末" if (mo, d) == (12, 31) else f"{y}年{mo}月{d}日"


def _inside_named_range(text: str, mentions: List[Dict[str, Any]], period: str) -> bool:
    """"2024年12月31日至2026年6月30日，实收资本各期均为9,653.7万元，无变动": a period between the two ends of a range the sentence names is
    named by it (the writer was refused three rounds for not writing 2025年末 as well)."""
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", str(period))
    if not m:
        return False
    target = (int(m.group(1)), int(m.group(2)), int(m.group(3)))

    def low(x: Dict[str, Any]) -> tuple:
        return (x["year"], x.get("month") or 1, x.get("day") or 1)

    def high(x: Dict[str, Any]) -> tuple:
        return (x["year"], x.get("month") or 12, x.get("day") or 31)
    ordered = sorted(mentions, key=lambda x: x["start"])
    return any(re.match(r"^\s*(?:至|到|-|－|—|–|~|及至)\s*$", text[a["end"]:b["start"]]) and low(a) <= target <= high(b)
               for a, b in zip(ordered, ordered[1:]))


def _line_remedy(contract: EvidenceContract, label: str, claim_periods: Iterable[str]) -> str:
    """Which figures the sentence would cite for a line it names: the line's own, else the lines under it, else the periods it has."""
    wanted = set(claim_periods)
    own = [f for f in contract.facts if f["role"] == ROLE_COMPONENT and _bare(f.get("line")) == label and f.get("basis") != BASIS_DERIVED]
    here = [f for f in own if not wanted or f["period"] in wanted]

    def show(items: List[Dict[str, Any]]) -> str:
        return "; ".join(f"{f['ref']} ({f['period']}, {f['value'] / 1e4:,.2f}万元)" for f in items[:4])
    if here:
        return f": cite {show(here)}"
    if own:
        return f": the account has no '{label}' figure for {', '.join(sorted(wanted))} (it has {show(own)}): leave the name out, or cite the lines that make it up"
    under = [f for f in contract.facts if f["role"] == ROLE_COMPONENT and _bare(f.get("line")).startswith(label + "-")
             and (not wanted or f["period"] in wanted)]
    if under:
        return f": no figure of that name exists; the lines under it are {show(under)}: cite those, or leave the name out"
    return ": no line of that name exists in the account; leave it out"


def _bracket_spans(text: str) -> List[Tuple[int, int]]:
    return [(m.start(), m.end()) for m in re.finditer(r"（[^（）]*）|\([^()]*\)", text)]


def _signed(text: str, span: Dict[str, Any]) -> float:
    before = text[max(0, span["start"] - 1):span["start"]]
    return -span["value"] if before in ("-", "−", "－", "负") else span["value"]


def _rest_analysis(text: str, amounts: List[Dict[str, Any]], total_word: "re.Match[str]", rest_word: "re.Match[str]",
                   contract: EvidenceContract, claim_periods: Iterable[str], nested_starts: set) -> Optional[Dict[str, Any]]:
    """"主要包括A、B，其余8项合计X": X is what is left of the period's total after the lines named before 其余. The lines inside a bracket
    ("物管成本73.9万元（其中物业费47.6万元…）") and lines the contract found nested in another are parts of a line already named. Returns
    {stated, parts, period, base, expected, slack}, or None when it cannot be tested (no period, the lines belong to more than one
    period, no total)."""
    spans = sorted(amounts, key=lambda a: a["start"])
    after = [a for a in spans if a["start"] >= total_word.end()]
    if not after:
        return None
    stated, rest_at = after[0], rest_word.start()
    boundary = max([m.end() for m in re.finditer(r"[。；;]", text[:rest_at])] or [0])
    brackets = _bracket_spans(text)
    inside = lambda pos: any(b0 <= pos < b1 for b0, b1 in brackets)               # noqa: E731
    # the lines are the ones after "其中" / "主要为" / "包括" (outside a bracket): an amount before it is the balance, not a part of it
    opener = next((m for m in re.finditer(r"其中|主要为|主要包括|主要由|包括", text[boundary:rest_at]) if not inside(boundary + m.start())), None)
    first = boundary + opener.end() if opener else boundary
    claimed = list(claim_periods)
    mentions = period_mentions(text)

    def period_of(a: Dict[str, Any]) -> Optional[str]:
        mention = attached_period(text, a["start"], a["end"], mentions)
        if mention is not None:
            return next((p for p in sorted(contract.periods) if mention_matches(mention, p)), None)
        return claimed[0] if len(claimed) == 1 else None
    period = period_of(stated)
    if period is None:
        return None
    base = contract.total(period, BASIS_RAW) if _has(text, _BEFORE_ADJUSTMENT) else None
    if base is None:
        base = contract.total(period)
    if base is None:
        return None
    parts = [a for a in spans if first <= a["start"] and a["end"] <= rest_at and a["start"] not in nested_starts and not inside(a["start"])
             and abs(a["value"] - abs(base)) > a["tolerance"] + 1.0]               # an amount equal to the total is the balance
    if not parts or any(period_of(a) != period for a in parts):
        return None
    named = sum(_signed(text, a) for a in parts)
    return {"stated": stated, "parts": parts, "period": period, "base": base, "named": named, "expected": base - named,
            "given": _signed(text, stated),
            "slack": sum(a["tolerance"] for a in parts) + stated["tolerance"] + max(1.0, 0.0005 * abs(base))}


def _stated_rest_problem(text: str, amounts: List[Dict[str, Any]], total_word: "re.Match[str]", rest_word: "re.Match[str]",
                         contract: EvidenceContract, claim_periods: Iterable[str], nested_starts: set) -> Optional[str]:
    """The refusal when the rest a sentence gives is not the remainder of the period's total after the lines it names, else None."""
    a = _rest_analysis(text, amounts, total_word, rest_word, contract, claim_periods, nested_starts)
    if a is None or abs(a["given"] - a["expected"]) <= a["slack"]:
        return None
    return (f"'{a['stated']['written']}' is given as what is left of {a['period']}'s total of {a['base'] / 1e4:,.2f}万元 after the "
            f"{len(a['parts'])} lines named ({a['named'] / 1e4:,.2f}万元), but the total less those lines is {a['expected'] / 1e4:,.2f}万元: "
            "write that figure for the rest (cite the D ref that sums the lines not named, if there is one), or name the lines that make "
            "the difference")


def _prefix_children(sources: List[Dict[str, Any]]) -> set:
    """ids of the lines whose name continues another's with a dash ("管理费用-行政-Leasing" under "管理费用-行政"): its parts, not another line."""
    names = {(s.get("id") or s.get("ref")): _bare(s.get("line")) for s in sources if s.get("line")}
    return {i for i, n in names.items() if any(n != m and n.startswith(m + "-") for m in names.values() if m)}


def _norm_figure(written: Any) -> str:
    return re.sub(r"[\s,，]", "", str(written or ""))


def _line_overlap(line: str, said: str) -> bool:
    a, b = _bare(line), _bare(said)
    if not a or not b:
        return False
    if a in b or b in a:
        return True
    return any(a[i:i + 3] in b for i in range(max(1, len(a) - 2)) if len(a[i:i + 3]) == 3)


def check_reading(claim: Dict[str, Any], reading: List[Dict[str, Any]], contracts: Dict[str, EvidenceContract]) -> Tuple[bool, str]:
    """The agent read a sentence the verifier refused for how it read the prose: for each figure, which cited fact it is, the period
    words and the line the sentence writes it under, and what it is (balance, part, subpart, rest, total, change). Python holds that
    reading to the evidence: every figure is the value of a cited fact (or the rest / a change Python computes), the period the sentence
    writes is that fact's, the line is that fact's, and the arithmetic of the roles holds (parts no more than the balance, subparts no
    more than their part, the rest what is left, a total the sum of its parts). (True, '') clears the reading refusals; the figures'
    own refusals (an amount no fact has, a unit, a sign, a percentage) are never cleared here."""
    account = str(claim.get("account") or "")
    contract = contracts.get(account)
    if contract is None:
        return False, "no contract"
    text = str(claim.get("text") or "")
    spans = text_amounts(text)
    if not spans:
        return False, "the sentence writes no amount"
    by_written = {_norm_figure(r.get("written")): r for r in reading if isinstance(r, dict)}
    declared = list(claim.get("periods") or [])
    cites = {str(c) for c in claim.get("cites") or []}
    parts: Dict[str, List[Tuple[Dict[str, Any], float]]] = {}
    children: Dict[str, List[float]] = {}
    values: Dict[str, float] = {}
    rest = None
    for sp in spans:
        key = _norm_figure(sp["written"])
        r = by_written.get(key)
        if r is None:
            return False, f"the reading has no entry for {sp['written']}"
        ref, role = str(r.get("ref") or ""), str(r.get("role") or "other").lower()
        item = None
        if ref not in ("rest", "movement", "sum"):
            item = contract.fact(ref) or next((d for d in contract.derived if d.get("ref") == ref), None)
            if item is None or ref not in cites:
                return False, f"{sp['written']} is read as {ref or '(no ref)'}, which the sentence does not cite"
            if item.get("value") is None or abs(abs(item["value"]) - sp["value"]) > sp["tolerance"]:
                return False, f"{sp['written']} is read as {ref}, whose value is {_fmt(item.get('value') or 0)} yuan"
            fact_periods = [item.get("period")] if item.get("period") and item.get("basis") != BASIS_DERIVED \
                else ([(item.get("periods") or [None])[-1]] if item.get("kind") == KIND_MOVEMENT else list(item.get("periods") or [item.get("period")]))
            for words in (r.get("under_period"), r.get("under_line")):
                if str(words or "").strip() and _norm_figure(words) not in _norm_figure(text):
                    return False, f"'{words}' is not in the sentence: the words a figure is written under are copied from it"
            said = period_mentions(str(r.get("under_period") or ""))
            if said:
                if not any(mention_matches(said[0], p) for p in fact_periods if p):
                    return False, (f"{sp['written']} is {ref}, a figure of {', '.join(str(p) for p in fact_periods)}; the sentence writes it under "
                                   f"{r.get('under_period')}")
            elif len(declared) > 1:
                return False, f"the sentence writes no period for {sp['written']} and the claim declares {len(declared)}"
            if item.get("role") == ROLE_COMPONENT and item.get("line") and item.get("basis") != BASIS_DERIVED:
                if not _line_overlap(item["line"], str(r.get("under_line") or "")):
                    return False, f"{sp['written']} is {ref} ({item['line']}); the sentence writes it under '{r.get('under_line') or ''}'"
        values[key] = sp["value"]
        period = (fact_periods[0] if item is not None and fact_periods else None) or (declared[-1] if declared else None)
        if role == "part":
            if str(r.get("of") or "").strip():
                children.setdefault(_norm_figure(r.get("of")), []).append(sp["value"])
            else:
                parts.setdefault(str(period), []).append((sp, sp["value"]))
        elif role == "subpart":
            children.setdefault(_norm_figure(r.get("of")), []).append(sp["value"])
        elif role == "rest":
            rest = (sp, str(period))
    for parent, kids in children.items():
        if parent in values and sum(kids) > values[parent] * 1.05 + 1.0:
            return False, f"the figures read as parts of {parent} add up to {_fmt(sum(kids))} yuan, more than it"
    for period, items in parts.items():
        base = contract.total(period)
        if base is None or base <= 0:
            continue
        made = sum(v for _sp, v in items)
        if rest is not None and rest[1] == period:
            slack = sum(sp["tolerance"] for sp, _v in items) + rest[0]["tolerance"] + max(1.0, 0.0005 * abs(base))
            if abs(rest[0]["value"] - (base - made)) > slack:
                return False, f"the rest is {_fmt(rest[0]['value'])} yuan, but the total {_fmt(base)} less the parts {_fmt(made)} is {_fmt(base - made)}"
        elif made > base + max(0.10 * base, 10000.0):
            return False, f"the parts add up to {_fmt(made)} yuan, more than the balance of {_fmt(base)}"
    return True, ""


def verify_claim(raw_claim: Dict[str, Any], contracts: Dict[str, EvidenceContract], *, default_account: str = "",
                 index: int = 0) -> Dict[str, Any]:
    claim = normalize_claim(raw_claim, default_account, index)
    account, text = claim["account"], claim["text"]
    contract = contracts.get(account)
    findings: List[ClaimFinding] = []
    checked: set = set()
    cleared = claim["judged"] if (claim.get("judged") or {}).get("verdict") == "faithful" else None

    def fail(code: str, dim: str, message: str) -> None:
        if cleared is not None and code in JUDGEABLE:      # the reading was done by the agent and checked against the evidence (check_reading)
            findings.append(ClaimFinding("BINDING_JUDGED", dim, "info", f"{code}: {message[:200]} (read by the agent, checked by Python)"))
            return
        findings.append(ClaimFinding(code, dim, "failed", message))

    def info(code: str, dim: str, message: str) -> None:
        findings.append(ClaimFinding(code, dim, "info", message))

    def review(code: str, dim: str, message: str) -> None:
        """Could not be tested. The claim is not failed and not verified: it waits for a person."""
        if cleared is not None and code in JUDGEABLE:
            findings.append(ClaimFinding("BINDING_JUDGED", dim, "info", f"{code}: {message[:200]} (read by the agent, checked by Python)"))
            return
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
    # "按财务口径为-240.4万元，与底稿差异240.4万元" citing the difference: the two totals it is the difference of are its inputs, and the
    # formula shows them; a sentence that quotes them is not made to cite them as well (2026-11-17, refused three rounds)
    by_id = {f.get("id"): f for f in contract.facts}
    for d in cited_derived:
        if d.get("kind") in (KIND_MOVEMENT, KIND_DIFFERENCE, KIND_ANNUALIZED):
            for iid in d.get("inputs") or []:
                f = by_id.get(iid)
                if f is not None and f.get("value") is not None and f.get("role") == ROLE_TOTAL:
                    allowed.append((abs(f["value"]), f, "input_of_cited"))
    by_period: Dict[str, List[Dict[str, Any]]] = {}
    for f in cited_facts:
        if f["role"] == ROLE_COMPONENT:
            by_period.setdefault(f["period"], []).append(f)
        elif f["role"] == ROLE_SUPPORT and not f.get("subtotal"):
            by_period.setdefault(f"{f['period']}|{f.get('table')}", []).append(f)     # one table's lines, never mixed with another's
    for group, parts in by_period.items():
        period = group.split("|")[0]
        # every cited line of a period, and any two or three of them: "土地使用权净值约1.21亿元（一期6105.03万元、二期5999.69万元）"
        # cites the houses of the same table too, and the whole sum was the only one offered (a needless red mark, 2026-11-09)
        subsets = [parts] if len(parts) >= 2 else []
        if 3 <= len(parts) <= 12:
            subsets += [list(c) for k in (2, 3) for c in itertools.combinations(parts, k) if len(c) < len(parts)]
        for chosen in subsets:
            allowed.append((abs(sum(p["value"] for p in chosen)), {"id": "sum(" + ",".join(p["ref"] for p in chosen) + ")",
                                                                  "period": period, "account": account, "line": "", "basis": BASIS_RAW,
                                                                  "role": "sum", "ref": "sum", "value": sum(p["value"] for p in chosen)},
                            "sum_of_cited"))
    for ref_n, figs in note_facts.items():
        for a in figs["amounts"]:
            allowed.append((abs(a["value"]), {"id": f"note:{ref_n}", "ref": ref_n, "account": account, "role": "note", "basis": "note",
                                              "line": "", "period": a.get("period"), "currency": a["currency"], "value": a["value"],
                                              "annualised": bool(a.get("annualised"))}, "note"))
    # "其余8项合计X": Python computes the rest, as it computes the change between two cited lines, so the writer need not cite a figure
    # that no fact carries (two right sentences of a trial were refused AMOUNT_UNMATCHED for writing the remainder the verifier asked for)
    _tw = re.search(r"合计|共计|总计|合共", text)
    _rw = _REST_BEFORE_TOTAL_RE.search(text[:_tw.start()]) if _tw else None
    if _tw and _rw:
        _ra = _rest_analysis(text, amounts, _tw, _rw, contract, claim_periods, set())
        if _ra is not None:
            allowed.append((abs(_ra["expected"]), {"id": "rest", "ref": "rest", "role": "rest", "account": account, "period": _ra["period"],
                                                  "line": "", "basis": BASIS_DERIVED, "value": _ra["expected"], "slack": _ra["slack"]},
                            "implicit_rest"))
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
        # the rest of lines written rounded is the total less their rounded figures, off by what the rounding takes (214.4万 written 21.4万
        # after 1.9万 and 1,628元 named, the lines being 1.93万 and 1,628元)
        hits = [(v, s, how) for v, s, how in allowed
                if abs(v - target) <= max(span["tolerance"], s.get("slack", 0.0) if how == "implicit_rest" else 0.0)
                and (s.get("currency") or "CNY") == currency]
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

            def fit(h: Tuple[float, Dict[str, Any], str]) -> Tuple[bool, bool, bool, bool]:
                src = h[1]
                periods = set(src.get("periods") or ([src["period"]] if src.get("period") else []))
                is_change = src.get("role") == "movement" or src.get("kind") == KIND_MOVEMENT
                return (bool(named) and _bare(src.get("line")) == named, as_change == is_change,
                        mention is not None and any(mention_matches(mention, p) for p in periods),
                        not span.get("sign") or ((span["sign"] == "-") == ((src.get("value") or 0) < 0)))
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
        # a figure within 1% of an allowed one is a miss by a hand-rounded step, not a scale slip: a subset of cited lines that happened to
        # sum to ten times the figure called a 28.0万 for 27.94万 a unit error three rounds running (2026-11-17)
        near_allowed = any(0 < abs(v - target) <= 0.01 * target for v, _s, _h in allowed)
        scaled = None if near_allowed else next(((v, s) for v, s, how in allowed if v and how != "sum_of_cited" and any(
            abs(target - v * 10 ** k) <= max(span["tolerance"], abs(target) * 0.01) for k in (-4, -3, -2, -1, 1, 2, 3, 4))), None)
        if scaled:
            factor = 10 ** round(math.log10(max(target, 1e-9) / max(scaled[0], 1e-9)))
            fail("UNIT_ERROR", "amounts",
                 f"'{span['written']}' is {factor:g}x the cited {scaled[1].get('ref')} "
                 f"({_fmt(scaled[0])} yuan): a unit or scale slip")
            continue
        equal = [o for o in others if o.get("value") is not None and abs(abs(o["value"]) - target) <= span["tolerance"]]
        here = attached_period(text, span["start"], span["end"], period_mentions(text))
        # the same figure sits in other periods too: point at the one in the period the sentence writes it under, else a declared one
        equal.sort(key=lambda o: (not (here is not None and o.get("period") and mention_matches(here, o["period"])),
                                  o.get("period") not in claim_periods))
        elsewhere = equal[0] if equal else None
        stated_in = next((n for n in contract.notes if any(abs(abs(a["value"]) - target) <= span["tolerance"]
                                                          and (a.get("currency") or "CNY") == currency
                                                          for a in note_figures(n, contract.periods)["amounts"])), None)
        if elsewhere is not None or stated_in is not None:
            where = (f"it equals {elsewhere.get('ref')} ({elsewhere.get('account')} {elsewhere.get('period')} {elsewhere.get('basis')} "
                     f"{elsewhere.get('line') or elsewhere.get('kind') or 'total'}), which is not cited") if elsewhere is not None else ""
            note = (f"note {stated_in.get('ref')} states it ({str(stated_in.get('text') or '')[:70]!r}), and is not cited either"
                    if stated_in is not None else "")
            fail("AMOUNT_ELSEWHERE", "amounts",
                 f"'{span['written']}' is not what the cited facts say; " + "; ".join(x for x in (where, note) if x)
                 + ("; cite the one the sentence means" if where and note else ""))
        else:
            # A figure the writer made by hand from a rounded one (176.0万 doubled: 352.0万, where Python's annualised figure is 351.9万)
            # was refused the same way three times running; the message did not say that the figure it wanted was in the evidence.
            near = min((o for o in others if o.get("value") and 0 < abs(abs(o["value"]) - target) <= 0.05 * target),
                       key=lambda o: abs(abs(o["value"]) - target), default=None)
            hint = ""
            if near is not None:
                hint = (f"; the nearest figure in EVIDENCE is {near.get('ref')} ({_fmt(abs(near['value']))} yuan, "
                        f"{near.get('line') or near.get('kind') or 'total'}, {near.get('period') or '/'.join(near.get('periods') or [])}): "
                        "write that figure as it is given, and cite it, instead of a figure made by hand")
            else:
                made = min(((v, s) for v, s, how in allowed if how in ("implicit_movement", "implicit_rest") and 0 < abs(v - target) <= 0.01 * target),
                           key=lambda x: abs(x[0] - target), default=None)
                if made is not None:
                    what = "the change between the two cited figures" if made[1].get("role") == "movement" else "what is left of the total"
                    hint = (f"; {what} is {_fmt(made[0])} yuan ({made[0] / 1e4:,.2f}万元): write that figure, not one made by hand "
                            "from rounded ones")
            fail("AMOUNT_UNMATCHED", "amounts",
                 f"'{span['written']}' ({_fmt(target)} yuan) is not accounted for by any cited fact" + hint)

    # -- a note's figure where the tab has the exact one -----------------------------------
    # "2023年在建工程余额1.87亿元" from a note whose author rounded 187,654,321 down: the tab's figure is 1.88亿元 (four sentences
    # across the arms of one trial, 2026-11-09). Held to the tab only for the period the sentence writes it under.
    for span, src, how in matched_sources:
        if how != "note" or (src.get("currency") or "CNY") != "CNY":
            continue
        mention = attached_period(text, span["start"], span["end"], period_mentions(text))
        if mention is None:
            continue
        exact = next((f for f in contract.facts if f.get("value") and mention_matches(mention, f["period"])
                      and abs(abs(f["value"]) - span["value"]) <= max(2 * span["tolerance"], 1.0)      # one unit of the last digit written
                      and abs(abs(f["value"]) - span["value"]) > span["tolerance"]), None)
        if exact is not None and not any(abs(abs(f["value"]) - span["value"]) <= span["tolerance"] for f in contract.facts
                                         if f.get("value") and mention_matches(mention, f["period"])):
            fail("NOTE_FIGURE_ROUNDED", "amounts",
                 f"'{span['written']}' is the note's rounding; the tab gives {_fmt(exact['value'])} yuan ({exact['ref']}, "
                 f"{exact['period']}): write the tab's figure")

    # -- gain or loss -----------------------------------------------------------------------
    # In an expense account a 损益 line is a loss when positive: "汇兑损益收益530.0万元" for +5,300,000 in 财务费用 was a loss
    # called a gain, and "收益-310.0万元" for a net gain the reverse (2026-11-09 review). A written minus turns the word.
    expense = bool(re.search(r"成本|费用|支出|损失", account)) and not re.search(r"收入|收益|待摊|预提|应付|预付", account)
    for span, src, how in matched_sources:
        if not expense or "损益" not in str(src.get("line") or "") or not src.get("value"):
            continue
        before = text[max(0, span["start"] - 14):span["start"]]
        before = before[max(before.rfind("，"), before.rfind("、"), before.rfind("；")) + 1:]
        says = "gain" if "收益" in before and "损失" not in before else ("loss" if "损失" in before and "收益" not in before else "")
        if not says:
            continue
        if span.get("sign") == "-":
            says = "loss" if says == "gain" else "gain"
        truth = "loss" if src["value"] > 0 else "gain"
        if says != truth:
            fail("GAIN_LOSS_REVERSED", "amounts",
                 f"'{span['written']}' is called a {'gain' if says == 'gain' else 'loss'}, but in this expense account "
                 f"{src.get('ref')} ({_fmt(src['value'])} yuan) is a net {truth}")

    # -- the rest ----------------------------------------------------------------------------
    # "主要包括A、B、C，其余6项合计0.13万元": the rest a D ref sums is what the evidence view did not show. With shown lines
    # left unnamed, the named lines and that rest fall short of the total and "其余" is false (three to four sentences a
    # trial on the client file, 2026-11-09).
    if re.search(r"其余|剩余|余下|其他\s*\d+\s*项|其他各项", text):
        for d in cited_derived:
            if d.get("kind") != KIND_SUM or d.get("value") is None or not d.get("inputs"):
                continue
            inputs = set(d["inputs"])
            named = [f for f in cited_facts if f["role"] == ROLE_COMPONENT and f["period"] == d.get("period")
                     and not f.get("nested_in") and f["id"] not in inputs]
            total = _total_fact(contract, d.get("period"), BASIS_RAW) or _total_fact(contract, d.get("period"))
            if not named or total is None:
                continue
            gap = total["value"] - sum(f["value"] for f in named) - d["value"]
            if abs(gap) > max(0.005 * abs(total["value"]), 1000.0):
                # the lines shown and not named are what the gap is made of (a credit balance among debits, a small line): name them
                cited_ids = {f["id"] for f in cited_facts}
                unnamed = sorted((f for f in contract.facts if f["role"] == ROLE_COMPONENT and f["period"] == d.get("period")
                                  and not f.get("nested_in") and f["id"] not in inputs and f["id"] not in cited_ids
                                  and abs(f.get("value") or 0) >= 100), key=lambda f: -abs(f["value"]))
                listed = "; ".join(f"{f['ref']} {_bare(f['line'])[:24]} {_fmt(f['value'])}" for f in unnamed[:4])
                fail("REST_INCOMPLETE", "arithmetic",
                     f"{d['ref']} sums only the lines the evidence did not show; with the {len(named)} lines named it leaves "
                     f"{_fmt(gap)} yuan of {d.get('period')}'s total unaccounted for: name the other lines or call the rest by "
                     "what it is" + (f". Lines shown and not named: {listed}" if listed else ""))

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
    # "占应付账款余额的80.0%": a cited line over a cited total of the same period is arithmetic on the citations (five right
    # shares were refused in one paragraph, 2026-11-09)
    totals_by_period = {f["period"]: f for f in cited_facts if f["role"] == ROLE_TOTAL and f.get("value")}
    for f in cited_facts:
        t = totals_by_period.get(f["period"])
        if f["role"] in _LINE_ROLES and t is not None and f is not t:
            share = f["value"] / t["value"] * 100
            pct_candidates.append({"pct": abs(share), "signed": share, "kind": "share", "periods": [f["period"]],
                                   "line": f["line"], "ref": f"share({f['ref']},{t['ref']})", "item": None, "base": None, "end": None})
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
                 + (f" (they give {shown})" if shown else " (no ratio is cited)")
                 + "; a ratio of two other figures is not one Python derived: cite the derived ratio, or write the two amounts and leave "
                   "the percentage out")
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
            fail("MOVEMENT_NOT_PERMITTED", "binding", f"{d['ref']}: {d.get('invalid_reason')}{_movement_remedy(contract, d)}")

    # -- binding: periods -------------------------------------------------------------------
    note_dates = [d for figs in note_facts.values() for d in figs["dates"]]
    for mention in period_mentions(text):
        if any((d["year"], d.get("month"), d.get("day")) == (mention["year"], mention.get("month"), mention.get("day")) for d in note_dates):
            continue                                   # a date the cited note itself states (an event, not a reporting period)
        if mention["kind"] == "year" and any(mention["year"] in figs["years"] for figs in note_facts.values()):
            continue                                   # a year the note names ("2025审计费"), not a period the claim reports
        if claim_periods and not any(mention_matches(mention, p) for p in claim_periods):
            have = next((p for p in contract.periods if mention_matches(mention, p)), None)
            total = _total_fact(contract, have) if have else None
            way = (f": the account has a {have} figure ({total['ref']}, {total['value'] / 1e4:,.2f}万元), so add {have} to the claim's "
                   f"periods and cite {total['ref']}, or take the period out of the sentence" if total is not None
                   else ": the evidence has no figure for it, take it out of the sentence" if have is None else "")
            fail("WRONG_PERIOD", "binding",
                 f"the text names '{mention['text']}', which is not among the periods this claim declares ({', '.join(sorted(claim_periods))}){way}")
        elif not claim_periods:
            fail("PERIODS_UNDECLARED", "binding", "the sentence names a period but the claim declares none")
            break
    # A movement into an earlier period, written without that period, reads as the change to the latest balance: "较2024年12月
    # 31日减少120.00万元" after the March 2026 balance quoted the 2025-on-2024 change (client file, 2026-11-09).
    mentioned = period_mentions(text)
    latest_period = contract.periods[-1] if contract.periods else None
    for d in cited_derived:
        end = (d.get("periods") or [None])[-1]
        if (d.get("kind") == KIND_MOVEMENT and end and latest_period and end != latest_period
                and _has(text, _UP + _DOWN + _CHANGE_CUES) and not any(mention_matches(m, end) for m in mentioned)
                and not _inside_named_range(text, mentioned, end)):
            begin = (d.get("periods") or [None])[0]
            example = f" (e.g. {_cn_period(end)}较{_cn_period(begin)}…)" if begin else ""
            fail("MOVEMENT_END_UNSTATED", "binding",
                 f"{d['ref']} is the change into {end}, but the sentence does not name {end}: a reader takes it as the change to the "
                 f"latest balance; name the period it ends in{example}")
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
            # "较2025年末减少187.6万元，主要系当期净利润-187.6万元": 当期 and the change from an earlier period's end are the latest period
            latest_said = (re.search(r"(?:当期|本期|本年度?|报告期|期内)[^，,。；;]{0,8}$", text[:span["start"]])
                           or any(m["end"] <= span["start"] and re.search(r"(?:较|比|相比|相较|对比)\s*$", text[max(0, m["start"] - 3):m["start"]])
                                  and _has(text[m["end"]:span["start"]], _UP + _DOWN) and "。" not in text[m["end"]:span["start"]]
                                  for m in mentions))
            if latest_said and source.get("period") == latest_declared:
                continue
            if len(claim_periods) > 1 and source.get("role") in (ROLE_TOTAL, ROLE_COMPONENT, ROLE_SUPPORT, "sum"):
                review("PERIOD_UNBOUND", "binding",
                       f"'{span['written']}' is not written under any period and the claim declares {len(claim_periods)}; "
                       f"put the period next to each amount, or write one claim per period ({source.get('ref')} is the {', '.join(sorted(src_periods))} figure)")
            continue
        if any(mention_matches(mention, p) for p in src_periods):
            continue
        between = text[mention["end"]:span["start"]]
        if re.search(r"[至达]", between) and source.get("period") == latest_declared:
            continue                                    # "由2024年末的100万元增至130万元": the amount after 至 is the later period
        if re.search(r"[由从]\s*(?:的)?\s*$", text[max(0, span["start"] - 4):span["start"]]) and \
                re.match(r"\s*(?:增加|减少|上升|下降|回升|回落|下滑|增|降|减|升)[至到]", text[span["end"]:span["end"] + 6]) and \
                any(mention_matches(mention, p) and p > str(source.get("period") or "") for p in contract.periods):
            continue                                    # "2024年由2.50亿元增至4.80亿元": the amount after 由 is where the year began
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
    # "其中A-60.0万元，B-9.0万元" naming every line is the same statement (every trial on the client file stated retained
    # earnings this way after the adjusted balance, 2026-11-09/10); a sentence that says the lines are before the adjustments
    # ("（未经调整）") states it rightly and was refused before
    if _has(text, ("构成", "组成", "其中", "consist", "made up of", "comprise", "of which")) and \
            not _has(text, _BEFORE_ADJUSTMENT):
        by_p: Dict[str, List[Dict[str, Any]]] = {}
        for f in cited_facts:
            if f["role"] == ROLE_COMPONENT:
                by_p.setdefault(f["period"], []).append(f)
        for period, parts in by_p.items():
            every = [f for f in contract.facts if f["period"] == period and f["role"] == ROLE_COMPONENT]
            adjusted = contract.total(period, BASIS_ADJUSTED)
            if adjusted is None or len(parts) < len(every):
                continue
            if (contract.permissions.get(period) or {}).get("detail_basis") == "ambiguous":
                continue            # the lines miss the total because some contain others, not because of an adjustment: nothing to say about the basis
            gap = adjusted - sum(f["value"] for f in parts)
            if abs(gap) > max(1.0, 0.0005 * abs(adjusted)):
                fail("COMPOSITION_BASIS", "binding",
                     f"{period}: the lines named are the tab's own and add up to {_fmt(adjusted - gap)} yuan, not to the balance of "
                     f"{_fmt(adjusted)} (a difference of {_fmt(gap)}, the adjustments): say they are before the adjustments, or do not "
                     "call them its make-up")

    # A make-up cannot be larger than what it makes up. "余额主要为仓库建造成本及土地成本，另有预提成本" over a balance of 4.86万 named lines of
    # 24,866万 and 5,850万: the lines were the tab's, before an adjustment that took most of the account away (a trial on four books,
    # 2026-11-17, found two such sentences unmarked). Lines nested in another are parts of it and are not counted; credit balances and
    # accumulated depreciation count with their sign, so naming them as well makes the sum right.
    if _has(text, ("主要为", "主要由", "主要构成", "构成", "组成", "其中", "包括")) and not re.search(r"其余|剩余|余下", text):
        named_lines: Dict[str, List[Dict[str, Any]]] = {}
        # the lines the sentence writes a figure for, and the lines it cites: "余额主要为仓库建造成本及土地相关成本，另有预提成本4.9万元"
        # writes one figure and cites three lines of 24,866万, 5,850万 and 4.86万 over a balance of 4.86万
        every = [(None, f, "cited") for f in cited_facts] + list(matched_sources)
        kids = _prefix_children([src for _s, src, _h in every if src.get("role") == ROLE_COMPONENT])
        in_bracket = {(src.get("id") or src.get("ref")) for span, src, _h in matched_sources
                      if span is not None and any(b0 <= span["start"] < b1 for b0, b1 in _bracket_spans(text))}
        in_bracket |= _lines_put_under_another(text, matched_sources, cited_facts)
        for _span, src, _how in every:
            ident = src.get("id") or src.get("ref")
            if (src.get("role") == ROLE_COMPONENT and src.get("basis") != BASIS_DERIVED and not src.get("nested_in")
                    and ident not in kids and ident not in in_bracket and src.get("period")):
                bucket = named_lines.setdefault(src["period"], [])
                if all(ident != (x.get("id") or x.get("ref")) for x in bucket):
                    bucket.append(src)
        for period, parts in sorted(named_lines.items()):
            if len(parts) < 2:
                continue
            base = (_total_fact(contract, period, BASIS_RAW) or {}).get("value") if _has(text, _BEFORE_ADJUSTMENT) else None
            if base is None:
                base = contract.total(period)
            if base is None or base <= 0:
                continue
            made = sum(f["value"] for f in parts)
            if made > base + max(0.10 * base, 10000.0):
                tab = contract.total(period, BASIS_RAW)
                own = (f" (the tab's own total for {period}, before the adjustments, is {tab / 1e4:,.2f}万元: lines like these are from it)"
                       if tab is not None and abs(tab - base) > max(1.0, 0.0005 * abs(base)) else "")
                fail("COMPOSITION_EXCEEDS_BALANCE", "binding",
                     f"{period}: the {len(parts)} lines named add up to {made / 1e4:,.2f}万元, more than the balance of {base / 1e4:,.2f}万元{own}, "
                     "so they cannot be what it is made of. Lines that offset them are not named (credit balances, accumulated depreciation, "
                     "an adjustment), or they are the tab's own before an adjustment: name the offsetting lines too, say they are the tab's "
                     "own (工作簿明细（未经调整）), or name fewer lines and do not call them its make-up")

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
            tab_total = _tab_detail_total(contract, source["period"]) if basis == "untied" else None
            if basis == "untied" and tab_total is not None and _has(text, _BEFORE_ADJUSTMENT):
                # The deck's total replaced the tab's and the lines were not updated: they add up to the tab's own total, and the
                # sentence says that is what they are. (The composition rule asks for exactly this wording; refusing it here too
                # left a trial's sentence unfixable, 2026-11-17.)
                if not any(f.code == "DETAIL_BEFORE_ADJUSTMENT" and source["period"] in f.message for f in findings):
                    info("DETAIL_BEFORE_ADJUSTMENT", "binding",
                         f"{source['period']}: the lines are the tab's own and add up to its own total ({_fmt(tab_total)} yuan), which the "
                         "deck replaced; the sentence says they are before the adjustments")
            elif basis in ("untied", "none"):
                message = (f"the tab's detail for {source['period']} does not add up to the account total; it may not be used to explain it")
                if tab_total is not None:
                    message += (f". The lines are the tab's own and add up to its total of {_fmt(tab_total)} yuan: they may be quoted "
                                "as the workbook's detail before the adjustments (write it, e.g. 工作簿明细（未经调整）), not as what makes "
                                "up the account's balance")
                else:
                    message += (". The lines add up to no total of the tab's either: do not list them; write the account's balance, the difference "
                                "from the tab where there is one, and say the detail is not provided (明细未提供)")
                if not any(f.code == "DETAIL_NOT_PERMITTED" and f.message == message for f in findings):
                    fail("DETAIL_NOT_PERMITTED", "binding", message)
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
                               or any(own[-k:] in text[max(0, span["start"] - 24):span["start"]] for k in range(4, min(len(own), 12) + 1))
                               or _short_name_before(own, text, span["start"], named or ""))
        if named and source.get("role") == ROLE_COMPONENT and source.get("line") and _bare(source["line"]) != named and not written_own:
            own_line = next((f for f in contract.facts if f["role"] == ROLE_COMPONENT and f["period"] == source.get("period")
                             and _bare(f.get("line")) == named and f["basis"] == source.get("basis")), None)
            fail("WRONG_LINE", "binding",
                 f"the sentence attaches '{span['written']}' to '{named}' but cites the '{source['line']}' figure"
                 + (f"; '{named}' is {own_line['ref']} ({_fmt(own_line['value'])} yuan): cite that for it, and cite the lines under it separately"
                    if own_line is not None else ""))
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
        if _GENERIC_LINE.match(label):
            continue        # 原值, 直线法: words a sentence uses for what they mean, not a line it must cite (2026-11-09)
        position = masked.find(label)
        if position >= 0:
            masked = masked.replace(label, " " * len(label))
            attached = any(0 <= start - (position + len(label)) <= 14 for start in figure_starts)
            as_cause = _name_on_cause_side(text, position, position + len(label))
            about_note = any(label in str(n.get("text") or "") for n in source_notes)
            has_figure = label in cited_lines or any(l.endswith("-" + label) or l.endswith("－" + label) for l in cited_lines)
            if not has_figure:           # "主要为银行存款-人民币户，其中A账户…，B账户…": the group is named and its accounts are cited
                group = {f.get("ref") for f in contract.facts if f["role"] == ROLE_COMPONENT and _bare(f.get("line")) == label}
                has_figure = any(f.get("nested_in") in group for f in cited_facts)
            if not has_figure and not about_note and (attached or (line_level and not as_cause)):
                fail("LINE_NOT_CITED", "binding", f"the sentence names '{label}' but cites none of its figures" + _line_remedy(contract, label, claim_periods))
    if _has(text, _NIL):
        # the period a nil word is about is the one written before it: "2024年12月31日应付账款余额为0元，2025年12月31日增至281.1万元" calls
        # 2024 nil, and refusing it for 2025 (a Financials period) was a refusal of a right sentence
        about = set()
        all_mentions = period_mentions(text)
        for word in _NIL:
            for at in (m.start() for m in re.finditer(re.escape(word), text, re.I)):
                near = [m for m in all_mentions if m["end"] <= at and "。" not in text[m["end"]:at]]
                if near:
                    last = max(near, key=lambda m: m["end"])
                    about |= {p for p in (claim_periods or contract.periods) if mention_matches(last, p)}
                else:
                    about |= set(claim_periods or contract.periods)
        for p in sorted(about):
            if p in contract.permissions and not contract.permissions[p].get("state_nil_or_first_arose", True):
                total = contract.total(p)
                tf = _total_fact(contract, p)
                diff = next((x for x in contract.derived if x.get("kind") == KIND_DIFFERENCE and p in (x.get("periods") or [x.get("period")])), None)
                way = (f": the deck's balance for {p} is {total / 1e4:,.2f}万元" + (f" ({tf['ref']})" if tf else "")
                       + "; state that balance" + (f" and the difference from the tab ({diff['ref']})" if diff else "")
                       + ", and do not say the account was nil or first arose" if total is not None else "")
                fail("NIL_NOT_PERMITTED", "binding",
                     f"{p}: the deck's total for this period comes from Financials; the account may not be called nil / first arisen from the tab{way}")
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
        counter = {a for n in contract.notes for m in re.finditer(r"对方科目[：:]([^）)；;]+)", str(n.get("text") or ""))
                   for a in re.split(r"[、,，]", m.group(1)) if a.strip()}
        counter |= {str(f.get("from_account")) for f in contract.facts if f.get("from_account")}
        counter = sorted((a.strip() for a in counter if a.strip() and a.strip() != account), key=len, reverse=True)
        for finding in verify_directions(text, line_names, amount_truths, direction_shares, percent_kinds, note_signs, counter):
            {"failed": fail, "review": review}[finding[0]](*finding[1:])
    for d in cited_derived:
        if d.get("kind") == KIND_MOVEMENT and not d.get("valid", True) and not any(f.code == "MOVEMENT_NOT_PERMITTED" for f in findings):
            fail("MOVEMENT_NOT_PERMITTED", "binding", f"{d['ref']}: {d.get('invalid_reason')}{_movement_remedy(contract, d)}")
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

    # -- names ----------------------------------------------------------------------------------------
    for name in names_not_in_evidence(text, contract):
        review("NAME_NOT_IN_EVIDENCE", "names", f"'{name}' is not a name the evidence carries: write the name as the tab or a note "
               "gives it, or leave it out")

    # -- subject -------------------------------------------------------------------------------------
    # "全部由甲公司持有，该公司注册资本为5000万美元": 该公司 read as the shareholder gave it the target's capital (every
    # trial on the client file, 2026-11-09/10). The writer names the subject instead.
    if re.search(r"该公司|该企业|该主体", text):
        review("AMBIGUOUS_SUBJECT", "names", "'该公司' can be read as the target or as the company named before it: write 目标公司 or "
               "that company's name")

    # -- terms a note states -------------------------------------------------------------------------
    # "二期是20BP" in the note became "LPR减20BP" in the sentence: the sign and the benchmark were the writer's (2026-11-09/10)
    note_pool = " ".join(str(n.get("text") or "") for n in source_notes) or " ".join(str(n.get("text") or "") for n in contract.notes)
    for m in re.finditer(r"(加|减|上浮|下浮|\+|-|＋|－)\s*(\d+(?:\.\d+)?)\s*(?:BP|bp|bps|个基点|基点)", text):
        sign = "+" if m.group(1) in ("加", "上浮", "+", "＋") else "-"
        words = ("加", "上浮", "\\+", "＋") if sign == "+" else ("减", "下浮", "-", "－")
        if not re.search(r"(?:" + "|".join(words) + r")\s*" + re.escape(m.group(2)) + r"\s*(?:BP|bp|bps|个基点|基点)", note_pool):
            fail("TERM_NOT_IN_NOTE", "citations", f"'{m.group(0)}': the notes give {m.group(2)}BP without "
                 f"{'加' if sign == '+' else '减'}; write the spread as the note does")

    # Five unmarked major errors in one frozen draft of the client file (Phase A, 2026-11-10), each confirmed against the team's
    # finished report. Each check is written for the kind of error, not the sentence; the draft the writer produced is the
    # test that it fires, and other books are the test that it does not fire where nothing is wrong.
    # (1) ageing and maturity: "账龄超1年的保证金" where the note says 到期时间为1年以上 (and 账龄小于1年 of the other part).
    every_note = " ".join(str(n.get("text") or "") for n in contract.notes)
    noted_terms = _term_spans(every_note + " " + " ".join(str(f.get("line") or "") for f in contract.facts))
    for term, sign, years in sorted(_term_spans(text) - noted_terms):
        if any(t != term and s_ == sign and y == years for t, s_, y in noted_terms):
            said, noted = ("账龄", "到期") if term == "age" else ("到期", "账龄")
            fail("TERM_NOT_IN_NOTE", "citations", f"'{said}{'超过' if sign == '>' else '小于'}{years}年': the notes put {years}年 "
                 f"on {noted}, not on {said}; write the term the note uses")

    # (2) cost lines as the make-up of a net balance: "其中，房屋-原值-一期账面余额为5.00亿元" after the net book value.
    marker = re.search(r"其中|构成|组成|consist|made up of|comprise|of which", text, re.IGNORECASE)
    gross = [f for f in cited_facts if f["role"] == ROLE_COMPONENT and re.search(r"原值|cost", str(f.get("line") or ""), re.IGNORECASE)]
    if marker and gross and not re.search(r"原值|cost", text[:marker.start()], re.IGNORECASE) and \
            not _has(text, ("累计折旧", "累计摊销", "减值准备", "accumulated depreciation", "accumulated amortisation")):
        fail("GROSS_AS_MAKE_UP", "binding",
             f"{', '.join(str(f['line']) for f in gross[:3])} {'is' if len(gross) == 1 else 'are'} at cost (原值); the balance is net "
             "of depreciation and amortisation, so cost lines are not its make-up: give the net amounts, or say these are cost "
             "and give the accumulated depreciation")

    # (3) no adjustment where the databook records one: "示意性调整后营业收入与账面一致" with an ADJ of +3.0万元 in the period.
    denied = _ADJUSTMENT_DENIED_RE.search(text)
    if denied:
        periods_said = claim_periods or set(contract.periods)
        recorded = [n for n in contract.notes if n.get("kind") == "adjustment" and (not n.get("period") or n["period"] in periods_said)
                    and any(abs(v) > 0 for v in (_plain_number(x) for x in re.findall(r"[+\-−－]?\s*\d[\d,]*(?:\.\d+)?(?=\s*(?:万元|亿元|元))",
                                                                                       str(n.get("text") or ""))) if v is not None)]
        moved = [p for p in sorted(periods_said) if contract.total(p, BASIS_RAW) is not None and contract.total(p, BASIS_ADJUSTED) is not None
                 and abs(contract.total(p, BASIS_RAW) - contract.total(p, BASIS_ADJUSTED)) > 1.0]
        if recorded or moved:
            where = f"{recorded[0]['ref']} records {str(recorded[0].get('text') or '')[:60]}" if recorded else \
                f"the tab's total and the adjusted total differ at {moved[0]}"
            fail("ADJUSTMENT_DENIED", "binding", f"'{denied.group(0)}', but {where}: state the adjustment, or leave the comparison out")

    # (4) an item put in one period range when the account's notes list it in another: "为应收2025年10-12月的水电费" while a
    # note says 应收2024年3月的水电费.
    for m in _ITEM_PERIOD_RE.finditer(text):
        year, first, last, item = int(m.group(1)), int(m.group(2)), int(m.group(3) or m.group(2)), m.group(4)
        for n in contract.notes:
            other = next(((int(y), int(mo)) for y, mo in re.findall(r"(\d{4})年\s*(\d{1,2})\s*月(?:份)?\s*的?\s*" + re.escape(item[:2]),
                                                                  str(n.get("text") or ""))
                          if not (int(y) == year and first <= int(mo) <= last)), None)
            if other is not None:
                review("NOTE_SCOPE", "binding", f"'{m.group(0).lstrip('为系均全部属于 ')}': {n['ref']} lists {item[:2]} for {other[0]}年{other[1]}月; "
                       "say which part is which period, or do not put all of it in one")
                break

    # The cross-book trial (three portfolio books, 2026-11-10): kinds that recurred unmarked across books and both writers.
    # (5) the databook described instead of the company: "原表为隐藏工作表", "工作底稿未更新" (five sentences, both writers).
    meta = _WORKBOOK_META_RE.search(text)
    if meta:
        fail("WORKBOOK_META", "format", f"'{meta.group(0)}' describes the databook, not the company: say what the reader lacks "
             "(e.g. 明细未提供), not how the workbook is kept")

    # (6) a stated total that the parts listed before it do not make: "房屋建筑物为500.0万元，机械设备为100.0万元，合计700.0万元"
    total_word = re.search(r"合计|共计|总计|合共", text)
    rest_word = _REST_BEFORE_TOTAL_RE.search(text[:total_word.start()]) if total_word else None
    if total_word and rest_word:
        # "…，其余8项小额项目合计14.6万元" gives the total of the rest, not of the lines before it (a trial on four books refused three
        # right sentences of that kind as SUM_STATED, 2026-11-17, and a wrong one, whose rest was a line the total already includes, was
        # only refused for that reason): it is checked as the remainder of the period's total
        kids = _prefix_children([src for _s, src, _h in matched_sources])
        nested_starts = {span["start"] for span, src, _h in matched_sources if src.get("nested_in") or (src.get("id") or src.get("ref")) in kids}
        problem = _stated_rest_problem(text, amounts, total_word, rest_word, contract, claim_periods, nested_starts)
        if problem:
            fail("REST_NOT_REMAINDER", "arithmetic", problem)
    elif total_word and not _has(text, ("等", "主要")):
        spans = sorted(amounts, key=lambda a: a["start"])
        after = [a for a in spans if a["start"] >= total_word.end()]
        parts = [a for a in spans if a["end"] <= total_word.start()]
        if after and len(parts) >= 2:
            signed = [-a["value"] if text[max(0, a["start"] - 1):a["start"]] in ("-", "−", "－") or text[max(0, a["start"] - 1):a["start"]] == "负"
                      else a["value"] for a in parts]
            stated, slack = after[0]["value"], sum(a["tolerance"] for a in parts) + after[0]["tolerance"]
            # "累计折旧合计6991.5万元（房屋建筑物5786.0万元、机械设备587.9万元、室外工程617.6万元）": the total of what the bracket after it lists
            shut = next((b for b in _bracket_spans(text) if 0 <= b[0] - after[0]["end"] <= 2), None)
            inside = [a for a in spans if shut and shut[0] <= a["start"] < shut[1]]
            if len(inside) >= 2 and abs(sum(a["value"] for a in inside) - stated) <= sum(a["tolerance"] for a in inside) + after[0]["tolerance"]:
                pass
            elif min(abs(abs(sum(signed)) - stated), abs(sum(abs(v) for v in signed) - stated)) > slack:
                fail("SUM_STATED", "arithmetic", f"'{after[0]['written']}' is given as the total of "
                     f"{', '.join(a['written'] for a in parts)}, which add up to {abs(sum(signed)) / 1e4:,.1f}万元: name every part, "
                     "or do not call it their total")

    # (7) a period's change put down to an adjustment larger than the change: "较上年末增加100.0万元，主要系管理层调整调增
    # 600.0万元所致". The adjustment is a period-end level; what it adds to the change is the difference from the one before.
    for m in _ADJUSTMENT_CAUSE_RE.finditer(text):
        adj = next((a for a in sorted(amounts, key=lambda a: a["start"]) if m.end() <= a["start"] <= m.end() + 3), None)
        if adj is None:
            continue
        moved = [a for a in amounts if a["end"] <= m.start() and _MOVE_WORD_RE.search(text[max(0, a["start"] - 8):a["start"]])]
        if not moved and m.group("why") == "系":
            continue                        # "其中80.0万元系管理层调整调增" names a part; with no change in the sentence it is no reason
        change = moved[-1]["value"] if moved else None
        if change is None:
            ordered = sorted(contract.periods)
            last = max(claim_periods) if claim_periods else (ordered[-1] if ordered else None)
            before = [p for p in ordered if last and p < last]
            if last and before and contract.total(last) is not None and contract.total(before[-1]) is not None:
                change = abs(contract.total(last) - contract.total(before[-1]))
        if change is not None and adj["value"] > change + adj["tolerance"] + max(1.0, 0.02 * change):
            review("ADJUSTMENT_AS_MOVEMENT", "causal", f"'{adj['written']}' is the adjustment at the period end, larger than the "
                   f"change it is given as the reason for ({change / 1e4:,.1f}万元); what it adds to the change is the difference "
                   "from the adjustment the period before: state that, or do not give it as the reason")
            break

    # -- cause --------------------------------------------------------------------------------------
    if claim["cause_cites"] or _CAUSE_RE.search(text) or _FORWARD_RE.search(text) or _MOVEMENT_CAUSE_RE.search(text):
        checked.add("causal")
        notes = []
        for ref in claim["cause_cites"]:
            note = contract.note(ref)
            if note is None:
                fail("UNKNOWN_REF", "citations", f"cited note '{ref}' does not exist")
            else:
                notes.append(note)
        if not claim["cause_cites"] and not _CAUSE_RE.search(text) and not _FORWARD_RE.search(text):
            # only the bare 系 made it a cause: "审计调整减少50.0万元，系应收借款利息重分类" gives the cited note's own
            # account of the adjustment, and the notes it cites are what it rests on
            notes = list(cited_notes)
        related = [n for n in notes if len(_bigrams(text) & _bigrams(n["text"])) >= 2]
        conflicts = [(n, c) for n in related for c in cause_conflicts(text, n["text"], contract)]
        if len(related) > 1:
            # a sentence listing several lines cites a note per line; each note naming its own line is no contradiction
            conflicts = [(n, c) for n, c in conflicts if not c.startswith("a different line")]
        for n, why in conflicts:
            fail("CAUSE_CONTRADICTS_NOTE", "causal", f"{n['ref']} says something the sentence reverses or changes: {why}")
        if related and not conflicts:
            # Sharing words with a note shows the note is ABOUT this. It does not show the note says this. Entailment cannot be
            # read here, so the cause stays marked (orange) unless the sentence is mostly the note's own words: at least
            # NOTE_FAITHFUL of its content bigrams are in the notes it cites. Marking every note-based cause put 21 to 52
            # orange sentences in each arm's deck on the client file, and a sentence-level review found almost all of them right
            # (2026-11-09); of 90 such sentences 77 reached 60%, and the rest were advice and causes the writer added.
            words = _bigrams(text) - _bigrams(contract.account)
            pool = set().union(*(_bigrams(n["text"]) for n in related + cited_notes))
            cover = len(words & pool) / len(words) if words else 0.0
            if cover >= NOTE_FAITHFUL:
                info("CAUSE_FROM_NOTE", "causal", f"{cover:.0%} of the sentence's wording is in "
                     + ", ".join(n["ref"] for n in related) + "; no contradiction was found")
            else:
                info("CAUSE_NOTE_RELATED_UNCHECKED", "causal",
                     "cites " + ", ".join(n["ref"] for n in related) + f", which is about the same thing, but only {cover:.0%} of its "
                     "wording is there and no contradiction was found; a person reads the note against it")
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


_SOFT_NEGATION = ("无", "不存在", "不")
_NOT_NEGATION = ("无形", "无锡", "无论", "无息")


def _soft_negation(text: str) -> bool:
    work = text
    for word in _NOT_NEGATION:
        work = work.replace(word, "")
    return any(w in work for w in _SOFT_NEGATION)


def cause_conflicts(claim_text: str, note_text: str, contract: Optional[EvidenceContract] = None) -> List[str]:
    """Ways a sentence reverses or changes what a related note says, found WITHIN the stretches around
    the words they share (so a negation elsewhere in a long note is not charged to the claim):
    polarity, done against expected, all against part, an opposite direction, a different named line.
    Screening for contradiction only; a sentence that passes is related to the note, not entailed by it."""
    shared = _bigrams(claim_text) & _bigrams(note_text)
    c, n = _cue_classes(_window_text(claim_text, shared)), _cue_classes(_window_text(note_text, shared))
    found: List[str] = []
    # "经检查入账凭证无异常" against "未发现异常": both negate. 无/不存在 do not start a conflict on their own (a note's "无重大异常"
    # just past a shared phrase is no contradiction of a sentence that leaves it out) but answer one (client file, 2026-11-09)
    c_soft, n_soft = _soft_negation(_window_text(claim_text, shared)), _soft_negation(_window_text(note_text, shared))
    if ("neg" in c and "neg" not in n and not n_soft) or ("neg" in n and "neg" not in c and not c_soft):
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
        # bare labels: a counterparty row is starred in the tab ("*甲公司") and never in prose, so with the star kept a sentence
        # naming the note's own counterparty was charged with "a different line" (2026-11-09 review)
        labels = {_bare(f["line"]) for f in contract.facts if f["role"] == ROLE_COMPONENT and f["line"] and f["line"] != contract.account}
        in_claim, in_note = {l for l in labels if l in claim_text}, {l for l in labels if l in note_text}
        if in_claim and in_note and not (in_claim & in_note):
            found.append(f"a different line ({', '.join(sorted(in_claim))} against {', '.join(sorted(in_note))})")
    return found


def _verdict(claim: Dict[str, Any], findings: List[ClaimFinding], checked: set) -> Dict[str, Any]:
    dimensions = {}
    for dim in ("citations", "amounts", "binding", "arithmetic", "causal", "coverage", "names"):
        mine = [f for f in findings if f.dimension == dim]
        if any(f.severity == "failed" for f in mine):
            dimensions[dim] = "failed"
        elif any(f.severity == "review" for f in mine):
            dimensions[dim] = "unverified"
        elif any(f.code == "CAUSE_NOTE_RELATED_UNCHECKED" for f in mine):
            dimensions[dim] = "related_unchecked"
        elif any(f.code == "CAUSE_FROM_NOTE" for f in mine):
            dimensions[dim] = "from_note"
        elif any(f.code == "CAUSE_INFERENCE" for f in mine):
            dimensions[dim] = "inference"
        else:
            dimensions[dim] = "ok" if dim in checked else "n/a"
    failed = any(f.severity == "failed" for f in findings)
    unverified = any(f.severity == "review" for f in findings)
    # A cause that cites a note on the same subject, with no contradiction found, is not marked whatever its wording overlap:
    # a sentence-level review of four arms (2026-11-09) found 36 of 60 needless marks were these, at overlaps from 21% to 59%,
    # and the one wrong sentence among them sat at 47%, so no threshold separates them. It stays on the record
    # (note_wording_unchecked). A cause with no note at all is still marked.
    marked = any(f.code == "CAUSE_INFERENCE" for f in findings)
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
#: refusals that rest on how the verifier READ the prose (which period, which line, which part of which): a rule can misread a list, a
#: bracket, a range, a make-up. They go to the agent, which reads the sentence figure by figure, and Python checks that reading against
#: the evidence (check_reading). A sentence refused for these alone is orange until the agent has confirmed it is wrong (red) or
#: read it right (cleared). Refusals about the figures themselves (a figure no cited fact has, a unit, a sign, a percentage, the
#: deck's basis) are never the agent's to clear.
JUDGEABLE = frozenset({"WRONG_PERIOD", "PERIOD_UNBOUND", "MOVEMENT_END_UNSTATED", "WRONG_LINE", "LINE_NOT_CITED", "COMPOSITION_BASIS",
                       "COMPOSITION_EXCEEDS_BALANCE", "SUM_STATED", "REST_NOT_REMAINDER", "REST_INCOMPLETE"})
#: flags on a verified sentence that a person should read: orange in the deck. note_amount, note_wording_unchecked and
#: detail_caveat were here until 2026-11-09: a sentence-level review of the client file found them on 33 to 55 right
#: sentences per arm against 9 to 22 real problems, so the orange told a person nothing. They stay on the candidate and in
#: the review file; a note-based cause the verifier cannot read as the note's own words is still orange through its verdict.
HIGHLIGHT_FLAGS = ("cause_not_proven", "lead_left_to_person", "review_objection")


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
            confirmed = (claim.get("judged") or {}).get("verdict") == "wrong"
            red = next((f for f in failed if f.get("dimension") in _RED_DIMENSIONS and (f.get("code") not in JUDGEABLE or confirmed)), None)
            first = red or next((f for f in verdict["findings"] if f["severity"] in ("failed", "review")), {})
            category = "hallucination" if red else "reasoning"
            supported, reason, code = False, first.get("message", ""), first.get("code")
        elif verdict["verdict"] == "verified_inference":
            marker = next((f for f in verdict["findings"] if f["code"] == "CAUSE_INFERENCE"), {})
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


def _wan(value: float) -> float:
    """A yuan figure as the writer is shown it, in 万元: one decimal, rounded half-up from the exact value, as the report writes
    it; under 1万元 four decimals, so the exact yuan can be written. Shown with two decimals, the writer copied 12.35 and the
    deck rounded that again to 12.4 for 12.348675 (client file, 2026-11-09: twelve sentences wrong only in the deck)."""
    from decimal import ROUND_HALF_UP, Decimal
    wan = Decimal(str(float(value))) / Decimal(10000)
    places = Decimal("0.1") if abs(wan) >= 1 else Decimal("0.0001")
    return float(wan.quantize(places, rounding=ROUND_HALF_UP))


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
                         "wan": [_wan(cells[c]["value"]) if c in cells else None for c in cols],
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
               "yuan": round(f["value"], 2), "wan": _wan(f["value"])}
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
               "wan": [_wan(e["facts"][p]["value"]) if p in e["facts"] else None for p in shown_periods]}
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
                rest["wan"].append(_wan(d["value"]))
            else:
                rest["refs"].append(refs[0])
                rest["wan"].append(_wan(by_ref[refs[0]]["value"]))
    rest_refs = set((rest or {}).get("refs") or [])
    omitted = len(left)
    tables = _supporting_view(contract, totals, shown_periods)
    related = [{"account": f.get("from_account") or f["line"], "ref": f["ref"], "period": f["period"], "wan": _wan(f["value"]),
                **({"flow": True} if str(f.get("statement") or "").upper().startswith("IS") else {})}
               for f in contract.facts if f["role"] == ROLE_RELATED]
    # this account with its related accounts of the same statement, one citable figure: the team writes the VAT on both asset
    # tabs, or a loan with its current portion, as one amount (the client file: 其他流动资产 with 其他非流动资产)
    group_totals = []
    own_statement = str(contract.statement_type or "").upper()[:2]
    for f in contract.facts:
        if f["role"] != ROLE_TOTAL or f["basis"] != BASIS_ADJUSTED:
            continue
        mates = [g for g in contract.facts if g["role"] == ROLE_RELATED and g["period"] == f["period"]
                 and str(g.get("statement") or "").upper()[:2] == own_statement]
        if mates:
            d = derive_sum(contract, [f["ref"]] + [g["ref"] for g in mates])
            if d.get("valid", True) and d.get("value") is not None:
                group_totals.append({"ref": d["ref"], "period": f["period"], "wan": _wan(d["value"]),
                                     "accounts": [contract.account] + [g.get("from_account") or g["line"] for g in mates]})
    rest_refs |= {g["ref"] for g in group_totals}
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
        **({"related_total": group_totals} if group_totals else {}),
        "components_not_shown": omitted,
        "material_not_shown": material_left,
    }

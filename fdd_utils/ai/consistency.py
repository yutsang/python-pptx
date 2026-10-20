"""Cross-account consistency, after every account is written (P4, report-only).

Every check the pipeline runs while writing looks at ONE account: its own
figures, its own clauses. Nothing puts two accounts' text side by side, and
nothing asks whether what the numbers say most loudly was said at all. This
module does both, deterministically, over the finished run:

* OBSERVATION_UNCOVERED -- an analyst observation (analysis.py: implied rate,
  留抵进项税, a negative component, one counterparty carrying an account, ...)
  that no commentary of the accounts it concerns mentions, by figure or by
  name. The arithmetic was done and the deck does not show it.
* CROSS_REF_AMOUNT -- account A names account B and puts an amount right after
  it, and that amount is neither one of B's figures nor one of A's own (its
  adjustment entries included). Grounding accepts any figure in the pool, so a
  real number attributed to the wrong account passes it.

Report-only on purpose: nothing here retries, patches or gates. A finding is a
place for the reader to look, and how often these fire on real runs decides
whether any of it should ever steer the pipeline.
"""
from __future__ import annotations

import re
from typing import Any, Dict, List, Optional

#: What a sentence covering each observation would say, when it does not quote
#: the figure itself.
_COVER_WORDS = {
    "IMPLIED_RATE": r"利率|融资成本|资金成本|利息率|interest rate|cost of (?:debt|borrowing)",
    "RECEIVABLE_DAYS": r"周转天数|周转期|回款周期|账期|DSO|days",
    "ADVANCE_COVER": r"预收[^。；;]{0,20}(?:收入|月)|相当于",
    "CIP_TRANSFER": r"转固|转入固定资产|结转(?:至|为)?固定资产|transferred to fixed",
    "DEPRECIATION_RATE": r"已提比例|折旧比例|成新率|累计折旧[^。；;]{0,12}(?:占|比例)",
    "CONCENTRATION": r"集中|占比|占本科目|占[^。；;]{0,6}\d+(?:\.\d+)?%|concentrat",
    "GOVERNMENT_COUNTERPARTY": r"税务局|政府|退还|可退|refund",
    "RELATED_PARTY": r"关联|related part|intra-?group",
    "INPUT_VAT_CREDIT": r"留抵|待抵扣|进项税|input VAT",
    "NEGATIVE_BALANCE": r"负数|负值|负余额|贷方余额|借方余额|反向|红字|negative|credit balance|debit balance",
    "STATIC_BALANCE": r"未发生变动|保持不变|无变动|各期均为|一直为|unchanged",
    "NO_DETAIL": r"未列明细|无明细|仅有合计|未提供明细|no breakdown",
    "UNTIED_EQUITY": r"实收资本|注册资本|未分配利润|paid-in|share capital",
}

#: An amount this close after another account's name is being attributed to it.
_CROSS_REF_WINDOW = 16


def _texts(results: Dict[str, Any]) -> Dict[str, str]:
    from ..financial_common import get_pipeline_result_text

    out: Dict[str, str] = {}
    for key, content in (results or {}).items():
        if str(key).startswith("__"):
            continue
        try:
            text = get_pipeline_result_text(content)
        except Exception:
            text = ""
        if text:
            out[str(key)] = str(text)
    return out


def _amount_matches(value: float, amount_text: str, target: float) -> bool:
    from .validator import display_half_ulp

    try:
        half = display_half_ulp(amount_text)
    except Exception:
        half = 0.0
    return abs(abs(value) - abs(target)) <= max(half * 1.01, 0.005 * abs(target), 1.0)


def observation_coverage(observations: List[Dict[str, Any]], texts: Dict[str, str]) -> List[Dict[str, Any]]:
    """Observations no commentary of their accounts mentions, by figure or by name."""
    from .validator import extract_amount_spans

    findings: List[Dict[str, Any]] = []
    for obs in observations:
        accounts = [a for a in (obs.get("accounts") or []) if a in texts]
        if not accounts:
            continue   # nothing was written for them at all; another check's business
        body = "\n".join(texts[a] for a in accounts)
        covered = False
        numbers = [float(n) for n in (obs.get("numbers") or []) if abs(float(n)) >= 1000]
        if numbers:
            for value, start, end in extract_amount_spans(body):
                if any(_amount_matches(value, body[start:end], n) for n in numbers):
                    covered = True
                    break
        words = _COVER_WORDS.get(str(obs.get("code") or ""))
        if not covered and words and re.search(words, body, re.IGNORECASE):
            covered = True
        if not covered:
            findings.append({
                "code": "OBSERVATION_UNCOVERED",
                "severity": obs.get("severity") or "medium",
                "accounts": accounts,
                "detail": f"{obs.get('code')}: {obs.get('finding')}",
            })
    return findings


def _name_pattern(names: List[str]) -> Optional["re.Pattern[str]"]:
    # Longest first, so 其他应付款 is taken whole and never read as 应付款.
    usable = sorted({n for n in names if len(n) >= 3 and re.search(r"[一-鿿]", n)},
                    key=len, reverse=True)
    if not usable:
        return None
    return re.compile("|".join(re.escape(n) for n in usable))


def cross_reference_amounts(texts: Dict[str, str], dfs: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Amounts written against another account's name that are not that account's."""
    from .validator import SourceIndex, extract_amount_spans, ground_amounts

    pattern = _name_pattern(list(dfs or {}))
    if pattern is None:
        return []
    indexes: Dict[str, Any] = {}

    def index_of(key: str):
        if key not in indexes:
            try:
                indexes[key] = SourceIndex.from_df(dfs[key])
            except Exception:
                indexes[key] = None
        return indexes[key]

    def grounded(amount_text: str, key: str) -> bool:
        source = index_of(key)
        if source is None:
            return True   # cannot judge -> say nothing
        try:
            verdict = ground_amounts(amount_text, source)
        except Exception:
            return True
        return bool(verdict) and verdict.get("category") != "hallucination"

    findings: List[Dict[str, Any]] = []
    for account, text in texts.items():
        if account not in dfs:
            continue
        for sentence in re.split(r"[。；;\n]", text):
            for mention in pattern.finditer(sentence):
                other = mention.group(0)
                if other == account or other not in dfs:
                    continue
                window = sentence[mention.end():mention.end() + _CROSS_REF_WINDOW]
                window = re.split(r"[，,、]", window, maxsplit=1)[0]
                spans = extract_amount_spans(window)
                if not spans:
                    continue
                _value, start, end = spans[0]
                amount_text = window[start:end]
                if grounded(amount_text, other) or grounded(amount_text, account):
                    continue
                findings.append({
                    "code": "CROSS_REF_AMOUNT",
                    "severity": "medium",
                    "accounts": [account, other],
                    "detail": (f"{account} writes {other} next to {amount_text}, which is neither one of "
                               f"{other}'s figures nor one of {account}'s own: 「{sentence.strip()[:120]}」"),
                })
    return findings


_UP = re.compile(r"增加|增长|上升|上涨|回升|扩大|新增|increase|rose|grew|up\b", re.IGNORECASE)
_DOWN = re.compile(r"减少|下降|下滑|回落|降至|收窄|缩减|decrease|fell|declin|down\b", re.IGNORECASE)


def _balance_moves(df: Any) -> List[float]:
    """Latest-vs-prior and first-to-last change of a balance-sheet total."""
    attrs = getattr(df, "attrs", None) or {}
    stype = str(((attrs.get("integrity") or {}).get("statement_type")) or "").upper()
    if stype and stype != "BS":
        return []
    totals = attrs.get("projection_totals_by_date") or {}
    series = [(str(p), float(v)) for p, v in totals.items()
              if re.fullmatch(r"\d{4}-\d{2}-\d{2}", str(p)) and isinstance(v, (int, float)) and v == v]
    series.sort()
    if len(series) < 2:
        return []
    return [series[-1][1] - series[-2][1], series[-1][1] - series[0][1]]


def summary_direction_findings(summaries: Dict[str, str], dfs: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Executive-summary clauses whose direction contradicts the account they name.

    Judged only where a clause names exactly ONE balance-sheet account and uses
    exactly one direction, and flagged only when BOTH the latest movement and
    the whole-period one go the other way -- a summary may legitimately
    describe either. The summary is the most-read text in the deck, and its
    grounding checks amounts, never direction.
    """
    pattern = _name_pattern(list(dfs or {}))
    if pattern is None:
        return []
    findings: List[Dict[str, Any]] = []
    for statement, text in (summaries or {}).items():
        for clause in re.split(r"[。；;，,\n]", str(text or "")):
            names = {m.group(0) for m in pattern.finditer(clause)}
            if len(names) != 1:
                continue
            up, down = bool(_UP.search(clause)), bool(_DOWN.search(clause))
            if up == down:
                continue
            account = names.pop()
            moves = _balance_moves(dfs.get(account))
            if not moves or any(abs(m) < 1e-6 for m in moves):
                continue
            if up and all(m < 0 for m in moves) or down and all(m > 0 for m in moves):
                findings.append({
                    "code": "SUMMARY_DIRECTION",
                    "severity": "high",
                    "accounts": [account],
                    "detail": (f"{statement} summary says {account} went {'up' if up else 'down'}, but it "
                               f"moved {moves[0]:+,.0f} in the latest period and {moves[1]:+,.0f} over "
                               f"the whole series: 「{clause.strip()[:120]}」"),
                })
    return findings


def cross_account_findings(results: Dict[str, Any], dfs: Dict[str, Any],
                           graph_facts: Optional[Dict[str, Any]] = None) -> List[Dict[str, Any]]:
    """Both checks over a finished run. Never raises; an empty list is an answer."""
    findings: List[Dict[str, Any]] = []
    try:
        texts = _texts(results)
    except Exception:
        return findings
    try:
        from .analysis import analyse, views_from_frames
        observations = [o.as_dict() for o in analyse(views_from_frames(dfs or {}, graph_facts or {}))
                        if o.severity in ("high", "medium")]
        findings.extend(observation_coverage(observations, texts))
    except Exception:
        pass
    try:
        findings.extend(cross_reference_amounts(texts, dfs or {}))
    except Exception:
        pass
    order = {"high": 0, "medium": 1, "low": 2}
    findings.sort(key=lambda f: (order.get(f["severity"], 9), f["code"]))
    return findings

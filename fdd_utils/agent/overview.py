"""The executive summary of each statement, written from the finished paragraphs and checked like them.

The deck's summary band held, on the agent path, a splice of each account's first sentence. This writes it instead: one call
per statement, given every finished account's latest balance and movement (with their refs) and its verified opening sentence.
Each sentence names its account and cites that account's refs, so the verifier checks it against that account's evidence,
exactly as a paragraph sentence is checked. A sentence that fails gets one repair; what still fails is left out, and the
result says how many were. About three thousand tokens in and a few hundred out per statement, so the stronger model is the
default and the cheaper one can be compared at almost no cost.
"""
from __future__ import annotations

import json
import re
from typing import Any, Dict, List, Optional

from ..ai import claims as cl
from ..ai import evidence_contract as ec

SECTIONS = {"BS": "资产负债表", "IS": "利润表"}


def overview_prompt(section: str, language: str, chars: int) -> str:
    chinese = language in ("Chi", "Chn")
    name = SECTIONS[section] if chinese else {"BS": "balance sheet", "IS": "income statement"}[section]
    size = f"at most {chars} Chinese characters" if chinese else f"at most {max(40, chars * 3 // 4)} words"
    return (
        f"You write the executive summary of the {name} section of a financial due diligence report, {size}, in "
        f"{'简体中文' if chinese else 'English'}.\n"
        "ACCOUNTS lists each finished account: its latest balance and movement (with refs) and the opening of its verified\n"
        "paragraph. Lead with what matters most to a buyer: the largest balances, the largest movements, and items the\n"
        "paragraphs flag. Amounts as 万元 or 亿元 with one decimal; periods as the account's period label gives them.\n"
        "Every sentence is one claim about ONE account and cites only that account's refs. No refs or internal names in the text.\n"
        'Reply with exactly one JSON object: {"claims": [{"account": "<account>", "kind": "amount|movement|descriptive", '
        '"text": "...", "periods": [ISO dates], "cites": ["F/D refs of that account"]}]}'
    )


def _account_line(session: Any, key: str, labels: Dict[str, str]) -> Optional[Dict[str, Any]]:
    c = session.contracts.get(key)
    cand = session.candidates.get(key)
    if c is None or cand is None or not cand.validated or not c.periods:
        return None
    cl.standard_derived(c)
    latest = c.periods[-1]
    total = next((f for f in c.facts if f["period"] == latest and f["role"] == ec.ROLE_TOTAL and f["basis"] == ec.BASIS_ADJUSTED), None)
    if total is None or abs(float(total["value"])) < 0.5:
        return None                     # a nil balance is not a summary point ("营业外支出为0.0万元" on the client file)
    move = next((d for d in c.derived if d.get("kind") == cl.KIND_MOVEMENT and (d.get("periods") or [])[-1:] == [latest]
                 and d.get("valid", True)), None)
    verified = [cl_text for cl_text, v in zip((x.get("text") for x in cand.claims), cand.outcome.get("verdicts") or [])
                if v.get("verdict") in ("verified", "verified_inference")]
    line = {"account": key, "period_label": labels.get(latest, latest),
            "balance": {"ref": total["ref"], "period": latest, "wan": cl._wan(total["value"])},
            "paragraph": "".join(str(t) for t in verified[:2])[:180],
            "flags": len(cand.flags or [])}
    if move is not None:
        line["movement"] = {"ref": move["ref"], "periods": move.get("periods"), "wan": cl._wan(float(move.get("value") or 0)),
                            "pct": move.get("pct")}
    return line


def _check(session: Any, claims: List[Dict[str, Any]], allowed: Dict[str, set]) -> List[Dict[str, Any]]:
    """Each claim against its own account's evidence; refs outside what that account was shown fail."""
    out = []
    for claim in claims:
        key = session.resolve_account(str(claim.get("account") or "")) or ""
        cites = [str(r) for r in claim.get("cites") or []]
        if key not in allowed or not cites or any(r not in allowed[key] for r in cites):
            out.append({"verdict": "failed", "findings": [{"code": "UNKNOWN_REF", "severity": "failed", "dimension": "citations",
                                                           "message": "cite only the refs listed for that account"}]})
            continue
        c = session.contracts[key]
        periods = sorted({p for r in cites for p in ((c.fact(r) or {}).get("periods") or [(c.fact(r) or {}).get("period")]) if p})
        out.append(cl.verify_claim(dict(claim, account=key, cites=cites, periods=periods or list(claim.get("periods") or [])),
                                   session.contracts))
    return out


def write_overview(session: Any, router: Any, *, role: str, language: str, accounts: List[str], period_labels: Dict[str, Dict[str, str]],
                   chars: int = 140) -> Dict[str, Any]:
    """{section: {"text", "claims", "dropped", "calls", "error"}} for BS and IS, written by `role` (primary / escalation)."""
    result: Dict[str, Any] = {}
    for section in SECTIONS:
        lines = [x for x in (_account_line(session, k, period_labels.get(k) or {}) for k in accounts
                             if str(session.contracts[k].statement_type or "").upper().startswith(section)) if x]
        entry: Dict[str, Any] = {"text": "", "claims": [], "dropped": 0, "calls": 0, "error": "", "accounts": len(lines)}
        result[section] = entry
        if not lines:
            continue
        allowed = {x["account"]: {x["balance"]["ref"]} | ({x["movement"]["ref"]} if "movement" in x else set()) for x in lines}
        system = overview_prompt(section, language, chars)
        user = "ACCOUNTS:\n" + "\n".join(json.dumps(x, ensure_ascii=False, separators=(",", ":")) for x in lines)
        claims: List[Dict[str, Any]] = []
        verdicts: List[Dict[str, Any]] = []
        for attempt in range(2):
            try:
                reply = router.call_as(role, system, user, 900, "overview")
            except Exception as exc:                                    # noqa: BLE001
                entry["error"] = f"{type(exc).__name__}: {str(exc)[:200]}"
                break
            entry["calls"] += 1
            text = str((reply or {}).get("content") or "")
            m = re.search(r"\{.*\}", text, re.S)
            try:
                data = json.loads(m.group(0)) if m else {}
            except ValueError:
                data = {}
            new = [c for c in (data.get("claims") or []) if isinstance(c, dict) and c.get("text")]
            if attempt == 0:
                claims = new
            else:                                   # a repair replaces only the claims that failed
                bad = [n for n, v in enumerate(verdicts) if v.get("verdict") not in ("verified", "verified_inference")]
                for n, c in zip(bad, new):
                    claims[n] = c
            verdicts = _check(session, claims, allowed)
            failed = [(n, v) for n, v in enumerate(verdicts) if v.get("verdict") not in ("verified", "verified_inference")]
            if not failed or not claims:
                break
            user = (user + "\n\nREJECTED (rewrite each, same order, as {\"claims\": [...]}):\n"
                    + "\n".join(f"  [{n}] {claims[n].get('text')} -- " + "; ".join(f.get("message", "")[:160] for f in v.get("findings") or []
                                                                                 if f.get("severity") in ("failed", "review"))
                                for n, v in failed))
        kept, seen, used = [], set(), 0
        budget = chars if language in ("Chi", "Chn") else chars * 4      # the band holds about this much (English: ~4 chars a word)
        for c, v in zip(claims, verdicts):
            text_ = str(c.get("text") or "").strip()
            if v.get("verdict") not in ("verified", "verified_inference") or text_ in seen:
                continue                # a sentence written twice was kept twice on the client file
            if kept and used + len(text_) > budget * 1.15:
                break                   # the band's size, not the prompt's request: one model wrote 271 characters for 140
            kept.append((c, v))
            seen.add(text_)
            used += len(text_)
        joiner = "" if language in ("Chi", "Chn") else " "
        entry.update(text=joiner.join(str(c["text"]).strip() for c, _v in kept), claims=[c for c, _v in kept],
                     dropped=len(claims) - len(kept))
    return result

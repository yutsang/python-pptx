"""A place for the sentences the verifier refused, kept so they can be looked at later, with no workbook and no model call.

The verifier reads two things: the account's evidence contract and the sentences. Both are plain data. One file per account that ended
a run with a refused sentence is therefore the whole case: the contract (every figure, its period, line and basis, every note), the
draft, and what was refused with the verifier's message. Replayed with today's code it says whether the refusal is still there, so a
fix is checked on the cases the paid runs produced instead of on the next paid run.

    capture(source, folder, entity=..., code=...)   after a run: writes <folder>/<entity>/<account>.json, replacing that account's file
    replay(path)                                    free: the refused sentences read again by the verifier as it is now

A case holds client figures and sentences. The folder is under fdd_utils/output (git-ignored); it stays on the machine that ran.
What the agent's own reading cleared (judge_binding) is not in a case: only sentences still refused at the end of the run are kept,
and only those are replayed, so the replay is the verifier alone.

Two kinds of case end up here. A right sentence the verifier refuses is a rule to fix (replay shows it open until it is fixed, then
fixed, and the file stays as a test of the rule). A wrong sentence, or one the databook cannot support, stays open for good: that is a
person's decision, and the case is the record of it.
"""
from __future__ import annotations

import datetime
import json
import re
from pathlib import Path
from typing import Any, Dict, List, Optional

from ..ai import claims as cl
from ..ai import evidence_contract as ec

SCHEMA = 1
_PASSED = ("verified", "verified_inference")


def _safe(name: Any) -> str:
    return re.sub(r'[\\/:*?"<>|\s]+', "_", str(name)).strip("_")[:80] or "account"


def _bad(verdict: Dict[str, Any]) -> List[Dict[str, Any]]:
    return [f for f in verdict.get("findings") or [] if f.get("severity") in ("failed", "review")]


def refused(outcome: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Every sentence of a verified paragraph that did not pass: its index, text, colour in the deck, codes and messages."""
    out = []
    for n, v in enumerate((outcome or {}).get("verdicts") or []):
        if v.get("verdict") in _PASSED:
            continue
        bad = _bad(v)
        red = any(f.get("severity") == "failed" and f.get("dimension") in cl._RED_DIMENSIONS for f in bad)
        out.append({"claim": n, "colour": "red" if red else "orange", "text": str(v.get("text") or ""),
                    "codes": sorted({str(f.get("code")) for f in bad}),
                    "messages": [f"{f.get('code')}: {str(f.get('message') or '')[:300]}" for f in bad]})
    return out


def capture(source: Any, folder: Path, *, entity: str, code: str = "") -> List[Path]:
    """Write one case for each account of `source` (an agent SourceSession) whose draft still has a refused sentence."""
    written: List[Path] = []
    target = Path(folder) / _safe(entity)
    for account, candidate in source.candidates.items():
        bad = refused(candidate.outcome)
        contract = source.contracts.get(account)
        if not bad or contract is None or not candidate.claims:
            continue
        target.mkdir(parents=True, exist_ok=True)
        path = target / f"{_safe(account)}.json"
        case = {"schema": SCHEMA, "entity": entity, "account": account, "code": code,
                "captured": datetime.datetime.now().isoformat(timespec="seconds"),
                "accepted": bool(candidate.accepted and candidate.validated),
                "contract": contract.to_dict(), "claims": list(candidate.claims), "refused": bad}
        path.write_text(json.dumps(case, ensure_ascii=False, indent=1), encoding="utf-8")
        written.append(path)
    return written


def replay(path: Path) -> Dict[str, Any]:
    """The refused sentences of one case read again by the verifier as it is now: status fixed (none refused any more), partly
    (fewer than before) or open, and for each still refused its codes and messages."""
    case = json.loads(Path(path).read_text(encoding="utf-8"))
    contract = ec.EvidenceContract.from_dict(case.get("contract"))
    account = str(case.get("account") or "")
    if contract is None:
        return {"path": str(path), "entity": case.get("entity"), "account": account, "status": "unreadable", "was": [], "now": []}
    claims = list(case.get("claims") or [])
    was, now = list(case.get("refused") or []), []
    for item in was:
        n = int(item["claim"])
        if n >= len(claims):
            continue
        claim = dict(claims[n], account=account, source_version=contract.source_version)
        verdict = cl.verify_claim(claim, {account: contract}, default_account=account, index=n)
        if verdict.get("verdict") in _PASSED:
            continue
        bad = _bad(verdict)
        now.append({"claim": n, "text": item.get("text") or str(claims[n].get("text") or ""), "codes": sorted({str(f.get("code")) for f in bad}),
                    "messages": [f"{f.get('code')}: {str(f.get('message') or '')[:300]}" for f in bad]})
    status = "fixed" if not now else ("partly" if len(now) < len(was) else "open")
    return {"path": str(path), "entity": case.get("entity"), "account": account, "code": case.get("code"), "captured": case.get("captured"),
            "status": status, "was": was, "now": now}


def cases(folder: Path) -> List[Path]:
    return sorted(Path(folder).glob("*/*.json"))


def summary(results: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Counts for a list of replays: by status, and which codes the open ones are refused with (the queue to fix, largest first)."""
    by_status: Dict[str, int] = {}
    by_code: Dict[str, int] = {}
    for r in results:
        by_status[r["status"]] = by_status.get(r["status"], 0) + 1
        for item in r["now"]:
            for code in item["codes"]:
                by_code[code] = by_code.get(code, 0) + 1
    return {"cases": len(results), "by_status": by_status, "open_by_code": dict(sorted(by_code.items(), key=lambda kv: -kv[1]))}

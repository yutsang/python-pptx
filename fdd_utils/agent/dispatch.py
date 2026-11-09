"""Which model writes which account: Python describes each account, the stronger model assigns.

Python cannot tell a hard paragraph from an easy one by itself, and a rule written ahead of the data guesses. So the state
measures what makes a paragraph hard (a conflict with Financials, adjustments, material notes to state, detail that does not
tie or is nested, part-year figures, how much of the statement the account is) and one call to the stronger model reads those
profiles and names the accounts that need it. Everything it does not name goes to the cheaper model, which still hands a
paragraph it cannot repair to the stronger one (routing.Policy.escalate_after). One call per databook: about two thousand
tokens in, a few hundred out.

If the reply cannot be read, a rule decides, and the plan says so.
"""
from __future__ import annotations

import json
import re
from typing import Any, Dict, List, Optional

from ..ai import claims as cl
from ..ai import evidence_contract as ec

PRO, FLASH = "pro", "flash"

DISPATCH_PROMPT = (
    "You assign each account of a financial due diligence report to the model that will write its paragraph.\n"
    "flash is cheap and reliable on plain paragraphs: a balance, a movement, a few detail lines, a note or two.\n"
    "pro costs twelve times as much. Give it the paragraphs that must reconcile or judge: the tab and Financials in conflict,\n"
    "adjustments to explain, several material notes to state, detail that does not tie or is nested, part-year figures to\n"
    "annualise, or an account large enough that an error would matter to the reader.\n"
    "Give pro as few as the profiles allow. A flash paragraph that fails is repaired by pro anyway.\n"
    'Reply with exactly one JSON object: {"pro": ["<account>", ...], "why": {"<account>": "<a few words>"}}.\n'
    "Every account you do not list goes to flash."
)


def account_profiles(session: Any, accounts: List[str]) -> List[Dict[str, Any]]:
    """What makes each account's paragraph hard, measured from its evidence. No model call."""
    from .controller import AgentController
    bare = AgentController.__new__(AgentController)       # _must_address reads only class-level word lists
    latest_totals: Dict[str, float] = {}
    for key in accounts:
        c = session.contracts[key]
        if c.periods:
            latest_totals[key] = abs(c.total(c.periods[-1]) or 0.0)
    by_statement: Dict[str, float] = {}
    for key in accounts:
        st = str(session.contracts[key].statement_type or "").upper()[:2]
        by_statement[st] = by_statement.get(st, 0.0) + latest_totals.get(key, 0.0)
    out = []
    for key in accounts:
        c = session.contracts[key]
        view = cl.evidence_view(c)
        latest = c.periods[-1] if c.periods else ""
        perm = c.permissions.get(latest) or {}
        st = str(c.statement_type or "").upper()[:2]
        conflict = any(not p.get("describe_total", True) for p in c.permissions.values()) or any(
            ch["kind"] == ec.CHECK_TOTAL_VS_FINANCIALS and ch["status"] == ec.MISMATCH for ch in c.checks)
        out.append({
            "account": key, "statement": st or "?",
            "share_pct": round(100 * latest_totals.get(key, 0.0) / by_statement[st], 1) if by_statement.get(st) else None,
            "must_address": len(bare._must_address(view)),
            "notes": sum(1 for n in c.notes if n.get("kind") != "related_note"),
            "adjustments": sum(1 for n in c.notes if n.get("kind") == "adjustment"),
            "conflict_with_financials": conflict,
            "detail": perm.get("detail_basis", "none"),
            "nested": any(r.get("nested_in") for r in view["components"]["lines"]),
            "lines": len(view["components"]["lines"]) + view.get("components_not_shown", 0),
            "tables": len(view.get("supporting_tables") or []),
            "annualised": bool((c.annualization or {}).get("annualized")),
        })
    return out


def rule_plan(profiles: List[Dict[str, Any]]) -> Dict[str, str]:
    """The fallback when the model's reply cannot be read: the same signals, a fixed threshold."""
    plan = {}
    for p in profiles:
        hard = (p["conflict_with_financials"] or p["adjustments"] > 0 or p["must_address"] >= 3
                or (p["detail"] in ("ambiguous", "untied") and p["lines"] >= 5) or p["nested"]
                or (p["annualised"] and (p["share_pct"] or 0) >= 10))
        plan[p["account"]] = PRO if hard else FLASH
    return plan


def dispatch(session: Any, router: Any, accounts: List[str], *, role: str = "escalation") -> Dict[str, Any]:
    """One call: the stronger model reads the profiles and names the accounts it should write. Returns
    {"plan": {account: pro|flash}, "why": {...}, "profiles": [...], "by": "model"|"rule", "error": str}."""
    profiles = account_profiles(session, accounts)
    user = "ACCOUNTS (one profile each):\n" + "\n".join(json.dumps(p, ensure_ascii=False, separators=(",", ":")) for p in profiles)
    plan, why, by, error = {}, {}, "model", ""
    try:
        reply = router.call_as(role, DISPATCH_PROMPT, user, 600, "dispatch")
        text = str((reply or {}).get("content") or "")
        m = re.search(r"\{.*\}", text, re.S)
        data = json.loads(m.group(0)) if m else None
        if not isinstance(data, dict) or not isinstance(data.get("pro"), list):
            raise ValueError(f"not a dispatch reply: {text[:120]!r}")
        named = {session.resolve_account(str(a)) or str(a) for a in data["pro"]}
        plan = {k: (PRO if k in named else FLASH) for k in accounts}
        why = {str(k): str(v)[:120] for k, v in (data.get("why") or {}).items()} if isinstance(data.get("why"), dict) else {}
    except Exception as exc:                                            # noqa: BLE001 -- any unreadable reply falls back
        plan, by, error = rule_plan(profiles), "rule", f"{type(exc).__name__}: {str(exc)[:200]}"
    return {"plan": plan, "why": why, "profiles": profiles, "by": by, "error": error,
            "pro": sorted(k for k, v in plan.items() if v == PRO)}

"""The numerical summary: every figure in it traces to a fact, and says how far it was checked.

No model writes this. For each account in scope: the latest balance (the figure
the deck uses), the change from the previous period, and what is known about
those figures -- tied to Financials or not, replaced by Financials, detail that
does or does not add up. Every number printed carries the ref of the fact or
derived fact it comes from, and `traceability` counts the numbers printed against
the numbers that resolve, so a summary that prints a figure nothing backs says so.
"""
from __future__ import annotations

import re
from typing import Any, Dict, List, Optional

from ..ai import claims as cl
from ..ai import evidence_contract as ec
from ..ai.claims import text_amounts, text_percents


def _wan(value: float) -> str:
    return f"{abs(value) / 1e4:,.1f}万元"


def _status_flags(contract: ec.EvidenceContract, period: str) -> List[str]:
    flags: List[str] = []
    tf = contract.check(period, ec.CHECK_TOTAL_VS_FINANCIALS) or {}
    dt = contract.check(period, ec.CHECK_DETAIL_VS_TOTAL) or {}
    status = tf.get("status")
    if status == ec.PASS:
        flags.append("与报表一致")
    elif status == ec.MISMATCH and tf.get("overridden"):
        flags.append("已按报表金额列示（本表未更新）")
    elif status == ec.MISMATCH:
        flags.append("本表与报表不一致，未确定采用哪一方")
    elif status == ec.NOT_COMPARABLE:
        flags.append("无独立报表可核对")
    elif status == ec.MISSING:
        flags.append("报表无此期间金额")
    if dt.get("status") == ec.MISMATCH:
        flags.append("明细不等于合计")
    elif dt.get("status") == ec.AMBIGUOUS:
        flags.append("明细层级不明")
    return flags


def build_numerical_summary(session: Any, accounts: Optional[List[str]] = None) -> Dict[str, Any]:
    """Rows, rendered lines, and the count of printed numbers that resolve to a fact."""
    rows: List[Dict[str, Any]] = []
    keys = [session.resolve_account(a) or a for a in (accounts or list(session.contracts))]
    for key in keys:
        contract = session.contracts.get(key)
        if contract is None:
            rows.append({"account": key, "error": "no such account"})
            continue
        in_scope = session.in_scope(key)
        if contract.issues and any(i["severity"] == "blocked" for i in contract.issues):
            rows.append({"account": key, "in_scope": in_scope, "issues": contract.issues, "figures": []})
            continue
        if not contract.periods:
            rows.append({"account": key, "in_scope": in_scope, "issues": [{"code": "NO_PERIODS", "message": "no periods"}], "figures": []})
            continue
        latest = contract.periods[-1]
        total = next((f for f in contract.facts if f["period"] == latest and f["role"] == ec.ROLE_TOTAL
                      and f["basis"] == ec.BASIS_ADJUSTED), None)
        figures: List[Dict[str, Any]] = []
        if total is not None:
            figures.append({"label": "balance", "period": latest, "ref": total["ref"], "value": total["value"], "text": _wan(total["value"])})
        movement = None
        if len(contract.periods) >= 2:
            movement = cl.derive_movement(contract, contract.periods[-2], latest)
        if movement is not None and movement["valid"] and movement.get("value") is not None:
            figures.append({"label": "change", "period": latest, "from": contract.periods[-2], "ref": movement["ref"],
                            "value": movement["value"], "text": ("+" if movement["value"] >= 0 else "-") + _wan(movement["value"])})
            if movement.get("pct") is not None:
                figures.append({"label": "change_pct", "ref": movement["ref"], "value": movement["pct"],
                                "text": f"{movement['pct'] * 100:+.1f}%"})
        decision = session.selection.by_key().get(key) if session.selection is not None else None
        rows.append({
            "account": key, "in_scope": in_scope, "route": getattr(decision, "route", None),
            "statement": contract.statement_type, "latest_period": latest, "figures": figures,
            "flags": _status_flags(contract, latest),
            "replaced_periods": [c["period"] for c in contract.checks if c.get("overridden")],
            "movement_note": (None if movement is None or movement["valid"] else movement["invalid_reason"]),
            "issues": [i for i in contract.issues if i["severity"] != "info"],
        })
    lines = []
    for row in rows:
        if "error" in row:
            lines.append(f"{row['account']}：{row['error']}")
            continue
        parts = [f"{row['account']}"]
        if not row.get("in_scope", True):
            parts.append("（未进入评论范围）")
        for fig in row["figures"]:
            label = {"balance": f"{row['latest_period']}余额", "change": f"较{fig.get('from')}变动", "change_pct": "幅度"}[fig["label"]]
            parts.append(f"{label} {fig['text']} [{fig['ref']}]")
        if row.get("movement_note"):
            parts.append(f"变动未计算：{row['movement_note']}")
        if row.get("flags"):
            parts.append("；".join(row["flags"]))
        lines.append(" | ".join(parts))
    return {"rows": rows, "lines": lines, "traceability": traceability(session, rows, lines)}


def traceability(session: Any, rows: List[Dict[str, Any]], lines: List[str]) -> Dict[str, Any]:
    """Every amount and percentage printed in `lines`, against the figures the rows
    declare. A number the rows do not declare, or whose ref does not resolve, is
    untraced and listed."""
    declared: Dict[str, List[float]] = {}
    declared_ratios: Dict[str, List[float]] = {}
    unresolved: List[str] = []
    for row in rows:
        contract = session.contracts.get(row.get("account"))
        for fig in row.get("figures") or []:
            if contract is None or contract.fact(fig["ref"]) is None:
                unresolved.append(f"{row['account']} {fig['ref']}")
            # amounts and ratios are kept apart by what the row says they are, not by how big they are: a
            # movement of +11,365.9% (a balance that went from 1.6万 to 182.7万) is a ratio of 113.7
            (declared_ratios if fig["label"] == "change_pct" else declared).setdefault(row["account"], []).append(abs(fig["value"]))
    printed = traced = 0
    untraced: List[str] = []
    for line in lines:
        account = line.split(" | ")[0]
        values = declared.get(account, [])
        for span in text_amounts(line):
            printed += 1
            if any(abs(abs(span["value"]) - v) <= span["tolerance"] for v in values):
                traced += 1
            else:
                untraced.append(f"{account}: {span['written']}")
        for span in text_percents(line):
            printed += 1
            if any(abs(abs(span["value"]) - v * 100) <= span["tolerance"] for v in declared_ratios.get(account, [])):
                traced += 1
            else:
                untraced.append(f"{account}: {span['written']}")
    return {"numbers_printed": printed, "traced": traced, "untraced": untraced, "unresolved_refs": unresolved,
            "complete": printed == traced and not unresolved}

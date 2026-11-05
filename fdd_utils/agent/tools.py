"""The tools an agent can call, each a wrapper over a capability that already exists.

Every tool takes JSON-able arguments and returns a ToolResult: the source version
it read, the facts it rests on (`basis`), any open issues, and an error when it did
not succeed. Three rules the wrappers enforce, because a model will not:

  * `ok` means the thing asked for HAPPENED. A validation that found a failed
    claim is `ok: false` with the verdicts in `data`; an exception is `ok: false`
    with its type and message; a stale session refuses to answer. Nothing that
    failed is returned as a success.
  * a tool that reads evidence refuses when the workbook has changed under the
    session (STALE_SOURCE) -- the figures it would return may no longer exist.
  * arguments are checked before the capability runs; an unknown tool or a
    missing argument is an error the caller can read, not an exception.
"""
from __future__ import annotations

import os
import time
import traceback
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, List, Optional

from ..ai import claims as cl
from ..ai import evidence_contract as ec
from ..delivery import Finding, banner, decide_delivery, labelled_path
from . import investigate as I
from .session import Candidate, SourceSession, capture_cite_ids, capture_context, claims_digest
from .summary import build_numerical_summary


@dataclass
class ToolResult:
    tool: str
    ok: bool
    data: Dict[str, Any] = field(default_factory=dict)
    error: Optional[Dict[str, str]] = None
    source_version: str = ""
    basis: List[str] = field(default_factory=list)
    issues: List[Dict[str, str]] = field(default_factory=list)
    seconds: float = 0.0

    def to_dict(self) -> Dict[str, Any]:
        # Order matters: a conversation truncates a long result from the END, so what a reader
        # must not lose (the verdict, the issues, the data) comes first and the ref list last.
        return {"tool": self.tool, "ok": self.ok, "error": self.error, "issues": self.issues, "data": self.data,
                "basis": self.basis, "source_version": self.source_version, "seconds": round(self.seconds, 3)}


class ToolError(Exception):
    def __init__(self, code: str, message: str, data: Optional[Dict[str, Any]] = None):
        super().__init__(message)
        self.code, self.message, self.data = code, message, data or {}


@dataclass
class Tool:
    name: str
    help: str
    args: Dict[str, Dict[str, Any]]            # arg -> {"type": str, "required": bool, "help": str}
    fn: Callable[..., Dict[str, Any]]
    reads_evidence: bool = True


_TYPES = {"str": str, "list": list, "dict": dict, "int": int, "bool": bool}


def _account(session: SourceSession, name: Any) -> str:
    key = session.resolve_account(str(name or ""))
    if key is None:
        choices = [k for k in session.contracts if str(name or "").strip() and str(name).strip().lower() in k.lower()][:8]
        raise ToolError("UNKNOWN_ACCOUNT", f"no single account matches {name!r}",
                        {"accounts": choices or list(session.contracts)[:40]})
    return key


# ---------------------------------------------------------------------------
# The tools
# ---------------------------------------------------------------------------

#: Below this share of accounts tying to Financials on their latest period the sheet probably belongs to
#: another entity or period (workbook/authority.py uses the same floor before it will replace any total).
PAIRING_FLOOR = 0.7


def financials_pairing(session: SourceSession) -> Dict[str, Any]:
    compared = agreeing = 0
    for contract in session.contracts.values():
        if not contract.periods:
            continue
        check = contract.check(contract.periods[-1], ec.CHECK_TOTAL_VS_FINANCIALS) or {}
        if check.get("status") in (ec.PASS, ec.MISMATCH):
            compared += 1
            agreeing += check["status"] == ec.PASS
    ratio = (agreeing / compared) if compared else None
    return {"compared": compared, "agreeing_latest": agreeing, "ratio": ratio,
            "suspect": compared >= 3 and ratio is not None and ratio < PAIRING_FLOOR}


def inspect_sources(session: SourceSession) -> Dict[str, Any]:
    """What the session read, how far it can be trusted, and what is left out."""
    sheets_total = len((session.digest or {}).get("sheets") or {})   # digest-ok: counts, not evidence
    hidden = list((session.digest or {}).get("hidden_sheets") or [])
    accounts = []
    for key, contract in session.contracts.items():
        decision = session.selection.by_key().get(key) if session.selection is not None else None
        matrix = contract.status_matrix()
        counts: Dict[str, int] = {}
        for row in matrix.values():
            for status in row.values():
                counts[status] = counts.get(status, 0) + 1
        accounts.append({
            "account": key, "in_scope": session.in_scope(key),
            "route": getattr(decision, "route", None), "reason": getattr(decision, "reason", None),
            "periods": contract.periods, "checks": counts,
            "replaced_periods": sorted({c["period"] for c in contract.checks if c.get("overridden")}),
            "issues": [i["code"] for i in contract.issues if i["severity"] != "info"],
        })
    selection = session.selection
    return {
        "financials_pairing": financials_pairing(session),
        "workbook": os.path.basename(session.workbook_path) or "(in memory)",
        "entity": session.entity_name, "language": session.language,
        "source_version": session.source_version, "statement_source": session.binding.statement_source or "unknown",
        "financials_sheet": session.financials_sheet, "sheets": {"total": sheets_total, "hidden": hidden},
        "stale": session.stale_reason(),
        "selection": None if selection is None else {
            "admitted": len(selection.admitted_keys),
            "excluded": [{"account": d.key, "reason": d.reason} for d in selection.excluded if d.reason != "no_activity"],
            "unmatched_financials": selection.unmatched_financials,
            "reconciliation_available": selection.reconciliation_available},
        "accounts": accounts,
        "base_commentary": sorted(k for k in (session.base_results or {}) if not str(k).startswith("__")) or None,
    }


def get_account_evidence(session: SourceSession, account: str, standard: bool = True) -> Dict[str, Any]:
    key = _account(session, account)
    contract = session.contracts[key]
    if standard:
        cl.standard_derived(contract)
    view = cl.evidence_view(contract)
    view["in_scope"] = session.in_scope(key)
    return view


def compute_analysis(session: SourceSession, account: str, op: str, period: str = "", period_from: str = "",
                     line: str = "", refs: Optional[List[str]] = None) -> Dict[str, Any]:
    key = _account(session, account)
    contract = session.contracts[key]
    op = str(op or "").strip().lower()
    if op == "movement":
        item = cl.derive_movement(contract, period_from, period)
    elif op == "share":
        item = cl.derive_share(contract, line, period)
    elif op == "sum":
        item = cl.derive_sum(contract, list(refs or []))
    elif op == "difference":
        item = cl.derive_difference(contract, period)
    elif op == "annualized":
        item = cl.derive_annualized(contract, period)
    elif op == "standard":
        cl.standard_derived(contract)
        return {"derived": contract.derived}
    else:
        raise ToolError("BAD_ARGS", f"unknown op {op!r}", {"ops": ["movement", "share", "sum", "difference", "annualized", "standard"]})
    if not item.get("valid", True):
        raise ToolError("NOT_COMPUTABLE", item.get("invalid_reason") or "not computable", {"derived": item})
    return {"derived": item}


def trace_claim(session: SourceSession, account: str, ref: str = "", claim: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    key = _account(session, account)
    contract = session.contracts[key]
    if ref:
        note = contract.note(ref)
        if note is not None:
            return {"found": True, "note": note, "source": {"sheet": note.get("sheet"), "period": note.get("period")},
                    "binding": {"source_version": contract.source_version}}
        trace = ec.trace_fact(contract, ref)
        if not trace.get("found"):
            raise ToolError("UNKNOWN_REF", f"no fact {ref!r} for {key}")
        return trace
    if not claim:
        raise ToolError("BAD_ARGS", "give a ref or a claim")
    verdict = cl.verify_claim(dict(claim, account=key), session.contracts)
    chain = [ec.trace_fact(contract, r) for r in claim.get("cites", [])]
    return {"verdict": verdict, "chain": chain}


def validate_candidate(session: SourceSession, account: str, claims: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Independent verification: deterministic, and not the author grading itself."""
    key = _account(session, account)
    if not isinstance(claims, list) or not claims:
        raise ToolError("BAD_ARGS", "claims must be a non-empty list")
    stamped = [dict(c, account=key, source_version=c.get("source_version") or session.source_version) for c in claims]
    outcome = cl.verify_candidate(stamped, session.contracts, default_account=key)
    cite_ids = capture_cite_ids(session.contracts[key], stamped)
    session.candidates[key] = Candidate(key, stamped, outcome, session.source_version, claims_digest(stamped), cite_ids=cite_ids,
                                        context=capture_context(session.contracts[key], stamped, cite_ids))
    legacy = {key: cl.to_legacy_result(key, stamped, outcome),
              "__run_health__": {"zero_successful_calls": False, "accounts_on_fallback_bullet": [], "accounts_on_passthrough": [],
                                 "accounts_on_error_text": [], "accounts_without_prompt": []}}
    preview = decide_delivery(ai_results=legacy, scope="account")
    data = {"outcome": outcome, "text": cl.render_claims(stamped), "delivery_preview": preview.status,
            "claims_digest": session.candidates[key].claims_digest}
    if not outcome["ok"]:
        raise ToolError("VALIDATION_FAILED",
                        f"{outcome['failed']} claim(s) failed and {outcome['unverified']} could not be verified, of {len(outcome['verdicts'])}: " +
                        ", ".join(f"{c}x{n}" for c, n in outcome["by_code"].items()), data)
    return data


def numerical_summary(session: SourceSession, accounts: Optional[List[str]] = None) -> Dict[str, Any]:
    summary = build_numerical_summary(session, accounts)
    if not summary["traceability"]["complete"]:
        raise ToolError("UNTRACED_NUMBERS", "the summary prints a number nothing backs", summary)
    return summary


def export_candidate(session: SourceSession, out_path: str = "", accounts: Optional[List[str]] = None) -> Dict[str, Any]:
    from .export import export_session_deck
    chosen = [_account(session, a) for a in accounts] if accounts else list(session.candidates)
    if not chosen:
        raise ToolError("NO_CANDIDATE", "nothing to export: validate a candidate first")
    unvalidated = [a for a in chosen if a not in session.candidates or not session.candidates[a].validated
                   or session.candidates[a].source_version != session.source_version]
    if unvalidated:
        raise ToolError("NOT_VALIDATED", "these candidates are not validated against the current sources: " + ", ".join(unvalidated),
                        {"accounts": unvalidated})
    return export_session_deck(session, chosen, out_path or "")


def refresh_sources(session: SourceSession) -> Dict[str, Any]:
    """Read the same workbook again and say what is different from the reading before it."""
    from . import updates as U
    try:
        before = session.refresh()
    except RuntimeError as exc:
        raise ToolError("CANNOT_REFRESH", str(exc))
    diff = U.diff_sources(before, session.contracts)
    effects = session.workspace.after_refresh(before, diff) if session.workspace is not None else None
    return {"source_version": session.source_version, "diff": diff, "candidate_effects": effects,
            "candidates_to_redo": sorted(a for a, c in session.candidates.items() if not c.validated)}


def diff_sources(session: SourceSession, baseline_deck: str = "") -> Dict[str, Any]:
    """What changed in the sources against a baseline -- and an error, not a guess, when there is none."""
    from . import updates as U
    from ..pptx import lineage as L
    if baseline_deck:
        internal, public = L.read_internal(baseline_deck), L.read_manifest(baseline_deck)
        if internal and internal.get("baseline"):
            return U.diff_sources(internal["baseline"], session.contracts)
        if public:
            return U.diff_by_hash(public, session.contracts)
    elif session.baseline is not None:
        return U.diff_sources(session.baseline, session.contracts)
    raise ToolError("NO_BASELINE", "there is no baseline for these sources (no earlier reading in this session, and no deck "
                    "with a manifest or lineage sidecar was given), so what changed since an earlier upload cannot be said")


def read_deck(session: SourceSession, path: str) -> Dict[str, Any]:
    from . import updates as U
    if not os.path.exists(path):
        raise ToolError("NO_SUCH_FILE", f"no deck at {path}")
    return U.analyze_returned(path, session)


def update_deck(session: SourceSession, path: str, out_path: str = "") -> Dict[str, Any]:
    """Edit a returned deck in place (to a copy): only paragraphs the user left alone whose cited facts
    changed, and only from validated candidates. Conflicts and relink cases are reported, untouched."""
    from . import updates as U
    from .export import layout_problems
    from ..pptx import lineage as L
    if not os.path.exists(path):
        raise ToolError("NO_SUCH_FILE", f"no deck at {path}")
    report = U.analyze_returned(path, session)
    if report["mode"] == "none":
        raise ToolError("NO_MANIFEST", "the deck has no manifest: nothing can be edited with confidence; inspect and relink first", report)
    public = L.read_manifest(path)
    by_id = {e["id"]: e for e in public["elements"]}
    texts, claims_by_element = {}, {}
    for action in report["actions"]:
        if action["action"] != "auto_update":
            continue
        element = by_id[action["element"]]
        candidate = session.candidates.get(element["account"])
        if candidate is not None and candidate.validated and candidate.source_version == session.source_version:
            from ..ai.claims import render_claims
            texts[element["id"]] = render_claims(candidate.claims)
            claims_by_element[element["id"]] = candidate.claims
    target = out_path or os.path.splitext(path)[0] + "_updated.pptx"
    result = U.apply_updates(path, session, report, texts, target)
    waiting = [a for a in report["actions"] if a["action"] in ("conflict", "inspect_and_relink")] + \
        [a for a in report["actions"] if a["action"] == "auto_update" and a["element"] not in {x["element"] for x in result["applied"]}]
    if os.path.exists(path + L.INTERNAL_SUFFIX):
        import shutil
        shutil.copyfile(path + L.INTERNAL_SUFFIX, target + L.INTERNAL_SUFFIX)
    L.restamp(target, session, result["applied"], texts, claims_by_element, fully_reconciled=not waiting)
    layout = layout_problems(target, result.get("touched_slides") or [])
    out = {"report": report["summary"], "applied": result["applied"], "skipped": result["skipped"], "out_path": target,
           "touched_slides": result.get("touched_slides"), "layout": layout, "waiting_for_a_person": len(waiting),
           "needs_candidates": sorted({by_id[a["element"]]["account"] for a in report["actions"]
                                       if a["action"] == "auto_update" and a["element"] not in texts})}
    if layout["problems"]:
        raise ToolError("LAYOUT_PROBLEM", f"{len(layout['problems'])} layout problem(s) on the updated slides", out)
    return out


_TOOLS: List[Tool] = [
    Tool("inspect_sources", "What was read: source version, statement source, which accounts are in scope and why not, hidden sheets, staleness.",
         {}, inspect_sources, reads_evidence=False),
    Tool("get_account_evidence", "One account's facts by basis (F refs), derived facts (D refs), notes (N refs), checks per period and what each period permits.",
         {"account": {"type": "str", "required": True, "help": "account key or a unique part of it"},
          "standard": {"type": "bool", "required": False, "help": "also compute each period-to-period movement"}}, get_account_evidence),
    Tool("compute_analysis", "Have Python compute a figure from facts: movement | share | sum | difference | annualized | standard. Returns a D ref to cite.",
         {"account": {"type": "str", "required": True, "help": ""}, "op": {"type": "str", "required": True, "help": "movement|share|sum|difference|annualized|standard"},
          "period": {"type": "str", "required": False, "help": "ISO period (the later one for movement)"},
          "period_from": {"type": "str", "required": False, "help": "movement: the earlier ISO period"},
          "line": {"type": "str", "required": False, "help": "share: the detail line"},
          "refs": {"type": "list", "required": False, "help": "sum: every F ref to add"}}, compute_analysis),
    Tool("trace_claim", "Trace a fact ref, or a claim, back to its sheet and row and the checks that touch it.",
         {"account": {"type": "str", "required": True, "help": ""}, "ref": {"type": "str", "required": False, "help": "F/D ref"},
          "claim": {"type": "dict", "required": False, "help": "a claim object to verify"}}, trace_claim),
    Tool("validate_candidate", "Check a paragraph of claims against what they cite. ok=false lists each failing claim and why.",
         {"account": {"type": "str", "required": True, "help": ""},
          "claims": {"type": "list", "required": True, "help": "[{kind, text, periods, cites, cause_cites, inference}]"}}, validate_candidate),
    Tool("numerical_summary", "Every account's latest balance and movement, each figure with its ref and check status.",
         {"accounts": {"type": "list", "required": False, "help": "default: all"}}, numerical_summary),
    Tool("export_candidate", "Export a deck from validated candidates (replacing those accounts in the base commentary). Refuses unvalidated or stale candidates.",
         {"out_path": {"type": "str", "required": False, "help": ""}, "accounts": {"type": "list", "required": False, "help": ""}},
         export_candidate),
]
_TOOLS += [
    Tool("refresh_sources", "Read the same workbook again (it changed on disk) and list what is different from the previous reading. Earlier drafts must be redone.",
         {}, refresh_sources, reads_evidence=False),
    Tool("diff_sources", "What changed in the sources against a baseline: the previous reading in this session, or an exported deck's lineage. Errors with NO_BASELINE if neither exists.",
         {"baseline_deck": {"type": "str", "required": False, "help": "an exported deck whose manifest/lineage is the baseline"}}, diff_sources),
    Tool("read_deck", "Read a returned deck against its manifest: what the user edited, moved, copied or deleted, and what the sources changed.",
         {"path": {"type": "str", "required": True, "help": "path to the .pptx"}}, read_deck),
    Tool("update_deck", "Update a returned deck's paragraphs whose cited facts changed (from validated candidates); conflicts and unlocated paragraphs are reported, not touched.",
         {"path": {"type": "str", "required": True, "help": ""}, "out_path": {"type": "str", "required": False, "help": ""}}, update_deck),
]
_TOOLS += [
    Tool("read_source_block", "The workbook's own cells around a fact (ref + account), a found remark, or a row of a visible sheet: sheet, cell range, "
         "values as saved, the sheet's header lines and the limits of the read. Context, not a citable fact.",
         {"account": {"type": "str", "required": False, "help": "needed with ref; alone it reads the top of the account's schedule"},
          "ref": {"type": "str", "required": False, "help": "an F or N ref of that account"},
          "sheet": {"type": "str", "required": False, "help": "a visible sheet, with row"},
          "row": {"type": "int", "required": False, "help": "Excel row number (1-based)"},
          "rows_before": {"type": "int", "required": False, "help": "default 2"}, "rows_after": {"type": "int", "required": False, "help": "default 8"},
          "max_cols": {"type": "int", "required": False, "help": "default 14"}}, I.read_source_block),
    Tool("search_notes", "Find remarks on the VISIBLE sheets that mention a word or phrase. Name the account and each hit is filed on it as a note "
         "(N ref) a cause can cite. Returns sheet, cell, the cell's text, and what was and was not searched.",
         {"query": {"type": "str", "required": True, "help": "words to find, e.g. a counterparty, a policy, 'adjustment'"},
          "account": {"type": "str", "required": False, "help": "file the hits on this account"},
          "max_hits": {"type": "int", "required": False, "help": "default 12, at most 30"}}, I.search_notes),
    Tool("cross_account_evidence", "Which other accounts the numbers tie this one to (tie-outs, shared names, ratios), with the test and its evidence; "
         "and the relationships that were looked for and did NOT tie, with what was compared and why (failed_checks: they explain a gap, they never support a link).",
         {"account": {"type": "str", "required": True, "help": ""},
          "max_failed": {"type": "int", "required": False, "help": "failed checks to list, default 12, at most 40"}}, I.cross_account_evidence),
    Tool("adjustment_evidence", "Per period: the tab's total, the Financials total, the one the deck uses, the difference, the adjustment entries (N refs) "
         "behind it, whether they add up, and whether the difference has no explanation on the visible sheets.",
         {"account": {"type": "str", "required": True, "help": ""}}, I.adjustment_evidence),
]
TOOLS: Dict[str, Tool] = {t.name: t for t in _TOOLS}


def register_tool(tool: Tool) -> None:
    TOOLS[tool.name] = tool


def describe_tools() -> str:
    lines = []
    for tool in TOOLS.values():
        args = ", ".join(f"{a}{'' if spec.get('required') else '?'}: {spec['type']}" for a, spec in tool.args.items())
        lines.append(f"  - {tool.name}({args}): {tool.help}")
    return "\n".join(lines)


def call_tool(session: SourceSession, name: str, args: Optional[Dict[str, Any]] = None) -> ToolResult:
    """Run one tool. Never raises; never reports a failure as a success."""
    started = time.perf_counter()
    args = dict(args or {})
    result = ToolResult(tool=str(name), ok=False, source_version=session.source_version)

    def fail(code: str, message: str, data: Optional[Dict[str, Any]] = None) -> ToolResult:
        result.error = {"code": code, "message": message}
        result.data = data or {}
        result.seconds = time.perf_counter() - started
        return result

    tool = TOOLS.get(str(name))
    if tool is None:
        return fail("UNKNOWN_TOOL", f"no tool {name!r}", {"tools": list(TOOLS)})
    unknown = [a for a in args if a not in tool.args]
    if unknown:
        return fail("BAD_ARGS", f"unknown argument(s) {unknown} for {tool.name}", {"expected": list(tool.args)})
    for arg, spec in tool.args.items():
        if spec.get("required") and arg not in args:
            return fail("BAD_ARGS", f"{tool.name} needs '{arg}'", {"expected": list(tool.args)})
        if arg in args and args[arg] is not None and not isinstance(args[arg], _TYPES[spec["type"]]):
            return fail("BAD_ARGS", f"'{arg}' must be a {spec['type']}")
    stale = session.stale_reason()
    if stale and tool.reads_evidence:
        return fail("STALE_SOURCE", stale, {"hint": "call refresh_sources"})
    try:
        data = tool.fn(session, **args)
    except ToolError as exc:
        result = fail(exc.code, exc.message, exc.data)
        session.log(tool=tool.name, ok=False, code=exc.code)
        return result
    except Exception as exc:                                   # noqa: BLE001 -- a tool crash is a result, not a crash
        session.log(tool=tool.name, ok=False, code="TOOL_EXCEPTION")
        return fail("TOOL_EXCEPTION", f"{type(exc).__name__}: {exc}", {"trace": traceback.format_exc()[-600:]})
    result.ok, result.data = True, data
    result.basis = sorted(_refs_in(data))[:60]
    if stale:
        result.issues.append({"code": "STALE_SOURCE", "message": stale})
    result.seconds = time.perf_counter() - started
    session.log(tool=tool.name, ok=True)
    return result


def _refs_in(data: Any, depth: int = 0) -> set:
    """Fact/derived/note refs that appear in a tool's data."""
    found: set = set()
    if depth > 6:
        return found
    if isinstance(data, dict):
        ref = data.get("ref")
        if isinstance(ref, str) and ref[:1] in "FDN" and ref[1:].isdigit():
            found.add(ref)
        for value in data.values():
            found |= _refs_in(value, depth + 1)
    elif isinstance(data, list):
        for value in data[:400]:
            found |= _refs_in(value, depth + 1)
    return found

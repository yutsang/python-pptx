"""Is this deck something to send? ready / review_required / blocked, and why.

Before this module the only gate on a deck was ``zero_successful_calls`` (no
model output at all). ``safe_to_export`` was computed and read by nothing, the
feedback loop's arbiter kept "the attempt with the fewest hallucinations" and
called the account terminal, and a deck built from schedule tabs alone looked
like a deck reconciled against Financials. A deck with an unresolved invented
figure in it could be exported, named like any other, and handed over.

This is the decision, kept apart from the pipeline that produced the text: it
reads what a run left behind (results, run state, reconciliation selection,
statement source) and answers for the whole deck and for each account.

  ready            nothing found that a reader of the deck would need told
  review_required  the deck can be read, but something a human should look at
                   changes what it says or leaves out (an account left out of
                   the deck for lack of a reconciliation row, a statement
                   table built from the tabs themselves, a degraded bullet)
  blocked          not deliverable: an unresolved invented figure, placeholder
                   text, no model output, a workbook that is not the run's

A blocked or review_required deck is still exported, as a DRAFT: sometimes the
draft is exactly what the user wants to read. What it cannot be is unlabelled.
The label is a file-name suffix, a sidecar ``<deck>.delivery.json`` and the
deck's own document properties; it never changes a slide.

Retry exhaustion does not clear anything. The arbiter chooses the least bad
attempt; if that attempt still carries an unresolved hallucination the account
is ``blocked``, whichever attempt won.
"""
from __future__ import annotations

import json
import os
import subprocess
import time
from dataclasses import dataclass, field
from typing import Any, Dict, Iterable, List, Optional

STATUS_READY = "ready"
STATUS_REVIEW = "review_required"
STATUS_BLOCKED = "blocked"
_RANK = {STATUS_READY: 0, STATUS_REVIEW: 1, STATUS_BLOCKED: 2}
_SEVERITY_TO_STATUS = {"info": STATUS_READY, "review": STATUS_REVIEW, "blocked": STATUS_BLOCKED}

SIDECAR_SUFFIX = ".delivery.json"
SENTINEL_PREFIX = "__"


@dataclass
class Finding:
    code: str
    severity: str            # "info" | "review" | "blocked"
    message: str
    account: str = ""        # "" = the deck as a whole

    def to_dict(self) -> Dict[str, str]:
        return {"code": self.code, "severity": self.severity, "account": self.account, "message": self.message}


@dataclass
class DeliveryDecision:
    status: str = STATUS_READY
    findings: List[Finding] = field(default_factory=list)
    #: what the decision was made from; an input that was not supplied is
    #: named here instead of being assumed clean
    inputs: Dict[str, bool] = field(default_factory=dict)
    assessed: bool = True

    def add(self, code: str, severity: str, message: str, account: str = "") -> None:
        self.findings.append(Finding(code, severity, message, account))
        self.status = _worse(self.status, _SEVERITY_TO_STATUS[severity])

    @property
    def accounts(self) -> Dict[str, str]:
        out: Dict[str, str] = {}
        for f in self.findings:
            if f.account:
                out[f.account] = _worse(out.get(f.account, STATUS_READY), _SEVERITY_TO_STATUS[f.severity])
        return out

    def with_severity(self, severity: str) -> List[Finding]:
        return [f for f in self.findings if f.severity == severity]

    @property
    def deliverable(self) -> bool:
        return self.status == STATUS_READY

    def reasons(self, limit: int = 0) -> List[str]:
        ordered = sorted(self.findings, key=lambda f: ("blocked", "review", "info").index(f.severity))
        lines = [f"[{f.severity}] {f.account + ': ' if f.account else ''}{f.message}" for f in ordered]
        return lines[:limit] if limit else lines

    def to_dict(self) -> Dict[str, Any]:
        return {
            "status": self.status,
            "deliverable": self.deliverable,
            "assessed": self.assessed,
            "inputs": dict(self.inputs),
            "accounts": self.accounts,
            "findings": [f.to_dict() for f in self.findings],
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "DeliveryDecision":
        decision = cls(status=STATUS_READY, inputs=dict(data.get("inputs") or {}), assessed=bool(data.get("assessed", True)))
        for item in data.get("findings") or []:
            decision.add(str(item.get("code")), str(item.get("severity")), str(item.get("message")),
                         str(item.get("account") or ""))
        return decision


def _worse(a: str, b: str) -> str:
    return a if _RANK[a] >= _RANK[b] else b


# ---------------------------------------------------------------------------
# The decision
# ---------------------------------------------------------------------------

def decide_delivery(
    *,
    selection: Any = None,
    ai_results: Optional[Dict[str, Any]] = None,
    bs_is_results: Optional[Dict[str, Any]] = None,
    identity_ok: Optional[bool] = None,
    identity_note: str = "",
    extra_findings: Optional[Iterable[Finding]] = None,
    expected_accounts: Optional[Iterable[str]] = None,
    scope: str = "deck",
) -> DeliveryDecision:
    """Combine what the run left behind into one answer.

    Silence is not a clean bill. A result set with nothing in it is ``blocked`` (there is no
    commentary to deliver), an account the selection admitted but the results lack is a finding
    (``expected_accounts`` defaults to the selection's admitted keys), and a missing statement
    source asks for review. ``scope="account"`` skips the deck-level checks (coverage, inputs) for a
    preview of one account's own text."""
    decision = DeliveryDecision(inputs={
        "selection": selection is not None,
        "ai_results": ai_results is not None,
        "bs_is_results": bs_is_results is not None,
        "identity": identity_ok is not None,
    })
    _selection_findings(decision, selection)
    _statement_findings(decision, bs_is_results, selection)
    _account_findings(decision, ai_results)
    if scope == "deck":
        if bs_is_results is None:
            decision.add("INPUT_MISSING", "review", "no statement source was supplied, so the deck's BS/IS tables and where they came from are unknown")
        if selection is None:
            decision.add("SELECTION_NOT_AVAILABLE", "info", "no account-selection record: which accounts were left out of the deck cannot be told")
        expected = list(expected_accounts) if expected_accounts is not None else (
            list(selection.admitted_keys) if selection is not None and getattr(selection, "reconciliation_available", False) else [])
        _coverage_findings(decision, ai_results, expected)
    if identity_ok is False:
        decision.add("IDENTITY_MISMATCH", "blocked",
                     identity_note or "the workbook or Financials sheet is not the one this run was made from")
    for finding in extra_findings or ():
        decision.add(finding.code, finding.severity, finding.message, finding.account)
    return decision


def combine_decisions(items: Iterable[Any]) -> DeliveryDecision:
    """One decision for a deck made of several decks: the worst status of any part.

    ``items`` are ``(label, DeliveryDecision-or-dict-or-None)``; a part with no
    decision counts as unassessed. Each finding is kept, with its account
    prefixed by the part it came from.
    """
    combined = DeliveryDecision()
    parts = 0
    for label, part in items:
        parts += 1
        if isinstance(part, dict):
            part = DeliveryDecision.from_dict(part)
        if part is None:
            part = unassessed_decision(f"{label}: no delivery decision was recorded for this part")
        for f in part.findings:
            combined.add(f.code, f.severity, f.message, f"{label}/{f.account}" if f.account else str(label))
        combined.assessed = combined.assessed and part.assessed
    combined.inputs = {"parts": parts}
    return combined


def unassessed_decision(reason: str = "the caller did not supply a delivery decision") -> DeliveryDecision:
    """For an export path nothing assessed: fail closed, say so."""
    decision = DeliveryDecision(assessed=False)
    decision.add("UNASSESSED", "review", reason)
    return decision


def _selection_findings(decision: DeliveryDecision, selection: Any) -> None:
    if selection is None:
        return
    for d in getattr(selection, "decisions", []) or []:
        if d.admitted:
            if d.route == "prior_period":
                decision.add("SELECTION_PRIOR_PERIOD", "review",
                             "latest column is a partial period that differs from the tab; only the previous full period ties out",
                             d.key)
            elif d.route == "unchecked_history":
                decision.add("SELECTION_UNCHECKED_HISTORY", "info",
                             "Financials line was not compared in the latest period (" + d.reason + "); admitted on earlier activity",
                             d.key)
            continue
        if d.reason == "no_activity":
            continue  # a dormant tab has nothing to say and nothing to be told about
        code = {"bs_diff": "SELECTION_BS_DIFF", "no_reconciliation_row": "SELECTION_NO_ROW",
                "no_data": "SELECTION_NO_DATA"}.get(d.reason, "SELECTION_EXCLUDED")
        decision.add(code, "review", f"left out of the deck: {d.detail}", d.key)
    for item in getattr(selection, "unmatched_financials", []) or []:
        decision.add("SELECTION_UNMATCHED_FINANCIALS", "review",
                     f"{item.get('statement')} line '{item.get('financials_account')}' has a balance but no schedule tab was found; the deck has no commentary for it")


def _accounts_in(ai_results: Optional[Dict[str, Any]]) -> List[str]:
    return [k for k, v in (ai_results or {}).items() if not str(k).startswith(SENTINEL_PREFIX) and isinstance(v, dict)]


def _coverage_findings(decision: DeliveryDecision, ai_results: Optional[Dict[str, Any]], expected: List[str]) -> None:
    if ai_results is None:
        return                                           # already blocked as NO_RESULTS
    present = _accounts_in(ai_results)
    if not present:
        decision.add("NO_CONTENT", "blocked", "the results hold no account commentary at all")
        return
    missing = [a for a in expected if a not in present]
    if missing and len(missing) == len(expected):
        decision.add("NO_CONTENT", "blocked", f"none of the {len(expected)} account(s) the selection admitted has commentary")
        return
    if len(missing) > 3:
        decision.add("ACCOUNT_MISSING", "review", f"{len(missing)} of the {len(expected)} accounts admitted to commentary have none in the results: "
                     + ", ".join(missing[:8]) + (" ..." if len(missing) > 8 else ""))
    else:
        for account in missing:
            decision.add("ACCOUNT_MISSING", "review", "admitted to commentary, but no commentary for it is in the results", account)


def _statement_findings(decision: DeliveryDecision, bs_is_results: Optional[Dict[str, Any]], selection: Any) -> None:
    source = (bs_is_results or {}).get("statement_source") if isinstance(bs_is_results, dict) else None
    no_financials = selection is not None and not getattr(selection, "reconciliation_available", True)
    if source == "synthesized" or (source is None and no_financials and bs_is_results):
        decision.add("STATEMENT_SYNTHESIZED", "review",
                     "the BS/IS tables are built from the schedule tabs themselves; no independent Financials sheet was reconciled")


def _account_findings(decision: DeliveryDecision, ai_results: Optional[Dict[str, Any]]) -> None:
    if ai_results is None:
        decision.add("NO_RESULTS", "blocked", "no commentary results were supplied")
        return
    from .ai.pipeline import PHASE_FAILED, count_defective_clauses
    health = ai_results.get("__run_health__") if isinstance(ai_results, dict) else None
    if isinstance(health, dict) and health.get("zero_successful_calls"):
        decision.add("NO_MODEL_OUTPUT", "blocked",
                     "no model call succeeded in this run; every bullet is deterministic filler")
    run = ai_results.get("__run__") if isinstance(ai_results, dict) else None
    states = (run or {}).get("accounts") or {} if isinstance(run, dict) else {}
    health = health if isinstance(health, dict) else {}
    for key in health.get("accounts_on_error_text") or []:
        decision.add("ACCOUNT_PLACEHOLDER_TEXT", "blocked", "the commentary is an error placeholder, not commentary", key)
    for key in health.get("accounts_on_fallback_bullet") or []:
        decision.add("ACCOUNT_FALLBACK_BULLET", "review", "deterministic fallback bullet, not model commentary", key)
    for key in health.get("accounts_on_passthrough") or []:
        decision.add("ACCOUNT_PASSTHROUGH", "review", "a stage failed and its input was passed through unchecked", key)
    for key in health.get("accounts_without_prompt") or []:
        decision.add("ACCOUNT_NO_PROMPT", "review", "no prompt could be built, so the account has no model commentary", key)
    for key, result in ai_results.items():
        if str(key).startswith(SENTINEL_PREFIX) or not isinstance(result, dict):
            continue
        validation = result.get("agent_4_validation") or {}
        reviews = validation.get("clause_reviews") or []
        phase = str((states.get(key) or {}).get("phase") or "")
        if not str(result.get("final") or result.get("subagent_4") or result.get("subagent_2") or result.get("subagent_1") or "").strip():
            decision.add("ACCOUNT_EMPTY_TEXT", "blocked", "the account's commentary is empty", key)
            continue
        if not reviews:
            # Text nothing was checked against is unverified, whatever phase the
            # run recorded for it (the sweep itself files this case as FAILED).
            why = "failed" if phase == PHASE_FAILED else ("empty" if not (result.get("final") or result.get("subagent_4")) else "never grounded")
            decision.add("ACCOUNT_UNGROUNDED", "blocked",
                         f"no clause was grounded against the source ({why}); the text is unverified", key)
            continue
        defective = count_defective_clauses(reviews)
        if defective:
            arbiter = result.get("feedback_arbiter") or {}
            tail = ""
            if arbiter:
                tail = (f"; retries were exhausted and the arbiter kept '{arbiter.get('chosen')}' "
                        f"(scores {arbiter.get('scores')})")
            elif result.get("feedback_retries"):
                tail = f"; {result.get('feedback_retries')} retry(ies) did not clear it"
            first = defective[0]
            decision.add("ACCOUNT_RESIDUAL_HALLUCINATION", "blocked",
                         f"{len(defective)} unresolved unsupported figure/fact clause(s), e.g. "
                         f"\"{str(first.get('clause') or '')[:60]}\" ({first.get('reason') or first.get('code') or 'unsupported'})" + tail,
                         key)
        else:
            reasoning = [r for r in reviews if isinstance(r, dict) and not r.get("supported")]
            if reasoning:
                decision.add("ACCOUNT_REASONING_FLAGS", "info",
                             f"{len(reasoning)} inference clause(s) marked as reasoning in the deck", key)


# ---------------------------------------------------------------------------
# The label
# ---------------------------------------------------------------------------

def delivery_suffix(decision: Optional[DeliveryDecision]) -> str:
    """Appended to the deck's file stem; empty for a deliverable deck."""
    if decision is None:
        return "_REVIEW"
    return {STATUS_READY: "", STATUS_REVIEW: "_REVIEW", STATUS_BLOCKED: "_DRAFT-BLOCKED"}[decision.status]


def labelled_path(path: str, decision: Optional[DeliveryDecision]) -> str:
    """``X.pptx`` -> ``X_REVIEW.pptx`` / ``X_DRAFT-BLOCKED.pptx``; unchanged when ready or already labelled."""
    suffix = delivery_suffix(decision)
    stem, ext = os.path.splitext(path)
    if not suffix or stem.endswith(suffix):
        return path
    return stem + suffix + ext


_LABELS = ("_DRAFT-BLOCKED", "_REVIEW")


def strip_label(path: str) -> str:
    stem, ext = os.path.splitext(path)
    for label in _LABELS:
        if stem.endswith(label):
            return stem[: -len(label)] + ext
    return path


def relabel_deck(path: str, decision: "DeliveryDecision") -> str:
    """Rewrite a deck's label after a check found something the first decision did not
    know (a layout problem, an edited account that never reached a slide): new file
    name, document properties and sidecar; the old file and sidecar are removed."""
    from pptx import Presentation
    new_path = labelled_path(strip_label(path), decision)
    presentation = Presentation(path)
    apply_to_presentation(presentation, decision)
    presentation.save(new_path)
    if os.path.abspath(new_path) != os.path.abspath(path):
        # files that belong to the deck by name travel with it (the lineage sidecar is read back by path)
        for suffix in (".lineage.json",):
            if os.path.exists(path + suffix):
                os.replace(path + suffix, new_path + suffix)
        for old in (path, sidecar_path(path)):
            if os.path.exists(old):
                os.remove(old)
    write_sidecar(new_path, decision)
    return new_path


def sidecar_path(deck_path: str) -> str:
    return deck_path + SIDECAR_SUFFIX


def _code_revision() -> str:
    try:
        root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        out = subprocess.run(["git", "-C", root, "rev-parse", "--short", "HEAD"], capture_output=True, text=True, timeout=5)
        return out.stdout.strip() or "unknown"
    except Exception:
        return "unknown"


#: python-pptx refuses a core property longer than this.
_PROPERTY_LIMIT = 250


def document_properties(decision: DeliveryDecision) -> Dict[str, str]:
    """Text for the deck's own properties (File > Info), the part that travels with the file.

    Short by necessity (the format allows 255 characters): the status, how many
    findings gate it, and the first reasons. The full list is in the sidecar.
    """
    head = {STATUS_READY: "READY", STATUS_REVIEW: "REVIEW REQUIRED", STATUS_BLOCKED: "DRAFT - NOT DELIVERABLE"}[decision.status]
    gating = [f for f in decision.findings if f.severity in ("blocked", "review")]
    gating.sort(key=lambda f: 0 if f.severity == "blocked" else 1)
    text = head
    if gating:
        text += f" ({len(gating)} finding(s); see .delivery.json): " + " | ".join(
            f"{f.account + ': ' if f.account else ''}{f.code}" for f in gating[:6])
    if len(text) > _PROPERTY_LIMIT:
        text = text[:_PROPERTY_LIMIT - 3] + "..."
    return {"category": head, "comments": text}


def apply_to_presentation(presentation: Any, decision: DeliveryDecision) -> None:
    """Write the label into an open python-pptx Presentation's core properties."""
    props = document_properties(decision)
    core = presentation.core_properties
    core.category = props["category"]
    core.comments = props["comments"]


def write_sidecar(deck_path: str, decision: DeliveryDecision) -> str:
    path = sidecar_path(deck_path)
    payload = decision.to_dict()
    payload["deck"] = os.path.basename(deck_path)
    payload["code_revision"] = _code_revision()
    payload["created"] = time.strftime("%Y-%m-%dT%H:%M:%S")
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=2)
    return path


def read_sidecar(deck_path: str) -> Optional[Dict[str, Any]]:
    path = sidecar_path(deck_path)
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def banner(decision: DeliveryDecision, limit: int = 8) -> str:
    """The text a console or the UI prints beside the file name."""
    head = {STATUS_READY: "DELIVERY: ready",
            STATUS_REVIEW: "DELIVERY: REVIEW REQUIRED -- read the reasons before sending",
            STATUS_BLOCKED: "DELIVERY: BLOCKED -- this is a draft, not a deliverable"}[decision.status]
    lines = [head]
    lines += ["  " + r for r in decision.reasons(limit=limit)]
    extra = len(decision.findings) - limit
    if extra > 0:
        lines.append(f"  ... and {extra} more in the .delivery.json beside the deck")
    return "\n".join(lines)

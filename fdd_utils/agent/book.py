"""A whole databook written by the agent: open the workbook, write every account in scope, write each statement's overview, export one deck.

The app, the command line (fdd_agent.py book) and the trials (ad-hoc/workbench/entity_trial.py) share this path; it adds nothing the
working session (workspace.py) does not already do, in the order the trials measured: Python decides the accounts in scope (the
reconciliation selection), each account is written from its own evidence by WorkSession.draft_entity (the controller reads the evidence,
the model writes once, the verifier checks, a rejected sentence is repaired alone), the executive summary of each statement is written
from the finished paragraphs, and one deck is exported from the accepted paragraphs and the drafts the verifier did not accept (their
rejected sentences marked red or orange, the account named in the delivery decision, which blocks). No text from an earlier run fills an
account the agent did not finish.

Cost control. A model call is paid for. The runaway guard of one book is agent.max_calls_per_account (default 40) times the accounts, not
agent.max_model_calls, which belongs to one ask or edit session; neither is a quota (a cap that stops a model midway costs the run again).
Every other cap in the agent block (tokens, seconds, cost) applies as it does everywhere.
"""
from __future__ import annotations

import hashlib
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

REPO_ROOT = Path(__file__).resolve().parents[2]
#: where an account that ends with a refused sentence is kept as one case (issues.py); git-ignored with the rest of fdd_utils/output
ISSUE_DIR = REPO_ROOT / "fdd_utils" / "output" / "agent_issues"
_CHINESE = ("Chi", "Chn", "chinese", "Chinese")


def code_fingerprint() -> Dict[str, Any]:
    """sha256 of the code a run depends on, line endings normalised: the company folder is not a git checkout, so this is how a run says
    which frozen version it used."""
    files = sorted(list((REPO_ROOT / "fdd_utils" / "agent").glob("*.py"))
                   + [REPO_ROOT / "fdd_utils" / "ai" / n for n in ("claims.py", "evidence_contract.py", "validator.py", "source_index.py")]
                   + [REPO_ROOT / "fdd_utils" / "workbook" / n for n in ("supporting.py", "inspector.py", "statements.py", "schedules.py")]
                   + [REPO_ROOT / "fdd_utils" / "financial_common.py", REPO_ROOT / "fdd_utils" / "report_lengths.yml"])
    total = hashlib.sha256()
    count = 0
    for f in files:
        if not f.exists():
            continue
        count += 1
        total.update(f.relative_to(REPO_ROOT).as_posix().encode())
        total.update(f.read_bytes().replace(b"\r\n", b"\n"))
    return {"sha256": total.hexdigest(), "short": total.hexdigest()[:12], "files": count}


def book_cap(config: Optional[Dict[str, Any]], accounts: int) -> Optional[int]:
    """The runaway guard on model calls for a book of this many accounts (the overviews and a review pass count as two more), or None."""
    from .flag import max_calls_per_account
    per = max_calls_per_account(config)
    return per * (accounts + 2) if per else None


def open_book(workbook: str, *, entity: str = "", sheet: Optional[str] = None, financials_from: Optional[str] = None,
              financials_sheet: Optional[str] = None, mapping_overrides: Optional[Dict[str, str]] = None, model_type: str = "",
              model_name: Optional[str] = None, language: str = "", config: Optional[Dict[str, Any]] = None, model: Any = None,
              escalation: Any = None, limits: Any = None, policy: Any = None, budget: Any = None, thinking: str = "", session_path: Optional[str] = None) -> Any:
    """A WorkSession over the workbook, writing with the model the caller names (the app's dropdown, --model) unless agent.models in
    config.yml names one, then that wins. `model` is a callable (a scripted writer) used instead of any provider."""
    from . import flag as F
    from .workspace import WorkSession
    if config is None:
        config = F._config(None)
    kwargs: Dict[str, Any] = {"budget": budget or F.agent_budget(config), "limits": limits or F.agent_limits(config),
                              "policy": policy or F.agent_policy(config), "config": config, "prepare": None,
                              "thinking": thinking}
    if model is not None:
        kwargs.update(model=model, escalation=escalation, model_name=model_name or "injected")
    else:
        configured = bool(F.role_config(config, "primary").get("provider"))
        kwargs.update(model_type="" if configured else model_type, model_name=None if configured else model_name)
    work = WorkSession.open(workbook, session_path=session_path, entity=entity or "", sheet=None if financials_from else sheet, financials_from=financials_from,
                            financials_sheet=financials_sheet if financials_from else None,
                            mapping_overrides=dict(mapping_overrides or {}) or None, **kwargs)
    if language:
        work.source.language = "Chi" if str(language).strip() in _CHINESE else "Eng"
    return work


@dataclass
class BookResult:
    work: Any
    accounts: List[str]
    rows: List[Dict[str, Any]] = field(default_factory=list)
    stopped: str = ""
    overview: Dict[str, str] = field(default_factory=dict)
    overview_error: str = ""
    export: Optional[Dict[str, Any]] = None
    ai_results: Dict[str, Any] = field(default_factory=dict)
    issue_cases: int = 0
    seconds: float = 0.0

    @property
    def deck(self) -> Dict[str, Any]:
        """The export record (path, delivery, layout) when a deck was written, even a draft one the delivery decision blocks."""
        return (self.export or {}).get("data") or {}

    @property
    def deck_path(self) -> str:
        return str(self.deck.get("path") or "")

    @property
    def delivery(self) -> Dict[str, Any]:
        return self.deck.get("delivery") or {}

    @property
    def error(self) -> Dict[str, Any]:
        return (self.export or {}).get("error") or {}

    @property
    def finished(self) -> List[str]:
        return [a for a, c in self.work.source.candidates.items() if c.accepted and c.validated]

    def books(self) -> Dict[str, Any]:
        return self.work.shared.summary()

    @property
    def marked(self) -> List[str]:
        """Drafts carried into the deck with the sentences the verifier refused marked (red or orange)."""
        from . import tools as T
        return T.drafts_to_mark(self.work.source)

    def summary_lines(self, name: str = "") -> List[str]:
        """What a run did, in a screenful: accounts finished, marked and not drafted, the deck and what blocks it, the books."""
        books = self.books()
        finished, marked = self.finished, self.marked
        lines = [f"BOOK  {name}  source version {self.work.source.source_version}",
                 f"accounts in scope {len(self.accounts)}   finished {len(finished)}   drafts carried with marked sentences {len(marked)}   "
                 f"not drafted {len([a for a in self.accounts if a not in finished and a not in marked])}"]
        if self.stopped:
            lines.append(f"STOPPED: {self.stopped}")
        if self.overview:
            lines.append("executive summary written for: " + ", ".join(sorted(self.overview)))
        elif self.overview_error:
            lines.append(f"executive summary NOT written: {self.overview_error}")
        if self.export is not None:
            if self.deck_path:
                lines.append(f"deck: {self.deck_path}   delivery: {self.delivery.get('status')}")
                for finding in (self.delivery.get("findings") or [])[:8]:
                    if finding.get("severity") in ("blocked", "review"):
                        lines.append(f"  [{finding['severity']}] {finding.get('account') or 'deck'}: {str(finding.get('message'))[:160]}")
            else:
                lines.append(f"NO DECK: [{self.error.get('code')}] {str(self.error.get('message'))[:200]}")
        lines.append(f"model calls {books['calls']}   charged tokens {books['charged_tokens']:,}"
                     + (f"   cost {books['cost']:.2f} yuan" if books.get("cost") is not None else "") + f"   {self.seconds}s")
        if self.issue_cases:
            lines.append(f"{self.issue_cases} account(s) with a refused sentence kept as cases (free to re-check: python ad-hoc/workbench/issue_cases.py replay)")
        return lines


def results_for_deck(work: Any) -> Dict[str, Any]:
    """The per-account results as the deck and the app's commentary tabs read them (claims.to_legacy_result): the accepted drafts and
    the drafts carried with their rejected sentences marked."""
    from .export import merged_results
    return merged_results(work.source, work.first_generation_accounts(), first_generation=True)


def run_book(work: Any, accounts: Optional[List[str]] = None, *, on_account: Optional[Callable[..., Any]] = None,
             on_review: Optional[Callable[..., Any]] = None, overview: bool = True, deck: bool = True, out_path: str = "",
             issues_dir: Optional[Any] = ISSUE_DIR, max_calls: Optional[int] = None) -> BookResult:
    """Write every account in scope (or `accounts`), the overviews, and export the deck. `on_account(key, turn)` is called as each
    account finishes. Returns what was written; a model that failed or a budget that ran out is in the result, not raised."""
    from . import routing as R
    started = time.perf_counter()
    keys = list(accounts or work.entity_accounts())
    cap = max_calls if max_calls is not None else book_cap(work.config, len(keys))
    if cap is not None:
        work.max_model_calls = cap
    result = BookResult(work=work, accounts=keys)
    work.deck["whole_book"] = {"at": time.strftime("%Y-%m-%dT%H:%M:%S"), "accounts": len(keys)}     # the session's drafts are now the deck
    done = work.draft_entity(keys, on_account=on_account, on_review=on_review)
    result.rows = done["rows"]
    result.stopped = next((r.get("error") or "" for r in done["rows"] if r.get("stop_reason") == "budget_session"), "")
    if overview and not result.stopped:
        try:
            written = work.write_overview("escalation", keys)
            result.overview = {s: e["text"] for s, e in written.items() if e.get("text")}
        except R.BudgetExhausted as exc:
            result.overview_error = str(exc)
    if deck:
        result.export = work.export(out_path=out_path, first_generation=True)
    result.ai_results = results_for_deck(work) if work.first_generation_accounts() else {}
    if issues_dir is not None:
        from . import issues
        try:
            result.issue_cases = len(issues.capture(work.source, issues_dir, entity=Path(work.source.workbook_path).stem,
                                                    code=code_fingerprint()["short"]))
        except OSError:
            result.issue_cases = 0
    result.seconds = round(time.perf_counter() - started, 1)
    return result


def redraft(work: Any, account: str, guidance: str = "") -> Dict[str, Any]:
    """Write one account again, with the person's guidance as the request (the app's per-account box). Returns the account's result as
    the commentary tabs read it, or {} when the model did not finish it (the earlier draft stands)."""
    from ..ai import claims as cl
    key = work.source.resolve_account(account) or account
    work.draft_entity([key], instruction=guidance.strip())
    candidate = work.source.candidates.get(key)
    if candidate is None or not candidate.claims:
        return {}
    return cl.to_legacy_result(key, candidate.claims, candidate.outcome, flags=candidate.flags)

"""A working session: what survives from one question to the next.

SourceSession (session.py) is one reading of the sources and the drafts made against it; the controller is one
run on top of it. A person investigating does not work in runs. They ask, follow up on what the answer said, say
"put that in the report", look at what would change, accept it, export, and then someone fixes a number in the
workbook. WorkSession is that thread:

    sources     the binding the session was read at, and a fresh reading when the file moves
    evidence    which accounts were consulted, which refs the model has been shown, the figures Python derived (in
                the order it derived them, so D3 stays D3) and the remarks a search filed
    issues      what the tool results raised, and which were investigated
    turns       every request with its answer, sources, unresolved items, tool path and usage
    candidates  the drafts, whether the user accepted them, and why one no longer stands
    exports     what was written, where, with what delivery decision

Nothing here is a database. A session lives in memory; if it is given a path it is written there as one JSON file
after each turn, and only to a path git would not track. That file holds the commentary and figures of the
workbook, so it is as sensitive as the workbook: it is never created by default and is never part of the repo.

When the sources change (the workbook was edited and saved), the session reads them again and decides, per draft,
whether it still stands: every ref a draft cites is looked up again by what it pointed at (the fact's id, the
derived figure's recipe, the note's text), and if any of those is gone or its value moved the draft is invalid and
says which; if none moved it is carried over, re-verified against the new reading. Answers given before the change
can no longer be turned into a draft: the model is not shown their refs again.
"""
from __future__ import annotations

import copy
import hashlib
import json
import os
import re
import subprocess
import time
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

from ..ai import claims as cl
from . import routing as R
from . import tools as T
from .controller import AgentController, AgentRun, Budget, Issue, Memory
from .session import (Candidate, SourceSession, _fact_label, account_profile, capture_cite_ids, capture_context, claims_digest,
                      context_changes, replay_derived)

SESSION_VERSION = 1
REPO_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_DIR = REPO_ROOT / "fdd_utils" / "output" / "agent_sessions"

# ---------------------------------------------------------------------------
# What a request is for
# ---------------------------------------------------------------------------

_REPORT_NOUN = re.compile(r"(報告|报告|簡報|简报|幻燈片|幻灯片|評論|评论|ppt|pptx|deck|slide|commentary|report)", re.I)
_REPORT_VERB = re.compile(r"(放入|放進|放进|加入|寫入|写入|寫進|写进|納入|纳入|更新|補充|补充|put|add|write|insert|include|move|update|place)", re.I)
_REWRITE = re.compile(r"(改寫|重寫|改写|重写|縮短|缩短|精簡|精简|擴充|扩充|潤色|润色|換成|换成|改成|翻譯|翻译|"
                      r"rewrite|shorten|expand|rephrase|polish|translate|reword|condense|simplify|shorter|longer)", re.I)


def classify_request(text: str) -> str:
    """'report' (put the analysis in the report as it stands), 'report_edit' (put it in, changed), or 'ask'.
    Plain pattern matching on the words, so that "put that in the report" costs no model call; anything
    the patterns do not recognise is an ordinary question."""
    t = str(text or "")
    if _REPORT_NOUN.search(t) and _REPORT_VERB.search(t):
        return "report_edit" if _REWRITE.search(t) else "report"
    return "ask"


def _now() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%S")


# ---------------------------------------------------------------------------
# One turn
# ---------------------------------------------------------------------------

@dataclass
class Turn:
    n: int
    mode: str                       # ask | edit | promote
    request: str
    ok: bool = False
    stop_reason: str = ""
    text: str = ""
    claims: List[Dict[str, Any]] = field(default_factory=list)
    sources: List[Dict[str, Any]] = field(default_factory=list)
    unresolved: List[str] = field(default_factory=list)
    verification: Dict[str, Any] = field(default_factory=dict)
    question: str = ""
    report: Dict[str, Any] = field(default_factory=dict)
    tool_path: List[str] = field(default_factory=list)
    usage: Dict[str, Any] = field(default_factory=dict)
    source_version: str = ""
    at: str = ""
    stale: bool = False
    accounts: List[str] = field(default_factory=list)
    notes: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return dict(self.__dict__)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "Turn":
        known = {k: v for k, v in data.items() if k in cls.__dataclass_fields__}
        return cls(**known)


def _sha(text: str) -> str:
    return hashlib.sha256(str(text or "").encode("utf-8")).hexdigest()[:16]


def _sentences(text: str) -> List[str]:
    return [s.strip() for s in re.split(r"(?<=[。；;！？!?])\s*|\n+", str(text or "")) if s.strip()]


class WorkSession:
    def __init__(self, source: SourceSession, *, model: Optional[Callable[..., Dict[str, Any]]] = None, model_type: str = "",
                 model_name: Optional[str] = None, budget: Optional[Budget] = None, path: Optional[str] = None,
                 max_model_calls: Optional[int] = None, session_id: str = "", policy: Optional[R.Policy] = None,
                 limits: Optional[R.Limits] = None, escalation: Any = None, thinking: str = "", config: Optional[Dict[str, Any]] = None,
                 prepare: Optional[bool] = None) -> None:
        self.source = source
        #: edit turns start as a prepared write (controller.py); None: agent.prepared_write in config.yml, true when unset
        self.prepare = prepare
        source.workspace = self
        self.model, self.model_type, self.model_name = model, model_type, model_name
        self.budget = budget
        self.path = str(path) if path else ""
        self.id = session_id or time.strftime("%Y%m%d-%H%M%S")
        self.created = _now()
        self.turns: List[Turn] = []
        self.memory = Memory()
        self.events: List[Dict[str, Any]] = []
        #: routing: which model answers (a callable for a scripted model, a {"provider","model","thinking"} dict for a configured one)
        self.policy = policy or R.Policy()
        #: which model writes each account (dispatch.py): "pro" starts the account's turn on the escalation model
        self.plan: Dict[str, str] = {}
        self.escalation = escalation
        self.thinking = thinking
        self.config = config
        self._clients: Dict[Tuple, Any] = {}
        #: the books for every model, every turn and every attempt; a cap on it is a cap on the whole piece of work
        limits = limits or R.Limits()
        if max_model_calls is not None:
            limits.max_calls = max_model_calls
        self.shared = R.SharedBudget(limits)
        #: set when the session was resumed against sources that had moved since it was saved
        self.opening_notice: Dict[str, Any] = {}
        #: the deck this work is about: the one last read, the last export and the last update, with what is still unfinished
        self.deck: Dict[str, Any] = {}
        if self.path:
            self._check_path(self.path)

    @property
    def model_calls(self) -> int:
        return self.shared.calls

    @property
    def max_model_calls(self) -> Optional[int]:
        return self.shared.limits.max_calls

    @max_model_calls.setter
    def max_model_calls(self, value: Optional[int]) -> None:
        self.shared.limits.max_calls = value

    @property
    def usage_total(self) -> Dict[str, Any]:
        """The session's usage in the shape the CLI and the panel print; computed from the books, which are the one record."""
        books = self.shared.summary()
        return {"model_calls": books["calls"], "failed_calls": books["failed_calls"], "prompt_tokens_est": books["estimated"]["prompt"],
                "completion_tokens_est": books["estimated"]["completion"], "provider_prompt_tokens": books["provider_reported"]["prompt"],
                "provider_completion_tokens": books["provider_reported"]["completion"], "provider_reported": bool(books["provider_reported"]["calls"]),
                "charged_tokens": books["charged_tokens"], "seconds": books["seconds"]}

    # -- opening -----------------------------------------------------------------------

    @classmethod
    def open(cls, workbook: str = "", *, session_path: Optional[str] = None, resume: bool = True, entity: str = "",
             sheet: Optional[str] = None, financials_from: Optional[str] = None, financials_sheet: Optional[str] = None,
             base_run: Optional[str] = None, mapping_overrides: Optional[Dict[str, str]] = None, **kwargs: Any) -> "WorkSession":
        """Start a session on a workbook, or pick one up from its file. Resuming reads the sources again; a
        session saved against other bytes than the workbook now holds says so in `opening_notice`."""
        saved = None
        if session_path and resume and os.path.exists(session_path):
            with open(session_path, encoding="utf-8") as fh:
                saved = json.load(fh)
            if saved.get("version") != SESSION_VERSION:
                raise ValueError(f"{session_path} is a version {saved.get('version')} session; this build reads version {SESSION_VERSION}")
        params = dict((saved or {}).get("source", {}).get("load_params") or {})
        if workbook:
            params["workbook"] = workbook
        for key, value in (("entity", entity), ("sheet", sheet), ("financials_from", financials_from),
                           ("financials_sheet", financials_sheet), ("base_run", base_run), ("mapping_overrides", mapping_overrides)):
            if value:
                params[key] = value
        if not params.get("workbook"):
            raise ValueError("give a workbook (or a session file that names one)")
        source = SourceSession.load(**{k: params.get(k) for k in ("workbook", "entity", "sheet", "financials_from", "financials_sheet", "base_run",
                                                                  "mapping_overrides")
                                       if k == "workbook" or params.get(k) is not None})
        model = (saved or {}).get("model") or {}
        # "" lets agent.models.primary decide (a hard-coded "local" here used to override the configured first model)
        kwargs.setdefault("model_type", model.get("type") or "")
        kwargs.setdefault("model_name", model.get("name"))
        work = cls(source, path=session_path, **kwargs)
        if saved:
            work._restore(saved)
        return work

    # -- the model -------------------------------------------------------------------------

    def _spec(self, role: str, selection: Dict[str, str]) -> R.RoleSpec:
        """A configured model as a role: the client, what thinking was asked for and whether this provider branch sends it, and the
        price the config carries for it (None when nobody wrote one down)."""
        from . import flag as F
        from .controller import AIClientModel
        provider, model, thinking = selection.get("provider") or "", selection.get("model") or "", selection.get("thinking") or ""
        effort, max_output = selection.get("reasoning_effort") or "", selection.get("max_output")
        key = (role, provider, model, thinking, effort, max_output)
        if key not in self._clients:
            self._clients[key] = AIClientModel(model_type=provider, model_name=model or None, language=self.source.language,
                                               thinking=thinking, ledger_stage=f"agent_{role}", reasoning_effort=effort, max_output=max_output)
        call = self._clients[key]
        if call.provider_fallback and not bool((F._agent_block(self.config).get("allow_provider_fallback"))):
            # AIClient quietly answers from another configured provider when the requested one is not configured. For the agent
            # that would send the workbook's figures, and the bill, somewhere nobody chose: refuse unless config allows it.
            raise ValueError(f"the {role} model asks for provider '{provider}', which is not configured here; the client would fall back to "
                             f"'{call.provider}'. Configure '{provider}' (or set agent.models.{role}), or set agent.allow_provider_fallback: true")
        return R.RoleSpec(role, call, provider=call.provider, model=call.chat_model or model, thinking=thinking or effort,
                          thinking_applied=call.thinking_applied, price=F.price_for(self.config, call.provider, model),
                          price_cached=F.cached_price_for(self.config, call.provider, model),
                          max_output=max_output, thinking_unverified=call.thinking_unverified,
                          provider_requested=provider if call.provider_fallback else "")

    def _router(self) -> R.RoutedModel:
        """The callable the controller calls, over this session's books. Built each turn (the policy and the model may have
        been changed between turns); the books are the session's own and carry on."""
        from . import flag as F
        if self.model is not None:                                       # a scripted or injected model: one role, no price
            primary = R.RoleSpec(R.ROLE_PRIMARY, self.model, provider=self.model_type, model=self.model_name or getattr(self.model, "label", "injected"))
            escalation = (R.RoleSpec(R.ROLE_ESCALATION, self.escalation, provider=self.model_type, model="injected-escalation")
                          if callable(self.escalation) else None)
        else:
            configured = F.role_config(self.config, "primary")
            # an explicit choice (CLI --model, the app's selection) wins; otherwise agent.models.primary; otherwise local
            explicit = bool(self.model_type)
            primary = self._spec(R.ROLE_PRIMARY, {"provider": self.model_type or configured["provider"] or "local",
                                                  "model": self.model_name or (configured["model"] if not explicit or self.model_type == configured["provider"] else ""),
                                                  "thinking": self.thinking or configured["thinking"],
                                                  "reasoning_effort": configured["reasoning_effort"] if not self.thinking else "",
                                                  "max_output": configured["max_output"]})
            escalation = None
            if self.policy.mode == "escalate":
                choice = self.escalation if isinstance(self.escalation, dict) else F.role_config(self.config, "escalation")
                if not choice.get("provider"):
                    raise ValueError("routing is 'escalate' but no escalation model is set (agent.models.escalation, or --escalation-model provider::model)")
                escalation = self._spec(R.ROLE_ESCALATION, choice)
        return R.RoutedModel(primary, escalation, policy=self.policy, budget=self.shared)

    # -- turns -----------------------------------------------------------------------------------

    def ensure_fresh(self) -> Optional[Dict[str, Any]]:
        """If the workbook was saved since it was read, read it again before the next turn (and say what it did)."""
        if self.source.stale_reason():
            return self.refresh()
        return None

    def _run(self, request: str, mode: str, account: str = "", memory: Optional[Memory] = None) -> Turn:
        from . import flag as F
        notice = self.ensure_fresh()
        prepare = self.prepare if self.prepare is not None else F.prepared_write(self.config)
        controller = AgentController(self.source, self._router(), budget=self.budget or _default_budget(),
                                     language=self.source.language, memory=memory or self.memory, prepare=prepare,
                                     seed=F.seed_facts(self.config), review=F.review_meaning(self.config), lead=F.lead_write(self.config),
                                     plan=F.plan_first(self.config),
                                     tool_loop_on_repeat=F.tool_loop_on_repeat(self.config))
        before = {a: c.claims_digest for a, c in self.source.candidates.items()}
        run = controller.run(request, mode=mode, account=account, export=False, turn_number=len(self.turns) + 1,
                             start_strong=bool(account) and self.plan.get(self.source.resolve_account(account) or account) == "pro")
        turn = self._record(run, mode, request, before)
        if notice:
            turn.notes.append(f"the workbook had changed: {self._effects_line(notice.get('candidate_effects'))}")
        self._autosave()
        return turn

    def ask(self, question: str) -> Turn:
        return self._run(question, "ask")

    def entity_accounts(self) -> List[str]:
        """The accounts a report on this entity comments on: those the selection admitted (the legacy pipeline's own rule), in the
        workbook's order. Decided by Python, not by a model."""
        return [k for k in self.source.contracts if not str(k).startswith("__") and self.source.in_scope(k)]

    def dispatch(self, accounts: Optional[List[str]] = None) -> Dict[str, Any]:
        """The stronger model names the accounts it should write (dispatch.py); the rest start on the primary. Needs routing
        'escalate'. One call; the plan is kept on the session and used by every later turn on those accounts."""
        from .dispatch import dispatch as _dispatch
        if self.policy.mode != "escalate":
            raise ValueError("dispatch needs routing 'escalate' (a primary and an escalation model)")
        keys = [k for k in (accounts or self.entity_accounts()) if k in self.source.contracts]
        record = _dispatch(self.source, self._router(), keys)
        self.plan = dict(record["plan"])
        if "dispatched" not in (self.policy.escalate_tasks or []):
            self.policy.escalate_tasks = list(self.policy.escalate_tasks or []) + ["dispatched"]
        self.events.append({"at": _now(), "type": "dispatch", "by": record["by"], "pro": record["pro"], "error": record["error"]})
        return record

    def write_overview(self, role: str = "escalation", accounts: Optional[List[str]] = None, chars: int = 140) -> Dict[str, Any]:
        """The executive summary of each statement from the finished paragraphs, each sentence verified against its account
        (overview.py). Kept on the source session, where a first-generation export puts it in the summary band."""
        from .overview import write_overview as _write
        router = self._router()
        if role not in router.roles:
            role = R.ROLE_PRIMARY
        keys = [k for k in (accounts or self.entity_accounts()) if k in self.source.contracts]
        ctl = AgentController(self.source, router, language=self.source.language)
        labels = {k: ctl._period_labels(k, {"periods": self.source.contracts[k].periods,
                                            "annualization": self.source.contracts[k].annualization}) for k in keys}
        result = _write(self.source, router, role=role, language=self.source.language, accounts=keys, period_labels=labels, chars=chars)
        self.source.overview = {s: e["text"] for s, e in result.items() if e.get("text")}
        self.events.append({"at": _now(), "type": "overview", "role": role,
                            "sections": {s: {"dropped": e["dropped"], "calls": e["calls"], "error": e["error"]} for s, e in result.items()}})
        return result

    def draft_entity(self, accounts: Optional[List[str]] = None, instruction: str = "", on_account: Optional[Callable[..., Any]] = None) -> Dict[str, Any]:
        """Every in-scope account written from its own evidence, one edit turn each, accepted when the verifier passes it.

        The state decides the order and the scope; each account is an independent task, so each turn starts with a fresh memory
        (the earlier accounts' answers are not resent in every prompt). Nothing is copied from a base run: an account the model
        could not finish has no paragraph, and the export says so. `on_account(key, turn)` is called after each account."""
        keys = [k for k in (accounts or self.entity_accounts()) if k in self.source.contracts]
        rows = []
        for key in keys:
            text = instruction or f"根据工作簿中的证据，撰写{key}的报告段落"
            try:
                turn = self._run(text, "edit", key, memory=Memory())
            except R.BudgetExhausted as exc:
                rows.append({"account": key, "ok": False, "stop_reason": "budget_session", "error": str(exc)})
                break
            if turn.ok:
                self.accept(key)
            rows.append({"account": key, "ok": turn.ok, "stop_reason": turn.stop_reason, "turn": turn.n,
                         "stop_class": ((turn.report or {}).get("routing") or {}).get("stop_class"),
                         "unresolved": list(turn.unresolved), "prepared_fallback": (turn.report or {}).get("prepared_fallback")})
            if on_account is not None:
                on_account(key, turn)
        self._autosave()
        return {"accounts": keys, "rows": rows, "ok": [r["account"] for r in rows if r["ok"]],
                "not_finished": [r["account"] for r in rows if not r["ok"]] + [k for k in keys[len(rows):]]}

    def edit(self, instruction: str, account: str = "", *, accept: bool = False, export: bool = False, out_path: str = "") -> Turn:
        """The model rewrites one account's paragraph from the evidence already gathered. `account` defaults to the
        account of the last answer; with several, the user is asked."""
        key = self.source.resolve_account(account) if account else self.infer_account()
        if not key:
            turn = Turn(n=len(self.turns) + 1, mode="edit", request=instruction, ok=False, stop_reason="needs_user", at=_now(),
                        question=("Which account is this for? " + ", ".join(self.answer_accounts())) if self.answer_accounts()
                        else "Which account is this for? There is no earlier answer to take it from.",
                        source_version=self.source.source_version)
            self.turns.append(turn)
            return turn
        turn = self._run(instruction, "edit", key)
        if turn.ok and (accept or export):
            self.accept(key)
        if turn.ok and export:
            turn.report["export"] = self.export(out_path=out_path)
        self._autosave()
        return turn

    def report(self, request: str = "", account: str = "", *, export: bool = False, out_path: str = "") -> Turn:
        """"Put that in the report." A plain placement turns the last answer's verified claims into a draft with no
        model call; a request to change the wording goes to the model as an edit, with the evidence it already has."""
        kind = classify_request(request) if request else "report"
        if kind == "report_edit":
            return self.edit(request, account, accept=True, export=export, out_path=out_path)
        turn = self.promote(account)
        if turn.ok:
            for key in turn.accounts:
                self.accept(key)
            if export:
                turn.report["export"] = self.export(out_path=out_path)
        self._autosave()
        return turn

    def promote(self, account: str = "", turn_number: Optional[int] = None) -> Turn:
        """Make a draft of an earlier answer's claims, unchanged. Deterministic; no model call."""
        notice = self.ensure_fresh()
        source_turn = self._answer_turn(turn_number)
        turn = Turn(n=len(self.turns) + 1, mode="promote", request=f"promote turn {source_turn.n}" if source_turn else "promote", at=_now(),
                    source_version=self.source.source_version, usage={"model_calls": 0})
        self.turns.append(turn)
        if source_turn is None:
            turn.stop_reason, turn.question = "no_answer", "There is no earlier answer to put in the report. Ask a question first."
            return turn
        if source_turn.stale:
            turn.stop_reason = "answer_is_stale"
            turn.question = ("That answer was written before the workbook changed, and its figures may no longer exist. "
                             "Ask again, or edit the account from the new sources.")
            return turn
        want = self.source.resolve_account(account) if account else None
        groups: Dict[str, List[Dict[str, Any]]] = {}
        for claim in source_turn.claims:
            key = self.source.resolve_account(str(claim.get("account") or ""))
            if key and (not want or key == want):
                groups.setdefault(key, []).append(claim)
        if not groups:
            turn.stop_reason, turn.question = "no_claims", "The answer holds no claim for that account."
            return turn
        results = []
        for key, claims in groups.items():
            result = T.call_tool(self.source, "validate_candidate", {"account": key, "claims": claims})
            candidate = self.source.candidates.get(key)
            if candidate is not None:
                candidate.origin, candidate.turn = "promoted", turn.n
            results.append({"account": key, "ok": result.ok, "error": (result.error or {}).get("code"),
                            "message": (result.error or {}).get("message")})
        turn.accounts = [r["account"] for r in results if r["ok"]]
        turn.ok = bool(turn.accounts) and all(r["ok"] for r in results)
        turn.stop_reason = "done" if turn.ok else "not_validated"
        turn.text = "\n".join(cl.render_claims(self.source.candidates[a].claims) for a in turn.accounts)
        turn.claims = [c for a in turn.accounts for c in self.source.candidates[a].claims]
        turn.unresolved = list(source_turn.unresolved)
        turn.report["promoted"] = results
        self.memory.history.append({"n": turn.n, "mode": "promote", "request": turn.request, "text": turn.text, "claims": turn.claims,
                                    "unresolved": turn.unresolved, "stale": False})
        self._autosave()
        return turn

    def _record(self, run: AgentRun, mode: str, request: str, digests_before: Dict[str, str]) -> Turn:
        answer = run.answer or {}
        turn = Turn(n=len(self.turns) + 1, mode=mode, request=request, ok=run.ok, stop_reason=run.stop_reason,
                    text=str(answer.get("text") or ""), claims=list(answer.get("claims") or []), sources=list(answer.get("sources") or []),
                    unresolved=list(answer.get("unresolved") or []), verification=dict(answer.get("verification") or {}),
                    question=run.question, report=dict(run.report), tool_path=run.tool_path, usage=dict(run.usage),
                    source_version=run.source_version, at=_now())
        turn.accounts = sorted({str(c.get("account")) for c in turn.claims if c.get("account")})
        self.turns.append(turn)
        for key, candidate in self.source.candidates.items():
            if digests_before.get(key) != candidate.claims_digest:
                candidate.turn, candidate.origin = turn.n, "edit"
        #: every call this turn made, from the books: which role and model, what the provider reported and what was estimated
        turn.usage["call_records"] = [r.to_dict() for r in self.shared.records if r.turn == turn.n]
        self.memory.history.append({"n": turn.n, "mode": mode, "request": request, "text": turn.text or turn.question,
                                    "claims": turn.claims, "unresolved": turn.unresolved, "stale": False})
        return turn

    # -- reading the thread ---------------------------------------------------------------------------

    def _answer_turn(self, number: Optional[int] = None) -> Optional[Turn]:
        for turn in reversed(self.turns):
            if number is not None and turn.n != number:
                continue
            if turn.mode in ("ask", "edit") and turn.ok and turn.claims:
                return turn
        return None

    def answer_accounts(self) -> List[str]:
        turn = self._answer_turn()
        return sorted({str(c.get("account")) for c in (turn.claims if turn else []) if c.get("account")})

    def infer_account(self) -> Optional[str]:
        accounts = self.answer_accounts()
        return accounts[0] if len(accounts) == 1 else None

    # -- drafts ----------------------------------------------------------------------------------------

    def accept(self, account: str) -> bool:
        """The user says this draft goes in the report. If the account moved around it since it was last looked at
        (`context_changes`), accepting is the acknowledgement: the reasons are kept in the event, not discarded."""
        key = self.source.resolve_account(account)
        candidate = self.source.candidates.get(key or "")
        if candidate is None or not candidate.validated or candidate.source_version != self.source.source_version:
            return False
        acknowledged, candidate.context_changes = list(candidate.context_changes), []
        candidate.accepted = True
        self.events.append({"at": _now(), "type": "accepted", "account": key, "digest": candidate.claims_digest, "acknowledged": acknowledged})
        self._autosave()
        return True

    def unaccept(self, account: str) -> bool:
        key = self.source.resolve_account(account)
        candidate = self.source.candidates.get(key or "")
        if candidate is None:
            return False
        if candidate.accepted:
            self.events.append({"at": _now(), "type": "acceptance_withdrawn", "account": key, "by": "user", "digest": candidate.claims_digest})
        candidate.accepted = False
        self._autosave()
        return True

    def base_text(self, account: str) -> str:
        from ..financial_common import get_pipeline_result_text
        result = (self.source.base_results or {}).get(account)
        return str(get_pipeline_result_text(result) or "") if result else ""

    def candidates_view(self) -> List[Dict[str, Any]]:
        """Each draft as a reader needs it: what it says, what it replaces, what changed, how it was checked,
        whether the user accepted it, and what is still open."""
        out = []
        blocking = {i.account: i for i in self.memory.issues.values() if i.severity == "blocking" and not i.resolved and i.account}
        for key, candidate in self.source.candidates.items():
            text = cl.render_claims(candidate.claims)
            before = self.base_text(key)
            old_sentences, new_sentences = _sentences(before), _sentences(text)
            verdicts = candidate.outcome.get("verdicts") or []
            counts = Counter(v.get("verdict") for v in verdicts)
            stale = candidate.source_version != self.source.source_version
            status = ("invalid" if candidate.invalid_reason or stale else
                      ("accepted" if candidate.accepted else ("validated" if candidate.validated else "failed")))
            turn = next((t for t in self.turns if t.n == candidate.turn), None)
            unresolved = list(turn.unresolved) if turn else []
            if key in blocking:
                unresolved.append(blocking[key].message)
            out.append({
                "account": key, "status": status, "origin": candidate.origin, "carried": candidate.carried, "accepted": candidate.accepted,
                "text": text, "replaces": before or None,
                "changes": {"added": [s for s in new_sentences if s not in old_sentences],
                            "removed": [s for s in old_sentences if s not in new_sentences],
                            "kept": sum(1 for s in new_sentences if s in old_sentences),
                            "replaces_existing": bool(before)},
                "verification": {"verified": counts.get("verified", 0), "inference": counts.get("verified_inference", 0),
                                 "failed": counts.get("failed", 0), "unverified": counts.get("unverified", 0), "claims": len(verdicts)},
                "invalid_reason": candidate.invalid_reason or ("written against earlier sources" if stale else ""),
                "context_changes": list(candidate.context_changes), "needs_review": bool(candidate.context_changes),
                "unresolved": unresolved, "source_version": candidate.source_version, "turn": candidate.turn,
                "sources": [{"ref": r, **{k: v for k, v in m.items() if k in ("id", "kind", "period", "line", "value")}}
                            for r, m in candidate.cite_ids.items()],
            })
        return out

    def unresolved(self) -> List[Dict[str, str]]:
        items: List[Dict[str, str]] = []
        last = self._answer_turn()
        for text in (last.unresolved if last else []):
            items.append({"from": f"turn {last.n}", "account": "", "message": text})
        for issue in self.memory.issues.values():
            if not issue.resolved:
                items.append({"from": issue.severity, "account": issue.account, "message": issue.message})
        return items

    # -- export and decks --------------------------------------------------------------------------------

    def first_generation_plan(self) -> Dict[str, Any]:
        """How to get a full first deck when there is no base run: the legacy pipeline writes every account (and a run this session
        can then be bound to). The agent path does not write a whole entity from nothing yet, and a deck of a few accounts is not one."""
        p = self.source.load_params or {}
        workbook = p.get("workbook") or self.source.workbook_path
        legacy = ["python", "inspect_databook.py", workbook, "--run-ai", "--export-pptx"]
        for flag, key in (("--entity", "entity"), ("--sheet", "sheet"), ("--financials-from", "financials_from"), ("--financials-sheet", "financials_sheet")):
            if p.get(key):
                legacy += [flag, str(p[key])]
        return {"why": "no base run: the agent path writes and checks single accounts; a whole entity's first draft comes from the legacy pipeline",
                "steps": [{"what": "write the whole entity with the legacy pipeline (real model calls; the run's tokens are in usage.jsonl)",
                           "command": " ".join(f'"{x}"' if " " in str(x) else str(x) for x in legacy)},
                          {"what": "bind this session to that run (the run must be shown to describe this workbook: identity or every account's fingerprint)",
                           "command": "in the session: /base-run <RUN_ID>   (or open it again with --base-run <RUN_ID>)"}],
                "not_offered": "a deck holding only the accounts edited here, presented as the report: it would leave every other account out"}

    def set_base_run(self, run: str) -> Dict[str, Any]:
        """Bind the session to a base run (a legacy run of the whole entity) and read the sources again. The run is checked against the
        workbook (check_base_run); the result is returned and recorded, and an export refuses a run that does not describe it."""
        self.source.load_params = dict(self.source.load_params or {}, base_run=run)
        result = self._reload("base_run", {"run": run})
        result["base_run_check"] = self.source.base_run_check
        return result

    @property
    def mapping_overrides(self) -> Dict[str, str]:
        return dict((self.source.load_params or {}).get("mapping_overrides") or {})

    def set_mapping_override(self, account: str, sheet: str) -> Dict[str, Any]:
        """The USER says which tab an account is read from. Only this method (the CLI's /override, the panel) changes a mapping; no tool
        does, so the model cannot. The sources are read again with it, the source version changes, and every draft is rebound."""
        overrides = self.mapping_overrides
        overrides[str(account)] = str(sheet)
        self.source.load_params = dict(self.source.load_params or {}, mapping_overrides=overrides)
        return self._reload("mapping_override", {"account": account, "sheet": sheet, "by": "user"})

    def clear_mapping_override(self, account: str) -> Dict[str, Any]:
        overrides = self.mapping_overrides
        removed = overrides.pop(str(account), None)
        self.source.load_params = dict(self.source.load_params or {}, mapping_overrides=overrides)
        return self._reload("mapping_override_cleared", {"account": account, "was": removed, "by": "user"})

    def _reload(self, why: str, detail: Dict[str, Any]) -> Dict[str, Any]:
        result = self.refresh()
        self.events.append({"at": _now(), "type": why, **detail, "ok": bool(result.get("ok")), "source_version": self.source.source_version})
        self._autosave()
        return result

    def export(self, out_path: str = "", accounts: Optional[List[str]] = None, first_generation: bool = False) -> Dict[str, Any]:
        """A deck from the accepted drafts. Over a base run, the drafts replace those accounts in its commentary. With
        first_generation=True the deck is the drafts alone (draft_entity): no base text fills an account the agent did not finish;
        the delivery decision lists each missing account and blocks the deck."""
        self.ensure_fresh()
        if first_generation:
            chosen = accounts or [a for a, c in self.source.candidates.items() if c.accepted and c.validated
                                  and c.source_version == self.source.source_version]
            if not accounts:
                # A draft the gate did not accept goes in too, its rejected sentences marked (red where the databook contradicts a
                # figure, orange otherwise) and the account named in the delivery decision, which blocks. Left out, its slot was
                # empty: on the client file (2026-11-09) 5 to 11 accounts per arm, most with one to three rejected sentences among
                # five to ten verified ones, and the person had nothing to correct.
                chosen += [a for a in T.drafts_to_mark(self.source) if a not in chosen]
            if not chosen:
                return {"ok": False, "error": {"code": "NO_ACCEPTED_CANDIDATE", "message": "no accepted draft to export"}}
            result = T.call_tool(self.source, "export_first_generation", {"accounts": chosen, "out_path": out_path})
            self.events.append({"at": _now(), "type": "export", "first_generation": True, "accounts": chosen, "ok": result.ok,
                                "error": (result.error or {}).get("code"), "path": (result.data or {}).get("path")})
            data = result.data or {}
            self.deck["last_export"] = {"at": _now(), "ok": result.ok, "path": data.get("path"), "accounts": chosen, "first_generation": True,
                                        "delivery": (data.get("delivery") or {}).get("status"), "error": (result.error or {}).get("code"),
                                        "source_version": self.source.source_version}
            if data.get("path"):
                self.deck["current"] = data.get("path")
            self._autosave()
            return result.to_dict()
        if not self.source.base_results:
            return {"ok": False, "error": {"code": "NO_BASE_COMMENTARY",
                    "message": "this session has no base run: a deck made now would hold only the edited accounts. "
                               "Write the whole entity first, then bind this session to that run (first_generation below).",
                    "first_generation": self.first_generation_plan()}}
        check = self.source.base_run_check or {}
        if check.get("status") in ("mismatch",):
            return {"ok": False, "error": {"code": "BASE_RUN_MISMATCH",
                    "message": "the base run does not describe this workbook: " + "; ".join(check.get("reasons") or [])[:400],
                    "first_generation": self.first_generation_plan()}}
        chosen = accounts or [a for a, c in self.source.candidates.items() if c.accepted and c.validated
                              and c.source_version == self.source.source_version]
        if not chosen:
            return {"ok": False, "error": {"code": "NO_ACCEPTED_CANDIDATE",
                    "message": "no accepted draft to export: accept one first (accept <account>)"}}
        result = T.call_tool(self.source, "export_candidate", {"accounts": chosen, "out_path": out_path})
        self.events.append({"at": _now(), "type": "export", "accounts": chosen, "ok": result.ok, "error": (result.error or {}).get("code"),
                            "path": (result.data or {}).get("path")})
        data = result.data or {}
        self.deck["last_export"] = {"at": _now(), "ok": result.ok, "path": data.get("path"), "accounts": chosen,
                                    "touched_slides": data.get("touched_slide_numbers"), "delivery": (data.get("delivery") or {}).get("status"),
                                    "error": (result.error or {}).get("code"), "source_version": self.source.source_version}
        if result.ok and data.get("path"):
            self.deck["current"] = data.get("path")
        self._autosave()
        return result.to_dict()

    def read_deck(self, path: str) -> Dict[str, Any]:
        self.ensure_fresh()
        result = T.call_tool(self.source, "read_deck", {"path": path})
        data = result.data or {}
        actions = data.get("actions") or []
        self.deck.update(current=path, read={
            "at": _now(), "ok": result.ok, "mode": data.get("mode"), "source_moved": data.get("source_moved"), "summary": data.get("summary"),
            "conflicts": [{k: a.get(k) for k in ("element", "action", "reason")} for a in actions
                          if a.get("action") in ("conflict", "inspect_and_relink", "table_edited")][:40],
            "source_version": self.source.source_version, "error": (result.error or {}).get("code")})
        self._autosave()
        return result.to_dict()

    def update_deck(self, path: str, out_path: str = "") -> Dict[str, Any]:
        """Update a returned deck from the drafts the user accepted (and only those)."""
        self.ensure_fresh()
        everything = self.source.candidates
        self.source.candidates = {a: c for a, c in everything.items() if c.accepted}
        try:
            result = T.call_tool(self.source, "update_deck", {"path": path, "out_path": out_path})
        finally:
            self.source.candidates = everything
        self.events.append({"at": _now(), "type": "update_deck", "ok": result.ok, "error": (result.error or {}).get("code")})
        data = result.data or {}
        self.deck["last_update"] = {"at": _now(), "ok": result.ok, "deck": path, "out_path": data.get("out_path"),
                                    "complete": bool(data.get("complete")), "applied": len(data.get("applied") or []),
                                    "unfinished": (data.get("unfinished") or [])[:40], "next_steps": data.get("next_steps") or [],
                                    "error": (result.error or {}).get("code"), "source_version": self.source.source_version}
        if result.ok and data.get("out_path"):
            self.deck["current"] = data.get("out_path")
        self._autosave()
        return result.to_dict()

    # -- when the sources move --------------------------------------------------------------------------------

    def refresh(self) -> Dict[str, Any]:
        result = T.call_tool(self.source, "refresh_sources")
        if not result.ok:
            return {"ok": False, "error": result.error}
        return {"ok": True, **result.data}

    def after_refresh(self, before: Dict[str, Any], diff: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Called by SourceSession.refresh via the refresh tool. Decides, per draft, whether it stands."""
        self.memory.seen_refs.clear()
        self.memory.consulted.clear()
        self.memory.issues.clear()
        for h in self.memory.history:
            h["stale"] = True
        for turn in self.turns:
            if turn.source_version != self.source.source_version:
                turn.stale = True
        effects = []
        for key, candidate in list(self.source.candidates.items()):
            effects.append(self._rebind(candidate))
        self.events.append({"at": _now(), "type": "sources_changed", "source_version": self.source.source_version,
                            "changed_accounts": list((diff or {}).get("changed_accounts") or []), "effects": effects})
        return effects

    def _rebind(self, candidate: Candidate) -> Dict[str, Any]:
        """Find what a draft rests on in the new reading, or say what is gone or moved. Three layers:

        1. what it cites: each figure by its id (value compared), each derived figure by its recipe AND the figures it was
           computed from, each note by its text AND the period and kind it belongs to
        2. whether it still verifies against the new reading (every sentence, again)
        3. the rest of what it rests on (session.capture_context): the other lines of the periods it is about, the
           checks and permissions, the remarks, the latest period. If any of that moved, the draft is kept but is no
           longer the one the user accepted: acceptance is withdrawn and the changes are named.

        Layers 1 and 2 invalidate. Layer 3 asks again."""
        key = candidate.account
        effect = {"account": key, "was_accepted": candidate.accepted}
        contract = self.source.contracts.get(key)
        if contract is None:
            candidate.invalid_reason = "the account is no longer in the workbook"
            candidate.accepted = False
            return {**effect, "status": "invalidated", "reasons": [candidate.invalid_reason]}
        replay_derived(contract, [m["recipe"] for m in candidate.cite_ids.values() if m.get("kind") == "derived" and m.get("recipe")])
        mapping: Dict[str, str] = {}
        reasons: List[str] = []
        for ref, meta in candidate.cite_ids.items():
            if meta.get("kind") == "note":
                note = next((n for n in contract.notes if _sha(n.get("text")) == meta.get("text_sha")), None)
                if note is None and meta.get("origin") == "search" and meta.get("text"):
                    note = self._refile_note(contract, meta)
                if note is None:
                    reasons.append(f"a note it cites is gone ({meta.get('id')})")
                else:
                    mapping[ref] = note["ref"]
                    if ("period" in meta and (meta.get("period") or None) != (note.get("period") or None)) or \
                            ("note_kind" in meta and (meta.get("note_kind") or None) != (note.get("kind") or None)):
                        reasons.append(f"a note it cites now belongs to {note.get('period') or 'no period'} ({note.get('kind')}), "
                                       f"not {meta.get('period') or 'no period'} ({meta.get('note_kind')})")
                continue
            item = contract.fact(meta.get("id") or "")
            label = " ".join(x for x in (key, str(meta.get("period") or ""), str(meta.get("line") or "")) if x)
            if item is None:
                reasons.append(f"{label}: the figure it cites is gone")
                continue
            old, new = meta.get("value"), item.get("value")
            if old is not None and new is not None and abs(float(new) - float(old)) > 0.5:
                reasons.append(f"{label}: {float(old):,.0f} is now {float(new):,.0f}")
            for input_id, old_input in (meta.get("inputs") or {}).items():
                source_fact = contract.fact(input_id)
                if source_fact is None:
                    reasons.append(f"{label}: a figure it was computed from is gone ({_fact_label(contract, input_id)})")
                elif old_input is not None and source_fact.get("value") is not None and abs(float(source_fact["value"]) - float(old_input)) > 0.5:
                    reasons.append(f"{label}: it was computed from {_fact_label(contract, input_id)}, which moved from "
                                   f"{float(old_input):,.0f} to {float(source_fact['value']):,.0f}")
            mapping[ref] = item["ref"]
        if reasons:
            candidate.invalid_reason = "; ".join(reasons)[:700]
            candidate.accepted = False
            return {**effect, "status": "invalidated", "reasons": reasons}
        claims = [dict(c, cites=[mapping.get(r, r) for r in c.get("cites") or []],
                       cause_cites=[mapping.get(r, r) for r in c.get("cause_cites") or []], source_version=self.source.source_version)
                  for c in candidate.claims]
        outcome = cl.verify_candidate(claims, self.source.contracts, default_account=key)
        if not outcome["ok"]:
            codes = sorted(outcome.get("by_code") or {})
            candidate.invalid_reason = "checked again against the new sources: " + ", ".join(codes)
            candidate.accepted = False
            return {**effect, "status": "invalidated", "reasons": [candidate.invalid_reason]}
        profile = account_profile(self.source, key)
        changes = context_changes(contract, candidate.context, profile_now=profile)
        candidate.claims, candidate.outcome = claims, outcome
        candidate.source_version, candidate.claims_digest = self.source.source_version, claims_digest(claims)
        candidate.cite_ids, candidate.carried, candidate.invalid_reason = capture_cite_ids(contract, claims), True, ""
        candidate.context, candidate.context_changes = capture_context(contract, claims, candidate.cite_ids, profile), changes
        if changes:
            if candidate.accepted:
                self.events.append({"at": _now(), "type": "acceptance_withdrawn", "account": key, "by": "source_change", "reasons": changes})
            candidate.accepted = False
            return {**effect, "status": "carried_changed", "reasons": changes}
        return {**effect, "status": "carried", "reasons": []}

    def _refile_note(self, contract: Any, meta: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """A remark a search filed, looked for again: the same text must still be in a visible cell."""
        from . import investigate as I
        scan = I._scan_text(self.source)
        want = re.sub(r"\s+", " ", meta.get("text") or "")
        cell = next((c for c in scan["cells"] if c["text"][:400] == want[:400]), None)
        if cell is None:
            return None
        number = len(contract.notes) + 1
        note = {"id": f"{contract.account}|note|{number}", "kind": "searched_remark", "text": cell["text"][:400], "period": None,
                "sheet": cell["sheet"], "address": cell["address"], "origin": "search", "query": meta.get("query", ""), "ref": f"N{number}"}
        contract.notes.append(note)
        return note

    @staticmethod
    def _effects_line(effects: Optional[List[Dict[str, Any]]]) -> str:
        effects = effects or []
        if not effects:
            return "no draft was affected"
        count = Counter(e["status"] for e in effects)
        parts = [f"{count['carried']} draft(s) stand unchanged"] if count["carried"] else []
        if count["carried_changed"]:
            parts.append(f"{count['carried_changed']} still verify but the account moved around them (accept again)")
        if count["invalidated"]:
            parts.append(f"{count['invalidated']} no longer stand")
        return ", ".join(parts)

    # -- status ------------------------------------------------------------------------------------------------

    def status(self) -> Dict[str, Any]:
        stale = self.source.stale_reason()
        views = self.candidates_view()
        return {"id": self.id, "source_version": self.source.source_version, "workbook": os.path.basename(self.source.workbook_path),
                "entity": self.source.entity_name or None, "base_run": os.path.basename(self.source.base_run_folder or "") or None,
                "stale": stale, "turns": len(self.turns), "model_calls": self.model_calls, "max_model_calls": self.max_model_calls,
                "usage": dict(self.usage_total), "budget": self.shared.summary(), "routing": self.policy.mode,
                "candidates": {v["account"]: v["status"] for v in views},
                "accounts_consulted": list(self.memory.consulted), "open_issues": len([i for i in self.memory.issues.values() if not i.resolved]),
                "mapping_overrides": self.mapping_overrides, "base_run_check": (self.source.base_run_check or {}).get("status"),
                "deck": {k: self.deck.get(k) for k in ("current",) if self.deck.get(k)} | (
                    {"last_update_complete": self.deck["last_update"].get("complete")} if self.deck.get("last_update") else {}),
                "session_file": self.path or None}

    # -- persistence --------------------------------------------------------------------------------------------

    @staticmethod
    def _check_path(path: str) -> None:
        """Refuse a session file git would track: it holds the workbook's commentary and figures."""
        target = Path(path).expanduser().resolve()
        try:
            target.relative_to(REPO_ROOT)
        except ValueError:
            return                                              # outside the repo: nothing to commit by accident
        under_output = str(target).startswith(str(REPO_ROOT / "fdd_utils" / "output"))
        try:
            # git's answer counts only when it is an answer: 0 = ignored, 1 = not ignored. Any other code (128 in a folder that is not a git
            # clone, e.g. a copied or downloaded tree) means git could not say, and that must not turn the output folder into a refusal.
            code = subprocess.run(["git", "check-ignore", "-q", str(target)], cwd=str(REPO_ROOT), capture_output=True, timeout=10).returncode
            ignored = (code == 0) if code in (0, 1) else under_output
        except Exception:                                       # noqa: BLE001
            ignored = under_output
        if not ignored:
            raise ValueError(f"{path} is inside the repo and not git-ignored. A session file holds figures and commentary from the workbook: "
                             f"use a path under fdd_utils/output/ (ignored) or outside the repo.")

    def to_dict(self) -> Dict[str, Any]:
        derived = {a: [m for m in (self._recipe(d) for d in c.derived)] for a, c in self.source.contracts.items() if c.derived}
        searched = {a: [{k: n.get(k) for k in ("id", "ref", "sheet", "address", "text", "query")} for n in c.notes if n.get("origin") == "search"]
                    for a, c in self.source.contracts.items() if any(n.get("origin") == "search" for n in c.notes)}
        return {
            "version": SESSION_VERSION, "id": self.id, "created": self.created, "saved": _now(),
            "model": {"type": self.model_type, "name": self.model_name},
            "source": {"binding": self.source.binding.to_dict(), "load_params": self.source.load_params,
                       "workbook": os.path.basename(self.source.workbook_path), "source_version": self.source.source_version},
            "turns": [t.to_dict() for t in self.turns],
            "memory": {"seen_refs": sorted(self.memory.seen_refs), "consulted": list(self.memory.consulted),
                       "issues": {k: v.to_dict() for k, v in self.memory.issues.items()}, "history": self.memory.history},
            "derived": derived, "searched": searched,
            "candidates": {a: {"claims": c.claims, "outcome": c.outcome, "source_version": c.source_version, "claims_digest": c.claims_digest,
                               "origin": c.origin, "accepted": c.accepted, "invalid_reason": c.invalid_reason, "cite_ids": c.cite_ids,
                               "turn": c.turn, "carried": c.carried, "context": c.context, "context_changes": c.context_changes,
                               "review": c.review, "flags": c.flags}
                           for a, c in self.source.candidates.items()},
            "exports": [{k: e.get(k) for k in ("path", "ok", "source_version", "touched_slide_numbers", "candidates")} | {"delivery": (e.get("delivery") or {}).get("status")}
                        for e in self.source.exports],
            "events": self.events[-200:], "usage": self.usage_total, "model_calls": self.model_calls, "max_model_calls": self.max_model_calls,
            "budget": self.shared.to_dict(), "routing": self.policy.mode, "deck": self.deck,
            "plan": self.plan, "overview": self.source.overview,
        }

    @staticmethod
    def _recipe(item: Dict[str, Any]) -> Dict[str, Any]:
        from .session import derived_recipe
        return derived_recipe(item)

    def save(self, path: Optional[str] = None) -> str:
        target = str(path or self.path or "")
        if not target:
            target = str(DEFAULT_DIR / f"{self.id}.agent-session.json")
        self._check_path(target)
        self.path = target
        Path(target).parent.mkdir(parents=True, exist_ok=True)
        data = json.dumps(self.to_dict(), ensure_ascii=False, indent=1, default=str)
        fd = os.open(target, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            fh.write(data)
        return target

    def _autosave(self) -> None:
        if self.path:
            self.save()

    def _restore(self, data: Dict[str, Any]) -> None:
        self.id, self.created = data.get("id") or self.id, data.get("created") or self.created
        self.turns = [Turn.from_dict(t) for t in data.get("turns") or []]
        self.events = list(data.get("events") or [])
        self.deck = dict(data.get("deck") or {})
        if data.get("budget"):
            self.shared.restore(data["budget"])
        if self.max_model_calls is None:
            self.max_model_calls = data.get("max_model_calls")
        same = (data.get("source") or {}).get("source_version") == self.source.source_version
        self.plan = dict(data.get("plan") or {})
        if self.plan and "dispatched" not in (self.policy.escalate_tasks or []):
            self.policy.escalate_tasks = list(self.policy.escalate_tasks or []) + ["dispatched"]
        if same:
            self.source.overview = dict(data.get("overview") or {})
        mem = data.get("memory") or {}
        self.memory.history = list(mem.get("history") or [])
        for a, recipes in (data.get("derived") or {}).items():
            contract = self.source.contracts.get(a)
            if contract is not None:
                replay_derived(contract, recipes)
        if same:
            for a, notes in (data.get("searched") or {}).items():
                contract = self.source.contracts.get(a)
                for n in notes:
                    if contract is not None and not any(x.get("ref") == n.get("ref") for x in contract.notes):
                        self._refile_note(contract, {"text": n.get("text"), "query": n.get("query")})
            self.memory.seen_refs |= set(mem.get("seen_refs") or [])
            self.memory.consulted.extend(mem.get("consulted") or [])
            self.memory.issues.update({k: Issue(**v) for k, v in (mem.get("issues") or {}).items()})
        for a, c in (data.get("candidates") or {}).items():
            self.source.candidates[a] = Candidate(
                a, c["claims"], c["outcome"], c["source_version"], c["claims_digest"], origin=c.get("origin", "edit"),
                accepted=bool(c.get("accepted")), invalid_reason=c.get("invalid_reason", ""), cite_ids=c.get("cite_ids") or {},
                turn=int(c.get("turn") or 0), carried=bool(c.get("carried")), context=c.get("context") or {},
                context_changes=list(c.get("context_changes") or []), review=c.get("review") or {}, flags=list(c.get("flags") or []))
        for e in data.get("exports") or []:
            self.source.exports.append({**e, "delivery": {"status": e.get("delivery")}, "restored": True})
        if not same:
            for h in self.memory.history:
                h["stale"] = True
            for t in self.turns:
                t.stale = True
            effects = [self._rebind(c) for c in list(self.source.candidates.values())]
            self.opening_notice = {"source_changed": True, "candidate_effects": effects, "summary": self._effects_line(effects)}
            self.events.append({"at": _now(), "type": "sources_changed_while_closed", "effects": effects})
        else:
            for c in self.source.candidates.values():
                contract = self.source.contracts.get(c.account)
                if not c.context and contract is not None:
                    # the sources are the ones the file was saved against, so what they hold now is what the draft was written against
                    c.context = capture_context(contract, c.claims, c.cite_ids, account_profile(self.source, c.account))
                outcome = cl.verify_candidate(c.claims, self.source.contracts, default_account=c.account)
                if outcome["ok"] != bool(c.outcome.get("ok")) and not c.invalid_reason:
                    c.invalid_reason = "the draft verified differently when the session was reopened"
        for turn in self.turns:
            if turn.source_version != self.source.source_version:
                turn.stale = True


def _default_budget() -> Budget:
    from .flag import agent_budget
    return agent_budget()

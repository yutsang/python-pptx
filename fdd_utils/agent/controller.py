"""One loop: the model picks the next tool, the controller enforces what code can.

No planner, no second model, no hand-off. The model sees the request, the tools
and what each tool returned; it answers with one JSON object per turn -- a tool
call, an answer, a question for the user, or "done". Everything that must hold
whatever the model does is enforced HERE, in code, because a guardrail that
exists only in the prompt is a suggestion:

  * a claim may cite only refs a tool returned in this run (UNSEEN_REF), so a
    figure cannot be made up and given a plausible-looking reference
  * every claim in an answer goes through the independent verifier
    (ai/claims.py); a failed claim sends the answer back, it is not softened
  * open issues come from tool RESULTS -- a total the deck replaced, tab and
    Financials in conflict, a stale source -- and a blocking one must be
    investigated or named as unresolved before an answer is accepted. This is how
    a result changes what the agent does next: the gate says what is still open
  * the user is asked only when the sources are ambiguous or in conflict; an
    ordinary gap is investigated, not handed back
  * budgets (steps, estimated tokens, seconds, steps without progress, answers
    sent back) end the run with a report of what is known, what is missing and
    what to do next -- never with a guess

Whether the model chooses well is the model's; whether what it says is allowed to
stand is not.
"""
from __future__ import annotations

import json
import re
import time
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, List, Optional, Tuple

from ..ai import claims as cl
from . import routing as R
from . import tools as T
from .session import SourceSession

STOP_ANSWERED = "answered"
STOP_DONE = "done"
STOP_NEEDS_USER = "needs_user"
STOP_STEPS = "budget_steps"
STOP_TOKENS = "budget_tokens"
STOP_TIME = "budget_time"
STOP_NO_PROGRESS = "no_progress"
STOP_REJECTED = "answer_not_acceptable"
STOP_MODEL_ERROR = "model_error"
STOP_SESSION = "budget_session"      # a cap that spans turns, roles and attempts (routing.SharedBudget) would be passed

Model = Callable[[str, str, int], Dict[str, Any]]


@dataclass
class Budget:
    max_steps: int = 12
    max_seconds: float = 300.0
    max_tokens: int = 80_000          # estimated, prompt + completion, across the run
    max_no_progress: int = 3
    max_rejections: int = 3           # answers the gate may send back


@dataclass
class Issue:
    id: str
    severity: str                     # "blocking" | "advisory"
    account: str
    message: str
    suggest: str = ""
    resolved: bool = False

    def to_dict(self) -> Dict[str, Any]:
        return dict(self.__dict__)


@dataclass
class AgentRun:
    request: str
    mode: str
    stop_reason: str = ""
    answer: Optional[Dict[str, Any]] = None
    question: str = ""
    steps: List[Dict[str, Any]] = field(default_factory=list)
    usage: Dict[str, Any] = field(default_factory=dict)
    issues: List[Dict[str, Any]] = field(default_factory=list)
    report: Dict[str, Any] = field(default_factory=dict)
    source_version: str = ""

    @property
    def ok(self) -> bool:
        return self.stop_reason in (STOP_ANSWERED, STOP_DONE)

    @property
    def tool_path(self) -> List[str]:
        return [s["tool"] for s in self.steps if s.get("tool")]

    def to_dict(self) -> Dict[str, Any]:
        return {"request": self.request, "mode": self.mode, "stop_reason": self.stop_reason, "ok": self.ok,
                "answer": self.answer, "question": self.question, "steps": self.steps, "usage": self.usage,
                "issues": self.issues, "report": self.report, "source_version": self.source_version}


@dataclass
class Memory:
    """What carries from one turn of a working session to the next (workspace.py owns it between turns).

    `seen_refs` is what the model may cite without fetching again; `issues` are the open questions the tool results
    raised, with what has since been resolved; `history` is the compact record of earlier turns the prompt shows. The
    controller reads and extends these objects in place, so a turn that ends early still leaves what it learned."""
    seen_refs: set = field(default_factory=set)
    consulted: List[str] = field(default_factory=list)
    issues: Dict[str, "Issue"] = field(default_factory=dict)
    history: List[Dict[str, Any]] = field(default_factory=list)


#: what `thinking` asks for, as AIClient.get_response takes it. Only some provider branches send either parameter at all
#: (see THINKING_HONOURED); asking for it on another is recorded as not applied, never reported as if it had been.
THINKING_EFFORTS = ("low", "medium", "high")
THINKING_HONOURED = {"local": ("allow_thinking",), "workbench": ("reasoning_effort",)}
#: provider branches that send a thinking parameter only when the provider block in config.yml declares its shape
#: (thinking_extra_body); what the gateway does with it is unverified until a probe has been run against it
THINKING_DECLARED = ("dashscope", "deepseek")

#: the name the agent path's calls carry in the client's logs and the usage ledger (`client_agent`); it is not a pipeline stage,
#: so nothing it spends is filed under the Validator. Per-agent settings, if any, live under agents.fdd_agent in config.yml.
AGENT_IDENTITY = "fdd_agent"


def thinking_params(model_type: str, thinking: str, *, declared: bool = False, reasoning_effort: str = "") -> Tuple[Dict[str, Any], bool]:
    """(keyword arguments for get_response, whether this provider branch acts on them). `thinking` is "" (the provider's
    default: nothing is sent, except that allow_thinking stays False on the branches that always had it), "off", "on" or an
    effort. `declared` says the provider block declares a thinking body (dashscope / deepseek only). An explicit
    `reasoning_effort` wins over an effort given as `thinking`."""
    want = str(thinking or "").strip().lower()
    effort = str(reasoning_effort or "").strip().lower() or (want if want in THINKING_EFFORTS else "")
    if model_type in THINKING_DECLARED:
        if not declared or not (want or effort):
            return {}, False
        params: Dict[str, Any] = {"allow_thinking": want in ("on", "true") or bool(effort)}
        if want in ("off", "false") and not effort:
            params["allow_thinking"] = False
        if effort:
            params["reasoning_effort"] = effort
        return params, True
    honoured = THINKING_HONOURED.get(model_type, ())
    params = {"allow_thinking": want in ("on", "true") or bool(effort)}
    if effort:
        params["reasoning_effort"] = effort
    sent = bool(want or effort) and (("allow_thinking" in honoured and want in ("off", "on", "true", "false")) or
                                     ("reasoning_effort" in honoured and bool(effort)))
    return params, sent


class AIClientModel:
    """The production model: AIClient.get_response, which the usage ledger and the
    response cache already sit behind. Constructed lazily so importing the
    controller needs no provider.

    The client is built under its own identity (AGENT_IDENTITY), not a pipeline stage's: the ledger's `client_agent` says
    fdd_agent and its `stage` says agent_<role>, so the agent's tokens never land in the Validator's column. `provider` is
    the branch the client actually uses: AIClient falls back to another configured provider when the requested one is not
    configured, and that difference is kept (`provider_fallback`) rather than hidden."""

    def __init__(self, model_type: str = "local", model_name: Optional[str] = None, language: str = "Chi", *,
                 thinking: str = "", ledger_stage: str = "agent", reasoning_effort: str = "", max_output: Optional[int] = None):
        from ..ai.client import AIClient
        self.client = AIClient(model_type=model_type, agent_name=AGENT_IDENTITY, language=language, model_name=model_name)
        self.requested_provider = model_type
        self.provider = str(getattr(self.client, "model_type", model_type) or model_type)
        self.provider_fallback = self.provider != model_type
        self.model_type, self.model_name, self.thinking, self.ledger_stage = self.provider, model_name, thinking, ledger_stage
        self.reasoning_effort, self.max_output = reasoning_effort, max_output
        declared = isinstance((getattr(self.client, "config_details", None) or {}).get("thinking_extra_body"), dict)
        self.params, self.thinking_applied = thinking_params(self.provider, thinking, declared=declared, reasoning_effort=reasoning_effort)
        #: a thinking parameter is sent because config declares its shape; no probe has shown the gateway acts on it
        self.thinking_unverified = self.thinking_applied and self.provider in THINKING_DECLARED

    @property
    def chat_model(self) -> str:
        return str((getattr(self.client, "config_details", None) or {}).get("chat_model") or self.model_name or "")

    def __call__(self, system: str, user: str, max_tokens: int) -> Dict[str, Any]:
        from .. import usage_ledger
        if self.max_output:
            max_tokens = min(int(max_tokens), int(self.max_output))
        with usage_ledger.scope(stage=self.ledger_stage):
            response = self.client.get_response(user, system, temperature=0.1, max_tokens=max_tokens, **self.params) or {}
        return {"content": str(response.get("content") or ""),
                "prompt_tokens": response.get("prompt_tokens"), "completion_tokens": response.get("completion_tokens"),
                "cached_prompt_tokens": response.get("cached_prompt_tokens") or response.get("prompt_cache_hit_tokens"),
                "reasoning_tokens": response.get("reasoning_tokens"), "token_source": response.get("token_usage_source"),
                "model_reported": response.get("model_reported"),
                "saved_prompt_tokens": response.get("saved_prompt_tokens"), "saved_completion_tokens": response.get("saved_completion_tokens")}


def estimate_tokens(text: str) -> int:
    cjk = len(re.findall(r"[一-鿿]", text))
    return int(cjk * 1.0 + (len(text) - cjk) / 3.5) + 1


_JSON_OBJ = re.compile(r"\{.*\}", re.S)


def parse_turn(text: str) -> Tuple[Optional[Dict[str, Any]], str]:
    raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", str(text or "").strip(), flags=re.S)
    m = _JSON_OBJ.search(raw)
    if not m:
        return None, "no JSON object in the reply"
    blob = m.group(0)
    for end in range(len(blob), 0, -1):
        try:
            obj = json.loads(blob[:end])
            if isinstance(obj, dict):
                return obj, ""
        except Exception:
            continue
    return None, "reply was not valid JSON"


def _truncate(obj: Any, limit: int) -> str:
    text = json.dumps(obj, ensure_ascii=False, default=str)
    return text if len(text) <= limit else text[:limit] + f" ...[truncated {len(text) - limit} chars; ask a narrower question]"


PROTOCOL = (
    "PROTOCOL. Reply with exactly ONE JSON object and nothing else:\n"
    '  {"tool": "<name>", "args": {...}}              call a tool\n'
    '  {"answer": {"claims": [...], "unresolved": ["..."]}}   answer (ask mode)\n'
    '  {"done": {"account": "<key>", "export": true|false, "summary": "..."}}   finish (edit mode)\n'
    '  {"ask_user": "<question>"}                     only when the sources are ambiguous or in conflict\n'
    "A claim is {\"account\", \"kind\": amount|movement|composition|share|difference|annualized|cause|descriptive,\n"
    '  "text", "periods": [ISO dates], "cites": [F/D refs], "cause_cites": [N refs], "inference": false}.\n'
    "Every number in a claim's text must be one a cited fact or derived fact carries, in the period the claim declares.\n"
    "Cite only refs that tools returned in THIS conversation. A cause needs a cited note (N ref) or inference=true.\n"
    "\"kind\" is a label. Every sentence is checked for everything its text states (amounts, periods, direction words, signs,\n"
    "percentages, causes); inference=true only waives the note behind a cause, never a figure, period, unit or sign."
)


def system_prompt(mode: str, session: SourceSession, language: str, remembered: bool = False) -> str:
    answer_lang = "繁體中文" if language in ("Chi", "Chn") else "English"
    rules = (
        "You work for a financial due diligence team on ONE workbook. You know nothing about it except what tools return.\n"
        "1. Use the tools. Never state a figure, period, flag or cause that a tool did not return.\n"
        "2. Python computes numbers: ask compute_analysis for a movement, share, sum or annualised figure and cite its D ref.\n"
        "3. A period can be marked 'replaced' (the deck uses the Financials figure, the tab was not updated): quote the statement\n"
        "   figure, do not explain it with the tab's detail, and look for the adjustment notes (N refs) that explain the difference.\n"
        "4. If evidence is missing, say so in 'unresolved'; do not fill the gap. Ask the user only if sources are ambiguous or conflict.\n"
        f"5. Write the text of claims in {answer_lang}. Amounts as 万元 or 亿元 with one decimal unless the figure is small.\n"
    )
    if mode == "edit":
        rules += ("EDIT MODE: you are given an account and an instruction. Read its evidence, write claims for the paragraph, call\n"
                  "validate_candidate until it is ok, then finish with done (export=true calls export_candidate first).\n")
    else:
        rules += "ASK MODE: investigate with tools, then answer with claims that cite what you found.\n"
    if remembered:
        rules += ("6. EARLIER IN THIS SESSION lists what has already been established and the refs you may cite again. Do not investigate it again.\n"
                  "   Investigate only what the new request needs. If the user asks to put an analysis in the report, write the claims for that\n"
                  "   paragraph from the refs shown and call validate_candidate; do not ask which account or which figures.\n"
                  "7. If a tool result lists a limit (not covered, truncated, hidden sheets skipped), say so under 'unresolved' when it matters.\n")
    return f"{rules}\nTools:\n{T.describe_tools()}\n\nWorkbook source version: {session.source_version}\n\n{PROTOCOL}"


class AgentController:
    def __init__(self, session: SourceSession, model: Model, *, budget: Optional[Budget] = None, language: Optional[str] = None,
                 memory: Optional[Memory] = None):
        self.session = session
        self.model = model
        self.budget = budget or Budget()
        self.language = language or session.language or "Chi"
        self.memory = memory
        self.seen_refs: set = memory.seen_refs if memory is not None else set()
        self.consulted: List[str] = memory.consulted if memory is not None else []
        self.issues: Dict[str, Issue] = memory.issues if memory is not None else {}
        self.transcript: List[Dict[str, Any]] = []
        self.rejections = 0
        #: routing.RoutedModel when the model is one (the working session always passes one); None for a bare callable
        self.router: Optional[R.RoutedModel] = model if isinstance(model, R.RoutedModel) else None
        #: after an escalation the new model reads the hand-off and the steps taken since, not the conversation so far
        self.handoff_text = ""
        self.view_from = 0
        self.escalations: List[Dict[str, Any]] = []
        self.last_draft: Optional[Dict[str, Any]] = None
        self._gate_codes: List[str] = []
        self._step_log: List[Tuple[int, List[str]]] = []
        #: what routing.effective_classes needs beyond the codes (filled as the turn goes)
        self._step_classes: Dict[int, Dict[str, str]] = {}
        self._unseen: Dict[str, Any] = {}
        self._cause_accounts: set = set()
        self._searched: set = set()
        self._notes_found: Dict[str, int] = {}
        self._tool_failures: Dict[str, int] = {}
        self._repeated_failure = False

    # -- the run -----------------------------------------------------------------------

    def run(self, request: str, *, mode: str = "ask", account: str = "", export: bool = False, turn_number: int = 0) -> AgentRun:
        run = AgentRun(request=request, mode=mode, source_version=self.session.source_version)
        started = time.perf_counter()
        usage: Dict[str, Any] = {"model_calls": 0, "prompt_tokens_est": 0, "completion_tokens_est": 0, "provider_prompt_tokens": 0,
                                 "provider_completion_tokens": 0, "provider_reported": False, "by_role": {}}
        router = self.router
        if router is not None:
            started_on = router.begin_turn(turn_number, R.task_kinds(request))
            if started_on is not None:
                self.escalations.append({**started_on, "handoff": ""})
        system = system_prompt(mode, self.session, self.language, remembered=bool(self.memory and self.memory.history))
        goal = request if not account else f"{request}\n[account: {account}; export: {str(export).lower()}]"
        no_progress, last_signature = 0, self._signature()
        resolved_account = self.session.resolve_account(account) if account else None
        if account and resolved_account is None:
            run.stop_reason, run.question = STOP_NEEDS_USER, f"The account '{account}' matches no single account in this workbook."
            return self._finish(run, usage, started)

        for step in range(1, self.budget.max_steps + 1):
            elapsed = time.perf_counter() - started
            if elapsed > self.budget.max_seconds:
                run.stop_reason = STOP_TIME
                break
            user = self._render(goal)
            est = estimate_tokens(system) + estimate_tokens(user)
            if usage["prompt_tokens_est"] + usage["completion_tokens_est"] + est + 1500 > self.budget.max_tokens:
                run.stop_reason = STOP_TOKENS
                break
            try:
                reply = self.model(system, user, 1500)
            except R.BudgetExhausted as exc:
                run.stop_reason, run.report["budget"] = STOP_SESSION, {"kind": exc.kind, "message": str(exc)}
                break
            except Exception as exc:                                # noqa: BLE001
                run.stop_reason, run.report["model_error"] = STOP_MODEL_ERROR, f"{type(exc).__name__}: {exc}"
                break
            text = str((reply or {}).get("content") or "")
            usage["model_calls"] += 1
            usage["prompt_tokens_est"] += est
            usage["completion_tokens_est"] += estimate_tokens(text)
            role = str((reply or {}).get("role") or "")
            row = usage["by_role"].setdefault(role or "model", {"calls": 0, "prompt_est": 0, "completion_est": 0, "provider_prompt": 0,
                                                                "provider_completion": 0, "model": (reply or {}).get("model")})
            row["calls"] += 1
            row["prompt_est"] += est
            row["completion_est"] += estimate_tokens(text)
            if (reply or {}).get("prompt_tokens") is not None:
                usage["provider_reported"] = True
                usage["provider_prompt_tokens"] += int(reply["prompt_tokens"] or 0)
                usage["provider_completion_tokens"] += int(reply.get("completion_tokens") or 0)
                row["provider_prompt"] += int(reply["prompt_tokens"] or 0)
                row["provider_completion"] += int(reply.get("completion_tokens") or 0)
            turn, error = parse_turn(text)
            record: Dict[str, Any] = {"step": step}
            if role:
                record["role"] = role
            codes: List[str] = []
            if turn is None:
                record["error"] = error
                codes = ["NOT_JSON"]
                if router is not None:
                    router.mark_last_reply_unused()
                self._note(f"SYSTEM: {error}. Reply with one JSON object only.")
                no_progress += 1
            elif "tool" in turn:
                name, args = str(turn.get("tool")), (turn.get("args") if isinstance(turn.get("args"), dict) else {})
                result = T.call_tool(self.session, name, args)
                self._absorb(name, args, result)
                record.update(tool=name, args=args, ok=result.ok, error=(result.error or {}).get("code"))
                codes = self._tool_codes(result)
                self.transcript.append({"tool": name, "args": args, "result": result.to_dict()})
            elif "answer" in turn and mode == "ask":
                accepted, feedback, answer = self._gate_answer(turn.get("answer"))
                record.update(action="answer", accepted=accepted)
                if accepted:
                    run.answer, run.stop_reason = answer, STOP_ANSWERED
                    run.steps.append(record)
                    break
                codes = list(self._gate_codes)
                self.rejections += 1
                self._note("CONTROLLER: answer not accepted. " + feedback)
                record["feedback"] = feedback
                if router is not None:
                    router.note_rejection()
                fetched = self._fetch_unread_evidence(codes)
                if fetched:
                    record["controller_fetched"] = fetched
                if self.rejections >= self.budget.max_rejections:
                    self._step_log.append((step, codes))
                    run.stop_reason = STOP_REJECTED
                    run.steps.append(record)
                    break
            elif "done" in turn and mode == "edit":
                accepted, feedback, answer = self._gate_done(turn.get("done"), resolved_account, export)
                record.update(action="done", accepted=accepted)
                if accepted:
                    run.answer, run.stop_reason = answer, STOP_DONE
                    run.steps.append(record)
                    break
                codes = list(self._gate_codes)
                self.rejections += 1
                self._note("CONTROLLER: not finished. " + feedback)
                record["feedback"] = feedback
                if router is not None:
                    router.note_rejection()
                if self.rejections >= self.budget.max_rejections:
                    self._step_log.append((step, codes))
                    run.stop_reason = STOP_REJECTED
                    run.steps.append(record)
                    break
            elif "ask_user" in turn:
                allowed, why = self._may_ask()
                record.update(action="ask_user", allowed=allowed)
                if allowed:
                    run.stop_reason, run.question = STOP_NEEDS_USER, str(turn.get("ask_user") or "")
                    run.steps.append(record)
                    break
                codes = ["ASK_NOT_ALLOWED"]
                self._note("CONTROLLER: do not ask yet. " + why + R.next_step_hint(codes))
                record["feedback"] = why
            else:
                record["error"] = "unrecognised or out-of-mode action"
                codes = ["BAD_ACTION"]
                if router is not None:
                    router.mark_last_reply_unused()
                self._note(f"SYSTEM: that action is not available in {mode} mode. {PROTOCOL.splitlines()[0]}")
            signature = self._signature()
            progress = signature != last_signature
            no_progress = 0 if progress else no_progress + 1
            last_signature = signature
            if router is not None:
                if not progress and not codes and not self._waiting_on_others():
                    codes = ["NO_PROGRESS"]             # a step that added nothing, with nothing the model was waiting on
                classes = R.effective_classes(codes, self._context(codes))
                self._step_log.append((step, codes))
                self._step_classes[step] = classes
                absent = {c: k for c, k in classes.items() if k == R.ABSENT}
                if absent:
                    record["absent"] = sorted(absent)
                    self._note("CONTROLLER:" + R.next_step_hint_for(absent))
                decision = router.observe(R.StepEvent(step=step, codes=codes, progress=progress, no_progress=no_progress, classes=classes))
                if decision.escalate:
                    handoff = self._build_handoff(goal, mode, resolved_account or "", step)
                    info = router.escalate(decision, handoff, step)
                    self.escalations.append({**info, "handoff": handoff})
                    record["escalated"] = {"to": info["to"], "model": info["model"], "reason": info["reason"]}
                    self.handoff_text, self.view_from = handoff, len(self.transcript)
                    if router.policy.fresh_rejections:
                        self.rejections = 0
                    no_progress, last_signature = 0, self._signature()
            run.steps.append(record)
            if no_progress >= self.budget.max_no_progress:
                run.stop_reason = STOP_NO_PROGRESS
                break
        else:
            run.stop_reason = STOP_STEPS
        if not run.stop_reason:
            run.stop_reason = STOP_STEPS
        return self._finish(run, usage, started)

    def _context(self, codes: List[str]) -> R.StepContext:
        """What this step's codes mean in context (routing.effective_classes): what was read, what exists, what the searches found."""
        repeats: Dict[str, int] = {}
        for _step, earlier in self._step_log:
            for code in earlier:
                repeats[code] = repeats.get(code, 0) + 1
        ctx = R.StepContext(fetched_accounts=set(self.consulted), cause_accounts=set(self._cause_accounts), searched=set(self._searched),
                            notes_found=dict(self._notes_found), repeated_tool_failure=self._repeated_failure, repeats=repeats)
        if any(c in ("UNSEEN_REF", "UNKNOWN_REF") for c in codes) and self._unseen:
            ctx.cited_refs_exist = bool(self._unseen.get("exist"))
            ctx.cited_ref_accounts = set(self._unseen.get("accounts") or [])
            ctx.cited_refs_hidden_by_paging = bool(self._unseen.get("paging"))
        self._repeated_failure = False
        return ctx

    def _fetch_unread_evidence(self, codes: List[str]) -> List[str]:
        """Evidence that exists but was not read is fetched here, by the controller, before the next model call: reading it is a
        tool call, free and deterministic, and no stronger model is needed to do it. At most two accounts per step."""
        if "UNSEEN_REF" not in codes or not self._unseen.get("exist"):
            return []
        fetched = []
        for account in [a for a in self._unseen.get("accounts") or [] if a not in self.consulted][:2]:
            result = T.call_tool(self.session, "get_account_evidence", {"account": account})
            self._absorb("get_account_evidence", {"account": account}, result)
            self.transcript.append({"tool": "get_account_evidence", "args": {"account": account}, "result": result.to_dict(), "by": "controller"})
            if result.ok:
                fetched.append(account)
        if fetched:
            self._note("CONTROLLER: the refs you cited exist; their evidence was not read, so it was fetched for you: " + ", ".join(fetched)
                       + ". Cite only refs it shows.")
        return fetched

    def _unseen_context(self, claims: List[Dict[str, Any]], unseen: List[str], default_account: str) -> None:
        """For refs no tool returned: do they exist in the account the claim names (refs are numbered per account), and if they do,
        were they among the components the evidence view left out (so paging, not a stronger model, would show them)?"""
        exist, accounts, paging = True, set(), False
        for claim in claims:
            account = self.session.resolve_account(str(claim.get("account") or default_account)) or ""
            contract = self.session.contracts.get(account)
            for ref in list(claim.get("cites") or []) + list(claim.get("cause_cites") or []):
                if ref not in unseen:
                    continue
                item = (contract.fact(ref) or contract.note(ref)) if contract is not None else None
                if item is None:
                    exist = False
                    continue
                accounts.add(account)
                if item.get("role") == "component" or item.get("line") not in (None, "", "(total)"):
                    paging = paging or account in self.consulted
        self._unseen = {"exist": exist, "accounts": sorted(accounts), "paging": paging}

    def _waiting_on_others(self) -> bool:
        """An issue is open that a tool or a person has to settle: a step that adds nothing then is not the model's failure."""
        return any(i.severity == "blocking" and not i.resolved for i in self.issues.values())

    def _tool_codes(self, result: T.ToolResult) -> List[str]:
        if result.ok:
            return []
        code = (result.error or {}).get("code", "TOOL_EXCEPTION")
        if code in R.WRAPPERS:
            inner = sorted(((result.data or {}).get("outcome") or {}).get("by_code") or {})
            return inner or [code]
        return [code]

    # -- what the model is shown ---------------------------------------------------------------

    def _note(self, text: str) -> None:
        self.transcript.append({"note": text})

    def _history_block(self) -> str:
        history = self.memory.history if self.memory is not None else []
        if not history:
            return ""
        lines = ["EARLIER IN THIS SESSION (checked against the sources as they were then; refs of a current entry may be cited again):"]
        for h in history[-4:-1]:
            warn = "  [the workbook has changed since: do not reuse its figures or refs]" if h.get("stale") else ""
            lines.append(f"  [{h['n']} {h['mode']}] {str(h['request'])[:120]}{warn}\n      -> {str(h['text'])[:300]}")
        last = history[-1]
        if last.get("stale"):
            lines.append(f"  [{last['n']} {last['mode']}] {str(last['request'])[:200]}  [the workbook has changed since: do not reuse its figures or refs]\n"
                         f"      -> {str(last['text'])[:300]}")
            return "\n".join(lines)
        lines.append(f"  [{last['n']} {last['mode']}] {str(last['request'])[:200]}\n      -> {str(last['text'])[:800]}")
        if last.get("claims"):
            lines.append("      its claims (reuse their refs and wording where the request allows): "
                         + _truncate([{k: c.get(k) for k in ("account", "kind", "periods", "cites", "cause_cites", "inference", "text") if c.get(k) not in (None, [], "")}
                                      for c in last["claims"]], 2600))
        if last.get("unresolved"):
            lines.append("      left unresolved then: " + "; ".join(str(u)[:160] for u in last["unresolved"][:5]))
        return "\n".join(lines)

    def _render(self, goal: str) -> str:
        parts = [f"REQUEST: {goal}"]
        remembered = self._history_block()
        if remembered:
            parts.append(remembered)
        if self.handoff_text:
            parts.append(self.handoff_text)
        open_issues = [i for i in self.issues.values() if not i.resolved]
        if open_issues:
            parts.append("OPEN ISSUES (from tool results):\n" + "\n".join(
                f"  [{i.severity}] {i.account or 'workbook'}: {i.message}" + (f"  -> {i.suggest}" if i.suggest else "")
                for i in open_issues))
        if self.consulted:
            parts.append("ACCOUNTS CONSULTED: " + ", ".join(self.consulted))
        visible = self.transcript[self.view_from:]
        tool_items = [t for t in visible if "tool" in t]
        recent = {id(t) for t in tool_items[-3:]}
        for item in visible:
            if "note" in item:
                parts.append(item["note"])
                continue
            body = _truncate(item["result"], 12000 if item["tool"] == "get_account_evidence" else 5000) if id(item) in recent else (
                f"ok={item['result']['ok']} refs={item['result']['basis'][:12]} error={item['result']['error']}")
            parts.append(f"TOOL {item['tool']}({json.dumps(item['args'], ensure_ascii=False)}) RETURNED:\n{body}")
        parts.append("Your next JSON object:")
        return "\n\n".join(parts)

    def _build_handoff(self, goal: str, mode: str, account: str, step: int) -> str:
        """The task state the escalation model starts from (routing.build_handoff): built from what the controller holds, not
        from the first model's words."""
        draft = self.last_draft or {}
        claims = [c for c in draft.get("claims") or [] if isinstance(c, dict)]
        seen, cited = set(), []
        for src in self._sources(claims):
            key = (src.get("account"), src.get("ref"))
            if key not in seen:
                seen.add(key)
                cited.append({k: src.get(k) for k in ("ref", "period", "basis", "line", "value")})
        actions = [{"tool": t["tool"], "args": t["args"], "ok": t["result"]["ok"], "error": (t["result"]["error"] or {}).get("code")}
                   for t in self.transcript if "tool" in t]
        counts: Dict[str, int] = {}
        for _step, codes in self._step_log:
            for code in codes:
                counts[code] = counts.get(code, 0) + 1
        failures = [{"code": c, "class": R.classify(c), "count": n} for c, n in sorted(counts.items(), key=lambda kv: -kv[1])]
        remaining = dict(self.router.budget.remaining()) if self.router is not None else {}
        remaining["steps in this turn"] = max(0, self.budget.max_steps - step)
        return R.build_handoff(goal=goal, mode=mode, account=account, source_version=self.session.source_version, draft=draft, cited=cited,
                               actions=actions, failures=failures, issues=[i.to_dict() for i in self.issues.values()], remaining=remaining,
                               max_tokens=self.router.policy.handoff_max_tokens if self.router is not None else 3000)

    # -- what a tool result changes ----------------------------------------------------------------

    def _signature(self) -> Tuple:
        # Only a draft that VERIFIED moves the work on. A rejected draft is a new attempt, not progress: counted as progress, a model
        # rewriting one failing sentence again and again (each a different digest) never reached the escalation rule in edit mode.
        return (len(self.seen_refs), sum(i.resolved for i in self.issues.values()), len(self.issues), len(self.consulted),
                tuple(sorted((a, c.claims_digest) for a, c in self.session.candidates.items() if c.validated)),
                len(self.session.exports))

    def _absorb(self, name: str, args: Dict[str, Any], result: T.ToolResult) -> None:
        self.seen_refs |= set(result.basis)
        if not result.ok:
            key = f"{name}|{json.dumps(args, sort_keys=True, ensure_ascii=False, default=str)}|{(result.error or {}).get('code')}"
            self._tool_failures[key] = self._tool_failures.get(key, 0) + 1
            self._repeated_failure = self._tool_failures[key] > 1
        if name == "search_notes" and result.ok and args.get("account"):
            account = self.session.resolve_account(str(args.get("account"))) or str(args.get("account"))
            self._searched.add(account)
            self._notes_found[account] = max(self._notes_found.get(account, 0), len((result.data or {}).get("hits") or []))
        if result.error and result.error["code"] == "STALE_SOURCE":
            self._issue("stale", "blocking", "", "the workbook changed on disk since this session read it; figures may no longer exist",
                        "tell the user the sources moved; do not answer from earlier results")
        if name == "inspect_sources" and result.ok and (result.data.get("financials_pairing") or {}).get("suspect"):
            pairing = result.data["financials_pairing"]
            self._issue("conflict:pairing", "blocking", "",
                        f"only {pairing['agreeing_latest']} of {pairing['compared']} accounts tie to the Financials sheet on their latest period; "
                        "it may belong to another entity or period", "ask the user to confirm the entity and the Financials sheet")
        if name == "inspect_sources" and result.ok:
            for item in (result.data.get("selection") or {}).get("excluded") or []:
                self._issue(f"excluded:{item['account']}", "advisory", item["account"],
                            f"left out of commentary ({item['reason']})", "")
            for u in (result.data.get("selection") or {}).get("unmatched_financials") or []:
                self._issue(f"unmatched:{u['financials_account']}", "advisory", "",
                            f"Financials line '{u['financials_account']}' has a balance but no schedule tab", "")
        if name == "get_account_evidence" and result.ok:
            account = self.session.resolve_account(str(args.get("account"))) or str(args.get("account"))
            if account not in self.consulted:
                self.consulted.append(account)
            self._notes_found[account] = max(self._notes_found.get(account, 0), len((result.data or {}).get("notes") or []))
            self._issues_from_evidence(account, result.data)
        if name == "validate_candidate" and isinstance(args.get("claims"), list):
            account = self.session.resolve_account(str(args.get("account"))) or str(args.get("account") or "")
            outcome = (result.data or {}).get("outcome") or {}
            self.last_draft = {"claims": [dict(c, account=c.get("account") or account) for c in args["claims"] if isinstance(c, dict)],
                               "objections": [] if result.ok else self._objections(outcome)}

    def _issues_from_evidence(self, account: str, view: Dict[str, Any]) -> None:
        notes = view.get("notes") or []
        adjustment_seen = any(n.get("kind") == "adjustment" for n in notes)
        for issue in view.get("issues") or []:
            if issue.get("severity") == "blocked":
                self._issue(f"source:{account}", "blocking", account, issue.get("message", ""), "")
        perms = view.get("permissions") or {}
        for period, perm in perms.items():
            if perm.get("describe_difference") and not perm.get("describe_total"):
                self._issue(f"conflict:{account}:{period}", "blocking", account,
                            f"{period}: tab and Financials disagree and no rule chose between them; state the difference, not either figure",
                            "compute_analysis(op=difference); name the conflict under 'unresolved' or ask the user which to follow")
        checks = view.get("checks") or {}
        replaced = [p for p, row in checks.items() if row.get("total_vs_financials") == "mismatch" and (perms.get(p) or {}).get("describe_total")]
        for period in replaced:
            iid = f"replaced:{account}:{period}"
            self._issue(iid, "blocking", account,
                        f"{period}: the deck uses the Financials figure; the tab was not updated, so its detail cannot explain the change",
                        "read the adjustment notes (N refs) for the reason, or list the reason as unresolved")
            if adjustment_seen:
                self.issues[iid].resolved = True       # the evidence that could explain it has been read
        if view.get("annualization", {}).get("annualized"):
            self._issue(f"annualized:{account}", "advisory", account,
                        "figures are part-year; an annual figure must be called annualised (年化), not a forecast", "")

    def _issue(self, issue_id: str, severity: str, account: str, message: str, suggest: str) -> None:
        if issue_id not in self.issues:
            self.issues[issue_id] = Issue(issue_id, severity, account, message, suggest)

    # -- the gates --------------------------------------------------------------------------------------

    @staticmethod
    def _objections(outcome: Dict[str, Any]) -> List[str]:
        out = []
        for v in outcome.get("verdicts") or []:
            if v.get("verdict") in ("failed", "unverified"):
                bad = [f for f in v.get("findings") or [] if f.get("severity") in ("failed", "review")]
                if bad:
                    out.append(f"claim {v.get('claim_id')} ({', '.join(sorted({f['code'] for f in bad}))}): {bad[0].get('message', '')}")
        return out

    def _gate_answer(self, answer: Any) -> Tuple[bool, str, Optional[Dict[str, Any]]]:
        self._gate_codes = []
        if not isinstance(answer, dict):
            self._gate_codes = ["BAD_ANSWER"]
            return False, "the answer must be an object with 'claims' and 'unresolved'." + R.next_step_hint(self._gate_codes), None
        claims = [c for c in (answer.get("claims") or []) if isinstance(c, dict)]
        unresolved = [str(u) for u in (answer.get("unresolved") or [])]
        if not claims and not unresolved:
            self._gate_codes = ["EMPTY_ANSWER"]
            return False, "the answer is empty: give claims, or say what is unresolved." + R.next_step_hint(self._gate_codes), None
        default_account = self.consulted[0] if len(self.consulted) == 1 else ""
        stamped = [dict(c, source_version=c.get("source_version") or self.session.source_version) for c in claims]
        unseen = sorted({r for c in stamped for r in list(c.get("cites") or []) + list(c.get("cause_cites") or [])
                         if r not in self.seen_refs})
        if unseen:
            self._unseen_context(stamped, unseen, default_account)
            self._gate_codes = ["UNSEEN_REF"]
            self.last_draft = {"claims": stamped, "objections": [f"UNSEEN_REF: {unseen}"]}
            return False, (f"UNSEEN_REF: you cite {unseen}, which no tool returned in this conversation. "
                           "Fetch the evidence (get_account_evidence / compute_analysis) and cite what it returns."
                           + R.next_step_hint(self._gate_codes)), None
        outcome = {"verdicts": [], "ok": True, "failed": 0, "verified": 0, "inference": 0, "by_code": {}}
        if stamped:
            outcome = cl.verify_candidate(stamped, self.session.contracts, default_account=default_account)
            if not outcome["ok"]:
                bad = [f"claim {v['claim_id']} ({', '.join(sorted({f['code'] for f in v['findings'] if f['severity'] in ('failed', 'review')}))}): "
                       + next((f['message'] for f in v['findings'] if f['severity'] in ('failed', 'review')), "")
                       for v in outcome["verdicts"] if v["verdict"] in ("failed", "unverified")]
                self._gate_codes = sorted(outcome.get("by_code") or {})
                self._cause_accounts = {v.get("account") or default_account for v in outcome["verdicts"]
                                        if any(f.get("code") == "CAUSE_UNSUPPORTED" for f in v.get("findings") or [])}
                self.last_draft = {"claims": stamped, "objections": self._objections(outcome)}
                return False, "the verifier did not accept: " + " | ".join(bad)[:900] + R.next_step_hint(self._gate_codes), None
        blocking = [i for i in self.issues.values() if i.severity == "blocking" and not i.resolved
                    and not any((i.account and i.account in u) or i.id in u for u in unresolved)]
        if blocking:
            self._gate_codes = [i.id for i in blocking]
            self.last_draft = {"claims": stamped, "objections": ["open blocking issue(s) not named under 'unresolved'"]}
            return False, ("open blocking issue(s) neither investigated nor named in 'unresolved': "
                           + " | ".join(f"{i.account or 'workbook'}: {i.message} ({i.suggest})" for i in blocking)[:900]
                           + R.next_step_hint(self._gate_codes)), None
        sources = self._sources(stamped)
        return True, "", {
            "text": cl.render_claims(stamped), "claims": stamped, "verdicts": outcome["verdicts"],
            "verification": {"verified": outcome["verified"], "inference": outcome["inference"], "failed": outcome["failed"]},
            "sources": sources, "unresolved": unresolved,
            "open_advisory": [i.to_dict() for i in self.issues.values() if not i.resolved and i.severity == "advisory"],
            "source_version": self.session.source_version,
        }

    def _gate_done(self, done: Any, account: Optional[str], export: bool) -> Tuple[bool, str, Optional[Dict[str, Any]]]:
        self._gate_codes = []
        if not isinstance(done, dict):
            self._gate_codes = ["BAD_ANSWER"]
            return False, "'done' must be an object." + R.next_step_hint(self._gate_codes), None
        key = self.session.resolve_account(str(done.get("account") or account or "")) or account
        candidate = self.session.candidates.get(key or "")
        if candidate is None or not candidate.validated or candidate.source_version != self.session.source_version:
            self._gate_codes = ["NO_VALIDATED_CANDIDATE"]
            return False, (f"no validated candidate for '{key}' against the current sources: call validate_candidate until it is ok."
                           + R.next_step_hint(self._gate_codes)), None
        if export:
            # a SUCCESSFUL export of THIS draft against THESE sources: a blocked or layout-failed export,
            # an export of an earlier draft of the account, or one made against older sources does not count
            proof = [e for e in self.session.exports if e.get("ok") and e["source_version"] == self.session.source_version
                     and key in (e.get("touched_slides") or {}) and (e.get("candidates") or {}).get(key) == candidate.claims_digest]
            if not proof:
                failed = [e for e in self.session.exports if not e.get("ok")]
                detail = f" (the last attempt ended {(failed[-1].get('delivery') or {}).get('status', 'unrecorded')})" if failed else ""
                self._gate_codes = ["EXPORT_MISSING"]
                return False, ("export was requested and no successful export of the current draft exists" + detail
                               + ": call export_candidate and resolve what it reports." + R.next_step_hint(self._gate_codes)), None
        return True, "", {"account": key, "text": cl.render_claims(candidate.claims), "claims": candidate.claims,
                          "verdicts": candidate.outcome["verdicts"], "exports": list(self.session.exports),
                          "summary": str(done.get("summary") or ""), "source_version": self.session.source_version}

    def _may_ask(self) -> Tuple[bool, str]:
        stale = self.session.stale_reason()
        if stale:
            return True, ""
        if self.session.binding.statement_source in ("", "unknown"):
            return True, ""
        conflicts = [i for i in self.issues.values() if i.id.startswith("conflict:") and not i.resolved]
        if conflicts:
            return True, ""
        return False, ("the sources are not ambiguous or in conflict, so this is an investigation, not a question. "
                       "Call a tool, or answer with what is unresolved.")

    def _sources(self, claims: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        out, seen = [], set()
        for claim in claims:
            contract = self.session.contracts.get(claim.get("account", ""))
            for ref in list(claim.get("cites") or []) + list(claim.get("cause_cites") or []):
                if contract is None or (claim.get("account"), ref) in seen:
                    continue
                seen.add((claim.get("account"), ref))
                item = contract.fact(ref) or contract.note(ref) or {}
                out.append({"account": claim.get("account"), "ref": ref, "period": item.get("period"), "basis": item.get("basis") or item.get("kind"),
                            "line": item.get("line"), "value": item.get("value"), "sheet": item.get("sheet"), "row": item.get("row"),
                            "formula": item.get("formula"), "text": (item.get("text") or "")[:120] or None})
        return out

    # -- ending ------------------------------------------------------------------------------------------

    def _finish(self, run: AgentRun, usage: Dict[str, Any], started: float) -> AgentRun:
        usage["seconds"] = round(time.perf_counter() - started, 2)
        run.usage = usage
        run.issues = [i.to_dict() for i in self.issues.values()]
        if not run.ok:
            run.report = {**run.report, **self._stop_report(run)}
        if self.router is not None:
            open_human = any(R.classify(i.id) == R.HUMAN and not i.resolved for i in self.issues.values())
            last_codes = [c for _step, codes in self._step_log[-3:] for c in codes]
            last_classes: Dict[str, str] = {}
            for step_n, _codes in self._step_log[-3:]:
                last_classes.update(self._step_classes.get(step_n) or {})
            run.report["routing"] = {
                "stop_class": self.router.stop_class(run.stop_reason, open_human=open_human, last_codes=last_codes, ok=run.ok,
                                                     last_classes=last_classes),
                "role_at_end": self.router.role, "mode": self.router.policy.mode,
                "models": {role: {"provider": spec.provider, "model": spec.model, "thinking": spec.thinking, "thinking_applied": spec.thinking_applied,
                                  "thinking_unverified": spec.thinking_unverified, "max_output": spec.max_output,
                                  "provider_requested": spec.provider_requested or None}
                           for role, spec in self.router.roles.items()},
                "escalations": self.escalations, "steps": self.router.turn_log}
        return run

    def _stop_report(self, run: AgentRun) -> Dict[str, Any]:
        """What is known, what is missing, what to do -- from the state, not from the model."""
        known = []
        for account in self.consulted:
            contract = self.session.contracts.get(account)
            if contract and contract.periods:
                latest = contract.periods[-1]
                total = contract.total(latest)
                known.append({"account": account, "latest_period": latest, "balance_yuan": total,
                              "checks": contract.status_matrix().get(latest)})
        gaps = [f"{i.account or 'workbook'}: {i.message}" for i in self.issues.values() if not i.resolved]
        next_steps = [i.suggest for i in self.issues.values() if not i.resolved and i.suggest]
        why = {
            STOP_STEPS: f"the step budget ({self.budget.max_steps}) ran out; raise it or narrow the question",
            STOP_TOKENS: "the token budget ran out; narrow the question",
            STOP_TIME: "the time budget ran out",
            STOP_NO_PROGRESS: f"{self.budget.max_no_progress} steps in a row added no evidence and resolved nothing",
            STOP_REJECTED: "the answer was sent back more times than allowed; the last objections are in the steps",
            STOP_NEEDS_USER: "the sources need a decision from the user",
            STOP_MODEL_ERROR: "the model call failed",
            STOP_SESSION: "a cap over the whole working session was reached: " + str((run.report.get("budget") or {}).get("message") or ""),
        }.get(run.stop_reason, run.stop_reason)
        return {"why_stopped": why, "known": known, "gaps": gaps, "next_steps": next_steps, "accounts_consulted": list(self.consulted)}

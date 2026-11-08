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

A paragraph for ONE known account (edit mode) starts as a prepared write when the
controller is built with prepare=True: the controller reads the account's
evidence itself (a free, deterministic tool call), computes what Python computes,
and asks the model once for the paragraph, with no tool menu. The verifier checks
it; a rejected draft goes back with only its rejected claims (the verified ones
are kept as they were, byte for byte) for a local repair. The model can say the
evidence is not enough ("need_evidence"), and then the turn continues in the tool
loop above with that evidence already read. Escalation is the same rule as in the
loop: consecutive capability failures move the repair to the stronger model.
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
#: the paragraph verifies, but the reviewer's objection to its meaning stands (the sentence is unchanged, or still objected to after
#: a rewrite): a draft, not a finished paragraph, and the state goes with the candidate into the deck's delivery decision
STOP_SEMANTIC_BLOCKED = "semantic_blocked"
#: a review was due (Python found something to judge) and did not happen: the call failed or the reply was not in the protocol
STOP_REVIEW_PENDING = "review_pending"

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

_EN_MONTHS = ("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")


def _date_words(period: str, chinese: bool) -> str:
    y, m, d = (int(x) for x in str(period)[:10].split("-"))
    return f"{y}年{m}月{d}日" if chinese else f"{d} {_EN_MONTHS[m - 1]} {y}"


def _money(yuan: float, chinese: bool) -> str:
    """How the house style writes an amount: 亿元 above a hundred million, 万元 otherwise (two decimals under ten thousand yuan);
    CNY millions in English. The verifier reads every form back at its printed precision."""
    if chinese:
        if abs(yuan) >= 1e8:
            return f"{yuan / 1e8:,.2f}亿元"
        return f"{yuan / 1e4:,.1f}万元" if abs(yuan) >= 1e4 or yuan == 0 else f"{yuan / 1e4:,.2f}万元"
    return f"CNY{yuan / 1e6:,.1f} million" if abs(yuan) >= 1e5 else f"CNY{yuan / 1e6:,.3f} million"


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


def _claim_sha(claim: Dict[str, Any]) -> str:
    """One claim's content (text, cites, periods, kind, inference): an objection is bound to this, not to a list position."""
    import hashlib
    body = {k: claim.get(k) for k in ("kind", "text", "periods", "cites", "cause_cites", "inference")}
    return hashlib.sha256(json.dumps(body, ensure_ascii=False, sort_keys=True, default=str).encode("utf-8")).hexdigest()[:16]


def cl_digest(claims: List[Dict[str, Any]]) -> str:
    return json.dumps(claims, ensure_ascii=False, sort_keys=True, default=str)


def _compact(obj: Any) -> str:
    """JSON with no padding: the same values, about a tenth fewer characters than the default separators."""
    return json.dumps(obj, ensure_ascii=False, default=str, separators=(",", ":"))


def _truncate(obj: Any, limit: int) -> str:
    text = _compact(obj)
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


WRITER_PROTOCOL = (
    "Reply with exactly ONE JSON object and nothing else:\n"
    '  {"draft": {"claims": [...], "unresolved": ["..."]}}       the paragraph, as claims\n'
    '  {"need_evidence": "<what is missing and why>"}           only if the evidence below cannot support a paragraph at all\n'
)
REPAIR_PROTOCOL = (
    "Reply with exactly ONE JSON object and nothing else:\n"
    '  {"fix": [{"index": <n>, "claim": {...}}, {"index": <n>, "drop": "<why; it goes to unresolved>"}],\n'
    '   "add": [{...claims to add...}], "unresolved": ["..."]}\n'
    "Change only the claims listed under REJECTED (by their index). The claims under KEPT were verified and stay exactly as they are.\n"
    '  {"need_evidence": "<what is missing and why>"}           only if the evidence below cannot settle an objection\n'
)
SEED_PROTOCOL = (
    "Reply with exactly ONE JSON object and nothing else:\n"
    '  {"add": [{...claims...}], "unresolved": ["..."]}\n'
    "Python wrote the claims under KEPT from the evidence and verified them; they open the paragraph unchanged. Add the rest:\n"
    "the main detail lines where detail is permitted, causes only from notes, and anything material the notes say. Do not restate the\n"
    "KEPT figures. An empty 'add' is right when the evidence holds nothing more worth saying.\n"
)
REVIEW_PROTOCOL = (
    "Reply with exactly ONE JSON object and nothing else:\n"
    '  {"objections": [{"index": <n>, "why": "..."}], "missing": [{"ref": "<N ref>", "why": "..."}]}\n'
    "Both lists empty when the paragraph is sound. Do not rewrite the paragraph."
)
REVIEW_SYSTEM = (
    "You review ONE paragraph of a financial due diligence report against EVIDENCE, which Python read from the workbook. Python has\n"
    "already checked every figure, period, sign and citation in it: do not re-check arithmetic. Judge meaning only:\n"
    "- a cause or explanation the cited notes do not state, or state differently;\n"
    "- a balance described as a flow, or a flow as a balance; one period's fact applied to other periods;\n"
    "- a parent line and its own children both counted; a detail line presented as the whole;\n"
    "- a note's party, condition, timing, restriction or amount misstated;\n"
    "- a material fact in the notes left out (give its N ref under 'missing').\n"
    "Object only where EVIDENCE shows the problem; a style preference is not an objection."
)
CLAIM_SCHEMA = (
    "A claim is {\"kind\": amount|movement|composition|share|difference|annualized|cause|descriptive, \"text\",\n"
    '  "periods": [ISO dates], "cites": [F/D refs], "cause_cites": [N refs], "inference": false}.\n'
    "Every number in a claim's text must be one a cited F or D ref carries, in the period the claim declares. Cite only refs shown\n"
    "in EVIDENCE. A cause needs a cited note (N ref) or inference=true; inference=true waives only the note, never a figure, period,\n"
    "unit or sign. Every sentence is checked for everything its text states (amounts, periods, direction words, signs, percentages, causes)."
)


def writer_system_prompt(language: str) -> str:
    """The prepared write's instructions: the same rules as the tool loop's, without the tool menu (nothing here calls a tool)."""
    answer_lang = "繁體中文" if language in ("Chi", "Chn") else "English"
    return (
        "You write ONE account's paragraph for a financial due diligence report. EVIDENCE is what Python read from the workbook and\n"
        "checked; you know nothing else about it.\n"
        "1. State only figures, periods, flags and causes EVIDENCE carries. Python computed every movement, difference and annualised\n"
        "   figure (D refs): cite the D ref, never compute a figure yourself.\n"
        "2. totals list each period under several bases; the deck uses basis authority_adjusted. A period whose permissions say\n"
        "   describe_total=false is in conflict: state the difference (its D ref), not either figure. Detail lines may explain a total\n"
        "   only where permissions say explain_with_detail=true. Read the notes (N refs) for causes and adjustments.\n"
        "3. The paragraph states the latest balance the deck uses (REQUIRED), the movement where a D ref exists, the main detail\n"
        "   lines where detail is permitted, and causes only from notes. A line marked nested_in is a sub-line of that line: never\n"
        "   add it to its parent.\n"
        "4. Notes are evidence too. A figure, percentage or date written in a note may be stated when the claim cites that note in\n"
        "   'cites' (an amount in the note's own currency, e.g. 万美元; never convert it). MUST ADDRESS lists the notes Python found\n"
        "   material: state what each says in a claim that cites it, or name its ref in 'unresolved' with the reason. None may be\n"
        "   left out silently.\n"
        "5. What the evidence cannot settle goes in 'unresolved'; do not fill the gap.\n"
        f"6. Write for the report, in {answer_lang}: amounts as 万元 or 亿元 with one decimal unless the figure is small; periods as\n"
        "   period_labels gives them (an income-statement line is a flow: 2024年度, 2026年1-3月, never '2024年12月31日营业成本');\n"
        "   no refs (F12, N3) and no internal names (authority_adjusted, raw_schedule) in the text. No percentage change over a\n"
        "   negative base or across a change of sign. The tab's lines are not the make-up of an adjusted balance unless they add\n"
        "   up to it.\n\n"
        + CLAIM_SCHEMA
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
                 memory: Optional[Memory] = None, prepare: bool = False, seed: bool = False, review: bool = False):
        self.session = session
        #: prepared write only: Python writes the balance and movement sentences and the model adds the rest (seed); the stronger
        #: model reads a verified paragraph for meaning when Python finds something a reviewer must judge (review)
        self.seed, self.review = seed, review
        #: edit mode on one account starts with a prepared write (see the module docstring); the tool loop is the fallback
        self.prepare = prepare
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
        #: a reply the prepared write received but the tool loop must handle (the writer called a tool): not paid for twice
        self._pending_reply: Optional[Tuple[str, str]] = None
        #: the prepared write's evidence and the notes it must address (set when it starts)
        self._view: Optional[Dict[str, Any]] = None
        self._must: List[Dict[str, str]] = []

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
        if mode == "edit" and resolved_account and self.prepare:
            if self._prepared_write(run, goal, resolved_account, export, usage, started):
                return self._finish(run, usage, started)
            run.report["prepared_fallback"] = run.report.get("prepared_fallback") or "the writer asked for the tool loop"
            no_progress, last_signature = 0, self._signature()

        for step in range(len(run.steps) + 1, self.budget.max_steps + 1):
            elapsed = time.perf_counter() - started
            if elapsed > self.budget.max_seconds:
                run.stop_reason = STOP_TIME
                break
            if self._pending_reply is not None:
                (text, role), self._pending_reply = self._pending_reply, None
            else:
                called = self._call_model(run, system, self._render(goal), 1500, usage)
                if called is None:
                    break
                text, role = called
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

    def _call_model(self, run: AgentRun, system: str, user: str, max_tokens: int, usage: Dict[str, Any], *, role: str = "",
                    purpose: str = "") -> Optional[Tuple[str, str]]:
        """One model call with the turn's own accounting. None when the call may not be made or failed; run.stop_reason says why.
        `role` + `purpose` call a named role of the router without moving the turn to it (a review)."""
        est = estimate_tokens(system) + estimate_tokens(user)
        if usage["prompt_tokens_est"] + usage["completion_tokens_est"] + est + max_tokens > self.budget.max_tokens:
            run.stop_reason = STOP_TOKENS
            return None
        try:
            reply = self.router.call_as(role, system, user, max_tokens, purpose) if (role and self.router is not None) \
                else self.model(system, user, max_tokens)
        except R.BudgetExhausted as exc:
            run.stop_reason, run.report["budget"] = STOP_SESSION, {"kind": exc.kind, "message": str(exc)}
            return None
        except Exception as exc:                                # noqa: BLE001
            run.stop_reason, run.report["model_error"] = STOP_MODEL_ERROR, f"{type(exc).__name__}: {exc}"
            return None
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
        return text, role

    # -- the prepared write -------------------------------------------------------------------------

    WRITE_ALLOWANCE = 3000                 # a paragraph of claims as JSON; a role's max_output still caps it

    def _prepare_evidence(self, account: str) -> Optional[Dict[str, Any]]:
        """What the controller reads before the model writes: the account's evidence (with the standard movements Python
        derives) and, for a part-year period, the annualised figure. Tool calls, made here, recorded as the controller's."""
        result = T.call_tool(self.session, "get_account_evidence", {"account": account})
        self._absorb("get_account_evidence", {"account": account}, result)
        self.transcript.append({"tool": "get_account_evidence", "args": {"account": account}, "result": result.to_dict(), "by": "controller"})
        if not result.ok:
            return None
        view = dict(result.data)
        view["period_labels"] = self._period_labels(account, view)
        if (view.get("annualization") or {}).get("annualized") and view.get("periods"):
            extra = T.call_tool(self.session, "compute_analysis", {"account": account, "op": "annualized", "period": view["periods"][-1]})
            self._absorb("compute_analysis", {"account": account, "op": "annualized"}, extra)
            if extra.ok:
                derived = (extra.data or {}).get("derived") or {}
                if derived and all(d.get("ref") != derived.get("ref") for d in view.get("derived") or []):
                    view["derived"] = list(view.get("derived") or []) + [derived]
        return view

    def _period_labels(self, account: str, view: Dict[str, Any]) -> Dict[str, str]:
        """How each period is written: a balance at its date; a flow (income statement) as the year, or the months it covers."""
        chinese = self.language in ("Chi", "Chn")
        contract = self.session.contracts.get(account)
        flow = str(getattr(contract, "statement_type", "") or "").upper().startswith("IS")
        months = (view.get("annualization") or {}).get("months")
        latest = (view.get("periods") or [""])[-1]
        out = {}
        for p in view.get("periods") or []:
            y, m, _d = (int(x) for x in p[:10].split("-"))
            if not flow:
                out[p] = _date_words(p, chinese)
            elif m == 12:
                out[p] = f"{y}年度" if chinese else f"FY{y}"
            elif p == latest and months and int(months) == m:
                out[p] = f"{y}年1-{m}月" if chinese else f"the {int(months)} months ended {_date_words(p, False)}"
            else:
                out[p] = f"截至{_date_words(p, True)}止期间" if chinese else f"the period ended {_date_words(p, False)}"
        return out

    _MATERIAL_WORDS = ("关联", "质保", "尾款", "保证金", "押金", "减值", "坏账", "抵押", "担保", "质押", "直线法", "重分类", "协议",
                       "合同", "账龄", "逾期", "诉讼", "变更", "利率", "未付", "受限", "冻结", "承诺", "实缴", "注册资本")
    _REMARK_VALUE = re.compile(r"\d{4}-\d{2}-\d{2}(?: annualised)?:\s*(-?[\d.]+)")

    def _must_address(self, view: Dict[str, Any]) -> List[Dict[str, str]]:
        """The notes a paragraph may not leave out silently, decided by Python: an adjustment of at least 5% of the latest balance; a
        note on a line that is at least 5% of it; a note that names a related party, a retention, a deposit, an impairment, a
        pledge, a straight-line accrual, a reclassification, an agreement, an ageing... unless its line is known to be under 1%.
        A remark that repeats a note on the same line, and a note contained in another, are not listed twice. At most eight."""
        latest = (view.get("periods") or [""])[-1]
        total = next((abs(float(t.get("yuan") or 0)) for t in view.get("totals") or []
                      if t.get("period") == latest and t.get("basis") == "authority_adjusted"), 0.0)
        notes = view.get("notes") or []
        label_value: Dict[str, float] = {}
        for n in notes:
            label = str(n.get("text") or "").split("|")[0].strip(" *")
            m = self._REMARK_VALUE.search(str(n.get("text") or ""))
            if m:
                label_value[label] = max(label_value.get(label, 0.0), abs(float(m.group(1))))
        picked: List[Dict[str, str]] = []
        seen_labels: set = set()
        ordered = sorted(notes, key=lambda n: {"adjustment": 0, "supporting_note": 1}.get(n.get("kind"), 2))
        for n in ordered:
            text = str(n.get("text") or "")
            label = text.split("|")[0].strip(" *")
            amounts = [abs(sp["value"]) for sp in cl.text_amounts(text) if cl._currency_at(text, sp["start"], sp["end"]) == "CNY"]
            weight = max(amounts + ([label_value[label]] if label in label_value else []), default=None)
            share = (weight / total) if (weight is not None and total) else None
            keyword = any(w in text for w in self._MATERIAL_WORDS)
            if n.get("kind") == "adjustment":
                why = f"adjustment of {share:.0%} of the balance" if share is not None and share >= 0.05 else ""
            elif share is not None and share >= 0.05:
                why = f"its line is {share:.0%} of the balance"
            elif keyword and (share is None or share >= 0.01):
                why = "names " + next(w for w in self._MATERIAL_WORDS if w in text)
            else:
                why = ""
            if not why or (n.get("kind") != "adjustment" and label and label in seen_labels):
                continue
            if any(text.split("|")[-1].strip() and text.split("|")[-1].strip() in other["text"] for other in picked):
                continue
            seen_labels.add(label)
            picked.append({"ref": n["ref"], "why": why, "text": text[:160]})
            if len(picked) >= 8:
                break
        return picked

    @staticmethod
    def _unaddressed(must: List[Dict[str, str]], claims: List[Dict[str, Any]], outcome: Dict[str, Any], unresolved: List[str]) -> List[Dict[str, str]]:
        """Must-address notes neither cited by a verified claim nor named in unresolved."""
        verdicts = outcome.get("verdicts") or []
        cited = {r for n, c in enumerate(claims) if n < len(verdicts) and verdicts[n].get("verdict") in ("verified", "verified_inference")
                 for r in list(c.get("cites") or []) + list(c.get("cause_cites") or [])}
        named = {r for u in unresolved for r in re.findall(r"(?<![A-Za-z])N\d+(?!\d)", str(u))}
        return [m for m in must if m["ref"] not in cited and m["ref"] not in named]

    def _flags(self, account: str, view: Dict[str, Any], must: List[Dict[str, str]], claims: List[Dict[str, Any]],
               outcome: Dict[str, Any], unresolved: List[str], escalated: bool) -> List[Dict[str, Any]]:
        """What a person should look at in this paragraph, decided by the state, not by a model: material notes deferred instead
        of stated, figures taken from a note's free text, causes marked as inference or only related to a note, detail or basis
        caveats, what is unresolved, and that the stronger model had to repair it."""
        flags: List[Dict[str, Any]] = []
        verdicts = outcome.get("verdicts") or []
        cited = {r for c in claims for r in list(c.get("cites") or []) + list(c.get("cause_cites") or [])}
        for m in must:
            if m["ref"] not in cited:
                flags.append({"kind": "deferred_note", "ref": m["ref"], "why": m["why"], "text": m["text"]})
        for n, v in enumerate(verdicts):
            codes = {f.get("code") for f in v.get("findings") or []}
            if "NOTE_SOURCED_AMOUNT" in codes:
                flags.append({"kind": "note_amount", "claim": n, "text": str(v.get("text"))[:160]})
            if v.get("verdict") == "verified_inference" or "CAUSE_NOTE_RELATED_UNCHECKED" in codes:
                flags.append({"kind": "cause_not_proven", "claim": n, "text": str(v.get("text"))[:160]})
            if "DETAIL_NESTING_UNKNOWN" in codes:
                flags.append({"kind": "detail_caveat", "claim": n, "text": str(v.get("text"))[:160]})
        for p, perm in (view.get("permissions") or {}).items():
            if not perm.get("describe_total", True) or (view.get("checks") or {}).get(p, {}).get("total_vs_financials") == "mismatch":
                flags.append({"kind": "basis_caveat", "period": p, "text": "tab and Financials differ in this period"})
        deferred = {m["ref"] for m in must}
        for u in unresolved:
            if not any(r in deferred for r in re.findall(r"(?<![A-Za-z])N\d+(?!\d)", str(u))):
                flags.append({"kind": "unresolved", "text": str(u)[:200]})
        if view.get("components_not_shown"):
            flags.append({"kind": "evidence_cut", "text": f"{view['components_not_shown']} detail line(s) were not shown to the writer"})
        if escalated:
            flags.append({"kind": "escalated", "text": "the cheaper model's draft was rejected; the stronger model repaired it"})
        return flags

    def _required_refs(self, view: Dict[str, Any]) -> List[str]:
        """What a paragraph must state: the latest balance the deck uses, or, where tab and Financials disagree and nothing chose,
        the difference Python computed. Without it a draft of one harmless sentence would pass the verifier."""
        periods = view.get("periods") or []
        if not periods:
            return []
        latest = periods[-1]
        perm = (view.get("permissions") or {}).get(latest) or {}
        if perm.get("describe_total", True):
            refs = [t["ref"] for t in view.get("totals") or [] if t.get("period") == latest and t.get("basis") == "authority_adjusted"]
            return refs[:1]
        refs = [d["ref"] for d in view.get("derived") or [] if d.get("kind") == "difference" and latest in (d.get("periods") or [d.get("period")])]
        return refs[:1]

    def _writer_user(self, goal: str, account: str, view: Dict[str, Any], required: List[str], claims: List[Dict[str, Any]],
                     rejected: Dict[int, List[str]], unresolved: List[str]) -> str:
        # what does not change between a draft and its repairs comes first, so a provider's prompt cache can serve it again
        parts = [f"EVIDENCE for {account} (cite only these refs):\n{_compact(view)}"]
        open_issues = [i for i in self.issues.values() if not i.resolved and (not i.account or i.account == account)]
        if open_issues:
            parts.append("OPEN ISSUES (from the evidence):\n" + "\n".join(f"  [{i.severity}] {i.message}" + (f" -> {i.suggest}" if i.suggest else "")
                                                                   for i in open_issues))
        earlier = self._earlier_answer(account)
        if earlier:
            parts.append(earlier)
        must = getattr(self, "_must", None) or []
        if must:
            parts.append("MUST ADDRESS (state it citing the ref, or name the ref in 'unresolved' with the reason):\n"
                         + "\n".join(f"  {m['ref']} ({m['why']}): {m['text']}" for m in must))
        if required:
            parts.append("REQUIRED: a claim must cite " + ", ".join(required) + " (the latest balance the deck uses, or the difference that replaces it).")
        parts.append(f"REQUEST: {goal}")
        if not claims:
            parts.append(WRITER_PROTOCOL + "Your JSON object:")
            return "\n\n".join(parts)
        if not rejected and not unresolved:          # the first prompt of a seeded write: Python's sentences, the model adds
            parts.append("KEPT (Python wrote these from EVIDENCE and verified them):\n" + "\n".join(f"  [{n}] {c.get('text')}" for n, c in enumerate(claims)))
            parts.append(SEED_PROTOCOL + CLAIM_SCHEMA + "\nYour JSON object:")
            return "\n\n".join(parts)
        kept = [f"  [{n}] {c.get('text')}" for n, c in enumerate(claims) if n not in rejected]
        bad = [f"  [{n}] {_compact({k: claims[n].get(k) for k in ('kind', 'text', 'periods', 'cites', 'cause_cites', 'inference')})}\n"
               f"      objections: {' | '.join(rejected[n])[:700]}" for n in sorted(rejected) if n < len(claims)]
        if kept:
            parts.append("KEPT (verified; unchanged whatever you reply):\n" + "\n".join(kept))
        if bad:
            parts.append("REJECTED (rewrite each from EVIDENCE, or drop it and say why):\n" + "\n".join(bad))
        extra = [o for n, objs in rejected.items() if n >= len(claims) for o in objs]
        if extra:
            parts.append("ALSO: " + " | ".join(extra)[:600])
        if unresolved:
            parts.append("UNRESOLVED so far: " + "; ".join(str(u)[:160] for u in unresolved[:8]))
        parts.append(REPAIR_PROTOCOL + "Your JSON object:")
        return "\n\n".join(parts)

    def _earlier_answer(self, account: str) -> str:
        """The last answer of this session about this account, when the request may refer to it ("make that shorter"). Only that
        one: the prepared write is about one account, and the rest of the history would be resent for nothing."""
        history = self.memory.history if self.memory is not None else []
        last = history[-1] if history else None
        if not last or last.get("stale"):
            return ""
        mine = [c for c in last.get("claims") or [] if (self.session.resolve_account(str(c.get("account") or "")) or "") == account]
        if not mine:
            return ""
        return ("EARLIER IN THIS SESSION (the answer this request may refer to; its refs are in EVIDENCE):\n"
                f"  {str(last.get('request'))[:200]}\n  -> {str(last.get('text'))[:800]}\n  its claims: "
                + _truncate([{k: c.get(k) for k in ("kind", "periods", "cites", "cause_cites", "inference", "text") if c.get(k) not in (None, [], "")}
                             for c in mine], 2600))

    @staticmethod
    def _clean_claim(claim: Any) -> Optional[Dict[str, Any]]:
        if not isinstance(claim, dict) or not str(claim.get("text") or "").strip():
            return None
        return {k: claim[k] for k in ("kind", "text", "periods", "cites", "cause_cites", "inference", "cross_account") if k in claim}

    def _merge_fix(self, claims: List[Dict[str, Any]], rejected: Dict[int, List[str]], reply: Dict[str, Any],
                   unresolved: List[str]) -> Tuple[List[Dict[str, Any]], List[str]]:
        """Apply a repair to the REJECTED claims only. A fix aimed at a kept claim is ignored and said so."""
        out: List[Optional[Dict[str, Any]]] = list(claims)
        ignored: List[str] = []
        for item in reply.get("fix") or []:
            if not isinstance(item, dict):
                continue
            try:
                n = int(item.get("index"))
            except (TypeError, ValueError):
                continue
            if n not in rejected or n >= len(out):
                ignored.append(f"index {n} is not a rejected claim; kept claims are not changed")
                continue
            if item.get("drop"):
                unresolved.append(str(item.get("drop"))[:300])
                out[n] = None
                continue
            fixed = self._clean_claim(item.get("claim"))
            if fixed is not None:
                out[n] = fixed
        merged = [c for c in out if c is not None]
        merged += [c for c in (self._clean_claim(a) for a in reply.get("add") or []) if c is not None]
        unresolved.extend(str(u) for u in reply.get("unresolved") or [] if str(u) not in unresolved)
        return merged, ignored

    def _seed_claims(self, account: str, view: Dict[str, Any]) -> List[Dict[str, Any]]:
        """The sentences Python can write itself: the latest balance the deck uses and the movement into it, in house style, kept
        only if the verifier passes them. Nothing for a period in conflict (the writer states the difference)."""
        periods = view.get("periods") or []
        if not periods:
            return []
        latest = periods[-1]
        perm = (view.get("permissions") or {}).get(latest) or {}
        total = next((t for t in view.get("totals") or [] if t.get("period") == latest and t.get("basis") == "authority_adjusted"), None)
        if total is None or not perm.get("describe_total", True):
            return []
        chinese = self.language in ("Chi", "Chn")
        contract = self.session.contracts.get(account)
        flow = str(getattr(contract, "statement_type", "") or "").upper().startswith("IS")
        when, amount = _date_words(latest, chinese), _money(float(total.get("yuan") or 0.0), chinese)
        label = self._period_labels(account, view).get(latest) or when
        if chinese:
            text = f"{label}，发生额为{amount}。" if flow else f"截至{when}，余额为{amount}。"
        else:
            text = f"For {label}, the amount was {amount}." if flow else f"As at {when}, the balance was {amount}."
        seeds = [{"kind": "amount", "periods": [latest], "cites": [total["ref"]], "text": text}]
        move = next((d for d in view.get("derived") or [] if d.get("kind") == "movement" and d.get("valid", True) and d.get("value")
                     and list(d.get("periods") or [])[-1:] == [latest]), None)
        if move is not None and perm.get("compute_movement", True):
            prev, value = move["periods"][0], float(move["value"])
            if chinese:
                words = f"较{_date_words(prev, True)}{'增加' if value > 0 else '减少'}{_money(abs(value), True)}。"
            else:
                words = f"This was {'an increase' if value > 0 else 'a decrease'} of {_money(abs(value), False)} from {_date_words(prev, False)}."
            seeds.append({"kind": "movement", "periods": list(move["periods"]), "cites": [move["ref"]], "text": words})
        stamped = [dict(c, account=account, source_version=self.session.source_version) for c in seeds]
        outcome = cl.verify_candidate(stamped, self.session.contracts, default_account=account)
        return [c for c, v in zip(seeds, outcome["verdicts"]) if v.get("verdict") == "verified"]

    @staticmethod
    def _review_reasons(view: Dict[str, Any], claims: List[Dict[str, Any]]) -> List[str]:
        """Whether a verified paragraph has anything a reviewer must judge that the verifier cannot: a cause, notes to read,
        a basis caveat, part-year figures. Python decides; a plain balance-and-movement paragraph is not sent."""
        reasons = []
        if any(c.get("cause_cites") or c.get("inference") or c.get("kind") == "cause" for c in claims):
            reasons.append("states a cause")
        if view.get("notes"):
            reasons.append(f"{len(view['notes'])} note(s) in the evidence")
        perms = (view.get("permissions") or {}).values()
        if any(not p.get("describe_total", True) or p.get("detail_basis") not in (None, "tied") for p in perms):
            reasons.append("a basis or detail caveat")
        if any(row.get("total_vs_financials") == "mismatch" for row in (view.get("checks") or {}).values()):
            reasons.append("a period the deck replaced")
        if (view.get("annualization") or {}).get("annualized"):
            reasons.append("part-year figures")
        return reasons

    def _review_meaning(self, run: AgentRun, account: str, view: Dict[str, Any], claims: List[Dict[str, Any]],
                        usage: Dict[str, Any]) -> Dict[str, Any]:
        """The stronger model reads a verified paragraph for meaning. Returns {"status": "passed" | "objected" | "unavailable", ...}.
        Unavailable: the call could not be made or failed, or the reply is not the protocol (both lists present). A review that did
        not happen is never a passed review."""
        user = (f"EVIDENCE for {account}:\n{_compact(view)}\n\nPARAGRAPH (figures already verified by Python):\n"
                + "\n".join(f"  [{n}] {_compact({k: c.get(k) for k in ('kind', 'text', 'periods', 'cites', 'cause_cites', 'inference')})}"
                             for n, c in enumerate(claims))
                + "\n\n" + REVIEW_PROTOCOL + "\nYour JSON object:")
        stop_before = run.stop_reason
        called = self._call_model(run, REVIEW_SYSTEM, user, 1500, usage, role=R.ROLE_ESCALATION, purpose="review")
        if called is None:
            why = run.report.get("model_error") or (run.report.get("budget") or {}).get("message") or run.stop_reason
            if run.stop_reason not in (STOP_SESSION, STOP_TOKENS):
                run.stop_reason = stop_before                 # a provider failure ends the review, not the turn's own accounting
            return {"status": "unavailable", "why": f"the review call was not made or failed: {str(why)[:200]}"}
        turn, error = parse_turn(called[0])
        if turn is None or not isinstance(turn.get("objections"), list) or not isinstance(turn.get("missing"), list):
            return {"status": "unavailable", "why": "the reply was not the review protocol (objections and missing, both lists)"
                                                     + (f": {error}" if error else "")}
        objections: List[Dict[str, Any]] = []
        for item in turn["objections"]:
            if not isinstance(item, dict):
                continue
            try:
                n = int(item.get("index"))
            except (TypeError, ValueError):
                continue
            if 0 <= n < len(claims) and str(item.get("why") or "").strip():
                objections.append({"index": n, "claim_sha": _claim_sha(claims[n]), "why": str(item["why"])[:400]})
        missing = [{"ref": str(m.get("ref") or ""), "why": str(m.get("why") or "")[:300]} for m in turn["missing"]
                   if isinstance(m, dict) and (m.get("ref") or m.get("why"))]
        return {"status": "objected" if objections or missing else "passed", "objections": objections, "missing": missing}

    @staticmethod
    def _objection_lines(review: Dict[str, Any]) -> List[str]:
        return ([f"REVIEW: [{o['index']}] {o['why']}" for o in review.get("objections") or []]
                + [f"REVIEW_MISSING {m['ref']}: {m['why']}" for m in review.get("missing") or []])

    def _required_unstated(self, required: List[str], claims: List[Dict[str, Any]], outcome: Dict[str, Any],
                           view: Dict[str, Any]) -> List[str]:
        """The required refs no verified claim STATES: it cites the ref and its text carries that figure at its printed precision
        (a nil balance may be written as nil). Citing the balance under "needs continued attention" is not stating it."""
        values = {t["ref"]: t.get("yuan") for t in view.get("totals") or []}
        values.update({d["ref"]: d.get("value") for d in view.get("derived") or []})
        verdicts = outcome.get("verdicts") or []
        out = []
        for ref in required:
            value = values.get(ref)
            stated = False
            for n, c in enumerate(claims):
                if ref not in (c.get("cites") or []) or n >= len(verdicts) or verdicts[n].get("verdict") not in ("verified", "verified_inference"):
                    continue
                if value is None:
                    stated = True
                    break
                spans = cl.text_amounts(str(c.get("text") or ""))
                if any(abs(abs(sp["value"]) - abs(float(value))) <= sp["tolerance"] for sp in spans) or (
                        abs(float(value)) < 0.5 and any(w in str(c.get("text") or "") for w in ("为零", "为0", "无余额", "nil", "zero"))):
                    stated = True
                    break
            if not stated:
                out.append(ref)
        return out

    def _settle_review(self, run: AgentRun, account: str, view: Dict[str, Any], claims: List[Dict[str, Any]],
                       pre_review: List[Dict[str, Any]], verified: bool, review: Dict[str, Any], unresolved: List[str],
                       usage: Dict[str, Any], export: bool, record: Dict[str, Any]) -> bool:
        """After the writer answered an objection. Settled: every objected sentence is gone and nothing new was written (a drop
        removes what was objected to), or the reviewer reads the changed paragraph once more and passes it. Not settled: the
        reply changed nothing, an objected sentence is still there word for word, the change did not verify (the verified text
        is kept), or the second reading still objects. An unsettled objection ends the turn as semantic_blocked."""
        if not verified or cl_digest(claims) == cl_digest(pre_review):
            if not verified:
                T.call_tool(self.session, "validate_candidate", {"account": account, "claims": pre_review})
            why = "the repair did not verify; the verified text is kept" if not verified else "the reply changed nothing"
            return self._end_with_review(run, account, pre_review, dict(review, status="objected", settled=why), unresolved, record,
                                         STOP_SEMANTIC_BLOCKED)
        now = {_claim_sha(c) for c in claims}
        still = [o for o in review["objections"] if o["claim_sha"] in now]
        if still:
            return self._end_with_review(run, account, claims, dict(review, status="objected", objections=still,
                                                                   settled="an objected sentence is unchanged"),
                                         unresolved, record, STOP_SEMANTIC_BLOCKED)
        before = {_claim_sha(c) for c in pre_review}
        if not review.get("missing") and all(_claim_sha(c) in before for c in claims):
            return self._end_with_review(run, account, claims, dict(review, status="passed_after_repair", settled="the objected sentences were dropped"),
                                         unresolved, record, STOP_DONE, export=export)
        again = self._review_meaning(run, account, view, claims, usage)
        if again["status"] == "passed":
            return self._end_with_review(run, account, claims, dict(review, status="passed_after_repair", settled="read again and passed",
                                                                   second=again), unresolved, record, STOP_DONE, export=export)
        if again["status"] == "unavailable":
            return self._end_with_review(run, account, claims, dict(review, status="unavailable", settled=again["why"], second=again),
                                         unresolved, record, STOP_REVIEW_PENDING if run.stop_reason in ("", None) else run.stop_reason)
        return self._end_with_review(run, account, claims, dict(again, status="objected", settled="read again and still objected",
                                                               first=review), unresolved, record, STOP_SEMANTIC_BLOCKED)

    def _end_with_review(self, run: AgentRun, account: str, claims: List[Dict[str, Any]], review: Optional[Dict[str, Any]],
                         unresolved: List[str], record: Dict[str, Any], stop: str, export: bool = False) -> bool:
        """End the prepared write. The review's state is written onto the candidate, where export reads it; an open objection
        or a review that did not happen is listed as unresolved and the turn is not a finished paragraph."""
        candidate = self.session.candidates.get(account)
        if candidate is not None and review is not None:
            candidate.review = {k: v for k, v in review.items()}
        if candidate is not None and getattr(self, "_view", None) is not None:
            candidate.flags = self._flags(account, self._view, self._must, candidate.claims, candidate.outcome, unresolved,
                                          escalated=bool(self.escalations))
            run.report["flags"] = list(candidate.flags)
        if review is not None:
            run.report["review"] = dict(review, outcome=review.get("status"))
            if review.get("status") in ("objected", "unavailable"):
                lines = self._objection_lines(review) if review.get("status") == "objected" else [f"REVIEW_PENDING: {review.get('why') or review.get('settled')}"]
                unresolved.extend(x for x in lines if x not in unresolved)
        if stop == STOP_DONE:
            accepted, _feedback, answer = self._gate_done({"account": account}, account, export)
            record["accepted"] = accepted
            if accepted:
                run.answer, run.stop_reason = dict(answer or {}, unresolved=list(unresolved)), STOP_DONE
                run.steps.append(record)
                return True
            stop = STOP_REJECTED
        # not finished: the draft is shown with what stands against it
        if candidate is not None:
            run.answer = {"account": account, "text": cl.render_claims(candidate.claims), "claims": candidate.claims,
                          "verdicts": candidate.outcome.get("verdicts"), "unresolved": list(unresolved),
                          "source_version": self.session.source_version}
        run.stop_reason = stop
        record["review_end"] = stop
        run.steps.append(record)
        return True

    def _prepared_write(self, run: AgentRun, goal: str, account: str, export: bool, usage: Dict[str, Any], started: float) -> bool:
        """The paragraph for one known account, written once from evidence the controller read, then repaired locally.
        True: the turn is finished (accepted, or stopped by a budget). False: continue in the tool loop."""
        view = self._prepare_evidence(account)
        if view is None:
            run.report["prepared_fallback"] = "the account's evidence could not be read"
            return False
        required = self._required_refs(view)
        self._view, self._must = view, self._must_address(view)
        system = writer_system_prompt(self.language)
        router = self.router
        claims: List[Dict[str, Any]] = []
        rejected: Dict[int, List[str]] = {}
        unresolved: List[str] = []
        bad_replies, last_failed = 0, None
        run.report["prepared"] = {"required": required, "evidence_chars": len(_compact(view))}
        if self.seed:
            claims = self._seed_claims(account, view)
            run.report["prepared"]["seeded"] = len(claims)
        can_review = self.review and self.router is not None and R.ROLE_ESCALATION in self.router.roles
        pre_review: Optional[List[Dict[str, Any]]] = None          # the verified paragraph the review objected to
        review: Optional[Dict[str, Any]] = None                     # this turn's review: status, reasons, objections
        while len(run.steps) < self.budget.max_steps:
            if time.perf_counter() - started > self.budget.max_seconds:
                run.stop_reason = STOP_TIME
                return True
            step = len(run.steps) + 1
            called = self._call_model(run, system, self._writer_user(goal, account, view, required, claims, rejected, unresolved),
                                      self.WRITE_ALLOWANCE, usage)
            if called is None:
                return True
            text, role = called
            turn, error = parse_turn(text)
            record: Dict[str, Any] = {"step": step, "action": "repair" if claims else "write", "prepared": True}
            if role:
                record["role"] = role
            codes: List[str] = []
            if turn is not None and ("need_evidence" in turn or "tool" in turn):
                record["fallback"] = str(turn.get("need_evidence") or f"tool {turn.get('tool')}")[:300]
                if "tool" in turn:
                    self._pending_reply = (text, role)          # the loop runs this tool call as its first step
                else:
                    run.steps.append(record)
                self._note("CONTROLLER: the evidence for " + account + " was read for you (above). "
                           + (f"The writer said: {record['fallback']}. " if "need_evidence" in turn else "")
                           + (f"Last objections: {' | '.join(o for objs in rejected.values() for o in objs)[:600]}. " if rejected else "")
                           + "Investigate with the tools, then validate_candidate and finish with done.")
                run.report["prepared_fallback"] = record["fallback"]
                return False
            if turn is None or not ("draft" in turn or "fix" in turn or "add" in turn):
                record["error"] = error or "neither a draft nor a fix"
                codes = ["NOT_JSON"] if turn is None else ["BAD_ACTION"]
                if router is not None:
                    router.mark_last_reply_unused()
                if review is not None and review["status"] == "objected" and pre_review is not None:
                    # the repair the review asked for came back unusable: the verified paragraph is kept, its objection stands
                    return self._settle_review(run, account, view, pre_review, pre_review, False, review, unresolved, usage, export, record)
                bad_replies += 1
                if bad_replies >= 2:
                    run.steps.append(record)
                    run.report["prepared_fallback"] = "two replies in a row were not a draft"
                    return False
            else:
                bad_replies = 0
                if "draft" in turn:
                    body = turn.get("draft") if isinstance(turn.get("draft"), dict) else {}
                    new = [c for c in (self._clean_claim(x) for x in body.get("claims") or []) if c is not None]
                    unresolved.extend(str(u) for u in body.get("unresolved") or [] if str(u) not in unresolved)
                    ignored: List[str] = []
                    if claims and new:            # a whole new draft in answer to a repair: the kept claims still stand
                        new = [c for n, c in enumerate(claims) if n not in rejected] + new
                else:
                    new, ignored = self._merge_fix(claims, rejected, turn, unresolved)
                if ignored:
                    record["ignored"] = ignored
                claims = new
                if not claims:
                    codes = ["EMPTY_ANSWER"]
                    rejected = {}
                else:
                    result = T.call_tool(self.session, "validate_candidate", {"account": account, "claims": claims})
                    self._absorb("validate_candidate", {"account": account, "claims": claims}, result)
                    outcome = (result.data or {}).get("outcome") or {}
                    missing = self._required_unstated(required, claims, outcome, view)
                    unaddressed = self._unaddressed(self._must, claims, outcome, unresolved)
                    record.update(claims=len(claims), verified=outcome.get("verified"), failed=outcome.get("failed"),
                                  unverified=outcome.get("unverified"))
                    if review is not None and review["status"] == "objected" and pre_review is not None:
                        # this draft answers the reviewer: settled only by a change the reviewer accepts, never by a repeat
                        return self._settle_review(run, account, view, claims, pre_review, result.ok and not missing and not unaddressed, review,
                                                   unresolved, usage, export, record)
                    if result.ok and not missing and not unaddressed:
                        if can_review:
                            reasons = self._review_reasons(view, claims)
                            if not reasons:
                                review = {"status": "not_needed", "reasons": []}
                            else:
                                review = dict(self._review_meaning(run, account, view, claims, usage), reasons=reasons)
                                if review["status"] == "objected":
                                    pre_review = list(claims)
                                    review["text_reviewed"] = cl.render_claims(claims)      # what the reviewer read, kept for the record
                                    rejected = {o["index"]: [f"REVIEW: {o['why']}"] for o in review["objections"]}
                                    if review["missing"]:
                                        rejected.setdefault(len(claims), []).extend(
                                            f"REVIEW_MISSING {m['ref']}: {m['why']}; add a claim that states it" for m in review["missing"])
                                    run.report["review"] = {k: v for k, v in review.items()}
                                    record.update(review=len(review["objections"]) + len(review["missing"]))
                                    run.steps.append(record)
                                    continue
                                if review["status"] == "unavailable":
                                    return self._end_with_review(run, account, claims, review, unresolved, record,
                                                                 STOP_REVIEW_PENDING if run.stop_reason in ("", None) else run.stop_reason)
                        return self._end_with_review(run, account, claims, review, unresolved, record, STOP_DONE, export=export)
                    else:
                        rejected = {}
                        for n, v in enumerate(outcome.get("verdicts") or []):
                            if v.get("verdict") not in ("verified", "verified_inference"):
                                bad = [f for f in v.get("findings") or [] if f.get("severity") in ("failed", "review")]
                                rejected[n] = [f"{f['code']}: {f.get('message', '')}" for f in bad] or ["not verified"]
                        if missing:
                            rejected[len(claims)] = [f"REQUIRED_NOT_STATED: no verified claim cites {', '.join(missing)} AND writes its figure; "
                                                     "add one ('add')"]
                        if unaddressed:
                            rejected.setdefault(len(claims), []).extend(
                                f"NOTE_UNADDRESSED: {m['ref']} ({m['why']}) is in MUST ADDRESS; state it citing {m['ref']}, or name {m['ref']} "
                                "in 'unresolved' with the reason" for m in unaddressed)
                        codes = sorted(outcome.get("by_code") or {}) + (["REQUIRED_NOT_STATED"] if missing else []) + \
                            (["NOTE_UNADDRESSED"] if unaddressed else [])
                        record["feedback"] = " | ".join(o for objs in rejected.values() for o in objs)[:900]
                if codes:
                    self.rejections += 1
                    if router is not None:
                        router.note_rejection()
                    if self.rejections >= self.budget.max_rejections:
                        self._step_log.append((step, codes))
                        run.stop_reason = STOP_REJECTED
                        run.steps.append(record)
                        return True
            # the same failure twice from the same draft: nothing new is happening; the tool loop can investigate
            failed_now = (tuple(sorted(codes)), cl_digest(claims))
            if codes and failed_now == last_failed:
                run.steps.append(record)
                run.report["prepared_fallback"] = "the same draft failed the same way twice"
                self._note("CONTROLLER: the draft below failed the same way twice: " + " | ".join(
                    o for objs in rejected.values() for o in objs)[:700] + ". Investigate with the tools, then validate_candidate and finish with done.")
                return False
            last_failed = failed_now
            if router is not None and codes:
                classes = R.effective_classes(codes, self._context(codes))
                self._step_log.append((step, codes))
                self._step_classes[step] = classes
                decision = router.observe(R.StepEvent(step=step, codes=codes, progress=False, no_progress=1, classes=classes))
                if decision.escalate:
                    handoff = (f"A prepared write of {account}: the evidence and the draft are in the next message; the repair moves to you. "
                               f"Failures so far: {', '.join(codes)}.")
                    info = router.escalate(decision, handoff, step)
                    self.escalations.append({**info, "handoff": handoff})
                    record["escalated"] = {"to": info["to"], "model": info["model"], "reason": info["reason"]}
                    if router.policy.fresh_rejections:
                        self.rejections = 0
            run.steps.append(record)
        run.stop_reason = STOP_STEPS
        return True

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
            STOP_SEMANTIC_BLOCKED: "the paragraph verifies, but the reviewer's objection to its meaning stands (listed under unresolved)",
            STOP_REVIEW_PENDING: "a review of the paragraph's meaning was due and did not happen",
        }.get(run.stop_reason, run.stop_reason)
        return {"why_stopped": why, "known": known, "gaps": gaps, "next_steps": next_steps, "accounts_consulted": list(self.consulted)}

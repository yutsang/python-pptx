"""Which model takes the next step, and what every step costs against ONE budget.

This is not a second agent. There is one controller and one loop (controller.py); this module decides, in code, which
configured model answers the next call, and keeps the books. Python still computes and verifies: the second model is
not asked to re-review the first one's work, it is handed a compact task state and continues the same task.

What a failure means for the next move (CLASSES) is settled here, per code, because it is the line between spending
more and doing something else:

    REWRITE   the sentence is wrong or cannot be bound, and the evidence already in the conversation settles it:
              reword it, recompute the figure, name the line, state the sign, move the statement to 'unresolved'.
              A sentence that cannot be verified is not a question for the user; it is a sentence to change.
    EVIDENCE  what the sentence needs is not in the conversation yet: fetch it with a tool; if the tools return nothing,
              say so under 'unresolved'. A stronger model cannot supply evidence the workbook does not hold.
    HUMAN     the sources conflict or are ambiguous and no rule chooses: only a person can. Named under 'unresolved' or
              asked, never retried.
    REFRESH   the workbook moved: read it again. No model call is needed.
    FORMAT    the reply was not a usable action (not JSON, an action this mode does not have, a step the gate refused).
    SYSTEM    a tool crashed or an argument was malformed: the harness's problem, not evidence of the model's limits.

Only REWRITE and FORMAT failures are "the model did not manage it with what it had". Two of them in a row, with nothing
the gate is waiting on from a person or from a tool, is the one trigger for the escalation model. EVIDENCE and HUMAN
never are: a different model does not change what the workbook says.

The books (SharedBudget) count every call of every role, in every turn, including calls that failed and replies that
were thrown away, in two columns that are never added together: what the provider reported and what this code
estimated. The cap is applied to the provider's number where it reported one and to the estimate where it did not.
What no code here can see is retries inside the provider SDK; they are named as unmeasured, not assumed zero.
"""
from __future__ import annotations

import hashlib
import json
import re
import time
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, List, Optional, Tuple

REWRITE, EVIDENCE, HUMAN, REFRESH, FORMAT, SYSTEM, INFO = "rewrite", "evidence", "human", "refresh", "format", "system", "info"
#: not a code's own class: what an EVIDENCE code becomes once the tools were asked and the workbook holds nothing (effective_classes)
ABSENT = "absent"
CAPABILITY_CLASSES = (REWRITE, FORMAT)

ROLE_PRIMARY, ROLE_ESCALATION = "primary", "escalation"

#: every code a claim verdict or a tool result can carry, and what it asks of the next step
CLASS: Dict[str, str] = {
    # -- a sentence the evidence in hand can repair --------------------------------------------------------------------
    "WRONG_PERIOD": REWRITE, "WRONG_LINE": REWRITE, "WRONG_OPERATION": REWRITE, "WRONG_BASIS": REWRITE,
    "WRONG_ACCOUNT": REWRITE, "NO_CONTRACT": REWRITE, "SIGN_ERROR": REWRITE, "UNIT_ERROR": REWRITE,
    "AMOUNT_UNMATCHED": REWRITE, "AMOUNT_ELSEWHERE": REWRITE, "AMOUNT_UNCITED": REWRITE,
    "PERCENT_UNMATCHED": REWRITE, "PERCENT_ZERO_DENOMINATOR": REWRITE, "PERCENT_SIGN_UNSTATED": REWRITE, "SIGN_UNCLEAR": REWRITE, "REQUIRED_NOT_STATED": REWRITE,
    "NOTE_UNADDRESSED": REWRITE, "PERCENT_OVER_SIGN_CHANGE": REWRITE, "COMPOSITION_BASIS": REWRITE, "INTERNAL_TERMS_IN_TEXT": FORMAT,
    "NOTE_SOURCED_AMOUNT": INFO,
    "LINE_UNBOUND": REWRITE, "LINE_NOT_CITED": REWRITE, "PERIOD_UNBOUND": REWRITE, "PERIODS_UNDECLARED": REWRITE,
    "MOVEMENT_END_UNSTATED": REWRITE, "NAME_NOT_IN_EVIDENCE": REWRITE,
    "REST_INCOMPLETE": REWRITE, "NOTE_FIGURE_ROUNDED": REWRITE, "GAIN_LOSS_REVERSED": REWRITE,
    "AMBIGUOUS_SUBJECT": REWRITE, "TERM_NOT_IN_NOTE": REWRITE,
    "DIRECTION_UNBOUND": REWRITE, "DIRECTION_UNSUPPORTED": REWRITE,
    # the permissions say what may be stated; the sentence states more. Say only what is permitted, and name the rest
    "MOVEMENT_NOT_PERMITTED": REWRITE, "DETAIL_NOT_PERMITTED": REWRITE, "NIL_NOT_PERMITTED": REWRITE, "TOTAL_NOT_DESCRIBABLE": REWRITE,
    "DERIVED_NOT_VALID": REWRITE, "ANNUALIZATION_UNSUPPORTED": REWRITE, "ANNUALIZATION_UNLABELLED": REWRITE,
    "ANNUALIZED_AS_FORECAST": REWRITE, "CAUSE_CONTRADICTS_NOTE": REWRITE, "CAUSE_NOTE_UNRELATED": REWRITE,
    # a statement about another account cannot be bound to a figure here: drop it or move it to 'unresolved'
    "CROSS_ACCOUNT_UNBOUND": REWRITE,
    # -- evidence not yet in the conversation ------------------------------------------------------------------------------
    "UNSEEN_REF": EVIDENCE, "UNKNOWN_REF": EVIDENCE, "CAUSE_UNSUPPORTED": EVIDENCE, "UNKNOWN_ACCOUNT": EVIDENCE,
    "NO_SUCH_SHEET": EVIDENCE, "NOT_FOUND": EVIDENCE, "NOT_COMPUTABLE": EVIDENCE,
    # -- only a person decides -------------------------------------------------------------------------------------------------
    "HIDDEN_SHEET": HUMAN, "FINANCIALS_PAIRING_SUSPECT": HUMAN, "OPEN_BLOCKING_ISSUE": HUMAN, "CANNOT_REFRESH": HUMAN,
    "DELIVERY_BLOCKED": HUMAN, "NO_BASELINE": HUMAN, "NO_MANIFEST": HUMAN, "NO_SUCH_FILE": HUMAN,
    # -- the sources moved ---------------------------------------------------------------------------------------------------------
    "STALE_SOURCE": REFRESH,
    # -- the reply or the step -----------------------------------------------------------------------------------------------------
    "NOT_JSON": FORMAT, "BAD_ACTION": FORMAT, "ASK_NOT_ALLOWED": FORMAT, "EMPTY_ANSWER": FORMAT, "BAD_ANSWER": FORMAT,
    "NO_VALIDATED_CANDIDATE": FORMAT, "EXPORT_MISSING": FORMAT, "NO_PROGRESS": FORMAT,
    "BAD_ARGS": FORMAT, "UNKNOWN_TOOL": FORMAT, "NOT_VALIDATED": FORMAT, "NO_CANDIDATE": FORMAT, "UNTRACED_NUMBERS": REWRITE,
    # the paragraph does not fit its slot: shorten it. A candidate that failed its checks: the objections say what to change
    "LAYOUT_PROBLEM": REWRITE, "VALIDATION_FAILED": REWRITE,
    "TOOL_EXCEPTION": SYSTEM, "NO_TEMPLATE": SYSTEM,
    # -- recorded, nothing to do ------------------------------------------------------------------------------------------------
    "DETAIL_NESTING_UNKNOWN": INFO, "CAUSE_NOTE_RELATED_UNCHECKED": INFO, "CAUSE_INFERENCE": INFO, "CAUSE_FROM_NOTE": INFO,
}

#: codes that wrap others: the classes are read from what they carry
WRAPPERS = ("VALIDATION_FAILED",)

#: what each class tells the model to do, appended to the gate's objection so the next step goes the right way
NEXT_STEP = {
    REWRITE: "Settle it with the evidence you already have: change the sentence, recompute the figure with compute_analysis, or move the "
             "statement to 'unresolved'. Do not ask the user.",
    EVIDENCE: "The evidence is not in this conversation yet: fetch it with a tool. If the tools return nothing, say that under 'unresolved' "
              "instead of stating it.",
    HUMAN: "No rule chooses between the sources: name it under 'unresolved' (or ask the user if the tools allow it). Do not pick one.",
    REFRESH: "The workbook was saved since it was read: call refresh_sources, then start from the new evidence.",
    FORMAT: "Reply with exactly one JSON object of the kind the protocol lists.",
    SYSTEM: "A tool failed on its own: retry once with the same arguments, then report it under 'unresolved'.",
    ABSENT: "The tools were asked and the workbook holds no such evidence: drop the statement or name it under 'unresolved'. "
            "Do not ask again and do not state it.",
}

_ISSUE_PREFIXES = (("conflict:", HUMAN), ("source:", HUMAN), ("stale", REFRESH), ("replaced:", EVIDENCE),
                   ("excluded:", INFO), ("unmatched:", INFO), ("annualized:", INFO))


def classify(code: str) -> str:
    """The class of a verifier or tool code, or of an issue id. A code nobody classified is treated as SYSTEM: not evidence of the
    model's limits, so never a reason to escalate. The test over every emitted code fails first, so this is a floor, not a plan."""
    code = str(code or "")
    if code in CLASS:
        return CLASS[code]
    for prefix, kind in _ISSUE_PREFIXES:
        if code.startswith(prefix):
            return kind
    return SYSTEM


def is_classified(code: str) -> bool:
    """Whether a person decided what this code means for the next step (as opposed to the SYSTEM floor)."""
    return str(code) in CLASS or any(str(code).startswith(prefix) for prefix, _ in _ISSUE_PREFIXES)


def classify_all(codes: List[str]) -> Dict[str, List[str]]:
    out: Dict[str, List[str]] = {}
    for code in codes:
        out.setdefault(classify(code), []).append(code)
    return out


def next_step_hint(codes: List[str]) -> str:
    """One line per class present, in the order that matters (a person's decision and missing evidence before wording)."""
    present = classify_all(codes)
    lines = [f"[{kind}] " + NEXT_STEP[kind] for kind in (HUMAN, EVIDENCE, REFRESH, REWRITE, FORMAT, SYSTEM) if present.get(kind)]
    return " NEXT: " + " ".join(lines) if lines else ""


def next_step_hint_for(classes: Dict[str, str]) -> str:
    """The same, from classes already settled in context (effective_classes), so 'the tools found nothing' reads as such."""
    present = sorted(set(classes.values()), key=lambda k: (HUMAN, ABSENT, EVIDENCE, REFRESH, REWRITE, FORMAT, SYSTEM, INFO).index(k)
                     if k in (HUMAN, ABSENT, EVIDENCE, REFRESH, REWRITE, FORMAT, SYSTEM, INFO) else 99)
    lines = [f"[{kind}] " + NEXT_STEP[kind] for kind in present if kind in NEXT_STEP]
    return " NEXT: " + " ".join(lines) if lines else ""


@dataclass
class StepContext:
    """What the controller knows about a step beyond its codes: which evidence was read, whether what was cited exists at all,
    what the searches found, and how often a code came back after the hint that should have settled it. One code can mean
    "go and read it" in one step and "the workbook does not have it" in the next; the code alone cannot tell them apart."""
    fetched_accounts: set = field(default_factory=set)        # accounts whose evidence the model has been shown
    cited_refs_exist: Optional[bool] = None                  # UNSEEN/UNKNOWN_REF: do the refs exist in any account (None: not asked)
    cited_ref_accounts: set = field(default_factory=set)      # the accounts those refs belong to
    cited_refs_hidden_by_paging: bool = False                 # they exist, but in components the evidence view left out
    cause_accounts: set = field(default_factory=set)          # accounts of claims that assert a cause with no note
    notes_found: Dict[str, int] = field(default_factory=dict) # account -> notes the evidence or a search returned
    searched: set = field(default_factory=set)               # accounts search_notes was run for
    repeated_tool_failure: bool = False                       # the same tool and arguments failed the same way before
    repeats: Dict[str, int] = field(default_factory=dict)     # code -> times it came back after its hint, this turn


def effective_classes(codes: List[str], ctx: Optional[StepContext] = None, *, give_up_after: int = 2) -> Dict[str, str]:
    """Each code's class in context. The rules the routing follows:

      evidence that exists but was not read  -> EVIDENCE (investigate first; no stronger model)
      evidence the tools were asked for and the workbook does not hold -> ABSENT (unresolved; never escalated)
      a ref that exists nowhere              -> REWRITE (the model invented it: that is a capability failure)
      evidence read, and the same code still back after its hint `give_up_after` times -> REWRITE (it is not using what it has)
      a source conflict or an authority no rule settles -> HUMAN (asked or named, never retried, never escalated)
    """
    ctx = ctx or StepContext()
    out: Dict[str, str] = {}
    for code in codes:
        kind = classify(code)
        if code in ("UNSEEN_REF", "UNKNOWN_REF") and ctx.cited_refs_exist is not None:
            if not ctx.cited_refs_exist:
                kind = REWRITE
            elif ctx.cited_ref_accounts and ctx.cited_ref_accounts <= ctx.fetched_accounts and not ctx.cited_refs_hidden_by_paging:
                kind = REWRITE if ctx.repeats.get(code, 0) >= give_up_after else EVIDENCE
            else:
                kind = EVIDENCE
        elif code == "CAUSE_UNSUPPORTED":
            accounts = ctx.cause_accounts or set()
            looked = bool(accounts) and all(a in ctx.searched and not ctx.notes_found.get(a) for a in accounts)
            kind = ABSENT if looked else EVIDENCE
        elif kind == EVIDENCE and code in ("NOT_FOUND", "NO_SUCH_SHEET", "NOT_COMPUTABLE") and ctx.repeated_tool_failure:
            kind = ABSENT
        elif kind == EVIDENCE and ctx.repeats.get(code, 0) >= give_up_after and ctx.fetched_accounts:
            kind = REWRITE
        out[code] = kind
    return out


def is_capability_failure(codes: List[str], classes: Optional[Dict[str, str]] = None) -> bool:
    """A step the model did not manage with what it had: only REWRITE / FORMAT codes, and nothing in it that waits on
    a person, on a tool, on a refresh or on the harness. `classes` (effective_classes) overrides each code's own class."""
    kinds = {(classes or {}).get(c) or classify(c) for c in codes}
    kinds.discard(INFO)
    if not kinds:
        return False
    return all(kind in CAPABILITY_CLASSES for kind in kinds)


_TASK_PATTERNS = {
    "cross_account": re.compile(r"(跨科目|勾稽|關聯|关联|對應|对应|cross[- ]?account|tie[- ]?out|related accounts?)", re.I),
    "adjustment_bridge": re.compile(r"(調整|调整|差異|差异|對賬|对账|bridge|adjust|reconcil)", re.I),
}


def task_kinds(request: str) -> List[str]:
    """The kinds of task a request names, for the configured 'start on the escalation model' list (agent.escalate_tasks).
    Word patterns only; a request that names none is an ordinary task."""
    return [kind for kind, rx in _TASK_PATTERNS.items() if rx.search(str(request or ""))]


# ---------------------------------------------------------------------------------------------------------------------------
# The books
# ---------------------------------------------------------------------------------------------------------------------------

class BudgetExhausted(RuntimeError):
    """Raised before a call that would break a cap. `kind` is calls | tokens | seconds | cost."""

    def __init__(self, kind: str, message: str):
        super().__init__(message)
        self.kind = kind


@dataclass
class Limits:
    max_calls: Optional[int] = None
    max_tokens: Optional[int] = None           # charged tokens over the whole session, every role
    max_seconds: Optional[float] = None
    max_cost: Optional[float] = None           # yuan, from the prices in config; needs a price for every role in use
    max_escalations: Optional[int] = None      # over the whole session

    def to_dict(self) -> Dict[str, Any]:
        return dict(self.__dict__)


_ADDRESS = re.compile(r"https?://\S+")


def _int(value: Any) -> int:
    try:
        return int(value or 0)
    except (TypeError, ValueError):
        return 0


@dataclass
class CallRecord:
    n: int
    role: str
    turn: int
    model: str
    provider: str
    ok: bool
    #: "provider" (the provider reported its usage), "estimate" (it did not), "response_cache" (never sent), "failed"
    usage_source: str
    prompt_est: int = 0
    completion_est: int = 0
    prompt_provider: Optional[int] = None
    completion_provider: Optional[int] = None
    cached_prompt: Optional[int] = None
    reasoning: Optional[int] = None
    saved_prompt: Optional[int] = None
    saved_completion: Optional[int] = None
    seconds: float = 0.0
    error: str = ""
    thinking: str = ""
    cost: Optional[float] = None
    reply_used: bool = True
    #: the model id the provider says answered (None when it does not say); `model` is the configured id
    model_reported: Optional[str] = None
    #: the output this call was allowed (the role's max_output or the step's), reserved against the cap before sending
    max_output: int = 0
    #: why this call was made again: "" (a new step), "failed" (the call before it failed), "discarded" (its reply was thrown away),
    #: "rejected" (the gate sent the answer back)
    after: str = ""
    #: the provider reported more reasoning tokens than completion tokens, so they cannot be part of completion: added to `charged`
    reasoning_separate: bool = False
    #: the cached part of the input was priced at the configured price_cached_in (False: at the full price_in, or nothing cached)
    cache_discounted: bool = False
    #: why this call was made when it is not the turn's own model at work: "review" (the stronger model reading a verified paragraph
    #: for meaning). "" for an ordinary step, a repair or an escalated step.
    purpose: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return dict(self.__dict__)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "CallRecord":
        return cls(**{k: v for k, v in data.items() if k in cls.__dataclass_fields__})

    @property
    def charged(self) -> int:
        """The number a cap is applied to: the provider's where it reported one, the estimate where it did not. A call
        served from the response cache was never sent and costs nothing; a call that failed is charged its prompt estimate
        (the provider may have read it) and nothing for the reply it never produced."""
        if self.usage_source == "response_cache":
            return 0
        if self.usage_source == "provider":
            # OpenAI-compatible usage counts reasoning INSIDE completion_tokens (completion_tokens_details.reasoning_tokens is a
            # part of it), so it is shown, never added; only a provider that reports more reasoning than completion is counting
            # it outside, and then it is added once
            return _int(self.prompt_provider) + _int(self.completion_provider) + (_int(self.reasoning) if self.reasoning_separate else 0)
        if self.usage_source == "failed":
            return self.prompt_est
        return self.prompt_est + self.completion_est


class SharedBudget:
    """One set of books for every model, every turn and every attempt of a working session."""

    def __init__(self, limits: Optional[Limits] = None):
        self.limits = limits or Limits()
        self.records: List[CallRecord] = []
        self.escalations: List[Dict[str, Any]] = []
        self.started = time.perf_counter()
        self._seconds_before = 0.0

    # -- reading ------------------------------------------------------------------------------------------------------------------
    @property
    def calls(self) -> int:
        return len(self.records)

    @property
    def charged_tokens(self) -> int:
        return sum(r.charged for r in self.records)

    @property
    def seconds(self) -> float:
        return self._seconds_before + (time.perf_counter() - self.started)

    @property
    def cost(self) -> Optional[float]:
        priced = [r.cost for r in self.records if r.cost is not None]
        return round(sum(priced), 4) if priced else None

    def summary(self) -> Dict[str, Any]:
        """The books, with the three kinds of number kept apart. `provider_reported` is what the provider said it
        processed; `estimated` is this code's estimate for EVERY call (so the two can be compared); `response_cache` is what
        the app's own cache answered without sending. `charged_tokens` is the one the cap sees. Nothing is summed across
        kinds except that one field, and it says how it was made."""
        sent = [r for r in self.records if r.usage_source != "response_cache"]
        reported = [r for r in sent if r.usage_source == "provider"]
        by_role: Dict[str, Dict[str, Any]] = {}
        for r in self.records:
            row = by_role.setdefault(r.role, {"calls": 0, "failed": 0, "model": r.model, "provider": r.provider, "charged_tokens": 0,
                                              "prompt_provider": 0, "completion_provider": 0, "prompt_est": 0, "completion_est": 0,
                                              "cached_prompt": 0, "reasoning": 0, "thinking": r.thinking, "models_reported": [],
                                              "retries": 0, "seconds": 0.0, "cost": None})
            if r.model_reported and r.model_reported not in row["models_reported"]:
                row["models_reported"].append(r.model_reported)
            row["retries"] += 1 if r.after else 0
            row["seconds"] = round(row["seconds"] + float(r.seconds or 0.0), 2)
            row["calls"] += 1
            row["failed"] += 0 if r.ok else 1
            row["charged_tokens"] += r.charged
            row["prompt_provider"] += _int(r.prompt_provider)
            row["completion_provider"] += _int(r.completion_provider)
            row["prompt_est"] += r.prompt_est
            row["completion_est"] += r.completion_est
            row["cached_prompt"] += _int(r.cached_prompt)
            row["reasoning"] += _int(r.reasoning)
            if r.cost is not None:
                row["cost"] = round((row["cost"] or 0.0) + r.cost, 6)
        return {
            "calls": self.calls, "failed_calls": sum(not r.ok for r in self.records),
            "replies_discarded": sum(r.ok and not r.reply_used for r in self.records),
            "escalations": len(self.escalations), "seconds": round(self.seconds, 2),
            "charged_tokens": self.charged_tokens,
            "charged_basis": "provider-reported where the provider reported, otherwise estimated (failed calls: prompt estimate only)",
            "provider_reported": {"calls": len(reported), "prompt": sum(_int(r.prompt_provider) for r in reported),
                                  "completion": sum(_int(r.completion_provider) for r in reported),
                                  "cached_prompt": sum(_int(r.cached_prompt) for r in reported),
                                  "reasoning": sum(_int(r.reasoning) for r in reported)},
            "estimated": {"calls": len(sent), "prompt": sum(r.prompt_est for r in sent), "completion": sum(r.completion_est for r in sent)},
            "calls_without_provider_usage": sum(1 for r in sent if r.usage_source in ("estimate", "failed")),
            "response_cache": {"hits": sum(r.usage_source == "response_cache" for r in self.records),
                               "saved_prompt": sum(_int(r.saved_prompt) for r in self.records),
                               "saved_completion": sum(_int(r.saved_completion) for r in self.records)},
            "cost": self.cost, "cost_basis": "prices from config (yuan per million tokens), provider-reported tokens where reported",
            "price_source": "config.yml available_models price_in / price_out, as written there; not checked against a provider bill",
            "cached_input_pricing": self._cached_pricing(reported),
            "reasoning_basis": ("reasoning tokens are shown, not added: OpenAI-compatible usage counts them inside completion. "
                                f"{sum(r.reasoning_separate for r in self.records)} call(s) reported more reasoning than completion and had it added once"),
            "retries": {"after_failed_call": sum(r.after == "failed" for r in self.records),
                        "after_discarded_reply": sum(r.after == "discarded" for r in self.records),
                        "after_rejection": sum(r.after == "rejected" for r in self.records)},
            "reservation": ("before each call the expected input (estimated) and the call's whole output allowance are reserved against "
                            "the caps; the provider's own count of the input can differ from the estimate, so a token cap is not a hard limit"),
            "unmeasured": ["retries made inside the provider SDK or gateway are not visible here"],
            "by_role": by_role, "limits": self.limits.to_dict(),
            "by_purpose": {p: {"calls": sum(1 for r in self.records if r.purpose == p),
                               "prompt_provider": sum(_int(r.prompt_provider) for r in self.records if r.purpose == p),
                               "completion_provider": sum(_int(r.completion_provider) for r in self.records if r.purpose == p),
                               "cost": (round(sum(r.cost for r in self.records if r.purpose == p and r.cost is not None), 6)
                                        if any(r.purpose == p and r.cost is not None for r in self.records) else None)}
                           for p in sorted({r.purpose for r in self.records if r.purpose})},
        }

    @staticmethod
    def _cached_pricing(reported: List[CallRecord]) -> str:
        cached = [r for r in reported if _int(r.cached_prompt)]
        if not cached:
            return "no input was reported as cached"
        discounted = sum(r.cache_discounted for r in cached)
        if discounted == len(cached):
            return "cached input priced at price_cached_in from config"
        full = sum(_int(r.cached_prompt) for r in cached if not r.cache_discounted)
        return (f"{full:,} cached input tokens priced at the full price_in (no price_cached_in in config for that model); "
                "if the provider bills cached input lower, the real charge is lower than this figure")

    # -- the gate before a call ------------------------------------------------------------------------------------------------
    def check(self, prompt_est: int, reserve_tokens: int, price: Optional[Tuple[float, float]]) -> None:
        lim = self.limits
        if lim.max_calls is not None and self.calls + 1 > lim.max_calls:
            raise BudgetExhausted("calls", f"this session's model-call cap ({lim.max_calls}) is used up")
        if lim.max_tokens is not None and self.charged_tokens + prompt_est + reserve_tokens > lim.max_tokens:
            raise BudgetExhausted("tokens", f"this session's token cap ({lim.max_tokens:,}) would be passed: "
                                  f"{self.charged_tokens:,} charged so far, this call needs up to {prompt_est + reserve_tokens:,}")
        if lim.max_seconds is not None and self.seconds > lim.max_seconds:
            raise BudgetExhausted("seconds", f"this session's time cap ({lim.max_seconds:.0f}s) is used up")
        if lim.max_cost is not None:
            if price is None:
                raise BudgetExhausted("cost", "a cost cap is set but this model has no price in config.yml (price_in / price_out)")
            spent = self.cost or 0.0
            if spent + (prompt_est * price[0] + reserve_tokens * price[1]) / 1e6 > lim.max_cost:
                raise BudgetExhausted("cost", f"this session's cost cap ({lim.max_cost:g} yuan) would be passed ({spent:.2f} spent)")

    def remaining(self) -> Dict[str, Any]:
        lim = self.limits
        return {"calls": None if lim.max_calls is None else max(0, lim.max_calls - self.calls),
                "tokens": None if lim.max_tokens is None else max(0, lim.max_tokens - self.charged_tokens),
                "seconds": None if lim.max_seconds is None else max(0, round(lim.max_seconds - self.seconds)),
                "cost": None if lim.max_cost is None else round(max(0.0, lim.max_cost - (self.cost or 0.0)), 2)}

    def can_escalate(self) -> bool:
        return self.limits.max_escalations is None or len(self.escalations) < self.limits.max_escalations

    def add(self, record: CallRecord) -> None:
        self.records.append(record)

    # -- persistence (the books survive a resumed session, so a cap is a cap over the whole work) ------------------------------
    def to_dict(self) -> Dict[str, Any]:
        return {"records": [r.to_dict() for r in self.records], "escalations": self.escalations,
                "seconds": round(self.seconds, 2)}

    def restore(self, data: Dict[str, Any]) -> None:
        self.records = [CallRecord.from_dict(r) for r in data.get("records") or []]
        self.escalations = list(data.get("escalations") or [])
        self._seconds_before = float(data.get("seconds") or 0.0)
        self.started = time.perf_counter()


# ---------------------------------------------------------------------------------------------------------------------------
# Roles and routing
# ---------------------------------------------------------------------------------------------------------------------------

@dataclass
class RoleSpec:
    role: str
    call: Callable[[str, str, int], Dict[str, Any]]
    provider: str = ""
    model: str = ""
    #: what was asked for: "" (provider default), "off", "on", or an effort level. Whether the provider branch honours it
    #: is a separate fact (thinking_applied), because several branches of AIClient send no such parameter.
    thinking: str = ""
    thinking_applied: bool = False
    price: Optional[Tuple[float, float]] = None        # yuan per million tokens, in / out, from config; None = unknown
    #: yuan per million for input the provider reports as cached (price_cached_in); None = priced at the full input price
    price_cached: Optional[float] = None
    #: the most output a call of this role may ask for (agent.models.<role>.max_output); None: the step's own allowance
    max_output: Optional[int] = None
    #: a thinking parameter is sent only because config declares its shape; no probe has shown the gateway acts on it
    thinking_unverified: bool = False
    #: the provider the role asked for, when the client fell back to another configured one ("" when it did not)
    provider_requested: str = ""


@dataclass
class Policy:
    mode: str = "single"                       # "single" or "escalate"
    escalate_after: int = 2                    # consecutive capability failures
    max_escalations_per_turn: int = 1
    handoff_max_tokens: int = 3000
    #: the escalation model's own rejection allowance; it is not charged for the primary's objections
    fresh_rejections: bool = True
    #: task kinds (task_kinds) that start on the escalation model, within max_escalations (agent.escalate_tasks); empty: none
    escalate_tasks: List[str] = field(default_factory=list)


@dataclass
class StepEvent:
    """What one controller step looked like, for the router. Built by the controller from the step's results."""
    step: int
    codes: List[str] = field(default_factory=list)
    progress: bool = False
    no_progress: int = 0
    #: each code's class in context (effective_classes); empty: the codes' own classes
    classes: Dict[str, str] = field(default_factory=dict)


@dataclass
class Decision:
    escalate: bool = False
    reason: str = ""
    trigger: str = ""
    codes: List[str] = field(default_factory=list)


def _sha(text: str) -> str:
    return hashlib.sha256(str(text or "").encode("utf-8")).hexdigest()[:12]


def estimate_tokens(text: str) -> int:
    cjk = len(re.findall(r"[一-鿿]", text))
    return int(cjk * 1.0 + (len(text) - cjk) / 3.5) + 1


class RoutedModel:
    """The callable the controller already takes (system, user, max_tokens) -> reply, over one or two models.

    Every call goes through the shared books; the controller tells it how each step went (`observe`) and it answers
    whether the next step goes to the escalation model (`decision`). A single-model configuration is the same object
    with no escalation role: the books and the records are identical, so a comparison between modes is like for like."""

    def __init__(self, primary: RoleSpec, escalation: Optional[RoleSpec] = None, *, policy: Optional[Policy] = None,
                 budget: Optional[SharedBudget] = None):
        self.roles: Dict[str, RoleSpec] = {ROLE_PRIMARY: primary}
        if escalation is not None:
            self.roles[ROLE_ESCALATION] = escalation
        self.policy = policy or Policy()
        if self.policy.mode == "escalate" and escalation is None:
            raise ValueError("routing is set to 'escalate' but no escalation model is configured (agent.models.escalation)")
        self.budget = budget or SharedBudget()
        if self.budget.limits.max_cost is not None:
            missing = [r.role for r in self.roles.values() if r.price is None]
            if missing:
                raise ValueError(f"a cost cap is set but no price is configured for: {', '.join(missing)} "
                                 "(price_in / price_out on the model's available_models entry)")
        self.role = ROLE_PRIMARY
        self.turn = 0
        self.turn_escalations = 0
        self._capability_run = 0
        self.turn_log: List[Dict[str, Any]] = []
        self._after = ""

    # -- a turn -----------------------------------------------------------------------------------------------------------------
    def begin_turn(self, turn: int, task: Optional[List[str]] = None) -> Optional[Dict[str, Any]]:
        """Start a turn on the first model, or, for a task kind listed in policy.escalate_tasks, on the escalation model (counted
        as an escalation, within the same caps). Returns the escalation record when it started there."""
        self.turn, self.role, self.turn_escalations, self._capability_run = turn, ROLE_PRIMARY, 0, 0
        self.turn_log, self._after = [], ""
        listed = sorted(set(task or []) & set(self.policy.escalate_tasks or []))
        if (listed and self.policy.mode == "escalate" and ROLE_ESCALATION in self.roles and self.budget.can_escalate()
                and self.policy.max_escalations_per_turn > 0):
            self.role = ROLE_ESCALATION
            self.turn_escalations = 1
            record = {"turn": turn, "step": 0, "after_call": self.budget.calls, "from": ROLE_PRIMARY, "to": ROLE_ESCALATION,
                      "reason": "configured_task", "trigger": "agent.escalate_tasks lists " + ", ".join(listed), "codes": [], "classes": [],
                      "model": self.roles[ROLE_ESCALATION].model, "handoff_chars": 0, "handoff_tokens_est": 0, "handoff_sha": ""}
            self.budget.escalations.append(record)
            return record
        return None

    def call_as(self, role: str, system: str, user: str, max_tokens: int, purpose: str) -> Dict[str, Any]:
        """One call by a named role for a stated purpose (a review), on the same books, without moving the turn to that role: it is
        not an escalation and does not count as one. The call is recorded with its purpose."""
        if role not in self.roles:
            raise KeyError(f"no {role} model configured")
        current, before = self.role, self.budget.calls
        self.role = role
        try:
            return self(system, user, max_tokens)
        finally:
            self.role = current
            if self.budget.calls > before:
                self.budget.records[-1].purpose = purpose       # a failed call keeps its purpose too

    def note_rejection(self) -> None:
        """The gate sent the last answer back: the next call is a retry of it, and is recorded as one."""
        self._after = "rejected"

    @property
    def escalated(self) -> bool:
        return self.role == ROLE_ESCALATION

    def __call__(self, system: str, user: str, max_tokens: int) -> Dict[str, Any]:
        spec = self.roles[self.role]
        if spec.max_output:
            max_tokens = min(int(max_tokens), int(spec.max_output))
        prompt_est = estimate_tokens(system) + estimate_tokens(user)
        self.budget.check(prompt_est, max_tokens, spec.price)
        started = time.perf_counter()
        thinking = spec.thinking + ("" if not spec.thinking or spec.thinking_applied else " (not sent by this provider branch)")
        if spec.thinking and spec.thinking_applied and spec.thinking_unverified:
            thinking += " (sent as config declares; gateway support unverified)"
        previous = next((r for r in reversed(self.budget.records) if r.turn == self.turn), None)
        after = self._after or ("failed" if previous is not None and not previous.ok else
                                ("discarded" if previous is not None and not previous.reply_used else ""))
        self._after = ""
        record = CallRecord(n=self.budget.calls + 1, role=spec.role, turn=self.turn, model=spec.model, provider=spec.provider, ok=True,
                            usage_source="estimate", prompt_est=prompt_est, thinking=thinking, max_output=int(max_tokens), after=after)
        try:
            reply = spec.call(system, user, max_tokens) or {}
        except Exception as exc:                                    # noqa: BLE001
            record.ok, record.usage_source = False, "failed"
            record.error = f"{type(exc).__name__}: {_ADDRESS.sub('<address>', str(exc))[:160]}"     # a gateway address is not for a file that gets pasted around
            record.seconds = round(time.perf_counter() - started, 2)
            self.budget.add(record)
            raise
        record.seconds = round(time.perf_counter() - started, 2)
        text = str(reply.get("content") or "")
        record.completion_est = estimate_tokens(text)
        if str(reply.get("token_source") or "") == "response_cache":
            record.usage_source = "response_cache"
            record.saved_prompt, record.saved_completion = reply.get("saved_prompt_tokens"), reply.get("saved_completion_tokens")
        elif reply.get("prompt_tokens") is not None:
            record.usage_source = "provider"
            record.prompt_provider, record.completion_provider = _int(reply.get("prompt_tokens")), _int(reply.get("completion_tokens"))
        record.cached_prompt, record.reasoning = reply.get("cached_prompt_tokens"), reply.get("reasoning_tokens")
        record.model_reported = reply.get("model_reported") or None
        if record.usage_source == "provider" and record.reasoning is not None and record.completion_provider is not None:
            record.reasoning_separate = _int(record.reasoning) > _int(record.completion_provider)
        if spec.price is not None and record.usage_source != "response_cache":
            tin = record.prompt_provider if record.usage_source == "provider" else record.prompt_est
            tout = record.completion_provider if record.usage_source == "provider" else record.completion_est
            if record.reasoning_separate:
                tout = _int(tout) + _int(record.reasoning)
            cached = min(_int(record.cached_prompt), _int(tin)) if record.usage_source == "provider" else 0
            if cached and spec.price_cached is not None:
                record.cache_discounted = True
                cost_in = (_int(tin) - cached) * spec.price[0] + cached * spec.price_cached
            else:
                cost_in = _int(tin) * spec.price[0]
            record.cost = round((cost_in + _int(tout) * spec.price[1]) / 1e6, 6)
        self.budget.add(record)
        return {**reply, "role": spec.role, "model": spec.model, "call_n": record.n}

    def mark_last_reply_unused(self) -> None:
        """The controller threw the reply away (unparseable, refused). It was paid for; it is counted and labelled."""
        if self.budget.records:
            self.budget.records[-1].reply_used = False

    # -- deciding ------------------------------------------------------------------------------------------------------------------
    def observe(self, event: StepEvent) -> Decision:
        """Called after each step. Escalates only on consecutive capability failures, once per turn, once the budget allows."""
        classes = {c: (event.classes or {}).get(c) or classify(c) for c in event.codes}
        if event.progress or not event.codes:
            self._capability_run = 0
        elif is_capability_failure(event.codes, classes):
            self._capability_run += 1
        else:
            self._capability_run = 0            # waiting on evidence, a person, a refresh or the harness: not the model's limit
        self.turn_log.append({"step": event.step, "role": self.role, "codes": list(event.codes), "classes": sorted(set(classes.values())),
                              "progress": event.progress, "capability_run": self._capability_run})
        if (self.policy.mode != "escalate" or self.role == ROLE_ESCALATION or ROLE_ESCALATION not in self.roles
                or self.turn_escalations >= self.policy.max_escalations_per_turn or not self.budget.can_escalate()
                or self._capability_run < max(1, self.policy.escalate_after)):
            return Decision()
        return Decision(True, reason="capability", trigger=f"{self._capability_run} consecutive steps the model did not manage with the evidence it had",
                        codes=sorted({c for e in self.turn_log[-self._capability_run:] for c in e["codes"]
                                      if c in classes and classes[c] in CAPABILITY_CLASSES or classify(c) in CAPABILITY_CLASSES}))

    def escalate(self, decision: Decision, handoff: str, step: int) -> Dict[str, Any]:
        self.role = ROLE_ESCALATION
        self.turn_escalations += 1
        self._capability_run = 0
        record = {"turn": self.turn, "step": step, "after_call": self.budget.calls, "from": ROLE_PRIMARY, "to": ROLE_ESCALATION,
                  "reason": decision.reason, "trigger": decision.trigger, "codes": decision.codes,
                  "classes": sorted(classify_all(decision.codes)), "model": self.roles[ROLE_ESCALATION].model,
                  "handoff_chars": len(handoff), "handoff_tokens_est": estimate_tokens(handoff), "handoff_sha": _sha(handoff)}
        self.budget.escalations.append(record)
        return record

    def stop_class(self, stop_reason: str, *, open_human: bool, last_codes: List[str], ok: bool,
                   last_classes: Optional[Dict[str, str]] = None) -> str:
        """Why a turn ended, in the terms the comparison needs: finished, ran out of budget, needs a person, the evidence is not
        there, the model did not manage with the evidence it had (and whether the escalation model was tried), or the provider failed."""
        if ok:
            return "ok"
        if stop_reason.startswith("budget") or stop_reason == "budget_session":
            return "budget"
        if stop_reason == "model_error":
            return "provider_error"
        if stop_reason == "needs_user" or open_human:
            return "needs_human"
        if stop_reason == "semantic_blocked":
            return "semantic_objection"           # verifies, but the reviewer's objection to its meaning stands
        if stop_reason == "review_pending":
            return "review_unavailable"           # a review was due and did not happen
        kinds: Dict[str, List[str]] = {}
        for c in last_codes:
            kind = (last_classes or {}).get(c) or classify(c)
            if kind != INFO:
                kinds.setdefault(kind, []).append(c)
        if kinds.get(ABSENT) or kinds.get(EVIDENCE) or kinds.get(REFRESH):
            return "evidence_insufficient"
        if kinds.get(SYSTEM):
            return "tool_error"
        return "capability_insufficient_after_escalation" if self.turn_escalations else "capability_insufficient"


# ---------------------------------------------------------------------------------------------------------------------------
# The hand-off: a compact task state, not the transcript
# ---------------------------------------------------------------------------------------------------------------------------

def _short(value: Any, limit: int = 160) -> str:
    text = value if isinstance(value, str) else json.dumps(value, ensure_ascii=False, default=str)
    return text if len(text) <= limit else text[:limit] + "..."


def build_handoff(*, goal: str, mode: str, account: str, source_version: str, draft: Optional[Dict[str, Any]], cited: List[Dict[str, Any]],
                  actions: List[Dict[str, Any]], failures: List[Dict[str, Any]], issues: List[Dict[str, Any]], remaining: Dict[str, Any],
                  max_tokens: int = 3000) -> str:
    """What the escalation model is told instead of the whole conversation: the task, the latest draft with the verifier's
    objections, the figures that draft cites (one line each, so it can reuse the refs), what was already done, what is
    open, and what is left of the budget. It is rebuilt from state, so it carries no reasoning of the first model to be
    re-reviewed or trusted. It is held to `max_tokens` by showing fewer of each list, never by cutting a line in half; only if
    that is not enough is the tail dropped, with a marker."""

    def render(n_claims: int, n_obj: int, n_cited: int, n_actions: int, n_issues: int) -> str:
        lines = [f"TASK STATE (handed over from the first model; the workbook source version is {source_version})",
                 f"  goal: {_short(goal, 400)}", f"  mode: {mode}" + (f"; account: {account}" if account else "")]
        if draft and draft.get("claims"):
            lines.append("  latest draft and what the verifier objected to:")
            for claim in draft["claims"][:n_claims]:
                lines.append("    - " + _short({k: claim.get(k) for k in ("kind", "periods", "cites", "cause_cites", "inference", "text")
                                                 if claim.get(k) not in (None, [], "")}, 360))
            for fail in (draft.get("objections") or [])[:n_obj]:
                lines.append(f"    ! {_short(fail, 260)}")
        if cited:
            lines.append("  figures the draft cites (you may cite these refs again):")
            for item in cited[:n_cited]:
                lines.append("    " + " ".join(f"{k}={_short(item[k], 60)}" for k in ("ref", "period", "basis", "line", "value") if item.get(k) not in (None, "")))
        if actions:
            lines.append("  already done (tool, outcome):")
            for a in actions[-n_actions:]:
                lines.append(f"    - {a.get('tool')}({_short(a.get('args') or {}, 90)}) -> {'ok' if a.get('ok') else 'error ' + str(a.get('error'))}")
        if failures:
            lines.append("  why it did not finish: " + "; ".join(f"{f['code']} [{f['class']}] x{f['count']}" for f in failures[:8]))
        open_issues = [i for i in issues if not i.get("resolved")]
        if open_issues:
            lines.append("  open issues:")
            for i in open_issues[:n_issues]:
                lines.append(f"    - [{i.get('severity')}] {i.get('account') or 'workbook'}: {_short(i.get('message'), 200)}")
        lines.append("  budget left: " + ", ".join(f"{k} {v}" for k, v in remaining.items() if v is not None))
        lines.append("  Continue this task. Fix what the verifier objected to using the figures above or a tool; do not start over and do not "
                     "re-check what already verified.")
        return "\n".join(lines)

    caps = [8, 8, 24, 8, 6]
    text = render(*caps)
    while estimate_tokens(text) > max_tokens and any(c > 1 for c in caps):
        caps = [max(1, c // 2) for c in caps]
        text = render(*caps)
    if estimate_tokens(text) > max_tokens:
        text = text[: int(max_tokens * 3)] + "\n  [hand-off cut to its size limit]"
    return text

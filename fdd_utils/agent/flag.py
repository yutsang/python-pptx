"""What the agent block of config.yml says, read from one place: the budgets of a run, the call caps, the routing, and the flags the
controller reads (prepared write, repair of an inference, the read of a sentence's binding, ...). The agent is the app's one path: there
is no switch."""
from __future__ import annotations

import os
from typing import Any, Dict, Optional

from .controller import Budget


def _config(config: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    if config is not None:
        return config
    try:
        from ..ai import FDDConfig
        return FDDConfig().config or {}
    except Exception:
        return {}


def agent_budget(config: Optional[Dict[str, Any]] = None) -> Budget:
    block = _config(config).get("agent") or {}
    defaults = Budget()
    return Budget(
        max_steps=int(block.get("max_steps", defaults.max_steps)), max_seconds=float(block.get("max_seconds", defaults.max_seconds)),
        max_tokens=int(block.get("max_tokens", defaults.max_tokens)), max_no_progress=int(block.get("max_no_progress", defaults.max_no_progress)),
        max_rejections=int(block.get("max_rejections", defaults.max_rejections)))


def max_calls_per_account(config: Optional[Dict[str, Any]] = None) -> Optional[int]:
    """agent.max_calls_per_account (default 40, 0 = no cap): the runaway guard of a whole book is this times its accounts (book.py). It is
    far above what an account needs (a write, a repair or two, a read of a sentence's binding); it is there for a loop that does not end."""
    value = (_config(config).get("agent") or {}).get("max_calls_per_account")
    if value in (None, ""):
        return 40
    return int(value) or None


def agent_max_model_calls(config: Optional[Dict[str, Any]] = None) -> Optional[int]:
    """A hard cap on model calls over one working session (agent.max_model_calls); None means no cap."""
    value = (_config(config).get("agent") or {}).get("max_model_calls")
    return int(value) if value not in (None, "", 0) else None


# ---------------------------------------------------------------------------------------------------------------------------
# Routing and the shared budget (routing.py). Everything below is read from the agent block; nothing here switches a default.
# ---------------------------------------------------------------------------------------------------------------------------

def _agent_block(config: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    return _config(config).get("agent") or {}


def prepared_write(config: Optional[Dict[str, Any]] = None) -> bool:
    """agent.prepared_write (default true): a paragraph for one known account is written once from evidence the controller read
    itself, then repaired claim by claim; false sends it through the tool loop from the first step, as before."""
    value = _agent_block(config).get("prepared_write")
    return True if value is None else bool(value)


def lead_write(config: Optional[Dict[str, Any]] = None) -> bool:
    """agent.lead_write (default false): the model leads each paragraph -- it reads the evidence and records what must be said,
    decides who writes, checks every claim against what it cites and the verifier's report, and decides. Python reads the workbook,
    computes, and verifies as a tool. With agent.models.escalation set, that model leads and the primary writes."""
    return bool(_agent_block(config).get("lead_write"))


def plan_first(config: Optional[Dict[str, Any]] = None) -> bool:
    """agent.plan_first (default false): in a prepared write the escalation model reads the evidence once and writes a short plan
    (the points to make and the traps to avoid); the primary model writes from it, the verifier decides, and a paragraph the
    primary cannot repair goes to the escalation model as in routing. The lead without its checking loop: the lead arms checked
    every claim and cost 6 to 8 yuan a databook for 8 to 11 accepted of 26 (2026-11-08). Off until entity_trial (arm 'planned')
    shows it helps."""
    return bool(_agent_block(config).get("plan_first"))


def tool_loop_on_repeat(config: Optional[Dict[str, Any]] = None) -> bool:
    """agent.tool_loop_on_repeat (default false): when a prepared write's draft fails the same way twice, hand the turn to the
    tool loop. On the client file the hand-over cost 0.3 to 0.9 yuan an account on the stronger model and finished one of four
    (2026-11-09); off, the write goes on to max_rejections and the draft reaches the deck with its rejected sentences marked."""
    return bool(_agent_block(config).get("tool_loop_on_repeat"))


def research(config: Optional[Dict[str, Any]] = None) -> bool:
    """agent.research (default false): a prepared write tells the writer what its evidence says nothing about (the rate and maturity of
    a loan, the ageing of receivables, a depreciation policy, what cash is made of) and lets it ask ONCE for up to
    agent.research_lookups read-only lookups (search_notes, get_account_components, adjustment_evidence, cross_account_evidence,
    read_source_block); what a search finds is filed as notes it may cite. When it still lacks something it does not hand the turn to
    the tool loop (0.3 to 0.9 yuan an account on the stronger model, one in four finished, 2026-11-09): the gap is written on the
    candidate and shown to the reader as a source gap. Compared with arm 'flash' on six books (124 accounts, 2026-11-17): the writer asked
    for no lookup at all (0 of 124), accepted accounts 120 against 120, accounts with a reason no note supports 20 against 22. Run
    by hand on four local books, the queries of the 17 gaps they have found a remark for 9, of which one says anything (a management
    remark on ageing): the workbooks mostly do not hold what the gaps ask for. Left off."""
    return bool(_agent_block(config).get("research"))


def repair_inference(config: Optional[Dict[str, Any]] = None) -> bool:
    """agent.repair_inference (default on): a draft that passes only as an inference, a reason with no note behind it, goes back to
    the writer once to cite the note or drop the reason, before the deck colours the sentence."""
    value = _agent_block(config).get("repair_inference")
    return True if value is None else bool(value)


def judge_binding(config: Optional[Dict[str, Any]] = None) -> bool:
    """agent.judge_binding (default on): a sentence the verifier refused only for how it READ the prose (the period or line a figure
    sits under, a make-up, a rest) is read by the agent on the last refusal, figure by figure, and Python holds the reading to the
    evidence. The stronger model reads when one is configured, else the writer's. Off, such a sentence ends orange, not cleared."""
    value = _agent_block(config).get("judge_binding")
    return True if value is None else bool(value)


def research_lookups(config: Optional[Dict[str, Any]] = None) -> int:
    """agent.research_lookups (default 3, at most 6): the lookups one account's writer may ask for, in all."""
    value = _agent_block(config).get("research_lookups")
    return max(0, min(6, int(value))) if value not in (None, "") else 3


def seed_facts(config: Optional[Dict[str, Any]] = None) -> bool:
    """agent.seed_facts (default false): in a prepared write Python writes the balance and movement sentences and the model adds
    the rest. Off until a paid comparison (entity_trial.py, arm 'seeded') shows it helps."""
    return bool(_agent_block(config).get("seed_facts"))


def review_meaning(config: Optional[Dict[str, Any]] = None) -> bool:
    """agent.review_meaning (default false): a verified paragraph with something only a reader can judge (a cause, notes, a basis
    caveat, part-year figures) is read once by the escalation model for meaning. Needs agent.models.escalation."""
    return bool(_agent_block(config).get("review_meaning"))


def review_marks_only(config: Optional[Dict[str, Any]] = None) -> bool:
    """agent.review_meaning: mark -- the review marks what it objects to and changes nothing: each objection colours its sentence
    orange in the deck with the reviewer's reason, the paragraph stays in. The repair mode (true) sends objections back to the
    writer and leaves an unsettled paragraph out of the deck. On the client file (2026-11-10) all ten wrong sentences a review
    found were meaning errors Python cannot see (a shareholder's capital, raw lines as the adjusted make-up, a gross loan as the
    long-term part, a note's broad period against its detail, a cause that cannot carry the movement), and every one was black."""
    return str(_agent_block(config).get("review_meaning") or "").strip().lower() == "mark"


def _number(value: Any, kind=int) -> Optional[Any]:
    """0, '' and None mean 'no cap'; anything else must be a number."""
    if value in (None, "", 0, 0.0, False):
        return None
    return kind(value)


def agent_limits(config: Optional[Dict[str, Any]] = None, *, max_model_calls: Optional[int] = None):
    """The caps that span a whole working session: every role, every turn, every attempt (routing.SharedBudget)."""
    from .routing import Limits
    block = _agent_block(config)
    return Limits(max_calls=max_model_calls if max_model_calls is not None else _number(block.get("max_model_calls")),
                  max_tokens=_number(block.get("max_session_tokens")), max_seconds=_number(block.get("max_session_seconds"), float),
                  max_cost=_number(block.get("max_session_cost"), float), max_escalations=_number(block.get("max_session_escalations")))


def agent_policy(config: Optional[Dict[str, Any]] = None, *, mode: Optional[str] = None):
    """single (default) or escalate. The environment (FDD_AGENT_ROUTING) and an explicit `mode` win over the file."""
    from .routing import Policy
    block = _agent_block(config)
    chosen = (mode or os.environ.get("FDD_AGENT_ROUTING") or block.get("routing") or "single").strip().lower()
    if chosen not in ("single", "escalate"):
        raise ValueError(f"agent.routing must be 'single' or 'escalate', not {chosen!r}")
    defaults = Policy()
    tasks = block.get("escalate_tasks") or []
    if isinstance(tasks, str):
        tasks = [t.strip() for t in tasks.split(",") if t.strip()]
    return Policy(mode=chosen, escalate_after=int(block.get("escalate_after", defaults.escalate_after)),
                  max_escalations_per_turn=int(block.get("max_escalations", defaults.max_escalations_per_turn)),
                  handoff_max_tokens=int(block.get("handoff_max_tokens", defaults.handoff_max_tokens)),
                  escalate_tasks=[str(t) for t in tasks])


def role_config(config: Optional[Dict[str, Any]], role: str) -> Dict[str, Any]:
    """{'provider', 'model', 'thinking', 'reasoning_effort', 'max_output'} configured for a role under agent.models; empty strings
    (and None for max_output) when unset."""
    entry = ((_agent_block(config).get("models") or {}).get(role)) or {}
    return {"provider": str(entry.get("provider") or ""), "model": str(entry.get("model") or ""), "thinking": str(entry.get("thinking") or ""),
            "reasoning_effort": str(entry.get("reasoning_effort") or ""), "max_output": _number(entry.get("max_output"))}


def parse_role(text: str) -> Dict[str, Any]:
    """'provider::model' (the form compare_models.py uses), or a bare provider."""
    provider, _, model = str(text or "").partition("::")
    return {"provider": provider.strip(), "model": model.strip(), "thinking": "", "reasoning_effort": "", "max_output": None}


def price_for(config: Optional[Dict[str, Any]], provider: str, model: str):
    """(price_in, price_out) in yuan per million tokens from the model's available_models entry, or None. Never a built-in
    table and never another provider's price: a model whose price nobody wrote down has no cost, not a borrowed one."""
    block = _config(config).get(provider)
    for entry in (block.get("available_models") or []) if isinstance(block, dict) else []:
        if isinstance(entry, dict) and str(entry.get("id")) == str(model) and entry.get("price_in") is not None:
            return (float(entry["price_in"]), float(entry.get("price_out") or 0.0))
    return None


def cached_price_for(config: Optional[Dict[str, Any]], provider: str, model: str) -> Optional[float]:
    """price_cached_in (yuan per million) for input the provider reports as served from its prompt cache, or None. Only what the
    entry says: with none, cached input is priced at the full price_in and the books say so; no discount is assumed."""
    block = _config(config).get(provider)
    for entry in (block.get("available_models") or []) if isinstance(block, dict) else []:
        if isinstance(entry, dict) and str(entry.get("id")) == str(model) and entry.get("price_cached_in") is not None:
            return float(entry["price_cached_in"])
    return None

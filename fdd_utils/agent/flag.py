"""Is the agent path switched on? One answer, read from one place.

Default off. The environment wins over the file so a single run can be switched either way
without editing config.yml (and so a run on a machine whose config.yml was edited by hand is
not silently different): FDD_AGENT=1 / true / on / yes enables, 0 / false / off / no disables.
Nothing in the legacy pipeline reads this; switching it off is the rollback.
"""
from __future__ import annotations

import os
from typing import Any, Dict, Optional

from .controller import Budget

_ON = {"1", "true", "on", "yes"}
_OFF = {"0", "false", "off", "no"}


def _config(config: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    if config is not None:
        return config
    try:
        from ..ai import FDDConfig
        return FDDConfig().config or {}
    except Exception:
        return {}


def agent_enabled(config: Optional[Dict[str, Any]] = None) -> bool:
    env = os.environ.get("FDD_AGENT", "").strip().lower()
    if env in _ON:
        return True
    if env in _OFF:
        return False
    return bool((_config(config).get("agent") or {}).get("enabled", False))


def agent_budget(config: Optional[Dict[str, Any]] = None) -> Budget:
    block = _config(config).get("agent") or {}
    defaults = Budget()
    return Budget(
        max_steps=int(block.get("max_steps", defaults.max_steps)), max_seconds=float(block.get("max_seconds", defaults.max_seconds)),
        max_tokens=int(block.get("max_tokens", defaults.max_tokens)), max_no_progress=int(block.get("max_no_progress", defaults.max_no_progress)),
        max_rejections=int(block.get("max_rejections", defaults.max_rejections)))


def agent_max_model_calls(config: Optional[Dict[str, Any]] = None) -> Optional[int]:
    """A hard cap on model calls over one working session (agent.max_model_calls); None means no cap."""
    value = (_config(config).get("agent") or {}).get("max_model_calls")
    return int(value) if value not in (None, "", 0) else None


# ---------------------------------------------------------------------------------------------------------------------------
# Routing and the shared budget (routing.py). Everything below is read from the agent block; nothing here switches a default.
# ---------------------------------------------------------------------------------------------------------------------------

def _agent_block(config: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    return _config(config).get("agent") or {}


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

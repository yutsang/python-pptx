"""What each configured agent role would send to its provider, shown without sending it; and, only when asked, one tiny real call.

The dry run builds the real client for a role (the same AIClientModel the session uses) and runs the client's own request
code with the network call replaced by a recorder, so what is shown is what that code would send: the model id, the output
allowance, temperature, and any thinking body or reasoning effort. Header VALUES and the address are not shown (the header
names are). Nothing leaves the machine and nothing is written to the usage ledger.

`send` makes one small real request per role (paid, a few dozen tokens) and reports what came back: whether the provider
accepted the parameters, the usage it reported (prompt, completion, reasoning, cache) and the model id it says answered.
That, not this file, is the evidence for whether a gateway honours a thinking parameter.
"""
from __future__ import annotations

import logging
import time
from types import SimpleNamespace
from typing import Any, Dict, Optional
from urllib.parse import urlparse


class _Captured(Exception):
    """Raised by the recorder in place of the network call."""


class _Recorder:
    def __init__(self) -> None:
        self.kwargs: Dict[str, Any] = {}
        self.chat = SimpleNamespace(completions=SimpleNamespace(create=self.create))

    def create(self, **kwargs: Any) -> Any:
        self.kwargs = kwargs
        raise _Captured("preview: the request was recorded, not sent")


def _host(client: Any) -> Optional[str]:
    base = str(getattr(client, "base_url", "") or "")
    host = urlparse(base).hostname if base else None
    if not host:
        return None
    parts = host.split(".")
    return ("*." + ".".join(parts[-2:])) if len(parts) > 2 else host


def preview(role: Any, *, system: str = "probe", user: str = "probe", max_tokens: Optional[int] = None) -> Dict[str, Any]:
    """The request `role` (an AIClientModel) would send, from the client's own code path, without sending it."""
    client = getattr(role, "client", None)
    out: Dict[str, Any] = {"provider": getattr(role, "provider", None), "provider_requested": getattr(role, "requested_provider", None),
                           "provider_fallback": bool(getattr(role, "provider_fallback", False)), "configured_model": getattr(role, "chat_model", None),
                           "thinking_requested": (getattr(role, "thinking", "") or getattr(role, "reasoning_effort", "") or None),
                           "thinking_sent": bool(getattr(role, "thinking_applied", False)),
                           "thinking_unverified": bool(getattr(role, "thinking_unverified", False)), "sent": False}
    if client is None or getattr(client, "use_heuristic", False) or getattr(client, "client", None) is None:
        out["error"] = "no provider client (heuristic mode or the provider is not configured)"
        return out
    real, recorder = client.client, _Recorder()
    out["endpoint_host"] = _host(real)
    params = dict(getattr(role, "params", {}) or {})
    level = client.logger.level
    client.client = recorder
    client.logger.setLevel(logging.CRITICAL)
    try:
        client._request(user, system, 0.1, int(max_tokens or getattr(role, "max_output", None) or 1500), None, None, None,
                        params.get("allow_thinking"), params.get("reasoning_effort"))
    except _Captured:
        pass
    except Exception as exc:                                            # noqa: BLE001
        out["error"] = f"the request could not be built: {type(exc).__name__}: {str(exc)[:200]}"
    finally:
        client.client = real
        client.logger.setLevel(level)
    kw = recorder.kwargs
    if kw:
        out.update(model=kw.get("model"), max_output=kw.get("max_tokens") or kw.get("max_completion_tokens"),
                   output_field="max_completion_tokens" if "max_completion_tokens" in kw else "max_tokens",
                   temperature=kw.get("temperature"), extra_body=kw.get("extra_body"), reasoning_effort=kw.get("reasoning_effort"),
                   header_names=sorted((kw.get("extra_headers") or {}).keys()), messages=len(kw.get("messages") or []))
    return out


def send(role: Any, *, max_tokens: int = 16) -> Dict[str, Any]:
    """One small real request through the role (paid). Reports acceptance, the usage the provider reported, and the model id."""
    started = time.perf_counter()
    try:
        reply = role("Reply with the single word OK.", "OK?", max_tokens)
    except Exception as exc:                                            # noqa: BLE001
        return {"ok": False, "error": f"{type(exc).__name__}: {str(exc)[:240]}", "seconds": round(time.perf_counter() - started, 2)}
    return {"ok": True, "seconds": round(time.perf_counter() - started, 2), "reply_chars": len(str(reply.get("content") or "")),
            "prompt_tokens": reply.get("prompt_tokens"), "completion_tokens": reply.get("completion_tokens"),
            "reasoning_tokens": reply.get("reasoning_tokens"), "cached_prompt_tokens": reply.get("cached_prompt_tokens"),
            "token_source": reply.get("token_source"), "model_reported": reply.get("model_reported"),
            "reasoning_inside_completion": (None if reply.get("reasoning_tokens") is None or reply.get("completion_tokens") is None
                                            else int(reply["reasoning_tokens"]) <= int(reply["completion_tokens"]))}

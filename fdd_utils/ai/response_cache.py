"""Answers the model has already given, kept by the exact request that got them.

A deck is rarely generated once. The same workbook is run again after a
template tweak, a changed threshold, a fixed check -- one real engagement file
went through the pipeline more than ten times in a day -- and on each of those
runs most accounts send the model a request identical, byte for byte, to one
it has already answered. Input tokens are 97% of what a run costs (measured:
736,756 in against 18,409 out for one 31-account run), so re-asking is where
the money and the minutes go.

The key is the request itself: provider, model id, both prompts and every
sampling parameter. Nothing about the account, the workbook or the code is
consulted, which is what makes it safe -- anything that changes what the model
is asked (an edited cell, a reworded instruction, a peer account's figure
quoted in this one's prompt, a different model) changes the key and is asked
afresh. What does NOT change the key is everything downstream of the answer:
the verifier, the repair pass, the deck. Those run again on every run, so a
fixed check is exercised against the stored drafts at no cost.

A feedback retry is a different request (it carries the first draft and the
reason it was rejected), so it is never served the draft it is meant to
replace.

Off unless processing.response_cache.enabled is true in config.yml, because
that file is per-machine and an unconfigured deployment must behave as it did.
FDD_RESPONSE_CACHE=0 / 1 in the environment overrides the file either way --
set 0 to draw fresh samples, e.g. to see how much a model varies run to run.
Entries live under fdd_utils/logs/response_cache/ (gitignored: they hold the
same client figures the run logs do) and are ignored once older than
max_age_days.
"""
from __future__ import annotations

import hashlib
import json
import os
import time
from typing import Any, Dict, Optional

from ..financial_common import package_file_path

_DIR = "response_cache"
_DEFAULT_MAX_AGE_DAYS = 30


def settings(config: Optional[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    """{"dir", "max_age_days"} when the cache is on, else None."""
    block = ((config or {}).get("processing") or {}).get("response_cache")
    if isinstance(block, bool):
        block = {"enabled": block}
    block = block if isinstance(block, dict) else {}
    enabled = bool(block.get("enabled", False))
    override = os.environ.get("FDD_RESPONSE_CACHE", "").strip().lower()
    if override in ("0", "false", "off", "no"):
        enabled = False
    elif override in ("1", "true", "on", "yes"):
        enabled = True
    if not enabled:
        return None
    return {
        "dir": str(block.get("dir") or os.path.join(package_file_path("logs"), _DIR)),
        "max_age_days": float(block.get("max_age_days") or _DEFAULT_MAX_AGE_DAYS),
    }


def request_key(**request: Any) -> str:
    """One hash over everything that decides what the model is asked."""
    blob = json.dumps(request, ensure_ascii=False, sort_keys=True, default=str)
    return hashlib.sha256(blob.encode("utf-8")).hexdigest()


def _path(cache: Dict[str, Any], key: str) -> str:
    return os.path.join(cache["dir"], key[:2], key + ".json")


def load(cache: Dict[str, Any], key: str) -> Optional[Dict[str, Any]]:
    path = _path(cache, key)
    try:
        if time.time() - os.path.getmtime(path) > cache["max_age_days"] * 86400:
            return None
        with open(path, encoding="utf-8") as handle:
            entry = json.load(handle)
    except (OSError, ValueError):
        return None
    return entry if isinstance(entry, dict) and str(entry.get("content") or "").strip() else None


def store(cache: Dict[str, Any], key: str, entry: Dict[str, Any]) -> None:
    """Written to a temp file and renamed, so a reader in another worker thread
    never opens half an entry. A failed write costs nothing but the saving."""
    if not str(entry.get("content") or "").strip():
        return  # an empty answer is a failure to retry, not a result to keep
    path = _path(cache, key)
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        temp = "%s.%s.tmp" % (path, os.getpid())
        with open(temp, "w", encoding="utf-8") as handle:
            json.dump(entry, handle, ensure_ascii=False)
        os.replace(temp, path)
    except OSError:
        pass

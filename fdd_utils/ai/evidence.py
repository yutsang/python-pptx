"""Which data an account was written from: a fingerprint of its frame.

A saved run (and a saved agent session) records each account's fingerprint, so an account whose workbook changed can be told from one
that merely did not finish, and a base run can be checked against the workbook it is laid out against (agent/session.py: check_base_run).
"""
from __future__ import annotations

import hashlib
import json
from typing import Any, Optional



#: attrs a RUN writes onto the frames while rendering (prompts.py stashes the
#: budget and the computed remainder there). Not workbook content, so not part
#: of a fingerprint: a frame that has been rendered once must still match the
#: same frame fresh from the workbook.
_RUN_WRITTEN_ATTRS = frozenset({
    "prompt_budget", "prompt_budget_exclude", "prompt_residual", "prompt_budget_residual",
})


def _canonical(value: Any, depth: int = 0) -> Any:
    """A JSON-able, process-independent picture of a frame and its attrs.
    Never a repr of an unknown object: one carrying an address would change
    in every process and make every resume look like a different workbook."""
    if depth > 8:
        return "..."
    if value is None or isinstance(value, bool) or isinstance(value, str):
        return value
    if hasattr(value, "itertuples") and hasattr(value, "columns"):  # a DataFrame, duck-typed
        attrs = getattr(value, "attrs", None) or {}
        return {
            "columns": [str(c) for c in value.columns],
            "rows": [[_canonical(v, depth + 1) for v in row] for row in value.itertuples(index=False)],
            "attrs": _canonical({k: v for k, v in attrs.items() if str(k) not in _RUN_WRITTEN_ATTRS}, depth + 1),
        }
    if isinstance(value, dict):
        return {str(k): _canonical(v, depth + 1) for k, v in value.items() if str(k) not in _RUN_WRITTEN_ATTRS}
    if isinstance(value, (set, frozenset)):
        return sorted((_canonical(v, depth + 1) for v in value), key=repr)
    if isinstance(value, (list, tuple)):
        return [_canonical(v, depth + 1) for v in value]
    iso = getattr(value, "isoformat", None)
    if callable(iso):
        try:
            return iso()
        except Exception:
            pass
    tolist = getattr(value, "tolist", None)
    if callable(tolist):
        try:
            return _canonical(tolist(), depth + 1)
        except Exception:
            pass
    try:
        number = float(value)
    except (TypeError, ValueError):
        return type(value).__name__
    if number != number or number in (float("inf"), float("-inf")):
        return None
    return round(number, 6)


def frame_fingerprint(df: Any) -> Optional[str]:
    """Which data an account was written from: a hash of its frame's values,
    labels and workbook-derived attrs (the nested analysis frame, notes,
    remarks, detail tables), taken before anything renders.

    A checkpoint line and an evidence file carry it, so a resume can tell an
    account whose workbook changed from one that merely died mid-run. Without
    it, a resume onto an edited workbook re-used every old draft and every old
    pool and said nothing.
    """
    if df is None:
        return None
    try:
        blob = json.dumps(_canonical(df), ensure_ascii=False, sort_keys=True, default=str)
    except Exception:
        return None
    return hashlib.sha256(blob.encode("utf-8")).hexdigest()[:16]


# -- building ----------------------------------------------------------------


# -- files -------------------------------------------------------------------



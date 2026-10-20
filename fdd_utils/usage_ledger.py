"""One line per LLM call and per user action, for usage and cost accounting.

The per-run folders (fdd_utils/logs/run_*/data.yml) hold the pipeline's own
stage calls, keyed account x stage -- a retry overwrites the attempt before it,
and calls made outside the pipeline (executive summaries, mapping help, the
re-prompt) are not there at all. This ledger is the complete record instead:

* ``fdd_utils/logs/usage.jsonl`` -- every call through AIClient.get_response,
  successful or not, with prompt / completion / total tokens as the provider
  reported them (or an estimate, labelled as one), the model, the stage, and
  which upload / session / account it was made for.
* ``fdd_utils/logs/activity.jsonl`` -- what users did: uploads (name, size,
  hash), Process Data, AI runs, exports, admin unlocks and downloads.

Attribution travels in a context variable: ``scope(file=..., account=...)``
around the work, and ``bind(fn)`` on anything handed to another thread (a
ThreadPoolExecutor task or a timeout thread does not inherit context on its
own). Recording never raises -- a full disk must not cost a user their deck.
"""
from __future__ import annotations

import contextlib
import contextvars
import json
import os
import threading
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, Iterator, Optional

LOG_DIR = Path(__file__).resolve().parent / "logs"
USAGE_PATH = LOG_DIR / "usage.jsonl"
ACTIVITY_PATH = LOG_DIR / "activity.jsonl"

_CONTEXT: contextvars.ContextVar[Dict[str, Any]] = contextvars.ContextVar("fdd_usage_context", default={})
_WRITE_LOCK = threading.Lock()


def current() -> Dict[str, Any]:
    return dict(_CONTEXT.get())


@contextlib.contextmanager
def scope(**fields: Any) -> Iterator[None]:
    """Attribution for every call and event inside the block (merged over any
    outer scope; None values are dropped)."""
    merged = {**_CONTEXT.get(), **{k: v for k, v in fields.items() if v is not None}}
    token = _CONTEXT.set(merged)
    try:
        yield
    finally:
        _CONTEXT.reset(token)


def set_context(**fields: Any) -> None:
    """Replace this thread's attribution outright (None values dropped). For a
    Streamlit rerun, which has no natural `with` block around it: called at the
    top of every rerun, so a value never outlives the rerun that set it."""
    _CONTEXT.set({k: v for k, v in fields.items() if v is not None})


def bind(fn: Callable[..., Any], **fields: Any) -> Callable[..., Any]:
    """fn, run in a copy of the CURRENT context (plus `fields`) whichever thread
    calls it. A fresh copy per call, so many executor tasks can share one bound
    function."""
    snapshot = contextvars.copy_context()
    extra = {k: v for k, v in fields.items() if v is not None}

    def run(*args: Any, **kwargs: Any) -> Any:
        ctx = snapshot.copy()

        def inner() -> Any:
            if extra:
                _CONTEXT.set({**_CONTEXT.get(), **extra})
            return fn(*args, **kwargs)

        return ctx.run(inner)

    return run


def _append(path: Path, record: Dict[str, Any]) -> None:
    try:
        line = json.dumps(record, ensure_ascii=False, default=str)
        with _WRITE_LOCK:
            path.parent.mkdir(parents=True, exist_ok=True)
            with path.open("a", encoding="utf-8") as handle:
                handle.write(line + "\n")
    except Exception:
        pass


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


_STAGE_NAMES = {"subagent_1": "Generator", "subagent_2": "Auditor",
                "subagent_3": "Refiner", "subagent_4": "Validator"}


def record_llm_call(**fields: Any) -> None:
    """One provider call (or a cache hit, which costs nothing and says so).

    `stage` comes from the caller's scope when it set one (the pipeline does,
    per account and stage); otherwise it is the AIClient's own agent name --
    the summary writer, mapping help, the re-prompt."""
    record = {"time": _now(), **current(), **fields}
    stage = record.get("stage") or record.get("client_agent") or ""
    record["stage"] = _STAGE_NAMES.get(str(stage), stage)
    _append(USAGE_PATH, record)


def record_event(event: str, **fields: Any) -> None:
    """One user action. `event` names it: upload, process_data, ai_run, export, admin_*."""
    _append(ACTIVITY_PATH, {"time": _now(), "event": event, **current(), **fields})


def read_jsonl(path: Path) -> list:
    rows = []
    try:
        with path.open("r", encoding="utf-8") as handle:
            for line in handle:
                line = line.strip()
                if line:
                    try:
                        rows.append(json.loads(line))
                    except json.JSONDecodeError:
                        continue
    except FileNotFoundError:
        pass
    return rows


# --- admin unlock -----------------------------------------------------------
#
# A .txt dropped into the databook uploader whose content is today's code
# unlocks the usage download for that browser session. The code is derived
# from a secret in config.yml (admin.unlock_secret, gitignored) and the date,
# so it changes daily and cannot be worked out from the source. No secret
# configured -> the feature is off and a .txt is rejected like any other
# unsupported file. Yesterday's and tomorrow's codes are accepted too: the
# server's clock may sit in another time zone from the person typing.

def admin_secret(config: Optional[Dict[str, Any]] = None) -> str:
    if config is None:
        try:
            from .financial_common import load_yaml_file, package_file_path
            config = load_yaml_file(package_file_path("config.yml"))
        except Exception:
            config = {}
    block = (config or {}).get("admin") or {}
    return str(block.get("unlock_secret") or os.environ.get("FDD_ADMIN_SECRET") or "").strip()


def admin_code(secret: str, day: Optional[datetime] = None) -> str:
    import hashlib

    day = day or datetime.now()
    return hashlib.sha256(f"{secret}|{day:%Y%m%d}".encode("utf-8")).hexdigest()[:10].upper()


def admin_code_matches(text: str, secret: Optional[str] = None) -> bool:
    import hmac
    from datetime import timedelta

    secret = admin_secret() if secret is None else secret
    candidate = str(text or "").strip().upper()
    if not secret or not candidate:
        return False
    today = datetime.now()
    return any(hmac.compare_digest(candidate, admin_code(secret, today + timedelta(days=offset)))
               for offset in (-1, 0, 1))


# --- reports ----------------------------------------------------------------

def build_usage_report() -> bytes:
    """An .xlsx: every call, every event, and totals by upload and by day."""
    import io

    import pandas as pd

    calls = pd.DataFrame(read_jsonl(USAGE_PATH))
    events = pd.DataFrame(read_jsonl(ACTIVITY_PATH))
    sheets: Dict[str, Any] = {}
    if not calls.empty:
        for col in ("prompt_tokens", "completion_tokens", "total_tokens", "cached_prompt_tokens",
                    "estimated_prompt_tokens", "estimated_output_tokens"):
            if col not in calls.columns:
                calls[col] = None
            calls[col] = pd.to_numeric(calls[col], errors="coerce")
        # A provider that reports no usage (a streaming local server, a failed
        # call) still cost something: count the estimate, and say it is one.
        calls["tokens_estimated"] = calls["prompt_tokens"].isna() & calls["estimated_prompt_tokens"].notna()
        calls["prompt_tokens"] = calls["prompt_tokens"].fillna(calls["estimated_prompt_tokens"])
        calls["completion_tokens"] = calls["completion_tokens"].fillna(calls["estimated_output_tokens"])
        calls["total_tokens"] = calls["total_tokens"].fillna(
            calls["prompt_tokens"].fillna(0) + calls["completion_tokens"].fillna(0))
        # A failed call is not billed as a rule: its estimate stays in
        # estimated_prompt_tokens for reference and is kept out of the sums.
        failed = calls["ok"].astype(str).str.lower().isin(["false", "0"]) if "ok" in calls.columns else False
        for col in ("prompt_tokens", "completion_tokens", "total_tokens", "cached_prompt_tokens"):
            calls[col] = calls[col].fillna(0).astype(int)
            calls.loc[failed, col] = 0
        calls.loc[failed, "tokens_estimated"] = False
        for col in ("file", "session", "model", "provider", "ok"):
            if col not in calls.columns:
                calls[col] = ""
        calls["day"] = calls["time"].astype(str).str[:10]
        calls["calls"] = 1
        sums = ["calls", "prompt_tokens", "completion_tokens", "total_tokens", "cached_prompt_tokens"]
        by_file = (calls.fillna({"file": "", "session": ""})
                   .groupby(["file", "session", "provider", "model"], dropna=False)
                   .agg(first_call=("time", "min"), last_call=("time", "max"),
                        **{c: (c, "sum") for c in sums}, failed=("ok", lambda s: int((s == False).sum())),  # noqa: E712
                        estimated_calls=("tokens_estimated", "sum"))
                   .reset_index().sort_values("first_call"))
        by_day = (calls.groupby(["day", "provider", "model"], dropna=False)[sums].sum()
                  .reset_index().sort_values("day"))
        sheets["by_upload"] = by_file
        sheets["by_day"] = by_day
        sheets["calls"] = calls.drop(columns=["calls"])
    else:
        sheets["calls"] = pd.DataFrame({"note": ["no LLM call recorded yet"]})
    sheets["events"] = events if not events.empty else pd.DataFrame({"note": ["no activity recorded yet"]})
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=name[:31], index=False)
    return buffer.getvalue()


def build_logs_zip() -> bytes:
    """Every file under fdd_utils/logs, zipped. Holds prompts and client figures."""
    import io
    import zipfile

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        if LOG_DIR.exists():
            for path in sorted(LOG_DIR.rglob("*")):
                if path.is_file():
                    try:
                        archive.write(path, arcname=str(path.relative_to(LOG_DIR.parent)))
                    except Exception:
                        continue
    return buffer.getvalue()

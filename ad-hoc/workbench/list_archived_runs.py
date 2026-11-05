#!/usr/bin/env python3
"""Inventory fdd_utils/logs/run_*/ — which archived runs are worth replaying, and which are dead.

The archive is the only free corpus this project has: 255 folders of real
pipeline output, no tokens to re-read. But nothing in a run folder says whether
the run WORKED. A run whose every LLM call returned 400 finishes normally, writes
a full results.yml, and ships deterministic fallback bullets as the deck's entire
commentary — indistinguishable, from the outside, from a clean run. Five such runs
are in this archive (the defect M8.1 fixes in production; this tool is how you see
them from the outside).

So each run gets a health verdict read out of processing.log:

  HEALTHY   every stage made LLM calls, no fallback bullets
  DEGRADED  calls succeeded but some accounts fell back or the run was partial
  DEAD      zero successful LLM calls — the commentary is entirely deterministic
            fallback text. Never replay one of these as a quality baseline.
  NO-LOG / NO-RESULTS   the folder is missing what a replay needs

Read from the log, because it is the only place the outcome survives:
  "[<Stage>] Processed: <key> | Duration: ..."  = one successful call
    (ai/logging.py:log_agent_complete, reached only after a response comes back)
  "AI unavailable after retries; using deterministic data-only fallback"
    = one account shipped on a fallback bullet (pipeline.py:745)
  "circuit breaker OPEN"  = a call the breaker refused to even attempt

Usage:
    python ad-hoc/workbench/list_archived_runs.py                  # all runs
    python ad-hoc/workbench/list_archived_runs.py --dead           # only the dead ones
    python ad-hoc/workbench/list_archived_runs.py --min-accounts 20 --healthy
"""

from __future__ import annotations

import argparse
import re
import sys
from collections import Counter
from pathlib import Path
from typing import Any, Dict, List

import yaml

REPO_ROOT = Path(__file__).resolve().parents[2]
LOGS_DIR = REPO_ROOT / "fdd_utils" / "logs"

_SUCCESS = re.compile(r"\] Processed: .+ \| Duration:")
_FALLBACK = re.compile(r"using deterministic data-only fallback")
_BREAKER = re.compile(r"circuit breaker OPEN")
_FAILED_CALL = re.compile(r"AI call attempt \d+ failed")
_STARTED = re.compile(r"Starting FDD pipeline with (\d+) items \| model=(\S+) \| language=(\S+)")
_COMPLETED = re.compile(r"=== Completed AI processing run")
_SUMMARY = re.compile(r"Summary: (\d+) items, ([\d.]+)s, (\d+) tokens")
_STAGE = re.compile(r"\[(Generator|Auditor|Validator|Refiner)\] Processed:")


def scan_log(path: Path) -> Dict[str, Any]:
    if not path.exists():
        return {"has_log": False}
    text = path.read_text(encoding="utf-8", errors="replace")
    started = _STARTED.search(text)
    summary = _SUMMARY.search(text)
    return {
        "has_log": True,
        "successful_calls": len(_SUCCESS.findall(text)),
        "fallback_bullets": len(_FALLBACK.findall(text)),
        "breaker_skips": len(_BREAKER.findall(text)),
        "failed_attempts": len(_FAILED_CALL.findall(text)),
        "planned_items": int(started.group(1)) if started else None,
        "model": started.group(2) if started else "?",
        "language": started.group(3) if started else "?",
        "completed": bool(_COMPLETED.search(text)),
        "tokens": int(summary.group(3)) if summary else 0,
        "stages_with_calls": sorted(set(_STAGE.findall(text))),
    }


def account_count(path: Path) -> int:
    """Accounts in results.yml. Parsed rather than counted from the log because a
    dead run still writes a full results.yml — the count says how much text is
    there, not whether any of it came from a model."""
    if not path.exists():
        return -1
    try:
        data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    except Exception:
        return -1
    return sum(1 for v in data.values() if isinstance(v, dict))


def verdict(info: Dict[str, Any], n_accounts: int, has_results: bool) -> str:
    if not has_results:
        return "NO-RESULTS"
    if not info.get("has_log"):
        return "NO-LOG"
    if info["successful_calls"] == 0:
        return "DEAD"
    planned = info.get("planned_items") or n_accounts
    # A healthy run makes calls in every stage of the active sequence. The
    # Validator is SELECTIVE (it only runs on accounts asserting a causal claim),
    # so a run can legitimately have fewer Validator calls than accounts — but
    # zero calls in a stage that was entered means the stage failed, not that it
    # was skipped.
    if info["fallback_bullets"] or info["breaker_skips"]:
        return "DEGRADED"
    if not info.get("completed"):
        return "DEGRADED"
    if planned and n_accounts and n_accounts < planned:
        return "DEGRADED"
    return "HEALTHY"


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--logs", default=str(LOGS_DIR), help="logs directory (default fdd_utils/logs)")
    ap.add_argument("--dead", action="store_true", help="show only DEAD runs")
    ap.add_argument("--healthy", action="store_true", help="show only HEALTHY runs")
    ap.add_argument("--min-accounts", type=int, default=0, help="hide runs with fewer accounts")
    ap.add_argument("--quiet", action="store_true", help="totals only, no per-run rows")
    args = ap.parse_args()

    logs = Path(args.logs)
    if not logs.is_dir():
        sys.exit(f"❌ No such logs directory: {logs}")

    rows: List[Dict[str, Any]] = []
    for run_dir in sorted(logs.glob("run_*")):
        if not run_dir.is_dir():
            continue
        results_path = run_dir / "results.yml"
        info = scan_log(run_dir / "processing.log")
        n = account_count(results_path)
        rows.append({
            "run": run_dir.name,
            "accounts": max(n, 0),
            "results": results_path.exists(),
            "data": (run_dir / "data.yml").exists(),
            "calls": info.get("successful_calls", 0),
            "fallbacks": info.get("fallback_bullets", 0),
            "breaker": info.get("breaker_skips", 0),
            "failed": info.get("failed_attempts", 0),
            "model": info.get("model", "?"),
            "lang": info.get("language", "?"),
            "tokens": info.get("tokens", 0),
            "verdict": verdict(info, max(n, 0), results_path.exists()),
        })

    shown = rows
    if args.dead:
        shown = [r for r in shown if r["verdict"] == "DEAD"]
    if args.healthy:
        shown = [r for r in shown if r["verdict"] == "HEALTHY"]
    shown = [r for r in shown if r["accounts"] >= args.min_accounts]

    if not args.quiet:
        header = (f"{'run':<24} {'acct':>4} {'res':>3} {'dat':>3} {'calls':>5} {'fbk':>4} "
                  f"{'brk':>4} {'fail':>5} {'model':<10} {'lang':<4} {'tokens':>8}  verdict")
        print(header)
        print("-" * len(header))
        for r in shown:
            print(f"{r['run']:<24} {r['accounts']:>4} {'Y' if r['results'] else '-':>3} "
                  f"{'Y' if r['data'] else '-':>3} {r['calls']:>5} {r['fallbacks']:>4} "
                  f"{r['breaker']:>4} {r['failed']:>5} {str(r['model'])[:10]:<10} "
                  f"{str(r['lang'])[:4]:<4} {r['tokens']:>8}  {r['verdict']}")

    print(f"\n--- {len(rows)} run folder(s) in {logs} ---")
    for name, count in Counter(r["verdict"] for r in rows).most_common():
        print(f"  {name:<12} {count:>4}")
    dead = [r for r in rows if r["verdict"] == "DEAD"]
    if dead:
        print(f"\n  DEAD runs shipped a deck built entirely from deterministic fallback bullets:")
        for r in dead:
            print(f"    {r['run']}  {r['accounts']} accounts, {r['fallbacks']} fallback bullets, "
                  f"{r['failed']} failed attempts, {r['breaker']} breaker skips, model={r['model']}")
    best = [r for r in rows if r["verdict"] == "HEALTHY" and r["results"]]
    if best:
        best.sort(key=lambda r: (r["accounts"], r["run"]), reverse=True)
        print(f"\n  Widest HEALTHY runs (best replay targets):")
        for r in best[:5]:
            print(f"    {r['run']}  {r['accounts']} accounts, {r['calls']} calls, "
                  f"model={r['model']}, lang={r['lang']}")


if __name__ == "__main__":
    main()

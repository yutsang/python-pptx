"""Concurrency smoke test for the KPMG Workbench provider.

The real pipeline shares ONE AIClient instance across a ThreadPoolExecutor
(see run_agent_stage in fdd_utils/ai.py — max_workers=2 for cloud providers,
4 for local). This script reproduces that exact pattern: one client, many
threads, and checks for:
  1. every request completes without exception (no silent drops)
  2. no 429 / rate-limit errors
  3. no cross-thread response mixing (each reply echoes its own unique marker)
  4. latency distribution (min/max/avg/p95) so you can judge whether
     max_workers can safely be raised above the current default of 2

Usage:
    python test_workbench_concurrency.py                       # 2 workers x 10 requests (matches prod default)
    python test_workbench_concurrency.py --workers 4 --requests 20
    python test_workbench_concurrency.py --model gpt-5-4-2026-03-05-gs-sdc
    python test_workbench_concurrency.py --sequential            # single-threaded (workers=1), same call path

Note on max_tokens: GPT-5.x is a REASONING model — max_completion_tokens is a
shared budget for internal reasoning tokens AND the visible answer. A small
budget (e.g. 32) can be entirely consumed by reasoning with nothing left for
the answer, producing an empty (but successful, no-exception) response. This
is NOT a thread-safety bug — do not mistake it for cross-thread mixing, which
is specifically "got a DIFFERENT marker than the one this thread sent".
"""
from __future__ import annotations

import argparse
import statistics
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

from fdd_utils.ai import AIClient, is_provider_ready, load_yaml_config, get_default_config_path

# Reasoning models need real headroom: internal reasoning tokens are billed
# against the same budget as the visible answer. 32-64 was observed to starve
# the answer entirely (empty content, still a 200/no-exception response).
DEFAULT_MAX_TOKENS = 512


def _call_one(client: AIClient, marker: str, max_tokens: int) -> dict:
    start = time.time()
    try:
        response = client.get_response(
            user_prompt=f"Reply with exactly this token and nothing else: {marker}",
            system_prompt="You are a terse echo assistant.",
            temperature=0,
            max_tokens=max_tokens,
        )
        elapsed = time.time() - start
        content = str(response.get("content") or "")
        other_markers = [m for m in _ALL_MARKERS if m != marker and m in content]
        return {
            "marker": marker,
            "ok": True,
            "elapsed": elapsed,
            "content": content,
            "marker_present": marker in content,
            "empty": not content.strip(),
            # Genuine cross-thread mixing = got SOMEONE ELSE'S marker, not "got nothing".
            "mixed": bool(other_markers),
            "other_markers": other_markers,
            "completion_tokens": response.get("completion_tokens"),
            "error": None,
        }
    except Exception as exc:
        elapsed = time.time() - start
        msg = str(exc)
        return {
            "marker": marker,
            "ok": False,
            "elapsed": elapsed,
            "content": "",
            "marker_present": False,
            "empty": True,
            "mixed": False,
            "other_markers": [],
            "completion_tokens": None,
            "error": msg,
            "rate_limited": "429" in msg or "rate limit" in msg.lower(),
        }


_ALL_MARKERS: list[str] = []


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--workers", type=int, default=2, help="concurrent threads (prod default for cloud providers is 2)")
    ap.add_argument("--requests", type=int, default=10, help="total requests to fire")
    ap.add_argument("--model", default=None, help="workbench model id; default = config's chat_model (GPT-5.5)")
    ap.add_argument("--max-tokens", type=int, default=DEFAULT_MAX_TOKENS,
                    help=f"max_completion_tokens per call (default {DEFAULT_MAX_TOKENS} — "
                         f"reasoning models need headroom beyond the visible answer)")
    ap.add_argument("--sequential", action="store_true",
                    help="run one request at a time (workers=1) — use this to rule out "
                         "concurrency entirely if you're still unsure about thread-safety")
    args = ap.parse_args()
    workers = 1 if args.sequential else args.workers

    config = load_yaml_config(get_default_config_path())
    if not is_provider_ready(config, "workbench"):
        print("❌ workbench provider is not configured. Set fdd_utils/config.yml -> workbench.api_key etc.")
        return 1

    print(f"Building ONE shared AIClient (mirrors how the real pipeline shares a "
          f"client across threads) — model={args.model or '(config default)'}")
    client = AIClient(
        model_type="workbench",
        agent_name="2_Auditor",  # low temperature agent config; irrelevant to this test
        language="Eng",
        model_name=args.model,
    )
    print(f"  resolved chat_model: {client.config_details.get('chat_model')}")
    mode = "SEQUENTIAL (1 worker)" if args.sequential else f"{workers} worker threads"
    print(f"\nFiring {args.requests} requests, {mode}, max_tokens={args.max_tokens}...\n")

    global _ALL_MARKERS
    _ALL_MARKERS = [f"MARK-{i:04d}" for i in range(args.requests)]
    results: list[dict] = []
    overall_start = time.time()
    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = {executor.submit(_call_one, client, m, args.max_tokens): m for m in _ALL_MARKERS}
        for future in as_completed(futures):
            r = future.result()
            results.append(r)
            if not r["ok"]:
                status = "❌"
                detail = r["error"]
            elif r["mixed"]:
                status = "🔴"  # genuine cross-thread mixing — the real bug this test guards against
                detail = f"got marker(s) from ANOTHER thread: {r['other_markers']}"
            elif r["empty"]:
                status = "⚠️ "
                detail = f"empty response (reasoning budget likely exhausted; completion_tokens={r['completion_tokens']})"
            elif r["marker_present"]:
                status = "✅"
                detail = "OK"
            else:
                status = "⚠️ "
                detail = f"non-empty but marker not found: {r['content'][:80]!r}"
            print(f"  {status} {r['marker']}  {r['elapsed']:.2f}s  {detail}")
    overall_elapsed = time.time() - overall_start

    ok = [r for r in results if r["ok"]]
    failed = [r for r in results if not r["ok"]]
    rate_limited = [r for r in failed if r.get("rate_limited")]
    mixed = [r for r in ok if r["mixed"]]          # genuine bug
    empty = [r for r in ok if r["empty"]]           # token-budget symptom, not a bug
    latencies = [r["elapsed"] for r in ok]

    print(f"\n{'='*70}\n  SUMMARY\n{'='*70}")
    print(f"  total requests        : {len(results)}")
    print(f"  succeeded             : {len(ok)}")
    print(f"  failed                : {len(failed)}  (rate-limited: {len(rate_limited)})")
    print(f"  empty responses       : {len(empty)}  (token-budget symptom — raise --max-tokens if >0)")
    print(f"  cross-thread mixing   : {len(mixed)}  (THE ACTUAL BUG THIS TEST GUARDS AGAINST — must be 0)")
    print(f"  wall time             : {overall_elapsed:.2f}s  "
          f"({len(results) / overall_elapsed:.2f} req/s effective throughput)")
    if latencies:
        sorted_lat = sorted(latencies)
        p95_idx = max(0, int(len(sorted_lat) * 0.95) - 1)
        print(f"  latency min/avg/max   : {min(latencies):.2f}s / {statistics.mean(latencies):.2f}s / {max(latencies):.2f}s")
        print(f"  latency p95           : {sorted_lat[p95_idx]:.2f}s")

    if failed:
        print("\n  Failures:")
        for r in failed:
            print(f"    {r['marker']}: {r['error']}")
    if empty:
        print(f"\n  ⚠️ {len(empty)} empty response(s) — likely max_tokens too small for this reasoning")
        print(f"     model (reasoning tokens + answer share one budget). Try: --max-tokens 1024")
    if mixed:
        print("\n  🔴 GENUINE cross-thread mixing (client/gateway returned another request's content):")
        for r in mixed:
            print(f"    {r['marker']} -> got {r['other_markers']}: {r['content'][:100]!r}")

    # PASS/FAIL is based ONLY on real failures: exceptions or genuine mixing.
    # Empty responses are reported but don't fail the run — they mean "retune
    # max_tokens", not "concurrency is unsafe".
    passed = not failed and not mixed
    print(f"\n  {'✅ PASS' if passed else '❌ FAIL'} — "
          f"{'safe to use at this concurrency' if passed else 'investigate failures/mixing above before deploying'}")
    if passed and empty:
        print(f"  (note: {len(empty)} empty response(s) reported above — retune --max-tokens, this did not fail the test)")
    return 0 if passed else 1


if __name__ == "__main__":
    raise SystemExit(main())

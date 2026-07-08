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
"""
from __future__ import annotations

import argparse
import statistics
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

from fdd_utils.ai import AIClient, is_provider_ready, load_yaml_config, get_default_config_path


def _call_one(client: AIClient, marker: str) -> dict:
    start = time.time()
    try:
        response = client.get_response(
            user_prompt=f"Reply with exactly this token and nothing else: {marker}",
            system_prompt="You are a terse echo assistant.",
            temperature=0,
            max_tokens=32,
        )
        elapsed = time.time() - start
        content = str(response.get("content") or "")
        return {
            "marker": marker,
            "ok": True,
            "elapsed": elapsed,
            "content": content,
            "marker_present": marker in content,
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
            "error": msg,
            "rate_limited": "429" in msg or "rate limit" in msg.lower(),
        }


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--workers", type=int, default=2, help="concurrent threads (prod default for cloud providers is 2)")
    ap.add_argument("--requests", type=int, default=10, help="total requests to fire")
    ap.add_argument("--model", default=None, help="workbench model id; default = config's chat_model (GPT-5.5)")
    args = ap.parse_args()

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
    print(f"\nFiring {args.requests} requests across {args.workers} worker threads...\n")

    markers = [f"MARK-{i:04d}" for i in range(args.requests)]
    results: list[dict] = []
    overall_start = time.time()
    with ThreadPoolExecutor(max_workers=args.workers) as executor:
        futures = {executor.submit(_call_one, client, m): m for m in markers}
        for future in as_completed(futures):
            r = future.result()
            results.append(r)
            status = "✅" if r["ok"] and r["marker_present"] else ("⚠️ " if r["ok"] else "❌")
            print(f"  {status} {r['marker']}  {r['elapsed']:.2f}s  "
                  f"{'OK' if r['ok'] else r['error']}")
    overall_elapsed = time.time() - overall_start

    ok = [r for r in results if r["ok"]]
    failed = [r for r in results if not r["ok"]]
    rate_limited = [r for r in failed if r.get("rate_limited")]
    mixed = [r for r in ok if not r["marker_present"]]
    latencies = [r["elapsed"] for r in ok]

    print(f"\n{'='*70}\n  SUMMARY\n{'='*70}")
    print(f"  total requests      : {len(results)}")
    print(f"  succeeded           : {len(ok)}")
    print(f"  failed              : {len(failed)}  (rate-limited: {len(rate_limited)})")
    print(f"  cross-thread mixing : {len(mixed)}  (should be 0 — each reply must echo its own marker)")
    print(f"  wall time           : {overall_elapsed:.2f}s  "
          f"({len(results) / overall_elapsed:.2f} req/s effective throughput)")
    if latencies:
        sorted_lat = sorted(latencies)
        p95_idx = max(0, int(len(sorted_lat) * 0.95) - 1)
        print(f"  latency min/avg/max : {min(latencies):.2f}s / {statistics.mean(latencies):.2f}s / {max(latencies):.2f}s")
        print(f"  latency p95         : {sorted_lat[p95_idx]:.2f}s")

    if failed:
        print("\n  Failures:")
        for r in failed:
            print(f"    {r['marker']}: {r['error']}")
    if mixed:
        print("\n  ⚠️ Cross-thread mixing detected (client may not be thread-safe, or the gateway")
        print("     is caching/reordering responses):")
        for r in mixed:
            print(f"    {r['marker']} -> {r['content'][:100]!r}")

    passed = not failed and not mixed
    print(f"\n  {'✅ PASS' if passed else '❌ FAIL'} — "
          f"{'safe to use at this concurrency' if passed else 'DO NOT deploy at this concurrency until fixed'}")
    return 0 if passed else 1


if __name__ == "__main__":
    raise SystemExit(main())

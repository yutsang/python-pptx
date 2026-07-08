"""Connectivity smoke test for the KPMG Workbench provider (GPT-5.5 / GPT-5.4).

Exercises the REAL AIClient wrapper (not a bare openai-SDK call), so this also
validates our header plumbing, the reasoning_effort defensive retry, and
response-content extraction — not just raw network reachability.

Prerequisites:
    fdd_utils/config.yml -> workbench.api_key must be set (gitignored, local only).

Usage:
    python test_workbench_connectivity.py            # tests every configured model
    python test_workbench_connectivity.py --model gpt-5-5-2026-04-24-gs-sdc
"""
from __future__ import annotations

import argparse
import sys
import time

from fdd_utils.ai import AIClient, WORKBENCH_AVAILABLE_MODELS, is_provider_ready, load_yaml_config, get_default_config_path


def _models_to_test(explicit: str | None) -> list[str]:
    if explicit:
        return [explicit]
    config = load_yaml_config(get_default_config_path())
    models = (config.get("workbench") or {}).get("available_models") or WORKBENCH_AVAILABLE_MODELS
    return [m.get("id") if isinstance(m, dict) else str(m) for m in models]


def _run_one(model_id: str) -> bool:
    print(f"\n{'='*70}\n  {model_id}\n{'='*70}")
    try:
        client = AIClient(
            model_type="workbench",
            agent_name="1_Generator",  # any agent; only used for temperature/max_tokens defaults
            language="Eng",
            model_name=model_id,
        )
    except Exception as exc:
        print(f"  ❌ AIClient init failed: {exc}")
        return False

    print(f"  resolved model_type : {client.model_type}")
    print(f"  resolved chat_model : {client.config_details.get('chat_model')}")
    print(f"  api_base            : {client.config_details.get('api_base')}")

    start = time.time()
    try:
        response = client.get_response(
            user_prompt="hi",
            system_prompt="You are a helpful assistant.",
            temperature=1,
            max_tokens=512,  # reasoning models spend part of this budget on hidden
                             # reasoning tokens before the visible answer; 64 was too
                             # tight and produced empty content that looked like a failure
        )
    except Exception as exc:
        print(f"  ❌ get_response raised: {type(exc).__name__}: {exc}")
        return False
    elapsed = time.time() - start

    content = str(response.get("content") or "").strip()
    if not content:
        print(f"  ⚠️  Empty response content after {elapsed:.2f}s "
              f"(completion_tokens={response.get('completion_tokens')}) — likely max_tokens "
              f"still too small for this model's reasoning overhead, try raising it further.")
        print(f"     Full payload: {response}")
        return False

    print(f"  ✅ OK in {elapsed:.2f}s")
    print(f"  content: {content[:200]!r}")
    print(f"  prompt_tokens={response.get('prompt_tokens')}  "
          f"completion_tokens={response.get('completion_tokens')}  "
          f"total_tokens={response.get('total_tokens')}")
    return True


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--model", default=None, help="test only this model id")
    args = ap.parse_args()

    config = load_yaml_config(get_default_config_path())
    if not is_provider_ready(config, "workbench"):
        print("❌ workbench provider is not configured.")
        print("   Set fdd_utils/config.yml -> workbench.api_key, api_base, api_version, chat_model.")
        return 1

    models = _models_to_test(args.model)
    if not models:
        print("❌ No workbench models found in config (workbench.available_models).")
        return 1

    results = {m: _run_one(m) for m in models}

    print(f"\n{'='*70}\n  SUMMARY\n{'='*70}")
    for model_id, ok in results.items():
        print(f"  {'✅ PASS' if ok else '❌ FAIL'}  {model_id}")

    return 0 if all(results.values()) else 1


if __name__ == "__main__":
    raise SystemExit(main())

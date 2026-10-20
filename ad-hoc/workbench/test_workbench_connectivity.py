"""Connectivity smoke test for every model the UI lists, across providers.

The models tested are exactly the ones the Streamlit model dropdown offers:
list_model_choices() over config.yml, i.e. every chat model whose provider
and entry are not `listed: false` (Workbench, DashScope, Local). Each one goes
through the REAL AIClient wrapper (not a bare SDK call), so header plumbing,
the reasoning_effort retry and response-content extraction are exercised too,
not just network reachability.

Embedding / rerank models are not chat models and never appear in the
dropdown; test_dashscope_connectivity.py still covers those endpoints.

Prerequisites:
    fdd_utils/config.yml -> each provider's api_key / api_base (gitignored, local only).

Usage:
    python ad-hoc/workbench/test_workbench_connectivity.py                  # every listed model
    python ad-hoc/workbench/test_workbench_connectivity.py --provider dashscope
    python ad-hoc/workbench/test_workbench_connectivity.py --model <model-id>
    python ad-hoc/workbench/test_workbench_connectivity.py --include-unlisted   # also listed: false ones
"""
from __future__ import annotations

# moved into ad-hoc/ -- put the repo root back on sys.path so
# `import fdd_utils...` still resolves when run from anywhere.
import sys as _sys
from pathlib import Path as _Path
_sys.path.insert(0, str(_Path(__file__).resolve().parents[2]))

import argparse
import time
from typing import Any, Dict, Optional

from fdd_utils.ai import AIClient, get_default_config_path, list_model_choices, load_yaml_config


def _run_one(choice: Dict[str, Any]) -> Optional[bool]:
    """True/False for pass/fail, None when the provider is not configured."""
    model_type, model_name = choice["model_type"], choice["model_name"]
    print(f"\n{'=' * 70}\n  {choice['label']}   [{choice['key']}]\n{'=' * 70}")
    if not choice.get("ready"):
        print(f"  ⏭️  skipped: provider '{model_type}' is not configured in config.yml "
              f"(api_key / api_base missing)")
        return None
    try:
        client = AIClient(
            model_type=model_type,
            agent_name="1_Generator",  # any agent; only used for temperature/max_tokens defaults
            language="Eng",
            model_name=model_name,
        )
    except Exception as exc:
        print(f"  ❌ AIClient init failed: {type(exc).__name__}: {exc}")
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
        print(f"  ❌ get_response raised after {time.time() - start:.2f}s: {type(exc).__name__}: {exc}")
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
    ap = argparse.ArgumentParser(description="Test connectivity of every model the UI lists")
    ap.add_argument("--provider", default=None, help="only this provider: workbench | dashscope | local")
    ap.add_argument("--model", default=None, help="only this model id")
    ap.add_argument("--include-unlisted", action="store_true",
                    help="also test models config.yml hides from the UI (listed: false)")
    args = ap.parse_args()

    config = load_yaml_config(get_default_config_path())
    choices = list_model_choices(config, include_unlisted=args.include_unlisted)
    if args.provider:
        choices = [c for c in choices if c["model_type"] == args.provider]
    if args.model:
        choices = [c for c in choices if c["model_name"] == args.model]
    if not choices:
        print("❌ No model matched. Listed choices in config.yml:")
        for c in list_model_choices(config, include_unlisted=True):
            print(f"   {c['key']}  listed={c['listed']}  ready={c['ready']}")
        return 1

    print(f"Testing {len(choices)} model(s) from {get_default_config_path()}"
          + ("  (including unlisted)" if args.include_unlisted else "  (listed only)"))
    results = {c["key"]: (c, _run_one(c)) for c in choices}

    print(f"\n{'=' * 70}\n  SUMMARY\n{'=' * 70}")
    for key, (choice, ok) in results.items():
        mark = "⏭️  SKIP" if ok is None else ("✅ PASS" if ok else "❌ FAIL")
        hidden = "" if choice.get("listed", True) else "  (unlisted)"
        print(f"  {mark}  {key}{hidden}")

    return 0 if all(ok is not False for _c, ok in results.values()) else 1


if __name__ == "__main__":
    raise SystemExit(main())

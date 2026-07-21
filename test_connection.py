#!/usr/bin/env python3
"""Standalone, isolated connectivity test for the configured AI provider
(default: 'workbench', KPMG's internal gateway).

Run this ON THE MACHINE where config.yml has the real api_key/api_base --
this repo's checked-in config.yml is a per-machine, gitignored file, so this
script only works with your actual local config.yml, not a copy of this file
elsewhere.

Makes exactly ONE minimal chat-completion call via the same AIClient class
production uses (same headers, same base_url, same api_version), plus a
separate raw HTTP reachability check against the bare host -- this lets you
tell apart:
  (a) can't reach the host at all (DNS/VPN/firewall) vs.
  (b) host is reachable but this specific request gets blocked (WAF rule,
      subscription key issue, gateway-side block).

Usage:
    python test_connection.py                  # test the 'workbench' provider
    python test_connection.py --provider openai # test a different provider
    python test_connection.py --agent agent_2   # use a different agent's config slice
    python test_connection.py --raw             # dump the full raw exception/response
"""
import argparse
import sys
import time
from urllib.parse import urlparse


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--provider", default="workbench", help="model_type from config.yml (default: workbench)")
    ap.add_argument("--agent", default="agent_1", help="agent_name whose config slice to use (default: agent_1)")
    ap.add_argument("--raw", action="store_true", help="print the full raw exception/response, not just a summary")
    args = ap.parse_args()

    print("=" * 78)
    print(f"  AI PROVIDER CONNECTIVITY TEST -- provider={args.provider!r}, agent={args.agent!r}")
    print("=" * 78)

    try:
        from fdd_utils.ai import AIClient
    except Exception as e:
        print(f"❌ FAILED to import fdd_utils.ai: {type(e).__name__}: {e}")
        return 1

    # --- Step 1: build the client (loads config.yml, validates keys) -------
    try:
        client = AIClient(model_type=args.provider, agent_name=args.agent, language="Eng", use_heuristic=False)
    except Exception as e:
        print(f"❌ FAILED to initialize AIClient (config/validation problem, not a network issue):")
        print(f"   {type(e).__name__}: {e}")
        return 1

    api_base = client.config_details.get("api_base", "<not set>")
    api_key = client.config_details.get("api_key", "") or ""
    redacted_key = (api_key[:4] + "..." + api_key[-4:]) if len(api_key) > 8 else ("<empty>" if not api_key else "<short>")
    print(f"  api_base: {api_base}")
    print(f"  api_key:  {redacted_key}")
    print(f"  model:    {client.model}")

    # --- Step 2: bare reachability check (no auth, no payload) -------------
    print("\n--- Step 1/2: raw host reachability (GET, no auth) ---")
    try:
        import httpx
        parsed = urlparse(api_base)
        host_root = f"{parsed.scheme}://{parsed.netloc}/"
        t0 = time.time()
        with httpx.Client(verify=False, timeout=httpx.Timeout(15.0, connect=10.0)) as h:
            resp = h.get(host_root)
        dt = time.time() - t0
        print(f"  GET {host_root} -> HTTP {resp.status_code} in {dt:.2f}s")
        body_preview = resp.text[:300].replace("\n", " ")
        if "request is blocked" in resp.text.lower() or "waf" in resp.text.lower():
            print(f"  ⚠️  Response body looks like a WAF block page: {body_preview!r}")
        elif args.raw:
            print(f"  body preview: {body_preview!r}")
    except Exception as e:
        print(f"  ❌ Could not even reach the host: {type(e).__name__}: {e}")
        print("  -> This points to network/VPN/DNS/firewall, not the gateway's WAF or your API key.")

    # --- Step 3: one real minimal chat completion via the production path --
    print("\n--- Step 2/2: one minimal chat completion (same code path as production) ---")
    try:
        t0 = time.time()
        result = client.get_response(
            user_prompt="Reply with exactly one word: OK",
            system_prompt="You are a connectivity test. Reply with exactly one word: OK",
            max_tokens=10,
        )
        dt = time.time() - t0
        content = (result or {}).get("content", "")
        print(f"  ✅ SUCCESS in {dt:.2f}s -- response content: {content!r}")
        print("\n  Gateway is reachable and answering normally. If your earlier batch run still")
        print("  failed, it was likely a transient outage/rate-limit, not a persistent block --")
        print("  try re-running the batch now.")
        return 0
    except Exception as e:
        dt = time.time() - t0
        print(f"  ❌ FAILED after {dt:.2f}s: {type(e).__name__}")
        # openai SDK errors usually carry the raw HTTP body on .response.text or .body
        body = None
        resp_obj = getattr(e, "response", None)
        if resp_obj is not None:
            body = getattr(resp_obj, "text", None)
            status = getattr(resp_obj, "status_code", None)
            print(f"  HTTP status: {status}")
        if body is None:
            body = getattr(e, "body", None)
        if body:
            body_str = str(body)
            if "request is blocked" in body_str.lower():
                print("  -> Confirmed: this specific call IS being blocked by the gateway's WAF")
                print("     (same block page as your batch run). Not an app-level/code issue --")
                print("     escalate to whoever manages the KPMG Workbench subscription/gateway.")
            print(f"  body: {body_str[:500] if not args.raw else body_str!r}")
        else:
            print(f"  {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())

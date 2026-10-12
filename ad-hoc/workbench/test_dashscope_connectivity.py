"""Standalone connectivity probe for the new DashScope / gateway model set.

Reads api_base / api_key / auth / available_models from
fdd_utils/config.yml -> dashscope: (gitignored). CLI flags only override.

Does NOT go through AIClient — so it can also exercise embedding + rerank
endpoints that the FDD chat pipeline never calls.

Usage:
    # 1) fill dashscope.api_base + dashscope.api_key in fdd_utils/config.yml
    # 2) run (no flags needed):
    python ad-hoc/workbench/test_dashscope_connectivity.py

    # optional filters / one-off overrides
    python ad-hoc/workbench/test_dashscope_connectivity.py --kind chat
    python ad-hoc/workbench/test_dashscope_connectivity.py --model dashscope_qwen3.7-max
    python ad-hoc/workbench/test_dashscope_connectivity.py --auth workbench

Prices: user-supplied in/out; blended cost assumes I:O = 1:5 (in + 5*out).
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import time
from pathlib import Path
from typing import Any

_REPO_ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO_ROOT))

# Fallback catalog when config.yml has no dashscope.available_models.
# `*_deepseek-v4-pro-0813` is treated as dashscope_deepseek-v4-pro-0813.
_DEFAULT_MODELS: list[dict[str, Any]] = [
    {
        "id": "dashscope_qwen3.7-max",
        "kind": "chat",
        "label": "Qwen3.7-Max",
        "fdd_role": "Generator / Auditor / Validator (best Chi quality)",
        "capability_rank": 1,
        "price_in": 12.0,
        "price_out": 36.0,
    },
    {
        "id": "dashscope_deepseek-v4-pro-0813",
        "kind": "chat",
        "label": "DeepSeek-V4-Pro",
        "fdd_role": "Generator / Auditor / Validator (strong reasoning)",
        "capability_rank": 2,
        "price_in": 12.0,
        "price_out": 24.0,
    },
    {
        "id": "kimi-k2.6_chat",
        "kind": "chat",
        "label": "Kimi-K2.6",
        "fdd_role": "Generator / Auditor / Validator (long-context)",
        "capability_rank": 3,
        "price_in": 6.5,
        "price_out": 27.0,
    },
    {
        "id": "dashscope_deepseek-v4-flash-0731",
        "kind": "chat",
        "label": "DeepSeek-V4-Flash",
        "fdd_role": "cheap draft / high-volume; quality floor for FDD",
        "capability_rank": 4,
        "price_in": 1.0,
        "price_out": 2.0,
    },
    {
        "id": "text_v4_embedding",
        "kind": "embedding",
        "label": "Text-V4-Embedding",
        "fdd_role": "NOT for commentary — embedding only (no FDD path today)",
        "capability_rank": None,
        "price_in": 0.5,
        "price_out": 0.0,
    },
    {
        "id": "qwen3_rerank",
        "kind": "rerank",
        "label": "Qwen3-Rerank",
        "fdd_role": "NOT for commentary — rerank only (no FDD path today)",
        "capability_rank": None,
        "price_in": 0.5,
        "price_out": 0.0,
    },
]

_KIND_DEFAULT_ROLE = {
    "chat": "chat (FDD pipeline candidate)",
    "embedding": "NOT for commentary — embedding only",
    "rerank": "NOT for commentary — rerank only",
}


def _blended_1_to_5(price_in: float, price_out: float) -> float:
    """Cost of 1 input unit + 5 output units (user's I:O = 1:5 estimate)."""
    return price_in + 5.0 * price_out


def _config_path() -> Path:
    from fdd_utils.ai import get_default_config_path
    return Path(get_default_config_path())


def _load_config_dashscope() -> tuple[dict[str, Any], Path]:
    path = _config_path()
    try:
        from fdd_utils.ai import load_yaml_config
        cfg = load_yaml_config(str(path))
        block = cfg.get("dashscope") or {}
        return (block if isinstance(block, dict) else {}), path
    except Exception:
        return {}, path


def _models_from_config(block: dict[str, Any]) -> list[dict[str, Any]]:
    """Merge config available_models over the built-in catalog (by id)."""
    by_id = {m["id"]: dict(m) for m in _DEFAULT_MODELS}
    raw = block.get("available_models") or []
    if not isinstance(raw, list) or not raw:
        return list(_DEFAULT_MODELS)

    ordered: list[dict[str, Any]] = []
    seen: set[str] = set()
    for i, entry in enumerate(raw):
        if isinstance(entry, str):
            mid = entry
            overlay: dict[str, Any] = {"id": mid}
        elif isinstance(entry, dict) and entry.get("id"):
            mid = str(entry["id"])
            overlay = dict(entry)
        else:
            continue
        base = by_id.get(mid, {
            "id": mid,
            "kind": overlay.get("kind") or "chat",
            "label": overlay.get("label") or mid,
            "fdd_role": _KIND_DEFAULT_ROLE.get(str(overlay.get("kind") or "chat"), "unknown"),
            "capability_rank": None,
            "price_in": 0.0,
            "price_out": 0.0,
        })
        merged = dict(base)
        for key in ("kind", "label", "fdd_role", "capability_rank", "price_in", "price_out"):
            if key in overlay and overlay[key] is not None:
                merged[key] = overlay[key]
        if merged.get("capability_rank") is None and merged.get("kind") == "chat":
            merged["capability_rank"] = i + 1
        ordered.append(merged)
        seen.add(mid)
    return ordered


def _resolve_creds(args: argparse.Namespace) -> tuple[str, str, str, dict[str, Any], Path, str]:
    """Prefer config.yml; CLI / env only override when explicitly set.

    Returns: api_base, api_key, auth, dashscope_block, config_path, source_note
    """
    block, path = _load_config_dashscope()

    def _pick(cli_val, env_key: str, cfg_key: str) -> tuple[str, str]:
        if cli_val is not None and str(cli_val).strip() != "":
            return str(cli_val).strip(), "cli"
        env_val = os.environ.get(env_key)
        if env_val and env_val.strip():
            return env_val.strip(), "env"
        cfg_val = block.get(cfg_key)
        if cfg_val is not None and str(cfg_val).strip():
            return str(cfg_val).strip(), "config.yml"
        return "", "missing"

    api_base, src_base = _pick(args.api_base, "DASHSCOPE_API_BASE", "api_base")
    api_key, src_key = _pick(args.api_key, "DASHSCOPE_API_KEY", "api_key")
    auth_raw, src_auth = _pick(args.auth, "DASHSCOPE_AUTH", "auth")
    auth = (auth_raw or "bearer").strip().lower()
    if not auth_raw:
        src_auth = "default(bearer)"

    api_base = api_base.rstrip("/")
    source_note = f"api_base={src_base}, api_key={src_key}, auth={src_auth}"
    return api_base, api_key, auth, block, path, source_note


def _headers(
    api_key: str,
    auth: str,
    charge_code: str = "0000",
    region_override: str = "westeurope",
) -> dict[str, str]:
    if auth == "workbench":
        return {
            "Authorization": f"Bearer {api_key}",
            "api-key": api_key,
            "Ocp-Apim-Subscription-Key": api_key,
            "x-kpmg-charge-code": charge_code,
            "x-kpmg-region-override": region_override,
            "Content-Type": "application/json",
        }
    return {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }


def _post(url: str, headers: dict[str, str], body: dict[str, Any], timeout: float) -> tuple[int, Any, float]:
    import urllib.error
    import urllib.request

    data = json.dumps(body).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers=headers, method="POST")
    start = time.time()
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read().decode("utf-8", errors="replace")
            elapsed = time.time() - start
            try:
                payload = json.loads(raw) if raw else {}
            except json.JSONDecodeError:
                payload = {"_raw": raw[:500]}
            return int(resp.status), payload, elapsed
    except urllib.error.HTTPError as exc:
        elapsed = time.time() - start
        raw = exc.read().decode("utf-8", errors="replace")
        try:
            payload = json.loads(raw) if raw else {"error": str(exc)}
        except json.JSONDecodeError:
            payload = {"_raw": raw[:800], "error": str(exc)}
        return int(exc.code), payload, elapsed
    except Exception as exc:
        elapsed = time.time() - start
        return 0, {"error": f"{type(exc).__name__}: {exc}"}, elapsed


def _chat_urls(api_base: str) -> list[str]:
    # Try the common OpenAI-compatible shapes this gateway family uses.
    bases = [api_base]
    if api_base.endswith("/v1"):
        bases.append(api_base[: -len("/v1")])
    urls: list[str] = []
    for base in bases:
        for suffix in ("/chat/completions", "/v1/chat/completions"):
            url = f"{base}{suffix}"
            if url not in urls:
                urls.append(url)
    return urls


def _embedding_urls(api_base: str) -> list[str]:
    bases = [api_base]
    if api_base.endswith("/v1"):
        bases.append(api_base[: -len("/v1")])
    urls: list[str] = []
    for base in bases:
        for suffix in ("/embeddings", "/v1/embeddings"):
            url = f"{base}{suffix}"
            if url not in urls:
                urls.append(url)
    return urls


def _rerank_urls(api_base: str) -> list[str]:
    bases = [api_base]
    if api_base.endswith("/v1"):
        bases.append(api_base[: -len("/v1")])
    urls: list[str] = []
    for base in bases:
        for suffix in (
            "/rerank",
            "/v1/rerank",
            "/services/rerank",
            "/api/v1/services/rerank",
        ):
            url = f"{base}{suffix}"
            if url not in urls:
                urls.append(url)
    return urls


def _test_chat(model: dict[str, Any], api_base: str, headers: dict[str, str], timeout: float) -> dict[str, Any]:
    body = {
        "model": model["id"],
        "messages": [
            {"role": "system", "content": "You are a helpful assistant."},
            {"role": "user", "content": "Reply with exactly: PONG"},
        ],
        "max_tokens": 64,
        "temperature": 0,
    }
    last: dict[str, Any] = {"ok": False, "detail": "no URL tried"}
    for url in _chat_urls(api_base):
        status, payload, elapsed = _post(url, headers, body, timeout)
        content = ""
        try:
            choices = payload.get("choices") or []
            if choices:
                msg = choices[0].get("message") or {}
                content = str(msg.get("content") or "").strip()
        except Exception:
            content = ""
        ok = status == 200 and bool(content)
        last = {
            "ok": ok,
            "kind": "chat",
            "url": url,
            "status": status,
            "elapsed_s": round(elapsed, 2),
            "content": content[:200],
            "usage": payload.get("usage"),
            "error": None if ok else payload,
        }
        if ok:
            return last
        # 404 on path → try next URL shape; other errors usually mean auth/model.
        if status not in (0, 404):
            return last
    return last


def _test_embedding(model: dict[str, Any], api_base: str, headers: dict[str, str], timeout: float) -> dict[str, Any]:
    body = {
        "model": model["id"],
        "input": "connectivity probe",
    }
    last: dict[str, Any] = {"ok": False, "detail": "no URL tried"}
    for url in _embedding_urls(api_base):
        status, payload, elapsed = _post(url, headers, body, timeout)
        dim = None
        try:
            data = payload.get("data") or []
            if data and isinstance(data[0], dict):
                emb = data[0].get("embedding") or []
                dim = len(emb) if isinstance(emb, list) else None
        except Exception:
            dim = None
        ok = status == 200 and bool(dim)
        last = {
            "ok": ok,
            "kind": "embedding",
            "url": url,
            "status": status,
            "elapsed_s": round(elapsed, 2),
            "dim": dim,
            "usage": payload.get("usage"),
            "error": None if ok else payload,
        }
        if ok:
            return last
        if status not in (0, 404):
            return last
    return last


def _test_rerank(model: dict[str, Any], api_base: str, headers: dict[str, str], timeout: float) -> dict[str, Any]:
    # Two common body shapes: OpenAI-ish {model,query,documents} and DashScope
    # {model,input:{query,documents}}. Try both against each URL.
    bodies = [
        {
            "model": model["id"],
            "query": "cash and cash equivalents",
            "documents": [
                "Cash at bank totalled 12.3 million.",
                "Inventory increased due to seasonality.",
                "Trade receivables days were 45.",
            ],
        },
        {
            "model": model["id"],
            "input": {
                "query": "cash and cash equivalents",
                "documents": [
                    "Cash at bank totalled 12.3 million.",
                    "Inventory increased due to seasonality.",
                    "Trade receivables days were 45.",
                ],
            },
        },
    ]
    last: dict[str, Any] = {"ok": False, "detail": "no URL tried"}
    for url in _rerank_urls(api_base):
        for body in bodies:
            status, payload, elapsed = _post(url, headers, body, timeout)
            results = (
                payload.get("results")
                or payload.get("data")
                or (payload.get("output") or {}).get("results")
                if isinstance(payload, dict)
                else None
            )
            ok = status == 200 and bool(results)
            last = {
                "ok": ok,
                "kind": "rerank",
                "url": url,
                "status": status,
                "elapsed_s": round(elapsed, 2),
                "n_results": len(results) if isinstance(results, list) else None,
                "error": None if ok else payload,
            }
            if ok:
                return last
            if status not in (0, 404):
                # Auth / model / schema error — keep trying other body shapes
                # on same URL once, then move on.
                continue
    return last


def _print_catalog(models: list[dict[str, Any]]) -> None:
    print("\nCatalog (capability order for FDD chat; utilities last)")
    print("-" * 78)
    print(f"{'rank':<5} {'id':<38} {'I':>6} {'O':>6} {'1in+5out':>10}  role")
    for m in models:
        blended = _blended_1_to_5(float(m.get("price_in") or 0), float(m.get("price_out") or 0))
        rank = m["capability_rank"] if m.get("capability_rank") is not None else "-"
        role = m.get("fdd_role") or m.get("label") or ""
        print(
            f"{str(rank):<5} {m['id']:<38} {float(m.get('price_in') or 0):>6.1f} "
            f"{float(m.get('price_out') or 0):>6.1f} {blended:>10.1f}  {role}"
        )


def main() -> int:
    ap = argparse.ArgumentParser(description="Probe DashScope / gateway model connectivity")
    ap.add_argument("--api-base", default=None, help="override config.yml dashscope.api_base")
    ap.add_argument("--api-key", default=None, help="override config.yml dashscope.api_key")
    ap.add_argument(
        "--auth",
        choices=("bearer", "workbench"),
        default=None,
        help="override config.yml dashscope.auth (bearer | workbench)",
    )
    ap.add_argument("--model", default=None, help="test only this model id")
    ap.add_argument(
        "--kind",
        choices=("all", "chat", "embedding", "rerank"),
        default="all",
        help="filter by API kind (default: all)",
    )
    ap.add_argument("--timeout", type=float, default=60.0)
    ap.add_argument("--catalog-only", action="store_true", help="print ranking/cost table and exit")
    args = ap.parse_args()

    api_base, api_key, auth, block, cfg_path, source_note = _resolve_creds(args)
    models = _models_from_config(block)
    _print_catalog(models)
    print(f"\nconfig   : {cfg_path}")
    print(f"source   : {source_note}")

    if args.catalog_only:
        return 0

    if not api_base or not api_key:
        print(
            "\n❌ Missing api_base / api_key in fdd_utils/config.yml -> dashscope:\n"
            "   dashscope:\n"
            "     api_key: \"<your-key>\"\n"
            "     api_base: \"https://<host>/...\"\n"
            "     auth: \"bearer\"   # or workbench\n"
            "   Then re-run: python ad-hoc/workbench/test_dashscope_connectivity.py"
        )
        return 1

    charge_code = str(block.get("charge_code") or "0000")
    region = str(block.get("region_override") or "westeurope")
    headers = _headers(api_key, auth, charge_code=charge_code, region_override=region)
    print(f"api_base : {api_base}")
    print(f"auth     : {auth}")
    print(f"key      : {api_key[:4]}…{api_key[-4:]}" if len(api_key) > 8 else "key      : (short)")

    selected = models
    if args.model:
        selected = [m for m in models if m["id"] == args.model]
        if not selected:
            print(f"❌ Unknown model id: {args.model}")
            print("   Known:", ", ".join(m["id"] for m in models))
            return 1
    if args.kind != "all":
        selected = [m for m in selected if m["kind"] == args.kind]

    results: dict[str, dict[str, Any]] = {}
    for model in selected:
        print(f"\n{'=' * 70}\n  {model['id']}  [{model['kind']}]\n{'=' * 70}")
        if model["kind"] == "chat":
            result = _test_chat(model, api_base, headers, args.timeout)
        elif model["kind"] == "embedding":
            result = _test_embedding(model, api_base, headers, args.timeout)
        else:
            result = _test_rerank(model, api_base, headers, args.timeout)
        results[model["id"]] = result
        if result.get("ok"):
            extra = ""
            if result.get("content"):
                extra = f"  content={result['content']!r}"
            elif result.get("dim"):
                extra = f"  dim={result['dim']}"
            elif result.get("n_results") is not None:
                extra = f"  n_results={result['n_results']}"
            print(f"  ✅ OK  {result.get('elapsed_s')}s  via {result.get('url')}{extra}")
        else:
            print(f"  ❌ FAIL  status={result.get('status')}  url={result.get('url')}")
            err = result.get("error")
            if err is not None:
                print(f"     {json.dumps(err, ensure_ascii=False)[:500]}")

    print(f"\n{'=' * 70}\n  SUMMARY\n{'=' * 70}")
    chat_ok = []
    meta_by_id = {m["id"]: m for m in models}
    for mid, result in results.items():
        mark = "✅ PASS" if result.get("ok") else "❌ FAIL"
        print(f"  {mark}  {mid}")
        meta = meta_by_id[mid]
        if result.get("ok") and meta["kind"] == "chat":
            chat_ok.append(mid)

    if chat_ok:
        print("\nFDD-ready chat models that responded:")
        for mid in chat_ok:
            meta = meta_by_id[mid]
            blended = _blended_1_to_5(float(meta.get("price_in") or 0), float(meta.get("price_out") or 0))
            print(f"  • {mid}  (rank {meta.get('capability_rank')}, 1in+5out≈{blended:.1f})")
    return 0 if results and all(r.get("ok") for r in results.values()) else 1


if __name__ == "__main__":
    raise SystemExit(main())

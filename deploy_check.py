"""Pre-launch check for a server install: is this checkout ready to serve users?

Reads the SAME fdd_utils/config.yml the app reads and checks, without calling
any model: the default provider and model, whether they are listed and have
credentials, the files git does not carry (config.yml, template.pptx), the
font metrics, a writable logs folder, the admin unlock secret, and the
library versions. Seconds, zero tokens. Run it on the server after every
config.yml edit -- hand edits to that file have failed silently before.

    python deploy_check.py
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

problems = 0


def line(ok: bool, text: str, fix: str = "") -> None:
    global problems
    print(("  ✅ " if ok else "  ❌ ") + text)
    if not ok:
        problems += 1
        if fix:
            print("       -> " + fix)


def main() -> int:
    from fdd_utils.ai import get_default_config_path, list_model_choices, load_yaml_config
    from fdd_utils.ai.config import default_model_choice_key, is_provider_ready

    print("Files git does not carry")
    cfg_path = Path(get_default_config_path())
    line(cfg_path.exists(), f"config.yml at {cfg_path}", "copy your config.yml into fdd_utils/")
    template = ROOT / "fdd_utils" / "template.pptx"
    line(template.exists(), f"template.pptx at {template}", "copy the team template into fdd_utils/template.pptx")
    for name in ("arial_eng.json", "msyh_chi.json"):
        metrics = ROOT / "fdd_utils" / "font_metrics" / name
        line(metrics.exists(), f"font metrics {name}", "git pull: these files are committed")

    config = load_yaml_config(str(cfg_path)) if cfg_path.exists() else {}
    print("\nModel the app starts on")
    default = config.get("default") or {}
    provider = str(default.get("ai_provider") or "")
    model = str(default.get("model") or "")
    print(f"     default.ai_provider = {provider or '(not set)'}   default.model = {model or '(not set)'}")
    choices = list_model_choices(config)
    start_key = default_model_choice_key(config, choices)
    start = next((c for c in choices if c["key"] == start_key), None)
    line(bool(start) and start["model_type"] == provider and (not model or start["model_name"] == model),
         f"a new session starts on: {start['label'] if start else 'nothing'}",
         "default.model must be one of that provider's available_models ids, and not listed: false")
    line(is_provider_ready(config, provider) if provider else False,
         f"{provider or 'default provider'} has api_key and api_base",
         f"fill {provider}.api_key and {provider}.api_base in config.yml")
    print("     models in the dropdown:")
    for c in choices:
        print(f"       - {c['label']}  [{c['key']}]  ready={c['ready']}")
    unready = [c["label"] for c in choices if not c["ready"]]
    line(not unready, "every listed model is configured" if not unready
         else f"listed but not configured: {', '.join(unready)}",
         "set listed: false on that provider (or model) to hide it from users")

    print("\nLogging and admin")
    logs = ROOT / "fdd_utils" / "logs"
    try:
        logs.mkdir(parents=True, exist_ok=True)
        probe = logs / ".write_test"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
        writable = True
    except Exception:
        writable = False
    line(writable, f"logs folder writable: {logs}", "give the app's user write access to fdd_utils/logs")
    from fdd_utils.usage_ledger import admin_code, admin_secret, admin_secret_is_custom
    print(f"     admin unlock code today: {admin_code(admin_secret(config))} "
          f"({'custom admin.unlock_secret' if admin_secret_is_custom(config) else 'built-in salt'})")
    cache = ((config.get("processing") or {}).get("response_cache"))
    print(f"     processing.response_cache = {cache!r} (a cached answer costs 0 tokens and is logged as such)")

    print("\nLibraries")
    import pandas
    import streamlit
    print(f"     python {sys.version.split()[0]}  pandas {pandas.__version__}  streamlit {streamlit.__version__}")
    try:
        import openai
        print(f"     openai {openai.__version__}")
    except Exception:
        line(False, "openai package importable", "pip install -r requirements.txt")

    print()
    print("READY" if not problems else f"{problems} problem(s) above -- fix them before users log in")
    return 0 if not problems else 1


if __name__ == "__main__":
    raise SystemExit(main())

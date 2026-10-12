"""Run the same accounts through several models and put the results side by side.

For deciding whether a model is fit to be offered before it goes on the UI
list: every candidate writes the same accounts from the same frames and
prompts, next to the baseline (the model in use until now), and each final
text is then scored by the SAME deterministic verifier -- not by the model's
own Validator stage, which would have each model marking its own work.

    python ad-hoc/workbench/compare_models.py --list
        no AI call: what the UI dropdown offers from this machine's config.yml,
        which one a session starts on, and what a comparison would run.

    python ad-hoc/workbench/compare_models.py for_test/<databook>.xlsx
        baseline + every DashScope chat model, on the default accounts.

    python ad-hoc/workbench/compare_models.py for_test/<databook>.xlsx \
        --accounts 应收账款,财务费用 --models dashscope::kimi-k2.6_chat --baseline workbench

Cost: one full Generator -> Auditor -> Validator pass per account per model.
Eight accounts is roughly a third of a normal run, per model.

What the scoreboard can and cannot say. Counted for every model alike:
unsupported figures, arithmetic the text gets wrong, retries the feedback loop
needed, empty outputs, time. NOT counted: whether the analysis is any good.
That is what the side-by-side text below the table is for, and it is also
written to for_test/model_compare/ so it can be read away from the console.
Output quotes account figures; do not paste it anywhere public.
"""
from __future__ import annotations

import argparse
import os
import re
import sys
import time
from typing import Any, Dict, List, Optional, Tuple

_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, _ROOT)

from fdd_utils.ai import (  # noqa: E402
    default_model_choice_key,
    get_default_config_path,
    is_provider_ready,
    list_model_choices,
    load_yaml_config,
)

#: Accounts where the baseline has gone wrong on a real deck (a wrong-period
#: component, an invented nature, a percentage that does not follow from its
#: figures), plus two it writes well. Used when --accounts is not given.
DEFAULT_ACCOUNTS = ["应收账款", "固定资产", "在建工程", "长期待摊费用", "长期应付款",
                    "营业收入", "财务费用", "投资收益"]

_FIGURE = re.compile(r"\d[\d,]*(?:\.\d+)?\s*(?:万元|亿元|千元|元|%|million|thousand|[kKmM]\b)")


def _spec(config: Dict[str, Any], text: str) -> Tuple[str, Optional[str]]:
    """'provider::id', a bare provider, or a bare model id -> (provider, id)."""
    text = text.strip()
    if "::" in text:
        provider, model = text.split("::", 1)
        return provider.strip(), (model.strip() or None) if model.strip() != "default" else None
    if isinstance(config.get(text), dict):
        return text, None
    for choice in list_model_choices(config, include_unlisted=True):
        if choice["model_name"] == text:
            return choice["model_type"], text
    raise SystemExit(f"'{text}' is neither a provider in config.yml nor an id in any available_models. "
                     f"Run with --list to see what is there.")


def _name(config: Dict[str, Any], provider: str, model: Optional[str]) -> str:
    return f"{provider}::{model or (config.get(provider) or {}).get('chat_model') or 'default'}"


def _baseline(config: Dict[str, Any], requested: Optional[str]) -> Tuple[str, Optional[str]]:
    if requested:
        return _spec(config, requested)
    configured = str((config.get("benchmark") or {}).get("baseline") or "").strip()
    if configured:
        return _spec(config, configured)
    for provider in ("workbench", "local", "openai", "deepseek"):
        if is_provider_ready(config, provider):
            return provider, None
    raise SystemExit("No baseline: none of workbench / local / openai / deepseek is configured. Pass --baseline.")


def _print_list(config: Dict[str, Any]) -> None:
    listed = list_model_choices(config)
    everything = list_model_choices(config, include_unlisted=True)
    default_key = default_model_choice_key(config, listed)
    print(f"config: {get_default_config_path()}")
    print(f"default block: {config.get('default') or '(none -- the first listed model is the default)'}\n")
    print("UI dropdown, in order:")
    for choice in listed:
        print(f"  {'*' if choice['key'] == default_key else ' '} {choice['label']:34s} {choice['key']:48s}"
              f" {'ready' if choice['ready'] else 'NOT CONFIGURED (api_key / api_base / chat_model empty)'}")
    if not listed:
        print("  (nothing -- every provider is hidden or has no available_models)")
    print("  (* = what a new session starts on)\n")
    hidden = [c for c in everything if not c["listed"]]
    print("Configured but hidden from the UI (listed: false) -- still callable here and from the CLI:")
    for choice in hidden:
        print(f"    {choice['label']:34s} {choice['key']:48s} {'ready' if choice['ready'] else 'NOT CONFIGURED'}")
    if not hidden:
        print("    (none)")
    try:
        provider, model = _baseline(config, None)
        print(f"\nA comparison would use as baseline: {_name(config, provider, model)}"
              f"   (override with --baseline, or benchmark.baseline in config.yml)")
    except SystemExit as exc:
        print(f"\n{exc}")
    print("and as candidates: " + (", ".join(
        c["key"] for c in everything if c["model_type"] == "dashscope") or "(no dashscope chat model in config)"))


def _score(text: str, df) -> Dict[str, Any]:
    from fdd_utils.ai.validator import verify_commentary
    reviews = verify_commentary(text, df) if text.strip() else []
    bad = [r for r in reviews if isinstance(r, dict) and not r.get("supported")]
    return {
        "chars": len(text),
        "figures": len(_FIGURE.findall(text)),
        "clauses": len(reviews),
        "hallucination": [r for r in bad if r.get("category") == "hallucination"],
        "other": [r for r in bad if r.get("category") != "hallucination"],
    }


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("databook", nargs="?", help="path to the .xlsx (omit with --list)")
    ap.add_argument("--list", action="store_true", help="print the UI model list and default from config.yml; no AI call")
    ap.add_argument("--accounts", default=None, help="comma-separated account names; default: " + ",".join(DEFAULT_ACCOUNTS))
    ap.add_argument("--models", default=None, help="comma-separated provider::id (or bare id); default: every dashscope chat model")
    ap.add_argument("--baseline", default=None, help="provider or provider::id of the model in use today")
    ap.add_argument("--entity", default="", help="entity name, as for inspect_databook.py")
    ap.add_argument("--sheet", default=None, help="Financials sheet name; default: auto-detected")
    args = ap.parse_args()

    config = load_yaml_config(get_default_config_path())
    if args.list or not args.databook:
        _print_list(config)
        return 0

    baseline = _baseline(config, args.baseline)
    if args.models:
        candidates = [_spec(config, m) for m in args.models.split(",") if m.strip()]
    else:
        candidates = [(c["model_type"], c["model_name"])
                      for c in list_model_choices(config, include_unlisted=True) if c["model_type"] == "dashscope"]
    runs: List[Tuple[str, Optional[str]]] = [baseline] + [c for c in candidates if c != baseline]
    for provider, model in runs:
        if not is_provider_ready(config, provider):
            # FDDConfig would quietly substitute the first ready provider, and
            # the column would then be another model's work under this name.
            raise SystemExit(f"{provider} is not configured in config.yml (api_key / api_base / chat_model) -- "
                             f"refusing to run '{_name(config, provider, model)}' on a substitute.")

    import pandas as pd
    from fdd_utils.ai import run_ai_pipeline_with_progress
    from fdd_utils.workbook import process_workbook_data

    sheet = args.sheet
    if not sheet:
        names = pd.ExcelFile(args.databook).sheet_names
        sheet = next((n for n in names if "financial" in n.lower()), None)
    state = process_workbook_data(temp_path=args.databook, entity_name=args.entity,
                                  selected_sheet=sheet, debug=False)
    dfs = state.get("dfs") or {}
    language = state.get("language", "Eng")
    wanted = [a.strip() for a in (args.accounts.split(",") if args.accounts else DEFAULT_ACCOUNTS) if a.strip()]
    keys = [a for a in wanted if a in dfs]
    missing = [a for a in wanted if a not in dfs]
    if not keys and not args.accounts:
        # The defaults are Chinese tab names; a workbook keyed in English has none.
        keys, missing = list(dfs)[:len(DEFAULT_ACCOUNTS)], []
    if missing:
        print(f"not in this workbook, skipped: {missing}")
    if not keys:
        raise SystemExit(f"none of the requested accounts is in this workbook. Available: {sorted(dfs)}")
    print(f"language={language}  accounts={keys}")
    print("models: " + ", ".join(_name(config, p, m) + (" [baseline]" if (p, m) == baseline else "") for p, m in runs))

    outcome: Dict[str, Dict[str, Any]] = {}
    for provider, model in runs:
        name = _name(config, provider, model)
        workers = (config.get(provider) or {}).get("max_workers") or (1 if provider == "local" else 4)
        print(f"\n>>> {name}  ({len(keys)} accounts, {workers} worker(s))", flush=True)
        started = time.time()
        try:
            results = run_ai_pipeline_with_progress(
                mapping_keys=list(keys), dfs=dfs, model_type=provider, model_name=model, language=language,
                use_multithreading=True, max_workers=int(workers),
                workbook_digest=(state.get("resolution") or {}).get("workbook_digest"),
            )
            error = None
        except Exception as exc:  # one gateway failing must not cost the others' runs
            results, error = {}, f"{type(exc).__name__}: {exc}"
            print(f"    FAILED: {error}")
        elapsed = time.time() - started
        per_account: Dict[str, Any] = {}
        for key in keys:
            record = results.get(key) if isinstance(results.get(key), dict) else {}
            text = str(record.get("final") or "").strip()
            per_account[key] = {"text": text, "retries": int(record.get("feedback_retries") or 0),
                                "generator": str(record.get("subagent_1") or "").strip(), **_score(text, dfs[key])}
        outcome[name] = {"seconds": elapsed, "error": error, "accounts": per_account}
        print(f"    {elapsed:.0f}s")

    lines: List[str] = []
    out = lines.append
    out("=" * 100)
    out("SCOREBOARD -- every text scored by the same deterministic verifier")
    out("=" * 100)
    out("%-46s %6s %6s %7s %7s %8s %7s %7s" % ("model", "secs", "empty", "retries", "halluc", "other", "figures", "chars"))
    for name, run in outcome.items():
        accounts = run["accounts"].values()
        out("%-46s %6.0f %6d %7d %7d %8d %7d %7d%s" % (
            name[:46], run["seconds"], sum(1 for a in accounts if not a["text"]),
            sum(a["retries"] for a in accounts), sum(len(a["hallucination"]) for a in accounts),
            sum(len(a["other"]) for a in accounts), sum(a["figures"] for a in accounts),
            sum(a["chars"] for a in accounts), "   ERROR: " + run["error"][:60] if run["error"] else ""))
    out("")
    out("empty   = accounts that came back with no commentary")
    out("retries = feedback-loop regenerations (each one is three more calls)")
    out("halluc  = clauses carrying a figure the account's own numbers do not support, or arithmetic that is wrong")
    out("other   = unsupported inference (a stated cause with no source) -- not a defect by itself")
    out("figures = amounts and percentages quoted; a low count next to the baseline means thinner commentary")
    out("")
    out("%-14s " % "per account" + " ".join("%-22s" % n.split("::", 1)[1][:22] for n in outcome))
    for key in keys:
        cells = []
        for run in outcome.values():
            a = run["accounts"][key]
            cells.append("%-22s" % ("EMPTY" if not a["text"] else
                                     f"H{len(a['hallucination'])} O{len(a['other'])} R{a['retries']} {a['figures']}fig {a['chars']}ch"))
        out("%-14s " % key + " ".join(cells))
    out("  (H = halluc, O = other, R = retries)")

    for key in keys:
        out("")
        out("=" * 100)
        out(f"  {key}")
        out("=" * 100)
        for name, run in outcome.items():
            a = run["accounts"][key]
            out(f"\n--- {name}{'  [baseline]' if name == _name(config, *baseline) else ''}"
                f"  ({a['chars']} chars, {a['retries']} retries) ---")
            out(a["text"] or "(empty)")
            for review in a["hallucination"] + a["other"]:
                out(f"    [{review.get('category')}] {str(review.get('clause'))[:110]}")
                out(f"        {str(review.get('reason'))[:240]}")

    report = "\n".join(lines)
    print("\n" + report)
    out_dir = os.path.join(os.path.dirname(os.path.abspath(args.databook)), "model_compare")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, time.strftime("compare_%Y%m%d_%H%M%S.txt"))
    with open(out_path, "w", encoding="utf-8") as handle:
        handle.write(report + "\n")
    print(f"\nsaved: {out_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

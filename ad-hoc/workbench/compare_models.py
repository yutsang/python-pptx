"""Run the same accounts through several models and put the results side by side.

For deciding whether a model is fit to be offered before it goes on the UI
list: every candidate writes the same accounts from the same frames and
prompts, next to the baseline (the model in use until now), and each final
text is then scored by the SAME deterministic verifier -- not by the model's
own Validator stage, which would have each model marking its own work.

    python ad-hoc/workbench/compare_models.py --list
        no AI call: what the UI dropdown offers from this machine's config.yml,
        which one a session starts on, and what a comparison would run.

    python ad-hoc/workbench/compare_models.py --show-config
        no AI call: the model blocks of config.yml as written, with every key,
        address and comment masked -- safe to paste when a block misbehaves.

    python ad-hoc/workbench/compare_models.py for_test/<databook>.xlsx
        baseline + every DashScope chat model, on the default accounts.

    python ad-hoc/workbench/compare_models.py for_test/<databook>.xlsx \
        --accounts 应收账款,财务费用 --models dashscope::kimi-k2.6_chat --baseline workbench

    python ad-hoc/workbench/compare_models.py for_test/<databook>.xlsx --all
        every account in the workbook.

Cost: one full Generator -> Auditor -> Validator pass per account per model.
Eight accounts is roughly a third of a normal run, per model. Models run one
after another, each with its own progress bar and its accounts on --workers
threads (default: the provider's max_workers, else 4). --parallel N runs N
models side by side instead, without bars. The report ends with a projection
of what every account would take.

What the scoreboard can and cannot say. Counted for every model alike:
unsupported figures, arithmetic the text gets wrong, retries the feedback loop
needed, empty outputs, time. NOT counted: whether the analysis is any good.
That is what the side-by-side text below the table is for, and it is also
written to for_test/model_compare/ so it can be read away from the console.
Output quotes account figures; do not paste it anywhere public.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from typing import Any, Dict, List, Optional, Tuple

# A child run draws its progress bar on the console and sends everything else
# -- the pipeline's log lines and warnings -- to its .log file. Done on the
# file descriptors, before anything is imported: a logging handler keeps
# whichever stderr existed when it was created, so swapping sys.stderr later
# would leave those lines scribbling over the bar.
_BAR_STREAM = None
if "--child-bar" in sys.argv:
    try:
        _BAR_STREAM = os.fdopen(os.dup(2), "w", errors="replace")
        os.dup2(1, 2)
    except OSError:
        _BAR_STREAM = None

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
    _print_provider_state(config)
    try:
        provider, model = _baseline(config, None)
        print(f"\nA comparison would use as baseline: {_name(config, provider, model)}"
              f"   (override with --baseline, or benchmark.baseline in config.yml)")
    except SystemExit as exc:
        print(f"\n{exc}")
    print("and as candidates: " + (", ".join(
        c["key"] for c in everything if c["model_type"] == "dashscope") or "(no dashscope chat model in config)"))


_SECRET_KEY = re.compile(r"key$|secret|token$|passw|base$|url|host|endpoint|charge|region|subscription|tenant|proxy", re.I)
#: Model ids are what a pasted config is compared on, and some are long enough
#: to look like a token. Shown unless they carry an address.
_ID_KEY = re.compile(r"^(?:id|label|kind|model|chat_model|thinking_model|embeddings_model|ai_provider|baseline)$")
_SECRET_VALUE = re.compile(r"://|[\w-]+\.[\w-]+\.[A-Za-z]{2,}|[A-Za-z0-9+/_=-]{24,}")
#: Top-level blocks that decide which model runs. The rest of config.yml
#: (agents, pptx, analysis ...) has nothing to do with it and is not printed.
_MODEL_BLOCKS = ("default", "benchmark", "workbench", "dashscope", "local", "deepseek", "openai")


def _print_masked_config() -> None:
    """The model blocks of config.yml as written, line-numbered, with every
    credential, address and comment replaced -- safe to paste.

    Read as TEXT, not parsed: a parsed file has already lost what most needs
    seeing -- a block written twice, a key indented into the wrong block.
    Masked by key name (anything like key / base / url / charge / region), and
    again by what the value looks like (a URL, a hostname, a long token), so a
    credential under an unexpected key name is still hidden. Comments are
    dropped whole: a commented-out line is where an old key tends to survive.
    """
    path = get_default_config_path()
    with open(path, encoding="utf-8") as handle:
        rows = handle.read().splitlines()
    print(f"config: {os.path.basename(path)}  ({len(rows)} lines; model blocks only, values masked)\n")
    showing = False
    for number, row in enumerate(rows, 1):
        top = re.match(r"^([A-Za-z_][\w-]*)\s*:", row)
        if top:
            showing = top.group(1) in _MODEL_BLOCKS
        if not showing or not row.strip():
            continue
        if row.lstrip().startswith("#"):
            continue
        entry = re.match(r"^(\s*(?:-\s+)?)([^:#\s][^:#]*?)\s*:\s*(.*)$", row)
        if not entry:
            # A bare list item or a continuation line: nothing names it, so it
            # is shown only if it does not look like an address or a token.
            body = row.split(" #", 1)[0].rstrip()
            print(f"{number:4d}  {'<hidden>' if _SECRET_VALUE.search(body) else body}")
            continue
        indent, key, value = entry.groups()
        value = re.sub(r"\s+#.*$", "", value).strip()
        bare = value.strip("\"'")
        if not value:
            shown = ""
        elif not bare:
            shown = "<EMPTY>"
        elif _ID_KEY.match(key) and "://" not in bare:
            shown = value
        elif _SECRET_KEY.search(key) or _SECRET_VALUE.search(bare):
            shown = f"<set, {len(bare)} chars>"
        else:
            shown = value
        print(f"{number:4d}  {indent}{key}: {shown}")


def _print_provider_state(config: Dict[str, Any]) -> None:
    """Why a provider is or is not usable, by key NAME only -- never a value.

    A block can hold real credentials and still read as unconfigured: one
    required key spelt differently, or the block written twice. YAML keeps
    only the LAST of two same-named top-level blocks and says nothing, so
    adding a second `workbench:` to carry `listed: false` silently discards
    the first one's api_key, api_base and available_models.
    """
    from fdd_utils.ai.config import PROVIDER_REQUIRED_KEYS
    print("\nProvider blocks in this file (key names only):")
    for provider in ("workbench", "dashscope", "local", "deepseek", "openai"):
        block = config.get(provider)
        if not isinstance(block, dict):
            print(f"    {provider:10s} no such block")
            continue
        missing = [k for k in PROVIDER_REQUIRED_KEYS.get(provider, []) if not str(block.get(k) or "").strip()]
        models = block.get("available_models") or []
        print(f"    {provider:10s} {'ready' if not missing else 'NOT READY -- empty or absent: ' + ', '.join(missing)}"
              f"   | listed={block.get('listed', True)}  available_models={len(models)}"
              f"  keys present: {', '.join(sorted(str(k) for k in block))}")
    try:
        with open(get_default_config_path(), encoding="utf-8") as handle:
            tops = re.findall(r"^([A-Za-z_][\w-]*)\s*:", handle.read(), flags=re.M)
    except OSError:
        tops = []
    repeated = sorted({k for k in tops if tops.count(k) > 1})
    for key in repeated:
        print(f"    !! '{key}:' is written {tops.count(key)} times at the top level. YAML keeps only the LAST one; "
              f"everything in the earlier block(s) is ignored. Merge them into one.")
    if "benchmark" not in tops:
        print("    (no top-level 'benchmark:' block -- the baseline is the first configured of workbench / local / ...)")


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


def _workers(config: Dict[str, Any], provider: str, requested: Optional[int]) -> int:
    """inspect_databook.py's rule, so a model is timed here as it would run there."""
    return int(requested or (config.get(provider) or {}).get("max_workers")
               or (config.get("processing") or {}).get("max_workers") or 4)


def _run_child(config, args, state, dfs, language, keys) -> int:
    """One model, in this process; results to --child-out as JSON."""
    from fdd_utils.ai import run_ai_pipeline_with_progress
    from fdd_utils.ai import SUBAGENT_SEQUENCE
    provider, model = _spec(config, args.child)
    workers = _workers(config, provider, args.workers)
    started = time.time()
    payload: Dict[str, Any] = {"workers": workers, "accounts": {}}
    bar, callback = None, None
    if args.child_bar and _BAR_STREAM is not None:
        import threading
        from tqdm import tqdm
        stages = len(SUBAGENT_SEQUENCE)
        bar = tqdm(total=stages * len(keys), file=_BAR_STREAM, ascii=True, dynamic_ncols=True,
                   desc=(model or provider)[:28], unit="step")
        lock, done, size, seen = threading.Lock(), {}, {}, {"n": 0}

        def callback(agent_num, agent_label, completed, total_eligible, overall_step, mapping_key):
            # The Validator runs only on accounts asserting a cause, so a stage
            # reports its own size; the total follows it instead of stalling
            # short of 100% (inspect_databook.py's bar, same reason).
            with lock:
                bar.set_postfix_str(f"{agent_label}: {mapping_key}"[:44])
                if total_eligible:
                    size[agent_num] = total_eligible
                done[agent_num] = max(done.get(agent_num, 0), completed)
                total = sum(size.get(i, len(keys)) for i in range(1, stages + 1))
                if total != bar.total:
                    bar.total = total
                moved = sum(done.values()) - seen["n"]
                if moved > 0:
                    bar.update(moved)
                    seen["n"] += moved
    try:
        results = run_ai_pipeline_with_progress(
            mapping_keys=list(keys), dfs=dfs, model_type=provider, model_name=model, language=language,
            use_multithreading=True, max_workers=workers, progress_callback=callback,
            workbook_digest=(state.get("resolution") or {}).get("workbook_digest"),
        )
        for key in keys:
            record = results.get(key) if isinstance(results.get(key), dict) else {}
            payload["accounts"][key] = {"text": str(record.get("final") or "").strip(),
                                        "retries": int(record.get("feedback_retries") or 0)}
    except Exception as exc:  # reported in this model's column; the others still run
        payload["error"] = f"{type(exc).__name__}: {exc}"
    if bar is not None:
        bar.close()
    payload["seconds"] = time.time() - started
    with open(args.child_out, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False)
    return 1 if payload.get("error") else 0


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("databook", nargs="?", help="path to the .xlsx (omit with --list)")
    ap.add_argument("--list", action="store_true", help="print the UI model list and default from config.yml; no AI call")
    ap.add_argument("--accounts", default=None, help="comma-separated account names; default: " + ",".join(DEFAULT_ACCOUNTS))
    ap.add_argument("--models", default=None, help="comma-separated provider::id (or bare id); default: every dashscope chat model")
    ap.add_argument("--baseline", default=None, help="provider or provider::id of the model in use today")
    ap.add_argument("--show-config", action="store_true",
                    help="print the model blocks of config.yml with keys, addresses and comments masked; no AI call")
    ap.add_argument("--all", action="store_true", help="every account in the workbook instead of the default eight")
    ap.add_argument("--parallel", type=int, default=1,
                    help="how many MODELS run at once (default 1: one after another, with a progress bar each). "
                         "Threading inside a model's run is --workers and is on regardless")
    ap.add_argument("--child-bar", action="store_true", help=argparse.SUPPRESS)
    ap.add_argument("--workers", type=int, default=None,
                    help="worker threads inside each model's run; default: <provider>.max_workers, "
                         "else processing.max_workers, else 4 -- as inspect_databook.py")
    ap.add_argument("--child", default=None, help=argparse.SUPPRESS)
    ap.add_argument("--child-out", default=None, help=argparse.SUPPRESS)
    ap.add_argument("--entity", default="", help="entity name, as for inspect_databook.py")
    ap.add_argument("--sheet", default=None, help="Financials sheet name; default: auto-detected")
    args = ap.parse_args()

    if args.show_config:
        _print_masked_config()
        return 0
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
    from fdd_utils.workbook import process_workbook_data

    sheet = args.sheet
    if not sheet:
        names = pd.ExcelFile(args.databook).sheet_names
        sheet = next((n for n in names if "financial" in n.lower()), None)
    state = process_workbook_data(temp_path=args.databook, entity_name=args.entity,
                                  selected_sheet=sheet, debug=False)
    dfs = state.get("dfs") or {}
    language = state.get("language", "Eng")
    if args.all:
        wanted = list(dfs)
    else:
        wanted = [a.strip() for a in (args.accounts.split(",") if args.accounts else DEFAULT_ACCOUNTS) if a.strip()]
    keys = [a for a in wanted if a in dfs]
    missing = [a for a in wanted if a not in dfs]
    if not keys and not args.accounts:
        # The defaults are Chinese tab names; a workbook keyed in English has none.
        keys, missing = list(dfs)[:len(DEFAULT_ACCOUNTS)], []

    if args.child:
        return _run_child(config, args, state, dfs, language, keys)

    if missing:
        print(f"not in this workbook, skipped: {missing}")
    if not keys:
        raise SystemExit(f"none of the requested accounts is in this workbook. Available: {sorted(dfs)}")
    parallel = max(1, min(int(args.parallel), len(runs)))
    print(f"language={language}  accounts ({len(keys)} of {len(dfs)}): {keys}")
    print("models: " + ", ".join(_name(config, p, m) + (" [baseline]" if (p, m) == baseline else "") for p, m in runs))
    print(f"{parallel} model at a time" if parallel == 1 else f"{parallel} models at a time (no progress bars)")

    out_dir = os.path.join(os.path.dirname(os.path.abspath(args.databook)), "model_compare")
    os.makedirs(out_dir, exist_ok=True)
    stamp = time.strftime("%Y%m%d_%H%M%S")
    started_all = time.time()
    outcome: Dict[str, Dict[str, Any]] = {}

    def _one(index: int, order: int, provider: str, model: Optional[str]) -> Tuple[str, Dict[str, Any]]:
        # A process per model, not a thread: the pipeline keeps one circuit
        # breaker and one logging setup per process, so two models sharing a
        # process would trip and reset each other's breaker. Starts are spaced
        # because a run folder is named to the second.
        name = _name(config, provider, model)
        time.sleep(index * 2.0)
        safe = re.sub(r"[^A-Za-z0-9_.-]+", "_", name)
        json_path = os.path.join(out_dir, f"compare_{stamp}_{safe}.json")
        log_path = os.path.join(out_dir, f"compare_{stamp}_{safe}.log")
        command = [sys.executable, os.path.abspath(__file__), args.databook,
                   "--child", f"{provider}::{model or 'default'}", "--child-out", json_path,
                   "--accounts", ",".join(keys), "--entity", args.entity]
        if sheet:
            command += ["--sheet", sheet]
        if args.workers:
            command += ["--workers", str(args.workers)]
        if parallel == 1:
            command += ["--child-bar"]
        print(f">>> [{order}/{len(runs)}] {name}  ({_workers(config, provider, args.workers)} worker threads)"
              + ("  -- reading the workbook, then the bar starts" if parallel == 1 else ""), flush=True)
        began = time.time()
        with open(log_path, "w", encoding="utf-8") as log:
            # stderr is left on the console only for the bar; the child points
            # its own stderr at the log (see _BAR_STREAM).
            code = subprocess.run(command, stdout=log, stderr=None if parallel == 1 else subprocess.STDOUT,
                                  env={**os.environ, "PYTHONIOENCODING": "utf-8"}).returncode
        payload: Dict[str, Any] = {}
        if os.path.exists(json_path):
            with open(json_path, encoding="utf-8") as handle:
                payload = json.load(handle)
            os.remove(json_path)
        error = payload.get("error")
        if code != 0 and not payload:
            with open(log_path, encoding="utf-8", errors="replace") as handle:
                tail = [line.strip() for line in handle.read().splitlines() if line.strip()][-3:]
            error = f"process exited {code}: " + " | ".join(tail)
        print(f"<<< finished {name}  {time.time() - began:.0f}s" + (f"  FAILED: {error[:200]}" if error else ""), flush=True)
        return name, {"seconds": payload.get("seconds", time.time() - began), "error": error,
                      "workers": payload.get("workers"), "records": payload.get("accounts") or {}, "log": log_path}

    with ThreadPoolExecutor(max_workers=parallel) as pool:
        futures = [pool.submit(_one, i if 1 < parallel > i else 0, i + 1, p, m) for i, (p, m) in enumerate(runs)]
        finished = dict(f.result() for f in futures)
    wall = time.time() - started_all
    for provider, model in runs:
        name = _name(config, provider, model)
        run = finished[name]
        per_account: Dict[str, Any] = {}
        for key in keys:
            record = run["records"].get(key) or {}
            text = str(record.get("text") or "").strip()
            per_account[key] = {"text": text, "retries": int(record.get("retries") or 0), **_score(text, dfs[key])}
        outcome[name] = {**run, "accounts": per_account}

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
    total_accounts = len(dfs)
    out(f"TIME -- this run: {wall:.0f}s wall for {len(keys)} account(s) x {len(outcome)} model(s), {parallel} at a time")
    out(f"        projected for every account in the workbook ({total_accounts}), from each model's own pace here:")
    projected = []
    for name, run in outcome.items():
        if run["error"]:
            out("          %-46s failed -- no pace to project from" % name[:46])
            continue
        each = run["seconds"] / max(len(keys), 1) * total_accounts
        projected.append(each)
        out("          %-46s %5.1f s/account at %s worker(s) -> about %4.1f min" % (
            name[:46], run["seconds"] / max(len(keys), 1), run.get("workers") or "?", each / 60))
    if projected:
        # Models running side by side finish when the slowest lane does; with
        # fewer lanes than models the lanes fill longest-first.
        lanes = [0.0] * parallel
        for each in sorted(projected, reverse=True):
            lanes[lanes.index(min(lanes))] += each
        out(f"        all {len(outcome)} models: about {max(lanes) / 60:.0f} min at --parallel {parallel}, "
            f"{sum(projected) / 60:.0f} min one after another")
        out("        (a straight scale-up: a full run has more large accounts and more retries than this sample,")
        out("         and each run here re-reads the workbook, which is counted in the wall time but not per model)")
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
    out_path = os.path.join(out_dir, f"compare_{stamp}.txt")
    with open(out_path, "w", encoding="utf-8") as handle:
        handle.write(report + "\n")
    print(f"\nsaved: {out_path}   (one .log per model beside it)")
    return 0


if __name__ == "__main__":
    sys.exit(main())

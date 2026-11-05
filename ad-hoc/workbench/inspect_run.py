"""One-shot report on a finished (or in-flight) AI run — paste the output back.

Everything here is read off disk: the run folder's processing.log, data.yml,
results.yml and audit.jsonl. No databook, no LLM call, no cost. Safe to run
while a run is still going; anything not written yet is reported as missing
rather than crashing.

    python ad-hoc/workbench/inspect_run.py            # newest run
    python ad-hoc/workbench/inspect_run.py --run 20260908_101500
    python ad-hoc/workbench/inspect_run.py --last 3   # compare runs

WARNING: sections 5 and 6 print commentary verbatim, so the output carries
client figures and account names. Fine to paste into a working conversation;
do not put it anywhere public.
"""
from __future__ import annotations

import os
import sys

# Runs from the repo root on Windows too: every ad-hoc script documented
# "PYTHONPATH=. python ...", which is Unix shell syntax that cmd.exe rejects
# before python starts. Invoked by path, sys.path[0] is this script's own
# directory, so put the repo root on it here instead of asking the caller.
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import argparse
import collections
import glob
import json
import os
import re
import sys

try:
    import yaml
except ImportError:  # pragma: no cover
    sys.exit("pyyaml is required: pip install pyyaml")

LOG_ROOT = "fdd_utils/logs"
RULE = "=" * 78


def _runs():
    return sorted(glob.glob(os.path.join(LOG_ROOT, "run_*")))


def _load_yaml(path):
    if not os.path.exists(path):
        return None
    try:
        with open(path, encoding="utf-8") as fh:
            return yaml.safe_load(fh)
    except Exception as exc:
        return {"__error__": str(exc)}


def _load_jsonl(path):
    if not os.path.exists(path):
        return None
    out = []
    with open(path, encoding="utf-8") as fh:
        for line in fh:
            line = line.strip()
            if not line:
                continue
            try:
                out.append(json.loads(line))
            except Exception:
                pass
    return out


def _accounts(results):
    """Real accounts only — the run-level sentinels are not accounts."""
    return {
        k: v for k, v in (results or {}).items()
        if isinstance(v, dict) and not str(k).startswith("__")
    }


def section_health(folder, results, log_text):
    print(f"\n{RULE}\n1. RUN HEALTH\n{RULE}")
    health = (results or {}).get("__run_health__")
    if isinstance(health, dict):
        for field in ("calls_succeeded", "calls_failed", "calls_skipped_breaker",
                      "fallback_accounts", "passthrough_accounts",
                      "no_prompt_accounts", "error_text_accounts",
                      "safe_to_export", "first_failure"):
            if field in health:
                print(f"  {field:<24} {health[field]}")
        stages = health.get("stage_successes")
        if stages:
            print(f"  stage_successes          {stages}")
    else:
        # Older run, or the tally has not been filed yet. Fall back to the log.
        ok = len(re.findall(r"\] Processed:", log_text))
        fail = len(re.findall(r"AI call attempt .* failed", log_text))
        fb = len(re.findall(r"deterministic data-only fallback", log_text))
        brk = len(re.findall(r"circuit breaker OPEN", log_text))
        print("  (no __run_health__ in results — read from processing.log)")
        print(f"  successful calls         {ok}")
        print(f"  failed attempts          {fail}")
        print(f"  fallback bullets         {fb}")
        print(f"  breaker skips            {brk}")
        if ok == 0 and (fb or fail):
            print("  *** ZERO successful LLM calls — this deck is fallback text ***")


def section_phases(audit):
    print(f"\n{RULE}\n2. PHASES AND STATE PATHS\n{RULE}")
    if audit is None:
        print("  audit.jsonl not written yet (run still going, or an older run)")
        return
    tally = collections.Counter(line.get("phase") for line in audit)
    print("  phase tally:", dict(tally))
    degraded = [(l.get("account"), l.get("degraded")) for l in audit if l.get("degraded")]
    print(f"  degraded accounts: {len(degraded)}")
    for acct, causes in degraded[:10]:
        print(f"    {acct}: {causes}")
    print("\n  sample state paths:")
    for line in audit[:3]:
        path = " -> ".join(
            f"{t.get('to')}" for t in (line.get("transitions") or [])
        ) or "(none)"
        print(f"    {line.get('account')}: PLANNED -> {path}")
    retried = [l for l in audit if l.get("feedback_retries")]
    print(f"\n  accounts that retried: {len(retried)}")
    for l in retried[:6]:
        print(f"    {l.get('account')}: retries={l.get('feedback_retries')} "
              f"arbiter={l.get('feedback_arbiter')}")


def section_defects(results):
    print(f"\n{RULE}\n3. DEFECTS BY CODE\n{RULE}")
    codes = collections.Counter()
    cats = collections.Counter()
    reviewed = unsupported = 0
    for _key, res in _accounts(results).items():
        val = res.get("agent_4_validation") or {}
        for r in (val.get("clause_reviews") or []):
            if not isinstance(r, dict):
                continue
            reviewed += 1
            if r.get("supported"):
                continue
            unsupported += 1
            codes[r.get("code") or "(no code)"] += 1
            cats[r.get("category") or "(none)"] += 1
    print(f"  clauses reviewed {reviewed}, unsupported {unsupported}"
          f" ({100*unsupported/max(reviewed,1):.2f}%)")
    print("  by code:     ", dict(codes) or "(none)")
    print("  by category: ", dict(cats) or "(none)")


def section_repairs(results):
    print(f"\n{RULE}\n4. LOCAL REPAIRS (repair_mode = patch_first)\n{RULE}")
    any_log = False
    for key, res in _accounts(results).items():
        log = res.get("repair_log")
        if not log:
            continue
        any_log = True
        print(f"\n  [{key}]")
        for entry in (log if isinstance(log, list) else [log]):
            if not isinstance(entry, dict):
                continue
            print(f"    mode={entry.get('mode')} code={entry.get('code')} "
                  f"verified={entry.get('verified')}")
            before, after = entry.get("before"), entry.get("after")
            if before or after:
                print(f"      before: {str(before)[:150]}")
                print(f"      after:  {str(after)[:150]}")
            if entry.get("reason"):
                print(f"      reason: {str(entry['reason'])[:150]}")
    if not any_log:
        print("  no repair_log on any account — either repair_mode is off,")
        print("  or no defect was routed to a patch this run.")


def section_flags(results, limit):
    print(f"\n{RULE}\n5. FLAGGED CLAUSES — the false-positive read\n{RULE}")
    print("  (this is the section that decides whether the narrowed pool is")
    print("   right. For each: is the flag correct, or is the figure fine?)\n")
    shown = 0
    for key, res in _accounts(results).items():
        val = res.get("agent_4_validation") or {}
        bad = [r for r in (val.get("clause_reviews") or [])
               if isinstance(r, dict) and not r.get("supported")]
        if not bad:
            continue
        print(f"  [{key}]")
        for r in bad:
            print(f"    ({r.get('category')}/{r.get('code') or '-'}) "
                  f"{str(r.get('clause'))[:140]}")
            print(f"        reason: {str(r.get('reason'))[:190]}")
            if r.get("expected") is not None:
                print(f"        expected: {r.get('expected')}")
            shown += 1
            if shown >= limit:
                print(f"\n  ... stopped at {limit}; rerun with --flags N for more")
                return
        print()
    if shown == 0:
        print("  no unsupported clauses this run.")


def section_cost(folder):
    print(f"\n{RULE}\n6. TOKENS AND PROMPT CACHE\n{RULE}")
    data = _load_yaml(os.path.join(folder, "data.yml"))
    if not isinstance(data, dict) or "__error__" in (data or {}):
        print("  data.yml not written yet (it is written at the very end of a run)")
        return
    per_stage = collections.defaultdict(lambda: [0, 0, 0, 0])  # prompt, completion, cached, calls
    # Provider-reported counts first, the local estimate second. A provider that
    # returns no usage block leaves prompt_tokens None, and reading only that
    # printed a table of zeros -- on the very run whose failure was
    # "Range of input length should be [1, 32768]", i.e. the one run where
    # prompt size was the whole question. estimated_prompt_tokens is computed
    # from the prompt text before the call, so it is there either way.
    biggest = []
    sources = collections.Counter()
    for key, account in (data.get("processing_results") or {}).items():
        for agent, rec in (account or {}).items():
            if not isinstance(rec, dict):
                continue
            prompt = rec.get("prompt_tokens")
            if prompt is None:
                prompt = rec.get("estimated_prompt_tokens") or 0
                sources["estimated"] += 1
            else:
                sources[rec.get("token_usage_source") or "provider"] += 1
            completion = (rec.get("completion_tokens")
                          if rec.get("completion_tokens") is not None
                          else rec.get("estimated_output_tokens") or 0)
            slot = per_stage[agent]
            slot[0] += prompt or 0
            slot[1] += completion or 0
            slot[2] += (rec.get("prompt_cache_hit_tokens")
                        or rec.get("cached_prompt_tokens") or 0)
            slot[3] += 1
            biggest.append((prompt or 0, key, agent))
    tp = tc = tk = 0
    print(f"  {'stage':<14}{'calls':>7}{'prompt tok':>13}{'completion':>12}{'cached':>11}{'cache %':>9}")
    for agent, (p, c, k, n) in sorted(per_stage.items()):
        tp += p; tc += c; tk += k
        print(f"  {agent:<14}{n:>7}{p:>13,}{c:>12,}{k:>11,}{100*k/max(p,1):>8.1f}%")
    if tp:
        print(f"  {'TOTAL':<14}{'':>7}{tp:>13,}{tc:>12,}{tk:>11,}{100*tk/max(tp,1):>8.1f}%")
        print(f"\n  prompt share of all tokens: {100*tp/max(tp+tc,1):.1f}%")
        print(f"  counted from: {dict(sources)}"
              + ("   (estimated = the provider returned no usage block)"
                 if sources.get("estimated") else ""))
        if tk == 0:
            print("  cached = 0 everywhere: either the provider reports no cache")
            print("  counters, or nothing is being cached. Both are worth knowing.")

    # The account that failed on context length is not the one with the most
    # tokens overall, it is the one whose SINGLE prompt was longest -- so name
    # them rather than leaving a per-stage total to be divided by hand.
    if biggest:
        biggest.sort(reverse=True)
        print("\n  largest single prompts (a provider limit applies to ONE call):")
        for tokens, key, agent in biggest[:6]:
            print(f"    {tokens:>8,}  {key} / {agent}")
        top = biggest[0][0]
        for limit in (32768, 16384, 8192):
            if top > limit * 0.8:
                print(f"\n    the largest is {100 * top / limit:.0f}% of a {limit:,}-token"
                      f" context window. A call over the limit comes back as\n"
                      f"    'Range of input length exceeds limited' and that account"
                      f" falls back to a deterministic bullet.")
                break


_TS = re.compile(r"^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}),(\d{3}) - .*?INFO - \[(\w+)\] Processed: (.+?) \| Duration: ([\d.]+)s")


def section_timeline(log_text):
    """R3-a: how much of the wall clock is the stage BARRIER, as opposed to
    LLM time or retries. All accounts go through the Generator before any goes
    through the Auditor; workers sit idle from the moment the last-but-one
    Generator finishes until the last one does. This section measures that
    idle from the log's own timestamps, so the decision to chain stages
    per-account (or not) rests on a number rather than a hunch."""
    import datetime as _dt
    print(f"\n{RULE}\n6c. STAGE TIMELINE (barrier idle vs LLM time)\n{RULE}")
    events = []
    for line in log_text.splitlines():
        m = _TS.match(line)
        if not m:
            continue
        ts = _dt.datetime.strptime(m.group(1), "%Y-%m-%d %H:%M:%S") + _dt.timedelta(milliseconds=int(m.group(2)))
        events.append((ts, m.group(3), m.group(4), float(m.group(5))))
    if not events:
        print("  no '[Stage] Processed:' lines with timestamps in processing.log")
        return
    by_stage = {}
    for ts, stage, acct, dur in events:
        by_stage.setdefault(stage, []).append((ts, dur))
    order = [s for s in ("Generator", "Auditor", "Refiner", "Validator") if s in by_stage]
    wall = (events[-1][0] - min(ts for ts, *_ in events)).total_seconds()
    print(f"  {'stage':<11}{'calls':>6}{'first done':>12}{'last done':>11}{'spread':>8}{'sum LLM':>9}{'tail idle*':>11}")
    total_idle = 0.0
    for i, stage in enumerate(order):
        rows = sorted(by_stage[stage])
        first, last = rows[0][0], rows[-1][0]
        spread = (last - first).total_seconds()
        llm = sum(d for _t, d in rows)
        # tail idle: from the second-to-last completion to the last one, the
        # window in which at most one worker is busy while the rest wait for
        # the barrier. A lower bound on barrier cost.
        tail = (rows[-1][0] - rows[-2][0]).total_seconds() if len(rows) > 1 else 0.0
        total_idle += tail
        print(f"  {stage:<11}{len(rows):>6}{first.strftime('%H:%M:%S'):>12}{last.strftime('%H:%M:%S'):>11}"
              f"{spread:>8.1f}{llm:>9.1f}{tail:>11.1f}")
    print(f"\n  wall clock (first start to last completion): {wall:.1f}s")
    print(f"  *tail idle summed over stages: {total_idle:.1f}s = {100*total_idle/max(wall,1):.0f}% of wall clock")
    print("  Rule from the plan: chain stages per account only if this is over 20%.")
    if wall and 100 * total_idle / wall > 20:
        print("  -> over 20%: R3 step 2 (per-account chaining) is worth building.")
    else:
        print("  -> under 20%: the barrier is not where the time goes; R3 stops here.")


def section_evidence(folder):
    print(f"\n{RULE}\n6b. EVIDENCE ON DISK (what each account was graded against)\n{RULE}")
    try:
        from fdd_utils.ai.evidence import list_evidence
        evs = list_evidence(folder)
    except Exception as exc:
        print(f"  could not load: {exc}")
        return
    if not evs:
        print("  no evidence/ folder -- an older run, or one that died before finalize")
        return
    sizes = sorted(ev.pool_size for ev in evs.values())
    print(f"  {len(evs)} account(s); pool size min/median/max {sizes[0]}/{sizes[len(sizes)//2]}/{sizes[-1]}")
    empty = [k for k, ev in evs.items() if not ev.facts]
    if empty:
        print(f"  *** EMPTY pool on {len(empty)} account(s): {empty[:6]} -- those verdicts grounded against nothing")
    import collections
    secs = collections.Counter(sec for ev in evs.values() for sec in (ev.shown or {}).get("sections", []))
    print(f"  guidance sections reaching prompts: {dict(secs)}")
    dropped = {k: ev.shown.get("components_dropped") for k, ev in evs.items()
               if (ev.shown or {}).get("components_dropped")}
    if dropped:
        print(f"  prompt budget dropped components on: {dropped}")
    big = sorted(((ev.shown or {}).get("prompt_tokens_est") or 0, k) for k, ev in evs.items())[-3:]
    print(f"  largest Generator prompts (est.): {[(k, t) for t, k in reversed(big)]}")


def section_extras(results):
    print(f"\n{RULE}\n7. DIRECTION FINDINGS / CLAIM CONTRACTS\n{RULE}")
    dir_n = 0
    contracts = collections.Counter()
    for _key, res in _accounts(results).items():
        dir_n += len(res.get("direction_findings") or [])
        for claim, verdict in (res.get("claim_contract") or {}).items():
            contracts[f"{claim}:{verdict}"] += 1
    print(f"  direction findings: {dir_n}"
          + ("  (0 across the archive so far — a real run is the first test)"
             if dir_n == 0 else ""))
    print(f"  claim contracts: {dict(contracts) or '(off, or not wired on this path)'}")


def report(folder, flags_limit):
    run_id = os.path.basename(folder).replace("run_", "")
    print(RULE)
    print(f"RUN {run_id}    {folder}")
    print(RULE)
    log_path = os.path.join(folder, "processing.log")
    log_text = open(log_path, encoding="utf-8", errors="replace").read() if os.path.exists(log_path) else ""
    results = _load_yaml(os.path.join(folder, "results.yml"))
    audit = _load_jsonl(os.path.join(folder, "audit.jsonl"))
    if results is None:
        print("\n  results.yml not written yet — the run is still going, or it died")
        print("  before finalize. The health section below reads processing.log.")
    accounts = _accounts(results)
    print(f"  accounts with results: {len(accounts)}")

    section_health(folder, results, log_text)
    section_phases(audit)
    if accounts:
        section_defects(results)
        section_repairs(results)
        section_flags(results, flags_limit)
    section_cost(folder)
    section_timeline(log_text)
    section_evidence(folder)
    if accounts:
        section_extras(results)


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--run", help="run id, e.g. 20260908_101500 (default: newest)")
    ap.add_argument("--last", type=int, default=1, help="report on the N newest runs")
    ap.add_argument("--flags", type=int, default=40,
                    help="max flagged clauses to print (default 40)")
    args = ap.parse_args()

    runs = _runs()
    if not runs:
        sys.exit(f"no runs under {LOG_ROOT}/")
    if args.run:
        want = os.path.join(LOG_ROOT, f"run_{args.run.replace('run_', '')}")
        if want not in runs:
            sys.exit(f"no such run: {want}")
        targets = [want]
    else:
        targets = runs[-args.last:]

    for folder in targets:
        report(folder, args.flags)
    print(f"\n{RULE}\nEnd. Output contains client figures — do not post publicly.\n{RULE}")


if __name__ == "__main__":
    main()

"""Where a round's tokens went, from fdd_utils/logs/usage.jsonl. Free, seconds.

The ledger has one line per AIClient.get_response call -- pipeline stages,
retries, patches, executive summaries, mapping help -- which no run folder
holds completely (data.yml keeps one attempt per account and stage). This
reads a time window of it and separates what costs money from what does not:

  * calls SENT to the provider, with their prompt / completion tokens as the
    provider reported them, or an estimate, counted and labelled apart;
  * failed calls (no usage reported; their prompt estimate is shown, not summed);
  * provider prompt-cache tokens -- part of a sent request, billed at the
    provider's cache rate, not free;
  * response-cache hits -- answered from fdd_utils' own cache, never sent;
  * repeats: calls beyond the first for the same file, account and stage,
    i.e. what retries and patches added;
  * the calls that belong to no account (executive summaries, mapping help).

    python ad-hoc/workbench/usage_summary.py                         # today
    python ad-hoc/workbench/usage_summary.py --since 2026-10-29 --until 2026-10-30
    python ad-hoc/workbench/usage_summary.py --since 2026-10-29T14:00 --file Crescent

Prints totals, then the stages, files and account x stage pairs that cost the
most. Nothing here reads prompts; account and file names are printed.
"""
from __future__ import annotations

import argparse
import json
import sys
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

ROOT = Path(__file__).resolve().parents[2]
DEFAULT_LEDGER = ROOT / "fdd_utils" / "logs" / "usage.jsonl"


def _int(value: Any) -> int:
    try:
        return int(value or 0)
    except (TypeError, ValueError):
        return 0


def _rows(path: Path, since: str, until: str, file_filter: str) -> Iterable[Dict[str, Any]]:
    with path.open(encoding="utf-8") as fh:
        for line in fh:
            line = line.strip()
            if not line:
                continue
            try:
                row = json.loads(line)
            except json.JSONDecodeError:
                continue
            stamp = str(row.get("time") or "")
            if stamp < since or (until and stamp >= until):
                continue
            if file_filter and file_filter not in str(row.get("file") or ""):
                continue
            yield row


def _tokens(row: Dict[str, Any]) -> Tuple[int, int, bool]:
    """(prompt, completion, estimated) for a call that was sent."""
    if row.get("prompt_tokens") is not None or row.get("completion_tokens") is not None:
        return _int(row.get("prompt_tokens")), _int(row.get("completion_tokens")), False
    return _int(row.get("estimated_prompt_tokens")), _int(row.get("estimated_output_tokens")), True


def _fmt(n: int) -> str:
    return f"{n / 1e6:.2f}m" if n >= 1e6 else (f"{n / 1e3:.1f}k" if n >= 1e3 else str(n))


def summarise(rows: List[Dict[str, Any]], top: int = 10) -> str:
    out: List[str] = []
    sent = [r for r in rows if r.get("ok") is not False and r.get("token_source") != "response_cache"]
    cached = [r for r in rows if r.get("token_source") == "response_cache"]
    failed = [r for r in rows if r.get("ok") is False]

    def totals(group: List[Dict[str, Any]]) -> Tuple[int, int, int]:
        prompt = completion = estimated = 0
        for r in group:
            p, c, est = _tokens(r)
            prompt += p
            completion += c
            estimated += int(est)
        return prompt, completion, estimated

    p, c, est = totals(sent)
    out.append(f"SENT to provider   {len(sent):5d} calls   prompt {_fmt(p):>8}   completion {_fmt(c):>8}   "
               f"total {_fmt(p + c):>8}" + (f"   ({est} call(s) estimated, no usage reported)" if est else ""))
    cache_hit = sum(_int(r.get("cached_prompt_tokens")) for r in sent)
    if cache_hit:
        out.append(f"  of the prompt, provider prompt-cache: {_fmt(cache_hit)} ({cache_hit / max(p, 1):.0%}) "
                   f"-- still sent and billed at the cache rate")
    if failed:
        out.append(f"FAILED             {len(failed):5d} calls   prompt estimate {_fmt(sum(_int(r.get('estimated_prompt_tokens')) for r in failed))}"
                   f" (not in the totals; whether the provider bills them is the provider's rule)")
    if cached:
        out.append(f"RESPONSE CACHE     {len(cached):5d} calls   not sent; saved prompt "
                   f"{_fmt(sum(_int(r.get('saved_prompt_tokens')) for r in cached))} / completion "
                   f"{_fmt(sum(_int(r.get('saved_completion_tokens')) for r in cached))}")
    runs = Counter((str(r.get("session") or ""), str(r.get("file") or "")) for r in sent)
    out.append(f"files x sessions   {len(runs)}   providers/models: "
               + ", ".join(f"{k} {v}" for k, v in Counter(f"{r.get('provider')}/{r.get('model')}"
                                                          for r in sent).most_common()))

    # Repeats: the first call for (session, file, account, stage) is the plan;
    # every further one was a retry, a patch or a re-run. Until 2026-10-30 every
    # CLI run logged session "cli", so two CLI runs of one workbook read as one
    # run with every call of the second counted here.
    seen: Counter = Counter()
    repeat_calls = repeat_tokens = 0
    for r in sorted(sent, key=lambda r: str(r.get("time") or "")):
        key = (r.get("session"), r.get("file"), r.get("account"), r.get("stage"))
        if r.get("account") and seen[key]:
            repeat_calls += 1
            rp, rc, _ = _tokens(r)
            repeat_tokens += rp + rc
        seen[key] += 1
    out.append(f"repeats            {repeat_calls:5d} calls   total {_fmt(repeat_tokens):>8}   "
               f"({repeat_tokens / max(p + c, 1):.0%} of sent tokens): a second+ call for the same file, "
               f"account and stage -- retries, patches, re-runs")

    out.append("\nBY STAGE (sent)                       calls    prompt  completion  avg prompt/call")
    by_stage: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    for r in sent:
        by_stage[str(r.get("stage") or r.get("client_agent") or "?")].append(r)
    for stage, group in sorted(by_stage.items(), key=lambda kv: -sum(sum(_tokens(r)[:2]) for r in kv[1])):
        sp, sc, _ = totals(group)
        label = stage if any(r.get("account") for r in group) else f"{stage} (no account)"
        out.append(f"  {label[:36]:36s} {len(group):5d}  {_fmt(sp):>8}  {_fmt(sc):>10}  {_fmt(sp // max(len(group), 1)):>10}")

    out.append("\nBY FILE (sent)                                  calls     total")
    by_file: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    for r in sent:
        by_file[str(r.get("file") or "(no file)")].append(r)
    for name, group in sorted(by_file.items(), key=lambda kv: -sum(sum(_tokens(r)[:2]) for r in kv[1]))[:top]:
        fp, fc, _ = totals(group)
        out.append(f"  {name[:46]:46s} {len(group):5d}  {_fmt(fp + fc):>8}")

    out.append(f"\nTOP {top} ACCOUNT x STAGE (sent)                  calls     total")
    pairs: Dict[Tuple[str, str], List[Dict[str, Any]]] = defaultdict(list)
    for r in sent:
        if r.get("account"):
            pairs[(str(r.get("account")), str(r.get("stage") or "?"))].append(r)
    for (account, stage), group in sorted(pairs.items(), key=lambda kv: -sum(sum(_tokens(r)[:2]) for r in kv[1]))[:top]:
        ap_, ac, _ = totals(group)
        out.append(f"  {(account + ' / ' + stage)[:46]:46s} {len(group):5d}  {_fmt(ap_ + ac):>8}")
    return "\n".join(out)


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--since", default=datetime.now().strftime("%Y-%m-%d"),
                    help="ISO date or date-time, inclusive (default: today)")
    ap.add_argument("--until", default="", help="ISO date or date-time, exclusive (default: no end)")
    ap.add_argument("--file", default="", help="only calls whose databook file name contains this")
    ap.add_argument("--ledger", default=str(DEFAULT_LEDGER))
    ap.add_argument("--top", type=int, default=10)
    args = ap.parse_args(argv)
    path = Path(args.ledger)
    if not path.exists():
        print(f"no usage ledger at {path} -- it is written by every model call since the usage ledger "
              f"was added; a run before that, or a stub run, leaves none here")
        return 1
    rows = list(_rows(path, args.since, args.until, args.file))
    print(f"USAGE {args.since} -> {args.until or 'now'}"
          + (f"  file~{args.file!r}" if args.file else "") + f"  ({len(rows)} ledger line(s), {path})")
    if not rows:
        print("nothing in this window")
        return 0
    print(summarise(rows, top=args.top))
    return 0


if __name__ == "__main__":
    sys.exit(main())

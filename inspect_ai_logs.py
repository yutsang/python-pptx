"""Summarize AI-pipeline utilization/health from processing.log file(s).

Parses every "[Stage] Processed: <account> | Duration: ... | ... | Output
chars: ... | Completion tokens: ... | ... | Expected max output tokens: ..."
line fdd_utils/ai.py's ContentGenerationLogger writes (one per account per
stage: Generator / Auditor / Refiner / Validator), and reports, per stage:
  - item count, total/avg duration
  - token utilization = completion_tokens / expected_max_output_tokens
    (the ratio previously used to catch reasoning_effort="high" burning
    its whole completion budget on hidden reasoning with zero visible
    output -- this IS the "utilisation" number to watch)
  - empty-output count (Output chars == 0 -- a hard failure)
  - near-cap count (utilization >= 95% -- at real risk of truncation)

Usage:
  python inspect_ai_logs.py <processing.log or a directory to search> [more...]

A directory argument is searched recursively for every processing.log /
run_*/processing.log it contains, and each is reported separately plus
an overall combined total across all of them.
"""
import glob
import re
import sys
from collections import defaultdict
from pathlib import Path

LINE_RE = re.compile(
    r"\[(?P<stage>\w+)\] Processed: (?P<account>.+?) \| "
    r"Duration: (?P<duration>[\d.]+)s \| "
    r"Model: (?P<model>\S+) \| "
    r"Prompt chars: (?P<prompt_chars>\d+) \| "
    r"Output chars: (?P<output_chars>\d+) \| "
    r"Tokens used: (?P<tokens_used>\d+) \| "
    r"Prompt tokens: (?P<prompt_tokens>\d+) \| "
    r"Completion tokens: (?P<completion_tokens>\d+) \| "
    r"Total tokens: (?P<total_tokens>\d+) \| "
    r"Estimated prompt tokens: (?P<est_prompt_tokens>\d+) \| "
    r"Estimated output tokens: (?P<est_output_tokens>\d+) \| "
    r"Expected max output tokens: (?P<expected_max>\d+) \| "
    r"Token source: (?P<token_source>\S+)"
)


def find_log_files(paths):
    """Resolve CLI args to actual processing.log files.

    Windows cmd.exe (unlike bash/zsh) never expands "*" wildcards itself --
    it hands the literal string "run_*/processing.log" to this script, so a
    plain Path(p).is_file() check on that string always fails. glob.glob()
    does the expansion ourselves instead, which works the same on every
    shell. A plain directory argument (no wildcard) is still the simplest
    option -- it's searched recursively for every processing.log under it.
    """
    files = []
    for p in paths:
        path = Path(p)
        if path.is_file():
            files.append(path)
        elif path.is_dir():
            files.extend(sorted(path.rglob("processing.log")))
        else:
            matches = sorted(glob.glob(p, recursive=True))
            if matches:
                files.extend(Path(m) for m in matches if Path(m).is_file())
            else:
                print(f"! Not found: {p}")
    return files


def parse_file(path: Path):
    rows = []
    with open(path, encoding="utf-8", errors="replace") as f:
        for line in f:
            m = LINE_RE.search(line)
            if m:
                d = m.groupdict()
                rows.append({
                    "stage": d["stage"],
                    "account": d["account"],
                    "duration": float(d["duration"]),
                    "model": d["model"],
                    "output_chars": int(d["output_chars"]),
                    "completion_tokens": int(d["completion_tokens"]),
                    "expected_max": int(d["expected_max"]),
                    "token_source": d["token_source"],
                })
    return rows


def report(rows, label):
    print(f"\n=== {label} ({len(rows)} stage-items) ===")
    if not rows:
        print("  (no 'Processed:' lines matched -- wrong file, or empty run)")
        return
    by_stage = defaultdict(list)
    for r in rows:
        by_stage[r["stage"]].append(r)

    for stage, items in by_stage.items():
        n = len(items)
        total_duration = sum(i["duration"] for i in items)
        utils = [
            (i["completion_tokens"] / i["expected_max"]) if i["expected_max"] else 0.0
            for i in items
        ]
        avg_util = sum(utils) / n
        empty = [i for i in items if i["output_chars"] == 0]
        near_cap = [i for i, u in zip(items, utils) if u >= 0.95]
        over_cap = [i for i, u in zip(items, utils) if u > 1.0]
        sources = defaultdict(int)
        for i in items:
            sources[i["token_source"]] += 1
        print(f"  [{stage}] {n} items | total {total_duration:.1f}s | avg {total_duration / n:.2f}s/item "
              f"| avg utilization {avg_util:.0%} (completion_tokens / expected_max_output_tokens)")
        if len(sources) > 1 or "provider_usage" not in sources:
            print(f"    ! TOKEN SOURCE NOT PURELY provider_usage: {dict(sources)} "
                  f"-- non-provider_usage entries use an ESTIMATED completion_tokens (chars/4 heuristic), "
                  f"not the real API-reported figure, which can make utilization/near-cap numbers meaningless")
        if empty:
            print(f"    ! EMPTY OUTPUT (0 chars): {len(empty)} item(s) -> {', '.join(i['account'] for i in empty)}")
        if over_cap:
            over_cap_labels = [f"{i['account']}({i['token_source']})" for i in over_cap]
            print(f"    ! OVER 100% (completion_tokens > expected_max_output_tokens -- should be architecturally "
                  f"impossible if token_source=provider_usage; check the token_source flag above): "
                  f"{len(over_cap)} item(s) -> {', '.join(over_cap_labels)}")
        elif near_cap:
            print(f"    ! NEAR TOKEN CAP (>=95% utilization, risk of truncation): {len(near_cap)} item(s) -> "
                  f"{', '.join(i['account'] for i in near_cap)}")


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    files = find_log_files(sys.argv[1:])
    if not files:
        print("No processing.log files found.")
        sys.exit(1)

    all_rows = []
    for f in files:
        rows = parse_file(f)
        all_rows.extend(rows)
        report(rows, str(f))

    if len(files) > 1:
        report(all_rows, f"COMBINED across {len(files)} file(s)")


if __name__ == "__main__":
    main()

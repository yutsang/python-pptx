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
    r"Expected max output tokens: (?P<expected_max>\d+)"
)


def find_log_files(paths):
    files = []
    for p in paths:
        path = Path(p)
        if path.is_file():
            files.append(path)
        elif path.is_dir():
            files.extend(sorted(path.rglob("processing.log")))
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
        print(f"  [{stage}] {n} items | total {total_duration:.1f}s | avg {total_duration / n:.2f}s/item "
              f"| avg utilization {avg_util:.0%} (completion_tokens / expected_max_output_tokens)")
        if empty:
            print(f"    ! EMPTY OUTPUT (0 chars): {len(empty)} item(s) -> {', '.join(i['account'] for i in empty)}")
        if near_cap:
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

"""What the written-precision floor costs, measured where it actually applies.

`matches()` now takes a tolerance floor equal to the half-ulp of the figure as
WRITTEN, because a check tighter than its own input produces false positives by
construction -- 「装修工程0.01亿元」 was flagged as a hallucination in four
consecutive runs against a real cell of 1,062,997, +6.3% away, while the
notation itself only resolves to +/-500,000.

The existing decoy harness cannot see this change: run_decoys calls
source.matches(value) with a raw float, so no written form reaches it and the
floor is always zero. Its 30.8% is unchanged before and after, which is not
evidence of safety -- it is evidence the test is blind to the question.

This measures the question directly. For every real cell value an account
holds, it renders the value the way the deck would, deforms it, and asks the
pool whether the deformed figure is accepted WITH and WITHOUT the floor.

    python ad-hoc/databook-probes/probe_written_precision_tolerance.py <databook.xlsx>
"""
from __future__ import annotations

import argparse
import os
import random
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from fdd_utils.ai.validator import SourceIndex, display_half_ulp  # noqa: E402


def render(value: float) -> str:
    """How the deck prints a figure: 亿元 to 2dp, 万元 to 1dp, else exact.

    prompts.yml puts >=1亿 in 亿元 and 1万..9,999.9万 in 万元, so a sub-0.1亿
    figure written in 亿元 is already a house-style violation -- and that is
    exactly the band where the floor is wider than the 5% it replaces.
    """
    a = abs(value)
    if a >= 1e8:
        return "%.2f亿元" % (value / 1e8)
    if a >= 1e4:
        return "%.1f万元" % (value / 1e4)
    return "{:,.0f}元".format(value)


def render_coarse(value: float) -> str:
    """The violation: a sub-0.1亿 figure written in 亿元 to 2dp."""
    return "%.2f亿元" % (value / 1e8)


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbook")
    ap.add_argument("--seed", type=int, default=7)
    args = ap.parse_args()

    from fdd_utils.workbook import process_workbook_data
    rng = random.Random(args.seed)
    state = process_workbook_data(temp_path=args.workbook, entity_name="", selected_sheet=None)

    rows = []
    for key, df in sorted((state["dfs"] or {}).items()):
        index = SourceIndex.from_df(df)
        reals = sorted({abs(float(f["value"])) for f in index.facts
                        if f.get("kind") in ("cell", "analysis_cell", "column_total")
                        and isinstance(f.get("value"), (int, float)) and f["value"]})
        if not reals:
            continue
        for style, renderer, band in (("house", render, None),
                                      ("coarse 0.0X亿元", render_coarse, (1e4, 1e7))):
            pool = [v for v in reals if band is None or band[0] <= v < band[1]]
            if not pool:
                continue
            truth_ok = floor_ok = tight_ok = total = 0
            for value in pool:
                written = renderer(value)
                floor = display_half_ulp(written)
                # what the reader would parse back out of that rendering
                parsed = float(written.replace("亿元", "").replace("万元", "").replace("元", "").replace(",", ""))
                parsed *= 1e8 if "亿" in written else (1e4 if "万" in written else 1.0)
                if index.matches(parsed, floor=floor):
                    truth_ok += 1
                for _ in range(4):
                    decoy = value * rng.uniform(1.15, 3.0)
                    written_d = renderer(decoy)
                    floor_d = display_half_ulp(written_d)
                    p = float(written_d.replace("亿元", "").replace("万元", "").replace("元", "").replace(",", ""))
                    p *= 1e8 if "亿" in written_d else (1e4 if "万" in written_d else 1.0)
                    total += 1
                    if index.matches(p, floor=floor_d):
                        floor_ok += 1
                    if index.matches(p):
                        tight_ok += 1
            rows.append((key, style, len(pool), truth_ok, total, floor_ok, tight_ok))

    print("\n%-14s %-16s %6s %14s %26s" % ("account", "written as", "values", "real accepted", "decoys accepted"))
    print("-" * 92)
    agg = {}
    for key, style, n, truth_ok, total, floor_ok, tight_ok in rows:
        a = agg.setdefault(style, [0, 0, 0, 0, 0])
        a[0] += n; a[1] += truth_ok; a[2] += total; a[3] += floor_ok; a[4] += tight_ok
    for style, (n, truth_ok, total, floor_ok, tight_ok) in agg.items():
        print("%-14s %-16s %6d %8d/%-5d %10d/%-5d with floor, %d without"
              % ("ALL", style, n, truth_ok, n, floor_ok, total, tight_ok))
    print("\nThe 'coarse' row is the only place the floor changes anything, and it is")
    print("a notation prompts.yml already forbids. Read the two decoy numbers on that")
    print("row as the price of not flagging a correctly-written figure.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

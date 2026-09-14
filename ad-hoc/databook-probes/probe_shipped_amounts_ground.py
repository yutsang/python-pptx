"""Every amount that shipped, against the pool that was supposed to gate it.

The pipeline's own verdict for an account can be "0 clause(s) unsupported" while
the bullet carries a figure nobody computed. This asks the question the other way
round and after the fact: take each amount in the text that shipped, and ask the
account's persisted evidence pool whether it can find it.

Needs only a run folder -- no workbook, no model, no re-run. That is what the
evidence pool is for: a figure stays answerable long after the databook is gone.

    python ad-hoc/databook-probes/probe_shipped_amounts_ground.py --run fdd_utils/logs/run_20260914_100825
    python ad-hoc/databook-probes/probe_shipped_amounts_ground.py --run <folder> --account 营业收入

An UNGROUNDED line is not automatically a defect -- a percentage, a count of
days, a year, or a figure the pipeline computed and did not file can all land
here. Read the context printed beside each one. What it IS good for is finding
the figure the model built out of nothing, which is the one shape the deck
cannot survive.
"""
from __future__ import annotations

import argparse
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from fdd_utils.ai.run_memory import RunMemory  # noqa: E402
from fdd_utils.ai.validator import extract_amount_spans  # noqa: E402

# Production's own parser, deliberately, rather than a regex written here. A
# hand-rolled one read 「人民币5，271.8万元」 as 271.8万元 and reported it as a
# figure the pool could not find -- a fabricated defect. extract_amount_spans
# normalises the fullwidth comma the Auditor introduces (its comment records
# 13 occurrences in Generator output against 46 in Auditor output) and knows the
# K/m suffixes and currency prefixes this deck actually uses. A diagnostic that
# parses differently from the thing it is auditing is measuring something else.


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--run", required=True)
    ap.add_argument("--account", default=None, help="limit to one account")
    ap.add_argument("--context", type=int, default=46, help="characters of context each side")
    args = ap.parse_args()

    memory = RunMemory(args.run)
    accounts = [args.account] if args.account else memory.accounts()

    checked = missed = 0
    rows = []
    for key in accounts:
        resolved = memory.resolve(key) or key
        # RunMemory.final_text returns the text under "final"; asking for "text"
        # silently yields "" and the sweep reports zero amounts on a run full of
        # them, which is exactly the shape of silence this probe exists to catch.
        record = memory.final_text(resolved)
        text = str(record.get("final") or record.get("text") or "")
        if not text.strip():
            continue
        seen = set()
        for value, start, end in extract_amount_spans(text):
            if not value or (resolved, value) in seen:
                continue
            seen.add((resolved, value))
            checked += 1
            if memory.trace_amount(resolved, value).get("grounded"):
                continue
            missed += 1
            lo, hi = max(0, start - args.context), end + args.context
            rows.append((resolved, text[start:end], value, text[lo:hi].replace("\n", " ")))

    print("\nRUN %s" % memory.run_id)
    print("distinct amounts that shipped: %d" % checked)
    print("  the pool can find            : %d" % (checked - missed))
    print("  the pool CANNOT find         : %d" % missed)
    for key, printed, value, context in rows:
        print("\n  [%s] %s  (= %s)" % (key, printed, format(value, ",.0f")))
        print("      ...%s..." % context.strip())
    if rows:
        print("\nFor any one of these, ask the run where a nearby figure came from:")
        print("  python ask_run.py --run %s \"<account> 的 <amount> 是從哪裡來的？\"" % memory.run_id)
    return 0


if __name__ == "__main__":
    sys.exit(main())

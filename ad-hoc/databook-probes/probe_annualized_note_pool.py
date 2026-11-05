"""Which annualised remark figures the pool should hold, and which it must not.

SourceIndex multiplies EVERY number found in an account's notes/remarks by the
annualisation factor and files the result. The reason is documented and real: a
remark-only sub-line (a stamp-duty figure that never gets its own table row) has
no pre-calculated annualised column, so a correctly annualised quote was being
flagged as a hallucination -- confirmed from a screenshot, 「印花税…2026年1-6月
年化后为人民币4,485元」 in red because 4,485 was never in the pool.

But the multiplication is applied to every number in the blob, including ones
that are ALREADY a full year. A real run then grounded 「较2025年度1,910.0万元
减少63%」 against 955.0万元 x 2 -- the model had annualised an annual base, which
prompts.yml:806 forbids in as many words, and the pool agreed with it.

Both cases below are from real runs. The first MUST stay groundable, the second
MUST NOT.

    python ad-hoc/databook-probes/probe_annualized_note_pool.py
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import pandas as pd  # noqa: E402

from fdd_utils.ai.validator import SourceIndex  # noqa: E402


def _frame(rows, columns, notes, months=6):
    """One account as SourceIndex sees it: an analysis frame plus a remark blob."""
    df = pd.DataFrame(rows, columns=columns)
    analysis = pd.DataFrame(rows, columns=columns)
    df.attrs["prompt_analysis_df"] = analysis
    df.attrs["supporting_notes"] = notes
    df.attrs["annualization_months"] = months
    df.attrs["integrity"] = {"annualization_months": months, "statement_type": "IS"}
    return df


CASES = [
    # (label, df, amount to look for, must_ground, why)
    (
        "remark-only sub-line, annualised correctly",
        _frame(
            rows=[["印花税", 0.0, 0.0, 2242.0]],
            columns=["Description", "2024-12-31", "2025-12-31", "2026-06-30"],
            notes=["印花税本期发生额为人民币2,242元"],
        ),
        4484.0, True,
        "2,242 sits ONLY in the partial column -- x2 is a real annualisation "
        "and is the screenshot case the widening was added for",
    ),
    (
        "an already-annual figure, doubled",
        _frame(
            rows=[["营业收入-租赁费", 10763000.0, 9550000.0, 4941000.0]],
            columns=["Description", "2024-12-31", "2025-12-31", "2026-06-30"],
            notes=["2025年度租赁费为人民币955.0万元"],
        ),
        19100000.0, False,
        "9,550,000 is the 2025 FULL YEAR -- doubling it is the error "
        "prompts.yml:806 forbids, and the pool must not bless it",
    ),
]


def main() -> int:
    ok = True
    for label, df, amount, must_ground, why in CASES:
        index = SourceIndex.from_df(df)
        hit = index.matches(amount)
        grounded = bool(hit)
        good = grounded == must_ground
        ok = ok and good
        print("%s  %s" % ("    " if good else "FAIL", label))
        print("        want %s, got %s" % ("grounded" if must_ground else "NOT grounded",
                                           "grounded" if grounded else "NOT grounded"))
        if grounded:
            fact = hit[0] if isinstance(hit, (list, tuple)) else hit
            if isinstance(fact, dict):
                print("        matched %s / %s" % (fact.get("kind"), fact.get("row_desc")))
        print("        %s" % why)
        print()
    print("ALL CASES PASS" if ok else "FAILURES ABOVE")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())

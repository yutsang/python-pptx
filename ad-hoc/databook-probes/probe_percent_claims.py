"""Sentences that state a percentage their own two amounts do not make.

`percent_claim_findings` costs a regeneration when it fires, so every case here
is a sentence a real run shipped (counterparty names removed), with whether a
reviewer should have been stopped by it.

    python ad-hoc/databook-probes/probe_percent_claims.py            # the cases
    python ad-hoc/databook-probes/probe_percent_claims.py --archive  # every archived run under fdd_utils/logs

--archive prints each sentence the check would flag in text the pipeline has
already shipped: the false-positive read before trusting it on a paid run.
Add a case whenever a run flags something it should not. Do not delete one.
"""
from __future__ import annotations

import glob
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from fdd_utils.ai.validator import percent_claim_findings  # noqa: E402

CASES = [
    ("2026年1-3月期间主营业务收入年化后为3,675.8万元，较2025年度的2,787.2万元下降67.0%，主要系出租率尚处于爬坡期所致。",
     True, "3,675.8 is the LARGER figure; -67.0% is the raw three months over the full year"),
    ("本科目于2026年1-3月（年化）为82.0万元，较2025年度75.7万元下降72.9%，主要系会计服务费及审计费大幅减少所致。",
     True, "82.0 against 75.7 is +8.3%"),
    ("2026年1-3月期间年化后为3,499.5万元，较2025年度4,003.5万元下降约12.3%。",
     False, "-12.6% written as about 12.3%: inside rounding plus a point"),
    ("2026年1-3月期间财务费用合计366.1万元，较2025年12月31日年化金额1,463.3万元下降75.0%。",
     False, "arithmetic is right; the fault is comparing three months with a year, which this check cannot see"),
    ("2024年末至2025年末，科目余额由199.6万元增至364.2万元，增长82.5%，主要系原值大幅增加所致。",
     False, "199.6 to 364.2 is +82.5%"),
    ("2024年至2025年期间，该科目由346.9万元增长至529.2万元，增幅52.5%，主要源于进项税额的增加。",
     False, "346.9 to 529.2 is +52.5%"),
    ("2024年12月31日余额为103.0万元，与2023年12月31日的244.0万元相比下降58%。",
     False, "103.0 against 244.0 is -57.8%"),
    ("本科目于2025年12月31日为183.5万元，2026年1至3月增长至295.2万元，增幅达61%。",
     False, "no 由/较 joining the two: not a pair this check reads"),
    ("房产税和土地使用税分别下降79.6%和75.0%，合计减少526.9万元。",
     False, "分别 makes it a series"),
    ("未分配利润由负2,873.9万元降至负6,372.7万元，降幅达121.7%。",
     False, "both negative: magnitude only, and 2,873.9 to 6,372.7 is 121.7%"),
]


def archive() -> int:
    import yaml
    flagged = scanned = 0
    for path in sorted(glob.glob(os.path.join("fdd_utils", "logs", "run_*", "results.yml"))):
        try:
            results = yaml.safe_load(open(path, encoding="utf-8")) or {}
        except Exception:
            continue
        for key, record in results.items():
            if not isinstance(record, dict) or str(key).startswith("__"):
                continue
            text = str(record.get("final") or record.get("subagent_4") or "")
            if not text:
                continue
            scanned += 1
            for finding in percent_claim_findings(text):
                flagged += 1
                print(f"{os.path.basename(os.path.dirname(path))} [{key}] stated {finding['stated_pct']:g}% "
                      f"actual {finding['actual_pct']:+.1f}%\n    {finding['clause'][:200]}")
    print(f"\n{flagged} sentence(s) flagged across {scanned} archived account text(s)")
    return 0


def main() -> int:
    if "--archive" in sys.argv:
        return archive()
    failed = 0
    for text, expect, why in CASES:
        got = bool(percent_claim_findings(text))
        failed += got != expect
        print(f"{'ok  ' if got == expect else 'FAIL'} {'flag ' if got else 'quiet'} {why}")
    print("\nALL CASES PASS" if not failed else f"\n{failed} CASE(S) FAILED")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())

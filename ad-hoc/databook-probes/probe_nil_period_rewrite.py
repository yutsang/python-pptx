"""Sentences a real portfolio deck shipped, and what they should read as.

prompts.yml line 212 already forbids writing a nil period inside a multi-period
enumeration, and gives a worked example of the wrong form. A seven-entity run
shipped roughly forty of exactly that form. This is the deterministic rewrite
that stops asking.

    python ad-hoc/databook-probes/probe_nil_period_rewrite.py

Every BEFORE below is copied from that deck's own text dump, with counterparty
names replaced. Add a case whenever a run ships one this misses; the last two
are sentences it must NOT touch.
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from fdd_utils.pptx.payloads import (  # noqa: E402
    rewrite_nil_periods_out_of_series as fix,
    round_to_house_precision,
)

CASES = [
    ("代理及佣金于2023年度、2024年度、2025年度及2026年1-6月期间分别为0万元、2.4万元、0万元及4.4万元。",
     "0万元 should leave the list", True),

    ("管理费用-行政分别为0万元、70.0万元、30.1万元及15.5万元。",
     "no 于…期间 frame, only the amounts -- nothing to pair them with", False),

    ("税金及附加-房产税于2023年度、2024年度、2025年度及2026年1-6月期间分别为0万元、255.9万元、267.5万元及97.9万元；",
     "one nil period at the front", True),

    ("折旧摊销于2023年度、2024年度、2025年度及2026年1-6月期间分别为0.0万元、1,732.9万元、1,731.9万元及865.9万元",
     "0.0 counts as nil too", True),

    # --- reversed order: amounts first, then the frame --------------------
    ("代理及佣金分别为0万元、15.0万元、4.4万元及0万元，于2023年度、2024年度、2025年度及2026年1至6月期间发生",
     "amounts BEFORE the frame -- seven of the eight survivors of the first pass", True),

    ("三代手续费返还分别为0万元、9.6万元、0万元、0万元，于2023年度、2024年度、2025年度及2026年1至6月期间发生",
     "one real period out of four", True),

    ("加计抵减分别为0万元、0.4万元、0万元、0万元，于同期发生",
     "KNOWN LIMIT: 于同期 names no periods to pair with", False),

    ("代理及佣金在2023年度、2024年度、2025年度及2026年1-6月期间分别为0万元、29.4万元、20.4万元及7.0万元",
     "在, not 于", True),

    ("营业成本-折旧成本：2023年度、2024年度、2025年度及2026年1-6月分别为0万元、1,654.9万元、1,655.1万元及827.6万元",
     "no lead word at all, just a colon -- and none is added back", True),

    # --- KNOWN LIMIT: no period frame next to the amounts, so nothing to pair --
    ("税金及附加-土地使用税同期分别为0万元、47.3万元、47.3万元及23.7万元",
     "KNOWN LIMIT: 同期 refers to a frame in an earlier sentence", False),

    ("管理费用-行政分别为0万元、194.0万元、108.5万元及61.9万元",
     "KNOWN LIMIT: no frame anywhere -- which period is nil cannot be known", False),

    # --- must not be touched ----------------------------------------------
    ("营业收入-租赁费于2024年度、2025年度及2026年1-6月分别为1,076.3万元、955.0万元及494.1万元。",
     "no nil period at all", False),

    ("展览费于2023年度、2024年度、2025年度及2026年1-6月期间分别为0万元、0万元、0万元及0万元。",
     "nil in EVERY period is a different rule -- the component should be dropped, not rewritten", False),
]


def main() -> int:
    ok = True
    for before, why, expect_change in CASES:
        after = fix(before)
        changed = after != before
        good = changed == expect_change
        ok = ok and good
        print("%s  %s" % ("    " if good else "FAIL", why))
        print("      BEFORE %s" % before)
        if changed:
            print("      AFTER  %s" % after)
        else:
            print("      AFTER  (unchanged)")
        print()
    # 万元 one decimal, 亿元 two -- the same "stated in the prompt, ignored in
    # the output" shape, from the same deck.
    print("-" * 72)
    for before, after_want in [
        ("押金余额较上年末增长4.9787万元", "押金余额较上年末增长5.0万元"),
        ("净值为1.9234亿元", "净值为1.92亿元"),
        ("余额为1,128.3万元", "余额为1,128.3万元"),
        ("合计为5,930元", "合计为5,930元"),
        ("占比约4.9787%", "占比约4.9787%"),
    ]:
        got = round_to_house_precision(before)
        good = got == after_want
        ok = ok and good
        print("%s  %s -> %s" % ("    " if good else "FAIL", before, got))
    print()
    print("ALL CASES PASS" if ok else "FAILURES ABOVE")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())

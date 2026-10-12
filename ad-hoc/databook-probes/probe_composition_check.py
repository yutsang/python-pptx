"""Every composition flag a real portfolio run produced, and what it should say.

`check_composition_adds_up` is the single loudest check in the deliverable: on a
seven-entity run it produced 14 of the 14 grounding warnings. That makes its
false positives expensive twice over -- each one costs a reviewer a lookup, and
they bury the ones that are real.

Each case below is a clause a real run shipped, with the counterparty names
replaced. EXPECT_FLAG says whether a reviewer should be asked to look at it.

    python ad-hoc/databook-probes/probe_composition_check.py

Add a case whenever a run flags something it should not. Do not delete one.
"""
from __future__ import annotations

import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from fdd_utils.ai.validator import check_composition_adds_up  # noqa: E402

CASES = [
    # (key, text, expect_flag, why)
    ("管理费用",
     "管理费用主要包括「管理费用-员工」及「管理费用-行政」，于2023年度、2024年度、2025年度及"
     "2026年1-6月期间分别为0万元、182.1万元、150.0万元及65.9万元。2026年1-6月（年化）期间，"
     "管理费用合计为131.8万元，较2025年度下降6.8%。",
     False,
     "0+182.1+150.0+65.9 are FOUR PERIODS of one account, not four components"),

    ("货币资金",
     "截至2026年06月30日，货币资金余额为1,128.3万元，主要为银行存款-人民币户余额1,128.3万元，"
     "其中某银行某支行0758余额632.1万元，某银行某支行7831余额503.2万元。",
     False,
     "the 1,128.3 parent restates the total before breaking it down"),

    ("预付款项",
     "截至2026年3月31日，该科目无余额；2025年12月31日余额为16.6万元，主要为非工程款16.6万元，"
     "其中某保险公司16.6万元。",
     False,
     "其中 marks the second 16.6 as a drill-down of the first, not a sibling"),

    ("应付账款",
     "截至2026年06月30日，应付账款余额为9.5万元，主要包括：1）应付合同保留金-工程类38.2万元；"
     "2）预提管理费用14.2万元；3）物管类3.5万元。其他小额应付款项（如法律服务费、维修费等）"
     "合计为负42.1万元。",
     True,
     "still short, but by 45% not 488% -- the 其他…合计为负 closing IS counted now"),

    # --- these MUST keep firing -------------------------------------------
    ("应收账款",
     "截至2026年06月30日，应收账款余额合计110.1万元，主要包括：1）某物流22.6万元；"
     "2）某供应链13.8万元；3）某科技12.7万元。",
     True,
     "49.1万元 of a 110.1万元 account is simply not described"),

    ("应付账款",
     "截至2026年6月30日，应付账款余额合计为224.3万元，主要包括：1）预提开发成本121.1万元；"
     "2）应付合同保留金-工程类29.5万元；3）物管类13.6万元。其余56.4万元为预提管理费用等。",
     True,
     "3.7万元 short with every component named -- real, but the least urgent of the 14"),

    ("货币资金",
     "截至2026年06月30日，货币资金余额为694.4万元，主要包括：某支行0682 721.7万元、"
     "某支行5515 403.7万元、某支行5515 403.0万元。",
     True,
     "three accounts sum to 220% of the total; the negative ones are not mentioned"),

    ("营业成本",
     # No 主要包括/包含 anywhere, so no enumeration is recognised at all. Recorded
     # as a case because it looks like one this check should catch and does not:
     # a bare 其中 breakdown is invisible to it.
     "营业成本余额合计968.3万元，其中折旧成本为827.6万元，物管费用为73.9万元。",
     False,
     "KNOWN BLIND SPOT: a bare 其中 breakdown carries no 主要包括 to anchor on"),

    ("其他应收款",
     "截至2024年12月31日，其他应收款余额为3,782.9万元，主要包括关联方-美元项下的3,649.7万元，"
     "系目标公司向关联方某境外公司提供的500万美元借款本金（按汇率折合人民币3,649.7万元），"
     "期限1年，利率6%。",
     False,
     "one component, restated in brackets as its own currency conversion -- not two"),

    ("长期借款",
     "截至2026年03月31日，长期借款余额为3.26亿元，主要包括工行借款3.63亿元，"
     "扣除一年内到期的非流动负债0.37亿元。",
     False,
     "扣除 takes the current portion OFF the loan: 3.63 - 0.37 = 3.26, not 4.00"),


    ("长期应付款",
     "截至2026年3月31日，长期应付款余额为1,062.8万元，主要包括：1）押金及保证金1,034.3万元；"
     "2）某企业发展585.2万元。",
     True,
     "items exceed the total: the prior year's balance and one tenant inside it, listed as peers"),


    ("长期应付款",
     "截至2026年3月31日，长期应付款余额为1,062.8万元，主要包括押金及保证金1,034.3万元。"
     "与其他应付款中账龄小于1年的租赁保证金合计1,109.2万元，两项合计构成租户租赁保证金总额。",
     False,
     "其他应付款 is an account name, not this account's closing 其他…合计 component"),

    ("应收账款",
     "截至2026年03月31日，应收账款余额为674.5万元，主要包括：1）租金收入434.2万元，2）物业管理费收入163.8万元，"
     "3）水电费收入39.2万元，其余27.3万元为截止性调整及补计提。",
     True,
     "the stated remainder is wrong: 37.3 is left, not 27.3"),

    ("长期待摊费用",
     "截至2026年3月31日，长期待摊费用余额为355.8万元，系原值653.3万元减累计摊销297.5万元后的净值。"
     "原值主要包括：1）项目管理费569.5万元（按5年摊销）；2）绿化改造工程83.9万元（按5年摊销）。",
     False,
     "the list is introduced as 原值's and ties to the 653.3 stated for it, not to the net balance"),

    ("长期待摊费用",
     "截至2026年3月31日，长期待摊费用余额为355.8万元，系原值653.3万元减累计摊销297.5万元后的净值。"
     "原值主要包括：1）项目管理费569.5万元；2）绿化改造工程183.9万元。",
     True,
     "same shape, but the items do NOT tie to the stated 原值 either -- still reported"),

]


def main() -> int:
    ok = True
    print("%-10s %-6s %-11s %s" % ("account", "want", "got", "why"))
    print("-" * 100)
    for key, text, expect, why in CASES:
        flags = check_composition_adds_up(key, text)
        got = bool(flags)
        pct = ""
        if flags:
            _m = re.search(r"\((\d+)%\)", flags[0])
            pct = (" gap=" + _m.group(1) + "%") if _m else ""
        mark = "  " if got == expect else "FAIL"
        print("%-10s %-6s %-11s %s%s" % (key, "flag" if expect else "quiet",
                                        ("flag" + pct) if got else "quiet", why,
                                        ("   <-- " + mark) if mark != "  " else ""))
        if flags and got != expect:
            print("           %s" % flags[0][:150])
        ok = ok and (got == expect)
    print("\n%s" % ("ALL CASES PASS" if ok else "FAILURES ABOVE"))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())

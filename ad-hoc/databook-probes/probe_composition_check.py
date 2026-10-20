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

    ("长期待摊费用",
     "截至2026年3月31日，长期待摊费用余额为355.8万元，系项目管理费和绿化改造工程费摊销后净值。"
     "原值合计653.3万元，其中项目管理费569.5万元、绿化改造工程83.9万元；减累计摊销297.5万元后账面净值355.8万元。",
     False,
     "原值合计…其中 introduces the gross cost's own breakdown, as 原值主要包括 does"),

    ("固定资产",
     "截至2026年03月31日，固定资产账面价值为6.65亿元，主要为一期房屋2.79亿元、二期房屋2.64亿元及土地-1.10亿元。",
     True,
     "the -1.10 is accumulated depreciation labelled as land; the list reaches 4.33 of 6.65"),

    ("固定资产",
     "截至2026年03月31日，固定资产账面净值为6.65亿元，包括：1）房屋一期2.79亿元；2）房屋二期2.64亿元；"
     "3）土地一期0.61亿元；4）土地二期0.60亿元；5）一期房屋改造0.01亿元。",
     False,
     "the same account listed correctly ties to its net book value"),

    ("应付账款",
     "截至2026年03月31日，应付账款余额为877.2万元，主要包括中国电子系统工程第四建设工程款731.3万元及其他应付款项124.7万元。",
     True,
     "a closing 其他… item that does not close: 877.2 - 731.3 = 145.9, not 124.7"),

    ("应付账款",
     "截至2026年03月31日，应付账款余额为877.2万元，主要包括中国电子系统工程第四建设工程款731.3万元及其他应付款项145.9万元。",
     False,
     "the same sentence with the right remainder"),

    ("货币资金",
     "截至2026年03月31日，货币资金余额为1.23亿元，均为银行存款，其中某银行美元资本金户0.84亿元、外债户0.37亿元。余额较2023年12月31日的1.81亿元下降，预计主要系基本户余额由2023年12月31日的0.58亿元降至2026年03月31日的0.01亿元所致。",
     False,
     "主要系…所致 is a cause, not a list: the two figures are one account in two periods"),

    ("长期借款",
     "截至2026年03月31日，长期借款余额为3.26亿元，系工行借款3.63亿元扣除重分类至一年内到期的非流动负债0.37亿元后的余额。工行借款主要包括一期物业借款本金1.40亿元及二期物业借款本金2.20亿元。",
     False,
     "the second list composes the 3.63 工行借款, within rounding (3.60 printed)"),

    ("其他应付款",
     "截至2026年03月31日，其他应付款余额82.7万元，主要包括押金及保证金46.4万元、预提费用30.1万元，另有第三方款项3.8万元及施工保证金2.5万元；其中押金及保证金主要为某物流32.0万元，预提费用主要为审计费24.1万元。",
     False,
     "a ；其中 drill-down after the list is inside its items, not more of them"),

    ("其他应付款",
     "截至2026年03月31日，其他应付款余额82.7万元，主要包括押金及保证金46.4万元、预提费用10.1万元；其中押金及保证金主要为某物流32.0万元。",
     True,
     "the same drill-down does not hide a real shortfall (56.5 of 82.7)"),

    # Signed components written the two ways real bullets write them
    # (Portfolio I run, 2026-10-16): a 负 in front, or 「…N万元抵减」 after.
    ("应交税费",
     "截至2026年06月30日，应交税费余额36.6万元，主要包括应交房产税44.0万元、应交土地使用税11.8万元，并由应交增值税留抵净额20.1万元抵减，其余0.9万元为管理层调整。",
     False,
     "44.0 + 11.8 - 20.1 + 0.9 = 36.6: the 抵减 item is negative"),

    ("长期应付款",
     "截至2026年06月30日，长期应付款余额为5,398.0万元，包括股东贷款本金7,700.0万元、应计利息1,121.6万元及管理层调整负3,423.6万元。",
     False,
     "7,700.0 + 1,121.6 - 3,423.6 = 5,398.0: 负 marks the adjustment negative"),

    ("应付账款",
     "截至2026年06月30日，应付账款余额合计9.5万元，主要包括工程类合同保留金38.2万元、预提管理费用14.2万元、物管类3.5万元、非工程类合同保留金2.1万元及其他零星项目4.8万元，并由某企业管理53.3万元借方余额抵减。",
     False,
     "a few words between the amount and 抵减 still make it a deduction"),

    ("预付款项",
     "截至2026年06月30日，预付款项余额为0.3万元，主要包括预付账款-物业运营类3.7万元，并被预付账款-营销费用类贷方余额3.4万元抵减，其中贷方余额主要为某公司。",
     False,
     "3.7 - 3.4 = 0.3"),

    ("应付账款",
     "截至2026年06月30日，应付账款余额合计81.7万元，主要包括预提管理费用82.2万元，并由某企业管理负数余额8.9万元抵减。",
     True,
     "82.2 - 8.9 = 73.3, not 81.7: a real shortfall stays reported"),

    ("固定资产",
     "截至2026年06月30日，固定资产原值2.00亿元，减累计折旧0.30亿元后账面净值为1.70亿元；原值包括房屋建筑物1.80亿元、机械设备0.13亿元及室外工程0.07亿元。",
     False,
     "a bare 包括 under 原值 composes the 2.00 gross cost, not the 1.70 net"),

    ("应付账款",
     "截至2026年06月30日，应付账款余额为60.0万元，主要包括合同保留金60.0万元、预提费用10.0万元、某服务费借方余额10.0万元；其中合同保留金主要对应甲公司40.0万元及乙公司20.0万元。",
     False,
     "60 + 10 - 10 = 60: a named 借方余额 inside a liability list is contra"),

    ("货币资金",
     "截至2026年06月30日，货币资金余额1,000.0万元，主要为银行存款人民币户1,000.0万元，其中某银行A、B及C账户分别为300.0万元、200.0万元及100.0万元，其余400.0万元为其他银行账户。",
     False,
     "分别 after three named accounts is not a period series; parent restates the total"),

    ("其他应收款",
     "截至2026年06月30日，其他应收款余额为4.0万元，主要为营运收费4.0万元，其中金额较大的应收对象为甲公司2.0万元、乙公司1.0万元及丙公司0.6万元。",
     False,
     "其中金额较大的 says the list is partial on purpose"),

    ("应付账款",
     "截至2026年06月30日，应付账款无余额；2025年12月31日余额为280.0万元，主要包括工程类、预提费用及预提开发成本，明细如下：\n2025年审计调整将6.0万元由应付账款重分类至其他应付款；2024年同口径调整6.0万元。",
     False,
     "the list names no amounts; the next line's adjustments are not its items"),

    ("货币资金",
     "截至2026年06月30日，货币资金余额为700.0万元，主要包括：1）银行存款人民币户1,400.0万元；2）银行存款美元户0.1万元；3）管理层调整抵减700.1万元。",
     False,
     "a numbered item written 抵减X is negative: 1,400.0 + 0.1 - 700.1 = 700.0"),

    ("应收账款",
     "截至2026年06月30日，应收账款余额合计100.0万元，主要包括：1）甲公司800.0万元；2）乙公司90.0万元；3）丙公司8.0万元；其余7.0万元为其他零星客户余额，并扣减管理层调整805.0万元。",
     False,
     "a deduction closing a numbered list: 800 + 90 + 8 + 7 - 805 = 100"),
    ("其他应收款",
     "截至2026年06月30日，其他应收款余额为5.0万元，主要为营运收费应收款120.4万元经管理层调整-115.3万元后的净额；营运收费应收款主要包括甲公司110.0万元、乙公司6.0万元、丙公司2.8万元，其余1.6万元为其他零星项目。",
     False,
     "a second list after ； breaks down the first item; 120.4 - 115.3 = 5.1 is 5.0 within display rounding"),

    ("预付款项",
     "截至2026年3月31日，预付款项无余额。2025年12月31日余额为16.6万元，全部为预付某保险公司的保险费，该余额已于"
     "2026年一季度内结转。2023年12月31日余额23.6万元中，除预付保险费外，还包括预缴所得税16.1万元及预付水电费7.5万元。",
     False,
     "the list breaks down the 2023 balance stated just before it, not the 2025 one (Crescent 2026-10-20: COMPOSITION_EXCEEDS)"),

    ("货币资金",
     "截至2026年3月31日，货币资金余额为1.23亿元，均为银行存款，无使用受限资金。余额主要存放于某银行，其中美元资本金户"
     "0.84亿元、外债户0.37亿元、基本户0.01亿元。2023年末以来余额逐年下降，由1.81亿元降至1.23亿元，主要系基本户资金由"
     "0.58亿元降至0.01亿元，预计与日常经营支出及资本金逐步结汇使用有关。",
     False,
     "主要系…由0.58亿元降至0.01亿元 is one account in two periods, not a two-item list (Crescent 2026-10-20: 52% gap, shipped red)"),

    ("应付账款",
     "截至2026年3月31日，应付账款余额合计877.2万元，其中对某建设公司的应付工程款为731.3万元，占总额的83.4%。"
     "该款项包括甲合同尾款1.1万元、乙总包质保尾款725.0万元，以及维修工程款5.1万元。"
     "其余主要应付款项为关联方某管理公司的管理费67.8万元、某物业的物业管理费51.5万元、某律所的法律服务费5.0万元、"
     "某监理尾款4.8万元，以及水电费9.9万元和其他若干小额质保金、维修款合计约7.0万元。",
     False,
     "731.3+67.8+51.5+5.0+4.8+9.9+7.0 = 877.3 ties; 「其余…为…67.8万元、」 opens a list, it is not a remainder "
     "(Crescent 2026-10-20: 67.8 was patched to 146.0 and shipped)"),

    ("其他应收款",
     "截至2026年06月30日，其他应收款余额为5.0万元，主要为营运收费应收款120.4万元经管理层调整-105.3万元后的净额。",
     True,
     "120.4 - 105.3 = 15.1, three times the 5.0 stated: rounding cannot explain it"),
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

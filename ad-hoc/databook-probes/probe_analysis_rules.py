"""Regression cases for the analyst rules, built from REAL run output.

The rules can only be judged against a real databook, and the workbooks that
exposed their defects are the client's -- gitignored, and not on the machine
that writes the rules. So each case here transcribes the figures a real run
PRINTED, and asserts what the rules must and must not say about them. That is
weaker than the workbook (it cannot catch an extraction defect) and stronger
than nothing (it pins every noise finding that has actually shipped).

    python ad-hoc/databook-probes/probe_analysis_rules.py

Add a case whenever a run reports something wrong. Do not delete one.
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from fdd_utils.ai.analysis import AccountView, analyse  # noqa: E402

P0, P1, P2 = "2024-12-31", "2025-12-31", "2026-06-30"


def view(key, stype, series, components=None, has_detail=True):
    return AccountView(key=key, statement_type=stype,
                       series={p: float(v) for p, v in series.items()},
                       components={k: {p: float(v) for p, v in cols.items()}
                                   for k, cols in (components or {}).items()},
                       has_detail=has_detail if components is None else bool(components))


def kunshan():
    """一间仓储实体, run of 2026-09-11. Figures as the run printed them.

    Six of its sixteen analyst findings were noise, and each is pinned below.
    """
    v = {}
    v["长期借款"] = view("长期借款", "BS", {P0: 100000000, P1: 135150000, P2: 135150000},
                       {"长期借款-抵押借款": {P2: 135150000}})
    # 财务费用 is not all interest: the tab carries five components.
    v["财务费用"] = view("财务费用", "IS", {P1: 5600000, P2: 2931898},
                       {"财务费用-利息支出": {P2: 2500000},
                        "财务费用-手续费": {P2: 131898},
                        "财务费用-汇兑损益": {P2: 300000}})
    v["营业收入"] = view("营业收入", "IS", {P1: 14000000, P2: 7771577},
                       {"营业收入-租赁费": {P2: 6000000}, "营业收入-物业费": {P2: 1771577}})
    v["应收账款"] = view("应收账款", "BS", {P2: 1803466},
                       {"C0170某某供应链管理集团有限公司": {P2: 28000},
                        "C0031某仓储客户": {P2: 700000},
                        "C0088某仓储客户": {P2: 620000},
                        "C0091某仓储客户": {P2: 455466}})
    v["预收款项"] = view("预收款项", "BS", {P2: 1133324},
                       {"预收账款-仓储服务": {P2: 1133324},
                        "C0170某某供应链管理集团有限公司": {P2: 32000},
                        "C0027某某(上海)供应链管理有限公司": {P2: -300000},
                        "C0044某客户": {P2: 1401324}})
    # 货币资金: a bank row annotated with the customer it collects for. The
    # 集团 substring reported the entity's own bank balance as intra-group.
    v["货币资金"] = view("货币资金", "BS", {P2: 11283455},
                       {"BANK-001某银行某支行0758,C0170某某供应链管理集团有限公司": {P2: 544000},
                        "BANK-003某银行某支行4968": {P2: -24161000},
                        "BANK-002某银行某支行": {P2: 34900455}})
    v["其他应收款"] = view("其他应收款", "BS", {P2: 2043151},
                        {"其他应收款-非关联公司-押金": {P2: 1876500},
                         "V0367某市财政局某分局": {P2: 1876500},
                         "C0170某某供应链管理集团有限公司": {P2: 1211},
                         "其他应收款-营运收费": {P2: 156651}})
    # The category rows and their members sit in the same column, so the top
    # "counterparty" was exactly 50% of the sum.
    v["其他应付款"] = view("其他应付款", "BS", {P2: 3566467},
                        {"其他应付款-非关联公司-押金": {P2: 2362000},
                         "其他应付款-非关联公司-其他": {P2: 1081016},
                         "应付利息": {P2: 123888},
                         "C0168某某网络科技有限公司": {P2: 600000},
                         "C0025某某信息科技有限公司": {P2: 450000},
                         "V0352某市地方税务局": {P2: 107000},
                         "C0180某某仓储物流有限公司": {P2: 43000}})
    v["应交税费"] = view("应交税费", "BS", {P2: 510790},
                       {"应交房产税": {P2: 431000},
                        "应交土地使用税": {P2: 71000},
                        "应缴税费-应交增值税-进项税额": {P2: -28037000}})
    # Balance flat across every period while 2.72亿 moved through the account.
    v["在建工程"] = view("在建工程", "BS", {P0: 12212, P1: 12212, P2: 12212},
                       {"在建工程-工程支出": {P2: 272012212},
                        "在建工程-转出成本": {P2: -272000000}})
    v["固定资产"] = view("固定资产", "BS", {P1: 198000000, P2: 191749229},
                       {"固定资产-房屋建筑物": {P2: 440000000},
                        "固定资产-机器设备": {P2: 17000000},
                        "累计折旧-房屋建筑物": {P2: -73068000}})
    return v, ["实收资本", "未分配利润", "营业外收入"]


def dongguan():
    """The workbook that first exercised RELATED_PARTY. Must keep working."""
    v = {}
    v["其他应收款"] = view("其他应收款", "BS", {P2: 71000000},
                        {"内部单位往来": {P2: 52090000},
                         "押金及保证金": {P2: 18910000}})
    return v, []


def check(title, views, unmatched, must, must_not):
    obs = analyse(views, unmatched)
    got = {o.code for o in obs}
    print("\n%s\n%s" % (title, "-" * len(title)))
    for o in obs:
        print("  [%-6s] %-24s %s" % (o.severity, o.code, o.finding))
    ok = True
    for code, why in must.items():
        if code not in got:
            print("  FAIL missing %s -- %s" % (code, why)); ok = False
    for code, why in must_not.items():
        hit = [o for o in obs if o.code == code and why(o)]
        if hit:
            print("  FAIL %s should not fire: %s" % (code, hit[0].finding)); ok = False
    return ok


def main() -> int:
    v, unmatched = kunshan()
    ok1 = check(
        "一间仓储实体 (real run 2026-09-11)", v, unmatched,
        must={
            "IMPLIED_RATE": "long-term loan + a finance-cost line are both present",
            "RECEIVABLE_DAYS": "AR and revenue are both present",
            "ADVANCE_COVER": "advances are 7% of annualised revenue",
            "DEPRECIATION_RATE": "gross and accumulated depreciation are both present",
            "CIP_TRANSFER": "2.72亿 left CIP through a 转出成本 line",
            "INPUT_VAT_CREDIT": "2803.7万 of input VAT sits in the tax account",
            "NEGATIVE_BALANCE": "a bank account at -2416.1万 is a real one",
            "UNTIED_EQUITY": "实收资本/未分配利润 have no tab",
        },
        must_not={
            # every noise finding the real run printed, by the exact shape of it
            "RELATED_PARTY": lambda o: "供应链管理集团" in o.finding or "货币资金" in o.accounts,
            "CONCENTRATION": lambda o: "-非关联公司-" in o.finding or "合计" in o.finding,
            "STATIC_BALANCE": lambda o: "在建工程" in o.accounts,
        },
    )
    # the sign rule must be silent on the two lines that are negative by design
    obs = analyse(*kunshan())
    for bad in ("转出成本", "进项税额"):
        if any(o.code == "NEGATIVE_BALANCE" and bad in o.finding for o in obs):
            print("  FAIL NEGATIVE_BALANCE still fires on %s" % bad); ok1 = False
    rate = next((o for o in obs if o.code == "IMPLIED_RATE"), None)
    if rate and "利息支出" not in rate.finding:
        print("  FAIL IMPLIED_RATE used the whole finance-cost account: %s" % rate.finding); ok1 = False

    v2, u2 = dongguan()
    ok2 = check("另一本底稿 (RELATED_PARTY must survive)", v2, u2,
                must={"RELATED_PARTY": "内部单位往来 is 73% of the account"}, must_not={})

    print("\n%s" % ("ALL CASES PASS" if (ok1 and ok2) else "FAILURES ABOVE"))
    return 0 if (ok1 and ok2) else 1


if __name__ == "__main__":
    sys.exit(main())

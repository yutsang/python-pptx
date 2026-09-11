"""Analyst observations over a run: cross-account arithmetic, concentration, scope.

The pipeline knows a great deal about a databook and says almost none of it.
`build_insight_summary` reports how the RUN behaved -- clauses flagged, retries,
reconciliation rows without a tab -- which is quality control, not diligence.
The questions it puts to the client are template-generated from nil-to-value
movements ("what drove X to appear from nil"), which is the shape of a question
rather than a finding.

This module is the other half: what a reviewer would notice reading the same
numbers. Every rule here is arithmetic over figures the run already holds, and
every observation carries the figures it was computed from, so the same
grounding that grades the deck grades these. No LLM, no external data, nothing
across a run boundary.

It is deliberately a CHECKLIST, not a model. An FDD reviewer opening a
warehouse-entity databook checks the same handful of things every time: what the
borrowing actually costs, whether receivables move with revenue, whether one
counterparty carries the account, whether construction in progress landed in
fixed assets, whether the equity can be tied to anything at all. Those are
written out below, each with the numeric test that fires it. A rule that cannot
be computed from the data stays silent rather than guessing -- silence is a
correct answer when the workbook does not carry the input.

Input is neutral (`AccountView`) so this runs two ways: during a run, off the
frames; and afterwards, off the persisted evidence pool, which is what lets the
chat agent answer "what should I be worried about" about a finished run.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Dict, Iterable, List, Optional, Tuple

# --- vocabulary -------------------------------------------------------------

#: Every observation carries one. Grep-able, stable, and the thing a reader
#: sorts by when the list gets long.
CODES = (
    "IMPLIED_RATE",          # what the borrowing actually costs
    "RECEIVABLE_DAYS",       # AR against revenue
    "ADVANCE_COVER",         # customer advances against revenue
    "CIP_TRANSFER",          # construction in progress landing in fixed assets
    "DEPRECIATION_RATE",     # accumulated depreciation against gross
    "CONCENTRATION",         # one counterparty carrying an account
    "GOVERNMENT_COUNTERPARTY",  # deposits with a bureau -- refundability
    "RELATED_PARTY",         # intra-group balances
    "NEGATIVE_BALANCE",      # a credit where a debit belongs, or the reverse
    "STATIC_BALANCE",        # no movement at all across every period
    "NO_DETAIL",             # a total with nothing behind it
    "UNTIED_EQUITY",         # equity lines with no supporting schedule
)

SEVERITY_ORDER = {"high": 0, "medium": 1, "low": 2}


@dataclass
class Observation:
    code: str
    severity: str
    accounts: List[str]
    finding: str                       # what is true, with the numbers in it
    question: str = ""                 # what to ask management, when there is one
    numbers: List[float] = field(default_factory=list)   # every figure quoted, base units
    basis: str = ""                    # the computation it came from

    def as_dict(self) -> Dict[str, Any]:
        return {
            "code": self.code, "severity": self.severity, "accounts": list(self.accounts),
            "finding": self.finding, "question": self.question,
            "numbers": [float(n) for n in self.numbers], "basis": self.basis,
        }


@dataclass
class AccountView:
    """One account as the rules need it. Built by an adapter, not by hand."""
    key: str
    statement_type: str = ""           # "BS" | "IS" | ""
    series: Dict[str, float] = field(default_factory=dict)          # period -> account total
    components: Dict[str, Dict[str, float]] = field(default_factory=dict)  # label -> period -> value
    notes: str = ""
    has_detail: bool = True

    @property
    def periods(self) -> List[str]:
        seen = list(self.series) or [p for cols in self.components.values() for p in cols]
        return sorted(dict.fromkeys(seen))

    def latest(self) -> Tuple[Optional[str], Optional[float]]:
        for p in reversed(self.periods):
            v = self.series.get(p)
            if v is not None:
                return p, float(v)
        return (self.periods[-1] if self.periods else None), None

    def component_latest(self) -> Dict[str, float]:
        period, _v = self.latest()
        if period is None:
            return {}
        return {k: float(cols.get(period, 0.0) or 0.0) for k, cols in self.components.items()}


# --- account identification -------------------------------------------------
#
# Matched on the account KEY and on its mapped English name, so a Chinese and an
# English workbook hit the same rule. Deliberately substring tests on terms that
# do not collide: 长期借款 / Long-term loans is not going to match anything else
# in a warehouse databook.

# Canonical mapping keys FIRST, because that is what dfs is keyed by on an
# English workbook ("AR", "OI", "NCA"), then the long forms a Chinese workbook
# uses. Two-letter keys are matched WHOLE, never as a substring: "OR" inside
# "Other current assets" would otherwise make every account an other-receivable.
_ROLE_KEYS = {
    "borrowing": ("Long-term loans", "Short-term loans", "NCL due within one year"),
    "interest": ("Fin Exp", "Financial expenses"),
    "revenue": ("OI", "Sales", "Operating income"),
    "receivable": ("AR",),
    "advance": ("Advances", "Advance payments received"),
    "cip": ("CIP",),
    "fixed_asset": ("NCA", "FA", "Investment properties"),
    "equity": ("Capital", "Share capital", "Paid-in capital", "Capital reserve",
               "R/E", "Retained earnings", "Reserve"),
    "other_receivable": ("OR",),
    "other_payable": ("OP",),
}

_ROLE_TERMS = {
    "borrowing": ("长期借款", "短期借款", "long-term loan", "long term loan", "bank loan", "bank borrowing"),
    "interest": ("财务费用", "finance cost", "financial expense", "interest expense"),
    "revenue": ("营业收入", "主营业务收入", "revenue", "operating income"),
    "receivable": ("应收账款", "accounts receivable", "trade receivable"),
    "advance": ("预收款项", "预收账款", "advances from customer", "contract liabilit", "deferred revenue"),
    "cip": ("在建工程", "construction in progress", "construction-in-progress"),
    "fixed_asset": ("固定资产", "fixed asset", "non-current asset", "investment propert", "property, plant"),
    "equity": ("实收资本", "资本公积", "未分配利润", "盈余公积",
               "share capital", "paid-in capital", "retained earning", "capital reserve"),
    "other_receivable": ("其他应收款", "other receivable"),
    "other_payable": ("其他应付款", "other payable"),
}

# Rows that are an adjustment or a balancing figure, not a real component. They
# legitimately sit on the opposite side of their account, so a sign test fires
# on every one of them: a first pass reported 管理层调整, 直线法调整, TB BS gap
# and 销项税额\已开票 as wrong-direction balances across two workbooks, which is
# four findings and no information. Same list prompts.py uses for the same
# reason (see adj_markers in _composition_guidance).
# A total row is not a counterparty. Left in, the largest "component" of every
# account was 「合计」 at exactly 50% of (components + total), which is an
# arithmetic artefact reported as a concentration risk.
_TOTALS = ("合计", "小计", "总计", "净值小计", "原值小计", "累计折旧小计",
           "total", "subtotal", "sub-total", "人民币千元", "人民币万元")

_ADJUSTMENT = ("管理层调整", "示意性调整", "直线法调整", "截止性调整", "重分类",
               "TB BS gap", "gap", "adjustment", "reclass", "rounding", "差额", "调整")

# Intra-group and shareholder markers. 非关联 is listed as a NEGATION because
# a real workbook labels arm's-length lines "其他应付款-非关联公司-押金", and a
# plain substring test on 关联公司 marks every one of them as related-party --
# the exact opposite of what the label says.
_RELATED = ("内部单位往来", "关联方", "关联公司", "关联企业", "母公司", "子公司", "同一控制",
            "股东借款", "股东往来", "投资基金", "合伙企业", "集团",
            "related part", "intercompany", "inter-company", "shareholder", "affiliate")
_NOT_RELATED = ("非关联", "非关聯", "third-party", "third party", "non-related")

_GOV = ("财政局", "税务局", "国家税务", "地方税务", "社保", "公积金中心", "管委会", "财政厅")
_CONTRA = ("累计折旧", "累计摊销", "减值准备", "跌价准备", "坏账准备",
           "accumulated depreciation", "accumulated amortis", "impairment", "provision")


def _has_role(view: AccountView, role: str, alias: str = "") -> bool:
    key = str(view.key).strip()
    if any(key.lower() == k.lower() for k in _ROLE_KEYS.get(role, ())):
        return True
    hay = ("%s %s" % (key, alias)).lower()
    return any(term.lower() in hay for term in _ROLE_TERMS.get(role, ()))


def _is_adjustment(label: str) -> bool:
    low = str(label).lower()
    return any(m.lower() in low for m in _ADJUSTMENT)


def _is_related(label: str) -> bool:
    low = str(label).lower()
    if any(n.lower() in low for n in _NOT_RELATED):
        return False
    return any(m.lower() in low for m in _RELATED)


def _is_total(label: str) -> bool:
    low = str(label).strip().lower()
    return any(m.lower() in low for m in _TOTALS)


def _is_contra(label: str) -> bool:
    low = str(label).lower()
    return any(m.lower() in low for m in _CONTRA)


def _find(views: Dict[str, AccountView], role: str,
          aliases: Optional[Dict[str, str]] = None) -> List[AccountView]:
    aliases = aliases or {}
    return [v for v in views.values() if _has_role(v, role, aliases.get(v.key, ""))]


def _period_months(a: str, b: str) -> Optional[float]:
    """Months between two ISO period ends, for annualising a stub."""
    m = re.match(r"(\d{4})-(\d{2})", str(a) or ""), re.match(r"(\d{4})-(\d{2})", str(b) or "")
    if not (m[0] and m[1]):
        return None
    ya, ma = int(m[0].group(1)), int(m[0].group(2))
    yb, mb = int(m[1].group(1)), int(m[1].group(2))
    months = (yb - ya) * 12 + (mb - ma)
    return float(months) if months > 0 else None


def _fmt(value: float) -> str:
    """A figure a Chinese FDD deck would print. Kept here rather than importing
    the display formatter, which needs a whole account's values to pick a unit."""
    a = abs(value)
    if a >= 1e8:
        return "%.2f亿元" % (value / 1e8)
    if a >= 1e4:
        return "%.1f万元" % (value / 1e4)
    return "{:,.0f}元".format(value)


# --- the rules --------------------------------------------------------------

def _rule_implied_rate(views: Dict[str, AccountView]) -> List[Observation]:
    """What the borrowing costs, against what the notes say it costs.

    The first thing a reviewer computes on a leveraged entity, and the pipeline
    has both halves already: the interest line and the loan balance. An implied
    rate far outside a plausible band means the interest is capitalised, the
    balance moved mid-period, or the rate in the notes is not the rate paid --
    each of which is a question worth asking."""
    out: List[Observation] = []
    loans = _find(views, "borrowing")
    costs = _find(views, "interest")
    if not loans or not costs:
        return out
    loan = max(loans, key=lambda v: abs(v.latest()[1] or 0.0))
    cost = costs[0]
    lp, lv = loan.latest()
    cp, cv = cost.latest()
    if not lv or not cv or lv == 0:
        return out
    # The interest line is a period figure; the loan is a balance. Annualise the
    # expense when the period is a stub, using the account's own period spacing.
    months = None
    periods = cost.periods
    if len(periods) >= 2:
        months = _period_months(periods[-2], periods[-1])
    annualised = abs(cv) * (12.0 / months) if months and months < 12 else abs(cv)
    rate = annualised / abs(lv) * 100.0
    if not (0.5 <= rate <= 25.0) or True:   # always report; the band decides severity
        severity = "medium" if 2.0 <= rate <= 8.0 else "high"
        note = ""
        if "LPR" in (loan.notes or "") or "lpr" in (loan.notes or "").lower():
            note = "借款条款提及 LPR 加点，"
        out.append(Observation(
            code="IMPLIED_RATE", severity=severity, accounts=[loan.key, cost.key],
            finding=("%s余额%s，%s年化后%s，隐含融资成本约%.1f%%。%s"
                     % (loan.key, _fmt(lv), cost.key, _fmt(annualised), rate, note)),
            question=("请确认该利率与借款合同条款是否一致，以及本期是否有利息资本化计入在建工程或固定资产。"
                      if severity == "high" or note else
                      "请确认借款合同的计息基准及是否存在资本化利息。"),
            numbers=[abs(lv), annualised], basis="interest / loan balance",
        ))
    return out


def _rule_receivable_days(views: Dict[str, AccountView]) -> List[Observation]:
    """Receivable days against revenue, and whether they moved."""
    out: List[Observation] = []
    ars, revs = _find(views, "receivable"), _find(views, "revenue")
    if not ars or not revs:
        return out
    ar, rev = ars[0], revs[0]
    ap, av = ar.latest()
    rp, rv = rev.latest()
    if not av or not rv or rv == 0:
        return out
    months = None
    if len(rev.periods) >= 2:
        months = _period_months(rev.periods[-2], rev.periods[-1])
    annual_rev = abs(rv) * (12.0 / months) if months and months < 12 else abs(rv)
    days = abs(av) / annual_rev * 365.0
    severity = "high" if days > 120 else ("medium" if days > 60 else "low")
    out.append(Observation(
        code="RECEIVABLE_DAYS", severity=severity, accounts=[ar.key, rev.key],
        finding=("%s余额%s，对应年化%s%s，应收账款周转天数约%.0f天。"
                 % (ar.key, _fmt(av), rev.key, _fmt(annual_rev), days)),
        question=("请提供应收账款账龄分析及主要客户的信用期，并说明是否存在逾期或需计提坏账的余额。"
                  if severity != "low" else ""),
        numbers=[abs(av), annual_rev], basis="AR / annualised revenue x 365",
    ))
    return out


def _rule_advance_cover(views: Dict[str, AccountView]) -> List[Observation]:
    """Customer advances as a share of revenue -- forward booking, or a refund
    exposure at completion."""
    out: List[Observation] = []
    advs, revs = _find(views, "advance"), _find(views, "revenue")
    if not advs or not revs:
        return out
    adv, rev = advs[0], revs[0]
    _ap, av = adv.latest()
    _rp, rv = rev.latest()
    if not av or not rv or rv == 0:
        return out
    months = None
    if len(rev.periods) >= 2:
        months = _period_months(rev.periods[-2], rev.periods[-1])
    annual_rev = abs(rv) * (12.0 / months) if months and months < 12 else abs(rv)
    share = abs(av) / annual_rev * 100.0
    if share < 3.0:
        return out
    out.append(Observation(
        code="ADVANCE_COVER", severity="medium" if share < 25 else "high",
        accounts=[adv.key, rev.key],
        finding=("%s余额%s，相当于年化收入%s的约%.0f%%。"
                 % (adv.key, _fmt(av), _fmt(annual_rev), share)),
        question="请说明该预收款对应的服务期间及交割后是否需要退还或由买方承担。",
        numbers=[abs(av), annual_rev], basis="advances / annualised revenue",
    ))
    return out


def _rule_cip_transfer(views: Dict[str, AccountView]) -> List[Observation]:
    """Construction in progress emptying while fixed assets do not fill.

    The normal pattern is CIP falling and FA rising by roughly the same amount
    in the same period. CIP that empties with no matching FA addition is either
    an expensed write-off or a transfer the extract did not capture, and either
    is worth a line."""
    out: List[Observation] = []
    cips, fas = _find(views, "cip"), _find(views, "fixed_asset")
    if not cips or not fas:
        return out
    cip, fa = cips[0], fas[0]
    periods = [p for p in cip.periods if p in fa.series or p in cip.series]
    if len(periods) < 2:
        return out
    prev, curr = periods[-2], periods[-1]
    cip_move = float(cip.series.get(curr, 0.0) or 0.0) - float(cip.series.get(prev, 0.0) or 0.0)
    fa_move = float(fa.series.get(curr, 0.0) or 0.0) - float(fa.series.get(prev, 0.0) or 0.0)
    if cip_move >= 0 or abs(cip_move) < 1000:
        return out
    covered = fa_move / abs(cip_move) * 100.0 if cip_move else 0.0
    if covered >= 60:
        return out
    out.append(Observation(
        code="CIP_TRANSFER", severity="high", accounts=[cip.key, fa.key],
        finding=("%s于%s至%s减少%s，同期%s仅变动%s，转固金额未能对应。"
                 % (cip.key, prev, curr, _fmt(abs(cip_move)), fa.key, _fmt(fa_move))),
        question="请提供在建工程转固定资产的明细及转固时点，并说明未转固部分的去向（费用化、处置或调整）。",
        numbers=[abs(cip_move), abs(fa_move)], basis="CIP movement vs FA movement, same periods",
    ))
    return out


def _rule_depreciation_rate(views: Dict[str, AccountView]) -> List[Observation]:
    """Accumulated depreciation against gross cost -- how far through its life
    the asset base is, and whether the charge is plausible."""
    out: List[Observation] = []
    for fa in _find(views, "fixed_asset"):
        latest = fa.component_latest()
        if not latest:
            continue
        gross = sum(v for k, v in latest.items() if v > 0 and not _is_contra(k))
        contra = sum(abs(v) for k, v in latest.items() if _is_contra(k))
        if gross <= 0 or contra <= 0:
            continue
        pct = contra / gross * 100.0
        severity = "medium" if pct < 60 else "high"
        out.append(Observation(
            code="DEPRECIATION_RATE", severity=severity, accounts=[fa.key],
            finding=("%s原值%s，累计折旧/摊销%s，已提比例约%.0f%%。"
                     % (fa.key, _fmt(gross), _fmt(contra), pct)),
            question=("请确认折旧年限、残值率及折旧方法，并说明是否存在已提足折旧但仍在使用的资产。"
                      if severity == "high" else
                      "请确认各类资产的折旧年限及残值率。"),
            numbers=[gross, contra], basis="sum(contra components) / sum(gross components)",
        ))
    return out


def _rule_concentration(views: Dict[str, AccountView], threshold: float = 50.0) -> List[Observation]:
    """One counterparty carrying an account. Reported for the balances where it
    is a risk statement -- receivables, advances, payables -- not for an asset
    register, where one large building is normal and says nothing."""
    out: List[Observation] = []
    for view in views.values():
        if not any(_has_role(view, r) for r in ("receivable", "advance", "other_receivable", "other_payable")):
            continue
        latest = {k: v for k, v in view.component_latest().items() if v and not _is_contra(k)}
        if len(latest) < 2:
            continue
        total = sum(abs(v) for v in latest.values())
        if total <= 0:
            continue
        name, value = max(latest.items(), key=lambda kv: abs(kv[1]))
        share = abs(value) / total * 100.0
        if share < threshold:
            continue
        # RELATED_PARTY already reports this line, and says more about it:
        # who it is and that it may not survive completion. Two findings on one
        # number reads as two problems.
        if _is_related(name):
            continue
        out.append(Observation(
            code="CONCENTRATION", severity="high" if share >= 75 else "medium",
            accounts=[view.key],
            finding=("%s前一大对手方「%s」%s，占本科目约%.0f%%（共%d个构成项）。"
                     % (view.key, name, _fmt(abs(value)), share, len(latest))),
            question="请说明该对手方的交易性质、结算安排及是否为关联方。",
            numbers=[abs(value), total], basis="largest component / sum of components",
        ))
    return out


def _rule_government_counterparty(views: Dict[str, AccountView]) -> List[Observation]:
    """A deposit held by a bureau. Refundability and timing are a standard
    completion-accounts question and the counterparty name is right there."""
    out: List[Observation] = []
    for view in views.values():
        latest = view.component_latest()
        hits = [(k, v) for k, v in latest.items() if v and any(g in str(k) for g in _GOV)]
        if not hits:
            continue
        total = sum(abs(v) for v in hits)
        names = "、".join(k for k, _v in hits[:3])
        out.append(Observation(
            code="GOVERNMENT_COUNTERPARTY", severity="medium", accounts=[view.key],
            finding=("%s中与政府部门往来合计%s（%s%s）。"
                     % (view.key, _fmt(total), names, "等" if len(hits) > 3 else "")),
            question="请说明该等款项的性质（保证金、退税、补贴）、退还条件及预计回收时点。",
            numbers=[total], basis="components whose counterparty name is a government body",
        ))
    return out


def _rule_related_party(views: Dict[str, AccountView]) -> List[Observation]:
    """Intra-group and shareholder balances, named and priced.

    The first thing a buyer asks of a target's balance sheet is which of it
    survives the transaction. An intra-group receivable is usually settled or
    waived at completion, a shareholder loan is usually repaid or capitalised,
    and neither is worth what the balance sheet says to a third party. The
    counterparty names are already in the component list."""
    out: List[Observation] = []
    for view in views.values():
        if view.statement_type == "IS":
            continue
        latest = {k: v for k, v in view.component_latest().items()
                  if v and not _is_contra(k) and _is_related(k)}
        if not latest:
            continue
        total_related = sum(abs(v) for v in latest.values())
        _p, total = view.latest()
        share = (total_related / abs(total) * 100.0) if total else 0.0
        names = "、".join(sorted(latest, key=lambda k: -abs(latest[k]))[:3])
        out.append(Observation(
            code="RELATED_PARTY", severity="high" if share >= 50 else "medium",
            accounts=[view.key],
            finding=("%s中关联方/股东往来合计%s%s（%s%s）。"
                     % (view.key, _fmt(total_related),
                        ("，占本科目约%.0f%%" % share) if total else "",
                        names, "等" if len(latest) > 3 else "")),
            question="请说明该等关联方往来的形成原因、计价与结算条款，以及交割前是否会结清、豁免或资本化。",
            numbers=[total_related] + ([abs(total)] if total else []),
            basis="components whose counterparty name is intra-group or a shareholder",
        ))
    return out


def _rule_negative_balance(views: Dict[str, AccountView]) -> List[Observation]:
    """A component sitting on the wrong side of its account."""
    out: List[Observation] = []
    for view in views.values():
        if view.statement_type != "BS":
            continue
        _p, total = view.latest()
        if not total:
            continue
        wrong = {k: v for k, v in view.component_latest().items()
                 if v and not _is_contra(k) and not _is_adjustment(k) and (v < 0) != (total < 0)}
        if not wrong:
            continue
        name, value = max(wrong.items(), key=lambda kv: abs(kv[1]))
        out.append(Observation(
            code="NEGATIVE_BALANCE", severity="medium", accounts=[view.key],
            finding=("%s合计%s，但构成项「%s」为%s，方向与科目相反（共%d项）。"
                     % (view.key, _fmt(total), name, _fmt(value), len(wrong))),
            question="请说明该反向余额的成因（多收、退款未结转、错记科目），以及是否应重分类。",
            numbers=[abs(value), abs(total)], basis="component sign vs account sign",
        ))
    return out


def _rule_static_balance(views: Dict[str, AccountView]) -> List[Observation]:
    """A balance that has not moved once across every period presented."""
    out: List[Observation] = []
    for view in views.values():
        # Share capital not moving is the normal case, not a finding.
        if _has_role(view, "equity"):
            continue
        values = [float(view.series[p]) for p in view.periods if view.series.get(p) is not None]
        values = [v for v in values if v]
        if len(values) < 3 or len(set(round(v, 2) for v in values)) != 1:
            continue
        out.append(Observation(
            code="STATIC_BALANCE", severity="low", accounts=[view.key],
            finding=("%s在%d个期间内余额均为%s，未发生任何变动。"
                     % (view.key, len(values), _fmt(values[0]))),
            question="请确认该余额是否仍具回收/清偿价值，以及为何未见结转或结算。",
            numbers=[values[0]], basis="identical balance in every presented period",
        ))
    return out


def _rule_no_detail(views: Dict[str, AccountView], floor: float = 1e6) -> List[Observation]:
    """A material total with nothing behind it -- a scope limitation, stated as
    one rather than left for the reader to notice."""
    out: List[Observation] = []
    for view in views.values():
        if view.has_detail and view.components:
            continue
        _p, total = view.latest()
        if not total or abs(total) < floor:
            continue
        out.append(Observation(
            code="NO_DETAIL", severity="medium", accounts=[view.key],
            finding="%s余额%s，底稿未提供任何构成明细。" % (view.key, _fmt(total)),
            question="请提供该科目的明细表及主要对手方/项目清单。",
            numbers=[abs(total)], basis="account total present, component set empty",
        ))
    return out


def _rule_untied_equity(views: Dict[str, AccountView], unmatched: Iterable[str]) -> List[Observation]:
    """Equity lines the reconciliation could not tie to a schedule."""
    names = [str(n) for n in (unmatched or [])
             if any(t.lower() in str(n).lower() for t in _ROLE_TERMS["equity"])
             or any(t.lower() == str(n).strip().lower() for t in _ROLE_KEYS["equity"])]
    if not names:
        return []
    return [Observation(
        code="UNTIED_EQUITY", severity="high", accounts=list(names),
        finding="权益类科目%s在底稿中没有对应明细表，无法核对。" % "、".join(names),
        question="请提供实收资本/资本公积/未分配利润的变动表及验资或股东决议文件。",
        numbers=[], basis="reconciliation rows with no supporting tab",
    )]


_RULES = (
    _rule_implied_rate,
    _rule_receivable_days,
    _rule_advance_cover,
    _rule_cip_transfer,
    _rule_depreciation_rate,
    _rule_concentration,
    _rule_related_party,
    _rule_government_counterparty,
    _rule_negative_balance,
    _rule_static_balance,
    _rule_no_detail,
)


def analyse(views: Dict[str, AccountView],
            unmatched_statement_rows: Optional[Iterable[str]] = None) -> List[Observation]:
    """Every observation the rules can make, most severe first.

    A rule that cannot compute stays silent; none of them guesses. Failures are
    contained per rule so one odd workbook cannot take the whole section down.
    """
    out: List[Observation] = []
    for rule in _RULES:
        try:
            out.extend(rule(views) or [])
        except Exception:
            continue
    try:
        out.extend(_rule_untied_equity(views, unmatched_statement_rows or []))
    except Exception:
        pass
    out.sort(key=lambda o: (SEVERITY_ORDER.get(o.severity, 9), o.code))
    return out


# --- adapters ---------------------------------------------------------------

def views_from_evidence(evidence: Dict[str, Any], facts: Optional[Dict[str, Any]] = None,
                        notes: Optional[Dict[str, str]] = None) -> Dict[str, AccountView]:
    """Rebuild the rules' input from persisted evidence -- no databook needed.

    Each `analysis_cell` fact carries the row it came from and the period column,
    so the analysis frame the Generator was shown reconstructs exactly. This is
    what lets the same observations be produced for a finished run months later.
    """
    facts = facts or {}
    series_all = (facts.get("series") or {})
    types = (facts.get("statement_types") or {})
    views: Dict[str, AccountView] = {}
    for key, ev in (evidence or {}).items():
        components: Dict[str, Dict[str, float]] = {}
        for fact in getattr(ev, "facts", None) or []:
            if fact.get("kind") != "analysis_cell":
                continue
            row, col = fact.get("row_desc"), fact.get("col_label")
            if not row or not col or _is_total(row):
                continue
            components.setdefault(str(row), {})[str(col)] = float(fact.get("value") or 0.0)
        views[key] = AccountView(
            key=key,
            statement_type=str(types.get(key) or ""),
            series={str(p): float(v) for p, v in (series_all.get(key) or {}).items()},
            components=components,
            notes=str((notes or {}).get(key) or ""),
            has_detail=bool(components),
        )
    return views


def views_from_frames(dfs: Dict[str, Any], facts: Optional[Dict[str, Any]] = None) -> Dict[str, AccountView]:
    """The same input, built during a run from the frames themselves."""
    facts = facts or {}
    series_all = (facts.get("series") or {})
    types = (facts.get("statement_types") or {})
    views: Dict[str, AccountView] = {}
    for key, df in (dfs or {}).items():
        attrs = getattr(df, "attrs", None) or {}
        analysis = attrs.get("prompt_analysis_df")
        components: Dict[str, Dict[str, float]] = {}
        if analysis is not None and getattr(analysis, "columns", None) is not None:
            try:
                label_col = analysis.columns[0]
                cols = [c for c in analysis.columns[1:] if not str(c).startswith("__")]
                declared = {str(c).strip()
                            for c in (getattr(analysis, "attrs", {}).get("component_descriptions") or [])}
                for _i, row in analysis.iterrows():
                    name = str(row[label_col]).strip()
                    # `component_descriptions` is the workbook's own answer to
                    # "which rows are components"; totals and section labels are
                    # not in it. Falling back to a name test only where the
                    # workbook did not record one.
                    if not name or (declared and name not in declared) or _is_total(name):
                        continue
                    slot = components.setdefault(name, {})
                    for col in cols:
                        value = row[col]
                        if isinstance(value, (int, float)) and value == value:
                            slot[str(col)] = slot.get(str(col), 0.0) + float(value)
            except Exception:
                components = {}
        notes_parts: List[str] = []
        for bucket in ("supporting_notes", "adjacent_detail_rows", "table_linked_remarks"):
            for item in (attrs.get(bucket) or []):
                notes_parts.append(str(item))
        views[key] = AccountView(
            key=key,
            statement_type=str(types.get(key) or (attrs.get("integrity") or {}).get("statement_type") or ""),
            series={str(p): float(v) for p, v in (series_all.get(key) or {}).items()},
            components=components,
            notes=" ".join(notes_parts)[:4000],
            has_detail=bool(components),
        )
    return views

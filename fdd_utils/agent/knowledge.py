"""What an account's paragraph is expected to know, and whether the evidence the writer was given holds it. Python; no model.

A buyer reading the paragraph on a loan expects its rate and maturity, on receivables their ageing, on a fixed asset its
depreciation policy, on cash what the balance is made of. On the client files these were the meanings the writer left out
(ageing, depreciation policy, loan terms) or the evidence did not carry (a cash balance whose nature the source states), and the
autonomous agent spent 53 calls on one loan whose terms were not in the workbook at all and never settled it. A paragraph that
stays silent on a gap looks complete; one that invents the missing term is wrong.

Here the controller decides from the account's own name which of these an account should have, looks for each in the free text
the writer is shown (the notes and the names of the detail lines), and lists what is not there. The list goes to the writer
("NOT IN EVIDENCE: do not state it; ask once with `look` if you think the workbook has it") and, whatever the writer then does,
to the reader as a source gap: the person checking the deck sees that the term is not in the workbook, not that the paragraph
forgot it.

The match is lexical and says so. A word in the evidence does not prove the paragraph states the term correctly (the verifier
and the reviewer do that); an absent word does not prove the workbook lacks it (the writer's lookup and the reader's check do
that). What this decides is only whether the writer was handed anything to say.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Dict, List, Tuple


@dataclass(frozen=True)
class Need:
    id: str
    label: str
    #: the account key (a canonical English name or the workbook's Chinese one) this applies to
    account: "re.Pattern[str]"
    #: facet name -> words that say it. A facet is covered when any of its words is in the evidence text.
    facets: Tuple[Tuple[str, "re.Pattern[str]"], ...]
    #: words to search the workbook with when the writer asks for the facet (search_notes terms; space separated, any may match)
    queries: Tuple[Tuple[str, str], ...]


def _re(pattern: str) -> "re.Pattern[str]":
    return re.compile(pattern, re.I)


NEEDS: Tuple[Need, ...] = (
    Need("loan_terms", "loan terms",
         _re(r"借款|贷款|貸款|loan|borrowing"),
         (("rate", _re(r"利率|年化|基点|基點|\bLPR\b|\bBP\b|\d\s*%|interest rate|per annum")),
          ("maturity", _re(r"到期|期限|还款|還款|偿还|償還|展期|结清|結清|\bmatur|repay|\bterm of\b"))),
         (("rate", "利率 年利率 LPR"), ("maturity", "到期 期限 还款"))),
    Need("ageing", "ageing of the balance",
         # trade and other receivables and payables, not every line that ends in "payable" (taxes, payroll, dividends have no ageing)
         _re(r"应收账款|應收賬款|应付账款|應付賬款|其他应收|其他應收|其他应付|其他應付|应收票据|应付票据|(?:accounts?|trade|other|notes?) (?:receivable|payable)|^AR$|^AP$|^OAR$|^OAP$"),
         (("ageing", _re(r"账龄|賬齡|逾期|\d\s*年以[内內上]|一年以[内內上]|ag(?:e)?ing|overdue|aged|past due")),),
         (("ageing", "账龄 逾期 aging ageing"),)),
    Need("depreciation_policy", "depreciation or amortisation policy",
         _re(r"固定资产|固定資產|投资性房地产|投資性房地產|无形资产|無形資產|使用权资产|使用權資產|长期待摊|長期待攤|fixed asset|investment propert|intangible|right-of-use|\bPPE\b|deferred expense"),
         (("policy", _re(r"年限|使用寿命|使用壽命|残值|殘值|直线法|直線法|折旧率|摊销率|useful li|straight-line|residual")),),
         (("policy", "折旧年限 使用寿命 直线法 摊销年限"),)),
    Need("cash_nature", "what the cash balance is made of",
         _re(r"货币资金|貨幣資金|现金|現金|银行存款|銀行存款|^cash|bank balance|cash and"),
         (("nature", _re(r"保证金|保證金|受限|冻结|凍結|质押|質押|定期|活期|通知存款|基本户|基本戶|结算户|結算戶|资本金户|restricted|pledged|frozen|time deposit|demand deposit|current account|escrow|deposits? with")),),
         (("nature", "受限 保证金 定期 活期 存款"),)),
)

#: word shown by a facet name, for the writer's prompt and the reader's flag
FACET_WORDS = {"rate": "interest rate", "maturity": "maturity or repayment date", "ageing": "ageing of the balance", "policy": "depreciation or amortisation policy",
               "nature": "nature of the balance (restricted, deposits, time or demand)"}


def needs_for(account: str) -> List[Need]:
    name = str(account or "").strip()
    return [n for n in NEEDS if n.account.search(name)]


def evidence_text(view: Dict[str, Any]) -> str:
    """The free text the writer is shown: each note and the name of each detail line. Refs, figures and keys are not text."""
    parts = [str(n.get("text") or "") for n in view.get("notes") or [] if isinstance(n, dict)]
    comp = view.get("components")
    lines = comp.get("lines") if isinstance(comp, dict) else comp
    parts += [str(c.get("line") or "") for c in lines or [] if isinstance(c, dict)]
    return "\n".join(parts)


def missing(account: str, view: Dict[str, Any]) -> List[Dict[str, str]]:
    """[{need, facet, label, words}] for every expected facet no word of which is in the evidence text. Empty for an account
    with no expectation, and for one whose evidence has them all."""
    text = evidence_text(view)
    out: List[Dict[str, str]] = []
    for need in needs_for(account):
        queries = dict(need.queries)
        for facet, words in need.facets:
            if not words.search(text):
                words = FACET_WORDS.get(facet, facet)
                out.append({"need": need.id, "facet": facet, "label": need.label if words == need.label else f"{need.label}: {words}",
                            "query": queries.get(facet, "")})
    return out


def describe(gaps: List[Dict[str, str]]) -> str:
    """The writer's line, or "" when nothing is missing."""
    if not gaps:
        return ""
    return ("NOT IN EVIDENCE (the notes and line names above say nothing on these; do not state them from habit):\n"
            + "\n".join(f"  - {g['label']}" + (f"  [search words: {g['query']}]" if g.get("query") else "") for g in gaps))

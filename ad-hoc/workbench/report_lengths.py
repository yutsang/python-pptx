"""Rebuild fdd_utils/report_lengths.yml from text dumps of finished decks. No model call; free.

    python ad-hoc/workbench/report_lengths.py <report text> [<report text> ...] [--write]

A dump is `inspect_pptx.py <deck> --dump-text` saved to a file. An account's paragraph is read the way score_reference.py reads
it (its name and a dash, up to the next such header; a 续 part is joined to it) and measured in characters with whitespace
removed. Each dump gives one number per account it writes, and the writer is told the median (fdd_utils/agent/lengths.py).
The table is rebuilt from exactly the dumps named, so name each report once. Without --write it only prints the table.
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import Dict, List

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from score_reference import _paragraphs, _read  # noqa: E402

# Account names looked for in a dump: the usual statement lines. A name not listed here is added with --account.
ACCOUNTS = (
    "货币资金 交易性金融资产 应收票据 应收账款 应收款项融资 预付款项 其他应收款 存货 合同资产 一年内到期的非流动资产 其他流动资产 "
    "长期应收款 长期股权投资 其他权益工具投资 投资性房地产 固定资产 在建工程 使用权资产 无形资产 开发支出 商誉 长期待摊费用 "
    "递延所得税资产 其他非流动资产 短期借款 应付票据 应付账款 预收款项 合同负债 应付职工薪酬 应交税费 其他应付款 "
    "一年内到期的非流动负债 其他流动负债 长期借款 应付债券 租赁负债 长期应付款 预计负债 递延收益 递延所得税负债 其他非流动负债 "
    "实收资本 股本 资本公积 盈余公积 未分配利润 营业收入 营业成本 税金及附加 销售费用 管理费用 研发费用 财务费用 其他收益 "
    "投资收益 公允价值变动收益 信用减值损失 资产减值损失 资产处置收益 营业外收入 营业外支出 所得税费用"
).split()

HEADER = """\
# Characters (whitespace removed) of each account's paragraph in finished team reports, one number per report.
# The writer is asked for about the median (fdd_utils/agent/lengths.py); an account not listed keeps the evidence rule.
# Built by: python ad-hoc/workbench/report_lengths.py <report text> [...] --write   (each report named once)
"""


def lengths(paths: List[str], extra: List[str]) -> Dict[str, List[int]]:
    names = list(dict.fromkeys(list(ACCOUNTS) + extra))
    table: Dict[str, List[int]] = {}
    for path in paths:
        paras = _paragraphs(_read(path), names)
        for name in names:
            chars = len(re.sub(r"\s", "", paras.get(name, "")))
            if chars:
                table.setdefault(name, []).append(chars)
    return table


def render(table: Dict[str, List[int]], reports: int) -> str:
    order = {name: i for i, name in enumerate(ACCOUNTS)}
    rows = sorted(table.items(), key=lambda kv: (order.get(kv[0], len(order)), kv[0]))
    return HEADER + f"reports: {reports}\naccounts:\n" + "".join(f"  {name}: [{', '.join(str(n) for n in seen)}]\n" for name, seen in rows)


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("reports", nargs="+", help="text dumps of finished decks")
    ap.add_argument("--account", action="append", default=[], help="an account name the built-in list lacks (repeatable)")
    ap.add_argument("--write", action="store_true", help="write fdd_utils/report_lengths.yml instead of printing it")
    args = ap.parse_args()
    text = render(lengths(args.reports, args.account), len(args.reports))
    if args.write:
        (ROOT / "fdd_utils" / "report_lengths.yml").write_text(text, encoding="utf-8")
        print(f"wrote fdd_utils/report_lengths.yml from {len(args.reports)} report(s)")
    else:
        print(text, end="")
    return 0


if __name__ == "__main__":
    sys.exit(main())

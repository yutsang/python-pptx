#!/usr/bin/env python3
"""Target schema for lease-contract extraction into 合同汇总模板.xlsx.

Captured from a real template dump (Sheet1, header row 3, columns A-AF).
Contains ONLY column letters + header labels — no filled example rows,
party names, or other client-specific values. Gold-row validation (if any)
must read those values at runtime from the local template file, never from
constants in this repo.

Layout conventions observed on the real template:
  - Row 2: title (e.g. 租赁台账 - 截至...)
  - Row 3: column headers (B-AF; column A is the source filename, no header)
  - Row 4+: project/entity divider rows (name in column B, other cols mostly empty)
    alternating with one data row per source PDF (main contract and each
    amendment/补充协议 get their own row; column A = source filename)
"""
from __future__ import annotations

from typing import Dict, List, Tuple

# (Excel column letter, header text as shown in the template).
# Column A has no header cell in row 3; it holds the source PDF filename.
TEMPLATE_COLUMNS: List[Tuple[str, str]] = [
    ("A", "源文件名"),  # not printed in header row 3; convention from data rows
    ("B", "甲方"),
    ("C", "乙方"),
    ("D", "行业"),
    ("E", "状态（存续/意向/终止）"),
    ("F", "租赁单元"),
    ("G", "租赁面积（平方米）"),
    ("H", "交付日"),
    ("I", "租赁开始 日"),  # space preserved as in the real template header cell
    ("J", "租赁结束日"),
    ("K", "提前终止日"),
    ("L", "免租开始日"),
    ("M", "免租结束日"),
    ("N", "免租期（月）"),
    ("O", "租金合同总额（含税）"),
    ("P", "物业费合同总额（含税）"),
    ("Q", "起始租金/日/平方米（不含税）"),
    ("R", "起始物业管理费/日/平方米（不含税）"),
    ("S", "租金涨幅"),
    ("T", "物业费涨幅"),
    ("U", "租赁保证金 (CNY)"),
    ("V", "支付方式"),
    ("W", "提前终止"),
    ("X", "收款账户"),
    ("Y", "收款账号"),
    ("Z", "违约条款"),
    ("AA", "续租条款"),
    ("AB", "特殊条款"),
    ("AC", "备注"),
    ("AD", "含税日租金"),
    ("AE", "含税日物业费"),
    ("AF", "合计"),
]

# Fields that are typically long free-text clauses copied from the contract
# (costly / noisy for vision; may need dedicated page selection).
LONG_TEXT_COLUMNS = ("V", "W", "Z", "AA", "AB")

# Numeric / short fields that are the highest-value extraction targets for
# the first pass (validate against gold rows in the local template first).
CORE_VALUE_COLUMNS = (
    "B", "C", "F", "G", "H", "I", "J", "K", "L", "M", "N",
    "O", "P", "Q", "R", "S", "T", "U", "X", "Y", "AD", "AE", "AF",
)


def column_map() -> Dict[str, str]:
    """Excel letter -> header label."""
    return {letter: header for letter, header in TEMPLATE_COLUMNS}


def header_row_labels() -> List[str]:
    """Labels for columns B-AF only (matches template row 3 cells)."""
    return [header for letter, header in TEMPLATE_COLUMNS if letter != "A"]


if __name__ == "__main__":
    print(f"{len(TEMPLATE_COLUMNS)} columns (A-AF):")
    for letter, header in TEMPLATE_COLUMNS:
        print(f"  {letter:>2}  {header}")

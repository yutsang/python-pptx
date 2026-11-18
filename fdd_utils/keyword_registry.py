from __future__ import annotations

from typing import Optional

SUMMARY_ACCOUNT_SKIP_KEYWORDS = (
    "subtotal",
    "sub-total",
    "sub total",
    "gross profit",
    "gross loss",
    "operating profit",
    "operating loss",
    "profit/(loss) before taxation",
    "profit before taxation",
    "loss before taxation",
    "net profit",
    "net loss",
    "合计",
    "总计",
    "小计",
    "毛利",
    "营业利润",
    "利润总额",
    "净利润",
    "亏损",
)

STATEMENT_ORDER_SKIP_KEYWORDS = (
    "total",
    "subtotal",
    "sub-total",
    "sub total",
    "合计",
    "总计",
    "小计",
)


# Markers indicating amounts are stated in thousands (apply ×1000 to recover the
# absolute figure). Must cover Simplified AND Traditional Chinese, else traditional
# databooks skip the multiplier and every reconciliation is off by 1000×. Every
# real client databook seen so far declares a CNY/RMB marker even when individual
# sub-accounts are USD/HKD-denominated (the client pre-converts to RMB-equivalent
# before it reaches the sheet), but USD'000/HKD'000 are included here in case a
# future databook is denominated in a different base currency at the sheet level —
# without this, such a sheet would silently fall back to a 1x multiplier and every
# figure on it would be 1000x too small.
UNIT_THOUSAND_MARKERS = ["cny'000", "rmb'000", "usd'000", "hkd'000", "人民币千元", "人民幣千元"]


def contains_thousand_unit_marker(text: str) -> bool:
    """True if the header text declares amounts in thousands (EN or zh-Hans/zh-Hant)."""
    blob = str(text or "").lower()
    return any(marker in blob for marker in UNIT_THOUSAND_MARKERS)


INDICATIVE_KEYWORDS = ["Indicative adjusted", "示意性调整后", "示意性調整後", "CNY'000", "人民币千元"]
BS_HEADER_KEYWORDS = ["示意性调整后资产负债表", "示意性調整後資產負債表", "Indicative adjusted balance sheet", "Indicative Adjusted Balance Sheet", "Balance sheet"]
IS_HEADER_KEYWORDS = ["示意性调整后利润表", "Indicative adjusted income statement", "Indicative Adjusted Income Statement"]
BS_END_KEYWORDS = ["负债及所有者权益总计", "Total liabilities and owners", "Total liabilities and owner"]
IS_END_KEYWORDS = ["净利润", "Net profit", "Net Profit"]
SUBTOTAL_KEYWORDS = ["小计", "合计", "总计", "subtotal", "sub-total", "total", "小計", "合計", "總計"]

REMARK_KEYWORDS = (
    "remark",
    "remarks",
    "note",
    "notes",
    "comment",
    "comments",
    "备注",
    "附注",
    "注释",
    "說明",
    "说明",
)

TABLE_END_KEYWORDS = {
    "Balance Sheet": [
        "负债及所有者权益总计",
        "Total liabilities and owners",
        "Total liabilities and owner",
    ],
    "Income Statement": [
        "净利润",
        "淨利潤",
        "Net profit",
        "Net Profit",
        "net income",
    ],
}

CATEGORY_TRANSLATIONS_ZH = {
    "Current assets": "流动资产",
    "Current Assets": "流动资产",
    "Non-current assets": "非流动资产",
    "Non-Current Assets": "非流动资产",
    "Non current assets": "非流动资产",
    "Assets": "资产",
    "Current liabilities": "流动负债",
    "Current Liabilities": "流动负债",
    "Non-current liabilities": "非流动负债",
    "Non-Current Liabilities": "非流动负债",
    "Non current liabilities": "非流动负债",
    "Liabilities": "负债",
    "Expenses": "费用",
    "Equity": "所有者权益",
    "Owner's equity": "所有者权益",
    "Owners' equity": "所有者权益",
    "Shareholders equity": "股东权益",
    "Shareholders' equity": "股东权益",
    "Revenue": "营业收入",
    "Sales": "销售收入",
    "Income": "收入",
    "Operating revenue": "营业收入",
    "Operating Revenue": "营业收入",
    "Cost of sales": "营业成本",
    "Cost of Sales": "营业成本",
    "Cost of goods sold": "销售成本",
    "COGS": "销售成本",
    "Operating expenses": "营业费用",
    "Operating Expenses": "营业费用",
    "Selling expenses": "销售费用",
    "Administrative expenses": "管理费用",
    "General and administrative": "管理费用",
    "G&A": "管理费用",
    "Other income": "其他收入",
    "Other Income": "其他收入",
    "Other expenses": "其他费用",
    "Other Expenses": "其他费用",
    "Finance costs": "财务费用",
    "Finance Costs": "财务费用",
    "Financial expenses": "财务费用",
    "Interest expense": "利息费用",
    "Tax": "税费",
    "Income tax": "所得税",
    "Taxes": "税费",
    "Tax expense": "所得税费用",
    "Gross profit": "毛利",
    "Operating profit": "营业利润",
    "Net profit": "净利润",
    "Profit before tax": "利润总额",
}

_CATEGORY_TRANSLATIONS_ZH_CASEFOLD = {
    key.lower(): value for key, value in CATEGORY_TRANSLATIONS_ZH.items()
}

# BS/IS statement-structure total/subtotal LINE labels -- distinct from
# CATEGORY_TRANSLATIONS_ZH above (bare category names like "Current assets")
# since these always carry a "Total ..." prefix and are pulled straight from
# extract_balance_sheet_and_income_statement's own row labels (raw Financials
# sheet text), not resolved through mappings.yml the way individual account
# rows are -- embed_financial_tables' embedded BS/IS table has no other
# translation path for these rows at all.
STATEMENT_TOTAL_LINE_TRANSLATIONS_ZH = {
    "Total current assets": "流动资产合计",
    "Total non-current assets": "非流动资产合计",
    "Total non current assets": "非流动资产合计",
    "Total assets": "资产总计",
    "Total current liabilities": "流动负债合计",
    "Total non-current liabilities": "非流动负债合计",
    "Total non current liabilities": "非流动负债合计",
    "Total liabilities": "负债合计",
    "Total owners' equity": "所有者权益合计",
    "Total owner's equity": "所有者权益合计",
    "Total equity": "所有者权益合计",
    "Total equity attributable to owners of the Company": "归属于母公司所有者权益合计",
    "Total liabilities and owners' equity": "负债及所有者权益总计",
    "Total liabilities and owner's equity": "负债及所有者权益总计",
    "Gross profit/(loss)": "毛利/(毛损)",
    "Operating profit/(loss)": "营业利润/(亏损)",
    "Profit before taxation": "利润总额",
    "Profit/(loss) before taxation": "利润总额",
    "Loss before taxation": "利润总额",
    "Net profit/(loss)": "净利润/(净亏损)",
    "Net loss": "净亏损",
}

_STATEMENT_TOTAL_LINE_TRANSLATIONS_ZH_CASEFOLD = {
    key.lower(): value for key, value in STATEMENT_TOTAL_LINE_TRANSLATIONS_ZH.items()
}


def translate_statement_line_to_chinese(label: str) -> Optional[str]:
    """Returns the Chinese translation of a BS/IS total/subtotal LINE label
    (e.g. 'Total current assets' -> '流动资产合计'), or None if `label` isn't
    one of the known statement-structure lines -- callers should fall back
    to their own per-account translation (mappings.yml aliases) first, since
    this only covers total/subtotal rows, not individual accounts."""
    if label in STATEMENT_TOTAL_LINE_TRANSLATIONS_ZH:
        return STATEMENT_TOTAL_LINE_TRANSLATIONS_ZH[label]
    # Source sheets are inconsistent about straight (') vs curly/smart (’)
    # apostrophes in labels like "Total owners' equity" -- normalize both to
    # straight before the casefold lookup so the dict only needs one spelling.
    normalized = str(label).strip().lower().replace("’", "'")
    return _STATEMENT_TOTAL_LINE_TRANSLATIONS_ZH_CASEFOLD.get(normalized)


def translate_category_to_chinese(category: str) -> str:
    if category in CATEGORY_TRANSLATIONS_ZH:
        return CATEGORY_TRANSLATIONS_ZH[category]
    return _CATEGORY_TRANSLATIONS_ZH_CASEFOLD.get(str(category).lower(), category)


__all__ = [
    "BS_END_KEYWORDS",
    "BS_HEADER_KEYWORDS",
    "CATEGORY_TRANSLATIONS_ZH",
    "INDICATIVE_KEYWORDS",
    "IS_END_KEYWORDS",
    "IS_HEADER_KEYWORDS",
    "REMARK_KEYWORDS",
    "STATEMENT_ORDER_SKIP_KEYWORDS",
    "STATEMENT_TOTAL_LINE_TRANSLATIONS_ZH",
    "SUBTOTAL_KEYWORDS",
    "SUMMARY_ACCOUNT_SKIP_KEYWORDS",
    "TABLE_END_KEYWORDS",
    "UNIT_THOUSAND_MARKERS",
    "contains_thousand_unit_marker",
    "translate_category_to_chinese",
    "translate_statement_line_to_chinese",
]

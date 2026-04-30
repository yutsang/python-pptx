"""
Two improvements based on pattern analysis:
  1. Upgrade IP / OP / Capital / Long-term loans / NCL due within one year
     with item-specific subagent_1_prompts drawn from observed patterns.
  2. Append the 8 missing items:
     CIP, DTA, DTL, Deferred income, Interest payable,
     Long-term payables, OCI, Other NCL.
"""
from pathlib import Path
import re

YAML = Path(__file__).resolve().parents[1] / "fdd_utils" / "mappings.yml"

# --------------------------------------------------------------------------
# 1.  Upgraded blocks for existing items
# --------------------------------------------------------------------------

UPGRADE_IP = """\
  subagent_1_prompts:
    Eng:
      system_prompt: |
        Role: Finance Due Diligence Consultant
        Write investment property commentary. Use only the provided data.
        Follow this structure when the data supports it:
        (a) Opening: total net book value at the reporting date, and its split (land use right vs. warehouses/buildings).
        (b) Land use right: location, area, permitted use, remaining term, original cost components (land premium + deed tax as % of premium). Depreciation method and useful life.
        (c) Warehouses / buildings: original cost, depreciation method, useful life, residual value rate.
        (d) Collateral: whether IPs are mortgaged/pledged and for which lender's loans.
        (e) Final accounting: whether completed, any differences vs construction costs on-book.
        (f) CAPEX note (if data supports): recommend the reader confirm future CAPEX plans with management and factor any capital commitments into the valuation model.
        Omit any sub-point for which no data is provided. Write as flowing sentences or short numbered lists. 3-5 sentences for concise, up to 8 for a complex multi-building portfolio. {period_reference_guidance}
        **LANGUAGE**: THIS IS AN ENGLISH DATABOOK. Output must be 100% in ENGLISH.
      user_prompt: |
        Use the patterns as style guidance (not templates), then write from the actual data.
        Patterns: {patterns}
        Data: {financial_figure}
        {rhs_guidance_block}
        {period_reference_guidance}
        {remarks_weight_instruction}
        {user_guidance_instruction}
        Structure: net book value → land use right details → building/warehouse details → collateral → final accounting → CAPEX flag. Follow the data; omit sub-points that are not supported.
        **LANGUAGE**: Output must be 100% in ENGLISH.
        Output only final text without pattern names.
    Chi:
      system_prompt: |
        角色: 财务尽职调查顾问
        撰写投资性房地产评论。仅使用提供的数据。
        在数据支持时，遵循以下结构：
        (a) 开篇：报告日账面价值总额及其拆分（土地使用权 vs. 仓库/建筑物）。
        (b) 土地使用权：所在地、面积、用途、剩余年限、原值构成（地价款+契税占地价款比例）。折旧方法与年限。
        (c) 仓库/建筑物：原值、折旧方法、折旧年限、净残值率。
        (d) 抵押情况：投资性房地产是否已抵押及对应贷款行。
        (e) 决算情况：是否已完成决算，决算结果与建造成本的差异是否已入账。
        (f) 资本性支出提示（如数据支持）：建议读者与管理层确认未来资本性支出计划，并在估值模型中考虑相关资本承诺。
        不提供数据支持的子项目可省略。整体约3-5句，复杂项目可延长至8句。{period_reference_guidance}
        **语言**: 这是一个中文数据簿。输出必须100%为中文。
      user_prompt: |
        以下模式仅作写作风格参考，请根据实际数据撰写最终评论。
        模式: {patterns}
        数据: {financial_figure}
        {rhs_guidance_block}
        {period_reference_guidance}
        {remarks_weight_instruction}
        {user_guidance_instruction}
        结构：账面价值 → 土地使用权详情 → 仓库/建筑物详情 → 抵押情况 → 决算情况 → 资本支出提示。严格依据数据，缺乏数据支持的子项目可省略。
        **语言**: 输出必须100%为中文。
        仅输出最终文本，不包含模式名称。
"""

UPGRADE_OP = """\
  subagent_1_prompts:
    Eng:
      system_prompt: |
        Role: Finance Due Diligence Consultant
        Write other payables commentary. Use only the provided data.
        Opening sentence: state the total balance and reporting date.
        Then list the main components as a numbered or bulleted set using the style shown in patterns — each component should have an amount and a brief description. Typical components: related-party borrowings (note interest-free and pre-deal settlement), construction payables, interest payables, property management fee payables, accrued expenses. Only include components that appear in the data.
        Keep under 5-6 components; group minor items if needed. Do not mention unsupported settlement dates or management representations beyond what the data shows. Write about 3-5 sentences total. {period_reference_guidance}
        **LANGUAGE**: THIS IS AN ENGLISH DATABOOK. Output must be 100% in ENGLISH.
      user_prompt: |
        Use the patterns as style guidance (not templates), then write from the actual data.
        Patterns: {patterns}
        Data: {financial_figure}
        {rhs_guidance_block}
        {period_reference_guidance}
        {remarks_weight_instruction}
        {user_guidance_instruction}
        Open with total balance, then list the main sub-components (amount + description). Related-party borrowings are the most important component — always note whether interest-free and whether settlement prior to transaction is mentioned in the data. Keep the output tight: 3-5 sentences.
        **LANGUAGE**: Output must be 100% in ENGLISH.
        Output only final text without pattern names.
    Chi:
      system_prompt: |
        角色: 财务尽职调查顾问
        撰写其他应付款评论。仅使用提供的数据。
        开篇陈述报告日余额总额，随后列示主要构成项（金额+简述），典型项目包括：关联方往来款（说明是否免息及是否在交割前结清）、工程款、银行利息、物业管理费、预提费用。仅列示数据中出现的项目，次要项目可合并。整体约3-5句。{period_reference_guidance}
        **语言**: 这是一个中文数据簿。输出必须100%为中文。
      user_prompt: |
        以下模式仅作写作风格参考，请根据实际数据撰写最终评论。
        模式: {patterns}
        数据: {financial_figure}
        {rhs_guidance_block}
        {period_reference_guidance}
        {remarks_weight_instruction}
        {user_guidance_instruction}
        以余额总额开篇，然后分列主要构成（金额+简述）。关联方往来款是最重要的构成项，须注明是否免息、是否在完成交割前结清（如数据支持）。整体约3-5句。
        **语言**: 输出必须100%为中文。
        仅输出最终文本，不包含模式名称。
"""

UPGRADE_CAPITAL = """\
  subagent_1_prompts:
    Eng:
      system_prompt: |
        Role: Finance Due Diligence Consultant
        Write paid-in capital / registered capital commentary. Use only the provided data.
        Cover in this order when data supports:
        (a) Consolidated-level paid-in capital balance (total, and name of the top entity if shown).
        (b) Registered capital of the main WFOE / onshore entity: amount, currency, whether fully paid-up and verified.
        (c) Verification source: reference the capital injection documentation or audit report checked, and state whether differences were found.
        (d) If relevant: consolidation elimination treatment for WFOE capital.
        Keep to 2-4 sentences. Do not add unsupported commentary about shareholder structure or ownership. {period_reference_guidance}
        **LANGUAGE**: THIS IS AN ENGLISH DATABOOK. Output must be 100% in ENGLISH.
      user_prompt: |
        Use the patterns as style guidance (not templates), then write from the actual data.
        Patterns: {patterns}
        Data: {financial_figure}
        {rhs_guidance_block}
        {period_reference_guidance}
        {remarks_weight_instruction}
        {user_guidance_instruction}
        Cover: (1) consolidated paid-in capital, (2) entity-level registered capital with verification, (3) cross-check result. Keep to 2-4 sentences.
        **LANGUAGE**: Output must be 100% in ENGLISH.
        Output only final text without pattern names.
    Chi:
      system_prompt: |
        角色: 财务尽职调查顾问
        撰写实收资本/注册资本评论。仅使用提供的数据。
        按顺序说明（有数据支持时）：
        (a) 合并层面实收资本余额（金额及境外控股主体名称）。
        (b) 主要境内实体注册资本：金额、币种、是否足额缴付及经验资。
        (c) 核查来源：所参考的增资协议或审计报告，以及是否发现差异。
        (d) 如适用：境内实体资本的抵销处理。
        整体约2-4句，不要加入未被数据支持的股权结构评论。{period_reference_guidance}
        **语言**: 这是一个中文数据簿。输出必须100%为中文。
      user_prompt: |
        以下模式仅作写作风格参考，请根据实际数据撰写最终评论。
        模式: {patterns}
        数据: {financial_figure}
        {rhs_guidance_block}
        {period_reference_guidance}
        {remarks_weight_instruction}
        {user_guidance_instruction}
        涵盖：①合并层实收资本，②境内实体注册资本及验资情况，③核查结果。整体约2-4句。
        **语言**: 输出必须100%为中文。
        仅输出最终文本，不包含模式名称。
"""

UPGRADE_LTLOANS = """\
  subagent_1_prompts:
    Eng:
      system_prompt: |
        Role: Finance Due Diligence Consultant
        Write long-term loans commentary. Use only the provided data.
        Cover when data supports: total balance, lender name, credit facility size vs. amount drawn, loan purpose, term (start → maturity), interest rate (e.g. LPR + spread or fixed rate), collateral (land use rights, equity pledge, receivables), any guarantors. Mention weighted average interest rate for the period if available. Keep to 2-5 sentences. {period_reference_guidance}
        **LANGUAGE**: THIS IS AN ENGLISH DATABOOK. Output must be 100% in ENGLISH.
      user_prompt: |
        Use the patterns as style guidance (not templates), then write from the actual data.
        Patterns: {patterns}
        Data: {financial_figure}
        {rhs_guidance_block}
        {period_reference_guidance}
        {remarks_weight_instruction}
        {user_guidance_instruction}
        State: balance, lender, term dates, interest rate, collateral. 2-5 sentences.
        **LANGUAGE**: Output must be 100% in ENGLISH.
        Output only final text without pattern names.
    Chi:
      system_prompt: |
        角色: 财务尽职调查顾问
        撰写长期借款评论。仅使用提供的数据。
        按顺序说明（有数据支持时）：余额、贷款行名称、授信额度与已提款金额、借款用途、借款期限（起止日期）、利率（LPR+基点或固定利率）、抵押/质押物（土地使用权、股权质押、应收账款质押）、担保人。整体约2-5句。{period_reference_guidance}
        **语言**: 这是一个中文数据簿。输出必须100%为中文。
      user_prompt: |
        以下模式仅作写作风格参考，请根据实际数据撰写最终评论。
        模式: {patterns}
        数据: {financial_figure}
        {rhs_guidance_block}
        {period_reference_guidance}
        {remarks_weight_instruction}
        {user_guidance_instruction}
        说明：余额、贷款行、期限日期、利率、抵押物。约2-5句。
        **语言**: 输出必须100%为中文。
        仅输出最终文本，不包含模式名称。
"""

UPGRADE_NCLDUE = """\
  subagent_1_prompts:
    Eng:
      system_prompt: |
        Role: Finance Due Diligence Consultant
        Write NCL-due-within-one-year commentary. Use only the provided data.
        Cover when data supports: amount, onshore/offshore split, lender name, specific maturity dates, interest rate, collateral. If the DBS or other significant loans are fully maturing soon, note this and suggest the reader consider funding sources and repayment/extension plans. Keep to 2-4 sentences. {period_reference_guidance}
        **LANGUAGE**: THIS IS AN ENGLISH DATABOOK. Output must be 100% in ENGLISH.
      user_prompt: |
        Use the patterns as style guidance (not templates), then write from the actual data.
        Patterns: {patterns}
        Data: {financial_figure}
        {rhs_guidance_block}
        {period_reference_guidance}
        {remarks_weight_instruction}
        {user_guidance_instruction}
        State: amounts, lender(s), maturity, interest rate, collateral. If large loans are maturing, flag the repayment planning point. 2-4 sentences.
        **LANGUAGE**: Output must be 100% in ENGLISH.
        Output only final text without pattern names.
    Chi:
      system_prompt: |
        角色: 财务尽职调查顾问
        撰写一年内到期的非流动负债评论。仅使用提供的数据。
        说明：金额（境内/境外分别列示）、贷款行名称、到期日期、利率、抵押物。若有大额贷款即将到期，建议读者关注还款/展期安排。整体约2-4句。{period_reference_guidance}
        **语言**: 这是一个中文数据簿。输出必须100%为中文。
      user_prompt: |
        以下模式仅作写作风格参考，请根据实际数据撰写最终评论。
        模式: {patterns}
        数据: {financial_figure}
        {rhs_guidance_block}
        {period_reference_guidance}
        {remarks_weight_instruction}
        {user_guidance_instruction}
        说明：金额、贷款行、到期日、利率、抵押物，如大额贷款即将到期需提示还款安排。约2-4句。
        **语言**: 输出必须100%为中文。
        仅输出最终文本，不包含模式名称。
"""


# --------------------------------------------------------------------------
# 2.  New items to append
# --------------------------------------------------------------------------

NEW_ITEMS = """\

CIP:
  type: BS
  category: Non-current assets
  aliases: ["CIP", "Construction in Progress", "Construction-in-progress", "在建工程"]
  subagent_1_prompts:
    Eng:
      system_prompt: |
        Role: Finance Due Diligence Consultant
        Write construction-in-progress commentary. Use only the provided data.
        Cover: total CIP balance, what it represents (project name/phase if shown), accounting policy for cost recognition (construction-progress-based accrual if stated), completion status and expected completion date. Flag if the management has not provided a construction contract log with contract/verified/settlement amounts — in that case note that construction payables and CAPEX commitments may be under-estimated. Keep to 2-4 sentences. {period_reference_guidance}
        **LANGUAGE**: THIS IS AN ENGLISH DATABOOK. Output must be 100% in ENGLISH.
      user_prompt: |
        Use the patterns as style guidance (not templates), then write from the actual data.
        Patterns: {patterns}
        Data: {financial_figure}
        {rhs_guidance_block}
        {period_reference_guidance}
        {remarks_weight_instruction}
        {user_guidance_instruction}
        State: balance, what it represents, completion status/expected date, cost-recognition policy. Flag under-estimation risk if no contract log was provided. 2-4 sentences.
        **LANGUAGE**: Output must be 100% in ENGLISH.
        Output only final text without pattern names.
    Chi:
      system_prompt: |
        角色: 财务尽职调查顾问
        撰写在建工程评论。仅使用提供的数据。
        说明：余额、代表内容（项目名称/阶段）、成本确认政策（如按工程进度确认）、完工进度及预计完工日期。若管理层未提供详细施工合同台账（含合同金额、核定金额、结算金额），需提示建造成本应付款及资本支出承诺可能被低估。整体约2-4句。{period_reference_guidance}
        **语言**: 这是一个中文数据簿。输出必须100%为中文。
      user_prompt: |
        以下模式仅作写作风格参考，请根据实际数据撰写最终评论。
        模式: {patterns}
        数据: {financial_figure}
        {rhs_guidance_block}
        {period_reference_guidance}
        {remarks_weight_instruction}
        {user_guidance_instruction}
        说明：余额、内容、完工状态/预计完工日期、成本确认政策；如未提供合同台账，提示低估风险。约2-4句。
        **语言**: 输出必须100%为中文。
        仅输出最终文本，不包含模式名称。
  patterns:
    Pattern 1: |
      the balance as at <DATE> represented some dynamic compaction work. According to internal accounting policy, construction costs were accounted for based on the construction progress maintained in the system by the engineering department, and the finance department would accrue the construction costs accordingly
    Pattern 2: |
      there was no balance as at <DATE> as construction was completed in <DATE>. The balance as at <DATE> represented the construction cost of Phase I and Phase II of the project

DTA:
  type: BS
  category: Non-current assets
  aliases: ["DTA", "Deferred Tax Assets", "Deferred tax assets", "Deferred tax asset", "递延所得税资产", "遞延所得稅資產"]
  subagent_1_prompts:
    Eng:
      system_prompt: |
        Role: Finance Due Diligence Consultant
        Write deferred tax assets commentary. Use only the provided data. List the main sources of DTA from the data — common sources in this sector: (1) revenue recognition timing difference from the rent-free period straight-line adjustment, (2) un-utilised tax losses carried forward, (3) accrued-but-not-yet-paid liabilities. State the balance at the reporting date and its main drivers. Keep to 2-3 sentences. Do not speculate on recoverability unless the data supports it. {period_reference_guidance}
        **LANGUAGE**: THIS IS AN ENGLISH DATABOOK. Output must be 100% in ENGLISH.
      user_prompt: |
        Use the patterns as style guidance (not templates), then write from the actual data.
        Patterns: {patterns}
        Data: {financial_figure}
        {rhs_guidance_block}
        {period_reference_guidance}
        {remarks_weight_instruction}
        {user_guidance_instruction}
        State: balance at reporting date, main sources. 2-3 sentences.
        **LANGUAGE**: Output must be 100% in ENGLISH.
        Output only final text without pattern names.
    Chi:
      system_prompt: |
        角色: 财务尽职调查顾问
        撰写递延所得税资产评论。仅使用提供的数据。列示DTA主要来源：①直线法收入确认时间性差异，②未使用税收亏损，③已计提未支付负债。陈述报告日余额及主要驱动因素。整体约2-3句，不对可回收性作出未被数据支持的判断。{period_reference_guidance}
        **语言**: 这是一个中文数据簿。输出必须100%为中文。
      user_prompt: |
        以下模式仅作写作风格参考，请根据实际数据撰写最终评论。
        模式: {patterns}
        数据: {financial_figure}
        {rhs_guidance_block}
        {period_reference_guidance}
        {remarks_weight_instruction}
        {user_guidance_instruction}
        陈述报告日余额及主要来源。约2-3句。
        **语言**: 输出必须100%为中文。
        仅输出最终文本，不包含模式名称。
  patterns:
    Pattern 1: |
      the balance as at <DATE> mainly arose from the impact of revenue recognition in the rent-free period, un-utilised tax losses, and liabilities that were accrued but not paid
    Pattern 2: |
      the balance as at <DATE> mainly arose from the impact of revenue recognition in the rent-free period and accrued but not-paid liabilities

DTL:
  type: BS
  category: Non-current liabilities
  aliases: ["DTL", "Deferred Tax Liabilities", "Deferred tax liabilities", "Deferred tax liability", "递延所得税负债", "遞延所得稅負債"]
  system_prompts:
  user_prompts:
  patterns:
    Pattern 1: |
      mainly generated from the impact of revenue recognition in the rent-free period
    Pattern 2: |
      系按直线法确认租金及物管费收入形成的递延所得税负债

Deferred income:
  type: BS
  category: Non-current liabilities
  aliases: ["Deferred income", "Deferred Income", "Government grants", "Government grant", "递延收益", "政府补贴"]
  subagent_1_prompts:
    Eng:
      system_prompt: |
        Role: Finance Due Diligence Consultant
        Write deferred income / government grants commentary. Use only the provided data.
        Cover: total balance, nature of grant (e.g. government subsidy for industrial development), year received, original amount, amortisation period (typically useful life of the related asset, e.g. 50 years or 20 years), annual amortisation amount. State when the grant was received and the amortisation method. Keep to 2-3 sentences. {period_reference_guidance}
        **LANGUAGE**: THIS IS AN ENGLISH DATABOOK. Output must be 100% in ENGLISH.
      user_prompt: |
        Use the patterns as style guidance (not templates), then write from the actual data.
        Patterns: {patterns}
        Data: {financial_figure}
        {rhs_guidance_block}
        {period_reference_guidance}
        {remarks_weight_instruction}
        {user_guidance_instruction}
        State: balance, grant type, year received, amortisation period and annual amount. 2-3 sentences.
        **LANGUAGE**: Output must be 100% in ENGLISH.
        Output only final text without pattern names.
    Chi:
      system_prompt: |
        角色: 财务尽职调查顾问
        撰写递延收益/政府补贴评论。仅使用提供的数据。
        说明：余额、补贴性质（如工业发展支持资金）、收到年份、原始金额、摊销期间（通常为相关资产使用寿命，如50年或20年）、每年摊销金额。整体约2-3句。{period_reference_guidance}
        **语言**: 这是一个中文数据簿。输出必须100%为中文。
      user_prompt: |
        以下模式仅作写作风格参考，请根据实际数据撰写最终评论。
        模式: {patterns}
        数据: {financial_figure}
        {rhs_guidance_block}
        {period_reference_guidance}
        {remarks_weight_instruction}
        {user_guidance_instruction}
        说明：余额、补贴类型、收到年份、摊销期间及年摊销金额。约2-3句。
        **语言**: 输出必须100%为中文。
        仅输出最终文本，不包含模式名称。
  patterns:
    Pattern 1: |
      represented property-related government grants, which were initially recorded as long-term payables and then recognised in other income on a straight-line basis over the useful life period of the Targeted properties
    Pattern 2: |
      递延收益为<DATE>收到的政府扶持资金<AMT>，其在相关资产使用期限内（即50年）平均分摊

Interest payable:
  type: BS
  category: Current liabilities
  aliases: ["Interest payable", "Interest Payable", "Accrued interest", "应付利息", "應付利息"]
  system_prompts:
  user_prompts:
  patterns:
    Pattern 1: |
      accrued and unpaid interest on bank loans amounting to <AMT>
    Pattern 2: |
      系银行借款尚未结清的利息费用<AMT>

Long-term payables:
  type: BS
  category: Non-current liabilities
  aliases: ["Long-term payables", "Long term payables", "Long-term payable", "长期应付款", "長期應付款"]
  system_prompts:
  user_prompts:
  patterns:
    Pattern 1: |
      represented property-related government grants, which were initially recorded as long-term payables and then recognised in other income on a straight-line basis over the useful life period of the Targeted properties. Historically, the entity received government subsidies amounting to <AMT>
    Pattern 2: |
      系2013年收到的政府扶持资金<AMT>，其在相关资产使用期限内（即50年）平均分摊

OCI:
  type: BS
  category: Equity
  aliases: ["OCI", "Other Comprehensive Income", "Other comprehensive income", "其他综合收益", "其他綜合收益",
            "Translation reserve", "Currency translation reserve", "Foreign currency translation reserve"]
  system_prompts:
  user_prompts:
  patterns:
    Pattern 1: |
      the balance represented the translation differences arising from the translation of offshore USD financial statements to CNY

Other NCL:
  type: BS
  category: Non-current liabilities
  aliases: ["Other NCL", "Other non-current liabilities", "Other Non-current Liabilities",
            "其他非流动负债", "其他非流動負債"]
  system_prompts:
  user_prompts:
  patterns:
    Pattern 1: |
      represented tenant deposits amounting to <AMT>
    Pattern 2: |
      the balance of <AMT> represented lease deposit of the subsidiaries under the principal tenant
"""


def _replace_block(content: str, key: str, old_marker: str, new_block: str) -> str:
    """Replace everything from old_marker up to (but not including) the next
    top-level key or end-of-file with new_block, within item `key`."""
    # Find the item's top-level section.
    item_start = content.find(f"\n{key}:\n")
    if item_start == -1:
        item_start = content.find(f"{key}:\n")  # first line
        if item_start != 0:
            print(f"  ⚠ Cannot find key '{key}' in file — skipping.")
            return content
    else:
        item_start += 1  # skip the leading newline

    # Find the start of the old marker within this item's section.
    marker_pos = content.find(old_marker, item_start)
    if marker_pos == -1:
        # Marker absent — insert before 'patterns:' block instead.
        pat_pos = content.find("\n  patterns:", item_start)
        if pat_pos == -1:
            print(f"  ⚠ Cannot find insertion point for '{key}' — skipping.")
            return content
        return content[:pat_pos + 1] + new_block + content[pat_pos + 1:]

    # Find end of old block (next top-level key after the marker).
    next_key = re.search(r"\n[A-Za-z]", content[marker_pos:])
    if next_key:
        end = marker_pos + next_key.start() + 1
    else:
        end = len(content)

    # Trim trailing whitespace / blank lines in old block.
    old_block = content[marker_pos:end]
    return content[:marker_pos] + new_block + content[end:]


def main():
    text = YAML.read_text(encoding="utf-8")

    upgrades = [
        ("IP",                  "  system_prompts:\n  user_prompts:",  UPGRADE_IP),
        ("OP",                  "  system_prompts:\n  user_prompts:",  UPGRADE_OP),
        ("Capital",             "  system_prompts:\n  user_prompts:",  UPGRADE_CAPITAL),
        ("Long-term loans",     "  system_prompts:\n  user_prompts:",  UPGRADE_LTLOANS),
        ("NCL due within one year", "  system_prompts:\n  user_prompts:", UPGRADE_NCLDUE),
    ]

    print("Upgrading existing items with subagent_1_prompts …")
    for key, marker, block in upgrades:
        text = _replace_block(text, key, marker, block)
        print(f"  ✓ {key}")

    print("\nAppending 8 new items …")
    text = text.rstrip() + "\n" + NEW_ITEMS
    for item in ["CIP", "DTA", "DTL", "Deferred income", "Interest payable",
                 "Long-term payables", "OCI", "Other NCL"]:
        print(f"  ✓ {item}")

    YAML.write_text(text, encoding="utf-8")
    print("\nDone.")


if __name__ == "__main__":
    main()

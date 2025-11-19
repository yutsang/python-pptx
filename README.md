# Financial Data Processing with AI

Automated financial content generation using 4-agent AI pipeline with smart number formatting.

---

## Quick Start

```python
from fdd_utils.process_databook import extract_data_from_excel
from fdd_utils.content_generation import run_ai_pipeline, extract_final_contents

# 1. Extract data from Excel
dfs, keys, _, lang = extract_data_from_excel(
    databook_path='databook.xlsx',
    entity_name='Company Name',  # or "" for single entity
    mode='All'  # "All", "BS", or "IS"
)

# 2. Check extraction succeeded
if not dfs or len(dfs) == 0:
    print("❌ Extraction failed! Run: python test_extraction.py")
    exit()

# 3. Run AI pipeline
results = run_ai_pipeline(
    mapping_keys=keys,
    dfs=dfs,
    model_type='local',  # or 'deepseek', 'openai'
    language=lang,
    use_multithreading=True
)

# 4. Get final contents
final_contents = extract_final_contents(results)
```

**Output**: `fdd_utils/logs/run_TIMESTAMP/results.yml`

---

## The 4 Agents

| Agent | Role | Temperature | What It Does |
|-------|------|-------------|--------------|
| **1_Generator** | Creates content | 0.7 | Generates financial descriptions using pre-formatted values |
| **2_Auditor** | Verifies accuracy | 0.3 | Validates that formatted values match source data |
| **3_Refiner** | Polishes content | 0.5 | Refines without over-shortening (preserves context) |
| **4_Validator** | Final check | 0.2 | Final quality control and format validation |

---

## Number Formatting

Values are **automatically formatted** in code before being sent to AI:

### Chinese (Chi)
| Value Range | Format | Example |
|-------------|--------|---------|
| < 10,000 | Raw | 5000 |
| 10,000 - 99,999,999 | 万 (1 d.p.) | 7.8万 |
| ≥ 100,000,000 | 亿 (2 d.p.) | 1.23亿 |

### English (Eng)
| Value Range | Format | Example |
|-------------|--------|---------|
| < 10,000 | Comma | 5,000 |
| 10,000 - 999,999 | K (1 d.p.) | 78.2K |
| ≥ 1,000,000 | million (2 d.p.) | 12.35 million |

### Special: Negative Retained Earnings
- **未分配利润** (negative) → **未弥补亏损** (positive display)
- **Retained Earnings** (negative) → **Accumulated Losses** (positive display)

---

## Financial Data Extraction

Extract Balance Sheet and Income Statement directly from specific worksheets:

```python
from fdd_utils.financial_extraction import (
    extract_balance_sheet_and_income_statement,
    filter_by_total_amount,
    get_account_total
)

# Extract BS and IS from specific sheets
results = extract_balance_sheet_and_income_statement(
    workbook_path="databook.xlsx",
    balance_sheet_name="示意性调整后资产负债表",      # Balance Sheet sheet name
    income_statement_name="示意性调整后利润表",       # Income Statement sheet name
    entity_keywords=["联洋"],                        # Optional: filter by entity
    debug=True                                       # Enable debug prints (shows what's happening)
)

# Access results
balance_sheet = results['balance_sheet']      # DataFrame or None
income_statement = results['income_statement']  # DataFrame or None

# Example: Filter to show only totals (remove sub-accounts)
if balance_sheet is not None:
    totals_only = filter_by_total_amount(balance_sheet)
    print(totals_only)

# Example: Get specific account value
if balance_sheet is not None:
    cash_total = get_account_total(balance_sheet, "货币资金")
    print(f"Cash: {cash_total:,.0f}")
```

**Returns**: Dictionary with `'balance_sheet'` and `'income_statement'` keys, each containing a DataFrame with columns:
- `Description`: Account name
- `Value`: Numeric value (already multiplied by 1000 if CNY'000)


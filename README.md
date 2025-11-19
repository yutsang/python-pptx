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

Extract Balance Sheet and Income Statement from a **single sheet** containing both statements:

```python
from fdd_utils.financial_extraction import (
    extract_balance_sheet_and_income_statement,
    filter_by_total_amount,
    get_account_total
)

# Extract BS and IS from single sheet
# Both statements are in the same sheet, separated by headers:
# - "示意性调整后资产负债表" or "Indicative adjusted balance sheet"
# - "示意性调整后利润表" or "Indicative adjusted income statement"
results = extract_balance_sheet_and_income_statement(
    workbook_path="databook.xlsx",
    sheet_name="Financial Statements",  # Sheet containing both BS and IS
    debug=True                           # Enable debug prints
)

# Access results
balance_sheet = results['balance_sheet']      # DataFrame or None
income_statement = results['income_statement']  # DataFrame or None
project_name = results['project_name']        # Extracted from headers (e.g., "东莞xx")

# Example: Filter to show only totals (remove sub-accounts)
if balance_sheet is not None:
    totals_only = filter_by_total_amount(balance_sheet)
    print(totals_only)

# Example: Get specific account value
if balance_sheet is not None:
    # Get most recent date value (auto-selects first date column)
    cash_total = get_account_total(balance_sheet, "货币资金")
    print(f"Cash (latest): {cash_total:,.0f}")
    
    # Get specific date value
    cash_2023 = get_account_total(balance_sheet, "货币资金", date_column='2023-12-31')
    print(f"Cash (2023): {cash_2023:,.0f}")
```

**Features**:
- Extracts both BS and IS from **single sheet**
- Auto-detects statement boundaries via headers
- Extracts project name (e.g., from "xxxx利润表 - 东莞xx")
- Gets **ALL columns** with "示意性调整后" or "Indicative adjusted"
- **Smart end detection**:
  - BS ends at "负债及所有者权益总计" or "Total liabilities and owners'equity"
  - IS ends at "净利润/（亏损）" or "Net profit/(loss)"
- Auto-multiplies by 1000 if "CNY'000" or "人民币千元" detected
- Converts dates: FY22→2022-12-31, 9M22→2022-09-30, 30-Sep-2022→2022-09-30
- **Removes empty rows** where all values are 0

**Returns**: Dictionary with keys:
- `'balance_sheet'`: DataFrame with `Description` column + ALL date columns (e.g., `2022-12-31`, `2021-12-31`)
- `'income_statement'`: DataFrame with `Description` column + ALL date columns
- `'project_name'`: String (project/entity name extracted from headers)

**Example Result**:
```
Balance Sheet columns: ['Description', '2024-12-31', '2023-12-31', '2022-12-31']
```


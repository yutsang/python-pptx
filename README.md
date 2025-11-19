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
from fdd_utils.financial_extraction import extract_balance_sheet_and_income_statement

# Extract BS and IS from single sheet
# Both statements are in the same sheet, separated by headers:
# - "示意性调整后资产负债表" or "Indicative adjusted balance sheet"  
# - "示意性调整后利润表" or "Indicative adjusted income statement"
results = extract_balance_sheet_and_income_statement(
    workbook_path="databook.xlsx",
    sheet_name="Sheet1",     # Sheet name containing both BS and IS
    debug=True               # Enable comprehensive debug prints
)

# Access results
balance_sheet = results['balance_sheet']      # DataFrame or None
income_statement = results['income_statement']  # DataFrame or None
project_name = results['project_name']        # Extracted from headers (e.g., "东莞联洋")

# Work with the data
if balance_sheet is not None:
    print(f"Balance Sheet: {len(balance_sheet)} rows")
    print(f"Columns: {list(balance_sheet.columns)}")
    print(balance_sheet.head())
    
    # Access specific account
    cash_row = balance_sheet[balance_sheet['Description'].str.contains('货币资金', na=False)]
    if not cash_row.empty:
        print(f"\n货币资金 (Cash):")
        print(cash_row)

if income_statement is not None:
    print(f"\nIncome Statement: {len(income_statement)} rows")
    print(income_statement.head())
```

**Features**:
- Extracts both BS and IS from **single sheet**
- Auto-detects statement boundaries via headers
- Extracts project name (e.g., from "xxxx利润表 - 东莞联洋")
- Gets **ONLY columns** with "示意性调整后" or "Indicative adjusted" (filters out 管理层数, 审定数, etc.)
- **Smart end detection**:
  - BS ends at "负债及所有者权益总计" or "Total liabilities and owners'equity"
  - IS ends at "净利润/（亏损）" or "Net profit/(loss)"
- Auto-multiplies by 1000 if "CNY'000" or "人民币千元" detected
- Converts dates: FY22→2022-12-31, 9M22→2022-09-30, 30-Sep-2022→2022-09-30
- **Smart column cleanup**: Removes date columns that have all zeros in Income Statement from BOTH statements
- **Removes empty rows**: Filters out rows where all values are 0

**Returns**: Dictionary with keys:
- `'balance_sheet'`: DataFrame with `Description` column + ALL date columns (e.g., `2022-12-31`, `2021-12-31`)
- `'income_statement'`: DataFrame with `Description` column + ALL date columns
- `'project_name'`: String (project/entity name extracted from headers)

**Example Result**:
```
Balance Sheet columns: ['Description', '2024-12-31', '2023-12-31', '2022-12-31']
```

---

## Data Reconciliation

Verify data accuracy by comparing two extraction methods:

```python
from fdd_utils.financial_extraction import extract_balance_sheet_and_income_statement
from fdd_utils.process_databook import extract_data_from_excel
from fdd_utils.reconciliation import reconcile_financial_statements, print_reconciliation_report

# Extract from both sources
# Source 1: BS/IS from single sheet
bs_is_results = extract_balance_sheet_and_income_statement(
    workbook_path="databook.xlsx",
    sheet_name="Financials"
)

# Source 2: Account-by-account extraction
dfs, keys, _, lang = extract_data_from_excel(
    databook_path="databook.xlsx",
    entity_name="",
    mode="All"
)

# Reconcile the two sources (uses LATEST date column only)
bs_recon, is_recon = reconcile_financial_statements(
    bs_is_results=bs_is_results,
    dfs=dfs,
    tolerance=1.0,  # Allow ±1 difference for rounding
    debug=True  # Shows which accounts are matched and total row detection
)

# Print report (show only mismatches)
print_reconciliation_report(bs_recon, is_recon, show_only_issues=True)

# Save to Excel
with pd.ExcelWriter('reconciliation.xlsx') as writer:
    bs_recon.to_excel(writer, sheet_name='BS Reconciliation', index=False)
    is_recon.to_excel(writer, sheet_name='IS Reconciliation', index=False)
```

**Reconciliation Output**:
```
Source_Account    Date         Source_Value  DFS_Account  DFS_Value    Match
货币资金          2024-12-31   4,119,178     货币资金     4,119,178    ✅ Match
应收账款          2024-12-31   13,034,797    应收账款     13,034,797   ✅ Match
投资性房地产      2024-12-31   168,526,613   投资性房地产 168,520,000  ❌ Diff: 6,613
```

**Features**:
- Uses **LATEST date column only** from BS/IS (last date column - most recent)
- **Strict account matching using ONLY mappings.yml aliases**:
  - Exact alias match (e.g., "货币资金" in aliases → key "Cash")
  - Cleans suffixes ('：', '(', ')') before matching
  - **NO name-based matching** - only uses defined aliases
- **Auto-skips total/subtotal/profit lines** (marked as "-"):
  - Chinese: xxx合计, xxx总计, xxx小计, 毛利, 营业利润, 净利润
  - English: Total xxx, Subtotal, Sub-total, Gross profit, Operating profit, Net profit
- Finds **total row** in DFS (looks for '合计', '总计', 'Total' keywords ONLY - skips '小计'/subtotal rows)
- **Smart matching logic**:
  - Source = 0 → Shows "-" (skipped)
  - Source ≠ 0 + Total row found → Compare values
  - Source ≠ 0 + Total row NOT found → ⚠️ Not Found
  - Total/profit lines → Shows "-" (skipped)
- **Income Statement expenses**: Negative values auto-converted to positive (category='Expenses' in mappings.yml)
- Shows: ✅ Match, ❌ Diff: X, ✅ Match (both zero), ⚠️ Not Found, ℹ️ Not Mapped

**Important**: Account matching ONLY works if the account name is in `mappings.yml` aliases. Add missing accounts to mappings.yml if needed.

**Example Output**:
```
Source_Account      Date         Source_Value  DFS_Account  DFS_Value    Match
货币资金            2024-05-31   4,119,178     货币资金     4,119,178    ✅ Match
应收账款            2024-05-31   0             -            -            -
流动资产合计        2024-05-31   9,246,577     -            -            -
投资性房地产        2024-05-31   168,526,613   投资性房地产 168,520,000  ❌ Diff: 6,613
管理费用            2024-05-31   1,234,567     管理费用     1,234,567    ✅ Match
净利润/（亏损）     2024-05-31   -85,061,858   -            -            -
```

Note: Total/profit lines (合计, 总计, 净利润, etc.) show "-" for DFS columns as they are not mapped.

**Example with Debug**:
```
    [MATCH] Searching for: '流动资产合计'
    [MATCH]   ⏭️  Skipped (total/profit line) → ℹ️ Not Mapped

    [MATCH] Searching for: '货币资金'
    [MATCH]   ✅ Exact alias match: alias='货币资金', key='Cash'
      Found total row: '货币资金合计'

    [MATCH] Searching for: '管理费用'
    [MATCH]   ✅ Found: key='GA', category='Expenses'
    [CONVERT] Expense: -1234567 → 1234567 (negative to positive)
```


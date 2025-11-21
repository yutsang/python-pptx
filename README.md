# Financial Data Processing with AI

Automated financial content generation using 4-agent AI pipeline with smart number formatting and reconciliation.

---

## Quick Start - Streamlit App

```bash
# Install dependencies
pip install -r fdd_utils/requirements.txt

# Run the app
streamlit run app.py
```

**Features**:
- 📤 Upload Excel databook
- 🤖 Select AI model (local/openai/deepseek)
- 📊 View Balance Sheet & Income Statement with reconciliation
- 🔄 Generate AI content for all accounts
- 📑 Export to PowerPoint (BS + IS combined)

---

## Quick Start - Python Code

```python
from fdd_utils.process_databook import extract_data_from_excel
from fdd_utils.content_generation import run_ai_pipeline, extract_final_contents

# 1. Extract data from Excel
dfs, keys, _, lang = extract_data_from_excel(
    databook_path='databook.xlsx',
    entity_name='Company Name',  # or "" for single entity
    mode='All'  # Always use "All" mode
)

# 2. Check extraction succeeded
if not dfs or len(dfs) == 0:
    print("❌ Extraction failed!")
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

### Special Handling
- **Negative Retained Earnings**: 未分配利润 (negative) → 未弥补亏损 (positive display)
- **Income Statement Expenses**: Displayed as negative, compared as positive for reconciliation

---

## Financial Data Extraction

Extract Balance Sheet and Income Statement from a **single sheet** containing both statements:

```python
from fdd_utils.financial_extraction import extract_balance_sheet_and_income_statement

results = extract_balance_sheet_and_income_statement(
    workbook_path="databook.xlsx",
    sheet_name="Financials",  # Sheet with both BS and IS
    debug=True                # Enable debug output
)

# Access results
balance_sheet = results['balance_sheet']
income_statement = results['income_statement']
project_name = results['project_name']  # e.g., "东莞联洋"
```

**Features**:
- Extracts both BS and IS from **single sheet**
- Auto-detects boundaries via headers ("示意性调整后资产负债表", "示意性调整后利润表")
- Extracts project name from headers
- Gets **ONLY "示意性调整后"** columns (filters out 管理层数, 审定数, etc.)
- Removes date columns with all zeros (based on Income Statement)
- Multiplies by 1000 if "CNY'000" or "人民币千元" detected
- Converts dates: FY22→2022-12-31, 9M22→2022-09-30

---

## Data Reconciliation

Verify data accuracy by comparing two extraction methods:

```python
from fdd_utils.reconciliation import reconcile_financial_statements, print_reconciliation_report

# Reconcile the two sources
bs_recon, is_recon = reconcile_financial_statements(
    bs_is_results=bs_is_results,
    dfs=dfs,
    tolerance=1.0,               # ±1 absolute difference
    materiality_threshold=0.005, # 0.5% materiality
    debug=True
)

# Print report
print_reconciliation_report(bs_recon, is_recon, show_only_issues=True)
```

**Example Output**:
```
Source_Account      Date         Source_Value  DFS_Account  DFS_Value    Diff        Match
货币资金            2024-05-31   4,119,178     货币资金     4,119,178    0           ✅ Match
应收账款            2024-05-31   0             -            -            -           -
流动资产合计        2024-05-31   9,246,577     -            -            -           -
投资性房地产        2024-05-31   168,526,613   投资性房地产 168,520,000  6,613       ✅ Immaterial
管理费用            2024-05-31   -1,234,567    管理费用     1,234,567    0           ✅ Match
```

**Features**:
- Uses **LATEST date** (last column)
- **Strict alias-only matching** via mappings.yml
- Skips total/subtotal/profit lines (shows "-")
- Skips accounts with source value = 0
- **Materiality threshold**: Diff < 0.5% → ✅ Immaterial
- Expenses: Negative display, positive comparison
- Skips '小计'/Subtotal rows in DFS when finding totals

**Match Status**:
- ✅ **Match**: Exact match (within tolerance)
- ✅ **Immaterial**: Diff < 0.5% of source value
- ❌ **Diff**: Material difference
- ⚠️ **Not Found**: Account not in mappings.yml or not extracted
- **-**: Skipped (total/subtotal/profit line or zero value)

---

## Recent Updates (Nov 2025)

### Number Formatting
- 万/K = 1 decimal place
- 亿/million = 2 decimal places
- Negative retained earnings → 未弥补亏损/Accumulated Losses

### Financial Extraction
- Extracts from single sheet with both BS and IS
- Filters for "示意性调整后" columns only
- Smart end detection (负债及所有者权益总计, 净利润)
- Removes empty date columns

### Reconciliation
- Latest date only (last column)
- Alias-only matching (no name guessing)
- Materiality threshold (0.5% default)
- Skips totals/subtotals/profit lines
- Diff column shows absolute difference

---

## Files Structure

```
fdd_utils/
├── process_databook.py       # Excel extraction + formatting
├── financial_extraction.py   # BS/IS from single sheet
├── reconciliation.py          # Data reconciliation
├── content_generation.py     # 4-agent AI pipeline
├── ai_helper.py              # AI model interface
├── mappings.yml              # Account aliases + prompts
├── prompts.yml               # Agent 2/3/4 prompts
├── config.yml                # AI parameters
└── logs/                     # Run outputs

app.py                        # Streamlit app
```

---

**Start**: `streamlit run app.py` 🚀

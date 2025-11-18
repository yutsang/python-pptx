# Financial Data Processing with AI

Automated financial content generation using 4-agent AI pipeline.

## Usage

```python
from fdd_utils.process_databook import extract_data_from_excel
from fdd_utils.content_generation import run_ai_pipeline, extract_final_contents

# Extract data
dfs, keys, _, lang = extract_data_from_excel(
    'inputs/databook.xlsx', 'Company Name', 'BS'
)

# Run pipeline
results = run_ai_pipeline(keys, dfs, 'local', lang)

# Get finals
final_contents = extract_final_contents(results)
```

**Output**: `fdd_utils/logs/run_TIMESTAMP/results.yml`

---

## The 4 Agents

| Agent | Name | Temperature | Role |
|-------|------|-------------|------|
| 1_Generator | **Generator** | 0.7 | Creates content |
| 2_Auditor | **Auditor** | 0.3 | Verifies accuracy |
| 3_Refiner | **Refiner** | 0.5 | Polishes content |
| 4_Validator | **Validator** | 0.2 | Final check |

---

## Features

✅ Chinese units (万元/亿元)  
✅ Filters sub-accounts ("应付利息_借款利息")  
✅ Uses totals not line items  
✅ Converts scientific notation (4.27e7)  
✅ Unified logging per run  
✅ Multi-threading enabled  

---

## Utilities

```python
# Extract Balance Sheet & Income Statement
from fdd_utils.financial_extraction import extract_balance_sheet_and_income_statement

results = extract_balance_sheet_and_income_statement(
    "inputs/databook.xlsx", 
    "示意性调整后资产负债表",
    "示意性调整后利润表"
)

# Format numbers
from fdd_utils.number_formatting import format_number_chinese
format_number_chinese(5000000, 'Chi')  # 人民币500.0万元
```

---

## Configuration

### AI Parameters
Edit `fdd_utils/config.yml`:

```yaml
agents:
  1_Generator:
    temperature: 0.7       # Higher = creative
    max_tokens: 2000
    frequency_penalty: 0.0
```

**Current settings:**
- 1_Generator: temp=0.7 (creative)
- 2_Auditor: temp=0.3 (precise)
- 3_Refiner: temp=0.5 (balanced)  
- 4_Validator: temp=0.2 (very precise)

### API Setup
```yaml
local:
  api_base: "http://localhost:1234"
  api_key: "local"
  chat_model: "your-model"
```

---

## Agent Prompts

**1_Generator**: `fdd_utils/mappings.yml` - Account-specific prompts  
**2_Auditor/3_Refiner/4_Validator**: `fdd_utils/prompts.yml` - Generic prompts

**Agent names**:
- 1_Generator (temp=0.7)
- 2_Auditor (temp=0.3)
- 3_Refiner (temp=0.5)
- 4_Validator (temp=0.2)

---

## FAQ

**Q: Agent 4 = Final always same?**  
A: Normal! Agent 4 validates. If good, outputs unchanged.

**Q: Adjust AI behavior?**  
A: Edit temperature in `config.yml`

**Q: Sub-accounts appearing?**  
A: Check `filter_details=True`

**Q: Check Agent 4 prompt?**  
A: See `fdd_utils/prompts.yml` lines 342-477

---

**Start**: Open `fdd_app.ipynb` 🚀


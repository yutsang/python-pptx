# Financial Data Processing with AI

Automated financial content generation from Excel databooks using multi-agent AI pipeline.

## Quick Start

### 1. Install
```bash
pip install -r fdd_utils/requirements.txt
```

### 2. Configure
Edit `fdd_utils/config.yml`:
```yaml
local:
  api_base: "http://localhost:1234"
  api_key: "local"
  chat_model: "your-model"
```

### 3. Run
Open `fdd_app.ipynb`:

```python
from fdd_utils.process_databook import extract_data_from_excel
from fdd_utils.content_generation import run_ai_pipeline, extract_final_contents

# Extract data
dfs, workbook_list, _, language = extract_data_from_excel(
    'inputs/databook.xlsx', 'Company Name', 'BS'
)

# Run AI pipeline (4 agents)
results = run_ai_pipeline(workbook_list, dfs, 'local', language)

# Get final contents
final_contents = extract_final_contents(results)
```

**Output**: `fdd_utils/logs/run_TIMESTAMP/results.yml`

---

## The 4 AI Agents

1. **Agent 1** - Content Generator: Creates draft from patterns
2. **Agent 2** - Value Checker: Verifies accuracy & totals
3. **Agent 3** - Content Refiner: Polishes (max 3 points, ≥25%)
4. **Agent 4** - Quality Controller: Final validation

---

## Key Features

✅ **Chinese units** - 万元/亿元 (not K/million)  
✅ **Sub-account filtering** - Removes "应付利息_借款利息"  
✅ **Total focus** - Uses totals not line items  
✅ **Scientific notation** - Converts 4.27e7 properly  
✅ **Unified logging** - All outputs in one subfolder  
✅ **Multi-threading** - Fast parallel processing  
✅ **Bilingual** - English & Chinese support  

---

## Important Notes

### Agent 4 = Final (This is Normal)
Agent 4 validates content quality. If Agent 3 output is already good, Agent 4 outputs it unchanged. This is **correct behavior** - it means the content passed validation.

### Chinese Number Formats
- 50,000 → 人民币5.0万元
- 5,000,000 → 人民币500.0万元  
- 500,000,000 → 人民币5.0亿元
- Negative R/E → "未弥补亏损" (not "未分配利润-XXX")

### Sub-Account Filtering
- ✅ Enabled by default
- ❌ Filters: "应付利息_借款利息", indented items, "其中:"
- ✅ Keeps: Main categories only

---

## Utilities

### Extract Balance Sheet & Income Statement
```python
from fdd_utils.financial_extraction import extract_balance_sheet_and_income_statement

results = extract_balance_sheet_and_income_statement(
    "inputs/databook.xlsx",
    "示意性调整后资产负债表",
    "示意性调整后利润表"
)
```

### Number Formatting
```python
from fdd_utils.number_formatting import format_number_chinese

format_number_chinese(5000000, 'Chi')  # 人民币500.0万元
```

---

## Configuration

### AI Parameters (config.yml)
All AI parameters are controlled in `fdd_utils/config.yml`:

```yaml
agents:
  agent_1:  # Content Generator
    temperature: 0.7       # Higher = more creative
    max_tokens: 2000
    top_p: 0.9
  
  agent_2:  # Value Checker
    temperature: 0.3       # Lower = more precise
    max_tokens: 2000
  
  agent_3:  # Content Refiner
    temperature: 0.5       # Balanced
    max_tokens: 2000
    frequency_penalty: 0.2 # Reduce repetition
  
  agent_4:  # Quality Controller
    temperature: 0.2       # Very precise
    max_tokens: 2000
```

**Parameters explained:**
- `temperature` (0.0-2.0): Creativity level
- `max_tokens`: Maximum response length
- `top_p` (0.0-1.0): Nucleus sampling
- `frequency_penalty` (-2.0-2.0): Reduce repetition
- `presence_penalty` (-2.0-2.0): Topic diversity

### Python Usage

```python
# Extract with filtering (default)
dfs, keys, _, lang = extract_data_from_excel(
    path, entity, mode, filter_details=True
)

# Pipeline with multi-threading
results = run_ai_pipeline(
    keys, dfs, 
    model_type='local',      # 'openai', 'local', 'deepseek'
    language='Chi',          # 'Chi' or 'Eng'
    use_multithreading=True,
    max_workers=None         # Use all CPU cores
)
```

---

## Files

```
├── fdd_app.ipynb              # START HERE
├── fdd_utils/
│   ├── ai_helper.py          # AI helper
│   ├── content_generation.py # 4-agent pipeline
│   ├── process_databook.py   # Excel extraction
│   ├── financial_extraction.py # Standalone BS/IS
│   ├── number_formatting.py  # Formatting utils
│   ├── config.yml            # AI settings
│   ├── prompts.yml           # Agent prompts
│   └── logs/run_TIMESTAMP/   # Output folder
└── inputs/                   # Your Excel files
```

---

## Tuning AI Parameters

Edit `fdd_utils/config.yml` to adjust each agent's behavior:

```yaml
agents:
  agent_1:
    temperature: 0.7  # 0.7 = creative, 0.3 = precise
```

**Current settings:**
- **Agent 1** (Generator): 0.7 - More creative for content generation
- **Agent 2** (Checker): 0.3 - Precise for accuracy verification  
- **Agent 3** (Refiner): 0.5 - Balanced for refinement
- **Agent 4** (Controller): 0.2 - Very precise for validation

Lower temperature = more consistent/precise. Higher = more creative/varied.

---

## Troubleshooting

**Q: Sub-accounts still appearing?**  
A: Check `filter_details=True` in `extract_data_from_excel()`

**Q: Scientific notation in reports?**  
A: System handles this automatically via prompts

**Q: Wrong units in Chinese?**  
A: Check `prompts.yml` - should use 万元/亿元

**Q: Agent 4 = Final always same?**  
A: Normal! Agent 4 only changes if needed

**Q: Want different AI behavior?**  
A: Adjust temperature in `config.yml` for each agent

---

**Ready to use!** Open `fdd_app.ipynb` 🚀


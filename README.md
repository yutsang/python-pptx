# Financial Data Processing with AI

Automated financial content generation from Excel databooks using AI agents and designed patterns.

## What It Does

Extracts financial data from Excel → Processes through 4 AI agents → Generates content for reports

## Quick Start

### 1. Install
```bash
pip install -r fdd_utils/requirements.txt
```

### 2. Configure
Edit `fdd_utils/config.yml` with your AI settings:
```yaml
local:
  api_base: "http://localhost:1234"
  api_key: "local"
```

### 3. Run
Open `fdd_app.ipynb` and run the cells:

```python
# 1. Extract data
from fdd_utils.process_databook import extract_data_from_excel

dfs, workbook_list, _, language = extract_data_from_excel(
    'inputs/your_databook.xlsx', 
    'Company Name', 
    'BS'
)

# 2. Run AI pipeline
from fdd_utils.content_generation import run_ai_pipeline, save_results

results = run_ai_pipeline(workbook_list, dfs, 'local', language)
save_results(results)

# 3. Get final contents
from fdd_utils.content_generation import extract_final_contents

final_contents = extract_final_contents(results)
# Ready to feed into your templates!
```

## The 4 AI Agents

1. **Agent 1**: Generates content from patterns + data
2. **Agent 2**: Verifies values and checks ≥25% rule
3. **Agent 3**: Refines content (max 3 points, ≥25% only)
4. **Agent 4**: Format checking (currency, quotes, numbering)

## Output Structure

```python
results = {
    'Cash': {
        'agent_1': 'Draft content...',
        'agent_2': 'Checked content...',
        'agent_3': 'Refined content...',
        'agent_4': 'Final content...',
        'final': 'Final content...'
    }
}
```

## Documentation

- **Full Guide:** `fdd_utils/HOW_TO_RUN.md`
- **Configuration:** `fdd_utils/config.yml`
- **Prompts:** `fdd_utils/prompts.yml` (Eng/Chi)
- **Patterns:** `fdd_utils/mappings.yml`

## Files

```
├── fdd_app.ipynb              # Main notebook - START HERE
├── fdd_utils/
│   ├── ai_helper.py           # AI helper class
│   ├── content_generation.py  # Main pipeline
│   ├── process_databook.py    # Excel extraction
│   ├── config.yml             # Settings
│   ├── prompts.yml            # AI prompts
│   ├── mappings.yml           # Account patterns
│   └── HOW_TO_RUN.md          # Detailed guide
└── inputs/                    # Put your Excel files here
```

## Key Features

✅ Multi-agent AI pipeline (4 agents)  
✅ Multi-threading for speed  
✅ English & Chinese support  
✅ Unified logging (one file per run)  
✅ Pattern-based content generation  
✅ Automatic value verification  
✅ Listing rules (max 3 points, ≥25%)  
✅ Format validation  

---

**Ready to use!** Open `fdd_app.ipynb` and start processing. 🚀


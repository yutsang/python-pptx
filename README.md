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

## Number Formatting (Updated Nov 2025)

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

## The 4 Agents

| Agent | Role | Temperature | What It Does |
|-------|------|-------------|--------------|
| **1_Generator** | Creates content | 0.7 | Generates financial descriptions using pre-formatted values |
| **2_Auditor** | Verifies accuracy | 0.3 | Validates that formatted values match source data |
| **3_Refiner** | Polishes content | 0.5 | Refines without over-shortening (preserves context) |
| **4_Validator** | Final check | 0.2 | Final quality control and format validation |

**Key**: AI agents now **validate** pre-formatted values instead of converting them.

---

## Features

✅ **Smart Formatting**: 万/K = 1 d.p., 亿/million = 2 d.p.  
✅ **Negative Earnings**: Auto-converts to loss terminology  
✅ **Filters Sub-accounts**: Removes "应付利息_借款利息" patterns  
✅ **Uses Totals**: Not line items  
✅ **Content Preservation**: Agent 3 keeps important context  
✅ **Unified Logging**: Per-run folders with full data  
✅ **Multi-threading**: Parallel processing enabled  

---

## Troubleshooting Extraction

### If `extract_data_from_excel` returns None or empty:

**Run the diagnostic tool**:
```bash
python test_extraction.py
```

**Common Issues**:
1. ❌ **File not found** → Check file path
2. ❌ **Sheet names don't match** → Check `mappings.yml` aliases
3. ❌ **Missing indicators** → Sheets need "Indicative adjusted" or "示意性调整后"
4. ❌ **Missing currency units** → Sheets need "CNY'000" or "人民币千元"
5. ❌ **Entity name wrong** → Use exact name from Excel or "" for single entity
6. ❌ **No valid dates** → Date row must have parseable dates

**Quick Fix Example**:
```python
# ❌ Wrong
dfs = extract_data_from_excel("databook.xlsx", "WrongName", "BS")

# ✅ Correct
dfs = extract_data_from_excel(
    databook_path="databook.xlsx",
    entity_name="",  # Empty for single entity
    mode="All"       # Get everything
)
```

**Debug Checklist**:
- [ ] File path exists: `os.path.exists("databook.xlsx")`
- [ ] Sheets have "Indicative adjusted" or "示意性调整后"
- [ ] Sheets have "CNY'000" or "人民币千元"
- [ ] Sheet names match `mappings.yml` aliases
- [ ] Entity name matches Excel (or use "")
- [ ] Valid dates in date row

---

## Configuration

### AI Parameters
Edit `fdd_utils/config.yml`:

```yaml
agents:
  1_Generator:
    temperature: 0.7       # Higher = more creative
    max_tokens: 2000
  2_Auditor:
    temperature: 0.3       # Lower = more precise
    max_tokens: 2000
```

### API Setup
```yaml
local:
  api_base: "http://localhost:1234"
  api_key: "local"
  chat_model: "your-model"

deepseek:
  api_key: "your-key"
  chat_model: "deepseek-chat"
```

---

## Agent Prompts

**Where to edit prompts**:
- **1_Generator**: `fdd_utils/mappings.yml` (account-specific)
- **2_Auditor/3_Refiner/4_Validator**: `fdd_utils/prompts.yml` (generic)

**Important**: Prompts now instruct AI that values are **pre-formatted**. Don't modify number conversion logic in prompts - it's handled in code.

---

## Examples

### Example 1: Chinese Databook
```python
dfs, keys, _, lang = extract_data_from_excel(
    databook_path="240624.联洋-databook.xlsx",
    entity_name="联洋",
    mode="All"
)

# Values automatically formatted with 万/亿
if '货币资金' in dfs:
    print(dfs['货币资金'])
    # Shows: 7.8万, 123.5万, 1.23亿 (already formatted)
```

### Example 2: English Databook
```python
dfs, keys, _, lang = extract_data_from_excel(
    databook_path="inputs/221128.Project TK.Databook.JW.xlsx",
    entity_name="Haining Wanpu",
    mode="BS"
)

# Values automatically formatted with K/million
if 'Cash' in dfs:
    print(dfs['Cash'])
    # Shows: 78.2K, 12.35 million (already formatted)
```

### Example 3: Using Utilities
```python
# Format numbers manually
from fdd_utils.process_databook import format_value_by_language

format_value_by_language(78200, 'Chi')      # → "7.8万"
format_value_by_language(123456789, 'Chi')  # → "1.23亿" (2 d.p.)
format_value_by_language(78200, 'Eng')      # → "78.2K"
format_value_by_language(12345678, 'Eng')   # → "12.35 million" (2 d.p.)
```

---

## Testing

### Test Number Formatting
```bash
python test_number_formatting.py
```

Expected: All ✅ (24 tests)
- Chinese formatting: 10/10
- English formatting: 12/12
- Edge cases: 2/2

### Test Extraction
```bash
python test_extraction.py
```

Diagnoses why extraction might fail and provides detailed feedback.

### Example Usage
```bash
python example_extraction.py
```

Shows working examples of extraction and AI pipeline.

---

## Recent Updates (Nov 2025)

### ✅ Decimal Places Updated
- **万/K**: 1 decimal place (e.g., 7.8万, 78.2K)
- **亿/million**: 2 decimal places (e.g., 1.23亿, 12.35 million)

### ✅ Agent 3 Content Preservation
- No longer over-shortens content
- Preserves important explanations and context
- Only removes truly redundant information

### ✅ Negative Retained Earnings
- Automatically converts terminology:
  - 未分配利润 (negative) → 未弥补亏损
  - Retained Earnings (negative) → Accumulated Losses
- Displays values as positive

---

## FAQ

**Q: Why does extraction return None?**  
A: Run `python test_extraction.py` for diagnosis. Most common: sheet names don't match mappings.yml or missing financial indicators.

**Q: Why are values already formatted?**  
A: Number formatting happens in code (not AI) for accuracy. AI only validates the formatted values.

**Q: Why do all agents show same output?**  
A: Normal if content is already good. Agents only fix issues, not force changes.

**Q: How do I adjust AI behavior?**  
A: Edit temperatures in `config.yml` (higher = more creative, lower = more precise).

**Q: Can I add custom accounts?**  
A: Yes! Add to `mappings.yml` with aliases and patterns.

**Q: Agent 3 makes content too short?**  
A: Fixed! Prompts updated to preserve context (Nov 2025).

**Q: How to handle multi-entity databooks?**  
A: Provide exact entity name: `extract_data_from_excel("file.xlsx", "Entity Name", "All")`

---

## Files Structure

```
fdd_utils/
├── process_databook.py       # Excel extraction + formatting
├── content_generation.py     # 4-agent AI pipeline
├── ai_helper.py              # AI model interface
├── mappings.yml              # Account aliases + Generator prompts
├── prompts.yml               # Auditor/Refiner/Validator prompts
├── config.yml                # AI parameters
└── logs/                     # Run outputs (timestamped)

Scripts:
├── test_number_formatting.py  # Test formatting logic
├── test_extraction.py         # Diagnose extraction issues
└── example_extraction.py      # Usage examples
```

---

## Getting Started

1. **Install dependencies**:
   ```bash
   pip install -r fdd_utils/requirements.txt
   ```

2. **Test extraction**:
   ```bash
   python test_extraction.py
   ```

3. **Run examples**:
   ```bash
   python example_extraction.py
   ```

4. **Use in Jupyter**:
   Open `fdd_app.ipynb` 🚀

---

## Support

For issues with:
- **Extraction**: Run `python test_extraction.py`
- **Formatting**: Run `python test_number_formatting.py`
- **AI Pipeline**: Check logs in `fdd_utils/logs/run_*/`

---

**Version**: November 2025  
**Features**: Smart formatting, negative earnings handling, content preservation

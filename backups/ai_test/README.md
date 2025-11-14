# AI Module Testing Suite

This folder contains a comprehensive testing framework for the AI financial report generation module.

## 📁 Folder Structure

```
ai_test/
├── config.json       # AI configuration (API keys, models, etc.)
├── ai_module.py      # Main AI module with generation capabilities
├── test_ai.py        # Test script with various test cases
└── README.md         # This file
```

## 🚀 Quick Start

### 1. Setup

Make sure you have the required dependencies:

```bash
pip install openai
```

### 2. Configure API Keys

Edit `config.json` and add your API keys:

```json
{
    "DEEPSEEK_API_KEY": "your-deepseek-api-key-here",
    "OPENAI_API_KEY": "your-openai-api-key-here",
    ...
}
```

### 3. Run Tests

**Interactive Mode (Recommended):**
```bash
python test_ai.py
```

**Run All Automated Tests:**
```bash
python test_ai.py --all
```

**Test AI Module Directly:**
```bash
python ai_module.py
```

## 📋 Test Modes

### Interactive Mode

The interactive mode provides a menu-driven interface:

1. **Test Cash Analysis (English)** - Test English mode with Cash data
2. **Test Cash Analysis (Chinese)** - Test Chinese translation mode
3. **Test AR Analysis (English)** - Test English mode with AR data
4. **Test AR Analysis (Chinese)** - Test Chinese translation mode
5. **Custom Test** - Enter your own data for testing
6. **Run All Tests** - Execute all predefined tests
7. **Exit** - Close the program

### Automated Mode

Runs all test cases automatically and generates a summary report.

## 🔧 Configuration Options

### AI Providers

The module supports multiple AI providers:

- **DeepSeek** (default) - Primary AI provider
- **OpenAI** - GPT-4 and GPT-3.5 models
- **Local AI** - Local AI models (requires local server)

### Language Modes

- **English Mode** - Generates financial commentary in English
- **Chinese Mode** - Translates to Simplified Chinese with proper formatting

### Parameters

- `temperature` - Controls randomness (0.0-2.0, default: 0.7)
- `max_tokens` - Maximum response length (default: 2000)
- `provider` - AI provider to use (deepseek/openai/local)
- `model` - Specific model name (optional)

## 📝 Sample Financial Data

The test suite includes sample data for:

1. **Cash Analysis**
   - Multiple bank accounts
   - Currency: CNY'000
   - Includes subtotals and totals

2. **Accounts Receivable**
   - Multiple customers
   - Bad debt provisions
   - Aging analysis

## 🧪 Test Workflow

Each test executes a multi-agent workflow:

### Agent 1: Content Generation
- Analyzes financial data
- Generates professional commentary
- Follows predefined patterns
- Includes specific figures and entities

### Agent 2: AI Proofreader
- Validates Agent 1 output
- Checks pattern compliance
- Verifies entity accuracy
- Ensures proper formatting
- Returns corrected content

## 📊 Test Results

After running tests, you'll see:

1. **Individual Test Results**
   - Agent 1 status (success/failure)
   - Agent 2 status (success/failure)
   - Token usage
   - Generated content preview

2. **Summary Report**
   - Total tests run
   - Tests passed/failed
   - Success rate
   - Total token usage

## 🛠️ Using the AI Module in Your Code

```python
from ai_module import AIModule

# Initialize
ai = AIModule()

# Generate content
result = ai.generate_content(
    system_prompt="You are a financial analyst.",
    user_prompt="Analyze this cash data...",
    provider='deepseek',
    mode='english'
)

print(result['content'])
print(f"Tokens used: {result['tokens']['total_tokens']}")

# Test multi-agent workflow
results = ai.test_multi_agent(
    financial_data=your_data,
    key="Cash",
    entity_name="Your Company",
    mode='english',
    provider='deepseek'
)
```

## 🔍 Debugging

Enable verbose output by checking the console logs during test execution. The module prints:

- ✅ Success messages (green)
- ❌ Error messages (red)
- 🤖 AI processing steps
- 📊 Token usage statistics
- 📝 Content previews

## 📄 Output Format

### Agent 1 Output
Professional financial commentary following patterns from `prompts.json`

### Agent 2 Output
JSON format with validation results:
```json
{
    "is_compliant": true,
    "issues": [],
    "corrected_content": "Final cleaned content...",
    "figure_checks": [...],
    "entity_checks": [...],
    "grammar_notes": [...]
}
```

## 🔐 Security Notes

- Never commit `config.json` with real API keys to version control
- Use environment variables for production deployments
- Keep API keys secure and rotate them regularly

## 🐛 Troubleshooting

**Issue: "Provider not available"**
- Check that API keys are configured in `config.json`
- Verify the provider name is correct (deepseek/openai/local)

**Issue: "Error loading prompts"**
- Ensure `fdd_utils/prompts.json` exists in parent directory
- Check file permissions

**Issue: "Connection error"**
- Verify internet connection
- Check API endpoint URLs in config
- For local AI, ensure local server is running

## 📚 Further Documentation

For more information about the main application, see:
- `/fdd_utils/` - Main utility modules
- `/common/assistant.py` - Full AI assistant implementation
- `/fdd_utils/prompts.json` - Complete prompt templates

## 🤝 Contributing

To add new test cases:

1. Edit `test_ai.py`
2. Add sample data constants
3. Create new test function
4. Add to test menu in `interactive_test()`

---

**Version:** 1.0  
**Last Updated:** 2025-01-11  
**Supported Languages:** English, Simplified Chinese


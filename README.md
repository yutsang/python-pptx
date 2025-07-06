# Financial Data Processor - Streamlit App

A Streamlit-based web application for processing and analyzing financial data from Excel files, built on the functionality from the original `old_ver` implementation.

## Features

- 📊 **Excel File Processing**: Upload and process financial Excel files
- 🔍 **Worksheet Section Viewing**: View filtered worksheet sections using the `process_and_filter_excel` function
- 🏢 **Entity Support**: Support for Haining, Nanjing, and Ningbo entities
- 📋 **Data Visualization**: Interactive tables and data statistics
- 💾 **Export Functionality**: Download processed data as markdown files
- ⚙️ **Configuration Management**: View and manage mapping and pattern configurations

### Enhanced Highlighting System (NEW)

The application now includes an advanced highlighting system in Agent 2 that performs pattern-based figure detection:

1. **Pattern Comparison**: Compares AI1 output with the most similar pattern from pattern.json
2. **Figure Extraction**: Extracts numbers from AI1 content and identifies pattern placeholders
3. **'000 Detection**: Automatically detects '000 notation in worksheet headers/titles
4. **Conversion Logic**: Converts balance sheet numbers based on notation (e.g., 9076 → 9076000)
5. **Row Highlighting**: Highlights all worksheet rows containing matching numbers in yellow

**Example Process:**
- AI1 Output: "CNY9.1M"
- Pattern Detection: Finds most similar pattern with [Amount] placeholder
- '000 Detection: Detects '000 notation in worksheet headers
- Conversion: 9.1M → 9,100,000 → 9,100 (worksheet number)
- Highlighting: All rows containing "9100" are highlighted in yellow

### AI Agents

- **Agent 1**: Content generation specialist
- **Agent 2**: Enhanced data integrity specialist with pattern-based highlighting
- **Agent 3**: Pattern compliance specialist

## Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd python-pptx
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Ensure configuration files are present in the `utils/` directory:
   - `utils/config.json`
   - `utils/mapping.json`
   - `utils/pattern.json`

## Usage

1. Run the Streamlit app:
```bash
streamlit run app.py
```

2. Open your browser and navigate to the provided URL (usually `http://localhost:8501`)

3. Upload an Excel file containing financial data

4. Select the entity (Haining, Nanjing, or Ningbo)

5. Enter entity helpers (comma-separated, e.g., "Wanpu,Limited,")

6. Click "Process Data" to analyze the file

## File Structure

```
python-pptx/
├── app.py                 # Main Streamlit application
├── requirements.txt       # Python dependencies
├── README.md             # This file
├── utils/                # Configuration and utility files
│   ├── config.json       # AI service configuration
│   ├── mapping.json      # Financial item mappings
│   ├── pattern.json      # Report patterns
│   └── ...
└── old_ver/              # Original implementation (untouched)
    ├── utils/
    ├── financial_processor/
    └── ...
```

## Expected Excel File Format

The application expects Excel files with the following structure:

- **Sheet Names**: BSHN (Haining), BSNJ (Nanjing), BSNB (Ningbo)
- **Financial Items**: Cash, AR, Prepayments, OR, Other CA, IP, Other NCA, AP, Taxes payable, OP, Capital, Reserve
- **Data Format**: Standard balance sheet format with descriptions and amounts

## Configuration

### Entity Mapping
- **Haining** → BSHN sheet
- **Nanjing** → BSNJ sheet  
- **Ningbo** → BSNB sheet

### Supported Financial Items
- **Current Assets**: Cash, AR, Prepayments, OR, Other CA
- **Non-current Assets**: IP, Other NCA
- **Liabilities**: AP, Taxes payable, OP
- **Equity**: Capital, Reserve

## Features from Original Implementation

This Streamlit app incorporates the core functionality from the original `old_ver` implementation:

- ✅ `process_and_filter_excel` function for worksheet section viewing
- ✅ Entity-based data filtering
- ✅ Financial item mapping and pattern matching
- ✅ Configuration file management
- ✅ Data processing and analysis capabilities

## Troubleshooting

1. **Configuration Files Missing**: Ensure all JSON configuration files are present in the `utils/` directory
2. **Excel File Format**: Verify your Excel file has the expected sheet names and structure
3. **Dependencies**: Make sure all required packages are installed using `pip install -r requirements.txt`

## Development

The original implementation files are preserved in the `old_ver/` directory and remain untouched as requested. The new Streamlit app provides a modern web interface while maintaining the core functionality of the original system.
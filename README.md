# Real Estate Due Diligence Automation

An enterprise-grade financial data processing system for real estate due diligence reports.

## 🎯 **Overview**

This system automates the creation of due diligence reports by:
- 📊 Processing Excel financial data with intelligent pattern matching
- 🤖 Using a 3-agent AI validation pipeline (Content Generation → Data Validation → Pattern Compliance)
- 📋 Generating professional PowerPoint presentations
- 🔍 Providing comprehensive data validation and error checking

## 🚀 **Quick Start**

### **Installation**
```bash
# Clone the repository
git clone <repository-url>
cd python-pptx

# Install dependencies
pip install -r requirements.txt

# Optional: Install AI dependencies for real AI processing
pip install -r utils/requirements_ai.txt
```

### **Running the Application**
```bash
# Run the main application
streamlit run streamlit_app.py
```

### **AI Configuration (Optional)**
For real AI processing, configure your API key:
```json
// config/config.json
{
    "OPENAI_API_KEY": "sk-your-api-key-here",
    "CHAT_MODEL": "gpt-4o-mini"
}
```

## 📊 **Key Features**

### **🤖 AI Processing Pipeline**
- **Agent 1**: Content Generation using pattern templates
- **Agent 2**: Data validation and accuracy checking
- **Agent 3**: Pattern compliance and format verification

### **📋 Data Processing**
- **Excel Processing**: Automated extraction of financial data
- **Entity Recognition**: Smart parsing of real estate entity names
- **Key-based Analysis**: Organized processing by financial categories
- **PowerPoint Export**: Automated presentation generation

### **🔍 UI Features**
- **Single Progress Bar**: Real-time processing status
- **2-Layer Tabs**: Organized AI results by Agent → Key structure
- **AI Logging**: Complete session tracking with timestamps
- **Export Options**: PowerPoint presentations and AI log summaries

## 🔧 **Configuration**

### **Directory Structure**
```
├── streamlit_app.py          # Main application
├── config/
│   ├── config.json           # AI and system configuration
│   ├── mapping.json          # Financial key mappings
│   ├── pattern.json          # Processing patterns
│   └── prompts.json          # AI agent prompts
├── utils/
│   ├── ai_config.py          # AI service initialization
│   ├── ai_logger.py          # AI interaction logging
│   └── requirements_ai.txt   # Optional AI dependencies
└── logging/                  # AI session logs with timestamps
```

### **Configuration Files**
- **config.json**: AI API keys and system settings
- **mapping.json**: Financial category mappings  
- **pattern.json**: Data processing patterns
- **prompts.json**: AI agent system prompts

## 📈 **Usage Workflow**

### **Processing Financial Data**
1. **Upload Excel File**: Financial data spreadsheet
2. **Select Entity**: Choose real estate entity for analysis  
3. **Filter Keys**: Select financial categories to process
4. **AI Processing**: Optional 3-agent analysis pipeline
5. **Export Results**: PowerPoint presentation generation

### **AI Processing Status**
- **🚀 Real AI**: When valid API key is configured
- **⚠️ No AI**: System works without AI, skips AI processing steps

## 🗂️ **AI Logging**

All AI interactions are automatically logged to `logging/` directory:

```
logging/ai_session_YYYYMMDD_HHMMSS/
├── session_summary.json      # Complete session metadata
├── detailed_interactions.jsonl # Line-by-line AI interactions  
├── summary.md                # Human-readable summary
└── Agent_X_Key_Y.json        # Individual interaction files
```

### **Log Contents**
- Input prompts (system and user)
- AI responses (full content)
- Processing times and connection status
- Success/error tracking
- Response type classification

## 🛠️ **Troubleshooting**

### **Common Issues**

#### **AI Services Not Available**
- Check if API key is configured in `config/config.json`
- Install AI dependencies: `pip install -r utils/requirements_ai.txt`
- System works without AI - will skip AI processing steps

#### **Excel Processing Errors**
- Ensure Excel file is in .xlsx or .xls format
- Check for required worksheets and data structure
- Verify entity name matches supported entities

#### **PowerPoint Generation Issues**
- Ensure template.pptx exists in utils/ directory
- Check file permissions for output directory

## 📄 **License**

This project is licensed under the MIT License - see the LICENSE file for details.

---

**🚀 Ready for professional real estate due diligence processing!** 🏠✨
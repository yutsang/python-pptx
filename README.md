# Real Estate Due Diligence Report Generator

This project generates professional due diligence reports for real estate projects using Excel databooks, predefined text patterns, and (optionally) AI for text generation. It supports both command-line and web (Streamlit) interfaces.

## Features
- Upload Excel databooks and extract financial tables
- Pattern-based text generation (AI or test mode)
- Quality assurance and auto-correction of generated text
- Editable output for each report section
- Export to PPTX (coming soon)

## Usage

### 1. Streamlit Web App
Launch the interactive UI:
```bash
streamlit run app.py
```
- Upload your Excel databook
- Select entity and helpers
- View each worksheet in a tab
- Click "Generate Text" to fill in report sections
- Edit the generated text as needed
- Click "Export to PPTX" (feature coming soon)

### 2. Command-Line Interface
Run the CLI for batch/scripted use:
```bash
python main.py -i <excel_file> -e <entity> [--helpers ...] [--ai] [--output <pptx>] [--config <config>] [--mapping <mapping>] [--pattern <pattern>]
```
Example:
```bash
python main.py -i utils/your_databook.xlsx -e Haining --helpers Wanpu Limited --ai --output report.pptx
```
- Omit `--ai` to use local/test mode (no AI required)

## Configuration & Data
- Place your Excel databooks in a known location (e.g., `utils/`)
- Edit `utils/config.json`, `utils/mapping.json`, and `utils/pattern.json` as needed
- All business logic is in `common/assistant.py`

## Project Structure
- `common/assistant.py` — All core logic (Excel scraping, pattern filling, QA, etc.)
- `utils/` — Config and pattern files
- `main.py` — CLI entry point
- `app.py` — Streamlit web app

## Clean Codebase
All legacy, duplicate, and unused scripts have been removed for clarity and maintainability. Only the above files are required for the workflow.

---
For questions or improvements, please open an issue or PR.
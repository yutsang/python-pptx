# PowerPoint Generator for Financial Reports

A Python-based PowerPoint presentation generator that converts structured markdown content into professionally formatted financial reports using PowerPoint templates.

## Features

- **Markdown to PPTX Conversion**: Transform nested markdown content into PowerPoint slides with proper hierarchy
- **Template-Driven Design**: Uses PowerPoint templates with predefined layouts for consistent formatting
- **Automatic Content Distribution**: Intelligently splits content across slides and sections while maintaining continuity markers
- **Advanced Formatting**:
  - Version-compatible paragraph styling (supports both legacy and modern python-pptx APIs)
  - XML-based bullet formatting with precise indentation controls
  - Dual-column slide support for complex layouts
- **Summary Section Handling**: Specialized summary slide generation with margin controls
- **Slide Management**: Automatic detection and removal of unused slides

## Installation
`pip install python-pptx lxml Pillow XlsxWriter`

## Usage
### Basic Example
`from pptx_generator import PowerPointGenerator`

### Initialize with your template
`generator = PowerPointGenerator("financial_template.pptx")`

### Your markdown content (simplified example)
```
md_content = """
Assets
Current Assets
Cash and equivalents: $45M
Accounts receivable: $30M
Liabilities
Short-Term Debt
Commercial paper: $15M
"""
```

### Generate presentation
```
generator.generate_full_report(
md_content=md_content,
summary_md="## Summary\nStrong liquidity position with current ratio of 2.4",
output_path="financial_report.pptx"
)
```
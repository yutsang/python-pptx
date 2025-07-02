import json, os, httpx
import pandas as pd
from tabulate import tabulate
from pathlib import Path
import re
from tqdm import tqdm
from typing import Dict, List, Optional
import numpy as np
import openpyxl

# AI-related imports (mocked if not available)
try:
    from openai import AzureOpenAI
    from azure.search.documents import SearchClient
    from azure.core.credentials import AzureKeyCredential
    AI_AVAILABLE = True
except ImportError:
    AzureOpenAI = None
    SearchClient = None
    AzureKeyCredential = None
    AI_AVAILABLE = False

# --- Config and AI Service Helpers ---
def load_config(file_path):
    """Load configuration from a JSON file."""
    with open(file_path) as config_file:
        config_details = json.load(config_file)
    return config_details

def initialize_ai_services(config_details):
    """Initialize Azure OpenAI and Azure Search clients using config details."""
    if not AI_AVAILABLE:
        raise RuntimeError("AI services not available on this machine.")
    httpx_client = httpx.Client(verify=False)
    if AzureOpenAI is None or SearchClient is None or AzureKeyCredential is None:
        raise RuntimeError("AI modules not available.")
    oai_client = AzureOpenAI(
        api_base=config_details['OPENAI_API_BASE'],
        api_key=config_details['OPENAI_API_KEY'],
        api_version=config_details['OPENAI_API_VERSION_COMPLETION'],
        client=httpx_client
    )
    search_client = SearchClient(
        endpoint=f"https://{config_details['AZURE_AI_SEARCH_SERVICE_ENDPOINT']}/",
        index_name=config_details['AZURE_SEARCH_INDEX_NAME'],
        credential=AzureKeyCredential(config_details['AZURE_SEARCH_API_KEY']),
        connection_verify=False,
        headers={"Host": f"{config_details['AZURE_SEARCH_SERVICE_NAME']}.search.windows.net"}
    )
    return oai_client, search_client

def generate_response(user_query, system_prompt, oai_client, context_content, openai_chat_model):
    """Generate a response from the AI model given a user query and system prompt."""
    conversation = [
        {"role": "system", "content": system_prompt},
        {"role": "assistant", "content": f"Context data: \n{context_content}"},
        {"role": "user", "content": user_query}
    ]
    response = oai_client.chat.completions.create(
        model=openai_chat_model,
        messages=conversation,
    )
    return response.choices[0].message.content

# --- Excel and Data Processing ---
def find_dense_blocks(df, min_rows=2, min_cols=3, density_threshold=0.6):
    blocks = []
    nrows, ncols = df.shape
    for row_start in range(nrows - min_rows + 1):
        for col_start in range(ncols - min_cols + 1):
            for row_end in range(row_start + min_rows, nrows + 1):
                for col_end in range(col_start + min_cols, ncols + 1):
                    block = df.iloc[row_start:row_end, col_start:col_end]
                    total_cells = block.size
                    non_empty_cells = block.notnull().values.sum()
                    if total_cells > 0 and (non_empty_cells / total_cells) >= density_threshold:
                        # Avoid duplicates
                        if not any((row_start >= b[0] and row_end <= b[1] and col_start >= b[2] and col_end <= b[3]) for b in blocks):
                            blocks.append((row_start, row_end, col_start, col_end))
    return blocks

def process_and_filter_excel(filename, tab_name_mapping, entity_name, entity_suffixes):
    try:
        main_dir = Path(__file__).parent.parent
        file_path = main_dir / filename
        wb = openpyxl.load_workbook(file_path, data_only=True)
        markdown_content = ""
        entity_keywords = [entity_name] + list(entity_suffixes)
        entity_keywords = [kw.strip().lower() for kw in entity_keywords if kw.strip()]
        for ws in wb.worksheets:
            if ws.title not in tab_name_mapping:
                continue
            for tbl in ws._tables.values():
                ref = tbl.ref
                min_col, min_row, max_col, max_row = openpyxl.utils.range_boundaries(ref)
                data = []
                for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col, values_only=True):
                    data.append(row)
                if not data or len(data) < 2:
                    continue
                df = pd.DataFrame(data[1:], columns=data[0])
                df = df.dropna(how='all').dropna(axis=1, how='all')
                df = df.applymap(lambda x: str(x) if x is not None else "")  # Convert every cell to string
                df = df.reset_index(drop=True)
                # Flatten all cell values to a list of strings
                all_cells = df.apply(lambda x: x.str.lower().strip(), axis=1).values.flatten().tolist()
                # Check if any entity keyword appears in any cell
                match_found = any(any(kw in cell for cell in all_cells) for kw in entity_keywords)
                if match_found:
                    print(f"[DEBUG] Table '{tbl.name}' in sheet '{ws.title}' included for entity keywords: {entity_keywords}")
                    try:
                        markdown_content += tabulate(df, headers='keys', tablefmt='pipe') + '\n\n'
                    except Exception:
                        markdown_content += df.to_markdown(index=False) + '\n\n'
                else:
                    print(f"[DEBUG] Table '{tbl.name}' in sheet '{ws.title}' skipped for entity keywords: {entity_keywords}")
        return markdown_content
    except Exception as e:
        print("An error occurred while processing the Excel file:", e)
    return ""

def find_financial_figures_with_context_check(filename, sheet_name, date_str, convert_thousands=False):
    try:
        file_path = Path(filename)
        xl = pd.ExcelFile(file_path)
        if sheet_name not in xl.sheet_names:
            print(f"Sheet '{sheet_name}' not found in the file.")
            return {}
        df = xl.parse(sheet_name)
        if not isinstance(df, pd.DataFrame):
            return {}
        df.columns = ['Description', 'Date_2020', 'Date_2021', 'Date_2022']
        date_column_map = {
            '31/12/2020': 'Date_2020',
            '31/12/2021': 'Date_2021',
            '30/09/2022': 'Date_2022'
        }
        if date_str not in date_column_map:
            print(f"Date '{date_str}' not recognized.")
            return {}
        date_column = date_column_map[date_str]
        # If convert_thousands and '000' in columns or first row, multiply numeric columns by 1000
        scale_factor = 1000 if (convert_thousands and any("'000" in str(col) for col in df.columns)) else 1
        financial_figure_map = {
            "Cash": "Cash at bank",
            "AR": "Accounts receivable",
            "Prepayments": "Prepayments",
            "OR": "Other receivables",
            "Other CA": "Other current assets",
            "IP": "Investment properties",
            "Other NCA": "Other non-current assets",
            "AP": "Accounts payable",
            "Taxes payable": "Taxes payable",
            "OP": "Other payables",
            "Capital": "Paid-in capital",
            "Reserve": "Surplus reserve"
        }
        financial_figures = {}
        for key, desc in financial_figure_map.items():
            if 'Description' in df.columns and date_column in df.columns:
                value = df.loc[df['Description'].str.contains(desc, case=False, na=False), date_column].values
                if value.size > 0:
                    financial_figures[key] = float(value[0]) * scale_factor
        return financial_figures
    except Exception as e:
        print(f"An error occurred while processing the Excel file: {e}")
    return {}

def get_tab_name(project_name):
    if project_name == 'Haining':
        return "BSHN"
    elif project_name == 'Nanjing':
        return "BSNJ"
    elif project_name == 'Ningbo':
        return "BSNB"

def get_financial_figure(financial_figures, key):
    figure = financial_figures.get(key, None)
    if figure is None:
        return f"{key} not found in the financial figures."
    if figure > 1000000:
        return f"{figure / 1000000:.1f}M"
    elif figure >= 1000:
        return f"{figure / 1000:,.0f}K"
    else:
        return f"{figure:.1f}"

def detect_string_in_file(file_content, target_string):
    try:
        return target_string in file_content
    except Exception:
        return False

def load_ip(file, key=None):
    try:
        with open(file, 'r') as f:
            data = json.load(f)
        if key is not None and key in data:
            return data[key]
        return data
    except FileNotFoundError:
        print(f"File {file} not found.")
    except json.JSONDecodeError:
        print(f"Error decoding JSON from file {file}.")
    return {}

# --- Pattern Filling and Main Processing ---
def process_keys(keys, entity_name, entity_helpers, input_file, mapping_file, pattern_file, config_file='utils/config.json', use_ai=True, convert_thousands=False):
    # Use test data if AI is not available
    if not use_ai or not AI_AVAILABLE:
        return generate_test_results(keys)
    financial_figures = find_financial_figures_with_context_check(input_file, get_tab_name(entity_name), '30/09/2022', convert_thousands=convert_thousands)
    system_prompt = """
        Role: system,
        Content: You are a senior financial analyst specializing in due diligence reporting. Your task is to integrate actual financial data from databooks into predefined report templates.
        CORE PRINCIPLES:
        1. SELECT exactly one appropriate non-nil pattern from the provided pattern options
        2. Replace all placeholder values with corresponding actual data
        3. Output only the financial completed pattern text, never show template structure
        4. ACCURACY: Use only provided - data - never estimate or extrapolate
        5. CLARITY: Write in clear business English, translating any foreign content
        6. FORMAT: Follow the exact template structure provided
        7. CURRENCY: Express figures to Thousands (K) or Millions (M) as appropriate
        8. CONCISENESS: Focus on material figures and key insights only
        OUTPUT REQUIREMENTS:
        - Choose the most suitable single pattern based on available data
        - Replace all placeholders with actaul figures from databook
        - Output ONLY the final text - no pattern names, no template structure, no explanations
        - If data is missing for a pattern, select a different pattern that has complete data
        - Never output JSON structure or pattern formatting
    """
    results = {}
    pbar = tqdm(keys, desc="Processing keys", unit="key")
    for key in pbar:
        config_details = load_config(config_file)
        oai_client, search_client = initialize_ai_services(config_details)
        openai_model = config_details['CHAT_MODEL']
        pattern = load_ip(pattern_file, key)
        mapping = {key: load_ip(mapping_file)}
        excel_tables = process_and_filter_excel(input_file, mapping, entity_name, entity_helpers)
        detect_zeros = "3. The figures in this table is already expressed in k, express the number in M (divide by 1000), rounded to 1 decimal place, if final figure less than 1M, express in K (no decimal places)." if detect_string_in_file(excel_tables, "'000") else ""
        user_query = f"""
        TASK: Select ONE pattern and complete it with actaul data
        AVAILABLE PATTERNS: {json.dumps(pattern, indent=2)}
        FINANCIAL FIGURE: {key}: {get_financial_figure(financial_figures, key)}
        DATA SOURCE: {excel_tables}
        SELECTION CRITERIA:
        - Choose the pattern with the most complete data coverage
        - Prioritize patterns that match the primary account category
        - use most recent data: latest available
        - {detect_zeros}
        REQUIRED OUTPUT FORMAT:
        - Only the completed pattern text
        - No pattern names or labels
        - No template structure
        - No JSON formatting
        - Replace ALL 'xxx' or placeholders with actaul data values
        - Do not use bullet point for listing
        - Check all numbers if they are in thousands (K) or millions (M) and express accordingly, do appropriate convertion (K or M) for any number > 1000
        - No foreign contents, if any, translate to English
        - Stick to Template format, no extra explanations or comments
        - For entity name to be filled into template, it should not be the reporting entity ({entity_name}) itself, it must be from the DATA SOURCE
        - For all listing figures, please check the total, together should be around the same or consituing majority of FINANCIAL FIGURE
        Example of CORRECT output format:
        "Cash at bank comprioses deposits of $2.3M held with major financial institutions as at 30/09/2022."
        Example of INCORRECT output format:
        "Pattern 1: Cash at bank comprises deposits of xxx held with xxx as at xxx."        
        """
        response_txt = generate_response(user_query, system_prompt, oai_client, excel_tables, openai_model)
        results[key] = response_txt
        pbar.set_postfix_str(f"key={key}")
    return results

def generate_test_results(keys):
    # Generate mock/test results for each key
    return {key: f"[TEST] Example output for {key}." for key in keys}

# --- QA Agent ---
class QualityAssuranceAgent:
    def __init__(self):
        self.excellent_threshold = 90
        self.good_threshold = 80
        self.acceptable_threshold = 70
        self.template_artifacts = [
            'Pattern 1:', 'Pattern 2:', 'Pattern 3:', '[', ']', '{', '}', 'xxx', 'XXX',
            'template', 'placeholder', 'PLACEHOLDER', 'TBD', 'TODO', 'FIXME'
        ]
        self.professional_terms = [
            'comprised', 'represented', 'indicated', 'demonstrated', 'reflected',
            'maintained', 'established', 'confirmed', 'verified', 'assessed',
            'evaluated', 'analyzed'
        ]
        self.risk_indicators = [
            'provision', 'impairment', 'restricted', 'covenant', 'collateral',
            'mortgage', 'guarantee', 'contingent'
        ]
    def validate_content(self, content: str) -> Dict:
        # Simple QA: check for template artifacts, paragraph structure, and number formatting
        issues = []
        score = 100
        for artifact in self.template_artifacts:
            if artifact.lower() in content.lower():
                issues.append(f"Template artifact found: '{artifact}'")
                score -= 5
        if not re.search(r'^##\s+\w+', content, re.MULTILINE):
            issues.append("Missing proper markdown headers")
            score -= 10
        if re.search(r'###\s+[^\n]+\n\s*\n', content):
            issues.append("Empty content sections detected")
            score -= 5
        paragraphs = [p.strip() for p in content.split('\n\n') if p.strip()]
        if len(paragraphs) < 3:
            issues.append("Insufficient content paragraphs")
            score -= 5
        return {"score": max(0, score), "issues": issues}
    def auto_correct(self, content: str) -> str:
        # Remove template artifacts and fix paragraph structure
        for artifact in self.template_artifacts:
            content = re.sub(re.escape(artifact), '', content, flags=re.IGNORECASE)
        # Ensure double newlines between paragraphs
        content = re.sub(r'\n{2,}', '\n\n', content)
        return content.strip() 
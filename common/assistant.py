import json, os, httpx
import pandas as pd
from tabulate import tabulate
from pathlib import Path
import re
from tqdm import tqdm
from typing import Dict, List, Optional
import numpy as np
import openpyxl
from utils.cache import get_cache_manager, cached_function

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
    if AzureOpenAI is None:
        raise RuntimeError("AI modules not available.")
    
    # Initialize OpenAI client
    oai_client = AzureOpenAI(
        azure_endpoint=config_details['OPENAI_API_BASE'],
        api_key=config_details['OPENAI_API_KEY'],
        api_version=config_details['OPENAI_API_VERSION_COMPLETION'],
        http_client=httpx_client
    )
    
    # Initialize search client only if all required configurations are available
    search_client = None
    try:
        # Check if all required Azure Search configurations are present
        required_search_configs = [
            'AZURE_AI_SEARCH_SERVICE_ENDPOINT',
            'AZAURE_AI_SEARCH_INDEX_NAME',
            'SEARCH_API_KEY',
            'AZURE_AI_SEARCH_SERVICE_NAME'
        ]
        
        if all(config_details.get(key) for key in required_search_configs) and SearchClient is not None and AzureKeyCredential is not None:
            # Configure search client settings
            search_client_configs = {
                'connection_verify': False,
                'headers': {"Host": f"{config_details['AZURE_AI_SEARCH_SERVICE_NAME']}.search.windows.net"}
            }
            
            search_client = SearchClient(
                'endpoint': f"https://{config_details['AZURE_AI_SEARCH_SERVICE_ENDPOINT']}/",
                'index_name': config_details['AZAURE_AI_SEARCH_INDEX_NAME'],
                'credential': AzureKeyCredential(config_details['SEARCH_API_KEY']),
                **search_client_configs
            )
        else:
            print("Azure Search configuration incomplete or modules not available. Search client will be None.")
    except Exception as e:
        print(f"Failed to initialize Azure Search client: {e}")
        search_client = None
    
    return oai_client, search_client

def generate_response(user_query, system_prompt, oai_client, context_content, openai_chat_model):
    """Generate a response from the AI model given a user query and system prompt with caching."""
    # Check cache first
    cache_manager = get_cache_manager()
    cached_response = cache_manager.get_cached_ai_response(user_query, system_prompt, context_content)
    if cached_response is not None:
        return cached_response
    
    conversation = [
        {"role": "system", "content": system_prompt},
        {"role": "assistant", "content": f"Context data: \n{context_content}"},
        {"role": "user", "content": user_query}
    ]
    response = oai_client.chat.completions.create(
        model=openai_chat_model,
        messages=conversation,
    )
    
    response_content = response.choices[0].message.content
    
    # Cache the response
    cache_manager.cache_ai_response(user_query, system_prompt, context_content, response_content)
    
    return response_content

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

def extract_tables_robust(worksheet, entity_keywords):
    """
    Robust table extraction using the original method from utils.py
    """
    tables = []
    
    try:
        # Method 1: Try to extract from openpyxl tables (works for individually formatted tables)
        if hasattr(worksheet, '_tables') and worksheet._tables:
            for tbl in worksheet._tables.values():
                try:
                    ref = tbl.ref
                    from openpyxl.utils import range_boundaries
                    min_col, min_row, max_col, max_row = range_boundaries(ref)
                    data = []
                    for row in worksheet.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col, values_only=True):
                        data.append(row)
                    if data and len(data) >= 2:
                        tables.append({
                            'data': data,
                            'method': 'openpyxl_table',
                            'name': tbl.name,
                            'range': ref
                        })
                except Exception as e:
                    print(f"Failed to extract table {tbl.name}: {e}")
                    continue
        
        # Method 2: Original method from utils.py - DataFrame splitting on empty rows
        try:
            # Convert worksheet to DataFrame
            all_data = []
            for row in worksheet.iter_rows(values_only=True):
                all_data.append(row)
            
            if all_data:
                df = pd.DataFrame(all_data)
                df = df.dropna(how='all').dropna(axis=1, how='all')
                
                if len(df) >= 2:
                    # Split dataframes on empty rows (original method)
                    empty_rows = df.index[df.isnull().all(1)]
                    start_idx = 0
                    dataframes = []
                    
                    for end_idx in empty_rows:
                        if end_idx > start_idx:
                            split_df = df[start_idx:end_idx]
                            if not split_df.dropna(how='all').empty:
                                dataframes.append(split_df)
                            start_idx = end_idx + 1
                    
                    if start_idx < len(df):
                        dataframes.append(df[start_idx:])
                    
                    # Filter dataframes by entity keywords (original method)
                    combined_pattern = '|'.join(re.escape(kw) for kw in entity_keywords)
                    
                    for i, data_frame in enumerate(dataframes):
                        # Check if dataframe contains entity keywords
                        mask = data_frame.apply(
                            lambda row: row.astype(str).str.contains(
                                combined_pattern, case=False, regex=True, na=False
                            ).any(),
                            axis=1
                        )
                        
                        if mask.any():
                            # Convert DataFrame to list format for consistency
                            table_data = [data_frame.columns.tolist()] + data_frame.values.tolist()
                            
                            # Check if table has meaningful content (not empty)
                            if table_data and len(table_data) > 1:
                                # Check if there's actual data beyond headers
                                has_data = False
                                for row in table_data[1:]:  # Skip header row
                                    if any(cell and str(cell).strip() for cell in row):
                                        has_data = True
                                        break
                                
                                if has_data:
                                    tables.append({
                                        'data': table_data,
                                        'method': 'original_split',
                                        'name': f'original_table_{i}',
                                        'range': f'dataframe_{i}'
                                    })
                            
        except Exception as e:
            print(f"Error in original table detection: {e}")
        
        return tables
        
    except Exception as e:
        print(f"Error in robust table extraction: {e}")
        return []



@cached_function(ttl=1800)  # Cache for 30 minutes
def process_and_filter_excel(filename, tab_name_mapping, entity_name, entity_suffixes):
    try:
        cache_manager = get_cache_manager()
        
        # For uploaded files, try content-based caching first
        original_filename = None
        file_content_hash = None
        
        # Check if this is a temporary uploaded file
        if filename.startswith('temp_ai_processing_'):
            original_filename = filename.replace('temp_ai_processing_', '')
            try:
                # Get file content hash for better caching
                with open(filename, 'rb') as f:
                    file_content = f.read()
                    file_content_hash = cache_manager.get_file_content_hash(file_content)
                
                # Try content-based cache first
                cached_result = cache_manager.get_cached_processed_excel_by_content(
                    file_content_hash, original_filename, entity_name, entity_suffixes
                )
                if cached_result is not None:
                    print(f"📋 Cache hit for {original_filename} (content-based)")
                    return cached_result
            except Exception as e:
                print(f"Content-based cache check failed: {e}")
        
        # Fallback to path-based caching for regular files
        cached_result = cache_manager.get_cached_processed_excel(filename, entity_name, entity_suffixes)
        if cached_result is not None:
            print(f"📋 Cache hit for {filename} (path-based)")
            return cached_result
            
        main_dir = Path(__file__).parent.parent
        file_path = main_dir / filename
        wb = openpyxl.load_workbook(file_path, data_only=True)
        markdown_content = ""
        entity_keywords = [entity_name] + list(entity_suffixes)
        entity_keywords = [kw.strip().lower() for kw in entity_keywords if kw.strip()]
        
        for ws in wb.worksheets:
            if ws.title not in tab_name_mapping:
                continue
            
            # Processing worksheet: {ws.title}
            
            # Use robust table extraction
            tables = extract_tables_robust(ws, entity_keywords)
            
            for table_info in tables:
                try:
                    data = table_info['data']
                    method = table_info['method']
                    table_name = table_info['name']
                    
                    if not data or len(data) < 2:
                        continue
                    
                    # Create DataFrame
                    df = pd.DataFrame(data[1:], columns=data[0])
                    df = df.dropna(how='all').dropna(axis=1, how='all')
                    df = df.applymap(lambda x: str(x) if x is not None else "")
                    df = df.reset_index(drop=True)
                    
                    # Check for entity keywords - handle mixed data types safely
                    all_cells = [str(cell).lower().strip() for cell in df.values.flatten()]
                    match_found = any(any(kw in cell for cell in all_cells) for kw in entity_keywords)
                    
                    if match_found:
                        # Table '{table_name}' (method: {method}) in sheet '{ws.title}' included for entity keywords: {entity_keywords}
                        try:
                            markdown_content += tabulate(df, headers='keys', tablefmt='pipe') + '\n\n'
                        except Exception:
                            markdown_content += df.to_markdown(index=False) + '\n\n'
                    else:
                        # Table skipped for entity keywords
                        pass
                        
                except Exception as e:
                    print(f"Error processing table {table_info.get('name', 'unknown')}: {e}")
                    continue
        
        # Cache the processed result - use content-based caching for uploaded files
        if file_content_hash and original_filename:
            cache_manager.cache_processed_excel_by_content(
                file_content_hash, original_filename, entity_name, entity_suffixes, markdown_content
            )
            print(f"📋 Cached result for {original_filename} (content-based)")
        else:
            cache_manager.cache_processed_excel(filename, entity_name, entity_suffixes, markdown_content)
            print(f"📋 Cached result for {filename} (path-based)")
        
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
        print(f"🔄 Using fallback mode for {len(keys)} keys")
        return generate_test_results(keys)
    
    print(f"🚀 Starting AI processing for {len(keys)} keys")
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
    
    # Fix tqdm progress bar to show proper total
    pbar = tqdm(keys, desc="Processing keys", unit="key", total=len(keys))
    
    for key_index, key in enumerate(pbar):
        # Update progress description to show current key and progress
        pbar.set_description(f"Processing {key}")
        
        config_details = load_config(config_file)
        
        # Try to initialize AI services with proper error handling
        try:
            oai_client, search_client = initialize_ai_services(config_details)
            openai_model = config_details['CHAT_MODEL']
        except RuntimeError as e:
            # AI services not available, return test results
            print(f"AI services not available: {e}")
            return generate_test_results(keys)
        
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
        
        # Print AI response [:20] for cmd checking (issue #2)
        print(f"📝 {key}: {response_txt[:20]}...")
        
        # Update progress bar with key information
        pbar.set_postfix_str(f"key={key} ({key_index + 1}/{len(keys)})")
    
    pbar.close()
    print(f"✅ Completed processing {len(keys)} keys")
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

# --- Data Validation Agent ---
class DataValidationAgent:
    def __init__(self):
        self.config_file = 'utils/config.json'
        self.financial_figure_map = {
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
    
    def validate_financial_data(self, content: str, excel_file: str, entity: str, key: str) -> Dict:
        """Validate that financial figures in content match the Excel data"""
        try:
            # Extract financial figures from Excel
            financial_figures = find_financial_figures_with_context_check(
                excel_file, 
                get_tab_name(entity), 
                '30/09/2022'
            )
            expected_figure = financial_figures.get(key)
            
            if not AI_AVAILABLE:
                return self._fallback_data_validation(content, expected_figure, key)
            
            # Use AI to validate data accuracy
            config_details = load_config(self.config_file)
            oai_client, _ = initialize_ai_services(config_details)
            
            system_prompt = """
            You are a financial data validation specialist. Your task is to verify that financial figures 
            mentioned in the content match the expected values from the financial statements.
            
            CRITICAL REQUIREMENTS:
            1. Extract all financial figures from the content
            2. Compare them with the expected figure from financial statements
            3. Check for accuracy, proper formatting (K/M), and consistency
            4. Identify any discrepancies or missing data
            5. Return a structured validation result
            6. Be very specific about what needs to be corrected
            """
            
            user_query = f"""
            VALIDATE FINANCIAL DATA ACCURACY:
            
            CONTENT TO VALIDATE: {content}
            EXPECTED FIGURE FOR {key}: {expected_figure}
            FINANCIAL STATEMENT DATA: {financial_figures}
            
            TASKS:
            1. Extract all financial figures from the content
            2. Compare with expected figure from financial statements
            3. Check if figures are properly formatted (K/M notation)
            4. Verify entity names match the data source
            5. Identify any discrepancies
            6. Check for placeholder values (xxx, [amount], etc.)
            
            RETURN FORMAT (JSON):
            {{
                "needs_correction": true/false,
                "issues": ["list of specific issues found"],
                "score": 0-100,
                "extracted_figures": ["list of figures found in content"],
                "expected_figure": "expected value",
                "discrepancies": ["list of discrepancies"],
                "suggestions": ["list of specific correction suggestions"]
            }}
            """
            
            response = generate_response(user_query, system_prompt, oai_client, "", config_details['CHAT_MODEL'])
            
            # Parse AI response
            try:
                result = json.loads(response)
                # Ensure all required fields are present
                result.setdefault('needs_correction', False)
                result.setdefault('issues', [])
                result.setdefault('score', 100)
                result.setdefault('suggestions', [])
                return result
            except Exception as parse_error:
                print(f"Failed to parse AI response: {parse_error}")
                return self._fallback_data_validation(content, expected_figure, key)
                
        except Exception as e:
            print(f"Data validation error: {e}")
            return {"needs_correction": False, "issues": [f"Validation error: {e}"], "score": 50, "suggestions": []}
    
    def _fallback_data_validation(self, content: str, expected_figure: float, key: str) -> Dict:
        """Fallback validation when AI is not available"""
        issues = []
        score = 100
        
        # Check if expected figure is mentioned
        if expected_figure:
            expected_str = str(expected_figure)
            if expected_str not in content and str(int(expected_figure)) not in content:
                issues.append(f"Expected figure {expected_figure} not found in content")
                score -= 30
        
        # Check for placeholder values
        if 'xxx' in content.lower() or 'placeholder' in content.lower():
            issues.append("Placeholder values found in content")
            score -= 20
        
        return {
            "needs_correction": score < 80,
            "issues": issues,
            "score": score,
            "suggestions": [],
            "extracted_figures": [],
            "expected_figure": expected_figure,
            "discrepancies": []
        }
    
    def correct_financial_data(self, content: str, issues: List[str]) -> str:
        """Correct financial data issues using AI"""
        try:
            if not AI_AVAILABLE:
                return content
            
            config_details = load_config(self.config_file)
            oai_client, _ = initialize_ai_services(config_details)
            
            system_prompt = """
            You are a financial data correction specialist. Your task is to fix financial data 
            accuracy issues in the content while maintaining the original structure and tone.
            
            REQUIREMENTS:
            1. Fix all identified data accuracy issues
            2. Ensure figures match financial statements exactly
            3. Maintain proper formatting (K/M notation)
            4. Keep the original writing style and structure
            5. Only correct the identified issues, don't rewrite unnecessarily
            """
            
            user_query = f"""
            CORRECT FINANCIAL DATA ISSUES:
            
            ORIGINAL CONTENT: {content}
            IDENTIFIED ISSUES: {issues}
            
            TASK: Fix the identified issues while maintaining the original content structure.
            
            REQUIREMENTS:
            - Fix all data accuracy issues
            - Ensure proper figure formatting
            - Maintain original writing style
            - Keep the same paragraph structure
            - Only correct what needs fixing
            
            RETURN: Only the corrected content text, no explanations or JSON.
            """
            
            corrected_content = generate_response(user_query, system_prompt, oai_client, "", config_details['CHAT_MODEL'])
            return corrected_content.strip()
            
        except Exception as e:
            print(f"Data correction error: {e}")
            return content

# --- Pattern Validation Agent ---
class PatternValidationAgent:
    def __init__(self):
        self.config_file = 'utils/config.json'
        self.pattern_file = 'utils/pattern.json'
    
    def validate_pattern_compliance(self, content: str, key: str) -> Dict:
        """Validate that content follows the expected pattern structure"""
        try:
            # Load patterns for the key
            patterns = load_ip(self.pattern_file, key)
            
            if not AI_AVAILABLE:
                return self._fallback_pattern_validation(content, patterns, key)
            
            # Use AI to validate pattern compliance
            config_details = load_config(self.config_file)
            oai_client, _ = initialize_ai_services(config_details)
            
            system_prompt = """
            You are a pattern compliance validation specialist. Your task is to verify that the content 
            follows the expected pattern structure and format.
            
            REQUIREMENTS:
            1. Compare content against available patterns
            2. Check for proper structure and formatting
            3. Verify all placeholders are filled
            4. Ensure professional tone and language
            5. Identify any pattern compliance issues
            6. Provide specific suggestions for improvement
            """
            
            user_query = f"""
            VALIDATE PATTERN COMPLIANCE:
            
            CONTENT TO VALIDATE: {content}
            KEY: {key}
            AVAILABLE PATTERNS: {json.dumps(patterns, indent=2)}
            
            TASKS:
            1. Check if content follows any of the available patterns
            2. Verify all placeholders are properly filled
            3. Check for proper financial terminology
            4. Ensure professional writing style
            5. Identify any structural or formatting issues
            6. Check for template artifacts (xxx, [placeholder], etc.)
            
            RETURN FORMAT (JSON):
            {{
                "needs_correction": true/false,
                "issues": ["list of pattern compliance issues"],
                "score": 0-100,
                "pattern_match": "name of matched pattern or 'none'",
                "missing_elements": ["list of missing pattern elements"],
                "suggestions": ["list of specific improvement suggestions"]
            }}
            """
            
            response = generate_response(user_query, system_prompt, oai_client, "", config_details['CHAT_MODEL'])
            
            # Parse AI response
            try:
                result = json.loads(response)
                # Ensure all required fields are present
                result.setdefault('needs_correction', False)
                result.setdefault('issues', [])
                result.setdefault('score', 100)
                result.setdefault('pattern_match', 'none')
                result.setdefault('missing_elements', [])
                result.setdefault('suggestions', [])
                return result
            except Exception as parse_error:
                print(f"Failed to parse AI response: {parse_error}")
                return self._fallback_pattern_validation(content, patterns, key)
                
        except Exception as e:
            print(f"Pattern validation error: {e}")
            return {"needs_correction": False, "issues": [f"Validation error: {e}"], "score": 50, "suggestions": []}
    
    def _fallback_pattern_validation(self, content: str, patterns: Dict, key: str) -> Dict:
        """Fallback validation when AI is not available"""
        issues = []
        score = 100
        
        # Check for placeholder values
        if 'xxx' in content.lower() or '{' in content or '}' in content:
            issues.append("Unfilled placeholders found in content")
            score -= 25
        
        # Check for pattern structure
        if not any(pattern.lower() in content.lower() for pattern in patterns.values()):
            issues.append("Content doesn't match expected pattern structure")
            score -= 20
        
        # Check for professional language
        professional_terms = ['represented', 'comprised', 'indicated', 'demonstrated']
        if not any(term in content.lower() for term in professional_terms):
            issues.append("Missing professional financial language")
            score -= 15
        
        return {
            "needs_correction": score < 80,
            "issues": issues,
            "score": score,
            "pattern_match": "none",
            "missing_elements": [],
            "suggestions": []
        }
    
    def correct_pattern_compliance(self, content: str, issues: List[str]) -> str:
        """Correct pattern compliance issues using AI"""
        try:
            if not AI_AVAILABLE:
                return content
            
            config_details = load_config(self.config_file)
            oai_client, _ = initialize_ai_services(config_details)
            
            system_prompt = """
            You are a pattern compliance correction specialist. Your task is to fix pattern 
            compliance issues in the content while maintaining accuracy and professionalism.
            
            REQUIREMENTS:
            1. Fix all identified pattern compliance issues
            2. Ensure content follows expected pattern structure
            3. Fill any missing placeholders appropriately
            4. Maintain professional financial language
            5. Keep the original meaning and accuracy
            """
            
            user_query = f"""
            CORRECT PATTERN COMPLIANCE ISSUES:
            
            ORIGINAL CONTENT: {content}
            IDENTIFIED ISSUES: {issues}
            
            TASK: Fix the identified pattern compliance issues while maintaining content accuracy.
            
            REQUIREMENTS:
            - Fix all pattern compliance issues
            - Ensure proper pattern structure
            - Fill missing placeholders appropriately
            - Maintain professional language
            - Keep original meaning intact
            
            RETURN: Only the corrected content text, no explanations or JSON.
            """
            
            corrected_content = generate_response(user_query, system_prompt, oai_client, "", config_details['CHAT_MODEL'])
            return corrected_content.strip()
            
        except Exception as e:
            print(f"Pattern correction error: {e}")
            return content 
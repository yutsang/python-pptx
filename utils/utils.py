import pandas as pd
from tabulate import tabulate
import json
from common.assistant import load_config, initialize_ai_services, generate_response
import warnings
import re, os, urllib3
from tqdm import tqdm
from pathlib import Path
from utils.cache import get_cache_manager, cached_function

urllib3.disable_warnings()
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action='ignore', category=UserWarning)
warnings.filterwarnings('ignore', message='Data Validation extension is not supported and will be removed', category=UserWarning, module='openpyxl')

def process_and_filter_excel(filename, tab_name_mapping, entity_name, entity_suffixes):
    """Process and filter Excel file with simple caching"""
    
    try:
        # Use simple cache instead of complex cache manager
        from utils.simple_cache import get_simple_cache
        cache = get_simple_cache()
        
        # Check cache first
        cached_result = cache.get_cached_excel_data(filename, entity_name)
        if cached_result is not None:
            return cached_result
            
        # Load the Excel file
        main_dir = Path(__file__).parent.parent
        file_path = main_dir / filename
        # Use context manager to ensure file handle is released on Windows
        with pd.ExcelFile(file_path) as xl:
            # Create a reverse mapping from values to keys
            reverse_mapping = {}
            for key, values in tab_name_mapping.items():
                for value in values:
                    reverse_mapping[value] = key
            
            # Initialize a string to store markdown content
            markdown_content = ""
            
            # Process each sheet according to the mapping
            for sheet_name in xl.sheet_names:
                if sheet_name in reverse_mapping:
                    df = xl.parse(sheet_name)
                    
                    # Split dataframes on empty rows
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
                    
                    # Filter dataframes by entity name
                    entity_keywords = [f"{entity_name}{suffix}" for suffix in entity_suffixes]
                    combined_pattern = '|'.join(re.escape(kw) for kw in entity_keywords)
                    
                    for data_frame in dataframes:
                        mask = data_frame.apply(
                            lambda row: row.astype(str).str.contains(
                                combined_pattern, case=False, regex=True, na=False
                            ).any(),
                            axis=1
                        )
                        if mask.any():
                            markdown_content += tabulate(data_frame, headers='keys', tablefmt='pipe') + '\n\n'
                        
                        if any(
                            data_frame.apply(
                                lambda row: row.astype(str).str.contains(keyword, case=False, na=False).any(),
                                axis=1
                            ).any() for keyword in entity_keywords
                        ):
                            markdown_content += tabulate(
                                data_frame, headers='keys', tablefmt='pipe', showindex=False
                            )
                            markdown_content += "\n\n"
        
        # Cache the processed result using simple cache
        cache.cache_excel_data(filename, entity_name, markdown_content)
        return markdown_content
    except Exception as e:
        print("An error occurred while processing the Excel file:", e)
    return ""

def load_ip(file, key):
    try:
        with open(file, 'r') as f:
            data = json.load(f)
        if key in data:
            return data[key]
    except FileNotFoundError:
        print(f"File {file} not found.")
    except json.JSONDecodeError:
        print(f"Error decoding JSON from file {file}.")
    return {}

def load_system_prompt(config_details):
    # Assemble the prompt from parts in the loaded config
    prompt = f"{config_details['promptTitle']}\n{config_details['promptIntro']}\n{config_details['promptTask']}\n"
    for rule in config_details['promptFormattingRules']:
        prompt += f"{rule}\n"
    prompt += config_details['promptFooter']
    
    return prompt

def detect_string_in_file(file_content, target_string):
    try:
        # Check if target string is in the file content
        if target_string in file_content:
            return True
        else:
            return False
    except Exception as e:
        pass
    
name_mapping = {
    'Cash': 'Cash at bank',
    'AR': 'Accounts receivables',
    'Prepayments': 'Prepayments',
    'OR': 'Other receivables',
    'Other CA': 'Other current assets',
    'IP': 'Investment properties',
    'Other NCA': 'Other non-current assets',
    'AP': 'Accounts payable',
    'Taxes payable': 'Taxes payables',
    'OP': 'Other payables',
    'Capital': 'Capital',
    'Reserve': 'Surplus reserve'
}

def find_financial_figures_with_context_check(filename, sheet_name, date_str):
    try:
        # Load the Excel file
        file_path = Path(filename)
        # Use context manager to ensure file handle is released
        with pd.ExcelFile(file_path) as xl:
            if sheet_name not in xl.sheet_names:
                print(f"Sheet '{sheet_name}' not found in the file.")
                return {}
            
            # Parse the sheet
            df = xl.parse(sheet_name)

            # Rename columns if seeing usual headers following context title
            df.columns = ['Description', 'Date_2020', 'Date_2021', 'Date_2022']
            
            # Filter for the specific date column
            date_column_map = {
                '31/12/2020': 'Date_2020',
                '31/12/2021': 'Date_2021',
                '30/09/2022': 'Date_2022'
            }
            
            if date_str not in date_column_map:
                print(f"Date '{date_str}' not recognized.")
                return {}
            
            date_column = date_column_map[date_str]
            
            # Scale factor based on the '000 notation
            scale_factor = 1000
            
            # Extract financial figures now if starts early perhaps descriptive context
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
            
            # Iterate only effectively without further row entity considerations
            for key, desc in financial_figure_map.items():
                value = df.loc[df['Description'].str.contains(desc, case=False, na=False), date_column].values
                if value.size > 0:
                    financial_figures[key] = float(value[0]) / scale_factor
            return financial_figures
    
    except Exception as e:
        print(f"An error occurred while processing the Excel file: {e}")
    return {}

# Example usage
excel_file_name = 'utils/221128.Project TK.Databook.JW_enhanced.xlsx'
project_name = 'Nanjing'
date = '30/09/2022'

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
    
# Main Function 
def process_keys(keys, entity_name, entity_helpers, input_file, mapping_file, pattern_file, config_file='utils/config.json'):

    financial_figures = find_financial_figures_with_context_check(input_file, get_tab_name(entity_name), date)
    
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
    
    # Dictionary to store responses for each key
    results = {}
    
    # Initialize tqdm progress bar
    pbar = tqdm(keys, desc="Processing keys", unit="key")
    
    # Iterate through each key and process them
    for key in pbar:
        
        # Load configuration and initialize services
        config_details = load_config(config_file)
        oai_client, _ = initialize_ai_services(config_details)
        
        # Load model details from config
        openai_model = config_details['DEEPSEEK_CHAT_MODEL']
        
        # Load pattern and mapping
        pattern = load_ip(pattern_file, key)
        mapping = {key: load_ip(mapping_file, key)}
        
        excel_tables = process_and_filter_excel(input_file, mapping, entity_name, entity_helpers)
        detect_zeros = "3. The figures in this table is already expressed in k, express the number in M " \
            "(divide by 1000), rounded to 1 decimal place, if final figure less than 1M, express in K (no decimal places)." if detect_string_in_file(excel_tables, "'000") else ""
            
        # User query construction using template and workbook details
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
        
        # Generate response using AI services
        # Use Excel tables as context content
        context_content = excel_tables
        response_txt = generate_response(user_query, system_prompt, oai_client, context_content, openai_model, entity_name)
        
        # Store the result in the dictionary
        results[key] = response_txt
        
        # Update progress bar description
        pbar.set_postfix_str(f"key={key}")
    
    return results
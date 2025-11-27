import pandas as pd
from tabulate import tabulate
import json
from common.assistant import load_config, initialize_ai_services, generate_response
import warnings
import re, os, urllib3
from tqdm import tqdm
from pathlib import Path


urllib3.disable_warnings()
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action='ignore', category=UserWarning)
warnings.filterwarnings('ignore', message='Data Validation extension is not supported and will be removed', category=UserWarning, module='openpyxl')

def process_and_filter_excel(filename, tab_name_mapping, entity_name, entity_suffixes):
    """Process and filter Excel file"""
    
    try:
            
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
                # Also map the key name directly to itself (for sheet names like "Cash", "AR")
                reverse_mapping[key] = key
            
            # Initialize a string to store markdown content
            markdown_content = ""
            
                        # Process each sheet according to the mapping
            for sheet_name in xl.sheet_names:
                if sheet_name in reverse_mapping:
                    df = xl.parse(sheet_name)
                    
                    # Detect latest date column for this sheet
                    latest_date_col = detect_latest_date_column(df)
                    if latest_date_col:
                        print(f"📅 Sheet {sheet_name}: Using latest date column {latest_date_col}")
                    
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
                            # Filter dataframe to show only description column and latest date column
                            filtered_df = data_frame.copy()
                            if latest_date_col and latest_date_col in filtered_df.columns:
                                # Keep only description column (first) and latest date column
                                desc_col = filtered_df.columns[0]
                                cols_to_keep = [desc_col, latest_date_col]
                                filtered_df = filtered_df[cols_to_keep]
                                print(f"📊 Sheet {sheet_name}: Filtered to show {desc_col} and {latest_date_col}")
                            
                            markdown_content += tabulate(filtered_df, headers='keys', tablefmt='pipe') + '\n\n'
                        
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

def detect_latest_date_column(df):
    """Detect the latest date column from a DataFrame, including xMxx format dates."""
    import re
    from datetime import datetime
    
    def parse_date(date_str):
        """Parse date string in various formats including xMxx."""
        if not date_str or pd.isna(date_str):
            return None
        
        date_str = str(date_str).strip()
        
        # Handle xMxx format (e.g., 9M22, 12M23) - END OF MONTH
        xmxx_match = re.match(r'^(\d+)M(\d{2})$', date_str)
        if xmxx_match:
            month = int(xmxx_match.group(1))
            year = 2000 + int(xmxx_match.group(2))  # Assume 20xx for 2-digit years
            # Use end of month, not beginning (last day of the month)
            if month == 12:
                return datetime(year, 12, 31)  # December 31st
            elif month in [1, 3, 5, 7, 8, 10]:
                return datetime(year, month, 31)  # 31-day months
            elif month in [4, 6, 9, 11]:
                return datetime(year, month, 30)  # 30-day months
            elif month == 2:
                # February - handle leap years
                if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0):
                    return datetime(year, 2, 29)  # Leap year
                else:
                    return datetime(year, 2, 28)  # Non-leap year
        
        # Handle standard date formats
        date_formats = [
            '%d/%m/%Y', '%d-%m-%Y', '%d/%m/%y', '%d-%m-%y',
            '%Y-%m-%d', '%m/%d/%Y', '%m-%d-%Y',
            '%d/%b/%Y', '%d-%b-%Y', '%b/%d/%Y', '%b-%d-%Y',
            '%d/%B/%Y', '%d-%B-%Y', '%B/%d/%Y', '%B-%d-%Y',
            # Chinese date formats
            '%Y年%m月%d日', '%Y年%m月', '%m月%d日', '%Y/%m/%d',
            '%Y.%m.%d', '%Y年%m月%d号',
            # Additional flexible formats
            '%Y%m%d', '%d%m%Y', '%m%d%Y'
        ]
        
        for fmt in date_formats:
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                continue
        
        return None
    
    # Get column names
    columns = df.columns.tolist()
    latest_date = None
    latest_column = None
    
    # First, try to find dates in column names
    for col in columns:
        col_str = str(col)
        parsed_date = parse_date(col_str)
        if parsed_date and (latest_date is None or parsed_date > latest_date):
            latest_date = parsed_date
            latest_column = col
    
    # If no dates found in column names, check the first few rows for datetime values
    if latest_column is None and len(df) > 0:
        # Check first 5 rows for date values (dates can be in different rows)
        for row_idx in range(min(5, len(df))):
            row = df.iloc[row_idx]
            for col in columns:
                val = row[col]
                
                # Check if it's already a datetime object
                if isinstance(val, (pd.Timestamp, datetime)):
                    date_val = val if isinstance(val, datetime) else val.to_pydatetime()
                    if latest_date is None or date_val > latest_date:
                        latest_date = date_val
                        latest_column = col
                # Check if it's a string that can be parsed as a date
                elif pd.notna(val):
                    parsed_date = parse_date(str(val))
                    if parsed_date and (latest_date is None or parsed_date > latest_date):
                        latest_date = parsed_date
                        latest_column = col
    
    return latest_column

def find_financial_figures_with_context_check(filename, sheet_name, date_str):
    try:
        # Load the Excel file
        file_path = Path(filename)
        # Use context manager to ensure file handle is released
        with pd.ExcelFile(file_path) as xl:
            # Handle both single sheet name and list of possible names
            if isinstance(sheet_name, list):
                found_sheet = None
                for possible_sheet in sheet_name:
                    if possible_sheet in xl.sheet_names:
                        found_sheet = possible_sheet
                        break

                if found_sheet is None:
                    print(f"None of the sheet names {sheet_name} found in the file. Available sheets: {xl.sheet_names}")
                    return {}
                sheet_name = found_sheet
                print(f"Found matching sheet: '{sheet_name}'")
            elif sheet_name not in xl.sheet_names:
                print(f"Sheet '{sheet_name}' not found in the file. Available sheets: {xl.sheet_names}")
                return {}

            # Parse the sheet
            df = xl.parse(sheet_name)

            # Detect latest date column automatically
            latest_date_col = detect_latest_date_column(df)
            if latest_date_col:
                # Use the latest date column instead of the requested date
                date_column = latest_date_col
                print(f"Using latest date column: {latest_date_col}")
            else:
                # Fallback to original logic only if date_str is provided
                if date_str is not None:
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
                else:
                    print("No date column detected and no fallback date provided.")
                    return {}
            
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
            
            # Find the description column (usually the first column)
            desc_column = None
            if 'Description' in df.columns:
                desc_column = 'Description'
            elif len(df.columns) > 0:
                desc_column = df.columns[0]  # Use first column as description column
            
            # Iterate only effectively without further row entity considerations
            for key, desc in financial_figure_map.items():
                if desc_column:
                    value = df.loc[df[desc_column].str.contains(desc, case=False, na=False), date_column].values
                    if value.size > 0:
                        financial_figures[key] = float(value[0]) / scale_factor
            return financial_figures
    
    except FileNotFoundError:
        print(f"❌ Excel file not found: {filename}")
        print(f"   Expected path: {file_path}")
        return {}
    except PermissionError:
        print(f"❌ Permission denied accessing Excel file: {filename}")
        print(f"   Check file permissions: {file_path}")
        return {}
    except Exception as e:
        print(f"❌ Unexpected error processing Excel file '{filename}': {e}")
        print(f"   File path: {file_path}")
        print(f"   Sheet name: {sheet_name}")
        return {}

# Example usage
excel_file_name = 'fdd_utils/221128.Project TK.Databook.JW_enhanced.xlsx'
project_name = 'Nanjing'
date = '30/09/2022'

def get_tab_name(project_name):
    if not project_name:
        return None

    project_name = project_name.strip()

    # Hardcoded mappings for known entities
    if project_name.lower() == 'haining':
        return "BSHN"
    elif project_name.lower() == 'nanjing':
        return "BSNJ"
    elif project_name.lower() == 'ningbo':
        return "BSNB"

    # For other entities, try to extract a meaningful sheet name
    # Remove common suffixes and use first word
    clean_name = project_name.split()[0] if project_name else None
    if clean_name:
        # Try different sheet name patterns
        possible_names = [
            f"BS{clean_name.upper()[:3]}",  # BSCLE, BSHAI, etc.
            f"{clean_name.upper()[:3]}",     # CLE, HAI, etc.
            clean_name.upper(),             # CLEANTECH, HAINING, etc.
            f"BS_{clean_name.upper()[:3]}",  # BS_CLE, etc.
            project_name                     # Original name as last resort
        ]
        return possible_names  # Return list of possible names

    # Fallback: return the project name itself to avoid None
    print(f"Warning: Could not extract sheet name from '{project_name}', using project name as fallback")
    return [project_name]
    
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
def process_keys(keys, entity_name, entity_helpers, input_file, mapping_file, pattern_file, config_file='fdd_utils/config.json'):

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
        mapping = load_ip(mapping_file)
        
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
        - If any foreign language content is encountered, translate it directly to English without using translation placeholders
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
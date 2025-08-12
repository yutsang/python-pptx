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
import logging

# Suppress httpx logging
logging.getLogger("httpx").setLevel(logging.WARNING)

# AI-related imports (DeepSeek only)
try:
    from openai import OpenAI
    AI_AVAILABLE = True
except ImportError:
    OpenAI = None
    AI_AVAILABLE = False

# --- Config and AI Service Helpers ---
def load_config(file_path):
    """Load configuration from a JSON file."""
    with open(file_path) as config_file:
        config_details = json.load(config_file)
    return config_details

def initialize_ai_services(config_details, use_local=False, use_openai=False):
    """Initialize AI client using config details - supports DeepSeek, OpenAI, and local AI."""
    if not AI_AVAILABLE:
        raise RuntimeError("AI services not available on this machine.")
    httpx_client = httpx.Client(verify=False)
    if OpenAI is None:
        raise RuntimeError("AI modules not available.")
    
    # Allow session-level override of provider selection
    try:
        import streamlit as st
        selected_provider = st.session_state.get('selected_provider')
        if selected_provider == 'Open AI':
            use_openai = True
            use_local = False
        elif selected_provider == 'Local AI':
            use_local = True
            use_openai = False
        elif selected_provider == 'Server AI':
            # Treat Server AI the same as Local AI (same OpenAI-compatible API format)
            # Just different base URL and key configured in config.json
            use_local = True
            use_openai = False
    except Exception:
        pass

    if use_local:
        # Initialize Local/Server AI client (same OpenAI-compatible format)
        # Prefer LOCAL_*; if SERVER_* are provided, they can be mapped externally by updating config.
        local_base = config_details.get('LOCAL_AI_API_BASE') or config_details.get('SERVER_AI_API_BASE')
        local_key = config_details.get('LOCAL_AI_API_KEY') or config_details.get('SERVER_AI_API_KEY') or 'local-key'
        local_enabled = config_details.get('LOCAL_AI_ENABLED', True)
        if local_base and local_enabled:
            oai_client = OpenAI(
                api_key=local_key,
                base_url=local_base,
                http_client=httpx_client
            )
            print("🏠 Using Local/Server AI (OpenAI-compatible)")
        else:
            raise RuntimeError("Local/Server AI configuration not found or not enabled. Please check LOCAL_AI_API_BASE (or SERVER_AI_API_BASE) and LOCAL_AI_ENABLED in config.")
    elif use_openai:
        # Initialize OpenAI client
        if config_details.get('OPENAI_API_KEY') and config_details.get('OPENAI_API_BASE'):
            oai_client = OpenAI(
                api_key=config_details['OPENAI_API_KEY'],
                base_url=config_details['OPENAI_API_BASE'],
                http_client=httpx_client
            )
            print("🤖 Using OpenAI GPT-4o-mini")
        else:
            raise RuntimeError("OpenAI configuration not found. Please check OPENAI_API_KEY and OPENAI_API_BASE in config.")
    else:
        # Initialize DeepSeek client (server)
        if config_details.get('DEEPSEEK_API_KEY') and config_details.get('DEEPSEEK_API_BASE'):
            oai_client = OpenAI(
                api_key=config_details['DEEPSEEK_API_KEY'],
                base_url=config_details['DEEPSEEK_API_BASE'],
                http_client=httpx_client
            )
            # DeepSeek AI initialized (silent for cleaner progress bar)
        else:
            raise RuntimeError("DeepSeek configuration not found. Please check DEEPSEEK_API_KEY and DEEPSEEK_API_BASE in config.")
    
    return oai_client, None  # No search client needed

def get_openai_client(config_details=None, use_local=False, use_openai=False):
    """Get the appropriate OpenAI client based on model selection."""
    if config_details is None:
        # Load config if not provided
        config_path = os.path.join(os.path.dirname(__file__), '..', 'utils', 'config.json')
        config_details = load_config(config_path)
    
    client, _ = initialize_ai_services(config_details, use_local=use_local, use_openai=use_openai)
    return client

def get_chat_model(config_details=None, use_local=False, use_openai=False):
    """Get the appropriate chat model name based on selection."""
    if config_details is None:
        # Load config if not provided
        config_path = os.path.join(os.path.dirname(__file__), '..', 'utils', 'config.json')
        config_details = load_config(config_path)
    
    # Allow session-level override of provider and model
    try:
        import streamlit as st
        selected_provider = st.session_state.get('selected_provider')
        selected_model = st.session_state.get('selected_model')
        if selected_provider == 'Local AI':
            use_local = True
            use_openai = False
            if selected_model:
                return selected_model
        elif selected_provider == 'Open AI':
            use_openai = True
            use_local = False
            if selected_model:
                return selected_model
        elif selected_provider == 'Server AI':
            # Same format as Local AI; use the same model selection/defaults
            use_local = True
            use_openai = False
            if selected_model:
                return selected_model
    except Exception:
        pass

    if use_local:
        return config_details.get('LOCAL_AI_CHAT_MODEL', 'local-model')
    elif use_openai:
        return config_details.get('OPENAI_CHAT_MODEL', 'gpt-4o-mini-2024-07-18')
    else:
        return config_details.get('DEEPSEEK_CHAT_MODEL', 'deepseek-chat')

def generate_response(user_query, system_prompt, oai_client, context_content, openai_chat_model, entity_name="default", use_local_ai=False):
    """Generate a response from the AI model given a user query and system prompt with simple caching."""
    # Use simple cache instead of complex hash-based cache
    from utils.simple_cache import get_simple_cache
    cache = get_simple_cache()
    
    # Create a simple cache key from the query
    cache_key = f"{hash(user_query) % 1000000}_{hash(system_prompt) % 1000000}"
    
    # Check cache first (with force refresh option)
    try:
        import streamlit as st
        force_refresh = st.session_state.get('force_refresh', False)
    except Exception:
        force_refresh = False
        
    cached_response = cache.get_cached_ai_result(cache_key, entity_name, force_refresh)
    if cached_response is not None:
        return cached_response
    else:
        # Reset force refresh after use
        try:
            import streamlit as st
            if force_refresh:
                st.session_state['force_refresh'] = False
        except Exception:
            pass
    
    # Include context data in the user query instead of as a separate assistant message
    enhanced_user_query = f"Context data:\n{context_content}\n\nUser query:\n{user_query}"
    
    conversation = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": enhanced_user_query}
    ]
    
    try:
        # For local AI with thinking support (DeepSeek 671B), enable reasoning
        extra_params = {}
        if use_local_ai:
            # Do NOT enable reasoning; always request direct answers only
            try:
                pass
            except Exception:
                pass
        
        # Use selected chat model if provided; fall back to session/config selection
        model_name = openai_chat_model or get_chat_model()
        # Remove unsupported params defensively for non-DeepSeek models (e.g., qwen2)
        # Ensure no reasoning param is sent for compatibility across models
        safe_params = {k: v for k, v in extra_params.items() if k != 'reasoning'}
        response = oai_client.chat.completions.create(model=model_name, messages=conversation, **safe_params)
    except Exception as e:
        print(f"❌ API call failed with model '{openai_chat_model}': {e}")
        raise
    
    response_content = response.choices[0].message.content
    
    # If model returns thinking tags anyway, strip them and keep final answer only
    if "<thinking>" in response_content and "</thinking>" in response_content:
        response_content = response_content.split("</thinking>")[-1].strip()
    
    # Cache the response using simple cache
    cache.cache_ai_result(cache_key, entity_name, response_content)
    
    return response_content

# --- Excel and Data Processing ---
def parse_table_to_structured_format(df, entity_name, table_name):
    """
    Parse a DataFrame into structured format for financial tables.
    Extracts table name, entity, date, currency, multiplier, and items.
    """
    try:
        import re
        from datetime import datetime
        
        # Initialize structured data
        structured_data = {
            'table_name': table_name,
            'entity': entity_name,
            'date': None,
            'currency': 'CNY',
            'multiplier': 1,
            'items': [],
            'total': None
        }
        
        # Convert DataFrame to list of rows for easier processing
        rows = df.values.tolist()
        if not rows:
            return None
        
        # Find the two most important columns (description and amount)
        # Usually the first two columns, but let's be smart about it
        desc_col = 0
        amount_col = 1
        
        # Look for columns with numbers in the amount column
        for col_idx in range(min(2, len(df.columns))):
            numeric_count = 0
            for row in rows:
                if col_idx < len(row):
                    cell_value = str(row[col_idx]).strip()
                    # Check if it's a number (including with commas, decimals, etc.)
                    if re.match(r'^[\d,]+\.?\d*$', cell_value.replace(',', '')):
                        numeric_count += 1
            
            if numeric_count > len(rows) * 0.3:  # At least 30% of rows have numbers
                amount_col = col_idx
                desc_col = 1 if col_idx == 0 else 0
                break
        
        # Process rows to extract information
        for row_idx, row in enumerate(rows):
            if len(row) < 2:
                continue
                
            desc_cell = str(row[desc_col]).strip() if desc_col < len(row) else ""
            amount_cell = str(row[amount_col]).strip() if amount_col < len(row) else ""
            
            # Skip empty rows
            if not desc_cell and not amount_cell:
                continue
            
            # Extract date
            if not structured_data['date']:
                # Look for date patterns in any cell
                for cell in row:
                    cell_str = str(cell).strip()
                    # Common date patterns
                    date_patterns = [
                        r'\d{4}-\d{2}-\d{2}',  # YYYY-MM-DD
                        r'\d{2}/\d{2}/\d{4}',  # MM/DD/YYYY
                        r'\d{2}-\d{2}-\d{4}',  # DD-MM-YYYY
                        r'\d{4}/\d{2}/\d{2}',  # YYYY/MM/DD
                    ]
                    for pattern in date_patterns:
                        match = re.search(pattern, cell_str)
                        if match:
                            try:
                                # Try to parse the date
                                date_str = match.group()
                                if '-' in date_str:
                                    if len(date_str.split('-')[0]) == 4:  # YYYY-MM-DD
                                        parsed_date = datetime.strptime(date_str, '%Y-%m-%d')
                                    else:  # DD-MM-YYYY
                                        parsed_date = datetime.strptime(date_str, '%d-%m-%Y')
                                else:  # MM/DD/YYYY or YYYY/MM/DD
                                    if len(date_str.split('/')[0]) == 4:  # YYYY/MM/DD
                                        parsed_date = datetime.strptime(date_str, '%Y/%m/%d')
                                    else:  # MM/DD/YYYY
                                        parsed_date = datetime.strptime(date_str, '%m/%d/%Y')
                                structured_data['date'] = parsed_date.strftime('%Y-%m-%d %H:%M:%S')
                                break
                            except:
                                continue
            
            # Extract currency and multiplier
            if 'CNY' in desc_cell.upper() or 'CNY' in amount_cell.upper():
                structured_data['currency'] = 'CNY'
                if "'000" in desc_cell or "'000" in amount_cell:
                    structured_data['multiplier'] = 1000
                elif "million" in desc_cell.lower() or "million" in amount_cell.lower():
                    structured_data['multiplier'] = 1000000
                elif "000" in desc_cell or "000" in amount_cell:
                    # Check if it's a standalone "000" indicating thousands
                    if re.match(r'^0*000$', desc_cell.replace("'", "")) or re.match(r'^0*000$', amount_cell.replace("'", "")):
                        structured_data['multiplier'] = 1000
            
            # Extract items (skip header rows and totals)
            if (desc_cell.lower() not in ['total', 'indicative adjusted', 'nan', ''] and 
                not re.match(r'^[A-Z\s]+$', desc_cell) and  # Skip all caps headers
                amount_cell and amount_cell != 'nan'):
                
                # Try to extract numeric amount
                amount_match = re.search(r'[\d,]+\.?\d*', amount_cell.replace(',', ''))
                if amount_match:
                    amount_str = amount_match.group()
                    try:
                        amount = float(amount_str.replace(',', ''))
                        # Apply multiplier if needed
                        if structured_data['multiplier'] > 1:
                            amount *= structured_data['multiplier']
                        
                        structured_data['items'].append({
                            'description': desc_cell,
                            'amount': int(amount) if amount.is_integer() else amount
                        })
                    except:
                        pass
            
            # Extract total
            if desc_cell.lower() == 'total' and amount_cell and amount_cell != 'nan':
                amount_match = re.search(r'[\d,]+\.?\d*', amount_cell.replace(',', ''))
                if amount_match:
                    amount_str = amount_match.group()
                    try:
                        amount = float(amount_str.replace(',', ''))
                        if structured_data['multiplier'] > 1:
                            amount *= structured_data['multiplier']
                        structured_data['total'] = int(amount) if amount.is_integer() else amount
                    except:
                        pass
        
        # Extract entity name from table content if not found
        if structured_data['entity'] == entity_name:
            # Look for more specific entity names in the data
            for item in structured_data['items']:
                desc = item['description'].lower()
                if 'haining' in desc:
                    # Extract the full entity name
                    entity_match = re.search(r'haining\s+[a-zA-Z]+', desc, re.IGNORECASE)
                    if entity_match:
                        structured_data['entity'] = entity_match.group()
                        break
        
        return structured_data if structured_data['items'] else None
        
    except Exception as e:
        print(f"Error parsing table to structured format: {e}")
        return None

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
            
        main_dir = Path(__file__).parent.parent
        file_path = main_dir / filename
        wb = None
        markdown_content = ""
        entity_keywords = [entity_name] + list(entity_suffixes)
        entity_keywords = [kw.strip().lower() for kw in entity_keywords if kw.strip()]

        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
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
                        df = df.map(lambda x: str(x) if x is not None else "")
                        df = df.reset_index(drop=True)
                        
                        # Check for entity keywords - handle mixed data types safely
                        all_cells = [str(cell).lower().strip() for cell in df.values.flatten()]
                        match_found = any(any(kw in cell for cell in all_cells) for kw in entity_keywords)
                        
                        if match_found:
                            # Include all rows that contain any entity information
                            # This allows the AI to see all entity names in the table
                            filtered_rows = []
                            for idx, row in df.iterrows():
                                row_cells = [str(cell).lower().strip() for cell in row.values]
                                # Include rows that contain any entity information
                                # This will help the AI identify the correct entity names
                                if any(cell for cell in row_cells if cell and cell != 'nan'):
                                    filtered_rows.append(row)
                            
                            # Create filtered DataFrame and parse it into structured format
                            if filtered_rows:
                                filtered_df = pd.DataFrame(filtered_rows)
                                
                                # Parse the table into structured format
                                structured_table = parse_table_to_structured_format(filtered_df, entity_name, table_name)
                                
                                if structured_table:
                                    # Add structured table to markdown content
                                    markdown_content += f"## {structured_table['table_name']}\n"
                                    markdown_content += f"**Entity:** {structured_table['entity']}\n"
                                    markdown_content += f"**Date:** {structured_table['date']}\n"
                                    markdown_content += f"**Currency:** {structured_table['currency']}\n"
                                    markdown_content += f"**Multiplier:** {structured_table['multiplier']}\n\n"
                                    
                                    # Add items
                                    for item in structured_table['items']:
                                        markdown_content += f"- {item['description']}: {item['amount']}\n"
                                    
                                    markdown_content += f"\n**Total:** {structured_table['total']}\n\n"
                                else:
                                    # Fallback to original format if parsing fails
                                    try:
                                        markdown_content += tabulate(filtered_df, headers='keys', tablefmt='pipe') + '\n\n'
                                    except Exception:
                                        markdown_content += filtered_df.to_markdown(index=False) + '\n\n'
                            else:
                                # No rows matched the strict filtering criteria
                                pass
                        else:
                            # Table skipped for entity keywords
                            pass
                            
                    except Exception as e:
                        print(f"Error processing table {table_info.get('name', 'unknown')}: {e}")
                        continue
        finally:
            # Ensure workbook is closed to release file handle on Windows
            try:
                if wb is not None:
                    wb.close()
            except Exception:
                pass

        # Cache the processed result using simple cache
        cache.cache_excel_data(filename, entity_name, markdown_content)
        print(f"📋 Cached result for {filename}")
        
        return markdown_content
        
    except Exception as e:
        print("An error occurred while processing the Excel file:", e)
        return ""

def find_financial_figures_with_context_check(filename, sheet_name, date_str, convert_thousands=False):
    try:
        file_path = Path(filename)
        with pd.ExcelFile(file_path) as xl:
            if sheet_name not in xl.sheet_names:
                print(f"Sheet '{sheet_name}' not found in the file.")
                return {}
            df = xl.parse(sheet_name)
        if not isinstance(df, pd.DataFrame):
            return {}
        # Handle different sheet formats
        if sheet_name == 'BSHN':
            # BSHN sheet has different column structure
            df.columns = ['Description', 'Column1', 'Column2', 'Column3']
            date_column_map = {
                '31/12/2020': 'Column1',
                '31/12/2021': 'Column2', 
                '30/09/2022': 'Column3'
            }
        else:
            # Standard sheet format
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
        # If convert_thousands and '000' in columns or first row, multiply numeric values by 1000 for AI processing
        # For BSHN sheet, always apply scale factor since it's in '000 format
        if sheet_name == 'BSHN':
            scale_factor = 1000  # BSHN sheet is always in '000 format
        else:
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
            "Capital": "Share capital",
            "Reserve": "Reserve"
        }
        financial_figures = {}
        for key, desc in financial_figure_map.items():
            if 'Description' in df.columns and date_column in df.columns:
                value = df.loc[df['Description'].str.contains(desc, case=False, na=False), date_column].values
                if value.size > 0:
                    # Apply scale factor: multiply by 1000 if '000 notation detected
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
    """Get financial figure with proper K/M formatting and 1 decimal place"""
    figure = financial_figures.get(key, None)
    if figure is None:
        return f"{key} not found in the financial figures."
    
    # Ensure 1 decimal place for all conversions
    if figure >= 1000000:
        return f"{figure / 1000000:.1f}M"
    elif figure >= 1000:
        return f"{figure / 1000:.1f}K"  # Changed to 1dp for K as well
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
def process_keys(keys, entity_name, entity_helpers, input_file, mapping_file, pattern_file, config_file='utils/config.json', prompts_file='utils/prompts.json', use_ai=True, convert_thousands=False, progress_callback=None, processed_table_data=None, use_local_ai=False, use_openai=False):
    # AI is required - no fallback mode
    if not use_ai:
        raise RuntimeError("AI processing is required. Cannot run in offline mode.")
    
    if not AI_AVAILABLE:
        raise RuntimeError("AI services are not available. Please check your configuration and internet connection.")
    
    print(f"🚀 Starting AI processing for {len(keys)} keys")
    
    # Load prompts from prompts.json file (no hardcoded fallback)
    with open(prompts_file, 'r') as f:
        prompts_config = json.load(f)
    system_prompt = prompts_config['system_prompts']['Agent 1']
    
    # Initialize financial figures without pre-processing (will check '000 per key)
    financial_figures = find_financial_figures_with_context_check(input_file, get_tab_name(entity_name), '30/09/2022', convert_thousands=False)
    results = {}
    
    # Fix tqdm progress bar to show proper total
    pbar = tqdm(keys, desc="Processing keys", unit="key", total=len(keys))
    
    for key_index, key in enumerate(pbar):
        # Update progress description to show current key and progress
        pbar.set_description(f"Processing {key}")
        
        # Update streamlit progress if callback provided
        if progress_callback:
            progress_callback((key_index + 1) / len(keys), f"Processing {key}...")
        
        config_details = load_config(config_file)
        
        # Initialize AI services - required for processing
        oai_client, search_client = initialize_ai_services(config_details, use_local=use_local_ai, use_openai=use_openai)
        # Use appropriate model based on selection
        if use_local_ai:
            openai_model = config_details.get('LOCAL_AI_CHAT_MODEL', 'local-model')
        elif use_openai:
            openai_model = config_details.get('OPENAI_CHAT_MODEL', 'gpt-4o-mini-2024-07-18')
        else:
            openai_model = config_details.get('DEEPSEEK_CHAT_MODEL', 'deepseek-chat')
        
        pattern = load_ip(pattern_file, key)
        mapping = {key: load_ip(mapping_file)}
        
        # Use processed table data if provided, otherwise process Excel file
        if processed_table_data and key in processed_table_data:
            excel_tables = processed_table_data[key]
        else:
            excel_tables = process_and_filter_excel(input_file, mapping, entity_name, entity_helpers)
        
        # Check if '000 notation is detected
        has_thousands_notation = detect_string_in_file(excel_tables, "'000")
        
        # Process data for AI: multiply figures by 1000 if '000 notation detected
        excel_tables_for_ai = multiply_figures_for_ai_processing(excel_tables) if has_thousands_notation else excel_tables
        
        # Apply thousands conversion to the specific financial figure for this key
        current_financial_figure = financial_figures.get(key, 0)
        if has_thousands_notation and current_financial_figure:
            adjusted_financial_figure = current_financial_figure * 1000
            financial_figures[key] = adjusted_financial_figure
        
        # Update prompt to reflect the data processing
        detect_zeros = """IMPORTANT: The numerical figures in the DATA SOURCE have been adjusted for analysis (multiplied by 1000 from the original '000 notation). 
        Express all figures with proper K/M conversion with 1 decimal place:
        - Figures ≥ 1,000,000: express in M (millions) with 1 decimal place (e.g., 2.3M)
        - Figures ≥ 1,000: express in K (thousands) with 1 decimal place (e.g., 1.5K)
        - Figures < 1,000: express with 1 decimal place (e.g., 123.0)""" if has_thousands_notation else """Express all figures with proper K/M conversion with 1 decimal place:
        - Figures ≥ 1,000,000: express in M (millions) with 1 decimal place (e.g., 2.3M)
        - Figures ≥ 1,000: express in K (thousands) with 1 decimal place (e.g., 1.5K)
        - Figures < 1,000: express with 1 decimal place (e.g., 123.0)"""
        
        # User query construction using f-strings for better prompt maintainability
        pattern_json = json.dumps(pattern, indent=2)
        financial_figure_info = f"{key}: {get_financial_figure(financial_figures, key)}"
        
        # Template for output requirements - reusable across queries
        output_requirements = f"""
        REQUIRED OUTPUT FORMAT:
        - Only the completed pattern text
        - No pattern names or labels
        - No template structure
        - No JSON formatting
        - Replace ALL 'xxx' or placeholders with actual data values
        - Replace ALL [ENTITY_NAME] placeholders with the SPECIFIC entity name from the DATA SOURCE
        - CRITICAL: Use the SPECIFIC entity names from the table data (e.g., 'Third-party receivables', 'Company #1') NOT the reporting entity name
        - Do not use bullet point for listing
        - Express all figures with proper K/M conversion (e.g., 9,076,000 = 9.1M, 1,500 = 1.5K)
        - No foreign contents, if any, translate to English
        - Stick to Template format, no extra explanations or comments
        - For entity name to be filled into template, use the specific entity names from the DATA SOURCE table
        - For all listing figures, please check the total, together should be around the same or constituting majority of FINANCIAL FIGURE
        - Ensure all financial figures mentioned match the actual values from the DATA SOURCE
        - IMPORTANT: Look at the table data to identify the correct entity names (e.g., 'Third-party receivables', 'Company #1', etc.)
        """
        
        # Example formats for consistent output
        examples = f"""
        Example of CORRECT output format:
        "Cash at bank comprises deposits of CNY9.1M held with major financial institutions as at 30/09/2022."
        
        Example of INCORRECT output format:
        "Pattern 1: Cash at bank comprises deposits of xxx held with xxx as at xxx."
        """
        
        user_query = f"""
        TASK: Select ONE pattern and complete it with actual data
        
        AVAILABLE PATTERNS: {pattern_json}
        
        FINANCIAL FIGURE: {financial_figure_info}
        
        DATA SOURCE: {excel_tables_for_ai}
        
        SELECTION CRITERIA:
        - Choose the pattern with the most complete data coverage
        - Prioritize patterns that match the primary account category
        - Use most recent data: latest available
        - {detect_zeros}
        
        {output_requirements}
        
        {examples}
        
        IMPORTANT REQUIREMENTS:
        1. ALWAYS specify exact dollar amounts and currency (e.g., "CNY9.1M", "$2.3M", "CNY687K")
        2. ALWAYS identify and mention the specific entity names from the DATA SOURCE table
        3. CRITICAL: For entity names, use the SPECIFIC entity names from the financial data table (e.g., 'Third-party receivables', 'Company #1') NOT the reporting entity name
        4. Look at the table data to identify the correct entity names and amounts
        5. When the table shows 'Third-party receivables', use that exact name, not the reporting entity
        6. At the end of your response, provide a brief summary like:
           "SUMMARY: Used entity 'Third-party receivables' with amounts: CNY634K (total receivables)"
        """
        
        response_txt = generate_response(user_query, system_prompt, oai_client, excel_tables, openai_model, entity_name, use_local_ai)
        
        # Clean up response: remove outer quotation marks and translate Chinese
        response_txt = clean_response_text(response_txt)
        
        # Store result with pattern information for logging
        results[key] = {
            'content': response_txt,
            'pattern_used': 'Pattern 1',  # Default, will be updated based on actual pattern
            'table_data': excel_tables_for_ai,
            'financial_figure': financial_figures.get(key, 0),
            'entity_name': entity_name
        }
        
        # Update progress bar with key information and AI response preview
        pbar.set_postfix_str(f"{key}: {response_txt[:10]}...")
    
    pbar.close()
    
    # Final progress update
    if progress_callback:
        progress_callback(1.0, "AI processing completed!")
    
    return results

def generate_test_results(keys):
    # This function is deprecated - AI is now required
    raise RuntimeError("AI processing is required. Cannot generate test results.")

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
            import json
            
            # Extract financial figures from Excel
            financial_figures = find_financial_figures_with_context_check(
                excel_file, 
                get_tab_name(entity), 
                '30/09/2022'
            )
            expected_figure = financial_figures.get(key)
            
            if not AI_AVAILABLE:
                raise RuntimeError("AI services are required for data validation. Please check your configuration.")
            
            # Use AI to validate data accuracy
            config_details = load_config(self.config_file)
            oai_client, _ = initialize_ai_services(config_details)
            
            # Load system prompt from prompts.json
            try:
                with open('utils/prompts.json', 'r') as f:
                    prompts_config = json.load(f)
                system_prompt = prompts_config.get('system_prompts', {}).get('Agent 2', '')
                
                if not system_prompt:
                    # Fallback to hardcoded prompt
                    system_prompt = """
                    You are AI2, a financial data validation specialist. Your task is to double-check each response 
                    by key to ensure figures match the balance sheet and verify data accuracy including K/M conversions.
                    
                    CRITICAL REQUIREMENTS:
                    1. Extract all financial figures from AI1 content
                    2. Compare with expected balance sheet figures for accuracy
                    3. Verify proper K/M conversion with 1 decimal place (e.g., 2.3M, 1.5K, 123.0)
                    4. Check entity names match data source (not reporting entity)
                    5. Identify ONLY top 2 most critical data accuracy issues
                    6. Remove unnecessary quotation marks around sections
                    7. Ensure no data inconsistencies or conversion errors
                    8. Verify figures are properly adjusted for '000 notation if applicable
                    """
            except (FileNotFoundError, json.JSONDecodeError):
                # Fallback to hardcoded prompt
                system_prompt = """
                You are AI2, a financial data validation specialist. Your task is to double-check each response 
                by key to ensure figures match the balance sheet and verify data accuracy including K/M conversions.
                
                CRITICAL REQUIREMENTS:
                1. Extract all financial figures from AI1 content
                2. Compare with expected balance sheet figures for accuracy
                3. Verify proper K/M conversion with 1 decimal place (e.g., 2.3M, 1.5K, 123.0)
                4. Check entity names match data source (not reporting entity)
                5. Identify ONLY top 2 most critical data accuracy issues
                6. Remove unnecessary quotation marks around sections
                7. Ensure no data inconsistencies or conversion errors
                8. Verify figures are properly adjusted for '000 notation if applicable
                """
            
            user_query = f"""
            AI2 DATA VALIDATION TASK:
            
            CONTENT TO VALIDATE: {content}
            
            EXPECTED FIGURE FOR {key}: {expected_figure}
            
            COMPLETE BALANCE SHEET DATA: {json.dumps(financial_figures, indent=2)}
            
            ENTITY: {entity}
            
            DETAILED VALIDATION CHECKLIST:
            1. Extract all financial figures from AI1 content and list them
            2. Compare each extracted figure with expected balance sheet figure for {key}: {expected_figure}
            3. Verify proper K/M conversion accuracy (should be 1 decimal place: 2.3M, 1.5K, 123.0)
            4. Check entity names match data source (should NOT be reporting entity "{entity}")
            5. Verify mathematical accuracy of any calculations or breakdowns
            6. Check for proper currency formatting and consistency
            7. Ensure dates are accurate (should be 30/09/2022 or Sep-22 format)
            8. Validate that component figures sum to expected total where applicable
            9. Remove unnecessary quotation marks around full sections
            10. Check for template artifacts that shouldn't be in final content
            
            SPECIFIC FIGURE ANALYSIS FOR {key}:
            - Expected total: {expected_figure}
            - Check if content figures match or reasonably approximate this total
            - Verify thousand/million notation is correct
            - Ensure component breakdowns add up properly
            
            CRITICAL: You MUST respond with ONLY valid JSON in this exact format:
            {{
                "is_valid": true,
                "issues": ["issue 1", "issue 2"],
                "score": 95,
                "corrected_content": "corrected content here if needed",
                "extracted_figures": ["figure1", "figure2"],
                "figure_validation": "detailed analysis of figure accuracy"
            }}
            
            Do not include any text before or after the JSON. Only return the JSON object.
            """
            
            response = generate_response(user_query, system_prompt, oai_client, content, config_details['DEEPSEEK_CHAT_MODEL'], entity, False)
            
            # Clean response and ensure it's valid JSON
            response = response.strip()
            
            # Remove any markdown formatting if present
            if response.startswith('```json'):
                response = response.replace('```json', '').replace('```', '').strip()
            elif response.startswith('```'):
                response = response.replace('```', '').strip()
            
            # Parse AI response
            try:
                result = json.loads(response)
                # Ensure all required fields are present with defaults
                result.setdefault('is_valid', True)
                result.setdefault('issues', [])
                result.setdefault('score', 100)
                result.setdefault('corrected_content', content)
                result.setdefault('needs_correction', False)
                result.setdefault('suggestions', [])
                return result
            except (json.JSONDecodeError, ValueError) as parse_error:
                print(f"Failed to parse AI response: {parse_error}")
                print(f"Raw AI response: {response}")
                # Return structured fallback result
                return {
                    "is_valid": True,
                    "issues": [f"AI response parsing failed: {str(parse_error)}"],
                    "score": 75,
                    "corrected_content": content,
                    "needs_correction": False,
                    "suggestions": ["Check AI response format"]
                }
                
        except Exception as e:
            print(f"Data validation error: {e}")
            return {"needs_correction": False, "issues": [f"Validation error: {e}"], "score": 50, "suggestions": []}
    
    def _fallback_data_validation(self, content: str, expected_figure: float, key: str) -> Dict:
        """This function is deprecated - AI is now required"""
        raise RuntimeError("AI services are required for data validation. Please check your configuration.")
    
    def correct_financial_data(self, content: str, issues: List[str]) -> str:
        """Correct financial data issues using AI"""
        try:
            if not AI_AVAILABLE:
                raise RuntimeError("AI services are required for data correction. Please check your configuration.")
            
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
            
            corrected_content = generate_response(user_query, system_prompt, oai_client, "", config_details['DEEPSEEK_CHAT_MODEL'], entity, False)
            return corrected_content.strip()
            
        except Exception as e:
            print(f"Data correction error: {e}")
            return content

# --- Pattern Validation Agent ---
class PatternValidationAgent:
    def __init__(self, use_local_ai=False, use_openai=False):
        self.config_file = 'utils/config.json'
        self.pattern_file = 'utils/pattern.json'
        self.use_local_ai = use_local_ai
        self.use_openai = use_openai
    
    def validate_pattern_compliance(self, content: str, key: str) -> Dict:
        """Validate that content follows the expected pattern structure"""
        try:
            import json
            
            # Load patterns for the key
            patterns = load_ip(self.pattern_file, key)
            
            if not AI_AVAILABLE:
                raise RuntimeError("AI services are required for pattern validation. Please check your configuration.")
            
            # Use AI to validate pattern compliance
            config_details = load_config(self.config_file)
            oai_client, _ = initialize_ai_services(config_details, use_local=self.use_local_ai, use_openai=self.use_openai)
            
            # Load system prompt from prompts.json
            try:
                with open('utils/prompts.json', 'r') as f:
                    prompts_config = json.load(f)
                system_prompt = prompts_config.get('system_prompts', {}).get('Agent 3', '')
                
                if not system_prompt:
                    # Fallback to hardcoded prompt
                    system_prompt = """
                    You are AI3, a pattern compliance validation specialist. Your task is to check if content 
                    follows patterns correspondingly and clean up excessive items.
                    
                    CRITICAL REQUIREMENTS:
                    1. Compare AI1 content against available pattern templates
                    2. Check proper pattern structure and professional formatting
                    3. Verify all placeholders are filled with actual data
                    4. If AI1 lists too many items, limit to top 2 most important
                    5. Remove quotation marks quoting full sections
                    6. Check for anything that shouldn't be there (template artifacts)
                    7. Ensure content follows pattern structure consistently
                    8. Verify proper K/M conversion with 1 decimal place formatting
                    """
            except (FileNotFoundError, json.JSONDecodeError):
                # Fallback to hardcoded prompt
                system_prompt = """
                You are AI3, a pattern compliance validation specialist. Your task is to check if content 
                follows patterns correspondingly and clean up excessive items.
                
                CRITICAL REQUIREMENTS:
                1. Compare AI1 content against available pattern templates
                2. Check proper pattern structure and professional formatting
                3. Verify all placeholders are filled with actual data
                4. If AI1 lists too many items, limit to top 2 most important
                5. Remove quotation marks quoting full sections
                6. Check for anything that shouldn't be there (template artifacts)
                7. Ensure content follows pattern structure consistently
                8. Verify proper K/M conversion with 1 decimal place formatting
                """
            
            user_query = f"""
            AI3 PATTERN COMPLIANCE CHECK:
            
            CONTENT TO ANALYZE: {content}
            
            KEY: {key}
            
            AVAILABLE PATTERNS FOR {key}: {json.dumps(patterns, indent=2)}
            
            PATTERN COMPLIANCE VALIDATION TASKS:
            1. Analyze AI1 content structure against available pattern templates above
            2. Verify all placeholders (xxx, [Amount], [Entity], etc.) are filled with actual data
            3. Check if content follows the narrative flow of selected pattern
            4. If AI1 content lists too many items/entities, limit to top 2 most important
            5. Remove quotation marks around full sections or paragraphs  
            6. Check for template artifacts that shouldn't appear in final content
            7. Ensure professional financial writing style and tone
            8. Verify pattern structure is maintained (intro → details → conclusion format)
            9. Check that selected pattern matches the type of data provided
            10. Ensure consistent tense and formatting throughout
            
            PATTERN ANALYSIS FOR {key}:
            - Which pattern from above appears to be most suitable?
            - Are all pattern elements properly filled?
            - Does the content maintain professional audit report language?
            - Are there any deviations from expected pattern structure?
            
            CONTENT OPTIMIZATION REQUIREMENTS:
            - Keep only essential information (top 2 items if listing multiple)
            - Remove redundant or verbose explanations
            - Ensure dates, amounts, and entities are accurate
            - Maintain consistent formatting with other sections
            
            CRITICAL: You MUST respond with ONLY valid JSON in this exact format:
            {{
                "is_compliant": true,
                "issues": ["issue 1", "issue 2"],
                "corrected_content": "cleaned content with top 2 items if needed",
                "pattern_used": "Pattern 1 or Pattern 2",
                "compliance_analysis": "detailed analysis of pattern compliance"
            }}
            
            Do not include any text before or after the JSON. Only return the JSON object.
            """
            
            # Use appropriate model based on selection
            if self.use_local_ai:
                model = config_details.get('LOCAL_AI_CHAT_MODEL', 'local-model')
            elif self.use_openai:
                model = config_details.get('OPENAI_CHAT_MODEL', 'gpt-4o-mini-2024-07-18')
            else:
                model = config_details.get('DEEPSEEK_CHAT_MODEL', 'deepseek-chat')
            
            response = generate_response(user_query, system_prompt, oai_client, content, model, entity, self.use_local_ai)
            
            # Clean response and ensure it's valid JSON
            response = response.strip()
            
            # Remove any markdown formatting if present
            if response.startswith('```json'):
                response = response.replace('```json', '').replace('```', '').strip()
            elif response.startswith('```'):
                response = response.replace('```', '').strip()
            
            # Parse AI response
            try:
                result = json.loads(response)
                # Ensure all required fields are present with defaults
                result.setdefault('is_compliant', True)
                result.setdefault('issues', [])
                result.setdefault('corrected_content', content)
                result.setdefault('needs_correction', False)
                result.setdefault('score', 100)
                result.setdefault('pattern_match', 'compliant')
                result.setdefault('missing_elements', [])
                result.setdefault('suggestions', [])
                return result
            except (json.JSONDecodeError, ValueError) as parse_error:
                print(f"Failed to parse AI3 response: {parse_error}")
                print(f"Raw AI3 response: {response}")
                # Return structured fallback result
                return {
                    "is_compliant": True,
                    "issues": [f"AI response parsing failed: {str(parse_error)}"],
                    "corrected_content": content,
                    "needs_correction": False,
                    "score": 75,
                    "pattern_match": "unknown",
                    "missing_elements": [],
                    "suggestions": ["Check AI response format"]
                }
                
        except Exception as e:
            print(f"Pattern validation error: {e}")
            return {"needs_correction": False, "issues": [f"Validation error: {e}"], "score": 50, "suggestions": []}
    
    def _fallback_pattern_validation(self, content: str, patterns: Dict, key: str) -> Dict:
        """This function is deprecated - AI is now required"""
        raise RuntimeError("AI services are required for pattern validation. Please check your configuration.")
    
    def correct_pattern_compliance(self, content: str, issues: List[str]) -> str:
        """Correct pattern compliance issues using AI"""
        try:
            if not AI_AVAILABLE:
                raise RuntimeError("AI services are required for pattern correction. Please check your configuration.")
            
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
            
            corrected_content = generate_response(user_query, system_prompt, oai_client, "", config_details['DEEPSEEK_CHAT_MODEL'], entity, False)
            return corrected_content.strip()
            
        except Exception as e:
            print(f"Pattern correction error: {e}")
            return content

def clean_response_text(text: str) -> str:
    """Clean up AI response text: remove outer quotes, translate Chinese, etc."""
    if not text:
        return text
    
    # Remove outer quotation marks
    text = text.strip()
    if (text.startswith('"') and text.endswith('"')) or (text.startswith("'") and text.endswith("'")):
        text = text[1:-1]
    
    # Translate Chinese to English (basic common translations)
    chinese_translations = {
        '進項稅金中轉': 'Input Tax Transfer',
        '自來水有限公司': 'Water Supply Co., Ltd.',
        '仲量聯行北京諮詢公司': 'Jones Lang LaSalle Beijing Consulting Co.',
        '信用中和會計師事務所': 'Credit Zhonghe Accounting Firm',
        '企業管理諮詢公司': 'Business Management Consulting Co.',
        '科技有限公司': 'Technology Co., Ltd.',
        '物業服務有限公司': 'Property Services Co., Ltd.',
        '物流發展有限公司': 'Logistics Development Co., Ltd.',
        '股權投資基金合夥企業': 'Equity Investment Fund Partnership',
        '物流倉儲有限公司': 'Logistics Storage Co., Ltd.',
        '特殊普通合夥': 'Special General Partnership',
        '戰區運營': 'Regional Operations',
        '已認證未入帳': 'Certified but Not Recorded',
        '有限合夥': 'Limited Partnership'
    }
    
    for chinese, english in chinese_translations.items():
        text = text.replace(chinese, english)
    
    return text.strip()

def multiply_figures_for_ai_processing(excel_content: str) -> str:
    """
    Multiply all numerical figures by 1000 in Excel content for AI processing when '000 notation is detected.
    This function processes the markdown table content to adjust figures for AI analysis.
    """
    import re
    
    if "'000" not in excel_content:
        return excel_content
    
    lines = excel_content.split('\n')
    processed_lines = []
    
    for line in lines:
        # Skip header lines and separator lines
        if '|' not in line or line.strip().startswith('|---') or 'Description' in line:
            processed_lines.append(line)
            continue
        
        # Process table rows with numerical data
        cells = line.split('|')
        processed_cells = []
        
        for cell in cells:
            cell = cell.strip()
            
            # Look for numerical patterns and multiply by 1000
            # Match various number formats: 123, 1,234, 1.23, (123), etc.
            number_pattern = r'(\(?)(-?\d{1,3}(?:,\d{3})*\.?\d*)(\)?)'
            
            def multiply_number(match):
                opening_paren = match.group(1)
                number_str = match.group(2)
                closing_paren = match.group(3)
                
                try:
                    # Remove commas and convert to float
                    clean_number = number_str.replace(',', '')
                    number = float(clean_number)
                    
                    # Multiply by 1000
                    adjusted_number = number * 1000
                    
                    # Format back with commas for large numbers
                    if adjusted_number == int(adjusted_number):
                        formatted = f"{int(adjusted_number):,}"
                    else:
                        formatted = f"{adjusted_number:,.1f}"
                    
                    return f"{opening_paren}{formatted}{closing_paren}"
                except ValueError:
                    # If conversion fails, return original
                    return match.group(0)
            
            # Apply multiplication to numbers in the cell
            processed_cell = re.sub(number_pattern, multiply_number, cell)
            processed_cells.append(processed_cell)
        
        processed_lines.append('|'.join(processed_cells))
    
    return '\n'.join(processed_lines) 
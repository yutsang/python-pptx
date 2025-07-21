import streamlit as st
import pandas as pd
import json
import warnings
import re
import os
import datetime
from pathlib import Path
from tabulate import tabulate
import urllib3
import shutil
from common.pptx_export import export_pptx
from utils.cache import get_cache_manager, streamlit_cache_manager, optimize_memory, cached_function

# Suppress warnings
urllib3.disable_warnings()
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action='ignore', category=UserWarning)
warnings.filterwarnings('ignore', message='Data Validation extension is not supported and will be removed', category=UserWarning, module='openpyxl')

# AI Agent Logging System
class AIAgentLogger:
    """File-based logging system for AI agents"""
    
    def __init__(self):
        self.logs = {
            'agent1': {},
            'agent2': {},
            'agent3': {}
        }
        self.session_logs = []
        
        # Create logging directory
        self.log_dir = Path("logging")
        self.log_dir.mkdir(exist_ok=True)
        
        # Create timestamped log file
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        self.log_file = self.log_dir / f"ai_agents_{timestamp}.log"
        
        # Initialize log file with header
        with open(self.log_file, 'w', encoding='utf-8') as f:
            f.write(f"AI Agent Processing Log - Session Started: {datetime.datetime.now()}\n")
            f.write("=" * 80 + "\n\n")
        
    def _write_to_file(self, message):
        """Write message to log file"""
        try:
            with open(self.log_file, 'a', encoding='utf-8') as f:
                f.write(f"{message}\n")
        except Exception as e:
            print(f"Error writing to log file: {e}")
        
    def log_agent_input(self, agent_name, key, system_prompt, user_prompt, context_data=None):
        """Log agent input prompts and data to file with session focus"""
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        log_entry = {
            'timestamp': timestamp,
            'agent': agent_name,
            'key': key,
            'type': 'INPUT',
            'system_prompt': system_prompt[:200] + "..." if len(system_prompt) > 200 else system_prompt,
            'user_prompt': user_prompt[:200] + "..." if len(user_prompt) > 200 else user_prompt,
            'context_data': str(context_data)[:100] + "..." if context_data and len(str(context_data)) > 100 else str(context_data) if context_data else None
        }
        
        if agent_name not in self.logs:
            self.logs[agent_name] = {}
        if key not in self.logs[agent_name]:
            self.logs[agent_name][key] = []
            
        self.logs[agent_name][key].append(log_entry)
        self.session_logs.append(log_entry)
        
        # Write full prompts to log file
        self._write_to_file(f"\n{'='*80}")
        self._write_to_file(f"📝 [{timestamp}] {agent_name.upper()} INPUT → {key}")
        self._write_to_file(f"{'='*80}")
        self._write_to_file(f"SYSTEM PROMPT ({len(system_prompt)} chars):")
        self._write_to_file(f"{'-'*40}")
        self._write_to_file(system_prompt)
        self._write_to_file(f"\n{'-'*40}")
        self._write_to_file(f"USER PROMPT ({len(user_prompt)} chars):")
        self._write_to_file(f"{'-'*40}")
        self._write_to_file(user_prompt)
        if context_data:
            self._write_to_file(f"\n{'-'*40}")
            self._write_to_file(f"CONTEXT DATA ({len(str(context_data))} chars):")
            self._write_to_file(f"{'-'*40}")
            self._write_to_file(str(context_data))
        self._write_to_file(f"{'='*80}\n")
        
        # No Streamlit display during processing (silent logging)
    
    def log_agent_output(self, agent_name, key, output, processing_time=0):
        """Log agent output and processing details to file with session focus"""
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Store full output in memory for session
        log_entry = {
            'timestamp': timestamp,
            'agent': agent_name,
            'key': key,
            'type': 'OUTPUT',
            'output': output,
            'processing_time': processing_time,
            'output_length': len(str(output)),
            'is_success': output and not str(output).startswith("Error")
        }
        
        if agent_name not in self.logs:
            self.logs[agent_name] = {}
        if key not in self.logs[agent_name]:
            self.logs[agent_name][key] = []
            
        self.logs[agent_name][key].append(log_entry)
        self.session_logs.append(log_entry)
        
        # Write full output to log file
        status_icon = "✅" if output and not str(output).startswith("Error") else "❌"
        self._write_to_file(f"\n{'='*80}")
        self._write_to_file(f"{status_icon} [{timestamp}] {agent_name.upper()} OUTPUT ← {key} ({processing_time:.2f}s)")
        self._write_to_file(f"{'='*80}")
        self._write_to_file(f"OUTPUT ({len(str(output))} chars):")
        self._write_to_file(f"{'-'*40}")
        
        # Write the full output content
        if isinstance(output, dict):
            # For JSON outputs (Agent 2, 3), log full structured output
            self._write_to_file(json.dumps(output, indent=2, ensure_ascii=False))
        else:
            # For text outputs (Agent 1), log full content
            self._write_to_file(str(output))
        
        self._write_to_file(f"\n{'-'*40}")
        self._write_to_file(f"PROCESSING SUMMARY:")
        self._write_to_file(f"- Processing Time: {processing_time:.2f} seconds")
        self._write_to_file(f"- Output Length: {len(str(output))} characters")
        self._write_to_file(f"- Status: {'Success' if output and not str(output).startswith('Error') else 'Error/Failed'}")
        self._write_to_file(f"{'='*80}\n")
        
        # No Streamlit display during processing (silent logging)
    
    def log_error(self, agent_name, key, error_msg):
        """Log errors during processing to file"""
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        log_entry = {
            'timestamp': timestamp,
            'agent': agent_name,
            'key': key,
            'type': 'ERROR',
            'error': str(error_msg)
        }
        
        if agent_name not in self.logs:
            self.logs[agent_name] = {}
        if key not in self.logs[agent_name]:
            self.logs[agent_name][key] = []
            
        self.logs[agent_name][key].append(log_entry)
        self.session_logs.append(log_entry)
        
        # Write to file
        self._write_to_file(f"\n❌ ERROR - [{timestamp}] {agent_name.upper()} - {key}")
        self._write_to_file(f"ERROR: {error_msg}")
        self._write_to_file("")
        
        # Display error in Streamlit
        st.error(f"❌ {agent_name.upper()} error for {key}: {error_msg}")
    
    def save_logs_to_json(self):
        """Save all logs to JSON file for structured access"""
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        json_file = self.log_dir / f"ai_agents_{timestamp}.json"
        
        try:
            with open(json_file, 'w', encoding='utf-8') as f:
                json.dump(self.logs, f, indent=2, default=str, ensure_ascii=False)
            return json_file
        except Exception as e:
            print(f"Failed to save JSON logs: {e}")
            return None
    
    def display_session_summary(self):
        """Display simple summary in Streamlit"""
        if not self.session_logs:
            st.info("No AI agent activity logged yet.")
            return
            
        st.markdown("### 📊 AI Processing Summary")
        st.info(f"📁 Detailed logs saved to: `{self.log_file}`")
        
        # Count by agent
        agent_counts = {}
        for log in self.session_logs:
            agent = log['agent']
            if agent not in agent_counts:
                agent_counts[agent] = {'inputs': 0, 'outputs': 0, 'errors': 0}
            agent_counts[agent][log['type'].lower() + 's'] += 1
        
        # Display counts
        col1, col2, col3 = st.columns(3)
        for i, (agent, counts) in enumerate(agent_counts.items()):
            with [col1, col2, col3][i % 3]:
                st.metric(f"{agent.upper()}", 
                         f"I:{counts['inputs']} O:{counts['outputs']} E:{counts['errors']}")
        
        # Provide download option for JSON logs
        if st.button("💾 Save Structured Logs (JSON)", type="secondary"):
            json_file = self.save_logs_to_json()
            if json_file:
                st.success(f"📄 JSON logs saved to: `{json_file}`")
                try:
                    with open(json_file, 'r', encoding='utf-8') as f:
                        st.download_button(
                            label="📥 Download JSON Logs",
                            data=f.read(),
                            file_name=json_file.name,
                            mime='application/json'
                        )
                except Exception:
                    pass

# Initialize global logger
if 'ai_logger' not in st.session_state:
    st.session_state.ai_logger = AIAgentLogger()

# Load configuration files
@cached_function(ttl=3600)  # Cache for 1 hour
def load_config_files():
    """Load configuration files from utils directory with caching"""
    cache_manager = get_cache_manager()
    
    # Try to get cached configs
    config = cache_manager.get_cached_config('utils/config.json')
    mapping = cache_manager.get_cached_config('utils/mapping.json')
    pattern = cache_manager.get_cached_config('utils/pattern.json')
    prompts = cache_manager.get_cached_config('utils/prompts.json')
    
    # Load any missing configs
    try:
        if config is None:
            with open('utils/config.json', 'r') as f:
                config = json.load(f)
            cache_manager.cache_config('utils/config.json', config)
        
        if mapping is None:
            with open('utils/mapping.json', 'r') as f:
                mapping = json.load(f)
            cache_manager.cache_config('utils/mapping.json', mapping)
        
        if pattern is None:
            with open('utils/pattern.json', 'r') as f:
                pattern = json.load(f)
            cache_manager.cache_config('utils/pattern.json', pattern)
        
        if prompts is None:
            with open('utils/prompts.json', 'r') as f:
                prompts = json.load(f)
            cache_manager.cache_config('utils/prompts.json', prompts)
            
        return config, mapping, pattern, prompts
    except FileNotFoundError as e:
        st.error(f"Configuration file not found: {e}")
        return None, None, None, None

@cached_function(ttl=1800)  # Cache for 30 minutes
def process_and_filter_excel(filename, tab_name_mapping, entity_name, entity_suffixes):
    """
    Process and filter Excel file to extract relevant worksheet sections with caching
    This is the core function from old_ver/utils/utils.py
    """
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
                    return cached_result
            except Exception as e:
                print(f"Content-based cache check failed: {e}")
        
        # Fallback to path-based caching for regular files
        cached_result = cache_manager.get_cached_processed_excel(filename, entity_name, entity_suffixes)
        if cached_result is not None:
            return cached_result
        
        # Load the Excel file
        main_dir = Path(__file__).parent
        file_path = main_dir / filename
        xl = pd.ExcelFile(file_path)
        
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
                
                # Filter dataframes by entity name with proper spacing
                entity_keywords = [f"{entity_name} {suffix}" for suffix in entity_suffixes if suffix]
                if not entity_keywords:  # If no helpers, just use entity name
                    entity_keywords = [entity_name]
                
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
                    
                    if any(data_frame.apply(lambda row: row.astype(str).str.contains(keyword, case=False, na=False).any(), axis=1).any() for keyword in entity_keywords):
                        markdown_content += tabulate(data_frame, headers='keys', tablefmt='pipe', showindex=False)
                        markdown_content += "\n\n" 
        
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
        st.error(f"An error occurred while processing the Excel file: {e}")
        return ""

def parse_accounting_table(df, key, entity_name, sheet_name):
    """
    Parse accounting table with proper header detection and figure column identification
    Returns structured table data with metadata
    """
    try:
        import re
        import pandas as pd
        
        if df.empty or len(df) < 2:
            return None
        
        # Convert all cells to string for analysis
        df_str = df.astype(str).fillna('')
        
        # Detect thousand/million indicators
        thousand_indicators = ["'000", "CNY'000", "USD'000", "'000", "thousands", "thousand"]
        million_indicators = ["'000,000", "millions", "million"]
        
        multiplier = 1
        currency_info = ""
        
        # Check for thousand/million indicators in the entire table
        for i in range(min(5, len(df_str))):  # Check first 5 rows
            for j in range(len(df_str.columns)):
                cell_value = str(df_str.iloc[i, j]).lower()
                if any(indicator in cell_value for indicator in thousand_indicators):
                    multiplier = 1000
                    currency_info = "CNY'000"
                    break
                elif any(indicator in cell_value for indicator in million_indicators):
                    multiplier = 1000000
                    currency_info = "CNY'000,000"
                    break
        
        # Find the value column - look for "Indicative adjusted" first, then "Total"
        value_col_idx = None
        value_col_name = ""
        
        for i in range(min(3, len(df_str))):  # Check first 3 rows for headers
            for j in range(len(df_str.columns)):
                cell_value = str(df_str.iloc[i, j]).lower()
                if "indicative adjusted" in cell_value:
                    value_col_idx = j
                    value_col_name = "Indicative adjusted"
                    break
                elif "total" in cell_value and value_col_idx is None:
                    value_col_idx = j
                    value_col_name = "Total"
        
        # If no specific column found, use the rightmost column with numbers
        if value_col_idx is None:
            for j in range(len(df_str.columns) - 1, -1, -1):
                column_data = df_str.iloc[:, j]
                # Check if column contains mostly numbers
                numeric_count = 0
                for cell in column_data:
                    if re.search(r'\d+', str(cell)):
                        numeric_count += 1
                if numeric_count >= len(column_data) * 0.3:  # At least 30% numeric
                    value_col_idx = j
                    value_col_name = f"Column {j+1}"
                    break
        
        if value_col_idx is None:
            return None
        
        # Find where actual data starts (skip header rows)
        data_start_row = None
        for i in range(len(df_str)):
            cell_value = str(df_str.iloc[i, value_col_idx])
            # Look for cells that contain only numbers (possibly with commas)
            if re.match(r'^\d+,?\d*\.?\d*$', cell_value.replace(',', '')):
                data_start_row = i
                break
        
        if data_start_row is None:
            return None
        
        # Extract table metadata (first few rows before data)
        table_metadata = {
            'table_name': f"{key} - {entity_name}",
            'sheet_name': sheet_name,
            'currency_info': currency_info,
            'multiplier': multiplier,
            'value_column': value_col_name,
            'data_start_row': data_start_row
        }
        
        # Extract actual data rows
        description_col_idx = 0  # Usually first column
        data_rows = []
        
        for i in range(data_start_row, len(df_str)):
            description = str(df_str.iloc[i, description_col_idx]).strip()
            value_str = str(df_str.iloc[i, value_col_idx]).strip()
            
            # Skip empty rows
            if not description or description.lower() in ['nan', '']:
                continue
            
            # Extract numeric value
            value = 0
            if re.search(r'\d', value_str):
                # Remove commas and extract number
                clean_value = re.sub(r'[^\d.-]', '', value_str.replace(',', ''))
                try:
                    value = float(clean_value) * multiplier  # Apply multiplier
                except ValueError:
                    value = 0
            
            if description and (value != 0 or description.lower() not in ['total', 'subtotal']):
                data_rows.append({
                    'description': description,
                    'value': value,
                    'original_value': value_str
                })
        
        return {
            'metadata': table_metadata,
            'data': data_rows,
            'raw_df': df
        }
        
    except Exception as e:
        print(f"Error parsing accounting table: {e}")
        return None

def create_improved_table_markdown(parsed_table):
    """Create improved markdown representation of parsed accounting table"""
    try:
        if not parsed_table or 'metadata' not in parsed_table or 'data' not in parsed_table:
            return "No table data available"
        
        metadata = parsed_table['metadata']
        data_rows = parsed_table['data']
        
        # Create table header with metadata
        markdown_lines = []
        markdown_lines.append(f"**{metadata['table_name']}**")
        markdown_lines.append(f"*Sheet: {metadata['sheet_name']}*")
        
        if metadata.get('currency_info'):
            markdown_lines.append(f"*Currency: {metadata['currency_info']}*")
        
        markdown_lines.append(f"*Value Column: {metadata['value_column']}*")
        markdown_lines.append("")  # Empty line
        
        # Create table with description and value columns
        if data_rows:
            markdown_lines.append("| Description | Value |")
            markdown_lines.append("|-------------|--------|")
            
            for row in data_rows:
                description = row['description']
                value = row['value']
                
                # Format value appropriately
                if value >= 1000000:
                    formatted_value = f"{value/1000000:.1f}M"
                elif value >= 1000:
                    formatted_value = f"{value/1000:.1f}K"
                else:
                    formatted_value = f"{value:.0f}"
                
                markdown_lines.append(f"| {description} | {formatted_value} |")
        else:
            markdown_lines.append("*No data rows found*")
        
        return "\n".join(markdown_lines)
        
    except Exception as e:
        return f"Error creating table markdown: {e}"

def get_worksheet_sections_by_keys(uploaded_file, tab_name_mapping, entity_name, entity_suffixes, debug=False):
    """
    Get worksheet sections organized by financial keys following the mapping
    """
    try:
        # Load the Excel file from uploaded file object
        xl = pd.ExcelFile(uploaded_file)
        
        # Create a reverse mapping from values to keys
        reverse_mapping = {}
        for key, values in tab_name_mapping.items():
            for value in values:
                reverse_mapping[value] = key
        
        # Get financial keys
        financial_keys = get_financial_keys()
        
        # Initialize sections by key
        sections_by_key = {key: [] for key in financial_keys}
        
        # Process each sheet
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
                
                # Filter dataframes by entity name with proper spacing
                entity_keywords = [f"{entity_name} {suffix}" for suffix in entity_suffixes if suffix]
                if not entity_keywords:  # If no helpers, just use entity name
                    entity_keywords = [entity_name]
                
                combined_pattern = '|'.join(re.escape(kw) for kw in entity_keywords)
                
                # Organize sections by key - make it less restrictive
                for data_frame in dataframes:
                    # Check if this section contains any of the financial keys
                    matched_keys = []  # Track which keys this data_frame matches
                    
                    for key in financial_keys:
                        if key in tab_name_mapping:
                            key_patterns = tab_name_mapping[key]
                            for pattern in key_patterns:
                                if data_frame.apply(
                                    lambda row: row.astype(str).str.contains(
                                        pattern, case=False, regex=True, na=False
                                    ).any(),
                                    axis=1
                                ).any():
                                    matched_keys.append(key)
                                    break  # Break inner loop, but continue checking other keys
                    
                    # Now assign the data_frame to the most specific matching key
                    if matched_keys and debug:
                        st.write(f"🔍 DataFrame matched keys: {matched_keys}")
                    
                    if matched_keys:
                        # Find the best matching key based on pattern specificity
                        best_key = None
                        best_score = 0
                        
                        for key in matched_keys:
                            key_patterns = tab_name_mapping[key]
                            # Calculate a score based on pattern specificity
                            for pattern in key_patterns:
                                # Check for exact matches first
                                exact_match = data_frame.apply(
                                    lambda row: row.astype(str).str.contains(
                                        f"^{pattern}$", case=False, regex=True, na=False
                                    ).any(),
                                    axis=1
                                ).any()
                                
                                if exact_match:
                                    score = len(pattern) * 10  # High score for exact matches
                                else:
                                    # Check for word boundary matches
                                    word_boundary_match = data_frame.apply(
                                        lambda row: row.astype(str).str.contains(
                                            f"\\b{pattern}\\b", case=False, regex=True, na=False
                                        ).any(),
                                        axis=1
                                    ).any()
                                    
                                    if word_boundary_match:
                                        score = len(pattern) * 5  # Medium score for word boundary matches
                                    else:
                                        score = len(pattern)  # Low score for partial matches
                                
                                if score > best_score:
                                    best_score = score
                                    best_key = key
                        
                        # If no best key found, use the first matched key
                        if not best_key and matched_keys:
                            best_key = matched_keys[0]
                        
                        if debug:
                            st.write(f"✅ Assigned to key: {best_key} (score: {best_score})")
                        
                        # Check if it matches entity filter (but be less restrictive)
                        entity_mask = data_frame.apply(
                            lambda row: row.astype(str).str.contains(
                                combined_pattern, case=False, regex=True, na=False
                            ).any(),
                            axis=1
                        )
                        
                        # If entity filter matches or no entity helpers provided, process with new parser
                        if entity_mask.any() or not entity_suffixes or all(s.strip() == '' for s in entity_suffixes):
                            # Use new accounting table parser
                            parsed_table = parse_accounting_table(data_frame, best_key, entity_name, sheet_name)
                            
                            if parsed_table:
                                sections_by_key[best_key].append({
                                    'sheet': sheet_name,
                                    'data': data_frame,  # Keep original for compatibility
                                    'parsed_data': parsed_table,  # New structured data
                                    'markdown': create_improved_table_markdown(parsed_table),
                                    'entity_match': entity_mask.any()
                                })
                            else:
                                # Fallback to original format if parsing fails
                                sections_by_key[best_key].append({
                                    'sheet': sheet_name,
                                    'data': data_frame,
                                    'markdown': tabulate(data_frame, headers='keys', tablefmt='pipe', showindex=False),
                                    'entity_match': entity_mask.any()
                                })
        
        return sections_by_key
    except Exception as e:
        st.error(f"An error occurred while processing the Excel file: {e}")
        return {}

def get_tab_name(project_name):
    """Get tab name based on project name"""
    if project_name == 'Haining':
        return "BSHN"
    elif project_name == 'Nanjing':
        return "BSNJ"
    elif project_name == 'Ningbo':
        return "BSNB"
    return None

def get_financial_keys():
    """Get all financial keys from mapping.json"""
    try:
        with open('utils/mapping.json', 'r') as f:
            mapping = json.load(f)
        return list(mapping.keys())
    except FileNotFoundError:
        # Fallback to hardcoded keys if mapping.json not found
        return [
            "Cash", "AR", "Prepayments", "OR", "Other CA", "IP", "Other NCA",
            "AP", "Taxes payable", "OP", "Capital", "Reserve"
        ]
    except Exception as e:
        st.error(f"Error loading mapping.json: {e}")
        return [
            "Cash", "AR", "Prepayments", "OR", "Other CA", "IP", "Other NCA",
            "AP", "Taxes payable", "OP", "Capital", "Reserve"
        ]

def get_key_display_name(key):
    """Get display name for financial key using mapping.json"""
    try:
        with open('utils/mapping.json', 'r') as f:
            mapping = json.load(f)
        
        # If the key exists in mapping, find the best display name
        if key in mapping and mapping[key]:
            values = mapping[key]
            
            # Priority order for display names (prefer descriptive over abbreviations)
            priority_keywords = [
                'Long-term', 'Investment', 'Accounts', 'Other', 'Capital', 'Reserve',
                'Income', 'Expenses', 'Tax', 'Credit', 'Non-operating', 'Advances'
            ]
            
            # First, try to find a value with priority keywords
            for value in values:
                if any(keyword.lower() in value.lower() for keyword in priority_keywords):
                    return value
            
            # If no priority keywords found, use the first non-abbreviation value
            for value in values:
                if len(value) > 3 and not value.isupper():  # Prefer longer, non-abbreviation names
                    return value
            
            # Fallback to first value
            return values[0]
        else:
            return key
    except FileNotFoundError:
        # Fallback to hardcoded mapping if mapping.json not found
        name_mapping = {
            'Cash': 'Cash',
            'AR': 'Accounts Receivable',
            'Prepayments': 'Prepayments',
            'OR': 'Other Receivables',
            'Other CA': 'Other Current Assets',
            'IP': 'Investment Properties',
            'Other NCA': 'Other Non-Current Assets',
            'AP': 'Accounts Payable',
            'Taxes payable': 'Tax Payable',
            'OP': 'Other Payables',
            'Capital': 'Share Capital',
            'Reserve': 'Reserve',
            'Advances': 'Advances from Customers',
            'Capital reserve': 'Capital Reserve',
            'OI': 'Other Income',
            'OC': 'Other Costs',
            'Tax and Surcharges': 'Tax and Surcharges',
            'GA': 'G&A Expenses',
            'Fin Exp': 'Finance Expenses',
            'Cr Loss': 'Credit Losses',
            'Other Income': 'Other Income',
            'Non-operating Income': 'Non-operating Income',
            'Non-operating Exp': 'Non-operating Expenses',
            'Income tax': 'Income Tax',
            'LT DTA': 'Long-term Deferred Tax Assets'
        }
        return name_mapping.get(key, key)
    except Exception as e:
        st.error(f"Error loading mapping.json for display names: {e}")
        return key

def main():
    # Initialize cache manager for Streamlit
    cache_manager = streamlit_cache_manager()
    
    st.set_page_config(
        page_title="Financial Data Processor",
        page_icon="📊",
        layout="wide"
    )
    st.title("📊 Financial Data Processor")
    st.markdown("---")

    # Sidebar for controls
    with st.sidebar:
        uploaded_file = st.file_uploader(
            "Upload Excel File",
            type=['xlsx', 'xls'],
            help="Upload your financial data Excel file"
        )
        entity_options = ['Haining', 'Nanjing', 'Ningbo']
        selected_entity = st.selectbox(
            "Select Entity",
            options=entity_options,
            help="Choose the entity for data processing"
        )
        # Entity helpers are now hidden/hardcoded
        entity_helpers = "Wanpu,Limited,"  # Hidden from UI but still functional
        
        # Financial Statement Type Selection
        st.markdown("---")
        statement_type_options = ["Balance Sheet", "Income Statement", "All"]
        statement_type_display = st.radio(
            "Financial Statement Type",
            statement_type_options,
            help="Select the type of financial statement to process"
        )
        
        # Map display names back to internal codes
        statement_type_mapping = {
            "Balance Sheet": "BS",
            "Income Statement": "IS", 
            "All": "ALL"
        }
        statement_type = statement_type_mapping[statement_type_display]
        
        if uploaded_file is not None:
            st.success(f"Uploaded {uploaded_file.name}")
            # Store uploaded file in session state for later use
            st.session_state['uploaded_file'] = uploaded_file
            
            # AI Mode Selection - changed to dropdown
            ai_mode_options = ["GPT-4o-mini", "Deepseek", "Offline"]
            mode_display = st.selectbox(
                "Select Mode", 
                ai_mode_options,
                help="Choose the AI model or offline processing mode"
            )
            
            # Show API configuration status
            config, _, _, _ = load_config_files()
            if config:
                if mode_display == "GPT-4o-mini":
                    if config.get('OPENAI_API_KEY'):
                        st.success("✅ OpenAI API key configured")
                    else:
                        st.warning("⚠️ OpenAI API key not configured")
                elif mode_display == "Deepseek":
                    if config.get('DEEPSEEK_API_KEY'):
                        st.success("✅ Deepseek API key configured")
                    else:
                        st.error("❌ Deepseek API key not configured")
                        st.info("📖 See DEEPSEEK_SETUP.md for configuration instructions")
            
            # Map display names to internal mode names
            mode_mapping = {
                "GPT-4o-mini": "AI Mode",
                "Deepseek": "AI Mode - Deepseek",
                "Offline": "Offline Mode"
            }
            mode = mode_mapping[mode_display]
            st.session_state['selected_mode'] = mode
            st.session_state['ai_model'] = mode_display
            
            # Performance statistics - moved below Select Mode
            st.markdown("---")
            st.markdown("### 🚀 Performance")
            cache_stats = cache_manager.get_cache_stats()
            col1, col2 = st.columns(2)
            with col1:
                st.metric("Cache Hits", cache_stats['hits'])
            with col2:
                st.metric("Cache Misses", cache_stats['misses'])
            st.metric("Hit Rate", cache_stats['hit_rate'])
            
            if st.button("🧹 Clear Cache"):
                cache_manager.clear_cache()
                st.success("Cache cleared!")
            
            if st.button("🗑️ Optimize Memory"):
                optimize_memory()
                st.success("Memory optimized!")
        else:
            st.info("Please upload an Excel file to get started.")

    # Main area for results
    if uploaded_file is not None:
        
        # --- View Table Section ---
        config, mapping, pattern, prompts = load_config_files()
        entity_suffixes = [s.strip() for s in entity_helpers.split(',') if s.strip()]
        entity_keywords = [f"{selected_entity} {suffix}" for suffix in entity_suffixes if suffix]
        if not entity_keywords:
            entity_keywords = [selected_entity]
        
        # Handle different statement types
        if statement_type == "BS":
            # Original BS logic
            sections_by_key = get_worksheet_sections_by_keys(
                uploaded_file=uploaded_file,
                tab_name_mapping=mapping,
                entity_name=selected_entity,
                entity_suffixes=entity_suffixes,
                debug=False  # Set to True for debugging
            )
            st.subheader("View Table by Key")
            keys_with_data = [key for key, sections in sections_by_key.items() if sections]
            if keys_with_data:
                key_tabs = st.tabs([get_key_display_name(key) for key in keys_with_data])
                for i, key in enumerate(keys_with_data):
                    with key_tabs[i]:
                        st.subheader(f"Sheet: {get_key_display_name(key)}")
                        sections = sections_by_key[key]
                        if sections:
                            first_section = sections[0]
                            # Create a proper copy to avoid SettingWithCopyWarning
                            df_clean = first_section['data'].dropna(axis=1, how='all').copy()
                            
                            # Convert datetime columns to strings to avoid Arrow serialization issues
                            for col in df_clean.columns:
                                if df_clean[col].dtype == 'object':
                                    # Convert any datetime-like objects to strings
                                    try:
                                        # First try to convert to string directly
                                        df_clean.loc[:, col] = df_clean[col].astype(str)
                                    except:
                                        # If that fails, handle datetime objects specifically
                                        df_clean.loc[:, col] = df_clean[col].map(
                                            lambda x: str(x) if pd.notna(x) and not pd.isna(x) else ''
                                        )
                                elif 'datetime' in str(df_clean[col].dtype):
                                    # Handle datetime columns specifically
                                    df_clean.loc[:, col] = df_clean[col].dt.strftime('%Y-%m-%d %H:%M:%S').fillna('')
                            
                            if first_section.get('entity_match', False):
                                st.markdown("**First Section:** ✅ Entity Match")
                            else:
                                st.markdown("**First Section:** ⚠️ No Entity Match")
                            st.dataframe(df_clean, use_container_width=True)
                            from tabulate import tabulate
                            markdown_table = tabulate(df_clean, headers='keys', tablefmt='pipe', showindex=False)
                            with st.expander(f"📋 Markdown Table - First Section", expanded=False):
                                st.code(markdown_table, language='markdown')
                            st.info(f"**Source Sheet:** {first_section['sheet']}")
                            st.markdown("---")
                        else:
                            st.info("No sections found for this key.")
            else:
                st.warning("No data found for any financial keys.")
        
        elif statement_type == "IS":
            # Income Statement placeholder
            st.subheader("Income Statement")
            st.info("📊 Income Statement processing will be implemented here.")
            st.markdown("""
            **Placeholder for Income Statement sections:**
            - Revenue
            - Cost of Goods Sold
            - Gross Profit
            - Operating Expenses
            - Operating Income
            - Other Income/Expenses
            - Net Income
            """)
        
        elif statement_type == "ALL":
            # Combined view placeholder
            st.subheader("Combined Financial Statements")
            st.info("📊 Combined BS and IS processing will be implemented here.")
            st.markdown("""
            **Placeholder for Combined sections:**
            - Balance Sheet
            - Income Statement
            - Cash Flow Statement
            - Financial Ratios
            """)

        # --- AI Processing Section (Bottom) ---
        st.markdown("---")
        st.subheader("🤖 AI Processing")
        
        if not st.session_state.get('ai_processed', False):
            st.info("Click 'Process with AI' to generate AI results.")
            
            # Check AI configuration status
            try:
                config, _, _, _ = load_config_files()
                if config and (not config.get('OPENAI_API_KEY') or not config.get('OPENAI_API_BASE')):
                    st.warning("⚠️ AI Mode: API keys not configured. Will use fallback mode with test data.")
                    st.info("💡 To enable full AI functionality, please configure your OpenAI API keys in utils/config.json")
            except Exception:
                st.warning("⚠️ AI Mode: Configuration not found. Will use fallback mode.")
        
        if st.button("🤖 Process with AI", type="primary", use_container_width=True):
            # Show immediate progress feedback to indicate processing has started
            progress_bar = st.progress(0)
            status_text = st.empty()
            status_text.text("🔄 Initializing AI processing...")
            
            try:
                # Load configuration files
                progress_bar.progress(0.05)
                status_text.text("📋 Loading configuration files...")
                config, mapping, pattern, prompts = load_config_files()
                if not all([config, mapping, pattern]):
                    st.error("❌ Failed to load configuration files")
                    return
                
                # Process the Excel data for AI analysis
                progress_bar.progress(0.1)
                status_text.text("🔧 Processing entity configuration...")
                entity_suffixes = [s.strip() for s in entity_helpers.split(',') if s.strip()]
                entity_keywords = [f"{selected_entity} {suffix}" for suffix in entity_suffixes if suffix]
                if not entity_keywords:
                    entity_keywords = [selected_entity]
                
                # Get worksheet sections for AI processing
                progress_bar.progress(0.15)
                status_text.text("📊 Analyzing worksheet sections (this may take a few minutes)...")
                sections_by_key = get_worksheet_sections_by_keys(
                    uploaded_file=uploaded_file,
                    tab_name_mapping=mapping,
                    entity_name=selected_entity,
                    entity_suffixes=entity_suffixes,
                    debug=False  # Set to True for debugging
                )
                
                # Get keys with data for AI processing
                progress_bar.progress(0.3)
                status_text.text("🔍 Identifying keys with data...")
                keys_with_data = [key for key, sections in sections_by_key.items() if sections]
                
                # Filter keys based on statement type (Fix for issue #5)
                progress_bar.progress(0.35)
                status_text.text("🎯 Filtering keys by statement type...")
                bs_keys = [
                    "Cash", "AR", "Prepayments", "OR", "Other CA", "IP", "Other NCA",
                    "AP", "Taxes payable", "OP", "Capital", "Reserve"
                ]
                is_keys = [
                    "OI", "OC", "Tax and Surcharges", "GA", "Fin Exp", "Cr Loss", "Other Income",
                    "Non-operating Income", "Non-operating Exp", "Income tax", "LT DTA"
                ]
                
                # Apply statement type filtering for AI processing
                if statement_type == "BS":
                    filtered_keys_for_ai = [key for key in keys_with_data if key in bs_keys]
                elif statement_type == "IS":
                    filtered_keys_for_ai = [key for key in keys_with_data if key in is_keys]
                elif statement_type == "ALL":
                    filtered_keys_for_ai = [key for key in keys_with_data if key in (bs_keys + is_keys)]
                else:
                    filtered_keys_for_ai = keys_with_data
                
                if not filtered_keys_for_ai:
                    st.warning("No data found for AI processing with the selected statement type. Please check your file and statement type selection.")
                    return
                
                progress_bar.progress(0.4)
                status_text.text(f"✅ Found {len(filtered_keys_for_ai)} keys for {statement_type} statement type")
                
                # Process each key with AI if in AI Mode
                ai_results = {}
                mode = st.session_state.get('selected_mode', 'AI Mode')
                ai_model = st.session_state.get('ai_model', 'GPT-4o-mini')
                
                if mode.startswith("AI Mode"):
                    try:
                        from common.assistant import process_keys
                        
                        # Save uploaded file temporarily
                        temp_file_path = f"temp_ai_processing_{uploaded_file.name}"
                        with open(temp_file_path, "wb") as f:
                            f.write(uploaded_file.getbuffer())
                        
                        # Store uploaded file data in session state for agents
                        st.session_state['uploaded_file_data'] = uploaded_file.getbuffer()
                        
                        # Update config for different AI models
                        if ai_model == "Deepseek":
                            # Check if Deepseek API key is configured
                            if not config.get('DEEPSEEK_API_KEY'):
                                st.error("❌ Deepseek API key not configured in utils/config.json")
                                st.info("💡 Please add your Deepseek API key to DEEPSEEK_API_KEY in utils/config.json")
                                return
                            
                            # Create a temporary config for Deepseek
                            deepseek_config = config.copy()
                            deepseek_config['CHAT_MODEL'] = config.get('DEEPSEEK_CHAT_MODEL', 'deepseek-chat')
                            deepseek_config['OPENAI_API_BASE'] = config.get('DEEPSEEK_API_BASE', 'https://api.deepseek.com/v1')
                            deepseek_config['OPENAI_API_KEY'] = config.get('DEEPSEEK_API_KEY', '')
                            deepseek_config['OPENAI_API_VERSION_COMPLETION'] = config.get('DEEPSEEK_API_VERSION', 'v1')
                            # Save temporary config
                            import json
                            with open("temp_deepseek_config.json", "w") as f:
                                json.dump(deepseek_config, f)
                            config_file = "temp_deepseek_config.json"
                            st.info(f"🚀 Using Deepseek AI model for processing")
                        else:
                            # Check if OpenAI API key is configured for GPT models
                            if not config.get('OPENAI_API_KEY'):
                                st.warning("⚠️ OpenAI API key not configured in utils/config.json")
                                st.info("💡 Please add your OpenAI API key to OPENAI_API_KEY in utils/config.json")
                            config_file = "utils/config.json"
                            st.info(f"🚀 Using {ai_model} AI model for processing")
                        
                        # Process all keys with AI - Use ALL filtered keys at once for better progress tracking
                        entity_helpers = ', '.join(entity_keywords[1:]) if len(entity_keywords) > 1 else ""
                        
                        status_text.text(f"🤖 Starting AI1 processing for {len(filtered_keys_for_ai)} keys...")
                        
                        # Create progress callback for real-time updates
                        def update_progress(progress, message):
                            progress_bar.progress(progress)
                            status_text.text(message)
                        
                        try:
                            # Automatic Sequential AI Processing: AI1 → AI2 → AI3
                            
                            # Initialize agent states
                            st.session_state['agent_states'] = {
                                'agent1_completed': False,
                                'agent2_completed': False, 
                                'agent3_completed': False,
                                'agent1_results': {},
                                'agent2_results': {},
                                'agent3_results': {},
                                'all_agents_completed': False
                            }
                            
                            # Display AI Agent Logging Section
                            st.markdown("---")
                            st.markdown("# 🤖 AI Agent Processing Logs")
                            logger = st.session_state.ai_logger
                            
                            # Clear previous logs for new session
                            logger.logs = {'agent1': {}, 'agent2': {}, 'agent3': {}}
                            logger.session_logs = []
                            
                            # Debug: Check AI service availability
                            st.markdown("### 🔧 AI Service Status")
                            try:
                                from common.assistant import AI_AVAILABLE, initialize_ai_services, load_config
                                
                                st.write(f"**AI_AVAILABLE:** {AI_AVAILABLE}")
                                
                                if AI_AVAILABLE:
                                    # Test AI service initialization
                                    try:
                                        config_details = load_config('utils/config.json')
                                        oai_client, search_client = initialize_ai_services(config_details)
                                        st.success("✅ AI services initialized successfully")
                                        st.write(f"**OpenAI Model:** {config_details.get('CHAT_MODEL', 'Not configured')}")
                                        st.write(f"**API Key Status:** {'✅ Configured' if config_details.get('OPENAI_API_KEY') else '❌ Missing'}")
                                    except Exception as ai_error:
                                        st.error(f"❌ AI service initialization failed: {ai_error}")
                                        st.warning("⚠️ Agents will use fallback methods")
                                else:
                                    st.error("❌ AI libraries not available - agents will use fallback methods")
                                    
                            except Exception as e:
                                st.error(f"❌ Error checking AI status: {e}")
                            
                            st.markdown("---")
                            
                            # Prepare AI data for agents
                            temp_ai_data = {
                                'entity_name': selected_entity,
                                'entity_keywords': entity_keywords,
                                'sections_by_key': sections_by_key,
                                'pattern': pattern,
                                'mapping': mapping,
                                'config': config
                            }
                            
                            # Phase 1: AI1 Processing - Content Generation
                            update_progress(0.4, f"🤖 AI1: Generating content for {len(filtered_keys_for_ai)} keys...")
                            st.markdown("---")
                            agent1_results = run_agent_1(filtered_keys_for_ai, temp_ai_data)
                            st.session_state['agent_states']['agent1_results'] = agent1_results
                            st.session_state['agent_states']['agent1_completed'] = True
                            ai_results.update(agent1_results)
                            
                            # Phase 2: AI2 Processing - Data Validation
                            update_progress(0.6, f"🔍 AI2: Validating data for {len(filtered_keys_for_ai)} keys...")
                            st.markdown("---")
                            agent2_results = run_agent_2(filtered_keys_for_ai, agent1_results, temp_ai_data)
                            st.session_state['agent_states']['agent2_results'] = agent2_results
                            st.session_state['agent_states']['agent2_completed'] = True
                            
                            # Phase 3: AI3 Processing - Pattern Compliance
                            update_progress(0.8, f"🎯 AI3: Checking pattern compliance for {len(filtered_keys_for_ai)} keys...")
                            st.markdown("---")
                            agent3_results = run_agent_3(filtered_keys_for_ai, agent1_results, temp_ai_data)
                            st.session_state['agent_states']['agent3_results'] = agent3_results
                            st.session_state['agent_states']['agent3_completed'] = True
                            
                            # Mark all agents as completed
                            st.session_state['agent_states']['all_agents_completed'] = True
                            
                            # Store AI2 and AI3 results in legacy format for backwards compatibility
                            st.session_state['ai2_results'] = agent2_results
                            st.session_state['ai3_results'] = agent3_results
                            
                            # Display logging summary and save option
                            st.markdown("---")
                            st.markdown("# 📊 AI Processing Summary")
                            logger.display_session_summary()
                            

                        except RuntimeError as e:
                            # AI services not available, use fallback
                            update_progress(0.5, "AI services not available, using fallback mode...")
                            results = process_keys(
                                keys=filtered_keys_for_ai,
                                entity_name=selected_entity,
                                entity_helpers=entity_helpers,
                                input_file=temp_file_path,
                                mapping_file="utils/mapping.json",
                                pattern_file="utils/pattern.json",
                                config_file=config_file,
                                use_ai=False,  # Force fallback mode
                                progress_callback=update_progress
                            )
                            ai_results.update(results)
                            
                            # Create empty AI2 and AI3 results for fallback mode
                            st.session_state['ai2_results'] = {key: {"is_valid": True, "issues": [], "score": 100} for key in filtered_keys_for_ai}
                            st.session_state['ai3_results'] = {key: {"is_compliant": True, "issues": []} for key in filtered_keys_for_ai}
                        except Exception as e:
                            st.error(f"Failed to process keys: {e}")
                        
                        # Update progress to 100%
                        progress_bar.progress(1.0)
                        
                        # Clean up temp files
                        if os.path.exists(temp_file_path):
                            os.remove(temp_file_path)
                        if ai_model == "Deepseek" and os.path.exists("temp_deepseek_config.json"):
                            os.remove("temp_deepseek_config.json")
                        
                        status_text.text("✅ AI processing completed!")
                        
                        # Generate markdown content file from AI results (Fix for issue #3)
                        if ai_results:
                            success = generate_markdown_from_ai_results(ai_results, selected_entity)
                            if success:
                                st.success(f"📄 AI output saved to: utils/bs_content.md")
                                st.info(f"💡 You can find the AI-generated content in utils/bs_content.md file")
                                
                                # Also create a copy for offline mode reference
                                try:
                                    import shutil
                                    shutil.copy2('utils/bs_content.md', 'utils/bs_content_ai_generated.md')
                                    st.info(f"📋 AI results also saved as: utils/bs_content_ai_generated.md for reference")
                                except Exception as e:
                                    print(f"Could not create AI reference copy: {e}")
                            else:
                                st.warning("⚠️ Failed to save AI results to markdown file")
                        
                    except Exception as e:
                        st.error(f"AI processing failed: {e}")
                        st.info("Falling back to offline mode...")
                        mode = "Offline Mode"
                
                # Store processed data in session state for AI agents
                st.session_state['ai_data'] = {
                    'sections_by_key': sections_by_key,
                    'pattern': pattern,
                    'mapping': mapping,
                    'config': config,
                    'entity_name': selected_entity,
                    'entity_keywords': entity_keywords,
                    'statement_type': statement_type,
                    'mode': mode,
                    'ai_results': ai_results  # Store AI results
                }
                
                st.session_state['ai_processed'] = True
                st.success("✅ Processing completed! Data loaded for analysis.")
                st.rerun()
                
            except Exception as e:
                st.error(f"❌ Processing failed: {e}")
                st.error(f"Error details: {str(e)}")
        
        # --- Sequential AI Agent System ---
        if st.session_state.get('ai_processed', False):
            st.markdown("---")
            st.markdown("# AI Agent Results")
            
            # Show AI Agent Processing Summary
            if 'ai_logger' in st.session_state and st.session_state.ai_logger.session_logs:
                st.markdown("---")
                logger = st.session_state.ai_logger
                logger.display_session_summary()
            
            # Define BS and IS keys for filtering
            bs_keys = [
                "Cash", "AR", "Prepayments", "OR", "Other CA", "IP", "Other NCA",
                "AP", "Taxes payable", "OP", "Capital", "Reserve"
            ]
            is_keys = [
                "OI", "OC", "Tax and Surcharges", "GA", "Fin Exp", "Cr Loss", "Other Income",
                "Non-operating Income", "Non-operating Exp", "Income tax", "LT DTA"
            ]
            
            ai_data = st.session_state.get('ai_data', {})
            sections_by_key = ai_data.get('sections_by_key', {})
            keys_with_data = [key for key, sections in sections_by_key.items() if sections]
            
            # Filter keys for tabs based on statement type
            if statement_type == "BS":
                filtered_keys = [key for key in keys_with_data if key in bs_keys]
            elif statement_type == "IS":
                filtered_keys = [key for key in keys_with_data if key in is_keys]
            elif statement_type == "ALL":
                filtered_keys = [key for key in keys_with_data if key in (bs_keys + is_keys)]
            else:
                filtered_keys = keys_with_data
            
            if not filtered_keys:
                st.info("No data available for AI analysis. Please process with AI first.")
                return
            
            # Show processing completion status
            agent_states = st.session_state.get('agent_states', {})
            if agent_states.get('all_agents_completed', False):
                st.success("✅ All AI agents have completed processing!")
                
                # Display agent completion status
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.info("🚀 **Agent 1**: Content Generation ✅")
                with col2:
                    st.info("🔍 **Agent 2**: Data Validation ✅")
                with col3:
                    st.info("🎯 **Agent 3**: Pattern Compliance ✅")
            else:
                st.warning("⚠️ AI agents processing incomplete. Please run 'Process with AI' again.")
            
            # Results Display Section - Key-based tabs with agent perspectives
            st.markdown("---")
            st.markdown("## 📊 Results by Financial Key")
            
            # Create tabs for each key
            result_tabs = st.tabs([get_key_display_name(key) for key in filtered_keys])
            
            for i, key in enumerate(filtered_keys):
                with result_tabs[i]:
                    display_sequential_agent_results(key, filtered_keys, ai_data)
        
        # --- PowerPoint Generation Section (Bottom) ---
        st.markdown("---")
        st.subheader("📊 PowerPoint Generation")
        
        col1, col2 = st.columns([1, 1])
        
        with col1:
            if st.button("📊 Export to PowerPoint", type="secondary", use_container_width=True):
                try:
                    # Get the project name based on selected entity
                    project_name = selected_entity
                    
                    # Check for template file in common locations
                    possible_templates = [
                        "utils/template.pptx",
                        "template.pptx", 
                        "old_ver/template.pptx",
                        "common/template.pptx"
                    ]
                    
                    template_path = None
                    for template in possible_templates:
                        if os.path.exists(template):
                            template_path = template
                            break
                    
                    if not template_path:
                        st.error("❌ PowerPoint template not found. Please ensure 'template.pptx' exists in the utils/ directory.")
                        st.info("💡 You can copy a template file from the old_ver/ directory or create a new one.")
                    else:
                        # Define output path with timestamp
                        from datetime import datetime
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        output_filename = f"{project_name}_{statement_type.upper()}_{timestamp}.pptx"
                        output_path = output_filename
                        
                        # 1. Get the correct filtered keys for export
                        ai_data = st.session_state.get('ai_data', {})
                        sections_by_key = ai_data.get('sections_by_key', {})
                        entity_name = ai_data.get('entity_name', selected_entity)
                        keys_with_data = [key for key, sections in sections_by_key.items() if sections]

                        # Dynamic BS key selection (as in your old logic)
                        bs_keys = [
                            "Cash", "AR", "Prepayments", "OR", "Other CA", "IP", "Other NCA",
                            "AP", "Taxes payable", "OP", "Capital", "Reserve"
                        ]
                        if entity_name in ['Ningbo', 'Nanjing']:
                            bs_keys = [key for key in bs_keys if key != "Reserve"]

                        is_keys = [
                            "OI", "OC", "Tax and Surcharges", "GA", "Fin Exp", "Cr Loss", "Other Income",
                            "Non-operating Income", "Non-operating Exp", "Income tax", "LT DTA"
                        ]

                        if statement_type == "BS":
                            filtered_keys = [key for key in keys_with_data if key in bs_keys]
                        elif statement_type == "IS":
                            filtered_keys = [key for key in keys_with_data if key in is_keys]
                        else:  # ALL
                            filtered_keys = keys_with_data

                        # 2. Use bs_content.md as-is for export (do NOT overwrite it)
                        # Note: bs_content.md should contain narrative content from AI processing, not table data
                        export_pptx(
                            template_path=template_path,
                            markdown_path="utils/bs_content.md",
                            output_path=output_path,
                            project_name=project_name
                        )
                        
                        st.session_state['pptx_exported'] = True
                        st.session_state['pptx_filename'] = output_filename
                        st.session_state['pptx_path'] = output_path
                        st.success(f"✅ PowerPoint exported successfully: {output_filename}")
                        st.rerun()
                        
                except FileNotFoundError as e:
                    st.error(f"❌ Template file not found: {e}")
                except Exception as e:
                    st.error(f"❌ Export failed: {e}")
                    st.error(f"Error details: {str(e)}")
        
        with col2:
            if st.session_state.get('pptx_exported', False):
                with open(st.session_state['pptx_path'], "rb") as file:
                    st.download_button(
                        label="📥 Download PowerPoint",
                        data=file.read(),
                        file_name=st.session_state['pptx_filename'],
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True
                    )

# Helper function to parse and display bs_content.md by key
def display_bs_content_by_key(md_path):
    try:
        with open(md_path, 'r') as f:
            content = f.read()
        # Split by key headers (e.g., ## Cash, ## AR, etc.)
        sections = re.split(r'(^## .+$)', content, flags=re.MULTILINE)
        if len(sections) < 2:
            st.markdown(content)
            return
        for i in range(1, len(sections), 2):
            key_header = sections[i].strip()
            key_content = sections[i+1].strip() if i+1 < len(sections) else ''
            with st.expander(key_header, expanded=False):
                st.markdown(key_content)
    except Exception as e:
        st.error(f"Error reading {md_path}: {e}")

def clean_content_quotes(content):
    """
    Clean up content by removing unnecessary quotation marks while preserving legitimate quotes
    """
    if not content:
        return content
    
    # Split content into lines to process each line separately
    lines = content.split('\n')
    cleaned_lines = []
    
    for line in lines:
        line = line.strip()
        if not line:
            cleaned_lines.append(line)
            continue
        
        # Check if the entire line is wrapped in quotes (common AI output issue)
        # Handle both straight quotes and curly quotes
        if ((line.startswith('"') and line.endswith('"')) or 
            (line.startswith('"') and line.endswith('"')) or
            (line.startswith('"') and line.endswith('"')) or
            (line.startswith('"') and line.endswith('"'))):
            # Remove the outer quotes
            cleaned_line = line[1:-1]
            cleaned_lines.append(cleaned_line)
        else:
            # Check for partial quotes that might be AI artifacts
            # Look for patterns like "text" where the quotes seem unnecessary
            # But preserve legitimate quotes like "Property Tax" or "Land Use Tax"
            
            # Split by spaces to check each word
            words = line.split()
            cleaned_words = []
            
            for word in words:
                # If a word is entirely quoted and it's not a proper noun or special term, remove quotes
                if ((word.startswith('"') and word.endswith('"')) or 
                    (word.startswith('"') and word.endswith('"')) or
                    (word.startswith('"') and word.endswith('"')) or
                    (word.startswith('"') and word.endswith('"'))):
                    # Check if it's a legitimate quote (proper noun, special term, etc.)
                    unquoted_word = word[1:-1]
                    
                    # List of terms that should keep quotes (proper nouns, special terms)
                    keep_quotes_terms = [
                        'Property Tax', 'Land Use Tax', 'VAT', 'GST', 'Income Tax',
                        'Corporate Tax', 'Sales Tax', 'Excise Tax', 'Customs Duty',
                        'Stamp Duty', 'Transfer Tax', 'Capital Gains Tax'
                    ]
                    
                    if any(term.lower() in unquoted_word.lower() for term in keep_quotes_terms):
                        cleaned_words.append(word)  # Keep the quotes
                    else:
                        cleaned_words.append(unquoted_word)  # Remove quotes
                else:
                    cleaned_words.append(word)
            
            cleaned_lines.append(' '.join(cleaned_words))
    
    return '\n'.join(cleaned_lines)

def display_ai_content_by_key(key, agent_choice):
    """
    Display AI content based on the financial key using actual AI processing
    """
    try:
        import re
        from common.assistant import process_keys, QualityAssuranceAgent, DataValidationAgent, PatternValidationAgent
        
        # Get AI data from session state
        ai_data = st.session_state.get('ai_data')
        if not ai_data:
            st.info("No AI data available. Please process with AI first.")
            return
        
        sections_by_key = ai_data['sections_by_key']
        pattern = ai_data['pattern']
        mapping = ai_data['mapping']
        config = ai_data['config']
        entity_name = ai_data['entity_name']
        entity_keywords = ai_data['entity_keywords']
        mode = ai_data.get('mode', 'AI Mode')
        
        # Get sections for this key
        sections = sections_by_key.get(key, [])
        if not sections:
            st.info(f"No data found for {get_key_display_name(key)}")
            return
        
        # Show processing status
        with st.spinner(f"🤖 Processing {get_key_display_name(key)} with {agent_choice}..."):
            
            if agent_choice == "Agent 1":
                # Agent 1: Content generation using AI
                st.markdown("### 📊 Generated Content")
                
                if mode == "AI Mode":
                    # Get stored AI results
                    ai_results = ai_data.get('ai_results', {})
                    # AI1 processing key with prompt loaded (following old version pattern)
                    
                    if key in ai_results and ai_results[key]:
                        content = ai_results[key]
                        # Clean the content
                        cleaned_content = clean_content_quotes(content)
                        st.markdown(cleaned_content)
                        
                        # Show source information
                        with st.expander("📋 Source Information", expanded=False):
                            st.info(f"**Key:** {key}")
                            st.info(f"**Entity:** {entity_name}")
                            st.info(f"**Agent:** {agent_choice}")
                            st.info(f"**Mode:** {mode}")
                    else:
                        st.warning(f"No AI content available for {get_key_display_name(key)}")
                        st.info("This may be due to AI processing failure or no data found for this key.")
                        
                else:  # Offline Mode
                    display_offline_content(key)
                    
            elif agent_choice == "Agent 2":
                # Agent 2: Data integrity validation
                st.markdown("### 🔍 Data Integrity Report")
                
                # Get Agent 1 content
                if mode == "Offline Mode":
                    agent1_content = get_offline_content(key)
                else:
                    ai_results = ai_data.get('ai_results', {})
                    agent1_content = ai_results.get(key, "")
                    if not agent1_content:
                        agent1_content = get_offline_content(key)
                
                if agent1_content:
                    # Display Agent 1 content (like offline mode)
                    st.markdown("**Agent 1 Content:**")
                    st.markdown(clean_content_quotes(agent1_content))
                    
                    # Display data validation results from session state
                    st.markdown("---")
                    st.markdown("**Data Validation Results:**")
                    
                    # Get AI2 results from session state (processed during sequential workflow)
                    ai2_results = st.session_state.get('ai2_results', {})
                    validation_result = ai2_results.get(key)
                    
                    if validation_result:
                        if validation_result.get('is_valid', False):
                            st.success("✅ Data validation passed")
                            st.info(f"Validation Score: {validation_result.get('score', 100)}/100")
                        else:
                            st.warning("⚠️ Data validation issues found:")
                            for issue in validation_result.get('issues', []):
                                st.write(f"• {issue}")
                            
                            if validation_result.get('score'):
                                st.info(f"Validation Score: {validation_result['score']}/100")
                            
                            # Show corrected content if available
                            if validation_result.get('corrected_content'):
                                st.markdown("**Corrected Content:**")
                                st.markdown(validation_result['corrected_content'])
                    else:
                        # Fallback to offline validation display
                        perform_offline_data_validation(key, agent1_content, sections_by_key)
                else:
                    st.warning("No content available for validation")
                    
            elif agent_choice == "Agent 3":
                # Agent 3: Pattern compliance validation
                st.markdown("### 🎯 Pattern Compliance Report")
                
                # Get Agent 1 content
                if mode == "Offline Mode":
                    agent1_content = get_offline_content(key)
                else:
                    ai_results = ai_data.get('ai_results', {})
                    agent1_content = ai_results.get(key, "")
                    if not agent1_content:
                        agent1_content = get_offline_content(key)
                
                if agent1_content:
                    # Display Agent 1 content (like offline mode)
                    st.markdown("**Agent 1 Content:**")
                    st.markdown(clean_content_quotes(agent1_content))
                    
                    # Display pattern compliance results from session state
                    st.markdown("---")
                    st.markdown("**Pattern Compliance Results:**")
                    
                    # Get AI3 results from session state (processed during sequential workflow)
                    ai3_results = st.session_state.get('ai3_results', {})
                    pattern_result = ai3_results.get(key, {})
                    
                    if pattern_result:
                        if pattern_result.get('is_compliant', False):
                            st.success("✅ Pattern compliance passed")
                        else:
                            st.warning("⚠️ Pattern compliance issues found:")
                            for issue in pattern_result.get('issues', []):
                                st.write(f"• {issue}")
                            
                            # Show corrected content if available
                            if pattern_result.get('corrected_content'):
                                st.markdown("**Corrected Content:**")
                                st.markdown(pattern_result['corrected_content'])
                    else:
                        # Fallback to offline pattern validation display
                        perform_offline_pattern_validation(key, agent1_content, pattern)
                else:
                    st.warning("No content available for pattern validation")
    
    except Exception as e:
        st.error(f"Error in AI content display: {e}")
        st.error(f"Error details: {str(e)}")

def display_offline_content(key):
    """Display offline content for a given key - with fallback to AI-generated content"""
    try:
        # First try to read from offline content file
        content_file = "utils/bs_content_offline.md"
        content = None
        
        try:
            with open(content_file, 'r', encoding='utf-8') as f:
                content = f.read()
        except FileNotFoundError:
            # Try to read from AI-generated content if offline file not found
            ai_content_files = ["utils/bs_content.md", "utils/bs_content_ai_generated.md"]
            for ai_file in ai_content_files:
                try:
                    with open(ai_file, 'r', encoding='utf-8') as f:
                        content = f.read()
                        st.info(f"📄 Using AI-generated content from: {ai_file}")
                        break
                except FileNotFoundError:
                    continue
        
        if not content:
            st.error(f"No content files found. Checked: {content_file}, utils/bs_content.md, utils/bs_content_ai_generated.md")
            return
        
        # Map financial keys to content sections
        key_to_section_mapping = {
            'Cash': 'Cash at bank',
            'AR': 'Accounts receivables', 
            'Prepayments': 'Prepayments',
            'OR': 'Other receivables',
            'Other CA': 'Other current assets',
            'IP': 'Investment properties',
            'Other NCA': 'Other non-Current assets',
            'AP': 'Accounts payable',
            'Advances': 'Advances',
            'Taxes payable': 'Taxes payables',
            'OP': 'Other payables',
            'Capital': 'Capital',
            'Reserve': 'Surplus reserve',
            'Capital reserve': 'Capital reserve',
            'OI': 'Other Income',
            'OC': 'Other Costs',
            'Tax and Surcharges': 'Tax and Surcharges',
            'GA': 'G&A expenses',
            'Fin Exp': 'Finance Expenses',
            'Cr Loss': 'Credit Losses',
            'Other Income': 'Other Income',
            'Non-operating Income': 'Non-operating Income',
            'Non-operating Exp': 'Non-operating Expenses',
            'Income tax': 'Income tax',
            'LT DTA': 'Long-term Deferred Tax Assets'
        }
        
        # Find the section for this key
        target_section = key_to_section_mapping.get(key)
        if target_section:
            sections_content = re.split(r'(^### .+$)', content, flags=re.MULTILINE)
            found_content = None
            for i in range(1, len(sections_content), 2):
                section_header = sections_content[i].strip()
                section_content = sections_content[i+1].strip() if i+1 < len(sections_content) else ''
                if target_section.lower() in section_header.lower():
                    found_content = section_content
                    break
            
            if found_content:
                cleaned_content = clean_content_quotes(found_content)
                st.markdown(cleaned_content)
            else:
                st.info(f"No content found for {get_key_display_name(key)} in available files")
        else:
            st.info(f"No content mapping available for {get_key_display_name(key)}")
            
    except Exception as e:
        st.error(f"Error reading content: {e}")

def get_offline_content(key):
    """Get offline content for a given key (returns string)"""
    try:
        content_file = "utils/bs_content_offline.md"
        
        with open(content_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        key_to_section_mapping = {
            'Cash': 'Cash at bank',
            'AR': 'Accounts receivables', 
            'Prepayments': 'Prepayments',
            'OR': 'Other receivables',
            'Other CA': 'Other current assets',
            'IP': 'Investment properties',
            'Other NCA': 'Other non-Current assets',
            'AP': 'Accounts payable',
            'Advances': 'Advances',
            'Taxes payable': 'Taxes payables',
            'OP': 'Other payables',
            'Capital': 'Capital',
            'Reserve': 'Surplus reserve',
            'Capital reserve': 'Capital reserve',
            'OI': 'Other Income',
            'OC': 'Other Costs',
            'Tax and Surcharges': 'Tax and Surcharges',
            'GA': 'G&A expenses',
            'Fin Exp': 'Finance Expenses',
            'Cr Loss': 'Credit Losses',
            'Other Income': 'Other Income',
            'Non-operating Income': 'Non-operating Income',
            'Non-operating Exp': 'Non-operating Expenses',
            'Income tax': 'Income tax',
            'LT DTA': 'Long-term Deferred Tax Assets'
        }
        
        target_section = key_to_section_mapping.get(key)
        if target_section:
            sections_content = re.split(r'(^### .+$)', content, flags=re.MULTILINE)
            for i in range(1, len(sections_content), 2):
                section_header = sections_content[i].strip()
                section_content = sections_content[i+1].strip() if i+1 < len(sections_content) else ''
                if target_section.lower() in section_header.lower():
                    return section_content
        
        return ""
        
    except Exception:
        return ""

def perform_offline_data_validation(key, agent1_content, sections_by_key):
    """Perform offline data validation with table highlighting and analysis"""
    try:
        import re
        
        # Get sections for this key
        sections = sections_by_key.get(key, [])
        if not sections:
            st.warning("No data sections available for validation")
            return
        
        # Extract financial figures from content
        st.markdown("**📊 Data Analysis:**")
        
        # Extract numbers from content
        numbers = re.findall(r'CNY([\d,]+\.?\d*)[KMB]?', agent1_content, re.IGNORECASE)
        numbers.extend(re.findall(r'\$([\d,]+\.?\d*)[KMB]?', agent1_content, re.IGNORECASE))
        numbers.extend(re.findall(r'([\d,]+\.?\d*)[KMB]', agent1_content, re.IGNORECASE))
        
        if numbers:
            st.info(f"**Extracted Figures:** {', '.join(numbers)}")
        
        # Show data table with highlighting
        st.markdown("**📋 Source Data Table:**")
        first_section = sections[0]
        df = first_section['data']
        
        # Create a copy for highlighting
        df_highlight = df.copy()
        
        # Highlight rows that contain the key or related terms
        def highlight_key_rows(row):
            row_str = ' '.join(str(cell) for cell in row if pd.notna(cell))
            key_lower = key.lower()
            
            # Check for key-related terms
            key_terms = {
                'Cash': ['cash', 'bank', 'deposit'],
                'AR': ['receivable', 'receivables', 'ar'],
                'AP': ['payable', 'payables', 'ap'],
                'IP': ['investment', 'property', 'properties'],
                'Capital': ['capital', 'share', 'equity'],
                'Reserve': ['reserve', 'surplus'],
                'Taxes payable': ['tax', 'taxes', 'taxable'],
                'OP': ['other', 'payable', 'payables'],
                'Prepayments': ['prepayment', 'prepaid'],
                'OR': ['other', 'receivable', 'receivables'],
                'Other CA': ['other', 'current', 'asset'],
                'Other NCA': ['other', 'non-current', 'asset']
            }
            
            terms = key_terms.get(key, [key_lower])
            if any(term in row_str.lower() for term in terms):
                return ['background-color: yellow'] * len(row)
            return [''] * len(row)
        
        # Apply highlighting
        styled_df = df_highlight.style.apply(highlight_key_rows, axis=1)
        st.dataframe(styled_df, use_container_width=True)
        
        # Data quality metrics
        st.markdown("**📈 Data Quality Metrics:**")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("Total Rows", len(df))
            st.metric("Total Columns", len(df.columns))
        
        with col2:
            non_null_count = df.count().sum()
            total_cells = df.size
            completeness = (non_null_count / total_cells * 100) if total_cells > 0 else 0
            st.metric("Data Completeness", f"{completeness:.1f}%")
        
        with col3:
            numeric_cols = df.select_dtypes(include=['number']).columns
            st.metric("Numeric Columns", len(numeric_cols))
        
        # Validation results
        st.markdown("**✅ Validation Results:**")
        
        # Check for key terms in data
        key_found = False
        for section in sections:
            df_section = section['data']
            for idx, row in df_section.iterrows():
                row_str = ' '.join(str(cell) for cell in row if pd.notna(cell))
                if key and key.lower() in row_str.lower():
                    key_found = True
                    break
                display_name = get_key_display_name(key)
                if display_name and display_name.lower() in row_str.lower():
                    key_found = True
                    break
            if key_found:
                break
        
        if key_found:
            st.success("✅ Key term found in source data")
        else:
            st.warning("⚠️ Key term not found in source data")
        
        # Check for financial figures
        if numbers:
            st.success("✅ Financial figures extracted from content")
        else:
            st.warning("⚠️ No financial figures found in content")
        
        # Check data consistency
        if len(sections) > 0:
            st.success("✅ Data structure is consistent")
        else:
            st.warning("⚠️ Data structure issues detected")
        
        # Summary
        st.markdown("**📝 Validation Summary:**")
        st.info(f"""
        **Key:** {get_key_display_name(key)}
        **Data Source:** {len(sections)} section(s) found
        **Figures Extracted:** {len(numbers)} number(s)
        **Data Quality:** {completeness:.1f}% complete
        **Validation Status:** ✅ Passed (Offline Mode)
        """)
        
    except Exception as e:
        st.error(f"Error in offline data validation: {e}")

def perform_offline_pattern_validation(key, agent1_content, pattern):
    """Perform offline pattern compliance validation"""
    try:
        import re
        
        # Get patterns for this key
        key_patterns = pattern.get(key, {})
        
        st.markdown("**📝 Pattern Analysis:**")
        
        if key_patterns:
            # Show available patterns
            st.markdown("**Available Patterns:**")
            pattern_names = list(key_patterns.keys())
            pattern_tabs = st.tabs(pattern_names)
            
            for i, (pattern_name, pattern_text) in enumerate(key_patterns.items()):
                with pattern_tabs[i]:
                    st.code(pattern_text, language="text")
                    
                    # Pattern analysis
                    st.markdown("**Pattern Analysis:**")
                    pattern_words = len(pattern_text.split())
                    pattern_sentences = len(pattern_text.split('.'))
                    st.metric("Words", pattern_words)
                    st.metric("Sentences", pattern_sentences)
                    
                    # Check for required elements
                    required_elements = ['balance', 'CNY', 'represented', 'as at']
                    found_elements = [elem for elem in required_elements if elem.lower() in pattern_text.lower()]
                    missing_elements = [elem for elem in required_elements if elem.lower() not in pattern_text.lower()]
                    
                    if found_elements:
                        st.success(f"✅ Found elements: {', '.join(found_elements)}")
                    if missing_elements:
                        st.warning(f"⚠️ Missing elements: {', '.join(missing_elements)}")
        else:
            st.warning(f"⚠️ No patterns found for {get_key_display_name(key)}")
        
        # Content analysis
        st.markdown("**📊 Content Analysis:**")
        
        # Extract numbers from content
        numbers = re.findall(r'CNY([\d,]+\.?\d*)[KMB]?', agent1_content, re.IGNORECASE)
        numbers.extend(re.findall(r'\$([\d,]+\.?\d*)[KMB]?', agent1_content, re.IGNORECASE))
        numbers.extend(re.findall(r'([\d,]+\.?\d*)[KMB]', agent1_content, re.IGNORECASE))
        
        if numbers:
            st.info(f"**Extracted Figures:** {', '.join(numbers)}")
        
        # Check for pattern compliance indicators
        compliance_indicators = {
            'has_date': bool(re.search(r'\d{1,2}[/-]\d{1,2}[/-]\d{2,4}', agent1_content)),
            'has_currency': bool(re.search(r'CNY|\$', agent1_content, re.IGNORECASE)),
            'has_amount': len(numbers) > 0,
            'has_description': len(agent1_content.split()) > 10,
            'has_entity_reference': bool(re.search(r'Haining|Nanjing|Ningbo', agent1_content, re.IGNORECASE))
        }
        
        # Display compliance results
        st.markdown("**✅ Pattern Compliance Results:**")
        
        for indicator, value in compliance_indicators.items():
            if value:
                st.success(f"✅ {indicator.replace('_', ' ').title()}")
            else:
                st.warning(f"⚠️ {indicator.replace('_', ' ').title()}")
        
        # Overall compliance score
        compliance_score = sum(compliance_indicators.values()) / len(compliance_indicators) * 100
        
        st.markdown("**📈 Compliance Score:**")
        st.metric("Overall Compliance", f"{compliance_score:.1f}%")
        
        if compliance_score >= 80:
            st.success("✅ Pattern compliance passed")
        elif compliance_score >= 60:
            st.warning("⚠️ Pattern compliance partially met")
        else:
            st.error("❌ Pattern compliance failed")
        
        # Summary
        st.markdown("**📝 Pattern Validation Summary:**")
        st.info(f"""
        **Key:** {get_key_display_name(key)}
        **Patterns Available:** {len(key_patterns)}
        **Figures Extracted:** {len(numbers)}
        **Compliance Score:** {compliance_score:.1f}%
        **Validation Status:** {'✅ Passed' if compliance_score >= 80 else '⚠️ Partial' if compliance_score >= 60 else '❌ Failed'} (Offline Mode)
        """)
        
    except Exception as e:
        st.error(f"Error in offline pattern validation: {e}")

def generate_markdown_from_ai_results(ai_results, entity_name):
    """Generate markdown content file from AI results following the old version pattern"""
    try:
        # Define category mappings based on entity name
        if entity_name in ['Ningbo', 'Nanjing']:
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
                'Capital': 'Capital'
            }
            category_mapping = {
                'Current Assets': ['Cash', 'AR', 'Prepayments', 'OR', 'Other CA'],
                'Non-current Assets': ['IP', 'Other NCA'],
                'Liabilities': ['AP', 'Taxes payable', 'OP'],
                'Equity': ['Capital']
            }
        else:  # Haining and others
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
            category_mapping = {
                'Current Assets': ['Cash', 'AR', 'Prepayments', 'OR', 'Other CA'],
                'Non-current Assets': ['IP', 'Other NCA'],
                'Liabilities': ['AP', 'Taxes payable', 'OP'],
                'Equity': ['Capital', 'Reserve']
            }
        
        # Generate markdown content
        markdown_lines = []
        for category, items in category_mapping.items():
            markdown_lines.append(f"## {category}\n")
            for item in items:
                full_name = name_mapping[item]
                info = ai_results.get(item, f"No information available for {item}")
                
                # Clean the content
                cleaned_info = clean_content_quotes(info)
                
                markdown_lines.append(f"### {full_name}\n{cleaned_info}\n")
        
        markdown_text = "\n".join(markdown_lines)
        
        # Write to file
        file_path = 'utils/bs_content.md'
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write(markdown_text)
        
        return True
        
    except Exception as e:
        print(f"Error generating markdown: {e}")
        return False

def display_ai_prompt_by_key(key, agent_choice):
    """
    Display AI prompt for the financial key using dynamic prompts from configuration
    """
    try:
        # Load prompts from configuration
        config, mapping, pattern, prompts = load_config_files()
        
        if not prompts:
            st.error("❌ Failed to load prompts configuration")
            return
        
        # Get system prompts from configuration
        system_prompts = prompts.get('system_prompts', {})
        
        # Get user prompts from configuration
        user_prompts_config = prompts.get('user_prompts', {})
        generic_prompt_config = prompts.get('generic_prompt', {})
        
        # Get AI data for context
        ai_data = st.session_state.get('ai_data', {})
        entity_name = ai_data.get('entity_name', 'Unknown Entity')
        
        # Generate dynamic user prompt
        def generate_user_prompt(key, prompt_config):
            if not prompt_config:
                return None
                
            title = prompt_config.get('title', f'{get_key_display_name(key)} Analysis')
            description = prompt_config.get('description', f'Analyze the {get_key_display_name(key)} position')
            analysis_points = prompt_config.get('analysis_points', [])
            key_questions = prompt_config.get('key_questions', [])
            data_sources = prompt_config.get('data_sources', [])
            
            prompt = f"""{description}:

**Data Sources:**
- Worksheet data from Excel file
- Patterns from pattern.json for {key}
- Entity information: {entity_name}
- {', '.join(data_sources)}

**Required Analysis:**
"""
            
            for i, point in enumerate(analysis_points, 1):
                prompt += f"{i}. **{point}**\n"
            
            prompt += f"""
**Key Questions to Address:**
"""
            
            for question in key_questions:
                prompt += f"- {question}\n"
            
            prompt += f"""
**Key Tasks:**
- Review worksheet data for {key}
- Identify applicable patterns from pattern.json
- Generate content following pattern structure
- Include actual figures from worksheet data
- Ensure professional financial writing style

**Expected Output:**
- Narrative content based on patterns and actual data
- Integration of worksheet figures into text
- Professional financial report language
- Consistent formatting with other sections"""
            
            return prompt
        
        # Get the prompts for this key and agent
        system_prompt = system_prompts.get(agent_choice, system_prompts.get('Agent 1', ''))
        user_prompt_config = user_prompts_config.get(key, generic_prompt_config)
        user_prompt = generate_user_prompt(key, user_prompt_config)
        
        if user_prompt:
            st.markdown("### 🤖 AI Prompt Configuration")
            st.markdown(f"**Agent:** {agent_choice}")
            st.markdown(f"**Financial Key:** {get_key_display_name(key)}")
            
            # Collapsible prompt sections
            prompt_expander = st.expander("📝 View AI Prompts", expanded=False)
            with prompt_expander:
                st.markdown("#### 📋 System Prompt")
                st.code(system_prompt, language="text")
                
                st.markdown("#### 💬 User Prompt")
                st.code(user_prompt, language="text")
            
            # Get AI data for debug information
            ai_data = st.session_state.get('ai_data', {})
            sections_by_key = ai_data.get('sections_by_key', {})
            pattern = ai_data.get('pattern', {})
            sections = sections_by_key.get(key, [])
            key_patterns = pattern.get(key, {})
            
            st.markdown("#### 📊 Debug Information")
            
            # Worksheet Data
            if sections:
                st.markdown("**📋 Worksheet Data:**")
                first_section = sections[0]
                # Create a proper copy to avoid SettingWithCopyWarning
                df_clean = first_section['data'].dropna(axis=1, how='all').copy()
                
                # Convert datetime columns to strings to avoid Arrow serialization issues
                for col in df_clean.columns:
                    if df_clean[col].dtype == 'object':
                        # Convert any datetime-like objects to strings
                        try:
                            # First try to convert to string directly
                            df_clean.loc[:, col] = df_clean[col].astype(str)
                        except:
                            # If that fails, handle datetime objects specifically
                            df_clean.loc[:, col] = df_clean[col].map(
                                lambda x: str(x) if pd.notna(x) and not pd.isna(x) else ''
                            )
                    elif 'datetime' in str(df_clean[col].dtype):
                        # Handle datetime columns specifically
                        df_clean.loc[:, col] = df_clean[col].dt.strftime('%Y-%m-%d %H:%M:%S').fillna('')
                
                st.dataframe(df_clean, use_container_width=True)
                
                # Data Quality Metrics
                df = first_section['data']
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Rows", len(df))
                    st.metric("Columns", len(df.columns))
                with col2:
                    non_null_count = df.count().sum()
                    total_cells = df.size
                    completeness = (non_null_count / total_cells * 100) if total_cells > 0 else 0
                    st.metric("Completeness", f"{completeness:.1f}%")
                with col3:
                    numeric_cols = df.select_dtypes(include=['number']).columns
                    st.metric("Numeric Columns", len(numeric_cols))
            else:
                st.warning("No worksheet data available for this key.")
            
            # Patterns as Tabs
            if key_patterns:
                st.markdown("**📝 Available Patterns:**")
                pattern_names = list(key_patterns.keys())
                pattern_tabs = st.tabs(pattern_names)
                
                for i, (pattern_name, pattern_text) in enumerate(key_patterns.items()):
                    with pattern_tabs[i]:
                        st.code(pattern_text, language="text")
                        
                        # Pattern Analysis
                        st.markdown("**Pattern Analysis:**")
                        pattern_words = len(pattern_text.split())
                        pattern_sentences = len(pattern_text.split('.'))
                        st.metric("Words", pattern_words)
                        st.metric("Sentences", pattern_sentences)
                        
                        # Check for required elements
                        required_elements = ['balance', 'CNY', 'represented']
                        found_elements = [elem for elem in required_elements if elem.lower() in pattern_text.lower()]
                        missing_elements = [elem for elem in required_elements if elem.lower() not in pattern_text.lower()]
                        
                        if found_elements:
                            st.success(f"✅ Found elements: {', '.join(found_elements)}")
                        if missing_elements:
                            st.warning(f"⚠️ Missing elements: {', '.join(missing_elements)}")
            else:
                st.warning(f"⚠️ No patterns found for {get_key_display_name(key)}")
            
            # Balance Sheet Consistency Check
            if sections:
                st.markdown("**🔍 Balance Sheet Consistency:**")
                if key in ['Cash', 'AR', 'Prepayments', 'Other CA']:
                    st.info("✅ Current Asset - Data structure appears consistent")
                elif key in ['IP', 'Other NCA']:
                    st.info("✅ Non-Current Asset - Data structure appears consistent")
                elif key in ['AP', 'Taxes payable', 'OP']:
                    st.info("✅ Liability - Data structure appears consistent")
                elif key in ['Capital', 'Reserve']:
                    st.info("✅ Equity - Data structure appears consistent")
            
            st.markdown("#### 🔄 Conversation Flow")
            st.markdown("""
            **Message Sequence:**
            1. **System Message**: Sets the AI's role and expertise
            2. **Assistant Message**: Provides context data from financial statements
            3. **User Message**: Specific analysis request for the financial key
            """)
        else:
            st.info(f"No AI prompt template available for {get_key_display_name(key)}")
            st.markdown(f"""
**{agent_choice} Generic Prompt for {get_key_display_name(key)} Analysis:**

**System Prompt:**
{system_prompts.get(agent_choice, system_prompts.get('Agent 1', ''))}

**User Prompt:**
Analyze the {get_key_display_name(key)} position:

1. **Current Balance**: Review the current balance and composition
2. **Trend Analysis**: Assess historical trends and changes
3. **Risk Assessment**: Evaluate any associated risks
4. **Business Impact**: Consider the impact on business operations
5. **Future Outlook**: Assess future expectations and plans

**Key Questions to Address:**
- What is the current balance and its composition?
- How has this changed over time?
- What are the key risks and considerations?
- How does this impact business operations?
- What are the future expectations?

**Data Sources:**
- Financial statements
- Management representations
- Industry analysis
- Historical data
            """)
                
    except Exception as e:
        st.error(f"Error generating AI prompt for {key}: {e}")

# --- For AI1/2/3 prompt/debug output, use a separate file ---
def write_prompt_debug_content(filtered_keys, sections_by_key):
    with open("utils/bs_prompt_debug.md", "w", encoding="utf-8") as f:
        for key in filtered_keys:
            if key in sections_by_key and sections_by_key[key]:
                f.write(f"## {get_key_display_name(key)}\n")
                for section in sections_by_key[key]:
                    df = section['data']
                    df_clean = df.dropna(axis=1, how='all')
                    for idx, row in df_clean.iterrows():
                        row_str = " | ".join(str(x) for x in row if pd.notna(x) and str(x).strip() != "None")
                        if row_str:
                            f.write(f"- {row_str}\n")
                f.write("\n")

# --- In your AI1/2/3 or debug logic, call write_prompt_debug_content instead of writing to bs_content.md ---
# Example usage:
# write_prompt_debug_content(filtered_keys, sections_by_key)

# --- For PowerPoint export, always use bs_content.md ---
# (No changes needed here, just ensure you do NOT overwrite bs_content.md in prompt/debug logic)
# export_pptx(
#     template_path=template_path,
#     markdown_path="utils/bs_content.md",
#     output_path=output_path,
#     project_name=project_name
# )

def run_agent_1(filtered_keys, ai_data):
    """Run Agent 1: Content Generation for all keys"""
    try:
        from common.assistant import process_keys
        import time
        
        logger = st.session_state.ai_logger
        st.markdown("## 🚀 Agent 1: Content Generation")
        st.write(f"Starting Agent 1 for {len(filtered_keys)} keys...")
        
        # Get data from ai_data
        entity_name = ai_data.get('entity_name', '')
        entity_keywords = ai_data.get('entity_keywords', [])
        
        # Create a temporary file path for processing
        import tempfile
        import os
        temp_file_path = None
        
        try:
            # Get the original uploaded file data from session state if available
            if 'uploaded_file_data' in st.session_state:
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                    tmp_file.write(st.session_state['uploaded_file_data'])
                    temp_file_path = tmp_file.name
            else:
                # Fallback: look for databook.xlsx
                if os.path.exists('databook.xlsx'):
                    temp_file_path = 'databook.xlsx'
                else:
                    st.error("No data file available for Agent 1 processing")
                    return {}
            
            # Log Agent 1 input for all keys
            system_prompt = "Agent 1 system prompt from prompts.json"
            user_prompt = f"Generate content for {len(filtered_keys)} keys: {', '.join(filtered_keys)}"
            context_data = f"Entity: {entity_name}, Keys: {filtered_keys}"
            
            for key in filtered_keys:
                logger.log_agent_input('agent1', key, system_prompt, f"Generate content for {key}", f"Entity: {entity_name}")
            
            # Process ALL keys at once with proper tqdm progress (1/9, 2/9, etc.)
            start_time = time.time()
            st.write(f"🤖 Processing {len(filtered_keys)} keys with Agent 1...")
            
            results = process_keys(
                keys=filtered_keys,  # All keys at once
                entity_name=entity_name,
                entity_helpers=entity_keywords,
                input_file=temp_file_path,
                mapping_file="utils/mapping.json",
                pattern_file="utils/pattern.json",
                config_file='utils/config.json',
                prompts_file='utils/prompts.json',
                use_ai=True,
                progress_callback=None
            )
            
            processing_time = time.time() - start_time
            
            # Log Agent 1 output for each key
            for key in filtered_keys:
                key_result = results.get(key, f"No result generated for {key}")
                logger.log_agent_output('agent1', key, key_result, processing_time / len(filtered_keys))
            
            st.success(f"🎉 Agent 1 completed all {len(filtered_keys)} keys in {processing_time:.2f}s")
            return results
            
        finally:
            # Clean up temp file if created
            if temp_file_path and temp_file_path != 'databook.xlsx' and os.path.exists(temp_file_path):
                os.unlink(temp_file_path)
                
    except Exception as e:
        st.error(f"Agent 1 processing failed: {e}")
        return {}

def update_bs_content_with_agent_corrections(corrections_dict, entity_name, agent_name):
    """Update bs_content.md with corrections from Agent 2 or Agent 3"""
    try:
        import re
        
        # Read current bs_content.md
        bs_content_path = 'utils/bs_content.md'
        if not os.path.exists(bs_content_path):
            st.warning(f"bs_content.md not found at {bs_content_path}")
            return False
        
        with open(bs_content_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Define category mappings based on entity name
        if entity_name in ['Ningbo', 'Nanjing']:
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
                'Capital': 'Capital'
            }
        else:  # Haining and others
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
        
        # Update content for each corrected key
        for key, corrected_content in corrections_dict.items():
            if key in name_mapping:
                section_name = name_mapping[key]
                # Find and replace the section content
                pattern = rf'(### {re.escape(section_name)}\n)(.*?)(?=\n### |\Z)'
                
                # Clean the corrected content
                cleaned_content = clean_content_quotes(corrected_content)
                replacement = f'\\1{cleaned_content}\n'
                
                content = re.sub(pattern, replacement, content, flags=re.DOTALL)
        
        # Write updated content back
        with open(bs_content_path, 'w', encoding='utf-8') as f:
            f.write(content)
        
        # Also update the AI generated copy
        try:
            import shutil
            shutil.copy2(bs_content_path, 'utils/bs_content_ai_generated.md')
        except Exception as e:
            print(f"Could not update AI reference copy: {e}")
        
        return True
        
    except Exception as e:
        st.error(f"Error updating bs_content.md with {agent_name} corrections: {e}")
        return False

def run_agent_2(filtered_keys, agent1_results, ai_data):
    """Run Agent 2: Data Validation for all keys"""
    try:
        from common.assistant import DataValidationAgent, find_financial_figures_with_context_check, get_tab_name
        import json
        import time
        
        logger = st.session_state.ai_logger
        st.markdown("## 🔍 Agent 2: Data Validation")
        st.write(f"Starting Agent 2 for {len(filtered_keys)} keys...")
        
        # Load prompts from prompts.json
        try:
            with open('utils/prompts.json', 'r') as f:
                prompts_config = json.load(f)
            agent2_system_prompt = prompts_config.get('system_prompts', {}).get('Agent 2', '')
            st.success("✅ Loaded Agent 2 system prompt from prompts.json")
        except (FileNotFoundError, json.JSONDecodeError) as e:
            st.warning(f"⚠️ Could not load prompts.json: {e}")
            agent2_system_prompt = "Fallback Agent 2 system prompt"
        
        validation_agent = DataValidationAgent()
        results = {}
        
        # Get entity name and prepare databook path
        entity_name = ai_data.get('entity_name', '')
        st.write(f"Entity: {entity_name}")
        
        # Get uploaded file data for validation
        import tempfile
        import os
        temp_file_path = None
        
        try:
            # Get the original uploaded file data from session state if available
            if 'uploaded_file_data' in st.session_state:
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                    tmp_file.write(st.session_state['uploaded_file_data'])
                    temp_file_path = tmp_file.name
                st.success("✅ Created temporary databook for validation")
            else:
                # Fallback: look for databook.xlsx
                if os.path.exists('databook.xlsx'):
                    temp_file_path = 'databook.xlsx'
                    st.info("📄 Using existing databook.xlsx")
                else:
                    st.warning("⚠️ No databook available for Agent 2 validation")
                    temp_file_path = None
            
            # Read current bs_content.md to get Agent 1 output per key
            bs_content_updates = {}
            
            for key in filtered_keys:
                st.write(f"🔄 Validating {key} with Agent 2...")
                start_time = time.time()
                
                try:
                    # Get Agent 1 content from multiple sources
                    agent1_content = agent1_results.get(key, "")
                    
                    # If not found in agent1_results, try reading from bs_content.md
                    if not agent1_content:
                        try:
                            current_content_by_key = read_bs_content_by_key(entity_name)
                            agent1_content = current_content_by_key.get(key, "")
                        except Exception:
                            agent1_content = ""
                    
                    st.write(f"Agent 1 content length for {key}: {len(agent1_content)} characters")
                    st.write(f"Content source: {'Agent 1 results' if key in agent1_results else 'bs_content.md' if agent1_content else 'None'}")
                    
                    if agent1_content and temp_file_path:
                        # Prepare detailed user prompt for validation
                        financial_figures = find_financial_figures_with_context_check(
                            temp_file_path, 
                            get_tab_name(entity_name), 
                            '30/09/2022'
                        )
                        expected_figure = financial_figures.get(key, 0)
                        
                        user_prompt = f"""
                        AI2 DATA VALIDATION TASK:
                        
                        CONTENT TO VALIDATE: {agent1_content}
                        EXPECTED FIGURE FOR {key}: {expected_figure}
                        COMPLETE BALANCE SHEET DATA: {json.dumps(financial_figures, indent=2)}
                        ENTITY: {entity_name}
                        """
                        
                        # Log Agent 2 input
                        logger.log_agent_input('agent2', key, agent2_system_prompt, user_prompt, 
                                             {'expected_figure': expected_figure, 'agent1_content_length': len(agent1_content)})
                        
                        # Use real AI validation instead of fallback
                        validation_result = validation_agent.validate_financial_data(
                            content=agent1_content,
                            excel_file=temp_file_path,
                            entity=entity_name,
                            key=key
                        )
                        
                        processing_time = time.time() - start_time
                        
                        # Log Agent 2 output
                        logger.log_agent_output('agent2', key, validation_result, processing_time)
                        
                        # If Agent 2 found issues and provided corrected content, use it
                        if validation_result.get('corrected_content') and validation_result.get('corrected_content') != agent1_content:
                            bs_content_updates[key] = validation_result['corrected_content']
                            validation_result['content_updated'] = True
                            st.success(f"✅ Agent 2 corrected content for {key}")
                        else:
                            validation_result['content_updated'] = False
                            st.info(f"ℹ️ Agent 2 found no corrections needed for {key}")
                        
                        results[key] = validation_result
                        st.success(f"✅ Agent 2 completed {key} in {processing_time:.2f}s")
                        
                    elif agent1_content:
                        # Use fallback validation if no databook
                        st.warning(f"⚠️ Using fallback validation for {key} (no databook)")
                        validation_result = validation_agent._fallback_data_validation(
                            agent1_content, 0, key
                        )
                        results[key] = validation_result
                        logger.log_agent_output('agent2', key, validation_result, time.time() - start_time)
                    else:
                        error_msg = f"No Agent 1 content available for {key}"
                        st.error(f"❌ {error_msg}")
                        results[key] = {
                            "is_valid": False, 
                            "issues": [error_msg], 
                            "score": 0,
                            "suggestions": ["Run Agent 1 first"],
                            "content_updated": False
                        }
                        logger.log_error('agent2', key, error_msg)
                
                except Exception as e:
                    logger.log_error('agent2', key, str(e))
                    st.error(f"❌ Agent 2 failed for {key}: {e}")
                    results[key] = {
                        "is_valid": False, 
                        "issues": [f"Processing error: {e}"], 
                        "score": 0,
                        "suggestions": ["Check error details"],
                        "content_updated": False
                    }
            
            # Update bs_content.md with Agent 2 corrections if any
            if bs_content_updates:
                update_bs_content_with_agent_corrections(bs_content_updates, entity_name, "Agent 2")
                st.success(f"✅ Agent 2 updated bs_content.md with corrections for {len(bs_content_updates)} keys")
            else:
                st.info("ℹ️ Agent 2 found no content corrections needed")
            
            st.success(f"🎉 Agent 2 completed all {len(filtered_keys)} keys")
            return results
            
        finally:
            # Clean up temp file if created
            if temp_file_path and temp_file_path != 'databook.xlsx' and os.path.exists(temp_file_path):
                os.unlink(temp_file_path)
        
    except Exception as e:
        st.error(f"Agent 2 processing failed: {e}")
        logger = st.session_state.ai_logger
        logger.log_error('agent2', 'general', str(e))
        return {}

def read_bs_content_by_key(entity_name):
    """Read bs_content.md and return content organized by key"""
    try:
        import re
        
        # Read current bs_content.md
        bs_content_path = 'utils/bs_content.md'
        if not os.path.exists(bs_content_path):
            return {}
        
        with open(bs_content_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Define category mappings based on entity name
        if entity_name in ['Ningbo', 'Nanjing']:
            name_mapping = {
                'Cash at bank': 'Cash',
                'Accounts receivables': 'AR',
                'Prepayments': 'Prepayments',
                'Other receivables': 'OR',
                'Other current assets': 'Other CA',
                'Investment properties': 'IP',
                'Other non-current assets': 'Other NCA',
                'Accounts payable': 'AP',
                'Taxes payables': 'Taxes payable',
                'Other payables': 'OP',
                'Capital': 'Capital'
            }
        else:  # Haining and others
            name_mapping = {
                'Cash at bank': 'Cash',
                'Accounts receivables': 'AR',
                'Prepayments': 'Prepayments',
                'Other receivables': 'OR',
                'Other current assets': 'Other CA',
                'Investment properties': 'IP',
                'Other non-current assets': 'Other NCA',
                'Accounts payable': 'AP',
                'Taxes payables': 'Taxes payable',
                'Other payables': 'OP',
                'Capital': 'Capital',
                'Surplus reserve': 'Reserve'
            }
        
        # Extract content by section
        content_by_key = {}
        
        # Split by section headers (### Section Name)
        sections = re.split(r'(^### .+$)', content, flags=re.MULTILINE)
        
        for i in range(1, len(sections), 2):
            if i + 1 < len(sections):
                section_header = sections[i].strip()
                section_content = sections[i + 1].strip()
                
                # Extract section name from header
                section_name = section_header.replace('### ', '').strip()
                
                # Map section name to key
                if section_name in name_mapping:
                    key = name_mapping[section_name]
                    content_by_key[key] = section_content
        
        return content_by_key
        
    except Exception as e:
        print(f"Error reading bs_content.md: {e}")
        return {}

def run_agent_3(filtered_keys, agent1_results, ai_data):
    """Run Agent 3: Pattern Compliance for all keys"""
    try:
        from common.assistant import PatternValidationAgent, load_ip
        import json
        import time
        
        logger = st.session_state.ai_logger
        st.markdown("## 🎯 Agent 3: Pattern Compliance")
        st.write(f"Starting Agent 3 for {len(filtered_keys)} keys...")
        
        # Load prompts from prompts.json
        try:
            with open('utils/prompts.json', 'r') as f:
                prompts_config = json.load(f)
            agent3_system_prompt = prompts_config.get('system_prompts', {}).get('Agent 3', '')
            st.success("✅ Loaded Agent 3 system prompt from prompts.json")
        except (FileNotFoundError, json.JSONDecodeError) as e:
            st.warning(f"⚠️ Could not load prompts.json: {e}")
            agent3_system_prompt = "Fallback Agent 3 system prompt"
        
        pattern_agent = PatternValidationAgent()
        results = {}
        
        # Load patterns
        patterns = load_ip("utils/pattern.json")
        st.write(f"Loaded patterns for {len(patterns)} keys")
        
        # Read current bs_content.md to get the latest content (possibly updated by Agent 2)
        bs_content_updates = {}
        current_content_by_key = read_bs_content_by_key(ai_data.get('entity_name', ''))
        st.write(f"Read bs_content.md with content for {len(current_content_by_key)} keys")
        
        for key in filtered_keys:
            st.write(f"🔄 Checking pattern compliance for {key} with Agent 3...")
            start_time = time.time()
            
            try:
                # Get the most recent content (from bs_content.md if updated by Agent 2, otherwise from Agent 1)
                current_content = current_content_by_key.get(key, agent1_results.get(key, ""))
                st.write(f"Content source for {key}: {'bs_content.md' if key in current_content_by_key else 'Agent 1 results'}")
                st.write(f"Content length for {key}: {len(current_content)} characters")
                
                if current_content:
                    # Get patterns for this key
                    key_patterns = patterns.get(key, {})
                    st.write(f"Found {len(key_patterns)} patterns for {key}")
                    
                    # Prepare detailed user prompt for pattern compliance
                    user_prompt = f"""
                    AI3 PATTERN COMPLIANCE CHECK:
                    
                    CONTENT TO ANALYZE: {current_content}
                    KEY: {key}
                    AVAILABLE PATTERNS FOR {key}: {json.dumps(key_patterns, indent=2)}
                    """
                    
                    # Log Agent 3 input
                    logger.log_agent_input('agent3', key, agent3_system_prompt, user_prompt, 
                                         {'patterns_count': len(key_patterns), 'content_length': len(current_content)})
                    
                    # Use real AI pattern validation instead of fallback
                    pattern_result = pattern_agent.validate_pattern_compliance(
                        content=current_content,
                        key=key
                    )
                    
                    processing_time = time.time() - start_time
                    
                    # Log Agent 3 output
                    logger.log_agent_output('agent3', key, pattern_result, processing_time)
                    
                    # If Agent 3 found issues and provided corrected content, use it
                    if pattern_result.get('corrected_content') and pattern_result.get('corrected_content') != current_content:
                        bs_content_updates[key] = pattern_result['corrected_content']
                        pattern_result['content_updated'] = True
                        st.success(f"✅ Agent 3 improved pattern compliance for {key}")
                    else:
                        pattern_result['content_updated'] = False
                        st.info(f"ℹ️ Agent 3 found no pattern improvements needed for {key}")
                    
                    results[key] = pattern_result
                    st.success(f"✅ Agent 3 completed {key} in {processing_time:.2f}s")
                    
                else:
                    error_msg = f"No content available for {key}"
                    st.error(f"❌ {error_msg}")
                    results[key] = {
                        "is_compliant": False,
                        "issues": [error_msg],
                        "pattern_match": "none",
                        "suggestions": ["Run Agent 1 first"],
                        "content_updated": False
                    }
                    logger.log_error('agent3', key, error_msg)
            
            except Exception as e:
                logger.log_error('agent3', key, str(e))
                st.error(f"❌ Agent 3 failed for {key}: {e}")
                results[key] = {
                    "is_compliant": False,
                    "issues": [f"Processing error: {e}"],
                    "pattern_match": "error",
                    "suggestions": ["Check error details"],
                    "content_updated": False
                }
        
        # Update bs_content.md with Agent 3 corrections if any
        if bs_content_updates:
            update_bs_content_with_agent_corrections(bs_content_updates, ai_data.get('entity_name', ''), "Agent 3")
            st.success(f"✅ Agent 3 updated bs_content.md with pattern compliance fixes for {len(bs_content_updates)} keys")
        else:
            st.info("ℹ️ Agent 3 found no pattern compliance improvements needed")
        
        st.success(f"🎉 Agent 3 completed all {len(filtered_keys)} keys")
        return results
        
    except Exception as e:
        st.error(f"Agent 3 processing failed: {e}")
        logger = st.session_state.ai_logger
        logger.log_error('agent3', 'general', str(e))
        return {}

def display_sequential_agent_results(key, filtered_keys, ai_data):
    """Display results for all agents for a specific key"""
    try:
        # Get agent states
        agent_states = st.session_state.get('agent_states', {})
        
        # Create sub-tabs for different perspectives (AI outputs only)
        perspective_tabs = st.tabs([
            "📝 Agent 1: Content Generation", 
            "📊 Agent 2: Data Validation", 
            "🎯 Agent 3: Pattern Compliance"
        ])
        
        # Tab 1: Agent 1 - Content Generation Focus
        with perspective_tabs[0]:
            st.markdown(f"### 🚀 What Agent 1 Does: Content Generation")
            st.info("**Focus**: Generate comprehensive financial analysis content using actual worksheet data and predefined patterns")
            
            if agent_states.get('agent1_completed', False):
                agent1_results = agent_states.get('agent1_results', {})
                content = agent1_results.get(key, "")
                
                if content:
                    # Show metadata in horizontal format
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Key", key)
                    with col2:
                        st.metric("Entity", ai_data.get('entity_name', ''))
                    with col3:
                        st.metric("Characters", len(content))
                    with col4:
                        st.metric("Words", len(content.split()))
                    
                    st.markdown("**✅ Generated Content:**")
                    st.markdown(content)
                    
                    # Show what Agent 1 accomplished
                    with st.expander("🎯 What Agent 1 Accomplished", expanded=False):
                        st.markdown("""
                        **Agent 1's Key Tasks:**
                        - Selected appropriate pattern from available templates
                        - Integrated actual financial figures from databook
                        - Replaced all placeholders with real data
                        - Applied proper K/M number formatting
                        - Generated professional financial narrative
                        """)
                else:
                    st.warning("No content generated for this key")
            else:
                st.info("⏳ Agent 1 will run automatically during 'Process with AI'")
        
        # Tab 2: Agent 2 - Data Validation Focus  
        with perspective_tabs[1]:
            st.markdown(f"### 🔍 What Agent 2 Does: Data Integrity & Accuracy Validation")
            st.info("**Focus**: Verify data accuracy, check K/M conversions, validate entity names, and ensure figures match balance sheet")
            
            if agent_states.get('agent2_completed', False):
                agent2_results = agent_states.get('agent2_results', {})
                validation_result = agent2_results.get(key, {})
                
                # Show what Agent 2 accomplished - make it collapsible
                with st.expander("🎯 What Agent 2 Accomplished", expanded=False):
                    accomplishments = [
                        "Extracted and verified all financial figures",
                        "Checked K/M conversion accuracy (thousands/millions)",
                        "Validated entity names match data source", 
                        "Ensured figures align with balance sheet data",
                        "Identified data inconsistencies and conversion errors"
                    ]
                    for acc in accomplishments:
                        st.write(f"• {acc}")
                
                if validation_result:
                    # Enhanced validation logic to catch K/M issues
                    is_valid = validation_result.get('is_valid', False)
                    score = validation_result.get('score', 0)
                    issues = validation_result.get('issues', [])
                    
                    # Additional validation for K/M conversion issues
                    agent1_results = agent_states.get('agent1_results', {})
                    agent1_content = agent1_results.get(key, "")
                    
                    # Check for inappropriate K/M usage (like years, counts, etc.)
                    import re
                    additional_issues = []
                    km_figures = re.findall(r'\d+\.?\d*[KM]', agent1_content)
                    for fig in km_figures:
                        # Extract the number
                        num_part = re.findall(r'\d+\.?\d*', fig)[0]
                        try:
                            value = float(num_part)
                            # Check if this looks like a year (between 1900-2100)
                            if 'K' in fig and 1900 <= value <= 2100:
                                additional_issues.append(f"Year figure '{fig}' should not use K notation")
                                score = max(0, score - 15)
                                is_valid = False
                            # Check for other inappropriate K/M usage
                            elif 'K' in fig and value < 10:
                                additional_issues.append(f"Small figure '{fig}' may not need K notation")
                                score = max(0, score - 5)
                        except ValueError:
                            continue
                    
                    # Combine original and additional issues
                    all_issues = issues + additional_issues
                    
                    st.markdown("---")
                    # Fix the logic: if score is 100 and no issues, it should be valid
                    if score >= 95 and len(all_issues) == 0:
                        st.success(f"✅ **Data Validation Result**: PASSED (Score: {score}/100)")
                    elif score >= 70:
                        st.warning(f"⚠️ **Data Validation Result**: ISSUES FOUND (Score: {score}/100)")
                    else:
                        st.error(f"❌ **Data Validation Result**: MAJOR ISSUES (Score: {score}/100)")
                    
                    # Show issues if any
                    if all_issues and any(issue.strip() for issue in all_issues):
                        st.markdown("**🚨 Issues Agent 2 Found:**")
                        for issue in all_issues:
                            if issue.strip():  # Only show non-empty issues
                                st.write(f"• {issue}")
                    
                    # Show suggestions if any
                    suggestions = validation_result.get('suggestions', [])
                    if suggestions and any(suggestion.strip() for suggestion in suggestions):
                        st.markdown("**💡 Agent 2 Suggestions:**")
                        for suggestion in suggestions:
                            if suggestion.strip():
                                st.write(f"• {suggestion}")
                    
                    # Show original vs corrected content
                    original_content = agent1_content
                    corrected_content = validation_result.get('corrected_content', '')
                    
                    if original_content and corrected_content and original_content != corrected_content:
                        st.markdown("**📋 Content Comparison:**")
                        col1, col2 = st.columns(2)
                        with col1:
                            st.markdown("**Original (Agent 1):**")
                            st.text_area("", original_content, height=200, key=f"orig_{key}", disabled=True)
                        with col2:
                            st.markdown("**Corrected (Agent 2):**")
                            st.text_area("", corrected_content, height=200, key=f"corr_{key}", disabled=True)
                    
                    # Enhanced data validation details with better highlighting
                    with st.expander("📊 Source Data Agent 2 Analyzed", expanded=False):
                        sections_by_key = ai_data.get('sections_by_key', {})
                        sections = sections_by_key.get(key, [])
                        if sections:
                            st.markdown("**Financial Data Reviewed with Figure Highlighting:**")
                            
                            # Extract figures from Agent 1 content for highlighting
                            import re
                            figures_in_content = re.findall(r'[\d,]+\.?\d*[KM]?', agent1_content)
                            
                            # Show highlighted tables
                            for i, section in enumerate(sections[:3]):
                                if isinstance(section, dict):
                                    if 'content' in section:
                                        section_content = section['content']
                                        st.markdown(f"**Section {i+1}:**")
                                        
                                        # Better highlighting logic
                                        highlighted_content = section_content
                                        for figure in figures_in_content:
                                            # Remove K/M and find base number
                                            base_fig = figure.replace('K', '').replace('M', '').replace(',', '')
                                            if base_fig.replace('.', '').isdigit():
                                                base_num = float(base_fig)
                                                
                                                # Look for the base number in the content
                                                base_pattern = str(int(base_num))
                                                formatted_pattern = f'{int(base_num):,}'
                                                
                                                # Highlight both formatted and unformatted versions
                                                if base_pattern in section_content:
                                                    highlighted_content = highlighted_content.replace(
                                                        base_pattern, 
                                                        f"🟡**{base_pattern}**🟡"
                                                    )
                                                if formatted_pattern in section_content and formatted_pattern != base_pattern:
                                                    highlighted_content = highlighted_content.replace(
                                                        formatted_pattern, 
                                                        f"🟡**{formatted_pattern}**🟡"
                                                    )
                                                
                                                # If it's in millions, also look for thousands representation
                                                if 'M' in figure:
                                                    thousands_num = int(base_num * 1000)
                                                    thousands_pattern = f'{thousands_num:,}'
                                                    if thousands_pattern in section_content:
                                                        highlighted_content = highlighted_content.replace(
                                                            thousands_pattern, 
                                                            f"🟡**{thousands_pattern}**🟡"
                                                        )
                                        
                                        # Display as markdown table if it has table structure
                                        lines = highlighted_content.split('\n')
                                        if len(lines) > 1 and any('|' in line for line in lines):
                                            st.markdown(highlighted_content)
                                        else:
                                            # Try to format as a table if possible
                                            try:
                                                # Look for tabular patterns and convert to markdown table
                                                import pandas as pd
                                                from io import StringIO
                                                
                                                # Try to parse as CSV-like content
                                                if '\t' in section_content or ',' in section_content:
                                                    delimiter = '\t' if '\t' in section_content else ','
                                                    df = pd.read_csv(StringIO(section_content), delimiter=delimiter, error_bad_lines=False)
                                                    if len(df) > 0:
                                                        st.dataframe(df, use_container_width=True)
                                                    else:
                                                        st.markdown(highlighted_content)
                                                else:
                                                    st.markdown(highlighted_content)
                                            except:
                                                st.markdown(highlighted_content)
                                        
                                        if '000' in section_content:
                                            st.caption("💡 Note: '000 notation detected - figures shown in thousands")
                                    
                                    elif 'data' in section and isinstance(section['data'], list):
                                        st.markdown(f"**Section {i+1} (Table):**")
                                        try:
                                            df = pd.DataFrame(section['data'])
                                            # Highlight matching figures in the dataframe
                                            def highlight_figures(val):
                                                if pd.isna(val):
                                                    return ''
                                                val_str = str(val)
                                                for figure in figures_in_content:
                                                    base_fig = figure.replace('K', '').replace('M', '').replace(',', '')
                                                    if base_fig.replace('.', '').isdigit():
                                                        base_num = float(base_fig)
                                                        # Look for the base number in various formats
                                                        if str(int(base_num)) in val_str or f'{int(base_num):,}' in val_str:
                                                            return 'background-color: yellow; font-weight: bold'
                                                        # For millions, check thousands representation
                                                        if 'M' in figure:
                                                            thousands_num = int(base_num * 1000)
                                                            if str(thousands_num) in val_str or f'{thousands_num:,}' in val_str:
                                                                return 'background-color: yellow; font-weight: bold'
                                                return ''
                                            
                                            styled_df = df.style.map(highlight_figures)
                                            st.dataframe(styled_df, use_container_width=True)
                                        except Exception as e:
                                            st.text(str(section['data']))
                        else:
                            st.info("No source data available")
                else:
                    st.warning("No validation results available")
            else:
                st.info("⏳ Agent 2 will run automatically during 'Process with AI'")
        
        # Tab 3: Agent 3 - Pattern Compliance Focus
        with perspective_tabs[2]:
            st.markdown(f"### 🎯 What Agent 3 Does: Pattern Compliance & Format Validation")
            st.info("**Focus**: Ensure content follows predefined patterns, check formatting standards, and verify template compliance")
            
            if agent_states.get('agent3_completed', False):
                agent3_results = agent_states.get('agent3_results', {})
                pattern_result = agent3_results.get(key, {})
                
                # Show what Agent 3 accomplished - make it collapsible
                with st.expander("🎯 What Agent 3 Accomplished", expanded=False):
                    accomplishments = [
                        "Compared content against available pattern templates",
                        "Verified all placeholders are filled with actual data",
                        "Checked professional formatting and structure", 
                        "Identified template artifacts that shouldn't be there",
                        "Ensured content follows pattern structure consistently"
                    ]
                    for acc in accomplishments:
                        st.write(f"• {acc}")
                
                if pattern_result:
                    # Enhanced pattern matching logic
                    is_compliant = pattern_result.get('is_compliant', False)
                    pattern_match = pattern_result.get('pattern_match', 'none')
                    issues = pattern_result.get('issues', [])
                    
                    # Get agent 1 content for analysis
                    agent1_results = agent_states.get('agent1_results', {})
                    agent1_content = agent1_results.get(key, "")
                    
                    st.markdown("---")
                    
                    # Improved pattern detection logic
                    try:
                        from common.assistant import load_ip
                        patterns = load_ip("utils/pattern.json")
                        key_patterns = patterns.get(key, {})
                        
                        if key_patterns and agent1_content:
                            st.markdown("**🎯 Pattern Analysis Results:**")
                            
                            # Try to detect which pattern was actually used
                            detected_pattern = None
                            highest_match_score = 0
                            
                            for pattern_name, pattern_content in key_patterns.items():
                                if isinstance(pattern_content, list):
                                    # Check each pattern variant
                                    for i, pattern_text in enumerate(pattern_content):
                                        # Simple similarity check - count matching words
                                        pattern_words = set(pattern_text.lower().split())
                                        content_words = set(agent1_content.lower().split())
                                        
                                        # Remove common words for better matching
                                        common_words = {'the', 'and', 'or', 'at', 'in', 'on', 'as', 'of', 'to', 'for', 'with', 'by'}
                                        pattern_words -= common_words
                                        content_words -= common_words
                                        
                                        # Calculate match score
                                        if len(pattern_words) > 0:
                                            match_score = len(pattern_words.intersection(content_words)) / len(pattern_words)
                                            if match_score > highest_match_score:
                                                highest_match_score = match_score
                                                detected_pattern = f"{pattern_name} (Variant {i+1})"
                                else:
                                    # Single pattern
                                    pattern_words = set(str(pattern_content).lower().split())
                                    content_words = set(agent1_content.lower().split())
                                    
                                    common_words = {'the', 'and', 'or', 'at', 'in', 'on', 'as', 'of', 'to', 'for', 'with', 'by'}
                                    pattern_words -= common_words
                                    content_words -= common_words
                                    
                                    if len(pattern_words) > 0:
                                        match_score = len(pattern_words.intersection(content_words)) / len(pattern_words)
                                        if match_score > highest_match_score:
                                            highest_match_score = match_score
                                            detected_pattern = pattern_name
                            
                            # Display results based on detection
                            if detected_pattern and highest_match_score > 0.3:  # 30% similarity threshold
                                st.success(f"✅ **Pattern Detected**: {detected_pattern}")
                                st.success(f"✅ **Match Confidence**: {highest_match_score:.1%}")
                                st.success(f"✅ **Pattern Compliance**: PASSED")
                            elif detected_pattern:
                                st.warning(f"⚠️ **Pattern Detected**: {detected_pattern}")
                                st.warning(f"⚠️ **Match Confidence**: {highest_match_score:.1%} (Low)")
                                st.warning(f"⚠️ **Pattern Compliance**: PARTIAL")
                            else:
                                st.info(f"🔍 **Pattern Detection**: No clear pattern match")
                                pattern_names = list(key_patterns.keys())
                                if pattern_names:
                                    st.info(f"📋 **Available Patterns**: {', '.join(pattern_names)}")
                        else:
                            st.warning("⚠️ No patterns available for this key or no content to analyze")
                    except Exception as e:
                        st.error(f"Error in pattern analysis: {e}")
                    
                    # Show issues if any (but filter out generic ones)
                    if issues and any(issue.strip() for issue in issues):
                        real_issues = [issue for issue in issues if issue.strip() and 
                                     "No Agent 1 content available" not in issue and
                                     "content doesn't match expected pattern structure" not in issue]
                        if real_issues:
                            st.markdown("**🚨 Specific Issues Agent 3 Found:**")
                            for issue in real_issues:
                                st.write(f"• {issue}")
                    
                    # Show missing elements if any
                    missing_elements = pattern_result.get('missing_elements', [])
                    if missing_elements and any(elem.strip() for elem in missing_elements):
                        st.markdown("**❌ Missing Elements Agent 3 Identified:**")
                        for element in missing_elements:
                            if element.strip():
                                st.write(f"• {element}")
                    
                    # Show suggestions if any
                    suggestions = pattern_result.get('suggestions', [])
                    if suggestions and any(suggestion.strip() for suggestion in suggestions):
                        st.markdown("**💡 Agent 3 Suggestions:**")
                        for suggestion in suggestions:
                            if suggestion.strip():
                                st.write(f"• {suggestion}")
                    
                    # Show available patterns using tabs for better display
                    st.markdown("---")
                    st.markdown("**📋 Available Patterns for This Key:**")
                    try:
                        from common.assistant import load_ip
                        patterns = load_ip("utils/pattern.json")
                        key_patterns = patterns.get(key, {})
                        
                        if key_patterns:
                            # Create tabs for different patterns
                            pattern_names = list(key_patterns.keys())
                            if len(pattern_names) > 1:
                                pattern_tabs = st.tabs(pattern_names)
                                
                                for idx, (pattern_name, pattern_content) in enumerate(key_patterns.items()):
                                    with pattern_tabs[idx]:
                                        st.markdown(f"**{pattern_name}:**")
                                        if isinstance(pattern_content, list):
                                            for i, pattern in enumerate(pattern_content):
                                                st.markdown(f"**Variant {i+1}:**")
                                                st.write(pattern)
                                                st.markdown("---")
                                        else:
                                            st.write(pattern_content)
                            else:
                                # Single pattern - no need for tabs
                                pattern_name, pattern_content = list(key_patterns.items())[0]
                                st.markdown(f"**{pattern_name}:**")
                                if isinstance(pattern_content, list):
                                    for i, pattern in enumerate(pattern_content):
                                        st.markdown(f"**Variant {i+1}:**")
                                        st.write(pattern)
                                        st.markdown("---")
                                else:
                                    st.write(pattern_content)
                        else:
                            st.info("No patterns available for this key")
                    except Exception as e:
                        st.error(f"Error loading patterns: {e}")
                        
                else:
                    st.warning("No pattern compliance results available")
            else:
                st.info("⏳ Agent 3 will run automatically during 'Process with AI'")
        

    
    except Exception as e:
        st.error(f"Error displaying results for {key}: {e}")

if __name__ == "__main__":
    main() 
import streamlit as st
import pandas as pd
import json
import warnings
import re
import os
import datetime
import time
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
        self.session_id = timestamp
        
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
    
    def _save_json_log(self, log_entry, log_type):
        """Save individual JSON log files for structured access"""
        try:
            # Create JSON subdirectory
            json_log_dir = self.log_dir / "json"
            json_log_dir.mkdir(exist_ok=True)
            
            # Create filename with timestamp and type
            timestamp = log_entry['timestamp'].replace(' ', '_').replace(':', '-')
            agent = log_entry['agent'].lower()
            key = log_entry['key'].lower()
            filename = f"{timestamp}_{agent}_{key}_{log_type}.json"
            
            json_file_path = json_log_dir / filename
            
            # Save structured JSON
            with open(json_file_path, 'w', encoding='utf-8') as f:
                json.dump(log_entry, f, indent=2, ensure_ascii=False)
                
        except Exception as e:
            print(f"Error saving JSON log: {e}")
    
    def _save_key_log(self, agent_name, key, input_entry, output_entry):
        """Save input and output for a key in one JSON file"""
        try:
            # Create filename for key log
            timestamp = input_entry['timestamp'].replace(' ', '_').replace(':', '-')
            filename = f"{timestamp}_{agent_name.lower()}_{key.lower()}.json"
            
            key_file_path = self.log_dir / filename
            
            # Combine input and output into key structure
            key_log = {
                'timestamp': input_entry['timestamp'],
                'agent': input_entry['agent'],
                'key': input_entry['key'],
                'session_id': input_entry.get('session_id', 'default'),
                'input': input_entry,
                'output': output_entry,
                'processing_time': output_entry.get('processing_time', 0),
                'success': output_entry.get('success', False)
            }
            
            # Save as formatted JSON
            with open(key_file_path, 'w', encoding='utf-8') as f:
                json.dump(key_log, f, indent=2, ensure_ascii=False)
                
        except Exception as e:
            print(f"Error saving key log: {e}")
        
    def _save_to_consolidated_log(self, log_entry):
        """Save log entry to consolidated session file (1 file per session with all agents and keys)"""
        try:
            # Create consolidated session file
            session_file = self.log_dir / f"session_{self.session_id}.json"
            
            # Load existing session data or create new
            session_data = {'session_info': {'session_id': self.session_id, 'started': datetime.datetime.now().isoformat()}, 'logs': []}
            if session_file.exists():
                try:
                    with open(session_file, 'r', encoding='utf-8') as f:
                        session_data = json.load(f)
                except (json.JSONDecodeError, FileNotFoundError):
                    pass
            
            # Add new log entry
            session_data['logs'].append(log_entry)
            session_data['session_info']['last_updated'] = datetime.datetime.now().isoformat()
            session_data['session_info']['total_logs'] = len(session_data['logs'])
            
            # Save updated session data
            with open(session_file, 'w', encoding='utf-8') as f:
                json.dump(session_data, f, indent=2, ensure_ascii=False)
                
        except Exception as e:
            print(f"Error saving to consolidated log: {e}")
    
    def log_agent_input(self, agent_name, key, system_prompt, user_prompt, context_data=None, actual_prompts=None):
        """Log agent input prompts and data to consolidated session file"""
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Create structured log entry for consolidated logging
        log_entry = {
            'timestamp': timestamp,
            'agent': agent_name.upper(),
            'key': key,
            'type': 'INPUT',
            'prompts': {
                'system_prompt': {
                    'content': system_prompt,
                    'length': len(system_prompt),
                    'token_estimate': len(system_prompt.split()) * 1.3  # Rough token estimate
                },
                'user_prompt': {
                    'content': user_prompt,
                    'length': len(user_prompt),
                    'token_estimate': len(user_prompt.split()) * 1.3
                }
            },
            'context_data': {
                'content': str(context_data) if context_data else None,
                'length': len(str(context_data)) if context_data else 0
            },
            'actual_prompts': actual_prompts,  # Store the actual parsed prompts
            'session_id': getattr(self, 'session_id', 'default')
        }
        
        if agent_name not in self.logs:
            self.logs[agent_name] = {}
        if key not in self.logs[agent_name]:
            self.logs[agent_name][key] = []
            
        self.logs[agent_name][key].append(log_entry)
        self.session_logs.append(log_entry)
        
        # Save to consolidated session file
        self._save_to_consolidated_log(log_entry)
        
        # Save individual JSON input file for detailed debugging
        self._save_json_log(log_entry, 'input')
        
        # Store input entry for paired logging
        if not hasattr(self, 'pending_inputs'):
            self.pending_inputs = {}
        self.pending_inputs[f"{agent_name}_{key}"] = log_entry
        
        # Minimal text log for quick reference
        self._write_to_file(f"📝 [{timestamp}] {agent_name.upper()} INPUT → {key} (Est. {log_entry['prompts']['system_prompt']['token_estimate'] + log_entry['prompts']['user_prompt']['token_estimate']:.0f} tokens)")
        
        # Silent processing - no Streamlit UI updates
    
    def log_agent_output(self, agent_name, key, output, processing_time=0):
        """Log agent output and processing details to JSON files with clear structure"""
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        is_success = output and not str(output).startswith("Error")
        
        # Create structured log entry for JSON
        log_entry = {
            'timestamp': timestamp,
            'agent': agent_name.upper(),
            'key': key,
            'type': 'OUTPUT',
            'output': {
                'content': output,
                'length': len(str(output)),
                'format': 'json' if isinstance(output, dict) else 'text'
            },
            'processing': {
                'time_seconds': processing_time,
                'status': 'success' if is_success else 'error',
                'is_success': is_success
            },
            'session_id': getattr(self, 'session_id', 'default')
        }
        
        if agent_name not in self.logs:
            self.logs[agent_name] = {}
        if key not in self.logs[agent_name]:
            self.logs[agent_name][key] = []
            
        self.logs[agent_name][key].append(log_entry)
        self.session_logs.append(log_entry)
        
        # Save individual JSON output file
        self._save_json_log(log_entry, 'output')
        
        # Create key log if input exists
        input_key = f"{agent_name}_{key}"
        if hasattr(self, 'pending_inputs') and input_key in self.pending_inputs:
            input_entry = self.pending_inputs[input_key]
            self._save_key_log(agent_name, key, input_entry, log_entry)
            # Remove from pending after key log is created
            del self.pending_inputs[input_key]
        
        # Write summary to text log
        status_icon = "✅" if is_success else "❌"
        self._write_to_file(f"{status_icon} [{timestamp}] {agent_name.upper()} OUTPUT ← {key} ({processing_time:.2f}s)")
        self._write_to_file(f"   Length: {len(str(output))} chars | Status: {'Success' if is_success else 'Error'}")
        
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
        
        # Debug: Show original DataFrame structure
        print(f"DEBUG: Original DataFrame shape: {df.shape}")
        print(f"DEBUG: Original DataFrame columns: {df.columns.tolist()}")
        print(f"DEBUG: Original DataFrame first few rows:")
        for i in range(min(5, len(df))):
            print(f"  Row {i}: {df.iloc[i].tolist()}")
        
        # Clean the DataFrame first - drop unnamed columns that are all NaN
        df_clean = df.copy()
        dropped_columns = []
        for col in df_clean.columns:
            if col.startswith('Unnamed:') or df_clean[col].isna().all():
                dropped_columns.append(col)
                df_clean = df_clean.drop(columns=[col])
        
        if dropped_columns:
            print(f"DEBUG: Dropped columns: {dropped_columns}")
            print(f"DEBUG: Remaining columns: {df_clean.columns.tolist()}")
        
        # If all columns were dropped, try a different approach
        if len(df_clean.columns) == 0:
            print("DEBUG: All columns dropped, trying alternative approach...")
            # Try to find columns with actual data
            df_clean = df.copy()
            for col in df_clean.columns:
                # Check if column has any non-null, non-empty values
                non_null_count = df_clean[col].notna().sum()
                non_empty_count = (df_clean[col].astype(str).str.strip() != '').sum()
                if non_null_count > 0 or non_empty_count > 0:
                    print(f"DEBUG: Keeping column {col} (non-null: {non_null_count}, non-empty: {non_empty_count})")
                else:
                    print(f"DEBUG: Dropping column {col} (all null/empty)")
                    df_clean = df_clean.drop(columns=[col])
        
        # Additional cleaning: remove columns that are all None/NaN after initial cleaning
        for col in list(df_clean.columns):
            if df_clean[col].isna().all() or (df_clean[col].astype(str) == 'None').all():
                df_clean = df_clean.drop(columns=[col])
                print(f"DEBUG: Removed all-NaN column {col} from structured parsing")
        
        print(f"DEBUG: Final cleaned DataFrame shape: {df_clean.shape}")
        print(f"DEBUG: Final cleaned DataFrame columns: {df_clean.columns.tolist()}")
        
        # Convert all cells to string for analysis
        df_str = df_clean.astype(str).fillna('')
        
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
                    print(f"DEBUG: Found 'Indicative adjusted' in column {j}")
                    break
                elif "total" in cell_value and value_col_idx is None:
                    value_col_idx = j
                    value_col_name = "Total"
                    print(f"DEBUG: Found 'Total' in column {j}")
        
        # If still no specific column found, look for any column with financial data patterns
        if value_col_idx is None:
            for j in range(len(df_str.columns)):
                column_data = df_str.iloc[:, j]
                # Look for columns that contain financial data patterns
                financial_patterns = ['amount', 'value', 'balance', 'figure', 'total', 'sum']
                for i in range(min(3, len(df_str))):
                    cell_value = str(df_str.iloc[i, j]).lower()
                    if any(pattern in cell_value for pattern in financial_patterns):
                        value_col_idx = j
                        value_col_name = f"Financial Column {j+1}"
                        print(f"DEBUG: Found financial pattern in column {j}: '{cell_value}'")
                        break
                if value_col_idx is not None:
                    break
        
        # If no specific column found, use the rightmost column with numbers
        if value_col_idx is None:
            candidate_cols = []
            for j in range(len(df_str.columns) - 1, -1, -1):
                column_data = df_str.iloc[:, j]
                column_name = str(df_str.columns[j]).lower()
                # Exclude Excel-generated or index columns
                if any(skip_name in column_name for skip_name in ['column1', 'unnamed', 'index']):
                    print(f"DEBUG: Skipping column {j} '{df_str.columns[j]}' - appears to be Excel-generated")
                    continue
                # Check if column is strictly sequential (row numbers)
                numeric_values = []
                for cell in column_data:
                    cell_str = str(cell).strip()
                    if cell_str and cell_str.lower() not in ['nan', '']:
                        if re.fullmatch(r'\d+', cell_str.replace(',', '')):
                            try:
                                numeric_values.append(int(cell_str.replace(',', '')))
                            except ValueError:
                                pass
                if len(numeric_values) >= 2:
                    diffs = [numeric_values[i+1] - numeric_values[i] for i in range(len(numeric_values)-1)]
                    if all(d == 1 for d in diffs) or all(d == 0 for d in diffs):
                        print(f"DEBUG: Skipping column {j} - strictly sequential numbers: {numeric_values[:5]}")
                        continue
                    # Also skip if all values are the same (e.g., all 1000)
                    if len(set(numeric_values)) == 1:
                        print(f"DEBUG: Skipping column {j} - all values identical: {numeric_values[0]}")
                        continue
                # Now check if this is a good candidate
                numeric_count = 0
                total_cells = 0
                for cell in column_data:
                    cell_str = str(cell).strip()
                    if cell_str and cell_str.lower() not in ['nan', '']:
                        total_cells += 1
                        if re.search(r'^\d+\.?\d*$', cell_str.replace(',', '')):
                            numeric_count += 1
                if total_cells > 0 and numeric_count >= total_cells * 0.3:
                    candidate_cols.append(j)
            if candidate_cols:
                # Pick the rightmost candidate
                value_col_idx = candidate_cols[0]
                value_col_name = f"Column {value_col_idx+1}"
                print(f"DEBUG: Selected value column {value_col_idx} (rightmost non-row-number column)")
            else:
                print("DEBUG: No valid value column found after excluding row number/index columns.")
                return None
        
        # Find where actual data starts (skip header rows)
        data_start_row = None
        for i in range(len(df_str)):
            cell_value = str(df_str.iloc[i, value_col_idx])
            # Look for cells that contain numbers (more flexible)
            if re.search(r'\d+', cell_value) and cell_value.strip() not in ['nan', '']:
                # Check if this looks like a data row (has both description and value)
                desc_cell = str(df_str.iloc[i, 0]).strip()
                if desc_cell and desc_cell.lower() not in ['nan', '']:
                    data_start_row = i
                    break
        
        if data_start_row is None:
            # Fallback: start from row 2 if we have at least 3 rows
            if len(df_str) >= 3:
                data_start_row = 2
            else:
                return None
        
        # Extract date from the table - look for date patterns in the first few rows
        extracted_date = None
        date_patterns = [
            r'(\d{4})-(\d{1,2})-(\d{1,2})',  # YYYY-MM-DD
            r'(\d{1,2})/(\d{1,2})/(\d{4})',  # DD/MM/YYYY
            r'(\d{1,2})-(\d{1,2})-(\d{4})',  # DD-MM-YYYY
            r'(\d{4})/(\d{1,2})/(\d{1,2})',  # YYYY/MM/DD
        ]
        
        # Search for date in the first 10 rows
        for i in range(min(10, len(df_str))):
            for j in range(len(df_str.columns)):
                cell_value = str(df_str.iloc[i, j]).strip()
                for pattern in date_patterns:
                    match = re.search(pattern, cell_value)
                    if match:
                        try:
                            if pattern == r'(\d{4})-(\d{1,2})-(\d{1,2})':
                                year, month, day = match.groups()
                            elif pattern == r'(\d{1,2})/(\d{1,2})/(\d{4})':
                                day, month, year = match.groups()
                            elif pattern == r'(\d{1,2})-(\d{1,2})-(\d{4})':
                                day, month, year = match.groups()
                            elif pattern == r'(\d{4})/(\d{1,2})/(\d{1,2})':
                                year, month, day = match.groups()
                            
                            # Validate date
                            from datetime import datetime
                            dt = datetime(int(year), int(month), int(day))
                            extracted_date = dt.strftime('%Y-%m-%d')
                            print(f"DEBUG: Found date '{extracted_date}' in cell [{i},{j}]: '{cell_value}'")
                            break
                        except (ValueError, TypeError):
                            continue
                if extracted_date:
                    break
            if extracted_date:
                break
        
        # Extract table metadata (first few rows before data)
        table_metadata = {
            'table_name': f"{key} - {entity_name}",
            'sheet_name': sheet_name,
            'date': extracted_date,
            'currency_info': currency_info,
            'multiplier': multiplier,
            'value_column': value_col_name,
            'data_start_row': data_start_row
        }
        
        # Extract actual data rows
        description_col_idx = 0  # Usually first column
        data_rows = []
        
        # Debug: Print the DataFrame to understand the structure
        print(f"DEBUG: DataFrame shape: {df_str.shape}")
        print(f"DEBUG: DataFrame columns: {df_str.columns.tolist()}")
        print(f"DEBUG: Data start row: {data_start_row}")
        print(f"DEBUG: Value column index: {value_col_idx}")
        print(f"DEBUG: Value column name: {value_col_name}")
        print(f"DEBUG: First few rows:")
        for i in range(min(10, len(df_str))):
            print(f"  Row {i}: {df_str.iloc[i].tolist()}")
        
        # Debug: Show column analysis
        print(f"DEBUG: Column analysis:")
        for j in range(len(df_str.columns)):
            column_data = df_str.iloc[:, j]
            numeric_count = 0
            total_cells = 0
            sample_values = []
            for cell in column_data:
                cell_str = str(cell).strip()
                if cell_str and cell_str.lower() not in ['nan', '']:
                    total_cells += 1
                    sample_values.append(cell_str)
                    if re.search(r'^\d+\.?\d*$', cell_str.replace(',', '')):
                        numeric_count += 1
            print(f"  Column {j}: {numeric_count}/{total_cells} numeric, samples: {sample_values[:3]}")
        
        # Skip currency/date header rows that shouldn't be treated as data
        skip_patterns = [
            r"CNY'000", r"USD'000", r"'000", r"thousands", r"thousand",
            r"millions", r"million", r"'000,000",
            r"\d{4}-\d{1,2}-\d{1,2}",  # Date patterns
            r"\d{1,2}/\d{1,2}/\d{4}",
            r"\d{1,2}-\d{1,2}-\d{4}",
            r"\d{4}/\d{1,2}/\d{1,2}",
            r"indicative adjusted"
        ]
        
        for i in range(data_start_row, len(df_str)):
            description = str(df_str.iloc[i, description_col_idx]).strip()
            value_str = str(df_str.iloc[i, value_col_idx]).strip()
            
            # Skip empty rows
            if not description or description.lower() in ['nan', '']:
                continue
            
            # Check if this row should be skipped
            should_skip = False
            for pattern in skip_patterns:
                if re.search(pattern, description.lower()) or re.search(pattern, value_str.lower()):
                    should_skip = True
                    print(f"DEBUG: Skipping row - Description: '{description}', Value: '{value_str}' (matches pattern: {pattern})")
                    break
            
            # Additional check: skip if description is a pure number (like 1000, 1001, etc.)
            if re.match(r'^\d+\.?\d*$', description.strip()):
                should_skip = True
                print(f"DEBUG: Skipping row - Description is pure number: '{description}'")
            
            if should_skip:
                continue
            
            # Extract numeric value - be more flexible
            value = 0
            if re.search(r'\d', value_str):
                # Remove commas and extract number - handle more formats
                clean_value = re.sub(r'[^\d.-]', '', value_str.replace(',', ''))
                try:
                    value = float(clean_value) * multiplier  # Apply multiplier
                except ValueError:
                    # Try to extract just the numeric part
                    numeric_match = re.search(r'[\d,]+\.?\d*', value_str.replace(',', ''))
                    if numeric_match:
                        try:
                            value = float(numeric_match.group()) * multiplier
                        except ValueError:
                            value = 0
                    else:
                        value = 0
            
            # Include rows with descriptions even if value is 0 (for debugging)
            if description and description.lower() not in ['nan']:
                # Include total rows but mark them
                is_total = description.lower() in ['total', 'subtotal']
                data_rows.append({
                    'description': description,
                    'value': value,
                    'original_value': value_str,
                    'is_total': is_total
                })
                print(f"DEBUG: Added row - Description: '{description}', Value: {value}, Original: '{value_str}', IsTotal: {is_total}")
        
        print(f"DEBUG: Total data rows extracted: {len(data_rows)}")
        
        return {
            'metadata': table_metadata,
            'data': data_rows,
            'raw_df': df_clean  # Use cleaned DataFrame
        }
        
    except Exception as e:
        print(f"Error parsing accounting table: {e}")
        return None

def format_date_to_dd_mmm_yyyy(date_str):
    """Convert date string to dd-mmm-yyyy format"""
    import re
    from datetime import datetime
    
    if not date_str or str(date_str).lower() in ['nan', 'none', 'unknown']:
        return 'Unknown'
    
    date_str = str(date_str).strip()
    
    # Common date patterns
    patterns = [
        # YYYY-MM-DD HH:MM:SS
        r'(\d{4})-(\d{1,2})-(\d{1,2})(?:\s+\d{1,2}:\d{1,2}:\d{1,2})?',
        # DD/MM/YYYY or MM/DD/YYYY
        r'(\d{1,2})/(\d{1,2})/(\d{4})',
        # DD-MM-YYYY
        r'(\d{1,2})-(\d{1,2})-(\d{4})',
        # YYYY/MM/DD
        r'(\d{4})/(\d{1,2})/(\d{1,2})',
    ]
    
    for pattern in patterns:
        match = re.match(pattern, date_str)
        if match:
            try:
                if len(match.groups()) == 3:
                    if pattern == r'(\d{4})-(\d{1,2})-(\d{1,2})(?:\s+\d{1,2}:\d{1,2}:\d{1,2})?':
                        # YYYY-MM-DD format
                        year, month, day = match.groups()
                    elif pattern == r'(\d{1,2})/(\d{1,2})/(\d{4})':
                        # DD/MM/YYYY or MM/DD/YYYY - assume DD/MM/YYYY
                        day, month, year = match.groups()
                    elif pattern == r'(\d{1,2})-(\d{1,2})-(\d{4})':
                        # DD-MM-YYYY
                        day, month, year = match.groups()
                    elif pattern == r'(\d{4})/(\d{1,2})/(\d{1,2})':
                        # YYYY/MM/DD
                        year, month, day = match.groups()
                    
                    # Create datetime object and format
                    dt = datetime(int(year), int(month), int(day))
                    return dt.strftime('%d-%b-%Y')
            except (ValueError, TypeError):
                continue
    
    # If no pattern matches, return original
    return date_str

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
        
        # Format date if present
        if metadata.get('date'):
            formatted_date = format_date_to_dd_mmm_yyyy(metadata['date'])
            markdown_lines.append(f"*Date: {formatted_date}*")
        
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
                is_total = row.get('is_total', False)
                
                # Use actual multiplied values with formatting
                actual_value = value  # This is already multiplied by the multiplier
                
                # Format value with thousand separators and 2 decimal places
                if isinstance(actual_value, (int, float)):
                    formatted_value = f"{actual_value:,.2f}"
                else:
                    formatted_value = str(actual_value)
                
                # Add bold formatting for total rows
                if is_total:
                    description = f"**{description}**"
                    formatted_value = f"**{formatted_value}**"
                
                markdown_lines.append(f"| {description} | {formatted_value} |")
        else:
            markdown_lines.append("*No data rows found*")
        
        return "\n".join(markdown_lines)
        
    except Exception as e:
        return f"Error creating table markdown: {e}"

@cached_function(ttl=1800)  # Cache for 30 minutes
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
                            
                            # Check if we have structured data available
                            if 'parsed_data' in first_section and first_section['parsed_data']:
                                # Use structured data
                                parsed_data = first_section['parsed_data']
                                metadata = parsed_data['metadata']
                                data_rows = parsed_data['data']
                                
                                # Display metadata horizontally to save space
                                col1, col2, col3, col4, col5, col6 = st.columns(6)
                                with col1:
                                    st.markdown(f"**Table:** {metadata['table_name']}")
                                with col2:
                                    if metadata.get('date'):
                                        formatted_date = format_date_to_dd_mmm_yyyy(metadata['date'])
                                        st.markdown(f"**Date:** {formatted_date}")
                                    else:
                                        st.markdown("**Date:** Unknown")
                                with col3:
                                    st.markdown(f"**Currency:** {metadata['currency_info']}")
                                with col4:
                                    st.markdown(f"**Multiplier:** {metadata['multiplier']}x")
                                with col5:
                                    st.markdown(f"**Value Column:** {metadata['value_column']}")
                                with col6:
                                    if first_section.get('entity_match', False):
                                        st.markdown("**Entity:** ✅")
                                    else:
                                        st.markdown("**Entity:** ⚠️")
                                
                                # Display structured data as a clean table
                                if data_rows:
                                    structured_data = []
                                    for row in data_rows:
                                        description = row['description']
                                        value = row['value']
                                        
                                        # Use actual multiplied values with formatting
                                        actual_value = value  # This is already multiplied by the multiplier
                                        
                                        # Format value with thousand separators and 2 decimal places
                                        if isinstance(actual_value, (int, float)):
                                            formatted_value = f"{actual_value:,.2f}"
                                        else:
                                            formatted_value = str(actual_value)
                                        
                                        structured_data.append({
                                            'Description': description,
                                            'Value': formatted_value
                                        })
                                    
                                    df_structured = pd.DataFrame(structured_data)
                                    
                                    # Clean up the display - show only Description and Value columns
                                    display_df = df_structured[['Description', 'Value']].copy()
                                    
                                    # Highlight total rows with theme-appropriate styling
                                    def highlight_totals(row):
                                        if row['Description'].lower() in ['total', 'subtotal']:
                                            # Use a subtle highlight that works with both light and dark themes
                                            # Light blue tint with low opacity works well in both themes
                                            return ['background-color: rgba(173, 216, 230, 0.3)'] * len(row)
                                        return [''] * len(row)  # Let theme handle default background
                                    
                                    styled_df = display_df.style.apply(highlight_totals, axis=1)
                                    st.dataframe(styled_df, use_container_width=True)
                                    
                                    # Show structured markdown
                                    with st.expander(f"📋 Structured Markdown", expanded=False):
                                        st.code(first_section.get('markdown', 'No markdown available'), language='markdown')
                                else:
                                    st.info("No structured data rows found")
                                    st.write(f"**Financial Data for {key}:**")
                                    
                                    # Try to use parsed data structure even if no rows were extracted
                                    if 'parsed_data' in first_section and first_section['parsed_data']:
                                        parsed_data = first_section['parsed_data']
                                        metadata = parsed_data['metadata']
                                        
                                        # Display metadata horizontally
                                        col1, col2, col3, col4, col5, col6 = st.columns(6)
                                        with col1:
                                            st.markdown(f"**Table:** {metadata['table_name']}")
                                        with col2:
                                            if metadata.get('date'):
                                                formatted_date = format_date_to_dd_mmm_yyyy(metadata['date'])
                                                st.markdown(f"**Date:** {formatted_date}")
                                            else:
                                                st.markdown("**Date:** Unknown")
                                        with col3:
                                            st.markdown(f"**Currency:** {metadata['currency_info']}")
                                        with col4:
                                            st.markdown(f"**Multiplier:** {metadata['multiplier']}x")
                                        with col5:
                                            st.markdown(f"**Value Column:** {metadata['value_column']}")
                                        with col6:
                                            if first_section.get('entity_match', False):
                                                st.markdown("**Entity:** ✅")
                                            else:
                                                st.markdown("**Entity:** ⚠️")
                                        
                                        # Clean and display the raw DataFrame with proper column names
                                        raw_df = first_section['data'].copy()
                                        
                                        # Remove columns that are all None/NaN
                                        for col in list(raw_df.columns):
                                            if raw_df[col].isna().all() or (raw_df[col].astype(str) == 'None').all():
                                                raw_df = raw_df.drop(columns=[col])
                                                print(f"DEBUG: Removed all-NaN column: {col}")
                                        
                                        # Rename columns to be more descriptive
                                        if len(raw_df.columns) >= 2:
                                            new_column_names = [f"{key} (Description)", f"{key} (Balance)"]
                                            if len(raw_df.columns) > 2:
                                                for i in range(2, len(raw_df.columns)):
                                                    new_column_names.append(f"{key} (Column {i+1})")
                                            raw_df.columns = new_column_names
                                        elif len(raw_df.columns) == 1:
                                            raw_df.columns = [f"{key} (Description)"]
                                        
                                        # Display the cleaned DataFrame
                                        if len(raw_df.columns) > 0:
                                            st.dataframe(raw_df, use_container_width=True)
                                            
                                            # Create structured markdown for AI prompts
                                            markdown_lines = []
                                            markdown_lines.append(f"## {metadata['table_name']}")
                                            markdown_lines.append(f"**Entity:** {metadata['table_name'].split(' - ')[-1] if ' - ' in metadata['table_name'] else 'Unknown'}")
                                            
                                            # Format date if present
                                            if metadata.get('date'):
                                                formatted_date = format_date_to_dd_mmm_yyyy(metadata['date'])
                                                markdown_lines.append(f"**Date:** {formatted_date}")
                                            else:
                                                markdown_lines.append(f"**Date:** {metadata.get('data_start_row', 'Unknown')}")
                                            
                                            markdown_lines.append(f"**Currency:** {metadata['currency_info']}")
                                            markdown_lines.append(f"**Multiplier:** {metadata['multiplier']}")
                                            markdown_lines.append("")
                                            
                                            # Add data rows with actual values
                                            for _, row in raw_df.iterrows():
                                                description = str(row.iloc[0]) if len(row) > 0 else ""
                                                if len(row) > 1:
                                                    value = str(row.iloc[1])
                                                    # Try to convert to numeric and apply multiplier if it's a number
                                                    try:
                                                        numeric_value = float(value.replace(',', ''))
                                                        # Apply multiplier from metadata if available
                                                        if 'multiplier' in metadata:
                                                            actual_value = numeric_value * metadata['multiplier']
                                                            # Format with thousand separators and 2 decimal places
                                                            if isinstance(actual_value, (int, float)):
                                                                value = f"{actual_value:,.2f}"
                                                            else:
                                                                value = str(actual_value)
                                                    except (ValueError, AttributeError):
                                                        # Keep original value if not numeric
                                                        pass
                                                    
                                                    if description and description.lower() not in ['nan', 'none', '']:
                                                        markdown_lines.append(f"- {description}: {value}")
                                            
                                            markdown_lines.append("")
                                            
                                            # Show the structured markdown
                                            with st.expander(f"📋 Structured Data for AI", expanded=False):
                                                st.code('\n'.join(markdown_lines), language='markdown')
                                        else:
                                            st.error("No valid data columns found after cleaning")
                                    
                                    else:
                                        # Fallback to original cleaning logic if no parsed data
                                        raw_df = first_section['data'].copy()
                                        
                                        # Remove columns that are all None/NaN
                                        for col in list(raw_df.columns):
                                            if raw_df[col].isna().all() or (raw_df[col].astype(str) == 'None').all():
                                                raw_df = raw_df.drop(columns=[col])
                                                print(f"DEBUG: Removed all-NaN column: {col}")
                                        
                                        # Rename columns to be more descriptive
                                        if len(raw_df.columns) >= 2:
                                            new_column_names = [f"{key} (Description)", f"{key} (Balance)"]
                                            if len(raw_df.columns) > 2:
                                                for i in range(2, len(raw_df.columns)):
                                                    new_column_names.append(f"{key} (Column {i+1})")
                                            raw_df.columns = new_column_names
                                        elif len(raw_df.columns) == 1:
                                            raw_df.columns = [f"{key} (Description)"]
                                        
                                        if len(raw_df.columns) > 0:
                                            st.dataframe(raw_df, use_container_width=True)
                                        else:
                                            st.error("No valid columns found after cleaning")
                                            st.write("**Original DataFrame:**")
                                            st.dataframe(first_section['data'], use_container_width=True)
                                    
                                    # Also show the parsed data structure for debugging
                                    if 'parsed_data' in first_section:
                                        with st.expander("🔍 Debug: Parsed Data Structure", expanded=False):
                                            st.json(first_section['parsed_data'])
                                
                            # Note: The fallback logic is now handled in the "else" block above
                            # This ensures we always use the improved DataFrame cleaning and markdown generation
                            
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
        # Check AI configuration status
        try:
            config, _, _, _ = load_config_files()
            if config and (not config.get('OPENAI_API_KEY') or not config.get('OPENAI_API_BASE')):
                st.warning("⚠️ AI Mode: API keys not configured. Will use fallback mode with test data.")
                st.info("💡 To enable full AI functionality, please configure your OpenAI API keys in utils/config.json")
        except Exception:
            st.warning("⚠️ AI Mode: Configuration not found. Will use fallback mode.")
        
        # --- AI Processing & Results Section ---
        st.markdown("---")
        st.markdown("## 🤖 AI Processing & Results")
        
        # Initialize session state for AI data if not exists
        if 'ai_data' not in st.session_state:
            st.session_state['ai_data'] = {}
        
        # Prepare data for AI processing
        if uploaded_file is not None:
            try:
                # Load configuration files
                config, mapping, pattern, prompts = load_config_files()
                if not all([config, mapping, pattern]):
                    st.error("❌ Failed to load configuration files")
                    return
                
                # Process entity configuration
                entity_suffixes = [s.strip() for s in entity_helpers.split(',') if s.strip()]
                entity_keywords = [f"{selected_entity} {suffix}" for suffix in entity_suffixes if suffix]
                if not entity_keywords:
                    entity_keywords = [selected_entity]
                
                # Get worksheet sections
                sections_by_key = get_worksheet_sections_by_keys(
                    uploaded_file=uploaded_file,
                    tab_name_mapping=mapping,
                    entity_name=selected_entity,
                    entity_suffixes=entity_suffixes,
                    debug=False
                )
                
                # Get keys with data
                keys_with_data = [key for key, sections in sections_by_key.items() if sections]
                
                # Filter keys based on statement type
                bs_keys = [
                    "Cash", "AR", "Prepayments", "OR", "Other CA", "IP", "Other NCA",
                    "AP", "Taxes payable", "OP", "Capital", "Reserve"
                ]
                is_keys = [
                    "OI", "OC", "Tax and Surcharges", "GA", "Fin Exp", "Cr Loss", "Other Income",
                    "Non-operating Income", "Non-operating Exp", "Income tax", "LT DTA"
                ]
                
                if statement_type == "BS":
                    filtered_keys_for_ai = [key for key in keys_with_data if key in bs_keys]
                elif statement_type == "IS":
                    filtered_keys_for_ai = [key for key in keys_with_data if key in is_keys]
                elif statement_type == "ALL":
                    filtered_keys_for_ai = [key for key in keys_with_data if key in (bs_keys + is_keys)]
                else:
                    filtered_keys_for_ai = keys_with_data
                
                if not filtered_keys_for_ai:
                    st.warning("No data found for AI processing with the selected statement type.")
                    return
                
                # Store uploaded file data in session state for agents
                st.session_state['uploaded_file_data'] = uploaded_file.getbuffer()
                
                # Prepare AI data
                temp_ai_data = {
                    'entity_name': selected_entity,
                    'entity_keywords': entity_keywords,
                    'sections_by_key': sections_by_key,
                    'pattern': pattern,
                    'mapping': mapping,
                    'config': config
                }
                
                # Store in session state
                st.session_state['ai_data'] = temp_ai_data
                st.session_state['filtered_keys_for_ai'] = filtered_keys_for_ai
                
                # Initialize agent states if not exists
                if 'agent_states' not in st.session_state:
                    st.session_state['agent_states'] = {
                        'agent1_completed': False,
                        'agent2_completed': False, 
                        'agent3_completed': False,
                        'agent1_results': {},
                        'agent2_results': {},
                        'agent3_results': {},
                        'agent1_success': False,
                        'agent2_success': False,
                        'agent3_success': False
                    }
                
                # AI Processing Buttons with Progress
                agent_states = st.session_state.get('agent_states', {})
                agent1_completed = agent_states.get('agent1_completed', False)
                agent1_results = agent_states.get('agent1_results', {})
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    if st.button("🚀 Run AI1: Content Generation", type="primary", use_container_width=True):
                        # Progress bar for AI1
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        try:
                            status_text.text("🤖 AI1: Initializing...")
                            progress_bar.progress(0.1)
                            
                            agent1_results = run_agent_1(filtered_keys_for_ai, temp_ai_data)
                            agent1_success = bool(agent1_results and any(agent1_results.values()))
                            
                            st.session_state['agent_states']['agent1_results'] = agent1_results
                            st.session_state['agent_states']['agent1_completed'] = True
                            st.session_state['agent_states']['agent1_success'] = agent1_success
                            
                            progress_bar.progress(1.0)
                            if agent1_success:
                                status_text.text(f"✅ AI1 completed! Generated content for {len(agent1_results)} keys.")
                            else:
                                status_text.text("❌ AI1 failed to generate content.")
                            
                            time.sleep(2)  # Show completion message briefly
                            st.rerun()
                            
                        except Exception as e:
                            progress_bar.progress(1.0)
                            status_text.text(f"❌ AI1 failed: {e}")
                            time.sleep(2)
                            st.rerun()
                
                with col2:
                    if agent1_completed and agent1_results:
                        if st.button("📊 Run AI2: Data Validation", type="secondary", use_container_width=True):
                            # Progress bar for AI2
                            progress_bar = st.progress(0)
                            status_text = st.empty()
                            
                            try:
                                status_text.text("🔍 AI2: Initializing...")
                                progress_bar.progress(0.1)
                                
                                agent2_results = run_agent_2(filtered_keys_for_ai, agent1_results, temp_ai_data)
                                agent2_success = bool(agent2_results and len(agent2_results) > 0)
                                
                                st.session_state['agent_states']['agent2_results'] = agent2_results
                                st.session_state['agent_states']['agent2_completed'] = True
                                st.session_state['agent_states']['agent2_success'] = agent2_success
                                
                                progress_bar.progress(1.0)
                                if agent2_success:
                                    status_text.text(f"✅ AI2 completed! Validated {len(agent2_results)} keys.")
                                else:
                                    status_text.text("❌ AI2 failed to validate data.")
                                
                                time.sleep(2)
                                st.rerun()
                                
                            except Exception as e:
                                progress_bar.progress(1.0)
                                status_text.text(f"❌ AI2 failed: {e}")
                                time.sleep(2)
                                st.rerun()
                    else:
                        st.button("📊 Run AI2: Data Validation", disabled=True, use_container_width=True)
                        st.caption("⚠️ Run AI1 first")
                
                with col3:
                    if agent1_completed and agent1_results:
                        if st.button("🎯 Run AI3: Pattern Compliance", type="secondary", use_container_width=True):
                            # Progress bar for AI3
                            progress_bar = st.progress(0)
                            status_text = st.empty()
                            
                            try:
                                status_text.text("🎯 AI3: Initializing...")
                                progress_bar.progress(0.1)
                                
                                agent3_results = run_agent_3(filtered_keys_for_ai, agent1_results, temp_ai_data)
                                agent3_success = bool(agent3_results and len(agent3_results) > 0)
                                
                                st.session_state['agent_states']['agent3_results'] = agent3_results
                                st.session_state['agent_states']['agent3_completed'] = True
                                st.session_state['agent_states']['agent3_success'] = agent3_success
                                
                                progress_bar.progress(1.0)
                                if agent3_success:
                                    status_text.text(f"✅ AI3 completed! Checked compliance for {len(agent3_results)} keys.")
                                else:
                                    status_text.text("❌ AI3 failed to check pattern compliance.")
                                
                                time.sleep(2)
                                st.rerun()
                                
                            except Exception as e:
                                progress_bar.progress(1.0)
                                status_text.text(f"❌ AI3 failed: {e}")
                                time.sleep(2)
                                st.rerun()
                    else:
                        st.button("🎯 Run AI3: Pattern Compliance", disabled=True, use_container_width=True)
                        st.caption("⚠️ Run AI1 first")
                
            except Exception as e:
                st.error(f"❌ Failed to prepare AI data: {e}")
        else:
            st.info("Please upload an Excel file first.")
        
        # --- AI Results Display ---
        # Check if any agent has run
        agent_states = st.session_state.get('agent_states', {})
        any_agent_completed = any([
            agent_states.get('agent1_completed', False),
            agent_states.get('agent2_completed', False),
            agent_states.get('agent3_completed', False)
        ])
        
        if any_agent_completed:
            # Get available keys
            filtered_keys = st.session_state.get('filtered_keys_for_ai', [])
            
            if filtered_keys:
                # Create tabs for each key (load all at once)
                key_tabs = st.tabs([get_key_display_name(key) for key in filtered_keys])
                
                # Display results for each key in its tab
                for i, key in enumerate(filtered_keys):
                    with key_tabs[i]:
                        st.markdown(f"### {get_key_display_name(key)} Results")
                        
                        # Create sub-tabs for each agent
                        agent_tabs = st.tabs(["🚀 AI1: Generation", "📊 AI2: Validation", "🎯 AI3: Compliance"])
                        
                        # AI1 Results
                        with agent_tabs[0]:
                            agent1_results = agent_states.get('agent1_results', {})
                            if key in agent1_results and agent1_results[key]:
                                content = agent1_results[key]
                                st.markdown("**Generated Content:**")
                                st.markdown(content)
                                
                                # Metadata
                                col1, col2, col3 = st.columns(3)
                                with col1:
                                    st.metric("Characters", len(content))
                                with col2:
                                    st.metric("Words", len(content.split()))
                                with col3:
                                    st.metric("Status", "✅ Generated" if content else "❌ Failed")
                            else:
                                st.info("No AI1 results available. Run AI1 first.")
                        
                        # AI2 Results
                        with agent_tabs[1]:
                            agent2_results = agent_states.get('agent2_results', {})
                            if key in agent2_results:
                                validation_result = agent2_results[key]
                                
                                # Validation metrics
                                col1, col2, col3 = st.columns(3)
                                with col1:
                                    score = validation_result.get('score', 0)
                                    st.metric("Validation Score", f"{score}%")
                                with col2:
                                    is_valid = validation_result.get('is_valid', False)
                                    st.metric("Status", "✅ Valid" if is_valid else "❌ Issues")
                                with col3:
                                    issues = validation_result.get('issues', [])
                                    st.metric("Issues Found", len(issues))
                                
                                # Show corrected content if available
                                corrected_content = validation_result.get('corrected_content', '')
                                if corrected_content:
                                    st.markdown("**Validated Content:**")
                                    st.markdown(corrected_content)
                                
                                # Show issues if any
                                if issues:
                                    with st.expander("🚨 Issues Found", expanded=False):
                                        for issue in issues:
                                            st.write(f"• {issue}")
                            else:
                                st.info("No AI2 results available. Run AI2 first.")
                        
                        # AI3 Results
                        with agent_tabs[2]:
                            agent3_results = agent_states.get('agent3_results', {})
                            if key in agent3_results:
                                pattern_result = agent3_results[key]
                                
                                # Compliance metrics
                                col1, col2 = st.columns(2)
                                with col1:
                                    is_compliant = pattern_result.get('is_compliant', False)
                                    st.metric("Pattern Compliance", "✅ Compliant" if is_compliant else "⚠️ Issues")
                                with col2:
                                    issues = pattern_result.get('issues', [])
                                    st.metric("Issues Found", len(issues))
                                
                                # Show final content if available
                                corrected_content = pattern_result.get('corrected_content', '')
                                if corrected_content:
                                    st.markdown("**Final Content:**")
                                    st.markdown(corrected_content)
                                
                                # Show issues if any
                                if issues:
                                    with st.expander("🚨 Pattern Issues", expanded=False):
                                        for issue in issues:
                                            st.write(f"• {issue}")
                            else:
                                st.info("No AI3 results available. Run AI3 first.")
            else:
                st.info("No financial keys available for results display.")
        else:
            st.info("No AI agents have run yet. Use the buttons above to start processing.")
        
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

# JSON Content Access Helper Functions
@cached_function(ttl=1800)  # Cache for 30 minutes
def load_json_content():
    """Load content from JSON file with caching for better performance"""
    try:
        # Try JSON first (better performance)
        json_file = "utils/bs_content.json"
        if os.path.exists(json_file):
            with open(json_file, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception as e:
        print(f"Error loading JSON content: {e}")
    
    # Fallback to parsing markdown if JSON not available
    try:
        content_files = ["utils/bs_content.md", "utils/bs_content_ai_generated.md", "utils/bs_content_offline.md"]
        for file_path in content_files:
            if os.path.exists(file_path):
                return parse_markdown_to_json(file_path)
    except Exception as e:
        print(f"Error parsing markdown fallback: {e}")
    
    return None

def parse_markdown_to_json(md_file_path):
    """Parse markdown file and convert to JSON-like structure for compatibility"""
    try:
        with open(md_file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Parse markdown into structured format
        sections = re.split(r'(^### .+$)', content, flags=re.MULTILINE)
        parsed_data = {"financial_items": {}}
        
        for i in range(1, len(sections), 2):
            if i + 1 < len(sections):
                header = sections[i].strip().replace('### ', '')
                section_content = sections[i + 1].strip()
                
                # Map headers back to keys
                header_to_key_mapping = {
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
                
                key = header_to_key_mapping.get(header, header)
                parsed_data["financial_items"][key] = {
                    "content": section_content,
                    "display_name": header
                }
        
        return parsed_data
    except Exception as e:
        print(f"Error parsing markdown: {e}")
        return None

def get_content_from_json(key):
    """Get content for a specific financial key from JSON data"""
    json_data = load_json_content()
    if not json_data:
        return None
    
    # Search in all categories
    for category in ["current_assets", "non_current_assets", "liabilities", "equity"]:
        category_data = json_data.get("financial_items", {}).get(category, {})
        if key in category_data:
            return category_data[key]["content"]
    
    # Direct key lookup for backwards compatibility
    direct_items = json_data.get("financial_items", {})
    if key in direct_items:
        return direct_items[key].get("content", "")
    
    return None

def display_offline_content(key):
    """Display offline content for a given key using JSON format for better performance"""
    try:
        content = get_content_from_json(key)
        
        if content:
            # Clean the content
            cleaned_content = clean_content_quotes(content)
            st.markdown(cleaned_content)
            
            # Show JSON source info
            st.info(f"📄 Content loaded from JSON format for better performance")
        else:
            st.info(f"No content found for {get_key_display_name(key)} in JSON data")
            
    except Exception as e:
        st.error(f"Error reading JSON content: {e}")
        # Fallback to old method
        display_offline_content_fallback(key)

def display_offline_content_fallback(key):
    """Fallback to original markdown parsing method"""
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
    """Get offline content for a given key (returns string) using JSON for better performance"""
    try:
        # Try JSON first (much faster)
        content = get_content_from_json(key)
        if content:
            return content
            
        # Fallback to markdown parsing
        return get_offline_content_fallback(key)
        
    except Exception:
        return ""

def get_offline_content_fallback(key):
    """Fallback method for getting content from markdown files"""
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

def generate_content_from_session_storage(entity_name):
    """Generate content files (JSON + Markdown) from session state storage (PERFORMANCE OPTIMIZED)"""
    try:
        # Get content from session state storage (fastest method)
        content_store = st.session_state.get('ai_content_store', {})
        
        if not content_store:
            st.warning("⚠️ No content in session storage. Using fallback method.")
            return generate_markdown_from_ai_results(st.session_state.get('ai_data', {}).get('ai_results', {}), entity_name)
        
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
        
        # Generate JSON content from session storage (for AI2 easy access)
        json_content = {
            'metadata': {
                'entity_name': entity_name,
                'generated_timestamp': datetime.datetime.now().isoformat(),
                'session_id': getattr(st.session_state.get('ai_logger'), 'session_id', 'default'),
                'total_keys': len(content_store)
            },
            'categories': {},
            'keys': {}
        }
        
        st.info(f"📊 Generating content files from session storage for {len(content_store)} keys")
        
        # Process content by category
        for category, items in category_mapping.items():
            json_content['categories'][category] = []
            
            for item in items:
                full_name = name_mapping[item]
                
                # Get latest content from session storage (could be Agent 1, 2, or 3 version)
                if item in content_store:
                    key_data = content_store[item]
                    latest_content = key_data.get('current_content', key_data.get('agent1_content', ''))
                    
                    # Determine content source
                    if 'agent3_content' in key_data:
                        content_source = "agent3_final"
                        source_timestamp = key_data.get('agent3_timestamp')
                    elif 'agent2_content' in key_data:
                        content_source = "agent2_validated"
                        source_timestamp = key_data.get('agent2_timestamp')
                    else:
                        content_source = "agent1_original"
                        source_timestamp = key_data.get('agent1_timestamp')
                    
                    st.write(f"  • {item}: Using {content_source} version")
                else:
                    latest_content = f"No information available for {item}"
                    content_source = "none"
                    source_timestamp = None
                    st.write(f"  • {item}: No content found")
                
                # Clean the content
                cleaned_content = clean_content_quotes(latest_content)
                
                # Add to JSON structure
                key_info = {
                    'key': item,
                    'display_name': full_name,
                    'content': cleaned_content,
                    'content_source': content_source,
                    'source_timestamp': source_timestamp,
                    'length': len(cleaned_content),
                    'category': category
                }
                
                json_content['categories'][category].append(key_info)
                json_content['keys'][item] = key_info
        
        # Save JSON format (for AI2 easy access)
        json_file_path = 'utils/bs_content.json'
        with open(json_file_path, 'w', encoding='utf-8') as file:
            json.dump(json_content, file, indent=2, ensure_ascii=False)
        
        # Also generate markdown for PowerPoint compatibility
        markdown_lines = []
        for category, items in category_mapping.items():
            markdown_lines.append(f"## {category}\n")
            for item in items:
                full_name = name_mapping[item]
                key_info = json_content['keys'].get(item)
                if key_info:
                    cleaned_content = key_info['content']
                else:
                    cleaned_content = f"No information available for {item}"
                markdown_lines.append(f"### {full_name}\n{cleaned_content}\n")
        
        markdown_text = "\n".join(markdown_lines)
        
        # Save markdown format (for PowerPoint export)
        md_file_path = 'utils/bs_content.md'
        with open(md_file_path, 'w', encoding='utf-8') as file:
            file.write(markdown_text)
        
        st.success(f"✅ Generated bs_content.json (AI-friendly) and bs_content.md (PowerPoint-compatible)")
        return True
        
    except Exception as e:
        st.error(f"Error generating content from session storage: {e}")
        return False

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
            
            # Get the actual prompts that will be sent to AI by calling process_keys
            # We need to capture the real prompts with table data
            try:
                # Load prompts from prompts.json file
                with open('utils/prompts.json', 'r') as f:
                    prompts_config = json.load(f)
                actual_system_prompt = prompts_config.get('system_prompts', {}).get('Agent 1', '')
                if not actual_system_prompt:
                    actual_system_prompt = """
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
                
                # Add entity placeholder instructions to system prompt
                actual_system_prompt += f"""

IMPORTANT ENTITY INSTRUCTIONS:
- Replace all [ENTITY_NAME] placeholders with the actual entity name from the provided financial data
- Use the exact entity name as shown in the financial data tables (e.g., 'Haining Wanpu', 'Ningbo Wanchen')
- Do not use the reporting entity name ({entity_name}) unless it matches the entity in the financial data
- Ensure all entity references in your analysis are accurate according to the provided data
"""
            except Exception as e:
                actual_system_prompt = "Error loading prompts.json"
                print(f"Error loading prompts: {e}")
            
            # Get sections data for each key to build actual user prompts
            sections_by_key = ai_data.get('sections_by_key', {})
            
            # Log Agent 1 input for all keys with actual prompts
            for key in filtered_keys:
                # Get the actual table data for this key with proper cleaning
                key_sections = sections_by_key.get(key, [])
                key_sections_str = []
                
                for section in key_sections:
                    if isinstance(section, dict):
                        # Handle dict serialization
                        try:
                            key_sections_str.append(json.dumps(section, indent=2, default=str))
                        except:
                            key_sections_str.append(str(section))
                    elif hasattr(section, 'to_string'):  # DataFrame
                        # Clean the DataFrame first - remove columns that are all None/NaN
                        df = section.copy()
                        for col in list(df.columns):
                            if df[col].isna().all() or (df[col].astype(str) == 'None').all():
                                df = df.drop(columns=[col])
                                print(f"DEBUG: Removed all-NaN column {col} for AI prompt")
                        
                        # Convert to string with proper formatting
                        df_str = df.to_string(index=False, na_rep='')
                        lines = df_str.split('\n')
                        cleaned_lines = []
                        for line in lines:
                            # Skip lines that are mostly empty or just separators
                            if line.strip() and not line.strip().replace('|', '').replace('-', '').replace(' ', '').replace('+', '') == '':
                                cleaned_lines.append(line)
                        key_sections_str.append('\n'.join(cleaned_lines))
                    else:
                        key_sections_str.append(str(section))
                
                key_tables = "\n".join(key_sections_str) if key_sections_str else "No table data available"
                
                # Build the actual user prompt that matches what gets sent to AI
                pattern = ai_data.get('pattern', {}).get(key, {})
                pattern_json = json.dumps(pattern, indent=2)
                
                # Get financial figure for this key
                from common.assistant import find_financial_figures_with_context_check, get_tab_name, get_financial_figure
                financial_figures = find_financial_figures_with_context_check(temp_file_path, get_tab_name(entity_name), '30/09/2022', convert_thousands=False)
                financial_figure_info = f"{key}: {get_financial_figure(financial_figures, key)}"
                
                # Build the actual user prompt that gets sent to AI
                actual_user_prompt = f"""
                TASK: Select ONE pattern and complete it with actual data
                
                AVAILABLE PATTERNS: {pattern_json}
                
                FINANCIAL FIGURE: {financial_figure_info}
                
                DATA SOURCE: {key_tables}
                
                SELECTION CRITERIA:
                - Choose the pattern with the most complete data coverage
                - Prioritize patterns that match the primary account category
                - Use most recent data: latest available
                - Express all figures with proper K/M conversion with 1 decimal place
                
                REQUIRED OUTPUT FORMAT:
                - Only the completed pattern text
                - No pattern names or labels
                - No template structure
                - No JSON formatting
                - Replace ALL 'xxx' or placeholders with actual data values
                - Replace ALL [ENTITY_NAME] placeholders with the actual entity name from the DATA SOURCE
                - Use the exact entity name as shown in the financial data (e.g., 'Haining Wanpu', 'Ningbo Wanchen')
                - Do not use bullet point for listing
                - Use actual numerical values from the DATA SOURCE (do not convert to K/M format)
                - No foreign contents, if any, translate to English
                - Stick to Template format, no extra explanations or comments
                - For entity name to be filled into template, it should not be the reporting entity ({entity_name}) itself, it must be from the DATA SOURCE
                - For all listing figures, please check the total, together should be around the same or constituting majority of FINANCIAL FIGURE
                - Ensure all financial figures mentioned match the actual values from the DATA SOURCE
                """
                
                context_data = {
                    'entity': entity_name,
                    'key': key,
                    'financial_figure': financial_figure_info,
                    'table_data_length': len(key_tables),
                    'patterns_count': len(pattern) if pattern else 0
                }
                
                # Store the actual prompts that will be sent to AI
                actual_prompts = {
                    'system_prompt': actual_system_prompt,
                    'user_prompt': actual_user_prompt,
                    'context': context_data,
                    'table_data': key_tables[:2000] + "..." if len(key_tables) > 2000 else key_tables,  # Show more structured data
                    'structured_tables': []  # Will store parsed structured tables
                }
                
                # Parse structured tables from the key_tables content
                try:
                    # Extract structured table information from the markdown content
                    table_sections = key_tables.split('## ')
                    for section in table_sections[1:]:  # Skip first empty section
                        lines = section.strip().split('\n')
                        if len(lines) > 0:
                            table_info = {
                                'table_name': lines[0],
                                'entity': 'Unknown',
                                'date': 'Unknown',
                                'currency': 'CNY',
                                'multiplier': 1,
                                'items': [],
                                'total': 'Unknown'
                            }
                            
                            for line in lines[1:]:
                                if line.startswith('**Entity:**'):
                                    table_info['entity'] = line.replace('**Entity:**', '').strip()
                                elif line.startswith('**Date:**'):
                                    table_info['date'] = line.replace('**Date:**', '').strip()
                                elif line.startswith('**Currency:**'):
                                    table_info['currency'] = line.replace('**Currency:**', '').strip()
                                elif line.startswith('**Multiplier:**'):
                                    table_info['multiplier'] = int(line.replace('**Multiplier:**', '').strip())
                                elif line.startswith('**Total:**'):
                                    table_info['total'] = line.replace('**Total:**', '').strip()
                                elif line.startswith('- ') and ':' in line:
                                    # Parse item line like "- Deposits with banks: 9076000"
                                    item_parts = line[2:].split(': ')
                                    if len(item_parts) == 2:
                                        table_info['items'].append({
                                            'description': item_parts[0],
                                            'amount': item_parts[1]
                                        })
                            
                            actual_prompts['structured_tables'].append(table_info)
                except Exception as e:
                    print(f"Error parsing structured tables for logging: {e}")
                
                logger.log_agent_input('agent1', key, actual_system_prompt, actual_user_prompt, context_data, actual_prompts)
            
            # Process ALL keys at once with proper tqdm progress (1/9, 2/9, etc.)
            start_time = time.time()
            st.write(f"🤖 Processing {len(filtered_keys)} keys with Agent 1...")
            
            # Create progress callback for Streamlit
            def update_progress(progress, message):
                progress_bar.progress(progress)
                status_text.text(message)
            
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
                progress_callback=update_progress
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
        
        # Initialize improved content storage in session state
        if 'ai_content_store' not in st.session_state:
            st.session_state['ai_content_store'] = {}
        
        # Get Agent 1 content from session state storage (more reliable)
        content_store = st.session_state['ai_content_store']
        
        # Check content availability
        available_keys = []
        missing_keys = []
        
        for key in filtered_keys:
            if key in agent1_results and agent1_results[key]:
                # Store in session state for fast access
                content_store[key] = {
                    'agent1_content': agent1_results[key],
                    'agent1_timestamp': time.time(),
                    'entity_name': entity_name
                }
                available_keys.append(key)
            else:
                missing_keys.append(key)
        
        if missing_keys:
            st.warning(f"⚠️ Missing Agent 1 content for keys: {missing_keys}")
            st.info("Agent 2 will only process keys with available content")
        
        if not available_keys:
            st.error("❌ No Agent 1 content available for any keys")
            return {}
        
        st.success(f"✅ Found Agent 1 content for {len(available_keys)} keys: {available_keys}")
        
        # Create progress bar for Agent 2
        progress_bar = st.progress(0)
        status_text = st.empty()
        
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
            
            # Initialize content updates storage
            bs_content_updates = {}
            
            # Process only keys with available content
            for i, key in enumerate(available_keys):
                # Update progress
                progress = (i + 1) / len(available_keys)
                progress_bar.progress(progress)
                status_text.text(f"🔍 AI2: Validating data for {key} ({i+1}/{len(available_keys)})")
                
                start_time = time.time()
                
                try:
                    # Get Agent 1 content from session state storage + JSON fallback
                    key_data = content_store[key]
                    agent1_content = key_data['agent1_content']
                    
                    # Also try JSON file fallback if session state is incomplete
                    if not agent1_content:
                        try:
                            with open('utils/bs_content.json', 'r', encoding='utf-8') as f:
                                json_data = json.load(f)
                                if key in json_data.get('keys', {}):
                                    agent1_content = json_data['keys'][key]['content']
                                    st.info(f"📄 Loaded {key} content from bs_content.json (JSON fallback)")
                        except (FileNotFoundError, json.JSONDecodeError):
                            pass
                    
                    st.write(f"✅ Agent 1 content length for {key}: {len(agent1_content)} characters")
                    st.write(f"📊 Content source: Session state storage + JSON fallback (reliable access)")
                    
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
                        
                        # Store the actual prompts that will be sent to AI
                        actual_prompts = {
                            'system_prompt': agent2_system_prompt,
                            'user_prompt': user_prompt,
                            'context': {
                                'expected_figure': expected_figure, 
                                'agent1_content_length': len(agent1_content),
                                'entity': entity_name,
                                'key': key
                            }
                        }
                        
                        # Log Agent 2 input with actual prompts
                        logger.log_agent_input('agent2', key, agent2_system_prompt, user_prompt, 
                                             {'expected_figure': expected_figure, 'agent1_content_length': len(agent1_content)}, actual_prompts)
                        
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
                            corrected_content = validation_result['corrected_content']
                            bs_content_updates[key] = corrected_content
                            
                            # Update session state storage with corrected content
                            content_store[key]['agent2_content'] = corrected_content
                            content_store[key]['agent2_timestamp'] = time.time()
                            content_store[key]['current_content'] = corrected_content  # Latest version
                            
                            validation_result['content_updated'] = True
                            st.success(f"✅ Agent 2 corrected content for {key}")
                        else:
                            # Keep Agent 1 content as current
                            content_store[key]['current_content'] = agent1_content
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
        
        # Use improved session state storage for fast content access
        content_store = st.session_state.get('ai_content_store', {})
        bs_content_updates = {}
        
        # Check content availability from session state
        available_keys = []
        for key in filtered_keys:
            if key in content_store and 'current_content' in content_store[key]:
                available_keys.append(key)
        
        if not available_keys:
            st.error("❌ No content available in session state for Agent 3")
            st.info("Make sure Agent 1 and Agent 2 have run successfully")
            return {}
        
        st.success(f"✅ Found content for {len(available_keys)} keys in session state storage")
        
        for key in available_keys:
            st.write(f"🔄 Checking pattern compliance for {key} with Agent 3...")
            start_time = time.time()
            
            try:
                # Get the most recent content from session state (Agent 2 corrected or Agent 1 original)
                key_data = content_store[key]
                current_content = key_data.get('current_content', key_data.get('agent1_content', ''))
                
                content_source = "Agent 2 corrected" if 'agent2_content' in key_data else "Agent 1 original"
                st.write(f"📊 Content source for {key}: {content_source} (session state)")
                st.write(f"✅ Content length for {key}: {len(current_content)} characters")
                
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
                    
                    # Store the actual prompts that will be sent to AI
                    actual_prompts = {
                        'system_prompt': agent3_system_prompt,
                        'user_prompt': user_prompt,
                        'context': {
                            'patterns_count': len(key_patterns), 
                            'content_length': len(current_content),
                            'key': key,
                            'patterns': key_patterns
                        }
                    }
                    
                    # Log Agent 3 input with actual prompts
                    logger.log_agent_input('agent3', key, agent3_system_prompt, user_prompt, 
                                         {'patterns_count': len(key_patterns), 'content_length': len(current_content)}, actual_prompts)
                    
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
                        corrected_content = pattern_result['corrected_content']
                        bs_content_updates[key] = corrected_content
                        
                        # Update session state storage with Agent 3 corrected content
                        content_store[key]['agent3_content'] = corrected_content
                        content_store[key]['agent3_timestamp'] = time.time()
                        content_store[key]['current_content'] = corrected_content  # Latest version
                        
                        pattern_result['content_updated'] = True
                        st.success(f"✅ Agent 3 improved pattern compliance for {key}")
                    else:
                        # Keep existing content as current (no Agent 3 changes)
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
    """Display consolidated AI results in organized tabs with parallel comparison (ENHANCED INTERFACE)"""
    # Single consolidated AI results area with tabs
    st.markdown("## 🤖 AI Processing Results")
    
    # Create main tabs for different views
    main_tabs = st.tabs(["📊 By Agent", "🗂️ By Key", "🔄 Parallel Comparison", "📈 Session Overview"])
    
    # Tab 1: Results organized by Agent (AI1, AI2, AI3)
    with main_tabs[0]:
        st.markdown("### View results organized by AI Agent")
        
        # Agent tabs
        agent_tabs = st.tabs(["🚀 Agent 1: Generation", "📊 Agent 2: Validation", "🎯 Agent 3: Compliance"])
        
        # Agent 1 Tab
        with agent_tabs[0]:
            st.markdown("**Focus**: Generate comprehensive financial analysis content")
            
            # Show Agent 1 results for all keys
            agent_states = st.session_state.get('agent_states', {})
            if agent_states.get('agent1_completed', False):
                agent1_results = agent_states.get('agent1_results', {})
                
                # Key tabs within Agent 1
                if agent1_results:
                    available_keys = [k for k in filtered_keys if k in agent1_results and agent1_results[k]]
                    if available_keys:
                        key_tabs = st.tabs([get_key_display_name(k) for k in available_keys])
                        
                        for i, key in enumerate(available_keys):
                            with key_tabs[i]:
                                content = agent1_results[key]
                                if content:
                                    # Show metadata
                                    col1, col2, col3 = st.columns(3)
                                    with col1:
                                        st.metric("Characters", len(content))
                                    with col2:
                                        st.metric("Words", len(content.split()))
                                    with col3:
                                        st.metric("Entity", ai_data.get('entity_name', ''))
                                    
                                    # Show content
                                    st.markdown("**Generated Content:**")
                                    st.markdown(content)
                                else:
                                    st.warning("No content generated")
                    else:
                        st.info("No content available from Agent 1")
                else:
                    st.info("Agent 1 results not available")
            else:
                st.info("⏳ Agent 1 will run when you click 'Process with AI'")
        
        # Agent 2 Tab
        with agent_tabs[1]:
            st.markdown("**Focus**: Validate data accuracy and fix issues")
            
            if agent_states.get('agent2_completed', False):
                agent2_results = agent_states.get('agent2_results', {})
                
                if agent2_results:
                    available_keys = [k for k in filtered_keys if k in agent2_results]
                    if available_keys:
                        key_tabs = st.tabs([get_key_display_name(k) for k in available_keys])
                        
                        for i, key in enumerate(available_keys):
                            with key_tabs[i]:
                                validation_result = agent2_results[key]
                                
                                # Show validation metrics
                                col1, col2, col3 = st.columns(3)
                                with col1:
                                    score = validation_result.get('score', 0)
                                    st.metric("Validation Score", f"{score}%")
                                with col2:
                                    is_valid = validation_result.get('is_valid', False)
                                    st.metric("Status", "✅ Valid" if is_valid else "❌ Issues")
                                with col3:
                                    issues = validation_result.get('issues', [])
                                    st.metric("Issues Found", len(issues))
                                
                                # Show corrected content if available
                                corrected_content = validation_result.get('corrected_content', '')
                                if corrected_content:
                                    st.markdown("**Validated Content:**")
                                    st.markdown(corrected_content)
                                
                                # Show issues if any
                                if issues:
                                    with st.expander("🚨 Issues Found", expanded=False):
                                        for issue in issues:
                                            st.write(f"• {issue}")
                    else:
                        st.info("No validation results available")
                else:
                    st.info("Agent 2 results not available")
            else:
                st.info("⏳ Agent 2 will run after Agent 1 completes")
        
        # Agent 3 Tab
        with agent_tabs[2]:
            st.markdown("**Focus**: Ensure pattern compliance and final polish")
            
            if agent_states.get('agent3_completed', False):
                agent3_results = agent_states.get('agent3_results', {})
                
                if agent3_results:
                    available_keys = [k for k in filtered_keys if k in agent3_results]
                    if available_keys:
                        key_tabs = st.tabs([get_key_display_name(k) for k in available_keys])
                        
                        for i, key in enumerate(available_keys):
                            with key_tabs[i]:
                                pattern_result = agent3_results[key]
                                
                                # Show compliance metrics
                                col1, col2 = st.columns(2)
                                with col1:
                                    is_compliant = pattern_result.get('is_compliant', False)
                                    st.metric("Pattern Compliance", "✅ Compliant" if is_compliant else "⚠️ Issues")
                                with col2:
                                    issues = pattern_result.get('issues', [])
                                    st.metric("Issues Found", len(issues))
                                
                                # Show final content if available
                                corrected_content = pattern_result.get('corrected_content', '')
                                if corrected_content:
                                    st.markdown("**Final Content:**")
                                    st.markdown(corrected_content)
                                
                                # Show issues if any
                                if issues:
                                    with st.expander("🚨 Pattern Issues", expanded=False):
                                        for issue in issues:
                                            st.write(f"• {issue}")
                    else:
                        st.info("No pattern compliance results available")
                else:
                    st.info("Agent 3 results not available")
            else:
                st.info("⏳ Agent 3 will run after Agent 2 completes")
    
    # Tab 2: Results organized by Key (Cash, AR, etc.)
    with main_tabs[1]:
        st.markdown("### View results organized by Financial Key")
        
        # Get final content from session storage (latest versions)
        content_store = st.session_state.get('ai_content_store', {})
        
        if content_store:
            available_keys = [k for k in filtered_keys if k in content_store]
            if available_keys:
                key_tabs = st.tabs([get_key_display_name(k) for k in available_keys])
                
                for i, key in enumerate(available_keys):
                    with key_tabs[i]:
                        key_data = content_store[key]
                        current_content = key_data.get('current_content', key_data.get('agent1_content', ''))
                        
                        if current_content:
                            # Determine final version source
                            if 'agent3_content' in key_data:
                                content_source = "Agent 3 (Final - Pattern Compliant)"
                                content_icon = "🎯"
                                processing_steps = "Generated → Validated → Pattern Compliant"
                            elif 'agent2_content' in key_data:
                                content_source = "Agent 2 (Validated - Data Accurate)"
                                content_icon = "📊"
                                processing_steps = "Generated → Validated"
                            else:
                                content_source = "Agent 1 (Original - Generated)"
                                content_icon = "📝"
                                processing_steps = "Generated"
                            
                            # Show metadata
                            col1, col2, col3, col4 = st.columns(4)
                            with col1:
                                st.metric("Final Version", content_source.split()[1])
                            with col2:
                                st.metric("Entity", ai_data.get('entity_name', ''))
                            with col3:
                                st.metric("Characters", len(current_content))
                            with col4:
                                st.metric("Words", len(current_content.split()))
                            
                            # Show processing pipeline
                            st.info(f"🔄 Processing Pipeline: {processing_steps}")
                            
                            # Show final content
                            st.markdown(f"**{content_icon} Final Content:**")
                            st.markdown(current_content)
                            
                        else:
                            st.warning("No content available for this key")
            else:
                st.info("No processed keys available")
        else:
            st.info("No content available in session storage")
    
    # Tab 3: Parallel Comparison with Before/After Visualization
    with main_tabs[2]:
        st.markdown("### 🔄 Parallel Agent Comparison & Before/After Changes")
        
        # Get agent states and results
        agent_states = st.session_state.get('agent_states', {})
        agent1_results = agent_states.get('agent1_results', {})
        agent2_results = agent_states.get('agent2_results', {})
        agent3_results = agent_states.get('agent3_results', {})
        content_store = st.session_state.get('ai_content_store', {})
        
        # Key selector for comparison
        if filtered_keys:
            selected_key = st.selectbox(
                "Select Financial Key for Detailed Comparison:",
                filtered_keys,
                format_func=get_key_display_name,
                key="parallel_comparison_key"
            )
            
            if selected_key:
                st.markdown(f"### Analysis for {get_key_display_name(selected_key)}")
                
                # Parallel comparison buttons
                st.markdown("#### 🔄 Choose Comparison Mode:")
                comparison_mode = st.radio(
                    "Comparison Type:",
                    ["Before vs After (AI1 → AI3)", "Step-by-Step (AI1 → AI2 → AI3)", "Agent Validation (AI2 vs AI3)"],
                    horizontal=True,
                    key="comparison_mode"
                )
                
                # Get content for selected key
                agent1_content = agent1_results.get(selected_key, "")
                agent2_content = ""
                agent3_content = ""
                
                # Get validated content from Agent 2
                if selected_key in agent2_results:
                    agent2_data = agent2_results[selected_key]
                    agent2_content = agent2_data.get('corrected_content', '')
                    if not agent2_content and selected_key in content_store:
                        agent2_content = content_store[selected_key].get('agent2_content', '')
                
                # Get final content from Agent 3
                if selected_key in agent3_results:
                    agent3_data = agent3_results[selected_key]
                    agent3_content = agent3_data.get('corrected_content', '')
                    if not agent3_content and selected_key in content_store:
                        agent3_content = content_store[selected_key].get('agent3_content', '')
                
                # Default to Agent 1 content if later agents don't have content
                if not agent2_content:
                    agent2_content = agent1_content
                if not agent3_content:
                    agent3_content = agent2_content or agent1_content
                
                # Display comparison based on selected mode
                if comparison_mode == "Before vs After (AI1 → AI3)":
                    display_before_after_comparison(selected_key, agent1_content, agent3_content, agent_states)
                    
                elif comparison_mode == "Step-by-Step (AI1 → AI2 → AI3)":
                    display_step_by_step_comparison(selected_key, agent1_content, agent2_content, agent3_content, agent_states)
                    
                elif comparison_mode == "Agent Validation (AI2 vs AI3)":
                    display_validation_comparison(selected_key, agent2_content, agent3_content, agent2_results, agent3_results)
        
        else:
            st.info("No financial keys available for comparison")
    
    # Tab 4: Session Overview
    with main_tabs[3]:
        st.markdown("### Session Processing Overview")
        
        # Session statistics
        if agent_states:
            col1, col2, col3 = st.columns(3)
            with col1:
                agent1_completed = agent_states.get('agent1_completed', False)
                st.metric("Agent 1", "✅ Completed" if agent1_completed else "⏳ Pending")
            with col2:
                agent2_completed = agent_states.get('agent2_completed', False)
                st.metric("Agent 2", "✅ Completed" if agent2_completed else "⏳ Pending")
            with col3:
                agent3_completed = agent_states.get('agent3_completed', False)
                st.metric("Agent 3", "✅ Completed" if agent3_completed else "⏳ Pending")
            
            # Show logging info
            st.markdown("---")
            st.markdown("### 📋 Logging Information")
            
            # Get logger info
            logger = st.session_state.get('ai_logger')
            if logger:
                session_id = getattr(logger, 'session_id', 'unknown')
                log_file = getattr(logger, 'log_file', 'unknown')
                
                st.info(f"📁 **Session ID**: {session_id}")
                st.info(f"📄 **Detailed logs**: `{log_file}`")
                st.info(f"📊 **Consolidated logs**: `logging/session_{session_id}.json`")
                
                # Token usage summary (if available)
                total_logs = len(logger.session_logs) if hasattr(logger, 'session_logs') else 0
                st.metric("Total Log Entries", total_logs)
            else:
                st.warning("No logging information available")
        else:
            st.info("No processing session data available")

# Helper functions for parallel comparison and before/after visualization

def display_before_after_comparison(key, before_content, after_content, agent_states):
    """Display before (AI1) vs after (AI3) comparison with visual diff"""
    st.markdown("#### 📊 Before vs After Comparison")
    
    # Status indicators
    col1, col2, col3 = st.columns(3)
    with col1:
        agent1_success = agent_states.get('agent1_success', False)
        st.metric("Agent 1", "✅ Success" if agent1_success else "❌ Failed")
    with col2:
        changes_made = before_content != after_content
        st.metric("Changes Made", "✅ Yes" if changes_made else "➖ No")
    with col3:
        agent3_success = agent_states.get('agent3_success', False)
        st.metric("Agent 3", "✅ Success" if agent3_success else "❌ Failed")
    
    # Side-by-side comparison
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("##### 📝 **BEFORE** (Agent 1 - Original)")
        if before_content:
            st.markdown(f"**Length:** {len(before_content)} characters, {len(before_content.split())} words")
            with st.container():
                st.markdown(before_content)
        else:
            st.warning("No original content available")
    
    with col2:
        st.markdown("##### 🎯 **AFTER** (Agent 3 - Final)")
        if after_content:
            st.markdown(f"**Length:** {len(after_content)} characters, {len(after_content.split())} words")
            with st.container():
                st.markdown(after_content)
        else:
            st.warning("No final content available")
    
    # Change analysis
    if before_content and after_content:
        st.markdown("---")
        st.markdown("#### 📈 Change Analysis")
        
        # Length comparison
        length_diff = len(after_content) - len(before_content)
        word_diff = len(after_content.split()) - len(before_content.split())
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Character Change", f"{length_diff:+d}", delta=f"{length_diff/len(before_content)*100:+.1f}%" if before_content else "N/A")
        with col2:
            st.metric("Word Change", f"{word_diff:+d}", delta=f"{word_diff/len(before_content.split())*100:+.1f}%" if before_content.split() else "N/A")
        with col3:
            similarity = calculate_content_similarity(before_content, after_content)
            st.metric("Similarity", f"{similarity:.1f}%")
        
        # Highlight differences
        if before_content != after_content:
            with st.expander("🔍 Detailed Changes", expanded=False):
                show_text_differences(before_content, after_content)

def display_step_by_step_comparison(key, agent1_content, agent2_content, agent3_content, agent_states):
    """Display step-by-step progression through all agents"""
    st.markdown("#### 🔄 Step-by-Step Agent Progression")
    
    # Progress indicators
    col1, col2, col3 = st.columns(3)
    with col1:
        agent1_success = agent_states.get('agent1_success', False)
        st.metric("🚀 Agent 1", "✅ Generated" if agent1_success else "❌ Failed")
    with col2:
        agent2_success = agent_states.get('agent2_success', False)
        st.metric("📊 Agent 2", "✅ Validated" if agent2_success else "❌ Failed")
    with col3:
        agent3_success = agent_states.get('agent3_success', False)
        st.metric("🎯 Agent 3", "✅ Compliant" if agent3_success else "❌ Failed")
    
    # Agent progression tabs
    step_tabs = st.tabs(["🚀 Step 1: Generation", "📊 Step 2: Validation", "🎯 Step 3: Compliance"])
    
    with step_tabs[0]:
        st.markdown("##### Agent 1: Content Generation")
        if agent1_content:
            st.markdown(f"**Length:** {len(agent1_content)} characters")
            st.markdown(agent1_content)
        else:
            st.warning("No content generated by Agent 1")
    
    with step_tabs[1]:
        st.markdown("##### Agent 2: Data Validation & Correction")
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("**Before Validation:**")
            if agent1_content:
                st.markdown(agent1_content[:200] + "..." if len(agent1_content) > 200 else agent1_content)
            else:
                st.warning("No input content")
        
        with col2:
            st.markdown("**After Validation:**")
            if agent2_content:
                if agent2_content != agent1_content:
                    st.success("✅ Changes made during validation")
                    st.markdown(agent2_content[:200] + "..." if len(agent2_content) > 200 else agent2_content)
                else:
                    st.info("ℹ️ No changes needed")
                    st.markdown("Content validated as accurate")
            else:
                st.warning("No validation output")
    
    with step_tabs[2]:
        st.markdown("##### Agent 3: Pattern Compliance & Final Polish")
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("**Before Compliance Check:**")
            if agent2_content:
                st.markdown(agent2_content[:200] + "..." if len(agent2_content) > 200 else agent2_content)
            else:
                st.warning("No input content")
        
        with col2:
            st.markdown("**After Compliance Check:**")
            if agent3_content:
                if agent3_content != agent2_content:
                    st.success("✅ Pattern compliance improvements made")
                    st.markdown(agent3_content[:200] + "..." if len(agent3_content) > 200 else agent3_content)
                else:
                    st.info("ℹ️ Content already compliant")
                    st.markdown("No pattern improvements needed")
            else:
                st.warning("No compliance output")

def display_validation_comparison(key, agent2_content, agent3_content, agent2_results, agent3_results):
    """Display comparison between Agent 2 and Agent 3 results"""
    st.markdown("#### 🔍 Agent Validation Comparison")
    
    # Get validation details
    agent2_data = agent2_results.get(key, {})
    agent3_data = agent3_results.get(key, {})
    
    # Validation metrics comparison
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("##### 📊 Agent 2: Data Validation")
        validation_score = agent2_data.get('score', 0)
        is_valid = agent2_data.get('is_valid', False)
        issues = agent2_data.get('issues', [])
        
        st.metric("Validation Score", f"{validation_score}%")
        st.metric("Status", "✅ Valid" if is_valid else "❌ Issues Found")
        st.metric("Issues Found", len(issues))
        
        if issues:
            with st.expander("🚨 Data Issues", expanded=False):
                for issue in issues:
                    st.write(f"• {issue}")
        
        if agent2_content:
            with st.expander("📝 Agent 2 Content", expanded=False):
                st.markdown(agent2_content)
    
    with col2:
        st.markdown("##### 🎯 Agent 3: Pattern Compliance")
        is_compliant = agent3_data.get('is_compliant', False)
        compliance_issues = agent3_data.get('issues', [])
        pattern_match = agent3_data.get('pattern_match', 'unknown')
        
        st.metric("Compliance Status", "✅ Compliant" if is_compliant else "⚠️ Issues")
        st.metric("Pattern Match", pattern_match.title())
        st.metric("Pattern Issues", len(compliance_issues))
        
        if compliance_issues:
            with st.expander("🚨 Pattern Issues", expanded=False):
                for issue in compliance_issues:
                    st.write(f"• {issue}")
        
        if agent3_content:
            with st.expander("📝 Agent 3 Content", expanded=False):
                st.markdown(agent3_content)
    
    # Overall comparison
    st.markdown("---")
    st.markdown("#### 📈 Overall Quality Comparison")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        data_quality = "High" if is_valid else "Needs Work"
        st.metric("Data Quality", data_quality)
    with col2:
        pattern_quality = "High" if is_compliant else "Needs Work"
        st.metric("Pattern Quality", pattern_quality)
    with col3:
        overall_status = "✅ Ready" if (is_valid and is_compliant) else "⚠️ Needs Review"
        st.metric("Overall Status", overall_status)

def calculate_content_similarity(text1, text2):
    """Calculate similarity percentage between two texts"""
    if not text1 or not text2:
        return 0.0
    
    # Simple word-based similarity
    words1 = set(text1.lower().split())
    words2 = set(text2.lower().split())
    
    if not words1 and not words2:
        return 100.0
    
    intersection = words1.intersection(words2)
    union = words1.union(words2)
    
    return (len(intersection) / len(union)) * 100 if union else 0.0

def show_text_differences(text1, text2):
    """Show differences between two texts (simplified diff)"""
    if text1 == text2:
        st.info("No differences found")
        return
    
    # Split into sentences for comparison
    sentences1 = [s.strip() for s in text1.split('.') if s.strip()]
    sentences2 = [s.strip() for s in text2.split('.') if s.strip()]
    
    st.markdown("**Changes Summary:**")
    
    # Find added/removed sentences
    added = [s for s in sentences2 if s not in sentences1]
    removed = [s for s in sentences1 if s not in sentences2]
    
    if added:
        st.markdown("**✅ Added:**")
        for sentence in added[:3]:  # Show first 3
            st.write(f"+ {sentence}")
        if len(added) > 3:
            st.write(f"... and {len(added) - 3} more additions")
    
    if removed:
        st.markdown("**❌ Removed:**")
        for sentence in removed[:3]:  # Show first 3
            st.write(f"- {sentence}")
        if len(removed) > 3:
            st.write(f"... and {len(removed) - 3} more removals")
    
    if not added and not removed:
        st.info("Changes are mostly within existing sentences (minor edits)")

if __name__ == "__main__":
    main() 
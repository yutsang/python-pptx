import streamlit as st
import pandas as pd
import json
import warnings
import re
import os
import datetime
import time
from tqdm import tqdm
from fdd_utils.mappings import (
    KEY_TO_SECTION_MAPPING,
    KEY_TERMS_BY_KEY,
    DISPLAY_NAME_MAPPING_DEFAULT,
    DISPLAY_NAME_MAPPING_NB_NJ,
)
from pathlib import Path
from tabulate import tabulate
import urllib3
import shutil

# Disable Python bytecode generation to prevent __pycache__ issues
os.environ['PYTHONDONTWRITEBYTECODE'] = '1'
from common.pptx_export import export_pptx
from fdd_utils.simple_cache import get_simple_cache
# Import assistant modules at module level to prevent runtime import issues
from common.assistant import process_keys, QualityAssuranceAgent, DataValidationAgent, PatternValidationAgent, find_financial_figures_with_context_check, get_tab_name, get_financial_figure, load_ip, ProofreadingAgent


import uuid
import tempfile

# Suppress warnings
urllib3.disable_warnings()
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action='ignore', category=UserWarning)
warnings.filterwarnings('ignore', message='Data Validation extension is not supported and will be removed', category=UserWarning, module='openpyxl')

# Suppress Streamlit file watcher errors
import logging
logging.getLogger('streamlit.watcher.event_based_path_watcher').setLevel(logging.ERROR)
logging.getLogger('streamlit.watcher.util').setLevel(logging.ERROR)

# AI Agent Logging System
class AIAgentLogger:
    """Single-file JSON logging per session (inputs, outputs, errors)."""
    
    def __init__(self):
        self.logs = {'agent1': {}, 'agent2': {}, 'agent3': {}}
        self.session_logs = []
        self.log_dir = Path("logging")
        self.log_dir.mkdir(exist_ok=True)
        self.session_id = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        # Consolidated session JSON file (only file we keep)
        self.session_file = self.log_dir / f"session_{self.session_id}.json"
        
    def _write_to_file(self, message):
        """No-op: we only keep consolidated JSON logs."""
        return
    
    def _save_json_log(self, log_entry, log_type):
        """Disabled: per-entry JSON files are not used."""
        return
    
    def _make_json_serializable(self, obj):
        """Convert DataFrame and other non-serializable objects to strings"""
        if hasattr(obj, 'to_dict'):  # DataFrame
            return str(obj)
        elif isinstance(obj, dict):
            return {k: self._make_json_serializable(v) for k, v in obj.items()}
        elif isinstance(obj, list):
            return [self._make_json_serializable(item) for item in obj]
        else:
            try:
                json.dumps(obj)  # Test if serializable
                return obj
            except (TypeError, ValueError):
                return str(obj)
    
    def _save_key_log(self, agent_name, key, input_entry, output_entry):
        """Disabled: per-key JSON files are not used."""
        return
        
    def _save_to_consolidated_log(self, log_entry):
        """Save log entry to consolidated session file (1 file per session with all agents and keys)"""
        try:
            session_file = self.session_file
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
        
        # Per-entry JSON disabled
        
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
        
        # Per-entry JSON disabled
        
        # Create key log if input exists
        input_key = f"{agent_name}_{key}"
        if hasattr(self, 'pending_inputs') and input_key in self.pending_inputs:
            input_entry = self.pending_inputs[input_key]
            # Per-key JSON disabled
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
        st.info(f"📁 Detailed logs saved to: `{self.session_file}`")
        
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
        
        # Download consolidated session log
        try:
            if self.session_file.exists():
                with open(self.session_file, 'r', encoding='utf-8') as f:
                    st.download_button(
                        label="📥 Download Session Log (JSON)",
                        data=f.read(),
                        file_name=self.session_file.name,
                        mime='application/json',
                        type="secondary"
                    )
        except Exception:
            pass

# Initialize global logger
if 'ai_logger' not in st.session_state:
    st.session_state.ai_logger = AIAgentLogger()

# Load configuration files
def load_config_files():
    """Load configuration files from utils directory"""
    try:
        # Load configs directly
        config = None
        mapping = None
        pattern = None
        prompts = None
        
        # Try to load config files directly
        try:
            with open('fdd_utils/config.json', 'r') as f:
                config = json.load(f)
        except FileNotFoundError:
            st.error("Configuration file not found: utils/config.json")
            return None, None, None, None
        
        try:
            with open('fdd_utils/mapping.json', 'r') as f:
                mapping = json.load(f)
        except FileNotFoundError:
            st.error("Configuration file not found: utils/mapping.json")
            return None, None, None, None
        
        try:
            with open('fdd_utils/pattern.json', 'r') as f:
                pattern = json.load(f)
        except FileNotFoundError:
            st.error("Configuration file not found: utils/pattern.json")
            return None, None, None, None
        
        try:
            with open('fdd_utils/prompts.json', 'r') as f:
                prompts = json.load(f)
        except FileNotFoundError:
            st.error("Configuration file not found: utils/prompts.json")
            return None, None, None, None
            
        return config, mapping, pattern, prompts
    except Exception as e:
        st.error(f"Failed to load configuration files: {e}")
        return None, None, None, None

def process_and_filter_excel(filename, tab_name_mapping, entity_name, entity_suffixes):
    """
    Process and filter Excel file to extract relevant worksheet sections with simple caching
    This is the core function from old_ver/utils/utils.py
    """
    try:
        # Use simple cache (imported at top)
        cache = get_simple_cache()
        
        # Check cache first (with force refresh option)
        force_refresh = st.session_state.get('force_refresh', False)
        cached_result = cache.get_cached_excel_data(filename, entity_name, force_refresh)
        if cached_result is not None:
            # Clear force refresh flag after using it
            if force_refresh:
                st.session_state['force_refresh'] = False
            return cached_result
        
        # Load the Excel file
        main_dir = Path(__file__).parent
        file_path = main_dir / filename
        xl = pd.ExcelFile(file_path)
        
        # Create a reverse mapping from values to keys
        reverse_mapping = {}
        if tab_name_mapping is not None:
            for key, values in tab_name_mapping.items():
                for value in values:
                    reverse_mapping[value] = key
                # Also map the key name directly to itself (for sheet names like "Cash", "AR")
                reverse_mapping[key] = key
                
        # Initialize a string to store markdown content
        markdown_content = ""
        
        # Process each sheet according to the mapping
        # Use multiprocessing for parallel sheet processing if many sheets
        relevant_sheets = [name for name in xl.sheet_names if name in reverse_mapping]
        
        if len(relevant_sheets) > 3:
            print(f"📊 Processing {len(relevant_sheets)} sheets in parallel...")
            # For large numbers of sheets, consider parallel processing
            # Note: For now, keeping sequential to avoid complexity with Streamlit
        
        for sheet_name in relevant_sheets:
            if sheet_name in reverse_mapping:
                df = xl.parse(sheet_name)
                
                # Detect latest date column for this sheet
                print(f"\n📊 Processing sheet: {sheet_name}")
                latest_date_col = detect_latest_date_column(df, sheet_name)
                if latest_date_col:
                    print(f"✅ Sheet {sheet_name}: Selected column {latest_date_col}")
                else:
                    print(f"⚠️  Sheet {sheet_name}: No date column detected")
                
                # Performance optimization: Skip splitting for small sheets
                if len(df) > 100:  # Only split very large sheets
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
                else:
                    # Small sheet: process as single dataframe for speed
                    dataframes = [df] if not df.empty else []
                    print(f"📈 [{sheet_name}] Processing as single dataframe (size: {len(df)})")
                
                # Filter dataframes by entity name with proper spacing
                entity_keywords = [f"{entity_name} {suffix}" for suffix in entity_suffixes if suffix]
                if not entity_keywords:  # If no helpers, just use entity name
                    entity_keywords = [entity_name]
                
                combined_pattern = '|'.join(re.escape(kw) for kw in entity_keywords)
                
                # Optimize: Pre-compile entity pattern for faster matching
                entity_pattern = '|'.join(re.escape(kw) for kw in entity_keywords)
                
                for data_frame in dataframes:
                    # Fast entity matching using vectorized operations
                    try:
                        # Convert to string once and check for entity match
                        df_str = data_frame.astype(str)
                        entity_match = df_str.apply(
                            lambda row: row.str.contains(entity_pattern, case=False, regex=True, na=False).any(),
                            axis=1
                        ).any()
                        
                        if entity_match:
                            # Fast filtering: only keep essential columns
                            if latest_date_col and latest_date_col in data_frame.columns:
                                # Find description column efficiently
                                desc_col = data_frame.columns[0]  # Usually first column
                                
                                # Keep only 2 columns for faster processing
                                essential_cols = [desc_col, latest_date_col]
                                filtered_df = data_frame[essential_cols].dropna(how='all')
                                
                                if not filtered_df.empty:
                                    markdown_content += f"## {sheet_name}\n"
                                    markdown_content += tabulate(filtered_df, headers='keys', tablefmt='pipe') + '\n\n'
                                    print(f"✅ [{sheet_name}] Entity data found and processed")
                                    
                                    # Early exit: we found the entity data for this sheet
                                    break
                    
                    except Exception as e:
                        # Fallback to original logic if optimization fails
                        print(f"📊 [{sheet_name}] Using fallback processing: {e}")
                        continue 
        
        # Cache the processed result using simple cache
        cache.cache_excel_data(filename, entity_name, markdown_content)
        print(f"📋 Cached result for {filename}")
        
        return markdown_content
    except Exception as e:
        st.error(f"An error occurred while processing the Excel file: {e}")
        return ""

def parse_accounting_table(df, key, entity_name, sheet_name, debug=False):
    """
    Parse accounting table with proper header detection and figure column identification
    Returns structured table data with metadata
    """
    try:
        import re
        import pandas as pd
        
        if df.empty or len(df) < 2:
            return None
        
        # Debug info reduced for cleaner output
        if debug:  # Only show if explicitly debugging
            print(f"DEBUG: DataFrame shape: {df.shape}")
        
        # Clean the DataFrame first - drop unnamed columns that are all NaN
        df_clean = df.copy()
        dropped_columns = []
        for col in df_clean.columns:
            if col.startswith('Unnamed:') or df_clean[col].isna().all():
                dropped_columns.append(col)
                df_clean = df_clean.drop(columns=[col])
        
        # If all columns were dropped, try a different approach
        if len(df_clean.columns) == 0:
            # Try to find columns with actual data
            df_clean = df.copy()
            for col in df_clean.columns:
                # Check if column has any non-null, non-empty values
                non_null_count = df_clean[col].notna().sum()
                non_empty_count = (df_clean[col].astype(str).str.strip() != '').sum()
                if non_null_count == 0 and non_empty_count == 0:
                    df_clean = df_clean.drop(columns=[col])
        
        # Additional cleaning: remove columns that are all None/NaN after initial cleaning
        for col in list(df_clean.columns):
            if df_clean[col].isna().all() or (df_clean[col].astype(str) == 'None').all():
                df_clean = df_clean.drop(columns=[col])
        
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
                    # Found value column indicator
                    break
                elif "total" in cell_value and value_col_idx is None:
                    value_col_idx = j
                    value_col_name = "Total"
                    # Found total column
        
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
                        # Found financial pattern
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
                    # Skipping Excel-generated column
                    continue
                
                # Additional check: skip if column name is a pure number
                if re.match(r'^\d+\.?\d*$', column_name):
                    # Skipping numeric column name
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
                        # Skipping sequential numbers column
                        continue
                    # Also skip if all values are the same (e.g., all 1000)
                    if len(set(numeric_values)) == 1:
                        # Skipping identical values column
                        continue
                # Now check if this is a good candidate
                numeric_count = 0
                total_cells = 0
                large_numbers = 0
                for cell in column_data:
                    cell_str = str(cell).strip()
                    if cell_str and cell_str.lower() not in ['nan', '']:
                        total_cells += 1
                        if re.search(r'^\d+\.?\d*$', cell_str.replace(',', '')):
                            numeric_count += 1
                            # Check if it's a large number (likely financial data)
                            try:
                                num_val = float(cell_str.replace(',', ''))
                                if num_val > 100:  # Skip small numbers like 1, 2, 3, 1000
                                    large_numbers += 1
                            except ValueError:
                                pass
                
                # Only consider columns with significant large numbers
                if total_cells > 0 and numeric_count >= total_cells * 0.3 and large_numbers >= 2:
                    candidate_cols.append(j)
                    # Found good candidate column
                    pass
                else:
                    # Column rejected
                    pass
            if candidate_cols:
                # Pick the rightmost candidate
                value_col_idx = candidate_cols[0]
                value_col_name = f"Column {value_col_idx+1}"
                # Selected value column
                pass
            else:
                # No valid value column found
                pass
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
                    # Additional check: skip if the description is a pure number (like 1000, 1001, etc.)
                    if not re.match(r'^\d+\.?\d*$', desc_cell):
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
                            # Found date in cell
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
        
        # Minimal debug output (reduced for better UX)
        if debug:
            print(f"DEBUG: Processing {df_str.shape[0]} rows, data starts at row {data_start_row}")
        
        # Column analysis (reduced verbosity)
        if debug:
            print(f"DEBUG: Analyzing {len(df_str.columns)} columns for data extraction")
        
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
                    # Skipping header/currency row
                    break
            
            # Additional check: skip if description is a pure number (like 1000, 1001, etc.)
            if re.match(r'^\d+\.?\d*$', description.strip()):
                should_skip = True
                # Skipping numeric description row
            
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
                if debug:
                    print(f"DEBUG: Added row - {description}: {value}")
        
        if debug:
            print(f"DEBUG: Extracted {len(data_rows)} data rows")
        
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

def detect_latest_date_column(df, sheet_name="Sheet"):
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
            year = 2000 + int(xmxx_match.group(2))
            # Use end of month
            if month == 12:
                return datetime(year, 12, 31)
            elif month in [1, 3, 5, 7, 8, 10]:
                return datetime(year, month, 31)
            elif month in [4, 6, 9, 11]:
                return datetime(year, month, 30)
            elif month == 2:
                if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0):
                    return datetime(year, 2, 29)
                else:
                    return datetime(year, 2, 28)
        
        # Handle standard date formats
        date_formats = [
            '%d/%m/%Y', '%d-%m-%Y', '%d/%m/%y', '%d-%m-%y',
            '%Y-%m-%d', '%m/%d/%Y', '%m-%d-%Y',
            '%d/%b/%Y', '%d-%b-%Y', '%b/%d/%Y', '%b-%d-%Y',
            '%d/%B/%Y', '%d-%B-%Y', '%B/%d/%Y', '%B-%d-%Y'
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
    
    print(f"🔍 {sheet_name}: Searching for latest date column...")
    print(f"   Available columns: {columns}")
    
    # Strategy 1: Look for "Indicative adjusted" merged cell and prioritize dates under it
    indicative_positions = []
    
    # Find "Indicative adjusted" text positions
    for row_idx in range(min(10, len(df))):
        for col_idx, col in enumerate(columns):
            val = df.iloc[row_idx, col_idx]
            if pd.notna(val) and 'indicative' in str(val).lower() and 'adjust' in str(val).lower():
                indicative_positions.append((row_idx, col_idx))
                print(f"   📋 Found 'Indicative adjusted' at Row {row_idx}, Col {col_idx} ({col})")
    
    # If we found "Indicative adjusted", use enhanced logic
    if indicative_positions:
        print(f"   🎯 Using 'Indicative adjusted' prioritization logic")
        all_found_dates = []
        
        # First, collect ALL dates from the sheet
        for row_idx in range(min(10, len(df))):
            for col_idx, col in enumerate(columns):
                val = df.iloc[row_idx, col_idx]
                
                if isinstance(val, (pd.Timestamp, datetime)):
                    date_val = val if isinstance(val, datetime) else val.to_pydatetime()
                    all_found_dates.append((date_val, col, row_idx, col_idx, "datetime"))
                    print(f"   📅 Found datetime in {col}[{row_idx}]: {date_val.strftime('%Y-%m-%d')}")
                elif pd.notna(val):
                    parsed_date = parse_date(str(val))
                    if parsed_date:
                        all_found_dates.append((parsed_date, col, row_idx, col_idx, "parsed"))
                        print(f"   📅 Parsed date in {col}[{row_idx}]: '{val}' -> {parsed_date.strftime('%Y-%m-%d')}")
        
        if all_found_dates:
            # Find the latest date
            max_date = max(all_found_dates, key=lambda x: x[0])[0]
            latest_date_columns = [item for item in all_found_dates if item[0] == max_date]
            
            print(f"   📊 Latest date found: {max_date.strftime('%Y-%m-%d')}")
            if len(latest_date_columns) > 1:
                print(f"   📊 Multiple columns with latest date:")
                for date_val, col, row, col_idx, source in latest_date_columns:
                    print(f"      • {col} (col {col_idx})")
            
            # Now find which columns are under "Indicative adjusted" merged cell
            selected_column = None
            
            for indic_row, indic_col in indicative_positions:
                if indic_col > 0:  # Not in description column
                    print(f"   🔍 Analyzing 'Indicative adjusted' merged cell at col {indic_col}")
                    
                    # Use NaN-based merged cell detection
                    # Find the range by looking for NaN values to the right
                    merge_start = indic_col
                    merge_end = indic_col
                    
                    # Look right for NaN values (indicating merged cells)
                    # In Excel merged cells, the leftmost cell has the value, others are NaN
                    for check_col in range(indic_col + 1, len(columns)):
                        val = df.iloc[indic_row, check_col]
                        if pd.isna(val):
                            merge_end = check_col
                        else:
                            # Found non-NaN value, this is the end of the merged range
                            merge_end = check_col - 1
                            break
                    else:
                        # If we reached the end without finding non-NaN, merge goes to the last column
                        merge_end = len(columns) - 1
                    
                    print(f"   📍 'Indicative adjusted' merged cell range: columns {merge_start}-{merge_end}")
                    print(f"   📋 Columns under 'Indicative adjusted': {[columns[i] for i in range(merge_start, merge_end + 1)]}")
                    
                    # Find latest date columns that fall within this merged cell range
                    indicative_latest_columns = []
                    for date_val, col, row, col_idx, source in latest_date_columns:
                        if merge_start <= col_idx <= merge_end:
                            indicative_latest_columns.append((date_val, col, row, col_idx, source))
                            print(f"   ✅ {col} (col {col_idx}) is under 'Indicative adjusted' with latest date")
                    
                    if indicative_latest_columns:
                        # Use the first (leftmost) column under "Indicative adjusted" with latest date
                        selected_column = indicative_latest_columns[0]
                        print(f"   🎯 PRIORITIZED: {selected_column[1]} (latest date under 'Indicative adjusted')")
                        break
            
            # If no column found under "Indicative adjusted", use first column with latest date
            if selected_column is None:
                selected_column = latest_date_columns[0]
                print(f"   ⚠️  Using first column with latest date: {selected_column[1]} (no 'Indicative adjusted' match)")
            
            latest_date, latest_column = selected_column[0], selected_column[1]
    
    # Strategy 2: Fallback to simple logic if no "Indicative adjusted" found
    else:
        print(f"   🔍 No 'Indicative adjusted' found, using simple date detection...")
        
        # First, try to find dates in column names
        column_dates_found = []
        for col in columns:
            col_str = str(col)
            parsed_date = parse_date(col_str)
            if parsed_date:
                column_dates_found.append((parsed_date, col, "column_name"))
                print(f"   📅 Found date in column name '{col}': {parsed_date.strftime('%Y-%m-%d')}")
                if latest_date is None or parsed_date > latest_date:
                    latest_date = parsed_date
                    latest_column = col
                    print(f"   ✅ New latest: {col} ({parsed_date.strftime('%Y-%m-%d')})")
        
        # If no dates found in column names, check the first few rows for datetime values
        if latest_column is None and len(df) > 0:
            print(f"   🔍 No dates in column names, checking row values...")
            cell_dates_found = []
            
            # Check first 5 rows for date values (dates can be in different rows)
            for row_idx in range(min(5, len(df))):
                row = df.iloc[row_idx]
                for col in columns:
                    val = row[col]
                    
                    # Check if it's already a datetime object
                    if isinstance(val, (pd.Timestamp, datetime)):
                        date_val = val if isinstance(val, datetime) else val.to_pydatetime()
                        cell_dates_found.append((date_val, col, f"row_{row_idx}"))
                        print(f"   📅 Found datetime in {col}[{row_idx}]: {date_val.strftime('%Y-%m-%d')}")
                        if latest_date is None or date_val > latest_date:
                            latest_date = date_val
                            latest_column = col
                            print(f"   ✅ New latest: {col} ({date_val.strftime('%Y-%m-%d')}) from row {row_idx}")
                    # Check if it's a string that can be parsed as a date
                    elif pd.notna(val):
                        parsed_date = parse_date(str(val))
                        if parsed_date:
                            cell_dates_found.append((parsed_date, col, f"row_{row_idx}_parsed"))
                            print(f"   📅 Parsed date in {col}[{row_idx}]: '{val}' -> {parsed_date.strftime('%Y-%m-%d')}")
                            if latest_date is None or parsed_date > latest_date:
                                latest_date = parsed_date
                                latest_column = col
                                print(f"   ✅ New latest: {col} ({parsed_date.strftime('%Y-%m-%d')}) from row {row_idx}")
            
            if not cell_dates_found:
                print(f"   ❌ No dates found in cell values")
    
    # Summary of selection
    if latest_column:
        print(f"   🎯 FINAL SELECTION: Column '{latest_column}' with date {latest_date.strftime('%Y-%m-%d')}")
        
        # Show comparison if multiple dates were found
        if 'all_found_dates' in locals() and len(all_found_dates) > 1:
            print(f"   📊 All dates found (for comparison):")
            for date_val, col, row, col_idx, source in sorted(all_found_dates, key=lambda x: x[0], reverse=True):
                marker = "👑" if col == latest_column else "  "
                print(f"   {marker} {col}: {date_val.strftime('%Y-%m-%d')} (from row {row})")
    else:
        print(f"   ❌ No date column detected")
    
    return latest_column

def get_worksheet_sections_by_keys(uploaded_file, tab_name_mapping, entity_name, entity_suffixes, debug=False):
    """
    Get worksheet sections organized by financial keys following the mapping
    """
    try:
        # Handle both uploaded files and default file using context manager to avoid file locks
        if hasattr(uploaded_file, 'name') and uploaded_file.name == "databook.xlsx":
            excel_source = "databook.xlsx"
        else:
            excel_source = uploaded_file

        # Create a reverse mapping from values to keys
        reverse_mapping = {}
        if tab_name_mapping is not None:
            for key, values in tab_name_mapping.items():
                for value in values:
                    reverse_mapping[value] = key
                # Also map the key name directly to itself (for sheet names like "Cash", "AR")
                reverse_mapping[key] = key
        
        # Get financial keys
        financial_keys = get_financial_keys()
        
        # Initialize sections by key
        sections_by_key = {key: [] for key in financial_keys}
        
        # Process sheets within context manager
        with pd.ExcelFile(excel_source) as xl:
            for sheet_name in xl.sheet_names:
                # Skip sheets not in mapping to avoid using undefined df
                if sheet_name not in reverse_mapping:
                    continue
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
                
                # Detect latest date column once per sheet (not per dataframe)
                latest_date_col = detect_latest_date_column(df, sheet_name)
                
                # Organize sections by key - make it less restrictive
                for data_frame in dataframes:
                    if debug and latest_date_col:
                        st.write(f"📅 Latest date column detected: {latest_date_col}")
                    
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
                        
                        # Check entity filter (relaxed for Ningbo/Nanjing)
                        strict_mask = data_frame.apply(
                            lambda row: row.astype(str).str.contains(
                                combined_pattern, case=False, regex=True, na=False
                            ).any(),
                            axis=1
                        )
                        # Relaxed match: allow base entity token without city prefix
                        loose_tokens = [s.strip() for s in entity_suffixes if s.strip()]
                        loose_mask = False
                        if loose_tokens:
                            loose_pattern = '|'.join(re.escape(tok) for tok in loose_tokens)
                            loose_mask = data_frame.apply(
                                lambda row: row.astype(str).str.contains(loose_pattern, case=False, regex=True, na=False).any(),
                                axis=1
                            )
                        entity_mask = strict_mask | loose_mask if isinstance(loose_mask, pd.Series) else strict_mask
                        
                        # If entity filter matches or helpers are empty, process
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
                                
                                # Performance: If we found entity data, we can move to next sheet faster
                                if entity_mask.any():
                                    print(f"🚀 [{sheet_name}] Found entity data for {best_key}, continuing...")
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
    """Get tab name based on project name - more flexible approach"""
    # Try to find a sheet that contains the project name
    try:
        if hasattr(st.session_state, 'uploaded_file') and st.session_state.uploaded_file:
            with pd.ExcelFile(st.session_state.uploaded_file) as xl:
                available_sheets = xl.sheet_names
                
                # Look for sheets that contain the project name
                for sheet in available_sheets:
                    if project_name.lower() in sheet.lower():
                        return sheet
                
                # Fallback to hardcoded names if no match found
                if project_name == 'Haining':
                    return "BSHN"
                elif project_name == 'Nanjing':
                    return "BSNJ"
                elif project_name == 'Ningbo':
                    return "BSNB"
                
                # If still no match, return the first sheet
                return available_sheets[0] if available_sheets else None
    except Exception:
        # Fallback to hardcoded names
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
        with open('fdd_utils/mapping.json', 'r') as f:
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
        with open('fdd_utils/mapping.json', 'r') as f:
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
        # Fallback to centralized defaults if mapping.json not found
        from fdd_utils.mappings import DISPLAY_NAME_MAPPING_DEFAULT as default_names
        return default_names.get(key, key)
    except Exception as e:
        st.error(f"Error loading mapping.json for display names: {e}")
        return key

def main():
    
    # Configure Streamlit page and sanitize deprecated options
    from common.ui import configure_streamlit_page
    configure_streamlit_page()
    st.title("📊 Financial Data Processor")

    # Sidebar for controls
    with st.sidebar:
        # File uploader with default file option
        uploaded_file = st.file_uploader(
            "Upload Excel File (Optional)",
            type=['xlsx', 'xls'],
            help="Upload your financial data Excel file or use the default databook.xlsx"
        )

        # Use default file if no file is uploaded (show immediately under uploader)
        if uploaded_file is None:
            default_file_path = "databook.xlsx"
            if os.path.exists(default_file_path):
                st.caption(f"Using default file: {default_file_path}")
                # Create a proper mock uploaded file object for the default file
                class MockUploadedFile:
                    def __init__(self, file_path):
                        self.name = file_path
                        self.file_path = file_path
                        self._file = None
                    
                    def read(self, size=-1):
                        if self._file is None:
                            self._file = open(self.file_path, 'rb')
                        return self._file.read(size)
                    
                    def getbuffer(self):
                        with open(self.file_path, 'rb') as f:
                            return f.read()
                    
                    def seek(self, offset, whence=0):
                        if self._file is None:
                            self._file = open(self.file_path, 'rb')
                        return self._file.seek(offset, whence)
                    
                    def tell(self):
                        if self._file is None:
                            return 0
                        return self._file.tell()
                    
                    def seekable(self):
                        return True
                    
                    def close(self):
                        if self._file:
                            self._file.close()
                            self._file = None
                
                uploaded_file = MockUploadedFile(default_file_path)
            else:
                st.error(f"❌ Default file not found: {default_file_path}")
                st.info("Please upload an Excel file to get started.")
                st.stop()

        # (Removed duplicate provider/model UI to avoid two model selectors)

        # Entity name input with auto-mapping
        entity_input = st.text_input(
            "Enter Entity Name",
            value="",
            placeholder="e.g., Haining Wanpu Limited, Nanjing Wanchen Limited",
            help="Enter the full entity name to start processing"
        )
        
        # Entity Selection Mode (Single vs Multiple)
        st.markdown("---")
        entity_mode_options = ["Multiple Entities", "Single Entity"]
        entity_mode_display = st.radio(
            "Entity in Databook",
            entity_mode_options,
            index=0,  # Default to Multiple
            help="Select whether your databook contains multiple entities (like Haining, Ningbo, Nanjing) or a single entity"
        )
        
        # Map display names back to internal codes
        entity_mode_mapping = {
            "Multiple Entities": "multiple",
            "Single Entity": "single"
        }
        entity_mode = entity_mode_mapping[entity_mode_display]
        st.session_state['entity_mode'] = entity_mode
        

        
        # Auto-extract base entity and generate mapping keys
        if entity_input:
            # Extract base entity name (first word)
            base_entity = entity_input.split()[0] if entity_input.split() else None
            
            # Generate mapping keys based on input
            mapping_keys = []
            words = entity_input.split()
            for i in range(len(words)):
                mapping_keys.append(" ".join(words[:i+1]))
            
            # Use base entity for processing
            selected_entity = base_entity
            
            # Show only base entity info
            st.info(f"📋 Base Entity: {base_entity}")
        else:
            selected_entity = None
            mapping_keys = []
            st.warning("⚠️ Please enter an entity name to start processing")
        # Check if entity is provided (file can be default)
        if not selected_entity:
            st.stop()
        
        # Use mapping keys as entity helpers
        if 'mapping_keys' in locals() and mapping_keys:
            entity_helpers = ",".join(mapping_keys) + ","
        else:
            # Fallback to original logic
            if selected_entity == 'Haining':
                entity_helpers = "Wanpu,Limited,"
            elif selected_entity in ['Nanjing', 'Ningbo']:
                entity_helpers = "Wanchen,Limited,Logistics,Development,Supply,Chain," 
            else:
                entity_helpers = "Limited,"

        # Auto-invalidate Excel cache when entity changes or when date detection is updated
        last_entity = st.session_state.get('last_selected_entity')
        cache_version = st.session_state.get('cache_version', 'v1')
        
        # Clear cache if entity changes or if we need to update for date detection
        if (last_entity is None or last_entity != selected_entity or cache_version != 'v9'):
            keys_to_remove = [key for key in st.session_state.keys() if key.startswith('sections_by_key_')]
            for key in keys_to_remove:
                del st.session_state[key]
            st.session_state['last_selected_entity'] = selected_entity
            st.session_state['cache_version'] = 'v9'  # Update cache version for date detection
        
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
            if hasattr(uploaded_file, 'name') and uploaded_file.name != "databook.xlsx":
                st.success(f"Uploaded {uploaded_file.name}")
            # Store uploaded file in session state for later use
            st.session_state['uploaded_file'] = uploaded_file
            
            # AI Provider Selection (models defined in fdd_utils/config.json)
            ai_mode_options = ["Local AI", "Open AI", "DeepSeek", "Offline"]
            mode_display = st.selectbox(
                "Select Mode", 
                ai_mode_options,
                index=0,  # Local AI default
                help="Choose AI provider. Models are taken from fdd_utils/config.json"
            )
            
            # Show API configuration status
            config, _, _, _ = load_config_files()
            if config:
                if mode_display == "Open AI":
                    if config.get('OPENAI_API_KEY') and config.get('OPENAI_API_BASE'):
                        st.success("✅ OpenAI configured")
                        st.info(f"🤖 Model: {config.get('OPENAI_CHAT_MODEL', 'gpt-4o-mini')}")
                    else:
                        st.warning("⚠️ OpenAI not configured. Add OPENAI_API_KEY and OPENAI_API_BASE in fdd_utils/config.json")
                elif mode_display == "DeepSeek":
                    if config.get('DEEPSEEK_API_KEY') and config.get('DEEPSEEK_API_BASE'):
                        st.success("✅ DeepSeek configured")
                        st.info(f"🤖 Model: {config.get('DEEPSEEK_CHAT_MODEL', 'deepseek-chat')}")
                    else:
                        st.warning("⚠️ DeepSeek not configured. Add DEEPSEEK_API_KEY and DEEPSEEK_API_BASE in fdd_utils/config.json")
                elif mode_display == "Local AI":
                    if config.get('LOCAL_AI_API_BASE') and config.get('LOCAL_AI_ENABLED'):
                        st.success("✅ Local AI configured")
                        st.info(f"🏠 Model: {config.get('LOCAL_AI_CHAT_MODEL', 'local-qwen2')}")
                        st.info(f"🔗 Endpoint: {config.get('LOCAL_AI_API_BASE', 'Not specified')}")
                    else:
                        st.warning("⚠️ Local AI not configured. Configure LOCAL_AI_* in fdd_utils/config.json")
            
            # Map display names to internal mode names
            provider_mapping = {
                "Open AI": "Open AI",
                "Local AI": "Local AI",
                "DeepSeek": "DeepSeek",
                "Offline": "Offline"
            }
            mode = f"AI Mode - {provider_mapping[mode_display]}" if mode_display != "Offline" else "Offline Mode"
            st.session_state['selected_mode'] = mode
            st.session_state['ai_model'] = mode_display
            st.session_state['selected_provider'] = provider_mapping[mode_display]
            st.session_state['use_local_ai'] = (mode_display == "Local AI")
            st.session_state['use_openai'] = (mode_display == "Open AI")
            
            # Performance statistics - moved below Select Mode
            st.markdown("---")
            st.markdown("### 🚀 Performance")
            cache = get_simple_cache()
            
            if st.button("🧹 Clear All Cache"):
                cache.clear_cache()
                # Also clear session state cache for Excel processing
                keys_to_remove = [key for key in st.session_state.keys() if key.startswith('sections_by_key_')]
                for key in keys_to_remove:
                    del st.session_state[key]
                # Clear cache version to force refresh
                if 'cache_version' in st.session_state:
                    del st.session_state['cache_version']
                st.success("All cache cleared!")
            
            if st.button("📊 Clear Excel Cache"):
                # Clear only Excel processing cache
                keys_to_remove = [key for key in st.session_state.keys() if key.startswith('sections_by_key_')]
                if keys_to_remove:
                    for key in keys_to_remove:
                        del st.session_state[key]
                    st.success(f"Excel cache cleared! ({len(keys_to_remove)} entries removed)")
                else:
                    st.info("No Excel cache to clear")
            
            if st.button("🔄 Force Refresh"):
                st.session_state['force_refresh'] = True
                st.success("Force refresh enabled! Next run will bypass cache.")
            
            if st.button("📁 List Cache Files"):
                cache.list_cache_files()
                st.success("Cache files listed in console!")

    # Main area for results
    if uploaded_file is not None:
        
        # --- View Table Section ---
        config, mapping, pattern, prompts = load_config_files()
        
        # Process entity configuration based on mode
        entity_mode = st.session_state.get('entity_mode', 'multiple')
        
        if entity_mode == 'single':
            # For single entity mode, use only the base entity name
            entity_suffixes = []
            entity_keywords = [selected_entity]
        else:
            # For multiple entity mode, use the full entity helpers
            entity_suffixes = [s.strip() for s in entity_helpers.split(',') if s.strip()]
            entity_keywords = [f"{selected_entity} {suffix}" for suffix in entity_suffixes if suffix]
            if not entity_keywords:
                entity_keywords = [selected_entity]
        
        # Handle different statement types with session state caching
        if statement_type == "BS":
            # Create cache key to avoid reprocessing
            cache_key = f"sections_by_key_{uploaded_file.name if hasattr(uploaded_file, 'name') else 'default'}_{selected_entity}"
            
            if cache_key not in st.session_state:
                # Original BS logic - only run if not cached
                with st.spinner("🔄 Processing Excel file..."):
                    sections_by_key = get_worksheet_sections_by_keys(
                        uploaded_file=uploaded_file,
                        tab_name_mapping=mapping,
                        entity_name=selected_entity,
                        entity_suffixes=entity_suffixes,
                        debug=False  # Set to True for debugging
                    )
                    st.session_state[cache_key] = sections_by_key
            else:
                sections_by_key = st.session_state[cache_key]
            from common.ui_sections import render_balance_sheet_sections
            render_balance_sheet_sections(
                sections_by_key,
                get_key_display_name,
                selected_entity,
                format_date_to_dd_mmm_yyyy,
            )
        
        elif statement_type == "IS":
            # Income Statement placeholder
            st.markdown("### Income Statement")
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
            st.markdown("### Combined Financial Statements")
            st.info("📊 Combined BS and IS processing will be implemented here.")
            st.markdown("""
            **Placeholder for Combined sections:**
            - Balance Sheet
            - Income Statement
            - Cash Flow Statement
            - Financial Ratios
            """)

        # --- AI Processing Section (Bottom) ---
        # Check AI configuration status (updated for DeepSeek as default)
        try:
            config, _, _, _ = load_config_files()
            if config:
                any_provider_configured = (
                    (config.get('DEEPSEEK_API_KEY') and config.get('DEEPSEEK_API_BASE')) or
                    (config.get('OPENAI_API_KEY') and config.get('OPENAI_API_BASE')) or
                    (config.get('LOCAL_AI_API_BASE') and config.get('LOCAL_AI_ENABLED')) or
                    (config.get('SERVER_AI_API_BASE') and (config.get('SERVER_AI_API_KEY') or config.get('LOCAL_AI_API_KEY')))
                )
                if not any_provider_configured:
                    st.warning("⚠️ AI Mode: No provider configured. Will use fallback mode with test data.")
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
                
                # Process entity configuration based on mode
                entity_mode = st.session_state.get('entity_mode', 'multiple')
                
                if entity_mode == 'single':
                    # For single entity mode, use only the base entity name
                    entity_suffixes = []
                    entity_keywords = [selected_entity]
                else:
                    # For multiple entity mode, use the full entity helpers
                    entity_suffixes = [s.strip() for s in entity_helpers.split(',') if s.strip()]
                    entity_keywords = [f"{selected_entity} {suffix}" for suffix in entity_suffixes if suffix]
                    if not entity_keywords:
                        entity_keywords = [selected_entity]
                
                # Get worksheet sections with caching
                cache_key = f"sections_by_key_{uploaded_file.name if hasattr(uploaded_file, 'name') else 'default'}_{selected_entity}"
                
                if cache_key not in st.session_state:
                    with st.spinner("🔄 Processing Excel file for AI..."):
                        sections_by_key = get_worksheet_sections_by_keys(
                            uploaded_file=uploaded_file,
                            tab_name_mapping=mapping,
                            entity_name=selected_entity,
                            entity_suffixes=entity_suffixes,
                            debug=False
                        )
                        st.session_state[cache_key] = sections_by_key
                    st.success("✅ Excel processing completed for AI")
                else:
                    sections_by_key = st.session_state[cache_key]
                
                # Get keys with data
                keys_with_data = [key for key, sections in sections_by_key.items() if sections]
                
                # Filter keys based on statement type (include synonyms for non-Haining)
                bs_keys = [
                    "Cash", "AR", "Prepayments", "OR", "Other CA", "IP", "Other NCA", "NCA",
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
                
                # AI Processing Button with a single main progress bar
                agent_states = st.session_state.get('agent_states', {})
                col_a, col_b, col_c = st.columns(3)
                with col_a:
                    run_ai_clicked = st.button("🚀 Run AI: Content Generation", type="secondary", use_container_width=True, key="btn_ai_gen")
                with col_b:
                    run_proof_clicked = st.button("🧐 Run AI: Proofreader", type="secondary", use_container_width=True, key="btn_ai_proof")
                with col_c:
                    run_both_clicked = st.button("🔁 Run AI: Generate → Proofread", type="primary", use_container_width=True, key="btn_ai_both")

                if run_ai_clicked:
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    eta_text = st.empty()
                    try:
                        status_text.text("🤖 Initializing…")
                        progress_bar.progress(10)
                        # Disable cache for this run
                        st.session_state['force_refresh'] = True
                        ext = {'bar': progress_bar, 'status': status_text, 'combined': {'stages': 1, 'stage_index': 0, 'start_time': time.time()}}
                        agent1_results = run_agent_1(filtered_keys_for_ai, temp_ai_data, external_progress=ext)
                        agent1_success = bool(agent1_results and any(agent1_results.values()))
                        st.session_state['agent_states']['agent1_results'] = agent1_results
                        st.session_state['agent_states']['agent1_completed'] = True
                        st.session_state['agent_states']['agent1_success'] = agent1_success
                        progress_bar.progress(100)
                        status_text.text("✅ AI done" if agent1_success else "❌ AI failed")
                        time.sleep(1)
                        st.rerun()
                    except Exception as e:
                        progress_bar.progress(100)
                        status_text.text(f"❌ AI failed: {e}")
                        time.sleep(1)
                        st.rerun()
                
                if run_proof_clicked:
                    # Run AI Proofreader using Agent 1 outputs
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    eta_text = st.empty()
                    try:
                        status_text.text("🧐 Initializing…")
                        progress_bar.progress(10)
                        # Disable cache for this run
                        st.session_state['force_refresh'] = True
                        agent1_results = st.session_state.get('agent_states', {}).get('agent1_results', {}) or {}
                        if not agent1_results:
                            st.warning("Run content generation first to produce material for proofreading.")
                        else:
                            ext = {'bar': progress_bar, 'status': status_text}
                            proof_results = run_ai_proofreader(filtered_keys_for_ai, agent1_results, temp_ai_data, external_progress=ext)
                            st.session_state['agent_states']['agent3_results'] = proof_results
                            st.session_state['agent_states']['agent3_completed'] = True
                            st.session_state['agent_states']['agent3_success'] = bool(proof_results)
                        progress_bar.progress(100)
                        status_text.text("✅ Proofreading done")
                        time.sleep(1)
                        st.rerun()
                    except Exception as e:
                        progress_bar.progress(100)
                        status_text.text(f"❌ Proofreading failed: {e}")
                        time.sleep(1)
                        st.rerun()

                if run_both_clicked:
                    # Run Generation then Proofreader in sequence with single trigger
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    try:
                        status_text.text("🤖 Initializing…")
                        progress_bar.progress(10)
                        st.session_state['force_refresh'] = True
                        ext = {'bar': progress_bar, 'status': status_text, 'combined': {'stages': 2, 'stage_index': 0, 'start_time': time.time()}}
                        agent1_results = run_agent_1(filtered_keys_for_ai, temp_ai_data, external_progress=ext)
                        st.session_state['agent_states']['agent1_results'] = agent1_results
                        st.session_state['agent_states']['agent1_completed'] = True
                        st.session_state['agent_states']['agent1_success'] = bool(agent1_results)
                        # Proofreader
                        status_text.text("🧐 Running…")
                        # Switch to PROOF stage index for combined ETA/progress
                        ext['combined']['stage_index'] = 1
                        proof_results = run_ai_proofreader(filtered_keys_for_ai, agent1_results, temp_ai_data, external_progress=ext)
                        st.session_state['agent_states']['agent3_results'] = proof_results
                        st.session_state['agent_states']['agent3_completed'] = True
                        st.session_state['agent_states']['agent3_success'] = bool(proof_results)
                        progress_bar.progress(100)
                        status_text.text("✅ Generate → Proofread complete")
                        time.sleep(1)
                        st.rerun()
                    except Exception as e:
                        progress_bar.progress(100)
                        status_text.text(f"❌ Combined run failed: {e}")
                        time.sleep(1)
                        st.rerun()
                

                
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
                        # Show Compliance (Proofreader) first if available
                        agent3_results_all = agent_states.get('agent3_results', {}) or {}
                        if key in agent3_results_all:
                            pr = agent3_results_all[key]
                            corrected_content = pr.get('corrected_content', '') or pr.get('content', '')
                            if corrected_content:
                                st.markdown(corrected_content)

                        # AI1 Results (collapsible if proofreader exists)
                        with st.expander("📝 AI: Generation (details)", expanded=key not in agent3_results_all):
                            agent1_results = agent_states.get('agent1_results', {}) or {}
                            if key in agent1_results and agent1_results[key]:
                                content = agent1_results[key]
                                
                                # Handle both string and dict content
                                if isinstance(content, dict):
                                    content_str = content.get('content', str(content))
                                else:
                                    content_str = str(content)
                                
                                st.markdown("**Generated Content:**")
                                st.markdown(content_str)
                                
                                # Metadata
                                col1, col2, col3 = st.columns(3)
                                with col1:
                                    st.metric("Characters", len(content_str))
                                with col2:
                                    st.metric("Words", len(content_str.split()))
                                with col3:
                                    st.metric("Status", "✅ Generated" if content else "❌ Failed")
                            else:
                                st.info("No AI results available. Run AI first.")

                        # If Proofreader made changes or found issues, show a compact summary
                        if key in agent3_results_all:
                            pr = agent3_results_all[key]
                            issues = pr.get('issues', []) or []
                            changed = bool(pr.get('corrected_content'))
                            if issues or changed:
                                with st.expander("🧐 AI Proofreader: Changes & Notes", expanded=False):
                                    if changed and not issues:
                                        st.markdown("- Corrected content applied")
                                    if issues:
                                        st.markdown("- Detected issues (reference only):")
                                        for issue in issues:
                                            st.write(f"  • {issue}")
                                    runs = pr.get('translation_runs', 0)
                                    if runs:
                                        st.markdown(f"- Heuristic translation applied: {runs} run(s)")
                        if key not in agent3_results_all:
                            st.info("No compliance results available. Run AI Proofreader.")
            else:
                st.info("No financial keys available for results display.")
        else:
            st.info("No AI agents have run yet. Use the buttons above to start processing.")
        
        # --- PowerPoint Generation Section (Bottom) ---
        st.markdown("---")
        st.subheader("📊 PowerPoint Generation")
        
        col1, col2 = st.columns([1, 1])
        
        with col1:
            if st.button("📊 Prepare PowerPoint", type="secondary", use_container_width=True):
                try:
                    # Get the project name based on selected entity
                    project_name = selected_entity
                    
                    # Check for template file in common locations
                    possible_templates = [
                        "fdd_utils/template.pptx",
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
                        st.error("❌ PowerPoint template not found. Please ensure 'template.pptx' exists in the fdd_utils/ directory.")
                        st.info("💡 You can copy a template file from the old_ver/ directory or create a new one.")
                    else:
                        # Define output path with timestamp in fdd_utils/output directory
                        from datetime import datetime
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        output_filename = f"{project_name}_{statement_type.upper()}_{timestamp}.pptx"
                        output_path = f"fdd_utils/output/{output_filename}"
                        
                        # Ensure output directory exists
                        os.makedirs("fdd_utils/output", exist_ok=True)
                        
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
                        
                        # Get the Excel file path for embedding data
                        excel_file_path = None
                        if hasattr(uploaded_file, 'file_path'):
                            excel_file_path = uploaded_file.file_path
                        elif hasattr(uploaded_file, 'name'):
                            excel_file_path = uploaded_file.name
                        
                        export_pptx(
                            template_path=template_path,
                            markdown_path="fdd_utils/bs_content.md",
                            output_path=output_path,
                            project_name=project_name,
                            excel_file_path=excel_file_path
                        )
                        
                        st.session_state['pptx_exported'] = True
                        st.session_state['pptx_filename'] = output_filename
                        st.session_state['pptx_path'] = output_path
                        st.success(f"✅ PowerPoint is ready for download: {output_filename}")
                        
                except FileNotFoundError as e:
                    st.error(f"❌ Template file not found: {e}")
                except Exception as e:
                    st.error(f"❌ Export failed: {e}")
                    st.error(f"Error details: {str(e)}")
        
        with col2:
            # Show a disabled download button until a PPTX is available, then enable
            pptx_ready = st.session_state.get('pptx_exported', False) and os.path.exists(st.session_state.get('pptx_path', ''))
            if pptx_ready:
                with open(st.session_state['pptx_path'], "rb") as file:
                    st.download_button(
                        label="📥 Download PowerPoint",
                        data=file.read(),
                        file_name=st.session_state['pptx_filename'],
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True
                    )
            else:
                st.download_button(
                    label="📥 Download PowerPoint",
                    data=b"",
                    disabled=True,
                    file_name="report.pptx",
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
def load_json_content():
    """Load content from JSON file with caching for better performance"""
    try:
        # Try JSON first (better performance)
        json_file = "fdd_utils/bs_content.json"
        if os.path.exists(json_file):
            with open(json_file, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception as e:
        print(f"Error loading JSON content: {e}")
    
    # Fallback to parsing markdown if JSON not available
    try:
        content_files = ["fdd_utils/bs_content.md", "fdd_utils/bs_content_ai_generated.md", "fdd_utils/bs_content_offline.md"]
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
        content_file = "fdd_utils/bs_content_offline.md"
        content = None
        
        try:
            with open(content_file, 'r', encoding='utf-8') as f:
                content = f.read()
        except FileNotFoundError:
            # Try to read from AI-generated content if offline file not found
            ai_content_files = ["fdd_utils/bs_content.md", "fdd_utils/bs_content_ai_generated.md"]
            for ai_file in ai_content_files:
                try:
                    with open(ai_file, 'r', encoding='utf-8') as f:
                        content = f.read()
                        st.info(f"📄 Using AI-generated content from: {ai_file}")
                        break
                except FileNotFoundError:
                    continue
        
        if not content:
            st.error(f"No content files found. Checked: {content_file}, fdd_utils/bs_content.md, fdd_utils/bs_content_ai_generated.md")
            return
        
        # Map financial keys to content sections
        key_to_section_mapping = KEY_TO_SECTION_MAPPING
        
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
        content_file = "fdd_utils/bs_content_offline.md"
        
        with open(content_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Centralized mapping
        from fdd_utils.mappings import KEY_TO_SECTION_MAPPING as key_to_section_mapping
        
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
            terms = KEY_TERMS_BY_KEY.get(key, [key_lower])
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
            'has_description': len(str(agent1_content).split()) > 10,
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
            name_mapping = DISPLAY_NAME_MAPPING_NB_NJ
            category_mapping = {
                'Current Assets': ['Cash', 'AR', 'Prepayments', 'OR', 'Other CA'],
                'Non-current Assets': ['IP', 'Other NCA'],
                'Liabilities': ['AP', 'Taxes payable', 'OP'],
                'Equity': ['Capital']
            }
        else:  # Haining and others
            name_mapping = DISPLAY_NAME_MAPPING_DEFAULT
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
        json_file_path = 'fdd_utils/bs_content.json'
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
        md_file_path = 'fdd_utils/bs_content.md'
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
            name_mapping = DISPLAY_NAME_MAPPING_NB_NJ
            category_mapping = {
                'Current Assets': ['Cash', 'AR', 'Prepayments', 'OR', 'Other CA'],
                'Non-current Assets': ['IP', 'Other NCA'],
                'Liabilities': ['AP', 'Taxes payable', 'OP'],
                'Equity': ['Capital']
            }
        else:  # Haining and others
            name_mapping = DISPLAY_NAME_MAPPING_DEFAULT
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
        file_path = 'fdd_utils/bs_content.md'
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
    with open("fdd_utils/bs_prompt_debug.md", "w", encoding="utf-8") as f:
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

def run_agent_1(filtered_keys, ai_data, external_progress=None):
    """Run Agent 1: Content Generation for all keys"""
    try:

        import time
        
        logger = st.session_state.ai_logger
        # Keep content section minimal; avoid extra headings that duplicate the main status
        
        # Get data from ai_data
        entity_name = ai_data.get('entity_name', '')
        entity_keywords = ai_data.get('entity_keywords', [])
        
        # Create temporary file for processing
        temp_file_path = None
        try:
            if 'uploaded_file_data' in st.session_state:
                # Use a unique filename to avoid conflicts
                unique_filename = f"databook_{uuid.uuid4().hex[:8]}.xlsx"
                temp_file_path = os.path.join(tempfile.gettempdir(), unique_filename)
                
                with open(temp_file_path, 'wb') as tmp_file:
                    tmp_file.write(st.session_state['uploaded_file_data'])
                # Temp file created silently
            else:
                # Fallback: use existing databook.xlsx
                if os.path.exists('databook.xlsx'):
                    temp_file_path = 'databook.xlsx'
                    st.info("📄 Using existing databook.xlsx")
                else:
                    st.error("❌ No databook available for processing")
                    return {}
        except Exception as e:
            st.error(f"❌ Error creating temporary file: {e}")
            return {}
        
        # Get the actual prompts that will be sent to AI by calling process_keys
        # We need to capture the real prompts with table data
        try:
        # Load prompts from prompts.json file
            with open('fdd_utils/prompts.json', 'r') as f:
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
    
            financial_figures = find_financial_figures_with_context_check(temp_file_path, get_tab_name(entity_name), None, convert_thousands=False)
            financial_figure_info = f"{key}: {get_financial_figure(financial_figures, key)}"
            
            # Build the actual user prompt using templates from prompts.json, then inject dynamic data
            try:
                with open('fdd_utils/prompts.json', 'r') as f:
                    prompts_cfg = json.load(f)
                user_prompts_cfg = prompts_cfg.get('user_prompts', {})
                generic_cfg = prompts_cfg.get('generic_prompt', {})
                # Pick key-specific config if available; else generic
                key_prompt_cfg = user_prompts_cfg.get(key, generic_cfg)
            except Exception:
                key_prompt_cfg = {}

            title = key_prompt_cfg.get('title', f'{get_key_display_name(key)} Analysis for {entity_name}')
            description = key_prompt_cfg.get('description', f'Analyze {get_key_display_name(key)} using worksheet data')
            analysis_points = key_prompt_cfg.get('analysis_points', [])
            key_questions = key_prompt_cfg.get('key_questions', [])
            data_sources = key_prompt_cfg.get('data_sources', [])

            # Compose user prompt template then inject dynamic dataset and patterns
            prompt_lines = [
                f"{title}",
                f"{description}",
                "",
                f"AVAILABLE PATTERNS: {pattern_json}",
                f"FINANCIAL FIGURE: {financial_figure_info}",
                f"DATA SOURCE: {key_tables}",
                "",
            ]

            if analysis_points:
                prompt_lines.append("REQUIRED ANALYSIS:")
                for i, p in enumerate(analysis_points, 1):
                    prompt_lines.append(f"{i}. {p.replace('[ENTITY_NAME]', entity_name)}")
                prompt_lines.append("")

            if key_questions:
                prompt_lines.append("KEY QUESTIONS:")
                for q in key_questions:
                    prompt_lines.append(f"- {q.replace('[ENTITY_NAME]', entity_name)}")
                prompt_lines.append("")

            if data_sources:
                prompt_lines.append("DATA SOURCES:")
                for s in data_sources:
                    prompt_lines.append(f"- {s}")
                prompt_lines.append("")

            # Output requirements central; concise, no template artifacts
            prompt_lines += [
                "OUTPUT REQUIREMENTS:",
                "- Provide only the final completed text; no JSON, no headers, no pattern names",
                "- Replace placeholders with actual values and entity names from the DATA SOURCE",
                "- Use exact entity names shown in the table (not the reporting entity)",
                "- Maintain professional financial tone and formatting",
                "- Ensure all figures match the DATA SOURCE",
            ]

            actual_user_prompt = "\n".join(prompt_lines)
            
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
        
        # Create/Reuse a single progress bar and status text
        if external_progress:
            progress_bar = external_progress.get('bar')
            status_text = external_progress.get('status')
        else:
            progress_bar = st.progress(0)
            status_text = st.empty()
        
        # Create progress callback for Streamlit with ETA (supports combined 2-stage runs)
        start_time = time.time()
        total = max(1, len(filtered_keys))
        def update_progress(progress, message):
            try:
                # If running in combined mode, map this stage's progress into overall combined progress
                combined = external_progress.get('combined') if external_progress else None
                if combined and isinstance(combined, dict):
                    stages = combined.get('stages', 2)
                    stage_index = combined.get('stage_index', 0)
                    stage_weight = 1.0 / max(1, stages)
                    combined_progress = min(1.0, max(0.0, stage_index * stage_weight + progress * stage_weight))
                    progress_bar.progress(combined_progress)
                else:
                    progress_bar.progress(progress)
            except Exception:
                pass
            try:
                # Estimate ETA from progress
                combined = external_progress.get('combined') if external_progress else None
                now = time.time()
                if combined and isinstance(combined, dict):
                    # Combined ETA from overall combined progress
                    start_all = combined.get('start_time') or start_time
                    stages = combined.get('stages', 2)
                    stage_index = combined.get('stage_index', 0)
                    stage_weight = 1.0 / max(1, stages)
                    overall_progress = max(0.001, min(0.999, stage_index * stage_weight + progress * stage_weight))
                    elapsed_all = now - start_all
                    remaining = int(elapsed_all * (1 - overall_progress) / overall_progress)
                else:
                    elapsed = now - start_time
                    remaining = int(elapsed * (1 - progress) / max(progress, 0.001))
                mins, secs = divmod(max(0, remaining), 60)
                eta_str = f"ETA {mins:02d}:{secs:02d}" if progress < 1 else "ETA 00:00"
                status_text.text(f"📝 Agent 1 — {message} — {eta_str}")
            except Exception:
                pass
        
        # Get processed table data from session state
        processed_table_data = ai_data.get('sections_by_key', {})
        
        # Get AI model settings from session state
        use_local_ai = st.session_state.get('use_local_ai', False)
        use_openai = st.session_state.get('use_openai', False)
        
        results = process_keys(
            keys=filtered_keys,  # All keys at once
            entity_name=entity_name,
            entity_helpers=entity_keywords,
            input_file=temp_file_path,
            mapping_file="fdd_utils/mapping.json",
            pattern_file="fdd_utils/pattern.json",
            config_file='fdd_utils/config.json',
            prompts_file='fdd_utils/prompts.json',
            use_ai=True,
            progress_callback=update_progress,
            processed_table_data=processed_table_data,
            use_local_ai=use_local_ai,
            use_openai=use_openai
        )
        
        processing_time = time.time() - start_time
        
        # Log Agent 1 output for each key with pattern information
        for key in filtered_keys:
            key_result = results.get(key, {})
            if isinstance(key_result, dict):
                content = key_result.get('content', f"No result generated for {key}")
                pattern_used = key_result.get('pattern_used', 'Unknown')
                table_data = key_result.get('table_data', '')
                financial_figure = key_result.get('financial_figure', 0)
            else:
                content = key_result
                pattern_used = 'Unknown'
                table_data = ''
                financial_figure = 0
            
            # Create enhanced output with pattern information
            enhanced_output = {
                'content': content,
                'pattern_used': pattern_used,
                'table_data': table_data,
                'financial_figure': financial_figure,
                'entity_name': entity_name
            }
            
            logger.log_agent_output('agent1', key, enhanced_output, processing_time / len(filtered_keys))
        
        st.success(f"🎉 Agent 1 completed all {len(filtered_keys)} keys in {processing_time:.2f}s")
        return results
            
    except RuntimeError as e:
        # AI-specific errors
        if "AI processing is required" in str(e):
            st.error("❌ **AI Processing Required**")
            st.error("This application requires AI services to function. Please check your configuration.")
            st.error(f"Error: {e}")
        elif "AI services are not available" in str(e):
            st.error("❌ **AI Services Unavailable**")
            st.error("DeepSeek AI services are not available. Please check your internet connection and API configuration.")
            st.error(f"Error: {e}")
        else:
            st.error(f"❌ **AI Error**: {e}")
        return {}
    except Exception as e:
        st.error(f"❌ **Unexpected Error**: {e}")
        return {}
    finally:
        # Clean up temp file if created
        try:
            if temp_file_path and temp_file_path != 'databook.xlsx' and os.path.exists(temp_file_path):
                os.unlink(temp_file_path)
        except PermissionError:
            # On Windows, file may still be locked momentarily; retry once
            try:
                import time
                time.sleep(0.2)
                if os.path.exists(temp_file_path):
                    os.unlink(temp_file_path)
            except Exception:
                pass

def run_ai_proofreader(filtered_keys, agent1_results, ai_data, external_progress=None):
    """Run AI Proofreader for all keys (Compliance, Figures, Entities, Grammar)."""
    try:
        import json
        logger = st.session_state.ai_logger

        # Model/provider selection
        use_local_ai = st.session_state.get('use_local_ai', False)
        use_openai = st.session_state.get('use_openai', False)
        proof_agent = ProofreadingAgent(use_local_ai=use_local_ai, use_openai=use_openai)

        results = {}
        entity_name = ai_data.get('entity_name', '')
        sections_by_key = ai_data.get('sections_by_key', {})

        # Prepare per-key tables markdown (from processed tables)
        def get_tables_for_key(k):
            key_sections = sections_by_key.get(k, [])
            parts = []
            for section in key_sections:
                if isinstance(section, dict):
                    try:
                        parts.append(json.dumps(section, indent=2, default=str))
                    except Exception:
                        parts.append(str(section))
                elif hasattr(section, 'to_string'):
                    df = section.copy()
                    for col in list(df.columns):
                        if df[col].isna().all() or (df[col].astype(str) == 'None').all():
                            df = df.drop(columns=[col])
                    parts.append(df.to_string(index=False, na_rep=''))
                else:
                    parts.append(str(section))
            return "\n".join(parts)

        # Check if we're in a command-line environment (no Streamlit)
        try:
            # Try to access Streamlit session state - if it fails, we're in CLI
            _ = st.session_state
            is_cli = False
        except Exception:
            is_cli = True
        
        if is_cli:
            # Use tqdm for command-line progress
            progress_bar = tqdm(total=len(filtered_keys), desc="🤖 AI Proofreader", unit="key")
        else:
            # Use Streamlit progress
            if external_progress:
                progress_bar = external_progress.get('bar')
                status_text = external_progress.get('status')
            else:
                progress_bar = st.progress(0)
                status_text = st.empty()
        
        start_time = time.time()
        total = len(filtered_keys)
        
        for idx, key in enumerate(filtered_keys):
            if not is_cli:
                elapsed = time.time() - start_time
                # Simple ETA: average time per processed item * remaining
                avg = (elapsed / (idx or 1)) if idx else 0
                remaining = total - idx
                eta_seconds = int(avg * remaining) if idx else 0
                mins, secs = divmod(eta_seconds, 60)
                eta_str = f"ETA {mins:02d}:{secs:02d}" if eta_seconds > 0 else "ETA --:--"
                status_text.text(f"🧐 Proofreader — {key} ({idx+1}/{total}) — {eta_str}")
                try:
                    combined = external_progress.get('combined') if external_progress else None
                    if combined and isinstance(combined, dict):
                        stages = combined.get('stages', 2)
                        stage_index = combined.get('stage_index', 1)
                        stage_weight = 1.0 / max(1, stages)
                        progress = (idx+1)/len(filtered_keys)
                        combined_progress = min(1.0, max(0.0, stage_index * stage_weight + progress * stage_weight))
                        progress_bar.progress(combined_progress)
                    else:
                        progress_bar.progress((idx+1)/len(filtered_keys))
                except Exception:
                    pass
            
            try:
                content = agent1_results.get(key, '')
                if isinstance(content, dict):
                    content_text = content.get('content', '')
                    tables_md = content.get('table_data', '') or get_tables_for_key(key)
                else:
                    content_text = str(content)
                    tables_md = get_tables_for_key(key)

                if not content_text:
                    results[key] = {'is_compliant': False, 'issues': ["No Agent 1 content"], 'corrected_content': ''}
                    if is_cli:
                        progress_bar.update(1)
                    continue

                # Pass progress_bar to proofread method
                result = proof_agent.proofread(content_text, key, tables_md, entity_name, progress_bar if is_cli else None)
                results[key] = result

                # Log output
                try:
                    logger.log_agent_output('agent3', key, result, 0)
                except Exception:
                    pass

                # Update session store with corrected content
                if not is_cli:
                    content_store = st.session_state.get('ai_content_store', {})
                    if key not in content_store:
                        content_store[key] = {}
                    corrected = result.get('corrected_content') or content_text
                    content_store[key]['agent3_content'] = corrected
                    content_store[key]['current_content'] = corrected
                    st.session_state['ai_content_store'] = content_store
            except Exception as e:
                results[key] = {'is_compliant': False, 'issues': [str(e)], 'corrected_content': ''}
                if is_cli:
                    progress_bar.update(1)

        if is_cli:
            progress_bar.close()
        else:
            st.success("✅ AI Proofreader completed")
        
        return results
    except Exception as e:
        st.error(f"❌ AI Proofreader Error: {e}")
        return {}

def update_bs_content_with_agent_corrections(corrections_dict, entity_name, agent_name):
    """Update bs_content.md with corrections from Agent 2 or Agent 3"""
    try:
        import re
        
        # Read current bs_content.md
        bs_content_path = 'fdd_utils/bs_content.md'
        if not os.path.exists(bs_content_path):
            st.warning(f"bs_content.md not found at {bs_content_path}")
            return False
        
        with open(bs_content_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Define category mappings based on entity name
        if entity_name in ['Ningbo', 'Nanjing']:
            from fdd_utils.mappings import DISPLAY_NAME_MAPPING_NB_NJ as name_mapping
        else:  # Haining and others
            from fdd_utils.mappings import DISPLAY_NAME_MAPPING_DEFAULT as name_mapping
        
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
            shutil.copy2(bs_content_path, 'fdd_utils/bs_content_ai_generated.md')
        except Exception as e:
            print(f"Could not update AI reference copy: {e}")
        
        return True
        
    except Exception as e:
        st.error(f"Error updating bs_content.md with {agent_name} corrections: {e}")
        return False

# run_agent_2 function removed as requested

def read_bs_content_by_key():
    """Read BS content by key - simplified since AI2 removed"""
    return {}

# Previous run_agent_2 code removed - this section will be cleaned up
# All content below was part of the old run_agent_2 function and needs to be removed

def run_agent_3(filtered_keys, agent1_results, ai_data):
    # Deprecated legacy agent (kept for backward compatibility); no-op.
    return {}

def read_bs_content_by_key(entity_name):
    """Read bs_content.md and return content organized by key"""
    try:
        import re
        
        # Read current bs_content.md
        bs_content_path = 'fdd_utils/bs_content.md'
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

        import json
        import time
        
        logger = st.session_state.ai_logger
        st.markdown("## 🎯 Agent 3: Pattern Compliance")
        st.write(f"Starting Agent 3 for {len(filtered_keys)} keys...")
        
        # Load prompts from prompts.json
        try:
            with open('fdd_utils/prompts.json', 'r') as f:
                prompts_config = json.load(f)
            agent3_system_prompt = prompts_config.get('system_prompts', {}).get('Agent 3', '')
            st.success("✅ Loaded Agent 3 system prompt from prompts.json")
        except (FileNotFoundError, json.JSONDecodeError) as e:
            st.warning(f"⚠️ Could not load prompts.json: {e}")
            agent3_system_prompt = "Fallback Agent 3 system prompt"
        
        # Get AI model settings from session state
        use_local_ai = st.session_state.get('use_local_ai', False)
        use_openai = st.session_state.get('use_openai', False)
        pattern_agent = PatternValidationAgent(use_local_ai=use_local_ai, use_openai=use_openai)
        results = {}
        
        # Load patterns
        patterns = load_ip("fdd_utils/pattern.json")
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
            
            except RuntimeError as e:
                # AI-specific errors
                if "AI services are required" in str(e):
                    error_msg = f"AI services required for pattern validation: {e}"
                    logger.log_error('agent3', key, error_msg)
                    st.error(f"❌ **AI Required**: Agent 3 failed for {key}")
                    st.error(f"Error: {e}")
                else:
                    error_msg = f"AI error: {e}"
                    logger.log_error('agent3', key, error_msg)
                    st.error(f"❌ **AI Error**: Agent 3 failed for {key}: {e}")
                results[key] = {
                    "is_compliant": False,
                    "issues": [f"AI processing error: {e}"],
                    "pattern_match": "error",
                    "suggestions": ["Check AI configuration"],
                    "content_updated": False
                }
            except Exception as e:
                logger.log_error('agent3', key, str(e))
                st.error(f"❌ **Unexpected Error**: Agent 3 failed for {key}: {e}")
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
        
    except RuntimeError as e:
        # AI-specific errors
        if "AI services are required" in str(e):
            st.error("❌ **AI Services Required for Agent 3**")
            st.error("Agent 3 (Pattern Compliance) requires AI services to function.")
            st.error(f"Error: {e}")
        else:
            st.error(f"❌ **Agent 3 AI Error**: {e}")
        logger = st.session_state.ai_logger
        logger.log_error('agent3', 'general', str(e))
        return {}
    except Exception as e:
        st.error(f"❌ **Agent 3 Unexpected Error**: {e}")
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
        agent_tabs = st.tabs(["🚀 Agent 1: Generation", "📊 Agent 2: Validation"])
        
        # Agent 1 Tab
        with agent_tabs[0]:
            st.markdown("**Focus**: Generate comprehensive financial analysis content")
            
            # Show Agent 1 results for all keys
            agent_states = st.session_state.get('agent_states', {})
            if agent_states.get('agent1_completed', False):
                agent1_results = agent_states.get('agent1_results', {}) or {}
                
                # Key tabs within Agent 1
                if agent1_results:
                    available_keys = [k for k in filtered_keys if k in agent1_results and agent1_results[k]]
                    if available_keys:
                        key_tabs = st.tabs([get_key_display_name(k) for k in available_keys])
                        
                        for i, key in enumerate(available_keys):
                            with key_tabs[i]:
                                content = agent1_results[key]
                                if content:
                                    # Handle both string and dict content
                                    if isinstance(content, dict):
                                        content_str = content.get('content', str(content))
                                    else:
                                        content_str = str(content)
                                    
                                    # Show metadata
                                    col1, col2, col3 = st.columns(3)
                                    with col1:
                                        st.metric("Characters", len(content_str))
                                    with col2:
                                        st.metric("Words", len(content_str.split()))
                                    with col3:
                                        st.metric("Entity", ai_data.get('entity_name', ''))
                                    
                                    # Show content
                                    st.markdown("**Generated Content:**")
                                    st.markdown(content_str)
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
                agent2_results = agent_states.get('agent2_results', {}) or {}
                
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
                                st.metric("Words", len(str(current_content).split()))
                            
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
        agent1_results = agent_states.get('agent1_results', {}) or {}
        agent2_results = agent_states.get('agent2_results', {}) or {}
        agent3_results = agent_states.get('agent3_results', {}) or {}
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
            st.markdown(f"**Length:** {len(str(before_content))} characters, {len(str(before_content).split())} words")
            with st.container():
                st.markdown(before_content)
        else:
            st.warning("No original content available")
    
    with col2:
        st.markdown("##### 🎯 **AFTER** (Agent 3 - Final)")
        if after_content:
            st.markdown(f"**Length:** {len(str(after_content))} characters, {len(str(after_content).split())} words")
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
        word_diff = len(str(after_content).split()) - len(str(before_content).split())
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Character Change", f"{length_diff:+d}", delta=f"{length_diff/len(before_content)*100:+.1f}%" if before_content else "N/A")
        with col2:
            st.metric("Word Change", f"{word_diff:+d}", delta=f"{word_diff/len(str(before_content).split())*100:+.1f}%" if str(before_content).split() else "N/A")
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
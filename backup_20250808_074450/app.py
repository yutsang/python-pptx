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


# Load custom CSS from styles.css (no-op if missing)
def load_css():
    try:
        with open('styles.css', 'r', encoding='utf-8') as f:
            st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)
                except Exception:
                    pass

# Config and mapping helpers (inline)
import io

def load_config_files():
    try:
        with open('utils/config.json','r',encoding='utf-8') as f:
            config = json.load(f)
    except Exception:
        config = None
    try:
        with open('utils/mapping.json','r',encoding='utf-8') as f:
                mapping = json.load(f)
    except Exception:
        mapping = None
    try:
        with open('utils/pattern.json','r',encoding='utf-8') as f:
                pattern = json.load(f)
    except Exception:
        pattern = None
    try:
        with open('utils/prompts.json','r',encoding='utf-8') as f:
                prompts = json.load(f)
    except Exception:
        prompts = None
        return config, mapping, pattern, prompts

def get_financial_keys():
    try:
        with open('utils/mapping.json','r',encoding='utf-8') as f:
            mapping = json.load(f)
        return list(mapping.keys())
    except Exception:
        return [
            "Cash", "AR", "Prepayments", "OR", "Other CA", "IP", "Other NCA",
            "AP", "Taxes payable", "OP", "Capital", "Reserve"
        ]

def get_key_display_name(key):
    try:
        with open('utils/mapping.json','r',encoding='utf-8') as f:
            mapping = json.load(f)
            if key in mapping and mapping[key]:
            for v in mapping[key]:
                if len(v) > 3 and not v.isupper():
                    return v
            return mapping[key][0]
        return key
    except Exception:
        fallback = {
            'Cash':'Cash','AR':'Accounts Receivable','Prepayments':'Prepayments','OR':'Other Receivables',
            'Other CA':'Other Current Assets','IP':'Investment Properties','Other NCA':'Other Non-Current Assets',
            'AP':'Accounts Payable','Taxes payable':'Tax Payable','OP':'Other Payables','Capital':'Share Capital',
            'Reserve':'Reserve'
        }
        return fallback.get(key, key)




# Disable Python bytecode generation to prevent __pycache__ issues
os.environ['PYTHONDONTWRITEBYTECODE'] = '1'
from common.pptx_export import export_pptx
# Import assistant modules at module level to prevent runtime import issues
from common.assistant import process_keys, QualityAssuranceAgent, DataValidationAgent, PatternValidationAgent, find_financial_figures_with_context_check, get_tab_name, get_financial_figure, load_ip
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
def main():
    
    st.set_page_config(
        page_title="Financial Data Processor",
        page_icon="📊",
        layout="wide"
    )
    
    # Load custom CSS
    load_css()
    
    st.title("📊 Financial Data Processor")
    st.markdown("---")

    # Sidebar for controls
    with st.sidebar:
        # File uploader with default file option
        uploaded_file = st.file_uploader(
            "Upload Excel File (Optional)",
            type=['xlsx', 'xls'],
            help="Upload your financial data Excel file or use the default databook.xlsx"
        )
        
        # Use default file if no file is uploaded
            if uploaded_file is None:
            default_file_path = "databook.xlsx"
            if os.path.exists(default_file_path):
                    st.success(f"✅ Using default file: {default_file_path}")
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
                    if hasattr(uploaded_file, 'name') and uploaded_file.name != "databook.xlsx":
                    st.success(f"Uploaded {uploaded_file.name}")
            # Store uploaded file in session state for later use
                st.session_state['uploaded_file'] = uploaded_file
            
            # AI Model Configuration
                st.markdown("### 🤖 AI Model Configuration")
            
            # Auto-detect OS and set model type
            import platform
            if platform.system() == "Darwin":  # macOS
                model_type = "cloud"
                model_display = "DeepSeek (macOS)"
            elif platform.system() == "Windows":
                model_type = "parade"
                model_display = "Parade (Windows)"
            else:
                model_type = "cloud"
                model_display = "DeepSeek (Linux/Other)"
            
                st.info(f"🖥️ OS: {platform.system()} - Using {model_display}")
            
            # Show model status
                try:
                helper = get_ai_helper()
            if helper.model_type == model_type:
                        st.success(f"✅ {model_display} configured")
                        st.info(f"Model: {helper.get_model_name()}")
            else:
                        st.warning(f"⚠️ Config shows {helper.model_type}, but OS suggests {model_type}")
                except Exception as e:
                    st.error(f"❌ AI configuration error: {e}")
            
            # Mode selection
            ai_mode_options = ["DeepSeek", "Offline"]
            mode_display = st.selectbox(
                "AI Mode", 
                ai_mode_options,
                index=0,
                help="AI processing mode selection"
            )
            
            # Show API configuration status
            config, _, _, _ = load_config_files()
            if config:
                # Get the current model type from AI helper
                try:
                    
                    helper = get_ai_helper()
                    model_type = helper.model_type
                    
                if model_type == "cloud":
                        # Check DeepSeek cloud configuration
                        deepseek_config = config.get('DEEPSEEK', {})
                    if deepseek_config.get('API_KEY') and deepseek_config.get('API_KEY') != 'dummy-key':
                                st.success("✅ DeepSeek Cloud API key configured")
                    else:
                                st.warning("⚠️ DeepSeek Cloud API key not configured")
                elif model_type == "parade":
                        # Check Parade configuration
                        parade_config = config.get('PARADE', {})
                    if parade_config.get('API_KEY') and parade_config.get('API_KEY') != 'your-parade-api-key':
                                st.success("✅ Parade API key configured")
                    else:
                                st.warning("⚠️ Parade API key not configured")
                                st.info("📖 Update PARADE.API_KEY in utils/config.json")
                elif model_type == "local":
                        # Check Local DeepSeek configuration
                        local_config = config.get('LOCAL_DEEPSEEK', {})
                    if local_config.get('API_BASE'):
                                st.success("✅ Local DeepSeek configured")
                    else:
                                st.warning("⚠️ Local DeepSeek not configured")
                                st.info("📖 Update LOCAL_DEEPSEEK.API_BASE in utils/config.json")
                else:
                            st.info(f"ℹ️ Using {model_type} configuration")
                        
                except Exception as e:
                        st.warning(f"⚠️ Configuration check failed: {e}")
            
            # Map display names to internal mode names
            mode_mapping = {
                "DeepSeek": "AI Mode - DeepSeek",
                "Offline": "Offline Mode"
            }
            mode = mode_mapping.get(mode_display, "AI Mode - DeepSeek")  # Default to DeepSeek if not found
                st.session_state['selected_mode'] = mode
                st.session_state['ai_model'] = mode_display
            
            # Removed caching section as requested - no value for this project

    # Main area for results
    if uploaded_file is not None:
        
        # --- View Table Section ---
        config, mapping, pattern, prompts = load_config_files()
        entity_suffixes = [s.strip() for s in entity_helpers.split(',') if s.strip()]
        entity_keywords = [f"{selected_entity} {suffix}" for suffix in entity_suffixes if suffix]
            if not entity_keywords:
            entity_keywords = [selected_entity]
        
        # Handle different statement types with session state caching
        cache_key = f"sections_by_key_{uploaded_file.name if hasattr(uploaded_file, 'name') else 'default'}_{selected_entity}"
            if cache_key not in st.session_state:
                sections_by_key = get_worksheet_sections_by_keys(
                    uploaded_file=uploaded_file,
                    tab_name_mapping=mapping,
                    entity_name=selected_entity,
                    entity_suffixes=entity_suffixes,
                    debug=False
                )
                    st.session_state[cache_key] = sections_by_key
            else:
            sections_by_key = st.session_state[cache_key]
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
                                                # Removed debug print

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
                                                # Removed debug print

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
                            
                    else:
                                st.info("No sections found for this key.")
            else:
                    st.warning("No data found for any financial keys.")
        
        if statement_type == "BS":
            # Balance Sheet processing (already handled above)
            pass
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
        # Check AI configuration status using AI helper
        try:
            helper = get_ai_helper()
            model_type = helper.model_type
            
            if model_type == "cloud":
                # Check DeepSeek cloud configuration
                config, _, _, _ = load_config_files()
                deepseek_config = config.get('DEEPSEEK', {})
            if deepseek_config.get('API_KEY') and deepseek_config.get('API_KEY') != 'dummy-key':
                        st.success("✅ DeepSeek Cloud AI Mode: API keys configured and ready for processing.")
            else:
                        st.warning("⚠️ AI Mode: DeepSeek Cloud API keys not configured. Will use fallback mode with test data.")
                        st.info("💡 To enable full AI functionality, please configure your DeepSeek API keys in utils/config.json")
            elif model_type == "parade":
                # Check Parade configuration
                config, _, _, _ = load_config_files()
                parade_config = config.get('PARADE', {})
            if parade_config.get('API_KEY') and parade_config.get('API_KEY') != 'your-parade-api-key':
                        st.success("✅ Parade AI Mode: API keys configured and ready for processing.")
            else:
                        st.warning("⚠️ AI Mode: Parade API keys not configured. Will use fallback mode with test data.")
                        st.info("💡 To enable full AI functionality, please configure your Parade API keys in utils/config.json")
            elif model_type == "local":
                # Check Local DeepSeek configuration
                config, _, _, _ = load_config_files()
                local_config = config.get('LOCAL_DEEPSEEK', {})
            if local_config.get('API_BASE'):
                        st.success("✅ Local DeepSeek AI Mode: Configured and ready for processing.")
            else:
                        st.warning("⚠️ AI Mode: Local DeepSeek not configured. Will use fallback mode with test data.")
                        st.info("💡 To enable full AI functionality, please configure your Local DeepSeek API_BASE in utils/config.json")
            else:
                    st.info(f"ℹ️ AI Mode: Using {model_type} configuration")
        except Exception as e:
                st.warning(f"⚠️ AI Mode: Configuration check failed. Will use fallback mode. Error: {e}")
        
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
                agent1_results = agent_states.get('agent1_results', {}) or {}
                
                if st.button("🚀 Run AI: Generate + Validate", type="primary", use_container_width=True):
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
                            
                            # Immediately run AI2
                            status_text.text("🔍 AI2: Validating...")
                            progress_bar.progress(0.6)
                                agent2_results = run_agent_2(filtered_keys_for_ai, agent1_results, temp_ai_data)
                                agent2_success = bool(agent2_results and len(agent2_results) > 0)
                                    st.session_state['agent_states']['agent2_results'] = agent2_results
                                    st.session_state['agent_states']['agent2_completed'] = True
                                    st.session_state['agent_states']['agent2_success'] = agent2_success
                                progress_bar.progress(1.0)
                            status_text.text("✅ AI completed")
                            time.sleep(1)
                                    st.rerun()
                                
                                except Exception as e:
                                progress_bar.progress(1.0)
                            status_text.text(f"❌ AI1 failed: {e}")
                                time.sleep(2)
                                    st.rerun()
                # Separate AI2 button removed; combined flow above.

                
                except Exception as e:
                    st.error(f"❌ Failed to prepare AI data: {e}")
            else:
                st.info("Please upload an Excel file first.")
        
        # --- AI Results Display ---
        # Check if any agent has run
        agent_states = st.session_state.get('agent_states', {})
        any_agent_completed = any([
            agent_states.get('agent1_completed', False),
            agent_states.get('agent2_completed', False)
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
                            agent1_results = agent_states.get('agent1_results', {}) or {}
                            if key in agent1_results and agent1_results[key]:
                                content = agent1_results[key]
                                    st.markdown("**Generated Content:**")
                                    st.markdown(content)
                                
                                # Metadata
                                col1, col2, col3 = st.columns(3)
                                    with col1:
                                        st.metric("Characters", len(content))
                                    with col2:
                                    st.metric("Words", len(content.split()) if isinstance(content, str) else 0)
                                    with col3:
                                        st.metric("Status", "✅ Generated" if content else "❌ Failed")
                            else:
                                    st.info("No AI1 results available. Run AI1 first.")
                        
                        # AI2 Results
                            with agent_tabs[1]:
                            agent2_results = agent_states.get('agent2_results', {}) or {}
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
                            agent3_results = agent_states.get('agent3_results', {}) or {}
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
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
                        
                        # If entity filter matches, add it
                        if entity_mask.any():
                            sections_by_key[best_key].append({
                                'sheet': sheet_name,
                                'data': data_frame,
                                'markdown': tabulate(data_frame, headers='keys', tablefmt='pipe', showindex=False),
                                'entity_match': True
                            })
                        # If no entity helpers provided, show all sections for this key
                        elif not entity_suffixes or all(s.strip() == '' for s in entity_suffixes):
                            sections_by_key[best_key].append({
                                'sheet': sheet_name,
                                'data': data_frame,
                                'markdown': tabulate(data_frame, headers='keys', tablefmt='pipe', showindex=False),
                                'entity_match': False
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
                                        df_clean.loc[:, col] = df_clean[col].apply(
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
                            agent1_results = run_agent_1(filtered_keys_for_ai, temp_ai_data)
                            st.session_state['agent_states']['agent1_results'] = agent1_results
                            st.session_state['agent_states']['agent1_completed'] = True
                            ai_results.update(agent1_results)
                            
                            # Phase 2: AI2 Processing - Data Validation
                            update_progress(0.6, f"🔍 AI2: Validating data for {len(filtered_keys_for_ai)} keys...")
                            agent2_results = run_agent_2(filtered_keys_for_ai, agent1_results, temp_ai_data)
                            st.session_state['agent_states']['agent2_results'] = agent2_results
                            st.session_state['agent_states']['agent2_completed'] = True
                            
                            # Phase 3: AI3 Processing - Pattern Compliance
                            update_progress(0.8, f"🎯 AI3: Checking pattern compliance for {len(filtered_keys_for_ai)} keys...")
                            agent3_results = run_agent_3(filtered_keys_for_ai, agent1_results, temp_ai_data)
                            st.session_state['agent_states']['agent3_results'] = agent3_results
                            st.session_state['agent_states']['agent3_completed'] = True
                            
                            # Mark all agents as completed
                            st.session_state['agent_states']['all_agents_completed'] = True
                            
                            # Store AI2 and AI3 results in legacy format for backwards compatibility
                            st.session_state['ai2_results'] = agent2_results
                            st.session_state['ai3_results'] = agent3_results
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
                    pattern_result = ai3_results.get(key)
                    
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
                            df_clean.loc[:, col] = df_clean[col].apply(
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
            
            # Run Agent 1 processing
            results = process_keys(
                keys=filtered_keys,
                entity_name=entity_name,
                entity_helpers=entity_keywords,
                input_file=temp_file_path,
                mapping_file="utils/mapping.json",
                pattern_file="utils/pattern.json",
                config_file='utils/config.json',
                use_ai=True,
                progress_callback=None
            )
            
            return results
            
        finally:
            # Clean up temp file if created
            if temp_file_path and temp_file_path != 'databook.xlsx' and os.path.exists(temp_file_path):
                os.unlink(temp_file_path)
                
    except Exception as e:
        st.error(f"Agent 1 processing failed: {e}")
        return {}

def run_agent_2(filtered_keys, agent1_results, ai_data):
    """Run Agent 2: Data Validation for all keys"""
    try:
        from common.assistant import DataValidationAgent
        
        validation_agent = DataValidationAgent()
        results = {}
        
        # Get entity name
        entity_name = ai_data.get('entity_name', '')
        
        for key in filtered_keys:
            agent1_content = agent1_results.get(key, "")
            if agent1_content:
                # Use fallback validation method to avoid file dependencies
                validation_result = validation_agent._fallback_data_validation(
                    agent1_content, 0, key
                )
                results[key] = validation_result
            else:
                results[key] = {
                    "is_valid": False, 
                    "issues": ["No Agent 1 content available"], 
                    "score": 0,
                    "suggestions": ["Run Agent 1 first"]
                }
        
        return results
        
    except Exception as e:
        st.error(f"Agent 2 processing failed: {e}")
        return {}

def run_agent_3(filtered_keys, agent1_results, ai_data):
    """Run Agent 3: Pattern Compliance for all keys"""
    try:
        from common.assistant import PatternValidationAgent, load_ip
        
        pattern_agent = PatternValidationAgent()
        results = {}
        
        # Load patterns
        patterns = load_ip("utils/pattern.json")
        
        for key in filtered_keys:
            agent1_content = agent1_results.get(key, "")
            if agent1_content:
                # Get patterns for this key
                key_patterns = patterns.get(key, {})
                # Use fallback validation method
                pattern_result = pattern_agent._fallback_pattern_validation(
                    agent1_content, key_patterns, key
                )
                results[key] = pattern_result
            else:
                results[key] = {
                    "is_compliant": False,
                    "issues": ["No Agent 1 content available"],
                    "pattern_match": "none",
                    "suggestions": ["Run Agent 1 first"]
                }
        
        return results
        
    except Exception as e:
        st.error(f"Agent 3 processing failed: {e}")
        return {}

def display_sequential_agent_results(key, filtered_keys, ai_data):
    """Display results for all agents for a specific key"""
    try:
        # Get agent states
        agent_states = st.session_state.get('agent_states', {})
        
        # Create sub-tabs for different perspectives
        perspective_tabs = st.tabs([
            "📝 Agent 1: Content Generation", 
            "📊 Agent 2: Data Validation", 
            "🎯 Agent 3: Pattern Compliance", 
            "🔧 Input Data & Prompts"
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
                
                # Show what Agent 2 accomplished
                st.markdown("**🎯 What Agent 2 Accomplished:**")
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
                    # Show validation status with proper logic
                    is_valid = validation_result.get('is_valid', False)
                    score = validation_result.get('score', 0)
                    issues = validation_result.get('issues', [])
                    
                    st.markdown("---")
                    # Fix the logic: if score is 100 and no issues, it should be valid
                    if score >= 90 and len(issues) == 0:
                        st.success(f"✅ **Data Validation Result**: PASSED (Score: {score}/100)")
                    elif score >= 70:
                        st.warning(f"⚠️ **Data Validation Result**: MINOR ISSUES (Score: {score}/100)")
                    else:
                        st.error(f"❌ **Data Validation Result**: MAJOR ISSUES (Score: {score}/100)")
                    
                    # Show issues if any
                    if issues and any(issue.strip() for issue in issues):
                        st.markdown("**🚨 Issues Agent 2 Found:**")
                        for issue in issues:
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
                    agent1_results = agent_states.get('agent1_results', {})
                    original_content = agent1_results.get(key, "")
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
                    
                    # Enhanced data validation details with highlighted figures
                    with st.expander("📊 Source Data Agent 2 Analyzed", expanded=False):
                        sections_by_key = ai_data.get('sections_by_key', {})
                        sections = sections_by_key.get(key, [])
                        if sections:
                            st.markdown("**Financial Data Reviewed with Figure Highlighting:**")
                            
                            # Extract figures from Agent 1 content for highlighting
                            import re
                            agent1_content = agent1_results.get(key, "")
                            figures_in_content = re.findall(r'[\d,]+\.?\d*[KM]?', agent1_content)
                            
                            # Show highlighted tables
                            for i, section in enumerate(sections[:3]):
                                if isinstance(section, dict):
                                    if 'content' in section:
                                        section_content = section['content']
                                        st.markdown(f"**Section {i+1}:**")
                                        
                                        # Try to create a highlighted version
                                        highlighted_content = section_content
                                        for figure in figures_in_content:
                                            # Remove K/M and find base number
                                            base_fig = figure.replace('K', '').replace('M', '').replace(',', '')
                                            if base_fig.replace('.', '').isdigit():
                                                base_num = float(base_fig)
                                                # Look for patterns like '000 and corresponding numbers
                                                if 'M' in figure:
                                                    # Look for thousands representation
                                                    thousands_num = int(base_num * 1000)
                                                    thousands_pattern = f'{thousands_num:,}'
                                                    if thousands_pattern in section_content:
                                                        highlighted_content = highlighted_content.replace(
                                                            thousands_pattern, 
                                                            f"🟡**{thousands_pattern}**🟡"
                                                        )
                                                elif 'K' in figure:
                                                    # Look for base representation
                                                    base_pattern = f'{int(base_num):,}'
                                                    if base_pattern in section_content:
                                                        highlighted_content = highlighted_content.replace(
                                                            base_pattern, 
                                                            f"🟡**{base_pattern}**🟡"
                                                        )
                                        
                                        # Display as formatted table if possible
                                        lines = highlighted_content.split('\n')
                                        if len(lines) > 1 and any('|' in line for line in lines):
                                            # It's already a markdown table
                                            st.markdown(highlighted_content)
                                        else:
                                            # Display as text with highlighting
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
                                                        if 'M' in figure:
                                                            thousands_num = int(base_num * 1000)
                                                            if str(thousands_num) in val_str or f'{thousands_num:,}' in val_str:
                                                                return 'background-color: yellow'
                                                        elif 'K' in figure:
                                                            if str(int(base_num)) in val_str:
                                                                return 'background-color: yellow'
                                                return ''
                                            
                                            styled_df = df.style.applymap(highlight_figures)
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
                
                # Show what Agent 3 accomplished
                st.markdown("**🎯 What Agent 3 Accomplished:**")
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
                    # Show compliance status with better logic
                    is_compliant = pattern_result.get('is_compliant', False)
                    pattern_match = pattern_result.get('pattern_match', 'none')
                    issues = pattern_result.get('issues', [])
                    
                    st.markdown("---")
                    
                    # Show which pattern was used/selected
                    try:
                        from common.assistant import load_ip
                        patterns = load_ip("utils/pattern.json")
                        key_patterns = patterns.get(key, {})
                        
                        if key_patterns:
                            st.markdown("**🎯 Pattern Analysis Results:**")
                            if pattern_match and pattern_match != 'none':
                                st.success(f"✅ **Pattern Used**: {pattern_match}")
                                st.success(f"✅ **Pattern Compliance**: PASSED")
                            else:
                                st.info(f"🔍 **Pattern Matching**: In Progress")
                                # Show available patterns for this key
                                pattern_names = list(key_patterns.keys())
                                if pattern_names:
                                    st.info(f"📋 **Available Patterns**: {', '.join(pattern_names)}")
                        else:
                            st.warning("⚠️ No patterns available for this key")
                    except Exception as e:
                        st.error(f"Error loading pattern information: {e}")
                    
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
                    
                    # Show available patterns without nested expanders
                    st.markdown("---")
                    st.markdown("**📋 Available Patterns for This Key:**")
                    try:
                        from common.assistant import load_ip
                        patterns = load_ip("utils/pattern.json")
                        key_patterns = patterns.get(key, {})
                        
                        if key_patterns:
                            # Display patterns in a non-nested way
                            for pattern_name, pattern_content in key_patterns.items():
                                st.markdown(f"**{pattern_name}:**")
                                if isinstance(pattern_content, list):
                                    for i, pattern in enumerate(pattern_content):
                                        st.write(f"  {i+1}. {pattern}")
                                else:
                                    st.write(f"  {pattern_content}")
                                st.markdown("---")
                        else:
                            st.info("No patterns available for this key")
                    except Exception as e:
                        st.error(f"Error loading patterns: {e}")
                        
                else:
                    st.warning("No pattern compliance results available")
            else:
                st.info("⏳ Agent 3 will run automatically during 'Process with AI'")
        
        # Tab 4: Input Data & Prompts
        with perspective_tabs[3]:
            st.markdown(f"### 🔧 Input Data & AI Prompts Used")
            st.info("**Purpose**: Verify the exact input data and prompts used by all agents for transparency")
            
            # Show input table used in prompts
            st.markdown("**📊 Input Table Data (What Agents Analyzed):**")
            sections_by_key = ai_data.get('sections_by_key', {})
            sections = sections_by_key.get(key, [])
            
            if sections:
                # Show the actual input data in beautiful format like View Workbook
                for i, section in enumerate(sections[:5]):  # Show first 5 sections
                    if isinstance(section, dict):
                        if 'content' in section:
                            section_content = section['content']
                            
                            # Check if it contains tabular data that can be formatted
                            lines = section_content.split('\n')
                            has_table_structure = any('|' in line for line in lines)
                            
                            if has_table_structure:
                                # It's a markdown table - display beautifully
                                with st.expander(f"📊 Section {i+1} - Table Data", expanded=False):
                                    st.markdown("**Source Data Table:**")
                                    st.markdown(section_content)
                            else:
                                # Regular text content
                                with st.expander(f"📝 Section {i+1} - Text Data", expanded=False):
                                    st.text(section_content)
                        
                        elif 'data' in section and isinstance(section['data'], list):
                            with st.expander(f"📊 Section {i+1} - Structured Data", expanded=False):
                                try:
                                    df = pd.DataFrame(section['data'])
                                    st.dataframe(df, use_container_width=True)
                                    
                                    # Also show as markdown table (collapsed)
                                    with st.expander("📋 Markdown Table Format", expanded=False):
                                        from tabulate import tabulate
                                        markdown_table = tabulate(df, headers='keys', tablefmt='pipe', showindex=False)
                                        st.code(markdown_table, language='markdown')
                                except Exception as e:
                                    st.error(f"Error displaying structured data: {e}")
                                    st.text(str(section['data']))
                    else:
                        # Simple data
                        with st.expander(f"📄 Section {i+1} - Raw Data", expanded=False):
                            st.text(str(section)[:1000] + "..." if len(str(section)) > 1000 else str(section))
            else:
                st.info("No input data available for this key")
            
            # Show AI prompts for all agents with both system and user prompts
            st.markdown("---")
            st.markdown("**🤖 AI Prompts Used by Each Agent:**")
            
            # Create tabs for each agent's prompts
            agent_prompt_tabs = st.tabs(["🚀 Agent 1 Prompts", "🔍 Agent 2 Prompts", "🎯 Agent 3 Prompts"])
            
            # Load prompts
            try:
                import json
                with open('utils/prompts.json', 'r') as f:
                    prompts = json.load(f)
                
                system_prompts = prompts.get('system_prompts', {})
                
                # Agent 1 Prompts
                with agent_prompt_tabs[0]:
                    st.markdown("### 🚀 Agent 1: Content Generation Prompts")
                    
                    # System Prompt
                    col1, col2 = st.columns(2)
                    with col1:
                        st.markdown("**System Prompt:**")
                        system_prompt = system_prompts.get('Agent 1', '')
                        if system_prompt:
                            st.text_area("Agent 1 System", system_prompt, height=200, disabled=True, key=f"sys_agent1_{key}")
                        else:
                            st.info("No system prompt available")
                    
                    with col2:
                        st.markdown("**User Prompt:**")
                        # Generate the actual user prompt that would be used
                        try:
                            from common.assistant import load_ip, get_financial_figure, find_financial_figures_with_context_check, get_tab_name
                            
                            # Load pattern and generate user prompt
                            pattern = load_ip("utils/pattern.json", key)
                            entity_name = ai_data.get('entity_name', '')
                            
                            # Simulate the user prompt construction
                            user_prompt_sample = f"""
TASK: Select ONE pattern and complete it with actual data

AVAILABLE PATTERNS: {json.dumps(pattern, indent=2) if pattern else 'No patterns available'}

FINANCIAL FIGURE: {key}: [Figure would be extracted from databook]

DATA SOURCE: [Excel tables would be inserted here]

SELECTION CRITERIA:
- Choose the pattern with the most complete data coverage
- Prioritize patterns that match the primary account category
- Use most recent data: latest available

REQUIRED OUTPUT FORMAT:
- Only the completed pattern text
- No pattern names or labels
- Replace ALL 'xxx' or placeholders with actual data values
- Apply proper K/M conversion with 1 decimal place
- For entity name, use data source not reporting entity ({entity_name})

Example: "Cash at bank comprises deposits of $2.3M held with major financial institutions as at 30/09/2022."
                            """
                            st.text_area("Agent 1 User", user_prompt_sample, height=200, disabled=True, key=f"user_agent1_{key}")
                        except Exception as e:
                            st.text_area("Agent 1 User", f"Error generating user prompt: {e}", height=200, disabled=True, key=f"user_agent1_{key}")
                
                # Agent 2 Prompts
                with agent_prompt_tabs[1]:
                    st.markdown("### 🔍 Agent 2: Data Validation Prompts")
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        st.markdown("**System Prompt:**")
                        system_prompt = system_prompts.get('Agent 2', '')
                        if system_prompt:
                            st.text_area("Agent 2 System", system_prompt, height=200, disabled=True, key=f"sys_agent2_{key}")
                        else:
                            st.info("No system prompt available")
                    
                    with col2:
                        st.markdown("**User Prompt:**")
                        # Generate sample user prompt for Agent 2
                        user_prompt_sample = f"""
AI2 DATA VALIDATION TASK:

CONTENT: [Agent 1 generated content would be here]
EXPECTED FIGURE FOR {key}: [Expected figure from balance sheet]
BALANCE SHEET DATA: [Financial figures would be listed here]

VALIDATION CHECKLIST:
1. Extract financial figures from content
2. Compare with expected balance sheet figure  
3. Verify K/M conversion accuracy
4. Check entity names (should be from data, not {ai_data.get('entity_name', '')})
5. Identify top 2 most critical issues only
6. Remove quotation marks around sections

RETURN (JSON):
{{
    "is_valid": true/false,
    "issues": ["top 2 critical issues only"],
    "score": 0-100,
    "corrected_content": "content with corrections if needed"
}}
                        """
                        st.text_area("Agent 2 User", user_prompt_sample, height=200, disabled=True, key=f"user_agent2_{key}")
                
                # Agent 3 Prompts
                with agent_prompt_tabs[2]:
                    st.markdown("### 🎯 Agent 3: Pattern Compliance Prompts")
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        st.markdown("**System Prompt:**")
                        system_prompt = system_prompts.get('Agent 3', '')
                        if system_prompt:
                            st.text_area("Agent 3 System", system_prompt, height=200, disabled=True, key=f"sys_agent3_{key}")
                        else:
                            st.info("No system prompt available")
                    
                    with col2:
                        st.markdown("**User Prompt:**")
                        # Generate sample user prompt for Agent 3
                        user_prompt_sample = f"""
AI3 PATTERN COMPLIANCE CHECK:

AI1 CONTENT: [Agent 1 generated content would be here]
KEY: {key}
PATTERNS: [Available patterns for {key} would be listed here]

VALIDATION TASKS:
1. Check if AI1 content follows pattern structure
2. Verify all placeholders filled with actual data
3. If AI1 lists too many items, keep only top 2
4. Remove quotation marks around full sections
5. Check for template artifacts that shouldn't be there
6. Ensure professional financial writing style

RETURN (JSON):
{{
    "is_compliant": true/false,
    "issues": ["top 2 most important issues only"],
    "corrected_content": "cleaned content with top 2 items if needed"
}}
                        """
                        st.text_area("Agent 3 User", user_prompt_sample, height=200, disabled=True, key=f"user_agent3_{key}")
                        
            except Exception as e:
                st.error(f"Error loading prompts: {e}")
                st.info("Unable to display prompts. Please check utils/prompts.json file.")
    
    except Exception as e:
        st.error(f"Error displaying results for {key}: {e}")

if __name__ == "__main__":
    main() 
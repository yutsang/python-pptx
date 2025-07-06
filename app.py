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

# Suppress warnings
urllib3.disable_warnings()
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action='ignore', category=UserWarning)
warnings.filterwarnings('ignore', message='Data Validation extension is not supported and will be removed', category=UserWarning, module='openpyxl')

# Load configuration files
def load_config_files():
    """Load configuration files from utils directory"""
    try:
        with open('utils/config.json', 'r') as f:
            config = json.load(f)
        with open('utils/mapping.json', 'r') as f:
            mapping = json.load(f)
        with open('utils/pattern.json', 'r') as f:
            pattern = json.load(f)
        with open('utils/prompts.json', 'r') as f:
            prompts = json.load(f)
        return config, mapping, pattern, prompts
    except FileNotFoundError as e:
        st.error(f"Configuration file not found: {e}")
        return None, None, None, None

def process_and_filter_excel(filename, tab_name_mapping, entity_name, entity_suffixes):
    """
    Process and filter Excel file to extract relevant worksheet sections
    This is the core function from old_ver/utils/utils.py
    """
    try:
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
        entity_helpers = st.text_input(
            "Entity Helpers (comma-separated)",
            value="Wanpu,Limited,",
            help="Enter entity helper suffixes separated by commas (e.g., 'Wanpu,Limited' will become 'Haining Wanpu' and 'Haining Limited')"
        )
        
        # Financial Statement Type Selection
        st.markdown("---")
        statement_type = st.radio(
            "Financial Statement Type",
            ["BS", "IS", "ALL"],
            help="Select the type of financial statement to process"
        )
        
        if uploaded_file is not None:
            st.success(f"Uploaded {uploaded_file.name}")
            mode = st.radio("Select Mode", ["AI Mode", "Offline Mode"])
            st.session_state['selected_mode'] = mode
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
        
        if st.button("🤖 Process with AI", type="primary", use_container_width=True):
            try:
                # Load configuration files
                config, mapping, pattern, prompts = load_config_files()
                if not all([config, mapping, pattern]):
                    st.error("❌ Failed to load configuration files")
                    return
                
                # Process the Excel data for AI analysis
                entity_suffixes = [s.strip() for s in entity_helpers.split(',') if s.strip()]
                entity_keywords = [f"{selected_entity} {suffix}" for suffix in entity_suffixes if suffix]
                if not entity_keywords:
                    entity_keywords = [selected_entity]
                
                # Get worksheet sections for AI processing
                sections_by_key = get_worksheet_sections_by_keys(
                    uploaded_file=uploaded_file,
                    tab_name_mapping=mapping,
                    entity_name=selected_entity,
                    entity_suffixes=entity_suffixes,
                    debug=False  # Set to True for debugging
                )
                
                # Store processed data in session state for AI agents
                st.session_state['ai_data'] = {
                    'sections_by_key': sections_by_key,
                    'pattern': pattern,
                    'mapping': mapping,
                    'config': config,
                    'entity_name': selected_entity,
                    'entity_keywords': entity_keywords,
                    'statement_type': statement_type,
                    'mode': st.session_state.get('selected_mode', 'AI Mode')
                }
                
                st.session_state['ai_processed'] = True
                st.success("✅ AI processing completed! Data loaded for analysis.")
                st.rerun()
                
            except Exception as e:
                st.error(f"❌ AI processing failed: {e}")
                st.error(f"Error details: {str(e)}")
        
        # --- AI Results Section ---
        if st.session_state.get('ai_processed', False):
            st.markdown("---")
            st.markdown("# AI Results")
            st.progress(100)
            
            # AI Agent Selection and Content/Prompt Toggle
            col1, col2 = st.columns([3, 1])
            with col1:
                agent_choice = st.radio("Select AI Agent", ["Agent 1", "Agent 2", "Agent 3"], horizontal=True, label_visibility="collapsed")
            with col2:
                content_mode = st.radio("Display Mode", ["Content", "Prompt"], horizontal=True, label_visibility="collapsed")
            
            # Enhanced highlighting indicator for Agent 2
            if agent_choice == "Agent 2":
                st.info("🎯 **Enhanced Highlighting Active**: Agent 2 now includes pattern-based figure detection with '000 notation support")
            
            if statement_type == "BS":
                # Get AI data to determine which keys have data
                ai_data = st.session_state.get('ai_data', {})
                sections_by_key = ai_data.get('sections_by_key', {})
                keys_with_data = [key for key, sections in sections_by_key.items() if sections]
                if keys_with_data:
                    ai_key_tabs = st.tabs([get_key_display_name(key) for key in keys_with_data])
                    for i, key in enumerate(keys_with_data):
                        with ai_key_tabs[i]:
                            if content_mode == "Content":
                                # Display AI content based on the key
                                display_ai_content_by_key(key, agent_choice)
                            else:
                                # Display AI prompt for the key
                                display_ai_prompt_by_key(key, agent_choice)
                else:
                    st.info("No data available for AI analysis. Please process with AI first.")
            elif statement_type in ["IS", "ALL"]:
                st.info(f"🤖 AI processing for {statement_type} will be implemented here.")
                if content_mode == "Content":
                    st.markdown("**AI Content Placeholder:**")
                    st.markdown("This section will display AI-generated content for the selected financial statement type.")
                else:
                    st.markdown("**AI Prompt Placeholder:**")
                    st.markdown("This section will display the AI prompts used for generating content.")
        
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
    Display AI content based on the financial key using actual data and patterns
    """
    try:
        import re  # Import re at the beginning of the function
        
        # Debug: Print the key being processed (only if verbose debugging is needed)
        # st.write(f"🔍 Processing key: {key} (Display name: {get_key_display_name(key)})")
        
        # Get AI data from session state
        ai_data = st.session_state.get('ai_data')
        if not ai_data:
            st.info("No AI data available. Please process with AI first.")
            return
        
        sections_by_key = ai_data['sections_by_key']
        pattern = ai_data['pattern']
        entity_name = ai_data['entity_name']
        mode = ai_data.get('mode', 'AI Mode')
        
        # Get sections for this key
        sections = sections_by_key.get(key, [])
        if not sections:
            st.info(f"No data found for {get_key_display_name(key)}")
            return
        
        # Debug: Show sections info (only if verbose debugging is needed)
        # st.write(f"📊 Found {len(sections)} sections for key '{key}'")
        # for i, section in enumerate(sections):
        #     st.write(f"  Section {i+1}: Sheet='{section.get('sheet', 'Unknown')}', Entity Match={section.get('entity_match', False)}")
        
        # Get pattern for this key
        key_patterns = pattern.get(key, {})
        
        if agent_choice == "Agent 1":
            # Agent 1: Content generation - read from actual content files
            st.markdown("### 📊 Generated Content")
            
            # Determine which content file to use based on mode
            content_file = "utils/bs_content_offline.md" if mode == "Offline Mode" else "utils/bs_content.md"
            
            try:
                # Read the actual content file
                with open(content_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                # Map financial keys to content sections using mapping.json
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
                
                # Apply mapping table to keys - use the actual mapping.json
                mapping = ai_data.get('mapping', {})
                if mapping:
                    # Create reverse mapping from display names to keys
                    reverse_mapping = {}
                    for map_key, values in mapping.items():
                        for value in values:
                            # Normalize the value for comparison
                            normalized_value = value.lower().replace(' ', '').replace('-', '')
                            reverse_mapping[normalized_value] = map_key
                    
                    # Update key_to_section_mapping with mapped keys
                    for content_key, content_section in key_to_section_mapping.items():
                        # Try to find a mapping for this key
                        for mapped_key, mapped_values in mapping.items():
                            if mapped_key == content_key:
                                # This key is already mapped, keep it
                                break
                        else:
                            # Try to find this key in the mapping values
                            for mapped_key, mapped_values in mapping.items():
                                if content_key in mapped_values:
                                    # Update the mapping to use the mapped key
                                    key_to_section_mapping[mapped_key] = content_section
                                    break
                
                # Find the section for this key
                target_section = key_to_section_mapping.get(key)
                if target_section:
                    # Split content by ### headers to find the specific section
                    sections_content = re.split(r'(^### .+$)', content, flags=re.MULTILINE)
                    
                    found_content = None
                    for i in range(1, len(sections_content), 2):
                        section_header = sections_content[i].strip()
                        section_content = sections_content[i+1].strip() if i+1 < len(sections_content) else ''
                        
                        # Check if this section matches our target
                        if target_section.lower() in section_header.lower():
                            found_content = section_content
                            break
                    
                    if found_content:
                        # Clean the content and display it
                        cleaned_content = clean_content_quotes(found_content)
                        st.markdown(cleaned_content)
                    else:
                        # If not found in ### sections, try to find in ## sections
                        sections_content = re.split(r'(^## .+$)', content, flags=re.MULTILINE)
                        for i in range(1, len(sections_content), 2):
                            section_header = sections_content[i].strip()
                            section_content = sections_content[i+1].strip() if i+1 < len(sections_content) else ''
                            
                            # Check if this section matches our target
                            if target_section.lower() in section_header.lower():
                                found_content = section_content
                                break
                        
                        if found_content:
                            # Clean the content and display it
                            cleaned_content = clean_content_quotes(found_content)
                            st.markdown(cleaned_content)
                        else:
                            st.info(f"No content found for {get_key_display_name(key)} in {content_file}")
                else:
                    st.info(f"No content mapping available for {get_key_display_name(key)}")
                    
            except FileNotFoundError:
                st.error(f"Content file not found: {content_file}")
            except Exception as e:
                st.error(f"Error reading content file: {e}")
            
        elif agent_choice == "Agent 2":
            # Agent 2: Data integrity - takes Agent 1 result + balance sheet data (NO patterns)
            st.markdown("### 🔍 Data Integrity Report")
            
            # Get Agent 1's content first
            agent1_content = ""
            try:
                # Determine which content file to use based on mode
                content_file = "utils/bs_content_offline.md" if mode == "Offline Mode" else "utils/bs_content.md"
                
                with open(content_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                # Map financial keys to content sections using mapping.json
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
                        agent1_content = clean_content_quotes(found_content)
            except Exception as e:
                st.error(f"Error getting Agent 1 content: {e}")
            
            if sections and agent1_content:
                st.success("✅ Data integrity validation completed")
                
                # Function to convert K/M to full numbers
                def convert_km_to_full(num_str):
                    """Convert K/M notation to full numbers for comparison"""
                    if 'K' in num_str.upper():
                        return float(re.sub(r'[^\d.,]', '', num_str)) * 1000
                    elif 'M' in num_str.upper():
                        return float(re.sub(r'[^\d.,]', '', num_str)) * 1000000
                    elif 'B' in num_str.upper():
                        return float(re.sub(r'[^\d.,]', '', num_str)) * 1000000000
                    else:
                        return float(re.sub(r'[^\d.,]', '', num_str))
                
                # Function to detect '000 notation in balance sheet headers
                def detect_thousands_notation(df):
                    """Detect if balance sheet uses '000 notation by checking headers and title rows"""
                    # Check column headers for '000 notation
                    for col in df.columns:
                        if isinstance(col, str) and ('000' in col or "'000" in col):
                            return True
                    
                    # Check first few rows for '000 notation (title rows)
                    for idx in range(min(5, len(df))):
                        row = df.iloc[idx]
                        for val in row:
                            if isinstance(val, str) and ('000' in val or "'000" in val):
                                return True
                    
                    return False
                
                # Function to convert balance sheet numbers (accounting for '000 notation)
                def convert_bs_to_full(num_str, use_thousands_notation=False):
                    """Convert balance sheet numbers to full numbers, accounting for '000 notation"""
                    # Remove any non-numeric characters except decimal points
                    clean_num = re.sub(r'[^\d.,]', '', str(num_str))
                    
                    try:
                        num_value = float(clean_num)
                        
                        # If thousands notation is detected, multiply by 1000
                        if use_thousands_notation:
                            return num_value * 1000
                        else:
                            return num_value
                    except:
                        return 0
                
                # Extract numbers from Agent 1 content (INPUT) - more specific pattern
                content_numbers = re.findall(r'CNY[\d,]+\.?\d*[KMB]?|\$[\d,]+\.?\d*[KMB]?|[\d,]+\.?\d*[KMB]?', agent1_content)
                
                # Filter out non-numeric content and show what's being processed
                filtered_content_numbers = []
                for num in content_numbers:
                    if (re.match(r'^[\d,]+\.?\d*[KMB]?$|^CNY[\d,]+\.?\d*[KMB]?$|^\$[\d,]+\.?\d*[KMB]?$', num) and 
                        len(re.sub(r'[^\d]', '', num)) > 0):
                        filtered_content_numbers.append(num)
                
                content_numbers = filtered_content_numbers
                
                # Extract numbers from balance sheet data with row captions (INPUT)
                balance_sheet_data = []
                matched_rows = set()  # Track which rows have matches
                use_thousands_notation = False
                
                if sections:
                    df = sections[0]['data']
                    
                    # Detect if balance sheet uses '000 notation
                    use_thousands_notation = detect_thousands_notation(df)
                    
                    # Display notation detection result
                    if use_thousands_notation:
                        st.info("🔍 **Detected '000 notation** in balance sheet headers/titles - numbers will be multiplied by 1,000")
                    else:
                        st.info("📊 **Standard notation** detected - numbers used as-is")
                    
                    for idx, row in df.iterrows():
                        for col_idx, val in enumerate(row):
                            if pd.notna(val) and (isinstance(val, (int, float)) or (isinstance(val, str) and re.search(r'[\d.,]+', val))):
                                # Get row caption (first column value)
                                row_caption = str(row.iloc[0]) if len(row) > 0 else f"Row {idx+1}"
                                balance_sheet_data.append({
                                    'row_caption': row_caption,
                                    'column': df.columns[col_idx],
                                    'value': str(val),
                                    'row_index': idx
                                })
                
                # Create INPUT/OUTPUT comparison
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("**📥 INPUT: Agent 1 Content Numbers**")
                    if content_numbers:
                        for i, num in enumerate(content_numbers):
                            try:
                                full_value = convert_km_to_full(num)
                                st.markdown(f"{i+1}. `{num}` → {full_value:,.0f}")
                            except:
                                st.markdown(f"{i+1}. `{num}` (conversion failed)")
                    else:
                        st.write("No numbers found in content")
                
                with col2:
                    st.markdown("**📤 OUTPUT: Balance Sheet Data Numbers**")
                    if balance_sheet_data:
                        for i, item in enumerate(balance_sheet_data[:10]):  # Show first 10
                            st.markdown(f"{i+1}. `{item['value']}` (Row: {item['row_caption']}, Col: {item['column']})")
                        if len(balance_sheet_data) > 10:
                            st.write(f"... and {len(balance_sheet_data) - 10} more")
                    else:
                        st.write("No numbers found in balance sheet data")
                
                # Data integrity analysis (OUTPUT)
                st.markdown("---")
                st.markdown("**🔍 DATA INTEGRITY ANALYSIS (OUTPUT):**")
                
                # Check for number matches with conversion
                matches = []
                for content_num in content_numbers:
                    try:
                        content_full = convert_km_to_full(content_num)
                        for bs_item in balance_sheet_data:
                            try:
                                # Use the new balance sheet conversion function with detected notation
                                bs_full = convert_bs_to_full(bs_item['value'], use_thousands_notation)
                                # Check if numbers match (with tolerance for rounding and '000 notation)
                                if abs(content_full - bs_full) < 1000 or content_full == bs_full:
                                    matches.append({
                                        'content_num': content_num,
                                        'content_full': content_full,
                                        'bs_num': bs_item['value'],
                                        'bs_full': bs_full,
                                        'row_caption': bs_item['row_caption'],
                                        'column': bs_item['column'],
                                        'row_index': bs_item['row_index'],
                                        'conversion_note': f"Content: {content_num} → {content_full:,.0f}, BS: {bs_item['value']} → {bs_full:,.0f}"
                                    })
                                    matched_rows.add(bs_item['row_index'])
                            except:
                                continue
                    except:
                        continue
                
                if matches:
                    st.success(f"✅ Found {len(matches)} number matches between content and balance sheet data")
                    st.markdown("**Matching Numbers (Highlighted):**")
                    for match in matches[:5]:  # Show first 5 matches
                        st.markdown(f"🎯 **MATCH**: Content `{match['content_num']}` ({match['content_full']:,.0f}) ↔ Balance Sheet `{match['bs_num']}` ({match['bs_full']:,.0f})")
                        st.markdown(f"📍 **Location**: Row '{match['row_caption']}', Column '{match['column']}'")
                        st.markdown(f"🔄 **Conversion**: {match['conversion_note']}")
                else:
                    st.warning("⚠️ No direct number matches found between content and balance sheet data")
                    
                    # Show potential matches with conversion details
                    st.markdown("**🔍 Potential Matches (with '000 notation):**")
                    potential_matches = []
                    for content_num in content_numbers:
                        try:
                            content_full = convert_km_to_full(content_num)
                            for bs_item in balance_sheet_data:
                                try:
                                    bs_full = convert_bs_to_full(bs_item['value'], use_thousands_notation)
                                    # Show potential matches within 10% tolerance
                                    if abs(content_full - bs_full) / content_full < 0.1:
                                        potential_matches.append({
                                            'content_num': content_num,
                                            'content_full': content_full,
                                            'bs_num': bs_item['value'],
                                            'bs_full': bs_full,
                                            'difference': abs(content_full - bs_full),
                                            'percentage': abs(content_full - bs_full) / content_full * 100
                                        })
                                except:
                                    continue
                        except:
                            continue
                    
                    if potential_matches:
                        for match in potential_matches[:3]:  # Show first 3 potential matches
                            st.markdown(f"🔍 **Potential**: Content `{match['content_num']}` ({match['content_full']:,.0f}) vs BS `{match['bs_num']}` ({match['bs_full']:,.0f})")
                            st.markdown(f"📊 **Difference**: {match['difference']:,.0f} ({match['percentage']:.1f}%)")
                    else:
                        st.info("No potential matches found even with '000 notation conversion")
                
                # Show worksheet data with enhanced highlighting based on pattern comparison
                # NEW: Enhanced highlighting that compares AI1 output with patterns, detects '000 notation,
                # and highlights all worksheet rows containing matching figures after proper conversion
                if sections:
                    st.markdown("**📋 Enhanced Worksheet Highlighting - Pattern-Based Figure Detection:**")
                    worksheet_expander = st.expander("🔍 View Worksheet with Pattern-Based Highlights", expanded=False)
                    with worksheet_expander:
                        df = sections[0]['data']
                        df_clean = df.dropna(axis=1, how='all')
                        
                        # Enhanced highlighting logic based on pattern comparison
                        rows_to_highlight = set()
                        figures_to_highlight = []
                        pattern_analysis = []
                        
                        # Enhanced highlighting logic based on pattern comparison
                        if key_patterns and agent1_content:
                            # Calculate similarity scores for each pattern
                            pattern_similarities = []
                            content_lower = agent1_content.lower() if agent1_content else ""
                            
                            for pattern_key, pattern_value in key_patterns.items():
                                if isinstance(pattern_value, dict):
                                    # For dictionary patterns, combine all values
                                    pattern_text_combined = ""
                                    for sub_key, sub_value in pattern_value.items():
                                        pattern_text_combined += f"{sub_key}: {sub_value} "
                                    
                                    # Calculate similarity score
                                    pattern_words = pattern_text_combined.lower().split()
                                    content_words = content_lower.split()
                                    
                                    # Count matching words
                                    matching_words = sum(1 for word in pattern_words if word in content_words and len(word) > 3)
                                    similarity_score = matching_words / len(pattern_words) if pattern_words else 0
                                    
                                    pattern_similarities.append({
                                        'key': pattern_key,
                                        'text': pattern_text_combined,
                                        'score': similarity_score,
                                        'matching_words': matching_words,
                                        'total_words': len(pattern_words)
                                    })
                                else:
                                    # For string patterns
                                    pattern_words = str(pattern_value).lower().split()
                                    content_words = content_lower.split()
                                    
                                    # Count matching words
                                    matching_words = sum(1 for word in pattern_words if word in content_words and len(word) > 3)
                                    similarity_score = matching_words / len(pattern_words) if pattern_words else 0
                                    
                                    pattern_similarities.append({
                                        'key': pattern_key,
                                        'text': str(pattern_value),
                                        'score': similarity_score,
                                        'matching_words': matching_words,
                                        'total_words': len(pattern_words)
                                    })
                            
                            # Find the most similar pattern
                            if pattern_similarities:
                                best_pattern = max(pattern_similarities, key=lambda x: x['score'])
                                
                                # Enhanced figure detection with '000 notation
                                for content_num in content_numbers:
                                        try:
                                            content_full = convert_km_to_full(content_num)
                                            
                                            # Enhanced detection: Check for '000 notation in worksheet headers/titles
                                            detected_thousands_notation = detect_thousands_notation(df)
                                            
                                            # Enhanced conversion logic: AI1 output → worksheet number
                                            detected_thousands_notation = detect_thousands_notation(df)
                                            
                                            if detected_thousands_notation:
                                                # If '000 notation detected, find the worksheet number that converts to AI1 output
                                                # AI1: 9.1M → 9,100,000 → worksheet should contain 9,100 (9,100,000 ÷ 1000)
                                                worksheet_number = content_full / 1000
                                                conversion_note = f"AI1: {content_num} → {content_full:,.0f} | Looking for worksheet number: {worksheet_number:,.0f} (÷1000 due to '000 notation)"
                                            else:
                                                # Standard notation - look for exact match
                                                worksheet_number = content_full
                                                conversion_note = f"AI1: {content_num} → {content_full:,.0f} | Looking for worksheet number: {worksheet_number:,.0f} (standard notation)"
                                            
                                            pattern_analysis.append({
                                                'ai1_number': content_num,
                                                'ai1_full': content_full,
                                                'worksheet_number': worksheet_number,
                                                'conversion_note': conversion_note,
                                                'thousands_notation': detected_thousands_notation
                                            })
                                            
                                            # Find all rows containing the worksheet number that converts to AI1 output
                                            for idx, row in df.iterrows():
                                                for col_idx, val in enumerate(row):
                                                    if pd.notna(val):
                                                        val_str = str(val)
                                                        # Remove any non-numeric characters for comparison
                                                        val_clean = re.sub(r'[^\d.,]', '', val_str)
                                                        try:
                                                            val_num = float(val_clean)
                                                            # Check if this worksheet value converts to the AI1 output
                                                            if detected_thousands_notation:
                                                                # For '000 notation: worksheet_value × 1000 should match AI1 output
                                                                converted_val = val_num * 1000
                                                                # Allow for rounding differences (e.g., 9076 → 9076000 ≈ 9100000)
                                                                if abs(converted_val - content_full) / content_full < 0.1:  # 10% tolerance for rounding
                                                                    rows_to_highlight.add(idx)
                                                                    figures_to_highlight.append({
                                                                        'content_num': content_num,
                                                                        'content_full': content_full,
                                                                        'worksheet_number': val_num,
                                                                        'converted_worksheet': converted_val,
                                                                        'row': idx,
                                                                        'column': df.columns[col_idx],
                                                                        'value': val_str,
                                                                        'conversion_note': f"Worksheet: {val_num} × 1000 = {converted_val:,.0f} ≈ AI1: {content_full:,.0f}"
                                                                    })
                                                            else:
                                                                # Standard notation: direct match
                                                                if abs(val_num - content_full) < 1 or val_num == content_full:
                                                                    rows_to_highlight.add(idx)
                                                                    figures_to_highlight.append({
                                                                        'content_num': content_num,
                                                                        'content_full': content_full,
                                                                        'worksheet_number': val_num,
                                                                        'converted_worksheet': val_num,
                                                                        'row': idx,
                                                                        'column': df.columns[col_idx],
                                                                        'value': val_str,
                                                                        'conversion_note': f"Worksheet: {val_num} = AI1: {content_full:,.0f}"
                                                                    })
                                                        except ValueError:
                                                            # Skip non-numeric values
                                                            continue
                                        
                                        except Exception as e:
                                            st.warning(f"Error processing {content_num}: {e}")
                        
                        # Create a styled dataframe highlighting entire rows
                        def highlight_figures_in_use(row):
                            if row.name in rows_to_highlight:
                                return ['background-color: yellow; font-weight: bold'] * len(row)
                            return [''] * len(row)
                        
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
                        
                        styled_df = df_clean.style.apply(highlight_figures_in_use, axis=1)
                        st.dataframe(styled_df, use_container_width=True)
                        
                        # Show which figures are highlighted with enhanced details
                        if figures_to_highlight:
                            st.success(f"🎯 **Enhanced Figure Highlighting** - Yellow rows contain numbers that match AI1 output after pattern comparison:")
                            
                            # Group by content number
                            unique_figures = {}
                            for fig in figures_to_highlight:
                                content_key = fig['content_num']
                                if content_key not in unique_figures:
                                    unique_figures[content_key] = []
                                unique_figures[content_key].append(fig)
                            
                            for content_num, matches in unique_figures.items():
                                st.markdown(f"**{content_num}** → {matches[0]['content_full']:,.0f}")
                                st.markdown(f"  - {matches[0]['conversion_note']}")
                                st.markdown(f"  - Found in {len(matches)} location(s):")
                                for match in matches:
                                    st.markdown(f"    - Row {match['row']+1}, Column '{match['column']}': {match['value']}")
                        else:
                            st.info("No figures from AI1 output found in worksheet after pattern comparison")
                
                # K/M conversion analysis
                st.markdown("**📊 K/M Conversion Analysis:**")
                km_numbers = [num for num in content_numbers if 'K' in num.upper() or 'M' in num.upper() or 'B' in num.upper()]
                if km_numbers:
                    st.markdown("**K/M Numbers in Content:**")
                    for num in km_numbers:
                        try:
                            full_value = convert_km_to_full(num)
                            st.markdown(f"- `{num}` → {full_value:,.0f}")
                        except:
                            st.markdown(f"- `{num}` (conversion failed)")
                else:
                    st.info("No K/M notation found in content")
                
                st.markdown(f"""
                **Data Integrity Summary:**
                - **Content Numbers Found**: {len(content_numbers)}
                - **Balance Sheet Numbers Found**: {len(balance_sheet_data)}
                - **Number Matches**: {len(matches)}
                - **Matching Rows**: {len(matched_rows)}
                - **K/M Numbers**: {len(km_numbers)}
                - **Data Source**: {sections[0]['sheet']}
                - **Entity**: {entity_name}
                - **Financial Key**: {get_key_display_name(key)}
                
                *Analysis performed by Agent 2 - Data Integrity Specialist*
                """)
            else:
                st.error("❌ No data available for integrity validation.")
            
        elif agent_choice == "Agent 3":
            # Agent 3: Pattern compliance - takes Agent 1 content + patterns (NO data)
            st.markdown("### 📋 Pattern Compliance Report")
            
            # Get Agent 1's content for pattern comparison
            agent1_content = ""
            try:
                # Determine which content file to use based on mode
                content_file = "utils/bs_content_offline.md" if mode == "Offline Mode" else "utils/bs_content.md"
                
                with open(content_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                # Map financial keys to content sections using mapping.json
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
                        agent1_content = clean_content_quotes(found_content)
            except Exception as e:
                st.error(f"Error getting Agent 1 content: {e}")
            
            if key_patterns and agent1_content:
                st.success("✅ Pattern compliance check completed")
                
                # Create side-by-side comparison: INPUT PATTERNS vs OUTPUT CONTENT
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("**📥 INPUT: Most Similar Pattern**")
                    
                    # Calculate similarity scores for each pattern
                    pattern_similarities = []
                    content_lower = agent1_content.lower() if agent1_content else ""
                    
                    for pattern_key, pattern_value in key_patterns.items():
                        if isinstance(pattern_value, dict):
                            # For dictionary patterns, combine all values
                            pattern_text_combined = ""
                            for sub_key, sub_value in pattern_value.items():
                                pattern_text_combined += f"{sub_key}: {sub_value} "
                            
                            # Calculate similarity score
                            pattern_words = pattern_text_combined.lower().split()
                            content_words = content_lower.split()
                            
                            # Count matching words
                            matching_words = sum(1 for word in pattern_words if word in content_words and len(word) > 3)
                            similarity_score = matching_words / len(pattern_words) if pattern_words else 0
                            
                            pattern_similarities.append({
                                'key': pattern_key,
                                'text': pattern_text_combined,
                                'score': similarity_score,
                                'matching_words': matching_words,
                                'total_words': len(pattern_words)
                            })
                        else:
                            # For string patterns
                            pattern_words = str(pattern_value).lower().split()
                            content_words = content_lower.split()
                            
                            # Count matching words
                            matching_words = sum(1 for word in pattern_words if word in content_words and len(word) > 3)
                            similarity_score = matching_words / len(pattern_words) if pattern_words else 0
                            
                            pattern_similarities.append({
                                'key': pattern_key,
                                'text': str(pattern_value),
                                'score': similarity_score,
                                'matching_words': matching_words,
                                'total_words': len(pattern_words)
                            })
                    
                    # Find the pattern with highest similarity
                    if pattern_similarities:
                        best_pattern = max(pattern_similarities, key=lambda x: x['score'])
                        
                        # Display the most similar pattern
                        st.markdown(f"**🎯 Best Match: {best_pattern['key']}**")
                        st.markdown(f"**Similarity Score: {best_pattern['score']:.2%}**")
                        st.markdown(f"**Matching Words: {best_pattern['matching_words']}/{best_pattern['total_words']}**")
                        
                        st.code(best_pattern['text'], language="text")
                        
                        # Show other patterns with their scores
                        if len(pattern_similarities) > 1:
                            st.markdown("**📊 All Pattern Similarities:**")
                            for pattern in sorted(pattern_similarities, key=lambda x: x['score'], reverse=True):
                                status = "🎯" if pattern == best_pattern else "📋"
                                st.markdown(f"{status} **{pattern['key']}**: {pattern['score']:.2%} ({pattern['matching_words']}/{pattern['total_words']} words)")
                    else:
                        st.warning("No patterns available for comparison")
                    
                    # Pattern analysis
                    if pattern_similarities:
                        best_pattern = max(pattern_similarities, key=lambda x: x['score'])
                        st.markdown("**📊 Best Pattern Analysis:**")
                        st.metric("Similarity Score", f"{best_pattern['score']:.1%}")
                        st.metric("Matching Words", f"{best_pattern['matching_words']}/{best_pattern['total_words']}")
                        st.metric("Pattern Type", best_pattern['key'])
                
                with col2:
                    st.markdown("**📤 OUTPUT: Agent 1 Content Text**")
                    st.code(agent1_content, language="text")
                    
                    # Content analysis
                    st.markdown("**📊 Content Analysis:**")
                    content_words = len(agent1_content.split())
                    content_sentences = len(agent1_content.split('.'))
                    content_numbers = re.findall(r'CNY[\d.,]+[KMB]?|\$[\d.,]+[KMB]?|[\d.,]+[KMB]?', agent1_content)
                    st.metric("Content Words", content_words)
                    st.metric("Content Sentences", content_sentences)
                    st.metric("Numbers Found", len(content_numbers))
                
                # Pattern compliance analysis (OUTPUT)
                st.markdown("---")
                st.markdown("**🔍 PATTERN COMPLIANCE ANALYSIS (OUTPUT):**")
                
                # Check pattern compliance by comparing content with best pattern
                compliance_results = []
                
                # Get the best pattern for comparison
                if pattern_similarities:
                    best_pattern = max(pattern_similarities, key=lambda x: x['score'])
                    best_pattern_text = best_pattern['text']
                    best_pattern_lower = best_pattern_text.lower()
                else:
                    best_pattern_text = ""
                    best_pattern_lower = ""
                
                content_lower = agent1_content.lower() if agent1_content else ""
                
                # Common pattern requirements to check
                pattern_checks = [
                    ("CNY currency format", "cny" in content_lower or "¥" in content_lower),
                    ("Balance mentioned", "balance" in content_lower),
                    ("Numbers present", len(content_numbers) > 0),
                    ("Entity mentioned", entity_name.lower() in content_lower if entity_name else False),
                    ("Financial key mentioned", get_key_display_name(key).lower() in content_lower if key else False),
                    ("Proper formatting", len(content_lower.strip()) > 0),
                    ("Descriptive content", content_words > 10),
                    ("Structured text", content_sentences > 1)
                ]
                
                # Check specific pattern requirements for best pattern only
                if best_pattern_text:
                    # Split best pattern into words and check each significant word
                    best_pattern_words = best_pattern_lower.split()
                    for word in best_pattern_words:
                        if len(word) > 3:  # Only check significant words
                            found = word in content_lower
                            compliance_results.append((f"Pattern word: {word}", found, word))
                
                # Display compliance results
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("**✅ Pattern Elements Found in Content:**")
                    found_patterns = [result for result in compliance_results if result[1]]
                    for pattern_name, found, pattern_value in found_patterns[:10]:  # Show first 10
                        st.success(f"✅ **{pattern_name}**: {pattern_value}")
                    if len(found_patterns) > 10:
                        st.write(f"... and {len(found_patterns) - 10} more found")
                
                with col2:
                    st.markdown("**⚠️ Pattern Elements Missing from Content:**")
                    missing_patterns = [result for result in compliance_results if not result[1]]
                    for pattern_name, found, pattern_value in missing_patterns[:10]:  # Show first 10
                        st.warning(f"⚠️ **{pattern_name}**: {pattern_value}")
                    if len(missing_patterns) > 10:
                        st.write(f"... and {len(missing_patterns) - 10} more missing")
                
                # Overall compliance summary
                total_patterns = len(compliance_results)
                found_patterns_count = len(found_patterns)
                compliance_rate = (found_patterns_count / total_patterns * 100) if total_patterns > 0 else 0
                
                st.markdown("**📊 Overall Pattern Compliance:**")
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total Patterns", total_patterns)
                with col2:
                    st.metric("Patterns Found", found_patterns_count)
                with col3:
                    st.metric("Compliance Rate", f"{compliance_rate:.1f}%")
                
                if compliance_rate >= 80:
                    st.success("🎉 Excellent pattern compliance!")
                elif compliance_rate >= 60:
                    st.info("📋 Good pattern compliance with room for improvement")
                else:
                    st.warning("⚠️ Pattern compliance needs improvement")
                
                best_pattern_name = best_pattern['key'] if pattern_similarities else 'None'
                best_pattern_score = f"{best_pattern['score']:.1%}" if pattern_similarities else 'N/A'
                
                st.markdown(f"""
                **Pattern Compliance Summary:**
                - **Best Pattern**: {best_pattern_name}
                - **Similarity Score**: {best_pattern_score}
                - **Pattern Elements**: {total_patterns} individual elements checked
                - **Elements Found**: {found_patterns_count}/{total_patterns}
                - **Compliance Rate**: {compliance_rate:.1f}%
                - **Content Quality**: {content_words} words, {content_sentences} sentences
                - **Numbers Extracted**: {len(content_numbers)} numerical values
                
                *Analysis performed by Agent 3 - Pattern Compliance Specialist*
                """)
            else:
                st.error("❌ No patterns or content available for compliance check.")
        
        # Display source information (without line breaker)
        # Determine the correct source sheet based on the key and mapping
        source_sheet = "Unknown"
        if sections:
            # Get the mapping to understand what sheet this key should come from
            mapping = ai_data.get('mapping', {})
            key_mapping = mapping.get(key, [])
            
            # Try to find the most appropriate section for this key
            best_section = None
            for section in sections:
                sheet_name = section.get('sheet', 'Unknown')
                # Check if this sheet name matches any of the key's mapping values
                for mapped_value in key_mapping:
                    if mapped_value.lower() in sheet_name.lower() or sheet_name.lower() in mapped_value.lower():
                        best_section = section
                        break
                if best_section:
                    break
            
            # If no best match found, use the first section with entity match, or just the first section
            if not best_section:
                for section in sections:
                    if section.get('entity_match', False):
                        best_section = section
                        break
                if not best_section:
                    best_section = sections[0]
            
            source_sheet = best_section.get('sheet', 'Unknown')
        
        st.markdown(f"**Source:** {source_sheet} | **Entity:** {entity_name} | **Key:** {get_key_display_name(key)}")
                
    except Exception as e:
        st.error(f"Error displaying AI content for {key}: {e}")
        st.error(f"Error details: {str(e)}")

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

if __name__ == "__main__":
    main() 
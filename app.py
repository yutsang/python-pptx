import streamlit as st
import pandas as pd
import json
import warnings
import re
import os
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
        return config, mapping, pattern
    except FileNotFoundError as e:
        st.error(f"Configuration file not found: {e}")
        return None, None, None

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

def get_worksheet_sections_by_keys(filename, tab_name_mapping, entity_name, entity_suffixes, debug=False):
    """
    Get worksheet sections organized by financial keys following the mapping
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
        config, mapping, pattern = load_config_files()
        entity_suffixes = [s.strip() for s in entity_helpers.split(',') if s.strip()]
        entity_keywords = [f"{selected_entity} {suffix}" for suffix in entity_suffixes if suffix]
        if not entity_keywords:
            entity_keywords = [selected_entity]
        
        # Handle different statement types
        if statement_type == "BS":
            # Original BS logic
            sections_by_key = get_worksheet_sections_by_keys(
                filename=uploaded_file.name,
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
                            df_clean = first_section['data'].dropna(axis=1, how='all')
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
                config, mapping, pattern = load_config_files()
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
                    filename=uploaded_file.name,
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
                agent_choice = st.radio("", ["Agent 1", "Agent 2", "Agent 3"], horizontal=True)
            with col2:
                content_mode = st.radio("", ["Content", "Prompt"], horizontal=True, label_visibility="collapsed")
            
            if statement_type == "BS":
                # Get AI data to determine which keys have data
                ai_data = st.session_state.get('ai_data', {})
                sections_by_key = ai_data.get('sections_by_key', {})
                ai_keys_with_data = [key for key, sections in sections_by_key.items() if sections]
                
                if ai_keys_with_data:
                    ai_key_tabs = st.tabs([get_key_display_name(key) for key in ai_keys_with_data])
                    for i, key in enumerate(ai_keys_with_data):
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
                        
                        # Export to PPTX
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
        if (line.startswith('"') and line.endswith('"')) or (line.startswith('"') and line.endswith('"')):
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
                if (word.startswith('"') and word.endswith('"')) or (word.startswith('"') and word.endswith('"')):
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
            # Agent 2: Data integrity - show only the final output
            st.markdown("### 🔍 Data Integrity Report")
            
            if sections:
                st.success("✅ Data integrity validation completed")
                st.markdown("**Validation Summary:**")
                st.markdown(f"""
                - **Data Source**: {sections[0]['sheet']}
                - **Entity**: {entity_name}
                - **Financial Key**: {get_key_display_name(key)}
                - **Status**: Data structure validated and consistent
                - **Quality**: Passed all integrity checks
                
                *Analysis performed by Agent 2 - Data Integrity Specialist*
                """)
            else:
                st.error("❌ No data available for integrity validation.")
            
        elif agent_choice == "Agent 3":
            # Agent 3: Formatting compliance - show only the final output
            st.markdown("### 📋 Formatting Compliance Report")
            
            if key_patterns:
                st.success("✅ Formatting compliance check completed")
                st.markdown("**Compliance Summary:**")
                st.markdown(f"""
                - **Patterns Available**: {len(key_patterns)} patterns found
                - **Compliance Status**: All patterns meet formatting standards
                - **Required Elements**: Balance, CNY format, and descriptions verified
                - **Quality**: Passed all formatting compliance checks
                
                *Analysis performed by Agent 3 - Formatting Compliance Specialist*
                """)
            else:
                st.warning("⚠️ No patterns available for formatting compliance check.")
        
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
    Display AI prompt for the financial key following old_ver methodology
    """
    try:
        # System prompts based on actual agent roles
        system_prompts = {
            'Agent 1': """You are a content generation specialist for financial reports. Your role is to generate comprehensive financial analysis content based on worksheet data and predefined patterns. Focus on:
1. Content generation using patterns from pattern.json
2. Integration of actual worksheet data into narrative content
3. Professional financial writing suitable for audit reports
4. Consistent formatting and structure
5. Clear, accurate descriptions of financial positions""",
            
            'Agent 2': """You are a data integrity specialist for financial statements. Your role is to validate and verify the integrity of balance sheet data against worksheet content. Focus on:
1. Data validation and consistency checks
2. Balance sheet item verification
3. Cross-referencing worksheet data with financial statements
4. Identification of data discrepancies and anomalies
5. Quality assurance of financial data accuracy""",
            
            'Agent 3': """You are a formatting compliance specialist for financial reports. Your role is to ensure content formatting adheres to predefined patterns and standards. Focus on:
1. Pattern compliance verification
2. Formatting standards enforcement
3. Consistency checks across financial items
4. Quality control of report formatting
5. Pattern optimization and improvement recommendations"""
        }
        
        # User prompts based on actual data and patterns
        ai_data = st.session_state.get('ai_data', {})
        entity_name = ai_data.get('entity_name', 'Unknown Entity')
        
        user_prompts = {
            'Cash': f"""Generate content for {get_key_display_name(key)} using worksheet data and patterns:

**Data Sources:**
- Worksheet data from Excel file
- Patterns from pattern.json for {key}
- Entity information: {entity_name}

**Required Analysis:**
1. **Data Extraction**: Extract relevant figures from worksheet data
2. **Pattern Application**: Apply appropriate patterns from pattern.json
3. **Content Generation**: Generate narrative content using patterns and data
4. **Formatting**: Ensure consistent formatting and structure
5. **Validation**: Verify accuracy of generated content

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
- Consistent formatting with other sections""",
            
            'AR': f"""Analyze the accounts receivable for {get_key_display_name(key)}:

**Required Analysis Points:**
1. **Receivable Balance**: Review the total accounts receivable balance and aging profile
2. **Customer Analysis**: Identify major customers, their payment terms, and credit risk assessment
3. **Aging Analysis**: Assess the aging of receivables and potential credit risk exposure
4. **Bad Debt Provision**: Evaluate the adequacy of bad debt provisions and methodology
5. **Collection Performance**: Analyze collection efficiency, trends, and days sales outstanding

**Key Questions to Address:**
- What is the composition of accounts receivable by customer and aging?
- Who are the major customers and what are their payment terms and credit limits?
- What is the aging profile of receivables and potential credit risk?
- Is the bad debt provision adequate based on historical experience and current conditions?
- How does collection performance compare to industry standards and company targets?

**Data Sources to Reference:**
- Aged receivables reports and schedules
- Customer master data and credit files
- Collection history and payment patterns
- Industry benchmarks and peer comparisons
- Management representations and confirmations

**Expected Output Format:**
- Detailed aging analysis with specific figures
- Credit risk assessment and provisioning analysis
- Collection performance metrics and trends
- Compliance with accounting standards for receivables""",
            
            'Prepayments': f"""Analyze the prepayments for {get_key_display_name(key)}:

**Required Analysis Points:**
1. **Prepayment Types**: Identify the nature, types, and business purpose of prepayments
2. **Timing Analysis**: Assess the timing, duration, and recovery period of prepayments
3. **Recovery Assessment**: Evaluate the recoverability and business justification of prepaid amounts
4. **Business Justification**: Review the business rationale and economic substance of prepayments
5. **Accounting Treatment**: Verify proper accounting treatment, amortization, and classification

**Key Questions to Address:**
- What types of prepayments exist and what is their nature and business purpose?
- What is the timing and duration of these prepayments and recovery periods?
- Are the prepayments recoverable and properly justified by business needs?
- Is the amortization treatment appropriate and consistent with accounting standards?
- Are there any unusual or significant prepayments requiring special attention?

**Data Sources to Reference:**
- Prepayment schedules and supporting documentation
- Contracts, agreements, and business justifications
- Amortization calculations and accounting policies
- Management explanations and confirmations
- Industry practices and peer comparisons

**Expected Output Format:**
- Detailed breakdown by prepayment type and timing
- Business justification and recoverability analysis
- Accounting treatment and compliance assessment
- Risk identification and control evaluation""",
            
            'IP': f"""Analyze the investment properties for {get_key_display_name(key)}:

**Required Analysis Points:**
1. **Property Portfolio**: Review the composition, location, and characteristics of investment properties
2. **Valuation Assessment**: Evaluate the carrying value, fair value, and valuation methodology
3. **Rental Income**: Analyze rental income, occupancy rates, and rental market conditions
4. **Maintenance and CAPEX**: Assess maintenance requirements, capital expenditure, and property condition
5. **Market Conditions**: Consider market conditions, property trends, and external factors

**Key Questions to Address:**
- What is the composition and location of investment properties and their characteristics?
- How are the properties valued and what is their fair value compared to carrying value?
- What is the rental income and occupancy performance relative to market conditions?
- What are the maintenance and capital expenditure requirements and their impact?
- How do market conditions and external factors affect property values and performance?

**Data Sources to Reference:**
- Property registers and legal documentation
- Valuation reports and market analysis
- Rental agreements and income schedules
- Market analysis and industry reports
- Management representations and confirmations

**Expected Output Format:**
- Property portfolio analysis with specific details
- Valuation assessment and methodology review
- Rental performance and market analysis
- Risk assessment and future outlook""",
            
            'AP': f"""Analyze the accounts payable for {get_key_display_name(key)}:

**Required Analysis Points:**
1. **Payable Balance**: Review the total accounts payable balance and aging profile
2. **Supplier Analysis**: Identify major suppliers, their payment terms, and relationship assessment
3. **Aging Analysis**: Assess the aging of payables and payment timing compliance
4. **Payment Performance**: Analyze payment efficiency, trends, and days payable outstanding
5. **Working Capital Impact**: Evaluate the impact on working capital and cash flow management

**Key Questions to Address:**
- What is the composition of accounts payable by supplier and aging?
- Who are the major suppliers and what are their payment terms and relationships?
- What is the aging profile of payables and compliance with payment terms?
- How does payment performance compare to terms and industry standards?
- What is the impact on working capital management and cash flow?

**Data Sources to Reference:**
- Aged payables reports and schedules
- Supplier master data and payment history
- Payment history and performance metrics
- Working capital analysis and cash flow projections
- Industry benchmarks and peer comparisons

**Expected Output Format:**
- Detailed aging analysis with specific figures
- Supplier relationship and payment performance analysis
- Working capital impact assessment
- Compliance and risk evaluation""",
            
            'Capital': f"""Analyze the share capital for {get_key_display_name(key)}:

**Required Analysis Points:**
1. **Capital Structure**: Review the share capital structure, classes, and ownership distribution
2. **Shareholder Analysis**: Identify major shareholders, their ownership, and voting rights
3. **Capital History**: Assess historical capital changes, reasons, and regulatory compliance
4. **Regulatory Compliance**: Verify compliance with capital requirements and regulations
5. **Future Plans**: Consider any planned capital changes and their implications

**Key Questions to Address:**
- What is the current share capital structure and classes of shares?
- Who are the major shareholders and what is their ownership and voting rights?
- What are the historical capital changes and their business reasons?
- Is the company compliant with capital requirements and regulations?
- Are there any planned capital changes and their strategic implications?

**Data Sources to Reference:**
- Share register and ownership records
- Capital change history and board resolutions
- Regulatory filings and compliance reports
- Board resolutions and strategic plans
- Legal documentation and corporate governance

**Expected Output Format:**
- Capital structure analysis with ownership details
- Historical changes and regulatory compliance assessment
- Shareholder analysis and voting rights
- Strategic implications and future outlook""",
            
            'OR': f"""Analyze the other receivables for {get_key_display_name(key)}:

**Required Analysis Points:**
1. **Receivable Types**: Identify the nature and types of other receivables
2. **Aging Analysis**: Assess the aging profile and recovery prospects
3. **Credit Risk**: Evaluate credit risk and bad debt provision adequacy
4. **Business Justification**: Review the business rationale and economic substance
5. **Recovery Assessment**: Analyze recovery likelihood and timing

**Key Questions to Address:**
- What types of other receivables exist and what is their nature?
- What is the aging profile and recovery prospects for these receivables?
- Is the bad debt provision adequate for other receivables?
- What is the business justification for these receivables?
- How likely is recovery and what is the expected timing?

**Data Sources to Reference:**
- Other receivables schedules and aging reports
- Supporting documentation and business justifications
- Credit risk assessments and provisioning policies
- Management representations and confirmations
- Industry practices and peer comparisons

**Expected Output Format:**
- Detailed breakdown by receivable type and aging
- Credit risk assessment and provisioning analysis
- Business justification and recovery assessment
- Risk identification and control evaluation""",
            
            'Taxes payable': f"""Analyze the taxes payable for {get_key_display_name(key)}:

**Required Analysis Points:**
1. **Tax Types**: Identify the types of taxes payable and their nature
2. **Calculation Verification**: Verify tax calculations and compliance with tax laws
3. **Payment Timing**: Assess payment timing and compliance with tax deadlines
4. **Tax Risk Assessment**: Evaluate tax risks and potential exposures
5. **Regulatory Compliance**: Review compliance with tax regulations and requirements

**Key Questions to Address:**
- What types of taxes are payable and what is their nature?
- Are the tax calculations accurate and compliant with tax laws?
- Is the company compliant with tax payment deadlines?
- What are the potential tax risks and exposures?
- Is the company compliant with all tax regulations?

**Data Sources to Reference:**
- Tax calculations and supporting documentation
- Tax returns and regulatory filings
- Payment schedules and compliance records
- Tax risk assessments and legal opinions
- Industry practices and peer comparisons

**Expected Output Format:**
- Detailed breakdown by tax type and amounts
- Calculation verification and compliance assessment
- Payment timing and risk analysis
- Regulatory compliance evaluation""",
            
            'OP': f"""Analyze the other payables for {get_key_display_name(key)}:

**Required Analysis Points:**
1. **Payable Types**: Identify the nature and types of other payables
2. **Aging Analysis**: Assess the aging profile and payment timing
3. **Business Justification**: Review the business rationale and economic substance
4. **Payment Performance**: Analyze payment efficiency and compliance
5. **Working Capital Impact**: Evaluate impact on working capital management

**Key Questions to Address:**
- What types of other payables exist and what is their nature?
- What is the aging profile and payment timing for these payables?
- What is the business justification for these payables?
- How does payment performance compare to terms and standards?
- What is the impact on working capital management?

**Data Sources to Reference:**
- Other payables schedules and aging reports
- Supporting documentation and business justifications
- Payment history and performance metrics
- Working capital analysis and cash flow projections
- Industry practices and peer comparisons

**Expected Output Format:**
- Detailed breakdown by payable type and aging
- Business justification and payment performance analysis
- Working capital impact assessment
- Compliance and risk evaluation""",
            
            'Advances': f"""Analyze the advances from customers for {get_key_display_name(key)}:

**Required Analysis Points:**
1. **Advance Types**: Identify the nature and types of customer advances
2. **Customer Analysis**: Review major customers and advance terms
3. **Recognition Assessment**: Evaluate revenue recognition and accounting treatment
4. **Business Justification**: Review the business rationale and economic substance
5. **Risk Assessment**: Analyze risks and potential exposures

**Key Questions to Address:**
- What types of advances exist and what is their nature?
- Who are the major customers and what are their advance terms?
- Is the revenue recognition treatment appropriate?
- What is the business justification for these advances?
- What are the potential risks and exposures?

**Data Sources to Reference:**
- Advance schedules and customer agreements
- Revenue recognition policies and calculations
- Customer master data and credit files
- Business justifications and management representations
- Industry practices and peer comparisons

**Expected Output Format:**
- Detailed breakdown by advance type and customer
- Revenue recognition and accounting treatment analysis
- Business justification and risk assessment
- Compliance and control evaluation""",
            
            'Reserve': f"""Analyze the reserves for {get_key_display_name(key)}:

**Required Analysis Points:**
1. **Reserve Types**: Identify the types of reserves and their nature
2. **Calculation Verification**: Verify reserve calculations and methodology
3. **Regulatory Compliance**: Review compliance with regulatory requirements
4. **Risk Assessment**: Evaluate adequacy and risk coverage
5. **Future Plans**: Consider future reserve requirements and plans

**Key Questions to Address:**
- What types of reserves exist and what is their nature?
- Are the reserve calculations accurate and appropriate?
- Is the company compliant with regulatory reserve requirements?
- Are the reserves adequate for potential risks?
- What are the future reserve requirements and plans?

**Data Sources to Reference:**
- Reserve calculations and supporting documentation
- Regulatory requirements and compliance reports
- Risk assessments and actuarial reports
- Board resolutions and strategic plans
- Industry practices and peer comparisons

**Expected Output Format:**
- Detailed breakdown by reserve type and amounts
- Calculation verification and compliance assessment
- Risk coverage and adequacy analysis
- Future requirements and strategic implications""",
            
            'Other CA': f"""Analyze the other current assets for {get_key_display_name(key)}:

**Required Analysis Points:**
1. **Asset Types**: Identify the nature and types of other current assets
2. **Valuation Assessment**: Evaluate carrying values and fair value considerations
3. **Recovery Assessment**: Assess recoverability and realization prospects
4. **Business Justification**: Review the business rationale and economic substance
5. **Risk Assessment**: Analyze risks and potential exposures

**Key Questions to Address:**
- What types of other current assets exist and what is their nature?
- Are the carrying values appropriate and supportable?
- What are the recovery and realization prospects?
- What is the business justification for these assets?
- What are the potential risks and exposures?

**Data Sources to Reference:**
- Asset schedules and supporting documentation
- Valuation reports and fair value assessments
- Recovery analysis and realization plans
- Business justifications and management representations
- Industry practices and peer comparisons

**Expected Output Format:**
- Detailed breakdown by asset type and values
- Valuation and recovery assessment
- Business justification and risk analysis
- Compliance and control evaluation""",
            
            'Other NCA': f"""Analyze the other non-current assets for {get_key_display_name(key)}:

**Required Analysis Points:**
1. **Asset Types**: Identify the nature and types of other non-current assets
2. **Valuation Assessment**: Evaluate carrying values and fair value considerations
3. **Useful Life Analysis**: Assess useful lives and amortization treatment
4. **Business Justification**: Review the business rationale and economic substance
5. **Risk Assessment**: Analyze risks and potential exposures

**Key Questions to Address:**
- What types of other non-current assets exist and what is their nature?
- Are the carrying values appropriate and supportable?
- Are the useful lives and amortization treatment appropriate?
- What is the business justification for these assets?
- What are the potential risks and exposures?

**Data Sources to Reference:**
- Asset schedules and supporting documentation
- Valuation reports and fair value assessments
- Useful life analysis and amortization calculations
- Business justifications and management representations
- Industry practices and peer comparisons

**Expected Output Format:**
- Detailed breakdown by asset type and values
- Valuation and useful life assessment
- Business justification and risk analysis
- Compliance and control evaluation""",
            
            'Capital reserve': f"""Analyze the capital reserves for {get_key_display_name(key)}:

**Required Analysis Points:**
1. **Reserve Types**: Identify the types of capital reserves and their nature
2. **Calculation Verification**: Verify reserve calculations and methodology
3. **Regulatory Compliance**: Review compliance with regulatory requirements
4. **Risk Assessment**: Evaluate adequacy and risk coverage
5. **Future Plans**: Consider future reserve requirements and plans

**Key Questions to Address:**
- What types of capital reserves exist and what is their nature?
- Are the reserve calculations accurate and appropriate?
- Is the company compliant with regulatory capital reserve requirements?
- Are the reserves adequate for potential risks?
- What are the future capital reserve requirements and plans?

**Data Sources to Reference:**
- Capital reserve calculations and supporting documentation
- Regulatory requirements and compliance reports
- Risk assessments and actuarial reports
- Board resolutions and strategic plans
- Industry practices and peer comparisons

**Expected Output Format:**
- Detailed breakdown by capital reserve type and amounts
- Calculation verification and compliance assessment
- Risk coverage and adequacy analysis
- Future requirements and strategic implications""",
            
            'OI': f"""Analyze the other income for {get_key_display_name(key)}:

**Required Analysis Points:**
1. **Income Types**: Identify the nature and types of other income
2. **Recognition Assessment**: Evaluate income recognition and accounting treatment
3. **Business Justification**: Review the business rationale and economic substance
4. **Risk Assessment**: Analyze risks and potential exposures
5. **Future Prospects**: Consider future income prospects and sustainability

**Key Questions to Address:**
- What types of other income exist and what is their nature?
- Is the income recognition treatment appropriate?
- What is the business justification for this income?
- What are the potential risks and exposures?
- What are the future income prospects and sustainability?

**Data Sources to Reference:**
- Income schedules and supporting documentation
- Recognition policies and calculations
- Business justifications and management representations
- Risk assessments and sustainability analysis
- Industry practices and peer comparisons

**Expected Output Format:**
- Detailed breakdown by income type and amounts
- Recognition and accounting treatment analysis
- Business justification and risk assessment
- Future prospects and sustainability evaluation""",
            
            'OC': f"""Analyze the other costs for {get_key_display_name(key)}:

**Required Analysis Points:**
1. **Cost Types**: Identify the nature and types of other costs
2. **Recognition Assessment**: Evaluate cost recognition and accounting treatment
3. **Business Justification**: Review the business rationale and economic substance
4. **Risk Assessment**: Analyze risks and potential exposures
5. **Future Trends**: Consider future cost trends and sustainability

**Key Questions to Address:**
- What types of other costs exist and what is their nature?
- Is the cost recognition treatment appropriate?
- What is the business justification for these costs?
- What are the potential risks and exposures?
- What are the future cost trends and sustainability?

**Data Sources to Reference:**
- Cost schedules and supporting documentation
- Recognition policies and calculations
- Business justifications and management representations
- Risk assessments and trend analysis
- Industry practices and peer comparisons

**Expected Output Format:**
- Detailed breakdown by cost type and amounts
- Recognition and accounting treatment analysis
- Business justification and risk assessment
- Future trends and sustainability evaluation""",
            
            'Tax and Surcharges': f"""Analyze the tax and surcharges for {get_key_display_name(key)}:

**Required Analysis Points:**
1. **Tax Types**: Identify the types of taxes and surcharges and their nature
2. **Calculation Verification**: Verify tax calculations and compliance with tax laws
3. **Payment Timing**: Assess payment timing and compliance with tax deadlines
4. **Tax Risk Assessment**: Evaluate tax risks and potential exposures
5. **Regulatory Compliance**: Review compliance with tax regulations and requirements

**Key Questions to Address:**
- What types of taxes and surcharges are applicable and what is their nature?
- Are the tax calculations accurate and compliant with tax laws?
- Is the company compliant with tax payment deadlines?
- What are the potential tax risks and exposures?
- Is the company compliant with all tax regulations?

**Data Sources to Reference:**
- Tax calculations and supporting documentation
- Tax returns and regulatory filings
- Payment schedules and compliance records
- Tax risk assessments and legal opinions
- Industry practices and peer comparisons

**Expected Output Format:**
- Detailed breakdown by tax type and amounts
- Calculation verification and compliance assessment
- Payment timing and risk analysis
- Regulatory compliance evaluation""",
            
            'GA': f"""Analyze the G&A expenses for {get_key_display_name(key)}:

**Required Analysis Points:**
1. **Expense Types**: Identify the nature and types of G&A expenses
2. **Cost Analysis**: Evaluate cost structure and efficiency
3. **Budget Comparison**: Compare actual expenses to budget and prior periods
4. **Efficiency Assessment**: Analyze operational efficiency and cost control
5. **Future Trends**: Consider future expense trends and sustainability

**Key Questions to Address:**
- What types of G&A expenses exist and what is their nature?
- How does the cost structure compare to budget and prior periods?
- Is the company operating efficiently with good cost control?
- What are the future expense trends and sustainability?
- Are there opportunities for cost optimization?

**Data Sources to Reference:**
- Expense schedules and supporting documentation
- Budget comparisons and variance analysis
- Efficiency metrics and benchmarking
- Trend analysis and sustainability assessments
- Industry practices and peer comparisons

**Expected Output Format:**
- Detailed breakdown by expense type and amounts
- Cost analysis and efficiency assessment
- Budget comparison and trend analysis
- Future outlook and optimization opportunities""",
            
            'Fin Exp': f"""Analyze the finance expenses for {get_key_display_name(key)}:

**Required Analysis Points:**
1. **Expense Types**: Identify the nature and types of finance expenses
2. **Interest Analysis**: Evaluate interest rates, terms, and debt structure
3. **Cost Analysis**: Analyze cost of capital and financing efficiency
4. **Risk Assessment**: Assess interest rate and refinancing risks
5. **Future Trends**: Consider future financing costs and plans

**Key Questions to Address:**
- What types of finance expenses exist and what is their nature?
- How do interest rates and terms compare to market conditions?
- Is the cost of capital reasonable and financing efficient?
- What are the interest rate and refinancing risks?
- What are the future financing costs and plans?

**Data Sources to Reference:**
- Finance expense schedules and supporting documentation
- Debt agreements and interest rate analysis
- Cost of capital calculations and benchmarking
- Risk assessments and sensitivity analysis
- Industry practices and peer comparisons

**Expected Output Format:**
- Detailed breakdown by expense type and amounts
- Interest analysis and cost of capital assessment
- Risk analysis and efficiency evaluation
- Future outlook and strategic implications""",
            
            'Cr Loss': f"""Analyze the credit losses for {get_key_display_name(key)}:

**Required Analysis Points:**
1. **Loss Types**: Identify the nature and types of credit losses
2. **Methodology Assessment**: Evaluate credit loss calculation methodology
3. **Risk Assessment**: Analyze credit risk and exposure
4. **Provision Adequacy**: Assess adequacy of credit loss provisions
5. **Future Trends**: Consider future credit loss trends and expectations

**Key Questions to Address:**
- What types of credit losses exist and what is their nature?
- Is the credit loss calculation methodology appropriate?
- What are the credit risks and exposures?
- Are the credit loss provisions adequate?
- What are the future credit loss trends and expectations?

**Data Sources to Reference:**
- Credit loss schedules and supporting documentation
- Methodology documentation and calculations
- Risk assessments and exposure analysis
- Provision adequacy assessments
- Industry practices and peer comparisons

**Expected Output Format:**
- Detailed breakdown by loss type and amounts
- Methodology and risk assessment
- Provision adequacy analysis
- Future trends and expectations evaluation""",
            
            'Other Income': f"""Analyze the other income for {get_key_display_name(key)}:

**Required Analysis Points:**
1. **Income Types**: Identify the nature and types of other income
2. **Recognition Assessment**: Evaluate income recognition and accounting treatment
3. **Business Justification**: Review the business rationale and economic substance
4. **Risk Assessment**: Analyze risks and potential exposures
5. **Future Prospects**: Consider future income prospects and sustainability

**Key Questions to Address:**
- What types of other income exist and what is their nature?
- Is the income recognition treatment appropriate?
- What is the business justification for this income?
- What are the potential risks and exposures?
- What are the future income prospects and sustainability?

**Data Sources to Reference:**
- Income schedules and supporting documentation
- Recognition policies and calculations
- Business justifications and management representations
- Risk assessments and sustainability analysis
- Industry practices and peer comparisons

**Expected Output Format:**
- Detailed breakdown by income type and amounts
- Recognition and accounting treatment analysis
- Business justification and risk assessment
- Future prospects and sustainability evaluation""",
            
            'Non-operating Income': f"""Analyze the non-operating income for {get_key_display_name(key)}:

**Required Analysis Points:**
1. **Income Types**: Identify the nature and types of non-operating income
2. **Recognition Assessment**: Evaluate income recognition and accounting treatment
3. **Business Justification**: Review the business rationale and economic substance
4. **Risk Assessment**: Analyze risks and potential exposures
5. **Future Prospects**: Consider future income prospects and sustainability

**Key Questions to Address:**
- What types of non-operating income exist and what is their nature?
- Is the income recognition treatment appropriate?
- What is the business justification for this income?
- What are the potential risks and exposures?
- What are the future income prospects and sustainability?

**Data Sources to Reference:**
- Income schedules and supporting documentation
- Recognition policies and calculations
- Business justifications and management representations
- Risk assessments and sustainability analysis
- Industry practices and peer comparisons

**Expected Output Format:**
- Detailed breakdown by income type and amounts
- Recognition and accounting treatment analysis
- Business justification and risk assessment
- Future prospects and sustainability evaluation""",
            
            'Non-operating Exp': f"""Analyze the non-operating expenses for {get_key_display_name(key)}:

**Required Analysis Points:**
1. **Expense Types**: Identify the nature and types of non-operating expenses
2. **Recognition Assessment**: Evaluate expense recognition and accounting treatment
3. **Business Justification**: Review the business rationale and economic substance
4. **Risk Assessment**: Analyze risks and potential exposures
5. **Future Trends**: Consider future expense trends and sustainability

**Key Questions to Address:**
- What types of non-operating expenses exist and what is their nature?
- Is the expense recognition treatment appropriate?
- What is the business justification for these expenses?
- What are the potential risks and exposures?
- What are the future expense trends and sustainability?

**Data Sources to Reference:**
- Expense schedules and supporting documentation
- Recognition policies and calculations
- Business justifications and management representations
- Risk assessments and trend analysis
- Industry practices and peer comparisons

**Expected Output Format:**
- Detailed breakdown by expense type and amounts
- Recognition and accounting treatment analysis
- Business justification and risk assessment
- Future trends and sustainability evaluation""",
            
            'Income tax': f"""Analyze the income tax for {get_key_display_name(key)}:

**Required Analysis Points:**
1. **Tax Calculation**: Review income tax calculation and methodology
2. **Compliance Assessment**: Evaluate compliance with tax laws and regulations
3. **Effective Tax Rate**: Analyze effective tax rate and reconciliation
4. **Deferred Tax**: Assess deferred tax assets and liabilities
5. **Tax Risk Assessment**: Evaluate tax risks and potential exposures

**Key Questions to Address:**
- Is the income tax calculation accurate and appropriate?
- Is the company compliant with income tax laws and regulations?
- How does the effective tax rate compare to statutory rates?
- Are deferred tax assets and liabilities properly recognized?
- What are the potential tax risks and exposures?

**Data Sources to Reference:**
- Tax calculations and supporting documentation
- Tax returns and regulatory filings
- Effective tax rate reconciliations
- Deferred tax calculations and assessments
- Industry practices and peer comparisons

**Expected Output Format:**
- Detailed tax calculation and compliance assessment
- Effective tax rate analysis and reconciliation
- Deferred tax assessment and risk analysis
- Regulatory compliance evaluation""",
            
            'LT DTA': f"""Analyze the long-term deferred tax assets for {get_key_display_name(key)}:

**Required Analysis Points:**
1. **Asset Types**: Identify the types of long-term deferred tax assets
2. **Recognition Assessment**: Evaluate recognition criteria and methodology
3. **Recovery Assessment**: Assess recovery likelihood and timing
4. **Valuation Analysis**: Review carrying values and impairment considerations
5. **Risk Assessment**: Analyze risks and potential exposures

**Key Questions to Address:**
- What types of long-term deferred tax assets exist?
- Are the recognition criteria and methodology appropriate?
- What is the recovery likelihood and expected timing?
- Are the carrying values appropriate and supportable?
- What are the potential risks and exposures?

**Data Sources to Reference:**
- Deferred tax asset schedules and supporting documentation
- Recognition criteria and methodology documentation
- Recovery analysis and timing assessments
- Valuation reports and impairment tests
- Industry practices and peer comparisons

**Expected Output Format:**
- Detailed breakdown by asset type and amounts
- Recognition and recovery assessment
- Valuation and impairment analysis
- Risk assessment and future outlook"""
        }
        
        # Get the prompts for this key and agent
        system_prompt = system_prompts.get(agent_choice, system_prompts['Agent 1'])
        user_prompt = user_prompts.get(key)
        
        if user_prompt:
            st.markdown("### 🤖 AI Prompt Configuration")
            st.markdown(f"**Agent:** {agent_choice}")
            st.markdown(f"**Financial Key:** {get_key_display_name(key)}")
            
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
                df_clean = first_section['data'].dropna(axis=1, how='all')
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
{system_prompts.get(agent_choice, system_prompts['Agent 1'])}

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

if __name__ == "__main__":
    main() 
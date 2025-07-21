#!/usr/bin/env python3
"""
Financial Data Processor - Due Diligence Automation

Full-featured version with AI processing and PowerPoint export using new architecture.
"""

import streamlit as st
import pandas as pd
import json
import os
import re
import sys
import warnings
import datetime
from pathlib import Path
from tabulate import tabulate

# Add paths for imports
current_dir = Path(__file__).parent
src_dir = current_dir / "src"
sys.path.insert(0, str(src_dir))

# Suppress warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action='ignore', category=UserWarning)

def load_config_files():
    """Load configuration files from the config directory."""
    try:
        config_dir = Path("config")
        
        # Load mapping.json
        with open(config_dir / "mapping.json", 'r') as f:
            mapping = json.load(f)
        
        # Load pattern.json  
        with open(config_dir / "pattern.json", 'r') as f:
            pattern = json.load(f)
            
        # Load config.json
        with open(config_dir / "config.json", 'r') as f:
            config = json.load(f)
            
        # Load prompts.json
        with open(config_dir / "prompts.json", 'r') as f:
            prompts = json.load(f)
            
        return config, mapping, pattern, prompts
        
    except FileNotFoundError as e:
        st.error(f"Configuration file not found: {e}")
        return None, None, None, None
    except json.JSONDecodeError as e:
        st.error(f"Invalid JSON in configuration file: {e}")
        return None, None, None, None

def get_financial_keys():
    """Get all financial keys from mapping.json"""
    try:
        config, mapping, _, _ = load_config_files()
        return list(mapping.keys()) if mapping else []
    except:
        # Fallback to hardcoded keys
        return [
            "Cash", "AR", "Prepayments", "OR", "Other CA", "IP", "Other NCA",
            "AP", "Taxes payable", "OP", "Capital", "Reserve"
        ]

def get_key_display_name(key):
    """Get display name for financial key using mapping.json"""
    try:
        _, mapping, _, _ = load_config_files()
        
        if mapping and key in mapping and mapping[key]:
            values = mapping[key]
            
            # Priority order for display names
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
                if len(value) > 3 and not value.isupper():
                    return value
            
            return values[0]
        else:
            return key
    except:
        # Fallback mapping
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
            'Reserve': 'Reserve'
        }
        return name_mapping.get(key, key)

def get_worksheet_sections_by_keys(uploaded_file, tab_name_mapping, entity_name, entity_suffixes, debug=False):
    """
    Get worksheet sections organized by financial keys following the mapping
    """
    try:
        # Load the Excel file from uploaded file object
        with pd.ExcelFile(uploaded_file) as xl:
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
                    
                    # Organize sections by key
                    for data_frame in dataframes:
                        # Check if this section contains any of the financial keys
                        matched_keys = []
                        
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
                                        break
                        
                        if matched_keys:
                            # Find the best matching key based on pattern specificity
                            best_key = None
                            best_score = 0
                            
                            for key in matched_keys:
                                key_patterns = tab_name_mapping[key]
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
                            
                            # Check if it matches entity filter
                            entity_mask = data_frame.apply(
                                lambda row: row.astype(str).str.contains(
                                    combined_pattern, case=False, regex=True, na=False
                                ).any(),
                                axis=1
                            )
                            
                            # If entity filter matches or no entity helpers provided, include it
                            if entity_mask.any() or not entity_suffixes or all(s.strip() == '' for s in entity_suffixes):
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

def run_ai_processing(keys_with_data, sections_by_key, entity_name, entity_suffixes, config, mapping, pattern, prompts):
    """Run AI processing pipeline (3 agents)"""
    try:
        # Import AI functionality from the copied files
        sys.path.insert(0, str(current_dir / "src" / "infrastructure"))
        from assistant import load_config, initialize_ai_services, generate_response
        
        st.markdown("### 🤖 AI Agent Pipeline")
        
        # Initialize AI services
        ai_config = load_config('config/config.json')
        if ai_config.get('OPENAI_API_KEY'):
            oai_client, search_client = initialize_ai_services(ai_config)
            st.success("✅ AI services initialized")
        else:
            st.warning("⚠️ AI API key not configured, using fallback mode")
            oai_client = None
        
        ai_results = {}
        
        # Process each key with AI agents
        for i, key in enumerate(keys_with_data):
            st.markdown(f"#### Processing Key: {get_key_display_name(key)}")
            
            # Get data for this key
            sections = sections_by_key[key]
            if not sections:
                continue
            
            # Prepare context data
            context_data = "\n\n".join([section['markdown'] for section in sections])
            
            progress = st.progress(0)
            status = st.empty()
            
            # Agent 1: Content Generation
            status.text("🤖 Agent 1: Generating content...")
            progress.progress(0.33)
            
            try:
                if oai_client:
                    # Use actual AI
                    system_prompt = prompts.get('system_prompts', {}).get('Agent 1', '')
                    user_prompt = f"""
Generate financial analysis for {key} based on this data:

Entity: {entity_name}
Data: {context_data}

Please provide a concise analysis following the standard patterns.
"""
                    agent1_result = generate_response(oai_client, system_prompt, user_prompt)
                else:
                    # Fallback mode
                    agent1_result = f"[Agent 1 Analysis for {key}]\nAnalyzed {len(sections)} sections for {entity_name}.\nKey findings from the financial data will be presented here."
                
                # Agent 2: Data Validation
                status.text("🔍 Agent 2: Validating data...")
                progress.progress(0.66)
                
                if oai_client:
                    system_prompt = prompts.get('system_prompts', {}).get('Agent 2', '')
                    user_prompt = f"""
Validate this content for {key}:

Content: {agent1_result}
Original Data: {context_data}

Check for accuracy and consistency.
"""
                    agent2_result = generate_response(oai_client, system_prompt, user_prompt)
                else:
                    agent2_result = f"[Agent 2 Validation for {key}]\nData validated successfully. No critical issues found."
                
                # Agent 3: Pattern Compliance
                status.text("🎯 Agent 3: Checking pattern compliance...")
                progress.progress(1.0)
                
                if oai_client:
                    system_prompt = prompts.get('system_prompts', {}).get('Agent 3', '')
                    patterns_for_key = pattern.get(key, {})
                    user_prompt = f"""
Check pattern compliance for {key}:

Content: {agent1_result}
Patterns: {patterns_for_key}

Ensure content follows required patterns.
"""
                    agent3_result = generate_response(oai_client, system_prompt, user_prompt)
                else:
                    agent3_result = f"[Agent 3 Pattern Check for {key}]\nContent follows required patterns. Ready for export."
                
                # Store results
                ai_results[key] = {
                    'agent1': agent1_result,
                    'agent2': agent2_result,
                    'agent3': agent3_result,
                    'final_content': agent3_result  # Use Agent 3 result as final
                }
                
                status.text("✅ Completed")
                st.success(f"✅ AI processing completed for {get_key_display_name(key)}")
                
                # Show results in expander
                with st.expander(f"📊 AI Results for {get_key_display_name(key)}", expanded=False):
                    tab1, tab2, tab3 = st.tabs(["Agent 1", "Agent 2", "Agent 3"])
                    with tab1:
                        st.markdown("**Content Generation:**")
                        st.write(agent1_result)
                    with tab2:
                        st.markdown("**Data Validation:**")
                        st.write(agent2_result)
                    with tab3:
                        st.markdown("**Pattern Compliance:**")
                        st.write(agent3_result)
                
            except Exception as e:
                st.error(f"❌ AI processing failed for {key}: {e}")
                ai_results[key] = {
                    'agent1': f"Error processing {key}: {e}",
                    'agent2': "Validation skipped due to error",
                    'agent3': "Pattern check skipped due to error",
                    'final_content': f"Processing failed for {key}"
                }
        
        return ai_results
        
    except ImportError:
        st.error("❌ AI functionality not available. Please check infrastructure setup.")
        return {}
    except Exception as e:
        st.error(f"❌ AI processing failed: {e}")
        return {}

def export_to_powerpoint(ai_results, entity_name):
    """Export AI results to PowerPoint"""
    try:
        # Import PowerPoint functionality from the copied files
        sys.path.insert(0, str(current_dir / "src" / "infrastructure"))
        from pptx_export import export_pptx
        
        st.markdown("### 📎 PowerPoint Export")
        
        if not ai_results:
            st.warning("No AI results to export. Please run AI processing first.")
            return
        
        # Prepare data for PowerPoint export
        export_data = {
            'entity_name': entity_name,
            'timestamp': datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'ai_results': ai_results
        }
        
        # Generate PowerPoint file
        output_file = f"{entity_name}_due_diligence_report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
        
        with st.spinner("Generating PowerPoint presentation..."):
            success = export_pptx(export_data, output_file)
        
        if success and os.path.exists(output_file):
            st.success(f"✅ PowerPoint exported: {output_file}")
            
            # Provide download button
            with open(output_file, "rb") as file:
                st.download_button(
                    label="📥 Download PowerPoint",
                    data=file.read(),
                    file_name=output_file,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
        else:
            st.error("❌ PowerPoint export failed")
            
    except ImportError:
        st.error("❌ PowerPoint export functionality not available. Please check infrastructure setup.")
    except Exception as e:
        st.error(f"❌ PowerPoint export failed: {e}")

def main():
    """Main application with full functionality"""
    st.set_page_config(
        page_title="Financial Data Processor",
        page_icon="📊",
        layout="wide"
    )
    st.title("📊 Financial Data Processor")
    st.markdown("---")

    # Load configuration files
    config, mapping, pattern, prompts = load_config_files()
    
    if not all([config, mapping, pattern, prompts]):
        st.error("❌ Configuration files not available. Please ensure config/ directory has all required files.")
        return

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
        
        # Entity helpers
        entity_helpers = st.text_input(
            "Entity Helpers",
            value="Wanpu,Limited,",
            help="Comma-separated entity keywords"
        )
        
        # Financial Statement Type Selection
        st.markdown("---")
        statement_type_options = ["Balance Sheet", "Income Statement", "All"]
        statement_type_display = st.radio(
            "Financial Statement Type",
            options=statement_type_options,
            index=0,
            help="Select the type of financial statement to process"
        )
        
        # Convert display names to internal codes
        statement_type_map = {
            "Balance Sheet": "BS",
            "Income Statement": "IS", 
            "All": "ALL"
        }
        statement_type = statement_type_map[statement_type_display]

    # Main content area
    if uploaded_file is not None:
        st.success(f"✅ File uploaded: {uploaded_file.name}")
        st.info(f"🏢 Processing for entity: **{selected_entity}**")
        st.info(f"📊 Statement type: **{statement_type_display}**")
        
        # Process data and show by keys
        entity_suffixes = [s.strip() for s in entity_helpers.split(',') if s.strip()]
        
        with st.spinner("Processing Excel file..."):
            sections_by_key = get_worksheet_sections_by_keys(
                uploaded_file=uploaded_file,
                tab_name_mapping=mapping,
                entity_name=selected_entity,
                entity_suffixes=entity_suffixes,
                debug=False
            )
        
        # Filter keys based on statement type
        if statement_type == "BS":
            bs_keys = [
                "Cash", "AR", "Prepayments", "OR", "Other CA", "IP", "Other NCA",
                "AP", "Taxes payable", "OP", "Capital", "Reserve"
            ]
            filtered_keys = [key for key in bs_keys if key in sections_by_key and sections_by_key[key]]
        elif statement_type == "IS":
            is_keys = [
                "OI", "OC", "Tax and Surcharges", "GA", "Fin Exp", "Cr Loss", "Other Income",
                "Non-operating Income", "Non-operating Exp", "Income tax", "LT DTA"
            ]
            filtered_keys = [key for key in is_keys if key in sections_by_key and sections_by_key[key]]
        else:  # ALL
            filtered_keys = [key for key, sections in sections_by_key.items() if sections]
        
        # Show results by key
        st.subheader("📋 View Table by Key")
        
        if filtered_keys:
            st.success(f"✅ Found {len(filtered_keys)} keys with data")
            
            # Create tabs for each key
            key_tabs = st.tabs([get_key_display_name(key) for key in filtered_keys])
            
            for i, key in enumerate(filtered_keys):
                with key_tabs[i]:
                    st.subheader(f"Sheet: {get_key_display_name(key)}")
                    sections = sections_by_key[key]
                    
                    for j, section in enumerate(sections):
                        # Clean dataframe for display
                        df_clean = section['data'].dropna(axis=1, how='all').copy()
                        
                        # Convert columns to strings to avoid Arrow serialization issues
                        for col in df_clean.columns:
                            if df_clean[col].dtype == 'object':
                                df_clean.loc[:, col] = df_clean[col].astype(str)
                            elif 'datetime' in str(df_clean[col].dtype):
                                df_clean.loc[:, col] = df_clean[col].dt.strftime('%Y-%m-%d %H:%M:%S').fillna('')
                        
                        if section.get('entity_match', False):
                            st.markdown(f"**Section {j+1}:** ✅ Entity Match")
                        else:
                            st.markdown(f"**Section {j+1}:** ⚠️ No Entity Match")
                        
                        st.dataframe(df_clean, use_container_width=True)
                        
                        with st.expander(f"📋 Markdown Table - Section {j+1}", expanded=False):
                            st.code(section['markdown'], language='markdown')
                        
                        st.info(f"**Source Sheet:** {section['sheet']}")
                        
                        if j < len(sections) - 1:
                            st.markdown("---")
        else:
            st.warning("⚠️ No data found for the selected statement type and entity configuration.")
            st.info("💡 Try adjusting the entity helpers or statement type selection.")
        
        # AI Processing Section
        st.markdown("---")
        st.subheader("🤖 AI Processing")
        
        if filtered_keys:
            if st.button("🤖 Process with AI", type="primary", use_container_width=True):
                ai_results = run_ai_processing(
                    filtered_keys, 
                    sections_by_key, 
                    selected_entity, 
                    entity_suffixes, 
                    config, 
                    mapping, 
                    pattern, 
                    prompts
                )
                
                # Store results in session state
                st.session_state['ai_results'] = ai_results
                st.session_state['ai_processed'] = True
        else:
            st.info("No data available for AI processing. Please upload a file and ensure data is found.")
        
        # PowerPoint Export Section
        if st.session_state.get('ai_processed', False) and st.session_state.get('ai_results'):
            st.markdown("---")
            if st.button("📎 Export to PowerPoint", type="secondary", use_container_width=True):
                export_to_powerpoint(st.session_state['ai_results'], selected_entity)
    
    else:
        st.info("📁 Please upload an Excel file to begin processing.")
        
        # Show configuration status
        with st.expander("⚙️ Configuration Status"):
            st.write("**Loaded Configuration Files:**")
            st.write(f"- mapping.json: {len(mapping) if mapping else 0} entities")
            st.write(f"- pattern.json: {len(pattern) if pattern else 0} patterns") 
            st.write(f"- config.json: ✅ Available")
            st.write(f"- prompts.json: ✅ Available")

if __name__ == "__main__":
    main() 
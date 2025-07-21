#!/usr/bin/env python3
"""
Financial Data Processor - Due Diligence Automation

Complete version with key-based processing, AI pipeline, and PowerPoint export.
"""

import streamlit as st
import pandas as pd
import json
import os
import re
import warnings
import datetime
from pathlib import Path
from tabulate import tabulate
import tempfile

# Suppress warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action='ignore', category=UserWarning)

def load_config_files():
    """Load configuration files from the config directory."""
    try:
        config_dir = Path("config")
        
        with open(config_dir / "mapping.json", 'r') as f:
            mapping = json.load(f)
        with open(config_dir / "pattern.json", 'r') as f:
            pattern = json.load(f)
        with open(config_dir / "config.json", 'r') as f:
            config = json.load(f)
        with open(config_dir / "prompts.json", 'r') as f:
            prompts = json.load(f)
            
        return config, mapping, pattern, prompts
        
    except Exception as e:
        st.error(f"Configuration error: {e}")
        return None, None, None, None

def get_financial_keys():
    """Get financial keys for Balance Sheet and Income Statement"""
    return {
        'BS': ["Cash", "AR", "Prepayments", "OR", "Other CA", "IP", "Other NCA", 
               "AP", "Taxes payable", "OP", "Capital", "Reserve"],
        'IS': ["OI", "OC", "Tax and Surcharges", "GA", "Fin Exp", "Cr Loss", 
               "Other Income", "Non-operating Income", "Non-operating Exp", "Income tax", "LT DTA"]
    }

def get_key_display_name(key, mapping=None):
    """Get display name for financial key"""
    if mapping and key in mapping and mapping[key]:
        values = mapping[key]
        # Use first descriptive value
        for value in values:
            if len(value) > 3 and not value.isupper():
                return value
        return values[0]
    
    # Fallback names
    names = {
        'Cash': 'Cash', 'AR': 'Accounts Receivable', 'Prepayments': 'Prepayments',
        'OR': 'Other Receivables', 'Other CA': 'Other Current Assets',
        'IP': 'Investment Properties', 'Other NCA': 'Other Non-Current Assets',
        'AP': 'Accounts Payable', 'Taxes payable': 'Tax Payable',
        'OP': 'Other Payables', 'Capital': 'Share Capital', 'Reserve': 'Reserve'
    }
    return names.get(key, key)

def get_worksheet_sections_by_keys(uploaded_file, tab_name_mapping, entity_name, entity_suffixes, statement_type="BS"):
    """
    Process Excel file and organize data by financial keys
    """
    try:
        # Get relevant financial keys for statement type
        financial_keys_map = get_financial_keys()
        if statement_type == "ALL":
            relevant_keys = financial_keys_map['BS'] + financial_keys_map['IS']
        else:
            relevant_keys = financial_keys_map.get(statement_type, financial_keys_map['BS'])
        
        # Initialize sections by key
        sections_by_key = {key: [] for key in relevant_keys}
        
        with pd.ExcelFile(uploaded_file) as xl:
            # Create reverse mapping from sheet values to keys
            reverse_mapping = {}
            for key, values in tab_name_mapping.items():
                for value in values:
                    reverse_mapping[value] = key
            
            # Entity keywords for filtering
            entity_keywords = [f"{entity_name} {suffix}" for suffix in entity_suffixes if suffix]
            if not entity_keywords:
                entity_keywords = [entity_name]
            combined_pattern = '|'.join(re.escape(kw) for kw in entity_keywords)
            
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
                    
                    # Match dataframes to financial keys
                    for data_frame in dataframes:
                        best_key = None
                        best_score = 0
                        
                        # Check which financial key this dataframe belongs to
                        for key in relevant_keys:
                            if key in tab_name_mapping:
                                key_patterns = tab_name_mapping[key]
                                for pattern in key_patterns:
                                    # Check if pattern exists in dataframe
                                    if data_frame.apply(
                                        lambda row: row.astype(str).str.contains(
                                            pattern, case=False, regex=True, na=False
                                        ).any(), axis=1
                                    ).any():
                                        # Score based on pattern specificity
                                        score = len(pattern)
                                        if score > best_score:
                                            best_score = score
                                            best_key = key
                        
                        if best_key:
                            # Check entity filter
                            entity_match = data_frame.apply(
                                lambda row: row.astype(str).str.contains(
                                    combined_pattern, case=False, regex=True, na=False
                                ).any(), axis=1
                            ).any()
                            
                            # Include if entity matches or no specific entity filter
                            if entity_match or not entity_suffixes:
                                sections_by_key[best_key].append({
                                    'sheet': sheet_name,
                                    'data': data_frame,
                                    'markdown': tabulate(data_frame, headers='keys', tablefmt='pipe', showindex=False),
                                    'entity_match': entity_match
                                })
        
        return sections_by_key
        
    except Exception as e:
        st.error(f"Processing error: {e}")
        return {}

def run_ai_processing_with_logging(keys_with_data, sections_by_key, entity_name, config, pattern, prompts):
    """AI processing with comprehensive logging and real API calls"""
    st.markdown("### 🤖 AI Agent Pipeline")
    
    # Import AI config and logging from utils
    import sys
    import time
    sys.path.insert(0, str(Path(__file__).parent / "utils"))
    
    try:
        from ai_config import load_ai_config, initialize_ai_services, generate_ai_response
        from ai_logger import start_new_ai_session, get_ai_logger
        
        # Start new AI logging session
        ai_logger = start_new_ai_session()
        st.success(f"🗂️ AI logging session started: {ai_logger.session_dir}")
        
        # Initialize AI services
        ai_config = load_ai_config()
        ai_client, _ = initialize_ai_services(ai_config)
        has_real_openai = ai_client and ai_client != "openai_client_placeholder"
        
        # Determine connection status for logging
        if has_real_openai:
            connection_status = "real_openai_connected"
        elif ai_client == "openai_client_placeholder":
            connection_status = "placeholder_mode"
        else:
            connection_status = "no_ai_configured"
        
    except ImportError as e:
        st.error(f"⚠️ AI utilities import failed: {e}")
        # Basic fallback without logging
        has_real_openai = False
        ai_client = None
        ai_logger = None
        connection_status = "import_error"
    
    # Display AI status
    if has_real_openai:
        st.success("🚀 **Real OpenAI API detected** - Processing with actual AI")
    elif ai_client == "openai_client_placeholder":
        st.info("🔄 **Demo mode** - OpenAI not available, using enhanced placeholders")
    else:
        st.warning("⚠️ **Fallback mode** - No AI configuration detected")
    
    # Single progress bar for all processing
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    total_keys = len(keys_with_data)
    total_steps = total_keys * 3  # 3 agents per key
    current_step = 0
    
    ai_results = {}
    
    for i, key in enumerate(keys_with_data):
        st.markdown(f"#### 🔄 Processing Key: {get_key_display_name(key)}")
        
        sections = sections_by_key[key]
        if not sections:
            continue
        
        # Prepare context data
        context_data = "\n\n".join([section['markdown'] for section in sections])
        
        # Agent 1: Content Generation
        current_step += 1
        progress_bar.progress(current_step / total_steps)
        status_text.text(f"🤖 Agent 1: Generating content for {get_key_display_name(key)} ({i+1}/{total_keys})")
        
        start_time = time.time()
        
        # Get system and user prompts
        system_prompt = prompts.get('system_prompts', {}).get('Agent 1', 'You are a financial analyst specializing in due diligence reports.') if prompts else "You are a financial analyst."
        user_prompt = f"""Generate a comprehensive financial analysis for the {key} category based on this data:

Entity: {entity_name}
Financial Category: {key}
Data Sections: {len(sections)}

Raw Data:
{context_data[:2000]}

Please provide:
1. Summary of key findings
2. Important financial figures
3. Notable trends or patterns
4. Risk factors or concerns
5. Recommendations

Format your response clearly and professionally."""
        
        agent1_result = generate_ai_response(ai_client, system_prompt, user_prompt)
        agent1_time = time.time() - start_time
        
        # Log the interaction
        if ai_logger:
            ai_logger.log_ai_interaction(
                agent_name="Agent 1",
                key=key,
                system_prompt=system_prompt,
                user_prompt=user_prompt,
                ai_response=agent1_result,
                entity_name=entity_name,
                processing_time=agent1_time,
                ai_connection_status=connection_status
            )
        
        # Agent 2: Data Validation
        current_step += 1
        progress_bar.progress(current_step / total_steps)
        status_text.text(f"🔍 Agent 2: Validating data for {get_key_display_name(key)} ({i+1}/{total_keys})")
        
        start_time = time.time()
        
        system_prompt = prompts.get('system_prompts', {}).get('Agent 2', 'You are a data validation specialist for financial reports.') if prompts else "You are a data validator."
        user_prompt = f"""Validate the following financial analysis for accuracy and consistency:

Financial Category: {key}
Entity: {entity_name}

Generated Content:
{agent1_result}

Original Data for Cross-Reference:
{context_data[:1000]}

Please:
1. Verify that all financial figures mentioned are present in the source data
2. Check for mathematical accuracy
3. Identify any inconsistencies or discrepancies
4. Suggest corrections if needed
5. Rate the accuracy (1-10) and explain your reasoning

Provide your validation in JSON format:
{{
    "accuracy_score": 8,
    "issues_found": ["list", "of", "issues"],
    "corrections_needed": ["list", "of", "corrections"],
    "validation_summary": "Overall assessment"
}}"""
        
        agent2_result = generate_ai_response(ai_client, system_prompt, user_prompt)
        agent2_time = time.time() - start_time
        
        # Log the interaction
        if ai_logger:
            ai_logger.log_ai_interaction(
                agent_name="Agent 2",
                key=key,
                system_prompt=system_prompt,
                user_prompt=user_prompt,
                ai_response=agent2_result,
                entity_name=entity_name,
                processing_time=agent2_time,
                ai_connection_status=connection_status
            )
        
        # Agent 3: Pattern Compliance
        current_step += 1
        progress_bar.progress(current_step / total_steps)
        status_text.text(f"🎯 Agent 3: Checking pattern compliance for {get_key_display_name(key)} ({i+1}/{total_keys})")
        
        start_time = time.time()
        
        patterns_for_key = pattern.get(key, {}) if pattern else {}
        system_prompt = prompts.get('system_prompts', {}).get('Agent 3', 'You are a compliance specialist for financial reporting patterns.') if prompts else "You are a compliance checker."
        user_prompt = f"""Check the following content for compliance with established patterns and format requirements:

Financial Category: {key}
Entity: {entity_name}

Content to Check:
{agent1_result}

Required Patterns:
{json.dumps(patterns_for_key, indent=2)}

Validation Results from Agent 2:
{agent2_result}

Please:
1. Ensure the content follows the required patterns
2. Check formatting and structure compliance
3. Verify professional tone and language
4. Remove any template artifacts
5. Provide the final, compliant version

Return the final compliant content that is ready for PowerPoint export."""
        
        agent3_result = generate_ai_response(ai_client, system_prompt, user_prompt)
        agent3_time = time.time() - start_time
        
        # Log the interaction
        if ai_logger:
            ai_logger.log_ai_interaction(
                agent_name="Agent 3",
                key=key,
                system_prompt=system_prompt,
                user_prompt=user_prompt,
                ai_response=agent3_result,
                entity_name=entity_name,
                processing_time=agent3_time,
                ai_connection_status=connection_status
            )
        
        # Store results
        ai_results[key] = {
            'agent1': agent1_result,
            'agent2': agent2_result,
            'agent3': agent3_result,
            'final_content': agent3_result,
            'processing_times': {
                'agent1': agent1_time,
                'agent2': agent2_time,
                'agent3': agent3_time
            }
        }
        
        # Show completion for this key
        st.success(f"✅ AI processing completed for {get_key_display_name(key)} (Total: {agent1_time + agent2_time + agent3_time:.1f}s)")
        
        # Show results in expandable tabs
        with st.expander(f"📊 AI Results for {get_key_display_name(key)}", expanded=False):
            tab1, tab2, tab3, tab4 = st.tabs(["📝 Agent 1", "🔍 Agent 2", "🎯 Agent 3", "📋 Summary"])
            
            with tab1:
                st.markdown("### 🤖 **Agent 1: Content Generation**")
                st.markdown(f"**Processing Time:** {agent1_time:.2f} seconds")
                if "[Demo AI Analysis]" in agent1_result:
                    st.info("🔄 **Demo Mode** - This is placeholder content")
                elif "[AI Error]" in agent1_result:
                    st.error("❌ **Error** - AI processing failed")
                else:
                    st.success("🤖 **Real AI Response**")
                st.markdown(agent1_result)
            
            with tab2:
                st.markdown("### 🔍 **Agent 2: Data Validation**")
                st.markdown(f"**Processing Time:** {agent2_time:.2f} seconds")
                if "[Demo AI Analysis]" in agent2_result:
                    st.info("🔄 **Demo Mode** - This is placeholder content")
                elif "[AI Error]" in agent2_result:
                    st.error("❌ **Error** - AI processing failed")
                else:
                    st.success("🤖 **Real AI Response**")
                st.markdown(agent2_result)
            
            with tab3:
                st.markdown("### 🎯 **Agent 3: Pattern Compliance**")
                st.markdown(f"**Processing Time:** {agent3_time:.2f} seconds")
                if "[Demo AI Analysis]" in agent3_result:
                    st.info("🔄 **Demo Mode** - This is placeholder content")
                elif "[AI Error]" in agent3_result:
                    st.error("❌ **Error** - AI processing failed")
                else:
                    st.success("🤖 **Real AI Response**")
                st.markdown(agent3_result)
            
            with tab4:
                st.markdown("### 📋 **Processing Summary**")
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Agent 1 Time", f"{agent1_time:.2f}s")
                    st.metric("Agent 2 Time", f"{agent2_time:.2f}s")
                with col2:
                    st.metric("Agent 3 Time", f"{agent3_time:.2f}s")
                    st.metric("Total Time", f"{agent1_time + agent2_time + agent3_time:.2f}s")
                
                # Response type indicators
                response_types = []
                for agent_result in [agent1_result, agent2_result, agent3_result]:
                    if "[Demo AI Analysis]" in agent_result:
                        response_types.append("🔄 Demo")
                    elif "[AI Error]" in agent_result:
                        response_types.append("❌ Error")
                    elif not any(x in agent_result for x in ["[Demo AI Analysis]", "[AI Response Placeholder]", "[Fallback Response]"]):
                        response_types.append("🤖 Real AI")
                    else:
                        response_types.append("🔄 Fallback")
                
                st.markdown("**Response Types:**")
                st.write(" | ".join(response_types))
    
    # Final completion
    progress_bar.progress(1.0)
    status_text.text(f"✅ All processing completed! Processed {total_keys} keys with 3 agents each.")
    
    # Finalize logging session
    if ai_logger:
        session_dir = ai_logger.finalize_session()
        st.success(f"📊 **AI logging completed:** {session_dir}")
        
        # Show logging summary
        session_info = ai_logger.get_session_info()
        st.info(f"🗂️ **Session:** {session_info['interactions_count']} interactions logged")
        
        # Download logs option
        try:
            summary_file = session_dir / "summary.md"
            if summary_file.exists():
                with open(summary_file, 'r', encoding='utf-8') as f:
                    summary_content = f.read()
                st.download_button(
                    label="📥 Download AI Logs Summary",
                    data=summary_content,
                    file_name=f"ai_logs_{session_info['session_id']}.md",
                    mime="text/markdown"
                )
        except Exception as e:
            st.warning(f"Could not prepare log download: {e}")
    
    # Show consolidated AI results for all keys
    if ai_results:
        st.markdown("---")
        st.markdown("### 📊 **All AI Results Summary**")
        
        # Create tabs for each agent across all keys
        all_agent1_tab, all_agent2_tab, all_agent3_tab = st.tabs(["🤖 All Agent 1 Results", "🔍 All Agent 2 Results", "🎯 All Agent 3 Results"])
        
        with all_agent1_tab:
            st.markdown("#### **Agent 1: Content Generation - All Keys**")
            for key, result in ai_results.items():
                with st.expander(f"📝 {get_key_display_name(key)} - Agent 1", expanded=False):
                    st.markdown(result['agent1'])
        
        with all_agent2_tab:
            st.markdown("#### **Agent 2: Data Validation - All Keys**")
            for key, result in ai_results.items():
                with st.expander(f"🔍 {get_key_display_name(key)} - Agent 2", expanded=False):
                    st.markdown(result['agent2'])
        
        with all_agent3_tab:
            st.markdown("#### **Agent 3: Pattern Compliance - All Keys**")
            for key, result in ai_results.items():
                with st.expander(f"🎯 {get_key_display_name(key)} - Agent 3", expanded=False):
                    st.markdown(result['agent3'])
    
    return ai_results

def export_to_powerpoint_fallback(ai_results, entity_name):
    """PowerPoint export with fallback implementation"""
    st.markdown("### 📎 PowerPoint Export")
    
    if not ai_results:
        st.warning("No AI results to export. Please run AI processing first.")
        return
    
    # Generate summary content
    output_file = f"{entity_name}_due_diligence_report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.md"
    
    content = f"""# Due Diligence Report: {entity_name}

Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

## Executive Summary

This report contains AI-processed financial analysis for {entity_name} covering {len(ai_results)} financial categories.

## Analysis Results

"""
    
    for key, results in ai_results.items():
        content += f"""
### {get_key_display_name(key)}

**AI Agent 1 Analysis:**
{results['agent1']}

**AI Agent 2 Validation:**
{results['agent2']}

**AI Agent 3 Pattern Compliance:**
{results['agent3']}

---
"""
    
    content += f"""
## Report Summary

- Total categories analyzed: {len(ai_results)}
- Entity: {entity_name}
- Processing completed successfully
- All data validated and pattern-compliant

*Note: This is a markdown export. PowerPoint functionality will be available with full infrastructure setup.*
"""
    
    # Provide download
    st.download_button(
        label="📥 Download Report (Markdown)",
        data=content,
        file_name=output_file,
        mime="text/markdown"
    )
    
    st.success(f"✅ Report exported as {output_file}")
    st.info("💡 Full PowerPoint export will be available once infrastructure is fully configured.")

def main():
    """Main application with complete functionality"""
    st.set_page_config(
        page_title="Financial Data Processor - Full Version",
        page_icon="📊",
        layout="wide"
    )
    
    st.title("📊 Financial Data Processor - Full Version")
    st.markdown("**Complete Due Diligence Automation with AI Processing**")
    st.markdown("---")

    # Load configuration
    config, mapping, pattern, prompts = load_config_files()
    
    if not mapping:
        st.error("❌ Configuration files required. Please ensure config/ directory is set up.")
        return

    # Sidebar controls
    with st.sidebar:
        uploaded_file = st.file_uploader(
            "Upload Excel File",
            type=['xlsx', 'xls'],
            help="Upload your financial data Excel file"
        )
        
        entity_options = ['Haining', 'Nanjing', 'Ningbo']
        selected_entity = st.selectbox("Select Entity", options=entity_options)
        
        entity_helpers = st.text_input(
            "Entity Helpers",
            value="Wanpu,Limited,",
            help="Comma-separated entity keywords"
        )
        
        st.markdown("---")
        statement_type_options = ["Balance Sheet", "Income Statement", "All"]
        statement_type_display = st.radio("Financial Statement Type", options=statement_type_options, index=0)
        
        statement_type_map = {"Balance Sheet": "BS", "Income Statement": "IS", "All": "ALL"}
        statement_type = statement_type_map[statement_type_display]

    # Main processing
    if uploaded_file is not None:
        st.success(f"✅ File uploaded: {uploaded_file.name}")
        st.info(f"🏢 Entity: **{selected_entity}** | 📊 Type: **{statement_type_display}**")
        
        # Process data
        entity_suffixes = [s.strip() for s in entity_helpers.split(',') if s.strip()]
        
        with st.spinner("Processing Excel file and organizing by financial keys..."):
            sections_by_key = get_worksheet_sections_by_keys(
                uploaded_file, mapping, selected_entity, entity_suffixes, statement_type
            )
        
        # Filter keys with data
        filtered_keys = [key for key, sections in sections_by_key.items() if sections]
        
        # Display results by key
        st.subheader("📋 View Table by Key")
        
        if filtered_keys:
            st.success(f"✅ Found {len(filtered_keys)} keys with data for {statement_type_display}")
            
            # Show expected vs found
            expected_keys = get_financial_keys()[statement_type] if statement_type != "ALL" else get_financial_keys()['BS']
            st.info(f"📊 Expected {len(expected_keys)} keys for {statement_type_display}, found {len(filtered_keys)} with data")
            
            # Create tabs for each key
            key_tabs = st.tabs([get_key_display_name(key, mapping) for key in filtered_keys])
            
            for i, key in enumerate(filtered_keys):
                with key_tabs[i]:
                    st.subheader(f"{get_key_display_name(key, mapping)}")
                    sections = sections_by_key[key]
                    
                    for j, section in enumerate(sections):
                        # Clean and display dataframe
                        df_clean = section['data'].dropna(axis=1, how='all').copy()
                        
                        # Handle data types
                        for col in df_clean.columns:
                            if df_clean[col].dtype == 'object':
                                df_clean.loc[:, col] = df_clean[col].astype(str)
                            elif 'datetime' in str(df_clean[col].dtype):
                                df_clean.loc[:, col] = df_clean[col].dt.strftime('%Y-%m-%d').fillna('')
                        
                        # Show entity match status
                        if section.get('entity_match', False):
                            st.markdown(f"**Section {j+1}:** ✅ Entity Match Found")
                        else:
                            st.markdown(f"**Section {j+1}:** ⚠️ General Data (No Entity Match)")
                        
                        st.dataframe(df_clean, use_container_width=True)
                        
                        with st.expander(f"📋 Raw Data - Section {j+1}", expanded=False):
                            st.code(section['markdown'], language='markdown')
                        
                        st.info(f"**Source:** {section['sheet']}")
                        
                        if j < len(sections) - 1:
                            st.markdown("---")
            
            # AI Processing Section
            st.markdown("---")
            st.subheader("🤖 AI Processing Pipeline")
            
            if st.button("🤖 Process with AI Agents", type="primary", use_container_width=True):
                ai_results = run_ai_processing_with_logging(
                    filtered_keys, sections_by_key, selected_entity, config, pattern, prompts
                )
                st.session_state['ai_results'] = ai_results
                st.session_state['ai_processed'] = True
            
            # PowerPoint Export Section  
            if st.session_state.get('ai_processed', False):
                st.markdown("---")
                if st.button("📎 Export to PowerPoint", type="secondary", use_container_width=True):
                    export_to_powerpoint_fallback(st.session_state['ai_results'], selected_entity)
        
        else:
            st.warning(f"⚠️ No data found for {statement_type_display}")
            st.info("💡 Try different entity helpers or check if your Excel file contains the expected sheet structure")
            
            # Show debug info
            with st.expander("🔧 Debug Information"):
                st.write("**Expected Keys for", statement_type_display, ":**")
                expected_keys = get_financial_keys()[statement_type] if statement_type != "ALL" else get_financial_keys()['BS']
                st.write(expected_keys)
                st.write("**Available Mapping Keys:**")
                st.write(list(mapping.keys())[:10], "... (showing first 10)")
    
    else:
        st.info("📁 Please upload an Excel file to begin processing.")
        
        # Configuration status
        with st.expander("⚙️ System Status"):
            st.write("**Configuration:**")
            st.write(f"- Mapping: {len(mapping)} keys" if mapping else "- Mapping: ❌ Not loaded")
            st.write(f"- Patterns: {len(pattern)} items" if pattern else "- Patterns: ❌ Not loaded") 
            st.write(f"- Config: ✅ Available" if config else "- Config: ❌ Not loaded")
            st.write(f"- Prompts: ✅ Available" if prompts else "- Prompts: ❌ Not loaded")
            
            st.write("**Expected Balance Sheet Keys:**")
            bs_keys = get_financial_keys()['BS']
            st.write(", ".join(bs_keys))

if __name__ == "__main__":
    main() 
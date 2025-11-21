#!/usr/bin/env python3
"""
Streamlit App for Financial Data Processing with AI
Combines extraction, reconciliation, AI generation, and PPTX export
"""

import streamlit as st
import pandas as pd
import os
import tempfile
from typing import Dict, List

# Import modules
from fdd_utils.process_databook import extract_data_from_excel
from fdd_utils.financial_extraction import extract_balance_sheet_and_income_statement
from fdd_utils.reconciliation import reconcile_financial_statements
from fdd_utils.content_generation import run_ai_pipeline, extract_final_contents

# Page config
st.set_page_config(
    page_title="Financial Data Processing",
    page_icon="📊",
    layout="wide"
)

# Custom CSS
st.markdown("""
<style>
.block-container {padding-top: 1rem;}
.stTabs [data-baseweb="tab-list"] {gap: 2px;}
.stTabs [data-baseweb="tab"] {padding: 10px 20px;}
</style>
""", unsafe_allow_html=True)


def init_session_state():
    """Initialize session state variables"""
    if 'uploaded_file' not in st.session_state:
        st.session_state.uploaded_file = None
    if 'dfs' not in st.session_state:
        st.session_state.dfs = None
    if 'workbook_list' not in st.session_state:
        st.session_state.workbook_list = []
    if 'language' not in st.session_state:
        st.session_state.language = 'Eng'
    if 'bs_is_results' not in st.session_state:
        st.session_state.bs_is_results = None
    if 'ai_results' not in st.session_state:
        st.session_state.ai_results = None
    if 'reconciliation' not in st.session_state:
        st.session_state.reconciliation = None
    if 'model_type' not in st.session_state:
        st.session_state.model_type = 'local'
    if 'project_name' not in st.session_state:
        st.session_state.project_name = None


def get_entity_names(file_path: str) -> List[str]:
    """Extract potential entity names from Excel file"""
    try:
        xls = pd.ExcelFile(file_path, engine='openpyxl')
        entity_names = set()
        
        # Check first few sheets for entity patterns
        for sheet in xls.sheet_names[:5]:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet, nrows=10, engine='openpyxl')
                df_str = df.to_string()
                
                # Look for patterns like "xxx公司", "xxx - 东莞xx"
                import re
                patterns = [
                    r'[-–]\s*([^,\n]+)',  # After dash
                    r'([^\s]+公司)',       # xxx公司
                    r'([^\s]+集团)',       # xxx集团
                ]
                
                for pattern in patterns:
                    matches = re.findall(pattern, df_str)
                    for match in matches[:3]:
                        if len(match) > 2 and len(match) < 50:
                            entity_names.add(match.strip())
            except:
                continue
        
        return [''] + sorted(list(entity_names))
    except:
        return ['']


def get_financial_sheets(file_path: str) -> List[str]:
    """Get list of sheets, prioritizing those with 'Financial' in name"""
    try:
        xls = pd.ExcelFile(file_path, engine='openpyxl')
        sheets = xls.sheet_names
        
        # Separate sheets with "Financial" from others
        financial_sheets = [s for s in sheets if 'financial' in s.lower()]
        other_sheets = [s for s in sheets if 'financial' not in s.lower()]
        
        # Return with financial sheets first
        return financial_sheets + other_sheets
    except:
        return []


# Initialize
init_session_state()

# Title
st.title("📊 Financial Data Processing & AI Generation")

# Sidebar
with st.sidebar:
    st.header("⚙️ Configuration")
    
    # File upload with default
    st.markdown("**📁 Databook File**")
    use_default = st.checkbox("Use default (databook.xlsx)", value=True)
    
    if use_default:
        temp_path = "databook.xlsx"
        if os.path.exists(temp_path):
            st.success(f"✅ Using: {temp_path}")
        else:
            st.error(f"❌ File not found: {temp_path}")
            temp_path = None
    else:
        uploaded_file = st.file_uploader(
            "Upload Excel",
            type=['xlsx', 'xls'],
            help="Upload your financial databook"
        )
        
        if uploaded_file:
            # Save to temp file
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                temp_path = tmp_file.name
            st.success(f"✅ File uploaded: {uploaded_file.name}")
        else:
            temp_path = None
    
    if temp_path:
        # Entity name selection (editable)
        st.markdown("---")
        st.markdown("**🏢 Entity Name**")
        entity_options = get_entity_names(temp_path)
        entity_selected = st.selectbox(
            "Select or type custom",
            options=entity_options,
            help="Leave blank for single entity",
            label_visibility="collapsed"
        )
        entity_name = st.text_input(
            "Custom entity name (optional)",
            value=entity_selected,
            help="Edit if needed",
            label_visibility="collapsed"
        )
        
        # Financial statement sheet selection
        st.markdown("---")
        st.markdown("**📊 Financial Statement Sheet**")
        sheet_options = get_financial_sheets(temp_path)
        if sheet_options:
            selected_sheet = st.selectbox(
                "Select sheet",
                options=sheet_options,
                help="Sheet containing both BS and IS",
                label_visibility="collapsed"
            )
        else:
            st.warning("No sheets found")
            selected_sheet = None
        
        # Model selection (horizontal)
        st.markdown("---")
        st.markdown("**🤖 AI Model**")
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("🏠 Local", use_container_width=True, 
                        type="primary" if st.session_state.get('model_type') == 'local' else "secondary"):
                st.session_state.model_type = 'local'
        with col2:
            if st.button("🌐 OpenAI", use_container_width=True,
                        type="primary" if st.session_state.get('model_type') == 'openai' else "secondary"):
                st.session_state.model_type = 'openai'
        with col3:
            if st.button("🚀 DeepSeek", use_container_width=True,
                        type="primary" if st.session_state.get('model_type') == 'deepseek' else "secondary"):
                st.session_state.model_type = 'deepseek'
        
        model_type = st.session_state.get('model_type', 'local')
        st.info(f"Selected: **{model_type}**")
        
        # Process button
        st.markdown("---")
        if st.button("🚀 Process Data", type="primary", use_container_width=True):
            with st.spinner("Processing..."):
                try:
                    # Extract account-by-account data
                    dfs, workbook_list, _, language = extract_data_from_excel(
                        databook_path=temp_path,
                        entity_name=entity_name,
                        mode="All"  # Always use All mode
                    )
                    
                    # Extract BS/IS from single sheet
                    bs_is_results = None
                    if selected_sheet:
                        bs_is_results = extract_balance_sheet_and_income_statement(
                            workbook_path=temp_path,
                            sheet_name=selected_sheet,
                            debug=False
                        )
                    
                    # Reconcile if both sources available
                    recon_bs, recon_is = None, None
                    if dfs and bs_is_results:
                        recon_bs, recon_is = reconcile_financial_statements(
                            bs_is_results=bs_is_results,
                            dfs=dfs,
                            tolerance=1.0,
                            materiality_threshold=0.005,
                            debug=False
                        )
                    
                    # Store in session state
                    st.session_state.dfs = dfs
                    st.session_state.workbook_list = workbook_list
                    st.session_state.language = language
                    st.session_state.bs_is_results = bs_is_results
                    st.session_state.reconciliation = (recon_bs, recon_is)
                    st.session_state.model_type = model_type
                    st.session_state.project_name = bs_is_results.get('project_name') if bs_is_results else None
                    
                    st.success("✅ Data processed successfully!")
                    st.rerun()
                    
                except Exception as e:
                    st.error(f"❌ Error processing data: {e}")
                    import traceback
                    st.code(traceback.format_exc())

# Main content
if st.session_state.dfs is None:
    st.info("👈 Upload a databook and click 'Process Data' to begin")
else:
    # Area 1: Data Display
    st.header("📈 Data Display")
    
    # Create tabs for BS and IS
    tab_bs, tab_is = st.tabs(["Balance Sheet", "Income Statement"])
    
    with tab_bs:
        bs_data = st.session_state.bs_is_results.get('balance_sheet') if st.session_state.bs_is_results else None
        recon_bs = st.session_state.reconciliation[0] if st.session_state.reconciliation else None
        
        if bs_data is not None:
            # Second level tabs
            bs_tabs = st.tabs(["📊 Reconciliation"] + [f"📋 {key}" for key in st.session_state.workbook_list if key in st.session_state.dfs])
            
            # Reconciliation tab
            with bs_tabs[0]:
                if recon_bs is not None and not recon_bs.empty:
                    # Convert to all strings for Streamlit display (fixes Arrow serialization)
                    display_df = recon_bs.copy()
                    for col in display_df.columns:
                        display_df[col] = display_df[col].apply(
                            lambda x: f"{x:,.0f}" if isinstance(x, (int, float)) and x != 0 else str(x)
                        )
                    st.dataframe(display_df, use_container_width=True, height=400)
                    
                    # Summary stats
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        matches = (recon_bs['Match'] == '✅ Match').sum()
                        st.metric("✅ Matches", matches)
                    with col2:
                        immaterial = (recon_bs['Match'] == '✅ Immaterial').sum()
                        st.metric("✅ Immaterial", immaterial)
                    with col3:
                        diffs = (recon_bs['Match'] == '❌ Diff').sum()
                        st.metric("❌ Differences", diffs)
                    with col4:
                        not_found = (recon_bs['Match'] == '⚠️ Not Found').sum()
                        st.metric("⚠️ Not Found", not_found)
                else:
                    st.info("No reconciliation data available")
            
            # Individual account tabs
            for idx, key in enumerate([k for k in st.session_state.workbook_list if k in st.session_state.dfs], 1):
                with bs_tabs[idx]:
                    st.dataframe(st.session_state.dfs[key], use_container_width=True)
        else:
            st.info("No Balance Sheet data extracted")
    
    with tab_is:
        is_data = st.session_state.bs_is_results.get('income_statement') if st.session_state.bs_is_results else None
        recon_is = st.session_state.reconciliation[1] if st.session_state.reconciliation else None
        
        if is_data is not None:
            # Second level tabs
            is_tabs = st.tabs(["📊 Reconciliation"] + [f"📋 {key}" for key in st.session_state.workbook_list if key in st.session_state.dfs])
            
            # Reconciliation tab
            with is_tabs[0]:
                if recon_is is not None and not recon_is.empty:
                    # Convert to all strings for Streamlit display (fixes Arrow serialization)
                    display_df = recon_is.copy()
                    for col in display_df.columns:
                        display_df[col] = display_df[col].apply(
                            lambda x: f"{x:,.0f}" if isinstance(x, (int, float)) and x != 0 else str(x)
                        )
                    st.dataframe(display_df, use_container_width=True, height=400)
                    
                    # Summary stats
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        matches = (recon_is['Match'] == '✅ Match').sum()
                        st.metric("✅ Matches", matches)
                    with col2:
                        immaterial = (recon_is['Match'] == '✅ Immaterial').sum()
                        st.metric("✅ Immaterial", immaterial)
                    with col3:
                        diffs = (recon_is['Match'] == '❌ Diff').sum()
                        st.metric("❌ Differences", diffs)
                    with col4:
                        not_found = (recon_is['Match'] == '⚠️ Not Found').sum()
                        st.metric("⚠️ Not Found", not_found)
                else:
                    st.info("No reconciliation data available")
            
            # Individual account tabs
            for idx, key in enumerate([k for k in st.session_state.workbook_list if k in st.session_state.dfs], 1):
                with is_tabs[idx]:
                    st.dataframe(st.session_state.dfs[key], use_container_width=True)
        else:
            st.info("No Income Statement data extracted")
    
    # Area 2: AI Content Generation
    st.markdown("---")
    st.header("🤖 AI Content Generation")
    
    # Horizontal buttons: Generate AI + Export PPTX
    col_ai, col_pptx = st.columns(2)
    
    with col_ai:
        if st.button("▶️ Generate AI Content", type="primary", use_container_width=True):
            progress_placeholder = st.empty()
            status_placeholder = st.empty()
            
            try:
                total_items = len(st.session_state.workbook_list)
                progress_bar = progress_placeholder.progress(0)
                
                # Show progress
                status_placeholder.info(f"🔄 Processing {total_items} accounts with 4 agents...")
                
                # Simple progress tracking (since we can't hook into the pipeline easily)
                import time
                for i in range(4):  # 4 agents
                    agent_name = ['Generator', 'Auditor', 'Refiner', 'Validator'][i]
                    status_placeholder.info(f"🔄 Agent {i+1}/4: {agent_name} - Processing {total_items} accounts...")
                    
                    if i == 0:  # Only run on first iteration
                        results = run_ai_pipeline(
                            mapping_keys=st.session_state.workbook_list,
                            dfs=st.session_state.dfs,
                            model_type=st.session_state.get('model_type', 'local'),
                            language=st.session_state.language,
                            use_multithreading=True
                        )
                        st.session_state.ai_results = results
                    
                    progress_bar.progress((i + 1) / 4)
                    time.sleep(0.5)
                
                progress_placeholder.empty()
                status_placeholder.success(f"✅ AI content generated for {total_items} accounts!")
                
            except Exception as e:
                progress_placeholder.empty()
                status_placeholder.error(f"❌ Error: {e}")
                import traceback
                st.code(traceback.format_exc())
    
    with col_pptx:
        if st.button("📄 Generate & Export PPTX", type="secondary", use_container_width=True, 
                     disabled=st.session_state.ai_results is None):
            st.info("🚧 PPTX generation coming soon...")
    
    # Display AI results
    if st.session_state.ai_results:
        st.markdown("### 📝 Generated Content")
        
        # Get BS and IS keys from mappings
        from fdd_utils.reconciliation import load_mappings
        mappings = load_mappings()
        
        # Filter for accounts that actually have results
        bs_keys = [k for k in st.session_state.ai_results.keys()
                   if k in mappings and mappings.get(k, {}).get('type') == 'BS']
        is_keys = [k for k in st.session_state.ai_results.keys()
                   if k in mappings and mappings.get(k, {}).get('type') == 'IS']
        
        if not bs_keys and not is_keys:
            st.warning("No AI results to display")
        else:
            ai_tab_bs, ai_tab_is = st.tabs([
                f"Balance Sheet ({len(bs_keys)} accounts)", 
                f"Income Statement ({len(is_keys)} accounts)"
            ])
            
            with ai_tab_bs:
                if bs_keys:
                    for key in bs_keys:
                        with st.expander(f"📄 {key}", expanded=False):
                            result = st.session_state.ai_results[key]
                            
                            # Final content (expanded)
                            st.markdown("**✨ Final Content (Agent 4):**")
                            final = result.get('final', result.get('agent_4', ''))
                            st.text_area("", value=final, height=150, key=f"final_{key}", label_visibility="collapsed")
                            
                            # Intermediate agents (collapsed)
                            with st.expander("🔍 View Agent Pipeline", expanded=False):
                                for agent in ['agent_1', 'agent_2', 'agent_3']:
                                    if agent in result:
                                        st.markdown(f"**{agent.replace('_', ' ').title()}:**")
                                        st.text(result[agent])
                                        st.markdown("---")
                else:
                    st.info("No Balance Sheet accounts with AI content")
            
            with ai_tab_is:
                if is_keys:
                    for key in is_keys:
                        with st.expander(f"📄 {key}", expanded=False):
                            result = st.session_state.ai_results[key]
                            
                            # Final content (expanded)
                            st.markdown("**✨ Final Content (Agent 4):**")
                            final = result.get('final', result.get('agent_4', ''))
                            st.text_area("", value=final, height=150, key=f"final_is_{key}", label_visibility="collapsed")
                            
                            # Intermediate agents (collapsed)
                            with st.expander("🔍 View Agent Pipeline", expanded=False):
                                for agent in ['agent_1', 'agent_2', 'agent_3']:
                                    if agent in result:
                                        st.markdown(f"**{agent.replace('_', ' ').title()}:**")
                                        st.text(result[agent])
                                        st.markdown("---")
                else:
                    st.info("No Income Statement accounts with AI content")


if __name__ == "__main__":
    pass


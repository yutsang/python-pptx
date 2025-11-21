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
    if 'bs_is_results' not in st.session_state:
        st.session_state.bs_is_results = None
    if 'ai_results' not in st.session_state:
        st.session_state.ai_results = None
    if 'reconciliation' not in st.session_state:
        st.session_state.reconciliation = None


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
    
    # File upload
    uploaded_file = st.file_uploader(
        "Upload Databook (Excel)",
        type=['xlsx', 'xls'],
        help="Upload your financial databook"
    )
    
    if uploaded_file:
        # Save to temp file
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            temp_path = tmp_file.name
        
        st.success(f"✅ File uploaded: {uploaded_file.name}")
        
        # Entity name selection
        st.markdown("---")
        entity_options = get_entity_names(temp_path)
        entity_name = st.selectbox(
            "Entity Name",
            options=entity_options,
            help="Leave blank for single entity databook"
        )
        
        # Financial statement sheet selection
        st.markdown("---")
        sheet_options = get_financial_sheets(temp_path)
        if sheet_options:
            selected_sheet = st.selectbox(
                "Financial Statement Sheet",
                options=sheet_options,
                help="Sheet containing both Balance Sheet and Income Statement"
            )
        else:
            st.warning("No sheets found in file")
            selected_sheet = None
        
        # Model selection
        st.markdown("---")
        st.subheader("🤖 AI Model")
        model_type = st.radio(
            "Select Model",
            options=['local', 'openai', 'deepseek'],
            index=0,
            help="AI model for content generation"
        )
        
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
                    st.dataframe(recon_bs, use_container_width=True, height=400)
                    
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
                    st.dataframe(recon_is, use_container_width=True, height=400)
                    
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
    
    if st.button("▶️ Generate AI Content", type="primary"):
        with st.spinner("Running AI pipeline..."):
            try:
                results = run_ai_pipeline(
                    mapping_keys=st.session_state.workbook_list,
                    dfs=st.session_state.dfs,
                    model_type=st.session_state.model_type,
                    language=st.session_state.language,
                    use_multithreading=True
                )
                
                st.session_state.ai_results = results
                st.success("✅ AI content generated!")
                st.rerun()
            except Exception as e:
                st.error(f"❌ Error generating content: {e}")
    
    # Display AI results
    if st.session_state.ai_results:
        ai_tab_bs, ai_tab_is = st.tabs(["Balance Sheet Accounts", "Income Statement Accounts"])
        
        # Get BS and IS keys from mappings
        from fdd_utils.reconciliation import load_mappings
        mappings = load_mappings()
        
        bs_keys = [k for k in st.session_state.workbook_list 
                   if k in mappings and mappings[k].get('type') == 'BS']
        is_keys = [k for k in st.session_state.workbook_list 
                   if k in mappings and mappings[k].get('type') == 'IS']
        
        with ai_tab_bs:
            for key in bs_keys:
                if key in st.session_state.ai_results:
                    with st.expander(f"📄 {key}", expanded=False):
                        result = st.session_state.ai_results[key]
                        
                        # Final content (expanded)
                        st.markdown("**Final Content (Agent 4):**")
                        final = result.get('final', result.get('agent_4', ''))
                        st.text_area("", value=final, height=150, key=f"final_{key}", label_visibility="collapsed")
                        
                        # Intermediate agents (collapsed)
                        with st.expander("🔍 View Agent Pipeline", expanded=False):
                            for agent in ['agent_1', 'agent_2', 'agent_3']:
                                if agent in result:
                                    st.markdown(f"**{agent.replace('_', ' ').title()}:**")
                                    st.text(result[agent])
                                    st.markdown("---")
        
        with ai_tab_is:
            for key in is_keys:
                if key in st.session_state.ai_results:
                    with st.expander(f"📄 {key}", expanded=False):
                        result = st.session_state.ai_results[key]
                        
                        # Final content (expanded)
                        st.markdown("**Final Content (Agent 4):**")
                        final = result.get('final', result.get('agent_4', ''))
                        st.text_area("", value=final, height=150, key=f"final_is_{key}", label_visibility="collapsed")
                        
                        # Intermediate agents (collapsed)
                        with st.expander("🔍 View Agent Pipeline", expanded=False):
                            for agent in ['agent_1', 'agent_2', 'agent_3']:
                                if agent in result:
                                    st.markdown(f"**{agent.replace('_', ' ').title()}:**")
                                    st.text(result[agent])
                                    st.markdown("---")
    
    # Area 3: PPTX Generation & Download
    st.markdown("---")
    st.header("📑 PowerPoint Generation")
    
    if st.session_state.ai_results:
        col1, col2 = st.columns([1, 1])
        
        with col1:
            if st.button("📝 Generate PowerPoint", type="primary", use_container_width=True):
                with st.spinner("Generating PowerPoint..."):
                    try:
                        # TODO: Implement PPTX generation
                        # Combine BS + IS into one PPTX
                        st.info("🚧 PPTX generation coming soon...")
                        st.session_state.pptx_file = None
                    except Exception as e:
                        st.error(f"❌ Error generating PPTX: {e}")
        
        with col2:
            if st.session_state.get('pptx_file'):
                with open(st.session_state.pptx_file, 'rb') as f:
                    st.download_button(
                        label="⬇️ Download PowerPoint",
                        data=f.read(),
                        file_name="financial_report.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True
                    )
    else:
        st.info("Generate AI content first to enable PowerPoint generation")


if __name__ == "__main__":
    pass


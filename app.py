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
from fdd_utils.content_generation import run_ai_pipeline, run_ai_pipeline_with_progress, extract_final_contents

# Import PPTX generation
try:
    import datetime
    import uuid
    import re
    from fdd_utils.pptx_generation import export_pptx, merge_presentations
    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False

# Page config
st.set_page_config(
    page_title="Financial Data Processing",
    page_icon="📊",
    layout="wide"
)

# Custom CSS for full width
st.markdown("""
<style>
.block-container {padding-top: 1rem; max-width: 100% !important;}
.stTabs [data-baseweb="tab-list"] {gap: 2px;}
.stTabs [data-baseweb="tab"] {padding: 10px 20px;}
.main .block-container {max-width: 100%; padding-left: 2rem; padding-right: 2rem;}
[data-testid="stAppViewContainer"] {max-width: 100%;}
</style>
""", unsafe_allow_html=True)


def load_latest_results_from_logs():
    """Load the most recent results from logs directory"""
    import yaml
    import os
    import glob
    
    logs_dir = 'fdd_utils/logs'
    if not os.path.exists(logs_dir):
        return None
    
    # Find all run directories
    run_dirs = glob.glob(os.path.join(logs_dir, 'run_*'))
    if not run_dirs:
        return None
    
    # Get the most recent one
    latest_run = max(run_dirs, key=os.path.getmtime)
    results_file = os.path.join(latest_run, 'results.yml')
    
    if os.path.exists(results_file):
        try:
            with open(results_file, 'r', encoding='utf-8') as f:
                results = yaml.safe_load(f)
            return results
        except Exception as e:
            print(f"Error loading results from {results_file}: {e}")
            return None
    return None


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
        # Try to load from latest log file
        latest_results = load_latest_results_from_logs()
        if latest_results:
            st.session_state.ai_results = latest_results
            st.session_state.results_loaded_from_logs = True
    if 'reconciliation' not in st.session_state:
        st.session_state.reconciliation = None
    if 'model_type' not in st.session_state:
        st.session_state.model_type = 'local'
    if 'project_name' not in st.session_state:
        st.session_state.project_name = None
    if 'last_run_folder' not in st.session_state:
        st.session_state.last_run_folder = None
    if 'results_loaded_from_logs' not in st.session_state:
        st.session_state.results_loaded_from_logs = False


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
        # Entity name selection (dropdown + text box)
        st.markdown("---")
        st.markdown("**🏢 Entity Name**")
        entity_options = get_entity_names(temp_path)
        
        # Dropdown for selection
        selected_entity = st.selectbox(
            "Select entity from list",
            options=[""] + entity_options,
            help="Select an entity from the list",
            label_visibility="collapsed",
            key="entity_dropdown"
        )
        
        # Text box for modification/custom input
        entity_name = st.text_input(
            "Or type/modify entity name",
            value=selected_entity if selected_entity else "",
            placeholder="Enter or modify entity name...",
            help="Type a custom entity name or modify the selected one",
            label_visibility="collapsed",
            key="entity_text_input"
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
        
        # Model selection (radio buttons)
        st.markdown("---")
        st.markdown("**🤖 AI Model**")
        model_type = st.radio(
            "Select model",
            options=['local', 'openai', 'deepseek'],
            index=0,
            help="AI model for content generation",
            label_visibility="collapsed"
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
    
    # Load mappings to filter accounts by type
    from fdd_utils.reconciliation import load_mappings
    mappings = load_mappings()
    
    # Separate accounts by type (BS or IS)
    bs_accounts = []
    is_accounts = []
    other_accounts = []
    
    for key in st.session_state.workbook_list:
        if key in st.session_state.dfs:
            acc_type = mappings.get(key, {}).get('type', '')
            if acc_type == 'BS':
                bs_accounts.append(key)
            elif acc_type == 'IS':
                is_accounts.append(key)
            else:
                other_accounts.append(key)
    
    # Create tabs for BS and IS
    tab_bs, tab_is = st.tabs(["Balance Sheet", "Income Statement"])
    
    with tab_bs:
        bs_data = st.session_state.bs_is_results.get('balance_sheet') if st.session_state.bs_is_results else None
        recon_bs = st.session_state.reconciliation[0] if st.session_state.reconciliation else None
        
        # Always show tabs if we have BS accounts, even if bs_data is None
        if bs_accounts:
            # Second level tabs - only show BS accounts
            tab_names = ["📊 Reconciliation"] + [f"📋 {key}" for key in bs_accounts]
            bs_tabs = st.tabs(tab_names)
            
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
            
            # Individual account tabs - only BS accounts
            for idx, key in enumerate(bs_accounts, 1):
                if idx < len(bs_tabs):
                    with bs_tabs[idx]:
                        if key in st.session_state.dfs:
                            st.dataframe(st.session_state.dfs[key], use_container_width=True)
                        else:
                            st.warning(f"Data not found for account: {key}")
        elif bs_data is not None:
            st.info("No Balance Sheet accounts found in workbook list")
        else:
            st.info("No Balance Sheet data extracted")
    
    with tab_is:
        is_data = st.session_state.bs_is_results.get('income_statement') if st.session_state.bs_is_results else None
        recon_is = st.session_state.reconciliation[1] if st.session_state.reconciliation else None
        
        # Always show tabs if we have IS accounts, even if is_data is None
        if is_accounts:
            # Second level tabs - only show IS accounts
            tab_names = ["📊 Reconciliation"] + [f"📋 {key}" for key in is_accounts]
            is_tabs = st.tabs(tab_names)
            
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
            
            # Individual account tabs - only IS accounts
            for idx, key in enumerate(is_accounts, 1):
                if idx < len(is_tabs):
                    with is_tabs[idx]:
                        if key in st.session_state.dfs:
                            st.dataframe(st.session_state.dfs[key], use_container_width=True)
                        else:
                            st.warning(f"Data not found for account: {key}")
        elif is_data is not None:
            st.info("No Income Statement accounts found in workbook list")
        else:
            st.info("No Income Statement data extracted")
    
    # Area 2: AI Content Generation
    st.markdown("---")
    st.header("🤖 AI Content Generation")
    
    # Horizontal buttons: Generate AI + Export PPTX
    col_ai, col_pptx = st.columns(2)
    
    # Progress area - full width, outside columns
    progress_container = st.container()
    
    with col_ai:
        generate_clicked = st.button("▶️ Generate AI Content", type="primary", use_container_width=True)
    
    # Progress display - full width, outside button columns
    if generate_clicked:
        with progress_container:
            st.markdown("### 🔄 AI Processing Progress")
            progress_placeholder = st.empty()
            status_placeholder = st.empty()
            eta_placeholder = st.empty()
            
            try:
                total_items = len(st.session_state.workbook_list)
                total_agents = 4
                total_steps = total_agents * total_items

                # Initialize progress tracking
                progress_state = {'current_step': 0, 'current_agent': 0, 'current_item': 0}
                
                # Progress callback function for real-time updates
                def update_progress(agent_num, agent_name, item_num, total_items_in_agent, completed_items):
                    progress_state['current_agent'] = agent_num
                    progress_state['current_item'] = item_num
                    
                    # Calculate overall progress
                    completed_steps = (agent_num - 1) * total_items + completed_items
                    progress = min(completed_steps / total_steps, 1.0)
                    
                    # Update progress bar
                    progress_placeholder.progress(progress)
                    
                    # Update status
                    status_placeholder.info(
                        f"🔄 Running Agent {agent_num}/4: {agent_name} | "
                        f"Processing item {item_num}/{total_items_in_agent}"
                    )
                    
                    # Calculate ETA
                    import time
                    if hasattr(update_progress, 'start_time'):
                        elapsed = time.time() - update_progress.start_time
                        if completed_steps > 0:
                            avg_time_per_step = elapsed / completed_steps
                            remaining_steps = total_steps - completed_steps
                            eta_seconds = avg_time_per_step * remaining_steps
                            eta_minutes = int(eta_seconds / 60)
                            eta_secs = int(eta_seconds % 60)
                            eta_placeholder.info(
                                f"📊 Progress: {completed_steps}/{total_steps} steps | "
                                f"ETA: {eta_minutes}m {eta_secs}s"
                            )
                    else:
                        update_progress.start_time = time.time()
                        eta_placeholder.info(f"📊 Progress: {completed_steps}/{total_steps} steps | ETA: Calculating...")
                
                # Show initial status
                import time
                start_time = time.time()
                update_progress.start_time = start_time
                
                status_placeholder.info(f"🚀 Starting AI pipeline for {total_items} accounts...")
                eta_placeholder.info(f"📊 Progress: 0/{total_steps} steps | ETA: Calculating...")
                progress_placeholder.progress(0)

                # Run the actual pipeline with progress updates
                results = run_ai_pipeline_with_progress(
                    mapping_keys=st.session_state.workbook_list,
                    dfs=st.session_state.dfs,
                    model_type=st.session_state.get('model_type', 'local'),
                    language=st.session_state.language,
                    use_multithreading=True,
                    progress_callback=update_progress
                )

                # Log results for debugging
                if results:
                    print(f"AI Pipeline completed: {len(results)} accounts processed")
                    content_count = 0
                    for key, value in results.items():
                        if isinstance(value, dict):
                            for agent_key, content in value.items():
                                if content and len(str(content).strip()) > 0:
                                    content_count += 1
                    print(f"Content generated: {content_count} items")
                    print(f"Results keys: {list(results.keys())}")
                    # Debug: Print first result structure
                    if results:
                        first_key = list(results.keys())[0]
                        print(f"First result structure for '{first_key}': {list(results[first_key].keys())}")
                else:
                    print("AI Pipeline failed: No results generated")

                st.session_state.ai_results = results

                # Show actual progress based on results
                if results and len(results) > 0:
                    # Check if any content was actually generated
                    has_content = False
                    total_content_items = 0
                    for key, value in results.items():
                        if isinstance(value, dict):
                            for agent_key, content in value.items():
                                if content and len(str(content).strip()) > 0:
                                    has_content = True
                                    total_content_items += 1

                    if has_content:
                        status_placeholder.success(f"✅ AI content generated successfully! {total_content_items} content items created for {len(results)} accounts.")
                    else:
                        status_placeholder.warning(f"⚠️ AI processing completed but no content was generated. This usually means the AI model is not properly configured.")
                        status_placeholder.info(f"💡 Selected model: **{st.session_state.get('model_type', 'local').upper()}** - Check configuration and try again.")
                else:
                    status_placeholder.error(f"❌ AI processing failed completely - no results generated. Check AI model setup.")

                progress_placeholder.progress(1.0)  # Complete the progress bar
                eta_placeholder.empty()
                status_placeholder.success(f"✅ AI content generated for {total_items} accounts! (Completed in {int(time.time() - start_time)}s)")
                
                # Force rerun to display results in UI
                st.rerun()

            except Exception as e:
                progress_placeholder.empty()
                eta_placeholder.empty()
                status_placeholder.error(f"❌ Error: {e}")
                import traceback
                st.code(traceback.format_exc())
    
    with col_pptx:
        if st.button("📄 Generate & Export PPTX", type="secondary", use_container_width=True,
                     disabled=st.session_state.ai_results is None):
            generate_pptx_presentation()
    
    # Display AI results
    if st.session_state.ai_results:
        st.markdown("### 📝 Generated Content")
        
        # Show if results were loaded from logs
        if st.session_state.get('results_loaded_from_logs'):
            st.info("ℹ️ Results loaded from latest log file. Click 'Generate AI Content' to create new results.")
            if st.button("🔄 Reload from Latest Log File"):
                latest_results = load_latest_results_from_logs()
                if latest_results:
                    st.session_state.ai_results = latest_results
                    st.session_state.results_loaded_from_logs = True
                    st.rerun()
                else:
                    st.warning("No results found in log files")
        
        # Debug: Show results info - always expanded for troubleshooting
        with st.expander("🔍 Debug: Results Info", expanded=True):
            st.write(f"**Total results:** {len(st.session_state.ai_results)}")
            if st.session_state.ai_results:
                all_keys = list(st.session_state.ai_results.keys())
                st.write(f"**Result keys:** {all_keys[:20]}{'...' if len(all_keys) > 20 else ''}")
                
                # Show detailed structure of first few results
                for idx, key in enumerate(all_keys[:3]):
                    result = st.session_state.ai_results[key]
                    st.write(f"**Result {idx+1} - Key: '{key}'**")
                    if isinstance(result, dict):
                        st.write(f"  - Type: dict with keys: {list(result.keys())}")
                        for agent_key in ['final', 'agent_4', 'agent_3', 'agent_2', 'agent_1']:
                            if agent_key in result:
                                content = result[agent_key]
                                content_str = str(content) if content else "None"
                                content_len = len(content_str.strip()) if content_str else 0
                                # Show more of the preview for debugging
                                preview = content_str[:200] if content_len > 0 else ""
                                st.write(f"  - {agent_key}: {content_len} chars" + (f" (preview: {preview}...)" if content_len > 0 else " (empty)"))
                    else:
                        st.write(f"  - Type: {type(result)}")
                        st.write(f"  - Value: {str(result)[:100]}")

        # Get BS and IS keys from mappings first
        from fdd_utils.reconciliation import load_mappings
        mappings = load_mappings()

        # Filter for accounts that actually have results
        bs_keys = []
        is_keys = []
        other_keys = []

        for k in st.session_state.ai_results.keys():
            result = st.session_state.ai_results[k]
            # Check if result is a dict with content
            if isinstance(result, dict):
                # Check if there's any actual content - be more lenient
                has_content = False
                content_preview = ""
                
                # First, check standard structure (agent_1, agent_2, etc.)
                for agent_key in ['final', 'agent_4', 'agent_3', 'agent_2', 'agent_1']:
                    if agent_key in result:
                        content = result[agent_key]
                        if content is not None:
                            content_str = str(content).strip()
                            # More lenient check - accept any non-empty string
                            if (len(content_str) > 0 and 
                                content_str.lower() not in ['none', 'null', 'nan', '', 'n/a', 'na'] and
                                not content_str.isspace()):
                                has_content = True
                                if not content_preview:
                                    content_preview = content_str[:100]
                                break
                
                # If no content found, check if this is log file structure (nested with 'output' key)
                if not has_content:
                    for agent_name in result.keys():
                        if isinstance(result[agent_name], dict):
                            # Check for 'output' key in nested structure
                            if 'output' in result[agent_name]:
                                output_content = result[agent_name]['output']
                                if output_content is not None:
                                    content_str = str(output_content).strip()
                                    if (len(content_str) > 0 and 
                                        content_str.lower() not in ['none', 'null', 'nan', '', 'n/a', 'na'] and
                                        not content_str.isspace()):
                                        has_content = True
                                        if not content_preview:
                                            content_preview = content_str[:100]
                                        # Also extract to top level for display
                                        if agent_name == 'agent_1':
                                            result['agent_1'] = output_content
                                        elif agent_name == 'agent_2':
                                            result['agent_2'] = output_content
                                        elif agent_name == 'agent_3':
                                            result['agent_3'] = output_content
                                        elif agent_name == 'agent_4':
                                            result['agent_4'] = output_content
                                            result['final'] = output_content
                                        break
                
                # Always categorize by type, even if no content (for debugging)
                if k in mappings:
                    acc_type = mappings[k].get('type')
                    if acc_type == 'BS':
                        if has_content:
                            bs_keys.append(k)
                        else:
                            # Add to bs_keys anyway for display, but mark as empty
                            bs_keys.append(k)
                    elif acc_type == 'IS':
                        if has_content:
                            is_keys.append(k)
                        else:
                            # Add to is_keys anyway for display, but mark as empty
                            is_keys.append(k)
                    else:
                        if has_content:
                            other_keys.append(k)
                else:
                    # If not in mappings, add to other
                    if has_content:
                        other_keys.append(k)

        # Check if there's any actual content (including error messages)
        has_content = False
        for key, value in st.session_state.ai_results.items():
            if isinstance(value, dict):
                final_content = value.get('final', '')
                if final_content and len(str(final_content).strip()) > 0:
                    has_content = True
                    break

        if not has_content:
            st.warning("⚠️ AI processing completed but no content was generated.")
            st.error("**Possible causes:**")
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("**For Local AI:**")
                st.markdown("- Local AI server not running")
                st.markdown("- Wrong server URL/port")
                st.markdown("- Model not loaded")
            with col2:
                st.markdown("**For Cloud AI:**")
                st.markdown("- Invalid API key")
                st.markdown("- Network connection issues")
                st.markdown("- Rate limits exceeded")

            st.info("💡 **Selected model:** " + st.session_state.get('model_type', 'local').upper())
            st.info("🔧 Configure your AI model and try again.")

            # Show sample structure anyway
            st.markdown("---")
            st.markdown("### 📋 Expected Content Structure")
            st.info("When AI models are properly configured, you will see content for each account like this:")

            # Show sample content structure
            sample_accounts = bs_keys[:2] + is_keys[:2]  # Show first 2 from each type
            for account in sample_accounts[:4]:  # Max 4 samples
                with st.expander(f"📄 {account} (Sample Structure)", expanded=False):
                    st.markdown("**✨ Final Content (Validator):**")
                    st.text_area("",
                        f"This would contain the final polished content for {account}...",
                        height=100, disabled=True, label_visibility="collapsed")

                    with st.expander("🔍 View Agent Pipeline (Sample)", expanded=False):
                        st.markdown("**Generator:** Raw generated content...")
                        st.markdown("**Auditor:** Verified and corrected content...")
                        st.markdown("**Refiner:** Improved and refined content...")
                        st.markdown("**Validator:** Final formatted content...")

        # Show summary
        st.info(f"📊 **Summary:** {len(bs_keys)} Balance Sheet accounts, {len(is_keys)} Income Statement accounts, {len(other_keys)} other accounts")
        
        if not bs_keys and not is_keys and not other_keys:
            st.warning("⚠️ No AI results to display with content")
            st.info(f"Found {len(st.session_state.ai_results)} results but none have content. Check debug info above.")
            
            # Show all keys even without content for debugging
            st.markdown("### 📋 All Results (including empty)")
            all_result_keys = list(st.session_state.ai_results.keys())
            for key in all_result_keys[:10]:  # Show first 10
                with st.expander(f"📄 {key} (no content)", expanded=False):
                    result = st.session_state.ai_results[key]
                    st.json(result if isinstance(result, dict) else {"value": str(result)})
        else:
            ai_tab_bs, ai_tab_is = st.tabs([
                f"Balance Sheet ({len(bs_keys)} accounts)",
                f"Income Statement ({len(is_keys)} accounts)"
            ])

            with ai_tab_bs:
                if bs_keys:
                    for key in bs_keys:
                        result = st.session_state.ai_results.get(key, {})
                        if not isinstance(result, dict):
                            result = {}
                        
                        # Check if has content
                        final = result.get('final', result.get('agent_4', ''))
                        has_final = final and str(final).strip() and str(final).lower() not in ['none', 'null', 'nan']
                        
                        expander_label = f"📄 {key}" + (" ⚠️ (empty)" if not has_final else "")
                        with st.expander(expander_label, expanded=False):
                            # Final content (expanded) - use actual agent name
                            st.markdown("**✨ Final Content (Validator):**")
                            if has_final:
                                st.text_area("", value=str(final), height=150, key=f"final_{key}", label_visibility="collapsed")
                            else:
                                st.warning("No final content available")
                                # Show what we have
                                st.markdown("**Available agent outputs:**")
                                for agent_key in ['agent_1', 'agent_2', 'agent_3', 'agent_4', 'final']:
                                    if agent_key in result:
                                        content = result[agent_key]
                                        if content:
                                            st.text(f"{agent_key}: {str(content)[:100]}...")
                                        else:
                                            st.text(f"{agent_key}: (empty)")

                            # Intermediate agents (collapsed) - use actual agent names
                            with st.expander("🔍 View Agent Pipeline", expanded=False):
                                agent_map = {
                                    'agent_1': 'Generator',
                                    'agent_2': 'Auditor',
                                    'agent_3': 'Refiner',
                                    'agent_4': 'Validator'
                                }
                                found_any = False
                                for agent in ['agent_1', 'agent_2', 'agent_3', 'agent_4']:
                                    if agent in result and result[agent] is not None:
                                        content_str = str(result[agent]).strip()
                                        if content_str and content_str.lower() not in ['none', 'null', 'nan']:
                                            found_any = True
                                            agent_name = agent_map.get(agent, agent.replace('_', ' ').title())
                                            st.markdown(f"**{agent_name}:**")
                                            st.text(content_str)
                                            st.markdown("---")
                                if not found_any:
                                    st.info("No agent outputs available")
                else:
                    st.info("No Balance Sheet accounts with AI content")

            with ai_tab_is:
                if is_keys:
                    for key in is_keys:
                        result = st.session_state.ai_results.get(key, {})
                        if not isinstance(result, dict):
                            result = {}
                        
                        # Check if has content
                        final = result.get('final', result.get('agent_4', ''))
                        has_final = final and str(final).strip() and str(final).lower() not in ['none', 'null', 'nan']
                        
                        expander_label = f"📄 {key}" + (" ⚠️ (empty)" if not has_final else "")
                        with st.expander(expander_label, expanded=False):
                            # Final content (expanded) - use actual agent name
                            st.markdown("**✨ Final Content (Validator):**")
                            if has_final:
                                st.text_area("", value=str(final), height=150, key=f"final_is_{key}", label_visibility="collapsed")
                            else:
                                st.warning("No final content available")
                                # Show what we have
                                st.markdown("**Available agent outputs:**")
                                for agent_key in ['agent_1', 'agent_2', 'agent_3', 'agent_4', 'final']:
                                    if agent_key in result:
                                        content = result[agent_key]
                                        if content:
                                            st.text(f"{agent_key}: {str(content)[:100]}...")
                                        else:
                                            st.text(f"{agent_key}: (empty)")

                            # Intermediate agents (collapsed) - use actual agent names
                            with st.expander("🔍 View Agent Pipeline", expanded=False):
                                agent_map = {
                                    'agent_1': 'Generator',
                                    'agent_2': 'Auditor',
                                    'agent_3': 'Refiner',
                                    'agent_4': 'Validator'
                                }
                                found_any = False
                                for agent in ['agent_1', 'agent_2', 'agent_3', 'agent_4']:
                                    if agent in result and result[agent] is not None:
                                        content_str = str(result[agent]).strip()
                                        if content_str and content_str.lower() not in ['none', 'null', 'nan']:
                                            found_any = True
                                            agent_name = agent_map.get(agent, agent.replace('_', ' ').title())
                                            st.markdown(f"**{agent_name}:**")
                                            st.text(content_str)
                                            st.markdown("---")
                                if not found_any:
                                    st.info("No agent outputs available")
                else:
                    st.info("No Income Statement accounts with AI content")
    else:
        st.info("🔄 Generate AI content to see results here")


def generate_pptx_presentation():
    """Generate PPTX presentation from AI results"""
    if not st.session_state.ai_results:
        st.error("❌ No AI results available. Generate AI content first.")
        return

    if not PPTX_AVAILABLE:
        st.error("❌ PPTX generation not available. Missing required modules.")
        return

    try:
        with st.spinner("📊 Generating PowerPoint presentation..."):
            # Get necessary data
            project_name = st.session_state.get('project_name', 'Project')
            language = st.session_state.get('language', 'Eng')

            # Load mappings
            from fdd_utils.reconciliation import load_mappings
            mappings = load_mappings()

            # Find template
            template_path = None
            for template in ["fdd_utils/template.pptx", "backups/fdd_utils/template.pptx", "template.pptx"]:
                if os.path.exists(template):
                    template_path = template
                    break

            if not template_path:
                st.error("❌ PowerPoint template not found. Please ensure template.pptx exists.")
                return

            # Create output directory
            output_dir = "fdd_utils/output"
            os.makedirs(output_dir, exist_ok=True)

            # Generate unique filename
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            unique_id = str(uuid.uuid4())[:8]
            sanitized_name = re.sub(r'[^\w\-_]', '_', project_name).strip('_')
            output_filename = f"{sanitized_name}_Combined_{timestamp}_{unique_id}.pptx"
            output_path = os.path.join(output_dir, output_filename)

            # Create temporary directory for intermediate files
            with tempfile.TemporaryDirectory() as temp_dir:
                bs_temp = os.path.join(temp_dir, "bs_temp.pptx")
                is_temp = os.path.join(temp_dir, "is_temp.pptx")
                bs_md = os.path.join(temp_dir, "bs_content.md")
                is_md = os.path.join(temp_dir, "is_content.md")

                # Generate markdown content
                bs_content = convert_ai_results_to_markdown(st.session_state.ai_results, mappings, 'BS')
                is_content = convert_ai_results_to_markdown(st.session_state.ai_results, mappings, 'IS')

                if not bs_content and not is_content:
                    st.error("❌ No content generated for PPTX")
                    return

                # Write markdown files
                if bs_content:
                    with open(bs_md, 'w', encoding='utf-8') as f:
                        f.write(bs_content)

                if is_content:
                    with open(is_md, 'w', encoding='utf-8') as f:
                        f.write(is_content)

                # Generate individual presentations
                generated_files = []

                if bs_content:
                    export_pptx(template_path, bs_md, bs_temp, project_name,
                              language='chinese' if language == 'Chn' else 'english',
                              statement_type='BS', row_limit=20)
                    generated_files.append(bs_temp)

                if is_content:
                    export_pptx(template_path, is_md, is_temp, project_name,
                              language='chinese' if language == 'Chn' else 'english',
                              statement_type='IS', row_limit=20)
                    generated_files.append(is_temp)

                # Merge presentations if both exist
                if len(generated_files) == 2:
                    merge_presentations(bs_temp, is_temp, output_path)
                elif len(generated_files) == 1:
                    # Copy single presentation
                    import shutil
                    shutil.copy(generated_files[0], output_path)
                else:
                    st.error("❌ No presentations generated")
                    return

            # Verify file was created
            if os.path.exists(output_path):
                # Read file for download
                with open(output_path, 'rb') as f:
                    pptx_data = f.read()

                st.success("✅ PowerPoint presentation generated successfully!")

                # Download button
                st.download_button(
                    label="📥 Download PowerPoint Presentation",
                    data=pptx_data,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )
            else:
                st.error("❌ PowerPoint file was not created")

    except Exception as e:
        st.error(f"❌ PPTX generation failed: {e}")
        import traceback
        st.code(traceback.format_exc())


def convert_ai_results_to_markdown(ai_results, mappings, statement_type='BS'):
    """Convert AI results to markdown format for PPTX generation"""
    if not ai_results:
        return ""

    content_lines = []

    # Filter accounts by statement type
    filtered_accounts = []
    for account_key in ai_results.keys():
        if account_key in mappings:
            acc_type = mappings[account_key].get('type')
            if statement_type == 'BS' and acc_type == 'BS':
                filtered_accounts.append(account_key)
            elif statement_type == 'IS' and acc_type == 'IS':
                filtered_accounts.append(account_key)

    for account_key in filtered_accounts:
        result = ai_results[account_key]
        final_content = result.get('final', result.get('agent_4', ''))

        # Always include accounts, even with empty/error content
        content_lines.append(f"## {account_key}")
        content_lines.append("")

        if final_content and final_content.strip():
            # Add the content
            content_lines.append(final_content.strip())
        else:
            # Add placeholder for empty content
            content_lines.append(f"[No content generated for {account_key}]")

        content_lines.append("")
        content_lines.append("---")
        content_lines.append("")

    return "\n".join(content_lines)




if __name__ == "__main__":
    pass


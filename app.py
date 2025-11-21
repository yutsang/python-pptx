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
        # Entity name selection (combined)
        st.markdown("---")
        st.markdown("**🏢 Entity Name**")
        entity_options = get_entity_names(temp_path)
        entity_name = st.selectbox(
            "Select entity or type custom",
            options=[""] + entity_options + ["[Custom]"],
            help="Select from list or choose '[Custom]' to type your own",
            label_visibility="collapsed"
        )

        if entity_name == "[Custom]":
            entity_name = st.text_input(
                "Type custom entity name",
                placeholder="Enter entity name...",
                help="Custom entity name",
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
            # Full width progress area
            st.markdown("### 🔄 AI Processing Progress")
            progress_placeholder = st.empty()
            status_placeholder = st.empty()
            eta_placeholder = st.empty()

            try:
                total_items = len(st.session_state.workbook_list)
                total_agents = 4
                total_steps = total_agents * total_items

                # Initialize progress
                progress_bar = progress_placeholder.progress(0)
                step_count = 0

                # Show initial status
                status_placeholder.info(f"🚀 Starting AI pipeline for {total_items} accounts...")
                eta_placeholder.info(f"📊 Progress: 0/{total_steps} steps | ETA: Calculating...")

                # Run the actual pipeline first
                import time
                start_time = time.time()

                status_placeholder.info(f"🔄 Running Agent 1/4: Generator...")
                eta_placeholder.info(f"📊 Processing {total_items} accounts through Generator...")

                results = run_ai_pipeline(
                    mapping_keys=st.session_state.workbook_list,
                    dfs=st.session_state.dfs,
                    model_type=st.session_state.get('model_type', 'local'),
                    language=st.session_state.language,
                    use_multithreading=True
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

                progress_bar.progress(1.0)  # Complete the progress bar

                progress_placeholder.empty()
                eta_placeholder.empty()
                status_placeholder.success(f"✅ AI content generated for {total_items} accounts! (Completed in {int(time.time() - start_time)}s)")

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

        # Get BS and IS keys from mappings first
        from fdd_utils.reconciliation import load_mappings
        mappings = load_mappings()

        # Filter for accounts that actually have results
        bs_keys = []
        is_keys = []

        for k in st.session_state.ai_results.keys():
            if k in mappings:
                acc_type = mappings[k].get('type')
                if acc_type == 'BS':
                    bs_keys.append(k)
                elif acc_type == 'IS':
                    is_keys.append(k)

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
                    st.markdown("**✨ Final Content (Agent 4):**")
                    st.text_area("",
                        f"This would contain the final polished content for {account}...",
                        height=100, disabled=True, label_visibility="collapsed")

                    with st.expander("🔍 View Agent Pipeline (Sample)", expanded=False):
                        st.markdown("**Generator (Agent 1):** Raw generated content...")
                        st.markdown("**Auditor (Agent 2):** Verified and corrected content...")
                        st.markdown("**Refiner (Agent 3):** Improved and refined content...")
                        st.markdown("**Validator (Agent 4):** Final formatted content...")

        # Continue with normal processing if we have content
        from fdd_utils.reconciliation import load_mappings
        mappings = load_mappings()

        # Filter for accounts that actually have results
        bs_keys = []
        is_keys = []

        for k in st.session_state.ai_results.keys():
            if k in mappings:
                acc_type = mappings[k].get('type')
                if acc_type == 'BS':
                    bs_keys.append(k)
                elif acc_type == 'IS':
                    is_keys.append(k)

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
                            if final and str(final).strip():
                                st.text_area("", value=str(final), height=150, key=f"final_{key}", label_visibility="collapsed")
                            else:
                                st.warning("No final content available")

                            # Intermediate agents (collapsed)
                            with st.expander("🔍 View Agent Pipeline", expanded=False):
                                for agent in ['agent_1', 'agent_2', 'agent_3']:
                                    if agent in result and result[agent] is not None:
                                        st.markdown(f"**{agent.replace('_', ' ').title()}:**")
                                        st.text(str(result[agent]))
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
                            if final and str(final).strip():
                                st.text_area("", value=str(final), height=150, key=f"final_is_{key}", label_visibility="collapsed")
                            else:
                                st.warning("No final content available")

                            # Intermediate agents (collapsed)
                            with st.expander("🔍 View Agent Pipeline", expanded=False):
                                for agent in ['agent_1', 'agent_2', 'agent_3']:
                                    if agent in result and result[agent] is not None:
                                        st.markdown(f"**{agent.replace('_', ' ').title()}:**")
                                        st.text(str(result[agent]))
                                        st.markdown("---")
                else:
                    st.info("No Income Statement accounts with AI content")
    else:
        st.info("🔄 Generate AI content to see results here")


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


if __name__ == "__main__":
    pass


#!/usr/bin/env python3
"""
Streamlit App for Financial Data Processing with AI
Combines extraction, reconciliation, AI generation, and PPTX export
"""

import streamlit as st
import pandas as pd
import os
import tempfile
import base64
import zipfile
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
    from fdd_utils.pptx_generation import export_pptx, merge_presentations, export_pptx_from_structured_data
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


def convert_ai_results_to_structured_data(ai_results, mappings, statement_type='BS', bs_is_results=None, dfs=None):
    """Convert AI results to structured data for PPTX generation"""
    if not ai_results:
        return []

    structured_data = []

    # Helper function to find mapping key from account name (checks aliases)
    def find_mapping_key_for_pptx(account_name):
        """Find the mapping key for an account name, checking aliases"""
        # First try direct key lookup
        if account_name in mappings:
            return account_name
        
        # Then check aliases
        for mapping_key, config in mappings.items():
            if mapping_key.startswith('_'):
                continue
            if isinstance(config, dict):
                aliases = config.get('aliases', [])
                if account_name in aliases:
                    return mapping_key
        return None

    # Helper to extract summary from content
    def extract_summary(content):
        """Extract summary from content - pass full content for AI to summarize later"""
        if not content:
            return ""
        content_str = str(content).strip()
        # Return full content - AI will control the summary length later (~200 words)
        # Don't truncate here as it may cut meaningful text mid-sentence
        return content_str

    # Use order from financial statements (bs_is_results) - this provides the proper presentation order
    # Get the financial statement DataFrame to extract account order
    financial_statement_df = None
    if bs_is_results:
        if statement_type == 'BS':
            financial_statement_df = bs_is_results.get('balance_sheet')
        elif statement_type == 'IS':
            financial_statement_df = bs_is_results.get('income_statement')
    
    # Create order map from financial statement (row order = presentation order)
    financial_statement_order = {}
    if financial_statement_df is not None and not financial_statement_df.empty:
        # Get account names from the first column (usually account names)
        # The first column typically contains account names
        first_col = financial_statement_df.iloc[:, 0] if len(financial_statement_df.columns) > 0 else None
        
        if first_col is not None:
            for idx, account_name_in_statement in enumerate(first_col):
                if pd.notna(account_name_in_statement) and str(account_name_in_statement).strip():
                    account_name_str = str(account_name_in_statement).strip()
                    # Skip totals, subtotals, and headers
                    skip_keywords = ['total', '合计', '总计', '小计', 'subtotal', 'sub-total', 'sub total']
                    if any(skip in account_name_str.lower() for skip in skip_keywords):
                        continue
                    
                    # Find matching mapping key for this account name
                    mapping_key = find_mapping_key_for_pptx(account_name_str)
                    if mapping_key:
                        financial_statement_order[mapping_key] = idx
                    # Also try direct match with account_key (for cases where account_key matches statement name)
                    financial_statement_order[account_name_str] = idx
                    
                    # Also check aliases - if any alias matches, use that mapping_key's order
                    for map_key, config in mappings.items():
                        if map_key.startswith('_'):
                            continue
                        if isinstance(config, dict):
                            aliases = config.get('aliases', [])
                            if account_name_str in aliases:
                                financial_statement_order[map_key] = idx
    
    # Filter and collect accounts with their mapping info
    account_info_list = []
    for account_key in ai_results.keys():
        # Use find_mapping_key to check aliases
        mapping_key = find_mapping_key_for_pptx(account_key)
        if mapping_key:
            acc_type = mappings[mapping_key].get('type')
            if statement_type == 'BS' and acc_type == 'BS':
                category = mappings[mapping_key].get('category', '')
                # Get order from financial statement, fallback to 9999
                order = financial_statement_order.get(mapping_key, financial_statement_order.get(account_key, 9999))
                account_info_list.append({
                    'account_key': account_key,
                    'mapping_key': mapping_key,
                    'category': category,
                    'order': order
                })
            elif statement_type == 'IS' and acc_type == 'IS':
                category = mappings[mapping_key].get('category', '')
                # Get order from financial statement, fallback to 9999
                order = financial_statement_order.get(mapping_key, financial_statement_order.get(account_key, 9999))
                account_info_list.append({
                    'account_key': account_key,
                    'mapping_key': mapping_key,
                    'category': category,
                    'order': order
                })
    
    # Sort by financial statement order (this preserves the proper presentation order)
    # Then group by category, but maintain order within category
    account_info_list.sort(key=lambda x: (x['order'], x['category'], x['mapping_key']))

    for account_info in account_info_list:
        account_key = account_info['account_key']
        mapping_key = account_info['mapping_key']
        category = account_info['category']
        result = ai_results[account_key]
        
        # Extract final content, handling nested structures
        def extract_final_content(result_dict):
            """Extract final content from result, handling nested structures"""
            # Try final first
            content = result_dict.get('final', '')
            if isinstance(content, dict):
                content = content.get('output', content.get('content', ''))
            
            # If no final, try agent_4
            if not content or not str(content).strip():
                content = result_dict.get('agent_4', '')
                if isinstance(content, dict):
                    content = content.get('output', content.get('content', ''))
            
            # If still no content, try other agents in reverse order
            if not content or not str(content).strip():
                for agent_key in ['agent_3', 'agent_2', 'agent_1']:
                    content = result_dict.get(agent_key, '')
                    if isinstance(content, dict):
                        content = content.get('output', content.get('content', ''))
                    if content and str(content).strip():
                        break
            
            # If still no content, check if result_dict itself is a string
            if not content or not str(content).strip():
                if isinstance(result_dict, str):
                    content = result_dict
            
            return content
        
        final_content = extract_final_content(result)
        
        # Extract financial data
        financial_data = None
        if dfs and account_key in dfs:
            df = dfs[account_key]
            if not df.empty:
                financial_data = df
        
        # Check if account has all zeros or values below 0.01 - exclude from display if so
        # Values below 0.01 are treated as insignificant/zero
        has_significant_balance = False
        if financial_data is not None and not financial_data.empty:
            # Check all numeric columns (skip first column which is usually account names)
            numeric_cols = financial_data.select_dtypes(include=[float, int]).columns
            if len(numeric_cols) > 0:
                # Check if any value is >= 0.01 (significant)
                for col in numeric_cols:
                    if (financial_data[col].abs() >= 0.01).any():
                        has_significant_balance = True
                        break
        else:
            # If no financial data, include it (might be a header or total row)
            has_significant_balance = True
        
        # Skip accounts with all zero or insignificant balances (< 0.01)
        if not has_significant_balance:
            import logging
            logger = logging.getLogger(__name__)
            logger.info(f"Skipping account {account_key} ({mapping_key}) - all balances are zero or below 0.01")
            continue
        
        # Detect if content is Chinese
        from fdd_utils.pptx_generation import detect_chinese_text
        commentary_text = str(final_content).strip() if final_content and str(final_content).strip() else f"[No content generated for {account_key}]"
        is_chinese = detect_chinese_text(commentary_text)
        
        # Get proper account name from financial statement for display
        # If databook is Chinese, use the name from financial statement
        display_name = mapping_key  # Default to mapping_key
        if financial_statement_df is not None and not financial_statement_df.empty:
            # Find the account name from financial statement that matches this mapping_key
            first_col = financial_statement_df.iloc[:, 0] if len(financial_statement_df.columns) > 0 else None
            if first_col is not None:
                for account_name_in_statement in first_col:
                    if pd.notna(account_name_in_statement):
                        account_name_str = str(account_name_in_statement).strip()
                        # Check if this matches our mapping_key or account_key
                        found_mapping_key = find_mapping_key_for_pptx(account_name_str)
                        if found_mapping_key == mapping_key or account_name_str == account_key:
                            display_name = account_name_str
                            break
        
        # Structure the data for PPTX
        account_data = {
            'account_name': account_key,
            'mapping_key': mapping_key,
            'display_name': display_name,  # Proper name from financial statement
            'category': category,
            'financial_data': financial_data,
            'commentary': commentary_text,
            'summary': extract_summary(final_content) if final_content else "",
            'is_chinese': is_chinese
        }
        
        structured_data.append(account_data)

    return structured_data


def generate_pptx_presentation():
    """Generate PPTX presentation from AI results"""
    if not st.session_state.ai_results:
        st.error("❌ No AI results available. Generate AI content first.")
        return

    if not PPTX_AVAILABLE:
        st.error("❌ PPTX generation not available. Missing required modules.")
        return

    # Get necessary data
    project_name = st.session_state.get('project_name', 'Project')
    entity_name = st.session_state.get('entity_name', project_name)
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

    # Generate timestamp for filenames
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    sanitized_entity = re.sub(r'[^\w\-_]', '_', str(entity_name)).strip('_')
    if not sanitized_entity or sanitized_entity == '_':
        sanitized_entity = 'Project'

    try:

            # Create temporary directory for intermediate files
            with tempfile.TemporaryDirectory() as temp_dir:
                bs_temp = os.path.join(temp_dir, "bs_temp.pptx")
                is_temp = os.path.join(temp_dir, "is_temp.pptx")
                bs_md = os.path.join(temp_dir, "bs_content.md")
                is_md = os.path.join(temp_dir, "is_content.md")

                # Generate structured data for PPTX
                bs_data = convert_ai_results_to_structured_data(
                    st.session_state.ai_results, 
                    mappings, 
                    'BS',
                    bs_is_results=st.session_state.bs_is_results,
                    dfs=st.session_state.dfs
                )
                is_data = convert_ai_results_to_structured_data(
                    st.session_state.ai_results, 
                    mappings, 
                    'IS',
                    bs_is_results=st.session_state.bs_is_results,
                    dfs=st.session_state.dfs
                )

                # Debug: Log data
                print(f"DEBUG: BS accounts: {len(bs_data)}")
                print(f"DEBUG: IS accounts: {len(is_data)}")

                if not bs_data and not is_data:
                    st.error("❌ No content generated for PPTX")
                    st.info(f"DEBUG: AI results keys: {list(st.session_state.ai_results.keys())[:10] if st.session_state.ai_results else 'None'}")
                    st.info(f"DEBUG: DFS keys: {list(st.session_state.dfs.keys())[:10] if st.session_state.dfs else 'None'}")
                    return

                # Generate ONE combined presentation with both BS and IS
                combined_output_path = os.path.join(output_dir, f"{sanitized_entity}_{timestamp}.pptx")
                from fdd_utils.pptx_generation import export_pptx_from_structured_data_combined
                # Pass language info to PPTX generation
                is_chinese_databook = (language == 'Chn')
                # Get temp_path and selected_sheet from session state
                excel_temp_path = st.session_state.get('temp_path')
                # Fallback to default if temp_path is missing or invalid
                if not excel_temp_path or not os.path.exists(excel_temp_path):
                    if os.path.exists("databook.xlsx"):
                        excel_temp_path = "databook.xlsx"
                        print(f"DEBUG: temp_path invalid, falling back to {excel_temp_path}")
                
                excel_selected_sheet = st.session_state.get('selected_sheet')
                print(f"DEBUG: Embedding tables from: {excel_temp_path}, sheet: {excel_selected_sheet}")
                export_pptx_from_structured_data_combined(
                    template_path, bs_data, is_data, combined_output_path, project_name,
                    language='chinese' if language == 'Chn' else 'english',
                    temp_path=excel_temp_path,
                    selected_sheet=excel_selected_sheet,
                    is_chinese_databook=is_chinese_databook
                )
                generated_files = [combined_output_path]
                output_files = [('Combined', combined_output_path)]

                if not generated_files:
                    st.error("❌ No presentations generated")
                    return

            # Wait for files to be fully written
            import time
            time.sleep(0.5)
            
            # Prepare download data - keep files separate, no merge
            # Store files in session state for download
            import time
            time.sleep(0.3)  # Wait for files to be ready
            
            if len(output_files) == 1:
                # Single file - download directly
                file_path = output_files[0][1]
                if os.path.exists(file_path):
                    with open(file_path, 'rb') as f:
                        download_data = f.read()
                    download_filename = os.path.basename(file_path)
                    
                    # Store in session state for download button
                    st.session_state.pptx_download_data = download_data
                    st.session_state.pptx_download_filename = download_filename
                    st.session_state.pptx_download_mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    st.session_state.pptx_ready = True
            else:
                # Multiple files - create zip and download
                zip_buffer = tempfile.NamedTemporaryFile(delete=False, suffix='.zip')
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    for file_type, file_path in output_files:
                        if os.path.exists(file_path):
                            zip_file.write(file_path, os.path.basename(file_path))
                zip_buffer.close()
                
                # Read zip file
                with open(zip_buffer.name, 'rb') as f:
                    zip_data = f.read()
                
                zip_filename = f"{sanitized_entity}_{timestamp}.zip"
                
                # Store in session state for download button
                st.session_state.pptx_download_data = zip_data
                st.session_state.pptx_download_filename = zip_filename
                st.session_state.pptx_download_mime = "application/zip"
                st.session_state.pptx_ready = True
                
                # Clean up temp zip file
                try:
                    os.unlink(zip_buffer.name)
                except:
                    pass
            
            # Rerun to show download button
            st.rerun()

    except Exception as e:
        st.error(f"❌ PPTX generation failed: {e}")
        import traceback
        st.code(traceback.format_exc())


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
    if 'reconciliation' not in st.session_state:
        st.session_state.reconciliation = None
    if 'model_type' not in st.session_state:
        st.session_state.model_type = 'local'
    if 'project_name' not in st.session_state:
        st.session_state.project_name = None
    if 'last_run_folder' not in st.session_state:
        st.session_state.last_run_folder = None
    if 'entity_name' not in st.session_state:
        st.session_state.entity_name = None
    if 'pptx_download_trigger' not in st.session_state:
        st.session_state.pptx_download_trigger = None
    if 'button_click_counter' not in st.session_state:
        st.session_state.button_click_counter = 0
    if 'pptx_ready' not in st.session_state:
        st.session_state.pptx_ready = False
    if 'temp_path' not in st.session_state:
        st.session_state.temp_path = None
    if 'selected_sheet' not in st.session_state:
        st.session_state.selected_sheet = None
    if 'prev_entity_dropdown' not in st.session_state:
        st.session_state.prev_entity_dropdown = ''


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
                            cleaned_match = match.strip()
                            if cleaned_match:  # Only add non-empty strings
                                entity_names.add(cleaned_match)
            except:
                continue
        
        # Filter out empty or whitespace-only strings from the set
        filtered_names = [name for name in entity_names if name and name.strip()]
        return [''] + sorted(list(filtered_names))
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

# Title with refresh button at top right
col_title, col_refresh = st.columns([10, 1])
with col_title:
    st.title("📊 Financial Data Processing & AI Generation")
with col_refresh:
    st.markdown("<br>", unsafe_allow_html=True)  # Align button with title
    if st.button("🔄", help="Refresh page and reset", use_container_width=True, key="refresh_main"):
        # Clear session state to reset
        for key in ['dfs', 'workbook_list', 'language', 'bs_is_results', 'ai_results', 
                    'reconciliation', 'entity_name', 'project_name', 'pptx_ready', 
                    'pptx_download_data', 'pptx_download_filename', 'pptx_download_mime',
                    'prev_uploaded_file', 'prev_entity_dropdown', 'selected_sheet']:
            if key in st.session_state:
                del st.session_state[key]
        st.rerun()

# Show entity name and financial sheet selection in main area (only if data not processed)
if st.session_state.dfs is None:
    # Entity name and financial sheet in parallel
    col_entity, col_sheet = st.columns(2)
    
    with col_entity:
        st.markdown("**🏢 Entity Name**")
        # Get temp_path from sidebar state or default
        temp_path = st.session_state.get('temp_path', None)
        if not temp_path:
            # Check if default file exists
            default_path = "databook.xlsx"
            if os.path.exists(default_path):
                temp_path = default_path
                st.session_state.temp_path = default_path
        
        if temp_path and os.path.exists(temp_path):
            entity_options = get_entity_names(temp_path)
            # Initialize entity_name in session state if not exists
            if 'entity_name' not in st.session_state:
                st.session_state.entity_name = ""
            
            # Track previous dropdown value to detect changes
            prev_dropdown = st.session_state.get('prev_entity_dropdown', '')
            
            # Dropdown for selection
            selected_entity = st.selectbox(
                label="Select entity from list",
                options=[""] + entity_options,
                help="Select an entity from the list",
                label_visibility="collapsed",
                key="entity_dropdown"
            )
            
            # Use widget value directly - don't set session state with widget key name
            # Get initial value from session state or dropdown
            initial_value = st.session_state.get('entity_name', '')
            if selected_entity and selected_entity != prev_dropdown:
                initial_value = selected_entity
                st.session_state.prev_entity_dropdown = selected_entity
            
            # Editable text box - widget manages its own state
            # Use value parameter to set initial value without triggering warning
            entity_name = st.text_input(
                label="Or type/modify entity name",
                value=initial_value if initial_value else "",
                placeholder="Enter or modify entity name...",
                help="Type a custom entity name or modify the selected one",
                label_visibility="collapsed",
                key="entity_text_input"
            )
            
            # Update session state from widget value (not the other way around)
            # This avoids the warning about widget keys
            if entity_name:
                st.session_state.entity_name = entity_name
            
            # Sync dropdown selection to text input - update session state and rerun
            if selected_entity and selected_entity != prev_dropdown:
                st.session_state.entity_name = selected_entity
                st.session_state.prev_entity_dropdown = selected_entity
        else:
            # Initialize if not exists
            if 'entity_name' not in st.session_state:
                st.session_state.entity_name = ""
            entity_name = st.text_input(
                label="Entity name",
                value=st.session_state.entity_name,
                placeholder="Enter entity name...",
                label_visibility="collapsed",
                key="entity_text_input"
            )
            if entity_name != st.session_state.entity_name:
                st.session_state.entity_name = entity_name
    
    with col_sheet:
        st.markdown("**📊 Financial Statement Sheet**")
        # Get temp_path from sidebar state or default
        temp_path = st.session_state.get('temp_path', None)
        if not temp_path:
            # Check if default file exists
            default_path = "databook.xlsx"
            if os.path.exists(default_path):
                temp_path = default_path
                st.session_state.temp_path = default_path
        
        if temp_path and os.path.exists(temp_path):
            sheet_options = get_financial_sheets(temp_path)
            if sheet_options:
                # Initialize selected_sheet in session state if not exists
                if 'selected_sheet' not in st.session_state:
                    st.session_state.selected_sheet = sheet_options[0] if sheet_options else None
                
                selected_sheet = st.selectbox(
                    label="Select sheet",
                    options=sheet_options,
                    index=0 if st.session_state.selected_sheet not in sheet_options else sheet_options.index(st.session_state.selected_sheet),
                    help="Sheet containing both BS and IS",
                    label_visibility="collapsed",
                    key="sheet_select"
                )
                st.session_state.selected_sheet = selected_sheet
            else:
                st.warning("No sheets found")
                selected_sheet = None
                st.session_state.selected_sheet = None
        else:
            st.info("Please select or upload a file")
            selected_sheet = None
            st.session_state.selected_sheet = None
    
    # Process button
    if st.button("🚀 Process Data", type="primary", use_container_width=True, key="process_data_main"):
        temp_path = st.session_state.get('temp_path', None)
        if not temp_path:
            st.error("Please upload a file first")
        else:
            st.session_state.entity_name = entity_name
            st.session_state.selected_sheet = selected_sheet
            st.session_state.process_data_clicked = True
            st.rerun()

# Sidebar - simplified
with st.sidebar:
    st.header("⚙️ Configuration")
    
    # File upload - no checkbox, always available
    st.markdown("**📁 Databook File**")
    
    # File uploader (always shown)
    uploaded_file = st.file_uploader(
        "Upload Excel (or use default databook.xlsx)",
        type=['xlsx', 'xls'],
        help="Upload your financial databook, or leave empty to use databook.xlsx",
        key="file_uploader"
    )
    
    # Determine which file to use
    if uploaded_file:
        # Save uploaded file to temp location
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            temp_path = tmp_file.name
        
        # Check if file has changed - clear all cache if so
        prev_file = st.session_state.get('prev_uploaded_file', None)
        if prev_file != uploaded_file.name:
            # File changed - clear all data and entity-related cache
            for key in ['dfs', 'workbook_list', 'language', 'bs_is_results', 'ai_results', 
                        'reconciliation', 'entity_name', 'project_name', 'pptx_ready', 
                        'pptx_download_data', 'pptx_download_filename', 'pptx_download_mime',
                        'prev_entity_dropdown', 'selected_sheet']:
                if key in st.session_state:
                    if key == 'entity_name':
                        st.session_state[key] = ""  # Clear but keep key
                    else:
                        del st.session_state[key]
            st.session_state.prev_uploaded_file = uploaded_file.name
        
        st.success(f"✅ Using uploaded file: {uploaded_file.name}")
        st.session_state.temp_path = temp_path
    else:
        # No upload - use default databook.xlsx if it exists
        default_path = "databook.xlsx"
        if os.path.exists(default_path):
            temp_path = default_path
            
            # Check if we switched from uploaded to default - clear all cache
            prev_file = st.session_state.get('prev_uploaded_file', None)
            if prev_file is not None:
                # Switched back to default - clear all data
                for key in ['dfs', 'workbook_list', 'language', 'bs_is_results', 'ai_results', 
                            'reconciliation', 'entity_name', 'project_name', 'pptx_ready', 
                            'pptx_download_data', 'pptx_download_filename', 'pptx_download_mime',
                            'prev_entity_dropdown', 'selected_sheet']:
                    if key in st.session_state:
                        if key == 'entity_name':
                            st.session_state[key] = ""  # Clear but keep key
                        else:
                            del st.session_state[key]
                st.session_state.prev_uploaded_file = None
            
            st.success(f"✅ Using default: {default_path}")
            st.session_state.temp_path = temp_path
        else:
            st.warning("⚠️ Default databook.xlsx not found. Please upload a file.")
            temp_path = None
            if 'temp_path' in st.session_state:
                del st.session_state.temp_path
    
    # Model selection (radio buttons)
    if temp_path:
        st.markdown("---")
        st.markdown("**🤖 AI Model**")
        # Initialize model_type in session state if not exists
        if 'model_type' not in st.session_state:
            st.session_state.model_type = 'local'
        
        # Widget with key automatically manages session state - don't modify it manually
        st.radio(
            label="Select model",
            options=['local', 'openai', 'deepseek'],
            index=['local', 'openai', 'deepseek'].index(st.session_state.model_type) if st.session_state.model_type in ['local', 'openai', 'deepseek'] else 0,
            help="AI model for content generation",
            label_visibility="collapsed",
            key="model_type"
        )
        
        # Show current model with Streamlit green success reminder
        model_display = {
            'local': 'qwen3-chat',
            'openai': 'gpt-4o-mini',
            'deepseek': 'deepseek-chat'
        }
        st.success(f"🤖 AI Mode: {model_display.get(st.session_state.model_type, st.session_state.model_type.upper())}")

# Process data if button was clicked
if st.session_state.get('process_data_clicked', False):
    st.session_state.process_data_clicked = False
    temp_path = st.session_state.get('temp_path', None)
    # If no temp_path but default exists, use it
    if not temp_path:
        default_path = "databook.xlsx"
        if os.path.exists(default_path):
            temp_path = default_path
            st.session_state.temp_path = default_path
    
    entity_name = st.session_state.get('entity_name', '')
    selected_sheet = st.session_state.get('selected_sheet', None)
    
    if temp_path:
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
                # model_type is already set by the radio widget, don't modify it here
                if 'model_type' not in st.session_state:
                    st.session_state.model_type = 'local'
                st.session_state.project_name = bs_is_results.get('project_name') if bs_is_results else None
                st.session_state.entity_name = entity_name
                
                st.success("✅ Data processed successfully!")
                st.rerun()
                
            except Exception as e:
                st.error(f"❌ Error processing data: {e}")
                import traceback
                st.code(traceback.format_exc())

# Main content
if st.session_state.dfs is None:
    st.info("👈 Upload a databook, set entity name and sheet, then click 'Process Data' to begin")
else:
    # Data Display header
    st.header("📈 Data Display")
    
    # Load mappings to filter accounts by type
    from fdd_utils.reconciliation import load_mappings
    mappings = load_mappings()
    
    # Helper function to find mapping key from account name (checks aliases)
    def find_mapping_key(account_name):
        """Find the mapping key for an account name, checking aliases"""
        # First try direct key lookup
        if account_name in mappings:
            return account_name
        
        # Then check aliases
        for mapping_key, config in mappings.items():
            if mapping_key.startswith('_'):
                continue
            if isinstance(config, dict):
                aliases = config.get('aliases', [])
                if account_name in aliases:
                    return mapping_key
        return None
    
    # Separate accounts by type (BS or IS)
    bs_accounts = []
    is_accounts = []
    other_accounts = []
    
    # Check if workbook_list exists and has items
    if st.session_state.workbook_list:
        for key in st.session_state.workbook_list:
            if key in st.session_state.dfs:
                # Find the mapping key (may be different from account name)
                mapping_key = find_mapping_key(key)
                if mapping_key:
                    acc_type = mappings[mapping_key].get('type', '')
                    if acc_type == 'BS':
                        bs_accounts.append(key)
                    elif acc_type == 'IS':
                        is_accounts.append(key)
                    else:
                        other_accounts.append(key)
                else:
                    other_accounts.append(key)
    else:
        # If workbook_list is empty, try to use all keys from dfs
        if st.session_state.dfs:
            for key in st.session_state.dfs.keys():
                # Find the mapping key (may be different from account name)
                mapping_key = find_mapping_key(key)
                if mapping_key:
                    acc_type = mappings[mapping_key].get('type', '')
                    if acc_type == 'BS':
                        bs_accounts.append(key)
                    elif acc_type == 'IS':
                        is_accounts.append(key)
                    else:
                        other_accounts.append(key)
                else:
                    other_accounts.append(key)
    
    # Create tabs for BS and IS
    tab_bs, tab_is = st.tabs(["Balance Sheet", "Income Statement"])
    
    with tab_bs:
        bs_data = st.session_state.bs_is_results.get('balance_sheet') if st.session_state.bs_is_results else None
        recon_bs = st.session_state.reconciliation[0] if st.session_state.reconciliation else None
        
        # Show tabs if we have BS accounts OR if we have dfs data
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
                    
                    # Summary stats with Reports (immaterial) on the right
                    col1, col2, col3, col4, col_reports = st.columns([1, 1, 1, 1, 1])
                    with col1:
                        matches = (recon_bs['Match'] == '✅ Match').sum()
                        st.metric("✅ Matches", matches)
                    with col2:
                        diffs = (recon_bs['Match'] == '❌ Diff').sum()
                        st.metric("❌ Differences", diffs)
                    with col3:
                        not_found = (recon_bs['Match'] == '⚠️ Not Found').sum()
                        st.metric("⚠️ Not Found", not_found)
                    with col4:
                        immaterial = (recon_bs['Match'] == '✅ Immaterial').sum()
                        st.metric("✅ Immaterial", immaterial)
                    with col_reports:
                        # Reports = total immaterial rows
                        st.metric("Reports", immaterial, help="Immaterial row matches")
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
        elif st.session_state.dfs:
            # If no BS accounts found but we have dfs, show all dfs keys
            all_dfs_keys = list(st.session_state.dfs.keys())
            if all_dfs_keys:
                st.warning(f"⚠️ No BS accounts found in mappings. Showing all {len(all_dfs_keys)} accounts from data.")
                tab_names = ["📊 Reconciliation"] + [f"📋 {key}" for key in all_dfs_keys]
                bs_tabs = st.tabs(tab_names)
                
                # Reconciliation tab
                with bs_tabs[0]:
                    if recon_bs is not None and not recon_bs.empty:
                        display_df = recon_bs.copy()
                        for col in display_df.columns:
                            display_df[col] = display_df[col].apply(
                                lambda x: f"{x:,.0f}" if isinstance(x, (int, float)) and x != 0 else str(x)
                            )
                        st.dataframe(display_df, use_container_width=True, height=400)
                    else:
                        st.info("No reconciliation data available")
                
                # Individual account tabs
                for idx, key in enumerate(all_dfs_keys, 1):
                    if idx < len(bs_tabs):
                        with bs_tabs[idx]:
                            if key in st.session_state.dfs:
                                st.dataframe(st.session_state.dfs[key], use_container_width=True)
                            else:
                                st.warning(f"Data not found for account: {key}")
            else:
                st.info("No data available in dfs")
        else:
            st.info("No Balance Sheet data available")
    
    with tab_is:
        is_data = st.session_state.bs_is_results.get('income_statement') if st.session_state.bs_is_results else None
        recon_is = st.session_state.reconciliation[1] if st.session_state.reconciliation else None
        
        # Show tabs if we have IS accounts OR if we have dfs data
        if is_accounts:
            # Second level tabs - only show IS accounts
            tab_names = ["📊 Reconciliation"] + [f"📋 {key}" for key in is_accounts]
            is_tabs = st.tabs(tab_names)
        elif st.session_state.dfs and not bs_accounts:
            # If no IS accounts found but we have dfs (and no BS accounts shown), show all dfs keys
            all_dfs_keys = list(st.session_state.dfs.keys())
            if all_dfs_keys:
                st.warning(f"⚠️ No IS accounts found in mappings. Showing all {len(all_dfs_keys)} accounts from data.")
                tab_names = ["📊 Reconciliation"] + [f"📋 {key}" for key in all_dfs_keys]
                is_tabs = st.tabs(tab_names)
                is_accounts = all_dfs_keys  # Use all keys for display
            else:
                st.info("No data available in dfs")
        else:
            st.info("No Income Statement data available")
        
        if is_accounts:
            
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
                    
                    # Summary stats with Reports (immaterial) on the right
                    col1, col2, col3, col4, col_reports = st.columns([1, 1, 1, 1, 1])
                    with col1:
                        matches = (recon_is['Match'] == '✅ Match').sum()
                        st.metric("✅ Matches", matches)
                    with col2:
                        diffs = (recon_is['Match'] == '❌ Diff').sum()
                        st.metric("❌ Differences", diffs)
                    with col3:
                        not_found = (recon_is['Match'] == '⚠️ Not Found').sum()
                        st.metric("⚠️ Not Found", not_found)
                    with col4:
                        immaterial = (recon_is['Match'] == '✅ Immaterial').sum()
                        st.metric("✅ Immaterial", immaterial)
                    with col_reports:
                        # Reports = total immaterial rows
                        st.metric("Reports", immaterial, help="Immaterial row matches")
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
    
    # Area 2: AI Content Generation - embedded in processed data
    st.markdown("---")
    col_header, col_pptx, col_download = st.columns([3, 1, 0.3])
    with col_header:
        st.header("🤖 AI Content Generation")
    with col_pptx:
        st.markdown("<br>", unsafe_allow_html=True)  # Align button with header
        pptx_key = f"pptx_btn_{st.session_state.button_click_counter}"
        if st.button("📄 Generate & Export PPTX", type="secondary", use_container_width=True,
                     disabled=st.session_state.ai_results is None, key=pptx_key):
            st.session_state.button_click_counter += 1
            generate_pptx_presentation()
    with col_download:
        st.markdown("<br>", unsafe_allow_html=True)  # Align with button
        # Download icon button - only show if PPTX is ready, directly downloads
        if st.session_state.get('pptx_ready', False) and 'pptx_download_data' in st.session_state:
            download_data = st.session_state.pptx_download_data
            download_filename = st.session_state.pptx_download_filename
            download_mime = st.session_state.pptx_download_mime
            
            download_icon_key = f"download_icon_{st.session_state.button_click_counter}"
            # Use download_button styled as a small icon
            st.download_button(
                label="📥",
                data=download_data,
                file_name=download_filename,
                mime=download_mime,
                help="Download generated PPTX",
                key=download_icon_key,
                use_container_width=True
            )
    
    # Generate AI Content button - embedded in processed data section
    ai_key = f"ai_btn_{st.session_state.button_click_counter}"
    generate_clicked = st.button("▶️ Generate AI Content", type="primary", use_container_width=True, key=ai_key)
    if generate_clicked:
        st.session_state.button_click_counter += 1
    
    # Progress area - full width
    progress_container = st.container()
    
    # Progress display - full width, outside button columns
    if generate_clicked:
        with progress_container:
            st.markdown("### 🔄 AI Processing Progress")
            progress_placeholder = st.empty()
            status_placeholder = st.empty()
            
            try:
                total_items = len(st.session_state.workbook_list)
                total_agents = 4
                total_steps = total_agents * total_items

                # Initialize progress tracking
                progress_state = {'current_step': 0, 'current_agent': 0, 'current_item': 0}
                
                # Progress callback function for real-time updates
                def update_progress(agent_num, agent_name, item_num, total_items_in_agent, completed_items, key_name=None):
                    progress_state['current_agent'] = agent_num
                    progress_state['current_item'] = item_num
                    
                    # Calculate overall progress
                    # The completed_items parameter already includes offset from previous agents
                    completed_steps = completed_items
                    progress = min(completed_steps / total_steps, 1.0) if total_steps > 0 else 0.0
                    
                    # Update progress bar
                    progress_placeholder.progress(progress)
                    
                    # Calculate ETA and combine all info into one status
                    import time
                    key_display = f" | Key: {key_name}" if key_name else ""
                    if hasattr(update_progress, 'start_time'):
                        elapsed = time.time() - update_progress.start_time
                        if completed_steps > 0:
                            avg_time_per_step = elapsed / completed_steps
                            remaining_steps = total_steps - completed_steps
                            eta_seconds = avg_time_per_step * remaining_steps
                            eta_minutes = int(eta_seconds / 60)
                            eta_secs = int(eta_seconds % 60)
                            status_placeholder.info(
                                f"🔄 Running Agent {agent_num}/4: {agent_name} | "
                                f"Processing item {item_num}/{total_items_in_agent}{key_display} | "
                                f"Progress: {completed_steps}/{total_steps} steps | "
                                f"ETA: {eta_minutes}m {eta_secs}s"
                            )
                        else:
                            status_placeholder.info(
                                f"🔄 Running Agent {agent_num}/4: {agent_name} | "
                                f"Processing item {item_num}/{total_items_in_agent}{key_display} | "
                                f"Progress: {completed_steps}/{total_steps} steps | "
                                f"ETA: Calculating..."
                            )
                    else:
                        update_progress.start_time = time.time()
                        status_placeholder.info(
                            f"🔄 Running Agent {agent_num}/4: {agent_name} | "
                            f"Processing item {item_num}/{total_items_in_agent}{key_display} | "
                            f"Progress: {completed_steps}/{total_steps} steps | "
                            f"ETA: Calculating..."
                        )
                
                # Show initial status
                import time
                start_time = time.time()
                update_progress.start_time = start_time
                
                status_placeholder.info(f"🚀 Starting AI pipeline for {total_items} accounts... | Progress: 0/{total_steps} steps | ETA: Calculating...")
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
                status_placeholder.success(f"✅ AI content generated for {total_items} accounts! (Completed in {int(time.time() - start_time)}s)")
                
                # Force rerun to display results in UI
                st.rerun()

            except Exception as e:
                progress_placeholder.empty()
                status_placeholder.error(f"❌ Error: {e}")
                import traceback
                st.code(traceback.format_exc())
    
    # Display AI results
    if st.session_state.ai_results:
        st.markdown("### 📝 Generated Content")
        
        # Get BS and IS keys from mappings first
        from fdd_utils.reconciliation import load_mappings
        mappings = load_mappings()

        # Filter for accounts that actually have results
        bs_keys = []
        is_keys = []
        other_keys = []

        # Helper function to find mapping key from account name (checks aliases)
        def find_mapping_key_for_result(account_name):
            """Find the mapping key for an account name, checking aliases"""
            # First try direct key lookup
            if account_name in mappings:
                return account_name
            
            # Then check aliases
            for mapping_key, config in mappings.items():
                if mapping_key.startswith('_'):
                    continue
                if isinstance(config, dict):
                    aliases = config.get('aliases', [])
                    if account_name in aliases:
                        return mapping_key
            return None
        
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
                
                # Find mapping key (may be different from account name)
                mapping_key = find_mapping_key_for_result(k)
                if mapping_key:
                    acc_type = mappings[mapping_key].get('type')
                    if acc_type == 'BS':
                        bs_keys.append(k)
                    elif acc_type == 'IS':
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
                    st.text_area(
                        label=f"Sample content for {account}",
                        value=f"This would contain the final polished content for {account}...",
                        height=100,
                        disabled=True
                    )

                    st.markdown("---")
                    st.markdown("**🔍 View Agent Pipeline (Sample):**")
                    st.markdown("**Generator:** Raw generated content...")
                    st.markdown("**Auditor:** Verified and corrected content...")
                    st.markdown("**Refiner:** Improved and refined content...")
                    st.markdown("**Validator:** Final formatted content...")

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
            # Create 3-layer tabs: BS/IS -> Account -> Agent
            tab_list = []
            if bs_keys:
                tab_list.append(f"Balance Sheet ({len(bs_keys)} accounts)")
            if is_keys:
                tab_list.append(f"Income Statement ({len(is_keys)} accounts)")
            if other_keys:
                tab_list.append(f"Other ({len(other_keys)} accounts)")
            
            ai_tabs = st.tabs(tab_list)
            tab_idx = 0
            
            # Helper function to create account tabs with collapsible agent boxes
            def create_account_agent_tabs(keys, prefix=""):
                """Create tabs for accounts, each with collapsible boxes for agents inside"""
                if not keys:
                    return
                
                # Create account tabs (second layer)
                account_tab_names = [f"📄 {key}" for key in keys]
                account_tabs = st.tabs(account_tab_names)
                
                agent_map = {
                    'agent_1': 'Generator',
                    'agent_2': 'Auditor',
                    'agent_3': 'Refiner',
                    'agent_4': 'Validator',
                    'final': 'Final (Validator)'
                }
                
                for acc_idx, key in enumerate(keys):
                    with account_tabs[acc_idx]:
                        result = st.session_state.ai_results.get(key, {})
                        if not isinstance(result, dict):
                            result = {}
                        
                        # Extract content from nested structure if needed
                        def extract_content(result_dict, agent_key):
                            """Extract content from result, handling nested structures"""
                            content = result_dict.get(agent_key, '')
                            
                            # If content is a dict, try to get 'output' key
                            if isinstance(content, dict):
                                content = content.get('output', '')
                            
                            return content
                        
                        # Show Final first
                        final_content = extract_content(result, 'final')
                        has_final = final_content and str(final_content).strip() and str(final_content).lower() not in ['none', 'null', 'nan']
                        
                        if has_final:
                            st.markdown("#### ✨ Final (Validator)")
                            # Dynamic height based on content length
                            content_length = len(str(final_content))
                            dynamic_height = min(max(100, int(content_length / 3)), 600)  # Between 100-600px
                            st.text_area(
                                label="Final content",
                                value=str(final_content),
                                height=dynamic_height,
                                key=f"{prefix}{key}_final_display",
                                label_visibility="collapsed"
                            )
                            st.markdown("---")
                        
                        # Combine agent 1-4 in one collapsible box
                        agent_contents = []
                        agent_names_list = []
                        for agent_key in ['agent_1', 'agent_2', 'agent_3', 'agent_4']:
                            content = extract_content(result, agent_key)
                            if content and str(content).strip() and str(content).lower() not in ['none', 'null', 'nan']:
                                agent_name = agent_map.get(agent_key, agent_key)
                                agent_contents.append((agent_name, str(content)))
                                agent_names_list.append(agent_name)
                        
                        if agent_contents:
                            with st.expander(f"🔍 Agent Pipeline ({', '.join(agent_names_list)})", expanded=False):
                                for agent_name, content in agent_contents:
                                    st.markdown(f"**{agent_name}:**")
                                    # Dynamic height based on content length
                                    content_length = len(str(content))
                                    dynamic_height = min(max(80, int(content_length / 4)), 400)  # Between 80-400px
                                    st.text_area(
                                        label=f"Content for {agent_name}",
                                        value=content,
                                        height=dynamic_height,
                                        key=f"{prefix}{key}_{agent_name}_pipeline",
                                        label_visibility="collapsed"
                                    )
                                    if agent_contents.index((agent_name, content)) < len(agent_contents) - 1:
                                        st.markdown("---")
                        
                        if not has_final and not agent_contents:
                            st.warning("No agent outputs available for this account")
            
            if bs_keys:
                with ai_tabs[tab_idx]:
                    create_account_agent_tabs(bs_keys, "bs_")
                tab_idx += 1
            
            if is_keys:
                with ai_tabs[tab_idx]:
                    create_account_agent_tabs(is_keys, "is_")
                tab_idx += 1
            
            if other_keys:
                with ai_tabs[tab_idx]:
                    for key in other_keys:
                        result = st.session_state.ai_results.get(key, {})
                        if not isinstance(result, dict):
                            result = {}
                        with st.expander(f"📄 {key}", expanded=False):
                            st.json(result)




if __name__ == "__main__":
    pass


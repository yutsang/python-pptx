import streamlit as st
import pandas as pd
from common import assistant
import tempfile
import os
import json
import re
from utils.utils import name_mapping
import sys
from common.pptx_export import export_pptx
import datetime
import numpy as np
# from pptx_export import ReportGenerator  # Uncomment after renaming 2.pptx.py
sys.path.append('.')

st.set_page_config(page_title="Due Diligence Report Generator", layout="wide")
st.title("Real Estate Due Diligence Report Generator")

# --- Upload Excel File ---
st.sidebar.header("Upload Excel Databook")
uploaded_file = st.sidebar.file_uploader("Choose an Excel file", type=["xlsx"]) 

# --- Entity and Helpers ---
st.sidebar.header("Entity & Helpers")
entity = st.sidebar.selectbox("Select Entity", ["Haining", "Nanjing", "Ningbo"])

# Predefined entity suffixes for better UX
entity_suffixes = {
    "Haining": [" Wanpu", " Wanpu Limited", " Limited", " Co.", " Ltd", " Corp"],
    "Nanjing": [" Wanpu", " Wanpu Limited", " Limited", " Co.", " Ltd", " Corp"],
    "Ningbo": [" Wanpu", " Wanpu Limited", " Limited", " Co.", " Ltd", " Corp"]
}

# Use predefined suffixes or allow custom input
use_custom_helpers = st.sidebar.checkbox("Use custom entity helpers", value=False)
if use_custom_helpers:
    helpers_input = st.sidebar.text_input("Entity Helpers (comma separated)", value="Wanpu, Limited, ...", placeholder="Wanpu, Limited, ...")
    helpers = [h.strip() for h in helpers_input.split(",") if h.strip()]
else:
    helpers = entity_suffixes.get(entity, ["Wanpu", "Limited", "Co.", "Ltd", "Corp"])
    st.sidebar.caption(f"Using default helpers: {', '.join(helpers)}")

# --- AI or Local Mode ---
st.sidebar.header("Generation Mode")
use_ai = st.sidebar.radio("Use AI for text generation?", ("No (Test Mode)", "Yes (AI Mode)")) == "Yes (AI Mode)"

# --- Template Upload ---
st.sidebar.header("Upload PPTX Template")
template_file = st.sidebar.file_uploader("Choose a PPTX template", type=["pptx"])
template_path = None
if template_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp:
        tmp.write(template_file.read())
        template_path = tmp.name

# --- '000 Scaling Toggle ---
st.sidebar.header("Scaling Option")
convert_thousands = st.sidebar.radio("Convert figures in '000 to base units?", ("Yes", "No")) == "Yes"
st.sidebar.caption("If Yes, convert figures in '000 to base units (multiply by 1,000).")

# --- Sheet Type Selector ---
sheet_type = st.sidebar.selectbox("Sheet Type", ["BS", "IS"], index=0)

# --- Table Filtering Options ---
st.sidebar.header("Table Filtering")
enable_filtering = st.sidebar.checkbox("Enable table filtering by entity", value=True)
show_all_tables = st.sidebar.checkbox("Show all tables (not just matching ones)", value=False)

# --- State ---
if 'results' not in st.session_state:
    st.session_state['results'] = None
if 'keys' not in st.session_state:
    st.session_state['keys'] = []
if 'excel_tables' not in st.session_state:
    st.session_state['excel_tables'] = {}

# --- Output File Name ---
default_output_name = f"{entity}_{sheet_type}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
output_file_name = st.sidebar.text_input("Output PPTX File Name", value=default_output_name)

# --- Helper Functions ---
def extract_tables_from_worksheet_robust(excel_path, sheet_name, entity_keywords):
    """
    Robust table extraction using the original method from utils.py
    """
    try:
        import openpyxl
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        ws = wb[sheet_name]
        
        tables = []
        
        # Method 1: Try to extract from openpyxl tables (works for individually formatted tables)
        if hasattr(ws, '_tables') and ws._tables:
            for tbl in ws._tables.values():
                try:
                    ref = tbl.ref
                    min_col, min_row, max_col, max_row = openpyxl.utils.range_boundaries(ref)
                    data = []
                    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col, values_only=True):
                        data.append(row)
                    if data and len(data) >= 2:
                        tables.append({
                            'data': data,
                            'method': 'openpyxl_table',
                            'name': tbl.name,
                            'range': ref
                        })
                except Exception as e:
                    print(f"Failed to extract table {tbl.name}: {e}")
                    continue
        
        # Method 2: Original method from utils.py - DataFrame splitting on empty rows
        try:
            # Load the sheet as DataFrame
            xl = pd.ExcelFile(excel_path)
            if sheet_name in xl.sheet_names:
                df = xl.parse(sheet_name)
                
                # Split dataframes on empty rows (original method)
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
                
                # Filter dataframes by entity keywords (original method)
                combined_pattern = '|'.join(re.escape(kw) for kw in entity_keywords)
                
                for i, data_frame in enumerate(dataframes):
                    # Check if dataframe contains entity keywords
                    mask = data_frame.apply(
                        lambda row: row.astype(str).str.contains(
                            combined_pattern, case=False, regex=True, na=False
                        ).any(),
                        axis=1
                    )
                    
                    if mask.any():
                        # Convert DataFrame to list format for consistency
                        table_data = [data_frame.columns.tolist()] + data_frame.values.tolist()
                        
                        # Check if table has meaningful content (not empty)
                        if table_data and len(table_data) > 1:
                            # Check if there's actual data beyond headers
                            has_data = False
                            for row in table_data[1:]:  # Skip header row
                                if any(cell and str(cell).strip() for cell in row):
                                    has_data = True
                                    break
                            
                            if has_data:
                                tables.append({
                                    'data': table_data,
                                    'method': 'original_split',
                                    'name': f'original_table_{i}',
                                    'range': f'dataframe_{i}'
                                })
                        
        except Exception as e:
            print(f"Error in original table detection: {e}")
        
        return tables
        
    except Exception as e:
        print(f"Error in robust table extraction for worksheet view: {e}")
        return []

def save_results_to_markdown(results, entity: str):
    """Save AI-generated results back to the markdown file"""
    try:
        # Read existing markdown content
        with open('utils/bs_content.md', 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Update each section with new content
        for key, new_content in results.items():
            # Get the heading name from name_mapping
            heading = name_mapping.get(key, key)
            if heading is None:
                heading = key
            
            # Create the new section content
            new_section = f'### {heading}\n{new_content.strip()}\n\n'
            
            # Find and replace the existing section
            pattern = re.compile(r'(###\s+' + re.escape(heading) + r'.*?)(?=\n###|\Z)', re.DOTALL | re.IGNORECASE)
            
            if pattern.search(content):
                # Replace existing section
                content = pattern.sub(new_section, content)
            else:
                # Add new section at the end
                content += '\n' + new_section
        
        # Write back to file
        with open('utils/bs_content.md', 'w', encoding='utf-8') as f:
            f.write(content)
            
    except Exception as e:
        print(f"Error saving to markdown: {e}")
        raise

# --- Main Workflow ---
if uploaded_file:
    # Save uploaded file to a temp location
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        tmp.write(uploaded_file.read())
        excel_path = tmp.name
    xl = pd.ExcelFile(excel_path)
    st.success(f"Excel file loaded: {uploaded_file.name}")
    # Load mapping.json for worksheet mapping
    with open('utils/mapping.json', 'r') as f:
        mapping = json.load(f)
    # Define keys (can be made dynamic)
    keys = [
        "Cash", "AR", "Prepayments", "OR", "Other CA", "IP", "Other NCA",
        "AP", "Taxes payable", "OP", "Capital", "Reserve"
    ]
    if entity in ['Ningbo', 'Nanjing']:
        keys = [key for key in keys if key != "Reserve"]
    st.session_state['keys'] = keys
    # Build a set of all possible worksheet names from mapping.json for the selected keys
    mapped_sheet_names = set()
    for key in keys:
        mapped_sheet_names.update(mapping.get(key, []))
    valid_sheets = [sheet for sheet in xl.sheet_names if sheet in mapped_sheet_names or sheet in name_mapping.values()]
    # --- View Worksheets Tabs (independent) ---
    tabs = st.tabs(valid_sheets)
    for i, tab in enumerate(tabs):
        with tab:
            sheet = valid_sheets[i]
            # Map worksheet name to key
            selected_key = None
            for key in keys:
                mapped_names = mapping.get(key, [])
                if sheet in mapped_names or sheet == name_mapping.get(key, key):
                    selected_key = key
                    break
            if selected_key is not None:
                # Use robust table extraction for better compatibility with different table formats
                # Create entity keywords that match the actual table content
                entity_keywords = [entity] + [f"{entity}{suffix}" for suffix in helpers]
                tables = extract_tables_from_worksheet_robust(excel_path, sheet, entity_keywords)
                
                if not tables:
                    st.info("No tables found in this sheet.")
                else:
                    st.success(f"Found {len(tables)} table(s) using robust detection")
                    
                    if enable_filtering:
                        matching_tables = []
                        for i, table_info in enumerate(tables):
                            try:
                                data = table_info['data']
                                method = table_info['method']
                                table_name = table_info['name']
                                
                                if not data or len(data) < 2:
                                    continue
                                
                                # Create DataFrame
                                df = pd.DataFrame(data[1:], columns=data[0])
                                df = df.dropna(how='all').dropna(axis=1, how='all')
                                
                                # Convert all columns to string to avoid ArrowTypeError with mixed data types
                                for col in df.columns:
                                    try:
                                        if df[col].dtype == 'object':
                                            df[col] = df[col].apply(lambda x: str(x) if pd.notna(x) else '')
                                        elif pd.api.types.is_datetime64_any_dtype(df[col]):
                                            df[col] = df[col].astype(str)
                                        elif pd.api.types.is_numeric_dtype(df[col]):
                                            df[col] = df[col].apply(lambda x: f"{x:,.2f}" if pd.notna(x) else '')
                                        else:
                                            df[col] = df[col].astype(str)
                                    except Exception as e:
                                        df[col] = df[col].apply(lambda x: str(x) if pd.notna(x) else '')
                                
                                # Drop Unnamed columns - handle mixed data types safely
                                df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed')]
                                df = df.reset_index(drop=True)
                                
                                # Check for entity keywords - handle mixed data types safely
                                # Include both the original data (which may contain headers) and the processed DataFrame
                                all_cells = [str(cell).lower().strip() for cell in df.values.flatten()]
                                # Also check the original data for headers/titles
                                original_cells = [str(cell).lower().strip() for cell in data[0]] if data and len(data) > 0 else []
                                all_cells.extend(original_cells)
                                
                                # Also check the table name itself for entity keywords
                                table_name_cells = [str(cell).lower().strip() for cell in table_name.split() if cell]
                                all_cells.extend(table_name_cells)
                                
                                # Check ALL rows of the original data for entity keywords
                                for row in data:
                                    row_cells = [str(cell).lower().strip() for cell in row if cell]
                                    all_cells.extend(row_cells)
                                
                                # Check for entity keywords in all collected cells
                                match_found = any(any(kw.lower() in cell for cell in all_cells) for kw in entity_keywords)
                                
                                if match_found:
                                    # Only append if DataFrame is not empty after dropna
                                    if not df.empty:
                                        matching_tables.append((i+1, method, df))
                                        
                            except Exception as e:
                                st.error(f"Error processing table {i+1}: {str(e)}")
                                continue
                        
                        # Only show matching tables
                        if matching_tables:
                            for table_num, method, df in matching_tables:
                                st.write(f"**Table {table_num}** (Detected by: {method})")
                                try:
                                    st.table(df)
                                except Exception as e:
                                    st.text(str(df.to_string()))
                        else:
                            st.info("No tables found matching the selected entity.")
                    else:
                        # Show all tables without filtering
                        for i, table_info in enumerate(tables):
                            try:
                                data = table_info['data']
                                method = table_info['method']
                                table_name = table_info['name']
                                
                                if not data or len(data) < 2:
                                    continue
                                
                                # Create DataFrame
                                df = pd.DataFrame(data[1:], columns=data[0])
                                df = df.dropna(how='all').dropna(axis=1, how='all')
                                
                                # Convert all columns to string
                                for col in df.columns:
                                    df[col] = df[col].astype(str)
                                
                                # Drop Unnamed columns - handle mixed data types safely
                                df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed')]
                                df = df.reset_index(drop=True)
                                
                                # In the unfiltered display block, add a check to skip empty DataFrames
                                if not df.empty:
                                    st.write(f"**Table {i+1}** (Detected by: {method})")
                                    try:
                                        st.table(df)
                                    except Exception as e:
                                        st.text(str(df.to_string()))
                                    
                            except Exception as e:
                                st.error(f"Error processing table {i+1}: {str(e)}")
                                continue
                
                # Store the first table for session state (for compatibility)
                if tables:
                    first_table_data = tables[0]['data']
                    if first_table_data and len(first_table_data) >= 2:
                        df_session = pd.DataFrame(first_table_data[1:], columns=first_table_data[0])
                        df_session = df_session.dropna(how='all').dropna(axis=1, how='all')
                        st.session_state['excel_tables'][sheet] = df_session
    # --- AI Generation Button ---
    st.subheader("Generate Report Text")
    if st.button("Generate Text (AI/Test)"):
        if use_ai:
            # Create progress bar for AI processing
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # Initialize progress tracking with more granular steps
            total_steps = len(keys) * 4  # 4 steps per key: generation, data validation, pattern validation, final QA
            current_step = 0
            
            try:
                status_text.text("🚀 Initializing AI services and loading data...")
                progress_bar.progress(0)
                
                # Step 1: Generate initial content
                status_text.text("🤖 Generating initial content with AI...")
                results = assistant.process_keys(
                    keys=keys,
                    entity_name=entity,
                    entity_helpers=helpers,
                    input_file=excel_path,
                    mapping_file='utils/mapping.json',
                    pattern_file='utils/pattern.json',
                    config_file='utils/config.json',
                    use_ai=use_ai,
                    convert_thousands=convert_thousands
                )
                
                current_step += len(keys)
                progress_bar.progress(current_step / total_steps)
                
                # Step 2: Data Accuracy Validation Agent
                status_text.text("🔍 Running data accuracy validation...")
                data_validation_agent = assistant.DataValidationAgent()
                validation_issues = []
                for i, key in enumerate(results):
                    status_text.text(f"📊 Validating data accuracy for {key} ({i+1}/{len(keys)})...")
                    try:
                        validation_result = data_validation_agent.validate_financial_data(
                            results[key], 
                            excel_path, 
                            entity, 
                            key
                        )
                        if validation_result['needs_correction']:
                            status_text.text(f"🔧 Correcting data accuracy issues for {key}...")
                            results[key] = data_validation_agent.correct_financial_data(
                                results[key], 
                                validation_result['issues']
                            )
                            validation_issues.append(f"{key}: {len(validation_result['issues'])} issues fixed")
                    except Exception as e:
                        st.warning(f"⚠️ Data validation failed for {key}: {str(e)}")
                    current_step += 1
                    progress_bar.progress(current_step / total_steps)
                
                # Step 3: Pattern Validation Agent
                status_text.text("📋 Running pattern compliance validation...")
                pattern_validation_agent = assistant.PatternValidationAgent()
                pattern_issues = []
                for i, key in enumerate(results):
                    status_text.text(f"📝 Validating pattern compliance for {key} ({i+1}/{len(keys)})...")
                    try:
                        pattern_result = pattern_validation_agent.validate_pattern_compliance(
                            results[key], 
                            key
                        )
                        if pattern_result['needs_correction']:
                            status_text.text(f"🔧 Correcting pattern compliance for {key}...")
                            results[key] = pattern_validation_agent.correct_pattern_compliance(
                                results[key], 
                                pattern_result['issues']
                            )
                            pattern_issues.append(f"{key}: {len(pattern_result['issues'])} issues fixed")
                    except Exception as e:
                        st.warning(f"⚠️ Pattern validation failed for {key}: {str(e)}")
                    current_step += 1
                    progress_bar.progress(current_step / total_steps)
                
                # Step 4: Final QA Review
                status_text.text("✨ Performing final quality assurance review...")
                qa_agent = assistant.QualityAssuranceAgent()
                qa_issues = []
                for i, key in enumerate(results):
                    status_text.text(f"🎯 Final QA review for {key} ({i+1}/{len(keys)})...")
                    try:
                        qa_result = qa_agent.validate_content(results[key])
                        if qa_result['score'] < 90:
                            results[key] = qa_agent.auto_correct(results[key])
                            qa_issues.append(f"{key}: QA score {qa_result['score']} improved")
                    except Exception as e:
                        st.warning(f"⚠️ QA review failed for {key}: {str(e)}")
                    current_step += 1
                    progress_bar.progress(current_step / total_steps)
                
                progress_bar.progress(1.0)
                status_text.text("🎉 AI processing completed successfully!")
                
                # Show summary of issues fixed
                if validation_issues or pattern_issues or qa_issues:
                    with st.expander("📋 Quality Assurance Summary", expanded=True):
                        if validation_issues:
                            st.write("**Data Accuracy Issues Fixed:**")
                            for issue in validation_issues:
                                st.write(f"• {issue}")
                        if pattern_issues:
                            st.write("**Pattern Compliance Issues Fixed:**")
                            for issue in pattern_issues:
                                st.write(f"• {issue}")
                        if qa_issues:
                            st.write("**QA Improvements:**")
                            for issue in qa_issues:
                                st.write(f"• {issue}")
                
                st.success("✅ AI processing completed with comprehensive validation and corrections!")
                
                # Save updated content back to markdown file
                try:
                    status_text.text("💾 Saving updated content to markdown file...")
                    save_results_to_markdown(results, entity)
                    st.success("✅ Content saved to markdown file!")
                except Exception as e:
                    st.warning(f"⚠️ Could not save to markdown file: {str(e)}")
                
            except Exception as e:
                st.error(f"❌ Error during AI processing: {str(e)}")
                status_text.text("Processing failed")
                # Show more detailed error information
                with st.expander("🔍 Error Details"):
                    st.code(str(e))
            finally:
                # Clear progress indicators after a delay
                import time
                time.sleep(3)
                progress_bar.empty()
                status_text.empty()
            
            st.session_state['results'] = results
        else:
            # Offline/test mode: split bs_content.md into sections for each mapped heading
            with open('utils/bs_content.md', 'r') as f:
                content = f.read()
            # Improved split: each section ends at the next heading of the same or higher level
            section_pattern = r'(### .+?)(?=\n### |\n## |\Z)'
            matches = re.findall(section_pattern, content, flags=re.DOTALL)
            section_map = {}
            for match in matches:
                heading_line = match.split('\n', 1)[0][4:].strip().lower()
                body = match.split('\n', 1)[1] if '\n' in match else ''
                section_map[heading_line] = body.strip('\n')
            # Map each key to its section using name_mapping, robust to case/whitespace
            mapped_results = {}
            for key in keys:
                heading = name_mapping.get(key, key).strip().lower()
                # Try exact match, then case-insensitive, then fallback
                section = section_map.get(heading)
                if section is None:
                    for h in section_map:
                        if h.replace(' ', '').lower() == heading.replace(' ', '').lower():
                            section = section_map[h]
                            break
                if not section:
                    section = f"No content found for section '{heading}'."
                mapped_results[key] = section
            st.session_state['results'] = mapped_results
    # --- Edit & Export Tabs (independent, manual mapping) ---
    if st.session_state['results']:
        st.subheader("Edit & Export")
        edit_tabs = st.tabs(valid_sheets)
        for i, tab in enumerate(edit_tabs):
            with tab:
                sheet = valid_sheets[i]
                # Manually map tab name (worksheet section) to the correct key using mapping.json and name_mapping
                mapped_key = None
                for key in keys:
                    mapped_names = mapping.get(key, [])
                    if sheet in mapped_names or sheet == name_mapping.get(key, key):
                        mapped_key = key
                        break
                if mapped_key is not None:
                    heading = name_mapping.get(mapped_key, mapped_key)
                    orig = st.session_state['results'].get(mapped_key, "")
                    edited = st.text_area(f"Edit text for {heading}", value=orig, height=150, key=f"edit_{mapped_key}")
                    if st.button(f"Save {heading}", key=f"save_{mapped_key}"):
                        # Update the section in bs_content.md
                        with open('utils/bs_content.md', 'r') as f:
                            content = f.read()
                        # Replace the section for this heading
                        pattern = re.compile(r'(###\s+' + re.escape(heading) + r'.*?)(?=\n###|\Z)', re.DOTALL | re.IGNORECASE)
                        new_section = f'### {heading}\n' + edited.strip() + '\n'
                        if pattern.search(content):
                            content = pattern.sub(new_section, content)
                        else:
                            content += '\n' + new_section
                        with open('utils/bs_content.md', 'w') as f:
                            f.write(content)
                        st.session_state['results'][mapped_key] = edited
                    # --- Export Button ---
                    if st.button("Export to PPTX", key=f"export_{mapped_key}"):
                        if template_path:
                            output_path = output_file_name
                            export_pptx(template_path, 'utils/bs_content.md', output_path, project_name=entity)
                            st.success(f"Exported to {output_path}")
                            with open(output_path, "rb") as f:
                                st.download_button("Download PPTX", data=f.read(), file_name=output_file_name)
                        else:
                            st.error("Please upload a PPTX template before exporting.")
    # Clean up temp file only after all processing is done
    if not st.session_state.get('results'):
        try:
            os.unlink(excel_path)
        except Exception:
            pass
else:
    st.info("Please upload an Excel databook to begin.") 
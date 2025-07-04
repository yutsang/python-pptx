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
                # Worksheet display/filtering code for 'sheet' only (no nested tabs)
                try:
                    df = xl.parse(sheet, header=0)
                    # If columns are Unnamed, use the next row as header
                    if any(str(col).startswith("Unnamed") for col in df.columns):
                        df = pd.read_excel(excel_path, sheet_name=sheet, header=1)
                except Exception as e:
                    # If parsing fails, try with different parameters
                    df = pd.read_excel(excel_path, sheet_name=sheet, header=None)
                    # Try to find a row that looks like headers
                    for j in range(min(5, len(df))):
                        if df.iloc[j].apply(lambda x: isinstance(x, str) and len(str(x)) > 0).sum() > len(df.columns) / 2:
                            df = pd.read_excel(excel_path, sheet_name=sheet, header=j)
                            break
                # Drop columns only if all values are NaN
                df = df.dropna(axis=1, how='all')
                # Convert all columns to string to avoid ArrowTypeError with mixed data types
                for col in df.columns:
                    try:
                        # Handle mixed data types more aggressively
                        if df[col].dtype == 'object':
                            # For object columns, check for mixed types and convert everything to string
                            df[col] = df[col].apply(lambda x: str(x) if pd.notna(x) else '')
                        elif pd.api.types.is_datetime64_any_dtype(df[col]):
                            df[col] = df[col].astype(str)
                        elif pd.api.types.is_numeric_dtype(df[col]):
                            # For numeric columns, convert to string but preserve formatting
                            df[col] = df[col].apply(lambda x: f"{x:,.2f}" if pd.notna(x) else '')
                        else:
                            # For all other types, convert to string
                            df[col] = df[col].astype(str)
                    except Exception as e:
                        # Ultimate fallback: convert everything to string
                        df[col] = df[col].apply(lambda x: str(x) if pd.notna(x) else '')
                # Optionally drop Unnamed columns if not needed
                df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
                # Split dataframes on empty rows and filter by entity name (from original utils.py)
                empty_rows = df.index[df.isnull().all(1)]
                start_idx = 0
                dataframes = []
                # Split on empty rows
                for end_idx in empty_rows:
                    if end_idx > start_idx:
                        split_df = df[start_idx:end_idx]
                        if not split_df.dropna(how='all').empty:
                            dataframes.append(split_df)
                        start_idx = end_idx + 1
                if start_idx < len(df):
                    dataframes.append(df[start_idx:])
                if not dataframes:
                    st.info("No tables found in this sheet.")
                else:
                    if enable_filtering:
                        entity_keywords = [f"{entity}{suffix}" for suffix in helpers]
                        for k, data_frame in enumerate(dataframes):
                            df_str = data_frame.astype(str).reset_index(drop=True)
                            header_indices = []
                            for idx, row in df_str.iterrows():
                                row_str = ' '.join(row.values).lower()
                                for pattern in entity_keywords:
                                    if pattern.lower() in row_str:
                                        header_indices.append(idx)
                                        break
                            for j, header_idx in enumerate(header_indices):
                                next_header = header_indices[j+1] if j+1 < len(header_indices) else len(df_str)
                                empty_after = df_str.iloc[header_idx+1:next_header].index[df_str.iloc[header_idx+1:next_header].apply(lambda r: all(str(cell).strip() == '' for cell in r), axis=1)]
                                section_end = empty_after[0] if len(empty_after) > 0 else next_header
                                section = df_str.iloc[header_idx:section_end]
                                section = section[~section.apply(lambda r: all(str(cell).strip() == '' for cell in r), axis=1)]
                                if len(section) > 0:
                                    # Set the first row as the new header, rest as data
                                    new_header = section.iloc[0].values.tolist()
                                    section_data = section.iloc[1:].copy()
                                    # Ensure all rows have the same number of columns as the header
                                    if section_data.shape[1] != len(new_header):
                                        section_data = section_data.reindex(columns=range(len(new_header)), fill_value='')
                                    section_data.columns = new_header
                                    section_data = section_data.reset_index(drop=True)
                                    try:
                                        st.table(section_data)
                                    except Exception as e:
                                        st.text(str(section_data.to_string()))
                    else:
                        for k, data_frame in enumerate(dataframes):
                            df_display = data_frame.copy()
                            for col in df_display.columns:
                                df_display[col] = df_display[col].astype(str)
                            try:
                                st.table(df_display)
                            except Exception as e:
                                st.text(str(df_display.to_string()))
                st.session_state['excel_tables'][sheet] = df
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
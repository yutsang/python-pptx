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
helpers = st.sidebar.text_input("Entity Helpers (comma separated)", value="Wanpu, Limited, ...", placeholder="Wanpu, Limited, ...").split(",")
helpers = [h.strip() for h in helpers if h.strip()]

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
    # Show tabs for each worksheet (using mapping.json and name_mapping)
    st.subheader("View Worksheets")
    # Only show worksheet tabs whose names match the mapped section headings
    valid_sheets = [sheet for sheet in xl.sheet_names if sheet in mapped_sheet_names or sheet in name_mapping.values()]
    # Use session state to synchronize selected tab
    if 'selected_tab' not in st.session_state:
        st.session_state['selected_tab'] = 0
    tabs = st.tabs(valid_sheets)
    for i, sheet in enumerate(valid_sheets):
        with tabs[i]:
            # Remove the select button (no longer needed)
            df = xl.parse(sheet, header=0)
            # If columns are Unnamed, use the next row as header
            if any(str(col).startswith("Unnamed") for col in df.columns):
                df = pd.read_excel(excel_path, sheet_name=sheet, header=1)
            # Drop columns only if all values are NaN
            df = df.dropna(axis=1, how='all')
            # Convert datetime columns to string to avoid ArrowTypeError (robust for object columns)
            for col in df.columns:
                # Convert if the dtype is datetime-like
                if pd.api.types.is_datetime64_any_dtype(df[col]):
                    df[col] = df[col].astype(str)
                # Convert if the column is object and contains any datetime-like objects
                elif df[col].dtype == 'object':
                    if df[col].apply(lambda x: isinstance(x, (pd.Timestamp, np.datetime64, datetime.datetime, datetime.date))).any():
                        df[col] = df[col].apply(lambda x: str(x) if isinstance(x, (pd.Timestamp, np.datetime64, datetime.datetime, datetime.date)) else x)
            # Optionally drop Unnamed columns if not needed
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
            st.dataframe(df, use_container_width=True)
            st.session_state['excel_tables'][sheet] = df
    # --- AI Generation Button ---
    st.subheader("Generate Report Text")
    if st.button("Generate Text (AI/Test)"):
        if not use_ai:
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
        else:
            # Run the assistant.process_keys logic and use second AI agent to review
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
            # Use second AI agent to review each result
            qa_agent = assistant.QualityAssuranceAgent()
            for key in results:
                qa_result = qa_agent.validate_content(results[key])
                if qa_result['score'] < 90:
                    results[key] = qa_agent.auto_correct(results[key])
            st.session_state['results'] = results
    # --- Show Editable Results ---
    if st.session_state['results']:
        st.subheader("Edit & Export")
        edited_results = {}
        # Synchronize tabs with View Worksheets
        tabs2 = st.tabs(keys)
        for i, key in enumerate(keys):
            with tabs2[i]:
                # Use name_mapping to get the correct section heading
                heading = name_mapping.get(key, key)
                orig = st.session_state['results'].get(key, "")
                edited = st.text_area(f"Edit text for {heading}", value=orig, height=150, key=f"edit_{key}")
                if st.button(f"Save {heading}", key=f"save_{key}"):
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
                    st.session_state['results'][key] = edited
                edited_results[key] = edited
        # --- Export Button ---
        if st.button("Export to PPTX"):
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
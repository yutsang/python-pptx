import streamlit as st
import pandas as pd
from common import assistant
import tempfile
import os

st.set_page_config(page_title="Due Diligence Report Generator", layout="wide")
st.title("Real Estate Due Diligence Report Generator")

# --- Upload Excel File ---
st.sidebar.header("Step 1: Upload Excel Databook")
uploaded_file = st.sidebar.file_uploader("Choose an Excel file", type=["xlsx"]) 

# --- Entity and Helpers ---
st.sidebar.header("Step 2: Entity & Helpers")
entity = st.sidebar.selectbox("Select Entity", ["Haining", "Nanjing", "Ningbo"])
helpers = st.sidebar.text_input("Entity Helpers (comma separated)", "").split(",")
helpers = [h.strip() for h in helpers if h.strip()]

# --- AI or Local Mode ---
st.sidebar.header("Step 3: Generation Mode")
use_ai = st.sidebar.radio("Use AI for text generation?", ("No (Test Mode)", "Yes (AI Mode)")) == "Yes (AI Mode)"

# --- State ---
if 'results' not in st.session_state:
    st.session_state['results'] = None
if 'keys' not in st.session_state:
    st.session_state['keys'] = []
if 'excel_tables' not in st.session_state:
    st.session_state['excel_tables'] = {}

# --- Main Workflow ---
if uploaded_file:
    # Save uploaded file to a temp location
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        tmp.write(uploaded_file.read())
        excel_path = tmp.name
    xl = pd.ExcelFile(excel_path)
    st.success(f"Excel file loaded: {uploaded_file.name}")
    # Show tabs for each worksheet
    st.subheader("Step 4: View Worksheets")
    tabs = st.tabs(xl.sheet_names)
    for i, sheet in enumerate(xl.sheet_names):
        with tabs[i]:
            df = xl.parse(sheet)
            st.dataframe(df)
            st.session_state['excel_tables'][sheet] = df
    # Define keys (can be made dynamic)
    keys = [
        "Cash", "AR", "Prepayments", "OR", "Other CA", "IP", "Other NCA",
        "AP", "Taxes payable", "OP", "Capital", "Reserve"
    ]
    if entity in ['Ningbo', 'Nanjing']:
        keys = [key for key in keys if key != "Reserve"]
    st.session_state['keys'] = keys
    # --- AI Generation Button ---
    st.subheader("Step 5: Generate Report Text")
    if st.button("Generate Text (AI/Test)"):
        # Run the assistant.process_keys logic
        results = assistant.process_keys(
            keys=keys,
            entity_name=entity,
            entity_helpers=helpers,
            input_file=excel_path,
            mapping_file='utils/mapping.json',
            pattern_file='utils/pattern.json',
            config_file='utils/config.json',
            use_ai=use_ai
        )
        st.session_state['results'] = results
    # --- Show Editable Results ---
    if st.session_state['results']:
        st.subheader("Step 6: Edit & Export")
        edited_results = {}
        tabs2 = st.tabs(keys)
        for i, key in enumerate(keys):
            with tabs2[i]:
                orig = st.session_state['results'].get(key, "")
                edited = st.text_area(f"Edit text for {key}", value=orig, height=150)
                edited_results[key] = edited
        # --- Export Button ---
        if st.button("Export to PPTX"):
            # Placeholder: implement export_to_pptx in assistant.py
            st.success("Export to PPTX (functionality to be implemented)")
            # pptx_path = assistant.export_to_pptx(edited_results, f"report_{entity}.pptx")
            # st.download_button("Download PPTX", data=open(pptx_path, 'rb').read(), file_name=f"report_{entity}.pptx")
    # Clean up temp file
    os.unlink(excel_path)
else:
    st.info("Please upload an Excel databook to begin.") 
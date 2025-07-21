#!/usr/bin/env python3
"""
Financial Data Processor - Due Diligence Automation

Self-contained version that works exactly like the original but uses config/ directory.
"""

import streamlit as st
import pandas as pd
import json
import os
import re
from pathlib import Path
from tabulate import tabulate

def load_config_files():
    """Load configuration files from the config directory."""
    try:
        config_dir = Path("config")
        
        # Load mapping.json
        with open(config_dir / "mapping.json", 'r') as f:
            mapping = json.load(f)
        
        # Load pattern.json  
        with open(config_dir / "pattern.json", 'r') as f:
            pattern = json.load(f)
            
        # Load config.json
        with open(config_dir / "config.json", 'r') as f:
            config = json.load(f)
            
        # Load prompts.json
        with open(config_dir / "prompts.json", 'r') as f:
            prompts = json.load(f)
            
        return config, mapping, pattern, prompts
        
    except FileNotFoundError as e:
        st.error(f"Configuration file not found: {e}")
        return None, None, None, None
    except json.JSONDecodeError as e:
        st.error(f"Invalid JSON in configuration file: {e}")
        return None, None, None, None

def process_and_filter_excel(filename, tab_name_mapping, entity_name, entity_suffixes):
    """
    Process and filter Excel file to extract relevant worksheet sections
    This matches the original processing logic exactly
    """
    try:
        # Load the Excel file
        xl = pd.ExcelFile(filename)
        
        # Create a reverse mapping from values to keys
        reverse_mapping = {}
        for key, values in tab_name_mapping.items():
            for value in values:
                reverse_mapping[value] = key
                
        # Initialize a string to store markdown content
        markdown_content = ""
        
        # Process each sheet according to the mapping
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
                
                # Filter dataframes by entity name with proper spacing
                entity_keywords = [f"{entity_name} {suffix}" for suffix in entity_suffixes if suffix]
                if not entity_keywords:  # If no helpers, just use entity name
                    entity_keywords = [entity_name]
                
                combined_pattern = '|'.join(re.escape(kw) for kw in entity_keywords)
                
                for data_frame in dataframes:
                    mask = data_frame.apply(
                        lambda row: row.astype(str).str.contains(
                            combined_pattern, case=False, regex=True, na=False
                        ).any(),
                        axis=1
                    )
                    if mask.any():
                        markdown_content += tabulate(data_frame, headers='keys', tablefmt='pipe') + '\n\n'
                    
                    if any(data_frame.apply(lambda row: row.astype(str).str.contains(keyword, case=False, na=False).any(), axis=1).any() for keyword in entity_keywords):
                        markdown_content += tabulate(data_frame, headers='keys', tablefmt='pipe', showindex=False)
                        markdown_content += "\n\n" 
        
        return markdown_content
    except Exception as e:
        st.error(f"An error occurred while processing the Excel file: {e}")
        return ""

def main():
    """Main application - matches original UI style"""
    st.set_page_config(
        page_title="Financial Data Processor",
        page_icon="📊",
        layout="wide"
    )
    st.title("📊 Financial Data Processor")
    st.markdown("---")

    # Load configuration files
    config, mapping, pattern, prompts = load_config_files()
    
    if not all([config, mapping, pattern, prompts]):
        st.error("❌ Configuration files not available. Please ensure config/ directory has all required files.")
        return

    # Sidebar for controls (matches original)
    with st.sidebar:
        uploaded_file = st.file_uploader(
            "Upload Excel File",
            type=['xlsx', 'xls'],
            help="Upload your financial data Excel file"
        )
        
        entity_options = ['Haining', 'Nanjing', 'Ningbo']
        selected_entity = st.selectbox(
            "Select Entity",
            options=entity_options,
            help="Choose the entity for data processing"
        )
        
        # Entity helpers (matches original default)
        entity_helpers = st.text_input(
            "Entity Helpers",
            value="Wanpu,Limited,",
            help="Comma-separated entity keywords"
        )
        
        # Financial Statement Type Selection (matches original)
        st.markdown("---")
        statement_type_options = ["Balance Sheet", "Income Statement", "All"]
        statement_type_display = st.radio(
            "Financial Statement Type",
            options=statement_type_options,
            index=0,
            help="Select the type of financial statement to process"
        )
        
        # Convert display names to internal codes
        statement_type_map = {
            "Balance Sheet": "BS",
            "Income Statement": "IS", 
            "All": "ALL"
        }
        statement_type = statement_type_map[statement_type_display]

    # Main content area
    if uploaded_file is not None:
        st.success(f"✅ File uploaded: {uploaded_file.name}")
        st.info(f"🏢 Processing for entity: **{selected_entity}**")
        st.info(f"📊 Statement type: **{statement_type_display}**")
        
        if st.button("🚀 Process Data", type="primary"):
            with st.spinner("Processing Excel file..."):
                try:
                    # Parse entity helpers
                    entity_suffixes = [s.strip() for s in entity_helpers.split(',') if s.strip()]
                    
                    # Save uploaded file temporarily
                    temp_file_path = f"temp_{uploaded_file.name}"
                    with open(temp_file_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                    
                    # Process data using the corrected function
                    result = process_and_filter_excel(
                        temp_file_path,
                        mapping,
                        selected_entity,
                        entity_suffixes
                    )
                    
                    # Clean up temp file
                    os.remove(temp_file_path)
                    
                    if result and result.strip():
                        st.success("✅ Processing completed!")
                        
                        # Show results in expandable section
                        with st.expander("📊 Processing Results", expanded=True):
                            st.markdown(result)
                            
                        # Download option
                        st.download_button(
                            label="📥 Download Results",
                            data=result,
                            file_name=f"{selected_entity}_processing_results.md",
                            mime="text/markdown"
                        )
                        
                    else:
                        st.warning("⚠️ No data found for the selected entity and configuration.")
                        st.info("💡 Try adjusting the entity helpers or check if the Excel file contains the expected sheet names.")
                        
                except Exception as e:
                    st.error(f"❌ Processing failed: {str(e)}")
                    # Clean up temp file on error
                    temp_file_path = f"temp_{uploaded_file.name}"
                    if os.path.exists(temp_file_path):
                        os.remove(temp_file_path)
    else:
        st.info("📁 Please upload an Excel file to begin processing.")
        
        # Show configuration status
        with st.expander("⚙️ Configuration Status"):
            st.write("**Loaded Configuration Files:**")
            st.write(f"- mapping.json: {len(mapping) if mapping else 0} entities")
            st.write(f"- pattern.json: {len(pattern) if pattern else 0} patterns") 
            st.write(f"- config.json: ✅ Available")
            st.write(f"- prompts.json: ✅ Available")

if __name__ == "__main__":
    main() 
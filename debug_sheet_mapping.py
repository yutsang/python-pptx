#!/usr/bin/env python3
"""
Debug script to investigate sheet mapping issues for AR
"""

import pandas as pd
import json
from pathlib import Path
from fdd_utils.excel_processing import get_worksheet_sections_by_keys
from fdd_utils.data_utils import load_config_files

def debug_sheet_mapping():
    """Debug the sheet mapping for AR specifically"""
    
    # Load configuration
    config, mapping, pattern, prompts = load_config_files()
    
    print("=== SHEET MAPPING DEBUG ===")
    print(f"Mapping keys: {list(mapping.keys())}")
    print(f"AR mapping: {mapping.get('AR', 'NOT FOUND')}")
    print(f"Taxes payable mapping: {mapping.get('Taxes payable', 'NOT FOUND')}")
    print(f"Tax and Surcharges mapping: {mapping.get('Tax and Surcharges', 'NOT FOUND')}")
    
    # Test with Excel file
    excel_file = "databook.xlsx"
    if not Path(excel_file).exists():
        print(f"❌ Excel file {excel_file} not found")
        return
    
    print(f"\n=== PROCESSING EXCEL FILE ===")
    
    # Get sections
    sections_by_key = get_worksheet_sections_by_keys(
        uploaded_file=excel_file,
        tab_name_mapping=mapping,
        entity_name="Company",
        entity_suffixes=["Haining Wanpu"],
        debug=True
    )
    
    print(f"\n=== RESULTS ===")
    print(f"Available keys: {list(sections_by_key.keys())}")
    
    # Check AR specifically
    if 'AR' in sections_by_key:
        ar_sections = sections_by_key['AR']
        print(f"\nAR has {len(ar_sections)} sections:")
        
        for i, section in enumerate(ar_sections):
            print(f"\n--- AR Section {i} ---")
            print(f"Sheet: {section.get('sheet', 'Unknown')}")
            
            if 'parsed_data' in section and section['parsed_data']:
                metadata = section['parsed_data']['metadata']
                print(f"Table name: {metadata.get('table_name', 'Unknown')}")
                print(f"Sheet name: {metadata.get('sheet_name', 'Unknown')}")
                print(f"Date: {metadata.get('date', 'None')}")
                print(f"Value column: {metadata.get('value_column', 'Unknown')}")
                
                # Show first few data rows
                data_rows = section['parsed_data']['data']
                print(f"Data rows: {len(data_rows)}")
                if data_rows:
                    print("First 3 rows:")
                    for j, row in enumerate(data_rows[:3]):
                        print(f"  Row {j}: {row['description']} = {row['value']}")
            else:
                print("No parsed_data available")
    else:
        print("❌ AR not found in sections_by_key")
    
    # Check what sheets are actually in the Excel file
    print(f"\n=== EXCEL SHEETS ===")
    with pd.ExcelFile(excel_file) as xl:
        print(f"Available sheets: {xl.sheet_names}")
        
        # Check each sheet for AR-related content
        for sheet_name in xl.sheet_names:
            df = xl.parse(sheet_name)
            all_text = ' '.join(df.astype(str).values.flatten()).lower()
            
            if 'ar' in all_text or 'accounts receivable' in all_text:
                print(f"\n--- Sheet: {sheet_name} (contains AR content) ---")
                print(f"Shape: {df.shape}")
                print(f"Columns: {list(df.columns)}")
                print(f"First few rows:")
                print(df.head(3))

if __name__ == "__main__":
    debug_sheet_mapping()

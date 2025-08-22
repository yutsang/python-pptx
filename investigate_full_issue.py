#!/usr/bin/env python3
"""
Debug script to investigate the issue where command line logging shows correct column selection
but Streamlit view shows "no data found"
"""

import pandas as pd
import json
import os
from pathlib import Path
from fdd_utils.excel_processing import get_worksheet_sections_by_keys, detect_latest_date_column
from fdd_utils.data_utils import get_financial_keys, load_config_files
from fdd_utils.mappings import KEY_TERMS_BY_KEY

def debug_data_flow():
    """Debug the data flow from Excel processing to sections_by_key"""
    
    print("🔍 DEBUGGING DATA FLOW ISSUE")
    print("=" * 50)
    
    # Check if databook.xlsx exists
    excel_file = "databook.xlsx"
    if not os.path.exists(excel_file):
        print(f"❌ {excel_file} not found in current directory")
        return
    
    print(f"✅ Found {excel_file}")
    
    # Load configuration files
    try:
        config, mapping, pattern, prompts = load_config_files()
        print(f"✅ Loaded config files - mapping keys: {list(mapping.keys()) if mapping else 'None'}")
    except Exception as e:
        print(f"❌ Failed to load config files: {e}")
        return
    
    # Test entity and suffixes
    test_entity = "Test Entity"  # Replace with actual entity name
    entity_suffixes = ["Inc", "LLC", "Corp"]  # Replace with actual suffixes
    
    print(f"🔍 Testing with entity: {test_entity}")
    print(f"🔍 Entity suffixes: {entity_suffixes}")
    
    # Test the get_worksheet_sections_by_keys function
    print("\n📊 Testing get_worksheet_sections_by_keys...")
    try:
        sections_by_key = get_worksheet_sections_by_keys(
            uploaded_file=excel_file,
            tab_name_mapping=mapping,
            entity_name=test_entity,
            entity_suffixes=entity_suffixes,
            debug=True
        )
        
        print(f"✅ get_worksheet_sections_by_keys completed")
        print(f"📊 Sections by key results:")
        
        total_sections = 0
        for key, sections in sections_by_key.items():
            if sections:
                print(f"  ✅ {key}: {len(sections)} sections")
                total_sections += len(sections)
            else:
                print(f"  ❌ {key}: 0 sections")
        
        print(f"\n📈 Total sections found: {total_sections}")
        
        if total_sections == 0:
            print("\n🔍 INVESTIGATING WHY NO SECTIONS FOUND:")
            
            # Check what sheets are available
            with pd.ExcelFile(excel_file) as xl:
                print(f"📋 Available sheets: {xl.sheet_names}")
                
                # Check each sheet for entity keywords
                for sheet_name in xl.sheet_names:
                    print(f"\n🔍 Analyzing sheet: {sheet_name}")
                    df = xl.parse(sheet_name)
                    print(f"  📊 DataFrame shape: {df.shape}")
                    print(f"  📋 Columns: {list(df.columns)}")
                    
                    # Check for entity keywords in the sheet
                    entity_keywords = [f"{test_entity} {suffix}" for suffix in entity_suffixes if suffix]
                    if not entity_keywords:
                        entity_keywords = [test_entity]
                    
                    print(f"  🔍 Looking for entity keywords: {entity_keywords}")
                    
                    # Check if any entity keywords are in the sheet
                    all_text = ' '.join(df.astype(str).values.flatten()).lower()
                    found_entities = []
                    for keyword in entity_keywords:
                        if keyword.lower() in all_text:
                            found_entities.append(keyword)
                    
                    if found_entities:
                        print(f"  ✅ Found entities: {found_entities}")
                    else:
                        print(f"  ❌ No entities found in this sheet")
                    
                    # Check for financial keys
                    financial_keys = get_financial_keys()
                    found_keys = []
                    for key in financial_keys:
                        key_terms = KEY_TERMS_BY_KEY.get(key, [key.lower()])
                        for term in key_terms:
                            if term.lower() in all_text:
                                found_keys.append(key)
                                break
                    
                    if found_keys:
                        print(f"  ✅ Found financial keys: {found_keys}")
                    else:
                        print(f"  ❌ No financial keys found in this sheet")
                    
                    # Test date column detection
                    print(f"  📅 Testing date column detection...")
                    latest_date_col = detect_latest_date_column(df, sheet_name, entity_keywords)
                    if latest_date_col:
                        print(f"  ✅ Latest date column: {latest_date_col}")
                    else:
                        print(f"  ❌ No date column detected")
        
    except Exception as e:
        print(f"❌ Error in get_worksheet_sections_by_keys: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    debug_data_flow() 
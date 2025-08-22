#!/usr/bin/env python3
"""
Focused debug script to check data structure and identify why sections_by_key is empty
"""

import pandas as pd
import json
import os
from pathlib import Path
from fdd_utils.excel_processing import get_worksheet_sections_by_keys
from fdd_utils.data_utils import get_financial_keys, load_config_files

def debug_sections_structure():
    """Debug the sections_by_key structure specifically"""
    
    print("🔍 DEBUGGING SECTIONS_BY_KEY STRUCTURE")
    print("=" * 50)
    
    # Check if databook.xlsx exists
    excel_file = "databook.xlsx"
    if not os.path.exists(excel_file):
        print(f"❌ {excel_file} not found")
        return
    
    # Load configuration files
    try:
        config, mapping, pattern, prompts = load_config_files()
        print(f"✅ Config loaded - mapping has {len(mapping) if mapping else 0} keys")
    except Exception as e:
        print(f"❌ Config load failed: {e}")
        return
    
    # Test with a real entity name (you'll need to replace this with actual entity from your data)
    # Let's try to find entities in the Excel file first
    print("\n🔍 SCANNING FOR ENTITIES IN EXCEL FILE...")
    
    with pd.ExcelFile(excel_file) as xl:
        print(f"📋 Available sheets: {xl.sheet_names}")
        
        # Scan for potential entity names
        potential_entities = set()
        for sheet_name in xl.sheet_names:
            df = xl.parse(sheet_name)
            # Look for common entity patterns in the first few rows
            for i in range(min(10, len(df))):
                for col in df.columns:
                    cell_value = str(df.iloc[i, df.columns.get_loc(col)])
                    # Look for patterns that might be entity names
                    if any(keyword in cell_value.lower() for keyword in ['inc', 'llc', 'corp', 'company', 'ltd']):
                        # Extract potential entity name
                        words = cell_value.split()
                        for word in words:
                            if any(keyword in word.lower() for keyword in ['inc', 'llc', 'corp', 'company', 'ltd']):
                                potential_entities.add(word)
        
        print(f"🔍 Potential entities found: {list(potential_entities)}")
    
    # Test with different entity names
    test_entities = list(potential_entities)[:3] if potential_entities else ["Test Company Inc"]
    
    for test_entity in test_entities:
        print(f"\n🧪 TESTING WITH ENTITY: {test_entity}")
        print("-" * 40)
        
        try:
            # Test with different entity suffixes
            entity_suffixes_list = [
                [],  # No suffixes
                ["Inc", "LLC"],  # Common suffixes
                ["Inc", "LLC", "Corp", "Company"]  # Extended suffixes
            ]
            
            for entity_suffixes in entity_suffixes_list:
                print(f"  🔍 Testing suffixes: {entity_suffixes}")
                
                sections_by_key = get_worksheet_sections_by_keys(
                    uploaded_file=excel_file,
                    tab_name_mapping=mapping,
                    entity_name=test_entity,
                    entity_suffixes=entity_suffixes,
                    debug=True
                )
                
                # Count sections
                total_sections = sum(len(sections) for sections in sections_by_key.values())
                print(f"    📊 Total sections: {total_sections}")
                
                if total_sections > 0:
                    print(f"    ✅ SUCCESS! Found {total_sections} sections")
                    print(f"    📋 Sections breakdown:")
                    for key, sections in sections_by_key.items():
                        if sections:
                            print(f"      - {key}: {len(sections)} sections")
                            # Show first section structure
                            if sections:
                                first_section = sections[0]
                                print(f"        First section keys: {list(first_section.keys())}")
                                if 'parsed_data' in first_section:
                                    print(f"        Has parsed_data: ✅")
                                if 'data' in first_section:
                                    print(f"        Has raw data: ✅ (shape: {first_section['data'].shape})")
                    break
                else:
                    print(f"    ❌ No sections found with these suffixes")
        
        except Exception as e:
            print(f"    ❌ Error: {e}")
    
    print(f"\n🎯 RECOMMENDATIONS:")
    print(f"1. Check if the entity name in your Streamlit app matches entities in the Excel file")
    print(f"2. Verify that entity suffixes are correctly configured")
    print(f"3. Ensure the Excel file contains the expected financial data")
    print(f"4. Check that the mapping configuration matches the sheet names")

if __name__ == "__main__":
    debug_sections_structure() 
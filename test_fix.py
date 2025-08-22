#!/usr/bin/env python3
"""
Test script to verify the fix for entity name mismatch issue
"""

import pandas as pd
import os
from fdd_utils.excel_processing import get_worksheet_sections_by_keys
from fdd_utils.data_utils import load_config_files

def test_entity_matching():
    """Test different entity names to find the correct one"""
    
    print("🔍 TESTING ENTITY NAME MATCHING")
    print("=" * 50)
    
    # Check if databook.xlsx exists
    excel_file = "databook.xlsx"
    if not os.path.exists(excel_file):
        print(f"❌ {excel_file} not found")
        return
    
    # Load configuration files
    try:
        config, mapping, pattern, prompts = load_config_files()
        print(f"✅ Config loaded")
    except Exception as e:
        print(f"❌ Config load failed: {e}")
        return
    
    # Test different entity names that might be used in the app
    test_entities = [
        "Company",  # What we found in the debug
        "Haining",  # From the app code
        "Nanjing",  # From the app code  
        "Ningbo",   # From the app code
        "Test Company Inc",  # Generic test
        "Test Entity"  # Another generic test
    ]
    
    print(f"🧪 Testing {len(test_entities)} different entity names...")
    
    for entity_name in test_entities:
        print(f"\n🔍 Testing entity: '{entity_name}'")
        print("-" * 40)
        
        try:
            # Test with no suffixes first
            sections_by_key = get_worksheet_sections_by_keys(
                uploaded_file=excel_file,
                tab_name_mapping=mapping,
                entity_name=entity_name,
                entity_suffixes=[],
                debug=False  # Reduce output
            )
            
            # Count sections
            total_sections = sum(len(sections) for sections in sections_by_key.values())
            print(f"  📊 Total sections: {total_sections}")
            
            if total_sections > 0:
                print(f"  ✅ SUCCESS! Found {total_sections} sections with '{entity_name}'")
                print(f"  📋 Keys with data:")
                for key, sections in sections_by_key.items():
                    if sections:
                        print(f"    - {key}: {len(sections)} sections")
                break
            else:
                print(f"  ❌ No sections found with '{entity_name}'")
                
                # Try with common suffixes
                entity_suffixes = ["Inc", "LLC", "Corp", "Company"]
                sections_by_key = get_worksheet_sections_by_keys(
                    uploaded_file=excel_file,
                    tab_name_mapping=mapping,
                    entity_name=entity_name,
                    entity_suffixes=entity_suffixes,
                    debug=False
                )
                
                total_sections = sum(len(sections) for sections in sections_by_key.values())
                if total_sections > 0:
                    print(f"  ✅ SUCCESS with suffixes! Found {total_sections} sections")
                    break
                else:
                    print(f"  ❌ Still no sections found with suffixes")
        
        except Exception as e:
            print(f"  ❌ Error: {e}")
    
    print(f"\n🎯 SOLUTION:")
    print(f"1. The entity name in your Streamlit app must match what's in the Excel file")
    print(f"2. Based on the debug output, 'Company' works and finds 77 sections")
    print(f"3. If your app uses 'Haining', 'Nanjing', or 'Ningbo', these don't match the Excel data")
    print(f"4. You need to either:")
    print(f"   a) Change the entity name in your app to 'Company'")
    print(f"   b) Update the Excel file to use the correct entity names")
    print(f"   c) Add entity suffixes to match the Excel data")

if __name__ == "__main__":
    test_entity_matching() 
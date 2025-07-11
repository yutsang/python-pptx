#!/usr/bin/env python3
"""
Debug script to diagnose why no content is generated for specific keys
"""

import pandas as pd
import json
import re
from pathlib import Path
from tabulate import tabulate
import openpyxl

def debug_excel_processing(filename, entity_name, entity_suffixes, key_to_debug=None):
    """Debug the Excel processing to see what data is being extracted"""
    
    print(f"🔍 DEBUGGING DATA EXTRACTION")
    print(f"File: {filename}")
    print(f"Entity: {entity_name}")
    print(f"Entity Suffixes: {entity_suffixes}")
    print("=" * 60)
    
    # Load mapping
    try:
        with open('utils/mapping.json', 'r') as f:
            mapping = json.load(f)
    except FileNotFoundError:
        print("❌ utils/mapping.json not found")
        return
    
    # Load the Excel file
    file_path = Path(filename)
    if not file_path.exists():
        print(f"❌ File {filename} not found")
        return
    
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True, keep_links=False)
    except Exception as e:
        print(f"❌ Error loading workbook: {e}")
        return
    
    # Prepare entity keywords
    entity_keywords = [entity_name] + list(entity_suffixes)
    entity_keywords = [kw.strip().lower() for kw in entity_keywords if kw.strip()]
    print(f"Entity keywords: {entity_keywords}")
    print()
    
    # Debug specific key or all keys
    keys_to_debug = [key_to_debug] if key_to_debug else mapping.keys()
    
    for key in keys_to_debug:
        print(f"🔍 DEBUGGING KEY: {key}")
        print("-" * 40)
        
        # Get tab name mapping for this key
        if key not in mapping:
            print(f"❌ Key '{key}' not found in mapping")
            continue
            
        tab_name_mapping = {key: mapping[key]}
        print(f"Tab mapping: {tab_name_mapping}")
        
        # Process each worksheet
        markdown_content = ""
        tables_found = 0
        matches_found = 0
        
        for ws in wb.worksheets:
            if ws.title not in mapping[key]:
                continue
                
            print(f"\n📊 Processing worksheet: {ws.title}")
            
            # Extract tables from worksheet
            from common.assistant import extract_tables_robust
            tables = extract_tables_robust(ws, entity_keywords)
            
            print(f"   Tables extracted: {len(tables)}")
            
            for i, table_info in enumerate(tables):
                tables_found += 1
                print(f"\n   📋 Table {i+1}: {table_info['name']}")
                
                data = table_info['data']
                if not data or len(data) < 2:
                    print("   ⚠️  No data or insufficient rows")
                    continue
                
                # Create DataFrame
                df = pd.DataFrame(data[1:], columns=data[0])
                df = df.dropna(how='all').dropna(axis=1, how='all')
                df = df.map(lambda x: str(x) if x is not None else "")
                df = df.reset_index(drop=True)
                
                print(f"   📐 DataFrame shape: {df.shape}")
                
                # Show first few rows for debugging
                print("   📝 First 3 rows of data:")
                print(df.head(3).to_string())
                
                # Test both old and new matching methods
                print(f"\n   🔍 Testing entity keyword matching:")
                
                # Method 1 (NEW - optimized): Vectorized approach
                df_str = df.astype(str).apply(lambda x: ' '.join(x), axis=1).str.lower()
                new_matches = []
                for kw in entity_keywords:
                    matches = df_str.str.contains(kw.lower(), regex=False, na=False)
                    if matches.any():
                        new_matches.append(kw)
                        print(f"     ✅ NEW method found: '{kw}'")
                
                # Method 2 (OLD - original): Row-by-row approach
                old_matches = []
                for kw in entity_keywords:
                    for idx, row in df.iterrows():
                        row_found = False
                        for cell in row:
                            if kw.lower() in str(cell).lower():
                                old_matches.append(kw)
                                row_found = True
                                break
                        if row_found:
                            break
                    if kw in old_matches:
                        print(f"     ✅ OLD method found: '{kw}'")
                
                # Compare methods
                new_match_found = len(new_matches) > 0
                old_match_found = len(old_matches) > 0
                
                if new_match_found != old_match_found:
                    print(f"     ⚠️  MISMATCH! NEW: {new_match_found}, OLD: {old_match_found}")
                    print(f"     NEW matches: {new_matches}")
                    print(f"     OLD matches: {old_matches}")
                    
                    # Show the concatenated strings for debugging
                    print(f"     📄 Sample concatenated strings:")
                    for idx in range(min(3, len(df_str))):
                        print(f"       Row {idx}: '{df_str.iloc[idx][:100]}...'")
                
                if new_match_found:
                    matches_found += 1
                    # Add to markdown content (using new method for consistency)
                    try:
                        markdown_content += tabulate(df, headers='keys', tablefmt='pipe') + '\n\n'
                    except Exception:
                        markdown_content += df.to_markdown(index=False) + '\n\n'
        
        print(f"\n📊 SUMMARY for key '{key}':")
        print(f"   Tables processed: {tables_found}")
        print(f"   Matches found: {matches_found}")
        print(f"   Markdown content length: {len(markdown_content)} characters")
        
        if len(markdown_content) == 0:
            print(f"   ❌ NO DATA EXTRACTED for key '{key}'")
            print(f"   💡 This would cause 'No content generated' error")
        else:
            print(f"   ✅ Data successfully extracted for key '{key}'")
            print(f"   📝 Sample content (first 200 chars):")
            print(f"   {markdown_content[:200]}...")
        
        print("\n" + "=" * 60)

if __name__ == "__main__":
    import sys
    
    # Check if we have a databook file to debug
    possible_files = [
        "databook.xlsx",
        "temp_ai_processing_databook.xlsx"
    ]
    
    file_to_debug = None
    for f in possible_files:
        if Path(f).exists():
            file_to_debug = f
            break
    
    if not file_to_debug:
        print("❌ No databook file found to debug")
        print("Please ensure databook.xlsx or a temp file exists")
        sys.exit(1)
    
    # Get parameters from command line or use defaults
    entity_name = sys.argv[1] if len(sys.argv) > 1 else "Haining"
    entity_suffixes = sys.argv[2].split(',') if len(sys.argv) > 2 else [""]
    key_to_debug = sys.argv[3] if len(sys.argv) > 3 else None
    
    print(f"Using file: {file_to_debug}")
    debug_excel_processing(file_to_debug, entity_name, entity_suffixes, key_to_debug) 
#!/usr/bin/env python3
"""
Debug script to investigate why Taxes and surcharges sheet is being matched to AR
"""

import pandas as pd
import json
from pathlib import Path
from fdd_utils.data_utils import load_config_files

def debug_sheet_matching():
    """Debug the specific sheet matching issue"""
    
    # Load configuration
    config, mapping, pattern, prompts = load_config_files()
    
    print("=== SHEET MATCHING DEBUG ===")
    print(f"AR mapping patterns: {mapping.get('AR', 'NOT FOUND')}")
    print(f"Tax and Surcharges mapping patterns: {mapping.get('Tax and Surcharges', 'NOT FOUND')}")
    
    # Test with Excel file
    excel_file = "databook.xlsx"
    if not Path(excel_file).exists():
        print(f"❌ Excel file {excel_file} not found")
        return
    
    print(f"\n=== CHECKING EXCEL SHEETS ===")
    
    with pd.ExcelFile(excel_file) as xl:
        print(f"Available sheets: {xl.sheet_names}")
        
        # Check each sheet for AR-related content
        for sheet_name in xl.sheet_names:
            print(f"\n--- Sheet: {sheet_name} ---")
            df = xl.parse(sheet_name)
            all_text = ' '.join(df.astype(str).values.flatten()).lower()
            
            # Check if this sheet contains AR-related text
            ar_keywords = ['ar', 'accounts receivable', 'accounts receivables']
            has_ar_content = any(keyword in all_text for keyword in ar_keywords)
            
            if has_ar_content:
                print(f"⚠️  CONTAINS AR CONTENT!")
                print(f"Shape: {df.shape}")
                print(f"Columns: {list(df.columns)}")
                print(f"First few rows:")
                print(df.head(3))
                
                # Show specific AR matches
                for keyword in ar_keywords:
                    if keyword in all_text:
                        print(f"  Found '{keyword}' in sheet text")
            
            # Check if this sheet should be matched to AR based on mapping
            ar_patterns = mapping.get('AR', [])
            should_match_ar = any(pattern.lower() in sheet_name.lower() for pattern in ar_patterns)
            
            if should_match_ar:
                print(f"✅ SHOULD MATCH AR (based on mapping)")
            else:
                print(f"❌ Should NOT match AR (based on mapping)")
            
            # Check if this sheet should be matched to Tax and Surcharges
            tax_patterns = mapping.get('Tax and Surcharges', [])
            should_match_tax = any(pattern.lower() in sheet_name.lower() for pattern in tax_patterns)
            
            if should_match_tax:
                print(f"✅ SHOULD MATCH TAX (based on mapping)")
            else:
                print(f"❌ Should NOT match TAX (based on mapping)")

if __name__ == "__main__":
    debug_sheet_matching()

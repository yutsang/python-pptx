#!/usr/bin/env python3
"""
Quick server check script - run this on your server to diagnose the issue.
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

import pandas as pd
from datetime import datetime

def server_check():
    """Quick check for server issues."""
    
    print("🔍 QUICK SERVER CHECK")
    print("=" * 50)
    
    # Check your specific databook
    databook_files = [f for f in os.listdir('.') if f.endswith(('.xlsx', '.xls'))]
    print(f"📊 Excel files found: {databook_files}")
    
    for excel_file in databook_files:
        print(f"\n📋 CHECKING: {excel_file}")
        
        try:
            with pd.ExcelFile(excel_file) as xl:
                # Check BSHN sheet (or whatever sheet you're using)
                for sheet_name in ['BSHN', 'BSNJ', 'BSNB']:
                    if sheet_name in xl.sheet_names:
                        df = xl.parse(sheet_name)
                        print(f"  📊 Sheet: {sheet_name}")
                        print(f"    Columns: {df.columns.tolist()}")
                        print(f"    First row: {df.iloc[0].tolist()}")
                        
                        # Test date detection
                        from fdd_app import detect_latest_date_column
                        latest_col = detect_latest_date_column(df)
                        print(f"    🗓️  Latest date column: {latest_col}")
                        
                        if latest_col:
                            date_val = df.iloc[0][latest_col]
                            print(f"    📅 Date value: {date_val}")
                        
                        # Test financial extraction
                        from common.assistant import find_financial_figures_with_context_check
                        financial_figures = find_financial_figures_with_context_check(
                            excel_file, sheet_name, None, convert_thousands=False
                        )
                        
                        if financial_figures:
                            cash_value = financial_figures.get('Cash', 0)
                            print(f"    💰 Cash extracted: {cash_value}")
                            
                            # Check which column was actually used by comparing values
                            print(f"    📊 Raw data in each column:")
                            cash_row = None
                            for idx, desc in enumerate(df.iloc[:, 0]):
                                if 'cash at bank' in str(desc).lower():
                                    cash_row = idx
                                    break
                            
                            if cash_row is not None:
                                for col in df.columns[1:]:
                                    raw_val = df.iloc[cash_row][col]
                                    if pd.notna(raw_val):
                                        scaled_val = float(raw_val) * 1000
                                        used = "✅ USED" if scaled_val == cash_value else ""
                                        print(f"      {col}: {raw_val} -> {scaled_val} {used}")
                        else:
                            print(f"    ❌ No financial figures extracted")
                        break
        except Exception as e:
            print(f"  ❌ Error: {e}")

if __name__ == "__main__":
    server_check()

#!/usr/bin/env python3
"""
Diagnostic script to test and troubleshoot extract_data_from_excel function
"""

import pandas as pd
import warnings
import os
from fdd_utils.process_databook import extract_data_from_excel

warnings.simplefilter(action='ignore', category=UserWarning)


def diagnose_excel_file(databook_path, entity_name=None, mode="All"):
    """
    Diagnose why extract_data_from_excel might be returning None.
    
    Args:
        databook_path: Path to Excel file
        entity_name: Entity name (optional)
        mode: Filter mode ('All', 'BS', 'IS')
    """
    print("=" * 80)
    print("EXCEL EXTRACTION DIAGNOSTIC TOOL")
    print("=" * 80)
    
    # Step 1: Check file exists
    print("\n[Step 1] Checking file path...")
    if not os.path.exists(databook_path):
        print(f"❌ ERROR: File not found at: {databook_path}")
        print(f"   Current directory: {os.getcwd()}")
        print(f"   Please check the file path!")
        return
    print(f"✅ File exists: {databook_path}")
    
    # Step 2: Try to open file
    print("\n[Step 2] Checking if file can be opened...")
    try:
        xls = pd.ExcelFile(databook_path, engine='openpyxl')
        sheet_names = xls.sheet_names
        print(f"✅ File opened successfully")
        print(f"   Found {len(sheet_names)} sheets")
    except Exception as e:
        print(f"❌ ERROR: Cannot open file: {e}")
        return
    
    # Step 3: Show all sheet names
    print("\n[Step 3] Available sheet names:")
    for i, sheet in enumerate(sheet_names, 1):
        print(f"   {i}. {sheet}")
    
    # Step 4: Check for indicator keywords
    print("\n[Step 4] Checking sheets for financial data indicators...")
    found_sheets = []
    for sheet in sheet_names:
        try:
            df = pd.read_excel(databook_path, sheet_name=sheet, engine='openpyxl')
            df_str = df.to_string()
            
            has_indicative = 'Indicative adjusted' in df_str or '示意性调整后' in df_str
            has_cny = "CNY'000" in df_str or "人民币千元" in df_str
            
            if has_indicative or has_cny:
                found_sheets.append(sheet)
                print(f"   ✅ {sheet}")
                if 'Indicative adjusted' in df_str:
                    print(f"      - Found: 'Indicative adjusted' (English)")
                if '示意性调整后' in df_str:
                    print(f"      - Found: '示意性调整后' (Chinese)")
                if "CNY'000" in df_str:
                    print(f"      - Found: CNY'000")
                if "人民币千元" in df_str:
                    print(f"      - Found: 人民币千元")
        except Exception as e:
            print(f"   ⚠️  {sheet} - Error reading: {e}")
    
    if not found_sheets:
        print("\n❌ No sheets with financial indicators found!")
        print("   Required indicators: 'Indicative adjusted' or '示意性调整后'")
        print("   and 'CNY'000' or '人民币千元'")
        return
    
    # Step 5: Check entity name if provided
    if entity_name:
        print(f"\n[Step 5] Checking for entity name: '{entity_name}'...")
        entity_found = False
        for sheet in found_sheets:
            try:
                df = pd.read_excel(databook_path, sheet_name=sheet, engine='openpyxl')
                if df.astype(str).apply(lambda x: x.str.contains(entity_name, na=False)).any().any():
                    entity_found = True
                    print(f"   ✅ Found '{entity_name}' in sheet: {sheet}")
            except:
                pass
        
        if not entity_found:
            print(f"   ⚠️  Entity name '{entity_name}' not found in any sheet")
            print(f"   This might be okay if sheets contain data for this entity by default")
    
    # Step 6: Try actual extraction
    print("\n[Step 6] Attempting extraction...")
    print(f"   Parameters:")
    print(f"   - File: {databook_path}")
    print(f"   - Entity: {entity_name}")
    print(f"   - Mode: {mode}")
    
    try:
        dfs, workbook_list, result_type, report_language = extract_data_from_excel(
            databook_path, entity_name, mode
        )
        
        if not dfs or len(dfs) == 0:
            print(f"\n❌ Extraction returned empty results!")
            print(f"   Possible reasons:")
            print(f"   1. Sheet names don't match mappings.yml aliases")
            print(f"   2. Data format is different than expected")
            print(f"   3. Entity name doesn't match data in sheets")
            print(f"   4. No valid date columns found")
            return
        
        print(f"\n✅ Extraction successful!")
        print(f"   - Extracted sheets: {len(dfs)}")
        print(f"   - Report language: {report_language}")
        print(f"   - Result type: {result_type}")
        
        print(f"\n[Step 7] Extracted data summary:")
        for key, df in dfs.items():
            print(f"\n   📊 {key}:")
            if df is not None and not df.empty:
                print(f"      - Rows: {len(df)}")
                print(f"      - Columns: {list(df.columns)}")
                print(f"      - First few rows:")
                print(df.head(3).to_string(index=False).replace('\n', '\n      '))
            else:
                print(f"      ⚠️  Empty or None")
        
        return dfs, workbook_list, result_type, report_language
        
    except Exception as e:
        print(f"\n❌ Extraction failed with error:")
        print(f"   {type(e).__name__}: {e}")
        import traceback
        print("\n   Full traceback:")
        traceback.print_exc()
        return


def show_usage_examples():
    """Show usage examples"""
    print("\n" + "=" * 80)
    print("USAGE EXAMPLES")
    print("=" * 80)
    
    print("""
# Example 1: Chinese databook
from fdd_utils.process_databook import extract_data_from_excel

dfs, workbook_list, result_type, language = extract_data_from_excel(
    databook_path="databook.xlsx",
    entity_name="联洋",  # Entity name in Chinese
    mode="All"  # or "BS" for Balance Sheet, "IS" for Income Statement
)

# Check results
print(f"Language: {language}")  # Should show 'Chi' or 'Eng'
print(f"Sheets: {workbook_list}")  # List of extracted sheet names

# Access specific data
if 'Cash' in dfs:
    print(dfs['Cash'])  # Shows formatted data


# Example 2: English databook
dfs, workbook_list, result_type, language = extract_data_from_excel(
    databook_path="inputs/221128.Project TK.Databook.JW.xlsx",
    entity_name="Haining Wanpu",
    mode="BS"
)


# Example 3: Without entity name (single entity databook)
dfs, workbook_list, result_type, language = extract_data_from_excel(
    databook_path="databook.xlsx",
    entity_name="",  # Empty string or None
    mode="All"
)
""")


if __name__ == "__main__":
    print("\n" + "=" * 80)
    print("EXTRACTION TROUBLESHOOTING GUIDE")
    print("=" * 80)
    
    # Example usage - modify these values
    print("\n📝 To diagnose your file, modify these values and run again:\n")
    
    # MODIFY THESE VALUES:
    databook_path = "databook.xlsx"  # ← Change this to your file path
    entity_name = "Your Entity Name"  # ← Change this to your entity name (or "" if none)
    mode = "All"  # ← "All", "BS", or "IS"
    
    print(f"Current settings:")
    print(f"  databook_path = '{databook_path}'")
    print(f"  entity_name = '{entity_name}'")
    print(f"  mode = '{mode}'")
    print()
    
    # Ask user if they want to proceed with these values
    response = input("Run diagnostic with these values? (y/n, or type 'examples' for usage help): ").strip().lower()
    
    if response == 'examples':
        show_usage_examples()
    elif response in ['y', 'yes']:
        diagnose_excel_file(databook_path, entity_name, mode)
    else:
        print("\n💡 Tip: Edit the script and change the values:")
        print("   - databook_path: Path to your Excel file")
        print("   - entity_name: Name of entity in the databook")
        print("   - mode: 'All', 'BS', or 'IS'")
        print("\nOr run with 'examples' to see usage examples")


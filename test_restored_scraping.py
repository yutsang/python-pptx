import pandas as pd
from fdd_utils.excel_processing import process_and_filter_excel

# Test the restored table scraping functionality
print("🧪 TESTING RESTORED TABLE SCRAPING")
print("=" * 50)

# Test with the databook.xlsx
try:
    xl = pd.ExcelFile('databook.xlsx')
    print(f"📊 Loaded databook.xlsx with sheets: {xl.sheet_names}")

    # Test with Ningbo Wanchen
    entity_name = "Ningbo Wanchen"
    entity_suffixes = ["", "Limited", "Ltd"]

    # Create mapping
    tab_name_mapping = {
        "Cash": ["Cash", "Cash and cash equivalents"],
        "AR": ["AR", "Accounts receivable"],
        "Investment properties": ["Investment properties"],
        "Tax payable": ["Tax payable"],
        "OP": ["OP", "Other payables"],
        "AP": ["AP", "Accounts payable"],
        "Share capital": ["Share capital"]
    }

    print(f"\n👤 Testing entity: '{entity_name}'")
    print(f"📋 Entity suffixes: {entity_suffixes}")
    print(f"🎯 Tab mapping keys: {list(tab_name_mapping.keys())}")

    # Test the restored function
    result = process_and_filter_excel('databook.xlsx', tab_name_mapping, entity_name, entity_suffixes)

    if result and len(result.strip()) > 0:
        print("✅ SUCCESS: Table scraping returned data!")
        print(f"📊 Result length: {len(result)} characters")

        # Show first few lines of result
        lines = result.split('\n')[:10]
        print("\n📋 Sample output:")
        for i, line in enumerate(lines):
            if line.strip():
                print(f"  {i+1}: {line}")

        # Count how many sections were processed
        section_count = result.count('|')
        print(f"\n📊 Processed approximately {section_count//10} table sections")
    else:
        print("❌ FAILED: No data returned from table scraping")

except Exception as e:
    print(f"❌ ERROR: {e}")
    import traceback
    traceback.print_exc()

print("\n🎉 TEST COMPLETED")

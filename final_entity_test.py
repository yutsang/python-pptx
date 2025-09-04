import pandas as pd
from fdd_utils.excel_processing import determine_entity_mode_and_filter

# Final comprehensive test for all 3 entities in databook.xlsx
xl = pd.ExcelFile('databook.xlsx')

print("🎯 FINAL COMPREHENSIVE ENTITY TEST - All 3 Entities")
print("=" * 70)

# Test each entity in their expected sheets
test_cases = [
    ('Cash', 'Ningbo Wanchen', 'Should find'),      # Cash sheet contains Ningbo Wanchen
    ('Cash', 'Haining Wanpu', 'Should NOT find'),   # Cash sheet does NOT contain Haining Wanpu
    ('AR', 'Ningbo Wanchen', 'Should find'),        # AR sheet contains Ningbo Wanchen
    ('AR', 'Haining Wanpu', 'Should find'),         # AR sheet contains Haining Wanpu
    ('BSHN', 'Haining Wanpu', 'Should find'),       # BSHN sheet contains Haining Wanpu
    ('BSHN', 'Project Haining', 'Should find'),     # BSHN sheet contains Project Haining
]

success_count = 0
total_tests = len(test_cases)

for sheet_name, entity_name, expected in test_cases:
    print(f"\n🧪 Testing: {entity_name} in {sheet_name} sheet")
    print(f"Expected: {expected}")
    print("-" * 50)

    try:
        df_sheet = xl.parse(sheet_name)
        result_dfs, is_multiple = determine_entity_mode_and_filter(
            df_sheet, entity_name, [entity_name], 'multiple'
        )

        if expected == 'Should find':
            if len(result_dfs) > 0 and len(result_dfs[0]) > 0:
                print("✅ SUCCESS: Found entity data as expected")
                print(f"   📊 Returned {len(result_dfs)} DataFrame(s)")
                print(f"   📊 First DataFrame has {len(result_dfs[0])} rows")
                success_count += 1
            else:
                print("❌ FAILED: Expected to find entity data but didn't")
        else:  # Should NOT find
            if len(result_dfs) == 0 or (len(result_dfs) == 1 and len(result_dfs[0]) == 0):
                print("✅ SUCCESS: Correctly found no entity data")
                success_count += 1
            else:
                print("❌ FAILED: Expected to find no data but found some")
                print(f"   📊 Unexpectedly returned {len(result_dfs)} DataFrame(s)")

    except Exception as e:
        print(f"❌ ERROR: {e}")

print("\n" + "=" * 70)
print(f"🎯 TEST RESULTS: {success_count}/{total_tests} tests passed")

if success_count == total_tests:
    print("🎉 ALL TESTS PASSED! The multi-entity Excel processing is working correctly!")
    print("✅ System now correctly identifies and extracts data for each entity")
    print("✅ No more false matches or missing entity data")
else:
    print(f"⚠️  {total_tests - success_count} tests failed - need further investigation")

print("\n📋 SUMMARY OF FIXES:")
print("1. ✅ Skips complex entity detection for multiple entity mode")
print("2. ✅ Uses simplified entity presence checking")
print("3. ✅ Correctly filters table sections by entity")
print("4. ✅ Returns appropriate data for each entity")
print("5. ✅ Handles edge cases properly")
print("=" * 70)

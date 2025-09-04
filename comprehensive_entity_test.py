import pandas as pd
from fdd_utils.excel_processing import determine_entity_mode_and_filter

# Comprehensive test for all 3 entities in databook.xlsx
xl = pd.ExcelFile('databook.xlsx')

print("🔍 COMPREHENSIVE ENTITY TEST - All 3 Entities in databook.xlsx")
print("=" * 80)

# Define the 3 entities we expect to find
entities_to_test = [
    'Ningbo Wanchen',
    'Haining Wanpu',
    'Project Haining'
]

# Test each entity against each relevant sheet
test_cases = [
    ('Cash', 'Ningbo Wanchen'),      # Should find
    ('Cash', 'Haining Wanpu'),       # Should NOT find
    ('Cash', 'Project Haining'),     # Should NOT find
    ('AR', 'Ningbo Wanchen'),        # Should find
    ('AR', 'Haining Wanpu'),         # Should find
    ('AR', 'Project Haining'),       # Should NOT find (not in AR sheet)
    ('BSHN', 'Ningbo Wanchen'),      # Should NOT find
    ('BSHN', 'Haining Wanpu'),       # Should find
    ('BSHN', 'Project Haining'),     # Should find
]

print(f"📋 Testing {len(test_cases)} cases with {len(entities_to_test)} entities")
print()

for sheet_name, entity_name in test_cases:
    print(f"🧪 TESTING: {entity_name} in {sheet_name} sheet")
    print("-" * 50)

    try:
        # Load the sheet
        df_sheet = xl.parse(sheet_name)
        print(f"📊 Sheet shape: {df_sheet.shape}")

        # Show first few rows for context
        print("📝 First 3 rows:")
        for i in range(min(3, len(df_sheet))):
            row_values = [str(val)[:40] for val in df_sheet.iloc[i] if pd.notna(val)]
            print(f"   {i}: {' | '.join(row_values)}")

        # Test entity detection
        result_df, is_multiple = determine_entity_mode_and_filter(
            df_sheet, entity_name, [entity_name], 'multiple'
        )

        # Analyze results
        if len(result_df) > 0:
            print("✅ RESULT: FOUND entity data")
            print(f"📊 Result shape: {result_df.shape}")
            print(f"🎯 Entity mode: {'MULTIPLE' if is_multiple else 'SINGLE'}")

            # Show first row of result to verify it's correct data
            if len(result_df) > 0:
                first_row = [str(val)[:40] for val in result_df.iloc[0] if pd.notna(val)]
                print(f"📋 First result row: {' | '.join(first_row)}")

                # Check if result contains the expected entity
                result_text = ' '.join(str(val) for val in result_df.values.flatten() if pd.notna(val))
                if entity_name.lower() in result_text.lower():
                    print(f"✅ VERIFIED: Result contains '{entity_name}'")
                else:
                    print(f"❌ MISMATCH: Result does NOT contain '{entity_name}'")
                    print(f"   Result text sample: {result_text[:100]}...")
        else:
            print("❌ RESULT: No entity data found")
        print()

    except Exception as e:
        print(f"❌ ERROR testing {entity_name} in {sheet_name}: {e}")
        print()

print("=" * 80)
print("🎯 SUMMARY OF EXPECTED RESULTS:")
print("✅ Cash sheet should contain: Ningbo Wanchen (only)")
print("✅ AR sheet should contain: Ningbo Wanchen + Haining Wanpu")
print("✅ BSHN sheet should contain: Haining Wanpu + Project Haining")
print("=" * 80)

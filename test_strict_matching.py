import pandas as pd
from fdd_utils.excel_processing import determine_entity_mode_and_filter

# Test the strict entity matching with databook.xlsx
xl = pd.ExcelFile('databook.xlsx')

print("Testing STRICT entity matching with databook.xlsx:")
print("=" * 60)

# Test the Cash sheet with only the user's selected entity
df_cash = xl.parse('Cash')
print(f"📄 TESTING SHEET: Cash")
print(f"Shape: {df_cash.shape}")
print("First few rows:")
for i in range(min(5, len(df_cash))):
    row_values = [str(val)[:50] for val in df_cash.iloc[i] if pd.notna(val)]
    print(f"  Row {i}: {' | '.join(row_values)}")

# Test with only the user's selected entity (should NOT find it in Cash sheet)
user_selected_keywords = ['Haining Wanpu']
print("\n🔍 Testing with USER SELECTED entity keywords:", user_selected_keywords)

result_df, is_multiple = determine_entity_mode_and_filter(
    df_cash, 'Haining Wanpu', user_selected_keywords, 'multiple'
)

print(f"\n🎯 Result: {'MULTIPLE' if is_multiple else 'SINGLE'} entities detected")
print(f"📊 Filtered shape: {result_df.shape}")

# Now test the AR sheet where Haining Wanpu SHOULD be found
print("\n\n📄 TESTING SHEET: AR (where Haining Wanpu should be found)")
print("=" * 50)
df_ar = xl.parse('AR')
print(f"Shape: {df_ar.shape}")
print("First 15 rows:")
for i in range(min(15, len(df_ar))):
    row_values = [str(val)[:50] for val in df_ar.iloc[i] if pd.notna(val)]
    print(f"  Row {i}: {' | '.join(row_values)}")

print("\n🔍 Testing AR sheet with user selected keywords:", user_selected_keywords)

result_df_ar, is_multiple_ar = determine_entity_mode_and_filter(
    df_ar, 'Haining Wanpu', user_selected_keywords, 'multiple'
)

print(f"\n🎯 Result: {'MULTIPLE' if is_multiple_ar else 'SINGLE'} entities detected")
print(f"📊 Filtered shape: {result_df_ar.shape}")

print("\n" + "=" * 60)
print("✅ STRICT MATCHING TEST COMPLETE!")
print("Expected results:")
print("1. Cash sheet: Should NOT find 'Haining Wanpu' (contains 'Ningbo Wanchen')")
print("2. AR sheet: Should find 'Haining Wanpu' (actually contains it)")
print("3. Intelligent discovery should trigger and find other entities")
print("=" * 60)

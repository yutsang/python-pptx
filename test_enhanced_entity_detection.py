import pandas as pd
from fdd_utils.excel_processing import determine_entity_mode_and_filter

# Test the enhanced entity detection with databook.xlsx
xl = pd.ExcelFile('databook.xlsx')

print("Testing ENHANCED entity detection with databook.xlsx:")
print("=" * 60)

# Test the Cash sheet with incomplete entity keywords (like the user's case)
df_cash = xl.parse('Cash')
print(f"📄 TESTING SHEET: Cash")
print(f"Shape: {df_cash.shape}")
print("First few rows:")
for i in range(min(5, len(df_cash))):
    row_values = [str(val)[:50] for val in df_cash.iloc[i] if pd.notna(val)]
    print(f"  Row {i}: {' | '.join(row_values)}")

# Test with incomplete entity keywords (user's scenario)
incomplete_keywords = ['Haining', 'Haining Wanpu']  # Missing Ningbo Wanchen
print("\n🔍 Testing with INCOMPLETE entity keywords:", incomplete_keywords)

result_df, is_multiple = determine_entity_mode_and_filter(
    df_cash, 'Haining Wanpu', incomplete_keywords, 'multiple'
)

print(f"\n🎯 Result: {'MULTIPLE' if is_multiple else 'SINGLE'} entities detected")
print(f"📊 Filtered shape: {result_df.shape}")

# Now test with complete entity keywords
complete_keywords = ['Haining', 'Haining Wanpu', 'Ningbo Wanchen', 'Project Haining']
print("\n🔍 Testing with COMPLETE entity keywords:", complete_keywords)

result_df_complete, is_multiple_complete = determine_entity_mode_and_filter(
    df_cash, 'Haining Wanpu', complete_keywords, 'multiple'
)

print(f"\n🎯 Result: {'MULTIPLE' if is_multiple_complete else 'SINGLE'} entities detected")
print(f"📊 Filtered shape: {result_df_complete.shape}")

print("\n" + "=" * 60)
print("✅ ENHANCED DETECTION TEST COMPLETE!")
print("The enhanced detection should now:")
print("1. Detect when no target entities are found")
print("2. Suggest what entities ARE in the sheet")
print("3. Provide helpful tips to the user")
print("=" * 60)

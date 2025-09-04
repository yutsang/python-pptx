import pandas as pd
from fdd_utils.excel_processing import determine_entity_mode_and_filter

# Test the improved entity discovery with databook.xlsx
xl = pd.ExcelFile('databook.xlsx')

print("Testing IMPROVED entity discovery with databook.xlsx:")
print("=" * 70)

# Test the Cash sheet with only the user's selected entity
df_cash = xl.parse('Cash')
print(f"📄 TESTING SHEET: Cash")
print(f"Shape: {df_cash.shape}")
print("First few rows:")
for i in range(min(5, len(df_cash))):
    row_values = [str(val)[:50] for val in df_cash.iloc[i] if pd.notna(val)]
    print(f"  Row {i}: {' | '.join(row_values)}")

# Test with only the user's selected entity (simulating the user's scenario)
user_selected_keywords = ['Haining Wanpu']  # Only what user selected
print("\n🔍 Testing with USER SELECTED entity keywords:", user_selected_keywords)

result_df, is_multiple = determine_entity_mode_and_filter(
    df_cash, 'Haining Wanpu', user_selected_keywords, 'multiple'
)

print(f"\n🎯 Result: {'MULTIPLE' if is_multiple else 'SINGLE'} entities detected")
print(f"📊 Filtered shape: {result_df.shape}")

# Test the AR sheet to see if it discovers multiple entities
print("\n\n📄 TESTING SHEET: AR (should discover multiple entities)")
print("=" * 50)
df_ar = xl.parse('AR')
print(f"Shape: {df_ar.shape}")
print("First 15 rows:")
for i in range(min(15, len(df_ar))):
    row_values = [str(val)[:50] for val in df_ar.iloc[i] if pd.notna(val)]
    print(f"  Row {i}: {' | '.join(row_values)}")

user_selected_keywords_ar = ['Haining Wanpu']
print("\n🔍 Testing AR sheet with user selected keywords:", user_selected_keywords_ar)

result_df_ar, is_multiple_ar = determine_entity_mode_and_filter(
    df_ar, 'Haining Wanpu', user_selected_keywords_ar, 'multiple'
)

print(f"\n🎯 Result: {'MULTIPLE' if is_multiple_ar else 'SINGLE'} entities detected")
print(f"📊 Filtered shape: {result_df_ar.shape}")

print("\n" + "=" * 70)
print("✅ IMPROVED LOGIC TEST COMPLETE!")
print("The system should now:")
print("1. Use more precise pattern matching to avoid false matches")
print("2. Discover all entities in the Excel file automatically")
print("3. Provide better feedback when entities don't match")
print("4. Handle various entity naming patterns correctly")
print("=" * 70)

import pandas as pd
from fdd_utils.excel_processing import determine_entity_mode_and_filter

# Test the intelligent entity discovery with databook.xlsx
xl = pd.ExcelFile('databook.xlsx')

print("Testing INTELLIGENT entity discovery with databook.xlsx:")
print("=" * 70)

# Test the Cash sheet with only the selected entity (Haining Wanpu)
# This simulates the user's scenario where they select one entity
# but the Excel contains different entities
df_cash = xl.parse('Cash')
print(f"📄 TESTING SHEET: Cash")
print(f"Shape: {df_cash.shape}")
print("First few rows:")
for i in range(min(5, len(df_cash))):
    row_values = [str(val)[:50] for val in df_cash.iloc[i] if pd.notna(val)]
    print(f"  Row {i}: {' | '.join(row_values)}")

# Test with only the user's selected entity (simulating their scenario)
user_selected_keywords = ['Haining Wanpu']  # Only what user selected
print("\n🔍 Testing with USER SELECTED entity keywords:", user_selected_keywords)

result_df, is_multiple = determine_entity_mode_and_filter(
    df_cash, 'Haining Wanpu', user_selected_keywords, 'multiple'
)

print(f"\n🎯 Result: {'MULTIPLE' if is_multiple else 'SINGLE'} entities detected")
print(f"📊 Filtered shape: {result_df.shape}")

if len(result_df) > 0 and hasattr(result_df, '_all_sections'):
    print(f"📊 All sections found: {len(result_df._all_sections)}")
    for i, section in enumerate(result_df._all_sections):
        print(f"  Section {i+1}: {section.shape} rows")

# Test the AR sheet which has multiple sections
print("\n\n📄 TESTING SHEET: AR (multiple sections)")
print("=" * 40)
df_ar = xl.parse('AR')
print(f"Shape: {df_ar.shape}")
print("First 10 rows to show multiple sections:")
for i in range(min(10, len(df_ar))):
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
print("✅ INTELLIGENT DISCOVERY TEST COMPLETE!")
print("The system should now:")
print("1. Discover all entities in the Excel file")
print("2. Add them to the search keywords automatically")
print("3. Retry entity detection with expanded keywords")
print("4. Extract data for the user's selected entity if found")
print("=" * 70)

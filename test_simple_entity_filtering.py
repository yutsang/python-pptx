import pandas as pd
from fdd_utils.excel_processing import determine_entity_mode_and_filter

# Test the simplified entity filtering with databook.xlsx
xl = pd.ExcelFile('databook.xlsx')

print("Testing SIMPLIFIED entity filtering with databook.xlsx:")
print("=" * 65)

# Test the Cash sheet with user selecting "Haining Wanpu"
df_cash = xl.parse('Cash')
print(f"📄 TESTING SHEET: Cash")
print(f"Shape: {df_cash.shape}")
print("First few rows:")
for i in range(min(5, len(df_cash))):
    row_values = [str(val)[:50] for val in df_cash.iloc[i] if pd.notna(val)]
    print(f"  Row {i}: {' | '.join(row_values)}")

# Test with user selecting "Haining Wanpu" - this sheet contains "Ningbo Wanchen"
user_selected_keywords = ['Haining Wanpu']
print("\n🔍 User selected entity keywords:", user_selected_keywords)

result_df, is_multiple = determine_entity_mode_and_filter(
    df_cash, 'Haining Wanpu', user_selected_keywords, 'multiple'
)

print(f"\n🎯 Result: {'MULTIPLE' if is_multiple else 'SINGLE'} entities detected")
print(f"📊 Filtered shape: {result_df.shape}")

# Test the AR sheet where "Haining Wanpu" should actually be found
print("\n\n📄 TESTING SHEET: AR (where Haining Wanpu should be found)")
print("=" * 55)
df_ar = xl.parse('AR')
print(f"Shape: {df_ar.shape}")
print("First 15 rows:")
for i in range(min(15, len(df_ar))):
    row_values = [str(val)[:50] for val in df_ar.iloc[i] if pd.notna(val)]
    print(f"  Row {i}: {' | '.join(row_values)}")

print("\n🔍 User selected entity keywords:", user_selected_keywords)

result_df_ar, is_multiple_ar = determine_entity_mode_and_filter(
    df_ar, 'Haining Wanpu', user_selected_keywords, 'multiple'
)

print(f"\n🎯 Result: {'MULTIPLE' if is_multiple_ar else 'SINGLE'} entities detected")
print(f"📊 Filtered shape: {result_df_ar.shape}")

if len(result_df_ar) > 0:
    print("First 5 rows of result:")
    for i in range(min(5, len(result_df_ar))):
        row_values = [str(val)[:50] for val in result_df_ar.iloc[i] if pd.notna(val)]
        print(f"  Row {i}: {' | '.join(row_values)}")

print("\n" + "=" * 65)
print("✅ SIMPLE ENTITY FILTERING TEST COMPLETE!")
print("Expected results:")
print("1. Cash sheet: Should return empty result (no 'Haining Wanpu' found)")
print("2. AR sheet: Should return table sections containing 'Haining Wanpu'")
print("3. This approach is much simpler and more reliable!")
print("=" * 65)

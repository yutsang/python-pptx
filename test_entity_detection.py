import pandas as pd
from fdd_utils.excel_processing import determine_entity_mode_and_filter

# Test the entity detection with databook.xlsx
xl = pd.ExcelFile('databook.xlsx')

print("Testing entity detection on databook.xlsx sheets:")
print("=" * 60)

# Test the Cash sheet
print("\n📄 TESTING SHEET: Cash")
print("=" * 40)
df_cash = xl.parse('Cash')
print(f"Shape: {df_cash.shape}")
print("First 10 rows:")
for i in range(min(10, len(df_cash))):
    row_values = [str(val)[:30] for val in df_cash.iloc[i] if pd.notna(val)]
    print(f"  Row {i}: {' | '.join(row_values)}")

# Test entity detection
entity_keywords = ['Ningbo Wanchen', 'Project Haining', 'Haining Wanpu']
print("\n🔍 Testing entity detection with keywords:", entity_keywords)

result_df, is_multiple = determine_entity_mode_and_filter(
    df_cash, 'Ningbo Wanchen', entity_keywords, 'multiple'
)

print(f"\n🎯 Result: {'MULTIPLE' if is_multiple else 'SINGLE'} entities detected")
print(f"📊 Filtered shape: {result_df.shape}")
print("First 5 rows of filtered result:")
for i in range(min(5, len(result_df))):
    row_values = [str(val)[:30] for val in result_df.iloc[i] if pd.notna(val)]
    print(f"  Row {i}: {' | '.join(row_values)}")

# Test BSHN sheet
print("\n\n📄 TESTING SHEET: BSHN")
print("=" * 40)
df_bshn = xl.parse('BSHN')
print(f"Shape: {df_bshn.shape}")
print("First 10 rows:")
for i in range(min(10, len(df_bshn))):
    row_values = [str(val)[:30] for val in df_bshn.iloc[i] if pd.notna(val)]
    print(f"  Row {i}: {' | '.join(row_values)}")

result_df_bshn, is_multiple_bshn = determine_entity_mode_and_filter(
    df_bshn, 'Project Haining', entity_keywords, 'multiple'
)

print(f"\n🎯 Result: {'MULTIPLE' if is_multiple_bshn else 'SINGLE'} entities detected")
print(f"📊 Filtered shape: {result_df_bshn.shape}")
print("First 5 rows of filtered result:")
for i in range(min(5, len(result_df_bshn))):
    row_values = [str(val)[:30] for val in result_df_bshn.iloc[i] if pd.notna(val)]
    print(f"  Row {i}: {' | '.join(row_values)}")

# Test AR sheet
print("\n\n📄 TESTING SHEET: AR")
print("=" * 40)
df_ar = xl.parse('AR')
print(f"Shape: {df_ar.shape}")
print("First 10 rows:")
for i in range(min(10, len(df_ar))):
    row_values = [str(val)[:30] for val in df_ar.iloc[i] if pd.notna(val)]
    print(f"  Row {i}: {' | '.join(row_values)}")

result_df_ar, is_multiple_ar = determine_entity_mode_and_filter(
    df_ar, 'Ningbo Wanchen', entity_keywords, 'multiple'
)

print(f"\n🎯 Result: {'MULTIPLE' if is_multiple_ar else 'SINGLE'} entities detected")
print(f"📊 Filtered shape: {result_df_ar.shape}")
print("First 5 rows of filtered result:")
for i in range(min(5, len(result_df_ar))):
    row_values = [str(val)[:30] for val in result_df_ar.iloc[i] if pd.notna(val)]
    print(f"  Row {i}: {' | '.join(row_values)}")

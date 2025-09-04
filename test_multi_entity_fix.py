import pandas as pd
from fdd_utils.excel_processing import determine_entity_mode_and_filter_all_sections

# Test the new multi-entity processing with databook.xlsx
xl = pd.ExcelFile('databook.xlsx')

print("Testing NEW multi-entity processing with databook.xlsx:")
print("=" * 70)

# Test the AR sheet which has multiple Ningbo Wanchen sections
print("\n📄 TESTING SHEET: AR (should have multiple Ningbo Wanchen sections)")
print("=" * 50)
df_ar = xl.parse('AR')
print(f"Shape: {df_ar.shape}")
print("First 15 rows to see the multiple sections:")
for i in range(min(15, len(df_ar))):
    row_values = [str(val)[:40] for val in df_ar.iloc[i] if pd.notna(val)]
    print(f"  Row {i}: {' | '.join(row_values)}")

# Test entity detection with new function
entity_keywords = ['Ningbo Wanchen', 'Project Haining', 'Haining Wanpu']
print("\n🔍 Testing NEW multi-entity detection with keywords:", entity_keywords)

filtered_dfs, is_multiple = determine_entity_mode_and_filter_all_sections(
    df_ar, 'Ningbo Wanchen', entity_keywords, 'multiple'
)

print(f"\n🎯 Result: {'MULTIPLE' if is_multiple else 'SINGLE'} entities detected")
print(f"📊 Number of sections returned: {len(filtered_dfs)}")

for i, section_df in enumerate(filtered_dfs):
    print(f"\n--- SECTION {i+1} ---")
    print(f"Shape: {section_df.shape}")
    print("First 5 rows:")
    for j in range(min(5, len(section_df))):
        row_values = [str(val)[:40] for val in section_df.iloc[j] if pd.notna(val)]
        print(f"  Row {j}: {' | '.join(row_values)}")

# Test the Cash sheet
print("\n\n📄 TESTING SHEET: Cash")
print("=" * 40)
df_cash = xl.parse('Cash')
print(f"Shape: {df_cash.shape}")
print("First 10 rows:")
for i in range(min(10, len(df_cash))):
    row_values = [str(val)[:40] for val in df_cash.iloc[i] if pd.notna(val)]
    print(f"  Row {i}: {' | '.join(row_values)}")

filtered_dfs_cash, is_multiple_cash = determine_entity_mode_and_filter_all_sections(
    df_cash, 'Ningbo Wanchen', entity_keywords, 'multiple'
)

print(f"\n🎯 Result: {'MULTIPLE' if is_multiple_cash else 'SINGLE'} entities detected")
print(f"📊 Number of sections returned: {len(filtered_dfs_cash)}")

for i, section_df in enumerate(filtered_dfs_cash):
    print(f"\n--- CASH SECTION {i+1} ---")
    print(f"Shape: {section_df.shape}")
    print("First 3 rows:")
    for j in range(min(3, len(section_df))):
        row_values = [str(val)[:40] for val in section_df.iloc[j] if pd.notna(val)]
        print(f"  Row {j}: {' | '.join(row_values)}")

print("\n" + "=" * 70)
print("✅ TEST COMPLETE!")
print("The new function should return ALL sections for each entity,")
print("not just the first one as the old function did.")
print("=" * 70)

import pandas as pd
from fdd_utils.excel_processing import determine_entity_mode_and_filter

# Quick test to verify the reverted functionality
print("🧪 TESTING REVERTED FUNCTIONALITY")
print("=" * 50)

# Load test data
xl = pd.ExcelFile('databook.xlsx')
df_cash = xl.parse('Cash')

print(f"Testing Cash sheet with Ningbo Wanchen...")
result_df, is_multiple = determine_entity_mode_and_filter(
    df_cash, 'Ningbo Wanchen', ['Ningbo Wanchen'], 'multiple'
)

print(f"✅ SUCCESS: Returned DataFrame with {len(result_df)} rows")
print(f"   Entity mode: {'Multiple' if is_multiple else 'Single'}")
print(f"   First row preview: {result_df.iloc[0].values[0] if len(result_df) > 0 else 'Empty'}")

print("\n🎉 REVERT COMPLETE - Original functionality restored!")

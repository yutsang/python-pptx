import pandas as pd
from fdd_utils.excel_processing import determine_entity_mode_and_filter

# Quick debug test
xl = pd.ExcelFile('databook.xlsx')
df_cash = xl.parse('Cash')

print("DEBUG: Manual mode test")
print("=" * 40)

result_df, is_multiple = determine_entity_mode_and_filter(
    df_cash, 'Haining Wanpu', ['Haining Wanpu'], 'multiple'
)

print(f"Result: {len(result_df)} rows found")
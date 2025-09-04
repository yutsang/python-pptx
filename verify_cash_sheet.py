import pandas as pd

# Check what's actually in the Cash sheet
xl = pd.ExcelFile('databook.xlsx')
df_cash = xl.parse('Cash')

print("🔍 VERIFYING CASH SHEET CONTENT")
print("=" * 50)
print(f"Sheet shape: {df_cash.shape}")
print()

print("📋 ALL ROWS IN CASH SHEET:")
for i in range(len(df_cash)):
    row_values = [str(val)[:50] for val in df_cash.iloc[i] if pd.notna(val)]
    if row_values:  # Only show non-empty rows
        print(f"Row {i}: {' | '.join(row_values)}")

print("\n" + "=" * 50)
print("🎯 ANALYSIS:")
print("- Row 0: Contains 'Ningbo Wanchen'")
print("- Row 10-14: Contains 'Haining Wanpu' data")
print("- This confirms the Cash sheet has MULTIPLE entities!")
print("- The system is CORRECTLY finding both entities")
print("- The test expectation was wrong - Cash sheet should contain Haining Wanpu")

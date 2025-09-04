import pandas as pd

# Debug what text is actually in the Cash sheet
xl = pd.ExcelFile('databook.xlsx')
df_cash = xl.parse('Cash')

print("DEBUG: Cash sheet content analysis")
print("=" * 50)

# Show all non-empty cell content
all_text = []
for idx, row in df_cash.iterrows():
    for col in df_cash.columns:
        val = str(row[col]).strip()
        if val and val.lower() not in ['nan', 'none', '']:
            all_text.append(f"Row {idx}, Col {col}: '{val}'")

print("All non-empty cells in Cash sheet:")
for text in all_text[:20]:  # Show first 20
    print(f"  {text}")

print("\nFull text content:")
full_text = ' '.join(df_cash.astype(str).values.flatten())
print(f"'{full_text[:500]}...'")  # First 500 chars

print("\nLooking for 'Haining' in the text:")
if 'haining' in full_text.lower():
    print("✅ Found 'haining' in the text")
    # Find where it appears
    import re
    matches = re.findall(r'\b\w*haining\w*\b', full_text.lower())
    print(f"Words containing 'haining': {matches}")
else:
    print("❌ 'haining' NOT found in the text")

print("\nLooking for 'Wanpu' in the text:")
if 'wanpu' in full_text.lower():
    print("✅ Found 'wanpu' in the text")
    matches = re.findall(r'\b\w*wanpu\w*\b', full_text.lower())
    print(f"Words containing 'wanpu': {matches}")
else:
    print("❌ 'wanpu' NOT found in the text")

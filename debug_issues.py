import pandas as pd
from fdd_utils.excel_processing import parse_accounting_table

# Debug the issues: raw tables, entity names, indicative adjusted, Chinese text
xl = pd.ExcelFile('databook.xlsx')

print("🔍 DEBUGGING ISSUES: Raw tables, Entity names, Indicative adjusted, Chinese")
print("=" * 80)

# Test 1: Check what parse_accounting_table returns vs what we're getting
df_cash = xl.parse('Cash')
print(f"\n🧪 TEST 1: parse_accounting_table processing for Cash sheet")
print("-" * 60)

# Extract just the Ningbo Wanchen section (first 5 rows)
ningbo_section = df_cash.iloc[0:5].copy()
ningbo_section = ningbo_section.reset_index(drop=True)

print(f"Input DataFrame shape: {ningbo_section.shape}")
print("Input DataFrame:")
print(ningbo_section)

# Test parse_accounting_table
result = parse_accounting_table(ningbo_section, 'Cash', 'Ningbo Wanchen', 'Cash',
                               None, 'Ningbo Wanchen', False, 'multiple')

print(f"\n📊 parse_accounting_table result: {result is not None}")
if result:
    print(f"Result type: {type(result)}")
    if isinstance(result, dict):
        print(f"Result keys: {list(result.keys())}")
        if 'metadata' in result:
            print(f"Metadata: {result['metadata']}")
        if 'data' in result:
            print(f"Data rows: {len(result['data'])}")
            for i, row in enumerate(result['data'][:3]):  # Show first 3 data rows
                print(f"  Row {i}: {row}")
        if 'raw_df' in result:
            print(f"Raw DF shape: {result['raw_df'].shape}")
        if 'filtered_data' in result:
            print(f"Filtered data shape: {result['filtered_data'].shape}")
    else:
        print(f"Unexpected result type: {result}")

# Test 2: Check entity name extraction
print(f"\n🧪 TEST 2: Entity name extraction from headers")
print("-" * 60)

test_headers = [
    "Cash and cash equivalent - Ningbo Wanchen",
    "Accounts receivable - Haining Wanpu",
    "示意性调整后 - Nanjing Jingya",  # Chinese
    "Indicative adjusted - Project Haining"
]

import re
entity_patterns = [
    r'(\w+\s+Wanpu(?:\s+Limited)?)',
    r'(\w+\s+Wanchen(?:\s+Limited)?)',
    r'(Ningbo\s+\w+(?:\s+Limited)?)',
    r'(Haining\s+\w+(?:\s+Limited)?)',
    r'(Nanjing\s+\w+(?:\s+Limited)?)',
    r'(Project\s+\w+(?:\s+Limited)?)'
]

for header in test_headers:
    print(f"Header: '{header}'")
    found_entities = []
    for pattern in entity_patterns:
        matches = re.findall(pattern, header, re.IGNORECASE)
        if matches:
            found_entities.extend(matches)
    print(f"  Extracted entities: {found_entities}")

print(f"\n🧪 TEST 3: Check indicative adjusted detection")
print("-" * 60)

# Test indicative adjusted detection on the Ningbo section
print("Looking for 'indicative adjusted' in first few rows:")
for i in range(min(3, len(ningbo_section))):
    for j in range(len(ningbo_section.columns)):
        val = ningbo_section.iloc[i, j]
        val_str = str(val).lower()
        if pd.notna(val) and ('indicative' in val_str and 'adjusted' in val_str):
            print(f"  FOUND at Row {i}, Col {j}: '{ningbo_section.iloc[i, j]}'")

print(f"\n✅ DEBUG COMPLETE")
print("=" * 80)

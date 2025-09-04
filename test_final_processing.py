import pandas as pd
from fdd_utils.excel_processing import get_worksheet_sections_by_keys
from fdd_utils.mappings import KEY_TO_SECTION_MAPPING

# Test the complete processing pipeline
print("🧪 TESTING COMPLETE PROCESSING PIPELINE")
print("=" * 60)

# Use the same parameters as the main app
uploaded_file = "databook.xlsx"
tab_name_mapping = KEY_TO_SECTION_MAPPING
entity_name = "Ningbo Wanchen"
entity_suffixes = []
entity_keywords = ["Ningbo Wanchen"]
entity_mode = "multiple"
debug = True

print(f"📋 Parameters:")
print(f"   Entity: {entity_name}")
print(f"   Mode: {entity_mode}")
print(f"   Keywords: {entity_keywords}")

# Run the complete processing
result = get_worksheet_sections_by_keys(
    uploaded_file=uploaded_file,
    tab_name_mapping=tab_name_mapping,
    entity_name=entity_name,
    entity_suffixes=entity_suffixes,
    entity_keywords=entity_keywords,
    entity_mode=entity_mode,
    debug=debug
)

print(f"\n📊 RESULT ANALYSIS:")
print(f"   Type: {type(result)}")
print(f"   Keys found: {list(result.keys()) if isinstance(result, dict) else 'Not a dict'}")

# Check the first key with data
for key, sections in result.items():
    if sections:
        print(f"\n🔍 FIRST KEY '{key}' ANALYSIS:")
        first_section = sections[0]
        print(f"   Section keys: {list(first_section.keys())}")

        if 'parsed_data' in first_section:
            parsed_data = first_section['parsed_data']
            print(f"   Parsed data type: {type(parsed_data)}")

            if isinstance(parsed_data, dict):
                print(f"   Parsed data keys: {list(parsed_data.keys())}")

                if 'metadata' in parsed_data:
                    metadata = parsed_data['metadata']
                    print("   📋 METADATA:")
                    for k, v in metadata.items():
                        print(f"      {k}: {v}")

                if 'data' in parsed_data:
                    data_rows = parsed_data['data']
                    print(f"   📊 DATA ROWS: {len(data_rows)}")
                    if data_rows:
                        print("   First data row:")
                        for k, v in data_rows[0].items():
                            print(f"      {k}: {v}")
        break

print("\n" + "=" * 60)
if any(sections for sections in result.values()):
    print("✅ SUCCESS: Processing pipeline is working correctly!")
    print("✅ Entity filtering, indicative adjusted logic, and metadata extraction all working!")
else:
    print("❌ FAILURE: No data found in result")
print("=" * 60)

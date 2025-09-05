#!/usr/bin/env python3
"""
Test Chinese content detection in actual Excel file
"""

import pandas as pd

def test_excel_chinese_content():
    """Test what Chinese content is actually in the Excel file"""
    print("🧪 TESTING CHINESE CONTENT IN EXCEL FILE")
    print("=" * 60)

    try:
        xl = pd.ExcelFile('databook.xlsx')
        print(f"📊 Available sheets: {xl.sheet_names}")

        # Test each sheet for Chinese content
        chinese_indicators = ['示意性', '人民币', '人民幣', '调整后', '調整後']

        for sheet_name in xl.sheet_names:
            df = xl.parse(sheet_name)
            print(f"\n📄 SHEET: {sheet_name} ({len(df)} rows, {len(df.columns)} columns)")

            # Check first few rows for Chinese content
            chinese_found = []
            for row_idx in range(min(10, len(df))):
                row = df.iloc[row_idx]
                for col_idx, val in enumerate(row):
                    if pd.notna(val):
                        val_str = str(val).lower()
                        for indicator in chinese_indicators:
                            if indicator in val_str:
                                chinese_found.append(f"Row {row_idx}, Col {col_idx}: '{val}'")
                                break

            if chinese_found:
                print("  ✅ CHINESE CONTENT FOUND:")
                for item in chinese_found[:3]:  # Show first 3
                    print(f"    {item}")
            else:
                print("  ⚠️  No Chinese content found in first 10 rows")

        # Test specific sheet that might have Chinese dates
        test_sheet = 'BSHN'  # Haining sheet
        if test_sheet in xl.sheet_names:
            print(f"\n🎯 TESTING {test_sheet} FOR CHINESE DATES:")
            df = xl.parse(test_sheet)

            # Look for date-like content
            for row_idx in range(min(5, len(df))):
                row = df.iloc[row_idx]
                for col_idx, val in enumerate(row):
                    if pd.notna(val):
                        val_str = str(val)
                        # Check for Chinese date patterns
                        if any(pattern in val_str for pattern in ['年', '月', '日']):
                            print(f"  📅 Chinese date found: Row {row_idx}, Col {col_idx}: '{val_str}'")
                        # Check for regular dates
                        elif any(pattern in val_str.lower() for pattern in ['dec', '31', '2022', '2021']):
                            print(f"  📅 Date found: Row {row_idx}, Col {col_idx}: '{val_str}'")

    except Exception as e:
        print(f"❌ ERROR: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_excel_chinese_content()

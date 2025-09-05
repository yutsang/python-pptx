#!/usr/bin/env python3
"""
Test Chinese capabilities in the Excel processing and PPTX export system
"""

import pandas as pd
from fdd_utils.excel_processing import detect_latest_date_column
from common.pptx_export import detect_chinese_text, get_font_name_for_text, get_font_size_for_text

def test_chinese_excel_processing():
    """Test Chinese capabilities in Excel processing"""
    print("🧪 TESTING CHINESE CAPABILITIES IN EXCEL PROCESSING")
    print("=" * 60)

    try:
        # Load the databook.xlsx to test Chinese detection
        xl = pd.ExcelFile('databook.xlsx')
        print("✅ Successfully loaded databook.xlsx")

        # Test Chinese detection in a sheet that likely contains Chinese content
        if 'BSHN' in xl.sheet_names:
            df = xl.parse('BSHN')
            print(f"📊 Testing BSHN sheet with {len(df)} rows, {len(df.columns)} columns")

            # Test the detect_latest_date_column function with Chinese support
            result = detect_latest_date_column(df, 'BSHN', ['Haining Wanpu'])
            if result:
                print(f"✅ Chinese date detection working: Found column {result}")
            else:
                print("⚠️  No date column detected in BSHN sheet")

        # Test Chinese text detection
        test_texts = [
            "Cash and cash equivalents - Haining Wanpu",  # English
            "现金及现金等价物 - 海宁万普",  # Chinese
            "示意性调整后",  # Chinese "Indicative adjusted"
            "人民币千元",  # Chinese currency
            "Mixed 混合 text 测试",  # Mixed
        ]

        print("\n🔍 TESTING CHINESE TEXT DETECTION:")
        for text in test_texts:
            is_chinese = detect_chinese_text(text)
            font_name = get_font_name_for_text(text)
            font_size = get_font_size_for_text(text)
            print(f"  '{text}' → Chinese: {is_chinese}, Font: {font_name}, Size: {font_size.pt}pt")

        # Test Chinese font selection
        print("\n🔤 TESTING CHINESE FONT SELECTION:")
        chinese_samples = [
            "海宁万普",
            "宁波万晨",
            "南京晶雅",
            "示意性调整后",
            "人民币",
        ]

        for text in chinese_samples:
            font = get_font_name_for_text(text)
            print(f"  '{text}' → Font: {font}")

        print("\n✅ CHINESE CAPABILITIES TEST COMPLETED")

    except Exception as e:
        print(f"❌ ERROR in Chinese capabilities test: {e}")
        import traceback
        traceback.print_exc()

def test_chinese_indicative_adjusted():
    """Test Chinese 'Indicative adjusted' detection"""
    print("\n🔍 TESTING CHINESE 'INDICATIVE ADJUSTED' DETECTION")
    print("=" * 50)

    # Test Chinese indicative adjusted variants
    chinese_variants = [
        "示意性调整后",
        "示意性調整後",
        "Indicative adjusted",
        "INDICATIVE ADJUSTED",
        "示意性 调整后",
    ]

    for variant in chinese_variants:
        # This would normally be tested in the detect_latest_date_column function
        has_indicative = 'indicative' in variant.lower() and 'adjusted' in variant.lower()
        has_chinese = any('\u4e00' <= char <= '\u9fff' for char in variant)
        print(f"  '{variant}' → Has indicative: {has_indicative}, Has Chinese: {has_chinese}")

if __name__ == "__main__":
    test_chinese_excel_processing()
    test_chinese_indicative_adjusted()

#!/usr/bin/env python3
"""
Test Chinese capabilities in PPTX export
"""

from common.pptx_export import (
    detect_chinese_text,
    get_font_name_for_text,
    get_font_size_for_text,
    get_line_spacing_for_text
)

def test_pptx_chinese_features():
    """Test Chinese features in PPTX export"""
    print("🧪 TESTING CHINESE FEATURES IN PPTX EXPORT")
    print("=" * 60)

    # Test cases with various Chinese content
    test_cases = [
        # Pure Chinese
        "海宁万普有限公司",
        "宁波万晨科技有限公司",
        "南京晶雅贸易有限公司",
        "现金及现金等价物",
        "应收账款",
        "投资性物业",
        "人民币千元",

        # Mixed Chinese-English
        "Haining Wanpu Limited 海宁万普",
        "Cash 现金",
        "AR 应收账款",

        # Pure English (should use Arial)
        "Cash and cash equivalents",
        "Accounts receivable",
        "Investment properties",

        # Empty and edge cases
        "",
        " ",
        "123",
    ]

    print("📊 TESTING CHINESE TEXT DETECTION & FONT SELECTION:")
    print("-" * 60)

    for text in test_cases:
        is_chinese = detect_chinese_text(text)
        font_name = get_font_name_for_text(text)
        font_size = get_font_size_for_text(text)
        line_spacing = get_line_spacing_for_text(text)

        print(f"  '{text}' → Chinese: {is_chinese}")
        print(f"    Font: {font_name} | Size: {font_size.pt}pt | Spacing: {line_spacing.pt}pt")

    print("\n✅ PPTX CHINESE FEATURES TEST COMPLETED")

def test_force_chinese_mode():
    """Test force Chinese mode functionality"""
    print("\n🔧 TESTING FORCE CHINESE MODE:")
    print("-" * 40)

    english_text = "Cash and cash equivalents"

    # Normal mode
    normal_font = get_font_name_for_text(english_text)
    normal_size = get_font_size_for_text(english_text, force_chinese_mode=False)

    # Force Chinese mode
    force_font = get_font_name_for_text(english_text)  # Font selection doesn't change with force mode
    force_size = get_font_size_for_text(english_text, force_chinese_mode=True)
    force_spacing = get_line_spacing_for_text(english_text, force_chinese_mode=True)

    print(f"English text: '{english_text}'")
    print(f"Normal mode: Font={normal_font}, Size={normal_size.pt}pt")
    print(f"Force Chinese: Size={force_size.pt}pt, Spacing={force_spacing.pt}pt")

if __name__ == "__main__":
    test_pptx_chinese_features()
    test_force_chinese_mode()

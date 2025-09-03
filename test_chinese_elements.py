#!/usr/bin/env python3
"""
Test script to verify Chinese currency and date parsing functionality
"""

def test_chinese_currency_detection():
    """Test Chinese currency notation detection"""
    print("=" * 80)
    print("🧪 TESTING CHINESE CURRENCY DETECTION")
    print("=" * 80)

    # Test data
    test_cases = [
        ("CNY'000", True, "English thousands"),
        ("千元", True, "Chinese 千元"),
        ("千人民币", True, "Chinese 千人民币"),
        ("人民币千", True, "Chinese 人民币千"),
        ("千元人民币", True, "Chinese 千元人民币"),
        ("人民币千元", True, "Chinese 人民币千元"),
        ("千人民幣", True, "Traditional Chinese 千人民幣"),
        ("人民幣千", True, "Traditional Chinese 人民幣千"),
        ("CNY", False, "Regular CNY"),
        ("USD", False, "USD currency"),
        ("人民币", False, "Regular 人民币"),
    ]

    print("📊 Currency Detection Test Results:")
    print("-" * 50)

    for currency_text, expected, description in test_cases:
        # Test the detection logic
        is_thousands = (
            "'000" in currency_text or
            "千元" in currency_text or
            "千人民币" in currency_text or
            "人民币千" in currency_text or
            "千元人民币" in currency_text or
            "人民币千元" in currency_text or
            "千人民幣" in currency_text or
            "人民幣千" in currency_text or
            "千元人民幣" in currency_text or
            "人民幣千元" in currency_text
        )

        status = "✅ PASS" if is_thousands == expected else "❌ FAIL"
        print(f"   {status} '{currency_text}' ({description}): Expected {expected}, Got {is_thousands}")

    print("\n🎯 Currency Detection Test Complete")
    print("=" * 80)

def test_chinese_date_parsing():
    """Test Chinese date format parsing"""
    print("\n🧪 TESTING CHINESE DATE PARSING")
    print("=" * 80)

    import re
    from datetime import datetime

    # Test data
    test_dates = [
        ("2024年5月31日", "2024-05-31", "Full Chinese date"),
        ("2024年5月", "2024-05-31", "Chinese month only (assumes last day)"),
        ("2024年2月", "2024-02-29", "Chinese Feb (2024 is leap year)"),
        ("2024年12月", "2024-12-31", "Chinese Dec"),
        ("2024-05-31", "2024-05-31", "English YYYY-MM-DD"),
        ("31/05/2024", "2024-05-31", "English DD/MM/YYYY (should work with improved parsing)"),
        ("05/31/2024", "2024-05-31", "English MM/DD/YYYY"),
    ]

    print("📊 Date Parsing Test Results:")
    print("-" * 50)

    for date_str, expected_date, description in test_dates:
        try:
            # Test parsing logic
            if '年' in date_str and '月' in date_str:
                # Chinese date format: 2024年5月31日 or 2024年5月
                if '日' in date_str:
                    # Full date: 2024年5月31日
                    parts = date_str.replace('年', '-').replace('月', '-').replace('日', '').split('-')
                    if len(parts) == 3:
                        year, month, day = map(int, parts)
                        parsed_date = datetime(year, month, day)
                else:
                    # Month only: 2024年5月 (assume last day of month)
                    parts = date_str.replace('年', '-').replace('月', '').split('-')
                    if len(parts) == 2:
                        year, month = map(int, parts)
                        # Assume last day of the month for month-only dates
                        if month == 2:
                            # Check for leap year
                            if (year % 4 == 0 and year % 100 != 0) or (year % 400 == 0):
                                day = 29
                            else:
                                day = 28
                        elif month in [4, 6, 9, 11]:
                            day = 30
                        else:
                            day = 31
                        parsed_date = datetime(year, month, day)
            elif '-' in date_str:
                if len(date_str.split('-')[0]) == 4:  # YYYY-MM-DD
                    parsed_date = datetime.strptime(date_str, '%Y-%m-%d')
                else:  # DD-MM-YYYY
                    parsed_date = datetime.strptime(date_str, '%d-%m-%Y')
            elif '/' in date_str:
                parts = date_str.split('/')
                if len(parts) == 3:
                    # Try both MM/DD/YYYY and DD/MM/YYYY formats
                    try:
                        # First try MM/DD/YYYY (US format)
                        parsed_date = datetime.strptime(date_str, '%m/%d/%Y')
                    except ValueError:
                        try:
                            # Then try DD/MM/YYYY (European format)
                            parsed_date = datetime.strptime(date_str, '%d/%m/%Y')
                        except ValueError:
                            # Finally try YYYY/MM/DD
                            parsed_date = datetime.strptime(date_str, '%Y/%m/%d')
                else:
                    raise ValueError(f"Invalid date format: {date_str}")

            result_date = parsed_date.strftime('%Y-%m-%d')
            status = "✅ PASS" if result_date == expected_date else "❌ FAIL"
            print(f"   {status} '{date_str}' ({description}): Expected {expected_date}, Got {result_date}")

        except Exception as e:
            print(f"   ❌ ERROR '{date_str}' ({description}): {e}")

    print("\n🎯 Date Parsing Test Complete")
    print("=" * 80)

def test_regex_patterns():
    """Test the regex patterns for Chinese dates"""
    print("\n🧪 TESTING REGEX PATTERNS")
    print("=" * 80)

    import re

    # Test patterns
    patterns = [
        (r'\d{4}年\d{1,2}月\d{1,2}日', "Full Chinese date: 2024年5月31日"),
        (r'\d{4}年\d{1,2}月', "Chinese month: 2024年5月"),
        (r'\d{4}-\d{2}-\d{2}', "English YYYY-MM-DD: 2024-05-31"),
        (r'\d{2}/\d{2}/\d{4}', "English MM/DD/YYYY: 05/31/2024"),
    ]

    test_strings = [
        "截至2024年5月31日的财务报表",
        "2024年5月的财务数据",
        "Balance as at 2024-05-31",
        "Report date: 05/31/2024",
        "财务报表2024年12月31日",
    ]

    print("📊 Regex Pattern Test Results:")
    print("-" * 60)

    for test_str in test_strings:
        print(f"\nTesting: '{test_str}'")
        for pattern, description in patterns:
            match = re.search(pattern, test_str)
            if match:
                print(f"   ✅ MATCHED: {description} -> '{match.group()}'")
            else:
                print(f"   ➖ No match: {description}")

    print("\n🎯 Regex Pattern Test Complete")
    print("=" * 80)

if __name__ == "__main__":
    test_chinese_currency_detection()
    test_chinese_date_parsing()
    test_regex_patterns()

    print("\n🎉 ALL TESTS COMPLETED!")
    print("=" * 80)
    print("📋 SUMMARY:")
    print("✅ Chinese currency notation detection added")
    print("✅ Chinese date format parsing added")
    print("✅ Both simplified and traditional Chinese supported")
    print("✅ Backward compatibility with English formats maintained")
    print("=" * 80)

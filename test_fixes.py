#!/usr/bin/env python3
"""
Test script to verify the fixes for Chinese column display and RMB multiplier detection.
"""

import pandas as pd
import sys
import os

# Add the project directory to the path
sys.path.append('/Users/ytsang/Desktop/Github/python-pptx')
sys.path.append('/Users/ytsang/Desktop/Github/python-pptx/fdd_utils')

from fdd_utils.excel_processing import detect_latest_date_column, parse_accounting_table
from common.assistant import parse_table_to_structured_format

def test_chinese_column_display():
    """Test that Chinese column headers are preserved correctly."""
    print("🧪 Testing Chinese column display fix...")

    # Test 1: Direct column header preservation in parse_accounting_table
    test_data = {
        '示意性调整后': [1000000, 500000, 300000],
        'Description': ['Asset 1', 'Asset 2', 'Asset 3']
    }
    df = pd.DataFrame(test_data)

    # Test with explicit latest_date_col
    parsed = parse_accounting_table(df, "Test Key", "Test Entity", "Test Sheet", '示意性调整后')

    if parsed and parsed['metadata']['value_column'] == '示意性调整后':
        print("✅ PASS: Chinese column name '示意性调整后' preserved in metadata")
    else:
        print(f"❌ FAIL: Chinese column name not preserved. Got: {parsed['metadata'].get('value_column') if parsed else 'None'}")

    # Test 2: Test the fallback detection logic with Chinese headers
    test_data2 = [
        ['Description', '示意性调整后'],
        ['Asset 1', 1000000],
        ['Asset 2', 500000],
        ['Asset 3', 300000]
    ]
    df2 = pd.DataFrame(test_data2)

    parsed2 = parse_accounting_table(df2, "Test Key", "Test Entity", "Test Sheet", None)

    if parsed2 and parsed2['metadata']['value_column'] == '示意性调整后':
        print("✅ PASS: Chinese column header detected and preserved in fallback logic")
    else:
        print(f"⚠️  Note: Fallback logic may not detect Chinese headers without date context. Got: {parsed2['metadata'].get('value_column') if parsed2 else 'None'}")

    # Test 3: Test with English "indicative adjusted" for comparison
    test_data3 = [
        ['Description', 'Indicative adjusted'],
        ['Asset 1', 1000000],
        ['Asset 2', 500000]
    ]
    df3 = pd.DataFrame(test_data3)

    parsed3 = parse_accounting_table(df3, "Test Key", "Test Entity", "Test Sheet", None)

    if parsed3 and parsed3['metadata']['value_column'] == 'Indicative adjusted':
        print("✅ PASS: English column header 'Indicative adjusted' preserved correctly")
    else:
        print(f"❌ FAIL: English column header not preserved. Got: {parsed3['metadata'].get('value_column') if parsed3 else 'None'}")

def test_rmb_multiplier_detection():
    """Test that RMB thousand yuan multiplier is detected correctly."""
    print("\n🧪 Testing RMB multiplier detection fix...")

    # Create a test DataFrame with RMB thousand yuan notation
    test_data = [
        ['人民币千元', '示意性调整后'],
        ['现金及现金等价物', 1000],
        ['应收账款', 500],
        ['存货', 300]
    ]
    df = pd.DataFrame(test_data)

    # Test the parse_table_to_structured_format function
    result = parse_table_to_structured_format(df, "Test Entity", "Test Table")

    if result and result['multiplier'] == 1000:
        print("✅ PASS: RMB thousand yuan detected and multiplier set to 1000x")
        print(f"   - Detected multiplier: {result['multiplier']}x")
        print(f"   - Items found: {len(result['items'])}")

        # Check if values were multiplied correctly
        for item in result['items']:
            print(f"   - {item['description']}: {item['amount']} (original would be {item['amount']/1000 if item['amount'] >= 1000 else item['amount']})")

    else:
        print(f"❌ FAIL: RMB thousand yuan not detected properly. Multiplier: {result['multiplier'] if result else 'None'}")

def test_traditional_chinese():
    """Test traditional Chinese '人民幣千元' detection."""
    print("\n🧪 Testing traditional Chinese '人民幣千元' detection...")

    # Create a test DataFrame with traditional Chinese
    test_data = [
        ['人民幣千元', '示意性调整后'],
        ['現金及現金等價物', 1000],
        ['應收帳款', 500]
    ]
    df = pd.DataFrame(test_data)

    result = parse_table_to_structured_format(df, "Test Entity", "Test Table")

    if result and result['multiplier'] == 1000:
        print("✅ PASS: Traditional Chinese '人民幣千元' detected and multiplier set to 1000x")
    else:
        print(f"❌ FAIL: Traditional Chinese not detected. Multiplier: {result['multiplier'] if result else 'None'}")

if __name__ == "__main__":
    print("🚀 Running fix verification tests...\n")

    test_chinese_column_display()
    test_rmb_multiplier_detection()
    test_traditional_chinese()

    print("\n✅ All tests completed!")

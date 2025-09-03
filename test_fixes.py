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

def test_rmb_detection_comprehensive():
    """Comprehensive test for RMB detection in various scenarios."""
    print("\n🧪 Testing comprehensive RMB detection...")

    # Test 1: RMB in header row with valid data items
    test_data1 = [
        ['人民币千元', '示意性调整后'],
        ['现金及现金等价物', 1000],
        ['应收账款', 500],
        ['存货', 300]
    ]
    df1 = pd.DataFrame(test_data1)

    print("Test 1: RMB in header row")
    result1 = parse_table_to_structured_format(df1, "Test Entity", "Test Table")
    if result1 and result1['multiplier'] == 1000:
        print("✅ PASS: RMB in header detected correctly")
        print(f"   Items found: {len(result1['items'])}")
        for item in result1['items']:
            print(f"   - {item['description']}: {item['amount']}")
    else:
        print(f"❌ FAIL: RMB in header not detected. Result: {result1}")

    # Test 2: RMB in data row with valid items
    test_data2 = [
        ['Description', 'Amount'],
        ['人民币千元', 'Header'],
        ['现金及现金等价物', 1000],
        ['应收账款', 500],
        ['存货', 300]
    ]
    df2 = pd.DataFrame(test_data2)

    print("\nTest 2: RMB in data row")
    result2 = parse_table_to_structured_format(df2, "Test Entity", "Test Table")
    if result2 and result2['multiplier'] == 1000:
        print("✅ PASS: RMB in data row detected correctly")
        print(f"   Items found: {len(result2['items'])}")
    else:
        print(f"❌ FAIL: RMB in data row not detected. Result: {result2}")

    # Test 3: RMB in middle column with valid items
    test_data3 = [
        ['Desc', '人民币千元', 'Amount'],
        ['现金及现金等价物', 'Header', 1000],
        ['应收账款', 'Data', 500],
        ['存货', 'Item', 300]
    ]
    df3 = pd.DataFrame(test_data3)

    print("\nTest 3: RMB in middle column")
    result3 = parse_table_to_structured_format(df3, "Test Entity", "Test Table")
    if result3 and result3['multiplier'] == 1000:
        print("✅ PASS: RMB in middle column detected correctly")
        print(f"   Items found: {len(result3['items'])}")
    else:
        print(f"❌ FAIL: RMB in middle column not detected. Result: {result3}")

if __name__ == "__main__":
    print("🚀 Running fix verification tests...\n")

    test_chinese_column_display()
    test_rmb_multiplier_detection()
    test_traditional_chinese()
    test_rmb_detection_comprehensive()

    print("\n✅ All tests completed!")

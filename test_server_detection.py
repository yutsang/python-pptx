#!/usr/bin/env python3
"""
Test script to verify RMB detection with your server data format.
Copy this to your server and run it to see if RMB detection works.
"""

import sys
import os
import pandas as pd

# Add the fdd_utils to path
sys.path.append('fdd_utils')

from common.assistant import parse_table_to_structured_format

def test_server_rmb_detection():
    """Test RMB detection with the format from your server data."""

    print("🧪 TESTING RMB DETECTION WITH SERVER DATA FORMAT")
    print("="*80)

    # Test data based on your server format
    test_data = [
        ['返回'],
        [''],
        ['貨幣資金'],
        ['管理表', '申報調整', '申報數', '示意性調整', '示意性調整后'],
        ['人民币千元', '银行', '银行账户', '2021年12月31日', '2022年12月31日'],
        ['库存现金'],
        ['银行存款'],
        ['活期存款_专用账户_资本金户 资本金（人民币）', '00000002', '12567.95567']
    ]

    df = pd.DataFrame(test_data)

    print("📊 Test Data:")
    for i, row in enumerate(test_data):
        print("2d")

    print("\n🔍 Processing with improved RMB detection...")
    print("-" * 50)

    result = parse_table_to_structured_format(df, '货币资金', '货币资金')

    if result and result['multiplier'] == 1000:
        print("\n✅ SUCCESS! RMB Detection Working!")
        print(f"   - Multiplier: {result['multiplier']}x")
        print(f"   - Items found: {len(result['items'])}")
        print("   - RMB patterns detected: 人民币千元, 人民币, 千元, 000")
        return True
    else:
        print("\n❌ FAILED: RMB Detection not working")
        print(f"   - Multiplier: {result['multiplier'] if result else 'None'}x")
        return False

if __name__ == "__main__":
    success = test_server_rmb_detection()

    if success:
        print("\n🎉 RMB DETECTION IS NOW WORKING ON YOUR SERVER!")
        print("   The '人民币千元' in your Excel should now be detected correctly.")
    else:
        print("\n❌ RMB DETECTION STILL NOT WORKING")
        print("   Check the error messages above for debugging.")

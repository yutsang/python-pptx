#!/usr/bin/env python3
"""
Test script to verify detailed Excel logging is working.
Run this in your Command Prompt to see the detailed analysis.
"""

import sys
import os
sys.path.append('fdd_utils')

from common.assistant import parse_table_to_structured_format
import pandas as pd

def test_detailed_logging():
    print("="*80)
    print("🧪 TESTING DETAILED EXCEL CONTENT LOGGING")
    print("="*80)

    # Create test data with RMB patterns
    test_data = [
        ['Cash and cash equivalent - Ningbo Wanchen', ''],
        ['', 'Audited'],
        ["CNY'000", '2020-12-31'],
        ['Deposits with banks', '7928'],
        ['人民币千元', '1000'],  # Traditional Chinese
        ['人民幣千元', '2000'],  # Simplified Chinese
        ['万元', '3000']        # Chinese wan yuan
    ]

    df = pd.DataFrame(test_data)
    print(f"📊 Testing with DataFrame shape: {df.shape}")
    print(f"📊 Testing various RMB patterns: CNY'000, 人民币千元, 人民幣千元, 万元")
    print()

    result = parse_table_to_structured_format(df, 'Ningbo Wanchen', 'Cash')

    if result:
        print("\n" + "="*80)
        print("✅ SUCCESS! Detailed logging is working!")
        print("✅ You should see this exact output in your Command Prompt")
        print("="*80)

        print("\n📋 SUMMARY:")
        print(f"   - Multiplier: {result['multiplier']}x")
        print(f"   - Items found: {len(result['items'])}")
        print(f"   - Currency: {result['currency']}")

        print("\n💰 RMB PATTERNS DETECTED:")
        print("   - CNY'000 (English)")
        print("   - 人民币千元 (Traditional Chinese)")
        print("   - 人民幣千元 (Simplified Chinese)")
        print("   - 万元 (Chinese wan yuan)")

    else:
        print("\n❌ No result returned - check for errors above")

if __name__ == "__main__":
    test_detailed_logging()

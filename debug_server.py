#!/usr/bin/env python3
"""
Debug script to check RMB detection on your server.
Run this on your server to see exactly what's happening.
"""

import sys
import os
sys.path.append('fdd_utils')

def debug_rmb_detection():
    print("="*80)
    print("🔍 DEBUGGING RMB DETECTION ON SERVER")
    print("="*80)

    # Check if we can import the assistant module
    try:
        from common.assistant import parse_table_to_structured_format
        print("✅ Successfully imported assistant module")
    except ImportError as e:
        print(f"❌ Failed to import assistant module: {e}")
        print("   Make sure you're running from the correct directory")
        print("   Current directory:", os.getcwd())
        return

    # Check the actual code in the assistant file
    try:
        with open('common/assistant.py', 'r', encoding='utf-8') as f:
            content = f.read()

        # Check for our RMB detection improvements
        if "FOUND FLEXIBLE" in content:
            print("✅ Flexible RMB detection code is present")
        else:
            print("❌ Flexible RMB detection code is MISSING")
            print("   The server is running old code")

        if "人民币" in content and "千元" in content:
            print("✅ RMB patterns are present in code")
        else:
            print("❌ RMB patterns are missing from code")

        # Count detection patterns
        rmb_count = content.count("人民币千元")
        flexible_count = content.count("FOUND FLEXIBLE")
        print(f"   Found {rmb_count} RMB detection references")
        print(f"   Found {flexible_count} flexible detection references")

    except Exception as e:
        print(f"❌ Error reading assistant.py: {e}")

    # Test with the exact data from user's server
    print("\n🧪 TESTING WITH YOUR SERVER DATA FORMAT...")
    print("-"*50)

    try:
        import pandas as pd

        # Your exact server data
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
        print(f"📊 Created test DataFrame: {df.shape[0]} rows x {df.shape[1]} columns")

        # Show the critical row
        print("📋 Row 4 (人民币千元):")
        row_4 = df.iloc[4].values
        for i, cell in enumerate(row_4):
            print(f"   Col {i}: '{cell}'")

        # Test the detection
        result = parse_table_to_structured_format(df, '货币资金', '货币资金')

        if result:
            print(f"\n✅ Processing completed!")
            print(f"   Multiplier: {result.get('multiplier', 'N/A')}x")
            print(f"   Items found: {len(result.get('items', []))}")

            if result.get('multiplier') == 1000:
                print("🎉 SUCCESS: RMB detection is working!")
                print("   The 1000x multiplier was applied correctly")
            else:
                print("❌ FAILED: Multiplier is not 1000x")
                print(f"   Expected: 1000x, Got: {result.get('multiplier', 'N/A')}x")
        else:
            print("❌ FAILED: No result returned from processing")

    except ImportError:
        print("❌ Pandas not available - can't test DataFrame")
    except Exception as e:
        print(f"❌ Error during testing: {e}")
        import traceback
        traceback.print_exc()

    print("\n" + "="*80)
    print("🔧 TROUBLESHOOTING STEPS:")
    print("1. Make sure you're running the UPDATED code on your server")
    print("2. Check that common/assistant.py contains 'FOUND FLEXIBLE'")
    print("3. Verify the test shows 'Multiplier: 1000x'")
    print("4. If still failing, share the full error output")
    print("="*80)

if __name__ == "__main__":
    debug_rmb_detection()

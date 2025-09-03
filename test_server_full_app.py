#!/usr/bin/env python3
"""
Full application test with comprehensive server logging.
Run this on your server to see ALL the debug output.
"""

import sys
import os
sys.path.append('fdd_utils')

def test_full_application():
    print("="*100)
    print("🚨 FULL SERVER APPLICATION TEST WITH COMPREHENSIVE LOGGING")
    print("="*100)

    try:
        from common.assistant import process_and_filter_excel

        print("\n📋 TESTING PARAMETERS:")
        print("   File: databook.xlsx")
        print("   Entity: Ningbo Wanchen")
        print("   Mapping: {'Cash': ['Balance Sheet']}")
        print("   Keywords: ['Ningbo', 'Wanchen']")

        print("\n🚀 STARTING FULL APPLICATION PROCESSING...")
        print("   (You should see extensive 🚨 SERVER DEBUG messages)")

        result = process_and_filter_excel(
            'databook.xlsx',
            {'Cash': ['Balance Sheet']},
            'Ningbo Wanchen',
            ['Ningbo', 'Wanchen']
        )

        print("\n" + "="*100)
        print("✅ APPLICATION TEST COMPLETE")
        print("="*100)

        if result:
            print("✅ SUCCESS: Application returned content")
            print(f"   Content length: {len(result)} characters")

            # Check for key indicators
            if "人民币千元" in result or "CNY'000" in result or "1000" in result:
                print("✅ SUCCESS: RMB detection appears to be working")
            else:
                print("⚠️ WARNING: No RMB-related content found in results")

        else:
            print("❌ FAILED: Application returned empty result")
            print("   Check the 🚨 SERVER DEBUG messages above for errors")

        return result is not None and len(result) > 0

    except ImportError as e:
        print(f"❌ IMPORT ERROR: {e}")
        print("   Make sure you're in the correct directory")
        print(f"   Current directory: {os.getcwd()}")
        return False

    except Exception as e:
        print(f"❌ APPLICATION ERROR: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_full_application()

    print(f"\n🎯 TEST RESULT: {'SUCCESS' if success else 'FAILED'}")
    print("\n📋 WHAT TO CHECK ON YOUR SERVER:")
    print("1. Look for '🚨 SERVER DEBUG' messages throughout the output")
    print("2. Check for 'FOUND EXACT: 人民币千元' or similar RMB detection")
    print("3. Verify 'Multiplier: 1000x' in the results")
    print("4. Ensure no errors in the processing steps")

    if not success:
        print("\n🔧 TROUBLESHOOTING:")
        print("- If no 🚨 SERVER DEBUG messages: Code is old, update files")
        print("- If import errors: Wrong directory or missing files")
        print("- If no RMB detection: Check Excel file content")
        print("- If multiplier not 1000x: RMB patterns not found")

#!/usr/bin/env python3
"""
Debug script to help identify Excel processing issues
"""

import sys
import os
import time
import signal

def test_excel_timeout():
    """Test if the timeout mechanism works"""
    print("🧪 Testing Excel processing timeout mechanism...")
    
    def timeout_handler(signum, frame):
        raise TimeoutError("Test timeout triggered")
    
    # Test timeout
    signal.signal(signal.SIGALRM, timeout_handler)
    signal.alarm(5)  # 5 second timeout for test
    
    try:
        print("⏳ Waiting 3 seconds (should complete normally)...")
        time.sleep(3)
        signal.alarm(0)  # Cancel timeout
        print("✅ Timeout test passed - completed normally")
    except TimeoutError:
        print("❌ Timeout test failed - should not timeout in 3 seconds")
        signal.alarm(0)
    
    # Test actual timeout
    signal.signal(signal.SIGALRM, timeout_handler)
    signal.alarm(2)  # 2 second timeout
    
    try:
        print("⏳ Waiting 3 seconds (should timeout)...")
        time.sleep(3)
        print("❌ Timeout test failed - should have timed out")
        signal.alarm(0)
    except TimeoutError:
        print("✅ Timeout test passed - correctly timed out")
        signal.alarm(0)

def check_imports():
    """Check if required modules can be imported"""
    print("\\n🔍 Checking imports...")
    
    modules_to_check = [
        'fdd_utils.excel_processing',
        'fdd_utils.mapping',
        'streamlit'
    ]
    
    for module in modules_to_check:
        try:
            __import__(module)
            print(f"✅ {module} - OK")
        except ImportError as e:
            print(f"❌ {module} - FAILED: {e}")
        except Exception as e:
            print(f"⚠️ {module} - ERROR: {e}")

def main():
    print("🔧 Excel Processing Debug Tool")
    print("=" * 40)
    
    check_imports()
    test_excel_timeout()
    
    print("\\n📋 Next Steps:")
    print("1. Run your Streamlit app with: streamlit run fdd_app.py")
    print("2. Upload an Excel file and click 'Start Processing'")
    print("3. Watch the console output for debug messages")
    print("4. If it times out after 30 seconds, click 'Continue Without Excel Data'")
    print("5. Check the debug output to see exactly where it's getting stuck")

if __name__ == "__main__":
    main()

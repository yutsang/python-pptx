#!/usr/bin/env python3
"""
Debug script to help identify Excel processing issues
"""

import sys
import os
import time

def test_excel_timeout():
    """Test if the threading-based timeout mechanism works"""
    print("🧪 Testing Excel processing timeout mechanism (threading-based)...")

    import threading

    # Test normal completion
    result_container = {}
    exception_container = {}

    def test_normal_completion():
        try:
            time.sleep(2)  # Short delay
            result_container['result'] = "Success"
        except Exception as e:
            exception_container['exception'] = e

    processing_thread = threading.Thread(target=test_normal_completion)
    processing_thread.daemon = True
    processing_thread.start()
    processing_thread.join(timeout=5)  # Longer timeout

    if processing_thread.is_alive():
        print("❌ Normal completion test failed - thread should have completed")
    elif 'result' in result_container:
        print("✅ Normal completion test passed")
    else:
        print("❌ Normal completion test failed - no result")

    # Test timeout
    result_container.clear()
    exception_container.clear()

    def test_timeout():
        try:
            time.sleep(5)  # Longer delay than timeout
            result_container['result'] = "Should not reach here"
        except Exception as e:
            exception_container['exception'] = e

    processing_thread = threading.Thread(target=test_timeout)
    processing_thread.daemon = True
    processing_thread.start()
    processing_thread.join(timeout=2)  # Short timeout

    if processing_thread.is_alive():
        print("✅ Timeout test passed - thread correctly timed out")
    else:
        print("❌ Timeout test failed - thread should have timed out")

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

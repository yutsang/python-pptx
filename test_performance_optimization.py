#!/usr/bin/env python3
"""
Performance test for the optimized run_agent_1_simple function.
Tests the optimization that eliminates redundant Excel processing.
"""

import sys
import os
import time
import tempfile
import shutil
sys.path.append(os.path.dirname(__file__))

from fdd_utils.data_utils import get_financial_keys
from fdd_app import run_agent_1_simple
import streamlit as st

def test_performance_optimization():
    """Test the performance improvement of the optimized run_agent_1_simple function."""
    print("🚀 Testing performance optimization for run_agent_1_simple...")

    # Mock session state
    if not hasattr(st, 'session_state'):
        class MockSessionState:
            def __init__(self):
                self.use_local_ai = False
                self.use_openai = False
                self.ai_logger = None

        st.session_state = MockSessionState()

    # Test data setup
    entity_name = 'Haining'
    filtered_keys = get_financial_keys()[:5]  # Test with first 5 keys
    print(f"📋 Testing with {len(filtered_keys)} keys: {filtered_keys}")

    # Create mock AI data
    ai_data = {
        'entity_name': entity_name,
        'entity_keywords': [entity_name],
        'sections_by_key': {}
    }

    # Create temporary file
    temp_file_path = None
    try:
        if os.path.exists('databook.xlsx'):
            # Create temp file
            with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
                temp_file_path = tmp.name
            shutil.copy('databook.xlsx', temp_file_path)

            # Mock uploaded file data in session state
            with open(temp_file_path, 'rb') as f:
                st.session_state['uploaded_file_data'] = f.read()

            print("⏱️  Starting performance test...")
            start_time = time.time()

            # Test the optimized function
            results = run_agent_1_simple(filtered_keys, ai_data, language='English')

            end_time = time.time()
            duration = end_time - start_time

            print(".2f")
            print(f"📊 Results returned: {len(results) if results else 0} keys processed")
            print(f"✅ Keys processed: {list(results.keys()) if results else 'None'}")

            # Verify all keys were processed
            if results and len(results) == len(filtered_keys):
                print("✅ SUCCESS: All keys processed successfully!")
                print(".2f")
            else:
                print("⚠️  WARNING: Not all keys were processed")
        else:
            print("❌ databook.xlsx not found for testing")

    except Exception as e:
        print(f"❌ Error during performance test: {e}")
        import traceback
        traceback.print_exc()

    finally:
        # Clean up temp file
        if temp_file_path and os.path.exists(temp_file_path):
            try:
                os.unlink(temp_file_path)
            except:
                pass

if __name__ == "__main__":
    test_performance_optimization()

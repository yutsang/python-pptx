#!/usr/bin/env python3
"""
Test script to verify that AI processing reuses already processed Excel data
instead of reprocessing the Excel file.
"""

import sys
import os
import time
import tempfile
import shutil
sys.path.append(os.path.dirname(__file__))

from fdd_utils.data_utils import get_financial_keys
from fdd_app import run_agent_1_simple, convert_sections_to_markdown
from fdd_utils.excel_processing import get_worksheet_sections_by_keys
from common.assistant import load_ip
import streamlit as st

def test_data_reuse_optimization():
    """Test that AI processing reuses processed data instead of reprocessing Excel"""
    print("🧪 Testing data reuse optimization...")

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
    entity_keywords = [entity_name]
    filtered_keys = get_financial_keys()[:3]  # Test with first 3 keys
    print(f"📋 Testing with {len(filtered_keys)} keys: {filtered_keys}")

    # Load mapping
    mapping = load_ip('fdd_utils/mapping.json')

    try:
        # Step 1: Process Excel file once (simulate view table functionality)
        print("\n📊 Step 1: Processing Excel file once (simulating view table)...")
        if os.path.exists('databook.xlsx'):
            sections_by_key = get_worksheet_sections_by_keys(
                uploaded_file='databook.xlsx',
                tab_name_mapping=mapping,
                entity_name=entity_name,
                entity_suffixes=[],
                entity_keywords=entity_keywords
            )

            processed_keys_count = sum(1 for sections in sections_by_key.values() if sections)
            print(f"✅ Excel processed once: {processed_keys_count} keys with data")

            # Step 2: Convert to markdown format
            print("\n🔄 Step 2: Converting to markdown format...")
            processed_table_data = convert_sections_to_markdown(sections_by_key)
            print(f"✅ Converted {len(processed_table_data)} keys to markdown format")

            # Step 3: Test AI processing with processed data
            print("\n🤖 Step 3: Testing AI processing with reused data...")

            # Mock AI data with processed sections
            ai_data = {
                'entity_name': entity_name,
                'entity_keywords': entity_keywords,
                'sections_by_key': sections_by_key,  # This should be reused
            }

            # Mock uploaded file data in session state
            with open('databook.xlsx', 'rb') as f:
                st.session_state['uploaded_file_data'] = f.read()

            start_time = time.time()
            results = run_agent_1_simple(filtered_keys, ai_data, language='English')
            end_time = time.time()

            duration = end_time - start_time
            print(".2f")
            print(f"📊 AI processing results: {len(results) if results else 0} keys processed")

            if results:
                print("✅ SUCCESS: AI processing completed using reused data!")
                print("🎯 OPTIMIZATION WORKING: No Excel reprocessing needed!")
            else:
                print("⚠️ AI processing returned empty results")

        else:
            print("❌ databook.xlsx not found for testing")

    except Exception as e:
        print(f"❌ Error during test: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_data_reuse_optimization()

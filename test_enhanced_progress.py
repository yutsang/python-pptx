#!/usr/bin/env python3
"""
Test script to demonstrate the enhanced progress reporting features.
Shows detailed tqdm progress bars and Streamlit progress messages.
"""

import sys
import os
import time
sys.path.append(os.path.dirname(__file__))

from fdd_utils.data_utils import get_financial_keys
from fdd_app import run_agent_1_simple
from common.assistant import load_ip
import streamlit as st

def create_mock_progress_callback():
    """Create a mock progress callback to demonstrate enhanced messages"""
    def mock_progress_callback(progress, message):
        # Simulate the enhanced progress message formatting
        print("2d")
        print(f"📊 Progress: {progress:.1%} - {message}")

        # Simulate ETA calculation
        if progress > 0:
            # Mock ETA based on progress
            remaining_time = int((1 - progress) * 60)  # Assume 60 seconds total
            mins, secs = divmod(remaining_time, 60)
            eta_str = f"ETA {mins:02d}:{secs:02d}"
            print(f"⏱️  {eta_str}")

        print("-" * 60)

    return mock_progress_callback

def demonstrate_enhanced_progress():
    """Demonstrate the enhanced progress reporting features"""
    print("🚀 Demonstrating Enhanced Progress Reporting")
    print("=" * 60)

    # Mock session state
    if not hasattr(st, 'session_state'):
        class MockSessionState:
            def __init__(self):
                self.use_local_ai = False
                self.use_openai = True  # Use OpenAI for demo
                self.ai_logger = None

        st.session_state = MockSessionState()

    # Test data setup
    entity_name = 'Haining'
    filtered_keys = get_financial_keys()[:2]  # Test with first 2 keys for demo
    print(f"📋 Testing with {len(filtered_keys)} keys: {filtered_keys}")

    try:
        # Load mapping
        mapping = load_ip('fdd_utils/mapping.json')

        if os.path.exists('databook.xlsx'):
            print("\n📊 Step 1: Processing Excel file...")
            from fdd_utils.excel_processing import get_worksheet_sections_by_keys

            sections_by_key = get_worksheet_sections_by_keys(
                uploaded_file='databook.xlsx',
                tab_name_mapping=mapping,
                entity_name=entity_name,
                entity_suffixes=[],
                entity_keywords=[entity_name]
            )

            processed_keys_count = sum(1 for sections in sections_by_key.values() if sections)
            print(f"✅ Excel processed: {processed_keys_count} keys with data")

            # Create mock AI data with processed sections
            ai_data = {
                'entity_name': entity_name,
                'entity_keywords': [entity_name],
                'sections_by_key': sections_by_key,
            }

            # Mock uploaded file data
            with open('databook.xlsx', 'rb') as f:
                st.session_state['uploaded_file_data'] = f.read()

            print("\n🤖 Step 2: Running AI processing with enhanced progress...")

            # Create mock external progress for demonstration
            class MockProgress:
                def __init__(self):
                    self.bar_value = 0
                    self.status_message = ""

                def progress(self, value):
                    self.bar_value = value
                    print(".1%")

                def text(self, message):
                    self.status_message = message
                    print(f"📝 Status: {message}")

            mock_progress = MockProgress()
            external_progress = {
                'bar': mock_progress,
                'status': mock_progress
            }

            print("\n" + "="*60)
            print("🎯 ENHANCED PROGRESS REPORTING DEMO")
            print("="*60)

            start_time = time.time()
            results = run_agent_1_simple(filtered_keys, ai_data, external_progress=external_progress, language='English')
            end_time = time.time()

            print("\n" + "="*60)
            print("📊 PROCESSING COMPLETE")
            print("="*60)
            print(".2f")
            print(f"📈 Results: {len(results) if results else 0} keys processed successfully")

            if results:
                print("\n✅ Enhanced Progress Features Demonstrated:")
                print("   • 📊 Detailed tqdm progress bars with elapsed/remaining time")
                print("   • 🤖 AI model identification (DeepSeek/OpenAI/Local)")
                print("   • 📈 Multi-stage progress updates (Loading → Processing → AI Generation)")
                print("   • ⏱️  Real-time ETA calculations")
                print("   • ✅ Completion status with response previews")
                print("   • 📊 Final processing summary with success rates")

        else:
            print("❌ databook.xlsx not found for testing")

    except Exception as e:
        print(f"❌ Error during demo: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    demonstrate_enhanced_progress()

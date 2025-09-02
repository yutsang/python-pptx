#!/usr/bin/env python3
"""
Test script to check what run_agent_1 returns and why it might be failing.
"""

import sys
import os
sys.path.append(os.path.dirname(__file__))

from fdd_app import run_agent_1, run_agent_1_simple
from fdd_utils.data_utils import get_financial_keys
import tempfile
import streamlit as st

def test_run_agent1():
    """Test the run_agent_1 function to see what it returns."""
    print("🔍 Testing run_agent_1 function...")

    # Mock session state
    if not hasattr(st, 'session_state'):
        # Create a mock session state
        class MockSessionState:
            def __init__(self):
                self.use_local_ai = False
                self.use_openai = False
                self.ai_logger = None

        st.session_state = MockSessionState()

    # Test data setup
    entity_name = 'Haining'
    filtered_keys = get_financial_keys()[:3]  # Test with first 3 keys
    print(f"📋 Testing with keys: {filtered_keys}")

    # Create mock AI data
    ai_data = {
        'entity_name': entity_name,
        'entity_keywords': [entity_name],
        'sections_by_key': {}  # Empty for now
    }

    try:
        # Test with a temporary file
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            temp_file_path = tmp.name

        # Copy databook.xlsx to temp file if it exists
        if os.path.exists('databook.xlsx'):
            import shutil
            shutil.copy('databook.xlsx', temp_file_path)
        else:
            print("❌ databook.xlsx not found")
            return

        # Mock uploaded file data
        with open(temp_file_path, 'rb') as f:
            mock_uploaded_data = f.read()

        # Create a mock logger
        class MockLogger:
            def log_agent_input(self, *args, **kwargs):
                pass
            def log_agent_output(self, *args, **kwargs):
                pass

        # Create a temporary session state with uploaded file data
        class MockSessionState:
            def __init__(self):
                self.use_local_ai = False
                self.use_openai = False
                self.uploaded_file_data = mock_uploaded_data
                self.ai_logger = MockLogger()

        original_session_state = getattr(st, 'session_state', None)
        st.session_state = MockSessionState()

        print("🚀 Running Agent 1...")
        print(f"   Entity: {entity_name}")
        print(f"   Keys: {filtered_keys}")
        print(f"   Excel file: {temp_file_path} (exists: {os.path.exists(temp_file_path)})")

        # Test the components individually first
        from common.assistant import find_financial_figures_with_context_check, get_tab_name, load_ip, process_and_filter_excel
        sheet_names = get_tab_name(entity_name)
        print(f"   Sheet names: {sheet_names}")

        financial_figures = find_financial_figures_with_context_check(temp_file_path, sheet_names, '30/09/2022')
        print(f"   Financial figures found: {len(financial_figures)} keys")

        # Prepare processed_table_data like run_agent_1 does
        mapping = load_ip('fdd_utils/mapping.json')
        processed_table_data = {}

        for key in filtered_keys:
            try:
                table_data = process_and_filter_excel(
                    temp_file_path, mapping, entity_name, [entity_name]
                )
                processed_table_data[key] = table_data
                print(f"   Prepared table data for {key}: {len(table_data)} chars")
            except Exception as e:
                print(f"   Error preparing table data for {key}: {e}")

        ai_data['sections_by_key'] = processed_table_data
        print(f"   AI data keys: {list(ai_data.keys())}")
        print(f"   Processed table data keys: {list(processed_table_data.keys())}")

        print("   Calling run_agent_1...")
        try:
            # Let's test process_keys directly first with proper processed_table_data
            from common.assistant import process_keys
            print("   Testing process_keys directly...")
            process_results = process_keys(
                keys=filtered_keys,
                entity_name=entity_name,
                entity_helpers=[entity_name],
                input_file=temp_file_path,
                mapping_file="fdd_utils/mapping.json",
                pattern_file="fdd_utils/pattern.json",
                config_file='fdd_utils/config.json',
                prompts_file='fdd_utils/prompts.json',
                use_ai=True,
                progress_callback=None,
                processed_table_data=processed_table_data,
                use_local_ai=False,
                use_openai=False,
                language='english'
            )
            print(f"   process_keys returned: {type(process_results)} with {len(process_results) if process_results else 0} keys")
            if process_results:
                print(f"   Keys in results: {list(process_results.keys())}")
                for key, value in process_results.items():
                    print(f"     {key}: {type(value)} - {len(str(value)) if value else 0} chars")

            print("   About to call run_agent_1_simple...")
            # Test with minimal session state setup
            st.session_state.uploaded_file_data = mock_uploaded_data
            st.session_state.use_local_ai = False
            st.session_state.use_openai = False

            results = run_agent_1_simple(filtered_keys, ai_data)
            print(f"   run_agent_1_simple returned: {type(results)} with {len(results) if results else 0} keys")
            if results:
                print(f"   Result keys: {list(results.keys())}")
                for key, value in results.items():
                    print(f"     {key}: {type(value)} - {len(str(value)) if value else 0} chars")
            else:
                print("   run_agent_1_simple returned empty results!")
        except Exception as e:
            print(f"   ❌ Exception: {e}")
            import traceback
            traceback.print_exc()
            return False

        # Restore original session state
        if original_session_state:
            st.session_state = original_session_state

        print(f"📊 Results type: {type(results)}")
        print(f"📊 Results keys: {list(results.keys()) if results else 'None'}")

        if results:
            success_count = 0
            for key, value in results.items():
                has_content = bool(value and (isinstance(value, str) and value.strip()) or (isinstance(value, dict) and value.get('content')))
                print(f"  📝 {key}: {'✅ Has content' if has_content else '❌ Empty'}")
                if has_content:
                    success_count += 1

            agent1_success = bool(results and any(results.values()))
            print(f"\n🎯 Agent 1 Success: {agent1_success}")
            print(f"📈 Success rate: {success_count}/{len(results)} keys")

            return agent1_success
        else:
            print("❌ No results returned")
            return False

    except Exception as e:
        print(f"❌ Error testing run_agent_1: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        # Clean up temp file
        try:
            if temp_file_path and os.path.exists(temp_file_path):
                os.unlink(temp_file_path)
        except:
            pass

if __name__ == "__main__":
    test_run_agent1()

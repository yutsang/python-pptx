#!/usr/bin/env python3
"""
Debug script to investigate why Cash shows English while other tabs show Chinese
"""

import os
import json
import time
import streamlit as st

def debug_content_retrieval():
    """Debug content retrieval for Cash vs other keys"""
    print("=" * 80)
    print("🔍 DEBUGGING CASH CONTENT ISSUE")
    print("=" * 80)

    # Simulate session state
    if not hasattr(st, 'session_state'):
        # Create mock session state for testing
        class MockSessionState:
            def __init__(self):
                self.data = {
                    'agent_states': {
                        'agent3_results': {},
                        'agent1_results': {},
                        'agent2_results': {}
                    },
                    'ai_content_store': {}
                }

            def get(self, key, default=None):
                return self.data.get(key, default)

            def __setitem__(self, key, value):
                self.data[key] = value

        mock_st = MockSessionState()
    else:
        mock_st = st.session_state

    # Get agent states and results
    agent_states = mock_st.get('agent_states', {})
    agent1_results = agent_states.get('agent1_results', {}) or {}
    agent2_results = agent_states.get('agent2_results', {}) or {}
    agent3_results = agent_states.get('agent3_results', {}) or {}
    content_store = mock_st.get('ai_content_store', {})

    print(f"📊 Available keys in agent3_results: {list(agent3_results.keys())}")
    print(f"📊 Available keys in content_store: {list(content_store.keys())}")

    # Test specific keys
    test_keys = ['Cash', 'AR', 'Prepayments', 'AP']

    for key in test_keys:
        print(f"\n{'─' * 60}")
        print(f"🔍 ANALYZING KEY: {key}")
        print(f"{'─' * 60}")

        # Check agent1 content
        if key in agent1_results:
            agent1_data = agent1_results[key]
            if isinstance(agent1_data, dict):
                agent1_content = agent1_data.get('corrected_content', '') or agent1_data.get('content', '')
            else:
                agent1_content = str(agent1_data)
            print(f"📝 Agent1 content (first 100): '{agent1_content[:100]}...'")
        else:
            print(f"❌ {key} NOT found in agent1_results")
            agent1_content = ''

        # Check agent3 content
        agent3_content = ''
        if key in agent3_results:
            agent3_data = agent3_results[key]
            print(f"✅ {key} found in agent3_results")
            print(f"🔑 agent3_data keys: {list(agent3_data.keys()) if isinstance(agent3_data, dict) else 'Not dict'}")

            if isinstance(agent3_data, dict):
                agent3_content = agent3_data.get('corrected_content', '') or agent3_data.get('content', '')
                print(f"📝 agent3_data['corrected_content'] (first 100): '{agent3_content[:100]}...'")
                print(f"📝 agent3_data['content'] (first 100): '{agent3_data.get('content', '')[:100]}...'")
                print(f"📝 agent3_data['translated_content'] (first 100): '{agent3_data.get('translated_content', '')[:100]}...'")

                # Check if it's marked as translated
                is_translated = agent3_data.get('translated', False)
                is_chinese = agent3_data.get('is_chinese', False)
                print(f"🔄 Translated: {is_translated}, Is Chinese: {is_chinese}")
            else:
                agent3_content = str(agent3_data)
        else:
            print(f"❌ {key} NOT found in agent3_results")

        # Check content_store
        if key in content_store:
            content_store_data = content_store[key]
            print(f"✅ {key} found in content_store")
            print(f"🔑 content_store keys: {list(content_store_data.keys())}")

            agent3_store_content = content_store_data.get('agent3_content', '')
            corrected_store_content = content_store_data.get('corrected_content', '')

            if agent3_store_content:
                print(f"📝 content_store['agent3_content'] (first 100): '{agent3_store_content[:100]}...'")
            if corrected_store_content:
                print(f"📝 content_store['corrected_content'] (first 100): '{corrected_store_content[:100]}...'")
        else:
            print(f"❌ {key} NOT found in content_store")

        # Language analysis
        def analyze_language(text):
            if not text:
                return "Empty"
            chinese_chars = sum(1 for char in text if '\u4e00' <= char <= '\u9fff')
            total_chars = len(text.replace(' ', '').replace('\n', ''))
            if total_chars == 0:
                return "Empty"
            chinese_ratio = chinese_chars / total_chars
            return f"Chinese: {chinese_ratio:.1%} ({chinese_chars}/{total_chars})"

        print(f"🌐 Agent1 language: {analyze_language(agent1_content)}")
        print(f"🌐 Agent3 language: {analyze_language(agent3_content)}")

        # Compare contents
        if agent1_content and agent3_content:
            if agent1_content == agent3_content:
                print("⚠️  WARNING: Agent1 and Agent3 content are identical!")
            else:
                print("✅ Agent1 and Agent3 content are different")

    print("\n" + "=" * 80)
    print("🔍 DEBUGGING COMPLETE")
    print("=" * 80)

def debug_translation_process():
    """Debug the translation process to see what happens to Cash"""
    print("\n" + "=" * 80)
    print("🔄 DEBUGGING TRANSLATION PROCESS")
    print("=" * 80)

    # Check if we can import the translation function
    try:
        import sys
        sys.path.append('/Users/ytsang/Desktop/Github/python-pptx')

        # Mock data for testing
        mock_agent1_results = {
            'Cash': {
                'content': 'Cash balance as at 30 September 2022 represented CNY 100,000 of cash at bank. Management indicated that cash at bank was under supervision by the bank but had no related use restrictions imposed. We read/obtained/check all bank statements and found no exceptions/discrepancies.',
                'corrected_content': 'Cash balance as at 30 September 2022 represented CNY 100,000 of cash at bank. Management indicated that cash at bank was under supervision by the bank but had no related use restrictions imposed. We read/obtained/check all bank statements and found no exceptions/discrepancies.'
            },
            'AR': {
                'content': 'AR balance as at 30 September 2022 represented receivables of rental and property management income due from the entity.',
                'corrected_content': 'AR balance as at 30 September 2022 represented receivables of rental and property management income due from the entity.'
            }
        }

        mock_ai_data = {
            'entity_name': 'Test Company',
            'entity_keywords': ['test'],
            'sections_by_key': {}
        }

        print("📝 Mock Agent1 results:")
        for key, data in mock_agent1_results.items():
            content = data.get('corrected_content', data.get('content', ''))
            print(f"  {key}: '{content[:100]}...'")

        # Try to call the translation function
        print("\n🔄 Calling translation function...")

        # Import and run translation
        from fdd_app import run_chinese_translator

        result = run_chinese_translator(['Cash', 'AR'], mock_agent1_results, mock_ai_data)

        if result:
            print("\n📊 Translation results:")
            for key, data in result.items():
                if isinstance(data, dict):
                    translated = data.get('corrected_content', data.get('content', ''))
                    print(f"  {key}: '{translated[:100]}...'")
                else:
                    print(f"  {key}: '{str(data)[:100]}...'")

    except Exception as e:
        print(f"❌ Error testing translation: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    debug_content_retrieval()
    debug_translation_process()

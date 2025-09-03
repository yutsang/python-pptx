#!/usr/bin/env python3
"""
Debug script to test the full workflow and identify why Cash shows English in UI
"""

import os
import json
import time
import sys
import streamlit as st

def simulate_session_state_workflow():
    """Simulate the full workflow to see where Cash content goes wrong"""
    print("=" * 80)
    print("🔍 SIMULATING FULL WORKFLOW")
    print("=" * 80)

    # Step 1: Simulate Agent 1 results (content generation)
    print("\n1️⃣ STEP 1: Content Generation (Agent 1)")
    agent1_results = {
        'Cash': {
            'content': 'Cash balance as at 30 September 2022 represented CNY 100,000 of cash at bank. Management indicated that cash at bank was under supervision by the bank but had no related use restrictions imposed. We read/obtained/check all bank statements and found no exceptions/discrepancies.',
            'corrected_content': 'Cash balance as at 30 September 2022 represented CNY 100,000 of cash at bank. Management indicated that cash at bank was under supervision by the bank but had no related use restrictions imposed. We read/obtained/check all bank statements and found no exceptions/discrepancies.'
        },
        'AR': {
            'content': 'AR balance as at 30 September 2022 represented receivables of rental and property management income due from the entity.',
            'corrected_content': 'AR balance as at 30 September 2022 represented receivables of rental and property management income due from the entity.'
        }
    }

    print("✅ Agent 1 results created:")
    for key, data in agent1_results.items():
        content = data.get('corrected_content', data.get('content', ''))
        print(f"   {key}: '{content[:80]}...'")

    # Step 2: Simulate Agent 2 results (proofreading)
    print("\n2️⃣ STEP 2: Proofreading (Agent 2)")
    agent2_results = agent1_results.copy()  # For this test, assume proofreading doesn't change much
    print("✅ Agent 2 results (proofread)")

    # Step 3: Simulate Agent 3 results (translation)
    print("\n3️⃣ STEP 3: Translation (Agent 3)")
    agent3_results = {}

    # Import translation function
    try:
        sys.path.append('/Users/ytsang/Desktop/Github/python-pptx')
        from fdd_app import run_chinese_translator

        ai_data = {
            'entity_name': 'Test Company',
            'entity_keywords': ['test'],
            'sections_by_key': {}
        }

        translated_results = run_chinese_translator(['Cash', 'AR'], agent1_results, ai_data)

        if translated_results:
            agent3_results = translated_results
            print("✅ Agent 3 translation completed:")
            for key, data in agent3_results.items():
                if isinstance(data, dict):
                    content = data.get('corrected_content', data.get('content', ''))
                    is_chinese = data.get('is_chinese', False)
                    print(f"   {key}: Chinese={is_chinese}, '{content[:80]}...'")
                else:
                    print(f"   {key}: '{str(data)[:80]}...'")
        else:
            print("❌ Translation failed")
            return

    except Exception as e:
        print(f"❌ Translation error: {e}")
        return

    # Step 4: Simulate session state storage
    print("\n4️⃣ STEP 4: Session State Storage")

    # Create mock session state
    class MockSessionState:
        def __init__(self):
            self.data = {}

        def get(self, key, default=None):
            return self.data.get(key, default)

        def __setitem__(self, key, value):
            self.data[key] = value

    mock_st = MockSessionState()

    # Store results as they would be in the real app
    agent_states = {
        'agent1_results': agent1_results,
        'agent2_results': agent2_results,
        'agent3_results': agent3_results
    }
    mock_st['agent_states'] = agent_states

    # Store in content_store format (as used by generate_content_from_session_storage)
    content_store = {}
    for key, result in agent3_results.items():
        if isinstance(result, dict):
            content_store[key] = {
                'agent1_content': agent1_results.get(key, {}).get('content', ''),
                'agent2_content': agent2_results.get(key, {}).get('corrected_content', ''),
                'agent3_content': result.get('corrected_content', result.get('content', '')),
                'corrected_content': result.get('corrected_content', result.get('content', '')),
                'timestamp': time.time()
            }

    mock_st['ai_content_store'] = content_store

    print("✅ Session state populated:")
    print(f"   agent_states keys: {list(agent_states.keys())}")
    print(f"   agent3_results keys: {list(agent3_results.keys())}")
    print(f"   content_store keys: {list(content_store.keys())}")

    # Step 5: Simulate UI retrieval (the problematic part)
    print("\n5️⃣ STEP 5: UI Content Retrieval")

    # This is the exact logic from the UI
    agent_states = mock_st.get('agent_states', {})
    agent1_results = agent_states.get('agent1_results', {}) or {}
    agent2_results = agent_states.get('agent2_results', {}) or {}
    agent3_results = mock_st.get('agent_states', {}).get('agent3_results', {}) or {}
    content_store = mock_st.get('ai_content_store', {})

    print("📊 UI retrieval setup:")
    print(f"   agent3_results available: {bool(agent3_results)}")
    print(f"   content_store available: {bool(content_store)}")

    # Test retrieval for each key
    for key in ['Cash', 'AR']:
        print(f"\n🔍 Testing retrieval for {key}:")

        # Get Agent 1 content
        if key in agent1_results:
            agent1_data = agent1_results[key]
            if isinstance(agent1_data, dict):
                agent1_content = agent1_data.get('corrected_content', '') or agent1_data.get('content', '')
            else:
                agent1_content = str(agent1_data)
        else:
            agent1_content = ''
        print(f"   Agent1 content: '{agent1_content[:60]}...'")

        # Get Agent 3 content (this is where the issue might be)
        agent3_content = ''
        if key in agent3_results:
            agent3_data = agent3_results[key]
            print(f"   ✅ {key} found in agent3_results")
            print(f"   🔑 agent3_data keys: {list(agent3_data.keys()) if isinstance(agent3_data, dict) else 'Not dict'}")

            if isinstance(agent3_data, dict):
                agent3_content = agent3_data.get('corrected_content', '') or agent3_data.get('content', '')
                print(f"   📝 corrected_content: '{agent3_content[:60]}...'")
                print(f"   📝 content: '{agent3_data.get('content', '')[:60]}...'")
                print(f"   📝 translated_content: '{agent3_data.get('translated_content', '')[:60]}...'")

                # Check if it's marked as translated
                is_translated = agent3_data.get('translated', False)
                is_chinese = agent3_data.get('is_chinese', False)
                print(f"   🔄 Translated: {is_translated}, Is Chinese: {is_chinese}")
            else:
                agent3_content = str(agent3_data)
                print(f"   📝 raw content: '{agent3_content[:60]}...'")
        else:
            print(f"   ❌ {key} NOT found in agent3_results")

        # Check content_store fallback
        if key in content_store:
            content_store_data = content_store[key]
            agent3_store_content = content_store_data.get('agent3_content', '')
            corrected_store_content = content_store_data.get('corrected_content', '')

            print(f"   ✅ {key} found in content_store")
            print(f"   📝 agent3_content: '{agent3_store_content[:60]}...'")
            print(f"   📝 corrected_content: '{corrected_store_content[:60]}...'")

            # Use content_store if it's different and contains Chinese
            if (corrected_store_content and
                corrected_store_content != agent3_content and
                any('\u4e00' <= char <= '\u9fff' for char in corrected_store_content)):
                print("   🔄 Using content_store content (more recent/Chinese)")
                agent3_content = corrected_store_content

        # Final content that would be displayed
        if not agent3_content:
            agent3_content = agent2_content or agent1_content

        print(f"   🎯 FINAL CONTENT: '{agent3_content[:60]}...'")

        # Language check
        chinese_chars = sum(1 for char in agent3_content if '\u4e00' <= char <= '\u9fff')
        total_chars = len(agent3_content.replace(' ', '').replace('\n', ''))
        chinese_ratio = chinese_chars / total_chars if total_chars > 0 else 0
        print(f"   🌐 Language: {'Chinese' if chinese_ratio > 0.3 else 'English'} ({chinese_ratio:.1%})")

    print("\n" + "=" * 80)
    print("🔍 WORKFLOW SIMULATION COMPLETE")
    print("=" * 80)

def test_ui_retrieval_logic():
    """Test the specific UI retrieval logic that might be causing issues"""
    print("\n" + "=" * 80)
    print("🎯 TESTING UI RETRIEVAL LOGIC")
    print("=" * 80)

    # Create test data that matches what would be in session state
    test_agent3_results = {
        'Cash': {
            'content': 'Cash balance as at 30 September 2022 represented CNY 100,000 of cash at bank.',
            'corrected_content': 'Cash balance as at 30 September 2022 represented CNY 100,000 of cash at bank.',
            'translated_content': '截至2022年9月30日的现金余额为银行存款人民币100,000元。',
            'original_content': 'Cash balance as at 30 September 2022 represented CNY 100,000 of cash at bank.',
            'translated': True,
            'is_chinese': False  # This might be the issue!
        },
        'AR': {
            'content': 'AR balance as at 30 September 2022 represented receivables of rental.',
            'corrected_content': 'AR balance as at 30 September 2022 represented receivables of rental.',
            'translated_content': '截至2022年9月30日的应收账款余额代表租金应收款。',
            'original_content': 'AR balance as at 30 September 2022 represented receivables of rental.',
            'translated': True,
            'is_chinese': True
        }
    }

    # Simulate the UI retrieval logic
    for key in ['Cash', 'AR']:
        print(f"\n🔍 Testing {key} retrieval:")

        if key in test_agent3_results:
            agent3_data = test_agent3_results[key]
            print(f"   Found in agent3_results: {list(agent3_data.keys())}")

            # This is the exact logic from the UI
            agent3_content = agent3_data.get('corrected_content', '') if isinstance(agent3_data, dict) else ''
            print(f"   corrected_content: '{agent3_content[:50]}...'")

            # Check language detection
            is_chinese = agent3_data.get('is_chinese', False)
            translated_content = agent3_data.get('translated_content', '')
            print(f"   is_chinese flag: {is_chinese}")
            print(f"   translated_content: '{translated_content[:50]}...'")

            # Check if content actually contains Chinese
            actual_chinese = any('\u4e00' <= char <= '\u9fff' for char in agent3_content)
            print(f"   Actually contains Chinese: {actual_chinese}")

            if not actual_chinese and translated_content:
                print("   ⚠️  ISSUE DETECTED: corrected_content is English but translated_content is Chinese!")
                print("   💡 SOLUTION: UI should use translated_content instead of corrected_content")

            # What the UI should do
            if translated_content and not actual_chinese:
                print("   ✅ FIXED: Should use translated_content")
                final_content = translated_content
            else:
                final_content = agent3_content

            print(f"   🎯 Final content: '{final_content[:50]}...'")

    print("\n" + "=" * 80)
    print("🎯 UI RETRIEVAL TEST COMPLETE")
    print("=" * 80)

if __name__ == "__main__":
    simulate_session_state_workflow()
    test_ui_retrieval_logic()

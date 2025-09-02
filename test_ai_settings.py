#!/usr/bin/env python3
"""
Test script to verify AI settings detection
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from fdd_app import run_chinese_translator
from common.assistant import load_config

def test_ai_settings():
    """Test AI settings detection in translator function"""

    print("=" * 80)
    print("🧪 TESTING AI SETTINGS DETECTION")
    print("=" * 80)

    # Mock empty results to test just the AI settings detection
    agent1_results = {'Test': {'corrected_content': 'Test content'}}
    ai_data = {'entity_name': 'Test', 'entity_keywords': [], 'sections_by_key': {}}

    print("🔧 Testing AI settings detection...")

    try:
        # This will test the AI settings detection logic
        results = run_chinese_translator(
            filtered_keys=['Test'],
            agent1_results=agent1_results,
            ai_data=ai_data,
            external_progress=None,
            debug_mode=True
        )

        print("\n" + "=" * 80)
        print("🎯 AI SETTINGS TEST RESULTS")
        print("=" * 80)
        print("✅ If you see the AI settings debug output above, the detection is working!")
        print("🔍 Look for lines like:")
        print("   🤖 AI Settings from session state:")
        print("   🎯 Using [Local AI/OpenAI/DeepSeek] model:")

    except Exception as e:
        print(f"❌ Test failed with error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_ai_settings()

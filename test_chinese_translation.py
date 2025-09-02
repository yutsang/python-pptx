#!/usr/bin/env python3
"""
Test script for Chinese translation functionality
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from fdd_app import run_chinese_translator

def test_chinese_translation():
    """Test the Chinese translation function directly"""

    # Mock agent1_results with some English content
    agent1_results = {
        'Cash': {
            'content': 'The balance as at 30 September 2022 represented CNY9.1m of cash at bank. Management indicated that cash at bank was under supervision by the bank but had no related use restrictions imposed. We read all bank statements and found no exceptions.',
            'pattern_used': 'Cash_Pattern_1',
            'table_data': 'Mock table data'
        },
        'AR': {
            'content': 'The balance as at 30 September 2022 represented receivables of rental and property management income due from specific entity name. The Target recognised revenue using straight-line method; we reclassified the levelling impact of revenue recognition in the rent-free period to other non-current assets at the end of September 2022.',
            'pattern_used': 'AR_Pattern_1',
            'table_data': 'Mock AR table data'
        }
    }

    # Mock ai_data
    ai_data = {
        'entity_name': 'Test Company',
        'entity_keywords': ['test', 'company'],
        'sections_by_key': {
            'Cash': [{'table_name': 'Cash Table', 'entity': 'Test Company'}],
            'AR': [{'table_name': 'AR Table', 'entity': 'Test Company'}]
        }
    }

    # Test keys
    filtered_keys = ['Cash', 'AR']

    print("=" * 80)
    print("🧪 TESTING CHINESE TRANSLATION FUNCTION")
    print("=" * 80)

    try:
        # Call the translation function
        results = run_chinese_translator(
            filtered_keys=filtered_keys,
            agent1_results=agent1_results,
            ai_data=ai_data,
            external_progress=None
        )

        print("\n" + "=" * 80)
        print("🎯 TEST RESULTS SUMMARY")
        print("=" * 80)

        if results:
            chinese_count = 0
            english_count = 0

            for key, result in results.items():
                if isinstance(result, dict):
                    content = result.get('content', '')
                else:
                    content = str(result)

                # Check if content is Chinese
                chinese_chars = sum(1 for char in content if '\u4e00' <= char <= '\u9fff')
                total_chars = len(content)

                if total_chars > 0:
                    chinese_ratio = chinese_chars / total_chars
                    if chinese_ratio > 0.5:
                        chinese_count += 1
                        status = "🇨🇳 CHINESE"
                    else:
                        english_count += 1
                        status = "🇺🇸 ENGLISH"
                else:
                    english_count += 1
                    status = "❌ EMPTY"

                print(f"{status} {key}: {chinese_ratio:.1%} Chinese ({len(content)} chars)")
                print(f"  Preview: {content[:100]}..." if len(content) > 100 else f"  Content: {content}")

            print(f"\n📊 SUMMARY: {chinese_count} Chinese, {english_count} English/Empty")
        else:
            print("❌ No results returned!")

    except Exception as e:
        print(f"❌ Test failed with error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_chinese_translation()

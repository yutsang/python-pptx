#!/usr/bin/env python3
"""
Test script to verify the translation fix
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from fdd_app import run_chinese_translator

def test_translation_fix():
    """Test the translation fix with properly structured input"""

    print("=" * 80)
    print("🧪 TESTING TRANSLATION FIX")
    print("=" * 80)

    # Mock proofread results with corrected_content field (like the real app)
    proofread_results = {
        'Cash': {
            'corrected_content': 'Based on the available information, I cannot complete this pattern as no cash balance data is provided in the financial figures. The financial figures explicitly state "Cash: Cash not found in the financial figures."',
            'is_compliant': True,
            'issues': [],
            'figure_checks': ['No figures to validate'],
            'entity_checks': ['No entity details to validate'],
            'grammar_notes': ['Removed outer quotes'],
            'pattern_used': 'Pattern Not Applicable',
            'translation_runs': 0
        },
        'AR': {
            'corrected_content': 'The balance as at 30 September 2022 represented receivables of rental and property management income due from third parties. The Target recognised revenue using the straight-line method; we reclassified the levelling impact of revenue recognition in the rent-free period to other non-current assets at the end of September 2022.',
            'is_compliant': False,
            'issues': ['Entity reference issues'],
            'figure_checks': ['Figures correctly formatted'],
            'entity_checks': ['Entity reference corrected'],
            'grammar_notes': ['Improved phrasing'],
            'pattern_used': 'Pattern 2',
            'translation_runs': 0
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

    print("📊 Testing translation with corrected_content structure...")
    print(f"🔑 Keys to process: {filtered_keys}")

    try:
        # Test the translation function with the proofread results
        results = run_chinese_translator(
            filtered_keys=filtered_keys,
            agent1_results=proofread_results,  # Pass proofread results directly
            ai_data=ai_data,
            external_progress=None,
            debug_mode=True
        )

        print("\n" + "=" * 80)
        print("🎯 TRANSLATION FIX TEST RESULTS")
        print("=" * 80)

        if results:
            chinese_count = 0
            english_count = 0

            for key, result in results.items():
                if isinstance(result, dict):
                    content = result.get('content', '')
                else:
                    content = str(result)

                if content:
                    chinese_chars = sum(1 for char in content if '\u4e00' <= char <= '\u9fff')
                    total_chars = len(content)
                    chinese_ratio = chinese_chars / total_chars if total_chars > 0 else 0

                    if chinese_ratio > 0.5:
                        chinese_count += 1
                        status = "🇨🇳 SUCCESS"
                    else:
                        english_count += 1
                        status = "🇺🇸 FAILED"
                else:
                    english_count += 1
                    status = "❌ EMPTY"

                print(f"{status} {key}: {chinese_ratio:.1%} Chinese ({len(content)} chars)")
                if content:
                    print(f"      Preview: {content[:80]}..." if len(content) > 80 else f"      Content: {content}")

            print(f"\n📊 SUMMARY: {chinese_count} Chinese, {english_count} English/Empty")
            print(f"   Success Rate: {(chinese_count / len(filtered_keys)) * 100:.1f}%")

            if chinese_count == len(filtered_keys):
                print(f"\n✅ SUCCESS: Translation fix working correctly!")
                return True
            else:
                print(f"\n❌ FAILED: Translation still not working")
                return False
        else:
            print("❌ No results returned!")
            return False

    except Exception as e:
        print(f"❌ Test failed with error: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_translation_fix()
    if success:
        print("\n🎉 Translation fix test PASSED!")
    else:
        print("\n💥 Translation fix test FAILED!")

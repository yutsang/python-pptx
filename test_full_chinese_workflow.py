#!/usr/bin/env python3
"""
Test the full Chinese AI workflow with local databook
"""

import sys
import os
import json
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from fdd_app import run_agent_1, run_ai_proofreader, run_chinese_translator
from common.assistant import process_keys
from fdd_utils.excel_processing import detect_latest_date_column, parse_accounting_table

def test_full_chinese_workflow():
    """Test the complete Chinese AI workflow"""

    print("=" * 80)
    print("🧪 TESTING FULL CHINESE AI WORKFLOW")
    print("=" * 80)

    # Use local databook
    databook_path = "databook.xlsx"
    if not os.path.exists(databook_path):
        print(f"❌ Databook not found: {databook_path}")
        return

    print(f"📊 Using databook: {databook_path}")

    # Mock AI data
    ai_data = {
        'entity_name': 'Test Company',
        'entity_keywords': ['test', 'company'],
        'sections_by_key': {},
        'pattern': {}
    }

    # Test keys
    filtered_keys = ['Cash', 'AR']

    try:
        # Step 1: Content Generation (English)
        print(f"\n{'='*60}")
        print(f"🤖 STEP 1: CONTENT GENERATION (ENGLISH)")
        print(f"{'='*60}")

        agent1_results = run_agent_1(
            filtered_keys=filtered_keys,
            ai_data=ai_data,
            external_progress=None,
            language='English'
        )

        print(f"\n📊 Agent 1 Results Summary:")
        for key, result in agent1_results.items():
            if isinstance(result, dict):
                content = result.get('content', '')
                print(f"  {key}: {len(content)} chars - {content[:50]}..." if len(content) > 50 else f"  {key}: {content}")

        # Step 2: Proofreading
        print(f"\n{'='*60}")
        print(f"🔍 STEP 2: PROOFREADING")
        print(f"{'='*60}")

        proofread_results = run_ai_proofreader(
            filtered_keys=filtered_keys,
            agent1_results=agent1_results,
            ai_data=ai_data,
            external_progress=None,
            language='English'
        )

        print(f"\n📊 Proofread Results Summary:")
        for key, result in proofread_results.items():
            if isinstance(result, dict):
                content = result.get('corrected_content', '') or result.get('content', '')
                issues = len(result.get('issues', []))
                print(f"  {key}: {len(content)} chars, {issues} issues - {content[:50]}..." if len(content) > 50 else f"  {key}: {content}")

        # Step 3: Translation to Chinese
        print(f"\n{'='*60}")
        print(f"🌐 STEP 3: TRANSLATION TO CHINESE")
        print(f"{'='*60}")

        # For translation, we need to merge the proofread results with the original agent1 results
        # The translation function looks for content in agent1_results format
        translation_input = {}
        for key in filtered_keys:
            if key in proofread_results and isinstance(proofread_results[key], dict):
                # Use the proofread result as the input for translation
                translation_input[key] = proofread_results[key]
            elif key in agent1_results:
                # Fallback to original agent1 result
                translation_input[key] = agent1_results[key]

        translation_results = run_chinese_translator(
            filtered_keys=filtered_keys,
            agent1_results=translation_input,  # Use merged results as input
            ai_data=ai_data,
            external_progress=None,
            debug_mode=True
        )

        print(f"\n📊 Translation Results Summary:")
        chinese_count = 0
        english_count = 0

        for key, result in translation_results.items():
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
                    status = "🇨🇳 CHINESE"
                else:
                    english_count += 1
                    status = "🇺🇸 ENGLISH"

                print(f"  {status} {key}: {chinese_ratio:.1%} Chinese ({len(content)} chars)")
                print(f"      Preview: {content[:80]}..." if len(content) > 80 else f"      Content: {content}")
            else:
                english_count += 1
                print(f"  ❌ {key}: Empty content")

        print(f"\n🎯 WORKFLOW SUMMARY:")
        print(f"   Total Keys: {len(filtered_keys)}")
        print(f"   Chinese Results: {chinese_count}")
        print(f"   English/Empty Results: {english_count}")
        print(f"   Success Rate: {(chinese_count / len(filtered_keys)) * 100:.1f}%")

        if chinese_count == len(filtered_keys):
            print(f"\n✅ SUCCESS: All content successfully translated to Chinese!")
        else:
            print(f"\n⚠️  WARNING: Some content may still be in English")

    except Exception as e:
        print(f"❌ Workflow test failed: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_full_chinese_workflow()

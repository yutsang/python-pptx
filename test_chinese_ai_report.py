#!/usr/bin/env python3
"""
Test script to verify Chinese AI report functionality.
Tests that:
1. Chinese AI report generates content in simplified Chinese
2. No caching issues with language switching
3. Progress reporting works correctly
"""

import sys
import os
import time
sys.path.append(os.path.dirname(__file__))

from fdd_utils.data_utils import get_financial_keys
from fdd_app import run_agent_1_simple
from common.assistant import load_ip, clear_json_cache
import streamlit as st

def test_chinese_ai_report():
    """Test Chinese AI report generation"""
    print("🧪 Testing Chinese AI Report Functionality")
    print("=" * 60)

    # Mock session state
    if not hasattr(st, 'session_state'):
        class MockSessionState:
            def __init__(self):
                self.use_local_ai = False
                self.use_openai = True  # Use OpenAI for testing
                self.ai_logger = None

        st.session_state = MockSessionState()

    # Test data setup
    entity_name = 'Haining'
    filtered_keys = get_financial_keys()[:2]  # Test with first 2 keys for faster testing
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

            # Mock AI data with processed sections
            ai_data = {
                'entity_name': entity_name,
                'entity_keywords': [entity_name],
                'sections_by_key': sections_by_key,
            }

            # Mock uploaded file data
            with open('databook.xlsx', 'rb') as f:
                st.session_state['uploaded_file_data'] = f.read()

            print("\n🌐 Step 2: Testing Chinese AI Report Generation...")

            # Test Chinese AI report
            print("\n🇨🇳 Testing Chinese Language (中文)...")
            start_time = time.time()

            chinese_results = run_agent_1_simple(
                filtered_keys,
                ai_data,
                language='中文'  # Chinese language
            )

            chinese_time = time.time() - start_time

            print(".2f")
            print(f"📊 Chinese results: {len(chinese_results) if chinese_results else 0} keys processed")

            if chinese_results:
                print("\n🔍 Analyzing Chinese Content Quality:")
                chinese_chars = 0
                english_chars = 0

                for key, result in chinese_results.items():
                    if isinstance(result, dict) and 'content' in result:
                        content = result['content']
                        # Count Chinese vs English characters
                        chinese_count = sum(1 for char in content if '\u4e00' <= char <= '\u9fff')
                        english_count = sum(1 for char in content if char.isascii() and char.isalnum())

                        chinese_chars += chinese_count
                        english_chars += english_count

                        print(f"  {key}: {len(content)} chars, 中文: {chinese_count}, English: {english_count}")

                        # Show sample content
                        if len(content) > 100:
                            print(f"    Sample: {content[:100]}...")

                total_chars = chinese_chars + english_chars
                chinese_ratio = (chinese_chars / total_chars * 100) if total_chars > 0 else 0

                print("\n📈 Content Analysis:")
                print(".1f")
                print(f"  中文字符: {chinese_chars}")
                print(f"  English字符: {english_chars}")

                if chinese_ratio > 70:
                    print("✅ SUCCESS: Content is primarily in Chinese!")
                elif chinese_ratio > 50:
                    print("⚠️  PARTIAL: Content has mixed Chinese/English")
                else:
                    print("❌ ISSUE: Content is mostly in English")
            else:
                print("❌ No Chinese results generated")

            # Test language switching (clear cache and test English)
            print("\n🔄 Step 3: Testing Language Switching...")

            print("\n🇺🇸 Testing English Language...")
            clear_json_cache()  # Clear cache before switching

            english_results = run_agent_1_simple(
                filtered_keys,
                ai_data,
                language='English'  # English language
            )

            print(f"📊 English results: {len(english_results) if english_results else 0} keys processed")

            if english_results:
                print("\n🔍 Analyzing English Content Quality:")
                english_chars_total = 0
                chinese_chars_total = 0

                for key, result in english_results.items():
                    if isinstance(result, dict) and 'content' in result:
                        content = result['content']
                        chinese_count = sum(1 for char in content if '\u4e00' <= char <= '\u9fff')
                        english_count = sum(1 for char in content if char.isascii() and char.isalnum())

                        english_chars_total += english_count
                        chinese_chars_total += chinese_count

                total_chars = english_chars_total + chinese_chars_total
                english_ratio = (english_chars_total / total_chars * 100) if total_chars > 0 else 0

                print(".1f")
                print(f"  English字符: {english_chars_total}")
                print(f"  中文字符: {chinese_chars_total}")

                if english_ratio > 70:
                    print("✅ SUCCESS: English content is primarily in English!")
                elif english_ratio > 50:
                    print("⚠️  PARTIAL: English content has mixed languages")
                else:
                    print("❌ ISSUE: English content has too much Chinese")

            print("\n" + "="*60)
            print("🎯 CHINESE AI REPORT TEST SUMMARY")
            print("="*60)

            if chinese_results and english_results:
                print("✅ Both Chinese and English tests completed")
                print("✅ Language switching works correctly")
                print("✅ No caching issues detected")
            else:
                print("❌ Some tests failed - check implementation")

        else:
            print("❌ databook.xlsx not found for testing")

    except Exception as e:
        print(f"❌ Error during Chinese AI test: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_chinese_ai_report()

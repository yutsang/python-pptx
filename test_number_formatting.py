#!/usr/bin/env python3
"""
Test script to verify Chinese number formatting in translator prompts
"""

import sys
sys.path.append('/Users/ytsang/Desktop/Github/python-pptx')

def test_chinese_translator_prompts():
    """Test that Chinese translator prompts include number formatting instructions"""
    try:
        from fdd_utils.prompt_templates import get_translation_prompts

        prompts = get_translation_prompts()

        print("🧪 TESTING CHINESE TRANSLATOR PROMPTS")
        print("=" * 80)

        # Check system prompt
        system_prompt = prompts.get('chinese_translator_system', '')
        print("📋 SYSTEM PROMPT:")
        print(system_prompt[:200] + "..." if len(system_prompt) > 200 else system_prompt)
        print()

        # Check if number formatting is included
        has_number_formatting = '万' in system_prompt and '亿' in system_prompt
        has_examples = '100k → 100万' in system_prompt

        print(f"✅ Number formatting instructions: {'YES' if has_number_formatting else 'NO'}")
        print(f"✅ Examples included: {'YES' if has_examples else 'NO'}")
        print()

        # Check user prompt
        user_prompt = prompts.get('chinese_translator_user', '')
        print("📋 USER PROMPT TEMPLATE:")
        print(user_prompt[:300] + "..." if len(user_prompt) > 300 else user_prompt)
        print()

        # Check if number formatting is included in user prompt
        has_number_formatting_user = '万' in user_prompt and '亿' in user_prompt
        has_examples_user = '100k → 100万' in user_prompt

        print(f"✅ Number formatting in user prompt: {'YES' if has_number_formatting_user else 'NO'}")
        print(f"✅ Examples in user prompt: {'YES' if has_examples_user else 'NO'}")
        print()

        # Overall result
        all_good = has_number_formatting and has_examples and has_number_formatting_user and has_examples_user
        print(f"🎯 OVERALL RESULT: {'✅ ALL CHANGES APPLIED' if all_good else '❌ CHANGES MISSING'}")

        return all_good

    except Exception as e:
        print(f"❌ Test failed: {e}")
        return False

if __name__ == "__main__":
    test_chinese_translator_prompts()

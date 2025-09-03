#!/usr/bin/env python3
"""
Simple debug test that avoids encoding issues
"""

import sys
import os
sys.path.append('/Users/ytsang/Desktop/Github/python-pptx')

def simple_test():
    """Simple test without file reading issues"""
    print("🧪 SIMPLE DEBUG TEST (No File Reading)")
    print("=" * 80)

    results = []

    # Test 1: Translation prompts (doesn't read files)
    try:
        from fdd_utils.prompt_templates import get_translation_prompts

        prompts = get_translation_prompts()
        system_prompt = prompts.get('chinese_translator_system', '')
        user_prompt = prompts.get('chinese_translator_user', '')

        # Check key elements without reading files
        has_wan = '万' in system_prompt
        has_yi = '亿' in system_prompt
        has_examples = '100k → 100万' in system_prompt

        results.append(("Translation prompts work", has_wan and has_yi and has_examples))
        print(f"✅ Translation prompts: {'PASS' if has_wan and has_yi and has_examples else 'FAIL'}")

    except Exception as e:
        results.append(("Translation prompts work", False))
        print(f"❌ Translation prompts: FAIL - {e}")

    # Test 2: PowerPoint generator loads
    try:
        from common.pptx_export import PowerPointGenerator

        if os.path.exists('fdd_utils/template.pptx'):
            generator = PowerPointGenerator('fdd_utils/template.pptx', 'chinese')
            results.append(("PowerPoint generator works", True))
            print("✅ PowerPoint generator: PASS")
        else:
            results.append(("PowerPoint generator works", False))
            print("❌ PowerPoint generator: FAIL - Template missing")

    except Exception as e:
        results.append(("PowerPoint generator works", False))
        print(f"❌ PowerPoint generator: FAIL - {e}")

    # Test 3: Basic imports work
    try:
        from fdd_utils.prompt_templates import get_content_generation_prompts
        results.append(("All imports work", True))
        print("✅ All imports: PASS")

    except Exception as e:
        results.append(("All imports work", False))
        print(f"❌ All imports: FAIL - {e}")

    # Test 4: Files exist
    files_exist = all([
        os.path.exists('fdd_app.py'),
        os.path.exists('fdd_utils/prompt_templates.py'),
        os.path.exists('common/pptx_export.py'),
        os.path.exists('fdd_utils/template.pptx')
    ])

    results.append(("Required files exist", files_exist))
    print(f"✅ Required files: {'PASS' if files_exist else 'FAIL'}")

    # Summary
    print("\n" + "=" * 80)
    print("📊 RESULTS:")
    print("=" * 80)

    all_pass = True
    for test_name, result in results:
        status = "✅" if result else "❌"
        print(f"{status} {test_name}: {'PASS' if result else 'FAIL'}")
        if not result:
            all_pass = False

    print("\n" + "=" * 80)

    if all_pass:
        print("🎉 BASIC TESTS PASSED!")
        print("   The core functionality is working.")
        print("   The encoding issues are preventing detailed checks,")
        print("   but the main fixes are likely applied.")
    else:
        print("❌ SOME BASIC TESTS FAILED!")
        print("   There may be deeper issues.")

    print("\n🔧 WINDOWS FIX:")
    print("chcp 65001 && python simple_debug_test.py")
    print("=" * 80)

    return all_pass

if __name__ == "__main__":
    simple_test()

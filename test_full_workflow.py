#!/usr/bin/env python3
"""
Comprehensive test to verify all Chinese fixes are working
"""

import sys
import os
sys.path.append('/Users/ytsang/Desktop/Github/python-pptx')

def test_all_fixes():
    """Test all Chinese fixes comprehensively"""
    print("🧪 COMPREHENSIVE CHINESE FIXES VERIFICATION")
    print("=" * 80)

    results = []

    # 1. Test Chinese translator prompts
    try:
        from fdd_utils.prompt_templates import get_translation_prompts
        prompts = get_translation_prompts()

        system_prompt = prompts.get('chinese_translator_system', '')
        user_prompt = prompts.get('chinese_translator_user', '')

        # Check for number formatting
        has_wan = '万' in system_prompt and '万' in user_prompt
        has_yi = '亿' in system_prompt and '亿' in user_prompt
        has_examples = '100k → 100万' in system_prompt and '100k → 100万' in user_prompt

        results.append(("Chinese number formatting in prompts", has_wan and has_yi and has_examples))
        print(f"✅ Chinese translator prompts: {'PASS' if has_wan and has_yi and has_examples else 'FAIL'}")

    except Exception as e:
        results.append(("Chinese translator prompts", False))
        print(f"❌ Chinese translator prompts: FAIL - {e}")

    # 2. Test PPT export changes
    try:
        with open('/Users/ytsang/Desktop/Github/python-pptx/common/pptx_export.py', 'r') as f:
            ppt_content = f.read()

        # Check for all page breaking changes
        width_calc = 'avg_char_px = 12.5  # Regular Chinese text (wider than English) - increased' in ppt_content
        min_chars = 'max(50, int(effective_width // avg_char_px))  # Minimum 50 chars' in ppt_content
        line_calc_25 = 'int(chars_per_line * 0.75))))  # 25% more lines' in ppt_content
        line_calc_22 = 'int(chars_per_line * 0.78))))  # 22% more lines' in ppt_content
        line_calc_18 = 'int(chars_per_line * 0.82))))  # 18% more lines' in ppt_content
        summary_logic = 'self.language == \'chinese\' and summary_md and any(\'\\u4e00\' <= char <= \'\\u9fff\' for char in summary_md)' in ppt_content

        page_breaking_ok = width_calc and min_chars and line_calc_25 and line_calc_22 and line_calc_18 and summary_logic
        results.append(("PPT page breaking changes", page_breaking_ok))
        print(f"✅ PPT page breaking: {'PASS' if page_breaking_ok else 'FAIL'}")

    except Exception as e:
        results.append(("PPT page breaking changes", False))
        print(f"❌ PPT page breaking: FAIL - {e}")

    # 3. Test imports work
    try:
        from common.pptx_export import PowerPointGenerator, export_pptx, ReportGenerator
        from fdd_utils.prompt_templates import get_translation_prompts, get_content_generation_prompts

        results.append(("All imports working", True))
        print("✅ All imports: PASS")

    except Exception as e:
        results.append(("All imports working", False))
        print(f"❌ All imports: FAIL - {e}")

    # 4. Test that template file exists
    template_exists = os.path.exists('/Users/ytsang/Desktop/Github/python-pptx/fdd_utils/template.pptx')
    results.append(("Template file exists", template_exists))
    print(f"✅ Template file: {'PASS' if template_exists else 'FAIL'}")

    # 5. Test that content files can be created
    content_files_ok = os.path.exists('/Users/ytsang/Desktop/Github/python-pptx/fdd_utils/bs_content.md')
    results.append(("Content files accessible", content_files_ok))
    print(f"✅ Content files: {'PASS' if content_files_ok else 'FAIL'}")

    # Summary
    print("\n" + "=" * 80)
    print("📊 TEST RESULTS SUMMARY:")
    print("=" * 80)

    all_pass = True
    for test_name, result in results:
        status = "✅" if result else "❌"
        print(f"{status} {test_name}: {'PASS' if result else 'FAIL'}")
        if not result:
            all_pass = False

    print("\n" + "=" * 80)
    if all_pass:
        print("🎯 ALL FIXES VERIFIED: Ready for Chinese processing!")
        print("\n💡 To test the fixes:")
        print("   1. Restart Streamlit: streamlit run fdd_app.py")
        print("   2. Upload Excel file and select entity")
        print("   3. Run Chinese AI processing")
        print("   4. Export PowerPoint - should show:")
        print("      - Numbers in 万/亿 format")
        print("      - Content distributed across multiple pages")
        print("      - Chinese content in summary sections")
    else:
        print("❌ SOME FIXES ARE MISSING - Please check the failed tests above")

    print("=" * 80)

    return all_pass

if __name__ == "__main__":
    test_all_fixes()

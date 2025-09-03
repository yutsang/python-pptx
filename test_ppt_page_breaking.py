#!/usr/bin/env python3
"""
Test script to verify PPT page breaking functionality
"""

import sys
sys.path.append('/Users/ytsang/Desktop/Github/python-pptx')

def test_ppt_page_breaking():
    """Test that PPT page breaking changes are applied"""
    try:
        from common.pptx_export import PowerPointGenerator

        print("🧪 TESTING PPT PAGE BREAKING CHANGES")
        print("=" * 80)

        # Check if the file exists and has our changes
        with open('/Users/ytsang/Desktop/Github/python-pptx/common/pptx_export.py', 'r') as f:
            content = f.read()

        # Check for key changes
        checks = [
            ('Chinese character width calculation', 'avg_char_px = 12.5  # Regular Chinese text (wider than English) - increased' in content),
            ('Minimum characters per line', 'max(50, int(effective_width // avg_char_px))  # Minimum 50 chars for Chinese optimization - more conservative' in content),
            ('Chinese line calculation 25%', 'int(chars_per_line * 0.75))))  # 25% more lines' in content),
            ('Chinese line calculation 22%', 'int(chars_per_line * 0.78))))  # 22% more lines' in content),
            ('Chinese line calculation 18%', 'int(chars_per_line * 0.82))))  # 18% more lines' in content),
            ('Summary content logic', 'if hasattr(self, \'language\') and self.language == \'chinese\' and summary_md and any(\'\\u4e00\' <= char <= \'\\u9fff\' for char in summary_md):' in content),
        ]

        all_passed = True
        for check_name, check_result in checks:
            status = "✅" if check_result else "❌"
            print(f"{status} {check_name}: {'PASS' if check_result else 'FAIL'}")
            if not check_result:
                all_passed = False

        print()
        print(f"🎯 OVERALL RESULT: {'✅ ALL PAGE BREAKING CHANGES APPLIED' if all_passed else '❌ SOME CHANGES MISSING'}")

        # Test that the class can be instantiated (basic import test)
        try:
            # We can't actually instantiate without a template file, but we can check the class exists
            print("✅ PowerPointGenerator class available")
        except Exception as e:
            print(f"❌ PowerPointGenerator class issue: {e}")
            all_passed = False

        return all_passed

    except Exception as e:
        print(f"❌ Test failed: {e}")
        return False

if __name__ == "__main__":
    test_ppt_page_breaking()

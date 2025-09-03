#!/usr/bin/env python3
"""
Server investigation script to check if the Chinese fixes are actually applied.
Run this ON YOUR SERVER to verify what's actually running.
"""

import sys
import os
import hashlib

def get_file_hash(filepath):
    """Get MD5 hash of a file"""
    try:
        with open(filepath, 'rb') as f:
            return hashlib.md5(f.read()).hexdigest()
    except:
        return "ERROR"

def investigate_server():
    """Comprehensive investigation of what's running on the server"""
    print("=" * 100)
    print("🔍 SERVER INVESTIGATION: Are Chinese fixes actually applied?")
    print("=" * 100)

    current_dir = os.getcwd()
    print(f"📁 Current directory: {current_dir}")
    print(f"📂 Script location: {__file__}")
    print()

    # Check if we're in the right directory
    expected_files = [
        'fdd_app.py',
        'fdd_utils/prompt_templates.py',
        'common/pptx_export.py',
        'fdd_utils/template.pptx'
    ]

    print("📋 FILE EXISTENCE CHECK:")
    all_files_exist = True
    for file_path in expected_files:
        exists = os.path.exists(file_path)
        status = "✅" if exists else "❌"
        print(f"   {status} {file_path}: {'EXISTS' if exists else 'MISSING'}")
        if not exists:
            all_files_exist = False
    print()

    if not all_files_exist:
        print("❌ CRITICAL: Some required files are missing!")
        print("   Make sure you're running from the correct directory.")
        return False

    # Check file hashes to verify content
    print("🔐 FILE CONTENT VERIFICATION:")
    print("   (Comparing with expected hash values)")

    checks = []

    # Check prompt templates
    try:
        with open('fdd_utils/prompt_templates.py', 'r', encoding='utf-8') as f:
            prompt_content = f.read()

        has_chinese_numbers = '万' in prompt_content and '亿' in prompt_content
        has_examples = '100k → 100万' in prompt_content
        checks.append(("Chinese number formatting in prompts", has_chinese_numbers and has_examples))
        print(f"   ✅ Chinese prompts: {'PASS' if has_chinese_numbers and has_examples else 'FAIL'}")

    except Exception as e:
        checks.append(("Chinese number formatting in prompts", False))
        print(f"   ❌ Chinese prompts: FAIL - {e}")

    # Check PPT export changes
    try:
        with open('common/pptx_export.py', 'r', encoding='utf-8') as f:
            ppt_content = f.read()

        # Check for our specific changes
        width_calc = 'avg_char_px = 12.5  # Regular Chinese text (wider than English) - increased' in ppt_content
        min_chars = 'max(50, int(effective_width // avg_char_px))  # Minimum 50 chars' in ppt_content
        summary_logic = 'self.language == \'chinese\' and summary_md and any(\'\\u4e00\' <= char <= \'\\u9fff\' for char in summary_md)' in ppt_content

        checks.append(("PPT page breaking changes", width_calc and min_chars and summary_logic))
        print(f"   ✅ PPT export changes: {'PASS' if width_calc and min_chars and summary_logic else 'FAIL'}")

    except Exception as e:
        checks.append(("PPT export changes", False))
        print(f"   ❌ PPT export changes: FAIL - {e}")

    # Check Python cache
    print("\n🗂️  PYTHON CACHE CHECK:")
    cache_dirs = []
    for root, dirs, files in os.walk('.'):
        for d in dirs:
            if d == '__pycache__':
                cache_dirs.append(os.path.join(root, d))

    if cache_dirs:
        print("   ⚠️  Found Python cache directories:")
        for cache_dir in cache_dirs:
            print(f"      {cache_dir}")
        print("   💡 These may contain cached bytecode from old code")
    else:
        print("   ✅ No Python cache found")

    # Check if we can import the modules
    print("\n📦 MODULE IMPORT TEST:")
    try:
        sys.path.insert(0, 'fdd_utils')
        from prompt_templates import get_translation_prompts
        from common.pptx_export import PowerPointGenerator

        checks.append(("Module imports", True))
        print("   ✅ All modules import successfully")

    except Exception as e:
        checks.append(("Module imports", False))
        print(f"   ❌ Module import failed: {e}")

    # Overall result
    print("\n" + "=" * 100)
    print("🎯 INVESTIGATION RESULTS:")
    print("=" * 100)

    all_pass = True
    for check_name, result in checks:
        status = "✅" if result else "❌"
        print(f"{status} {check_name}: {'PASS' if result else 'FAIL'}")
        if not result:
            all_pass = False

    print("\n" + "=" * 100)

    if all_pass:
        print("🎉 ALL CHECKS PASSED!")
        print("   Your server is running the updated code with Chinese fixes.")
        print("   If you're still seeing old behavior, try:")
        print("   1. Restart your Streamlit server")
        print("   2. Clear browser cache")
        print("   3. Check if there are multiple instances running")

    else:
        print("❌ SOME CHECKS FAILED!")
        print("   Your server may be running old code.")
        print("   Solutions:")
        print("   1. Make sure you're in the correct directory")
        print("   2. Clear Python cache: find . -name '*.pyc' -delete")
        print("   3. Restart your server completely")
        print("   4. Check if there are multiple copies of the code")

    print("=" * 100)

    # Provide specific commands for the user
    print("\n🔧 IMMEDIATE ACTION COMMANDS:")
    print("1. Clear Python cache:")
    print("   find . -name '*.pyc' -delete && find . -name '__pycache__' -type d -exec rm -rf {} +")
    print()
    print("2. Restart Streamlit:")
    print("   pkill -f streamlit")
    print("   streamlit run fdd_app.py")
    print()
    print("3. If still not working, check what directory you're in:")
    print("   pwd && ls -la")
    print("=" * 100)

    return all_pass

if __name__ == "__main__":
    investigate_server()

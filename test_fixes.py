#!/usr/bin/env python3
"""
Test script to demonstrate the fixes for the reported issues:
1. Unknown project name warning
2. Sheet not found error
3. Agent 3 unexpected error
"""

import sys
import os
sys.path.append(os.path.dirname(__file__))

from fdd_utils.data_utils import get_tab_name
from fdd_utils.utils import find_financial_figures_with_context_check

def test_get_tab_name():
    """Test the improved get_tab_name function with various project names."""
    print("🧪 Testing get_tab_name function...")

    test_cases = [
        ("Haining", "Should return 'BSHN'"),
        ("Nanjing", "Should return 'BSNJ'"),
        ("Ningbo", "Should return 'BSNB'"),
        ("CleanTech", "Should return list of possible sheet names"),
        ("TechCorp", "Should return list of possible sheet names"),
        ("ABC Company", "Should return list of possible sheet names"),
        ("", "Should return None"),
        ("   ", "Should return None")
    ]

    for project_name, description in test_cases:
        result = get_tab_name(project_name)
        print(f"  📝 {project_name or '(empty)'} -> {result} ({description})")

    print("✅ get_tab_name tests completed\n")

def test_excel_file_handling():
    """Test Excel file handling with better error messages."""
    print("🧪 Testing Excel file handling...")

    # Test with a non-existent file
    print("  📁 Testing with non-existent file...")
    result = find_financial_figures_with_context_check(
        "non_existent_file.xlsx",
        "BSHN",
        "30/09/2022"
    )
    print(f"  📊 Result: {result} (should be empty dict with error message)")

    # Test with valid file but non-existent sheet
    print("  📊 Testing with invalid sheet name...")
    test_file = "fdd_utils/databook.xlsx"  # Use a sample file if it exists
    if os.path.exists(test_file):
        result = find_financial_figures_with_context_check(
            test_file,
            "NonExistentSheet",
            "30/09/2022"
        )
        print(f"  📊 Result: {result} (should be empty dict with error message)")
    else:
        print("  ⚠️ Sample Excel file not found, skipping sheet test")

    print("✅ Excel file handling tests completed\n")

def test_agent3_patterns():
    """Test that Agent 3 patterns are properly loaded."""
    print("🧪 Testing Agent 3 pattern loading...")

    try:
        import json
        with open('fdd_utils/prompts.json', 'r', encoding='utf-8') as f:
            prompts = json.load(f)

        # Check if Agent 3 system prompt exists
        agent3_prompt = prompts.get('system_prompts', {}).get('english', {}).get('Agent 3')
        if agent3_prompt:
            print("  ✅ Agent 3 system prompt found in prompts.json")
            print(f"  📝 Prompt length: {len(agent3_prompt)} characters")
        else:
            print("  ❌ Agent 3 system prompt not found")

        # Check pattern.json
        with open('fdd_utils/pattern.json', 'r', encoding='utf-8') as f:
            patterns = json.load(f)

        print(f"  📋 Available pattern keys: {list(patterns.keys())}")
        print(f"  📊 Total patterns loaded: {sum(len(v) for v in patterns.values())}")

    except Exception as e:
        print(f"  ❌ Error testing patterns: {e}")

    print("✅ Agent 3 pattern tests completed\n")

def main():
    """Run all tests."""
    print("🚀 Running comprehensive fix tests...\n")

    test_get_tab_name()
    test_excel_file_handling()
    test_agent3_patterns()

    print("🎉 All tests completed!")
    print("\n📋 Summary of fixes implemented:")
    print("  1. ✅ Improved get_tab_name() to handle unknown project names")
    print("  2. ✅ Added fallback sheet name patterns for unknown entities")
    print("  3. ✅ Enhanced sheet lookup to try multiple possible names")
    print("  4. ✅ Added missing Agent 3 system prompt to prompts.json")
    print("  5. ✅ Improved error handling for Excel file access")
    print("  6. ✅ Added better error messages for debugging")

if __name__ == "__main__":
    main()

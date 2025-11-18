#!/usr/bin/env python3
"""
Test script to verify number formatting functionality
Tests both Chinese (万/亿) and English (K/million) formatting
"""

import sys
import os

# Add fdd_utils to path
sys.path.insert(0, os.path.dirname(__file__))

from fdd_utils.process_databook import format_value_by_language


def test_chinese_formatting():
    """Test Chinese number formatting (万/亿)"""
    print("=" * 60)
    print("Testing Chinese Formatting (万/亿)")
    print("=" * 60)
    
    test_cases = [
        (5000, "5000"),
        (50000, "5.0万"),
        (500000, "50.0万"),
        (5000000, "500.0万"),
        (50000000, "5000.0万"),
        (100000000, "1.0亿"),
        (500000000, "5.0亿"),
        (1234567890, "12.3亿"),
        (-50000, "-5.0万"),
        (-100000000, "-1.0亿"),
    ]
    
    for value, expected in test_cases:
        result = format_value_by_language(value, 'Chi')
        status = "✅" if result == expected else "❌"
        print(f"{status} {value:>15,} -> {result:>15} (Expected: {expected})")


def test_english_formatting():
    """Test English number formatting (K/million)"""
    print("\n" + "=" * 60)
    print("Testing English Formatting (K/million)")
    print("=" * 60)
    
    test_cases = [
        (5000, "5,000"),
        (10000, "10.0K"),
        (50000, "50.0K"),
        (99999, "100.0K"),  # Note: rounds to 100.0K
        (100000, "100.0K"),  # Still in K format (< 1 million)
        (500000, "500.0K"),  # Still in K format (< 1 million)
        (1000000, "1.0 million"),
        (5000000, "5.0 million"),
        (50000000, "50.0 million"),
        (123456789, "123.5 million"),  # Note: rounded to 1 decimal
        (-50000, "-50.0K"),
        (-5000000, "-5.0 million"),
    ]
    
    for value, expected in test_cases:
        result = format_value_by_language(value, 'Eng')
        status = "✅" if result == expected else "❌"
        print(f"{status} {value:>15,} -> {result:>20} (Expected: {expected})")


def test_edge_cases():
    """Test edge cases"""
    print("\n" + "=" * 60)
    print("Testing Edge Cases")
    print("=" * 60)
    
    # Zero
    result_chi = format_value_by_language(0, 'Chi')
    result_eng = format_value_by_language(0, 'Eng')
    print(f"Zero (Chi): {result_chi} (Expected: 0)")
    print(f"Zero (Eng): {result_eng} (Expected: 0)")
    
    # Very large numbers
    result_chi = format_value_by_language(9999999999, 'Chi')
    result_eng = format_value_by_language(9999999999, 'Eng')
    print(f"Very large (Chi): {result_chi}")
    print(f"Very large (Eng): {result_eng}")


if __name__ == "__main__":
    print("\n" + "=" * 60)
    print("NUMBER FORMATTING TEST SUITE")
    print("=" * 60 + "\n")
    
    test_chinese_formatting()
    test_english_formatting()
    test_edge_cases()
    
    print("\n" + "=" * 60)
    print("TEST COMPLETED")
    print("=" * 60 + "\n")


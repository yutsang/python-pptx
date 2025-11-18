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
    """Test Chinese number formatting (万/亿) - 万 with 1 d.p., 亿 with 2 d.p."""
    print("=" * 60)
    print("Testing Chinese Formatting (万=1 d.p., 亿=2 d.p.)")
    print("=" * 60)
    
    test_cases = [
        (5000, "5000"),
        (50000, "5.0万"),
        (500000, "50.0万"),
        (5000000, "500.0万"),
        (50000000, "5000.0万"),
        (100000000, "1.00亿"),  # 2 decimal places for 亿
        (500000000, "5.00亿"),  # 2 decimal places for 亿
        (1234567890, "12.35亿"),  # 2 decimal places for 亿
        (-50000, "-5.0万"),
        (-100000000, "-1.00亿"),  # 2 decimal places for 亿
    ]
    
    for value, expected in test_cases:
        result = format_value_by_language(value, 'Chi')
        status = "✅" if result == expected else "❌"
        print(f"{status} {value:>15,} -> {result:>15} (Expected: {expected})")


def test_english_formatting():
    """Test English number formatting (K/million) - K with 1 d.p., million with 2 d.p."""
    print("\n" + "=" * 60)
    print("Testing English Formatting (K=1 d.p., million=2 d.p.)")
    print("=" * 60)
    
    test_cases = [
        (5000, "5,000"),
        (10000, "10.0K"),
        (50000, "50.0K"),
        (99999, "100.0K"),  # Note: rounds to 100.0K
        (100000, "100.0K"),  # Still in K format (< 1 million)
        (500000, "500.0K"),  # Still in K format (< 1 million)
        (1000000, "1.00 million"),  # 2 decimal places for million
        (5000000, "5.00 million"),  # 2 decimal places for million
        (50000000, "50.00 million"),  # 2 decimal places for million
        (123456789, "123.46 million"),  # 2 decimal places for million
        (-50000, "-50.0K"),
        (-5000000, "-5.00 million"),  # 2 decimal places for million
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
    print(f"Very large (Chi): {result_chi} (Expected: 100.00亿)")
    print(f"Very large (Eng): {result_eng} (Expected: 10000.00 million)")
    
    # Retained earnings test
    print("\n" + "-" * 60)
    print("Retained Earnings Special Handling")
    print("-" * 60)
    print("Note: This is handled at DataFrame level in process_databook.py")
    print("Negative 未分配利润 → 未弥补亏损 (value shown as positive)")
    print("Negative Retained Earnings → Accumulated Losses (value shown as positive)")


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


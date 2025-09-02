#!/usr/bin/env python3
"""
Debug script to identify the root cause of AI processing failures.
"""

import sys
import os
sys.path.append(os.path.dirname(__file__))

from common.assistant import (
    load_config, initialize_ai_services, process_keys,
    find_financial_figures_with_context_check, get_tab_name,
    load_ip, process_and_filter_excel
)
import json
import tempfile

def debug_ai_processing():
    """Debug the complete AI processing pipeline."""
    print("🔍 Starting AI processing debug...")

    # Test 1: Check file availability
    print("\n📁 Checking file availability...")
    required_files = [
        'fdd_utils/config.json',
        'fdd_utils/mapping.json',
        'fdd_utils/pattern.json',
        'fdd_utils/prompts.json'
    ]

    for file_path in required_files:
        if os.path.exists(file_path):
            print(f"  ✅ {file_path} exists")
        else:
            print(f"  ❌ {file_path} missing")

    # Test 2: Load configuration
    print("\n⚙️ Testing configuration loading...")
    try:
        config = load_config('fdd_utils/config.json')
        print("  ✅ Config loaded successfully")
    except Exception as e:
        print(f"  ❌ Config loading failed: {e}")
        return

    # Test 3: Initialize AI client
    print("\n🤖 Testing AI client initialization...")
    try:
        client, _ = initialize_ai_services(config)
        print("  ✅ AI client initialized successfully")
    except Exception as e:
        print(f"  ❌ AI client initialization failed: {e}")
        return

    # Test 4: Test Excel file access (if available)
    print("\n📊 Testing Excel file access...")
    excel_files = [
        'databook.xlsx',
        'fdd_utils/databook.xlsx'
    ]

    excel_file = None
    for ef in excel_files:
        if os.path.exists(ef):
            excel_file = ef
            print(f"  ✅ Found Excel file: {ef}")
            break

    if not excel_file:
        print("  ⚠️ No Excel file found for testing")
        return

    # Test 5: Test sheet name resolution
    print("\n📋 Testing sheet name resolution...")
    test_entities = ['Haining', 'Nanjing', 'Ningbo', 'TestCompany']

    for entity in test_entities:
        sheet_names = get_tab_name(entity)
        print(f"  📝 {entity} -> {sheet_names}")

    # Test 6: Test financial figure extraction
    print("\n💰 Testing financial figure extraction...")
    try:
        entity_name = 'Haining'
        sheet_names = get_tab_name(entity_name)
        financial_figures = find_financial_figures_with_context_check(
            excel_file, sheet_names, '30/09/2022'
        )
        print(f"  ✅ Financial figures extracted: {len(financial_figures)} keys")
        if financial_figures:
            print(f"     Sample: {list(financial_figures.keys())[:3]}")
    except Exception as e:
        print(f"  ❌ Financial figure extraction failed: {e}")

    # Test 7: Test data processing
    print("\n🔄 Testing data processing...")
    try:
        entity_name = 'Haining'
        entity_helpers = ['Haining']
        mapping = load_ip('fdd_utils/mapping.json')
        excel_tables = process_and_filter_excel(
            excel_file, mapping, entity_name, entity_helpers
        )
        print(f"  ✅ Data processing successful: {len(excel_tables)} chars")
    except Exception as e:
        print(f"  ❌ Data processing failed: {e}")

    # Test 8: Test pattern loading
    print("\n📋 Testing pattern loading...")
    test_keys = ['Cash', 'AR', 'AP']
    for key in test_keys:
        try:
            patterns = load_ip('fdd_utils/pattern.json', key)
            if patterns:
                print(f"  ✅ {key}: {len(patterns)} patterns loaded")
            else:
                print(f"  ⚠️ {key}: No patterns found")
        except Exception as e:
            print(f"  ❌ {key}: Pattern loading failed: {e}")

    # Test 9: Test prompt loading
    print("\n💬 Testing prompt loading...")
    try:
        with open('fdd_utils/prompts.json', 'r', encoding='utf-8') as f:
            prompts = json.load(f)
        agent1_prompt = prompts.get('system_prompts', {}).get('english', {}).get('Agent 1')
        if agent1_prompt:
            print(f"  ✅ Agent 1 prompt loaded: {len(agent1_prompt)} chars")
        else:
            print("  ❌ Agent 1 prompt not found")
    except Exception as e:
        print(f"  ❌ Prompt loading failed: {e}")

    print("\n🎯 Debug completed!")

if __name__ == "__main__":
    debug_ai_processing()

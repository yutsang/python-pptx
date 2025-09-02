#!/usr/bin/env python3
"""
Detailed debug script to test the actual AI processing pipeline.
"""

import sys
import os
sys.path.append(os.path.dirname(__file__))

from common.assistant import (
    load_config, initialize_ai_services, generate_response,
    find_financial_figures_with_context_check, get_tab_name,
    load_ip, process_and_filter_excel, get_financial_figure
)
import json

def test_actual_ai_processing():
    """Test the actual AI processing that would happen in process_keys."""
    print("🔍 Testing actual AI processing pipeline...")

    # Setup
    entity_name = 'Haining'
    key = 'Cash'
    excel_file = 'databook.xlsx'

    if not os.path.exists(excel_file):
        print(f"❌ Excel file {excel_file} not found")
        return

    try:
        # Load configuration
        config = load_config('fdd_utils/config.json')
        client, _ = initialize_ai_services(config)

        # Get financial figures
        sheet_names = get_tab_name(entity_name)
        financial_figures = find_financial_figures_with_context_check(
            excel_file, sheet_names, '30/09/2022'
        )

        # Get table data
        mapping = load_ip('fdd_utils/mapping.json')
        excel_tables = process_and_filter_excel(
            excel_file, mapping, entity_name, [entity_name]
        )

        # Load patterns and prompts
        patterns = load_ip('fdd_utils/pattern.json', key)
        with open('fdd_utils/prompts.json', 'r', encoding='utf-8') as f:
            prompts_config = json.load(f)
        system_prompt = prompts_config['system_prompts']['english']['Agent 1']

        # Create user prompt (similar to what process_keys does)
        pattern_json = json.dumps(patterns, indent=2)
        financial_figure_info = get_financial_figure(financial_figures, key)

        user_prompt = f"""
        TASK: Select ONE pattern and complete it with actual data

        AVAILABLE PATTERNS: {pattern_json}

        FINANCIAL FIGURE: {key}: {financial_figure_info}

        DATA SOURCE: {excel_tables[:2000]}  # Limit for testing

        SELECTION CRITERIA:
        - Choose the pattern with the most complete data coverage
        - Prioritize patterns that match the primary account category
        - Use most recent data: latest available

        REQUIRED OUTPUT FORMAT:
        - Only the completed pattern text
        - No pattern names or labels
        - No template structure
        - No JSON formatting
        - Replace ALL 'xxx' or placeholders with actual data values
        - Do not use bullet point for listing

        Example of CORRECT output format:
        "Cash at bank comprises deposits of $2.3M held with major financial institutions as at 30/09/2022."

        Example of INCORRECT output format:
        "Pattern 1: Cash at bank comprises deposits of xxx held with xxx as at xxx."
        """

        print(f"📝 Testing AI call for key: {key}")
        print(f"💰 Financial figure: {financial_figure_info}")
        print(f"📊 Table data length: {len(excel_tables)} characters")
        print(f"📋 Patterns available: {len(patterns)}")
        print(f"🤖 System prompt length: {len(system_prompt)} characters")
        print(f"💬 User prompt length: {len(user_prompt)} characters")

        # Make the actual AI call
        response = generate_response(
            user_query=user_prompt,
            system_prompt=system_prompt,
            oai_client=client,
            context_content=excel_tables,
            openai_chat_model='deepseek-chat',
            entity_name=entity_name,
            use_local_ai=False
        )

        print("✅ AI call successful!")
        print(f"📄 Response length: {len(response)} characters")
        print(f"📄 Response preview: {response[:200]}...")

        return True

    except Exception as e:
        print(f"❌ AI processing failed: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    test_actual_ai_processing()

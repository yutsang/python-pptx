"""
Content processing utilities for FDD application
Functions for loading, parsing, and processing financial content
"""

import os
import json
import re
import streamlit as st
from datetime import datetime

def display_bs_content_by_key(md_path):
    """
    Display balance sheet content by key from markdown file
    """
    try:
        with open(md_path, 'r', encoding='utf-8') as f:
            content = f.read()
        # Split by key headers (e.g., ## Cash, ## AR, etc.)
        sections = re.split(r'(^## .+$)', content, flags=re.MULTILINE)
        if len(sections) < 2:
            st.markdown(content)
            return
        for i in range(1, len(sections), 2):
            key_header = sections[i].strip()
            key_content = sections[i+1].strip() if i+1 < len(sections) else ''
            with st.expander(key_header, expanded=False):
                st.markdown(key_content)
    except Exception as e:
        st.error(f"Error reading {md_path}: {e}")

def clean_content_quotes(content):
    """
    Clean up content by removing unnecessary quotation marks while preserving legitimate quotes
    """
    if not content:
        return content

    # Split content into lines to process each line separately
    lines = content.split('\n')
    cleaned_lines = []

    for line in lines:
        line = line.strip()

        # Skip empty lines
        if not line:
            cleaned_lines.append('')
            continue

        # Remove quotation marks that wrap entire sections (not legitimate quotes)
        if line.startswith('"') and line.endswith('"'):
            # Check if this is a legitimate quote by looking for sentence structure
            inner_content = line[1:-1].strip()
            if not inner_content.startswith(('The', 'A', 'An', 'In', 'On', 'At', 'By', 'For')):
                line = inner_content

        # Remove quotation marks from lines that are entirely wrapped in quotes
        # but preserve quotes around actual quoted text
        if len(line) > 2 and line.startswith('"') and line.endswith('"'):
            # Check for punctuation before the closing quote to identify legitimate quotes
            if not line[-2] in '.!?' and not any(char in line[:-1] for char in '.!?'):
                line = line[1:-1]

        cleaned_lines.append(line)

    return '\n'.join(cleaned_lines)

def load_json_content():
    """Load content from JSON file with caching for better performance"""
    try:
        # Try JSON first (better performance)
        json_file = "fdd_utils/bs_content.json"
        if os.path.exists(json_file):
            with open(json_file, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception as e:
        print(f"Error loading JSON content: {e}")

    # Fallback to parsing markdown if JSON not available
    try:
        content_files = ["fdd_utils/bs_content.md", "fdd_utils/bs_content_ai_generated.md"]
        for file_path in content_files:
            if os.path.exists(file_path):
                return parse_markdown_to_json(file_path)
    except Exception as e:
        print(f"Error parsing markdown fallback: {e}")

    return None

def parse_markdown_to_json(md_file_path):
    """Parse markdown file and convert to JSON-like structure for compatibility"""
    try:
        with open(md_file_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # Parse markdown into structured format
        sections = re.split(r'(^### .+$)', content, flags=re.MULTILINE)
        parsed_data = {"financial_items": {}}

        for i in range(1, len(sections), 2):
            if i + 1 < len(sections):
                header = sections[i].strip().replace('### ', '')
                section_content = sections[i + 1].strip()

                # Map headers back to keys
                header_to_key_mapping = {
                    'Cash at bank': 'Cash',
                    'Accounts Receivable': 'AR',
                    'Prepayments': 'Prepayments',
                    'Other Receivables': 'OR',
                    'Other current assets': 'Other CA',
                    'Investment Property': 'IP',
                    'Other non-current assets': 'Other NCA',
                    'Accounts Payable': 'AP',
                    'Taxes payable': 'Taxes payable',
                    'Other Payables': 'OP',
                    'Capital': 'Capital',
                    'Reserves': 'Reserve',
                    'Operating Income': 'OI',
                    'Operating Cost': 'OC',
                    'Tax and Surcharges': 'Tax and Surcharges',
                    'Management Fees': 'GA',
                    'Finance Expense': 'Fin Exp',
                    'Credit Loss': 'Cr Loss',
                    'Other Income': 'Other Income',
                    'Non-operating Income': 'Non-operating Income',
                    'Non-operating Expense': 'Non-operating Exp',
                    'Income Tax': 'Income tax',
                    'Deferred Tax Assets': 'LT DTA'
                }

                key = header_to_key_mapping.get(header, header)
                parsed_data["financial_items"][key] = {
                    "content": section_content,
                    "header": header
                }

        return parsed_data
    except Exception as e:
        print(f"Error parsing markdown file {md_file_path}: {e}")
        return None

def get_content_from_json(key):
    """Get content for a specific key from loaded JSON data"""
    content = load_json_content()
    if not content:
        return f"No content available for {key}"

    # Try different paths to find the content
    if "financial_items" in content and key in content["financial_items"]:
        return content["financial_items"][key].get("content", f"No content for {key}")

    # Fallback for different JSON structures
    if key in content:
        item_data = content[key]
        if isinstance(item_data, dict):
            return item_data.get("content", item_data.get("corrected_content", str(item_data)))
        elif isinstance(item_data, str):
            return item_data

    return f"No content found for key: {key}"

def generate_content_from_session_storage(entity_name):
    """Generate content files (JSON + Markdown) from session state storage (PERFORMANCE OPTIMIZED)"""
    try:
        # Get content from session state storage (fastest method)
        content_store = st.session_state.get('ai_content_store', {})
        current_statement_type = st.session_state.get('current_statement_type', 'BS')

        if not content_store:
            st.error("❌ No AI-generated content available. Please run AI processing first.")
            return

        # Get current statement type from session state
        current_statement_type = st.session_state.get('current_statement_type', 'BS')

        # Get category mappings from centralized config
        from fdd_utils.category_config import get_category_mapping
        category_mapping, name_mapping = get_category_mapping(current_statement_type, entity_name)

        # Generate JSON content from session storage (for AI2 easy access)
        json_content = {
            'metadata': {
                'entity_name': entity_name,
                'generated_timestamp': datetime.now().isoformat(),
                'session_id': getattr(st.session_state.get('ai_logger'), 'session_id', 'default'),
                'statement_type': current_statement_type,
                'total_keys': len(content_store)
            },
            'financial_items': {}
        }

        # Process each key in the content store
        for key, content_data in content_store.items():
            if isinstance(content_data, dict):
                # Get the most appropriate content (prioritize corrected_content)
                final_content = (content_data.get('corrected_content') or
                               content_data.get('content') or
                               content_data.get('agent3_final') or
                               'No content available')

                json_content['financial_items'][key] = {
                    'content': final_content,
                    'display_name': name_mapping.get(key, key),
                    'category': get_category_from_key(key, category_mapping),
                    'last_updated': content_data.get('timestamp', datetime.now().isoformat()),
                    'source_agent': content_data.get('source', 'unknown')
                }

        # Save JSON file
        json_filename = f"fdd_utils/bs_content.json"
        with open(json_filename, 'w', encoding='utf-8') as f:
            json.dump(json_content, f, indent=2, ensure_ascii=False)

        # Generate markdown file for human readability
        markdown_content = f"# Balance Sheet Content - {entity_name}\n\n"
        markdown_content += f"**Generated:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        markdown_content += f"**Statement Type:** {current_statement_type}\n\n"

        # Group by categories
        categorized_content = {}
        for category_name, keys_in_category in category_mapping.items():
            categorized_content[category_name] = []
            for key in keys_in_category:
                if key in json_content['financial_items']:
                    categorized_content[category_name].append(key)

        # Generate markdown by category
        for category_name, keys_in_category in categorized_content.items():
            if keys_in_category:
                markdown_content += f"## {category_name}\n\n"
                for key in keys_in_category:
                    item_data = json_content['financial_items'][key]
                    markdown_content += f"### {item_data['display_name']} ({key})\n\n"
                    markdown_content += f"{item_data['content']}\n\n"

        # Save markdown file
        md_filename = f"fdd_utils/bs_content.md"
        with open(md_filename, 'w', encoding='utf-8') as f:
            f.write(markdown_content)

        st.success(f"✅ Content files generated successfully!\n- JSON: {json_filename}\n- Markdown: {md_filename}")

    except Exception as e:
        st.error(f"Error generating content files: {e}")
        print(f"Error details: {e}")

def get_category_from_key(key, category_mapping):
    """Get category name for a given key"""
    for category_name, keys in category_mapping.items():
        if key in keys:
            return category_name
    return "Other"

def generate_markdown_from_ai_results(ai_results, entity_name):
    """Generate markdown content file from AI results following the old version pattern"""
    try:
        if not ai_results:
            st.error("❌ No AI results available to generate markdown")
            return

        markdown_content = f"# AI Generated Financial Content - {entity_name}\n\n"
        markdown_content += f"**Generated:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"

        # Process each key in the AI results
        for key, result in ai_results.items():
            if isinstance(result, dict) and 'content' in result:
                content = result['content']
                # Clean up the content
                content = clean_content_quotes(content)
                markdown_content += f"## {key}\n\n{content}\n\n"

        # Save to file
        filename = f"fdd_utils/bs_content_ai_generated.md"
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(markdown_content)

        st.success(f"✅ AI results saved to {filename}")

    except Exception as e:
        st.error(f"Error generating markdown from AI results: {e}")

def convert_sections_to_markdown(sections_by_key):
    """Convert sections dictionary to markdown format"""
    try:
        markdown_content = "# Financial Data Sections\n\n"
        markdown_content += f"**Generated:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"

        for key, sections in sections_by_key.items():
            markdown_content += f"## {key}\n\n"
            if sections:
                for section in sections:
                    markdown_content += f"{section}\n\n"
            else:
                markdown_content += "*No data available*\n\n"

        return markdown_content
    except Exception as e:
        print(f"Error converting sections to markdown: {e}")
        return "# Error generating markdown\n\nError occurred during conversion."

def update_bs_content_with_agent_corrections(corrections_dict, entity_name, agent_name):
    """Update balance sheet content with agent corrections"""
    try:
        # Load existing content
        content = load_json_content()
        if not content:
            st.error("❌ No existing content found to update")
            return

        # Apply corrections
        updated_count = 0
        for key, correction in corrections_dict.items():
            if key in content.get('financial_items', {}):
                content['financial_items'][key]['content'] = correction
                content['financial_items'][key]['last_updated'] = datetime.now().isoformat()
                content['financial_items'][key]['last_agent'] = agent_name
                updated_count += 1

        # Save updated content
        json_filename = f"fdd_utils/bs_content.json"
        with open(json_filename, 'w', encoding='utf-8') as f:
            json.dump(content, f, indent=2, ensure_ascii=False)

        st.success(f"✅ Updated {updated_count} items with {agent_name} corrections")

    except Exception as e:
        st.error(f"Error updating content with corrections: {e}")

def read_bs_content_by_key(entity_name=None):
    """Read balance sheet content by key for display"""
    try:
        content = load_json_content()
        if not content:
            return {}

        # Extract financial items
        financial_items = content.get('financial_items', {})
        result = {}

        for key, item_data in financial_items.items():
            if isinstance(item_data, dict):
                result[key] = item_data.get('content', 'No content available')
            else:
                result[key] = str(item_data)

        return result

    except Exception as e:
        print(f"Error reading BS content by key: {e}")
        return {}

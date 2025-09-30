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

    # Handle case where content is a dict instead of string
    if isinstance(content, dict):
        # Convert dict to string representation
        content = str(content)

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
        # Get current statement type to determine which file to load
        current_statement_type = st.session_state.get('current_statement_type', 'BS')

        # Try JSON first (better performance)
        if current_statement_type == "IS":
            json_file = "fdd_utils/is_content.json"
            md_fallback_files = ["fdd_utils/is_content.md", "fdd_utils/is_content_ai_generated.md"]
        else:
            json_file = "fdd_utils/bs_content.json"
            md_fallback_files = ["fdd_utils/bs_content.md", "fdd_utils/bs_content_ai_generated.md"]

        if os.path.exists(json_file):
            with open(json_file, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception as e:
        print(f"Error loading JSON content: {e}")

    # Fallback to parsing markdown if JSON not available
    try:
        for file_path in md_fallback_files:
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
        # Get current statement type from session state first
        current_statement_type = st.session_state.get('current_statement_type', 'BS')

        print(f"📝 CONTENT GEN: Processing {entity_name} for {current_statement_type} statement")

        # Get content from session state storage (fastest method)
        content_store = st.session_state.get('ai_content_store', {})

        if not content_store:
            print(f"❌ CONTENT GEN: No AI content found in session state")
            return

        # Filter content based on statement type
        if current_statement_type == "IS":
            is_keys = ['OI', 'OC', 'Tax and Surcharges', 'GA', 'Fin Exp', 'Cr Loss', 'Other Income', 'Non-operating Income', 'Non-operating Exp', 'Income tax', 'LT DTA']
            relevant_content = {k: v for k, v in content_store.items() if k in is_keys}
            print(f"📝 CONTENT GEN: Found {len(relevant_content)} IS keys: {list(relevant_content.keys())}")

            # Debug: Show all available keys in content_store for IS mode
            all_keys = list(content_store.keys())
            print(f"📝 CONTENT GEN: All available keys in content_store: {all_keys}")
            print(f"📝 CONTENT GEN: Expected IS keys: {is_keys}")

            # Check if the expected IS keys are in the content_store but maybe with different casing
            for expected_key in is_keys:
                for actual_key in all_keys:
                    if expected_key.lower() == actual_key.lower():
                        print(f"📝 CONTENT GEN: Found case-insensitive match: '{expected_key}' -> '{actual_key}'")

        elif current_statement_type == "BS":
            bs_keys = ['Cash', 'AR', 'Prepayments', 'OR', 'Other CA', 'Other NCA', 'IP', 'NCA', 'AP', 'Taxes payable', 'OP', 'Capital', 'Reserve']
            relevant_content = {k: v for k, v in content_store.items() if k in bs_keys}
            print(f"📝 CONTENT GEN: Found {len(relevant_content)} BS keys: {list(relevant_content.keys())}")
        else:
            relevant_content = content_store
            print(f"📝 CONTENT GEN: Processing all {len(relevant_content)} keys")

        print(f"🔍 DEBUG CONTENT GENERATION: content_store type: {type(content_store)}")
        print(f"🔍 DEBUG CONTENT GENERATION: content_store length: {len(content_store)}")

        # Debug: Print first few items in content_store
        if content_store:
            sample_keys = list(content_store.keys())[:3]
            for key in sample_keys:
                content_item = content_store[key]
                print(f"🔍 DEBUG CONTENT GENERATION: Sample key '{key}' content: {type(content_item)}")
                if isinstance(content_item, dict):
                    print(f"🔍 DEBUG CONTENT GENERATION:   Keys: {list(content_item.keys())}")
                    if 'current_content' in content_item:
                        content_preview = content_item['current_content'][:100] + "..." if len(content_item['current_content']) > 100 else content_item['current_content']
                        print(f"🔍 DEBUG CONTENT GENERATION:   current_content preview: {content_preview}")
                    else:
                        print(f"🔍 DEBUG CONTENT GENERATION:   No current_content found")

        # Get detected language from session state
        detected_language = st.session_state.get('ai_data', {}).get('detected_language', 'chinese')

        # Get category mappings from centralized config
        from fdd_utils.category_config import get_category_mapping
        category_mapping, name_mapping = get_category_mapping(current_statement_type, entity_name, detected_language)

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
                # Get the most appropriate content (prioritize current_content, then agent3, then agent1)
                final_content = (content_data.get('current_content') or
                               content_data.get('agent3_content') or
                               content_data.get('agent1_content') or
                               content_data.get('corrected_content') or
                               content_data.get('content') or
                               'No content available')

                # Debug: Print content info
                print(f"🔍 DEBUG CONTENT GENERATION: Key '{key}' - final_content length: {len(final_content) if final_content else 0}")
                print(f"🔍 DEBUG CONTENT GENERATION: Key '{key}' - content_data keys: {list(content_data.keys())}")

                # Apply item limiting to enforce "top 2 only" rule
                limited_content = limit_commentary_items(final_content)

                json_content['financial_items'][key] = {
                    'content': limited_content,
                    'original_content': final_content,  # Keep original for reference
                    'display_name': name_mapping.get(key, key),
                    'category': get_category_from_key(key, category_mapping),
                    'last_updated': content_data.get('timestamp', datetime.now().isoformat()),
                    'source_agent': content_data.get('source', 'unknown')
                }

        # Save JSON file based on statement type
        if current_statement_type == "IS":
            json_filename = f"fdd_utils/is_content.json"
            md_filename = f"fdd_utils/is_content.md"
        else:
            json_filename = f"fdd_utils/bs_content.json"
            md_filename = f"fdd_utils/bs_content.md"

        with open(json_filename, 'w', encoding='utf-8') as f:
            json.dump(json_content, f, indent=2, ensure_ascii=False)

        # Generate markdown file for human readability (without metadata headers for PowerPoint)
        markdown_content = ""

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
                    # Apply item limiting to enforce "top 2 only" rule
                    limited_content = limit_commentary_items(item_data['content'])
                    markdown_content += f"{limited_content}\n\n"

        # Save markdown file
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

        markdown_content = ""

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
        markdown_content = ""

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

        # Save updated content based on statement type
        current_statement_type = st.session_state.get('current_statement_type', 'BS')
        if current_statement_type == "IS":
            json_filename = f"fdd_utils/is_content.json"
        else:
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


def limit_commentary_items(content: str, max_items: int = 2) -> str:
    """
    Limit the number of items in commentary lists to max_items (default: 2)
    This enforces the "top 2 only" rule from the AI prompts.
    """
    if not content:
        return content

    lines = content.split('\n')
    result_lines = []

    for line in lines:
        # Skip empty lines
        if not line.strip():
            result_lines.append(line)
            continue

        # Check if this line contains a list of items (patterns like "item1, item2, item3" or "item1 and item2")
        # Look for patterns with commas or "and" that suggest multiple items
        list_patterns = [
            r'([^,]+),\s*([^,]+),\s*([^,]+)',  # item1, item2, item3
            r'([^,]+),\s*([^,]+),\s*and\s+([^,]+)',  # item1, item2, and item3
            r'([^,]+)\s+and\s+([^,]+)\s+and\s+([^,]+)',  # item1 and item2 and item3
        ]

        limited_line = line
        for pattern in list_patterns:
            matches = re.finditer(pattern, line, re.IGNORECASE)
            for match in matches:
                # Extract all groups (items)
                items = [group.strip() for group in match.groups()]

                # If we have more than max_items, keep only the first max_items
                if len(items) > max_items:
                    if 'and' in line.lower():
                        # For "and" patterns, replace with "and" between first two items
                        limited_items = items[:max_items]
                        if len(items) > max_items:
                            limited_line = limited_line.replace(match.group(0), f"{limited_items[0]} and {limited_items[1]}")
                    else:
                        # For comma patterns, replace with comma between first two items
                        limited_items = items[:max_items]
                        if len(items) > max_items:
                            limited_line = limited_line.replace(match.group(0), f"{limited_items[0]}, {limited_items[1]}")

        result_lines.append(limited_line)

    return '\n'.join(result_lines)

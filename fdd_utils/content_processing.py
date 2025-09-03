"""
Content processing utilities for FDD application
"""

import json
import os
from datetime import datetime
import streamlit as st

def display_bs_content_by_key(md_path):
    """Display BS content by key from markdown file"""
    try:
        if not os.path.exists(md_path):
            st.error(f"Content file not found: {md_path}")
            return

        with open(md_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # Parse markdown and display by sections
        sections = content.split('## ')
        for section in sections[1:]:  # Skip first empty section
            lines = section.strip().split('\n')
            if lines:
                section_title = lines[0].strip()
                st.markdown(f"## {section_title}")

                for line in lines[1:]:
                    line = line.strip()
                    if line.startswith('### '):
                        st.markdown(line)
                    elif line and not line.startswith('#'):
                        st.write(line)

    except Exception as e:
        st.error(f"Error displaying content: {e}")

def load_json_content():
    """Load JSON content from bs_content.json"""
    try:
        json_file_path = 'fdd_utils/bs_content.json'
        if os.path.exists(json_file_path):
            with open(json_file_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        return None
    except Exception as e:
        st.error(f"Error loading JSON content: {e}")
        return None

def parse_markdown_to_json(md_file_path):
    """Parse markdown file to JSON format for AI processing"""
    try:
        if not os.path.exists(md_file_path):
            return None

        with open(md_file_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # Parse markdown structure
        lines = content.split('\n')
        json_content = {
            'metadata': {
                'generated_at': datetime.now().isoformat(),
                'format_version': '1.0',
                'description': 'Parsed markdown content'
            },
            'categories': {},
            'keys': {}
        }

        current_category = None
        current_key = None

        for line in lines:
            line = line.strip()
            if line.startswith('## '):
                current_category = line[3:].strip()
                json_content['categories'][current_category] = []
            elif line.startswith('### '):
                if current_category:
                    current_key = line[4:].strip()
                    key_info = {
                        'key': current_key,
                        'display_name': current_key,
                        'content': '',
                        'content_source': 'markdown_parsed',
                        'category': current_category
                    }
                    json_content['categories'][current_category].append(key_info)
                    json_content['keys'][current_key] = key_info
            elif current_key and line and not line.startswith('#'):
                if json_content['keys'][current_key]['content']:
                    json_content['keys'][current_key]['content'] += '\n'
                json_content['keys'][current_key]['content'] += line

        return json_content

    except Exception as e:
        st.error(f"Error parsing markdown: {e}")
        return None

def get_content_from_json(key):
    """Get content for a specific key from JSON file"""
    try:
        json_content = load_json_content()
        if json_content and 'keys' in json_content and key in json_content['keys']:
            return json_content['keys'][key].get('content', '')
        return None
    except Exception as e:
        print(f"Error getting content from JSON: {e}")
        return None

def generate_content_from_session_storage(entity_name):
    """Generate content files (JSON + Markdown) from session state storage (PERFORMANCE OPTIMIZED)"""
    try:
        # Get content from session state storage (fastest method)
        content_store = st.session_state.get('ai_content_store', {})
        current_statement_type = st.session_state.get('current_statement_type', 'BS')

        if not content_store:
            st.error("❌ No AI-generated content available. Please run AI processing first.")
            return

        # Define category mappings based on entity name and statement type
        if current_statement_type == "IS":
            # Income Statement categories
            if entity_name in ['Ningbo', 'Nanjing']:
                name_mapping = {
                    'OI': '营业收入', 'OC': '营业成本', 'Tax and Surcharges': '税金及附加',
                    'GA': '管理费用', 'Fin Exp': '财务费用', 'Cr Loss': '信用损失',
                    'Other Income': '其他收益', 'Non-operating Income': '营业外收入',
                    'Non-operating Exp': '营业外支出', 'Income tax': '所得税', 'LT DTA': '递延所得税资产'
                }
                category_mapping = {
                    'Revenue': ['OI', 'Other Income', 'Non-operating Income'],
                    'Expenses': ['OC', 'GA', 'Fin Exp', 'Cr Loss', 'Non-operating Exp'],
                    'Taxes': ['Tax and Surcharges', 'Income tax'],
                    'Other': ['LT DTA']
                }
            else:  # Haining and others
                name_mapping = {
                    'OI': '营业收入', 'OC': '营业成本', 'Tax and Surcharges': '税金及附加',
                    'GA': '管理费用', 'Fin Exp': '财务费用', 'Cr Loss': '信用损失',
                    'Other Income': '其他收益', 'Non-operating Income': '营业外收入',
                    'Non-operating Exp': '营业外支出', 'Income tax': '所得税', 'LT DTA': '递延所得税资产'
                }
                category_mapping = {
                    'Revenue': ['OI', 'Other Income', 'Non-operating Income'],
                    'Expenses': ['OC', 'GA', 'Fin Exp', 'Cr Loss', 'Non-operating Exp'],
                    'Taxes': ['Tax and Surcharges', 'Income tax'],
                    'Other': ['LT DTA']
                }
        else:
            # Balance Sheet categories (default)
            if entity_name in ['Ningbo', 'Nanjing']:
                name_mapping = {
                    'Cash': '现金', 'AR': '应收账款', 'Prepayments': '预付款项',
                    'OR': '其他应收款', 'Other CA': '其他流动资产', 'Other NCA': '其他非流动资产',
                    'IP': '投资性房地产', 'NCA': '无形资产', 'AP': '应付账款',
                    'Taxes payable': '应交税费', 'OP': '其他应付款', 'Capital': '股本',
                    'Reserve': '资本公积'
                }
                category_mapping = {
                    'Current Assets': ['Cash', 'AR', 'Prepayments', 'OR', 'Other CA'],
                    'Non-current Assets': ['IP', 'Other NCA'],
                    'Liabilities': ['AP', 'Taxes payable', 'OP'],
                    'Equity': ['Capital']
                }
            else:  # Haining and others
                name_mapping = {
                    'Cash': '现金', 'AR': '应收账款', 'Prepayments': '预付款项',
                    'OR': '其他应收款', 'Other CA': '其他流动资产', 'Other NCA': '其他非流动资产',
                    'IP': '投资性房地产', 'NCA': '无形资产', 'AP': '应付账款',
                    'Taxes payable': '应交税费', 'OP': '其他应付款', 'Capital': '股本',
                    'Reserve': '资本公积'
                }
                category_mapping = {
                    'Current Assets': ['Cash', 'AR', 'Prepayments', 'OR', 'Other CA'],
                    'Non-current Assets': ['IP', 'Other NCA'],
                    'Liabilities': ['AP', 'Taxes payable', 'OP'],
                    'Equity': ['Capital', 'Reserve']
                }

        # Generate JSON content from session storage (for AI2 easy access)
        json_content = {
            'metadata': {
                'entity_name': entity_name,
                'generated_timestamp': datetime.now().isoformat(),
                'session_id': getattr(st.session_state.get('ai_logger'), 'session_id', 'default'),
                'total_keys': len(content_store)
            },
            'categories': {},
            'keys': {}
        }

        st.info(f"📊 Generating content files from session storage for {len(content_store)} keys")

        # Filter content store based on current statement type
        if current_statement_type == "IS":
            # For IS, only process IS-related keys
            is_keys = ['OI', 'OC', 'Tax and Surcharges', 'GA', 'Fin Exp', 'Cr Loss', 'Other Income', 'Non-operating Income', 'Non-operating Exp', 'Income tax', 'LT DTA']
            filtered_content_store = {k: v for k, v in content_store.items() if k in is_keys}
        elif current_statement_type == "BS":
            # For BS, only process BS-related keys
            bs_keys = ['Cash', 'AR', 'Prepayments', 'OR', 'Other CA', 'Other NCA', 'IP', 'NCA', 'AP', 'Taxes payable', 'OP', 'Capital', 'Reserve']
            filtered_content_store = {k: v for k, v in content_store.items() if k in bs_keys}
        else:
            # For ALL, use all content
            filtered_content_store = content_store

        # Process content by category
        for category, items in category_mapping.items():
            json_content['categories'][category] = []

            for item in items:
                full_name = name_mapping[item]

                # Get latest content from session storage (could be Agent 1, 2, or 3 version)
                if item in filtered_content_store:
                    key_data = filtered_content_store[item]
                    latest_content = key_data.get('current_content', key_data.get('agent1_content', ''))

                    # Determine content source
                    if 'agent3_content' in key_data:
                        content_source = "agent3_final"
                        source_timestamp = key_data.get('agent3_timestamp')
                    elif 'agent2_content' in key_data:
                        content_source = "agent2_validated"
                        source_timestamp = key_data.get('agent2_timestamp')
                    else:
                        content_source = "agent1_original"
                        source_timestamp = key_data.get('agent1_timestamp')

                    st.write(f"  • {item}: Using {content_source} version")
                else:
                    latest_content = f"No information available for {item}"
                    content_source = "not_available"
                    source_timestamp = None

                # Clean content
                cleaned_content = latest_content.strip() if latest_content else ""

                # Skip items with no information or empty content
                if not cleaned_content or "no information available" in cleaned_content.lower():
                    continue

                key_info = {
                    'key': item,
                    'display_name': full_name,
                    'content': cleaned_content,
                    'content_source': content_source,
                    'source_timestamp': source_timestamp,
                    'length': len(cleaned_content),
                    'category': category
                }

                json_content['categories'][category].append(key_info)
                json_content['keys'][item] = key_info

        # Save JSON format (for AI2 easy access)
        if current_statement_type == "IS":
            json_file_path = 'fdd_utils/is_content.json'
            md_file_path = 'fdd_utils/is_content.md'
        elif current_statement_type == "BS":
            json_file_path = 'fdd_utils/bs_content.json'
            md_file_path = 'fdd_utils/bs_content.md'
        else:  # ALL - create comprehensive file
            json_file_path = 'fdd_utils/all_content.json'
            md_file_path = 'fdd_utils/all_content.md'

        with open(json_file_path, 'w', encoding='utf-8') as file:
            json.dump(json_content, file, indent=2, ensure_ascii=False)

        # Also generate markdown for PowerPoint compatibility
        markdown_lines = []
        for category, items in category_mapping.items():
            markdown_lines.append(f"## {category}\n")
            for item in items:
                full_name = name_mapping[item]
                key_info = json_content['keys'].get(item)
                if key_info:
                    cleaned_content = key_info['content']
                else:
                    cleaned_content = f"No information available for {item}"
                markdown_lines.append(f"### {full_name}\n{cleaned_content}\n")

        markdown_text = "\n".join(markdown_lines)

        # For ALL statement type, create comprehensive content file
        if current_statement_type == "ALL":
            # Also create BS and IS specific files for PowerPoint compatibility
            bs_keys = ['Cash', 'AR', 'Prepayments', 'OR', 'Other CA', 'Other NCA', 'IP', 'NCA', 'AP', 'Taxes payable', 'OP', 'Capital', 'Reserve']
            is_keys = ['OI', 'OC', 'Tax and Surcharges', 'GA', 'Fin Exp', 'Cr Loss', 'Other Income', 'Non-operating Income', 'Non-operating Exp', 'Income tax', 'LT DTA']

            # Create BS content file
            bs_content_store = {k: v for k, v in content_store.items() if k in bs_keys}
            if bs_content_store:
                bs_json_content = {
                    'metadata': {
                        'generated_at': datetime.now().strftime('%Y-%m-%d'),
                        'format_version': '1.0',
                        'description': 'Balance Sheet content data'
                    },
                    'categories': {},
                    'keys': {}
                }

                # Process BS content
                bs_category_mapping = category_mapping.copy()
                # Filter to only include BS categories
                bs_category_mapping = {k: v for k, v in bs_category_mapping.items()
                                     if any(item in bs_keys for item in v)}

                for category, items in bs_category_mapping.items():
                    bs_json_content['categories'][category] = []
                    for item in items:
                        if item in bs_content_store:
                            full_name = name_mapping[item]
                            key_data = bs_content_store[item]
                            latest_content = key_data.get('current_content', key_data.get('agent1_content', ''))

                            if latest_content and "no information available" not in latest_content.lower():
                                cleaned_content = latest_content.strip()
                                key_info = {
                                    'key': item,
                                    'display_name': full_name,
                                    'content': cleaned_content,
                                    'content_source': 'agent1_original',
                                    'source_timestamp': key_data.get('agent1_timestamp'),
                                    'length': len(cleaned_content),
                                    'category': category
                                }

                                bs_json_content['categories'][category].append(key_info)
                                bs_json_content['keys'][item] = key_info

                # Save BS content files
                with open('fdd_utils/bs_content.json', 'w', encoding='utf-8') as file:
                    json.dump(bs_json_content, file, indent=2, ensure_ascii=False)

                # Generate BS markdown
                bs_markdown_lines = []
                for category, items in bs_category_mapping.items():
                    bs_markdown_lines.append(f"## {category}\n")
                    for item in items:
                        if item in bs_content_store:
                            full_name = name_mapping[item]
                            key_info = bs_json_content['keys'].get(item)
                            if key_info and key_info['content'] and "no information available" not in key_info['content'].lower():
                                cleaned_content = key_info['content']
                            else:
                                cleaned_content = f"No information available for {item}"
                            bs_markdown_lines.append(f"### {full_name}\n{cleaned_content}\n")

                bs_markdown_text = "\n".join(bs_markdown_lines)
                with open('fdd_utils/bs_content.md', 'w', encoding='utf-8') as file:
                    file.write(bs_markdown_text)

        with open(md_file_path, 'w', encoding='utf-8') as file:
            file.write(markdown_text)

        st.success(f"✅ Content files generated successfully!")

    except Exception as e:
        st.error(f"❌ Error generating content files: {e}")
        import traceback
        st.error(traceback.format_exc())

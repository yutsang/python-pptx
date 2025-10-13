#!/usr/bin/env python3
"""
Data utility functions for FDD application.
Moved from fdd_app.py for better organization.
"""

import re
import json
from datetime import datetime
from pathlib import Path


def get_tab_name(project_name):
    """Get tab name based on project name."""
    if not project_name:
        return None

    project_name = project_name.strip()

    # Hardcoded mappings for known entities
    if project_name.lower() == 'haining':
        return "BSHN"
    elif project_name.lower() == 'nanjing':
        return "BSNJ"
    elif project_name.lower() == 'ningbo':
        return "BSNB"

    # For other entities, try to extract a meaningful sheet name
    # Remove common suffixes and use first word
    clean_name = project_name.split()[0] if project_name else None
    if clean_name:
        # Try different sheet name patterns
        possible_names = [
            f"BS{clean_name.upper()[:3]}",  # BSCLE, BSHAI, etc.
            f"{clean_name.upper()[:3]}",     # CLE, HAI, etc.
            clean_name.upper(),             # CLEANTECH, HAINING, etc.
            f"BS_{clean_name.upper()[:3]}",  # BS_CLE, etc.
            project_name                     # Original name as last resort
        ]
        return possible_names  # Return list of possible names

    # Fallback: return the project name itself to avoid None
    print(f"Warning: Could not extract sheet name from '{project_name}', using project name as fallback")
    return [project_name]


def get_financial_keys():
    """Get list of financial keys."""
    return [
        'Cash', 'AR', 'Prepayments', 'OR', 'IP', 'NCA', 'Other NCA', 'Other CA',
        'AP', 'Advances', 'Taxes payable', 'OP', 'Capital', 'Reserve', 'Capital reserve',
        'OI', 'OC', 'Tax and Surcharges', 'GA', 'Fin Exp', 'Cr Loss', 'Other Income',
        'Non-operating Income', 'Non-operating Exp', 'Income tax', 'LT DTA'
    ]


def detect_chinese_content(content):
    """Detect if content contains significant Chinese characters."""
    if not content:
        return False

    chinese_chars = sum(1 for char in str(content) if '\u4e00' <= char <= '\u9fff')
    total_chars = len(str(content).replace(' ', '').replace('\n', ''))

    if total_chars == 0:
        return False

    chinese_ratio = chinese_chars / total_chars
    return chinese_ratio > 0.3

def get_key_display_name(key, use_excel_tab_name=False, content=None):
    """Get display name for a financial key with Chinese content detection."""
    # Auto-detect if we should use Excel tab names based on content
    if content is not None and detect_chinese_content(content):
        use_excel_tab_name = True

    display_mapping = {
        'Cash': 'Cash and Cash Equivalents',
        'AR': 'Accounts Receivable',
        'Prepayments': 'Prepayments',
        'OR': 'Other Receivables',
        'IP': 'Investment Properties',
        'NCA': 'Non-Current Assets',
        'Other NCA': 'Other Non-Current Assets',
        'Other CA': 'Other Current Assets',
        'AP': 'Accounts Payable',
        'Advances': 'Advances from Customers',
        'Taxes payable': 'Taxes Payable',
        'OP': 'Other Payables',
        'Capital': 'Share Capital',
        'Reserve': 'Reserves',
        'Capital reserve': 'Capital Reserve',
        'OI': 'Operating Income',
        'OC': 'Operating Cost',
        'Tax and Surcharges': 'Tax and Surcharges',
        'GA': 'General and Administrative',
        'Fin Exp': 'Finance Expenses',
        'Cr Loss': 'Credit Loss',
        'Other Income': 'Other Income',
        'Non-operating Income': 'Non-operating Income',
        'Non-operating Exp': 'Non-operating Expenses',
        'Income tax': 'Income Tax',
        'LT DTA': 'Long-term Deferred Tax Assets'
    }

    if use_excel_tab_name:
        # For Excel tab names, use a simpler mapping
        excel_tab_mapping = {
            'Cash': 'Cash',
            'AR': 'AR',
            'Prepayments': 'Prepayments',
            'OR': 'OR',
            'IP': 'IP',
            'NCA': 'NCA',
            'Other NCA': 'Other NCA',
            'Other CA': 'Other CA',
            'AP': 'AP',
            'Advances': 'Advances',
            'Taxes payable': 'Taxes payable',
            'OP': 'OP',
            'Capital': 'Capital',
            'Reserve': 'Reserve',
            'Capital reserve': 'Capital reserve',
            'OI': 'OI',
            'OC': 'OC',
            'Tax and Surcharges': 'Tax and Surcharges',
            'GA': 'GA',
            'Fin Exp': 'Fin Exp',
            'Cr Loss': 'Cr Loss',
            'Other Income': 'Other Income',
            'Non-operating Income': 'Non-operating Income',
            'Non-operating Exp': 'Non-operating Exp',
            'Income tax': 'Income tax',
            'LT DTA': 'LT DTA'
        }
        return excel_tab_mapping.get(key, key)

    # Apply proper case formatting for display names
    display_name = display_mapping.get(key, key)
    return display_name.title()


def format_date_to_dd_mmm_yyyy(date_str):
    """Format date string to DD-MMM-YYYY format."""
    if not date_str:
        return 'Unknown'
    
    # Handle datetime objects directly
    if hasattr(date_str, 'strftime'):
        try:
            return date_str.strftime('%d-%b-%Y')
        except:
            pass
    
    date_str = str(date_str).strip()
    
    # Common date patterns
    patterns = [
        ('%Y-%m-%d', '%d-%b-%Y'),  # 2022-09-30 -> 30-Sep-2022
        ('%d/%m/%Y', '%d-%b-%Y'),  # 30/09/2022 -> 30-Sep-2022
        ('%m/%d/%Y', '%d-%b-%Y'),  # 09/30/2022 -> 30-Sep-2022
        ('%d-%m-%Y', '%d-%b-%Y'),  # 30-09-2022 -> 30-Sep-2022
        ('%Y/%m/%d', '%d-%b-%Y'),  # 2022/09/30 -> 30-Sep-2022
        # Chinese date formats
        ('%Y年%m月%d日', '%d-%b-%Y'),  # 2024年5月31日 -> 31-May-2024
        ('%Y年%m月', '%b-%Y'),       # 2024年5月 -> May-2024
        ('%m月%d日', '%d-%b'),       # 5月31日 -> 31-May
        ('%Y.%m.%d', '%d-%b-%Y'),    # 2024.5.31 -> 31-May-2024
    ]
    
    for input_format, output_format in patterns:
        try:
            dt = datetime.strptime(date_str, input_format)
            return dt.strftime(output_format)
        except ValueError:
            continue
    
    # If no pattern matches, return as-is
    return date_str


def extract_entity_names_from_databook(databook_path='databook.xlsx', exclude_patterns=None):
    """Extract entity names from databook Excel headers using pattern ' - EntityName'.
    
    Args:
        databook_path: Path to Excel file
        exclude_patterns: List of patterns to exclude (e.g., ['年', '2024', 'Indicative'])
    """
    import pandas as pd
    import os
    import re
    
    # Default exclusion patterns
    if exclude_patterns is None:
        exclude_patterns = [
            r'\d{4}年',          # Pattern like "2024年"
            r'\d{4}-\d{2}-\d{2}', # Dates like "2024-12-31"
            r'Indicative',       # Indicative adjusted columns
            r'示意性',           # Chinese "Indicative"
            r'^$',               # Empty strings
            r'^\s*$',            # Whitespace only
        ]
    
    entity_names = set()
    
    try:
        # Check if databook exists (in parent directory if path is relative)
        if not os.path.isabs(databook_path):
            base_path = Path(__file__).parent.parent
            full_path = base_path / databook_path
        else:
            full_path = Path(databook_path)
            
        if not full_path.exists():
            print(f"Databook not found at {full_path}")
            return []
        
        # Scan all sheets for entity name patterns
        xl = pd.ExcelFile(full_path)
        for sheet_name in xl.sheet_names:
            try:
                df = xl.parse(sheet_name)
                for i in range(len(df)):
                    for j in range(len(df.columns)):
                        val = df.iloc[i, j]
                        if pd.notna(val) and isinstance(val, str):
                            val_str = str(val)
                            # Look for pattern ' - EntityName'
                            if ' - ' in val_str:
                                parts = val_str.split(' - ')
                                if len(parts) >= 2:
                                    entity_name = parts[-1].strip()
                                    
                                    # Filter out unwanted patterns
                                    should_exclude = False
                                    for pattern in exclude_patterns:
                                        if re.search(pattern, entity_name):
                                            should_exclude = True  # CRITICAL: Exclude if pattern matches
                                            break
                                    
                                    # Filter reasonable entity names
                                    if not should_exclude and len(entity_name) > 2 and len(entity_name) < 50:
                                        entity_names.add(entity_name)
            except Exception as e:
                print(f"Error processing sheet {sheet_name}: {e}")
                continue
        
        return sorted(list(entity_names))
        
    except Exception as e:
        print(f"Error extracting entity names from databook: {e}")
        return []


def load_config_files():
    """Load configuration files."""
    base_path = Path(__file__).parent
    
    try:
        # Load config
        config_path = base_path / 'config.json'
        if config_path.exists():
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
        else:
            config = {}

        # Load mapping
        mapping_path = base_path / 'mapping.json'
        if mapping_path.exists():
            with open(mapping_path, 'r', encoding='utf-8') as f:
                mapping = json.load(f)
        else:
            mapping = {}

        # Load pattern
        pattern_path = base_path / 'pattern.json'
        if pattern_path.exists():
            with open(pattern_path, 'r', encoding='utf-8') as f:
                pattern = json.load(f)
        else:
            pattern = {}

        # Load prompts
        prompts_path = base_path / 'prompts.json'
        if prompts_path.exists():
            with open(prompts_path, 'r', encoding='utf-8') as f:
                prompts = json.load(f)
        else:
            prompts = {}
        
        return config, mapping, pattern, prompts
        
    except UnicodeDecodeError as e:
        print(f"Encoding error loading config files (try saving files as UTF-8): {e}")
        return {}, {}, {}, {}
    except Exception as e:
        print(f"Error loading config files: {e}")
        return {}, {}, {}, {}

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
    if project_name == 'Haining':
        return "BSHN"
    elif project_name == 'Nanjing':
        return "BSNJ"
    elif project_name == 'Ningbo':
        return "BSNB"
    return project_name


def get_financial_keys():
    """Get list of financial keys."""
    return [
        'Cash', 'AR', 'Prepayments', 'OR', 'IP', 'NCA', 'Other NCA', 'Other CA',
        'AP', 'Advances', 'Taxes payable', 'OP', 'Capital', 'Reserve', 'Capital reserve',
        'OI', 'OC', 'Tax and Surcharges', 'GA', 'Fin Exp', 'Cr Loss', 'Other Income',
        'Non-operating Income', 'Non-operating Exp', 'Income tax', 'LT DTA'
    ]


def get_key_display_name(key):
    """Get display name for a financial key."""
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
        'OI': 'Other Income',
        'OC': 'Other Comprehensive Income',
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
    return display_mapping.get(key, key)


def format_date_to_dd_mmm_yyyy(date_str):
    """Format date string to DD-MMM-YYYY format."""
    if not date_str:
        return 'Unknown'
    
    date_str = str(date_str).strip()
    
    # Common date patterns
    patterns = [
        ('%Y-%m-%d', '%d-%b-%Y'),  # 2022-09-30 -> 30-Sep-2022
        ('%d/%m/%Y', '%d-%b-%Y'),  # 30/09/2022 -> 30-Sep-2022
        ('%m/%d/%Y', '%d-%b-%Y'),  # 09/30/2022 -> 30-Sep-2022
        ('%d-%m-%Y', '%d-%b-%Y'),  # 30-09-2022 -> 30-Sep-2022
        ('%Y/%m/%d', '%d-%b-%Y'),  # 2022/09/30 -> 30-Sep-2022
    ]
    
    for input_format, output_format in patterns:
        try:
            dt = datetime.strptime(date_str, input_format)
            return dt.strftime(output_format)
        except ValueError:
            continue
    
    # If no pattern matches, return as-is
    return date_str


def load_config_files():
    """Load configuration files."""
    base_path = Path(__file__).parent
    
    try:
        # Load config
        config_path = base_path / 'config.json'
        if config_path.exists():
            with open(config_path, 'r') as f:
                config = json.load(f)
        else:
            config = {}
        
        # Load mapping
        mapping_path = base_path / 'mapping.json'
        if mapping_path.exists():
            with open(mapping_path, 'r') as f:
                mapping = json.load(f)
        else:
            mapping = {}
        
        # Load pattern
        pattern_path = base_path / 'pattern.json'
        if pattern_path.exists():
            with open(pattern_path, 'r') as f:
                pattern = json.load(f)
        else:
            pattern = {}
        
        # Load prompts
        prompts_path = base_path / 'prompts.json'
        if prompts_path.exists():
            with open(prompts_path, 'r') as f:
                prompts = json.load(f)
        else:
            prompts = {}
        
        return config, mapping, pattern, prompts
        
    except Exception as e:
        print(f"Error loading config files: {e}")
        return {}, {}, {}, {}

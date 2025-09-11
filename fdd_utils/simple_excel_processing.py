#!/usr/bin/env python3
"""
Simplified Excel processing for FDD application.
This version focuses on extracting financial data from any sheet that contains it.
"""

import pandas as pd
import re
from datetime import datetime


def simple_process_excel(uploaded_file, tab_name_mapping, entity_name, entity_keywords=None):
    """
    Simplified Excel processing that looks for financial data in any sheet.
    
    Args:
        uploaded_file: The uploaded Excel file
        tab_name_mapping: Mapping of financial keys to their patterns
        entity_name: Target entity name
        entity_keywords: List of entity keywords
    
    Returns:
        dict: sections_by_key containing extracted financial data
    """
    sections_by_key = {}
    
    try:
        # Read Excel file
        xl = pd.ExcelFile(uploaded_file)
        print(f"📊 Found {len(xl.sheet_names)} sheets in Excel file")
        print(f"📋 Sheet names: {xl.sheet_names}")
        
        # Skip obvious non-financial sheets
        skip_patterns = [
            'cover', 'overview', 'summary', 'snapshot', 'choice', 'check', 
            'violations', 'navi', '历史沿革', '工程台账', '关联方', 'rent roll', 
            'property', 'briefing'
        ]
        
        for sheet_name in xl.sheet_names:
            print(f"\n🔍 Processing sheet: '{sheet_name}'")
            
            # Skip non-financial sheets
            if any(skip_word.lower() in sheet_name.lower() for skip_word in skip_patterns):
                print(f"⏭️ Skipping non-financial sheet: {sheet_name}")
                continue
            
            try:
                # Load the sheet
                df = xl.parse(sheet_name)
                
                if df.empty or len(df) == 0:
                    print(f"⚠️ Sheet '{sheet_name}' is empty, skipping...")
                    continue
                
                print(f"📊 Sheet '{sheet_name}' has {len(df)} rows and {len(df.columns)} columns")
                
                # Check if this sheet contains financial data
                has_financial_data = check_for_financial_indicators(df)
                
                if not has_financial_data:
                    print(f"⏭️ No financial indicators found in sheet '{sheet_name}', skipping...")
                    continue
                
                print(f"✅ Found financial data in sheet '{sheet_name}'")
                
                # Extract financial keys from the content
                financial_keys = extract_financial_keys_from_content(df, tab_name_mapping)
                
                if not financial_keys:
                    # Use a generic key based on sheet name
                    generic_key = f"Sheet_{sheet_name.replace(' ', '_').replace('->', '_')}"
                    financial_keys = [generic_key]
                    print(f"💡 Using generic key: {generic_key}")
                
                # Process the data for each financial key
                for key in financial_keys:
                    if key not in sections_by_key:
                        sections_by_key[key] = []
                    
                    # Create a simple parsed data structure
                    parsed_data = create_simple_parsed_data(df, sheet_name, entity_name)
                    
                    if parsed_data:
                        sections_by_key[key].append({
                            'sheet_name': sheet_name,
                            'entity_name': entity_name,
                            'parsed_data': parsed_data,
                            'raw_data': df
                        })
                        print(f"✅ Added data to key '{key}' from sheet '{sheet_name}'")
                
            except Exception as e:
                print(f"❌ Error processing sheet '{sheet_name}': {e}")
                continue
        
        # Summary
        keys_with_data = [key for key, sections in sections_by_key.items() if sections]
        print(f"\n🎉 Processing complete!")
        print(f"📊 Found {len(keys_with_data)} financial keys with data: {keys_with_data}")
        
        return sections_by_key
        
    except Exception as e:
        print(f"❌ Error processing Excel file: {e}")
        return {}


def check_for_financial_indicators(df):
    """Check if a DataFrame contains financial data indicators."""
    try:
        # Look for financial indicators in first 10 rows
        for row_idx in range(min(10, len(df))):
            for col_idx in range(min(10, len(df.columns))):
                try:
                    cell_value = str(df.iloc[row_idx, col_idx]).lower()
                    
                    # Check for financial indicators
                    financial_indicators = [
                        'indicative adjusted', '示意性调整', '示意性調整',
                        'adjusted', '人民币千元', '人民幣千元', 'cny\'000',
                        'rmb', 'financial', 'balance sheet', 'income statement',
                        'cash flow', 'assets', 'liabilities', 'equity'
                    ]
                    
                    if any(indicator in cell_value for indicator in financial_indicators):
                        return True
                        
                except:
                    continue
        
        return False
        
    except:
        return False


def extract_financial_keys_from_content(df, tab_name_mapping):
    """Extract financial keys by looking at the actual content."""
    found_keys = []
    
    try:
        # Convert DataFrame content to searchable text
        all_text = ""
        for row_idx in range(min(20, len(df))):
            for col_idx in range(len(df.columns)):
                try:
                    cell_value = str(df.iloc[row_idx, col_idx])
                    if cell_value and cell_value != 'nan':
                        all_text += cell_value.lower() + " "
                except:
                    continue
        
        # Look for financial terms in the content
        for financial_key, patterns in tab_name_mapping.items():
            for pattern in patterns:
                if pattern.lower() in all_text:
                    if financial_key not in found_keys:
                        found_keys.append(financial_key)
                        print(f"🔍 Found '{pattern}' -> key '{financial_key}'")
                    break
        
        return found_keys
        
    except Exception as e:
        print(f"❌ Error extracting financial keys: {e}")
        return []


def create_simple_parsed_data(df, sheet_name, entity_name):
    """Create a simple parsed data structure from the DataFrame."""
    try:
        # Find the data area (skip empty rows at the top)
        data_start_row = 0
        for row_idx in range(len(df)):
            row_values = df.iloc[row_idx]
            if row_values.notna().sum() > 2:  # Row has at least 3 non-empty cells
                data_start_row = row_idx
                break
        
        # Use the DataFrame from the data start
        data_df = df.iloc[data_start_row:].copy()
        
        if data_df.empty:
            return None
        
        # Create metadata
        metadata = {
            'table_name': sheet_name,
            'entity_name': entity_name,
            'date': datetime.now().strftime('%Y-%m-%d'),
            'currency': 'CNY',
            'unit': 'thousands'
        }
        
        # Convert DataFrame to list of dictionaries for data rows
        data_rows = []
        for _, row in data_df.iterrows():
            row_dict = {}
            for col_name, value in row.items():
                if pd.notna(value):
                    row_dict[str(col_name)] = str(value)
            if row_dict:  # Only add non-empty rows
                data_rows.append(row_dict)
        
        return {
            'metadata': metadata,
            'data': data_rows
        }
        
    except Exception as e:
        print(f"❌ Error creating parsed data: {e}")
        return None
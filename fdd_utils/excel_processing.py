#!/usr/bin/env python3
"""
Excel processing utilities for FDD application.
Moved from fdd_app.py for better organization.
"""

import pandas as pd
import re
import sys
from datetime import datetime
from pathlib import Path
import streamlit as st
from tabulate import tabulate



def detect_latest_date_column(df, sheet_name="Sheet", entity_keywords=None):
    """Detect the latest date column from a DataFrame, focusing on 'Indicative adjusted' (English/Chinese) with merged cell handling."""

    def col_num_to_letter(n):
        """Convert column number to Excel-style letter (0=A, 1=B, ..., 25=Z, 26=AA, etc.)"""
        if n < 0:
            return "?"
        result = ""
        n = n + 1  # Convert to 1-based
        while n > 0:
            n = n - 1
            result = chr(65 + (n % 26)) + result
            n = n // 26
        return result

    from datetime import datetime
    timestamp = datetime.now().strftime('%H:%M:%S.%f')[:-3]

    def parse_date(date_str):
        """Parse date string in various formats including xMxx."""
        if not date_str or pd.isna(date_str):
            return None
        
        date_str = str(date_str).strip()
        
        # AVOID CONFUSING NUMBERS WITH DATES
        # Skip values that look like large numbers with commas (e.g., "2,021,000")
        # These should NOT be interpreted as dates
        if ',' in date_str and date_str.replace(',', '').replace('.', '').isdigit():
            num_val = float(date_str.replace(',', ''))
            if num_val > 10000:  # Large numbers are definitely not dates
                return None
        
        # Handle Chinese date range format: 2024年1-5月 (use the END month)
        chinese_range_match = re.match(r'^(\d{4})年(\d{1,2})-(\d{1,2})月$', date_str)
        if chinese_range_match:
            year = int(chinese_range_match.group(1))
            end_month = int(chinese_range_match.group(3))  # Use end month
            # Use last day of the end month
            if end_month == 12:
                return datetime(year, 12, 31)
            elif end_month in [1, 3, 5, 7, 8, 10]:
                return datetime(year, end_month, 31)
            elif end_month in [4, 6, 9, 11]:
                return datetime(year, end_month, 30)
            elif end_month == 2:
                if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0):
                    return datetime(year, 2, 29)
                else:
                    return datetime(year, 2, 28)
        
        # Handle xMxx format (e.g., 9M22, 12M23) - END OF MONTH
        xmxx_match = re.match(r'^(\d+)M(\d{2})$', date_str)
        if xmxx_match:
            month = int(xmxx_match.group(1))
            year = 2000 + int(xmxx_match.group(2))  # Assume 20xx for 2-digit years
            # Use end of month, not beginning (last day of the month)
            if month == 12:
                return datetime(year, 12, 31)  # December 31st
            elif month in [1, 3, 5, 7, 8, 10]:
                return datetime(year, month, 31)  # 31-day months
            elif month in [4, 6, 9, 11]:
                return datetime(year, month, 30)  # 30-day months
            elif month == 2:
                # February - handle leap years
                if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0):
                    return datetime(year, 2, 29)  # Leap year
                else:
                    return datetime(year, 2, 28)  # Non-leap year
        
        # Handle standard date formats
        date_formats = [
            '%d/%m/%Y', '%d-%m-%Y', '%d/%m/%y', '%d-%m-%y',
            '%Y-%m-%d', '%m/%d/%Y', '%m-%d-%Y',
            '%d/%b/%Y', '%d-%b-%Y', '%b/%d/%Y', '%b-%d-%Y',
            '%d/%B/%Y', '%d-%B-%Y', '%B/%d/%Y', '%B-%d-%Y',
            # Chinese date formats
            '%Y年%m月%d日', '%Y年%m月', '%m月%d日', '%Y/%m/%d',
            '%Y.%m.%d', '%Y年%m月%d日', '%Y年%m月%d号',
            # Additional flexible formats
            '%Y%m%d', '%d%m%Y', '%m%d%Y'
        ]
        
        for fmt in date_formats:
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                continue
        
        return None
    
    # Get column names
    columns = df.columns.tolist()
    latest_date = None
    latest_column = None

    # Strategy: Find rightmost column that has data (not empty)
    # If "Indicative adjusted" is found, look for data columns to the right of it
    # Otherwise, look for rightmost column with actual data values
    
    # print(f"🔍 {sheet_name}: Searching for 'Indicative adjusted' (English/Chinese) column...")

    # Step 1: Find "Indicative adjusted" (English/Chinese) positions
    indicative_positions = []
    
    # Search in first 10 rows for "Indicative adjusted" (English and Chinese)
    for row_idx in range(min(10, len(df))):
        for col_idx, col in enumerate(columns):
            val = df.iloc[row_idx, col_idx]
            val_str = str(val).lower()
            # Check for English "indicative adjusted" or Chinese "示意性調整後" / "示意性调整后"
            if pd.notna(val) and (
                ('indicative' in val_str and 'adjusted' in val_str) or
                '示意性調整後' in val_str or
                '示意性调整后' in val_str
            ):
                indicative_positions.append((row_idx, col_idx))
                print(f"   📋 FOUND 'Indicative adjusted' (English/Chinese) at Row {row_idx}, Col {col_idx} ({col}): '{df.iloc[row_idx, col_idx]}'")
                print(f"   📋 Column letter: {col_num_to_letter(col_idx)} (actual position in Excel)")

    if not indicative_positions:
        print(f"   ⚠️  No 'Indicative adjusted' (English/Chinese) found, using fallback date detection")
        # Fallback: find any date column
    else:
        # Compact output
        # Step 2: For each "Indicative adjusted" (English/Chinese) position, find the merged range and get the date
        for instance_idx, (indic_row, indic_col) in enumerate(indicative_positions):
            col_name = columns[indic_col]
            col_letter = col_num_to_letter(indic_col)
            # print(f"   🔍 [{instance_idx+1}/{len(indicative_positions)}] Processing 'Indicative adjusted' at Row {indic_row}, Col {indic_col} ({col_name}) - {col_letter}")

            # Find merged range: go right until we hit a non-empty cell or reach the end
            merge_start = indic_col
            merge_end = indic_col

            # Check if this is a merged cell by looking right
            for check_col in range(indic_col + 1, len(columns)):
                val = df.iloc[indic_row, check_col]
                if pd.isna(val) or str(val).strip() == '':
                    merge_end = check_col
                else:
                    break
            else:
                merge_end = len(columns) - 1

            # print(f"   📍 Merged range: columns {merge_start}-{merge_end} (searching for latest date)")
            # print(f"   📍 Indicative adjusted found at column {indic_col} (0-indexed)")

            # Find the date value in the row below the "Indicative adjusted" header
            date_row = indic_row + 1
            if date_row < len(df):
                # Look for date in the merged range
                for col_idx in range(merge_start, merge_end + 1):
                    val = df.iloc[date_row, col_idx]

                    if isinstance(val, (pd.Timestamp, datetime)):
                        date_val = val if isinstance(val, datetime) else val.to_pydatetime()
                        if latest_date is None or date_val > latest_date:
                            old_date = latest_date
                            latest_date = date_val
                            latest_column = columns[col_idx]
                            # print(f"   📅 [{instance_idx+1}] FOUND DATE in merged range: {latest_column} = {date_val.strftime('%Y-%m-%d')} (previous: {old_date.strftime('%Y-%m-%d') if old_date else 'None'})")
                    elif pd.notna(val):
                        parsed_date = parse_date(str(val))
                        if parsed_date:
                            if latest_date is None or parsed_date > latest_date:
                                old_date = latest_date
                                latest_date = parsed_date
                                latest_column = columns[col_idx]
                                # print(f"   📅 [{instance_idx+1}] PARSED DATE in merged range: {latest_column} = {parsed_date.strftime('%Y-%m-%d')} (previous: {old_date.strftime('%Y-%m-%d') if old_date else 'None'})")

            # If no date found in the row below, check a few more rows down
            if latest_column is None:
                for check_row in range(indic_row + 2, min(indic_row + 8, len(df))):  # Extended search range
                    for col_idx in range(merge_start, merge_end + 1):
                        val = df.iloc[check_row, col_idx]

                        if isinstance(val, (pd.Timestamp, datetime)):
                            date_val = val if isinstance(val, datetime) else val.to_pydatetime()
                            if latest_date is None or date_val > latest_date:
                                latest_date = date_val
                                latest_column = columns[col_idx]
                                print(f"   📅 Found date in row {check_row}: {latest_column} = {date_val.strftime('%Y-%m-%d')}")
                        elif pd.notna(val):
                            parsed_date = parse_date(str(val))
                            if parsed_date:
                                if latest_date is None or parsed_date > latest_date:
                                    latest_date = parsed_date
                                    latest_column = columns[col_idx]
                                    print(f"   📅 Parsed date in row {check_row}: {latest_column} = {parsed_date.strftime('%Y-%m-%d')}")

            # If still no date found in merged range, search entire sheet for dates (Chinese files might have dates elsewhere)
            if latest_column is None:
                # Extended search
                for row_idx in range(min(15, len(df))):  # Search first 15 rows
                    for col_idx in range(len(columns)):
                        val = df.iloc[row_idx, col_idx]

                        if isinstance(val, (pd.Timestamp, datetime)):
                            date_val = val if isinstance(val, datetime) else val.to_pydatetime()
                            if latest_date is None or date_val > latest_date:
                                latest_date = date_val
                                latest_column = columns[col_idx]
                                print(f"   📅 Found date in extended search: Row {row_idx}, Col {col_idx} = {date_val.strftime('%Y-%m-%d')}")
                        elif pd.notna(val) and str(val).strip():
                            parsed_date = parse_date(str(val))
                            if parsed_date:
                                if latest_date is None or parsed_date > latest_date:
                                    latest_date = parsed_date
                                    latest_column = columns[col_idx]
                                    print(f"   📅 Parsed date in extended search: Row {row_idx}, Col {col_idx} = {parsed_date.strftime('%Y-%m-%d')}")
    
    if latest_column:
        from datetime import datetime
        timestamp = datetime.now().strftime('%H:%M:%S.%f')[:-3]
        col_idx = columns.index(latest_column) if latest_column in columns else -1

        # If no date was found but we have indicative positions, find rightmost column with data
        if latest_column is None and indicative_positions:
            print(f"   🔍 No date found in indicative ranges, looking for rightmost data column...")

            # Strategy: Find the rightmost column that has actual data (not headers/empty)
            rightmost_data_col = None
            rightmost_col_idx = -1

            # Check columns from right to left, starting from the rightmost
            for col_idx_check in range(len(columns) - 1, -1, -1):
                col_name = columns[col_idx_check]

                # Skip obvious header columns
                if any(skip in str(col_name).lower() for skip in ['total', '合计', 'sum', 'summary']):
                    continue

                # Check if this column has actual data (numbers, not just headers)
                has_data = False
                for row_idx in range(min(50, len(df))):  # Check first 50 rows
                    val = df.iloc[row_idx, col_idx_check]
                    if pd.notna(val) and str(val).strip():
                        # Check if it's a number or meaningful text (not header)
                        val_str = str(val).strip()
                        if (val_str.replace(',', '').replace('.', '').replace('-', '').isdigit() and
                            len(val_str) > 2):  # At least 3 digits for meaningful numbers
                            has_data = True
                            break
                        elif not any(header_word in val_str.lower() for header_word in
                                   ['total', '合计', 'sum', 'subtotal', '小计', 'header', 'title']):
                            # Check if it's Chinese date format like "2024年1-5月"
                            if re.match(r'\d{4}年\d{1,2}月|\d{4}年\d{1,2}-\d{1,2}月', val_str):
                                has_data = True
                                break

                if has_data:
                    rightmost_data_col = col_name
                    rightmost_col_idx = col_idx_check
                    print(f"   📊 Found data in column {col_idx_check}: {col_name}")
                    break

            if rightmost_data_col:
                latest_column = rightmost_data_col
                col_idx = rightmost_col_idx
                print(f"   🎯 Selected rightmost data column: {latest_column}")

        # Extract column number from column name (e.g., "Unnamed: 20" -> 20)
        if latest_column and latest_column.startswith('Unnamed: '):
            try:
                actual_col_idx = int(latest_column.split(': ')[1])
                col_letter = col_num_to_letter(actual_col_idx)
            except (ValueError, IndexError):
                col_letter = 'unknown'
        else:
            col_letter = col_num_to_letter(col_idx) if col_idx >= 0 else 'unknown'

        if latest_date:
            print(f"   🎯 [{timestamp}] FINAL SELECTION: Column '{latest_column}' ({col_letter}) with date {latest_date.strftime('%Y-%m-%d')}")
        else:
            print(f"   🎯 [{timestamp}] FINAL SELECTION: Column '{latest_column}' ({col_letter}) - rightmost data column")
        print(f"   ✅ Column {col_letter} (list index {col_idx}) selected as target column")
    else:
        from datetime import datetime
        timestamp = datetime.now().strftime('%H:%M:%S.%f')[:-3]
        print(f"   ❌ [{timestamp}] No column detected")

    return latest_column


def determine_entity_mode_and_filter(df, entity_name, entity_keywords, manual_mode='single'):
    """Determine if we're in single entity or multiple entity mode and filter accordingly.

    Args:
        df: DataFrame to process
        entity_name: Target entity name
        entity_keywords: List of entity keywords to search for
        manual_mode: Manual selection - 'single' or 'multiple'
    """
    # Minimal output

    # If manual mode is 'single', skip entity filtering and return original table
    if manual_mode == 'single':
        return df, False

    # Step 1: FIRST - Scan entire sheet for multiple entities to determine mode
    # Entity detection
    entities_found_in_sheet = []
    entity_patterns = [kw.lower() for kw in entity_keywords if kw.strip()]
    if entity_name and entity_name.lower() not in entity_patterns:
        entity_patterns.append(entity_name.lower())

    for row_idx in range(len(df)):
        row = df.iloc[row_idx]
        row_text = ' '.join(str(val).lower() for val in row if pd.notna(val))

        for pattern in entity_patterns:
            if pattern in row_text:
                # Use the original entity name if it matches the pattern, otherwise use the pattern
                if entity_name and entity_name.lower() == pattern:
                    entity_name_found = entity_name
                else:
                    # Find the original keyword that matches this pattern
                    original_keyword = None
                    for kw in entity_keywords:
                        if kw.lower() == pattern:
                            original_keyword = kw
                            break
                    entity_name_found = original_keyword if original_keyword else pattern.title()
                
                if entity_name_found not in entities_found_in_sheet:
                    entities_found_in_sheet.append(entity_name_found)
                    # Entity found
                break  # Stop checking other patterns for this row

    # Also check with the provided entity keywords
    for row_idx in range(len(df)):
        row = df.iloc[row_idx]
        row_text = ' '.join(str(val).lower() for val in row if pd.notna(val))

        for keyword in entity_keywords:
            if keyword.lower() in row_text and keyword not in entities_found_in_sheet:
                entities_found_in_sheet.append(keyword)
                # Entity found

    total_entities_detected = len(entities_found_in_sheet)
    # Compact entity summary

    # Determine if this is truly a multiple entity file
    is_actually_multiple = total_entities_detected > 1

    # Entity mode validation (reduced output)

    # Step 2: Identify table sections by finding empty rows or major delimiters
    table_sections = []
    current_section_start = None

    for row_idx in range(len(df)):
        # Check if this row is mostly empty (indicates table boundary)
        row_values = df.iloc[row_idx]
        non_empty_count = row_values.notna().sum()
        has_content = non_empty_count > 0

        if has_content and current_section_start is None:
            # Start of a new section
            current_section_start = row_idx
        elif not has_content and current_section_start is not None:
            # End of current section
            table_sections.append((current_section_start, row_idx - 1))
            current_section_start = None
        elif row_idx == len(df) - 1 and current_section_start is not None:
            # Last row, close the current section
            table_sections.append((current_section_start, row_idx))

    print(f"   📊 Found {len(table_sections)} table sections in the worksheet")

    # Step 2: Analyze each section to find entity associations
    # Use both user-specified keywords AND all entities found in the sheet
    all_entity_keywords = list(set(entity_keywords + entities_found_in_sheet))
    print(f"   🔑 ALL ENTITY KEYWORDS FOR SECTION ANALYSIS: {all_entity_keywords}")

    entity_sections = {}

    # Always use the original complex detection logic for all modes
    print(f"   🎯 USING ORIGINAL COMPLEX DETECTION LOGIC")

    # Also try to find any potential entity names in the Excel file
    all_potential_entities = set()

    for section_idx, (start_row, end_row) in enumerate(table_sections):
        # Include both cell values and column names in the section text
        section_values = []
        for val in df.iloc[start_row:end_row+1].values.flatten():
            if pd.notna(val) and str(val).strip():
                section_values.append(str(val))

        # Also include column names that might contain entity information
        for col_name in df.columns:
            if pd.notna(col_name) and str(col_name).strip():
                col_name_str = str(col_name).strip()
                section_values.append(col_name_str)
                # Debug: show when we're checking column names
                if any(kw.lower() in col_name_str.lower() for kw in all_entity_keywords):
                    print(f"   📋 Checking column name: '{col_name_str}'")

        section_text = ' '.join(val.lower() for val in section_values)

        # Check which entities are mentioned in this section - MORE ROBUST MATCHING
        for keyword in all_entity_keywords:
            # Try different matching strategies
            keyword_lower = keyword.lower()
            section_lower = section_text.lower()
            keyword_parts = keyword_lower.replace('-', ' ').split()

            # 1. Exact match
            if keyword_lower in section_lower:
                if keyword not in entity_sections:
                    entity_sections[keyword] = []
                entity_sections[keyword].append((start_row, end_row, section_idx))
                print(f"   ✅ EXACT MATCH: Section {section_idx} contains entity '{keyword}'")
                continue

            # 2. Partial word match (split by spaces and hyphens) - STRICT MATCHING
            matches_found = 0
            total_parts = len(keyword_parts)

            # For multi-word entities, require them to appear together
            if total_parts > 1:
                # Check if the full phrase appears
                if keyword_lower in section_lower:
                    matches_found = total_parts
                else:
                    # Check if all parts appear (but not necessarily together)
                    for part in keyword_parts:
                        if len(part) > 2:
                            # Use word boundaries to avoid substring matches
                            if f' {part} ' in f' {section_lower} ' or section_lower.startswith(f'{part} ') or section_lower.endswith(f' {part}'):
                                matches_found += 1
                    # For multi-word, require 100% match
                    if matches_found < total_parts:
                        matches_found = 0
            else:
                # For single words, use the original logic
                part = keyword_parts[0]
                if len(part) > 2:
                    if f' {part} ' in f' {section_lower} ' or section_lower.startswith(f'{part} ') or section_lower.endswith(f' {part}'):
                        matches_found = 1

            # Require exact match for multi-word entities
            if matches_found >= total_parts:
                if keyword not in entity_sections:
                    entity_sections[keyword] = []
                entity_sections[keyword].append((start_row, end_row, section_idx))
                print(f"   ✅ PARTIAL MATCH: Section {section_idx} contains entity '{keyword}' (matched {matches_found}/{total_parts} parts)")
                continue

        # 3. Special handling for patterns like "Cash and cash equivalents - Ningbo"
        # Extract entity names that appear after common separators
        separators = [' - ', ' – ', ' and ', ' or ', ' for ']
        for separator in separators:
            if separator in section_text:
                parts = section_text.split(separator)
                for part in parts[1:]:  # Look in parts after the separator
                    part = part.strip()
                    if part and any(kw_part.lower() in part.lower() for kw_part in keyword_parts):
                        if keyword not in entity_sections:
                            entity_sections[keyword] = []
                        entity_sections[keyword].append((start_row, end_row, section_idx))
                        print(f"   ✅ SEPARATOR MATCH: Section {section_idx} contains entity '{keyword}' after '{separator.strip()}'")
                        break

    # Try to identify potential entity names in the section - IMPROVED DETECTION
    # Include column names in the word analysis
    section_words = section_text.split()

    # Also add column names as potential words to check
    for col_name in df.columns:
        if pd.notna(col_name):
            col_name_str = str(col_name)
            section_words.extend(col_name_str.split())
    for word in section_words:
        word = word.strip('.,;:!?()[]{}"\'')
        # Look for capitalized words that might be company names
        if (len(word) > 2 and word[0].isupper() and
            not word.lower().endswith(('the', 'and', 'for', 'with', 'from', 'into', 'onto', 'data', 'table', 'sheet', 'file',
                                      'total', 'amount', 'value', 'date', 'cash', 'flow', 'income', 'expense', 'asset', 'liability',
                                      'equity', 'revenue', 'cost', 'profit', 'loss', 'balance', 'sheet', 'statement',
                                      'adjusted', 'indicative', 'audited', 'mgt', 'acc', 'cny', '000', '000', 'equivalent'))):
            # Additional check: avoid common financial terms
            if not any(financial_term in word.lower() for financial_term in
                      ['assets', 'liabilities', 'equity', 'receivable', 'payable', 'deposits', 'banks', 'provision',
                       'bad', 'doubtful', 'properties', 'investment', 'capital', 'share', 'current', 'prepayments']):
                all_potential_entities.add(word)

    # Step 3: Determine entity mode and handle fallback
    # For multiple entity mode, these will be 0 since we skipped entity detection
    unique_entities = len(entity_sections)
    detected_multiple = unique_entities > 1

    # Entity summary: {unique_entities} entities detected

    # For single entity mode only, show troubleshooting if no entities found
    # Troubleshooting reduced

    # Set entity mode based on detection results
    if manual_mode == 'multiple':
        # For multiple entity mode, use the detected multiple status
        is_multiple_entity = detected_multiple
        print(f"   🎯 MULTIPLE ENTITY MODE: {'CONFIRMED' if detected_multiple else 'NOT DETECTED'} ({unique_entities} entities found)")
    else:
        is_multiple_entity = False

    # Step 4: Find the best matching entity section (single entity mode only)
    target_entity = None
    best_section = None

    # For multiple entity mode, skip this and use simplified filtering below
    if manual_mode != 'multiple':
        # Priority 1: Exact match with entity_name
        if entity_name in entity_sections:
            target_entity = entity_name
            best_section = entity_sections[entity_name][0]  # Take first matching section
            print(f"   ✅ EXACT MATCH: Found entity '{entity_name}' in section {best_section[2]} (rows {best_section[0]}-{best_section[1]})")

        # Priority 2: Partial match with any entity keyword
        elif entity_sections:
            # Find the entity that best matches our target
            for keyword in entity_keywords:
                if keyword in entity_sections:
                    target_entity = keyword
                    best_section = entity_sections[keyword][0]
                    print(f"   ✅ PARTIAL MATCH: Found entity '{keyword}' in section {best_section[2]} (rows {best_section[0]}-{best_section[1]})")
                    break

        # Priority 3: If no entity found but we have sections, use the first substantial section (single entity mode only)
        elif table_sections:
            # Look for the first section with substantial data (more than just headers)
            for start_row, end_row, section_idx in [(s[0], s[1], i) for i, s in enumerate(table_sections)]:
                section_size = end_row - start_row + 1
                if section_size > 3:  # At least 4 rows (header + 3 data rows)
                    best_section = (start_row, end_row, section_idx)
                    print(f"   ⚠️  NO ENTITY MATCH: Using first substantial section {section_idx} (rows {start_row}-{end_row})")
                    break

        # Step 5: Extract the selected table section (single entity mode only)
        if best_section:
            start_row, end_row, section_idx = best_section
            df_filtered = df.iloc[start_row:end_row+1].copy()

            # Reset index to make it cleaner but preserve column names
            df_filtered = df_filtered.reset_index(drop=True)

            print(f"   🎯 SELECTED TABLE: Section {section_idx}, rows {start_row}-{end_row} ({len(df_filtered)} rows extracted)")
            if target_entity:
                print(f"   📋 ENTITY: {target_entity}")

            return df_filtered, is_multiple_entity
        else:
            print(f"   ⚠️  No suitable table section found, using original table")
            return df, is_multiple_entity

    # MULTI-ENTITY MODE: Return the entire dataframe without filtering
    # The calling function will handle section selection
    print(f"   🎯 MULTIPLE ENTITY MODE: Returning entire dataframe for section-by-section processing")
    return df, is_multiple_entity

def determine_entity_mode_and_filter_all_sections(df, entity_name, entity_keywords, manual_mode='single'):
    """Determine if we're in single entity or multiple entity mode and return ALL matching sections.

    This function is similar to determine_entity_mode_and_filter but returns ALL sections
    that match the target entity, rather than just the first one. This is needed for
    multi-entity Excel files where the same entity may have multiple tables in the same sheet.

    Args:
        df: DataFrame to process
        entity_name: Target entity name
        entity_keywords: List of entity keywords to search for
        manual_mode: Manual selection - 'single' or 'multiple'

    Returns:
        tuple: (list_of_filtered_dfs, is_multiple_entity)
            - list_of_filtered_dfs: List of DataFrames, one for each matching section
            - is_multiple_entity: Boolean indicating if multiple entities were detected
    """
        # Entity analysis
    # Processing info

    # If manual mode is 'single', skip entity filtering and return original table as single-item list
    if manual_mode == 'single':
        print(f"   🎯 SINGLE ENTITY MODE: Using original table (no filtering applied)")
        return [df], False

    # Step 1: Identify table sections by finding empty rows or major delimiters
    table_sections = []
    current_section_start = None

    for row_idx in range(len(df)):
        # Check if this row is mostly empty (indicates table boundary)
        row_values = df.iloc[row_idx]
        non_empty_count = row_values.notna().sum()
        has_content = non_empty_count > 0

        if has_content and current_section_start is None:
            # Start of a new section
            current_section_start = row_idx
        elif not has_content and current_section_start is not None:
            # End of current section
            table_sections.append((current_section_start, row_idx - 1))
            current_section_start = None
        elif row_idx == len(df) - 1 and current_section_start is not None:
            # Last row, close the current section
            table_sections.append((current_section_start, row_idx))

    print(f"   📊 Found {len(table_sections)} table sections in the worksheet")

    # Step 2: Analyze each section to find entity associations
    entity_sections = {}
    all_potential_entities = set()

    # Multiple entity mode logic is handled above in the main flow

    # Step 3b: For multiple entity mode, apply simplified entity filtering
    if manual_mode == 'multiple':
        # Multiple entity filtering
        # Processing sections

        # Filter table sections to only include those that contain the target entity
        filtered_sections = []

        for section_idx, (start_row, end_row) in enumerate(table_sections):
            # Get all text from this table section
            section_values = []
            for val in df.iloc[start_row:end_row+1].values.flatten():
                if pd.notna(val) and str(val).strip():
                    section_values.append(str(val))

            # Also check column names which might contain entity info
            for col_name in df.columns:
                if pd.notna(col_name) and str(col_name).strip():
                    section_values.append(str(col_name).strip())

            section_text = ' '.join(section_values).lower()

            # Check if this section contains the target entity
            entity_found_in_section = False
            for keyword in entity_keywords:
                keyword_lower = keyword.lower()

                # Simple exact match check
                if keyword_lower in section_text:
                    entity_found_in_section = True
                    print(f"   ✅ ENTITY FOUND: Section {section_idx} (rows {start_row}-{end_row}) contains '{keyword}'")
                    break

                # Check for separator patterns like "Cash - Entity Name"
                separators = [' - ', ' – ', ' and ', ' or ', ' for ']
                for separator in separators:
                    if separator in section_text:
                        parts = section_text.split(separator)
                        for part in parts[1:]:  # Look in parts after the separator
                            part = part.strip()
                            if part and keyword_lower in part:
                                entity_found_in_section = True
                                print(f"   ✅ ENTITY FOUND: Section {section_idx} (rows {start_row}-{end_row}) contains '{keyword}' after '{separator.strip()}'")
                                break
                        if entity_found_in_section:
                            break
                if entity_found_in_section:
                    break

            if entity_found_in_section:
                filtered_sections.append((start_row, end_row, section_idx))

        # Update table_sections to only include sections with the target entity
        if filtered_sections:
            # Found matching sections
            table_sections = filtered_sections
            is_multiple_entity = len(filtered_sections) > 1
        else:
            # No matching sections
            # If no sections contain the target entity, return original table
            return [df], False
    else:
        is_multiple_entity = False

    # Step 4: Use the filtered table sections
    matching_sections = table_sections

    if matching_sections:
        pass  # Using filtered sections
    else:
        # No sections found
        return [df], is_multiple_entity

    # Step 5: Extract ALL matching table sections
    filtered_dfs = []
    if matching_sections:
        for i, (start_row, end_row, section_idx) in enumerate(matching_sections):
            df_filtered = df.iloc[start_row:end_row+1].copy()
            # Reset index to make it cleaner but preserve column names
            df_filtered = df_filtered.reset_index(drop=True)
            filtered_dfs.append(df_filtered)
            print(f"   🎯 EXTRACTED TABLE {i+1}: Section {section_idx}, rows {start_row}-{end_row} ({len(df_filtered)} rows)")

        print(f"   📊 TOTAL EXTRACTED: {len(filtered_dfs)} tables for entity")
        return filtered_dfs, is_multiple_entity
    else:
        print(f"   ⚠️  No suitable table sections found, using original table")
        return [df], is_multiple_entity

def separate_balance_sheet_and_income_statement_tables(df, entity_keywords):
    """
    Separate balance sheet and income statement tables from a combined Excel sheet.
    
    Looks for table headers like:
    - English: "Indicative adjusted balance sheet - Project Ningbo", "Indicative adjusted income statement - Project Ningbo"  
    - Chinese: "经示意性调整后资产负债表 - 东莞岭南", "经示意性调整后利润表 - 东莞岭南"
    
    Returns:
        tuple: (balance_sheet_df, income_statement_df, metadata)
    """
    print(f"🔍 Processing {df.shape[0]} rows × {df.shape[1]} columns")

    balance_sheet_sections = []
    income_statement_sections = []

    # Define keywords for table identification
    bs_keywords_english = ['balance sheet', 'statement of financial position']
    is_keywords_english = ['income statement', 'profit and loss', 'statement of comprehensive income']
    bs_keywords_chinese = ['资产负债表', '财务状况表']
    is_keywords_chinese = ['利润表', '损益表', '综合收益表']

    # Exact patterns for table detection (complete phrases only)
    bs_patterns = ['示意性调整后资产负债表', 'indicative adjusted balance sheet']
    is_patterns = ['示意性调整后利润表', 'indicative adjusted income statement']
    
    # First check column headers for table identification
    for col_idx, col_name in enumerate(df.columns):
        if pd.notna(col_name):
            col_str = str(col_name).lower()
            
            # Check if this column header contains "Indicative adjusted" or "示意性调整后"
            is_indicative_header = (
                ('indicative' in col_str and 'adjusted' in col_str) or
                '示意性调整后' in col_str or
                '经示意性调整后' in col_str
            )
            
            if is_indicative_header:
                # Determine if it's balance sheet or income statement using exact patterns only
                is_balance_sheet = any(pattern in col_str for pattern in bs_patterns)
                is_income_statement = any(pattern in col_str for pattern in is_patterns)

                if is_balance_sheet:
                    print(f"✅ BS found")
                    balance_sheet_sections.append({
                        'header_row': -1,  # Column header
                        'header_col': col_idx,
                        'header_text': str(col_name),
                        'type': 'balance_sheet'
                    })
                elif is_income_statement:
                    print(f"✅ IS found")
                    income_statement_sections.append({
                        'header_row': -1,  # Column header
                        'header_col': col_idx,
                        'header_text': str(col_name),
                        'type': 'income_statement'
                    })
    
    # Then search through ENTIRE dataframe for table headers (not just first 20 rows)
    if not balance_sheet_sections and not income_statement_sections:
        print(f"🔍 Scanning sheet for financial statement headers...")

        for row_idx in range(len(df)):  # Check ALL rows, not just first 20
            for col_idx in range(len(df.columns)):
                cell_value = df.iloc[row_idx, col_idx]
                if pd.notna(cell_value):
                    cell_str = str(cell_value).lower()

                    # Check for EXACT phrases only
                    is_balance_sheet = any(pattern in cell_str for pattern in bs_patterns)
                    is_income_statement = any(pattern in cell_str for pattern in is_patterns)

                    if is_balance_sheet:
                        print(f"✅ Found Balance Sheet header at row {row_idx}")
                        balance_sheet_sections.append({
                            'header_row': row_idx,
                            'header_col': col_idx,
                            'header_text': str(cell_value),
                            'type': 'balance_sheet'
                        })
                    elif is_income_statement:
                        print(f"✅ Found Income Statement header at row {row_idx}")
                        income_statement_sections.append({
                            'header_row': row_idx,
                            'header_col': col_idx,
                            'header_text': str(cell_value),
                            'type': 'income_statement'
                        })
    
    # Extract table data for each identified section
    def extract_table_data(sections, table_type):
        """Extract table data for given sections"""
        if not sections:
            return None
            
        # Use the first section found (most common case)
        section = sections[0]
        header_row = section['header_row']
        
        # If this is balance sheet extraction, check if there's an income statement section
        # and use that as the boundary
        end_boundary = len(df)
        if table_type == 'Balance Sheet' and income_statement_sections:
            is_header_row = income_statement_sections[0]['header_row']
            if is_header_row > header_row:
                end_boundary = is_header_row
                print(f"📍 BS-IS boundary: BS ends at row {is_header_row-1}, IS starts at {is_header_row}")
        
        # If header is in column header (header_row = -1), start from row 0
        if header_row == -1:
            data_start_row = 0
        else:
            # Find the data start row (usually 1-2 rows after header)
            data_start_row = header_row + 1
            
            # Look for the actual data start by finding non-empty rows
            for check_row in range(header_row + 1, min(header_row + 5, len(df))):
                row_data = df.iloc[check_row]
                if row_data.notna().sum() > len(df.columns) * 0.3:  # At least 30% non-empty cells
                    data_start_row = check_row
                    break
        
        # Find the data end row (look for next table header or use boundary)
        data_end_row = end_boundary

        # For column header cases, we need to extract the entire table
        if header_row == -1:
            # Look for table end indicators - scan every row for IS headers
            for check_row in range(data_start_row + 1, len(df)):
                row_data = df.iloc[check_row]

                # Check ALL cells in this row for IS header
                found_is_header = False
                for col_check in range(len(row_data)):
                    cell_val = row_data.iloc[col_check]
                    if pd.notna(cell_val):
                        cell_text = str(cell_val)
                        cell_text_lower = cell_text.lower()

                        # Look for EXACT income statement phrases only
                        if ('示意性调整后利润表' in cell_text_lower or
                            'indicative adjusted income statement' in cell_text_lower):
                            data_end_row = check_row
                            found_is_header = True
                            print(f"📍 {table_type}: Found IS header at row {check_row}")
                            break

                if found_is_header:
                    break

                # Check for empty row separators
                non_empty_count = row_data.notna().sum()
                if non_empty_count == 0:  # Completely empty row
                    # Look ahead to see if next section starts
                    next_section_start = None
                    for look_ahead in range(check_row + 1, min(check_row + 5, len(df))):
                        next_row = df.iloc[look_ahead]
                        if next_row.notna().sum() > len(df.columns) * 0.3:
                            next_section_start = look_ahead
                            break

                    if next_section_start:
                        data_end_row = check_row
                        break
        else:
            # For embedded headers, look for next table header or empty rows
            for check_row in range(data_start_row + 1, len(df)):
                row_data = df.iloc[check_row]

                # Check if this row contains EXACT table headers
                row_text = ' '.join(str(val).lower() for val in row_data if pd.notna(val))
                if ('示意性调整后利润表' in row_text or
                    'indicative adjusted income statement' in row_text):
                    data_end_row = check_row
                    print(f"📍 {table_type}: Found IS header at row {check_row}")
                    break
                elif ('示意性调整后资产负债表' in row_text or
                      'indicative adjusted balance sheet' in row_text):
                    data_end_row = check_row
                    print(f"📍 Another BS header at row {check_row}")
                    break
                
                # Don't stop on empty rows - continue until we find the actual section boundary
                # (Empty rows within a section are normal and should be filtered out later)
        
        # Extract the table data
        table_df = df.iloc[data_start_row:data_end_row].copy()
        print(f"📊 {table_type}: {table_df.shape[0]}×{table_df.shape[1]} data extracted")
        
        # Remove completely empty rows and columns
        table_df = table_df.dropna(how='all').dropna(axis=1, how='all')
        
        # Intelligent column selection - focus on "Indicative adjusted" columns
        selected_cols = []
        desc_col_idx = None
        indicative_cols = []
        
        # Step 1: Find description column (avoid tick symbol columns)
        # Look for columns with actual English/Chinese text, not tick symbols
        desc_col_idx = None

        # First, try to find the column containing "示意性调整后资产负债表"
        for col_idx, col in enumerate(table_df.columns):
            sample_values = table_df[col].dropna().head(10)
            if len(sample_values) > 0:
                # Check if this column contains the BS header text
                has_bs_header = sum(1 for val in sample_values if isinstance(val, str) and '示意性调整后资产负债表' in str(val))
                if has_bs_header > 0:
                    desc_col_idx = col_idx
                    print(f"🎯 Found description column: contains '示意性调整后资产负债表' at col {col_idx}")
                    break

        # If not found, look for columns with significant text content (English/Chinese)
        if desc_col_idx is None:
            for col_idx, col in enumerate(table_df.columns):
                sample_values = table_df[col].dropna().head(10)  # Check more rows
                if len(sample_values) > 0:
                    # Check for tick symbols (✓, √, etc.) - SKIP these columns
                    tick_symbols = ['✓', '√', '✔', '☑', '✓', '■', '□', '●', '○', '★', '☆']
                    has_tick_symbols = sum(1 for val in sample_values if isinstance(val, str) and any(tick in val for tick in tick_symbols))

                    # Check for text content (including Chinese and English)
                    text_count = sum(1 for val in sample_values if isinstance(val, str) and len(val.strip()) > 2)

                    # Consider it a description column if:
                    # - Has NO tick symbols AND has significant text content (including Chinese/English)
                    # - Requirement: At least 20% of first 10 rows must contain text (len > 2 chars)
                    # - This avoids tick symbol columns but allows sparse text columns
                    if has_tick_symbols == 0 and text_count >= len(sample_values) * 0.2:
                        desc_col_idx = col_idx
                        print(f"🎯 Found description column: text content at col {col_idx} ({text_count}/{len(sample_values)} rows have text)")
                        break

        # Final fallback: use first column if no good description column found
        if desc_col_idx is None:
            desc_col_idx = 0
            print(f"⚠️ Using first column as description (fallback)")
        else:
            print(f"✅ Description column set to: {table_df.columns[desc_col_idx]}")
        
        # Step 2: Find ALL "Indicative adjusted" or "示意性调整后" columns
        # Multiple columns with 示意性调整后 indicate which data columns should be included
        for row_idx in range(min(5, len(table_df))):  # Check first 5 rows
            for col_idx, col in enumerate(table_df.columns):
                if col_idx == desc_col_idx:
                    continue  # Skip description column

                cell_value = table_df.iloc[row_idx, col_idx]
                if pd.notna(cell_value):
                    cell_str = str(cell_value).lower()
                    if ('indicative' in cell_str and 'adjusted' in cell_str) or '示意性调整后' in cell_str:
                        # Found an "Indicative adjusted" header - ALL columns with this indicator should be included
                        if col_idx not in indicative_cols:
                            indicative_cols.append(col_idx)

        print(f"📊 COLUMN DETECTION: {len(indicative_cols)} data columns found")
        
        # Step 3: Build final column selection
        selected_cols = [table_df.columns[desc_col_idx]]  # Always include description
        
        if indicative_cols:
            # Add indicative adjusted columns
            for col_idx in indicative_cols:
                if col_idx != desc_col_idx:  # Don't duplicate description column
                    selected_cols.append(table_df.columns[col_idx])
            print(f"✅ DATA COLUMNS: {len(selected_cols)} columns selected for analysis")
        else:
            # Fallback: use description + last 2-3 columns (most recent dates)
            remaining_cols = [col for i, col in enumerate(table_df.columns) if i != desc_col_idx]
            selected_cols.extend(remaining_cols[-3:])  # Last 3 columns
            print(f"⚠️ FALLBACK: Using {len(selected_cols)} columns (no 示意性调整后 indicators found)")
        
        # Remove duplicates while preserving order
        seen = set()
        final_cols = []
        for col in selected_cols:
            if col not in seen:
                final_cols.append(col)
                seen.add(col)
        
        if len(final_cols) > 1:  # Only filter if we have more than description column
            table_df = table_df[final_cols]
        
        # Preprocess data for Income Statement tables
        if table_type == 'Income Statement' and not table_df.empty:
            print(f"   🔧 Preprocessing Income Statement data...")
            
            # 1. Remove rows containing "示意性调整后" or "Indicative adjusted"
            mask = table_df.apply(lambda row: not any(
                '示意性调整后' in str(val) or '示意性調整後' in str(val) or 
                '经示意性调整后' in str(val) or '經示意性調整後' in str(val) or
                'indicative adjusted' in str(val).lower()
                for val in row
            ), axis=1)
            rows_before = len(table_df)
            table_df = table_df[mask]
            print(f"   ✅ Removed {rows_before - len(table_df)} rows with '示意性调整后'")
            
            # 1.5. Limit table to rows up to and including "净利润" row
            net_profit_keywords = ['净利润', '淨利潤', 'net profit', 'net income']
            net_profit_row_idx = None
            for idx, row in table_df.iterrows():
                row_str = ' '.join(str(val).lower() for val in row if pd.notna(val))
                if any(keyword in row_str for keyword in net_profit_keywords):
                    net_profit_row_idx = idx
                    print(f"   📍 Found '净利润' at row index {idx}")
                    break
            
            if net_profit_row_idx is not None:
                # Keep only rows up to and including 净利润
                table_df = table_df.loc[:net_profit_row_idx]
                print(f"   ✅ Limited table to {len(table_df)} rows (up to 净利润)")
            
            # Log processed table to markdown file for verification
            try:
                import os
                from datetime import datetime
                
                # Create output directory if it doesn't exist
                log_dir = 'fdd_utils/output/table_logs'
                os.makedirs(log_dir, exist_ok=True)
                
                # Create daily log file (one file per day)
                daily_log = f"{log_dir}/daily_log_{datetime.now().strftime('%Y%m%d')}.md"
                
                # Generate markdown table
                md_content = f"\n---\n\n# Income Statement Table - {table_type}\n\n"
                md_content += f"**Processed at**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
                md_content += f"**Rows**: {len(table_df)}\n"
                md_content += f"**Columns**: {len(table_df.columns)}\n\n"
                md_content += "## Table Data\n\n"
                md_content += table_df.to_markdown(index=False) + "\n\n"
                
                # Append to daily log file
                with open(daily_log, 'a', encoding='utf-8') as f:
                    f.write(md_content)
                
                print(f"   📝 Logged processed table to: {daily_log}")
            except Exception as e:
                print(f"   ⚠️ Table logging failed: {e}")
            
            # 2. Round numerical columns to 1 decimal place and add thousand separators
            import numpy as np
            for col in table_df.columns:
                if table_df[col].dtype in ['float64', 'int64', 'float32', 'int32']:
                    # Round to 1 decimal and format with thousand separators
                    # Skip if values look like years (2020-2030 range)
                    def format_with_check(x):
                        if pd.notna(x):
                            # Don't format if it's a year value (between 2000 and 2100)
                            if 2000 <= x <= 2100:
                                return x
                            return f"{x:,.1f}"
                        return x
                    table_df[col] = table_df[col].apply(format_with_check)
                elif table_df[col].dtype == 'object':
                    # Try to convert string numbers to formatted numbers
                    def format_number(val):
                        if pd.isna(val) or val == '':
                            return val
                        try:
                            # Skip if already contains Chinese characters or is descriptive text
                            if any('\u4e00' <= char <= '\u9fff' for char in str(val)):
                                return val
                            # Skip if it looks like a date or year (e.g., "2024", "2021-01-01", "2024年1-5月")
                            val_str = str(val)
                            if any(char in val_str for char in ['年', '月', '日', '-', '/']):
                                return val
                            # Try to parse as number
                            num = float(str(val).replace(',', '').replace(' ', ''))
                            # Don't format if it's a year value (between 2000 and 2100)
                            if 2000 <= num <= 2100:
                                return val
                            return f"{num:,.1f}"
                        except (ValueError, TypeError):
                            return val
                    table_df[col] = table_df[col].apply(format_number)
            print(f"   ✅ Applied number formatting: 1 decimal place with thousand separators")
        
        return {
            'data': table_df,
            'header_info': section,
            'data_range': (data_start_row, data_end_row-1)
        }
    
    balance_sheet_data = extract_table_data(balance_sheet_sections, 'Balance Sheet')
    income_statement_data = extract_table_data(income_statement_sections, 'Income Statement')

    # Summary of extraction results
    bs_rows = balance_sheet_data['data'].shape[0] if balance_sheet_data else 0
    is_rows = income_statement_data['data'].shape[0] if income_statement_data else 0
    print(f"📊 Extraction complete: BS={bs_rows} rows, IS={is_rows} rows")

    metadata = {
        'bs_sections_found': len(balance_sheet_sections),
        'is_sections_found': len(income_statement_sections),
        'separation_successful': balance_sheet_data is not None or income_statement_data is not None
    }

    # Summary
    if balance_sheet_data or income_statement_data:
        print(f"📊 SUMMARY: BS sections={len(balance_sheet_sections)}, IS sections={len(income_statement_sections)}")

    return balance_sheet_data, income_statement_data, metadata


def extract_income_statement_table_directly(df, entity_keywords):
    """
    Extract Income Statement table directly using the same logic as Balance Sheet extraction.
    This is a fallback when the separation logic doesn't work properly for IS mode.
    """
    print(f"🔄 DIRECT IS EXTRACTION: Starting extraction for {df.shape}")

    # Look for IS headers in the dataframe
    is_sections = []

    # Check column headers first
    for col_idx, col_name in enumerate(df.columns):
        if pd.notna(col_name):
            col_str = str(col_name).lower()
            if ('示意性调整后利润表' in col_str or
                'indicative adjusted income statement' in col_str):
                print(f"✅ IS header found in column: {col_name}")
                is_sections.append({
                    'header_row': -1,  # Column header
                    'header_col': col_idx,
                    'header_text': str(col_name),
                    'type': 'income_statement'
                })

    # If not found in column headers, search through the dataframe
    if not is_sections:
        print(f"🔍 SCANNING DATAFRAME for IS headers in {df.shape}...")
        found_headers = []
        for row_idx in range(len(df)):
            for col_idx in range(len(df.columns)):
                cell_value = df.iloc[row_idx, col_idx]
                if pd.notna(cell_value):
                    cell_str = str(cell_value).lower()
                    if ('示意性调整后利润表' in cell_str or
                        'indicative adjusted income statement' in cell_str):
                        print(f"✅ IS header found at row {row_idx}, col {col_idx}: {cell_value}")
                        found_headers.append((row_idx, col_idx, cell_value))
                        is_sections.append({
                            'header_row': row_idx,
                            'header_col': col_idx,
                            'header_text': str(cell_value),
                            'type': 'income_statement'
                        })

        if not found_headers:
            print(f"❌ DIRECT IS EXTRACTION: No IS headers found in {len(df)} rows")
            # Debug: Show some sample data to understand the structure
            print(f"📊 Sample data (first 5 rows, first 3 columns):")
            for i in range(min(5, len(df))):
                row_data = []
                for j in range(min(3, len(df.columns))):
                    cell_val = df.iloc[i, j]
                    if pd.notna(cell_val):
                        row_data.append(str(cell_val)[:50])
                    else:
                        row_data.append("NaN")
                print(f"   Row {i}: {row_data}")

    if not is_sections:
        print(f"❌ DIRECT IS EXTRACTION: No IS headers found")
        return None

    # Use the first IS section found
    section = is_sections[0]
    header_row = section['header_row']

    # If header is in column header (header_row = -1), start from row 0
    if header_row == -1:
        data_start_row = 0
        print(f"📊 DIRECT IS: Column header case, starting from row 0")
    else:
        # Find the data start row (usually 1-2 rows after header)
        data_start_row = header_row + 1
        print(f"📊 DIRECT IS: Header at row {header_row}, starting from row {data_start_row}")

        # Look for the actual data start by finding non-empty rows
        for check_row in range(header_row + 1, min(header_row + 5, len(df))):
            row_data = df.iloc[check_row]
            if row_data.notna().sum() > len(df.columns) * 0.3:  # At least 30% non-empty cells
                data_start_row = check_row
                print(f"📊 DIRECT IS: Found data start at row {check_row}")
                break

    # For IS, we want to extract from the IS header to the end of the table
    # Unlike BS which has boundaries with IS, IS usually goes to the end
    data_end_row = len(df)

    # Look for table end indicators
    print(f"🔍 DIRECT IS: Scanning for table end from row {data_start_row + 1} to {len(df)-1}")

    for check_row in range(data_start_row + 1, len(df)):
        row_data = df.iloc[check_row]

        # Check for empty row separators (if we find a completely empty row)
        non_empty_count = row_data.notna().sum()
        if non_empty_count == 0:  # Completely empty row
            # Look ahead to see if next section starts
            next_section_start = None
            for look_ahead in range(check_row + 1, min(check_row + 5, len(df))):
                next_row = df.iloc[look_ahead]
                if next_row.notna().sum() > len(df.columns) * 0.3:
                    next_section_start = look_ahead
                    break

            if next_section_start:
                data_end_row = check_row
                print(f"📊 DIRECT IS: Found table boundary at row {check_row}")
                break

    # Extract the table data
    table_df = df.iloc[data_start_row:data_end_row].copy()
    print(f"📊 DIRECT IS: Extracted table from rows {data_start_row} to {data_end_row-1} → {table_df.shape}")

    # Remove completely empty rows and columns
    table_df = table_df.dropna(how='all').dropna(axis=1, how='all')

    if table_df.empty:
        print(f"❌ DIRECT IS EXTRACTION: Table is empty after cleaning")
        return None

    print(f"✅ DIRECT IS EXTRACTION: Final table shape: {table_df.shape}")
    
    # Preprocess Income Statement data
    print(f"   🔧 Preprocessing Income Statement data...")
    
    # 1. Remove rows containing "示意性调整后" or "Indicative adjusted"
    mask = table_df.apply(lambda row: not any(
        '示意性调整后' in str(val) or '示意性調整後' in str(val) or 
        '经示意性调整后' in str(val) or '經示意性調整後' in str(val) or
        'indicative adjusted' in str(val).lower()
        for val in row
    ), axis=1)
    rows_before = len(table_df)
    table_df = table_df[mask]
    print(f"   ✅ Removed {rows_before - len(table_df)} rows with '示意性调整后'")
    
    # 1.5. Limit table to rows up to and including "净利润" row
    net_profit_keywords = ['净利润', '淨利潤', 'net profit', 'net income']
    net_profit_row_idx = None
    for idx, row in table_df.iterrows():
        row_str = ' '.join(str(val).lower() for val in row if pd.notna(val))
        if any(keyword in row_str for keyword in net_profit_keywords):
            net_profit_row_idx = idx
            print(f"   📍 Found '净利润' at row index {idx}")
            break
    
    if net_profit_row_idx is not None:
        # Keep only rows up to and including 净利润
        table_df = table_df.loc[:net_profit_row_idx]
        print(f"   ✅ Limited table to {len(table_df)} rows (up to 净利润)")
    
    # Log processed table to markdown file for verification
    try:
        import os
        from datetime import datetime
        
        # Create output directory if it doesn't exist
        log_dir = 'fdd_utils/output/table_logs'
        os.makedirs(log_dir, exist_ok=True)
        
        # Create daily log file (one file per day)
        daily_log = f"{log_dir}/daily_log_{datetime.now().strftime('%Y%m%d')}.md"
        
        # Generate markdown table
        md_content = f"\n---\n\n# Income Statement Table - Direct Extraction\n\n"
        md_content += f"**Processed at**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
        md_content += f"**Rows**: {len(table_df)}\n"
        md_content += f"**Columns**: {len(table_df.columns)}\n\n"
        md_content += "## Table Data\n\n"
        md_content += table_df.to_markdown(index=False) + "\n\n"
        
        # Append to daily log file
        with open(daily_log, 'a', encoding='utf-8') as f:
            f.write(md_content)
        
        print(f"   📝 Logged processed table to: {daily_log}")
    except Exception as e:
        print(f"   ⚠️ Table logging failed: {e}")
    
    # 2. Round numerical columns to 1 decimal place and add thousand separators
    import numpy as np
    for col in table_df.columns:
        if table_df[col].dtype in ['float64', 'int64', 'float32', 'int32']:
            # Round to 1 decimal and format with thousand separators
            # Skip if values look like years (2020-2030 range)
            def format_with_check(x):
                if pd.notna(x):
                    # Don't format if it's a year value (between 2000 and 2100)
                    if 2000 <= x <= 2100:
                        return x
                    return f"{x:,.1f}"
                return x
            table_df[col] = table_df[col].apply(format_with_check)
        elif table_df[col].dtype == 'object':
            # Try to convert string numbers to formatted numbers
            def format_number(val):
                if pd.isna(val) or val == '':
                    return val
                try:
                    # Skip if already contains Chinese characters or is descriptive text
                    if any('\u4e00' <= char <= '\u9fff' for char in str(val)):
                        return val
                    # Skip if it looks like a date or year (e.g., "2024", "2021-01-01", "2024年1-5月")
                    val_str = str(val)
                    if any(char in val_str for char in ['年', '月', '日', '-', '/']):
                        return val
                    # Try to parse as number
                    num = float(str(val).replace(',', '').replace(' ', ''))
                    # Don't format if it's a year value (between 2000 and 2100)
                    if 2000 <= num <= 2100:
                        return val
                    return f"{num:,.1f}"
                except (ValueError, TypeError):
                    return val
            table_df[col] = table_df[col].apply(format_number)
    print(f"   ✅ Applied number formatting: 1 decimal place with thousand separators")

    return {
        'data': table_df,
        'header_info': section,
        'data_range': (data_start_row, data_end_row-1)
    }


def filter_to_indicative_adjusted_columns(df):
    """
    Filter dataframe to only show description column + Indicative adjusted columns.
    This is used when no BS/IS separation occurs but we still want to filter columns.
    """
    print(f"🔍 FILTER COLUMNS: Input df shape: {df.shape}")
    print(f"🔍 FILTER COLUMNS: Input df columns: {list(df.columns)}")

    # Reduced output
    
    # Find description column - must contain actual text (not tick symbols)
    desc_col_idx = 0
    for col_idx, col in enumerate(df.columns):
        sample_values = df[col].dropna().head(5)
        if len(sample_values) > 0:
            # Check for actual English/Chinese text content (not symbols)
            text_count = 0
            for val in sample_values:
                if isinstance(val, str):
                    val_clean = str(val).strip()
                    # Must contain actual words (English or Chinese characters)
                    has_english = any(c.isalpha() for c in val_clean)
                    has_chinese = any('\u4e00' <= c <= '\u9fff' for c in val_clean)
                    # Exclude tick symbols and short non-descriptive text
                    if (has_english or has_chinese) and len(val_clean) > 2 and val_clean not in ['✓', '√', '○', '●', '■', '□']:
                        text_count += 1
            
            if text_count >= len(sample_values) * 0.6:  # 60% or more are real text
                desc_col_idx = col_idx
                break
    
    # Find ALL "Indicative adjusted" columns - scan MORE rows and ALL columns
    indicative_cols = []
    print(f"🔍 SCANNING for ALL 示意性调整后 columns in {len(df)} rows...")
    
    for row_idx in range(min(10, len(df))):  # Check first 10 rows (was only 5)
        for col_idx, col in enumerate(df.columns):
            if col_idx == desc_col_idx:
                continue  # Skip description column
                
            cell_value = df.iloc[row_idx, col_idx]
            if pd.notna(cell_value):
                cell_str = str(cell_value).lower()
                if ('indicative' in cell_str and 'adjusted' in cell_str) or '示意性调整后' in cell_str:
                    # Include this column if not already added
                    if col_idx not in indicative_cols:
                        indicative_cols.append(col_idx)
                        print(f"🎯 Found 示意性调整后 at row {row_idx}, col {col_idx}")
    
    print(f"📊 TOTAL 示意性调整后 COLUMNS FOUND: {len(indicative_cols)}")
    
    # Build final column selection
    if indicative_cols:
        selected_cols = [df.columns[desc_col_idx]]  # Always include description
        for col_idx in indicative_cols:
            if col_idx != desc_col_idx:  # Don't duplicate description column
                selected_cols.append(df.columns[col_idx])
        
        # Remove duplicates while preserving order
        seen = set()
        final_cols = []
        for col in selected_cols:
            if col not in seen:
                final_cols.append(col)
                seen.add(col)
        
        filtered_df = df[final_cols]
        print(f"✅ Filtered: {df.shape} → {filtered_df.shape}")
        print(f"✅ Filtered columns: {list(filtered_df.columns)}")
        return filtered_df
    else:
        print(f"⚠️ No 示意性调整后 columns found, returning original df: {df.shape}")
        return df


def find_indicative_adjusted_column_and_dates(df, entity_keywords):
    """Find 'Indicative adjusted' column and extract dates according to new logic."""

    def parse_date(date_str):
        """Parse date string in various formats including xMxx."""
        if not date_str or pd.isna(date_str):
            return None

        date_str = str(date_str).strip()
        
        # AVOID CONFUSING NUMBERS WITH DATES
        # Skip values that look like large numbers with commas (e.g., "2,021,000")
        # These should NOT be interpreted as dates
        if ',' in date_str and date_str.replace(',', '').replace('.', '').isdigit():
            num_val = float(date_str.replace(',', ''))
            if num_val > 10000:  # Large numbers are definitely not dates
                return None

        # Handle Chinese date range format: 2024年1-5月 (use the END month)
        chinese_range_match = re.match(r'^(\d{4})年(\d{1,2})-(\d{1,2})月$', date_str)
        if chinese_range_match:
            year = int(chinese_range_match.group(1))
            end_month = int(chinese_range_match.group(3))  # Use end month
            # Use last day of the end month
            if end_month == 12:
                return datetime(year, 12, 31)
            elif end_month in [1, 3, 5, 7, 8, 10]:
                return datetime(year, end_month, 31)
            elif end_month in [4, 6, 9, 11]:
                return datetime(year, end_month, 30)
            elif end_month == 2:
                if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0):
                    return datetime(year, 2, 29)
                else:
                    return datetime(year, 2, 28)

        # Handle xMxx format (e.g., 9M22, 12M23) - END OF MONTH
        xmxx_match = re.match(r'^(\d+)M(\d{2})$', date_str)
        if xmxx_match:
            month = int(xmxx_match.group(1))
            year = 2000 + int(xmxx_match.group(2))  # Assume 20xx for 2-digit years
            # Use end of month, not beginning (last day of the month)
            if month == 12:
                return datetime(year, 12, 31)  # December 31st
            elif month in [1, 3, 5, 7, 8, 10]:
                return datetime(year, month, 31)  # 31-day months
            elif month in [4, 6, 9, 11]:
                return datetime(year, month, 30)  # 30-day months
            elif month == 2:
                # February - handle leap years
                if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0):
                    return datetime(year, 2, 29)  # Leap year
                else:
                    return datetime(year, 2, 28)  # Non-leap year

        # Handle standard date formats
        date_formats = [
            '%d/%m/%Y', '%d-%m-%Y', '%d/%m/%y', '%d-%m-%y',
            '%Y-%m-%d', '%m/%d/%Y', '%m-%d-%Y',
            '%d/%b/%Y', '%d-%b-%Y', '%b/%d/%Y', '%b-%d-%Y',
            '%d/%B/%Y', '%d-%B-%Y', '%B/%d/%Y', '%B-%d-%Y',
            # Chinese date formats
            '%Y年%m月%d日', '%Y年%m月', '%m月%d日', '%Y/%m/%d',
            '%Y.%m.%d', '%Y年%m月%d日', '%Y年%m月%d号',
            # Additional flexible formats
            '%Y%m%d', '%d%m%Y', '%m%d%Y'
        ]

        for fmt in date_formats:
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                continue

        return None

    print(f"   🔍 SEARCHING for 'Indicative adjusted' (English/Chinese) with new logic...")

    # Step 1: Find "Indicative adjusted" (English/Chinese) positions
    indicative_positions = []

    # Search in first 10 rows for "Indicative adjusted" (English and Chinese)
    for row_idx in range(min(10, len(df))):
        for col_idx, col in enumerate(df.columns):
            val = df.iloc[row_idx, col_idx]
            val_str = str(val).lower()
            # Check for English "indicative adjusted" or Chinese "示意性調整後" / "示意性调整后"
            if pd.notna(val) and (
                ('indicative' in val_str and 'adjusted' in val_str) or
                '示意性調整後' in val_str or
                '示意性调整后' in val_str
            ):
                indicative_positions.append((row_idx, col_idx))
                print(f"   📋 FOUND 'Indicative adjusted' at Row {row_idx}, Col {col_idx} ({col}): '{df.iloc[row_idx, col_idx]}'")
                break  # Found one, move to next row

    if not indicative_positions:
        print(f"   ❌ No 'Indicative adjusted' (English/Chinese) found")
        return None, None, None

    print(f"   📊 Found {len(indicative_positions)} instances of 'Indicative adjusted'")

    # Step 2: Process the first "Indicative adjusted" position (usually the most relevant)
    indic_row, indic_col = indicative_positions[0]
    col_name = df.columns[indic_col]

    # Step 3: Find merged range by looking right until we hit a different value or end
    merge_start = indic_col
    merge_end = indic_col

    # Check if this is a merged cell by looking right
    header_value = str(df.iloc[indic_row, indic_col]).strip()
    for check_col in range(indic_col + 1, len(df.columns)):
        check_value = str(df.iloc[indic_row, check_col]).strip()
        if pd.isna(df.iloc[indic_row, check_col]) or check_value == '' or check_value == header_value:
            merge_end = check_col
        else:
            break

    print(f"   📍 MERGED RANGE: columns {merge_start}-{merge_end} ({col_name} to {df.columns[merge_end]})")

    # Step 4: Look one row below for dates
    date_row = indic_row + 1
    if date_row >= len(df):
        print(f"   ❌ No date row found below 'Indicative adjusted'")
        return None, None, None

    dates_found = []
    latest_date = None
    latest_col = None
    latest_row_number = None

    # Look for dates in the merged range
    for col_idx in range(merge_start, merge_end + 1):
        val = df.iloc[date_row, col_idx]

        if isinstance(val, (pd.Timestamp, datetime)):
            date_val = val if isinstance(val, datetime) else val.to_pydatetime()
            if latest_date is None or date_val > latest_date:
                latest_date = date_val
                latest_col = df.columns[col_idx]
                # Extract row number from column name (e.g., "Unnamed: 23" -> 23)
                if isinstance(latest_col, str) and latest_col.startswith('Unnamed: '):
                    try:
                        latest_row_number = int(latest_col.split(': ')[1])
                    except (ValueError, IndexError):
                        latest_row_number = None
                elif isinstance(latest_col, int):
                    # If column name is an integer, use it directly as the row number
                    latest_row_number = latest_col
                else:
                    latest_row_number = None
        elif pd.notna(val):
            # Use the local parse_date function defined above
            parsed_date = parse_date(str(val))
            if parsed_date and (latest_date is None or parsed_date > latest_date):
                latest_date = parsed_date
                latest_col = df.columns[col_idx]
                # Extract row number from column name
                if isinstance(latest_col, str) and latest_col.startswith('Unnamed: '):
                    try:
                        latest_row_number = int(latest_col.split(': ')[1])
                    except (ValueError, IndexError):
                        latest_row_number = None
                elif isinstance(latest_col, int):
                    # If column name is an integer, use it directly as the row number
                    latest_row_number = latest_col
                else:
                    latest_row_number = None

    # ALWAYS USE RIGHTMOST COLUMN from 示意性调整后 section
    # This ensures we use the most recent data column
    if latest_date:
        # But prioritize the RIGHTMOST column even if it's not the latest date parsed
        # This handles cases where dates are in formats we can't parse
        rightmost_col_idx = merge_end
        rightmost_col = df.columns[rightmost_col_idx]
        
        # Extract row number from column name
        if isinstance(rightmost_col, str) and rightmost_col.startswith('Unnamed: '):
            try:
                rightmost_row_number = int(rightmost_col.split(': ')[1])
            except (ValueError, IndexError):
                rightmost_row_number = None
        elif isinstance(rightmost_col, int):
            rightmost_row_number = rightmost_col
        else:
            rightmost_row_number = None
        
        print(f"   🎯 USING RIGHTMOST COLUMN: '{rightmost_col}' (col {rightmost_col_idx}, Excel row: {rightmost_row_number})")
        print(f"   📅 Date in that column: {df.iloc[date_row, rightmost_col_idx]}")
        
        # Try to extract date from the rightmost column
        rightmost_val = df.iloc[date_row, rightmost_col_idx]
        if isinstance(rightmost_val, (pd.Timestamp, datetime)):
            date_val = rightmost_val if isinstance(rightmost_val, datetime) else rightmost_val.to_pydatetime()
            return date_val.strftime('%Y-%m-%d'), rightmost_col, rightmost_row_number
        elif pd.notna(rightmost_val):
            parsed_date = parse_date(str(rightmost_val))
            if parsed_date:
                return parsed_date.strftime('%Y-%m-%d'), rightmost_col, rightmost_row_number
            else:
                # If can't parse, return the string value as-is
                return str(rightmost_val), rightmost_col, rightmost_row_number
        
        # Fallback to the latest parsed date
        return latest_date.strftime('%Y-%m-%d'), latest_col, latest_row_number
    else:
        # If no valid date found, use the RIGHTMOST column in the merged range (most recent)
        # This handles cases like "2024年1-5月" which can't be parsed as standard dates
        print(f"   ⚠️ No parseable dates found - using RIGHTMOST column as fallback")
        rightmost_col = df.columns[merge_end]
        # Extract row number from column name
        if isinstance(rightmost_col, str) and rightmost_col.startswith('Unnamed: '):
            try:
                rightmost_row_number = int(rightmost_col.split(': ')[1])
            except (ValueError, IndexError):
                rightmost_row_number = None
        elif isinstance(rightmost_col, int):
            rightmost_row_number = rightmost_col
        else:
            rightmost_row_number = None
        print(f"   🎯 FALLBACK: Using rightmost column '{rightmost_col}' (Excel row: {rightmost_row_number})")
        return None, rightmost_col, rightmost_row_number

def parse_accounting_table(df, key, entity_name, sheet_name, latest_date_col=None, actual_entity=None, debug=False, manual_mode='single'):
    """
    Parse accounting table with proper header detection and figure column identification
    Returns structured table data with metadata
    """
        # Parse accounting table
    import re  # Import re inside function to avoid scope issues
    try:
        
        if df.empty or len(df) < 2:
            return None
        
        # Debug info reduced for cleaner output
        if debug:  # Only show if explicitly debugging
            print(f"DEBUG: DataFrame shape: {df.shape}")
        
        # NEW LOGIC: Step 1 - Determine entity mode and filter table if needed
        entity_keywords = [entity_name]  # Start with base entity name
        if hasattr(st, 'session_state') and hasattr(st.session_state, 'get') and 'ai_data' in st.session_state:
            ai_data = st.session_state['ai_data']
            if 'entity_keywords' in ai_data:
                entity_keywords = ai_data['entity_keywords']

        # Use the manual_mode parameter passed from the calling function
        if manual_mode == 'multiple':
            # For multiple entity mode, get ALL matching sections for the entity
            filtered_dfs, is_multiple_entity = determine_entity_mode_and_filter_all_sections(df, entity_name, entity_keywords, manual_mode)
            if filtered_dfs:
                df_filtered = filtered_dfs[0]  # Use the first matching section for now, but we have all sections available
                print(f"   📊 MULTIPLE SECTIONS FOUND: {len(filtered_dfs)} sections for entity '{entity_name}'")
                # Store all sections for potential future use
                if not hasattr(df_filtered, '_all_sections'):
                    df_filtered._all_sections = filtered_dfs
            else:
                df_filtered = df  # Fallback
                is_multiple_entity = False
        else:
            # For single entity mode, use the original function
            df_filtered, is_multiple_entity = determine_entity_mode_and_filter(df, entity_name, entity_keywords, manual_mode)

        # NEW LOGIC: Step 2 - Separate Balance Sheet and Income Statement tables
        bs_data, is_data, separation_metadata = separate_balance_sheet_and_income_statement_tables(df_filtered, entity_keywords)
        
        # Store separation results in session state for later use
        if hasattr(st, 'session_state'):
            if 'separated_tables' not in st.session_state:
                st.session_state['separated_tables'] = {}
            st.session_state['separated_tables'][sheet_name] = {
                'balance_sheet': bs_data,
                'income_statement': is_data,
                'metadata': separation_metadata
            }
        
        # For current processing, choose the appropriate table based on context
        # If we have a specific statement type preference, use that table
        current_statement_type = 'BS'  # Default
        if hasattr(st, 'session_state') and hasattr(st.session_state, 'get'):
            current_statement_type = st.session_state.get('current_statement_type', 'BS')
        
        # Select the appropriate data based on statement type
        if current_statement_type == 'BS' and bs_data:
            df_to_process = bs_data['data']
            print(f"🎯 USING BS DATA ONLY: {df_to_process.shape}")
            print(f"🎯 BS data range: {bs_data.get('data_range', 'None')}")
        elif current_statement_type == 'IS' and is_data:
            df_to_process = is_data['data']
            print(f"🎯 USING IS DATA ONLY: {df_to_process.shape}")
            print(f"🎯 IS data range: {is_data.get('data_range', 'None')}")
            print(f"🎯 IS header info: {is_data.get('header_info', 'None')}")
        elif current_statement_type == 'IS':
            # Special handling for IS mode - try to extract IS table directly
            print(f"🔄 IS MODE: Attempting direct IS table extraction...")
            print(f"🔄 IS MODE: Input df_filtered shape: {df_filtered.shape}")
            print(f"🔄 IS MODE: Input df_filtered columns: {list(df_filtered.columns)}")

            # Debug: Show first few rows of input data
            print(f"🔄 IS MODE: First 10 rows of input data:")
            for i in range(min(10, len(df_filtered))):
                print(f"   Row {i}: {list(df_filtered.iloc[i])}")

            is_table_data = extract_income_statement_table_directly(df_filtered, entity_keywords)
            if is_table_data:
                df_to_process = is_table_data['data']
                print(f"✅ DIRECT IS EXTRACTION: {df_to_process.shape}")
                print(f"✅ IS data range: {is_table_data.get('data_range', 'None')}")
                print(f"✅ IS columns: {list(df_to_process.columns)}")
                print(f"✅ IS first few rows:")
                for i in range(min(5, len(df_to_process))):
                    print(f"   Row {i}: {list(df_to_process.iloc[i])}")
            else:
                # Fallback to original data
                df_to_process = df_filtered
                print(f"⚠️ DIRECT IS EXTRACTION FAILED - using original: {df_to_process.shape}")
        else:
            # Fallback to original data if separation didn't work or no preference
            df_to_process = df_filtered
            print(f"⚠️ NO SEPARATION - using original: {df_to_process.shape}")
            print(f"⚠️ current_statement_type: {current_statement_type}")
            print(f"⚠️ bs_data exists: {bs_data is not None}")
            print(f"⚠️ is_data exists: {is_data is not None}")
        
        # Apply column filtering to the selected table
        df_to_process = filter_to_indicative_adjusted_columns(df_to_process)
        
        # Continue with date detection on the selected table
        extracted_date, selected_column, row_number = find_indicative_adjusted_column_and_dates(df_to_process, entity_keywords)

        # Keep the original dataframe structure - don't filter columns
        df_clean = df_filtered.copy()
        
        # Debug: Print the original and cleaned dataframe info
        # print(f"   🔍 DEBUG: Original df shape: {df.shape}")
        # print(f"   🔍 DEBUG: Cleaned df_clean shape: {df_clean.shape}")
        # print(f"   🔍 DEBUG: Original df columns: {list(df.columns)}")
        # print(f"   🔍 DEBUG: Cleaned df_clean columns: {list(df_clean.columns)}")
        # print(f"   🔍 DEBUG: Original df first few rows:")
        # for i in range(min(5, len(df))):
        #     print(f"      Row {i}: {list(df.iloc[i])}")
        # print(f"   🔍 DEBUG: Cleaned df_clean first few rows:")
        # for i in range(min(5, len(df_clean))):
        #     print(f"      Row {i}: {list(df_clean.iloc[i])}")
        
        # If all columns were dropped, try a different approach
        if len(df_clean.columns) == 0:
            # Try to find columns with actual data
            df_clean = df.copy()
            for col in df_clean.columns:
                # Check if column has any non-null, non-empty values
                non_null_count = df_clean[col].notna().sum()
                non_empty_count = (df_clean[col].astype(str).str.strip() != '').sum()
                if non_null_count == 0 and non_empty_count == 0:
                    df_clean = df_clean.drop(columns=[col])
        
        # Ensure the latest_date_col is in the cleaned DataFrame
        if latest_date_col and latest_date_col not in df_clean.columns:
            # Add the column back from the original DataFrame
            df_clean[latest_date_col] = df[latest_date_col]
        
        # Convert to string for easier processing
        df_str = df_clean.astype(str)
        
        # Extract multiplier and currency info from the first few rows
        multiplier = 1
        currency_info = "CNY"
        
        # Look for thousand/million indicators in first 3 rows
        thousand_indicators = ["'000", "'000", "cny'000", "thousands", "人民币千元", "人民幣千元"]
        million_indicators = ["'000,000", "millions", "cny'000,000", "人民币万元", "人民幣萬元"]


        for i in range(min(3, len(df_str))):
            for j in range(len(df_str.columns)):
                cell_value = str(df_str.iloc[i, j]).lower()
                original_cell_value = str(df_str.iloc[i, j])

                # Check thousand indicators
                if any(indicator in cell_value for indicator in thousand_indicators):
                    multiplier = 1000
                    currency_info = "CNY'000"
                    print(f"   💰 FOUND: '{original_cell_value}' at Row {i}, Col {j} - Detected as {currency_info}")
                    break
                # Check million indicators
                elif any(indicator in cell_value for indicator in million_indicators):
                    multiplier = 1000000
                    currency_info = "CNY'000,000"
                    print(f"   💰 FOUND: '{original_cell_value}' at Row {i}, Col {j} - Detected as {currency_info}")
                    break

        # If no currency indicators found, show what we looked at
        if multiplier == 1:
            print(f"   ⚠️  No currency indicators found in first 3 rows of sheet '{sheet_name}'")
            print(f"   💡 Looking for: {thousand_indicators + million_indicators}")
            print(f"   📝 First row content: {' | '.join(str(df_str.iloc[0, j]) for j in range(min(5, len(df_str.columns))))}")
        
        # Find the value column - use detected latest_date_col if provided
        value_col_idx = None
        value_col_name = ""
        
        if latest_date_col and latest_date_col in df_str.columns:
            # Use the detected latest date column
            value_col_idx = df_str.columns.get_loc(latest_date_col)
            # Get the actual column header text from the Excel file
            value_col_name = str(df_str.columns[value_col_idx])
            print(f"   🎯 Using detected latest date column: {latest_date_col} (col {value_col_idx}) - Column name: '{value_col_name}'")
            # print(f"   🔍 DEBUG: Column names: {list(df_str.columns)}")
            # print(f"   🔍 DEBUG: latest_date_col: {latest_date_col}")
            # print(f"   🔍 DEBUG: value_col_idx: {value_col_idx}")
        else:
            # Fallback to original logic
            print(f"   ⚠️  No latest_date_col provided, using fallback detection")
            for i in range(min(3, len(df_str))):  # Check first 3 rows for headers
                for j in range(len(df_str.columns)):
                    cell_value = str(df_str.iloc[i, j]).lower()
                    # Check for English "indicative adjusted" or Chinese "示意性調整後" / "示意性调整后"
                    if ("indicative adjusted" in cell_value or
                        "示意性調整後" in cell_value or
                        "示意性调整后" in cell_value):
                        value_col_idx = j
                        # Get the actual column header text from the Excel file
                        value_col_name = str(df_str.iloc[i, j])
                        print(f"   🎯 Found indicative adjusted column: '{value_col_name}' (col {j})")
                        # Found value column indicator
                        break
                    elif "total" in cell_value and value_col_idx is None:
                        value_col_idx = j
                        value_col_name = "Total"
                        # Found total column
        
        # If still no specific column found, look for any column with financial data patterns
        if value_col_idx is None:
            candidate_cols = []
            for j in range(len(df_str.columns)):
                column_data = df_str.iloc[:, j]
                numeric_count = 0
                total_cells = 0
                large_numbers = 0
                for cell in column_data:
                    cell_str = str(cell).strip()
                    if cell_str and cell_str.lower() not in ['nan', '']:
                        total_cells += 1
                        if re.search(r'^\d+\.?\d*$', cell_str.replace(',', '')):
                            numeric_count += 1
                            # Check if it's a large number (likely financial data)
                            try:
                                num_val = float(cell_str.replace(',', ''))
                                if num_val > 100:  # Skip small numbers like 1, 2, 3, 1000
                                    large_numbers += 1
                            except ValueError:
                                pass
                
                # Only consider columns with significant large numbers
                if total_cells > 0 and numeric_count >= total_cells * 0.3 and large_numbers >= 2:
                    candidate_cols.append(j)
                    # Found good candidate column
        
        # If no specific column found, use the rightmost column with numbers
        if value_col_idx is None:
            candidate_cols = []
            for j in range(len(df_str.columns) - 1, -1, -1):
                column_data = df_str.iloc[:, j]
                numeric_count = 0
                total_cells = 0
                large_numbers = 0
                for cell in column_data:
                    cell_str = str(cell).strip()
                    if cell_str and cell_str.lower() not in ['nan', '']:
                        total_cells += 1
                        if re.search(r'^\d+\.?\d*$', cell_str.replace(',', '')):
                            numeric_count += 1
                            # Check if it's a large number (likely financial data)
                            try:
                                num_val = float(cell_str.replace(',', ''))
                                if num_val > 100:  # Skip small numbers like 1, 2, 3, 1000
                                    large_numbers += 1
                            except ValueError:
                                pass
                
                # Only consider columns with significant large numbers
                if total_cells > 0 and numeric_count >= total_cells * 0.3 and large_numbers >= 2:
                    candidate_cols.append(j)
                    # Found good candidate column
            
            if candidate_cols:
                value_col_idx = candidate_cols[0]  # Use the rightmost good column
                value_col_name = "Financial Data"
        
        if value_col_idx is None:
            return None
        
        # SMART DESCRIPTION COLUMN DETECTION - Support MULTIPLE description columns
        # The column with "人民币千元" is actually the FIRST DESCRIPTION COLUMN (not a unit column!)
        # Structure is: [Desc1] → [Desc2] → [Date Headers] → [Value Columns]
        desc_col_indices = []
        first_desc_col_idx = None
        
        # Find the FIRST description column (contains "人民币千元" as header)
        for i in range(min(5, len(df_str))):
            for j in range(len(df_str.columns)):
                cell_value = str(df_str.iloc[i, j]).lower()
                if any(indicator in cell_value for indicator in ['人民币千元', '人民幣千元', "cny'000"]):
                    first_desc_col_idx = j
                    print(f"   📍 Found FIRST description column at {j}: '{df_str.iloc[i, j]}'")
                    break
            if first_desc_col_idx is not None:
                break
        
        # Check if the next column (after first desc) contains dates or is another description
        def is_date_value(val):
            """Check if a value looks like a date"""
            if pd.isna(val) or val == '' or str(val).lower() == 'nan':
                return False
            val_str = str(val).strip()
            # Check for date patterns
            date_patterns = [
                r'\d{4}年\d{1,2}月\d{1,2}日',  # Chinese: 2021年12月31日
                r'\d{4}年\d{1,2}月',            # Chinese: 2024年5月
                r'\d{4}年\d{1,2}-\d{1,2}月',    # Chinese: 2024年1-5月
                r'\d{4}/\d{1,2}/\d{1,2}',       # English: 2021/12/31
                r'\d{1,2}/\d{1,2}/\d{4}',       # English: 31/12/2021
                r'\d{4}-\d{1,2}-\d{1,2}',       # ISO: 2021-12-31
            ]
            for pattern in date_patterns:
                if re.search(pattern, val_str):
                    return True
            return False
        
        if first_desc_col_idx is not None:
            # Always include the first description column
            desc_col_indices.append(first_desc_col_idx)
            
            # Check the next column (adjacent to first desc column)
            next_col_idx = first_desc_col_idx + 1
            if next_col_idx < len(df_str.columns):
                # Check if this column contains dates (skip if so) or is another description column
                has_dates = False
                for i in range(min(5, len(df_str))):
                    if is_date_value(df_str.iloc[i, next_col_idx]):
                        has_dates = True
                        print(f"   📅 Found date in column {next_col_idx}: '{df_str.iloc[i, next_col_idx]}' - skipping")
                        break
                
                if not has_dates:
                    # This is likely a second description column (e.g., bank names, customer names)
                    desc_col_indices.append(next_col_idx)
                    print(f"   ✅ Using TWO description columns: {desc_col_indices[0]} (main category) and {desc_col_indices[1]} (sub-category)")
                else:
                    print(f"   ✅ Using ONE description column: {desc_col_indices[0]} (main category)")
        
        # Fallback to column 0 and 1 if not detected
        if not desc_col_indices:
            desc_col_indices = [0, 1]
            print(f"   ⚠️  Using fallback description columns: {desc_col_indices}")
        
        # Validate description column indices
        desc_col_indices = [idx for idx in desc_col_indices if idx < len(df_str.columns)]
        if not desc_col_indices:
            desc_col_indices = [0]
            print(f"   ⚠️  Description columns out of bounds, adjusted to: {desc_col_indices}")
        
        print(f"   📊 Final description column indices: {desc_col_indices}")
        
        # For backward compatibility, keep desc_col_idx as the first description column
        desc_col_idx = desc_col_indices[0]
        
        # Find data start row (first row with actual numeric data)
        data_start_row = None
        for i in range(len(df_str)):
            cell_value = str(df_str.iloc[i, value_col_idx])
            # Look for cells that contain numbers (more flexible)
            if re.search(r'\d+', cell_value) and cell_value.strip() not in ['nan', '']:
                # Check if this looks like a data row (has both description and value)
                desc_cell = str(df_str.iloc[i, desc_col_idx]).strip()
                if desc_cell and desc_cell.lower() not in ['nan', '']:
                    # Simplified check: just make sure it's not a pure number and not a header
                    if (not re.match(r'^\d+\.?\d*$', desc_cell) and
                        desc_cell.lower() not in ['cny\'000', 'audited', 'mgt acc', 'indicative adjusted', 'indicative adjustment']):
                        data_start_row = i
                        # print(f"   🔍 Found data start row at {i}: desc='{desc_cell}', value='{cell_value}'")
                        break
        
        # Debug: Print the data start row and first few rows for verification
        # print(f"   🔍 DEBUG: data_start_row = {data_start_row}")
        # if data_start_row is not None:
        #     for i in range(data_start_row, min(data_start_row + 5, len(df_str))):
        #         desc_cell = str(df_str.iloc[i, 0]).strip()
        #         value_cell = str(df_str.iloc[i, value_col_idx]).strip()
        #         print(f"   🔍 DEBUG: Row {i}: desc='{desc_cell}', value='{value_cell}'")
        
        if data_start_row is None:
            # Fallback: start from row 3 if we have at least 4 rows (typical structure)
            if len(df_str) >= 4:
                data_start_row = 3
            elif len(df_str) >= 3:
                data_start_row = 2
            else:
                return None
        
        # Debug: Print the data start row and first few rows for verification
        # print(f"   🔍 DEBUG: Final data_start_row = {data_start_row}")
        # if data_start_row is not None:
        #     for i in range(data_start_row, min(data_start_row + 5, len(df_str))):
        #         desc_cell = str(df_str.iloc[i, 0]).strip()
        #         value_cell = str(df_str.iloc[i, value_col_idx]).strip()
        #         print(f"   🔍 DEBUG: Final Row {i}: desc='{desc_cell}', value='{value_cell}'")
        
        # Use the extracted date from new logic if available, otherwise fallback to old logic
        if extracted_date is None:
            extracted_date = None

            if latest_date_col and latest_date_col in df_str.columns and value_col_idx is not None:
                # Look for date in the detected latest date column
                from datetime import datetime
                col_idx = df_str.columns.get_loc(latest_date_col)
                for i in range(min(5, len(df_str))):
                    val = df_str.iloc[i, col_idx]
                    print(f"   🔍 Checking row {i}, value: {repr(val)} (type: {type(val)})")

                    if isinstance(val, datetime):
                        extracted_date = val.strftime('%Y-%m-%d')
                        print(f"   📅 Extracted date from detected column {latest_date_col}: {extracted_date}")
                        break
                    elif isinstance(val, pd.Timestamp):
                        date_val = val.to_pydatetime()
                        extracted_date = date_val.strftime('%Y-%m-%d')
                        print(f"   📅 Extracted timestamp from detected column {latest_date_col}: {extracted_date}")
                        break
                    elif pd.notna(val):
                        # Try to convert to datetime if it's a different type
                        try:
                            date_val = pd.to_datetime(val)
                            extracted_date = date_val.strftime('%Y-%m-%d')
                            print(f"   📅 Converted and extracted date from detected column {latest_date_col}: {extracted_date}")
                            break
                        except:
                            print(f"   ⚠️  Could not convert {repr(val)} to date")
                            continue
        
        # Fallback: Extract date using pattern matching if not found in detected column
        if not extracted_date:
            print(f"   🔍 Fallback: Searching for date patterns in table")
            extracted_date = None
        date_patterns = [
            r'(\d{4})-(\d{1,2})-(\d{1,2})',  # YYYY-MM-DD
            r'(\d{1,2})/(\d{1,2})/(\d{4})',  # DD/MM/YYYY
            r'(\d{1,2})-(\d{1,2})-(\d{4})',  # DD-MM-YYYY
            r'(\d{4})/(\d{1,2})/(\d{1,2})',  # YYYY/MM/DD
        ]
        
        for i in range(min(5, len(df_str))):
            for j in range(len(df_str.columns)):
                cell_value = str(df_str.iloc[i, j])
                for pattern in date_patterns:
                    match = re.search(pattern, cell_value)
                    if match:
                        try:
                            # Parse different date formats
                            if pattern == r'(\d{4})-(\d{1,2})-(\d{1,2})':
                                year, month, day = match.groups()
                            elif pattern == r'(\d{1,2})/(\d{1,2})/(\d{4})':
                                day, month, year = match.groups()
                            elif pattern == r'(\d{1,2})-(\d{1,2})-(\d{4})':
                                day, month, year = match.groups()
                            elif pattern == r'(\d{4})/(\d{1,2})/(\d{1,2})':
                                year, month, day = match.groups()
                            
                            # Validate date
                            from datetime import datetime
                            dt = datetime(int(year), int(month), int(day))
                            extracted_date = dt.strftime('%Y-%m-%d')
                            print(f"   📅 Fallback date extracted: {extracted_date}")
                            break
                        except (ValueError, TypeError):
                            continue
                if extracted_date:
                    break
            if extracted_date:
                break
        
        # Extract table metadata (first few rows before data)
        # Use actual entity found in data if available, otherwise use entity_name
        display_entity = actual_entity if actual_entity else entity_name

        # NEW LOGIC: Generate clean table name without row numbers
        table_name = f"{key} - {display_entity}"

        table_metadata = {
            'table_name': table_name,
            'sheet_name': sheet_name,
            'date': extracted_date,
            'currency_info': currency_info,
            'multiplier': multiplier,
            'value_column': value_col_name,
            'data_start_row': data_start_row,
            'selected_column': selected_column,
            'excel_row_number': row_number,
            'entity_mode': 'multiple' if is_multiple_entity else 'single'
        }
        
        # Extract data rows - FILTER FOR SELECTED ENTITY ONLY
        data_rows = []
        
        # Initialize forward-fill tracking for first description column (accounting convention)
        last_first_desc = None

        # For single entity mode, we need to identify which rows belong to the selected entity
        if manual_mode == 'single' and actual_entity:
            # First, try to find the entity in the current dataframe
            entity_found_in_current_df = False
            for row_idx in range(len(df_str)):
                row_text = ' '.join(str(val).lower() for val in df_str.iloc[row_idx] if pd.notna(val))
                if actual_entity.lower() in row_text:
                    entity_found_in_current_df = True
                    print(f"   ✅ Found {actual_entity} in current dataframe at row {row_idx}")
                    break

            # If entity is not found in current dataframe, it means we already extracted the correct section
            # Trust the extraction logic and process all data (since it's already filtered)
            if not entity_found_in_current_df:
                print(f"   ℹ️  {actual_entity} not found in current section (already pre-filtered), processing all data")
                # Don't modify data_start_row/data_end_row - process the entire provided dataframe
            else:
                # Entity found, apply section filtering
                entity_start_row = None
                entity_end_row = None

                for row_idx in range(len(df_str)):
                    row_text = ' '.join(str(val).lower() for val in df_str.iloc[row_idx] if pd.notna(val))
                    if actual_entity.lower() in row_text:
                        if entity_start_row is None:
                            entity_start_row = row_idx
                        # Continue until we find the next entity or end of sheet
                        for next_row in range(row_idx + 1, len(df_str)):
                            next_row_text = ' '.join(str(val).lower() for val in df_str.iloc[next_row] if pd.notna(val))
                            if any(other_entity.lower() in next_row_text for other_entity in ['ningbo', 'haining', 'nanjing'] if other_entity != actual_entity.lower().split()[0]):
                                entity_end_row = next_row - 1
                                break
                        if entity_end_row is None:
                            entity_end_row = len(df_str) - 1
                        break

                if entity_start_row is not None:
                    data_start_row = max(data_start_row, entity_start_row + 1)  # Start after entity header
                    data_end_row = min(len(df_str), entity_end_row)
                    print(f"   🎯 FILTERED FOR {actual_entity}: rows {data_start_row}-{data_end_row}")
                else:
                    print(f"   ⚠️  Could not find {actual_entity} section, using all data")

        for i in range(data_start_row, len(df_str)):
            # Stop if we've reached the end of the selected entity's section
            if 'data_end_row' in locals() and i > data_end_row:
                break

            # Get description from ALL detected description columns
            # IMPORTANT: Implement forward-fill for first description column (accounting convention)
            # When first column is blank, it means "same as above" to avoid redundancy
            # KEEP AS SEPARATE COLUMNS - DON'T MERGE!
            desc_values = {}  # Use dict to store separate description values
            
            for idx, desc_idx in enumerate(desc_col_indices):
                if desc_idx < len(df_str.columns):
                    cell_val = str(df_str.iloc[i, desc_idx]).strip()
                    
                    # Forward-fill logic for FIRST description column only
                    if idx == 0:  # First description column
                        if cell_val and cell_val.lower() not in ['nan', '', 'none']:
                            # Check if this is a total/subtotal row - reset forward-fill
                            if any(keyword in cell_val.lower() for keyword in ['total', 'subtotal', '小计', '合计', '总计']):
                                last_first_desc = None  # Reset forward-fill at section boundaries
                                desc_values[f'description_{idx}'] = cell_val
                            else:
                                # Valid value - update last known value and use it
                                last_first_desc = cell_val
                                desc_values[f'description_{idx}'] = cell_val
                        else:
                            # Empty value - use last known value (forward-fill)
                            if last_first_desc:
                                desc_values[f'description_{idx}'] = last_first_desc
                            else:
                                desc_values[f'description_{idx}'] = ''
                    else:
                        # For other description columns, just use the value as-is
                        if cell_val and cell_val.lower() not in ['nan', '', 'none']:
                            desc_values[f'description_{idx}'] = cell_val
                        else:
                            desc_values[f'description_{idx}'] = ''
            
            # Create description string for backward compatibility (merged version)
            desc_parts = [v for v in desc_values.values() if v]
            if desc_parts:
                desc_cell = ' | '.join(desc_parts)
            else:
                desc_cell = ''
            
            value_cell = str(df_str.iloc[i, value_col_idx]).strip()

            # Debug: Print what we're extracting
            # print(f"   🔍 DEBUG: Extracting from row {i}: desc_cols={desc_col_indices}='{desc_cell}', value_col={value_col_idx}='{value_cell}'")

            # Skip empty rows
            if desc_cell.lower() in ['nan', ''] and value_cell.lower() in ['nan', '']:
                continue

            # Skip header rows
            if desc_cell.lower() in ['description', 'item', 'account', 'details']:
                continue
            
            # Process value
            if value_cell and value_cell.lower() not in ['nan', '']:
                try:
                    # FIRST: Check if it's a date format BEFORE attempting numeric conversion
                    if re.search(r'\d{4}年\d{1,2}月\d{1,2}日', value_cell) or \
                       re.search(r'\d{4}/\d{1,2}/\d{1,2}', value_cell) or \
                       re.search(r'\d{1,2}/\d{1,2}/\d{4}', value_cell):
                        # This is a date row - skip it since date is already shown in info bar
                        continue
                    else:
                        # Try to parse as number only if it's not a date
                        clean_value = re.sub(r'[^\d.-]', '', value_cell)
                        if clean_value:
                            numeric_value = float(clean_value)
                            # Apply multiplier
                            final_value = numeric_value * multiplier

                            # Determine if this is a total row
                            is_total = 'total' in desc_cell.lower()

                            # print(f"   📊 ADDING ITEM: {desc_cell} = {final_value}")
                            row_data = {
                                'description': desc_cell,  # Merged description for backward compatibility
                                'value': final_value,
                                'is_total': is_total
                            }
                            
                            # Add SEPARATE description columns for multi-column display
                            for desc_key, desc_val in desc_values.items():
                                row_data[desc_key] = desc_val
                            
                            # Store number of description columns for rendering
                            row_data['num_desc_columns'] = len(desc_col_indices)
                            
                            data_rows.append(row_data)
                        else:
                            # Skip non-numeric, non-date values
                            continue
                except (ValueError, TypeError):
                    # For non-numeric values that might be dates, skip them
                    if re.search(r'\d{4}年\d{1,2}月\d{1,2}日', value_cell) or \
                       re.search(r'\d{4}/\d{1,2}/\d{1,2}', value_cell) or \
                       re.search(r'\d{1,2}/\d{1,2}/\d{4}', value_cell):
                        # This is a date row - skip it since date is already shown in info bar
                        continue
                    else:
                        # Skip rows with non-numeric, non-date values
                        continue
        
        if not data_rows:
            return None
        
        # Include filtered data if entity filtering was applied
        result = {
            'metadata': table_metadata,
            'data': data_rows,
            'raw_df': df_clean
        }

        # Add filtered data if entity filtering was applied
        # Check if entity filtering was actually applied
        # For manual_mode='multiple', we should always include filtered_data since entity filtering was attempted
        entity_filtering_applied = (manual_mode == 'multiple')

        if entity_filtering_applied:
            result['filtered_data'] = df_clean
            print(f"   🔄 ENTITY FILTERING APPLIED: {df.shape} → {df_clean.shape} rows")
            print(f"   📦 Adding filtered_data to result: {df_clean.shape}")
        else:
            print(f"   📋 Single entity mode: using original data")

        return result
        
    except Exception as e:
        print(f"Error parsing accounting table: {e}")
        return None



def process_and_filter_excel(filename, tab_name_mapping, entity_name, entity_suffixes):
    """
    Process and filter Excel file to extract relevant worksheet sections
    This is the core function from old_ver/utils/utils.py
    """
    import re  # Import re inside function to avoid scope issues
    try:
        
        # Load the Excel file
        main_dir = Path(__file__).parent.parent
        file_path = main_dir / filename
        xl = pd.ExcelFile(file_path)
        
        # Create a reverse mapping from values to keys
        reverse_mapping = {}
        if tab_name_mapping is not None:
            for key, values in tab_name_mapping.items():
                for value in values:
                    reverse_mapping[value] = key
                # Also map the key name directly to itself (for sheet names like "Cash", "AR")
                reverse_mapping[key] = key
        
        # Get the sheet names that are relevant
        relevant_sheets = [sheet for sheet in xl.sheet_names if sheet in reverse_mapping]
        
        # Process sheets and extract data
        markdown_content = ""
        
        for sheet_name in relevant_sheets:
            if sheet_name in reverse_mapping:
                df = xl.parse(sheet_name)
                
                # Detect latest date column for this sheet
                # Processing sheet
                # Create entity keywords for this function call
                sheet_entity_keywords = [f"{entity_name} {suffix}" for suffix in entity_suffixes if suffix]
                if not sheet_entity_keywords:
                    sheet_entity_keywords = [entity_name]
                latest_date_col = detect_latest_date_column(df, sheet_name, sheet_entity_keywords)
                if latest_date_col:
                    print(f"✅ Sheet {sheet_name}: Selected column {latest_date_col}")
                else:
                    print(f"⚠️  Sheet {sheet_name}: No date column detected")
                
                # RESTORED WORKING LOGIC: Simple table splitting that worked well
                if len(df) > 50:  # Split medium to large sheets
                    # Split dataframes on empty rows - SIMPLIFIED APPROACH
                    empty_rows = df.index[df.isnull().all(1)]
                    start_idx = 0
                    dataframes = []

                    for end_idx in empty_rows:
                        if end_idx > start_idx:
                            split_df = df.iloc[start_idx:end_idx]
                            if not split_df.dropna(how='all').empty:
                                dataframes.append(split_df)
                            start_idx = end_idx + 1

                    if start_idx < len(df):
                        dataframes.append(df.iloc[start_idx:])
                else:
                    # Process small sheets as single dataframe
                    print(f"📈 [{reverse_mapping[sheet_name]}] Processing as single dataframe (size: {len(df)})")
                    dataframes = [df]
                
                # RESTORED WORKING ENTITY MATCHING: Simple and effective approach
                entity_keywords = [entity_name]  # Start with base entity name
                # Add variations that were working before
                entity_keywords.extend([f"{entity_name} {suffix}" for suffix in entity_suffixes if suffix])

                print(f"🔑 Entity keywords for {sheet_name}: {entity_keywords}")

                # Track if we found entity data in this sheet
                entity_found_in_sheet = False

                # ENHANCED ENTITY MATCHING: Scan entire sheet for entity names
                entity_found_in_sheet = False
                selected_data_frame = None

                # Scan the entire sheet for entity names (not just first 5 rows)
                for row_idx in range(len(df)):
                    row = df.iloc[row_idx]
                    row_text = ' '.join(str(val).lower() for val in row if pd.notna(val))

                    # Check if this row contains any entity keyword
                    entity_in_row = any(kw.lower() in row_text for kw in entity_keywords)

                    if entity_in_row:
                        print(f"   📍 Found entity '{[kw for kw in entity_keywords if kw.lower() in row_text][0]}' in row {row_idx}: {row_text[:100]}...")
                        entity_found_in_sheet = True

                        # Find the table section that starts after this entity row
                        # Look for the next substantial data block (skip empty rows)
                        table_start_row = row_idx + 1

                        # Skip empty rows to find the actual table start
                        while table_start_row < len(df):
                            table_row = df.iloc[table_start_row]
                            has_data = any(pd.notna(val) and str(val).strip() for val in table_row)
                            if has_data:
                                break
                            table_start_row += 1

                        # Extract the table section from this point
                        if table_start_row < len(df):
                            # Find where this table section ends (next empty row or end of sheet)
                            table_end_row = table_start_row
                            while table_end_row < len(df):
                                end_row = df.iloc[table_end_row]
                                has_data = any(pd.notna(val) and str(val).strip() for val in end_row)
                                if not has_data:
                                    break
                                table_end_row += 1

                            # Extract this specific entity table section
                            selected_data_frame = df.iloc[table_start_row:table_end_row+1].copy()
                            print(f"   🎯 Extracted entity table: rows {table_start_row}-{table_end_row} ({len(selected_data_frame)} rows)")
                            break

                # If no entity-specific section found, fall back to original logic
                if not entity_found_in_sheet:
                    for data_frame in dataframes:
                        all_text = ' '.join(data_frame.astype(str).values.flatten()).lower()
                        entity_match = any(kw.lower() in all_text for kw in entity_keywords)

                        if entity_match:
                            print(f"🚀 [{reverse_mapping[sheet_name]}] Found entity data, processing section...")
                            entity_found_in_sheet = True
                            selected_data_frame = data_frame
                            break

                # Process the selected dataframe
                if entity_found_in_sheet:
                    # Use the current column selection improvements
                    if latest_date_col and latest_date_col in selected_data_frame.columns:
                        # Find the actual description column (contains text like "Deposits with banks", "Total")
                        desc_col = None
                        for col in selected_data_frame.columns:
                            # Check if this column contains the descriptions we want
                            col_values = selected_data_frame[col].dropna().astype(str)
                            if any('deposits' in val.lower() or 'total' in val.lower() or 'banks' in val.lower()
                                   for val in col_values if val.strip()):
                                desc_col = col
                                break

                        # Fallback to column 1 if we can't find the description column
                        if desc_col is None:
                            desc_col = selected_data_frame.columns[min(1, len(selected_data_frame.columns)-1)]

                        essential_cols = [desc_col, latest_date_col]
                        filtered_df = selected_data_frame[essential_cols].dropna(how='all')
                    else:
                        # Fallback to full dataframe
                        filtered_df = selected_data_frame

                    if not filtered_df.empty:
                        # Convert to markdown format (RESTORED WORKING FORMAT)
                        try:
                            markdown_content += tabulate(filtered_df, headers='keys', tablefmt='pipe') + '\n\n'
                        except Exception:
                            markdown_content += filtered_df.to_markdown(index=False) + '\n\n'
                        print(f"✅ [{reverse_mapping[sheet_name]}] Processed section with {len(filtered_df)} rows")

                # Early exit if entity data found in this sheet
                if entity_found_in_sheet:
                    print(f"✅ [{reverse_mapping[sheet_name]}] Entity data found and processed")
        
        return markdown_content
        
    except Exception as e:
        print("An error occurred while processing the Excel file:", e)
        return ""


def get_worksheet_sections_by_keys(uploaded_file, tab_name_mapping, entity_name, entity_suffixes, entity_keywords=None, entity_mode='multiple', debug=False):
        # Processing worksheet sections

    """
    Get worksheet sections organized by financial keys with enhanced entity filtering and latest date detection.
    For single entity mode, entity filtering is skipped as there's only one entity table.
    """
    import re  # Import re inside function to avoid scope issues
    try:
        # Handle both uploaded files and default file using context manager to avoid file locks
        if hasattr(uploaded_file, 'name') and uploaded_file.name == "databook.xlsx":
            excel_source = "databook.xlsx"
        else:
            excel_source = uploaded_file

        # Create a reverse mapping from values to keys
        reverse_mapping = {}
        if tab_name_mapping is not None:
            for key, values in tab_name_mapping.items():
                for value in values:
                    reverse_mapping[value] = key
                # Also map the key name directly to itself (for sheet names like "Cash", "AR")
                reverse_mapping[key] = key

        # Get financial keys
        from fdd_utils.data_utils import get_financial_keys
        financial_keys = get_financial_keys()

        # Initialize sections by key
        sections_by_key = {key: [] for key in financial_keys}

        # Process sheets within context manager
        processed_sheets = set()  # Track processed sheets to avoid duplicates
        with pd.ExcelFile(excel_source) as xl:
            print(f"\n{'='*80}")
            print(f"📊 STARTING EXCEL PROCESSING")
            print(f"{'='*80}")

            for sheet_name in xl.sheet_names:
                # Skip sheets not in mapping to avoid using undefined df
                if sheet_name not in reverse_mapping:
                    continue

                # Skip already processed sheets to avoid duplicates
                if sheet_name in processed_sheets:
                    print(f"⏭️  Skipping already processed sheet: {sheet_name}")
                    continue

                processed_sheets.add(sheet_name)
                df = xl.parse(sheet_name)

                print(f"\n{'='*60}")
                print(f"📋 PROCESSING SHEET: {sheet_name}")
                print(f"{'='*60}")

                # Use entity_keywords passed from main app, or generate fallback
                if entity_keywords is None:
                    # Fallback: generate entity_keywords from entity_suffixes
                    entity_keywords = []
                    for suffix in entity_suffixes:
                        if suffix == entity_name:
                            entity_keywords.append(entity_name)
                        else:
                            entity_keywords.append(f"{entity_name} {suffix}")

                    if not entity_keywords:
                        entity_keywords = [entity_name]

                combined_pattern = '|'.join(re.escape(kw) for kw in entity_keywords)

                # Detect latest date column once per sheet (not per dataframe)
                latest_date_col = detect_latest_date_column(df, sheet_name, entity_keywords)

                # For multiple entity mode, we need to handle the entire sheet differently
                if entity_mode == 'multiple':
                    print(f"   🎯 MULTIPLE ENTITY MODE: Processing sheet '{sheet_name}' for SELECTED entity '{entity_name}'")
                    # Get sections ONLY for the selected entity
                    entity_sections, _ = determine_entity_mode_and_filter_all_sections(df, entity_name, entity_keywords, entity_mode)
                    
                    print(f"   📊 Found {len(entity_sections)} sections for selected entity '{entity_name}' in sheet '{sheet_name}'")

                    # Process each section for the selected entity only
                    for section_idx, data_frame in enumerate(entity_sections):
                        print(f"   🎯 Processing entity '{entity_name}' section {section_idx+1}")
                        print(f"   📊 SECTION DATA DEBUG:")
                        print(f"      📋 DataFrame shape: {data_frame.shape}")
                        print(f"      📋 First row content: {data_frame.iloc[0].fillna('').astype(str).str.cat(sep=' ')}")

                        # Print FULL table content (not just first 3 rows)
                        print(f"      📋 FULL TABLE CONTENT:")
                        for i in range(len(data_frame)):
                            row_content = data_frame.iloc[i].fillna('').astype(str).str.cat(sep=' | ')
                            print(f"         Row {i}: {row_content}")

                        # Check if this section contains any of the financial keys
                        matched_keys = []  # Track which keys this data_frame matches

                        # Get all text from the dataframe for searching
                        all_text = ' '.join(data_frame.astype(str).values.flatten()).lower()
                        # Skip duplicate processing message - already shown above
                        print(f"   🔍 Processing dataframe section: {data_frame.shape[0]}×{data_frame.shape[1]}")
                        print(f"   🔍 Looking for '人民币千元' in content...")

                        # Check for RMB patterns
                        rmb_found = False
                        if "人民币千元" in all_text:
                            print(f"   💰 FOUND: '人民币千元' detected in tab '{sheet_name}'")
                            rmb_found = True
                        elif "人民幣千元" in all_text:
                            print(f"   💰 FOUND: '人民幣千元' (Traditional) detected in tab '{sheet_name}'")
                            rmb_found = True
                        elif "cny'000" in all_text.lower():
                            print(f"   💰 FOUND: 'CNY'000' detected in tab '{sheet_name}'")
                            rmb_found = True
                        else:
                            print(f"   ❌ RMB patterns NOT found in tab '{sheet_name}'")

                        # Check each financial key - prioritize exact sheet name matches
                        for financial_key in financial_keys:
                            # First, check if the sheet name exactly matches this key
                            if sheet_name.lower() == financial_key.lower():
                                matched_keys.append(financial_key)
                                print(f"   ✅ Exact match: {sheet_name} -> {financial_key}")
                                continue

                            # Check if the sheet name matches any of the key's sheet patterns
                            # Use more restrictive matching to avoid substring conflicts
                            if financial_key in tab_name_mapping:
                                sheet_patterns = tab_name_mapping[financial_key]
                                for pattern in sheet_patterns:
                                    # Use word boundary matching to avoid substring conflicts
                                    # e.g., "AR" should not match "Share capital" which contains "AR"
                                    pattern_lower = pattern.lower()
                                    sheet_lower = sheet_name.lower()

                                    # Check for exact word match or exact pattern match
                                    if (pattern_lower == sheet_lower or
                                        pattern_lower in sheet_lower.split() or
                                        sheet_lower.startswith(pattern_lower + ' ') or
                                        sheet_lower.endswith(' ' + pattern_lower) or
                                        ' ' + pattern_lower + ' ' in ' ' + sheet_lower + ' '):
                                        matched_keys.append(financial_key)
                                        break

                            # Only use exact sheet name matching - no fallback to KEY_TERMS_BY_KEY
                            # This prevents multiple keys from matching the same sheet
                            pass

                        # Process this dataframe for each matched key
                        print(f"   🔍 Sheet '{sheet_name}' matched keys: {matched_keys}")
                        for best_key in matched_keys:
                            print(f"   🎯 Processing key '{best_key}' for tab '{sheet_name}'")
                            print(f"   🔍 Sheet '{sheet_name}' matches key '{best_key}' - processing...")
                            # Initialize actual_entity_found for this key
                            actual_entity_found = None

                            # Check for entity keywords in the dataframe
                            try:
                                # Vectorized entity matching - much faster than row-by-row
                                mask_series = data_frame.apply(
                                    lambda row: row.astype(str).str.contains(
                                        combined_pattern, case=False, regex=True, na=False
                                    ).any(),
                                    axis=1
                                )
                                entity_mask = mask_series
                            except Exception:
                                # Fallback to simpler matching
                                entity_mask = data_frame.astype(str).apply(
                                    lambda x: x.str.contains(
                                        combined_pattern, case=False, regex=True, na=False
                                    ).any()
                                )
                            # entity_mask is already defined above as mask_series

                        # Intelligent entity detection - automatically handle single vs multiple entity scenarios
                        # entity_mode parameter is now 'auto' and the logic adapts automatically
                        section_text = ' '.join(data_frame.astype(str).values.flatten()).lower()

                        # Check if entity keywords are found in the content
                        entity_found = any(entity_keyword.lower() in section_text for entity_keyword in entity_keywords)

                        if entity_found:
                            # Entity found - use normal processing
                            actual_entity_found = None
                            print(f"   🔍 Entity found in content: entity_found={entity_found}")
                            print(f"   🔍 Entity keywords: {entity_keywords}")
                        else:
                            # No entity found - use intelligent detection
                            if len(section_text.strip()) > 50:  # Has substantial content
                                # Check if this looks like a financial table (has numbers and/or Chinese characters)
                                has_numbers = any(char.isdigit() for char in section_text)
                                has_chinese = any('\u4e00' <= char <= '\u9fff' for char in section_text)

                                if has_numbers or has_chinese:
                                    # Likely valid financial content - assume correct entity
                                    entity_found = True
                                    actual_entity_found = entity_name
                                    print(f"   🔍 No entity found but valid financial content detected - assuming correct entity for {best_key}")
                                    print(f"   🔍 Content has numbers: {has_numbers}, Chinese: {has_chinese}")
                                else:
                                    print(f"   🔍 No entity found and content doesn't appear to be financial data")
                                    print(f"   🔍 Entity keywords: {entity_keywords}")
                                    print(f"   🔍 Section text sample: {section_text[:200]}...")
                            else:
                                print(f"   🔍 No entity found in minimal content")
                                print(f"   🔍 Entity keywords: {entity_keywords}")
                                print(f"   🔍 Section text sample: {section_text[:200]}...")

                        # CRITICAL FIX: Only process if this section matches the SELECTED entity
                        # Check if this section contains the selected entity name
                        print(f"      📋 Looking for entity keywords: {entity_keywords}")
                        print(f"      📋 Section text (first 200 chars): {section_text[:200]}")
                        print(f"      📋 Section text (full): {section_text}")
                        
                        section_contains_selected_entity = any(entity_keyword.lower() in section_text for entity_keyword in entity_keywords)
                        
                        print(f"      🎯 DECISION RESULT: section_contains_selected_entity = {section_contains_selected_entity}")
                        
                        # Detailed keyword matching
                        for entity_keyword in entity_keywords:
                            keyword_found = entity_keyword.lower() in section_text
                            print(f"      🔍 Keyword '{entity_keyword}' found: {keyword_found}")
                            if keyword_found:
                                # Find where the keyword appears
                                keyword_pos = section_text.find(entity_keyword.lower())
                                context_start = max(0, keyword_pos - 50)
                                context_end = min(len(section_text), keyword_pos + len(entity_keyword) + 50)
                                context = section_text[context_start:context_end]
                                print(f"         📍 Found at position {keyword_pos}, context: '{context}'")
                        
                        if section_contains_selected_entity:
                            print(f"   ✅ SECTION MATCHES SELECTED ENTITY: '{entity_name}' found in section")
                            # Find the actual entity name from the section text
                            if actual_entity_found is None:  # Not already set by intelligent detection
                                actual_entity_found = None
                                # First try to find the exact entity keyword
                                for entity_keyword in entity_keywords:
                                    if entity_keyword.lower() in section_text:
                                        actual_entity_found = entity_keyword
                                        break

                            # If still not found, try to extract the actual entity name from the data
                            if actual_entity_found is None:
                                # Look for entity patterns in the section text
                                import re
                                # Common patterns for entity names
                                entity_patterns = [
                                    r'(\w+\s+Wanpu(?:\s+Limited)?)',
                                    r'(\w+\s+Wanchen(?:\s+Limited)?)',
                                    r'(Ningbo\s+\w+(?:\s+Limited)?)',
                                    r'(Haining\s+\w+(?:\s+Limited)?)',
                                    r'(Nanjing\s+\w+(?:\s+Limited)?)'
                                ]

                                for pattern in entity_patterns:
                                    matches = re.findall(pattern, section_text, re.IGNORECASE)
                                    if matches:
                                        actual_entity_found = matches[0]
                                        print(f"   🔍 Extracted actual entity name: {actual_entity_found}")
                                        break
                        else:
                            print(f"   ❌ SECTION SKIPPED: Does not contain selected entity '{entity_name}'")
                            print(f"   🔍 Section text sample: {section_text[:100]}...")
                            print(f"   🔍 FULL SECTION TEXT: {section_text}")
                            print(f"   🔍 ENTITY KEYWORDS SEARCHED: {entity_keywords}")
                            continue  # Skip this section entirely

                        # Use new accounting table parser with detected latest date column
                        print(f"🔍 About to parse table for key '{best_key}' in sheet '{sheet_name}':")
                        print(f"   📊 DataFrame shape: {data_frame.shape}")
                        print(f"   🎯 entity_mode: {entity_mode}")
                        print(f"   👤 entity_name: '{entity_name}'")
                        print(f"   🔑 entity_keywords: {entity_keywords}")
                        print(f"   📍 actual_entity_found: {actual_entity_found}")

                        # Check if we have separated tables for this sheet
                        current_statement_type = 'BS'  # Default
                        if hasattr(st, 'session_state') and hasattr(st.session_state, 'get'):
                            current_statement_type = st.session_state.get('current_statement_type', 'BS')
                        
                        # Get separated table data if available
                        separated_tables = getattr(st.session_state, 'separated_tables', {}) if hasattr(st, 'session_state') else {}
                        
                        df_to_parse = data_frame
                        if sheet_name in separated_tables:
                            sheet_separated = separated_tables[sheet_name]
                            if current_statement_type == 'BS' and sheet_separated['balance_sheet']:
                                df_to_parse = sheet_separated['balance_sheet']['data']
                                print(f"   🎯 Using separated BALANCE SHEET data for {sheet_name}")
                            elif current_statement_type == 'IS' and sheet_separated['income_statement']:
                                df_to_parse = sheet_separated['income_statement']['data']
                                print(f"   🎯 Using separated INCOME STATEMENT data for {sheet_name}")
                        
                        # Parse the table - this will handle entity filtering internally
                        # Convert entity_mode to manual_mode for parse_accounting_table
                        manual_mode = entity_mode
                        parsed_table = parse_accounting_table(df_to_parse, best_key, entity_name, sheet_name, latest_date_col, actual_entity_found, debug, manual_mode)

                        # IMPORTANT: Extract the filtered data from parsed_table if entity filtering was applied
                        # This ensures we use the filtered data for all subsequent processing
                        if parsed_table and 'filtered_data' in parsed_table:
                            filtered_data = parsed_table['filtered_data']
                            print(f"   🔄 Using filtered data from entity filtering: {filtered_data.shape} rows")
                        else:
                            filtered_data = data_frame  # Fallback to original data
                            print(f"   📋 Using original data (no entity filtering applied): {filtered_data.shape} rows")

                        print(f"   🔍 parse_accounting_table returned: {parsed_table is not None}")
                        if parsed_table:
                            print(f"   🔍 parsed_table keys: {list(parsed_table.keys()) if isinstance(parsed_table, dict) else 'not a dict'}")

                        if parsed_table:
                            # Entity validation - use the filtered data for validation if available
                            if 'filtered_data' in parsed_table:
                                validation_data = parsed_table['filtered_data']
                                print(f"   🔍 Using filtered data for entity validation: {validation_data.shape}")
                            else:
                                validation_data = data_frame

                            if actual_entity_found is not None:
                                # Entity was already validated in the intelligent detection above
                                is_selected_entity = True
                                section_text = ' '.join(validation_data.astype(str).values.flatten()).lower()
                            else:
                                # Perform final validation check
                                section_text = ' '.join(validation_data.astype(str).values.flatten()).lower()
                                is_selected_entity = any(entity_keyword.lower() in section_text for entity_keyword in entity_keywords)

                            # Debug: Print which entity was found
                            print(f"   🔍 Section for {best_key}: entity_found={entity_found}, is_selected_entity={is_selected_entity}")
                            print(f"   🔍 Looking for: {entity_keywords}")
                            print(f"   🔍 Found entity: {actual_entity_found}")
                            print(f"   🔍 Section text sample: {section_text[:100]}...")

                            # VALIDATION: Check for content mismatch (e.g., AR key showing taxes content)
                            if best_key == 'AR':
                                # Check if this section contains taxes content
                                if 'tax' in section_text or 'surcharge' in section_text:
                                    print(f"⚠️ WARNING: AR key matched to sheet '{sheet_name}' but contains taxes content!")
                                    print(f"   Section text sample: {section_text[:200]}...")
                                    # Skip this section to avoid incorrect mapping
                                    continue

                            # Use filtered data if available, otherwise use original
                            display_data = parsed_table.get('filtered_data', data_frame)

                            section_data = {
                                'sheet': sheet_name,
                                'sheet_name': sheet_name,  # Add sheet_name for compatibility
                                'entity_name': actual_entity_found or entity_name,  # Use detected entity or fallback to selected entity
                                'data': display_data,  # Use filtered data if available
                                'parsed_data': parsed_table,
                                'markdown': create_improved_table_markdown(parsed_table),
                                'entity_match': True,
                                'is_selected_entity': is_selected_entity,
                                'detected_entity': actual_entity_found  # Store the detected entity separately
                            }
                            
                            # Only add this section if it's the correct sheet for this key AND matches the selected entity
                            # This prevents wrong sheets from being assigned to keys
                            if (sheet_name.lower() == best_key.lower() or
                                (best_key in tab_name_mapping and
                                 any(pattern.lower() in sheet_name.lower() for pattern in tab_name_mapping[best_key]))):

                                # Only add sections that match the selected entity
                                if entity_found and is_selected_entity:
                                    sections_by_key[best_key].append(section_data)
                                    print(f"   ✅ SUCCESS: Added section for '{best_key}' from tab '{sheet_name}'")
                                    print(f"   📊 Section contains {len(parsed_table.get('data', []))} data rows")
                                    print(f"   💰 RMB detected: {'YES' if rmb_found else 'NO'}")
                                    print(f"   🔍 Entity found: {actual_entity_found}")
                                    print(f"   📊 Total sections for {best_key}: {len(sections_by_key[best_key])}")
                                else:
                                    print(f"   ❌ SKIPPED: Section for '{best_key}' from tab '{sheet_name}' (entity not matched)")
                                    print(f"   🔍 entity_found: {entity_found}, is_selected_entity: {is_selected_entity}")
                                    print(f"   🔍 actual_entity_found: {actual_entity_found}")
                                    print(f"   🔑 entity_keywords: {entity_keywords}")
                        else:
                            # Fallback to original format if parsing fails
                            print(f"   ⚠️  parse_accounting_table failed for {best_key}, using fallback")
                            try:
                                markdown_content = tabulate(data_frame, headers='keys', tablefmt='pipe') + '\n\n'
                            except Exception:
                                markdown_content = data_frame.to_markdown(index=False) + '\n\n'
                            
                            sections_by_key[best_key].append({
                                'sheet': sheet_name,
                                'sheet_name': sheet_name,  # Add sheet_name for compatibility
                                'entity_name': entity_name,  # Fallback case - use selected entity name
                                'data': data_frame,
                                'markdown': markdown_content,
                                'entity_match': True,
                                'is_selected_entity': True  # Force this to True for fallback
                            })
                else:
                    # SINGLE ENTITY MODE: Use the original approach with empty row splitting
                    print(f"   🎯 SINGLE ENTITY MODE: Using traditional section splitting for sheet '{sheet_name}'")
                    empty_rows = df.index[df.isnull().all(1)]
                    start_idx = 0
                    dataframes = []
                    for end_idx in empty_rows:
                        if end_idx > start_idx:
                            split_df = df[start_idx:end_idx]
                            if not split_df.dropna(how='all').empty:
                                dataframes.append(split_df)
                            start_idx = end_idx + 1
                    if start_idx < len(df):
                        dataframes.append(df[start_idx:])

                    # Process each split dataframe (original logic)
                    for data_frame in dataframes:
                        # Check if this section contains any of the financial keys
                        matched_keys = []  # Track which keys this data_frame matches

                        # Get all text from the dataframe for searching
                        all_text = ' '.join(data_frame.astype(str).values.flatten()).lower()

                        # Skip duplicate processing message - already shown above
                        print(f"   🔍 Processing dataframe section: {data_frame.shape[0]}×{data_frame.shape[1]}")

                        # Check for RMB patterns
                        rmb_found = False
                        if "人民币千元" in all_text:
                            print(f"   💰 FOUND: '人民币千元' detected in tab '{sheet_name}'")
                            rmb_found = True
                        elif "人民幣千元" in all_text:
                            print(f"   💰 FOUND: '人民幣千元' (Traditional) detected in tab '{sheet_name}'")
                            rmb_found = True
                        elif "cny'000" in all_text.lower():
                            print(f"   💰 FOUND: 'CNY'000' detected in tab '{sheet_name}'")
                            rmb_found = True
                        else:
                            print(f"   ❌ RMB patterns NOT found in tab '{sheet_name}'")

                        # Check each financial key - prioritize exact sheet name matches
                        for financial_key in financial_keys:
                            # First, check if the sheet name exactly matches this key
                            if sheet_name.lower() == financial_key.lower():
                                matched_keys.append(financial_key)
                                print(f"   ✅ Exact match: {sheet_name} -> {financial_key}")
                                continue

                            # Check if the sheet name matches any of the key's sheet patterns
                            # Use more restrictive matching to avoid substring conflicts
                            if financial_key in tab_name_mapping:
                                sheet_patterns = tab_name_mapping[financial_key]
                                for pattern in sheet_patterns:
                                    # Use word boundary matching to avoid substring conflicts
                                    # e.g., "AR" should not match "Share capital" which contains "AR"
                                    pattern_lower = pattern.lower()
                                    sheet_lower = sheet_name.lower()

                                    # Check for exact word match or exact pattern match
                                    if (pattern_lower == sheet_lower or
                                        pattern_lower in sheet_lower.split() or
                                        sheet_lower.startswith(pattern_lower + ' ') or
                                        sheet_lower.endswith(' ' + pattern_lower) or
                                        ' ' + pattern_lower + ' ' in ' ' + sheet_lower + ' '):
                                        matched_keys.append(financial_key)
                                        break

                            # Only use exact sheet name matching - no fallback to KEY_TERMS_BY_KEY
                            # This prevents multiple keys from matching the same sheet
                            pass

                        # Process this dataframe for each matched key
                        print(f"   🔍 Sheet '{sheet_name}' matched keys: {matched_keys}")
                        for best_key in matched_keys:
                            print(f"   🎯 Processing key '{best_key}' for tab '{sheet_name}'")
                            print(f"   🔍 Sheet '{sheet_name}' matches key '{best_key}' - processing...")
                            # Initialize actual_entity_found for this key
                            actual_entity_found = None

                            # Check for entity keywords in the dataframe
                            try:
                                # Vectorized entity matching - much faster than row-by-row
                                mask_series = data_frame.apply(
                                    lambda row: row.astype(str).str.contains(
                                        combined_pattern, case=False, regex=True, na=False
                                    ).any(),
                                    axis=1
                                )
                                entity_mask = mask_series
                            except Exception:
                                # Fallback to simpler matching
                                entity_mask = data_frame.astype(str).apply(
                                    lambda x: x.str.contains(
                                        combined_pattern, case=False, regex=True, na=False
                                    ).any()
                                )

                            # Intelligent entity detection - automatically handle single vs multiple entity scenarios
                            section_text = ' '.join(data_frame.astype(str).values.flatten()).lower()

                            # Check if entity keywords are found in the content
                            entity_found = any(entity_keyword.lower() in section_text for entity_keyword in entity_keywords)

                            if entity_found:
                                # Entity found - use normal processing
                                actual_entity_found = None
                                print(f"   🔍 Entity found in content: entity_found={entity_found}")
                                print(f"   🔍 Entity keywords: {entity_keywords}")
                            else:
                                # No entity found - use intelligent detection
                                if len(section_text.strip()) > 50:  # Has substantial content
                                    # Check if this looks like a financial table (has numbers and/or Chinese characters)
                                    has_numbers = any(char.isdigit() for char in section_text)
                                    has_chinese = any('\u4e00' <= char <= '\u9fff' for char in section_text)

                                    if has_numbers or has_chinese:
                                        # Likely valid financial content - assume correct entity
                                        entity_found = True
                                        actual_entity_found = entity_name
                                        print(f"   🔍 No entity found but valid financial content detected - assuming correct entity for {best_key}")
                                        print(f"   🔍 Content has numbers: {has_numbers}, Chinese: {has_chinese}")
                                    else:
                                        print(f"   🔍 No entity found and content doesn't appear to be financial data")
                                        print(f"   🔍 Entity keywords: {entity_keywords}")
                                        print(f"   🔍 Section text sample: {section_text[:200]}...")
                                else:
                                    print(f"   🔍 No entity found in minimal content")
                                    print(f"   🔍 Entity keywords: {entity_keywords}")
                                    print(f"   🔍 Section text sample: {section_text[:200]}...")

                            # Only process if entity is found in this section
                            if entity_found:
                                # Find the actual entity name from the section text
                                if actual_entity_found is None:  # Not already set by intelligent detection
                                    actual_entity_found = None
                                    # First try to find the exact entity keyword
                                    for entity_keyword in entity_keywords:
                                        if entity_keyword.lower() in section_text:
                                            actual_entity_found = entity_keyword
                                            break

                                # If still not found, try to extract the actual entity name from the data
                                if actual_entity_found is None:
                                    # Look for entity patterns in the section text
                                    import re
                                    # Common patterns for entity names
                                    entity_patterns = [
                                        r'(\w+\s+Wanpu(?:\s+Limited)?)',
                                        r'(\w+\s+Wanchen(?:\s+Limited)?)',
                                        r'(Ningbo\s+\w+(?:\s+Limited)?)',
                                        r'(Haining\s+\w+(?:\s+Limited)?)',
                                        r'(Nanjing\s+\w+(?:\s+Limited)?)'
                                    ]

                                    for pattern in entity_patterns:
                                        matches = re.findall(pattern, section_text, re.IGNORECASE)
                                        if matches:
                                            actual_entity_found = matches[0]
                                            print(f"   🔍 Extracted actual entity name: {actual_entity_found}")
                                            break

                                # Use new accounting table parser with detected latest date column
                                print(f"🔍 About to parse table for key '{best_key}' in sheet '{sheet_name}':")
                                print(f"   📊 DataFrame shape: {data_frame.shape}")
                                print(f"   🎯 entity_mode: {entity_mode}")
                                print(f"   👤 entity_name: '{entity_name}'")
                                print(f"   🔑 entity_keywords: {entity_keywords}")
                                print(f"   📍 actual_entity_found: {actual_entity_found}")

                                # Check if we have separated tables for this sheet
                                current_statement_type = 'BS'  # Default
                                if hasattr(st, 'session_state') and hasattr(st.session_state, 'get'):
                                    current_statement_type = st.session_state.get('current_statement_type', 'BS')
                                
                                # Get separated table data if available
                                separated_tables = getattr(st.session_state, 'separated_tables', {}) if hasattr(st, 'session_state') else {}
                                
                                df_to_parse = data_frame
                                if sheet_name in separated_tables:
                                    sheet_separated = separated_tables[sheet_name]
                                    if current_statement_type == 'BS' and sheet_separated['balance_sheet']:
                                        df_to_parse = sheet_separated['balance_sheet']['data']
                                        print(f"   🎯 Using separated BALANCE SHEET data for {sheet_name}")
                                    elif current_statement_type == 'IS' and sheet_separated['income_statement']:
                                        df_to_parse = sheet_separated['income_statement']['data']
                                        print(f"   🎯 Using separated INCOME STATEMENT data for {sheet_name}")
                                
                                # Parse the table - this will handle entity filtering internally
                                manual_mode = entity_mode
                                parsed_table = parse_accounting_table(df_to_parse, best_key, entity_name, sheet_name, latest_date_col, actual_entity_found, debug, manual_mode)

                                # Extract the filtered data from parsed_table if entity filtering was applied
                                if parsed_table and 'filtered_data' in parsed_table:
                                    filtered_data = parsed_table['filtered_data']
                                    print(f"   🔄 Using filtered data from entity filtering: {filtered_data.shape} rows")
                                else:
                                    filtered_data = data_frame  # Fallback to original data
                                    print(f"   📋 Using original data (no entity filtering applied): {filtered_data.shape} rows")

                                print(f"   🔍 parse_accounting_table returned: {parsed_table is not None}")
                                if parsed_table:
                                    print(f"   🔍 parsed_table keys: {list(parsed_table.keys()) if isinstance(parsed_table, dict) else 'not a dict'}")

                                if parsed_table:
                                    # Entity validation - use the filtered data for validation if available
                                    if 'filtered_data' in parsed_table:
                                        validation_data = parsed_table['filtered_data']
                                        print(f"   🔍 Using filtered data for entity validation: {validation_data.shape}")
                                    else:
                                        validation_data = data_frame

                                    if actual_entity_found is not None:
                                        # Entity was already validated in the intelligent detection above
                                        is_selected_entity = True
                                        section_text = ' '.join(validation_data.astype(str).values.flatten()).lower()
                                    else:
                                        # Perform final validation check
                                        section_text = ' '.join(validation_data.astype(str).values.flatten()).lower()
                                        is_selected_entity = any(entity_keyword.lower() in section_text for entity_keyword in entity_keywords)

                                    # Debug: Print which entity was found
                                    print(f"   🔍 Section for {best_key}: entity_found={entity_found}, is_selected_entity={is_selected_entity}")
                                    print(f"   🔍 Looking for: {entity_keywords}")
                                    print(f"   🔍 Found entity: {actual_entity_found}")
                                    print(f"   🔍 Section text sample: {section_text[:100]}...")

                                    # VALIDATION: Check for content mismatch (e.g., AR key showing taxes content)
                                    if best_key == 'AR':
                                        # Check if this section contains taxes content
                                        if 'tax' in section_text or 'surcharge' in section_text:
                                            print(f"⚠️ WARNING: AR key matched to sheet '{sheet_name}' but contains taxes content!")
                                            print(f"   Section text sample: {section_text[:200]}...")
                                            # Skip this section to avoid incorrect mapping
                                            continue

                                    # Use filtered data if available, otherwise use original
                                    display_data = parsed_table.get('filtered_data', data_frame)

                                    section_data = {
                                        'sheet': sheet_name,
                                        'sheet_name': sheet_name,  # Add sheet_name for compatibility
                                        'entity_name': actual_entity_found or entity_name,  # Use detected entity or fallback to selected entity
                                        'data': display_data,
                                        'parsed_data': parsed_table,
                                        'markdown': create_improved_table_markdown(parsed_table),
                                        'entity_match': True,
                                        'is_selected_entity': is_selected_entity,
                                        'detected_entity': actual_entity_found  # Store the detected entity separately
                                    }

                                    # Only add this section if it's the correct sheet for this key AND matches the selected entity
                                    if (sheet_name.lower() == best_key.lower() or
                                        (best_key in tab_name_mapping and
                                         any(pattern.lower() in sheet_name.lower() for pattern in tab_name_mapping[best_key]))):

                                        # Only add sections that match the selected entity
                                        if entity_found and is_selected_entity:
                                            sections_by_key[best_key].append(section_data)
                                            print(f"   ✅ SUCCESS: Added section for '{best_key}' from tab '{sheet_name}'")
                                            print(f"   📊 Section contains {len(parsed_table.get('data', []))} data rows")
                                            print(f"   💰 RMB detected: {'YES' if rmb_found else 'NO'}")
                                            print(f"   🔍 Entity found: {actual_entity_found}")
                                            print(f"   📊 Total sections for {best_key}: {len(sections_by_key[best_key])}")
                                        else:
                                            print(f"   ❌ SKIPPED: Section for '{best_key}' from tab '{sheet_name}' (entity not matched)")
                                            print(f"   🔍 entity_found: {entity_found}, is_selected_entity: {is_selected_entity}")
                                            print(f"   🔍 actual_entity_found: {actual_entity_found}")
                                            print(f"   🔑 entity_keywords: {entity_keywords}")
                                else:
                                    # Fallback to original format if parsing fails
                                    print(f"   ⚠️  parse_accounting_table failed for {best_key}, using fallback")
                                    try:
                                        markdown_content = tabulate(data_frame, headers='keys', tablefmt='pipe') + '\n\n'
                                    except Exception:
                                        markdown_content = data_frame.to_markdown(index=False) + '\n\n'

                                    sections_by_key[best_key].append({
                                        'sheet': sheet_name,
                                        'sheet_name': sheet_name,  # Add sheet_name for compatibility
                                        'entity_name': entity_name,  # Fallback case - use selected entity name
                                        'data': data_frame,
                                        'markdown': markdown_content,
                                        'entity_match': True,
                                        'is_selected_entity': True  # Force this to True for fallback
                                    })

        # Print summary of processed sections
        total_sections = sum(len(sections) for sections in sections_by_key.values())
        print(f"\n🎉 EXCEL PROCESSING COMPLETE!")
        print(f"   📊 SUMMARY: Processed {len(sections_by_key)} financial keys with {total_sections} total sections")

        # Detailed summary
        print(f"   📋 DETAILED RESULTS:")
        keys_with_data = 0
        for key, sections in sections_by_key.items():
            if sections:
                keys_with_data += 1
                print(f"   ✅ {key}: {len(sections)} sections found")
            else:
                print(f"   ❌ {key}: No sections found")

        print(f"   🎯 SUCCESS: {keys_with_data} out of {len(sections_by_key)} financial keys have data")
        print(f"   💡 TIP: If '人民币千元' was not detected, check if your Excel file contains this exact text or similar currency indicators")
        print(f"\n" + "="*80)
        print(f"   🎉 EXCEL TAB PROCESSING FINISHED!")
        print(f"   📊 You can now view the processed data in the application tabs above.")
        print(f"   💰 Check each financial key tab to see the extracted data.")
        print(f"="*80)

        return sections_by_key

    except Exception as e:
        print(f"Error in get_worksheet_sections_by_keys: {e}")
        return {}


def create_improved_table_markdown(parsed_table):
    """Create improved markdown representation of parsed accounting table"""
    try:
        if not parsed_table or 'metadata' not in parsed_table or 'data' not in parsed_table:
            return "No structured data available"
        
        metadata = parsed_table['metadata']
        data_rows = parsed_table['data']
        
        markdown_lines = []
        
        # Table header with metadata
        markdown_lines.append(f"**{metadata['table_name']}**")
        markdown_lines.append(f"*Sheet: {metadata['sheet_name']}*")
        
        # Format date if present
        if metadata.get('date'):
            from fdd_utils.data_utils import format_date_to_dd_mmm_yyyy
            formatted_date = format_date_to_dd_mmm_yyyy(metadata['date'])
            markdown_lines.append(f"*Date: {formatted_date}*")
        
        if metadata.get('currency_info'):
            markdown_lines.append(f"*Currency: {metadata['currency_info']}*")
        
        if metadata.get('value_column'):
            markdown_lines.append(f"*Value Column: {metadata['value_column']}*")
        
        markdown_lines.append("")  # Empty line
        
        # Create table with MULTIPLE description columns if available
        # Check if we have multiple description columns
        num_desc_cols = data_rows[0].get('num_desc_columns', 1) if data_rows else 1
        
        if num_desc_cols > 1:
            # Multiple description columns - create 3-column table
            markdown_lines.append("| Category | Subcategory | Value |")
            markdown_lines.append("|----------|-------------|-------|")
            
            for row in data_rows:
                desc0 = row.get('description_0', '')
                desc1 = row.get('description_1', '')
                value = row['value']
                is_total = row.get('is_total', False)
                
                # Format value with commas (1 decimal place for financial figures)
                if isinstance(value, (int, float)):
                    formatted_value = f"{value:,.1f}"
                elif isinstance(value, str) and (re.search(r'\d{4}年\d{1,2}月\d{1,2}日', value) or \
                                              re.search(r'\d{4}/\d{1,2}/\d{1,2}', value) or \
                                              re.search(r'\d{1,2}/\d{1,2}/\d{4}', value)):
                    # This is a date, preserve original format
                    formatted_value = value
                else:
                    formatted_value = str(value)
                
                # Bold total rows
                if is_total:
                    markdown_lines.append(f"| **{desc0}** | **{desc1}** | **{formatted_value}** |")
                else:
                    markdown_lines.append(f"| {desc0} | {desc1} | {formatted_value} |")
        else:
            # Single description column - create 2-column table
            markdown_lines.append("| Description | Value |")
            markdown_lines.append("|-------------|-------|")
            
            for row in data_rows:
                desc = row['description']
                value = row['value']
                is_total = row.get('is_total', False)
                
                # Format value with commas (1 decimal place for financial figures)
                if isinstance(value, (int, float)):
                    formatted_value = f"{value:,.1f}"
                elif isinstance(value, str) and (re.search(r'\d{4}年\d{1,2}月\d{1,2}日', value) or \
                                              re.search(r'\d{4}/\d{1,2}/\d{1,2}', value) or \
                                              re.search(r'\d{1,2}/\d{1,2}/\d{4}', value)):
                    # This is a date, preserve original format
                    formatted_value = value
                else:
                    formatted_value = str(value)
                
                # Bold total rows
                if is_total:
                    markdown_lines.append(f"| **{desc}** | **{formatted_value}** |")
                else:
                    markdown_lines.append(f"| {desc} | {formatted_value} |")
        
        return "\n".join(markdown_lines)
        
    except Exception as e:
        return f"Error creating table markdown: {e}"

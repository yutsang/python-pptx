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
    print(f"🔍 [{timestamp}] DETECT_LATEST_DATE_COLUMN called for sheet '{sheet_name}' with {len(df)} rows, {len(df.columns)} columns")
    print(f"🔍 [{timestamp}] Entity keywords: {entity_keywords}")

    def parse_date(date_str):
        """Parse date string in various formats including xMxx."""
        if not date_str or pd.isna(date_str):
            return None
        
        date_str = str(date_str).strip()
        
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
        print(f"   📊 Found {len(indicative_positions)} instances of 'Indicative adjusted' (English/Chinese)")
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
                print(f"   🔍 Extended search: Looking for dates in entire sheet...")
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
        # Extract column number from column name (e.g., "Unnamed: 20" -> 20)
        if latest_column and latest_column.startswith('Unnamed: '):
            try:
                actual_col_idx = int(latest_column.split(': ')[1])
                col_letter = col_num_to_letter(actual_col_idx)
            except (ValueError, IndexError):
                col_letter = 'unknown'
        else:
            col_letter = col_num_to_letter(col_idx) if col_idx >= 0 else 'unknown'
        print(f"   🎯 [{timestamp}] FINAL SELECTION: Column '{latest_column}' ({col_letter}) with date {latest_date.strftime('%Y-%m-%d')}")
        print(f"   ✅ Column {col_letter} (list index {col_idx}) selected as latest date column")
    else:
        from datetime import datetime
        timestamp = datetime.now().strftime('%H:%M:%S.%f')[:-3]
        print(f"   ❌ [{timestamp}] No date column detected")
    
    return latest_column


def determine_entity_mode_and_filter(df, entity_name, entity_keywords):
    """Determine if we're in single entity or multiple entity mode and filter accordingly."""
    print(f"   🔍 ENTITY ANALYSIS: entity_name='{entity_name}', entity_keywords={entity_keywords}")

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

    for section_idx, (start_row, end_row) in enumerate(table_sections):
        section_text = ' '.join(
            str(val).lower() for val in df.iloc[start_row:end_row+1].values.flatten()
            if pd.notna(val) and str(val).strip()
        )

        # Check which entities are mentioned in this section
        for keyword in entity_keywords:
            if keyword.lower() in section_text:
                if keyword not in entity_sections:
                    entity_sections[keyword] = []
                entity_sections[keyword].append((start_row, end_row, section_idx))
                print(f"   📍 Section {section_idx} (rows {start_row}-{end_row}) contains entity '{keyword}'")

    # Step 3: Determine entity mode
    unique_entities = len(entity_sections)
    is_multiple_entity = unique_entities > 1
    print(f"   🎯 ENTITY MODE: {'MULTIPLE' if is_multiple_entity else 'SINGLE'} (found {unique_entities} unique entities)")

    # Step 4: Find the best matching entity section
    target_entity = None
    best_section = None

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

    # Priority 3: If no entity found but we have sections, use the first substantial section
    elif table_sections:
        # Look for the first section with substantial data (more than just headers)
        for start_row, end_row, section_idx in [(s[0], s[1], i) for i, s in enumerate(table_sections)]:
            section_size = end_row - start_row + 1
            if section_size > 3:  # At least 4 rows (header + 3 data rows)
                best_section = (start_row, end_row, section_idx)
                print(f"   ⚠️  NO ENTITY MATCH: Using first substantial section {section_idx} (rows {start_row}-{end_row})")
                break

    # Step 5: Extract the selected table section
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

def find_indicative_adjusted_column_and_dates(df, entity_keywords):
    """Find 'Indicative adjusted' column and extract dates according to new logic."""

    def parse_date(date_str):
        """Parse date string in various formats including xMxx."""
        if not date_str or pd.isna(date_str):
            return None

        date_str = str(date_str).strip()

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

    if latest_date:
        print(f"   🎯 LATEST DATE FOUND: {latest_date.strftime('%Y-%m-%d')} in column '{latest_col}' (Excel row: {latest_row_number})")
        return latest_date.strftime('%Y-%m-%d'), latest_col, latest_row_number
    else:
        print(f"   ❌ No valid dates found in the date row")
        return None, None, None

def parse_accounting_table(df, key, entity_name, sheet_name, latest_date_col=None, actual_entity=None, debug=False):
    """
    Parse accounting table with proper header detection and figure column identification
    Returns structured table data with metadata
    """
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

        df_filtered, is_multiple_entity = determine_entity_mode_and_filter(df, entity_name, entity_keywords)

        # NEW LOGIC: Step 2 - Find Indicative adjusted column and dates
        extracted_date, selected_column, row_number = find_indicative_adjusted_column_and_dates(df_filtered, entity_keywords)

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

        # Debug: Print what we're looking for
        print(f"   🔍 DEBUG: Looking for currency indicators in first 3 rows of sheet '{sheet_name}'")

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
        
        # Find data start row (first row with actual numeric data)
        data_start_row = None
        for i in range(len(df_str)):
            cell_value = str(df_str.iloc[i, value_col_idx])
            # Look for cells that contain numbers (more flexible)
            if re.search(r'\d+', cell_value) and cell_value.strip() not in ['nan', '']:
                # Check if this looks like a data row (has both description and value)
                desc_cell = str(df_str.iloc[i, 0]).strip()
                if desc_cell and desc_cell.lower() not in ['nan', '']:
                    # Simplified check: just make sure it's not a pure number and not a header
                    if (not re.match(r'^\d+\.?\d*$', desc_cell) and 
                        desc_cell.lower() not in ['cny\'000', 'audited', 'mgt acc', 'indicative adjusted', 'indicative adjustment']):
                        data_start_row = i
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

        # NEW LOGIC: Generate dynamic table name based on entity mode and selected column
        if is_multiple_entity and selected_column and row_number:
            # Multiple entity mode with Indicative adjusted column found
            table_name = f"{key} - {display_entity} (Row {row_number})"
        elif selected_column and row_number:
            # Single entity mode with Indicative adjusted column found
            table_name = f"{key} - {display_entity} (Row {row_number})"
        else:
            # Fallback to original logic
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
        
        # Extract data rows
        data_rows = []
        for i in range(data_start_row, len(df_str)):
            # Get description from column 1 (index 1) which contains the actual descriptions
            # Based on the Excel structure, descriptions are in column 1, not column 0
            desc_col_idx = 1
            desc_cell = str(df_str.iloc[i, desc_col_idx]).strip()
            value_cell = str(df_str.iloc[i, value_col_idx]).strip()
            
            # Debug: Print what we're extracting
            # print(f"   🔍 DEBUG: Extracting from row {i}: desc_col={desc_col_idx}='{desc_cell}', value_col={value_col_idx}='{value_cell}'")
            
            # Skip empty rows
            if desc_cell.lower() in ['nan', ''] and value_cell.lower() in ['nan', '']:
                continue
            
            # Skip header rows
            if desc_cell.lower() in ['description', 'item', 'account', 'details']:
                continue
            
            # Process value
            if value_cell and value_cell.lower() not in ['nan', '']:
                try:
                    # Clean and convert numeric value
                    clean_value = re.sub(r'[^\d.-]', '', value_cell)
                    if clean_value:
                        numeric_value = float(clean_value)
                        # Apply multiplier
                        final_value = numeric_value * multiplier
                        
                        # Determine if this is a total row
                        is_total = 'total' in desc_cell.lower()
                        
                        data_rows.append({
                            'description': desc_cell,
                            'value': final_value,
                            'original_value': value_cell,
                            'is_total': is_total
                        })
                except (ValueError, TypeError):
                    # Skip rows with non-numeric values
                    continue
        
        if not data_rows:
            return None
        
        return {
            'metadata': table_metadata,
            'data': data_rows,
            'raw_df': df_clean
        }
        
    except Exception as e:
        print(f"Error parsing accounting table: {e}")
        return None


def test_new_table_logic():
    """Test function for the new table logic implementation."""
    import pandas as pd
    from datetime import datetime

    print("🧪 TESTING NEW TABLE LOGIC IMPLEMENTATION")
    print("=" * 60)

    # Test Case 1: Single Entity Table
    print("\n📋 TEST CASE 1: Single Entity Table")
    print("-" * 40)

    single_entity_data = {
        'Description': ['示意性调整后', 'Revenue', 'Cost', 'Profit'],
        'Unnamed: 20': ['', '2023-12-31', 'Amount', 'Amount'],
        'Unnamed: 21': ['', '2023-11-30', 'Amount', 'Amount'],
        'Unnamed: 22': ['', '2023-10-31', 'Amount', 'Amount'],
        'Unnamed: 23': ['', '2023-09-30', 'Amount', 'Amount']
    }

    df_single = pd.DataFrame(single_entity_data)
    print(f"📊 Single Entity DataFrame: {len(df_single)} rows, {len(df_single.columns)} columns")

    entity_keywords = ['Entity A']
    df_filtered, is_multiple_entity = determine_entity_mode_and_filter(df_single, 'Entity A', entity_keywords)
    print(f"🎯 Entity Mode: {'MULTIPLE' if is_multiple_entity else 'SINGLE'}")

    # Test Case 2: Multiple Entity Tables
    print("\n📋 TEST CASE 2: Multiple Entity Tables")
    print("-" * 40)

    # Create a worksheet with multiple entity tables separated by truly empty rows
    multi_entity_data = [
        ['Entity A Financial Data', '', '', '', ''],  # Entity A table header
        ['示意性调整后', '', '', '', ''],  # Indicative adjusted header
        ['', '2023-12-31', '2023-11-30', '2023-10-31', '2023-09-30'],  # Date row
        ['Revenue', 'Amount', 'Amount', 'Amount', 'Amount'],  # Data rows
        ['Cost', 'Amount', 'Amount', 'Amount', 'Amount'],
        ['Profit', 'Amount', 'Amount', 'Amount', 'Amount'],
        [None, None, None, None, None],  # Truly empty row separator
        [None, None, None, None, None],  # Another empty row
        ['Entity B Financial Data', '', '', '', ''],  # Entity B table header
        ['示意性调整后', '', '', '', ''],  # Indicative adjusted header
        ['', '2024-01-31', '2024-02-29', '2024-03-31', '2024-04-30'],  # Date row
        ['Revenue', 'Amount', 'Amount', 'Amount', 'Amount'],  # Data rows
        ['Cost', 'Amount', 'Amount', 'Amount', 'Amount'],
        ['Profit', 'Amount', 'Amount', 'Amount', 'Amount'],
        [None, None, None, None, None],  # Truly empty row separator
        [None, None, None, None, None],  # Another empty row
        ['Entity C Financial Data', '', '', '', ''],  # Entity C table header
        ['示意性调整后', '', '', '', ''],  # Indicative adjusted header
        ['', '2024-02-28', '2024-03-31', '2024-04-30', '2024-05-31'],  # Date row
        ['Revenue', 'Amount', 'Amount', 'Amount', 'Amount'],  # Data rows
        ['Cost', 'Amount', 'Amount', 'Amount', 'Amount'],
        ['Profit', 'Amount', 'Amount', 'Amount', 'Amount']
    ]

    df_multi = pd.DataFrame(multi_entity_data)
    print(f"📊 Multiple Entity DataFrame: {len(df_multi)} rows, {len(df_multi.columns)} columns")

    # Test selecting Entity B table
    entity_keywords_multi = ['Entity A', 'Entity B', 'Entity C']
    df_filtered_multi, is_multiple_entity_multi = determine_entity_mode_and_filter(df_multi, 'Entity B', entity_keywords_multi)
    print(f"🎯 Entity Mode: {'MULTIPLE' if is_multiple_entity_multi else 'SINGLE'}")
    print(f"📊 Filtered DataFrame: {len(df_filtered_multi)} rows, {len(df_filtered_multi.columns)} columns")

    # Test Indicative adjusted column detection on filtered table
    if len(df_filtered_multi) > 0:
        extracted_date, selected_column, row_number = find_indicative_adjusted_column_and_dates(df_filtered_multi, entity_keywords_multi)
        print(f"📅 Date Extraction: {extracted_date}")
        print(f"📋 Selected Column: {selected_column}")
        print(f"📍 Excel Row Number: {row_number}")

        # Test table name generation
        display_entity = 'Entity B'
        if is_multiple_entity_multi and selected_column and row_number:
            table_name = f"BS - {display_entity} (Row {row_number})"
        elif selected_column and row_number:
            table_name = f"BS - {display_entity} (Row {row_number})"
        else:
            table_name = f"BS - {display_entity}"

        print(f"📋 Generated Table Name: {table_name}")

    print("\n" + "=" * 60)
    print("✅ MULTIPLE ENTITY TEST COMPLETED")


if __name__ == "__main__":
    test_new_table_logic()


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
                print(f"\n📊 Processing sheet: {sheet_name}")
                # Create entity keywords for this function call
                sheet_entity_keywords = [f"{entity_name} {suffix}" for suffix in entity_suffixes if suffix]
                if not sheet_entity_keywords:
                    sheet_entity_keywords = [entity_name]
                latest_date_col = detect_latest_date_column(df, sheet_name, sheet_entity_keywords)
                if latest_date_col:
                    print(f"✅ Sheet {sheet_name}: Selected column {latest_date_col}")
                else:
                    print(f"⚠️  Sheet {sheet_name}: No date column detected")
                
                # Performance optimization: Skip splitting for small sheets
                if len(df) > 100:  # Only split very large sheets
                    # Split dataframes on empty rows
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
                else:
                    # Process small sheets as single dataframe
                    print(f"📈 [{reverse_mapping[sheet_name]}] Processing as single dataframe (size: {len(df)})")
                    dataframes = [df]
                
                # Vectorized entity matching for better performance
                entity_keywords = [f"{entity_name} {suffix}" for suffix in entity_suffixes if suffix]
                if not entity_keywords:  # If no helpers, just use entity name
                    entity_keywords = [entity_name]
                
                combined_pattern = '|'.join(re.escape(kw) for kw in entity_keywords)
                
                # Track if we found entity data in this sheet
                entity_found_in_sheet = False
                
                for data_frame in dataframes:
                    # Fast filtering: only keep essential columns
                    if latest_date_col and latest_date_col in data_frame.columns:
                        # Find description column efficiently
                        desc_col = data_frame.columns[0]  # Usually first column
                        
                        # Keep only 2 columns for faster processing
                        essential_cols = [desc_col, latest_date_col]
                        filtered_df = data_frame[essential_cols].dropna(how='all')
                        
                        if not filtered_df.empty:
                            # Vectorized entity matching
                            all_text = ' '.join(filtered_df.astype(str).values.flatten()).lower()
                            entity_match = any(kw.lower() in all_text for kw in entity_keywords)
                            
                            if entity_match:
                                print(f"🚀 [{reverse_mapping[sheet_name]}] Found entity data for {reverse_mapping[sheet_name]}, continuing...")
                                entity_found_in_sheet = True
                                
                                # Convert to markdown format
                                try:
                                    markdown_content += tabulate(filtered_df, headers='keys', tablefmt='pipe') + '\n\n'
                                except Exception:
                                    markdown_content += filtered_df.to_markdown(index=False) + '\n\n'
                
                # Early exit if entity data found in this sheet
                if entity_found_in_sheet:
                    print(f"✅ [{reverse_mapping[sheet_name]}] Entity data found and processed")
        
        return markdown_content
        
    except Exception as e:
        print("An error occurred while processing the Excel file:", e)
        return ""


def get_worksheet_sections_by_keys(uploaded_file, tab_name_mapping, entity_name, entity_suffixes, entity_keywords=None, entity_mode='multiple', debug=False):
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
        with pd.ExcelFile(excel_source) as xl:
            for sheet_name in xl.sheet_names:
                # Skip sheets not in mapping to avoid using undefined df
                if sheet_name not in reverse_mapping:
                    continue
                df = xl.parse(sheet_name)

                # Split dataframes on empty rows
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
                
                # Organize sections by key - make it less restrictive
                for data_frame in dataframes:

                    # Check if this section contains any of the financial keys
                    matched_keys = []  # Track which keys this data_frame matches

                    # Get all text from the dataframe for searching
                    all_text = ' '.join(data_frame.astype(str).values.flatten()).lower()

                    print(f"📋 PROCESSING TAB: {sheet_name}")
                    print(f"   🔍 Available financial keys: {financial_keys}")
                    print(f"   🔍 Entity mode: {entity_mode}")
                    print(f"   🔍 Entity name: {entity_name}")
                    print(f"   🔍 Entity keywords: {entity_keywords}")
                    print(f"   📊 DataFrame shape: {data_frame.shape}")
                    print(f"   📝 Sample content (first 3 rows):")
                    for idx in range(min(3, len(data_frame))):
                        row_content = ' | '.join(str(val)[:50] for val in data_frame.iloc[idx] if pd.notna(val))
                        print(f"      Row {idx}: {row_content}")
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
                    for best_key in matched_keys:
                        print(f"   🎯 Processing key '{best_key}' for tab '{sheet_name}'")
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
                            parsed_table = parse_accounting_table(data_frame, best_key, entity_name, sheet_name, latest_date_col, actual_entity_found)
                            
                            print(f"   🔍 parse_accounting_table returned: {parsed_table is not None}")
                            if parsed_table:
                                print(f"   🔍 parsed_table keys: {list(parsed_table.keys()) if isinstance(parsed_table, dict) else 'not a dict'}")
                            
                            if parsed_table:
                                # Entity validation - use the same intelligent logic as above
                                if actual_entity_found is not None:
                                    # Entity was already validated in the intelligent detection above
                                    is_selected_entity = True
                                    section_text = ' '.join(data_frame.astype(str).values.flatten()).lower()
                                else:
                                    # Perform final validation check
                                    section_text = ' '.join(data_frame.astype(str).values.flatten()).lower()
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
                                
                                section_data = {
                                    'sheet': sheet_name,
                                    'data': data_frame,  # Keep original for compatibility
                                    'parsed_data': parsed_table,
                                    'markdown': create_improved_table_markdown(parsed_table),
                                    'entity_match': True,
                                    'is_selected_entity': is_selected_entity
                                }
                                
                                # Only add this section if it's the correct sheet for this key AND matches the selected entity
                                # This prevents wrong sheets from being assigned to keys
                                if (sheet_name.lower() == best_key.lower() or 
                                    (best_key in tab_name_mapping and 
                                     any(pattern.lower() in sheet_name.lower() for pattern in tab_name_mapping[best_key]))):
                                    
                                    # Only add sections that match the selected entity
                                    # Temporarily disable entity matching to debug
                                    sections_by_key[best_key].append(section_data)
                                    print(f"   ✅ SUCCESS: Added section for '{best_key}' from tab '{sheet_name}'")
                                    print(f"   📊 Section contains {len(parsed_table.get('data', []))} data rows")
                                    print(f"   💰 RMB detected: {'YES' if rmb_found else 'NO'}")
                                    print(f"   🔍 Entity found: {actual_entity_found}")
                                    print(f"   📊 Total sections for {best_key}: {len(sections_by_key[best_key])}")
                                    if not is_selected_entity:
                                        print(f"   ⚠️  Note: entity mismatch (found: {actual_entity_found}, expected: {entity_keywords})")
                            else:
                                # Fallback to original format if parsing fails
                                print(f"   ⚠️  parse_accounting_table failed for {best_key}, using fallback")
                                try:
                                    markdown_content = tabulate(data_frame, headers='keys', tablefmt='pipe') + '\n\n'
                                except Exception:
                                    markdown_content = data_frame.to_markdown(index=False) + '\n\n'
                                
                                sections_by_key[best_key].append({
                                    'sheet': sheet_name,
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
        
        # Create table
        markdown_lines.append("| Description | Value |")
        markdown_lines.append("|-------------|-------|")
        
        for row in data_rows:
            desc = row['description']
            value = row['value']
            is_total = row.get('is_total', False)
            
            # Format value with commas
            if isinstance(value, (int, float)):
                formatted_value = f"{value:,.2f}"
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

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
from fdd_utils.simple_cache import get_simple_cache


def detect_latest_date_column(df, sheet_name="Sheet", entity_keywords=None):
    """Detect the latest date column from a DataFrame, including xMxx format dates."""
    
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
            '%d/%B/%Y', '%d-%B-%Y', '%B/%d/%Y', '%B-%d-%Y'
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
    
    print(f"🔍 {sheet_name}: Searching for latest date column...")
    print(f"   Available columns: {columns}")
    
    # Strategy 1: Extract entity-related tables first, then find "Indicative adjusted" columns
    print(f"   🎯 STEP 1: Extracting entity-related tables from {sheet_name}")
    
    # Extract entity-related tables using the same logic as the main processing
    entity_tables = []
    entity_keywords_list = entity_keywords or []
    
    # Split dataframe into sections based on empty rows
    empty_rows = df.index[df.isnull().all(1)]
    start_idx = 0
    sections = []
    
    for end_idx in empty_rows:
        if end_idx > start_idx:
            section_df = df[start_idx:end_idx]
            if not section_df.dropna(how='all').empty:
                sections.append((start_idx, end_idx, section_df))
            start_idx = end_idx + 1
    
    if start_idx < len(df):
        sections.append((start_idx, len(df), df[start_idx:]))
    
    # Filter sections to only those containing entity keywords
    for start_row, end_row, section_df in sections:
        # Check if this section contains entity keywords
        section_has_entity = False
        all_cells = [str(cell).lower() for cell in section_df.values.flatten()]
        
        for keyword in entity_keywords_list:
            if any(keyword.lower() in cell for cell in all_cells):
                section_has_entity = True
                entity_tables.append((start_row, end_row, section_df))
                print(f"   ✅ Entity table found: Rows {start_row}-{end_row} (contains '{keyword}')")
                break
    
    if not entity_tables:
        print(f"   ⚠️  No entity-specific tables found, using first section as fallback")
        if sections:
            entity_tables = [sections[0]]
    
    print(f"   📊 Found {len(entity_tables)} entity-related tables")
    
    # Strategy 2: Within entity tables, find "Indicative adjusted" and get the correct column
    indicative_positions = []
    all_found_dates = []
    
    for start_row, end_row, entity_df in entity_tables:
        print(f"   🔍 STEP 2: Searching entity table (rows {start_row}-{end_row}) for 'Indicative adjusted'")
        
        # Find "Indicative adjusted" positions within this entity table
        for local_row_idx in range(len(entity_df)):
            global_row_idx = start_row + local_row_idx
            for col_idx, col in enumerate(columns):
                val = entity_df.iloc[local_row_idx, col_idx]
                if pd.notna(val) and 'indicative' in str(val).lower() and 'adjusted' in str(val).lower():
                    indicative_positions.append((global_row_idx, col_idx))
                    print(f"     📋 Found 'Indicative adjusted' at Row {global_row_idx}, Col {col_idx} ({col})")
        
        # Find dates within this entity table
        for local_row_idx in range(len(entity_df)):
            global_row_idx = start_row + local_row_idx
            for col_idx, col in enumerate(columns):
                val = entity_df.iloc[local_row_idx, col_idx]
                
                if isinstance(val, (pd.Timestamp, datetime)):
                    date_val = val if isinstance(val, datetime) else val.to_pydatetime()
                    all_found_dates.append((date_val, col, global_row_idx, col_idx, "datetime"))
                    print(f"     📅 Found datetime in {col}[{global_row_idx}]: {date_val.strftime('%Y-%m-%d')}")
                elif pd.notna(val):
                    parsed_date = parse_date(str(val))
                    if parsed_date:
                        all_found_dates.append((parsed_date, col, global_row_idx, col_idx, "parsed"))
                        print(f"     📅 Parsed date in {col}[{global_row_idx}]: '{val}' -> {parsed_date.strftime('%Y-%m-%d')}")
    
    # Strategy 3: Prioritize "Indicative adjusted" columns with latest dates
    if indicative_positions and all_found_dates:
        print(f"   🎯 STEP 3: Prioritizing 'Indicative adjusted' columns with latest dates")
        
        # Find the latest date
        max_date = max(all_found_dates, key=lambda x: x[0])[0]
        latest_date_columns = [item for item in all_found_dates if item[0] == max_date]
        
        print(f"   📊 Latest date found: {max_date.strftime('%Y-%m-%d')}")
        if len(latest_date_columns) > 1:
            print(f"   📊 Multiple columns with latest date:")
            for date_val, col, row, col_idx, source in latest_date_columns:
                print(f"      • {col} (col {col_idx})")
        
        # Find which "Indicative adjusted" positions have the latest date
        selected_column = None
        
        for indic_row, indic_col in indicative_positions:
            print(f"   🔍 Checking 'Indicative adjusted' at col {indic_col}")
            
            # Check if this exact column has the latest date
            exact_match = None
            for date_val, col, row, col_idx, source in latest_date_columns:
                if col_idx == indic_col:
                    exact_match = (date_val, col, row, col_idx, source)
                    print(f"     ✅ EXACT match: {col} (col {col_idx}) has latest date")
                    break
            
            if exact_match:
                selected_column = exact_match
                print(f"   🎯 SELECTED: {selected_column[1]} (EXACT 'Indicative adjusted' column with latest date)")
                break
            else:
                # Check merged range for this "Indicative adjusted"
                print(f"     🔍 EXACT column doesn't have latest date, checking merged range...")
                
                # Detect merged range
                merge_start = indic_col
                merge_end = indic_col
                
                for check_col in range(indic_col + 1, len(columns)):
                    val = df.iloc[indic_row, check_col]
                    if pd.isna(val):
                        merge_end = check_col
                    else:
                        merge_end = check_col - 1
                        break
                else:
                    merge_end = len(columns) - 1
                
                print(f"     📍 Merged range: columns {merge_start}-{merge_end}")
                
                # Find latest date within this merged range
                range_matches = []
                for date_val, col, row, col_idx, source in latest_date_columns:
                    if merge_start <= col_idx <= merge_end:
                        range_matches.append((date_val, col, row, col_idx, source))
                        print(f"     ✅ {col} (col {col_idx}) is in merged range with latest date")
                
                if range_matches:
                    selected_column = range_matches[0]  # Use first match in range
                    print(f"   🎯 SELECTED: {selected_column[1]} (latest date in 'Indicative adjusted' merged range)")
                    break
        
        if selected_column:
            latest_date, latest_column = selected_column[0], selected_column[1]
        else:
            # No "Indicative adjusted" column found with latest date, use first latest
            selected_column = latest_date_columns[0]
            latest_date, latest_column = selected_column[0], selected_column[1]
            print(f"   ⚠️  No 'Indicative adjusted' match, using: {latest_column}")
    
    # Strategy 2: Fallback to simple logic if no "Indicative adjusted" found
    else:
        print(f"   🔍 No 'Indicative adjusted' found, using simple date detection...")
        
        # First, try to find dates in column names
        column_dates_found = []
        for col in columns:
            col_str = str(col)
            parsed_date = parse_date(col_str)
            if parsed_date:
                column_dates_found.append((parsed_date, col, "column_name"))
                print(f"   📅 Found date in column name '{col}': {parsed_date.strftime('%Y-%m-%d')}")
                if latest_date is None or parsed_date > latest_date:
                    latest_date = parsed_date
                    latest_column = col
                    print(f"   ✅ New latest: {col} ({parsed_date.strftime('%Y-%m-%d')})")
        
        # If no dates found in column names, check the first few rows for datetime values
        if latest_column is None and len(df) > 0:
            print(f"   🔍 No dates in column names, checking row values...")
            cell_dates_found = []
            
            # Check first 5 rows for date values (dates can be in different rows)
            for row_idx in range(min(5, len(df))):
                row = df.iloc[row_idx]
                for col in columns:
                    val = row[col]
                    
                    # Check if it's already a datetime object
                    if isinstance(val, (pd.Timestamp, datetime)):
                        date_val = val if isinstance(val, datetime) else val.to_pydatetime()
                        cell_dates_found.append((date_val, col, f"row_{row_idx}"))
                        print(f"   📅 Found datetime in {col}[{row_idx}]: {date_val.strftime('%Y-%m-%d')}")
                        if latest_date is None or date_val > latest_date:
                            latest_date = date_val
                            latest_column = col
                            print(f"   ✅ New latest: {col} ({date_val.strftime('%Y-%m-%d')}) from row {row_idx}")
                    # Check if it's a string that can be parsed as a date
                    elif pd.notna(val):
                        parsed_date = parse_date(str(val))
                        if parsed_date:
                            cell_dates_found.append((parsed_date, col, f"row_{row_idx}_parsed"))
                            print(f"   📅 Parsed date in {col}[{row_idx}]: '{val}' -> {parsed_date.strftime('%Y-%m-%d')}")
                            if latest_date is None or parsed_date > latest_date:
                                latest_date = parsed_date
                                latest_column = col
                                print(f"   ✅ New latest: {col} ({parsed_date.strftime('%Y-%m-%d')}) from row {row_idx}")
            
            if not cell_dates_found:
                print(f"   ❌ No dates found in cell values")
    
    # Summary of selection
    if latest_column:
        print(f"   🎯 FINAL SELECTION: Column '{latest_column}' with date {latest_date.strftime('%Y-%m-%d')}")
        
        # Show comparison if multiple dates were found
        if 'all_found_dates' in locals() and len(all_found_dates) > 1:
            print(f"   📊 All dates found (for comparison):")
            for date_val, col, row, col_idx, source in sorted(all_found_dates, key=lambda x: x[0], reverse=True):
                marker = "👑" if col == latest_column else "  "
                print(f"   {marker} {col}: {date_val.strftime('%Y-%m-%d')} (from row {row})")
    else:
        print(f"   ❌ No date column detected")
    
    return latest_column


def parse_accounting_table(df, key, entity_name, sheet_name, latest_date_col=None, debug=False):
    """
    Parse accounting table with proper header detection and figure column identification
    Returns structured table data with metadata
    """
    try:
        
        if df.empty or len(df) < 2:
            return None
        
        # Debug info reduced for cleaner output
        if debug:  # Only show if explicitly debugging
            print(f"DEBUG: DataFrame shape: {df.shape}")
        
        # Clean the DataFrame first - drop unnamed columns that are all NaN
        # BUT preserve the detected latest_date_col if it's an "Unnamed:" column
        df_clean = df.copy()
        dropped_columns = []
        for col in df_clean.columns:
            # Don't drop the detected latest_date_col even if it's "Unnamed:"
            if col == latest_date_col:
                continue
            if col.startswith('Unnamed:') or df_clean[col].isna().all():
                dropped_columns.append(col)
                df_clean = df_clean.drop(columns=[col])
        
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
        thousand_indicators = ["'000", "'000", "cny'000", "thousands"]
        million_indicators = ["'000,000", "millions", "cny'000,000"]
        
        for i in range(min(3, len(df_str))):
            for j in range(len(df_str.columns)):
                cell_value = str(df_str.iloc[i, j]).lower()
                if any(indicator in cell_value for indicator in thousand_indicators):
                    multiplier = 1000
                    currency_info = "CNY'000"
                    break
                elif any(indicator in cell_value for indicator in million_indicators):
                    multiplier = 1000000
                    currency_info = "CNY'000,000"
                    break
        
        # Find the value column - use detected latest_date_col if provided
        value_col_idx = None
        value_col_name = ""
        
        if latest_date_col and latest_date_col in df_str.columns:
            # Use the detected latest date column
            value_col_idx = df_str.columns.get_loc(latest_date_col)
            value_col_name = "Indicative adjusted"  # This is the detected column
            print(f"   🎯 Using detected latest date column: {latest_date_col} (col {value_col_idx})")
        else:
            # Fallback to original logic
            print(f"   ⚠️  No latest_date_col provided, using fallback detection")
            for i in range(min(3, len(df_str))):  # Check first 3 rows for headers
                for j in range(len(df_str.columns)):
                    cell_value = str(df_str.iloc[i, j]).lower()
                    if "indicative adjusted" in cell_value:
                        value_col_idx = j
                        value_col_name = "Indicative adjusted"
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
                    # Additional check: skip if the description is a pure number (like 1000, 1001, etc.)
                    if not re.match(r'^\d+\.?\d*$', desc_cell):
                        data_start_row = i
                        break
        
        if data_start_row is None:
            # Fallback: start from row 2 if we have at least 3 rows
            if len(df_str) >= 3:
                data_start_row = 2
            else:
                return None
        
        # Extract date from the detected latest date column
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
        table_metadata = {
            'table_name': f"{key} - {entity_name}",
            'sheet_name': sheet_name,
            'date': extracted_date,
            'currency_info': currency_info,
            'multiplier': multiplier,
            'value_column': value_col_name,
            'data_start_row': data_start_row
        }
        
        # Extract data rows
        data_rows = []
        for i in range(data_start_row, len(df_str)):
            desc_cell = str(df_str.iloc[i, 0]).strip()
            value_cell = str(df_str.iloc[i, value_col_idx]).strip()
            
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


def process_and_filter_excel(filename, tab_name_mapping, entity_name, entity_suffixes):
    """
    Process and filter Excel file to extract relevant worksheet sections with simple caching
    This is the core function from old_ver/utils/utils.py
    """
    try:
        # Use simple cache instead of complex cache manager
        cache = get_simple_cache()
        
        # Check cache first (with force refresh option)
        try:
            force_refresh = st.session_state.get('force_refresh', False) if 'streamlit' in sys.modules else False
        except Exception:
            force_refresh = False
        
        cached_result = cache.get_cached_excel_data(filename, entity_name, force_refresh)
        if cached_result is not None:
            # Clear force refresh flag after using it
            try:
                if force_refresh and 'streamlit' in sys.modules:
                    st.session_state['force_refresh'] = False
            except Exception:
                pass
            return cached_result
        
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
        
        # Cache the processed result using simple cache
        cache.cache_excel_data(filename, entity_name, markdown_content)
        print(f"💾 Cached Excel data for {filename}")
        
        return markdown_content
        
    except Exception as e:
        print("An error occurred while processing the Excel file:", e)
        return ""


def get_worksheet_sections_by_keys(uploaded_file, tab_name_mapping, entity_name, entity_suffixes, debug=False):
    """
    Get worksheet sections organized by financial keys with enhanced entity filtering and latest date detection.
    """
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
                
                # Filter dataframes by entity name with proper spacing
                entity_keywords = [f"{entity_name} {suffix}" for suffix in entity_suffixes if suffix]
                if not entity_keywords:  # If no helpers, just use entity name
                    entity_keywords = [entity_name]
                
                combined_pattern = '|'.join(re.escape(kw) for kw in entity_keywords)
                
                # Detect latest date column once per sheet (not per dataframe)
                latest_date_col = detect_latest_date_column(df, sheet_name, entity_keywords)
                
                # Organize sections by key - make it less restrictive
                for data_frame in dataframes:
                    if debug and latest_date_col and 'streamlit' in sys.modules:
                        try:
                            st.write(f"📅 Latest date column detected: {latest_date_col}")
                        except Exception:
                            pass
                    
                    # Check if this section contains any of the financial keys
                    matched_keys = []  # Track which keys this data_frame matches
                    
                    # Get all text from the dataframe for searching
                    all_text = ' '.join(data_frame.astype(str).values.flatten()).lower()
                    
                    # Check each financial key
                    for financial_key in financial_keys:
                        # Get the key terms for this financial key
                        from fdd_utils.mappings import KEY_TERMS_BY_KEY
                        key_terms = KEY_TERMS_BY_KEY.get(financial_key, [financial_key.lower()])
                        
                        # Check if any key terms appear in the dataframe
                        key_found = any(term.lower() in all_text for term in key_terms)
                        
                        if key_found:
                            matched_keys.append(financial_key)
                    
                    # Process this dataframe for each matched key
                    for best_key in matched_keys:
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
                        
                        # If entity filter matches or helpers are empty, process
                        if entity_mask.any() or not entity_suffixes or all(s.strip() == '' for s in entity_suffixes):
                            # Use new accounting table parser with detected latest date column
                            parsed_table = parse_accounting_table(data_frame, best_key, entity_name, sheet_name, latest_date_col)
                            
                            if parsed_table:
                                sections_by_key[best_key].append({
                                    'sheet': sheet_name,
                                    'data': data_frame,  # Keep original for compatibility
                                    'parsed_data': parsed_table,
                                    'markdown': create_improved_table_markdown(parsed_table),
                                    'entity_match': True
                                })
                            else:
                                # Fallback to original format if parsing fails
                                try:
                                    markdown_content = tabulate(data_frame, headers='keys', tablefmt='pipe') + '\n\n'
                                except Exception:
                                    markdown_content = data_frame.to_markdown(index=False) + '\n\n'
                                
                                sections_by_key[best_key].append({
                                    'sheet': sheet_name,
                                    'data': data_frame,
                                    'markdown': markdown_content,
                                    'entity_match': True
                                })
        
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

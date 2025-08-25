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
    """Detect the latest date column from a DataFrame, focusing on 'Indicative adjusted' with merged cell handling."""
    
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
    
    # print(f"🔍 {sheet_name}: Searching for 'Indicative adjusted' column...")
    
    # Step 1: Find "Indicative adjusted" positions
    indicative_positions = []
    
    # Search in first 10 rows for "Indicative adjusted"
    for row_idx in range(min(10, len(df))):
        for col_idx, col in enumerate(columns):
            val = df.iloc[row_idx, col_idx]
            if pd.notna(val) and 'indicative' in str(val).lower() and 'adjusted' in str(val).lower():
                indicative_positions.append((row_idx, col_idx))
                                                # print(f"   📋 Found 'Indicative adjusted' at Row {row_idx}, Col {col_idx} ({col})")
    
    if not indicative_positions:
        print(f"   ⚠️  No 'Indicative adjusted' found, using fallback date detection")
        # Fallback: find any date column
        for col in columns:
            for row_idx in range(min(5, len(df))):
                val = df.iloc[row_idx, col_idx]
                if isinstance(val, (pd.Timestamp, datetime)):
                    date_val = val if isinstance(val, datetime) else val.to_pydatetime()
                    if latest_date is None or date_val > latest_date:
                        latest_date = date_val
                        latest_column = col
                        print(f"   📅 Fallback: Found date in {col}: {date_val.strftime('%Y-%m-%d')}")
        return latest_column
    
    # Step 2: For each "Indicative adjusted" position, find the merged range and get the date
    for indic_row, indic_col in indicative_positions:
        print(f"   🔍 Processing 'Indicative adjusted' at col {indic_col}")
        
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
        
        print(f"   📍 Merged range: columns {merge_start}-{merge_end}")
        
        # Find the date value in the row below the "Indicative adjusted" header
        date_row = indic_row + 1
        if date_row < len(df):
            # Look for date in the merged range
            for col_idx in range(merge_start, merge_end + 1):
                val = df.iloc[date_row, col_idx]
                
                if isinstance(val, (pd.Timestamp, datetime)):
                    date_val = val if isinstance(val, datetime) else val.to_pydatetime()
                    if latest_date is None or date_val > latest_date:
                        latest_date = date_val
                        latest_column = columns[col_idx]
                        print(f"   📅 Found date in merged range: {latest_column} = {date_val.strftime('%Y-%m-%d')}")
                elif pd.notna(val):
                    parsed_date = parse_date(str(val))
                    if parsed_date:
                        if latest_date is None or parsed_date > latest_date:
                            latest_date = parsed_date
                            latest_column = columns[col_idx]
                            print(f"   📅 Parsed date in merged range: {latest_column} = {parsed_date.strftime('%Y-%m-%d')}")
        
        # If no date found in the row below, check a few more rows down
        if latest_column is None:
            for check_row in range(indic_row + 2, min(indic_row + 5, len(df))):
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
    
    if latest_column:
        print(f"   🎯 FINAL SELECTION: Column '{latest_column}' with date {latest_date.strftime('%Y-%m-%d')}")
    else:
        print(f"   ❌ No date column detected")
    
    return latest_column


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
        
        # Keep the original dataframe structure - don't filter columns
        df_clean = df.copy()
        
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
            # print(f"   🔍 DEBUG: Column names: {list(df_str.columns)}")
            # print(f"   🔍 DEBUG: latest_date_col: {latest_date_col}")
            # print(f"   🔍 DEBUG: value_col_idx: {value_col_idx}")
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
        # Use actual entity found in data if available, otherwise use entity_name
        display_entity = actual_entity if actual_entity else entity_name
        table_metadata = {
            'table_name': f"{key} - {display_entity}",
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


def get_worksheet_sections_by_keys(uploaded_file, tab_name_mapping, entity_name, entity_suffixes, entity_keywords=None, debug=False):
    """
    Get worksheet sections organized by financial keys with enhanced entity filtering and latest date detection.
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
                    
                    print(f"   🔍 Processing sheet: {sheet_name}")
                    print(f"   🔍 Available financial keys: {financial_keys}")
                    
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
                        
                        # Check if this section contains the selected entity
                        section_text = ' '.join(data_frame.astype(str).values.flatten()).lower()
                        entity_found = any(entity_keyword.lower() in section_text for entity_keyword in entity_keywords)
                        
                        print(f"   🔍 Entity check for {best_key}: entity_found={entity_found}")
                        print(f"   🔍 Entity keywords: {entity_keywords}")
                        print(f"   🔍 Section text sample: {section_text[:200]}...")
                        
                        # Only process if entity is found in this section
                        if entity_found:
                            # Find the actual entity name from the section text
                            actual_entity_found = None
                            # First try to find the exact entity keyword
                            for entity_keyword in entity_keywords:
                                if entity_keyword.lower() in section_text:
                                    actual_entity_found = entity_keyword
                                    break
                            
                            # If not found, try to extract the actual entity name from the data
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
                                # Check if this section contains the selected entity
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
                                    print(f"   ✅ Added section for {best_key} with entity: {actual_entity_found}")
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

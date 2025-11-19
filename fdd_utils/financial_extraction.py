"""
Standalone Financial Extraction Helper Module
Extracts Balance Sheet and Income Statement data from Excel workbooks
Based on the methods from backups/fdd_utils/excel_processing.py
"""

import pandas as pd
import re
from datetime import datetime
from typing import Dict, Tuple, Optional, List
import warnings

warnings.simplefilter(action='ignore', category=UserWarning)


def parse_date(date_str):
    """
    Parse date string in various formats including xMxx and Chinese formats.
    
    Args:
        date_str: Date string in various formats
        
    Returns:
        datetime object or None if parsing fails
    """
    if not date_str or pd.isna(date_str):
        return None
    
    date_str = str(date_str).strip()
    
    # AVOID CONFUSING NUMBERS WITH DATES
    if ',' in date_str and date_str.replace(',', '').replace('.', '').isdigit():
        num_val = float(date_str.replace(',', ''))
        if num_val > 10000:
            return None
    
    # Handle Chinese date range format: 2024年1-5月 (use the END month)
    chinese_range_match = re.match(r'^(\d{4})年(\d{1,2})-(\d{1,2})月$', date_str)
    if chinese_range_match:
        year = int(chinese_range_match.group(1))
        end_month = int(chinese_range_match.group(3))
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
    
    # Handle xMxx format (e.g., 9M22, 12M23)
    xmxx_match = re.match(r'^(\d+)M(\d{2})$', date_str)
    if xmxx_match:
        month = int(xmxx_match.group(1))
        year = 2000 + int(xmxx_match.group(2))
        if month == 12:
            return datetime(year, 12, 31)
        elif month in [1, 3, 5, 7, 8, 10]:
            return datetime(year, month, 31)
        elif month in [4, 6, 9, 11]:
            return datetime(year, month, 30)
        elif month == 2:
            if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0):
                return datetime(year, 2, 29)
            else:
                return datetime(year, 2, 28)
    
    # Handle standard date formats
    date_formats = [
        '%d/%m/%Y', '%d-%m-%Y', '%d/%m/%y', '%d-%m-%y',
        '%Y-%m-%d', '%m/%d/%Y', '%m-%d-%Y',
        '%d/%b/%Y', '%d-%b-%Y', '%b/%d/%Y', '%b-%d-%Y',
        '%d/%B/%Y', '%d-%B-%Y', '%B/%d/%Y', '%B-%d-%Y',
        '%Y年%m月%d日', '%Y年%m月', '%m月%d日', '%Y/%m/%d',
        '%Y.%m.%d', '%Y年%m月%d日', '%Y年%m月%d号',
        '%Y%m%d', '%d%m%Y', '%m%d%Y'
    ]
    
    for fmt in date_formats:
        try:
            return datetime.strptime(date_str, fmt)
        except (ValueError, TypeError):
            continue
    
    return None


def detect_table_header_row(df: pd.DataFrame, keywords: List[str] = None) -> Optional[int]:
    """
    Detect the header row containing indicative keywords.
    
    Args:
        df: DataFrame to search
        keywords: List of keywords to search for (default: indicative adjusted keywords)
        
    Returns:
        Row index of header or None if not found
    """
    if keywords is None:
        keywords = ['Indicative adjusted', '示意性调整后', "CNY'000", "人民币千元"]
    
    for idx, row in df.iterrows():
        row_str = ' '.join(row.astype(str).values)
        if any(keyword in row_str for keyword in keywords):
            return idx
    
    return None


def find_date_columns(df: pd.DataFrame, header_row_idx: int) -> Tuple[List[int], List[datetime], int]:
    """
    Find date columns and return the most recent date column index.
    
    Args:
        df: DataFrame to search
        header_row_idx: Index of the header row
        
    Returns:
        Tuple of (date_column_indices, parsed_dates, most_recent_column_index)
    """
    if header_row_idx >= len(df) - 1:
        return [], [], None
    
    date_row_idx = header_row_idx + 1
    date_row = df.iloc[date_row_idx]
    
    parsed_dates = []
    date_indices = []
    
    for col_idx, value in enumerate(date_row):
        parsed_date = parse_date(value)
        if parsed_date:
            parsed_dates.append(parsed_date)
            date_indices.append(col_idx)
    
    if not parsed_dates:
        return [], [], None
    
    # Find most recent date
    most_recent_idx = parsed_dates.index(max(parsed_dates))
    most_recent_col_idx = date_indices[most_recent_idx]
    
    return date_indices, parsed_dates, most_recent_col_idx


def extract_financial_table(
    df: pd.DataFrame,
    table_name: str,
    entity_keywords: Optional[List[str]] = None,
    debug: bool = False
) -> Optional[pd.DataFrame]:
    """
    Extract financial table (Balance Sheet or Income Statement) from a worksheet.
    Gets ALL date columns under "Indicative adjusted" or "示意性调整后", not just the latest.
    
    Args:
        df: DataFrame containing the financial data
        table_name: Name of the table (e.g., "Balance Sheet", "Income Statement")
        entity_keywords: Optional list of entity name components to search for
        debug: If True, print debugging information
        
    Returns:
        Cleaned DataFrame with Description column and ALL date columns
    """
    if debug:
        print(f"\n[DEBUG] Extracting {table_name}...")
        print(f"[DEBUG] DataFrame shape: {df.shape}")
    
    # Detect header row with "Indicative adjusted" or "示意性调整后"
    header_row_idx = None
    for idx, row in df.iterrows():
        row_str = ' '.join(row.astype(str).values)
        if 'Indicative adjusted' in row_str or '示意性调整后' in row_str:
            header_row_idx = idx
            break
    
    if header_row_idx is None:
        if debug:
            print(f"[DEBUG] ❌ No 'Indicative adjusted' or '示意性调整后' header found!")
        return None
    
    if debug:
        print(f"[DEBUG] ✅ Header row found at index: {header_row_idx}")
    
    # Find description column (has "CNY'000" or "人民币千元")
    desc_col_idx = None
    header_row = df.iloc[header_row_idx]
    for col_idx, value in enumerate(header_row):
        value_str = str(value)
        if "CNY'000" in value_str or "人民币千元" in value_str:
            desc_col_idx = col_idx
            break
    
    if desc_col_idx is None:
        if debug:
            print(f"[DEBUG] ❌ No description column found with 'CNY'000' or '人民币千元'")
        return None
    
    if debug:
        print(f"[DEBUG] ✅ Description column at index: {desc_col_idx}")
    
    # Find ALL date columns (row after header)
    date_row_idx = header_row_idx + 1
    date_row = df.iloc[date_row_idx]
    
    date_columns = []  # List of (col_idx, parsed_date, date_string)
    for col_idx in range(desc_col_idx + 1, len(date_row)):
        date_str = date_row.iloc[col_idx]
        if pd.isna(date_str) or str(date_str).strip() == '':
            break  # Stop at first empty column
        
        parsed_date = parse_date(date_str)
        if parsed_date:
            date_columns.append((col_idx, parsed_date, str(date_str)))
    
    if not date_columns:
        if debug:
            print(f"[DEBUG] ❌ No valid date columns found!")
            print(f"[DEBUG] Date row: {date_row.values}")
        return None
    
    if debug:
        print(f"[DEBUG] ✅ Found {len(date_columns)} date columns:")
        for col_idx, parsed_date, date_str in date_columns:
            print(f"[DEBUG]    Column {col_idx}: {date_str} → {parsed_date.strftime('%Y-%m-%d')}")
    
    # Check if CNY'000 multiplier needed
    date_row_str = ' '.join(date_row.astype(str).values)
    multiply_by_1000 = "CNY'000" in date_row_str or "人民币千元" in date_row_str
    
    if debug and multiply_by_1000:
        print(f"[DEBUG] Will multiply values by 1000 (CNY'000 detected)")
    
    # Extract data starting from row after dates
    data_start_row = date_row_idx + 1
    
    # Build result dataframe with Description + ALL date columns
    result_rows = []
    for row_idx in range(data_start_row, len(df)):
        row = df.iloc[row_idx]
        
        # Stop at empty row
        if row.isnull().all():
            break
        
        description = row.iloc[desc_col_idx]
        
        # Skip if description is null
        if pd.isna(description) or str(description).strip() == '':
            continue
        
        # Build row dict with description and all date values
        row_dict = {'Description': str(description).strip()}
        
        has_any_value = False
        for col_idx, parsed_date, date_str in date_columns:
            value = row.iloc[col_idx]
            
            # Try to convert to float
            try:
                numeric_value = float(value)
                if multiply_by_1000:
                    numeric_value *= 1000
                numeric_value = round(numeric_value, 0)
                
                # Use formatted date as column name
                col_name = parsed_date.strftime('%Y-%m-%d')
                row_dict[col_name] = int(numeric_value)
                
                if numeric_value != 0:
                    has_any_value = True
            except (ValueError, TypeError):
                # Use formatted date as column name
                col_name = parsed_date.strftime('%Y-%m-%d')
                row_dict[col_name] = None
        
        # Only add row if it has at least one non-zero value
        if has_any_value:
            result_rows.append(row_dict)
    
    if not result_rows:
        if debug:
            print(f"[DEBUG] ❌ No valid data rows found!")
        return None
    
    result_df = pd.DataFrame(result_rows)
    
    # Remove rows where all values are None or 0
    value_cols = [col for col in result_df.columns if col != 'Description']
    result_df = result_df.dropna(subset=value_cols, how='all')
    
    if debug:
        print(f"[DEBUG] ✅ Final DataFrame: {len(result_df)} rows × {len(result_df.columns)} columns")
        print(f"[DEBUG] Columns: {list(result_df.columns)}")
        print(f"[DEBUG] Sample data:")
        print(result_df.head(3).to_string())
    
    return result_df


def extract_balance_sheet_and_income_statement(
    workbook_path: str,
    sheet_name: str,
    debug: bool = False
) -> Dict[str, any]:
    """
    Extract Balance Sheet and Income Statement from a SINGLE Excel worksheet.
    Both BS and IS are in the same sheet, separated by header rows.
    
    Args:
        workbook_path: Path to Excel workbook
        sheet_name: Worksheet name containing both BS and IS
        debug: If True, print debugging information
        
    Returns:
        Dictionary with keys:
        - 'balance_sheet': DataFrame or None
        - 'income_statement': DataFrame or None  
        - 'project_name': String (extracted from headers) or None
        
    Example:
        >>> results = extract_balance_sheet_and_income_statement(
        ...     workbook_path="databook.xlsx",
        ...     sheet_name="Financial Statements",
        ...     debug=True
        ... )
        >>> print(results['balance_sheet'])
        >>> print(results['income_statement'])
        >>> print(results['project_name'])
    """
    results = {
        'balance_sheet': None,
        'income_statement': None,
        'project_name': None
    }
    
    if debug:
        print("=" * 80)
        print("FINANCIAL EXTRACTION - DEBUG MODE")
        print("=" * 80)
        print(f"Workbook: {workbook_path}")
        print(f"Sheet: {sheet_name}")
    
    try:
        # Load Excel file
        df = pd.read_excel(workbook_path, sheet_name=sheet_name, engine='openpyxl')
        
        if debug:
            print(f"\n[DEBUG] ✅ Sheet loaded: {df.shape}")
        
        # Find Balance Sheet section
        bs_start_row = None
        bs_keywords = ["示意性调整后资产负债表", "Indicative adjusted balance sheet", 
                       "Indicative Adjusted Balance Sheet"]
        
        for idx, row in df.iterrows():
            row_str = ' '.join(row.astype(str).values).lower()
            if any(kw.lower() in row_str for kw in bs_keywords):
                bs_start_row = idx
                if debug:
                    print(f"[DEBUG] ✅ Balance Sheet starts at row {idx}: {df.iloc[idx].values[0]}")
                break
        
        # Find Income Statement section  
        is_start_row = None
        is_keywords = ["示意性调整后利润表", "Indicative adjusted income statement",
                       "Indicative Adjusted Income Statement"]
        
        for idx, row in df.iterrows():
            row_str = ' '.join(row.astype(str).values).lower()
            if any(kw.lower() in row_str for kw in is_keywords):
                is_start_row = idx
                if debug:
                    print(f"[DEBUG] ✅ Income Statement starts at row {idx}: {df.iloc[idx].values[0]}")
                break
        
        # Extract project name (from header row pattern)
        project_name = None
        if bs_start_row is not None:
            header_text = str(df.iloc[bs_start_row].values[0])
            # Pattern: "xxxx利润表 - 东莞xx" or "Balance Sheet - Project Name"
            if ' - ' in header_text:
                project_name = header_text.split(' - ', 1)[1].strip()
            elif '-' in header_text:
                parts = header_text.split('-')
                if len(parts) > 1:
                    project_name = parts[-1].strip()
            
            if debug and project_name:
                print(f"[DEBUG] ✅ Project name extracted: '{project_name}'")
        
        results['project_name'] = project_name
        
        # Extract Balance Sheet
        if bs_start_row is not None:
            # Determine end row (either IS start or end of sheet)
            bs_end_row = is_start_row if is_start_row else len(df)
            df_bs = df.iloc[bs_start_row:bs_end_row].copy()
            
            results['balance_sheet'] = extract_financial_table(
                df_bs, "Balance Sheet", None, debug
            )
        
        # Extract Income Statement
        if is_start_row is not None:
            # IS goes to end of sheet
            df_is = df.iloc[is_start_row:].copy()
            
            results['income_statement'] = extract_financial_table(
                df_is, "Income Statement", None, debug
            )
        
        if debug:
            print("\n" + "=" * 80)
            print("EXTRACTION RESULTS:")
            print("=" * 80)
            print(f"Project Name: {results['project_name'] or '❌ Not found'}")
            print(f"Balance Sheet: {'✅ Extracted' if results['balance_sheet'] is not None else '❌ None'}")
            print(f"Income Statement: {'✅ Extracted' if results['income_statement'] is not None else '❌ None'}")
            if results['balance_sheet'] is not None:
                print(f"  - Balance Sheet rows: {len(results['balance_sheet'])}")
            if results['income_statement'] is not None:
                print(f"  - Income Statement rows: {len(results['income_statement'])}")
        
    except Exception as e:
        print(f"❌ Error extracting financial data: {e}")
        if debug:
            import traceback
            print("\n[DEBUG] Full traceback:")
            traceback.print_exc()
    
    return results


def filter_by_total_amount(df: pd.DataFrame, filter_keywords: Optional[List[str]] = None) -> pd.DataFrame:
    """
    Filter dataframe to show only total amounts, not detail line items.
    
    Args:
        df: DataFrame with Description and Value columns
        filter_keywords: Keywords that indicate detail items to filter out
        
    Returns:
        Filtered DataFrame with only major categories
    """
    if df is None or df.empty:
        return df
    
    if filter_keywords is None:
        # Default keywords for sub-account filtering (Chinese and English)
        filter_keywords = ['_', '其中:', '其中：', 'including:', 'including：', '  -', '   ']
    
    # Filter out rows that contain filtering keywords
    filtered_df = df.copy()
    for keyword in filter_keywords:
        filtered_df = filtered_df[~filtered_df['Description'].str.contains(keyword, na=False)]
    
    return filtered_df


def get_account_total(df: pd.DataFrame, account_name: str, date_column: str = None) -> Optional[float]:
    """
    Get the total value for a specific account name.
    
    Args:
        df: DataFrame with Description and date columns
        account_name: Name of the account to search for
        date_column: Specific date column to get value from (e.g., '2024-12-31')
                    If None, returns the most recent date column value
        
    Returns:
        Total value for specified date or None if not found
    """
    if df is None or df.empty:
        return None
    
    # Search for exact match first
    matches = df[df['Description'] == account_name]
    if matches.empty:
        # Search for partial match
        matches = df[df['Description'].str.contains(account_name, na=False)]
    
    if matches.empty:
        return None
    
    # Get the row
    row = matches.iloc[0]
    
    # If no specific date column specified, use the most recent (last date column)
    if date_column is None:
        date_cols = [col for col in df.columns if col != 'Description']
        if not date_cols:
            return None
        # Most recent is typically the first date column
        date_column = date_cols[0]
    
    # Return value for that date
    if date_column in row.index:
        return row[date_column]
    
    return None


# Example usage and testing
if __name__ == "__main__":
    # Example: Extract BS and IS from single sheet
    print("="*80)
    print("Example: Extract Balance Sheet and Income Statement from Single Sheet")
    print("="*80)
    
    workbook_path = "databook.xlsx"
    sheet_name = "Financial Statements"  # Sheet containing both BS and IS
    
    results = extract_balance_sheet_and_income_statement(
        workbook_path=workbook_path,
        sheet_name=sheet_name,
        debug=True  # Enable debugging
    )
    
    print(f"\n{'='*80}")
    print("EXTRACTION SUMMARY")
    print(f"{'='*80}")
    
    # Show project name
    if results['project_name']:
        print(f"✅ Project Name: {results['project_name']}")
    else:
        print("❌ Project Name: Not found")
    
    # Show Balance Sheet
    if results['balance_sheet'] is not None:
        print(f"\n✅ Balance Sheet Extracted:")
        print(f"   Total rows: {len(results['balance_sheet'])}")
        print(f"   Columns: {list(results['balance_sheet'].columns)}")
        print(f"   Sample data:")
        print(results['balance_sheet'].head(5))
    else:
        print("\n❌ Balance Sheet: Not found")
    
    # Show Income Statement
    if results['income_statement'] is not None:
        print(f"\n✅ Income Statement Extracted:")
        print(f"   Total rows: {len(results['income_statement'])}")
        print(f"   Columns: {list(results['income_statement'].columns)}")
        print(f"   Sample data:")
        print(results['income_statement'].head(5))
    else:
        print("\n❌ Income Statement: Not found")
    
    # Example: Filter by total amounts only
    if results['balance_sheet'] is not None:
        print(f"\n{'='*80}")
        print("Filter to show only totals (no sub-accounts)")
        print(f"{'='*80}")
        filtered_bs = filter_by_total_amount(results['balance_sheet'])
        print(f"Filtered from {len(results['balance_sheet'])} to {len(filtered_bs)} rows")
    
    # Example: Get specific account total
    if results['balance_sheet'] is not None:
        print(f"\n{'='*80}")
        print("Get specific account total")
        print(f"{'='*80}")
        
        # Get most recent value
        cash_total = get_account_total(results['balance_sheet'], "货币资金")
        if cash_total:
            print(f"货币资金 (Cash) - Latest: {cash_total:,.0f}")
        
        # Get specific date value
        date_cols = [col for col in results['balance_sheet'].columns if col != 'Description']
        if len(date_cols) > 1:
            cash_specific = get_account_total(results['balance_sheet'], "货币资金", date_column=date_cols[1])
            if cash_specific:
                print(f"货币资金 (Cash) - {date_cols[1]}: {cash_specific:,.0f}")


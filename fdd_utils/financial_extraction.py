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
    entity_keywords: Optional[List[str]] = None
) -> Optional[pd.DataFrame]:
    """
    Extract financial table (Balance Sheet or Income Statement) from a worksheet.
    
    Args:
        df: DataFrame containing the financial data
        table_name: Name of the table (e.g., "Balance Sheet", "Income Statement")
        entity_keywords: Optional list of entity name components to search for
        
    Returns:
        Cleaned DataFrame with financial data or None if extraction fails
    """
    # Detect header row
    header_row_idx = detect_table_header_row(df)
    if header_row_idx is None:
        return None
    
    # Find date columns
    date_indices, parsed_dates, most_recent_col_idx = find_date_columns(df, header_row_idx)
    if most_recent_col_idx is None:
        return None
    
    # Find description column (usually contains account names)
    desc_col_idx = None
    header_row = df.iloc[header_row_idx]
    for col_idx, value in enumerate(header_row):
        value_str = str(value)
        if "CNY'000" in value_str or "人民币千元" in value_str:
            desc_col_idx = col_idx
            break
    
    if desc_col_idx is None:
        return None
    
    # Extract data starting from row after dates
    date_row_idx = header_row_idx + 1
    data_start_row = date_row_idx + 1
    
    # Build result dataframe
    result_rows = []
    for row_idx in range(data_start_row, len(df)):
        row = df.iloc[row_idx]
        
        # Stop at empty row
        if row.isnull().all():
            break
        
        description = row.iloc[desc_col_idx]
        value = row.iloc[most_recent_col_idx]
        
        # Skip if description or value is null
        if pd.isna(description) or pd.isna(value):
            continue
        
        # Try to convert value to float
        try:
            numeric_value = float(value)
            if numeric_value != 0:
                result_rows.append({
                    'Description': str(description).strip(),
                    'Value': numeric_value
                })
        except (ValueError, TypeError):
            continue
    
    if not result_rows:
        return None
    
    result_df = pd.DataFrame(result_rows)
    
    # Check if values need to be multiplied by 1000 (CNY'000)
    date_row = df.iloc[date_row_idx]
    date_row_str = ' '.join(date_row.astype(str).values)
    if "CNY'000" in date_row_str or "人民币千元" in date_row_str:
        result_df['Value'] = result_df['Value'] * 1000
    
    return result_df


def extract_balance_sheet_and_income_statement(
    workbook_path: str,
    balance_sheet_name: str = "示意性调整后资产负债表",
    income_statement_name: str = "示意性调整后利润表",
    entity_keywords: Optional[List[str]] = None
) -> Dict[str, pd.DataFrame]:
    """
    Extract Balance Sheet and Income Statement from specified Excel workbook.
    
    Args:
        workbook_path: Path to Excel workbook
        balance_sheet_name: Worksheet name for balance sheet (default: Chinese name)
        income_statement_name: Worksheet name for income statement (default: Chinese name)
        entity_keywords: Optional list of entity name components to filter by
        
    Returns:
        Dictionary with keys 'balance_sheet' and 'income_statement', 
        each containing a DataFrame or None if not found
        
    Example:
        >>> results = extract_balance_sheet_and_income_statement(
        ...     "databook.xlsx",
        ...     balance_sheet_name="示意性调整后资产负债表",
        ...     income_statement_name="示意性调整后利润表",
        ...     entity_keywords=["联洋"]
        ... )
        >>> print(results['balance_sheet'])
        >>> print(results['income_statement'])
    """
    results = {
        'balance_sheet': None,
        'income_statement': None
    }
    
    try:
        # Load Excel file
        excel_file = pd.ExcelFile(workbook_path, engine='openpyxl')
        
        # Extract Balance Sheet
        if balance_sheet_name in excel_file.sheet_names:
            df_bs = pd.read_excel(workbook_path, sheet_name=balance_sheet_name, engine='openpyxl')
            results['balance_sheet'] = extract_financial_table(df_bs, "Balance Sheet", entity_keywords)
        
        # Extract Income Statement
        if income_statement_name in excel_file.sheet_names:
            df_is = pd.read_excel(workbook_path, sheet_name=income_statement_name, engine='openpyxl')
            results['income_statement'] = extract_financial_table(df_is, "Income Statement", entity_keywords)
        
    except Exception as e:
        print(f"Error extracting financial data: {e}")
    
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


def get_account_total(df: pd.DataFrame, account_name: str) -> Optional[float]:
    """
    Get the total value for a specific account name.
    
    Args:
        df: DataFrame with Description and Value columns
        account_name: Name of the account to search for
        
    Returns:
        Total value or None if not found
    """
    if df is None or df.empty:
        return None
    
    # Search for exact match first
    matches = df[df['Description'] == account_name]
    if not matches.empty:
        return matches.iloc[0]['Value']
    
    # Search for partial match
    matches = df[df['Description'].str.contains(account_name, na=False)]
    if not matches.empty:
        return matches.iloc[0]['Value']
    
    return None


# Example usage and testing
if __name__ == "__main__":
    # Example 1: Extract from Chinese databook
    print("="*60)
    print("Example 1: Extract Balance Sheet and Income Statement")
    print("="*60)
    
    workbook_path = "inputs/240624.联洋-databook.xlsx"
    entity_keywords = ["联洋"]
    
    results = extract_balance_sheet_and_income_statement(
        workbook_path,
        balance_sheet_name="示意性调整后资产负债表",
        income_statement_name="示意性调整后利润表",
        entity_keywords=entity_keywords
    )
    
    if results['balance_sheet'] is not None:
        print("\n✅ Balance Sheet Extracted:")
        print(results['balance_sheet'].head(10))
        print(f"Total rows: {len(results['balance_sheet'])}")
    else:
        print("\n❌ Balance Sheet not found")
    
    if results['income_statement'] is not None:
        print("\n✅ Income Statement Extracted:")
        print(results['income_statement'].head(10))
        print(f"Total rows: {len(results['income_statement'])}")
    else:
        print("\n❌ Income Statement not found")
    
    # Example 2: Filter by total amounts only
    print("\n" + "="*60)
    print("Example 2: Filter to show only totals (no detail items)")
    print("="*60)
    
    if results['balance_sheet'] is not None:
        filtered_bs = filter_by_total_amount(results['balance_sheet'])
        print("\nFiltered Balance Sheet (totals only):")
        print(filtered_bs)
    
    # Example 3: Get specific account total
    print("\n" + "="*60)
    print("Example 3: Get specific account total")
    print("="*60)
    
    if results['balance_sheet'] is not None:
        cash_total = get_account_total(results['balance_sheet'], "货币资金")
        if cash_total:
            print(f"\n货币资金 (Cash) Total: {cash_total:,.2f}")


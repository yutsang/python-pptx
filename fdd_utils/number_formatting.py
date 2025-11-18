"""
Number Formatting Utilities for Financial Reports
Handles Chinese units (万, 亿) and scientific notation conversion
"""

import re
from typing import Union


def format_number_chinese(value: Union[float, int], language: str = 'Chi') -> str:
    """
    Format number using Chinese units (万, 亿) for Chinese language,
    or K/million for English language.
    
    Args:
        value: Numeric value to format
        language: 'Chi' for Chinese units, 'Eng' for English units
        
    Returns:
        Formatted string with appropriate units
        
    Examples:
        >>> format_number_chinese(5000, 'Chi')
        '5千元'
        >>> format_number_chinese(50000, 'Chi')
        '5.0万元'
        >>> format_number_chinese(5000000, 'Chi')
        '500.0万元'
        >>> format_number_chinese(50000000, 'Chi')
        '5000.0万元'  # or '0.5亿元' depending on preference
        >>> format_number_chinese(500000000, 'Chi')
        '5.0亿元'
    """
    if pd.isna(value) or value == 0:
        return "0"
    
    # Handle negative numbers
    is_negative = value < 0
    abs_value = abs(value)
    
    if language == 'Chi':
        # Chinese formatting
        if abs_value < 10000:  # Less than 1万
            result = f"{abs_value:,.0f}元"
        elif abs_value < 100000000:  # Less than 1亿 (10,000 to 99,999,999)
            wan_value = abs_value / 10000
            result = f"{wan_value:.1f}万元"
        else:  # 1亿 or more
            yi_value = abs_value / 100000000
            result = f"{yi_value:.1f}亿元"
        
        # Add "人民币" prefix and handle negative
        if is_negative:
            return f"人民币-{result}"
        else:
            return f"人民币{result}"
    
    else:  # English formatting
        if abs_value < 10000:
            result = f"{abs_value:,.0f}"
        elif abs_value < 100000:  # 10,000 to 99,999 -> use K
            k_value = abs_value / 1000
            result = f"{k_value:.1f}K"
        else:  # 100,000 or more -> use million
            m_value = abs_value / 1000000
            result = f"{m_value:.1f} million"
        
        # Add CNY prefix and handle negative
        if is_negative:
            return f"CNY -{result}"
        else:
            return f"CNY {result}"


def convert_scientific_to_normal(value: Union[str, float, int]) -> float:
    """
    Convert scientific notation string to normal float.
    Handles strings like "4.27e7" and converts to 42700000.
    
    Args:
        value: Value in scientific notation or normal format
        
    Returns:
        Float value in normal notation
        
    Examples:
        >>> convert_scientific_to_normal("4.27e7")
        42700000.0
        >>> convert_scientific_to_normal(4.27e7)
        42700000.0
        >>> convert_scientific_to_normal("123456")
        123456.0
    """
    if isinstance(value, str):
        # Try to convert string to float (handles scientific notation)
        try:
            return float(value)
        except ValueError:
            return 0.0
    elif isinstance(value, (int, float)):
        return float(value)
    else:
        return 0.0


def detect_and_format_amount(
    value: Union[str, float, int],
    language: str = 'Chi'
) -> str:
    """
    Detect scientific notation and format appropriately.
    
    Args:
        value: Value to format (may be in scientific notation)
        language: Target language for formatting
        
    Returns:
        Formatted string
        
    Examples:
        >>> detect_and_format_amount("4.27e7", "Chi")
        '人民币4270.0万元'
        >>> detect_and_format_amount(4.27e7, "Chi")
        '人民币4270.0万元'
    """
    # Convert to normal float
    numeric_value = convert_scientific_to_normal(value)
    
    # Format using appropriate units
    return format_number_chinese(numeric_value, language)


def format_retained_earnings(value: Union[float, int], language: str = 'Chi') -> tuple:
    """
    Format retained earnings with special handling for negative values.
    Returns both the account name and formatted value.
    
    Args:
        value: Retained earnings value (may be negative)
        language: Target language
        
    Returns:
        Tuple of (account_name, formatted_value)
        
    Examples:
        >>> format_retained_earnings(-5000000, 'Chi')
        ('未弥补亏损', '人民币500.0万元')
        >>> format_retained_earnings(5000000, 'Chi')
        ('未分配利润', '人民币500.0万元')
    """
    if language == 'Chi':
        if value < 0:
            # For negative R/E, use "未弥补亏损" (Unrecovered Losses)
            account_name = "未弥补亏损"
            formatted_value = format_number_chinese(abs(value), language)
        else:
            # For positive R/E, use "未分配利润" (Retained Earnings)
            account_name = "未分配利润"
            formatted_value = format_number_chinese(value, language)
    else:  # English
        if value < 0:
            account_name = "Accumulated Losses"
            formatted_value = format_number_chinese(abs(value), language)
        else:
            account_name = "Retained Earnings"
            formatted_value = format_number_chinese(value, language)
    
    return account_name, formatted_value


def format_dataframe_values(df, language: str = 'Chi'):
    """
    Format all numeric values in a DataFrame using appropriate units.
    
    Args:
        df: DataFrame with numeric values
        language: Target language for formatting
        
    Returns:
        DataFrame with formatted values
    """
    import pandas as pd
    
    if df is None or df.empty:
        return df
    
    df_formatted = df.copy()
    
    # Find numeric columns
    for col in df_formatted.columns:
        if pd.api.types.is_numeric_dtype(df_formatted[col]):
            df_formatted[col + '_formatted'] = df_formatted[col].apply(
                lambda x: detect_and_format_amount(x, language)
            )
    
    return df_formatted


# Import pandas for type checking
try:
    import pandas as pd
except ImportError:
    pass


if __name__ == "__main__":
    # Test examples
    print("="*60)
    print("Number Formatting Tests")
    print("="*60)
    
    # Test 1: Scientific notation
    print("\nTest 1: Scientific Notation Conversion")
    print(f"4.27e7 (Chi): {detect_and_format_amount('4.27e7', 'Chi')}")
    print(f"4.27e7 (Eng): {detect_and_format_amount('4.27e7', 'Eng')}")
    print(f"4.3e6 (Chi): {detect_and_format_amount('4.3e6', 'Chi')}")
    print(f"4.3e6 (Eng): {detect_and_format_amount('4.3e6', 'Eng')}")
    
    # Test 2: Chinese units
    print("\nTest 2: Chinese Units")
    test_values = [5000, 50000, 500000, 5000000, 50000000, 500000000]
    for val in test_values:
        print(f"{val:,} -> Chi: {format_number_chinese(val, 'Chi')} | Eng: {format_number_chinese(val, 'Eng')}")
    
    # Test 3: Negative retained earnings
    print("\nTest 3: Retained Earnings (Negative)")
    account, value = format_retained_earnings(-5000000, 'Chi')
    print(f"Chi: {account} = {value}")
    
    account, value = format_retained_earnings(-5000000, 'Eng')
    print(f"Eng: {account} = {value}")
    
    # Test 4: Positive retained earnings
    print("\nTest 4: Retained Earnings (Positive)")
    account, value = format_retained_earnings(5000000, 'Chi')
    print(f"Chi: {account} = {value}")
    
    account, value = format_retained_earnings(5000000, 'Eng')
    print(f"Eng: {account} = {value}")


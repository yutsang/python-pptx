"""
Reconciliation Module
Compares financial data from two sources to verify accuracy
"""

import pandas as pd
import yaml
from typing import Dict, Tuple, Optional


def load_mappings(mappings_file: str = 'fdd_utils/mappings.yml') -> dict:
    """Load mappings configuration."""
    with open(mappings_file, 'r', encoding='utf-8') as f:
        return yaml.safe_load(f)


def find_account_in_dfs(account_name: str, dfs: Dict[str, pd.DataFrame], mappings: dict, debug: bool = False) -> Tuple[Optional[str], Optional[pd.DataFrame]]:
    """
    Find an account in dfs using mappings aliases with improved matching.
    
    Args:
        account_name: Account name from BS/IS
        dfs: Dictionary of DataFrames from extract_data_from_excel
        mappings: Mappings configuration
        debug: Enable debug output
        
    Returns:
        Tuple of (mapping_key, dataframe) or (None, None) if not found
    """
    if debug:
        print(f"    [MATCH] Searching for: '{account_name}'")
    
    # Try exact match in dfs keys first
    if account_name in dfs:
        if debug:
            print(f"    [MATCH]   ✅ Exact match in dfs keys")
        return account_name, dfs[account_name]
    
    # Remove common suffixes/prefixes for better matching
    account_clean = account_name.strip()
    for suffix in ['：', ':', '合计', '总计', 'Total', 'total']:
        account_clean = account_clean.replace(suffix, '').strip()
    
    # Try to find via mappings
    for key, config in mappings.items():
        if key.startswith('_'):  # Skip special keys
            continue
        
        if not isinstance(config, dict):
            continue
        
        aliases = config.get('aliases', [])
        
        # Check if account_name exactly matches any alias
        if account_name in aliases or account_clean in aliases:
            if key in dfs:
                if debug:
                    print(f"    [MATCH]   ✅ Found via alias match: key='{key}'")
                return key, dfs[key]
        
        # Check if any alias contains or is contained in account_name
        for alias in aliases:
            alias_clean = alias.strip()
            # Exact substring match
            if alias_clean.lower() in account_clean.lower() or account_clean.lower() in alias_clean.lower():
                if key in dfs:
                    if debug:
                        print(f"    [MATCH]   ✅ Found via partial match: alias='{alias}', key='{key}'")
                    return key, dfs[key]
    
    # Last resort: try direct partial matching in dfs keys
    for dfs_key in dfs.keys():
        if account_clean.lower() in dfs_key.lower() or dfs_key.lower() in account_clean.lower():
            if debug:
                print(f"    [MATCH]   ✅ Found via direct key match: '{dfs_key}'")
            return dfs_key, dfs[dfs_key]
    
    if debug:
        print(f"    [MATCH]   ❌ Not found")
        print(f"    [MATCH]   Available dfs keys: {list(dfs.keys())[:10]}...")
    
    return None, None


def get_total_from_dfs(dfs_df: pd.DataFrame, date_col: str, debug: bool = False) -> Optional[float]:
    """
    Get total value from DFS dataframe.
    Looks for rows with keywords like 'Total', '合计', '总计' or uses the last row.
    
    Args:
        dfs_df: DataFrame from dfs
        date_col: Date column to get value from
        debug: Enable debug output
        
    Returns:
        Total value or None
    """
    if dfs_df is None or dfs_df.empty:
        return None
    
    if date_col not in dfs_df.columns:
        return None
    
    # Keywords for total rows
    total_keywords = ['合计', '总计', 'Total', 'total', '小计']
    
    # Try to find total row
    desc_col = dfs_df.columns[0]  # First column is description
    for idx, row in dfs_df.iterrows():
        desc = str(row[desc_col]).lower()
        if any(keyword.lower() in desc for keyword in total_keywords):
            if debug:
                print(f"      Found total row: '{row[desc_col]}'")
            return row[date_col]
    
    # If no total row found, use the last non-zero row
    non_zero_rows = dfs_df[dfs_df[date_col] != 0]
    if not non_zero_rows.empty:
        last_value = non_zero_rows.iloc[-1][date_col]
        if debug:
            print(f"      Using last non-zero row value: {last_value}")
        return last_value
    
    # Fallback: use first row
    return dfs_df[date_col].iloc[0]


def reconcile_financial_statements(
    bs_is_results: Dict,
    dfs: Dict[str, pd.DataFrame],
    mappings_file: str = 'fdd_utils/mappings.yml',
    tolerance: float = 1.0,
    debug: bool = False
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Reconcile Balance Sheet and Income Statement between two data sources.
    Only uses the LATEST date column from BS/IS for comparison.
    
    Args:
        bs_is_results: Results from extract_balance_sheet_and_income_statement
                      with keys 'balance_sheet', 'income_statement', 'project_name'
        dfs: Dictionary of DataFrames from extract_data_from_excel
        mappings_file: Path to mappings.yml file
        tolerance: Tolerance for matching (default: 1.0, allows ±1 difference)
        debug: If True, print debugging information
        
    Returns:
        Tuple of (bs_reconciliation_df, is_reconciliation_df)
        Each DataFrame has columns:
        - Source_Account: Account name from BS/IS
        - Date: Date column (latest only)
        - Source_Value: Value from BS/IS
        - DFS_Account: Matched account key from dfs
        - DFS_Value: Total value from dfs
        - Match: '✅ Match' or '❌ Diff: X' or '⚠️ Not Found'
        - Difference: Absolute difference
    """
    if debug:
        print("=" * 80)
        print("RECONCILIATION - DEBUG MODE")
        print("=" * 80)
    
    mappings = load_mappings(mappings_file)
    
    bs_recon_rows = []
    is_recon_rows = []
    
    # Reconcile Balance Sheet
    if bs_is_results.get('balance_sheet') is not None:
        bs_df = bs_is_results['balance_sheet']
        date_cols = [col for col in bs_df.columns if col != 'Description']
        
        # Use only the LATEST date column (LAST one, as dates are typically oldest to newest)
        latest_date = date_cols[-1] if date_cols else None
        
        if debug:
            print(f"\n[RECON] Reconciling Balance Sheet...")
            print(f"[RECON]   Accounts to check: {len(bs_df)}")
            print(f"[RECON]   Available dates: {date_cols}")
            print(f"[RECON]   Using latest date (last column): {latest_date}")
        
        if latest_date:
            for idx, row in bs_df.iterrows():
                account_name = row['Description']
                source_value = row[latest_date]
                
                # Find matching account in dfs
                dfs_key, dfs_df = find_account_in_dfs(account_name, dfs, mappings, debug=debug and idx < 10)
                
                # Get total value from dfs
                dfs_value = get_total_from_dfs(dfs_df, latest_date, debug and idx < 10) if dfs_df is not None else None
                
                # Check match
                if dfs_value is None:
                    match_status = '⚠️ Not Found'
                    difference = None
                else:
                    difference = abs(source_value - dfs_value)
                    if difference <= tolerance:
                        match_status = '✅ Match'
                    else:
                        match_status = f'❌ Diff: {difference:,.0f}'
                
                bs_recon_rows.append({
                    'Source_Account': account_name,
                    'Date': latest_date,
                    'Source_Value': source_value,
                    'DFS_Account': dfs_key or 'Not Found',
                    'DFS_Value': dfs_value if dfs_value is not None else 0,
                    'Match': match_status,
                    'Difference': difference if difference is not None else 0
                })
    
    # Reconcile Income Statement
    if bs_is_results.get('income_statement') is not None:
        is_df = bs_is_results['income_statement']
        date_cols = [col for col in is_df.columns if col != 'Description']
        
        # Use only the LATEST date column (LAST one, as dates are typically oldest to newest)
        latest_date = date_cols[-1] if date_cols else None
        
        if debug:
            print(f"\n[RECON] Reconciling Income Statement...")
            print(f"[RECON]   Accounts to check: {len(is_df)}")
            print(f"[RECON]   Available dates: {date_cols}")
            print(f"[RECON]   Using latest date (last column): {latest_date}")
        
        if latest_date:
            for idx, row in is_df.iterrows():
                account_name = row['Description']
                source_value = row[latest_date]
                
                # Find matching account in dfs
                dfs_key, dfs_df = find_account_in_dfs(account_name, dfs, mappings, debug=debug and idx < 10)
                
                # Get total value from dfs
                dfs_value = get_total_from_dfs(dfs_df, latest_date, debug and idx < 10) if dfs_df is not None else None
                
                # Check match
                if dfs_value is None:
                    match_status = '⚠️ Not Found'
                    difference = None
                else:
                    difference = abs(source_value - dfs_value)
                    if difference <= tolerance:
                        match_status = '✅ Match'
                    else:
                        match_status = f'❌ Diff: {difference:,.0f}'
                
                is_recon_rows.append({
                    'Source_Account': account_name,
                    'Date': latest_date,
                    'Source_Value': source_value,
                    'DFS_Account': dfs_key or 'Not Found',
                    'DFS_Value': dfs_value if dfs_value is not None else 0,
                    'Match': match_status,
                    'Difference': difference if difference is not None else 0
                })
    
    # Create DataFrames
    bs_recon_df = pd.DataFrame(bs_recon_rows) if bs_recon_rows else pd.DataFrame()
    is_recon_df = pd.DataFrame(is_recon_rows) if is_recon_rows else pd.DataFrame()
    
    if debug:
        print("\n" + "=" * 80)
        print("RECONCILIATION SUMMARY")
        print("=" * 80)
        
        if not bs_recon_df.empty:
            matches = (bs_recon_df['Match'] == '✅ Match').sum()
            mismatches = bs_recon_df['Match'].str.contains('❌').sum()
            not_found = (bs_recon_df['Match'] == '⚠️ Not Found').sum()
            print(f"Balance Sheet: {len(bs_recon_df)} comparisons")
            print(f"  ✅ Matches: {matches}")
            print(f"  ❌ Mismatches: {mismatches}")
            print(f"  ⚠️  Not Found: {not_found}")
        
        if not is_recon_df.empty:
            matches = (is_recon_df['Match'] == '✅ Match').sum()
            mismatches = is_recon_df['Match'].str.contains('❌').sum()
            not_found = (is_recon_df['Match'] == '⚠️ Not Found').sum()
            print(f"\nIncome Statement: {len(is_recon_df)} comparisons")
            print(f"  ✅ Matches: {matches}")
            print(f"  ❌ Mismatches: {mismatches}")
            print(f"  ⚠️  Not Found: {not_found}")
    
    return bs_recon_df, is_recon_df


def print_reconciliation_report(bs_recon_df: pd.DataFrame, is_recon_df: pd.DataFrame, 
                                show_only_issues: bool = False):
    """
    Print a formatted reconciliation report.
    
    Args:
        bs_recon_df: Balance Sheet reconciliation DataFrame
        is_recon_df: Income Statement reconciliation DataFrame
        show_only_issues: If True, only show mismatches and not found items
    """
    print("\n" + "=" * 100)
    print("RECONCILIATION REPORT")
    print("=" * 100)
    
    # Balance Sheet
    if not bs_recon_df.empty:
        print("\n📊 BALANCE SHEET RECONCILIATION")
        print("-" * 100)
        
        df_to_show = bs_recon_df.copy()
        if show_only_issues:
            df_to_show = df_to_show[df_to_show['Match'] != '✅ Match']
        
        if not df_to_show.empty:
            # Format for display
            df_display = df_to_show.copy()
            df_display['Source_Value'] = df_display['Source_Value'].apply(lambda x: f"{x:,.0f}")
            df_display['DFS_Value'] = df_display['DFS_Value'].apply(lambda x: f"{x:,.0f}")
            
            print(df_display.to_string(index=False))
        else:
            print("✅ All accounts match perfectly!")
    
    # Income Statement
    if not is_recon_df.empty:
        print("\n\n📈 INCOME STATEMENT RECONCILIATION")
        print("-" * 100)
        
        df_to_show = is_recon_df.copy()
        if show_only_issues:
            df_to_show = df_to_show[df_to_show['Match'] != '✅ Match']
        
        if not df_to_show.empty:
            # Format for display
            df_display = df_to_show.copy()
            df_display['Source_Value'] = df_display['Source_Value'].apply(lambda x: f"{x:,.0f}")
            df_display['DFS_Value'] = df_display['DFS_Value'].apply(lambda x: f"{x:,.0f}")
            
            print(df_display.to_string(index=False))
        else:
            print("✅ All accounts match perfectly!")
    
    print("\n" + "=" * 100)


# Example usage
if __name__ == "__main__":
    from fdd_utils.financial_extraction import extract_balance_sheet_and_income_statement
    from fdd_utils.process_databook import extract_data_from_excel
    
    print("=" * 80)
    print("RECONCILIATION EXAMPLE")
    print("=" * 80)
    
    # Example: Extract from both sources
    databook_path = "databook.xlsx"
    
    # Source 1: Financial extraction (BS/IS from single sheet)
    bs_is_results = extract_balance_sheet_and_income_statement(
        workbook_path=databook_path,
        sheet_name="Financials",
        debug=False
    )
    
    # Source 2: DFS extraction (account by account)
    dfs, workbook_list, _, language = extract_data_from_excel(
        databook_path=databook_path,
        entity_name="",
        mode="All"
    )
    
    # Reconcile
    bs_recon, is_recon = reconcile_financial_statements(
        bs_is_results=bs_is_results,
        dfs=dfs,
        tolerance=1.0,
        debug=True
    )
    
    # Print report
    print_reconciliation_report(bs_recon, is_recon, show_only_issues=True)
    
    # Save to Excel
    if not bs_recon.empty:
        with pd.ExcelWriter('reconciliation_report.xlsx') as writer:
            bs_recon.to_excel(writer, sheet_name='Balance Sheet', index=False)
            if not is_recon.empty:
                is_recon.to_excel(writer, sheet_name='Income Statement', index=False)
        print("\n✅ Reconciliation report saved to: reconciliation_report.xlsx")


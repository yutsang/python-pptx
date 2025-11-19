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


def find_account_in_dfs(account_name: str, dfs: Dict[str, pd.DataFrame], mappings: dict) -> Tuple[Optional[str], Optional[pd.DataFrame]]:
    """
    Find an account in dfs using mappings aliases.
    
    Args:
        account_name: Account name from BS/IS
        dfs: Dictionary of DataFrames from extract_data_from_excel
        mappings: Mappings configuration
        
    Returns:
        Tuple of (mapping_key, dataframe) or (None, None) if not found
    """
    # Try exact match first
    if account_name in dfs:
        return account_name, dfs[account_name]
    
    # Try to find via aliases
    for key, config in mappings.items():
        if key.startswith('_'):  # Skip special keys
            continue
        
        if not isinstance(config, dict):
            continue
        
        aliases = config.get('aliases', [])
        
        # Check if account_name is in aliases
        if account_name in aliases:
            # Find the corresponding key in dfs
            if key in dfs:
                return key, dfs[key]
        
        # Check if any alias matches partially
        for alias in aliases:
            if alias.lower() in account_name.lower() or account_name.lower() in alias.lower():
                if key in dfs:
                    return key, dfs[key]
    
    return None, None


def reconcile_financial_statements(
    bs_is_results: Dict,
    dfs: Dict[str, pd.DataFrame],
    mappings_file: str = 'fdd_utils/mappings.yml',
    tolerance: float = 1.0,
    debug: bool = False
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Reconcile Balance Sheet and Income Statement between two data sources.
    
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
        - Date: Date column
        - Source_Value: Value from BS/IS
        - DFS_Account: Matched account key from dfs
        - DFS_Value: Value from dfs
        - Match: '✅ Match' or '❌ Mismatch' or '⚠️ Not Found'
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
        
        if debug:
            print(f"\n[RECON] Reconciling Balance Sheet...")
            print(f"[RECON]   Accounts to check: {len(bs_df)}")
            print(f"[RECON]   Date columns: {date_cols}")
        
        for idx, row in bs_df.iterrows():
            account_name = row['Description']
            
            # Find matching account in dfs
            dfs_key, dfs_df = find_account_in_dfs(account_name, dfs, mappings)
            
            if debug and idx < 5:
                print(f"[RECON]   '{account_name}' → dfs key: '{dfs_key}'")
            
            for date_col in date_cols:
                source_value = row[date_col]
                
                # Get corresponding value from dfs
                dfs_value = None
                if dfs_df is not None and not dfs_df.empty:
                    # dfs has columns like: [Description, '2024-12-31', '2024-12-31_formatted']
                    if date_col in dfs_df.columns:
                        # Get the total (usually first row or sum)
                        dfs_value = dfs_df[date_col].iloc[0] if len(dfs_df) > 0 else None
                
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
                    'Date': date_col,
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
        
        if debug:
            print(f"\n[RECON] Reconciling Income Statement...")
            print(f"[RECON]   Accounts to check: {len(is_df)}")
            print(f"[RECON]   Date columns: {date_cols}")
        
        for idx, row in is_df.iterrows():
            account_name = row['Description']
            
            # Find matching account in dfs
            dfs_key, dfs_df = find_account_in_dfs(account_name, dfs, mappings)
            
            if debug and idx < 5:
                print(f"[RECON]   '{account_name}' → dfs key: '{dfs_key}'")
            
            for date_col in date_cols:
                source_value = row[date_col]
                
                # Get corresponding value from dfs
                dfs_value = None
                if dfs_df is not None and not dfs_df.empty:
                    if date_col in dfs_df.columns:
                        dfs_value = dfs_df[date_col].iloc[0] if len(dfs_df) > 0 else None
                
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
                    'Date': date_col,
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


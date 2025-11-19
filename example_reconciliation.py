#!/usr/bin/env python3
"""
Example: Reconcile financial data from two sources
Compares BS/IS extraction vs. DFS account-by-account extraction
"""

from fdd_utils.financial_extraction import extract_balance_sheet_and_income_statement
from fdd_utils.process_databook import extract_data_from_excel
from fdd_utils.reconciliation import reconcile_financial_statements, print_reconciliation_report
import pandas as pd


def run_reconciliation_example():
    """Run complete reconciliation example"""
    
    print("=" * 100)
    print("DATA RECONCILIATION EXAMPLE")
    print("=" * 100)
    
    # Configuration
    databook_path = "databook.xlsx"
    sheet_name = "Financials"  # Sheet with both BS and IS
    entity_name = ""  # Entity name for DFS extraction
    
    print(f"\nDatabook: {databook_path}")
    print(f"Sheet: {sheet_name}")
    
    # STEP 1: Extract BS/IS from single sheet
    print("\n" + "=" * 100)
    print("STEP 1: Extracting Balance Sheet & Income Statement (Single Sheet Method)")
    print("=" * 100)
    
    bs_is_results = extract_balance_sheet_and_income_statement(
        workbook_path=databook_path,
        sheet_name=sheet_name,
        debug=False  # Set to True to see detailed extraction
    )
    
    if bs_is_results['balance_sheet'] is not None:
        print(f"✅ Balance Sheet: {len(bs_is_results['balance_sheet'])} rows × {len(bs_is_results['balance_sheet'].columns)} columns")
        print(f"   Columns: {list(bs_is_results['balance_sheet'].columns)}")
    else:
        print("❌ Balance Sheet: Not extracted")
    
    if bs_is_results['income_statement'] is not None:
        print(f"✅ Income Statement: {len(bs_is_results['income_statement'])} rows × {len(bs_is_results['income_statement'].columns)} columns")
        print(f"   Columns: {list(bs_is_results['income_statement'].columns)}")
    else:
        print("❌ Income Statement: Not extracted")
    
    if bs_is_results['project_name']:
        print(f"✅ Project Name: {bs_is_results['project_name']}")
    
    # STEP 2: Extract using DFS method (account by account)
    print("\n" + "=" * 100)
    print("STEP 2: Extracting Data (Account-by-Account Method)")
    print("=" * 100)
    
    dfs, workbook_list, _, language = extract_data_from_excel(
        databook_path=databook_path,
        entity_name=entity_name,
        mode="All"
    )
    
    if dfs and len(dfs) > 0:
        print(f"✅ Extracted {len(dfs)} accounts")
        print(f"   Accounts: {workbook_list[:10]}...")  # Show first 10
    else:
        print("❌ DFS extraction failed")
        return
    
    # STEP 3: Reconcile
    print("\n" + "=" * 100)
    print("STEP 3: Reconciliation")
    print("=" * 100)
    
    bs_recon, is_recon = reconcile_financial_statements(
        bs_is_results=bs_is_results,
        dfs=dfs,
        tolerance=1.0,  # Allow ±1 for rounding errors
        debug=True
    )
    
    # STEP 4: Print report
    print("\n" + "=" * 100)
    print("STEP 4: Reconciliation Report")
    print("=" * 100)
    
    # Show all items
    print("\n📋 FULL REPORT (All Items)")
    print_reconciliation_report(bs_recon, is_recon, show_only_issues=False)
    
    # Show only issues
    print("\n⚠️  ISSUES ONLY (Mismatches & Not Found)")
    print_reconciliation_report(bs_recon, is_recon, show_only_issues=True)
    
    # STEP 5: Save to Excel
    print("\n" + "=" * 100)
    print("STEP 5: Saving Report")
    print("=" * 100)
    
    output_file = 'reconciliation_report.xlsx'
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        if not bs_recon.empty:
            bs_recon.to_excel(writer, sheet_name='BS Reconciliation', index=False)
        if not is_recon.empty:
            is_recon.to_excel(writer, sheet_name='IS Reconciliation', index=False)
    
    print(f"✅ Reconciliation report saved to: {output_file}")
    
    # Summary
    print("\n" + "=" * 100)
    print("SUMMARY")
    print("=" * 100)
    
    if not bs_recon.empty:
        total = len(bs_recon)
        matches = (bs_recon['Match'] == '✅ Match').sum()
        mismatches = bs_recon['Match'].str.contains('❌').sum()
        not_found = (bs_recon['Match'] == '⚠️ Not Found').sum()
        match_rate = (matches / total * 100) if total > 0 else 0
        
        print(f"\n📊 Balance Sheet:")
        print(f"   Total comparisons: {total}")
        print(f"   ✅ Matches: {matches} ({match_rate:.1f}%)")
        print(f"   ❌ Mismatches: {mismatches}")
        print(f"   ⚠️  Not Found: {not_found}")
    
    if not is_recon.empty:
        total = len(is_recon)
        matches = (is_recon['Match'] == '✅ Match').sum()
        mismatches = is_recon['Match'].str.contains('❌').sum()
        not_found = (is_recon['Match'] == '⚠️ Not Found').sum()
        match_rate = (matches / total * 100) if total > 0 else 0
        
        print(f"\n📈 Income Statement:")
        print(f"   Total comparisons: {total}")
        print(f"   ✅ Matches: {matches} ({match_rate:.1f}%)")
        print(f"   ❌ Mismatches: {mismatches}")
        print(f"   ⚠️  Not Found: {not_found}")


if __name__ == "__main__":
    run_reconciliation_example()


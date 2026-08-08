from __future__ import annotations

# re-added: bound by an import in another section of the pre-split module
from typing import Any, Dict, Iterable, List, Optional

"""
Reconciliation Module
Compares financial data from two sources to verify accuracy
"""

from .mapping import iter_account_mappings, load_mappings, normalize_mapping_label, should_skip_account_label
from .statements import extract_balance_sheet_and_income_statement
from .schedules import _parse_statement_date_label
from .databook import extract_data_from_excel

import pandas as pd
from pathlib import Path
from typing import Dict, Tuple, Optional

should_skip_mapping = should_skip_account_label


def _normalize_account_name(account_name: str) -> str:
    """Normalize account labels for tolerant matching."""
    return normalize_mapping_label(account_name)


def find_reconciliation_example(repo_root: Optional[Path] = None) -> Optional[dict]:
    """Return the first readable local workbook/sheet combination for the reconciliation demo."""
    base_dir = Path(repo_root) if repo_root is not None else Path(__file__).resolve().parents[2]
    candidates = [
        {
            'workbook_path': base_dir / 'databook-sample-a_rebuilt.xlsx',
            'sheet_name': 'Financials',
            'entity_name': 'Sample Entity A',
        },
        {
            'workbook_path': base_dir / 'databook-sample-b.xlsx',
            'sheet_name': 'Financials',
            'entity_name': '',
        },
        {
            'workbook_path': base_dir / 'databook.xlsx',
            'sheet_name': 'Financials',
            'entity_name': '',
        },
    ]

    for candidate in candidates:
        workbook_path = candidate['workbook_path']
        if not workbook_path.exists():
            continue
        try:
            workbook = pd.ExcelFile(workbook_path)
        except (OSError, ValueError, ImportError):
            continue
        if candidate['sheet_name'] in workbook.sheet_names:
            return {
                'workbook_path': str(workbook_path),
                'sheet_name': candidate['sheet_name'],
                'entity_name': candidate['entity_name'],
            }

    return None


def _resolve_mapping_alias(account_name: str, mappings: dict) -> Tuple[Optional[str], Optional[dict], Optional[str]]:
    account_clean = account_name.strip()
    account_normalized = _normalize_account_name(account_clean)
    for mapping_key, config in iter_account_mappings(mappings):
        aliases = config.get('aliases', [])
        normalized_aliases = {_normalize_account_name(alias) for alias in aliases}
        if (
            account_name in aliases
            or account_clean in aliases
            or account_normalized in normalized_aliases
        ):
            return mapping_key, config, config.get('category', None)
    return None, None, None


def find_account_in_dfs(
    account_name: str,
    dfs: Dict[str, pd.DataFrame],
    mappings: dict,
    debug: bool = False,
) -> Tuple[Optional[str], Optional[pd.DataFrame], Optional[str], Optional[str], str, str]:
    """
    Find an account in dfs by:
    1. Finding which mapping KEY the source account belongs to (via aliases)
    2. Preferring a direct normalized match against the available dfs keys
    3. Falling back to a normalized alias match if no direct key exists
    
    Args:
        account_name: Account name from BS/IS
        dfs: Dictionary of DataFrames from extract_data_from_excel
        mappings: Mappings configuration
        debug: Enable debug output
        
    Returns:
        Tuple of (dfs_key, dfs_df, category, mapping_key, mapping_status, mapping_note)
        where dfs_key is the actual matched key in dfs.
    """
    if debug:
        print(f"    [MATCH] Source account: '{account_name}'")
    
    # Check if this account should skip mapping
    if should_skip_mapping(account_name):
        if debug:
            print(f"    [MATCH]   ⏭️  Skipped (total/profit line)")
        return 'SKIP', None, None, None, 'Skipped', 'Skipped total/subtotal/profit line.'
    
    # Remove common suffixes for better matching
    account_clean = account_name.strip()
    account_normalized = _normalize_account_name(account_clean)
    
    if debug:
        print(f"    [MATCH]   Cleaned: '{account_clean}'")
        print(f"    [MATCH]   DFS keys: {list(dfs.keys())}")

    mapping_key, mapping_config, category = _resolve_mapping_alias(account_name, mappings)

    # Helper: collect the normalised identifiers for a dfs entry (tab key + block_title).
    def _dfs_names(key: str, df: pd.DataFrame) -> List[str]:
        names = [_normalize_account_name(key)]
        block_title = df.attrs.get("block_title") or ""
        if block_title:
            norm_bt = _normalize_account_name(block_title)
            if norm_bt and norm_bt != names[0]:
                names.append(norm_bt)
        return names

    # STEP 0: Always honor an exact source-account to DFS-key/block_title match first.
    for dfs_key, dfs_df_candidate in dfs.items():
        if account_normalized in _dfs_names(dfs_key, dfs_df_candidate):
            if debug:
                print(f"    [MATCH]   Step 0: ✅ Exact source-to-tab match '{dfs_key}'")
            if mapping_key:
                return (
                    dfs_key,
                    dfs_df_candidate,
                    category,
                    mapping_key,
                    'Mapped',
                    f"Mapped in mappings.yml as '{mapping_key}' and matched the workbook tab directly.",
                )
            return (
                dfs_key,
                dfs_df_candidate,
                None,
                None,
                'Tab-only match',
                "Matched by workbook tab name only. Add this account to mappings.yml if it should be classified into BS or IS.",
            )

    # STEP 1: Find which mapping KEY this source account belongs to
    if mapping_key and isinstance(mapping_config, dict):
        aliases = mapping_config.get('aliases', [])
        normalized_aliases = {_normalize_account_name(alias) for alias in aliases}
        if debug:
            print(f"    [MATCH]   Step 1: Found in mappings.yml")
            print(f"    [MATCH]     Mapping key: '{mapping_key}'")
            print(f"    [MATCH]     Category: '{category}'")
            print(f"    [MATCH]     Aliases: {aliases}")

        # STEP 2b: Fall back to any alias match against DFS keys or block_titles.
        for dfs_key, dfs_df_candidate in dfs.items():
            if normalized_aliases & set(_dfs_names(dfs_key, dfs_df_candidate)):
                if debug:
                    print(f"    [MATCH]   Step 2b: ✅ DFS key '{dfs_key}' matches alias!")
                return (
                    dfs_key,
                    dfs_df_candidate,
                    category,
                    mapping_key,
                    'Mapped',
                    f"Mapped in mappings.yml as '{mapping_key}'.",
                )

        # No alias matched any dfs key
        if debug:
            print(f"    [MATCH]   Step 2: ❌ No alias matches any dfs key")
            print(f"    [MATCH]     Aliases: {aliases}")
            print(f"    [MATCH]     DFS keys: {list(dfs.keys())}")
        return (
            None,
            None,
            category,
            mapping_key,
            'Mapped but missing tab',
            f"Mapped in mappings.yml as '{mapping_key}', but no matching workbook tab was found.",
        )

    # Not found in any mapping aliases
    if debug:
        print(f"    [MATCH]   ❌ '{account_name}' not in any mappings.yml aliases")

    return (
        None,
        None,
        None,
        None,
        'Missing mapping',
        "This account is not covered by mappings.yml. Add a mapping or alias if it should reconcile to a schedule tab.",
    )


def get_total_from_dfs(dfs_df: pd.DataFrame, date_col: str, debug: bool = False) -> Optional[float]:
    """
    Get total value from DFS dataframe.
    ONLY looks for rows with 'Total', '合计', '总计'.
    Skips subtotal rows ('Subtotal', '小计').
    No fallback - returns None if no total row found.
    
    Args:
        dfs_df: DataFrame from dfs
        date_col: Date column to get value from
        debug: Enable debug output
        
    Returns:
        Total value or None if no total row found
    """
    if dfs_df is None or dfs_df.empty:
        return None

    attrs = dfs_df.attrs or {}

    projection_totals_by_date = attrs.get('projection_totals_by_date') or {}
    if isinstance(projection_totals_by_date, dict):
        projection_total = projection_totals_by_date.get(date_col)
        if projection_total is not None:
            if debug:
                print(f"      Using normalized projection total for '{date_col}': {projection_total:,.0f}")
            return projection_total

    def _scan_table_total(target_date_col: str) -> Optional[float]:
        if target_date_col not in dfs_df.columns:
            return None
        desc_col = dfs_df.columns[0]
        total_keywords = ['合计', '总计', 'total']
        skip_keywords = ['小计', 'subtotal', 'sub-total', 'sub total']
        for _, row in dfs_df.iterrows():
            desc = str(row[desc_col])
            desc_lower = desc.lower()
            if any(skip_kw in desc_lower for skip_kw in skip_keywords):
                continue
            if any(keyword in desc_lower for keyword in total_keywords):
                return row[target_date_col]
        return None

    auxiliary_check_totals_by_date = attrs.get('auxiliary_check_totals_by_date') or {}
    if isinstance(auxiliary_check_totals_by_date, dict):
        auxiliary_total = auxiliary_check_totals_by_date.get(date_col)
        if auxiliary_total is not None:
            main_total = _scan_table_total(date_col)
            if isinstance(main_total, (int, float)) and main_total not in (0, 0.0):
                if abs(auxiliary_total) > 0 and (main_total > 0) != (auxiliary_total > 0):
                    auxiliary_total = abs(auxiliary_total) * (1 if main_total > 0 else -1)
            if debug:
                print(f"      Using auxiliary check total for '{date_col}': {auxiliary_total:,.0f}")
            return auxiliary_total
    
    if date_col not in dfs_df.columns:
        original_column_label = attrs.get('projection_original_column_label')
        original_values_by_description = attrs.get('projection_original_values_by_description') or {}
        if original_column_label == date_col and isinstance(original_values_by_description, dict):
            desc_col = dfs_df.columns[0]
            for _, row in dfs_df.iterrows():
                desc = str(row[desc_col])
                desc_lower = desc.lower()
                if any(skip_kw in desc_lower for skip_kw in ['小计', 'subtotal', 'sub-total', 'sub total']):
                    continue
                if any(keyword in desc_lower for keyword in ['合计', '总计', 'total']):
                    original_total = original_values_by_description.get(desc)
                    if original_total is not None:
                        if debug:
                            print(
                                f"      Using original projection total for '{date_col}' "
                                f"from annualized dataframe row '{desc}': {original_total:,.0f}"
                            )
                        return original_total
        return None
    
    # Try to find total row
    total_value = _scan_table_total(date_col)
    if total_value is not None:
        if debug:
            desc_col = dfs_df.columns[0]
            total_rows = dfs_df[
                dfs_df[desc_col].astype(str).str.lower().str.contains('合计|总计|total', regex=True, na=False)
                & ~dfs_df[desc_col].astype(str).str.lower().str.contains('小计|subtotal|sub-total|sub total', regex=True, na=False)
            ]
            if not total_rows.empty:
                print(f"      Found total row: '{total_rows.iloc[0][desc_col]}' → value: {total_value:,.0f}")
        return total_value
    
    # Fallback: when no explicit total row exists, the schedule may use a
    # parent-first structure where the first non-breakdown row IS the total.
    # Use projection_totals_by_date if available (populated from sum-detected rows
    # marked as total via _detect_implicit_breakdowns_from_sum), otherwise try
    # the first row with a non-zero value as the best-effort total.
    if date_col in dfs_df.columns:
        desc_col = dfs_df.columns[0]
        row_types = dfs_df.attrs.get("row_types_by_description") or {}
        for _, row in dfs_df.iterrows():
            desc = str(row[desc_col])
            # Skip rows that are breakdowns or explicitly subtotal/total (already checked above)
            if row_types.get(desc) in ("breakdown", "subtotal"):
                continue
            candidate = row[date_col]
            if isinstance(candidate, (int, float)) and abs(candidate) > 0:
                if debug:
                    print(f"      Using first-row fallback total for '{date_col}': {candidate:,.0f} from '{desc}'")
                return float(candidate)

    if debug:
        desc_col = dfs_df.columns[0]
        print(f"      ❌ No total row found (no '合计', '总计', or 'Total' in descriptions)")
        print(f"      Available descriptions: {dfs_df[desc_col].tolist()}")

    return None


def _should_compare_income_statement_as_absolute(account_name: str, category: Optional[str]) -> bool:
    category_text = str(category or "").strip().lower()
    account_text = str(account_name or "").strip().lower()
    if any(keyword in category_text for keyword in ("expense", "loss", "cost")):
        return True
    return any(keyword in account_text for keyword in ("loss", "expense", "cost", "损失", "费用", "成本"))


def _integrity_metadata(dfs_df: Optional[pd.DataFrame]) -> Dict[str, str]:
    if dfs_df is None:
        return {
            'Projection_Stage': '-',
            'Projection_Date': '-',
            'Integrity_Flag': '-',
        }

    integrity = dfs_df.attrs.get('integrity') or {}
    preferred_stage = integrity.get('preferred_stage')
    effective_stage = integrity.get('effective_stage')
    effective_date = integrity.get('effective_date')
    non_zero_rows = integrity.get('non_zero_rows')

    flag = '-'
    if preferred_stage and effective_stage and preferred_stage != effective_stage:
        flag = f'Fallback from {preferred_stage} to {effective_stage}'
    elif non_zero_rows == 0:
        flag = 'Zero-value projection'

    return {
        'Projection_Stage': effective_stage or preferred_stage or '-',
        'Projection_Date': effective_date or '-',
        'Integrity_Flag': flag,
    }


def _latest_column_is_partial_period(date_cols: List[str], min_full_period_months: float = 10.0) -> bool:
    """True if the last column spans clearly fewer months than the prior
    column-to-column cadence — e.g. columns ...-12-31, ...-12-31, ...-01-31,
    where the final "2026-01-31" is a 1-month interim cut, not a year-end.
    Used to widen reconciliation to also check the prior (full-period) column
    rather than relying solely on a partial period that may not tie out the
    same way a full year does.
    """
    if len(date_cols) < 2:
        return False
    last_date = _parse_statement_date_label(date_cols[-1])
    prev_date = _parse_statement_date_label(date_cols[-2])
    if last_date is None or prev_date is None:
        return False
    months_gap = (last_date.year - prev_date.year) * 12 + (last_date.month - prev_date.month)
    return 0 < months_gap < min_full_period_months


def _reconcile_against_prior_period(
    account_name: str,
    date_cols: List[str],
    row: pd.Series,
    dfs_df: Optional[pd.DataFrame],
    tolerance: float,
    materiality_threshold: float,
    debug: bool = False,
    use_absolute: bool = False,
) -> Optional[Dict[str, Any]]:
    """Re-check Financials vs. Tab for the prior (second-to-last) column,
    used when the latest column is a partial period. Returns None if there's
    no prior column, no matching tab data, or the prior check itself can't
    be evaluated. use_absolute mirrors the IS expense-account convention of
    comparing magnitudes rather than signed values.
    """
    if len(date_cols) < 2 or dfs_df is None:
        return None
    prior_date = date_cols[-2]
    if prior_date not in row.index:
        return None
    prior_source_value = row[prior_date]
    prior_dfs_value = get_total_from_dfs(dfs_df, prior_date, debug)
    if prior_dfs_value is None:
        return None
    prior_source_for_comparison = abs(prior_source_value) if use_absolute else prior_source_value
    prior_dfs_for_comparison = abs(prior_dfs_value) if use_absolute else prior_dfs_value
    prior_diff = abs(prior_source_for_comparison - prior_dfs_for_comparison)
    if prior_diff <= tolerance:
        prior_match = True
    elif prior_source_for_comparison != 0 and (prior_diff / abs(prior_source_for_comparison)) <= materiality_threshold:
        prior_match = True
    else:
        prior_match = False
    return {
        "date": prior_date,
        "financials_value": prior_source_value,
        "tab_value": prior_dfs_value,
        "diff": prior_diff,
        "match": prior_match,
    }


def reconcile_financial_statements(
    bs_is_results: Dict,
    dfs: Dict[str, pd.DataFrame],
    mappings_file: str = 'fdd_utils/mappings.yml',
    mappings: Optional[dict] = None,
    tolerance: float = 1.0,
    materiality_threshold: float = 0.005,
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
        mappings: Optional preloaded effective mappings, including dynamic mappings
        tolerance: Absolute tolerance for matching (default: 1.0, allows ±1 difference)
        materiality_threshold: Percentage threshold for immaterial differences (default: 0.005 = 0.5%)
        debug: If True, print debugging information
        
    Returns:
        Tuple of (bs_reconciliation_df, is_reconciliation_df)
        Each DataFrame has columns:
        - Financials_Account: Account name from BS/IS
        - Date: Date column (latest only)
        - Financials_Value: Value from BS/IS (expenses converted to positive)
        - Tab_Account: Actual workbook tab name (e.g., '货币资金', not mapping key 'Cash')
        - Tab_Value: Total value from the schedule tab
        - Match: '✅ Match', '❌ Diff: X', '⚠️ Not Found', or '-' (skipped)
    """
    if debug:
        print("=" * 80)
        print("RECONCILIATION - DEBUG MODE")
        print("=" * 80)
    
    mappings = mappings or load_mappings(mappings_file)
    
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
            # Use the 2 most recent columns for the zero-source check:
            # only skip if both are zero (loosened to keep items with adjacent-period data).
            recent_dates = date_cols[-2:] if len(date_cols) >= 2 else date_cols

            for idx, row in bs_df.iterrows():
                account_name = row['Description']
                source_value = row[latest_date]

                # Skip only when ALL of the most-recent columns are zero
                recent_values = [row[d] for d in recent_dates]
                if all(v == 0 for v in recent_values):
                    integrity_fields = _integrity_metadata(None)
                    bs_recon_rows.append({
                        'Financials_Account': account_name,
                        'Date': latest_date,
                        'Financials_Value': source_value,
                        'Tab_Account': '-',
                        'Tab_Value': '-',
                        'Diff': '-',
                        'Match': '-',
                        'Mapping_Key': '-',
                        'Mapping_Status': 'Zero source',
                        'Mapping_Note': 'Most recent period values are all 0, so schedule mapping was skipped.',
                        **integrity_fields,
                    })
                    continue

                # Flag: latest period is 0 but an adjacent period has data
                zero_with_adjacent = source_value == 0 and any(v != 0 for v in recent_values)

                # Find matching account in dfs (ONLY via mappings.yml)
                dfs_key, dfs_df, category, mapping_key, mapping_status, mapping_note = find_account_in_dfs(account_name, dfs, mappings, debug=debug)

                # Handle skipped accounts (totals/profit lines)
                if dfs_key == 'SKIP':
                    integrity_fields = _integrity_metadata(None)
                    bs_recon_rows.append({
                        'Financials_Account': account_name,
                        'Date': latest_date,
                        'Financials_Value': source_value,
                        'Tab_Account': '-',
                        'Tab_Value': '-',
                        'Diff': '-',
                        'Match': '-',
                        'Mapping_Key': '-',
                        'Mapping_Status': mapping_status,
                        'Mapping_Note': mapping_note,
                        **integrity_fields,
                    })
                    continue

                # Get total value from dfs
                dfs_value = get_total_from_dfs(dfs_df, latest_date, debug) if dfs_df is not None else None

                # When the tab exists but has no column for this date and the
                # source is 0 (zero_with_adjacent), treat the missing value as 0
                # — the schedule simply has no data for that period.
                if dfs_value is None and zero_with_adjacent and dfs_df is not None:
                    dfs_value = 0.0

                # Check match
                if dfs_value is None:
                    match_status = '⚠️ Not Found'
                    difference = None
                else:
                    difference = abs(source_value - dfs_value)

                    # Check if within absolute tolerance
                    if difference <= tolerance:
                        match_status = '⚠️ Match' if zero_with_adjacent else '✅ Match'
                    else:
                        # Check if within materiality threshold (percentage)
                        if source_value != 0:
                            pct_diff = difference / abs(source_value)
                            if pct_diff <= materiality_threshold:
                                match_status = '✅ Immaterial'
                            else:
                                match_status = '❌ Diff'
                        else:
                            match_status = '❌ Diff'

                # The latest column can be an interim/partial-year cut (e.g.
                # "2026-01-31" after year-end columns) that doesn't tie out the
                # same way a full period does. If it shows a real diff, also
                # check the prior (full-period) column — if THAT reconciles,
                # note it rather than reporting an unexplained mismatch.
                if match_status == '❌ Diff' and _latest_column_is_partial_period(date_cols):
                    prior_check = _reconcile_against_prior_period(
                        account_name, date_cols, row, dfs_df, tolerance, materiality_threshold, debug,
                    )
                    if prior_check and prior_check["match"]:
                        match_status = '⚠️ Interim Diff (prior period matches)'
                        mapping_note = (
                            f"{mapping_note} Latest column '{latest_date}' is an interim/partial period and "
                            f"differs from the tab; prior column '{prior_check['date']}' reconciles "
                            f"(Financials={prior_check['financials_value']:,.0f}, Tab={prior_check['tab_value']:,.0f})."
                        ).strip()

                integrity_fields = _integrity_metadata(dfs_df)
                bs_recon_rows.append({
                    'Financials_Account': account_name,
                    'Date': latest_date,
                    'Financials_Value': source_value,
                    'Tab_Account': dfs_key or 'Not Found',
                    'Tab_Value': dfs_value if dfs_value is not None else 0,
                    'Diff': difference if difference is not None else '-',
                    'Match': match_status,
                    'Mapping_Key': mapping_key or '-',
                    'Mapping_Status': mapping_status,
                    'Mapping_Note': mapping_note,
                    **integrity_fields,
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
            # Use the 2 most recent columns for the zero-source check:
            # only skip if both are zero (loosened to keep items with adjacent-period data).
            recent_dates = date_cols[-2:] if len(date_cols) >= 2 else date_cols

            for idx, row in is_df.iterrows():
                account_name = row['Description']
                source_value_raw = row[latest_date]

                # Skip only when ALL of the most-recent columns are zero
                recent_values = [row[d] for d in recent_dates]
                if all(v == 0 for v in recent_values):
                    integrity_fields = _integrity_metadata(None)
                    is_recon_rows.append({
                        'Financials_Account': account_name,
                        'Date': latest_date,
                        'Financials_Value': source_value_raw,
                        'Tab_Account': '-',
                        'Tab_Value': '-',
                        'Diff': '-',
                        'Match': '-',
                        'Mapping_Key': '-',
                        'Mapping_Status': 'Zero source',
                        'Mapping_Note': 'Most recent period values are all 0, so schedule mapping was skipped.',
                        **integrity_fields,
                    })
                    continue

                # Flag: latest period is 0 but an adjacent period has data
                zero_with_adjacent = source_value_raw == 0 and any(v != 0 for v in recent_values)

                # Find matching account in dfs (ONLY via mappings.yml)
                dfs_key, dfs_df, category, mapping_key, mapping_status, mapping_note = find_account_in_dfs(account_name, dfs, mappings, debug=debug)

                # Handle skipped accounts (totals/profit lines)
                if dfs_key == 'SKIP':
                    integrity_fields = _integrity_metadata(None)
                    is_recon_rows.append({
                        'Financials_Account': account_name,
                        'Date': latest_date,
                        'Financials_Value': source_value_raw,
                        'Tab_Account': '-',
                        'Tab_Value': '-',
                        'Diff': '-',
                        'Match': '-',
                        'Mapping_Key': '-',
                        'Mapping_Status': mapping_status,
                        'Mapping_Note': mapping_note,
                        **integrity_fields,
                    })
                    continue

                # Get total value from dfs
                dfs_value = get_total_from_dfs(dfs_df, latest_date, debug) if dfs_df is not None else None

                # When the tab exists but has no column for this date and the
                # source is 0 (zero_with_adjacent), treat the missing value as 0
                # — the schedule simply has no data for that period.
                if dfs_value is None and zero_with_adjacent and dfs_df is not None:
                    dfs_value = 0.0

                # For expense/loss-style lines: keep the original sign in display but compare on absolute value.
                source_for_comparison = source_value_raw
                dfs_for_comparison = dfs_value
                if _should_compare_income_statement_as_absolute(account_name, category):
                    source_for_comparison = abs(source_value_raw)
                    if dfs_for_comparison is not None:
                        dfs_for_comparison = abs(dfs_for_comparison)
                    if debug and (source_value_raw < 0 or (dfs_value is not None and dfs_value < 0)):
                        print(
                            f"    [CONVERT] Compare absolute values: source {source_value_raw:,.0f} → {source_for_comparison:,.0f}; "
                            f"dfs {dfs_value if dfs_value is not None else 'None'} → {dfs_for_comparison if dfs_for_comparison is not None else 'None'}"
                        )

                # Check match
                if dfs_value is None:
                    match_status = '⚠️ Not Found'
                    difference = None
                else:
                    difference = abs(source_for_comparison - dfs_for_comparison)

                    # Check if within absolute tolerance
                    if difference <= tolerance:
                        match_status = '⚠️ Match' if zero_with_adjacent else '✅ Match'
                    else:
                        # Check if within materiality threshold (percentage)
                        if source_for_comparison != 0:
                            pct_diff = difference / abs(source_for_comparison)
                            if pct_diff <= materiality_threshold:
                                match_status = '✅ Immaterial'
                            else:
                                match_status = '❌ Diff'
                        else:
                            match_status = '❌ Diff'

                # See the matching BS comment above: the latest column can be
                # an interim/partial period. If it shows a real diff, also
                # check the prior (full-period) column before reporting an
                # unexplained mismatch.
                if match_status == '❌ Diff' and _latest_column_is_partial_period(date_cols):
                    prior_check = _reconcile_against_prior_period(
                        account_name, date_cols, row, dfs_df, tolerance, materiality_threshold, debug,
                        use_absolute=_should_compare_income_statement_as_absolute(account_name, category),
                    )
                    if prior_check and prior_check["match"]:
                        match_status = '⚠️ Interim Diff (prior period matches)'
                        mapping_note = (
                            f"{mapping_note} Latest column '{latest_date}' is an interim/partial period and "
                            f"differs from the tab; prior column '{prior_check['date']}' reconciles "
                            f"(Financials={prior_check['financials_value']:,.0f}, Tab={prior_check['tab_value']:,.0f})."
                        ).strip()

                integrity_fields = _integrity_metadata(dfs_df)
                is_recon_rows.append({
                    'Financials_Account': account_name,
                    'Date': latest_date,
                    'Financials_Value': source_value_raw,  # Keep original negative value
                    'Tab_Account': dfs_key or 'Not Found',
                    'Tab_Value': dfs_value if dfs_value is not None else 0,
                    'Diff': difference if difference is not None else '-',
                    'Match': match_status,
                    'Mapping_Key': mapping_key or '-',
                    'Mapping_Status': mapping_status,
                    'Mapping_Note': mapping_note,
                    **integrity_fields,
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
            df_display['Financials_Value'] = df_display['Financials_Value'].apply(
                lambda x: f"{x:,.0f}" if isinstance(x, (int, float)) else x
            )
            df_display['Tab_Value'] = df_display['Tab_Value'].apply(
                lambda x: f"{x:,.0f}" if isinstance(x, (int, float)) else x
            )
            df_display['Diff'] = df_display['Diff'].apply(
                lambda x: f"{x:,.0f}" if isinstance(x, (int, float)) else x
            )
            
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
            df_display['Financials_Value'] = df_display['Financials_Value'].apply(
                lambda x: f"{x:,.0f}" if isinstance(x, (int, float)) else x
            )
            df_display['Tab_Value'] = df_display['Tab_Value'].apply(
                lambda x: f"{x:,.0f}" if isinstance(x, (int, float)) else x
            )
            df_display['Diff'] = df_display['Diff'].apply(
                lambda x: f"{x:,.0f}" if isinstance(x, (int, float)) else x
            )
            
            print(df_display.to_string(index=False))
        else:
            print("✅ All accounts match perfectly!")
    
    print("\n" + "=" * 100)


# Example usage
if __name__ == "__main__":
    print("=" * 80)
    print("RECONCILIATION EXAMPLE")
    print("=" * 80)
    
    example = find_reconciliation_example()
    if not example:
        raise FileNotFoundError(
            "No local reconciliation example workbook with a 'Financials' sheet was found."
        )

    databook_path = example["workbook_path"]
    sheet_name = example["sheet_name"]
    entity_name = example["entity_name"]
    print(f"Using workbook: {databook_path}")
    print(f"Using financial sheet: {sheet_name}")
    
    # Source 1: Financial extraction (BS/IS from single sheet)
    bs_is_results = extract_balance_sheet_and_income_statement(
        workbook_path=databook_path,
        sheet_name=sheet_name,
        debug=False
    )
    
    # Source 2: DFS extraction (account by account)
    dfs, workbook_list, _, language = extract_data_from_excel(
        databook_path=databook_path,
        entity_name=entity_name,
        mode="All"
    )
    
    # Reconcile
    bs_recon, is_recon = reconcile_financial_statements(
        bs_is_results=bs_is_results,
        dfs=dfs,
        tolerance=1.0,
        materiality_threshold=0.005,  # 0.5% materiality
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
# --- end workbook/reconcile.py ---

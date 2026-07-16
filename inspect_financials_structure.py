"""Check whether a databook's "Financials" sheet (the one
embed_financial_tables() actually reads for the embedded BS/IS table) has
category-header rows ("流动资产" / "非流动负债" / etc. -- a label with no
figures at all) baked into it, or whether it goes straight from one line
item to the next subtotal with no header row in between.

Prints ONLY row labels and a blank/has-data flag per row -- no financial
figures at all, safe to paste back even for a real client databook.

Usage:
    python inspect_financials_structure.py "databook.xlsx" [--sheet "Financials"]
"""
import argparse
import sys

import pandas as pd

from fdd_utils.workbook import extract_balance_sheet_and_income_statement


def _is_blank(value) -> bool:
    if pd.isna(value):
        return True
    if isinstance(value, str) and not value.strip():
        return True
    return False


def _report(label: str, df) -> None:
    print(f"\n=== {label} ({len(df) if df is not None else 0} rows) ===")
    if df is None or df.empty:
        print("  (no data)")
        return
    numeric_cols = list(df.columns[1:])
    header_rows = []
    for row_idx, row in df.iterrows():
        first_col = str(row.iloc[0]).strip()
        if not first_col:
            continue
        all_blank = all(_is_blank(row[col]) for col in numeric_cols)
        marker = "  [HEADER -- label only, no figures]" if all_blank else ""
        print(f"  {first_col}{marker}")
        if all_blank:
            header_rows.append(first_col)
    print(f"\n  -> {len(header_rows)} category-header-shaped row(s) found: {header_rows}")


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("databook", help="Path to the .xlsx databook")
    ap.add_argument("--sheet", default="Financials", help="Financials sheet name (default: 'Financials')")
    args = ap.parse_args()

    print(f"Reading sheet: {args.sheet!r} (pass --sheet to use a different tab name)")
    result = extract_balance_sheet_and_income_statement(args.databook, args.sheet, debug=False)
    if not result:
        print("Extraction returned nothing.")
        return 1

    _report("Balance Sheet", result.get("balance_sheet"))
    _report("Income Statement", result.get("income_statement"))
    return 0


if __name__ == "__main__":
    sys.exit(main())

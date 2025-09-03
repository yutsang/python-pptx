#!/usr/bin/env python3
"""
Simple script to show raw Excel table content without analysis.
Just displays what's actually in each Excel tab/sheet.
"""

import pandas as pd
import sys
import os

def show_excel_content_simple(excel_file='databook.xlsx'):
    """Show raw content of all Excel sheets without any analysis."""

    if not os.path.exists(excel_file):
        print(f"❌ File '{excel_file}' not found!")
        return

    print(f"📊 EXCEL FILE: {excel_file}")
    print("="*80)

    try:
        # Read all sheets
        excel_data = pd.read_excel(excel_file, sheet_name=None, header=None)

        for sheet_name, df in excel_data.items():
            print(f"\n🔍 SHEET: '{sheet_name}'")
            print(f"   Size: {df.shape[0]} rows x {df.shape[1]} columns")
            print("-" * 50)

            # Show first 10 rows or all if less than 10
            rows_to_show = min(10, len(df))

            for i in range(rows_to_show):
                row_data = df.iloc[i].fillna('')  # Replace NaN with empty string
                row_values = [str(cell).strip() for cell in row_data.values]

                # Show row number and content
                print("2d")

                # If row has actual data, show it
                if any(row_values):  # Only show rows that have some content
                    # Truncate long values for readability
                    display_values = []
                    for val in row_values:
                        if len(val) > 30:
                            display_values.append(val[:27] + "...")
                        else:
                            display_values.append(val)

                    print(f"   Data: {' | '.join(display_values)}")

            # Show if there are more rows
            if len(df) > rows_to_show:
                print(f"   ... and {len(df) - rows_to_show} more rows")

            print()

    except Exception as e:
        print(f"❌ Error reading Excel file: {e}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        excel_file = sys.argv[1]
    else:
        excel_file = 'databook.xlsx'

    show_excel_content_simple(excel_file)

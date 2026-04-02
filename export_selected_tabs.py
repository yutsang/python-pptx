#!/usr/bin/env python3
"""
CLI wrapper for exporting selected workbook tabs to plain text.
"""

from __future__ import annotations

import argparse

from fdd_utils.workbook import export_selected_tabs_to_file


def main() -> int:
    parser = argparse.ArgumentParser(description="Export selected workbook tabs to a text file.")
    parser.add_argument("--workbook", required=True, help="Path to the Excel workbook.")
    parser.add_argument("--tabs", nargs="+", required=True, help="One or more sheet names to export.")
    parser.add_argument("--output", help="Optional output text file path.")
    args = parser.parse_args()

    output_path = export_selected_tabs_to_file(
        workbook_path=args.workbook,
        selected_tabs=args.tabs,
        output_path=args.output,
    )
    print(output_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
import argparse

from fdd_utils.workbook import export_selected_tabs_to_file


def parse_args():
    parser = argparse.ArgumentParser(
        description="Export selected Excel tabs into a compact plain-text file."
    )
    parser.add_argument(
        "--workbook",
        required=True,
        help="Path to the source Excel workbook.",
    )
    parser.add_argument(
        "--tabs",
        required=True,
        help="Comma-separated list of worksheet names to export.",
    )
    parser.add_argument(
        "--output",
        help="Optional output text file path. Defaults beside the workbook.",
    )
    return parser.parse_args()


def main():
    args = parse_args()
    tabs = [tab.strip() for tab in args.tabs.split(",") if tab.strip()]
    output_path = export_selected_tabs_to_file(
        workbook_path=args.workbook,
        selected_tabs=tabs,
        output_path=args.output,
    )
    print(output_path)


if __name__ == "__main__":
    main()

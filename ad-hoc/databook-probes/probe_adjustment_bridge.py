"""Print the audit and indicative adjustment entries each account is handed.

    python ad-hoc/databook-probes/probe_adjustment_bridge.py <databook.xlsx>

Free and quick: runs the workbook processing only, no model call. Lists, per
account, the lines authority.adjustment_bridge read off the period bridge
sheets (columns 审定数 / 示意性调整后) and the ADJ register -- the block the
Generator and Auditor see as 【本科目的审计调整及示意性调整（工作底稿）】.
An account missing here that has adjustments in the workbook means its row
label on the bridge sheet does not match the account name. Output quotes
client figures; do not paste it anywhere public.
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..")))

import pandas as pd  # noqa: E402

from fdd_utils.workbook import process_workbook_data  # noqa: E402
from fdd_utils.workbook.authority import BRIDGE_KEY  # noqa: E402


def main() -> int:
    if len(sys.argv) != 2:
        print(__doc__)
        return 1
    book = sys.argv[1]
    sheet = next((n for n in pd.ExcelFile(book).sheet_names if "financial" in n.lower()), None)
    state = process_workbook_data(temp_path=book, entity_name="", selected_sheet=sheet, debug=False)
    dfs = state.get("dfs") or {}
    bridged = {k: f.attrs.get(BRIDGE_KEY) for k, f in dfs.items() if f.attrs.get(BRIDGE_KEY)}
    print(f"{len(bridged)} of {len(dfs)} account(s) carry adjustment entries\n")
    for account in dfs:
        lines = (bridged.get(account) or {}).get("lines") or []
        print(f"{account}: {'(none)' if not lines else ''}")
        for line in lines:
            print(f"    {line}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

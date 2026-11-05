"""Print the audit and indicative adjustment entries each account is handed.

    python ad-hoc/databook-probes/probe_adjustment_bridge.py <databook.xlsx>
    python ad-hoc/databook-probes/probe_adjustment_bridge.py <databook.xlsx> --register 投资收益

--register prints, for every place an adjustment-register sheet carries an
amount against that account, the register rows from 8 above to 2 below with
their texts and figures -- the layout the description is read from.

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
import re
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..")))

import pandas as pd  # noqa: E402

from fdd_utils.workbook import process_workbook_data  # noqa: E402
from fdd_utils.workbook.authority import BRIDGE_KEY  # noqa: E402


def main() -> int:
    if len(sys.argv) not in (2, 4):
        print(__doc__)
        return 1
    book = sys.argv[1]
    if len(sys.argv) == 4 and sys.argv[2] == "--register":
        return dump_register(book, sys.argv[3])
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


def dump_register(book: str, account: str) -> int:
    from fdd_utils.workbook.digest import DIGEST_KEY  # noqa: E402
    sheet = next((n for n in pd.ExcelFile(book).sheet_names if "financial" in n.lower()), None)
    state = process_workbook_data(temp_path=book, entity_name="", selected_sheet=sheet, debug=False)
    digest = (state.get("resolution") or {}).get(DIGEST_KEY) or {}
    for name, reg in (digest.get("sheets") or {}).items():
        if not re.search(r"ADJ|调整", name, re.I):
            continue
        rows = {int(k): str(v).strip() for k, v in (reg.get("row_labels") or {}).items()}
        cells = reg.get("cells") or {}
        values: dict = {}
        for r, c, v in zip(cells.get("r", []), cells.get("c", []), cells.get("v", [])):
            values.setdefault(int(r), []).append((int(c), v))
        texts: dict = {}
        for t in reg.get("texts") or []:
            texts.setdefault(int(t[0]), []).append((int(t[1]), str(t[2]).strip()))
        hits = [r for r, label in rows.items() if label == account and values.get(r)]
        for hit in hits:
            print(f"\n=== {name}: row {hit} labelled {account} ===")
            for r in range(hit - 8, hit + 3):
                if r in texts or r in values:
                    print(f"  row {r:>4}{' <--' if r == hit else '    '} texts={sorted(texts.get(r, []))} "
                          f"values={sorted(values.get(r, []))}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

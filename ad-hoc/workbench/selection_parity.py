"""Old vs new account selection on a real workbook. Deterministic, zero tokens.

    python ad-hoc/workbench/selection_parity.py <databook.xlsx> [--sheet NAME]

Runs process_workbook_data once, then asks both the pre-selection-module
function (kept below, verbatim) and the current one which accounts go to
commentary. The current one is narrower by construction -- it may only refuse
what the old one let in through the history-activity step -- so the script
exits 1 if it ever admits something the old one did not, and prints every
account whose fate changed with the route and reason.
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))


def legacy_derive(reconciliation, available_keys, dfs=None):
    keys = [k for k in dict.fromkeys(str(k).strip() for k in available_keys) if k]
    key_set, seen, matched = set(keys), set(), []
    bs_inc = {"✅ Match", "⚠️ Match", "✅ Immaterial"}
    is_inc = bs_inc | {"❌ Diff"}
    for idx, frame in enumerate(reconciliation or ()):
        if frame is None or frame.empty:
            continue
        for value in frame[frame["Match"].isin(is_inc if idx == 1 else bs_inc)]["Tab_Account"].tolist():
            key = str(value or "").strip()
            if key and key != "-" and key not in seen and key in key_set:
                seen.add(key)
                matched.append(key)
    for key in keys:
        if key in seen:
            continue
        attrs = getattr(dfs.get(key), "attrs", {}) or {}
        per = attrs.get("any_period_nonzero_by_description") or {}
        if isinstance(per, dict) and any(bool(v) for v in per.values()):
            seen.add(key)
            matched.append(key)
    return matched


def financials_sheets(path: str):
    from fdd_utils.workbook.visibility import find_visible_financials_sheet
    found = find_visible_financials_sheet(path)      # a hidden sheet is never chosen
    return [found] if found else []


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbook")
    ap.add_argument("--sheet", default=None, help="Financials sheet (default: first match)")
    args = ap.parse_args()

    from fdd_utils.ui.views import explain_reconciliation_selection
    from fdd_utils.workbook import process_workbook_data

    sheets = [args.sheet] if args.sheet else financials_sheets(args.workbook)[:1]
    state = process_workbook_data(temp_path=args.workbook, entity_name="", selected_sheet=sheets[0] if sheets else None)
    dfs, recon = state["dfs"], state["reconciliation"]
    old = legacy_derive(recon, dfs.keys(), dfs=dfs)
    result = explain_reconciliation_selection(recon, dfs.keys(), state.get("resolution"), dfs=dfs)
    new = result.admitted_keys
    print(f"{Path(args.workbook).name}: {len(dfs)} tab(s), sheet={sheets[:1]}")
    print(f"  admitted old={len(old)} new={len(new)}  (bs_is_results source: "
          f"{'financials' if sheets else 'none -> synthesized'})")
    dropped = [k for k in old if k not in new]
    added = [k for k in new if k not in old]
    by_key = result.by_key()
    for key in dropped:
        d = by_key[key]
        print(f"  NO LONGER ADMITTED  {key}: {d.reason} -- {d.detail}")
    for key in added:
        print(f"  NEWLY ADMITTED (bug) {key}")
    review = [d for d in result.decisions if d.needs_review]
    for d in review:
        print(f"  admitted for review  {d.key}: {d.route} ({d.reason})")
    if result.unmatched_financials:
        print("  Financials lines with no tab:", ", ".join(u["financials_account"] for u in result.unmatched_financials))
    return 1 if added else 0


if __name__ == "__main__":
    sys.exit(main())

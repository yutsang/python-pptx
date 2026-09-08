"""Why did BS/IS extraction fail on a roll-up ("主表") workbook?

A real portfolio run lost both statements on every entity with nothing but

    Error extracting financial data: sequence item 0: expected str instance, float found

and no frame, because the handler only logged the traceback under debug. That
is now fixed, but the fastest way to see the cause is to call the extractor
directly: no AI, no export, seconds, free.

    python ad-hoc/databook-probes/probe_financials_extraction.py "<roll-up.xlsx>"
    python ad-hoc/databook-probes/probe_financials_extraction.py "<roll-up.xlsx>" --sheet "<entity>Financials"

Run it from the repo root. It puts the root on sys.path itself, so no
PYTHONPATH is needed -- the other ad-hoc scripts all require `PYTHONPATH=.`,
which is Unix shell syntax that cmd.exe rejects outright.

With no --sheet it tries every sheet whose name ends in "Financials", which is
how a roll-up file names its one sheet per entity. For each it prints either the
extracted shape or the FULL traceback of whatever went wrong.

Output names sheets and prints row counts, so it carries client identifiers.
Fine to paste into a working conversation; do not put it anywhere public.
"""
from __future__ import annotations

import argparse
import os
import sys
import traceback

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import pandas as pd

from fdd_utils.workbook.statements import extract_balance_sheet_and_income_statement

RULE = "=" * 78


def probe(path: str, sheet: str) -> bool:
    print(f"\n{RULE}\nSHEET  {sheet}\n{RULE}")
    try:
        raw = pd.read_excel(path, sheet_name=sheet, header=None, engine="openpyxl")
        print(f"  raw shape       {raw.shape}")
        print(f"  column dtypes   {sorted({str(d) for d in raw.dtypes})}")
    except Exception:
        print("  could not even read the sheet:")
        traceback.print_exc()
        return False

    try:
        out = extract_balance_sheet_and_income_statement(path, sheet, debug=False)
    except Exception:
        # The production path swallows this into results; re-raise here so the
        # frame is visible. Anything that escapes to this handler is a bug in
        # the extractor rather than in the workbook.
        print("  extractor RAISED past its own handler:")
        traceback.print_exc()
        return False

    bs, is_ = out.get("balance_sheet"), out.get("income_statement")
    print(f"  project_name    {out.get('project_name')!r}")
    print(f"  balance_sheet   {None if bs is None else f'{len(bs)} rows x {len(bs.columns)} cols'}")
    print(f"  income_stmt     {None if is_ is None else f'{len(is_)} rows x {len(is_.columns)} cols'}")
    if bs is None and is_ is None:
        print("  *** BOTH EMPTY — the extractor caught something internally.")
        print("  *** Re-run with the line below to see the traceback it logged:")
        print("  ***   PYTHONPATH=. python -c \"import logging;logging.basicConfig("
              "level=logging.DEBUG);from fdd_utils.workbook.statements import *;"
              f"extract_balance_sheet_and_income_statement({path!r},{sheet!r})\"")
        return False
    return True


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbook", help="the roll-up (主表) .xlsx")
    ap.add_argument("--sheet", action="append",
                    help="sheet to probe; repeatable. Default: every *Financials sheet")
    args = ap.parse_args()

    names = pd.ExcelFile(args.workbook, engine="openpyxl").sheet_names
    targets = args.sheet or [n for n in names if str(n).strip().endswith("Financials")]
    if not targets:
        print(f"no sheet name ends in 'Financials'. Sheets present: {names}")
        return 2

    print(f"{RULE}\n{args.workbook}\n{len(names)} sheet(s), probing {len(targets)}\n{RULE}")
    ok = sum(probe(args.workbook, sheet) for sheet in targets)
    print(f"\n{RULE}\n{ok}/{len(targets)} sheet(s) extracted something. "
          f"Paste the first traceback above.\n{RULE}")
    return 0 if ok == len(targets) else 1


if __name__ == "__main__":
    sys.exit(main())

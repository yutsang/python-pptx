#!/usr/bin/env python3
"""Finds the ready-made detail tables in a databook -- the ones a report
paragraph hands off to with '明细如下：'.

The reference deck tabulates its income-statement expense accounts instead
of narrating them (管理费用, 营业成本, 财务费用 all end '明细如下：' and stop).
Building that means knowing where those tables already live and what shape
they are, rather than assembling one from scratch. This looks in the two
places they can be:

  A. INSIDE an account's own sheet -- a second block below or beside the
     main schedule, e.g. a per-period breakdown by cost category;
  B. in a DEDICATED sheet named after the account, e.g. '成本明细-物业费',
     'AR SUPPORTING', '租户保证金核对'.

For each candidate it reports where it starts, its shape, its header row,
its row labels, and whether the period columns line up with the account's
own -- which is what decides whether it can be dropped into a slide as-is.

Read-only.

Usage:
    python inspect_detail_tables.py "for_test/x.xlsx"
    python inspect_detail_tables.py "for_test/x.xlsx" --account 管理费用 --rows 40
"""
import argparse
import re
import sys
import warnings

warnings.filterwarnings("ignore")
sys.path.insert(0, ".")

import pandas as pd

# Sheets that are supporting detail rather than a primary account schedule.
_SUPPORTING_NAME_HINTS = (
    "明细", "明細", "supporting", "support", "核对", "核對", "清单", "清單",
    "detail", "breakdown", "je", "调整", "調整",
)
# Navigation / structural sheets that are never a detail table.
_SKIP_NAME_HINTS = ("-->", ">>", "_tm_", "upslide", "封面", "目录", "目錄")
_DATE_HDR_RE = re.compile(r"(20\d{2})\s*年|20\d{2}-\d{2}|\b20\d{2}\b|1-\d+月|\d+M\d{2}")


def _cells(df, idx):
    return [str(v).strip() for v in df.iloc[idx].tolist() if str(v).strip() not in ("", "nan")]


def _looks_header(df, idx) -> bool:
    """A row carrying several period labels -- the top of a periodised table."""
    vals = _cells(df, idx)
    if len(vals) < 3:
        return False
    hits = sum(1 for v in vals if _DATE_HDR_RE.search(v))
    return hits >= 2


def _numeric_share(df, idx) -> float:
    vals = df.iloc[idx].tolist()
    nums = sum(1 for v in vals if isinstance(v, (int, float)) and not pd.isna(v))
    filled = sum(1 for v in vals if str(v).strip() not in ("", "nan"))
    return nums / filled if filled else 0.0


def scan_sheet_for_blocks(df, sheet_name, max_rows=200):
    """Header rows in a sheet, with the run of data rows under each. More than
    one means the sheet holds a main schedule AND a separate detail block."""
    blocks = []
    n = min(len(df), max_rows)
    idx = 0
    while idx < n:
        if _looks_header(df, idx):
            start = idx
            data_rows = 0
            j = idx + 1
            blank = 0
            while j < n and blank < 2:
                if not _cells(df, j):
                    blank += 1
                    j += 1
                    continue
                blank = 0
                if _looks_header(df, j):
                    break
                if _numeric_share(df, j) > 0.2:
                    data_rows += 1
                j += 1
            if data_rows >= 2:
                blocks.append({"header_row": start, "end_row": j, "data_rows": data_rows,
                               "header": _cells(df, start)[:8]})
            idx = max(j, idx + 1)
        else:
            idx += 1
    return blocks


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("path", help="path to the databook .xlsx")
    ap.add_argument("--account", default=None, help="only look at this account's sheet")
    ap.add_argument("--rows", type=int, default=25, help="max detail rows to print per table")
    args = ap.parse_args()

    xl = pd.ExcelFile(args.path)
    names = xl.sheet_names
    print(f"Workbook: {args.path!r}  ({len(names)} sheets)\n")

    account_sheets = [
        s for s in names
        if not any(h in s.lower() for h in _SKIP_NAME_HINTS)
        and not any(h in s.lower() for h in _SUPPORTING_NAME_HINTS)
        and s not in ("Financials", "Overview", "TB", "mapping", "NAV")
    ]
    supporting_sheets = [
        s for s in names
        if any(h in s.lower() for h in _SUPPORTING_NAME_HINTS)
        and not any(h in s.lower() for h in _SKIP_NAME_HINTS)
    ]

    print("=" * 78)
    print("[A] DEDICATED SUPPORTING SHEETS  (candidate ready-made detail tables)")
    print("=" * 78)
    if not supporting_sheets:
        print("  none found by name")
    for s in supporting_sheets:
        try:
            df = pd.read_excel(args.path, sheet_name=s, header=None)
        except Exception as exc:
            print(f"  {s!r}: could not read ({type(exc).__name__})")
            continue
        blocks = scan_sheet_for_blocks(df, s)
        print(f"\n  {s!r}  shape={df.shape}  periodised block(s)={len(blocks)}")
        for b in blocks[:2]:
            print(f"    header row {b['header_row']}: {b['header']}")
            print(f"    {b['data_rows']} data row(s), ends row {b['end_row']}")
            labels = []
            for r in range(b["header_row"] + 1, min(b["end_row"], b["header_row"] + 1 + args.rows)):
                c = _cells(df, r)
                if c:
                    labels.append(c[0][:26])
            print(f"    row labels: {labels[:12]}")

    print("\n" + "=" * 78)
    print("[B] SECOND BLOCKS INSIDE AN ACCOUNT'S OWN SHEET")
    print("=" * 78)
    targets = [args.account] if args.account else account_sheets
    found_any = False
    for s in targets:
        if s not in names:
            print(f"  {s!r}: sheet not found")
            continue
        try:
            df = pd.read_excel(args.path, sheet_name=s, header=None)
        except Exception:
            continue
        blocks = scan_sheet_for_blocks(df, s)
        if len(blocks) < 2 and not args.account:
            continue  # one block = just the main schedule, nothing extra
        found_any = True
        print(f"\n  {s!r}  shape={df.shape}  periodised block(s)={len(blocks)}")
        for i, b in enumerate(blocks, 1):
            role = "main schedule" if i == 1 else "ADDITIONAL detail block"
            print(f"    block {i} ({role}): header row {b['header_row']}, "
                  f"{b['data_rows']} data row(s)")
            print(f"      header: {b['header']}")
            labels = []
            for r in range(b["header_row"] + 1, min(b["end_row"], b["header_row"] + 1 + args.rows)):
                c = _cells(df, r)
                if c:
                    labels.append(c[0][:26])
            print(f"      row labels: {labels[:12]}")
    if not found_any and not args.account:
        print("  no account sheet carries a second periodised block")

    print("\n" + "=" * 78)
    print("READ THIS BEFORE BUILDING ANYTHING")
    print("=" * 78)
    print("  The deck tabulates 管理费用 / 营业成本 / 财务费用. Check above whether a")
    print("  ready-made table for those exists, and if so WHERE -- a dedicated sheet")
    print("  means referencing it, a second block in the account's own sheet means")
    print("  extracting it, and neither means the table is assembled in the report")
    print("  from the schedule's own detail rows. Those are three different builds,")
    print("  so this needs answering before any of them starts.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

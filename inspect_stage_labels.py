#!/usr/bin/env python3
"""Dumps the raw stage-label header text of every breakdown tab, next to
what canonical_stage_label() resolves it to -- to confirm or rule out a
suspected label-matching bug before touching any pipeline code.

Background: a real production run showed 12/14 BS accounts landing on
'Mgt acc' instead of the preferred 'Indicative adjusted', when the
expected/baseline behaviour (Kunshan) is 100% Indicative adjusted with
zero fallbacks. An investigation confirmed `_choose_projection`'s fallback
logic itself is correct (only fires when Indicative adjusted is genuinely
all-zero) -- but found two gaps in `canonical_stage_label` (fdd_utils/
workbook.py) that could produce exactly this uniform symptom:
  1. CANONICAL_STAGE_LABELS lists both fully-simplified '示意性调整后' and
     fully-traditional '示意性調整後', but NOT mixed-glyph variants like
     '示意性调整後' (simplified 调整 + traditional 後) or '示意性調整后'
     (traditional 調整 + simplified 后) -- these mixed forms fall through
     to matching "Indicative adjustment" (the DELTA stage) instead, or
     None entirely.
  2. Any stage header longer than 35 characters returns None outright
     (e.g. a long bilingual header with notes appended).

This script does NOT fix anything -- it just shows the raw text next to
the resolved canonical stage, so we can see directly whether a populated
Indicative-adjusted column exists under a label canonical_stage_label
fails to recognise, on the ACTUAL file that showed the problem.

Follow-up finding (same session): on a real file, EVERY stage label
resolved correctly (no unrecognized text at all), yet the final extracted
data still only had 'Mgt acc' columns -- 'Indicative adjusted' was
entirely ABSENT, not just zero-valued. An investigation traced this to
normalize_financial_schedule's column-building loop (fdd_utils/workbook.py
~3928-3944): a column only survives if it has BOTH a resolved stage label
AND a parseable date in the date row at that same column -- so this script
now ALSO dumps the date row at exactly the columns where a stage label was
found, to show directly whether the OTHER 4 stage blocks (Audit adjustment/
Audited/Indicative adjustment/Indicative adjusted) have their own dates or
share a blank cell (a "write the dates once under the first block" layout,
which the code would then correctly, silently drop every later block for).

Usage:
    python inspect_stage_labels.py "databooks/xx.xlsx"
    python inspect_stage_labels.py "databooks/xx.xlsx" --sheet 长期借款
"""
import argparse
import sys

from fdd_utils.workbook import (
    load_workbook_frames,
    canonical_stage_label,
    _cell_text,
    _date_row_index,
    stage_row_indices,
    parse_date,
)


def dump_sheet_stage_labels(df, sheet_name: str):
    rows = stage_row_indices(df, parse_date)
    if not rows:
        print(f"  (no stage row detected on {sheet_name!r})")
        return 0
    unresolved = 0
    missing_dates = 0
    for row_idx in rows:
        print(f"  stage row {row_idx}:")
        date_row_idx = _date_row_index(df, row_idx)
        print(f"  paired date row: {date_row_idx}")
        for col in range(df.shape[1]):
            raw_val = df.iat[row_idx, col]
            raw_text = _cell_text(raw_val)
            if not raw_text or not raw_text.strip():
                continue
            resolved = canonical_stage_label(raw_val)
            flag = "" if resolved else "  ⚠️ UNRECOGNIZED"
            date_text = "(no date row found)"
            date_flag = ""
            if date_row_idx is not None:
                date_raw = df.iat[date_row_idx, col]
                date_text = _cell_text(date_raw)
                parsed = parse_date(date_raw)
                if not parsed:
                    date_flag = "  ⚠️ NO PARSEABLE DATE HERE -- this column will be DROPPED entirely"
                    missing_dates += 1
            print(f"    col {col}: stage={raw_text!r} -> {resolved!r}{flag}   "
                  f"| date={date_text!r}{date_flag}")
            if not resolved:
                unresolved += 1
    if missing_dates:
        print(f"  ⚠️ {missing_dates} stage-labeled column(s) have no parseable date paired with them "
              f"-- these silently produce ZERO output columns regardless of the stage label being fine.")
    return unresolved


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("path", help="path to the databook .xlsx")
    ap.add_argument("--sheet", default=None, help="only check this one sheet (default: all sheets)")
    args = ap.parse_args()

    print(f"Loading {args.path!r}...")
    frames = load_workbook_frames(args.path)

    sheet_names = [args.sheet] if args.sheet else list(frames.keys())
    if args.sheet and args.sheet not in frames:
        print(f"❌ sheet {args.sheet!r} not found. Available: {list(frames.keys())}")
        return 1

    total_unresolved = 0
    for name in sheet_names:
        print(f"\n--- {name} ---")
        total_unresolved += dump_sheet_stage_labels(frames[name], name)

    print(f"\n{'='*70}")
    if total_unresolved:
        print(f"⚠️ {total_unresolved} stage-row cell(s) with text that canonical_stage_label "
              f"could NOT resolve to any known stage -- these are exactly the kind of cell "
              f"that would cause a silent fallback away from Indicative adjusted.")
    else:
        print("✅ Every non-empty stage-row cell resolved to a known canonical stage.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

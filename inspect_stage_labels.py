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
    stage_row_indices,
    parse_date,
)


def dump_sheet_stage_labels(df, sheet_name: str):
    rows = stage_row_indices(df, parse_date)
    if not rows:
        print(f"  (no stage row detected on {sheet_name!r})")
        return 0
    unresolved = 0
    for row_idx in rows:
        print(f"  stage row {row_idx}:")
        for col in range(df.shape[1]):
            raw_val = df.iat[row_idx, col]
            raw_text = _cell_text(raw_val)
            if not raw_text or not raw_text.strip():
                continue
            resolved = canonical_stage_label(raw_val)
            flag = "" if resolved else "  ⚠️ UNRECOGNIZED"
            print(f"    col {col}: {raw_text!r} -> {resolved!r}{flag}")
            if not resolved:
                unresolved += 1
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

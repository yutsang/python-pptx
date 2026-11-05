#!/usr/bin/env python3
"""Is a detail table simply TALLER than the slot it has to share?

inspect_pptx.py reports a slot at 153% fill and inspect_table_bands.py reports
which paragraph a table landed on. Neither answers the question that decides
what to do about it: whether the packer made a bad call that better packing
could fix, or whether a 44-row table cannot fit on a page whatever the packer
does. Those need opposite fixes -- retune the packer, or cap the table -- and
guessing between them has cost real re-runs.

So this prints, for every table in an exported deck: its drawn height, the
height of the text slot it is floated over, the ratio, and how many of its rows
would have fitted. A table over 100% of its slot is a content problem: no
packing decision can place it, and the only lever is fewer rows.

    python inspect_table_height.py "<deck.pptx>"
    python inspect_table_height.py "<deck.pptx>" --slide 2

Reads the .pptx only. Seconds, free, no AI. It measures the file as saved --
PowerPoint auto-grows a row whose cell wraps, so a table flagged for wrapping
is TALLER on screen than the height reported here, never shorter.

Output names accounts and slides, so it carries client identifiers. Fine to
paste into a working conversation; do not put it anywhere public.
"""
from __future__ import annotations

import argparse
import os
import sys

from pptx import Presentation

RULE = "=" * 78
EMU_PER_PT = 12700.0


def _pt(emu) -> float:
    return (emu or 0) / EMU_PER_PT


def _overlaps_vertically(a_top: float, a_bot: float, b_top: float, b_bot: float) -> bool:
    return a_top < b_bot and b_top < a_bot


def report(path: str, only_slide: int | None) -> int:
    prs = Presentation(path)
    slot_h = _pt(prs.slide_height)
    print(f"{RULE}\n{os.path.basename(path)}   slide height {slot_h:.0f}pt\n{RULE}")

    worst = 0.0
    over = 0
    for index, slide in enumerate(prs.slides, start=1):
        if only_slide and index != only_slide:
            continue
        tables = [s for s in slide.shapes if getattr(s, "has_table", False)]
        if not tables:
            continue
        bodies = [
            s for s in slide.shapes
            if getattr(s, "has_text_frame", False) and not getattr(s, "has_table", False)
            and "mainbullets" in (s.name or "").lower()
        ]
        print(f"\n--- Slide {index} ---")
        for shape in tables:
            table = shape.table
            rows = len(table.rows)
            row_pts = [_pt(r.height) for r in table.rows]
            drawn = sum(row_pts)
            # The shape's own height is what python-pptx wrote; the row heights
            # are what PowerPoint actually lays out, and they disagree whenever
            # a row was set shorter than its text needs. Report both -- the
            # larger is the one that decides whether it fits.
            declared = _pt(shape.height)
            top, bottom = _pt(shape.top), _pt(shape.top) + max(drawn, declared)

            host = None
            for body in bodies:
                b_top, b_bot = _pt(body.top), _pt(body.top) + _pt(body.height)
                if _overlaps_vertically(top, bottom, b_top, b_bot):
                    host = body
                    break

            print(f"  [{shape.name}] {rows} rows x {len(table.columns)} cols")
            print(f"      rows sum to {drawn:7.1f}pt   shape declares {declared:7.1f}pt"
                  f"   {'(rows win)' if drawn > declared + 0.5 else ''}")
            if host is None:
                print("      no textMainBullets slot overlaps it — floated outside a "
                      "commentary column, so nothing here to share with.")
                continue
            host_h = _pt(host.height)
            used = max(drawn, declared)
            ratio = used / host_h if host_h else 0.0
            worst = max(worst, ratio)
            median = sorted(row_pts)[len(row_pts) // 2] if row_pts else 0.0
            fits = int(host_h / median) if median else 0
            flag = ""
            if ratio > 1.0:
                flag = "   *** TALLER THAN THE WHOLE SLOT ***"
                over += 1
            elif ratio > 0.75:
                flag = "   (leaves under a quarter of the slot for text)"
            print(f"      slot '{host.name}' is {host_h:.1f}pt -> table takes "
                  f"{100 * ratio:5.1f}%{flag}")
            print(f"      median row {median:.1f}pt, so about {fits} row(s) fit a full "
                  f"slot ({rows} drawn{'' if rows <= fits else f', {rows - fits} over'})")

    print(f"\n{RULE}")
    if over:
        print(f"{over} table(s) are taller than the slot they share. For those the packer\n"
              f"has no placement that works -- the lever is fewer rows, not better packing.\n"
              f"Worst is {100 * worst:.0f}% of a slot.")
    else:
        print("No table exceeds its slot. An overflow on these slides is the packer's\n"
              "call about what to put beside the table, not the table's size.")
    print(RULE)
    return 1 if over else 0


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pptx", help="an exported .pptx")
    ap.add_argument("--slide", type=int, help="only this slide number")
    args = ap.parse_args()
    if not os.path.exists(args.pptx):
        sys.exit(f"no such file: {args.pptx}")
    return report(args.pptx, args.slide)


if __name__ == "__main__":
    sys.exit(main())

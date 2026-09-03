#!/usr/bin/env python3
"""Rebuild the combined portfolio decks from previews already on disk,
leaving the 主表 roll-up out.

inspect_databook.py now excludes the roll-up from the merge (77a2949), but that
only changes what the NEXT batch produces. The eight near-empty roll-up pages
sit in the combined decks already exported, and re-running a batch to remove
them costs twelve minutes and real tokens when every preview needed is already
on disk.

    python ad-hoc/pptx-probes/remerge_without_rollup.py PREVIEW_DIR
    python ad-hoc/pptx-probes/remerge_without_rollup.py PREVIEW_DIR --dry-run

WRITES one new Portfolio_<label>_<stamp>.pptx per portfolio. The stamp is
taken at run time, so nothing already in the folder is overwritten or deleted;
the old combined decks stay where they are until you remove them yourself.

Previews merge in filename order, which is the order the batch merged them in
(confirmed on a real run: every per-slide row matched between the combined
deck and the previews concatenated this way).
"""
import re
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from fdd_utils.pptx import combine_presentations

#: Portfolio_I_20260903_102743.pptx -- an output of a previous merge, never an input
_COMBINED = re.compile(r"^(?:Portfolio_.+?|Combined)_\d{8}_\d{6}\.pptx$", re.I)
#: Project Mint.Portfolio I.南通通海.preview.pptx -> "I"
_LABEL = re.compile(r"Portfolio[ _]([^.\s_]+)", re.I)
#: The roll-up is the source entities read Financials from, not an entity.
ROLLUP_MARK = "主表"


def main() -> int:
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    dry = "--dry-run" in sys.argv
    if len(args) != 1 or not Path(args[0]).is_dir():
        print(__doc__)
        return 2
    folder = Path(args[0])

    groups: dict = {}
    skipped: list = []
    for f in sorted(folder.glob("*.pptx")):
        if f.name.startswith("~$") or _COMBINED.match(f.name):
            continue
        m = _LABEL.search(f.name)
        if not m:
            continue
        label = m.group(1).upper()
        if ROLLUP_MARK in f.name:
            skipped.append((label, f.name))
            continue
        groups.setdefault(label, []).append(f)

    if not groups:
        print(f"No per-entity .preview decks found in {folder}.")
        return 1

    stamp = time.strftime("%Y%m%d_%H%M%S")
    for label in sorted(groups):
        paths = groups[label]
        out = folder / f"Portfolio_{label}_{stamp}.pptx"
        print(f"\nPortfolio {label}: {len(paths)} entity deck(s)")
        for p in paths:
            print(f"    + {p.name}")
        for lab, name in skipped:
            if lab == label:
                print(f"    - {name}   (roll-up, left out)")
        if dry:
            print(f"  would write {out.name}")
            continue
        combine_presentations([str(p) for p in paths], str(out))
        print(f"  wrote {out.resolve()}")

    if not dry:
        print("\nThe previously combined decks are untouched -- delete them yourself "
              "once you have checked these.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

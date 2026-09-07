#!/usr/bin/env python3
"""N1 acceptance probe: does an unresolved sheet explain itself correctly?

`resolution["unresolved_sheets"]` carries a dict per sheet with a `reason` from
a closed set. Two of those reasons cannot be observed on a healthy databook —
every tab resolves — so this builds the broken cases instead: it copies the
workbook to a temp file, renames or hides tabs, re-runs the resolver, and prints
what came back. Free, offline, seconds; no LLM, no PowerPoint.

The databook path is a REQUIRED argument: every real databook in this repo is
gitignored client data and must never be named in a committed file.

    PYTHONPATH=. python ad-hoc/workbench/probe_unresolved_reasons.py <databook.xlsx>
    PYTHONPATH=. python ad-hoc/workbench/probe_unresolved_reasons.py <databook.xlsx> --tab Cash --other "Long-term deferred expenses"

Cases, in order:
  baseline          the file as it is
  renamed           --other's tab name AND its in-sheet title replaced with a
                    string no alias recognises   -> no_alias_scored_above_45
                    (renaming the tab alone is not enough: the dynamic
                    exact-name discovery still finds it by its in-sheet title)
  duplicate         --other renamed to "<tab> " (trailing space), which
                    normalizes to the same name as --tab, and its title blanked
                    so it can win nothing else    -> sheet_taken_by_<key>
  duplicate+hidden  the same, with --tab's real sheet marked hidden, to show the
                    hidden-sheet score penalty in the candidate list

`sheet_kind_excluded` and `normalization_error` need no construction — the first
shows up on any book with a nav/notes tab, the second on any sheet that resolves
and then fails to normalize.
"""
from __future__ import annotations

import argparse
import os
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..")))

from openpyxl import load_workbook  # noqa: E402

from fdd_utils.workbook import clear_workbook_caches  # noqa: E402
from fdd_utils.workbook.databook import extract_normalized_data_from_excel  # noqa: E402
from fdd_utils.workbook.inspector import profile_workbook  # noqa: E402

RENAMED_TAB_NAME = "ZZ working file"


def mutate(src: str, out_dir: str, tag: str, renames=None, hide=(), blank_title_rows=()) -> str:
    dst = os.path.join(out_dir, f"case_{tag}.xlsx")
    shutil.copy(src, dst)
    workbook = load_workbook(dst)
    for sheet_name, row_idx in blank_title_rows:  # row_idx is 0-based, as the profile reports it
        for col in range(1, 13):
            workbook[sheet_name].cell(row=row_idx + 1, column=col).value = None
    for old_name, new_name in (renames or {}).items():
        workbook[old_name].title = new_name
    for sheet_name in hide:
        workbook[sheet_name].sheet_state = "hidden"
    workbook.save(dst)
    workbook.close()
    return dst


def report(tag: str, path: str, focus_key: str) -> None:
    clear_workbook_caches()
    _, _, _, _, resolution = extract_normalized_data_from_excel(path)
    profile = resolution.get("workbook_profile") or {}
    winner = (resolution.get("resolved") or {}).get(focus_key) or {}
    print(f"\n--- {tag}")
    print(f"    resolved={len(resolution.get('resolved') or {})} "
          f"hidden={profile.get('hidden_sheets')} "
          f"{focus_key} -> {winner.get('sheet_name')!r} score={winner.get('score')}")
    for entry in resolution.get("unresolved_sheets") or []:
        print(f"    UNRESOLVED {entry['sheet_name']!r:24} kind={entry['sheet_kind']:19} "
              f"reason={entry['reason']:26} best_key={entry['best_candidate_key']} "
              f"score={entry['best_score']} alias={entry['best_alias_score']} "
              f"missed_by={entry['floor_missed_by']} taken_by={entry['taken_by_sheet']!r}")
    for candidate in (resolution.get("candidate_map") or {}).get(focus_key, [])[:4]:
        print(f"      candidate {candidate['sheet_name']!r:24} score={candidate['score']:>7} "
              f"alias={candidate['alias_score']:>7} below_floor={candidate['below_floor']}")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("databook", help="path to a databook (client data — never hardcode one)")
    parser.add_argument("--tab", default="Cash", help="the tab that should keep winning its mapping key")
    parser.add_argument("--other", default="Long-term deferred expenses",
                        help="a second tab, sacrificed to build the broken cases")
    args = parser.parse_args()

    if not os.path.exists(args.databook):
        sys.exit(f"No such databook: {args.databook}")
    profiles = profile_workbook(args.databook)
    for tab in (args.tab, args.other):
        if tab not in profiles:
            sys.exit(f"Tab {tab!r} is not in this workbook. Tabs: {sorted(profiles)}")
    other_title_row = profiles[args.other].get("title_row_idx")
    if other_title_row is None:
        sys.exit(f"Tab {args.other!r} has no detected title row; pick another --other")

    with tempfile.TemporaryDirectory(prefix="n1_unresolved_") as out_dir:
        report("baseline", args.databook, args.tab)
        report(
            f"renamed: {args.other!r} tab+title -> {RENAMED_TAB_NAME!r}",
            mutate(args.databook, out_dir, "renamed",
                   renames={args.other: RENAMED_TAB_NAME},
                   blank_title_rows=[(args.other, other_title_row)]),
            args.tab,
        )
        duplicate = {args.other: args.tab + " "}
        blanked = [(args.other, other_title_row)]
        report(
            f"duplicate: {args.other!r} -> {args.tab + ' '!r}",
            mutate(args.databook, out_dir, "dupe", renames=duplicate, blank_title_rows=blanked),
            args.tab,
        )
        report(
            f"duplicate + the real {args.tab!r} tab hidden",
            mutate(args.databook, out_dir, "dupehidden", renames=duplicate,
                   blank_title_rows=blanked, hide=[args.tab]),
            args.tab,
        )


if __name__ == "__main__":
    main()

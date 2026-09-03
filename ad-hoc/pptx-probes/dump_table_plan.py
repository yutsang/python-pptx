#!/usr/bin/env python3
"""Print every detail table's rows EXACTLY as the renderer will draw them,
without exporting a deck and without an AI call.

--export-pptx refuses to run without --run-ai, because the payloads need
commentary. The detail tables need none of that: they are built on the
deterministic path, land in df.attrs["presentation_detail_table"], and are
turned into rows by _build_presentation_table_plan. So a change to table
CONTENT can be checked in seconds and for free, instead of paying twelve
minutes and real tokens for a deck whose text is irrelevant to the question.

    python ad-hoc/pptx-probes/dump_table_plan.py DATABOOK.xlsx
    python ad-hoc/pptx-probes/dump_table_plan.py DATABOOK.xlsx \
        --financials-from ROLLUP.xlsx --financials-sheet <entity>Financials

Each row is printed with the shape it will render in:

    label            a plain component row
    [label]          a SECTION HEADING (drawn in the heading style)
        label        a member of the heading above it (indented)
      label          a child of the row above it (indented, tinted)
    = label          the total row

Two things worth reading off it:
  - a heading that is obviously a component (an asset class with a nil
    balance) means the section-label detection has misfired, and one such
    heading stands the whole table's prefix grouping down;
  - the same label twice with different numbers means a block boundary is
    missing between them.

Reads only. Writes nothing. No AI, no template, no export.
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from fdd_utils.workbook import process_workbook_data
from fdd_utils.pptx.helpers import _build_presentation_table_plan

_MARK = {"group": "  [", "grouped": "      ", "child": "        ",
         "data": "    ", "total": "  = "}


def main() -> int:
    ap = argparse.ArgumentParser(add_help=True, description=__doc__)
    ap.add_argument("databook")
    ap.add_argument("--entity", default=None,
                    help="entity name; defaults to the last dot-separated part "
                         "of the filename, which is how these files are named")
    ap.add_argument("--sheet", default=None)
    ap.add_argument("--financials-from", default=None)
    ap.add_argument("--financials-sheet", default=None)
    args = ap.parse_args()

    path = Path(args.databook)
    if not path.is_file():
        print(f"Not a file: {path}")
        return 2
    entity = args.entity or path.stem.split(".")[-1].strip()
    print(f"databook  {path.name}\nentity    {entity}\n")

    state = process_workbook_data(
        temp_path=str(path), entity_name=entity, selected_sheet=args.sheet,
        debug=False, financials_from=args.financials_from,
        financials_sheet=args.financials_sheet,
    )
    dfs = state.get("dfs") or {}
    is_chinese = str(state.get("language") or "").strip().lower() in ("chi", "chinese")

    shown = 0
    for account, df in sorted(dfs.items()):
        table = (getattr(df, "attrs", None) or {}).get("presentation_detail_table")
        if not (table and table.get("rows")):
            continue
        shown += 1
        rows = table["rows"]
        headings = [r for r in rows if r.get("is_header")]
        plan = _build_presentation_table_plan(table, is_chinese, 1)

        print(f"{'=' * 74}\n{account}   (title {table.get('title')!r}, "
              f"source {table.get('synthesized_from') or 'sheet block'})")
        print(f"  {len(rows)} extracted row(s), {len(headings)} kept as section "
              f"heading(s) -> prefix grouping "
              f"{'STANDS DOWN' if headings else 'runs'}"
              f"   -> {len(plan)} rendered row(s)")
        for entry in plan:
            close = "]" if entry["kind"] == "group" else ""
            print(f"{_MARK.get(entry['kind'], '    ')}{entry['label']}{close}")

        seen: dict = {}
        for entry in plan:
            if entry["kind"] in ("group", "total"):
                continue
            seen.setdefault(entry["label"], []).append(entry)
        repeats = {k: v for k, v in seen.items() if len(v) > 1}
        if repeats:
            print(f"  ⚠️  {len(repeats)} label(s) appear more than once: "
                  f"{', '.join(sorted(repeats))}")
        print()

    if not shown:
        print("No account carried a presentation_detail_table. Run\n"
              f"    python inspect_databook.py \"{path}\"\n"
              "and read section 3c, which says per account why not.")
        return 1
    print(f"{shown} table(s).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

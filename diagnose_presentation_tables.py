#!/usr/bin/env python3
"""Traces why presentation_tables.enabled=true produced ZERO native
detail tables on a real export (only the pre-existing BS/IS overview grid
tables showed up), even though the 4 known accounts (营业成本, 税金及附加,
管理费用, 财务费用) have real presentation_detail_table data confirmed by
inspect_table_candidates.py.

Runs the REAL pipeline functions -- process_workbook_data,
build_pptx_structured_payloads, PowerPointGenerator's own packing/table
methods -- with placeholder (non-AI) commentary, printing at every stage so
whichever one silently drops the tables is visible instead of guessed at.

Read-only except for writing the diagnostic .pptx export.

Usage:
    python diagnose_presentation_tables.py "for_test/Crescent-databook.xlsx"
"""
import argparse
import sys
import warnings

warnings.filterwarnings("ignore")
sys.path.insert(0, ".")

from fdd_utils.workbook import process_workbook_data, load_mappings
from fdd_utils.pptx import PowerPointGenerator, build_pptx_structured_payloads

TARGET_ACCOUNTS = ("营业成本", "税金及附加", "管理费用", "财务费用")


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("path")
    ap.add_argument("--entity", default="x")
    ap.add_argument("--out", default="diagnose_presentation_tables_output.pptx")
    args = ap.parse_args()

    print("=" * 78)
    print("STAGE 0: config")
    print("=" * 78)
    gen = PowerPointGenerator(template_path="fdd_utils/template.pptx", language="chinese")
    print(f"  pptx_settings.presentation_tables = {gen.pptx_settings.get('presentation_tables')!r}")
    print(f"  _presentation_tables_enabled() = {gen._presentation_tables_enabled()}")
    if not gen._presentation_tables_enabled():
        print("  ^^ STOP: config is not actually enabled in this process. If you set "
              "enabled: true and still see this, something is loading a different "
              "config.yml, or a cached config from before the edit -- check "
              "fdd_utils/config.yml's real path and restart whatever process runs this.")
        return 1

    print("\n" + "=" * 78)
    print(f"STAGE 1: process_workbook_data({args.path!r})")
    print("=" * 78)
    result = process_workbook_data(temp_path=args.path, entity_name=args.entity, selected_sheet=None)
    dfs = result["dfs"]
    print(f"  {len(dfs)} account(s), language={result.get('language')}")
    for acc in TARGET_ACCOUNTS:
        df = dfs.get(acc)
        if df is None:
            print(f"  {acc:12s}: NOT FOUND in dfs -- key mismatch. Available keys containing "
                  f"a similar substring: {[k for k in dfs if acc[:2] in k]}")
            continue
        table = (df.attrs or {}).get("presentation_detail_table")
        has_rows = bool(table and table.get("rows"))
        print(f"  {acc:12s}: found, df.shape={df.shape}, presentation_detail_table rows="
              f"{len(table['rows']) if has_rows else 0}")

    print("\n" + "=" * 78)
    print("STAGE 2: build_pptx_structured_payloads (the REAL function ui.py calls)")
    print("=" * 78)
    mappings = load_mappings()
    # Placeholder ai_results, KEYED EXACTLY LIKE dfs -- isolates stage 2's own
    # logic from any real-pipeline key-naming mismatch between ai_results and
    # dfs, which is checked separately at the end if this stage looks clean.
    ai_results = {
        key: {"agent_4_validation": {}, "final_content": f"[placeholder] {key} 明细如下：",
              "agent_1_output": f"[placeholder] {key} 明细如下："}
        for key in dfs
    }
    payloads = build_pptx_structured_payloads(ai_results, mappings, dfs=dfs)
    is_items = payloads["IS"]
    bs_items = payloads["BS"]
    print(f"  BS items: {len(bs_items)}  IS items: {len(is_items)}")
    for item in is_items:
        if item["mapping_key"] in TARGET_ACCOUNTS:
            fd = item.get("financial_data")
            has_attrs = hasattr(fd, "attrs") and bool((fd.attrs or {}).get("presentation_detail_table", {}).get("rows"))
            print(f"  {item['mapping_key']:12s}: in IS payload, financial_data is "
                  f"{type(fd).__name__}, presentation_detail_table survives={has_attrs}")
    found_keys = {item["mapping_key"] for item in is_items}
    missing = [a for a in TARGET_ACCOUNTS if a not in found_keys]
    if missing:
        print(f"  ^^ MISSING from IS payload entirely: {missing} -- these never reached "
              f"stage 2's output at all (check account_type in mappings.yml, or "
              f"find_mapping_key resolution for these exact keys).")

    print("\n" + "=" * 78)
    print("STAGE 3: PowerPointGenerator table extraction + packing")
    print("=" * 78)
    gen.load_template()
    for item in is_items:
        if item["mapping_key"] in TARGET_ACCOUNTS:
            table = gen._presentation_table_for_account(item)
            print(f"  {item['mapping_key']:12s}: _presentation_table_for_account -> "
                  f"{'FOUND (' + str(len(table['rows'])) + ' rows)' if table else 'None'}")

    prepared = gen._prepare_structured_data_for_slides(is_items)
    tables_enabled = gen._presentation_tables_enabled()
    table_items, normal_items = [], []
    for item in prepared:
        table = gen._presentation_table_for_account(item) if tables_enabled else None
        if table:
            item = dict(item)
            item["commentary"] = gen._truncate_for_table_lead_in(item.get("commentary", ""), bool(item.get("is_chinese")))
            item["_presentation_table"] = table
            table_items.append(item)
        else:
            normal_items.append(item)
    print(f"\n  table_items={len(table_items)} {[i['mapping_key'] for i in table_items]}")
    print(f"  normal_items={len(normal_items)}")

    max_slides = int(gen.pptx_settings.get("max_commentary_slides_per_statement", 4) or 4)
    dist = gen._distribute_content_across_slots(normal_items, max_slides=max_slides, start_slide=5, statement_type="IS")
    print(f"\n  normal packing result ({len(dist)} slot(s)):")
    for s, n, accs in dist:
        print(f"    slide_idx={s} slot={n!r} accounts={[a['mapping_key'] for a in accs]}")

    full = gen._append_table_accounts_to_distribution(table_items, dist, max_slides=max_slides, start_slide=5)
    print(f"\n  after appending table accounts ({len(full)} slot(s)):")
    for s, n, accs in full:
        tag = " <-- TABLE ACCOUNT" if accs and accs[0].get("_presentation_table") else ""
        print(f"    slide_idx={s} slot={n!r} accounts={[a['mapping_key'] for a in accs]}{tag}")

    print("\n" + "=" * 78)
    print("STAGE 4: real apply_structured_data_to_slides + export")
    print("=" * 78)
    gen.apply_structured_data_to_slides(
        is_items, start_slide=5, project_name=args.entity, statement_type="IS",
        is_chinese_databook=(result.get("language") == "Chi"),
    )
    if bs_items:
        gen2 = PowerPointGenerator(template_path="fdd_utils/template.pptx", language="chinese")
        gen2.pptx_settings["presentation_tables"] = {"enabled": True}
        gen2.load_template()
        gen2.apply_structured_data_to_slides(
            bs_items, start_slide=1, project_name=args.entity, statement_type="BS",
            is_chinese_databook=(result.get("language") == "Chi"),
        )
        # Merge BS slides into gen's presentation isn't straightforward across
        # instances -- BS is checked separately in its own right below instead.
        unused_bs = sorted(set(getattr(gen2, "_unused_slides_to_remove", []) or []), reverse=True)
        if unused_bs:
            gen2._remove_slides(unused_bs)
        gen2.presentation.save(args.out.replace(".pptx", "_BS.pptx"))
        print(f"  BS export saved separately: {args.out.replace('.pptx', '_BS.pptx')}")

    unused = sorted(set(getattr(gen, "_unused_slides_to_remove", []) or []), reverse=True)
    print(f"  _unused_slides_to_remove BEFORE removal: {[u + 1 for u in unused]} (1-indexed)")
    table_shapes_before_removal = sum(
        1 for slide in gen.presentation.slides for shape in slide.shapes if getattr(shape, "has_table", False)
    )
    print(f"  table shapes present BEFORE slide removal: {table_shapes_before_removal}")

    if unused:
        gen._remove_slides(unused)
    gen.presentation.save(args.out)
    print(f"\n  Saved: {args.out}  ({len(gen.presentation.slides)} slides after removal)")

    table_shapes_after = sum(
        1 for slide in gen.presentation.slides for shape in slide.shapes if getattr(shape, "has_table", False)
    )
    print(f"  table shapes present AFTER slide removal: {table_shapes_after}")
    print("\nCompare the two 'table shapes present' counts above: if BEFORE > AFTER, "
          "the unused-slide removal step deleted a slide that had a real table on it.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

"""M0 probe: which deterministic guidance sections actually reach the Generator prompt.

Free, no LLM. Prints a PRESENT/ABSENT matrix per account, then re-runs with the
proposed M0 fix applied in memory (attach the unfiltered prompt_analysis_df to
the detail_analysis variant) and prints the same matrix for comparison.
"""
import sys
from fdd_utils.workbook import process_workbook_data
from fdd_utils.ai.prompts import PromptEngine

# (label, english marker, chinese marker)
MARKERS = [
    ("trend_summary",   "Trend summary",                "趋势摘要"),
    ("sig_movements",   "Significant movements",        "重大变动"),
    ("remainder",       "[REMAINDER ALREADY COMPUTED]", "【余额差额已算好】"),
    ("comp_nature",     "[COMPONENT NATURE",            "【构成项性质"),
    ("hierarchy",       "[VERIFIED HIERARCHY",          "【本科目已核对的层级关系】"),
    # controls: these read df/attrs directly and should be alive today
    ("CTRL material",   "[MATERIAL MOVEMENT]",          "【重大变动提示】"),
    ("CTRL insight",    "[DATA INSIGHT",                "【数据洞察"),
]


def matrix(dfs, language, pe, keys):
    rows = []
    for k in keys:
        df = dfs.get(k)
        try:
            sys_p, usr_p = pe.render_prompt("subagent_1", language, k, df, data_format="markdown")
            blob = (sys_p or "") + "\n" + (usr_p or "")
        except Exception as exc:  # noqa: BLE001
            rows.append((k, f"RENDER FAILED: {type(exc).__name__}: {exc}", None))
            continue
        hits = {}
        for label, eng, chi in MARKERS:
            hits[label] = ("E" if eng in blob else "") + ("C" if chi in blob else "")
        rows.append((k, hits, len(blob)))
    return rows


def show(title, rows):
    print(f"\n=== {title} ===")
    labels = [m[0] for m in MARKERS]
    print(f"{'account':<22} " + " ".join(f"{l:>14}" for l in labels) + "   chars")
    for k, hits, size in rows:
        if isinstance(hits, str):
            print(f"{k:<22} {hits}")
            continue
        cells = " ".join(f"{(hits[l] or '-'):>14}" for l in labels)
        print(f"{k:<22} {cells}   {size}")


def main(path, entity=""):
    state = process_workbook_data(temp_path=path, entity_name=entity, selected_sheet=None)
    dfs = state["dfs"]
    language = state.get("language") or "Eng"
    print(f"language={language}  accounts={len(dfs)}")
    pe = PromptEngine()

    keys = list(dfs.keys())[:8]
    first = dfs[keys[0]]
    print("variant:", first.attrs.get("selected_variant"),
          "| has nested prompt_analysis_df:", first.attrs.get("prompt_analysis_df") is not None,
          "| has significant_movements:", first.attrs.get("significant_movements") is not None,
          "| component_descriptions:", len(first.attrs.get("component_descriptions") or []))

    show("BEFORE (current code)", matrix(dfs, language, pe, keys))

    # --- apply the proposed M0 fix in memory -------------------------------
    # Rebuild the nested frame the same way schedules.py does, from the
    # normalized results, and attach it to each detail_analysis variant.
    from fdd_utils.workbook.databook import extract_normalized_data_from_excel
    out = extract_normalized_data_from_excel(path, entity_name=entity or None)
    normalized = out[0] if isinstance(out, tuple) else out
    by_display = {}
    for sheet, norm in normalized.items():
        if not isinstance(norm, dict):
            continue
        dk = str(norm.get("display_key") or sheet)
        by_display[dk] = norm.get("prompt_analysis_df")

    attached = 0
    for k, df in dfs.items():
        nested = by_display.get(k)
        if nested is not None and df.attrs.get("prompt_analysis_df") is None:
            df.attrs["prompt_analysis_df"] = nested
            attached += 1
    print(f"\nattached nested frame to {attached}/{len(dfs)} accounts")

    show("AFTER (M0 fix applied in memory)", matrix(dfs, language, pe, keys))

    # copy-safety smoke test on a real frame
    import copy
    d = dfs[keys[0]]
    for op, fn in (("copy", lambda x: x.copy()),
                   ("deepcopy", lambda x: copy.deepcopy(x)),
                   ("reset_index", lambda x: x.reset_index(drop=True))):
        try:
            out = fn(d)
            print(f"{op}: ok, nested preserved =", out.attrs.get("prompt_analysis_df") is not None)
        except RecursionError:
            print(f"{op}: RecursionError")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "")

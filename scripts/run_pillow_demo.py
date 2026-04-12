"""
scripts/run_pillow_demo.py

Headless harness: runs each databook through the full FDD pipeline once
(workbook -> AI subagents -> structured payloads), then exports the same
PPTX twice — once with the legacy CPL heuristic, once with the new
Pillow-based text fitting — for side-by-side visual comparison.

The Pillow path is toggled via the FDD_USE_PILLOW_FITTING environment
variable, which the new helper in pptx.py reads at call time.

Outputs land in fdd_utils/output/pillow_pptx_checks/:
    <databook>_legacy.pptx
    <databook>_pillow.pptx
    summary.json   (paths, account counts, durations, errors)

Run with:
    python scripts/run_pillow_demo.py
"""

from __future__ import annotations

import json
import os
import re
import sys
import time
import traceback
from pathlib import Path
from typing import Any, Dict, List, Optional

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from fdd_utils.workbook import (  # noqa: E402
    build_workbook_preflight,
    extract_entity_names_from_preflight,
    get_financial_sheet_options,
    get_effective_mappings,
    load_mappings,
    process_workbook_data,
)
from fdd_utils.ai import run_ai_pipeline_with_progress  # noqa: E402
from fdd_utils.pptx import (  # noqa: E402
    build_pptx_structured_payloads,
    export_pptx_from_structured_data_combined,
)


DATABOOKS: List[str] = [
    "[DEMO]240627.东莞岭南-databook.xlsx",
    "databook-Foshan Wanyuan_rebuilt.xlsx",
    "databook.xlsx",
    "Project Gold Kunshan.databook.xlsx",
]

# Skip databooks whose pillow+legacy PPTX already exist (cache-friendly restart).
SKIP_IF_EXISTS: bool = True

OUTPUT_DIR = REPO_ROOT / "fdd_utils" / "output" / "pillow_pptx_checks"
TEMPLATE_PATH = REPO_ROOT / "fdd_utils" / "template.pptx"


def _sanitize(name: str) -> str:
    stem = re.sub(r"\.(xlsx|xls)$", "", name, flags=re.IGNORECASE)
    return re.sub(r"[^\w\-]+", "_", stem).strip("_") or "databook"


def _pick_entity(workbook_path: Path) -> tuple[str, Optional[str]]:
    """Pick a usable (entity_name, sheet_name) pair via the same preflight
    helpers the Streamlit UI uses."""
    preflight = build_workbook_preflight(str(workbook_path))
    entities = extract_entity_names_from_preflight(preflight)
    sheets = get_financial_sheet_options(preflight)
    entity = entities[0] if entities else workbook_path.stem
    sheet = sheets[0] if sheets else None
    return entity, sheet


def _run_pipeline(workbook_path: Path) -> Dict[str, Any]:
    """workbook -> {dfs, bs_is_results, ai_results, mappings, language, entity}"""
    entity, sheet = _pick_entity(workbook_path)
    print(f"  entity={entity!r} sheet={sheet!r}")

    processed = process_workbook_data(
        temp_path=str(workbook_path),
        entity_name=entity,
        selected_sheet=sheet,
        debug=False,
    )

    dfs = processed.get("dfs") or {}
    workbook_list = processed.get("workbook_list") or []
    bs_is_results = processed.get("bs_is_results")
    language = processed.get("language") or "Eng"
    resolution = processed.get("resolution")
    mappings = get_effective_mappings(load_mappings(), resolution)

    selected_dfs = {key: dfs[key] for key in workbook_list if key in dfs}
    if not selected_dfs:
        selected_dfs = dict(dfs)
    mapping_keys = list(selected_dfs.keys())
    print(f"  mapping_keys={len(mapping_keys)} accounts")

    if not mapping_keys:
        raise RuntimeError(f"No accounts extracted from {workbook_path.name}")

    ai_started = time.perf_counter()
    ai_results = run_ai_pipeline_with_progress(
        mapping_keys=mapping_keys,
        dfs=selected_dfs,
        model_type="deepseek",
        language=language,
        use_multithreading=True,
    )
    ai_elapsed = time.perf_counter() - ai_started
    print(f"  AI pipeline finished in {ai_elapsed:.1f}s")

    return {
        "entity": entity,
        "sheet": sheet,
        "language": language,
        "dfs": selected_dfs,
        "bs_is_results": bs_is_results,
        "mappings": mappings,
        "ai_results": ai_results,
        "ai_elapsed_s": ai_elapsed,
        "temp_path": str(workbook_path),
    }


def _export_pptx(
    pipeline: Dict[str, Any],
    output_path: Path,
    *,
    use_pillow: bool,
) -> Dict[str, Any]:
    """Export a single PPTX, toggling FDD_USE_PILLOW_FITTING for this call."""
    payloads = build_pptx_structured_payloads(
        ai_results=pipeline["ai_results"],
        mappings=pipeline["mappings"],
        bs_is_results=pipeline["bs_is_results"],
        dfs=pipeline["dfs"],
        commentary_modes=None,
    )
    bs_data = payloads.get("BS", [])
    is_data = payloads.get("IS", [])

    prev = os.environ.get("FDD_USE_PILLOW_FITTING")
    os.environ["FDD_USE_PILLOW_FITTING"] = "1" if use_pillow else "0"
    started = time.perf_counter()
    error: Optional[str] = None
    try:
        export_pptx_from_structured_data_combined(
            str(TEMPLATE_PATH),
            bs_data,
            is_data,
            str(output_path),
            pipeline["entity"],
            language="chinese" if pipeline["language"] == "Chn" else "english",
            temp_path=pipeline["temp_path"],
            selected_sheet=pipeline["sheet"],
            is_chinese_databook=(pipeline["language"] == "Chn"),
            bs_is_results=pipeline["bs_is_results"],
            model_type="deepseek",
        )
    except Exception as exc:
        error = f"{type(exc).__name__}: {exc}"
        traceback.print_exc()
    finally:
        if prev is None:
            os.environ.pop("FDD_USE_PILLOW_FITTING", None)
        else:
            os.environ["FDD_USE_PILLOW_FITTING"] = prev
    elapsed = time.perf_counter() - started

    return {
        "path": str(output_path),
        "exists": output_path.exists(),
        "size_bytes": output_path.stat().st_size if output_path.exists() else 0,
        "use_pillow": use_pillow,
        "elapsed_s": elapsed,
        "bs_count": len(bs_data),
        "is_count": len(is_data),
        "error": error,
    }


def main() -> int:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    summary: List[Dict[str, Any]] = []

    for name in DATABOOKS:
        workbook_path = REPO_ROOT / name
        print()
        print("=" * 78)
        print(f"DATABOOK: {name}")
        print("=" * 78)
        if not workbook_path.exists():
            print(f"  SKIP: {workbook_path} not found")
            summary.append({"workbook": name, "error": "not_found"})
            continue

        sanitized_pre = _sanitize(name)
        legacy_pre = OUTPUT_DIR / f"{sanitized_pre}_legacy.pptx"
        pillow_pre = OUTPUT_DIR / f"{sanitized_pre}_pillow.pptx"
        if SKIP_IF_EXISTS and legacy_pre.exists() and pillow_pre.exists():
            print(f"  SKIP: both PPTX already exist for {name}")
            summary.append({
                "workbook": name,
                "skipped": True,
                "legacy": {"path": str(legacy_pre), "size_bytes": legacy_pre.stat().st_size},
                "pillow": {"path": str(pillow_pre), "size_bytes": pillow_pre.stat().st_size},
            })
            continue

        record: Dict[str, Any] = {"workbook": name}
        try:
            pipeline = _run_pipeline(workbook_path)
        except Exception as exc:
            traceback.print_exc()
            record["error"] = f"{type(exc).__name__}: {exc}"
            summary.append(record)
            continue

        record["entity"] = pipeline["entity"]
        record["sheet"] = pipeline["sheet"]
        record["language"] = pipeline["language"]
        record["account_count"] = len(pipeline["dfs"])
        record["ai_elapsed_s"] = round(pipeline["ai_elapsed_s"], 2)

        sanitized = _sanitize(name)
        legacy_path = OUTPUT_DIR / f"{sanitized}_legacy.pptx"
        pillow_path = OUTPUT_DIR / f"{sanitized}_pillow.pptx"

        print()
        print("  -> exporting LEGACY (CPL heuristic)")
        record["legacy"] = _export_pptx(pipeline, legacy_path, use_pillow=False)
        print(
            f"     {record['legacy']['path']} "
            f"({record['legacy']['size_bytes']} bytes, "
            f"{record['legacy']['elapsed_s']:.1f}s)"
        )

        print()
        print("  -> exporting PILLOW (real font metrics)")
        record["pillow"] = _export_pptx(pipeline, pillow_path, use_pillow=True)
        print(
            f"     {record['pillow']['path']} "
            f"({record['pillow']['size_bytes']} bytes, "
            f"{record['pillow']['elapsed_s']:.1f}s)"
        )

        summary.append(record)

    summary_path = OUTPUT_DIR / "summary.json"
    with summary_path.open("w", encoding="utf-8") as handle:
        json.dump(summary, handle, ensure_ascii=False, indent=2)

    print()
    print("=" * 78)
    print(f"Summary written to {summary_path}")
    print("=" * 78)
    for record in summary:
        name = record.get("workbook", "?")
        if record.get("error"):
            print(f"  [FAIL] {name}: {record['error']}")
            continue
        legacy = record.get("legacy", {})
        pillow = record.get("pillow", {})
        print(
            f"  [OK]   {name}: legacy={Path(legacy.get('path','?')).name} "
            f"({legacy.get('size_bytes', 0)}B), "
            f"pillow={Path(pillow.get('path','?')).name} "
            f"({pillow.get('size_bytes', 0)}B), "
            f"AI={record.get('ai_elapsed_s', 0)}s"
        )

    failures = [r for r in summary if r.get("error") or (r.get("legacy", {}).get("error") or r.get("pillow", {}).get("error"))]
    return 0 if not failures else 1


if __name__ == "__main__":
    sys.exit(main())

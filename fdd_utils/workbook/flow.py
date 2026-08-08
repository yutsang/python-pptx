from __future__ import annotations


from .mapping import get_effective_mappings, load_mappings
from .statements import extract_balance_sheet_and_income_statement, synthesize_balance_sheet_and_income_statement
from .databook import build_dataframe_variants_from_normalized_results, extract_normalized_data_from_excel
from .reconcile import reconcile_financial_statements
import contextlib
import io
import logging
import os
import time
from typing import Any, Dict

logger = logging.getLogger(__name__)


def process_workbook_data(
    *,
    temp_path: str,
    entity_name: str,
    selected_sheet: str | None,
    mapping_overrides: Dict[str, str] | None = None,
    debug: bool = False,
    financials_from: str | None = None,
    financials_sheet: str | None = None,
) -> Dict[str, Any]:
    process_started = time.perf_counter()
    normalized_results, normalized_workbook_list, _, language, resolution = extract_normalized_data_from_excel(
        databook_path=temp_path,
        entity_name=entity_name,
        mode="All",
        mapping_overrides=mapping_overrides or None,
    )
    logger.debug(
        "Built normalized workbook payload for %s in %.2fs",
        os.path.basename(temp_path),
        time.perf_counter() - process_started,
    )

    dataframe_variants = build_dataframe_variants_from_normalized_results(
        normalized_results=normalized_results,
        workbook_list=normalized_workbook_list,
        report_language=language,
        variant_specs=[
            {
                "name": f"{view_name}_{variant}",
                "variant": variant,
                "filter_details": filter_details,
                "keep_zero_rows": False,
            }
            for view_name, filter_details in (("display", False), ("detail", True))
            for variant in ("original", "default", "analysis")
        ],
    )
    display_dfs_original = dataframe_variants.get("display_original", {}).get("dfs", {})
    display_workbook_list = dataframe_variants.get("display_original", {}).get("workbook_list", [])
    display_dfs = dataframe_variants.get("display_default", {}).get("dfs", {})
    dfs_original = dataframe_variants.get("detail_original", {}).get("dfs", {})
    # Use analysis variant (all Indicative adjusted periods) as primary for AI
    dfs = dataframe_variants.get("detail_analysis", {}).get("dfs", {})
    if not dfs:
        dfs = dataframe_variants.get("detail_default", {}).get("dfs", {})
    workbook_list = dataframe_variants.get("detail_analysis", {}).get("workbook_list", [])
    if not workbook_list:
        workbook_list = dataframe_variants.get("detail_default", {}).get("workbook_list", [])

    debug_buffer = io.StringIO() if debug else None
    debug_ctx = contextlib.redirect_stdout(debug_buffer) if debug_buffer else contextlib.nullcontext()

    # Some portfolios keep each sub-entity's own databook free of any
    # Financials-pattern sheet at all -- the real summary for that entity
    # instead lives inside a sibling roll-up ("主表") workbook, one sheet per
    # entity. financials_from/financials_sheet let the caller point the BS/IS
    # extraction at that sibling file's named sheet instead, while dfs (the
    # breakdown tabs above) still come from temp_path as normal either way.
    _financials_workbook_path = financials_from or temp_path
    _financials_sheet_name = financials_sheet if financials_from else selected_sheet

    bs_is_results = None
    if _financials_sheet_name:
        bs_started = time.perf_counter()
        with debug_ctx:
            bs_is_results = extract_balance_sheet_and_income_statement(
                workbook_path=_financials_workbook_path,
                sheet_name=_financials_sheet_name,
                debug=debug,
            )
        logger.debug(
            "Extracted financial summary sheet %s (from %s) in %.2fs",
            _financials_sheet_name,
            os.path.basename(_financials_workbook_path),
            time.perf_counter() - bs_started,
        )

    recon_bs, recon_is = None, None
    if dfs_original and bs_is_results:
        recon_started = time.perf_counter()
        effective_mappings = get_effective_mappings(load_mappings(), resolution)
        with debug_ctx:
            recon_bs, recon_is = reconcile_financial_statements(
                bs_is_results=bs_is_results,
                dfs=dfs_original,
                mappings=effective_mappings,
                tolerance=1.0,
                materiality_threshold=0.005,
                debug=debug,
            )
        logger.debug(
            "Reconciled %s account tabs in %.2fs",
            len(dfs),
            time.perf_counter() - recon_started,
        )

    # No literal "Financials"-style sheet found (bs_is_results stayed None
    # above) -- synthesize a BS/IS summary purely from the schedule tabs that
    # DID get mapped, so the PPTX embedded BS/IS summary table still has
    # something to render instead of silently skipping it. Deliberately built
    # AFTER the reconciliation block above (using the real bs_is_results,
    # which is still None here) rather than feeding this synthetic version
    # into reconcile_financial_statements -- comparing each schedule tab's
    # own total against itself would be a trivial self-match, not a real
    # cross-check.
    if bs_is_results is None and dfs_original:
        effective_mappings = get_effective_mappings(load_mappings(), resolution)
        synthesized = synthesize_balance_sheet_and_income_statement(dfs_original, effective_mappings)
        if synthesized.get("balance_sheet") is not None or synthesized.get("income_statement") is not None:
            bs_is_results = synthesized

    logger.debug(
        "Finished Process Data for %s in %.2fs",
        os.path.basename(temp_path),
        time.perf_counter() - process_started,
    )

    return {
        "dfs": dfs,
        "display_dfs": display_dfs,
        "dfs_variants": {
            "default": dfs,
            "original": dfs_original,
        },
        "display_df_variants": {
            "default": display_dfs,
            "original": display_dfs_original,
        },
        "workbook_list": workbook_list,
        "display_workbook_list": display_workbook_list,
        "language": language,
        "bs_is_results": bs_is_results,
        "reconciliation": (recon_bs, recon_is),
        "resolution": resolution,
        "project_name": bs_is_results.get("project_name") if bs_is_results else None,
        "entity_name": entity_name,
        "display_dfs_original": display_dfs_original,
        "debug_output": debug_buffer.getvalue() if debug_buffer else "",
    }

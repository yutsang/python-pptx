"""The deck, in the app: export it from the agent's working session, or (the replayed demo) from recorded results; and the Bridge Lab page.

The agent's deck is written by the session itself (agent/export.py: the production exporter with its packing and splitting, the
delivery decision, the lineage manifest and the layout check), so the app only asks for it. The results-shaped export below it is
what the demo, which has no session, uses.
"""
from __future__ import annotations

import datetime as dt_module
import logging
import os
import re
from typing import Any

import pandas as pd
import streamlit as st

from ..pptx import build_pptx_structured_payloads
from ..workbook.entity import batch_extract_entity_data  # noqa: F401  (the app and the CLI take it from here)
from .ai_panel import build_selected_pipeline_dfs, effective_mappings_from_session

logger = logging.getLogger(__name__)


def generate_pptx_presentation(
    *,
    session_state: Any,
    pptx_available: bool,
) -> None:
    if not session_state.ai_results:
        st.error("❌ No AI results available. Generate AI content first.")
        return

    if not pptx_available:
        st.error("❌ PPTX generation not available. Missing required modules.")
        return

    if session_state.get("agent_work") is not None:
        from .. import usage_ledger
        from . import agent_run
        with st.spinner("Writing the deck and checking the slides…"):
            exported = agent_run.export_deck(session_state)
        if session_state.get("pptx_ready"):
            usage_ledger.record_event("export_pptx", mode="single", delivery=(session_state.get("delivery_decision") or {}).get("status"))
        else:
            error = exported.get("error") or {}
            st.error(f"❌ PPTX generation failed: [{error.get('code')}] {error.get('message')}")
        return

    project_name = session_state.get("project_name", "Project")
    entity_name = session_state.get("entity_name", project_name)
    language = session_state.get("language", "Eng")
    mappings = effective_mappings_from_session(session_state)

    template_path = None
    for template in ["fdd_utils/template.pptx", "template.pptx"]:
        if os.path.exists(template):
            template_path = template
            break
    if not template_path:
        st.error("❌ PowerPoint template not found. Please add `fdd_utils/template.pptx` or `template.pptx`.")
        return

    output_dir = "fdd_utils/output"
    os.makedirs(output_dir, exist_ok=True)

    timestamp = dt_module.datetime.now().strftime("%Y%m%d_%H%M%S")
    sanitized_entity = re.sub(r"[^\w\-_]", "_", str(entity_name)).strip("_") or "Project"
    selected_pipeline_dfs = build_selected_pipeline_dfs(session_state)

    try:
        combined_output_path = os.path.join(output_dir, f"{sanitized_entity}_{timestamp}.pptx")
        from fdd_utils.pptx import export_pptx_from_structured_data_combined

        structured_payloads = build_pptx_structured_payloads(
            ai_results=session_state.ai_results,
            mappings=mappings,
            bs_is_results=session_state.bs_is_results,
            dfs=selected_pipeline_dfs,
        )
        bs_data = structured_payloads.get("BS", [])
        is_data = structured_payloads.get("IS", [])
        logger.debug("PPTX payload account counts | BS=%s | IS=%s", len(bs_data), len(is_data))

        if not bs_data and not is_data:
            st.error("❌ No content generated for PPTX")
            logger.debug(
                "PPTX payload is empty | ai_results_keys=%s | dfs_keys=%s",
                list(session_state.ai_results.keys())[:10] if session_state.ai_results else "None",
                list(selected_pipeline_dfs.keys())[:10] if selected_pipeline_dfs else "None",
            )
            return

        # embed_financial_tables reads (temp_path, selected_sheet) for its
        # currency-unit-label detection and as a fresh-extraction fallback --
        # when this entity's financials came from an uploaded roll-up
        # workbook (the "進階：主表" expander), that source -- not this
        # entity's own file/selected_sheet -- is the one that actually holds
        # them. Mirrors process_workbook_data's own precedence (financials_from
        # or temp_path).
        _rollup_temp_path = session_state.get("rollup_temp_path")
        _financials_workbook_path = _rollup_temp_path or session_state.get("temp_path")
        _financials_sheet_name = (
            session_state.get("rollup_sheet") if _rollup_temp_path else session_state.get("selected_sheet")
        )

        from .. import usage_ledger
        from ..delivery import decide_delivery, labelled_path
        delivery = decide_delivery(
            selection=session_state.get("reconciliation_selection"),
            ai_results=session_state.ai_results,
            bs_is_results=session_state.get("bs_is_results"),
        )
        combined_output_path = labelled_path(combined_output_path, delivery)
        session_state.delivery_decision = delivery.to_dict()
        usage_ledger.record_event("export_pptx", mode="single", delivery=delivery.status)
        with st.spinner("Generating PPTX…"):
            export_pptx_from_structured_data_combined(
                template_path,
                bs_data,
                is_data,
                combined_output_path,
                entity_name,
                language="chinese" if language == "Chn" else "english",
                temp_path=_financials_workbook_path,
                selected_sheet=_financials_sheet_name,
                is_chinese_databook=(language == "Chn"),
                bs_is_results=session_state.get("bs_is_results"),
                model_type=None,
                skip_summary_ai=True,
                pre_generated_summaries=session_state.get("section_summaries") or None,
                mappings=mappings,
                delivery=delivery,
            )
        if os.path.exists(combined_output_path):
            with open(combined_output_path, "rb") as handle:
                session_state.pptx_download_data = handle.read()
            session_state.pptx_download_filename = os.path.basename(combined_output_path)
            session_state.pptx_download_mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            session_state.pptx_ready = True

    except Exception as exc:
        st.error(f"❌ PPTX generation failed: {exc}")
        import traceback

        st.code(traceback.format_exc())


# --- Bridge Lab (experimental, isolated testing page) ---
# Lets the project team upload a workbook, pick one tab, and auto-detect a
# 'Base'/'Change' bridge helper block (the same convention as the real
# 成都-量价桥图 tab) to generate a NATIVE Excel waterfall chart -- not PPTX,
# because the team's actual downstream tool (UpSlide) links PowerPoint
# charts FROM Excel chart objects; it has no concept of a python-pptx chart,
# so a python-pptx output would be a dead end for their real workflow.
# Entirely separate code path from the main upload/process/AI/PPTX flow --
# gated behind its own session_state flag so it can't interfere with it.
import io as _bridge_lab_io

from openpyxl import Workbook as _bridge_lab_Workbook
from openpyxl import load_workbook as _bridge_lab_load_workbook

from ..bridge_chart_prototype import build_excel_waterfall_chart, find_bridge_blocks
from ..generate_bridge_waterfall_batch import build_bridges_for_ab_tab


def render_bridge_lab_toggle() -> None:
    with st.sidebar:
        st.markdown("---")
        if st.session_state.get("show_bridge_lab"):
            if st.button("← 返回主流程", use_container_width=True, key="bridge_lab_back_btn"):
                st.session_state["show_bridge_lab"] = False
                st.rerun()
        else:
            if st.button("🧪 橋圖測試 (Bridge Lab)", use_container_width=True, key="bridge_lab_enter_btn"):
                st.session_state["show_bridge_lab"] = True
                st.rerun()


def _bridge_lab_show_block(index: int, block, note: str = "") -> bool:
    """Renders one detected bridge block as a table plus a chart, and returns
    whether it's usable (check passed, or unverified -- never a hard mismatch).

    Previously printed one st.write per item, which for 11 items across
    several transitions filled the page with lines nobody reads. A table and
    the chart itself are what actually get checked, so that is all this shows;
    the per-item source notes still travel into the downloaded workbook."""
    header = f"區塊 {index + 1}：{block.items[0].label} → {block.items[-1].label}"
    with st.expander(header, expanded=True):
        if note:
            st.caption(note)

        rows = []
        running = 0.0
        for it in block.items:
            if it.kind == "total":
                running = it.value
                rows.append({"項目": it.label, "類型": "合計", "金額 (千元)": round(it.value, 1),
                             "累計": round(running, 1)})
            else:
                running += it.value
                rows.append({"項目": it.label, "類型": "變動", "金額 (千元)": round(it.value, 1),
                             "累計": round(running, 1)})
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

        try:
            _bridge_lab_render_preview_chart(block, header)
        except Exception as exc:  # a preview failure must not block the download
            st.caption(f"（圖表預覽無法產生：{exc}）")

        if block.check_ok is True:
            st.success("✅ 核對一致")
            return True
        if block.check_ok is False:
            st.error("❌ 核對不一致 -- 不會為此區塊生成圖表，請檢查來源表格")
            return False
        st.warning("⚠️ 無法核對一致性，仍會生成圖表，請自行核對數字")
        return True


def _bridge_lab_render_preview_chart(block, title: str) -> None:
    """On-screen waterfall preview using the same invisible-base-series maths
    as the downloaded Excel chart, so what is checked here is what ships."""
    from ..bridge_chart_prototype import _compute_waterfall_series

    categories, base_vals, total_vals, inc_vals, dec_vals = _compute_waterfall_series(block)
    frame = pd.DataFrame(
        {
            "（基準）": base_vals,
            "合計": total_vals,
            "增加": inc_vals,
            "減少": dec_vals,
        },
        index=categories,
    )
    st.bar_chart(frame, stack=True, height=340,
                 color=("#00000000", "#00338D", "#6D2077", "#00A3A1"))


def render_bridge_lab() -> None:
    st.title("🧪 橋圖測試 (Bridge Chart Lab)")
    st.caption(
        "實驗性功能，與主流程完全獨立。支援兩類 tab：(1) 已建好 Base/Change 輔助區塊的橋圖表，"
        "(2) AB- 原始數據表（自動計算價/量/天數因子分解）。選擇 tab 後自動偵測、生成原生 Excel "
        "疊加圖，供下載後透過 UpSlide 帶入 PPT。"
    )

    uploaded = st.file_uploader("上傳 Excel 檔案 (.xlsx)", type=["xlsx"], key="bridge_lab_upload")
    if not uploaded:
        st.info("請先上傳一個 .xlsx 檔案。")
        return

    file_bytes = uploaded.getvalue()
    try:
        wb_values = _bridge_lab_load_workbook(_bridge_lab_io.BytesIO(file_bytes), data_only=True)
    except Exception as exc:
        st.error(f"無法讀取此 Excel 檔案：{exc}")
        return

    sheet_names = wb_values.sheetnames
    selected_sheet = st.selectbox("選擇 tab", sheet_names, key="bridge_lab_sheet")

    if not st.button("偵測並生成", type="primary", key="bridge_lab_detect_btn"):
        return

    ws = wb_values[selected_sheet]

    # Route by tab type. A pre-built Base/Change helper block (like a
    # <entity>-量价桥图 tab) is read directly; otherwise try to treat it as a
    # raw AB-* data tab and COMPUTE the factor decomposition. Both paths
    # converge on a list of BridgeBlock objects rendered identically.
    renderable = []  # list of BridgeBlock that passed (or lack) their check
    prebuilt = find_bridge_blocks(ws)
    if prebuilt:
        st.success(f"偵測到 {len(prebuilt)} 個預建 Base/Change 橋圖區塊。")
        for i, block in enumerate(prebuilt):
            if _bridge_lab_show_block(i, block):
                renderable.append(block)
    else:
        ab_blocks, results = build_bridges_for_ab_tab(ws, selected_sheet)
        if ab_blocks is None:
            st.warning(
                f"「{selected_sheet}」既不是 Base/Change 結構的橋圖表，也不是可識別的 AB- 原始數據表"
                "（找不到 Year/Days 標籤列或分期區塊）。請確認選對了 tab。"
            )
            return
        if not results:
            st.warning(
                f"「{selected_sheet}」偵測為 AB- 原始數據表，但算不出任何年度轉換（可能只有單一年度資料）。"
            )
            return
        st.success(
            f"偵測為 AB- 原始數據表，已計算出 {len(results)} 個年度轉換橋圖"
            "（價/量/天數因子分解，末期採 LTM 滾動12個月口徑）。"
        )
        for i, res in enumerate(results):
            note = "註：末期為不完整年度，已改用 LTM 滾動12個月窗口比較" if res.is_ltm else ""
            if _bridge_lab_show_block(i, res.bridge, note=note):
                renderable.append(res.bridge)

    if not renderable:
        st.error("沒有通過核對、可生成圖表的區塊。")
        return

    # STANDALONE output workbook -- we build a brand-new Workbook and NEVER
    # re-save the user's upload, so their original file's formulas, cached
    # values, and existing native charts are physically untouched (openpyxl
    # round-tripping a workbook silently drops every formula's cached result,
    # which is what made an earlier version appear to "change" old tabs).
    out_wb = _bridge_lab_Workbook()
    out_wb.remove(out_wb.active)
    out_ws = out_wb.create_sheet("Bridge_Output")
    next_row = 1
    for block in renderable:
        title = f"{block.items[0].label} → {block.items[-1].label}"
        next_row = build_excel_waterfall_chart(out_ws, block, title, start_row=next_row)

    out_buffer = _bridge_lab_io.BytesIO()
    out_wb.save(out_buffer)
    st.download_button(
        "下載含橋圖的 Excel",
        data=out_buffer.getvalue(),
        file_name="bridge_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="bridge_lab_download_btn",
    )
    st.caption(
        "此為獨立的新檔案，只含生成的「Bridge_Output」分頁（結構化數據表＋橋圖）；"
        "你上傳的原始檔案完全沒有被修改。"
    )

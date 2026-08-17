from __future__ import annotations

from .helpers import (  # lifted out of PowerPointGenerator
    _account_cost_key,
    _account_is_chinese,
    _add_commentary_slot_shape,
    _apply_east_asian_line_breaking,
    _build_clause_segments,
    _build_presentation_table_plan,
    _category_to_rgb,
    _expand_commentary_to_cover_summary,
    _explanation_render_text,
    _fill_content_shape,
    _fit_table_columns,
    _force_no_autofit,
    _format_table_value,
    _insert_category_header_rows,
    _is_commentary_text_shape,
    _jieba_word_boundary_snap,
    _measurer_family,
    _merge_contd_pairs,
    _planning_std_lh_pt,
    _prepare_structured_data_for_slides,
    _presentation_table_for_account,
    _process_markdown_content,
    _read_table_style_id,
    _real_font_size_pt,
    _real_line_spacing,
    _real_para_gap_pt,
    _rendered_bullet_label,
    _resolve_font_metrics_path,
    _resolve_table_target_shape,
    _set_cell_border,
    _set_paragraph_left_indent,
    _set_table_style_id,
    _shape_has_table,
    _shape_name,
    _sublist_text_for_table,
    _table_source_multiplier,
    _table_unit_label,
    _textbox_usable_and_inset_pt,
    _truncate_text_at_boundary,
    find_content_shape,
    find_shape_by_name,
    replace_text_preserve_formatting,
)

# re-added: bound by an import in another section of the pre-split module
from ..keyword_registry import (
    STATEMENT_ORDER_SKIP_KEYWORDS,
    SUMMARY_ACCOUNT_SKIP_KEYWORDS,
    translate_category_to_chinese,
    translate_statement_line_to_chinese,
)
from ..financial_common import (
    contains_chinese_text,
    contains_predominantly_chinese_text,
    get_pipeline_result_text,
    load_yaml_file,
    package_file_path,
)
from ..workbook import find_mapping_key
import traceback

"""
PowerPoint Generation Module for Financial Reports
Based on the backup methods but implemented fresh for the new system
"""

from .text import detect_chinese_text, get_font_name_for_text, get_font_size_for_text, get_line_spacing_for_text, get_space_after_for_text, get_space_before_for_text
from .payloads import _load_pptx_settings, _looks_like_blocked_ai_content, _merge_nested_dict, _normalize_slide_commentary_text, _split_text_sentences, _translate_statement_row_label

import os
import re
import logging
import threading
import time
from typing import Any, Dict, List, Optional, Tuple
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE_TYPE

logger = logging.getLogger(__name__)
logger.setLevel(logging.WARNING)




class _MeasurementMixin:
    """How tall text and tables come out: chars per line, slot capacity, paragraph
    height, and the table-block estimates the packer charges against a slot.

    Mixed into PowerPointGenerator; `self` is the generator.
    """

    def _estimate_chars_per_line(
        self,
        slot_name: str,
        is_chinese: bool,
        shape=None,
        *,
        statement_type: Optional[str] = None,
    ) -> int:
        """Return the number of characters that fit on one line.

        When the actual shape is available we measure directly from its width
        and the text-frame insets, then divide by the known average character
        width for the font in use.  This removes all dependency on the
        ``chars_per_line`` config block for shapes that exist in the template.

        Falls back to the config-based estimate only when no shape is supplied.
        """
        if shape is not None and hasattr(shape, "width"):
            width_pt = shape.width * 72 / 914400
            # Read actual text-frame left/right insets; default is 0.1 in = 7.2 pt.
            left_pt = right_pt = 7.2
            try:
                tf = shape.text_frame
                if tf.margin_left is not None:
                    left_pt  = tf.margin_left  * 72 / 914400
                if tf.margin_right is not None:
                    right_pt = tf.margin_right * 72 / 914400
            except Exception:
                pass
            effective_pt = max(10.0, width_pt - left_pt - right_pt)
            avg_char = self._AVG_CHAR_WIDTH_CHI if is_chinese else self._AVG_CHAR_WIDTH_ENG
            return max(16, int(effective_pt * self._WORD_WRAP_SLACK / avg_char))

        # No shape — fall back to config values.
        packing = self._packing_settings(statement_type)
        chars_per_line = packing.get("chars_per_line") or {}
        slot_key = slot_name if slot_name in {"single", "L", "R"} else "default"
        language_key = "chi" if is_chinese else "eng"
        base_value = (
            ((chars_per_line.get(slot_key) or {}).get(language_key))
            or ((chars_per_line.get("default") or {}).get(language_key))
            or (32 if is_chinese else 60)
        )
        return int(base_value)


    def _slot_capacity_pt(self, slide_idx: int, slot_name: str, start_slide: int) -> float:
        """The REAL usable height of one commentary slot, read from the
        template shape it will actually render into. Falls back to the
        shared constant when the shape can't be resolved.

        A single shared constant can't be right for both slot types -- a
        real template's first-page "single" slot is ~330pt while its
        continuation L/R slots are ~359pt, so one number is either 8%
        too small for one or too large for the other."""
        try:
            slide = self.presentation.slides[start_slide - 1 + slide_idx]
            shape = self._resolve_commentary_slot_shape(slide, slot_name)
            if shape is not None:
                usable_pt, _inset_pt = _textbox_usable_and_inset_pt(shape)
                if usable_pt > 1.0:
                    return usable_pt
        except Exception as exc:
            logger.debug("Could not resolve real capacity for slot %s/%s: %s",
                         slide_idx, slot_name, exc)
        return self._TABLE_SLOT_CAPACITY_PT


    def _measurement_slot_shape(self):
        """A representative commentary slot shape, used ONLY to give text
        measurement a real width when a block's height must be estimated
        before its column has been chosen.

        Passing shape=None instead makes the measurer fall back to a default
        width, and on real content that is worth a whole wrapped line: a
        measured lead-in came out 24.1pt against a real 35.4pt, which is
        exactly the amount a table then overshot the column by. Resolving a
        specific slot would be circular (the estimate decides the column),
        but it doesn't need to be: every commentary slot in the template is
        the same 4.78in wide -- only their HEIGHTS differ -- and wrapping
        depends on width alone."""
        cached = getattr(self, "_measurement_slot_shape_cache", "unset")
        if cached != "unset":
            return cached
        shape = None
        try:
            for slide in self.presentation.slides:
                for candidate in slide.shapes:
                    if (_shape_name(candidate) or "").startswith("textMainBullets"):
                        shape = candidate
                        break
                if shape is not None:
                    break
        except Exception as exc:
            logger.debug("Could not resolve a measurement slot shape: %s", exc)
            shape = None
        self._measurement_slot_shape_cache = shape
        return shape


    def _measure_paragraph_pt(self, text: str, shape, is_chinese: bool) -> float:
        """Rendered height of ONE paragraph's own lines (excluding its
        space_after), measured with the same wrap rules the renderer uses:
        a "■ ..." bullet hangs, so its FIRST line spans the full box width
        and only continuation lines are narrower."""
        try:
            from fdd_utils.text_metrics import get_measurer, text_box_from_shape
            packing = self._packing_settings()
            measurer = get_measurer(
                _measurer_family(is_chinese, packing),
                _real_font_size_pt(is_chinese), is_cjk=is_chinese,
                line_spacing=_real_line_spacing(is_chinese),
                metrics_path=_resolve_font_metrics_path(is_chinese, packing),
            )
            box = text_box_from_shape(shape)
            hang_w = max(10.0, box.width_pt - self._BULLET_HANGING_INDENT_PT)
            n = max(1, len(measurer.wrap(
                text, hang_w,
                first_line_width_pt=box.width_pt if text.lstrip().startswith("■") else None,
            )))
            return n * measurer.line_height_pt()
        except Exception as exc:
            logger.debug("Could not measure paragraph: %s", exc)
            return _planning_std_lh_pt(is_chinese)


    def _calculate_max_lines_for_textbox(
        self,
        shape,
        *,
        is_chinese: bool = False,
        slot_name: str = "single",
        statement_type: Optional[str] = None,
    ):
        """Return the number of 'line units' that fit in this text box.

        Measures the effective height directly from the shape and its
        text-frame insets (top/bottom margins), then divides by the standard
        line height used by ``_calculate_content_lines``:

            std_line_h = font_size × line_spacing + PARA_SPACE_AFTER

        Both capacity and content are expressed in the same unit so the fill
        ratios are accurate without any fudge factors.
        """
        packing = self._packing_settings(statement_type)
        if not shape or not hasattr(shape, "height"):
            return int(packing.get("minimum_slot_lines", 20) or 20)

        font_size_pt = _real_font_size_pt(is_chinese)
        line_spacing = _real_line_spacing(is_chinese)
        family       = _measurer_family(is_chinese, packing)

        # ── Real font metrics via text_metrics ───────────────────────────────────
        # Prefer the client's font (metrics.json) so line height matches what the
        # client's PowerPoint renders; else the resolved system font. text_box_from_shape
        # reads bodyPr tIns/bIns directly from shape XML.
        try:
            from fdd_utils.text_metrics import get_measurer, text_box_from_shape
            _mpath   = _resolve_font_metrics_path(is_chinese, packing)
            measurer = get_measurer(
                family, font_size_pt, is_cjk=is_chinese, line_spacing=line_spacing,
                metrics_path=_mpath,
            )
            self._log_measurer_source_once(measurer, _mpath, is_chinese)
            box      = text_box_from_shape(shape)
            std_lh   = measurer.line_height_pt() + _real_para_gap_pt(is_chinese)
            # Float, not int(...) floored -- content cost (_calculate_content_
            # lines / _compute_slot_used_lines) is ALSO measured in these same
            # std_lh units and never rounds, so "N lines fit" and "the content
            # that's been packed in totals N.xx units" are already compared as
            # floats everywhere except here. Flooring only this side throws
            # away up to a full std_lh unit (line + its trailing gap) of real,
            # physically-there box height on every single slot -- verified via
            # a real Windows client-metrics export: a box the DP stopped
            # filling at "believed" 95.2% (15.23 / floor(16.77)=16) was
            # actually only 90.8% full relative to its own true height
            # (15.23 / 16.77) once the discarded 0.77-line margin is counted.
            # This value is exactly self-consistent with the box's real pt
            # height by construction (capacity_float * std_lh == box.height_pt),
            # so packing right up to it can never overflow the box -- the
            # floor was never a safety margin, just an unnecessary display-
            # style rounding that leaked into the fit-decision math.
            return max(1.0, box.height_pt / std_lh)
        except Exception:
            pass   # font file missing — fall through to heuristic

        # ── Heuristic fallback ───────────────────────────────────────────────────
        height_pt    = shape.height * 72 / 914400
        top_pt = bottom_pt = 3.6          # OOXML default tIns/bIns = 0.05" = 3.6 pt
        try:
            tf = shape.text_frame
            if tf.margin_top    is not None: top_pt    = tf.margin_top    * 72 / 914400
            if tf.margin_bottom is not None: bottom_pt = tf.margin_bottom * 72 / 914400
        except Exception:
            pass
        effective_pt = max(1.0, height_pt - top_pt - bottom_pt)
        std_lh       = font_size_pt * line_spacing + self._PARA_SPACE_AFTER
        max_rows     = int(effective_pt / std_lh)
        return max(int(packing.get("minimum_slot_lines", 20) or 20), max_rows)


    def _calculate_content_lines(
        self,
        category: str,
        mapping_key: str,
        commentary: str,
        *,
        slot_name: str = "single",
        shape=None,
        is_chinese: Optional[bool] = None,
        statement_type: Optional[str] = None,
        whole_box: bool = False,
    ) -> float:
        """Return the physical height of this content expressed in std_lh units.

        whole_box=True means this call measures ALL the text in one shape,
        so the final paragraph's trailing space_after can be dropped -- it
        renders as invisible padding at the bottom of the frame and pushes
        nothing. Default False, because the packer calls this ONCE PER
        ACCOUNT for slots that hold several: there, every account's
        trailing gap except the very last one is real spacing that
        separates it from the next account. Dropping it per-account
        under-counted a 7-account slot by 6 gaps (~1.3 lines), which is
        exactly the residual that had the DP reporting 100% on a slot
        PowerPoint renders at 103%.

        Returns a *float* (no ceil) so that the DP and greedy fill can track
        actual physical space consumed.  Using ceil was inflating every
        multi-line account to the next integer boundary, causing the DP to
        report 100 % fill when the box was only ~75 % physically used.

        One "unit" = std_lh = line_h + PARA_SPACE_AFTER (17 pt for English).
        Capacity from _calculate_max_lines_for_textbox is int(box_height/std_lh),
        so comparing float content against int capacity gives physically accurate
        fill ratios.
        """
        is_chinese = any('\u4e00' <= c <= '\u9fff' for c in commentary) if is_chinese is None else is_chinese

        # Memoize per-instance — Pillow font measurement runs ~80 calls per
        # paragraph and the same (account, slot, shape) tuple is asked many
        # times across the greedy distribute, the DP optimizer, and any
        # post-processing logging. Cache lookup is keyed on shape width
        # (the only shape attribute that affects line wrapping).
        if not hasattr(self, "_content_lines_cache"):
            self._content_lines_cache = {}
        shape_w = int(getattr(shape, "width", 0) or 0) if shape is not None else 0
        cache_key = (
            bool(category), mapping_key, commentary, slot_name, shape_w,
            is_chinese, str(statement_type or ""), whole_box,
        )
        cached = self._content_lines_cache.get(cache_key)
        if cached is not None:
            return cached

        font_size_pt = _real_font_size_pt(is_chinese)
        line_spacing = _real_line_spacing(is_chinese)
        para_gap     = _real_para_gap_pt(is_chinese)
        packing = self._packing_settings(statement_type)
        family = _measurer_family(is_chinese, packing)

        # -- Real glyph metrics via text_metrics --
        # Uses get_measurer (client metrics.json when configured, else the
        # resolved system font) -- the SAME measurer _calculate_max_lines_for_textbox
        # uses for capacity. Previously this called get_font()/wrap_paragraph()
        # directly with a hardcoded "Arial"/"Microsoft YaHei" system font,
        # ignoring font_metrics_path_eng/chi entirely -- capacity and content
        # were measured with two different rulers whenever a client metrics
        # file was configured, which is exactly the kind of quiet mismatch
        # that produces a "few rows off" fill-ratio gap despite already having
        # real font metrics available.
        if shape is not None:
            try:
                from fdd_utils.text_metrics import get_measurer, text_box_from_shape
                _mpath   = _resolve_font_metrics_path(is_chinese, packing)
                measurer = get_measurer(
                    family, font_size_pt, is_cjk=is_chinese, line_spacing=line_spacing,
                    metrics_path=_mpath,
                )
                self._log_measurer_source_once(measurer, _mpath, is_chinese)
                box      = text_box_from_shape(shape)
                line_h   = measurer.line_height_pt()
                std_lh   = line_h + para_gap

                total_pt = 0.0
                if category:
                    total_pt += line_h          # category header: no space_after

                paras = [p for p in commentary.split('\n') if p.strip()] if commentary else []
                key_prefix = f"\u25a0 {mapping_key} - "
                # Hanging indent: wrapped CONTINUATION lines render 0.15"
                # narrower than the box (see _BULLET_HANGING_INDENT_PT), but
                # line 1 of the account's own first paragraph (p_key) spans
                # the FULL box width -- first_line_indent=-0.15" cancels
                # left_indent=0.15" for that one line only. Only paras[0]
                # (rendered as p_key) gets this exception; paras[1:] render
                # as p_text with first_line_indent=0, i.e. narrow throughout.
                wrap_w = max(10.0, box.width_pt - self._BULLET_HANGING_INDENT_PT)
                if paras:
                    first_wrapped = measurer.wrap(
                        key_prefix + paras[0], wrap_w, first_line_width_pt=box.width_pt,
                    )
                    total_pt += len(first_wrapped) * line_h + para_gap
                    for para in paras[1:]:
                        wrapped = measurer.wrap(para, wrap_w)
                        total_pt += len(wrapped) * line_h + para_gap
                else:
                    total_pt += line_h + para_gap

                # Drop the LAST paragraph's trailing gap: space_after on the
                # final paragraph renders as invisible padding at the
                # bottom of the frame and pushes nothing, so counting it
                # inflates every block by one para_gap. Confirmed against
                # PowerPoint's own BoundHeight on a real export -- a
                # 27-line, 8-paragraph box measured 309.6pt where this
                # function claimed ~327pt, and inspect_pptx.py's matching
                # copy of the same over-count was reporting a false
                # "102% OVERFLOW RISK" on a box that is really 93.5% full.
                if paras and whole_box:
                    total_pt -= para_gap

                # Return float -- no ceil so actual physical proportion is preserved.
                result = total_pt / std_lh
                self._content_lines_cache[cache_key] = result
                return result
            except Exception:
                pass    # font file missing -- fall through to heuristic

        # ── Heuristic fallback (no shape or font unavailable) ───────────────────
        space_after  = self._PARA_SPACE_AFTER
        std_lh       = font_size_pt * line_spacing + space_after
        cpl          = self._estimate_chars_per_line(slot_name, is_chinese, shape=shape,
                                                     statement_type=statement_type)
        total_pt     = 0.0
        if category:
            total_pt += font_size_pt * line_spacing
        paras = [p for p in commentary.split('\n') if p.strip()] if commentary else []
        key_pfx_len  = len(str(mapping_key)) + 5
        if paras:
            first_len    = key_pfx_len + len(paras[0])
            first_w      = max(1, (first_len + cpl - 1) // cpl)
            total_pt    += first_w * font_size_pt * line_spacing + space_after
            for para in paras[1:]:
                w         = max(1, (len(para) + cpl - 1) // cpl)
                total_pt += w * font_size_pt * line_spacing + space_after
        else:
            total_pt += font_size_pt * line_spacing + space_after
        result = total_pt / std_lh
        self._content_lines_cache[cache_key] = result
        return result


    def _log_measurer_source_once(self, measurer, metrics_path: Optional[str], is_chinese: bool) -> None:
        """INFO-log the text-measurement source once per language per export, so a
        server log shows unambiguously whether client-font metrics are active."""
        key = "CHI" if is_chinese else "ENG"
        logged = getattr(self, "_measurer_sources_logged", None)
        if logged is None:
            logged = self._measurer_sources_logged = set()
        if key in logged:
            return
        logged.add(key)
        detail = f" ({metrics_path})" if measurer.source == "client-metrics" else ""
        logger.info("Text measurement [%s]: %s%s", key, measurer.source, detail)


    def _estimate_lead_in_pt(self, item: Dict[str, Any], is_chinese_databook: bool = False) -> float:
        """Heuristic (no real shape) estimate of one account's lead-in block
        alone -- category header + "key - commentary" bullet -- in points.
        Shared by _estimate_table_account_block_height_pt (as its lead-in
        arm) and directly by trailing plain-text accounts appended to a
        table slot's leftover space (see _append_table_accounts_to_
        distribution's trailing_items), which have no table/source/
        explanation arms at all."""
        is_chinese = bool(item.get("is_chinese"))
        lead_in_units = self._calculate_content_lines(
            # The rendered label, NOT the raw mapping_key -- see
            # _rendered_bullet_label for why that difference is a whole
            # wrapped line on real Chinese content.
            item.get("category") or "", _rendered_bullet_label(item, is_chinese_databook),
            item.get("commentary", ""), slot_name="single",
            shape=self._measurement_slot_shape(), is_chinese=is_chinese,
            # whole_box=False: this block sits INSIDE the column's shared
            # frame with more content after it, so it really does carry its
            # own trailing paragraph gap. Only the frame's last block gets
            # that gap refunded, and that refund belongs at frame level
            # (slot_cost), not here -- charging it per account is the exact
            # bug that took 15 rounds to find in the capacity investigation.
            whole_box=False,
        )
        lead_std_lh_pt = _planning_std_lh_pt(is_chinese)
        # No safety factor and no insets. Both existed because a lead-in used
        # to render into its OWN textbox, which carries insets and needed
        # margin against a height this file computed but PowerPoint laid out.
        # Since _render_table_accounts_stack writes every lead-in as ordinary
        # paragraphs in the column's single shared frame, a lead-in costs
        # exactly what any other bullet costs -- its measured text height.
        # (Do NOT add a category-header line here: _calculate_content_lines
        # already charges one line pitch for it when `category` is non-empty,
        # verified at 10.80pt, matching the header paragraph's space_after=0.
        # Adding it again over-charged every account by a full line.)
        return lead_in_units * lead_std_lh_pt


    def _estimate_table_account_block_height_pt(
        self, item: Dict[str, Any], table: Dict[str, Any], is_chinese_databook: bool,
    ) -> float:
        """Heuristic (no real shape) estimate of one table-bearing account's
        whole block -- lead-in + table + source line + explanation -- for
        the bin-packing decision in _append_table_accounts_to_distribution.
        Mirrors the real render-time formula (_add_presentation_detail_
        table_below_text / _render_presentation_table) closely enough to be
        directionally right, without needing a real shape to measure
        against (none exists yet at planning time)."""
        return sum(self._estimate_table_account_parts_pt(item, table, is_chinese_databook))


    def _table_block_reserved_pt(self, table: Dict[str, Any], *, with_source: bool = True) -> float:
        """Vertical space one table block occupies inside the column's shared
        frame: the gap above it, the table itself, and the source caption
        _render_presentation_table draws underneath (whose box is
        _TABLE_SOURCE_LINE_PT + 2 tall).

        Single definition on purpose -- the renderer reserves this and the
        planner charges it, and the two silently drifting is exactly the
        class of bug that produced a table overhanging its column. That is
        also why `with_source` exists rather than the renderer simply skipping
        the caption: both sides have to agree on which tables carry one.

        with_source=False for every table in a column except the last. The
        caption names the source for the page, not for one table, and a column
        holding three subtables was printing "资料来源：管理层信息；毕马威分析"
        three times.

        No _TABLE_GAP_BELOW_PT: separation from the paragraph that follows
        is already provided by that paragraph's own space_before/after in
        the shared frame, and adding it here was pushing the reservation
        over a line boundary -- the extra blank line reported under every
        subtable."""
        return (
            self._TABLE_GAP_ABOVE_PT
            + self._presentation_table_height_pt(table)
            + ((self._TABLE_SOURCE_LINE_PT + 2.0) if with_source else 0.0)
        )


    def _estimate_table_account_parts_pt(
        self, item: Dict[str, Any], table: Dict[str, Any], is_chinese_databook: bool,
    ) -> Tuple[float, float, float]:
        """The same estimate broken into its three independently-placeable
        parts -- (lead-in, table+source, explanation) -- so flow() can try
        splitting an account at either boundary rather than only
        accepting or rejecting it whole."""
        lead_in_pt = self._estimate_lead_in_pt(item, is_chinese_databook)

        # Exactly what the renderer reserves -- single definition, so the two
        # can't drift (see _table_block_reserved_pt).
        table_pt = self._table_block_reserved_pt(table)

        post_table_text = item.get("_post_table_text", "")
        explain_pt = 0.0
        if post_table_text:
            explain_units = self._calculate_content_lines(
                # Measure the marker-prefixed text the renderer actually
                # writes (_append_explanation_to_frame), not the bare text --
                # "➢ " is two more characters and on a full line that is
                # enough to force one extra wrap, measured at 10.8pt (a whole
                # line) of under-estimate on a real-shaped account.
                "", "", _explanation_render_text(post_table_text, is_chinese_databook),
                slot_name="single", shape=self._measurement_slot_shape(),
                is_chinese=is_chinese_databook, whole_box=False,
            )
            # Like _estimate_lead_in_pt: no safety factor, no insets. The
            # explanation is ordinary flowing text in the column's shared
            # frame now, not a floating textbox of its own.
            explain_pt = explain_units * _planning_std_lh_pt(is_chinese_databook)

        return lead_in_pt, table_pt, explain_pt


    @classmethod
    def _presentation_table_height_pt(cls, table: Dict[str, Any]) -> float:
        """Total table height in points, from the SAME row-height constants
        _render_presentation_table sets as real row heights -- so the space
        reserved during packing and the space actually drawn can't drift
        apart from each other."""
        # Counted off the PLAN, not off table["rows"], because the plan is
        # what the renderer walks and it no longer maps one-to-one onto the
        # rows: _group_plan_rows_by_category inserts a heading row above each
        # run of same-category items. Re-deriving the count here is exactly
        # the drift this docstring warns about -- one row of silent
        # disagreement per category is enough to push a tall table off the
        # bottom of its slot. is_chinese_databook/source_multiplier change
        # only labels and scale, never the row COUNT, so any value does.
        from .helpers import _build_presentation_table_plan
        pts = cls._TABLE_TITLE_ROW_PT + cls._TABLE_HEADER_ROW_PT
        for entry in _build_presentation_table_plan(table, False, 1):
            kind = entry.get("kind")
            if kind == "total":
                pts += cls._TABLE_TOTAL_ROW_PT
            elif kind in ("child", "grouped", "group"):
                pts += cls._TABLE_CHILD_ROW_PT
            else:
                pts += cls._TABLE_DATA_ROW_PT
        return pts


    def _measure_presentation_table_column_widths_pt(
        self, plan: List[Dict[str, Any]], periods: List[str], period_labels: Dict[str, str],
        unit_label: str, is_chinese_databook: bool, available_pt: Optional[float] = None,
    ) -> List[float]:
        """Real-glyph-metrics column widths for a presentation (subtable)
        breakdown -- guarantees every cell's content fits on ONE line at its
        real rendered font size, rather than 2c6e2b1's crude per-character-
        count estimate. That estimate is what actually broke: it under-
        guessed real glyph width, a column ended up narrower than its
        content needed, PowerPoint wrapped the cell to a second line, and
        the row auto-grew past the fixed _TABLE_*_ROW_PT height that
        assumed one line -- text looked shrunk/cramped and the table's
        actual rendered height stopped matching _presentation_table_height_
        pt's estimate (the two symptoms that got 2c6e2b1 reverted).
        Reuses the SAME text_metrics measurer commentary sizing already
        trusts instead of a second, separate heuristic, so "guaranteed
        no-wrap" is a real guarantee, not another guess.

        Returns one width in points per column (label column first),
        summing to at most `available_pt` -- when content genuinely needs
        less than that (the common case, and the whole point: the real
        deck's own subtables don't fill their column either), the returned
        widths sum to LESS, and the caller renders a narrower, left-aligned
        table rather than stretching to fill the slot."""
        from fdd_utils.text_metrics import get_measurer

        packing = self._packing_settings()
        family = _measurer_family(is_chinese_databook, packing)
        metrics_path = _resolve_font_metrics_path(is_chinese_databook, packing)
        measurer_header = get_measurer(family, 7.5, is_cjk=is_chinese_databook, metrics_path=metrics_path)
        measurer_data = get_measurer(family, 7.0, is_cjk=is_chinese_databook, metrics_path=metrics_path)

        # Column 0 (label): widest of the unit-label header and every row's
        # own label text (plus child indent when that row is indented).
        label_candidates_pt = [measurer_header.text_width_pt(unit_label)]
        for entry in plan:
            # "grouped" is indented by the renderer exactly as "child" is, so
            # it has to be measured that way too -- a label measured without
            # its indent fits a column it will then overflow.
            indent_pt = (self._TABLE_CHILD_INDENT_PT
                         if entry.get("kind") in ("child", "grouped") else 0.0)
            label_candidates_pt.append(measurer_data.text_width_pt(entry.get("label", "")) + indent_pt)
        widths_pt = [max(label_candidates_pt) + self._TABLE_CELL_PADDING_PT]

        # One column per period: widest of its header label and every row's
        # own formatted value in that column.
        for period in periods:
            candidates_pt = [measurer_header.text_width_pt(period_labels.get(period, period))]
            for entry in plan:
                value = entry.get("values", {}).get(period)
                text_val = _format_table_value(value, is_numeric_column=True) if value is not None else ""
                candidates_pt.append(measurer_data.text_width_pt(text_val))
            widths_pt.append(max(candidates_pt) + self._TABLE_CELL_PADDING_PT)

        # available_pt=None returns the RAW measured need with no slot
        # clamp -- used by _precompute_uniform_table_column_widths, which
        # needs comparable per-table numbers before any per-slot scaling.
        return self._clamp_column_widths_to_available(widths_pt, available_pt)


    def _clamp_column_widths_to_available(
        self, widths_pt: List[float], available_pt: Optional[float],
    ) -> List[float]:
        """Floors each column at the legibility minimum, then scales the
        whole set down proportionally if it would overflow the slot."""
        widths_pt = [max(self._TABLE_MIN_COLUMN_PT, w) for w in widths_pt]
        total_pt = sum(widths_pt)
        if available_pt is not None and total_pt > available_pt and total_pt > 0:
            scale = available_pt / total_pt
            widths_pt = [w * scale for w in widths_pt]
        return widths_pt


    def _precompute_uniform_table_column_widths(
        self, table_items: List[Dict[str, Any]], is_chinese_databook: bool,
    ) -> None:
        """Computes one shared set of column widths per column-count, as
        the element-wise MAX of every table's own measured need, and
        caches it on the instance for _render_presentation_table.

        Sizing each subtable independently (80e7a70) correctly guaranteed
        no cell ever wraps, but made two subtables on the SAME page render
        at visibly different widths whenever their longest label differed
        -- flagged by the user as "表格要固定一種format 不要一個大一個細".
        Taking the max (rather than, say, a fixed width) keeps the
        no-wrap guarantee intact by construction: every table gets at
        least the width its own content needed.

        Deliberately grouped by column count -- tables with different
        period counts aren't visually comparable side by side anyway, and
        forcing them to a shared width would mean padding or squeezing
        columns that have no counterpart."""
        by_cols: Dict[int, List[float]] = {}
        for item in (table_items or []):
            table = item.get("_presentation_table")
            if not table:
                continue
            periods = table.get("periods") or []
            try:
                plan = _build_presentation_table_plan(
                    table, is_chinese_databook, _table_source_multiplier(item),
                )
                widths = self._measure_presentation_table_column_widths_pt(
                    plan, periods, table.get("period_labels") or {},
                    _table_unit_label(is_chinese_databook), is_chinese_databook,
                    available_pt=None,   # raw need; the per-slot clamp happens at render
                )
            except Exception as exc:
                logger.debug("Could not measure uniform width for %s: %s", item.get("mapping_key"), exc)
                continue
            n_cols = len(widths)
            existing = by_cols.get(n_cols)
            by_cols[n_cols] = widths if existing is None else [max(a, b) for a, b in zip(existing, widths)]
        self._uniform_table_col_widths_pt = by_cols


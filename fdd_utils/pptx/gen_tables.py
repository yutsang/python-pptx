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





_ISO_DATE = re.compile(r"^\s*(\d{4})-(\d{2})-(\d{2})\s*$")


def _format_header_period(label: str, is_chinese: bool) -> str:
    """Column headers read 2023年12月31日 in the analyst deliverable, not
    2023-12-31. Pure text at render time -- it touches no fill, border or font,
    so it is not the class of change that produced blank pages here before.
    Anything that is not a bare ISO date comes back untouched, which leaves the
    annualised and period labels alone.
    """
    if not is_chinese:
        return label
    m = _ISO_DATE.match(str(label or ""))
    if not m:
        return label
    y, mth, d = m.groups()
    return f"{y}年{int(mth)}月{int(d)}日"


class _TablesMixin:
    """Drawing tables onto slides: the statement table, the presentation detail
    tables and their lead-ins, continuation headings and post-table explanations.

    Mixed into PowerPointGenerator; `self` is the generator.
    """

    def _calculate_table_bounds(self, slide, target_shape=None, statement_type: str = "BS") -> Dict[str, int]:
        """Use target geometry when available, otherwise derive bounds from slide layout.

        The table top is aligned with the textMainBullets commentary body so the
        financial table and the Commentary blue box sit on the same horizontal
        baseline. The blue "Commentary"/"Table" label boxes (TextBox 10/11 in
        the template) act as headers above this baseline and are not covered
        by the table.
        """
        if target_shape is not None:
            return {
                "left": target_shape.left,
                "top": target_shape.top,
                "width": target_shape.width,
                "height": target_shape.height,
            }

        slide_width = getattr(self.presentation, "slide_width", Inches(10))
        slide_height = getattr(self.presentation, "slide_height", Inches(7.5))

        title_like_shapes = []
        body_like_shapes = []
        subtitle_shapes = []
        label_shapes = []  # Blue "Commentary"/"Table" label boxes — used as baseline anchor.
        for shape in slide.shapes:
            if not getattr(shape, "has_text_frame", False):
                continue
            name = _shape_name(shape).lower()
            try:
                label_text = (shape.text_frame.text or "").strip().lower()
            except Exception:
                label_text = ""
            if "subtitle" in name:
                subtitle_shapes.append(shape)
                body_like_shapes.append(shape)
            elif any(token in name for token in ["title", "projtitle"]):
                title_like_shapes.append(shape)
            elif any(token in name for token in ["text-commentary", "textmainbullets", "content"]):
                body_like_shapes.append(shape)
            elif label_text in ("commentary", "table"):
                label_shapes.append(shape)

        left = Inches(0.5)
        width = max(Inches(5.5), slide_width - Inches(1.0))
        top = Inches(1.45 if statement_type.upper() == "BS" else 1.6)
        height = slide_height - top - Inches(0.45)

        if title_like_shapes:
            bottom = max(shape.top + shape.height for shape in title_like_shapes)
            top = max(top, bottom + Inches(0.15))

        generic_is_layout = statement_type.upper() == "IS" and subtitle_shapes and target_shape is None
        if generic_is_layout:
            earliest_subtitle_top = min(shape.top for shape in subtitle_shapes)
            left = Inches(0.5)
            top = Inches(1.5)
            width = min(slide_width - left - Inches(0.35), int((slide_width - Inches(1.0)) * 0.5))
            height = max(Inches(2.0), earliest_subtitle_top - top - Inches(0.12))

        # Horizontal extent: widen left to include any body/label shape that
        # starts further left than the default, but STOP before any such
        # shape that sits to the right (a commentary column) — never extend
        # the table's right edge OUT to it. body_like_shapes/label_shapes are
        # always commentary text or its label, so there's no case where the
        # table should legitimately render under/over one of them. The old
        # `right_edge = max(anchor right edges)` did exactly that on a layout
        # with a single right-positioned commentary box and no left-side
        # counterpart (e.g. the first BS/IS slide, which has no separate
        # "Table" label once the table's own title row replaces it) —
        # right_edge became the commentary box's own right edge, stretching
        # the table full-width so it rendered on top of the commentary text
        # instead of stopping at the gutter before it.
        horizontal_anchors = list(body_like_shapes) + list(label_shapes)
        if horizontal_anchors:
            left = min(left, min(shape.left for shape in horizontal_anchors))
            anchors_to_the_right = [shape for shape in horizontal_anchors if shape.left > left]
            if anchors_to_the_right and not generic_is_layout:
                gutter = Inches(0.15)
                right_edge = min(shape.left for shape in anchors_to_the_right) - gutter
                width = min(width, max(Inches(2.0), right_edge - left))

        # Vertical alignment: anchor the table TOP to the "Commentary" /
        # "Table" blue label box (TextBox 10/11 in the template). This puts
        # the navy title row of the table at the exact same visible level as
        # the Commentary label on the right, replacing the need for a separate
        # "Table" label above the table.
        if label_shapes:
            label_top = min(shape.top for shape in label_shapes)
            top = max(top, label_top)
            if not generic_is_layout:
                height = max(Inches(2.5), slide_height - top - Inches(0.35))
        elif body_like_shapes:
            earliest_body_top = min(shape.top for shape in body_like_shapes)
            if not generic_is_layout:
                # No label shapes (dynamically added slide): fall back to the
                # commentary body top as the anchor.
                top = max(top, earliest_body_top)
                height = max(Inches(2.5), slide_height - top - Inches(0.35))
            else:
                height = min(height, max(Inches(2.0), earliest_body_top - top - Inches(0.12)))

        width = min(width, slide_width - left - Inches(0.25))
        height = max(Inches(2.5), min(height, slide_height - top - Inches(0.2)))
        return {
            "left": int(left),
            "top": int(top),
            "width": int(width),
            "height": int(height),
        }


    def _resolve_table_style_id(self) -> Optional[str]:
        """The table style GUID to apply to BS/IS tables.

        Priority: explicit config (pptx.table_style_id) > the style GUID of any
        table already present in the template (i.e. the firm's UpSlide style).
        Cached after first resolution. Returns None if none found (keep default).
        """
        if hasattr(self, "_cached_table_style_id"):
            return self._cached_table_style_id
        style_id = None
        try:
            style_id = (str(self.pptx_settings.get("table_style_id") or "").strip()) or None
        except Exception:
            style_id = None
        if not style_id and self.presentation is not None:
            for slide in self.presentation.slides:
                for shape in slide.shapes:
                    if getattr(shape, "has_table", False):
                        sid = _read_table_style_id(shape.table._tbl)
                        if sid:
                            style_id = sid
                            break
                if style_id:
                    break
        if style_id:
            logger.info("Applying table style GUID to BS/IS tables: %s", style_id)
        self._cached_table_style_id = style_id
        return style_id


    def _add_table_to_slide(self, slide, df, bounds: Dict[str, int], table_name: str = None):
        total_rows = len(df) + 2 if table_name else len(df) + 1
        graphic_frame = slide.shapes.add_table(
            total_rows,
            len(df.columns),
            bounds["left"],
            bounds["top"],
            bounds["width"],
            bounds["height"],
        )
        # Auto-apply the firm's UpSlide table style (detected from the
        # template, or set via config) so new tables match existing ones.
        style_id = self._resolve_table_style_id()
        if style_id:
            try:
                _set_table_style_id(graphic_frame.table._tbl, style_id)
            except Exception as exc:
                logger.debug("Could not apply table style %s: %s", style_id, exc)
        return graphic_frame


    def _presentation_tables_enabled(self) -> bool:
        try:
            return bool((self.pptx_settings.get("presentation_tables") or {}).get("enabled", False))
        except Exception:
            return False


    def _presentation_table_style(self) -> str:
        """"table" (default, native PPTX table + dedicated stack shapes,
        see _render_table_accounts_stack) or "sublist" (config-gated
        fallback: the table dict is converted to plain indented text lines
        appended to the account's own commentary, and the account is NEVER
        pulled out of the normal packing pool at all -- it inherits every
        existing text module -- packing, splitting, cross-column
        continuation, overflow handling -- for free, at the cost of losing
        the native-table look and per-period-per-component detail). See
        _sublist_text_for_table."""
        try:
            value = (self.pptx_settings.get("presentation_tables") or {}).get("style", "table")
        except Exception:
            value = "table"
        value = str(value or "table").strip().lower()
        return value if value in ("table", "sublist") else "table"


    def _render_continuation_heading(
        self, slide, left: int, top: int, width: int, display_name: str,
        is_chinese_databook: bool,
    ) -> int:
        """Draws the compact "科目名（续）" / "name (cont'd)" line that stands
        in for an account's intro sentence when only its TABLE renders in
        this slot (the intro finished the previous column -- see
        place_table_item). Returns the bottom EMU position.

        Deliberately its own small renderer rather than reusing
        _fill_text_main_bullets_with_category_and_key with empty
        commentary: that path always writes a " - " separator after the
        account name, which would render as a dangling dash here."""
        suffix = "（续）" if is_chinese_databook else " (cont'd)"
        box = slide.shapes.add_textbox(left, top, width, Pt(self._TABLE_CONTINUATION_HEADING_PT))
        tf = box.text_frame
        tf.word_wrap = True
        # Zero the default 3.6pt top/bottom insets -- a one-line heading is
        # tight enough that they alone pushed it over its own box height.
        try:
            tf.margin_top = 0
            tf.margin_bottom = 0
        except Exception:
            pass
        p = tf.paragraphs[0]
        try:
            p.left_indent = Inches(0.15)
            p.first_line_indent = Inches(-0.15)
            p.space_before = Pt(0)
            p.space_after = Pt(0)
            p.line_spacing = 1.0
            _apply_east_asian_line_breaking(p)
        except Exception:
            pass
        run_bullet = p.add_run()
        run_bullet.text = '■ '
        run_bullet.font.size = Pt(9)
        run_bullet.font.name = 'Arial'
        self._set_east_asian_typeface(run_bullet)
        self._declare_run_language(run_bullet)
        try:
            run_bullet.font.color.rgb = RGBColor(128, 128, 128)
        except Exception:
            pass
        run_name = p.add_run()
        run_name.text = f"{display_name}{suffix}"
        run_name.font.size = Pt(9)
        run_name.font.name = 'Arial'
        self._set_east_asian_typeface(run_name)
        self._declare_run_language(run_name)
        run_name.font.bold = True
        try:
            run_name.font.color.rgb = RGBColor(0, 0, 0)
        except Exception:
            pass
        _force_no_autofit(tf)
        return top + int(Pt(self._TABLE_CONTINUATION_HEADING_PT))


    def _render_table_accounts_stack(
        self, slide, bullets_shape, account_data_list: List[Dict[str, Any]],
        is_chinese_databook: bool, statement_type: Optional[str] = None,
    ) -> None:
        """Renders one or more accounts stacked within a single slot -- one
        or more, since _append_table_accounts_to_distribution now packs
        multiple small table accounts into the same column when there's
        room, rather than always giving every table account a whole column
        regardless of how little of it real content needs (a real Crescent
        export showed 税金及附加 -- 3 components, one short explanation --
        leaving most of its column empty), AND now also flows plain
        (non-table) trailing accounts into a table's own leftover space --
        e.g. 投资收益/营业外支出 after 财务费用, matching how the real deck's
        own last IS page reads (table, then more prose, same column).

        Each account gets its own dedicated lead-in textbox (mimicking the
        "■ key - text" bullet style via _fill_text_main_bullets_with_
        category_and_key, but writing into a FRESH text frame rather than
        the slot's shared one), sized to its own measured content. A
        table-bearing account then gets its table and any post-table
        explanation immediately below (_render_presentation_table); a plain
        trailing account (no `_presentation_table`) stops after its own
        lead-in -- there is nothing else to render for it. Either way the
        next account in account_data_list starts at a running vertical
        offset below whatever the previous one drew. This is deliberately
        NOT the shared-text-frame/paragraph-flow mechanism ordinary
        accounts use below -- that assumes a slot's whole content lives in
        one text frame with nothing else interleaved, which doesn't hold
        once a table (a separate shape) sits between one account's lead-in
        and the next account's own lead-in: writing a second account's
        lead-in as another paragraph in the same shared frame would render
        it flowing directly under the FIRST account's lead-in, ignoring
        that account's table sitting (as a different shape) in between.

        bullets_shape's own text frame is left empty (matches how an
        intentionally-unused slot is already handled elsewhere) -- every
        real line of content here lives in a shape this method creates.
        Pixel-perfect positioning isn't achievable from python-pptx alone
        -- PowerPoint itself does the actual text layout, not this library
        -- so each account's own height is a reasoned estimate with the
        same margin the rest of this feature uses, not a ground-truth
        read-back.
        """
        left = int(bullets_shape.left)
        width = int(bullets_shape.width)
        current_category = None

        # ---- One shared text frame for the whole column ------------------
        # A table is a separate shape and cannot flow inside a text frame, so
        # its vertical space is RESERVED with blank paragraphs and the table
        # is floated over them. Everything else -- each account's lead-in,
        # its post-table explanation, and the NEXT account's lead-in -- stays
        # in one continuous flow inside bullets_shape's own frame.
        #
        # The previous design gave every account its own textbox stacked down
        # the column. That could never reuse the space a table left over, and
        # when a table did not fit it stranded the account's "...明细如下："
        # with nothing under it. Reserving the space inside the shared frame
        # removes both: text simply resumes below the table, in the same box.
        tf = bullets_shape.text_frame
        try:
            # MUST come first. The template ships these slots carrying
            # "Placeholder – placeholder"; the previous per-account-textbox
            # renderer cleared it and the shared-frame rewrite dropped that
            # call. Left in, it is a real rendered LINE that every table's
            # position was then computed without: a real deck put the table
            # 10.8pt (exactly one line) above where the text actually
            # started, so it clipped the lead-in by 5.8pt on every table
            # column. It also leaks the literal placeholder text into the
            # deliverable.
            tf.clear()
            tf.word_wrap = True
            _force_no_autofit(tf)
            from pptx.enum.text import MSO_VERTICAL_ANCHOR
            tf.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP
        except Exception:
            pass
        template_empty_p = tf.paragraphs[0]._p

        shape_name = _shape_name(bullets_shape) or ""
        slot_name = ("L" if shape_name.endswith("_L")
                     else "R" if shape_name.endswith("_R") else "single")
        std_lh_pt = _planning_std_lh_pt(is_chinese_databook)
        line_pitch_pt = max(1.0, std_lh_pt - _real_para_gap_pt(is_chinese_databook))
        _usable_pt, _inset_pt = _textbox_usable_and_inset_pt(bullets_shape)
        top_inset_pt = _inset_pt / 2.0

        cursor_pt = 0.0
        deferred_tables: List[Tuple[Dict[str, Any], Dict[str, Any], float]] = []

        def _measured_content_pt() -> float:
            """Height of everything CURRENTLY in the frame, measured from the
            frame's own paragraphs.

            Deliberately re-measures the whole frame instead of accumulating
            per-block estimates. A block estimate and the frame's real
            rendered height only have to disagree by one wrapped line for the
            table floated over the band below to land on top of real text --
            and the two are measured by different code paths, on different
            machines, with different font metrics (this repo's own checker
            falls back to system fonts where production uses client metrics),
            so they WILL drift. Measuring what is actually in the frame is
            the only version that can't. Same walk inspect_pptx.py and
            inspect_table_bands.py perform, so all three agree by
            construction."""
            total = 0.0
            for para in tf.paragraphs:
                ptext = para.text or ""
                sizes = [r.font.size.pt for r in para.runs if r.font.size is not None]
                gap = para.space_after.pt if para.space_after is not None else 0.0
                if not ptext.strip():
                    total += (max(sizes) if sizes else 9.0) * POWERPOINT_LINE_PITCH_FACTOR + gap
                    continue
                total += self._measure_paragraph_pt(ptext, bullets_shape, is_chinese_databook) + gap
            return total

        from fdd_utils.text_metrics import POWERPOINT_LINE_PITCH_FACTOR

        def _reserve_blank_lines(count: int, font_pt: float = 9.0) -> None:
            """`count` blank paragraphs of exactly font_pt * 1.2 each -- no
            paragraph gap and no text, so the reserved height is simply
            count * that pitch and the table dropped on top of it lands
            where the arithmetic says it will."""
            for _ in range(max(0, int(count))):
                blank_p = tf.add_paragraph()
                try:
                    blank_p.space_before = Pt(0)
                    blank_p.space_after = Pt(0)
                    blank_p.line_spacing = 1.0
                    _apply_east_asian_line_breaking(blank_p)
                except Exception:
                    pass
                blank_run = blank_p.add_run()
                blank_run.text = " "
                blank_run.font.size = Pt(font_pt)
                blank_run.font.name = 'Arial'
                self._set_east_asian_typeface(blank_run)
                self._declare_run_language(blank_run)

        def _reserve_exact(needed_pt: float) -> float:
            """Reserve `needed_pt` almost exactly, and return what was
            actually reserved.

            Rounding every reservation UP to a whole 9pt line left up to a
            full blank line visible under a table -- reported and manually
            confirmed on a real deck. A line's height is just its font size
            x 1.2, so the leftover fraction gets its own blank paragraph at
            whatever font size makes it come out right (floored at
            PowerPoint's own 1pt minimum). Over-reservation drops from up to
            10.8pt to at most ~1.2pt."""
            full = int(needed_pt // line_pitch_pt)
            _reserve_blank_lines(full)
            reserved = full * line_pitch_pt
            remainder = needed_pt - reserved
            if remainder > 0.05:
                font_pt = max(1.0, remainder / POWERPOINT_LINE_PITCH_FACTOR)
                _reserve_blank_lines(1, font_pt)
                reserved += font_pt * POWERPOINT_LINE_PITCH_FACTOR
            return reserved

        for account_data in account_data_list:
            table = account_data.get("_presentation_table")
            category = account_data.get('category', '')
            mapping_key = account_data.get('mapping_key', account_data.get('account_name', ''))
            display_name = (
                account_data.get('display_name_zh') or account_data.get('display_name', mapping_key)
                if is_chinese_databook
                else account_data.get('display_name', mapping_key)
            )
            commentary = account_data.get('commentary', '')
            clause_reviews = account_data.get('clause_reviews', [])
            is_chinese = account_data.get('is_chinese', False)

            # Which parts of this account render HERE. Absent (the normal
            # case) means the whole account. Set by
            # _append_table_accounts_to_distribution's flow().
            render_parts = account_data.get("_render_parts")
            wants_lead = render_parts is None or "lead" in render_parts
            wants_table = render_parts is None or "table" in render_parts
            wants_expl = render_parts is None or "expl" in render_parts
            post_table_text = account_data.get("_post_table_text", "")

            category_to_write = category if (category and category != current_category) else None
            if category:
                current_category = category

            if wants_lead:
                try:
                    self._fill_text_main_bullets_with_category_and_key(
                        tf, category_to_write, display_name, commentary, is_chinese,
                        is_chinese_databook=is_chinese_databook, needs_continuation=False,
                        font_size_pt=9, clause_reviews=clause_reviews,
                    )
                    # The template ships one empty paragraph; drop it as soon
                    # as real content exists so every measurement below sees
                    # only what will render.
                    if (template_empty_p is not None
                            and template_empty_p.getparent() is not None
                            and not (template_empty_p.text or "").strip()):
                        template_empty_p.getparent().remove(template_empty_p)
                        template_empty_p = None
                    cursor_pt = _measured_content_pt()
                except Exception as exc:
                    logger.warning("Could not render lead-in for account %s: %s", mapping_key, exc)
            else:
                # Continued fragment: a compact "科目名（续）" line stands in
                # for the intro sentence that ended the previous column.
                self._append_continuation_line_to_frame(
                    tf, display_name, is_chinese_databook, std_lh_pt,
                )
                cursor_pt = _measured_content_pt()

            if table and wants_table:
                deferred_tables.append(
                    (table, account_data, cursor_pt + self._TABLE_GAP_ABOVE_PT)
                )
                cursor_pt += _reserve_exact(
                    self._table_block_reserved_pt(table)
                )

            if post_table_text and wants_expl:
                self._append_explanation_to_frame(
                    tf, post_table_text, is_chinese_databook, std_lh_pt,
                    shape=bullets_shape, slot_name=slot_name, statement_type=statement_type,
                )
                cursor_pt = _measured_content_pt()

        # The frame's own template paragraph was never counted in cursor_pt,
        # so it has to go before the tables are positioned against it -- one
        # stray empty line at the top would shift every reserved band down.
        try:
            if (template_empty_p is not None
                    and template_empty_p.getparent() is not None
                    and not (template_empty_p.text or "").strip()):
                template_empty_p.getparent().remove(template_empty_p)
        except Exception:
            pass

        base_top = int(bullets_shape.top) + int(Pt(top_inset_pt))
        for table, account_data, y_pt in deferred_tables:
            try:
                self._render_presentation_table(
                    slide, left, base_top + int(Pt(y_pt)), width, table, is_chinese_databook,
                    source_multiplier=_table_source_multiplier(account_data),
                    # The explanation is now real flowing text in the shared
                    # frame, not a floating box under the table.
                    post_table_text="",
                )
            except Exception as exc:
                logger.warning(
                    "Could not render presentation table for %s: %s",
                    account_data.get("mapping_key", ""), exc,
                )


    def _render_presentation_table(
        self, slide, left: int, top: int, width: int, table: Dict[str, Any],
        is_chinese_databook: bool, source_multiplier: float = 1, post_table_text: str = "",
    ) -> int:
        """Renders `table` (extract_presentation_detail_table's return
        shape) as a native table at the given EMU position, a source-line
        caption beneath it, and -- when the model wrote any (see ai.py's
        _detail_table_guidance and _split_table_commentary above) -- the
        "-"/"➢" explanatory bullets a real deliverable places below that,
        naming each named component's provider/charging basis/contract
        terms. Returns the bottom EMU position.

        table["rows"]/["total_row"] hold values in the SAME raw-yuan
        internal scale every account's df uses (normalize_financial_schedule
        multiplies by source_multiplier -- 1000 whenever the sheet's own
        header says CNY'000/千元 -- so cross-account math and the block's
        own tie-out against the account total both work in one consistent
        unit). The table itself is rendered under its own "人民币千元"/
        "CNY'000" header for a human to read directly, so those raw-yuan
        figures are divided back down here, at render time only -- dividing
        upstream in extract_presentation_detail_table would break its own
        tie-out against the (also raw-yuan-scale) account total.
        """
        periods = table.get("periods") or []
        period_labels = table.get("period_labels") or {}
        plan = _build_presentation_table_plan(table, is_chinese_databook, source_multiplier)

        n_cols = 1 + len(periods)
        n_rows = 2 + len(plan)  # title + header + plan rows
        height = int(self._presentation_table_height_pt(table) * 12700)

        unit_label = _table_unit_label(is_chinese_databook)
        available_pt = width / 12700.0
        # Prefer the deck-wide uniform widths when they've been precomputed
        # (see _precompute_uniform_table_column_widths): sizing each table
        # to its OWN content made otherwise-identical subtables render at
        # visibly different widths on the same page (2.94in vs 2.586in on
        # a real export), which the user flagged as "一個大一個細". The
        # uniform set is the element-wise MAX across every table sharing
        # this column count, so no table's own content can wrap under it.
        uniform = (getattr(self, "_uniform_table_col_widths_pt", None) or {}).get(n_cols)
        if uniform:
            column_widths_pt = self._clamp_column_widths_to_available(list(uniform), available_pt)
        else:
            column_widths_pt = self._measure_presentation_table_column_widths_pt(
                plan, periods, period_labels, unit_label, is_chinese_databook, available_pt,
            )
        table_width = max(int(Inches(0.7)), int(sum(column_widths_pt) * 12700))

        graphic_frame = slide.shapes.add_table(n_rows, n_cols, left, top, table_width, height)
        table_shape = graphic_frame.table
        for col_idx, col_width_pt in enumerate(column_widths_pt):
            table_shape.columns[col_idx].width = Pt(col_width_pt)
        style_id = self._resolve_table_style_id()
        if style_id:
            try:
                _set_table_style_id(table_shape._tbl, style_id)
            except Exception as exc:
                logger.debug("Could not apply table style %s: %s", style_id, exc)

        DARK_BLUE = RGBColor(0x00, 0x33, 0x8D)
        WHITE = RGBColor(255, 255, 255)
        BLACK = RGBColor(0, 0, 0)
        GREY_TOTAL_FILL = RGBColor(0xD9, 0xD9, 0xD9)
        CHILD_BLUE = RGBColor(0x1F, 0x4E, 0x96)
        CHILD_ROW_FILL = RGBColor(0xF2, 0xF2, 0xF2)
        # The reference report's two banner rows are DIFFERENT blues sitting
        # directly on top of each other: a dark navy statement title, then a
        # lighter blue period header. Rendering both in the same navy (tried
        # first) merges them into one slab and loses the distinction the
        # reference draws; a separating rule loses the adjacency instead.
        # Different fills, no rule between them.
        HEADER_BAND_BLUE = RGBColor(0x1F, 0x4E, 0x96)

        def _set_cell(cell, text, *, bold=False, color=BLACK, fill=None, size_pt=7.0,
                      align=PP_ALIGN.LEFT, indent_emu=0):
            cell.text = text
            if not cell.text_frame.paragraphs:
                cell.text_frame.add_paragraph()
            p = cell.text_frame.paragraphs[0]
            if not p.runs:
                p.add_run()
            for run in p.runs:
                run.font.name = 'Arial'
                self._set_east_asian_typeface(run)
                self._declare_run_language(run)
                run.font.size = Pt(size_pt)
                run.font.bold = bold
                try:
                    run.font.color.rgb = color
                except Exception:
                    pass
            try:
                p.line_spacing = 1.0
                _apply_east_asian_line_breaking(p)
                p.alignment = align
                if indent_emu:
                    _set_paragraph_left_indent(p, indent_emu)
            except Exception:
                pass
            try:
                cell.margin_left = Inches(0.04)
                cell.margin_right = Inches(0.04)
                cell.margin_top = Inches(0.01)
                cell.margin_bottom = Inches(0.01)
                if fill is not None:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = fill
                else:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = WHITE
            except Exception:
                pass

        # Row 0: title band, merged across every column.
        table_shape.rows[0].height = Pt(self._TABLE_TITLE_ROW_PT)
        if n_cols > 1:
            table_shape.cell(0, 0).merge(table_shape.cell(0, n_cols - 1))
        title_text = table.get("title") or ""
        _set_cell(table_shape.cell(0, 0), title_text, bold=True, color=WHITE, fill=DARK_BLUE, size_pt=8.0)

        # Row 1: period header -- navy fill, white bold text, matching the
        # reference deck (its "人民币千元 | 2023年 | ..." row is filled navy,
        # not white).
        #
        # This was reverted on 2026-08-04 as a PRECAUTION while isolating a
        # blank-page bug, not because it was implicated: those blank pages
        # were the BS/IS overview pages, which this function never draws on
        # -- they come from _fill_table_placeholder, whose own revert note
        # names the same two slides and which stays reverted. Slides 4-5,
        # the only ones this function touches, never went blank. Re-applied
        # here alone so the two can be told apart if it ever recurs.
        table_shape.rows[1].height = Pt(self._TABLE_HEADER_ROW_PT)
        _set_cell(table_shape.cell(1, 0), unit_label, bold=True,
                  color=WHITE, fill=HEADER_BAND_BLUE, size_pt=7.5)
        for j, period in enumerate(periods, start=1):
            _set_cell(table_shape.cell(1, j), period_labels.get(period, period),
                      bold=True, color=WHITE, fill=HEADER_BAND_BLUE, size_pt=7.5,
                      align=PP_ALIGN.CENTER)

        # Data / child / total rows.
        for row_idx, entry in enumerate(plan, start=2):
            kind = entry["kind"]
            is_total = kind == "total"
            is_child = kind == "child"
            row_h = (self._TABLE_TOTAL_ROW_PT if is_total
                     else self._TABLE_CHILD_ROW_PT if is_child
                     else self._TABLE_DATA_ROW_PT)
            table_shape.rows[row_idx].height = Pt(row_h)
            label_color = CHILD_BLUE if is_child else BLACK
            # Reference deck (IMG_0229, pixel-sampled): indented child
            # rows carry a light NEUTRAL-grey band (~5-7% darker than
            # the white parent rows, no blue cast) on top of their blue
            # text; parent/plain rows stay white.
            label_fill = (GREY_TOTAL_FILL if is_total
                          else CHILD_ROW_FILL if is_child
                          else None)
            _set_cell(table_shape.cell(row_idx, 0), entry["label"], bold=is_total,
                      color=label_color, fill=label_fill, size_pt=7.0,
                      indent_emu=int(Inches(0.12)) if is_child else 0)
            for j, period in enumerate(periods, start=1):
                value = entry["values"].get(period)
                text_val = _format_table_value(value, is_numeric_column=True) if value is not None else ""
                _set_cell(table_shape.cell(row_idx, j), text_val, bold=is_total,
                          color=label_color, fill=label_fill, size_pt=7.0, align=PP_ALIGN.RIGHT)

        # Thin vertical borders between columns, a rule under the header
        # row, and top+bottom rules bracketing the total row -- matches the
        # BS/IS grid table's own convention (no rule on every data row).
        try:
            for r in range(n_rows):
                for c in range(n_cols):
                    cell = table_shape.cell(r, c)
                    if c > 0:
                        _set_cell_border(cell, 'left', color_rgb=RGBColor(0xBF, 0xBF, 0xBF), width=Pt(0.5))
            for c in range(n_cols):
                _set_cell_border(table_shape.cell(1, c), 'bottom', color_rgb=BLACK, width=Pt(1))
            total_row_idx = next((i for i, e in enumerate(plan, start=2) if e["kind"] == "total"), None)
            if total_row_idx is not None:
                for c in range(n_cols):
                    _set_cell_border(table_shape.cell(total_row_idx, c), 'top', color_rgb=BLACK, width=Pt(1))
                    _set_cell_border(table_shape.cell(total_row_idx, c), 'bottom', color_rgb=BLACK, width=Pt(1.25))
        except Exception as exc:
            logger.debug("Could not apply presentation-table borders: %s", exc)

        bottom = top + height
        source_box = slide.shapes.add_textbox(left, bottom, width, Pt(self._TABLE_SOURCE_LINE_PT + 2))
        source_tf = source_box.text_frame
        source_tf.word_wrap = True
        # Same spAutoFit problem as the explanation box below: left on, this
        # one-line 7pt caption grows itself and pushes everything under it down.
        _force_no_autofit(source_tf)
        source_p = source_tf.paragraphs[0]
        source_run = source_p.add_run()
        source_run.text = "资料来源：管理层信息；毕马威分析" if is_chinese_databook else "Source: Management information; KPMG analysis"
        source_run.font.name = 'Arial'
        self._set_east_asian_typeface(source_run)
        self._declare_run_language(source_run)
        source_run.font.size = Pt(7.0)
        source_run.font.italic = True
        try:
            source_run.font.color.rgb = RGBColor(0x59, 0x59, 0x59)
        except Exception:
            pass
        after_source = bottom + int(Pt(self._TABLE_SOURCE_LINE_PT + 2))

        if not post_table_text:
            return after_source

        return self._render_post_table_explanation(
            slide, left, after_source, width, post_table_text, is_chinese_databook,
        )


    def _render_post_table_explanation(
        self, slide, left: int, top: int, width: int, post_table_text: str,
        is_chinese_databook: bool,
    ) -> int:
        """Draws the "➢"/"-" explanatory bullets that follow a presentation
        table. Extracted from _render_presentation_table so the same
        rendering can also be used on its own, when the explanation is
        carried into the NEXT column while its table stays behind (see
        _append_table_accounts_to_distribution's flow()). Returns the
        bottom EMU position."""
        marker = "➢ " if is_chinese_databook else "- "
        raw_lines = [ln.strip() for ln in post_table_text.split("\n") if ln.strip()]
        if not raw_lines:
            raw_lines = [post_table_text.strip()]
        lines = [ln if ln.startswith(("➢", "-", "•")) else f"{marker}{ln}" for ln in raw_lines]

        # Placeholder height, resized below via the same real-metrics
        # measurement _render_table_accounts_stack uses for the lead-in --
        # needs a real shape to measure against first.
        explain_box = slide.shapes.add_textbox(left, top, width, Pt(200))
        explain_tf = explain_box.text_frame
        explain_tf.word_wrap = True
        # python-pptx's add_textbox ships <a:spAutoFit/> ("resize shape to fit
        # text"), which makes PowerPoint DISCARD the height computed below and
        # substitute its own -- consistently taller, because it charges the
        # last paragraph's space_after plus both insets. That is the phantom
        # blank line under this box, and it is also why the box cannot be
        # dragged smaller by hand: PowerPoint snaps it straight back.
        _force_no_autofit(explain_tf)
        for i, line_text in enumerate(lines):
            p = explain_tf.paragraphs[0] if i == 0 else explain_tf.add_paragraph()
            run = p.add_run()
            run.text = line_text
            run.font.name = 'Arial'
            self._set_east_asian_typeface(run)
            self._declare_run_language(run)
            run.font.size = Pt(9.0)
            try:
                run.font.color.rgb = RGBColor(0, 0, 0)
            except Exception:
                pass
            p.line_spacing = 1.0
            _apply_east_asian_line_breaking(p)
            p.space_after = Pt(2)

        try:
            combined = "\n".join(lines)
            used_units = self._calculate_content_lines(
                "", "", combined, slot_name="single", shape=explain_box,
                is_chinese=is_chinese_databook, whole_box=True,
            )
            capacity_units = self._calculate_max_lines_for_textbox(
                explain_box, is_chinese=is_chinese_databook, slot_name="single",
            )
            # See _render_table_accounts_stack's lead-in sizing: derive
            # std_lh from the USABLE height capacity was measured against,
            # then add the insets back on top of the text's own height.
            usable_pt, inset_pt = _textbox_usable_and_inset_pt(explain_box)
            std_lh_pt = (
                (usable_pt / capacity_units) if capacity_units > 0 else
                _real_font_size_pt(is_chinese_databook) * _real_line_spacing(is_chinese_databook)
                + _real_para_gap_pt(is_chinese_databook)
            )
            explain_height_pt = (
                max(used_units, 1.0) * std_lh_pt * self._TABLE_RENDER_HEIGHT_SAFETY_FACTOR + inset_pt
            )
            explain_box.height = int(explain_height_pt * 12700)
        except Exception as exc:
            logger.debug("Could not size presentation-table explanatory text: %s", exc)

        return top + int(explain_box.height)


    def _embed_statement_table(
        self, slide, df, statement_type: str, table_name: str = None, currency_unit: str = None,
        mappings: Optional[Dict[str, Any]] = None, is_chinese_mode: bool = False,
    ):
        # Insert category-header rows BEFORE anything below sizes the table
        # off len(df) -- the table's own row count (_add_table_to_slide /
        # the TablePlaceholder branch) is derived directly from len(df), so
        # doing this here means the extra rows are already accounted for by
        # the time the table shape itself gets created, not squeezed in
        # after the fact.
        df = _insert_category_header_rows(df, mappings, is_chinese_mode)

        target_shape = _resolve_table_target_shape(slide, statement_type)
        bounds = self._calculate_table_bounds(slide, target_shape=target_shape, statement_type=statement_type)
        target_name = _shape_name(target_shape) if target_shape is not None else "(new table)"
        logger.info(
            f"Resolved {statement_type} table target on slide using {target_name} "
            f"at left={bounds['left']} top={bounds['top']} width={bounds['width']} height={bounds['height']}"
        )

        # Remove the redundant "Table" label box (TextBox 11 in the template).
        # The table's navy title row now lives at the same vertical position
        # as that label, so keeping the label would double-print the header.
        # Leave the right-side "Commentary" label intact — there's still a
        # commentary box on this slide that needs its header.
        for shape in list(slide.shapes):
            if not getattr(shape, "has_text_frame", False):
                continue
            try:
                label_text = (shape.text_frame.text or "").strip().lower()
            except Exception:
                continue
            if label_text != "table":
                continue
            try:
                sp = shape._element
                sp.getparent().remove(sp)
            except Exception as e:
                logger.debug("Could not remove 'Table' label shape: %s", e)

        if target_shape is None:
            table_shape = self._add_table_to_slide(slide, df, bounds, table_name=table_name)
            self._fill_table_placeholder(
                table_shape,
                df,
                table_name=table_name,
                currency_unit=currency_unit,
                bounds=bounds,
                mappings=mappings,
                is_chinese_mode=is_chinese_mode,
            )
            return

        self._fill_table_placeholder(
            target_shape,
            df,
            table_name=table_name,
            currency_unit=currency_unit,
            bounds=bounds,
            mappings=mappings,
            is_chinese_mode=is_chinese_mode,
        )


    def _append_continuation_line_to_frame(
        self, text_frame, display_name: str, is_chinese_databook: bool, std_lh_pt: float,
    ) -> float:
        """Writes "科目名（续）" as a paragraph in the shared column frame and
        returns the height it consumes. The standalone
        _render_continuation_heading textbox is only used by callers that
        still draw their own boxes; inside the shared frame a heading has to
        be a real paragraph or the text after it would not flow past it."""
        label = f"{display_name}（续）" if is_chinese_databook else f"{display_name} (cont'd)"
        para = text_frame.add_paragraph()
        try:
            para.space_before = Pt(0)
            para.space_after = Pt(3)
            para.line_spacing = 1.0
            _apply_east_asian_line_breaking(para)
        except Exception:
            pass
        run = para.add_run()
        run.text = label
        run.font.size = Pt(9)
        run.font.name = 'Arial'
        self._set_east_asian_typeface(run)
        self._declare_run_language(run)
        run.font.bold = True
        try:
            run.font.color.rgb = RGBColor(0, 0, 0)
        except Exception:
            pass
        return std_lh_pt


    def _append_explanation_to_frame(
        self, text_frame, post_table_text: str, is_chinese_databook: bool, std_lh_pt: float,
        *, shape=None, slot_name: str = "single", statement_type: Optional[str] = None,
    ) -> float:
        """Writes the "➢"/"-" explanatory bullets that follow a presentation
        table as real paragraphs in the shared column frame, and returns the
        height they consume. Same content _render_post_table_explanation
        draws into its own textbox -- but as flowing text, so the next
        account's lead-in continues underneath instead of needing a whole
        new column."""
        lines = _explanation_render_text(post_table_text, is_chinese_databook).split("\n")

        for line_text in lines:
            para = text_frame.add_paragraph()
            try:
                para.space_before = Pt(0)
                para.space_after = Pt(3)
                para.line_spacing = 1.0
                _apply_east_asian_line_breaking(para)
            except Exception:
                pass
            run = para.add_run()
            run.text = line_text
            run.font.size = Pt(9)
            run.font.name = 'Arial'
            self._set_east_asian_typeface(run)
            self._declare_run_language(run)
            try:
                run.font.color.rgb = RGBColor(0, 0, 0)
            except Exception:
                pass

        return self._calculate_content_lines(
            "", "", "\n".join(lines), slot_name=slot_name, shape=shape,
            is_chinese=is_chinese_databook, statement_type=statement_type,
        ) * std_lh_pt


    def _fill_table_placeholder(
        self, shape, df, table_name: str = None, currency_unit: str = None, bounds: Dict[str, int] = None,
        mappings: Optional[Dict[str, Any]] = None, is_chinese_mode: bool = False,
    ):
        """Fill table placeholder with DataFrame data, preserving original formatting
        Args:
            shape: Table shape or placeholder
            df: DataFrame with data
            table_name: Name of the table (e.g., "示意性调整后资产负债表 - xxxx")
            currency_unit: Currency unit (e.g., "人民币千元" or "CNY'000") to replace "Description"
        """
        try:
            # Debug: Log DataFrame content
            logger.info("Filling table with DF shape: %s", df.shape)
            if not df.empty:
                logger.info("First row data: %s", df.iloc[0].to_dict())
                # Check if any data is non-zero
                numeric_cols = df.select_dtypes(include=['number']).columns
                if len(numeric_cols) > 0:
                    non_zero_count = (df[numeric_cols] != 0).sum().sum()
                    logger.info("Non-zero values in DF: %s", non_zero_count)
            
            # Find parent slide
            slide = None
            for s in self.presentation.slides:
                for shp in s.shapes:
                    if shp == shape:
                        slide = s
                        break
                if slide:
                    break
            
            if bounds is None:
                bounds = {
                    "left": shape.left,
                    "top": shape.top,
                    "width": shape.width,
                    "height": shape.height,
                }

            # Adjust position and size from resolved layout bounds
            try:
                shape.left = bounds["left"]
                shape.top = bounds["top"]
                shape.width = bounds["width"]
                shape.height = bounds["height"]
            except Exception as e:
                logger.warning("Could not adjust table position/width: %s", e)

            # Check if shape is a TablePlaceholder (textbox placeholder)
            from pptx.shapes.placeholder import TablePlaceholder
            
            table = None
            # Check if it's a TablePlaceholder (either a real pptx placeholder
            # of that type, OR a plain textbox/autoshape named "Table
            # Placeholder" used as a template bounds marker). isinstance()
            # returning False is NOT an exception -- the previous try/except
            # here never actually reached the name-based fallback, since
            # isinstance() simply returns False for a non-match rather than
            # raising, so a plain-shape placeholder (e.g. added directly via
            # python-pptx, not through the placeholder XML machinery) was
            # silently treated as "not a table placeholder" and fell through
            # to the broken "insert as plain text" error path below.
            is_table_placeholder = isinstance(shape, TablePlaceholder) or (
                'Table' in getattr(shape, 'name', '') and 'Placeholder' in getattr(shape, 'name', '')
            )
            
            if is_table_placeholder:
                # It's a table placeholder - insert a table into it
                logger.info("Found TablePlaceholder (%s), inserting table with %s rows and %s columns", shape.name if hasattr(shape, 'name') else 'unnamed', len(df), len(df.columns))
                try:
                    left = bounds["left"]
                    top = bounds["top"]
                    width = bounds["width"]
                    height = bounds["height"]
                    
                    # Find the slide containing this shape (already found above)
                    if slide:
                        # Remove the placeholder shape
                        sp = shape._element
                        slide.shapes._spTree.remove(sp)
                        
                        # Add new table at the same position
                        # Need: 1 row for title (if table_name), 1 for header, N for data
                        total_rows = len(df) + 2 if table_name else len(df) + 1
                        table_shape = slide.shapes.add_table(
                            rows=total_rows,
                            cols=len(df.columns),
                            left=left,
                            top=top,
                            width=width,
                            height=height
                        )
                        table = table_shape.table
                        logger.info("Inserted new table: %s rows, %s columns", len(table.rows), len(table.columns))
                except Exception as e:
                    logger.error("Could not insert table into placeholder: %s", e)
                    logger.debug(traceback.format_exc())
            elif hasattr(shape, 'table'):
                # Try to access existing table
                try:
                    table = shape.table
                    logger.info("Found existing table with %s rows and %s columns", len(table.rows), len(table.columns))
                except ValueError:
                    # Shape doesn't contain a table
                    logger.warning("Shape has table attribute but doesn't contain a table")
                    table = None
            
            if table:
                # Every cell's fill/font/border is set explicitly below, but a
                # freshly-created/template table shape defaults to
                # first_row=True, horz_banding=True -- PowerPoint's built-in
                # table-style theme then auto-applies its own "first row" and
                # alternating row-banding colors on top of (or in place of,
                # for rows this code doesn't explicitly touch) the direct
                # formatting, which is exactly the kind of stray theme
                # colour/stripe the plain reference table doesn't have.
                # Disable both so only the explicit per-cell styling renders.
                try:
                    table.first_row = False
                    table.horz_banding = False
                except Exception:
                    pass

                # Colors -- matches the company-format reference (real KPMG
                # deliverable page): navy #00338D fills the title band and
                # accents total-row borders; only the LAST column is
                # highlighted light blue (every row, header+data+totals);
                # every other cell stays unfilled (white bg, dark text) --
                # total rows are called out with border weight only, no fill.
                DARK_BLUE = RGBColor(0x00, 0x33, 0x8D)
                WHITE = RGBColor(255, 255, 255)
                BLACK = RGBColor(0, 0, 0)
                # Reference format (IMG_0035) calls out only the four
                # statement-level grand totals with a solid grey fill.
                GREY_TOTAL_FILL = RGBColor(0xD9, 0xD9, 0xD9)

                _fit_table_columns(table, df)

                # Smaller/tighter than before across the board -- reference
                # format (IMG_0035) reads noticeably more compact than this
                # table previously rendered at.
                #
                # Category header rows (_insert_category_header_rows --
                # "Revenue"/"Expenses"/"Current assets"/etc.) don't count
                # toward this tier threshold: they're a single short label
                # with no figures, not a real data row, and letting them
                # count at full weight was dragging the WHOLE table (every
                # real data row too) down to a smaller font/row-height tier
                # than the actual data alone would need -- e.g. the IS
                # table's 16 real rows sit comfortably under the 20-row
                # threshold on their own, but 16 rows + 6 "Revenue"/
                # "Expenses" header rows = 22 pushed it into the tighter
                # tier for no real reason.
                _tier_keywords = list(
                    {'total', '合计', '总计', '小计', 'subtotal', 'sub-total', 'sub total'}
                    | set(SUMMARY_ACCOUNT_SKIP_KEYWORDS)
                )
                _category_header_row_count = 0
                for _, _tier_row in df.iterrows():
                    _tier_label = str(_tier_row.iloc[0]).strip()
                    if not _tier_label or any(kw in _tier_label.lower() for kw in _tier_keywords):
                        continue
                    if all(pd.isna(v) or (isinstance(v, str) and not v.strip()) for v in _tier_row.iloc[1:]):
                        _category_header_row_count += 1
                total_visible_rows = (len(df) - _category_header_row_count) + 1 + (1 if table_name else 0)
                if total_visible_rows >= 26:
                    data_font_size = Pt(6)
                    data_row_height = Inches(0.13)
                elif total_visible_rows >= 20:
                    data_font_size = Pt(6.5)
                    data_row_height = Inches(0.15)
                else:
                    data_font_size = Pt(7)
                    data_row_height = Inches(0.17)

                # Title band and column-header row scale WITH the data tier
                # instead of a fixed Pt(8)/Pt(7)/0.25" -- on a dense table
                # (total_visible_rows >= 26, data shrunk to 6pt/0.13") a
                # hardcoded 0.25"/Pt(7) header reads nearly 2x taller and a
                # full point bigger than every row beneath it, exactly the
                # "header row太大" mismatch a user screenshot flagged.
                header_font_size = Pt(data_font_size.pt + 1)
                header_row_height = Inches(data_row_height.inches + 0.03)
                title_font_size = Pt(data_font_size.pt + 2)
                title_row_height = Inches(data_row_height.inches + 0.05)

                # Hard clamp: the three fixed tiers above are picked from
                # ROW COUNT alone and don't know this table's actual
                # available height (varies by template/slide) -- a table
                # with unusually many rows for its placeholder could still
                # render taller than its bounds and spill onto the slide.
                # Scale every row height (and, floored at a legibility
                # minimum, font size) down uniformly so the whole table's
                # summed row heights never exceed bounds["height"].
                try:
                    _available_h_in = max(0.1, (bounds.get("height", 0) or 0) / 914400)
                    _needed_h_in = (
                        (title_row_height.inches if table_name else 0.0)
                        + header_row_height.inches
                        + len(df) * data_row_height.inches
                    )
                    if _needed_h_in > _available_h_in > 0:
                        _scale = _available_h_in / _needed_h_in
                        data_row_height = Inches(max(0.08, data_row_height.inches * _scale))
                        header_row_height = Inches(max(0.09, header_row_height.inches * _scale))
                        title_row_height = Inches(max(0.10, title_row_height.inches * _scale))
                        data_font_size = Pt(max(4.5, data_font_size.pt * _scale))
                        header_font_size = Pt(max(5.0, header_font_size.pt * _scale))
                        title_font_size = Pt(max(5.5, title_font_size.pt * _scale))
                except Exception:
                    pass

                # Add table name as first row if provided
                if table_name:
                    # Insert a new row at the top for table name
                    try:
                        # Ensure table has at least one row
                        if len(table.rows) == 0:
                            table.rows.add_row()
                            
                        name_row = table.rows[0]  # Use first row for name
                        name_row.height = title_row_height
                        
                        # Merge all cells in first row for table name
                        if len(table.columns) > 1:
                            name_row.cells[0].merge(name_row.cells[len(table.columns) - 1])
                        name_cell = name_row.cells[0]
                        name_cell.text = table_name
                        # Company-format reference (real KPMG deliverable page):
                        # the statement title sits in a navy-filled band spanning
                        # the full table width, white bold text, left-aligned --
                        # the date/header row directly beneath it is the one
                        # left unfilled (white bg, dark text), not the title.
                        if name_cell.text_frame.paragraphs:
                            p = name_cell.text_frame.paragraphs[0]
                            p.alignment = PP_ALIGN.LEFT
                            if p.runs:
                                run = p.runs[0]
                            else:
                                run = p.add_run()
                            run.font.name = 'Arial'
                            self._set_east_asian_typeface(run)
                            self._declare_run_language(run)
                            run.font.size = title_font_size
                            run.font.bold = True
                            run.font.color.rgb = WHITE

                            name_cell.fill.solid()
                            name_cell.fill.fore_color.rgb = DARK_BLUE

                        # Shift data down - we'll use rows starting from index 1
                        data_start_row = 1
                    except:
                        data_start_row = 0
                else:
                    data_start_row = 0
                
                # Fill header row with formatting
                max_cols = min(len(df.columns), len(table.columns))
                header_row_idx = data_start_row
                
                # Ensure header row exists
                if len(table.rows) <= header_row_idx:
                    table.rows.add_row()
                
                # Set header row height slightly taller for readability
                try:
                    table.rows[header_row_idx].height = header_row_height
                except:
                    pass
                    
                # TEMPORARILY REVERTED 2026-08-04 (was HEADER_ROW_BLUE fill
                # across every column, white text, white L/R borders) while
                # isolating a real blank-page bug: two real Chinese exports
                # in a row rendered COMPLETELY BLANK in real PowerPoint on
                # exactly the two slides this function's table lands on
                # (BS/IS overview "financials" pages), with abnormally slow
                # open, while python-pptx itself read every shape/cell back
                # intact from the same files. The Commentary-band deletion
                # was the first suspect and has since been replaced with an
                # in-place text swap (see _localize_label_bands) that keeps
                # every shape's structure untouched -- but the blank pages
                # persisted even after that fix, which rules the band out
                # and leaves this function's per-cell colour/border changes
                # (the only other thing ebf2179 touched on these same two
                # pages) as the remaining suspect. Reverted to the
                # known-safe pre-ebf2179 styling (white bg, black text,
                # only the last column tinted) as a controlled test -- if
                # blank pages persist even with this reverted too, the
                # cause is elsewhere entirely (not this session's colour
                # work), and this revert should itself be re-reverted once
                # that's confirmed.
                #
                # A touch more saturated than a bare tint (#DCE6F1) so the
                # highlighted column is unmistakable at a glance, not just
                # visible on close inspection -- a user photo of a fresh
                # export still read as "hasn't taken effect" at the lighter
                # shade.
                LIGHT_BLUE_HIGHLIGHT = RGBColor(0xBD, 0xD7, 0xEE)
                # A lighter blue than the navy title band directly above it,
                # matching the reference report's two-tone banner. Merging
                # them into one navy (tried first) removed the distinction the
                # reference draws. What made the header read as a SEPARATE
                # element before was not the shade but the black cell borders
                # drawn across it -- those are white hairlines now.
                HEADER_ROW_BLUE = RGBColor(0x1F, 0x4E, 0x96)
                _header_blue = bool(
                    self.pptx_settings.get("financial_table_header_blue", True)
                )

                for col_idx, col_name in enumerate(df.columns[:max_cols]):
                    if col_idx < len(table.columns):
                        cell = table.cell(header_row_idx, col_idx)
                        # Replace "Description" with currency unit if found
                        if currency_unit and (col_name.lower() == 'description' or '描述' in str(col_name) or '项目' in str(col_name)):
                            cell.text = currency_unit
                        else:
                            cell.text = _format_header_period(str(col_name), bool(is_chinese_mode))
                        # Company-format reference: the date/header row sits
                        # directly under the navy title band with NO fill of
                        # its own (white bg, dark bold text) -- only the title
                        # row above it is filled navy. The first column's own
                        # header ("Description"/currency unit) is left-aligned
                        # like a row label; the date/period columns stay
                        # centered above their right-aligned numeric data.
                        if cell.text_frame.paragraphs:
                            p = cell.text_frame.paragraphs[0]
                            p.alignment = PP_ALIGN.LEFT if col_idx == 0 else PP_ALIGN.CENTER

                            if p.runs:
                                run = p.runs[0]
                            else:
                                run = p.add_run()

                            run.font.name = 'Arial'
                            self._set_east_asian_typeface(run)
                            self._declare_run_language(run)
                            run.font.size = header_font_size
                            run.font.bold = True
                            run.font.color.rgb = WHITE if _header_blue else BLACK
                            p.line_spacing = 1.0
                            _apply_east_asian_line_breaking(p)

                        # Blue header band with white text, matching the
                        # reference report (IMG_0224): the date/period row
                        # sits under the navy title band on its own blue
                        # fill, not on white.
                        #
                        # READ THIS BEFORE TOUCHING IT. This exact fill is
                        # what ebf2179 applied and 93c41e4 reverted: two
                        # real Chinese exports came back with the BS and IS
                        # first pages rendering COMPLETELY BLANK in real
                        # PowerPoint (slow file-open, while python-pptx still
                        # saw every shape and all text intact), and those are
                        # exactly the two pages carrying this grid. It is not
                        # reproducible on a Mac -- no real PowerPoint, and a
                        # local template whose bands are plain text boxes
                        # rather than the native placeholders production uses.
                        # Re-applied here on the user's explicit sign-off
                        # against the reference photo, but behind
                        # pptx_settings["financial_table_header_blue"] so a
                        # recurrence is one config value to flip, not a code
                        # revert -- and so the answer stays a single known
                        # switch rather than another round of bisecting.
                        cell.fill.solid()
                        if _header_blue:
                            cell.fill.fore_color.rgb = HEADER_ROW_BLUE
                        else:
                            # Only the LAST column is highlighted light blue
                            # (matches the company-format reference's
                            # "adjusted figures" column) -- every other column
                            # is EXPLICITLY set to solid white, not left
                            # unset. Leaving fill untouched makes the cell
                            # "inherit" from the table's built-in style GUID
                            # (still referenced even with first_row/
                            # horz_banding disabled), which can be a themed/
                            # tinted colour -- confirmed via
                            # inspect_pptx_tables.py showing fill=inherited on
                            # every un-highlighted cell.
                            # The deliverable tints NO column: the stage is
                            # already named in the header, so a highlight on
                            # the last one is redundant and reads as emphasis
                            # the analyst did not intend. Still set solid WHITE
                            # rather than leaving it unset -- an unset fill
                            # inherits the table style GUID's themed colour,
                            # which is the reason this branch is explicit.
                            cell.fill.fore_color.rgb = WHITE

                        # Vertical (column-separating) borders plus ONE
                        # bottom rule under the header row -- horizontal
                        # rules on every data row read as cluttered once the
                        # table has 20+ rows; the total/subtotal rows below
                        # get their own explicit top/bottom rule so those
                        # separators are still there where they matter.
                        # Black column separators drawn ACROSS a navy header
                        # band read as a grid pasted over the title, which is
                        # the "header 同藍色黐唔埋" complaint. Use white hairlines
                        # between the date columns instead -- visible against
                        # navy, invisible as a boundary -- and keep the black
                        # rule only where it does real work: under the band,
                        # separating it from the data.
                        # No vertical separators: the deliverable's table has
                        # none in the data area at all -- the columns are held
                        # apart by alignment and cell margins, and the only
                        # rules are horizontal, under the header band and above
                        # each total. Drawing them made the table read as a
                        # spreadsheet rather than a report page.
                        _set_cell_border(cell, "bottom", color_rgb="000000", width=Pt(0.5))

                        try:
                            cell.margin_left = Inches(0.04)
                            cell.margin_right = Inches(0.04)
                            cell.margin_top = Inches(0.02)
                            cell.margin_bottom = Inches(0.02)
                        except Exception:
                            pass

                        logger.debug("Filled header cell %s: %s", col_idx, cell.text)
                
                # Fill data rows with formatting - show ALL rows (no limit)
                # Check if table has enough rows, if not, limit to available rows
                max_rows = len(df)  # Show all rows
                rows_needed = max_rows + data_start_row + 1  # +1 for header row
                available_rows = len(table.rows)
                
                if available_rows < rows_needed:
                    logger.warning("Table has %s rows but needs %s. Will only fill %s data rows.", available_rows, rows_needed, available_rows - data_start_row - 1)
                    max_rows = min(max_rows, available_rows - data_start_row - 1)
                    if max_rows < 0:
                        max_rows = 0
                
                logger.info("Table has %s rows available, will fill %s data rows", available_rows, max_rows)
                
                # Now fill all rows with Arial 9 font
                # Check for title, date, total, and subtotal rows to highlight
                logger.info("Filling %s data rows, starting at row index %s, table has %s rows", max_rows, header_row_idx + 1, len(table.rows))
                for row_idx in range(max_rows):
                    if row_idx >= len(df):
                        break
                    df_row = df.iloc[row_idx]
                    first_col_value = str(df_row.iloc[0]) if len(df_row) > 0 else ""

                    # A section header ("流动资产" / "非流动负债" / etc.) is a
                    # row the source Financials sheet gives a label but no
                    # figures at all -- blank on every numeric column, not
                    # just a zero (a genuine zero renders as "-" via
                    # _format_table_value, so it's distinguishable from a
                    # truly blank/missing value). "Blank" covers both
                    # pd.isna() (None/NaN) AND an empty/whitespace-only
                    # string -- a real extracted sheet's blank cell isn't
                    # guaranteed to come through as NaN specifically (could
                    # already be "" depending on how the extractor read it),
                    # and pd.isna("") is False, so a NaN-only check silently
                    # never fires on that variant. Reference format (real
                    # company deliverable, IMG_0035): these sit flush left
                    # with no indent, everything else between one and the
                    # next header/total is indented under it.
                    def _is_blank_cell(v) -> bool:
                        if pd.isna(v):
                            return True
                        if isinstance(v, str) and not v.strip():
                            return True
                        return False

                    is_category_header_row = bool(first_col_value.strip()) and all(
                        _is_blank_cell(df_row[col]) for col in df.columns[1:max_cols] if col in df_row.index
                    )

                    # Check if this is a title, total, or subtotal row
                    is_special_row = False
                    is_total_row = False
                    first_col_lower = first_col_value.lower()
                    # 'total'/合计 etc. cover BS lines (Total assets, Total
                    # current assets, ...); IS running-total lines (Gross
                    # profit, Operating profit, Net profit, ...) don't contain
                    # the word "total" at all, so pull in the canonical
                    # SUMMARY_ACCOUNT_SKIP_KEYWORDS list (already used
                    # elsewhere in the codebase for "is this a statement
                    # subtotal/total line, not a leaf account" detection)
                    # rather than hand-rolling a second copy of it here.
                    total_keywords = list(
                        {'total', '合计', '总计', '小计', 'subtotal', 'sub-total', 'sub total'}
                        | set(SUMMARY_ACCOUNT_SKIP_KEYWORDS)
                    )
                    # NOTE: previously also matched a 'date_keywords' list
                    # ('date'/'日期'/'年'/'月') into special_keywords AND a
                    # separate is_date_row flag that forced a fixed Pt(7)
                    # font. The actual date/period COLUMN HEADER row (e.g.
                    # "2023-12-31") is a different row entirely, handled
                    # above via header_row_idx -- this per-data-row loop only
                    # ever sees ACCOUNT LABELS, and '年'/'月' as bare
                    # substrings match ordinary Chinese account names like
                    # "一年内到期的非流动负债" ("non-current liabilities due
                    # within one year"), wrongly bolding them AND overriding
                    # their font size to Pt(7) regardless of the table's
                    # actual density tier (data_font_size could be 6/6.5/7pt)
                    # -- a real, visible font-size/weight mismatch against
                    # every other leaf row on the same table.
                    special_keywords = total_keywords + ['title', '标题']

                    if any(keyword in first_col_lower for keyword in special_keywords):
                        is_special_row = True

                    if any(keyword in first_col_lower for keyword in total_keywords):
                        is_total_row = True

                    # A blank-values row that ALSO matches a total keyword is
                    # a total row that simply has no figures yet (e.g. a
                    # subtotal for a section with no populated accounts),
                    # not a section header -- keep those two mutually
                    # exclusive so the total-row border/fill logic below
                    # still applies to it.
                    is_category_header_row = is_category_header_row and not is_total_row
                    if is_category_header_row:
                        is_special_row = True

                    # Statement-terminal lines (grand totals) get the heavier
                    # two-border/grey-fill treatment; everything else that only
                    # matches total_keywords (subsection subtotals like "Total
                    # current assets", or IS running subtotals like "Gross
                    # profit") gets a thin top border only, matching the
                    # company-format reference where those two tiers look
                    # visually distinct.
                    grand_total_keywords = [
                        'total assets', 'total liabilities', "total owners", "total owner's",
                        '总资产', '负债合计', '负债总计', '所有者权益合计',
                        '股东权益合计', '资产总计', 'net profit', 'net loss', '净利润', '净亏损',
                        'ebitda',
                    ]
                    # "Total equity attributable to owners of the Company" is a
                    # subtotal, not the statement-level grand total, even though
                    # it contains "total ... owners" as a substring (same trap
                    # applies to the Chinese "归属于母公司所有者权益合计" pattern) --
                    # exclude any label signalling it's scoped to a sub-group.
                    attributable_signal_keywords = ['attributable', '归属', '母公司']
                    is_attributable_subtotal = any(
                        keyword in first_col_lower for keyword in attributable_signal_keywords
                    )
                    is_grand_total_row = (
                        is_total_row
                        and not is_attributable_subtotal
                        and any(keyword in first_col_lower for keyword in grand_total_keywords)
                    )
                    
                    # Data row index = header_row_idx + 1 + row_idx
                    data_row_idx = header_row_idx + 1 + row_idx
                    if data_row_idx >= len(table.rows):
                        logger.warning("Data row index %s exceeds table rows %s, skipping", data_row_idx, len(table.rows))
                        break
                    
                    # Set data row height based on table density
                    try:
                        table.rows[data_row_idx].height = data_row_height
                    except:
                        pass
                    
                    # Log first row processing
                    if row_idx == 0:
                        logger.info("Processing first data row: %s", df_row.values[:3])

                    # Unit scaling policy for the financial table:
                    #   The extractor is called with multiply_values=False
                    #   (embed_financial_tables), so numeric values flow through
                    #   in the ORIGINAL source units declared by the workbook
                    #   header. If the header says CNY'000 / 人民币千元, values
                    #   already represent thousands and must NOT be multiplied.
                    #   Same for CNY'M / 人民币百万 (millions). The column
                    #   header shows the unit so the reader interprets them
                    #   correctly. Any accidental scaling here would double-count.
                    for col_idx, col_name in enumerate(df.columns[:max_cols]):
                        if col_idx >= len(table.columns):
                            break
                        cell = table.cell(data_row_idx, col_idx)

                        # Get value from DataFrame safely
                        value = df_row[col_name] if col_name in df_row.index else ""
                        text_val = _format_table_value(value, is_numeric_column=col_idx > 0)

                        # Description column only -- the source Financials
                        # sheet's own row labels (e.g. "Cash at bank and on
                        # hand") stay whatever language that sheet was
                        # authored in, even when the REPORT is Chinese, since
                        # nothing upstream of this table translates them.
                        if col_idx == 0 and is_chinese_mode:
                            text_val = _translate_statement_row_label(text_val, mappings)

                        # Set text
                        cell.text = text_val
                        
                        # Log first cell value of first row
                        if row_idx == 0 and col_idx < 2:
                            logger.info("Setting cell (%s, %s) to: '%s'", data_row_idx, col_idx, text_val)
                        
                        # Apply formatting: Arial 7pt (reduced from 9pt) for all cells
                        # Note: Always access paragraphs[0] AFTER setting text
                        if not cell.text_frame.paragraphs:
                            cell.text_frame.add_paragraph()
                            
                        p = cell.text_frame.paragraphs[0]
                        if not p.runs:
                            p.add_run()
                            
                        # cell.text = text_val above already wrote the text into
                        # the run; setting run.text again was a redundant XML
                        # roundtrip. Just apply the font formatting.
                        for run in p.runs:
                            run.font.name = 'Arial'
                            self._set_east_asian_typeface(run)
                            self._declare_run_language(run)
                            run.font.size = data_font_size
                            try:
                                run.font.color.rgb = BLACK
                            except Exception:
                                pass
                            run.font.bold = is_special_row
                        try:
                            p.line_spacing = 1.0
                            _apply_east_asian_line_breaking(p)
                        except Exception:
                            pass

                        # Give cells a small internal margin so text doesn't hug the border
                        try:
                            cell.margin_left = Inches(0.04)
                            cell.margin_right = Inches(0.04)
                            cell.margin_top = Inches(0.01)
                            cell.margin_bottom = Inches(0.01)
                        except Exception:
                            pass

                        # First column left-aligned, numeric columns right-aligned.
                        # Within the label column, a leaf line item (neither a
                        # section header nor a total/subtotal) is indented
                        # under whichever header it follows -- headers and
                        # every tier of total stay flush with the left margin
                        # (reference format, IMG_0035). Set via raw XML
                        # (_set_paragraph_left_indent) -- see that method's
                        # docstring for why paragraph.left_indent itself is a
                        # no-op in this python-pptx version.
                        try:
                            p.alignment = PP_ALIGN.LEFT if col_idx == 0 else PP_ALIGN.RIGHT
                            if col_idx == 0:
                                should_indent = not is_category_header_row and not is_total_row
                                _set_paragraph_left_indent(p, int(Inches(0.12)) if should_indent else 0)
                        except Exception:
                            pass

                        # No borders on a data cell at all. The deliverable
                        # draws none: not vertically, which would read as a
                        # spreadsheet, and not horizontally under every row,
                        # which is visually busy past 20 rows. Total and
                        # subtotal rows get their own explicit rules below,
                        # applied AFTER this, so the separators remain exactly
                        # where they carry meaning.

                        # Only the LAST column is highlighted light blue, on
                        # EVERY row including totals -- every other cell is
                        # EXPLICITLY set to solid white (not left unset --
                        # an untouched cell.fill inherits the table's built-in
                        # style GUID, which can render as a themed/tinted
                        # colour even with first_row/horz_banding disabled).
                        # The four statement-level grand totals (total
                        # assets/liabilities/equity, and liabilities+equity)
                        # are the one exception -- reference format (IMG_0035)
                        # calls those out with a solid grey fill across the
                        # whole row; every other total/subtotal tier stays
                        # white, called out by border weight only (below).
                        try:
                            cell.fill.solid()
                            if is_grand_total_row:
                                cell.fill.fore_color.rgb = GREY_TOTAL_FILL
                            else:
                                cell.fill.fore_color.rgb = (
                                    WHITE  # deliverable tints no column
                                )
                        except Exception:
                            pass

                        # Thin top border on every total/subtotal row; grand
                        # totals additionally get a heavier bottom border,
                        # matching the reference's two-tier total styling.
                        if is_total_row:
                            try:
                                _set_cell_border(cell, 'top', color_rgb="00338D", width=Pt(0.75))
                                if is_grand_total_row:
                                    _set_cell_border(cell, 'bottom', color_rgb="00338D", width=Pt(2.25))
                            except Exception:
                                pass
                    
                    logger.debug("Filled table row %s (data_row_idx: %s, special: %s)", row_idx + 1, data_row_idx, is_special_row)
                
                logger.info("Updated table with Excel data (formatting preserved)")
            else:
                # If no table, this is an error - table placeholder should be a table shape
                logger.error("Table Placeholder is not a table shape! Cannot embed financial table.")
                logger.error("Shape type: %s, has_table: %s", type(shape), hasattr(shape, 'table'))
                logger.error("Shape name: %s", shape.name if hasattr(shape, 'name') else 'unnamed')
                # Check if shape has table attribute but it's None
                if hasattr(shape, 'table'):
                    logger.error("shape.table is: %s", shape.table)
                # Try to create a table representation in text frame as last resort
                if shape.has_text_frame:
                    shape.text_frame.clear()
                    # Convert DataFrame to formatted text table - show ALL rows
                    try:
                        # Show all rows, no limit
                        text_table = df.to_string(index=False)
                    except:
                        text_table = str(df)
                    
                    p = shape.text_frame.paragraphs[0] if shape.text_frame.paragraphs else shape.text_frame.add_paragraph()
                    p.text = text_table
                    logger.warning("Added text table representation with all %s rows (%s chars) - NOT IDEAL, should be table format", len(df), len(text_table))
        except Exception as e:
            logger.error("Could not fill table placeholder: %s", e)
            logger.error(traceback.format_exc())
            # Fallback: add text representation - show ALL rows
            if shape.has_text_frame:
                shape.text_frame.clear()
                # Show all rows, not just first 10
                text_repr = df.to_string(index=False)
                p = shape.text_frame.paragraphs[0] if shape.text_frame.paragraphs else shape.text_frame.add_paragraph()
                p.text = text_repr


    def embed_financial_tables(
        self,
        excel_path: str,
        sheet_name: str,
        project_name: str,
        language: str,
        bs_is_results: Optional[Dict[str, Any]] = None,
        mappings: Optional[Dict[str, Any]] = None,
    ):
        """Embed financial tables: BS to page 1, IS to page 5"""
        try:
            import pandas as pd
            from fdd_utils.workbook import extract_balance_sheet_and_income_statement
            
            logger.info("Embedding financial tables from %s, sheet: %s", excel_path, sheet_name)

            # Only bail here if there's NEITHER a precomputed bs_is_results
            # NOR enough to extract one fresh -- excel_path/sheet_name being
            # blank is fine on its own when the caller already did the
            # extraction (e.g. from a roll-up workbook, or a synthesized
            # BS/IS with no literal Financials sheet at all).
            if bs_is_results is None and (not excel_path or not sheet_name):
                logger.warning("Missing excel_path (%s) or sheet_name (%s), and no precomputed BS/IS results -- skipping table embedding", excel_path, sheet_name)
                return
            
            # Use the precomputed results when available (they came from the
            # default extractor, multiply_values=True, so numeric values are
            # already in actual units). Only extract fresh if nothing was
            # passed in. Either way, at display time we rescale to the source
            # unit (CNY'000 / 人民币千元 / CNY'M / 人民币百万) so cells line up
            # with the header. This avoids the fragility of a second extract
            # that could silently return empty and lose the table.
            if bs_is_results is None:
                try:
                    logger.info("No precomputed BS/IS; extracting fresh")
                    bs_is_results = extract_balance_sheet_and_income_statement(
                        excel_path,
                        sheet_name,
                        debug=False,
                    )
                except Exception as exc:
                    logger.warning("Fresh BS/IS extraction failed: %s", exc)
                    return

            if not bs_is_results:
                logger.warning("No BS/IS data available for PPTX tables")
                return

            # Values as received have been multiplied by 1000 if the source
            # sheet declared CNY'000 / 人民币千元, and left as-is otherwise.
            # Rescale once we know the unit label (detected further below).
            values_pre_multiplied = True
            
            # Extract BS and IS DataFrames from results. Copy them — the rescale
            # block below divides numeric columns by 1000 IN PLACE, and
            # bs_is_results is the same object as session_state.bs_is_results,
            # which survives across re-exports (only cleared on a new file
            # upload). Without the copy, a second export in the same session
            # would divide already-divided values by 1000 again, showing
            # table figures 1000x too small.
            bs_df = bs_is_results.get('balance_sheet')
            is_df = bs_is_results.get('income_statement')
            bs_df = bs_df.copy() if bs_df is not None else None
            is_df = is_df.copy() if is_df is not None else None
            
            # Table titles follow the standard FDD phrasing regardless of what
            # the source Excel calls the sheet. Language-aware so Chinese decks
            # stay consistent with English decks.
            is_chinese_mode = str(language or "").strip().lower().startswith(("chi", "zh", "cn"))
            # The analyst deliverable's band reads just 资产负债表. It carries no
            # 示意性调整后 prefix (the stage is already stated in the column
            # headers) and no project suffix -- the suffix put the databook's
            # FILE NAME in the band, e.g. "...资产负债表 - Crescent-databook",
            # which is a working artefact, not a table title.
            if is_chinese_mode:
                bs_table_name = "资产负债表"
                is_table_name = "利润表"
            else:
                bs_table_name = "Balance sheet"
                is_table_name = "Income statement"

            # Detect currency unit from the sheet header. Currency markers live
            # in the first 20 rows (table titles / unit row); reading the full
            # sheet via iterrows() was a ~1-3s hit on big workbooks. Cap to
            # nrows=20 and use vectorised astype(str) instead of iterrows.
            currency_unit = None
            try:
                excel_df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, nrows=20)
                blob = ' '.join(
                    excel_df.fillna('').astype(str).agg(' '.join, axis=1).tolist()
                )
                if '人民币百万' in blob:
                    currency_unit = '人民币百万'
                elif "CNY'M" in blob or 'CNY million' in blob or 'CNY mn' in blob.lower():
                    currency_unit = "CNY'M"
                elif '人民币千元' in blob:
                    currency_unit = '人民币千元'
                elif "CNY'000" in blob or "CNY 000" in blob:
                    currency_unit = "CNY'000"
            except Exception:
                pass

            # An English-labelled source databook (e.g. Kunshan) will only
            # ever have "CNY'000"/"CNY'M" markers to detect, even when the
            # REPORT is being generated in Chinese -- normalise the unit
            # label itself to match the report language, same as the table
            # title above, so a Chinese table never shows an English header.
            if is_chinese_mode:
                if currency_unit == "CNY'000":
                    currency_unit = "人民币千元"
                elif currency_unit == "CNY'M":
                    currency_unit = "人民币百万"

            logger.info("Extracted BS: %s, IS: %s", bs_df.shape if bs_df is not None else 'None', is_df.shape if is_df is not None else 'None')
            logger.info("Table names - BS: %s, IS: %s, Currency: %s", bs_table_name, is_table_name, currency_unit)

            # If the values came from the precomputed (multiply_values=True)
            # pipeline, rescale to the source unit so the cells match the
            # header. "CNY'000" / "人民币千元" → divide by 1000,
            # The workbook extractor multiplies by 1000 ONLY when the source
            # header is CNY'000 / 人民币千元 (it does not touch millions).
            # Divide by 1000 here so the displayed cells read as thousands
            # (matching the "CNY'000" header). For any other unit the
            # values pass through unchanged.
            if values_pre_multiplied and currency_unit and (
                "千" in currency_unit or "'000" in currency_unit or "000" in currency_unit
            ):
                logger.info("Rescaling values by 1/1000 to match unit %s", currency_unit)
                for _df in (bs_df, is_df):
                    if _df is None or _df.empty:
                        continue
                    for _col in _df.columns:
                        if pd.api.types.is_numeric_dtype(_df[_col]):
                            _df[_col] = _df[_col] / 1000.0
            
            # Target the ACTUAL first commentary slide of each statement (a slide
            # object recorded during apply_structured_data_to_slides), not a
            # hard-coded slides[0]/slides[4]. Commentary adds slides and unused
            # ones are removed, so fixed indices drift — and the BS table could
            # land on a slide that no longer corresponds to BS page 1. The slide
            # OBJECT survives that reshuffle because it is a used (kept) slide.
            tracked = getattr(self, "_statement_table_slides", {}) or {}

            # Embed BS table on the first BS commentary slide.
            bs_slide = tracked.get("BS")
            if bs_slide is None and len(self.presentation.slides) > 0:
                bs_slide = self.presentation.slides[0]  # fallback
            if bs_df is not None and not bs_df.empty and bs_slide is not None:
                logger.info("Embedding BS table on tracked slide (shapes: %s)",
                            [getattr(s, 'name', '?') for s in bs_slide.shapes])
                self._embed_statement_table(
                    bs_slide, bs_df, "BS",
                    table_name=bs_table_name, currency_unit=currency_unit,
                    mappings=mappings, is_chinese_mode=is_chinese_mode,
                )
            else:
                logger.warning(
                    "Skipping BS table — bs_df empty=%s, target slide=%s. If bs_df is "
                    "empty but the databook DOES have a balance sheet, the session's "
                    "bs_is_results is stale: re-run Process Data, then export.",
                    bs_df is None or getattr(bs_df, 'empty', True), bs_slide is not None,
                )

            # Embed IS table on the first IS commentary slide.
            is_slide = tracked.get("IS")
            if is_slide is None and len(self.presentation.slides) > 4:
                is_slide = self.presentation.slides[4]  # fallback
            if is_df is not None and not is_df.empty and is_slide is not None:
                logger.info("Embedding IS table on tracked slide (shapes: %s)",
                            [getattr(s, 'name', '?') for s in is_slide.shapes])
                self._embed_statement_table(
                    is_slide, is_df, "IS",
                    table_name=is_table_name, currency_unit=currency_unit,
                    mappings=mappings, is_chinese_mode=is_chinese_mode,
                )
            elif is_df is not None and not is_df.empty:
                logger.error("No target slide found for IS table (slides=%s)", len(self.presentation.slides))
                    
        except Exception as e:
            logger.error("Error embedding financial tables: %s", e)
            logger.error(traceback.format_exc())


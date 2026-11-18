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


class _SummaryMixin:
    """The executive summary band (coSummaryShape): sizing it to the box, asking the
    model for text that fits, and checking what comes back.

    Mixed into PowerPointGenerator; `self` is the generator.
    """

    def _summary_settings(self) -> Dict[str, Any]:
        return dict(self.pptx_settings.get("executive_summary") or {})


    _SUMMARY_BAND_CLEARANCE_PT = 3.0

    @classmethod
    def _summary_band_spare_pt(cls, slide, shape) -> float:
        """Vertical space below the band that nothing else claims.

        The template draws coSummaryShape at whatever height its author picked
        and then leaves a gap before the table beneath it. Sizing the summary
        to the shape's own height therefore measures the drawing, not the
        space -- it can only ever fill the box someone drew, however much room
        sits under it. On the template this was written against that gap is
        17.7pt, a full 1.6 lines unused on every summary page, which is the
        band visibly not reaching the table.

        Returns the extra height available, measured to the top of the nearest
        shape that overlaps this one horizontally, less a small clearance so
        the last line does not touch the table. 0 when nothing is below.
        """
        try:
            top = float(shape.top)
            bottom = top + float(shape.height)
            left = float(shape.left)
            right = left + float(shape.width)
            nearest = None
            for other in slide.shapes:
                if other is shape or other.top is None:
                    continue
                o_top = float(other.top)
                if o_top < bottom:
                    continue
                o_left = float(other.left)
                if o_left >= right or (o_left + float(other.width)) <= left:
                    continue                      # different column, ignore
                nearest = o_top if nearest is None else min(nearest, o_top)
            if nearest is None:
                return 0.0
            spare_pt = (nearest - bottom) / 12700.0 - cls._SUMMARY_BAND_CLEARANCE_PT
            return max(0.0, spare_pt)
        except Exception:
            return 0.0

    def _grow_summary_band_into_spare(self, slide, shape) -> float:
        """Take the unused space under the band. Returns the pt gained.

        Only ever grows. The target character count is computed from the same
        measurement, so the box and the text it is asked for stay in step.
        """
        spare = self._summary_band_spare_pt(slide, shape)
        if spare <= 1.0:
            return 0.0
        try:
            shape.height = int(shape.height + spare * 12700)
            logger.info("Summary band grown by %.1fpt into unused space above the table", spare)
            return spare
        except Exception as exc:
            logger.debug("Could not grow the summary band: %s", exc)
            return 0.0

    def _summary_length_targets(self, is_chinese: bool) -> Tuple[int, int]:
        """(target_chars_chi, target_words_eng) sized to the REAL summary box
        rather than to a static config number.

        The static values had to be re-tuned by hand for every template and
        were consistently short -- a real deck filled about 2 of the 4 lines
        the box can hold. Measuring the actual `coSummaryShape` width and
        multiplying by `target_lines` self-corrects, and keeps the AI's own
        length instruction honest about the space it is writing into.

        Falls back to the configured values whenever the shape can't be
        resolved or measured."""
        settings = self._summary_settings()
        static_chi = int(settings.get("target_chars_chi", 144))
        static_eng = int(settings.get("target_words_eng", 110))
        target_lines = max(1, int(settings.get("target_lines", 4) or 4))
        try:
            shape = None
            spare_pt = 0.0
            for slide in self.presentation.slides:
                shape = find_shape_by_name(slide.shapes, "coSummaryShape")
                if shape is not None:
                    spare_pt = self._summary_band_spare_pt(slide, shape)
                    break
            if shape is None:
                return static_chi, static_eng
            from fdd_utils.text_metrics import get_measurer, text_box_from_shape
            box = text_box_from_shape(shape)
            packing = self._packing_settings()
            font_pt = _real_font_size_pt(is_chinese)
            measurer = get_measurer(
                _measurer_family(is_chinese, packing), font_pt, is_cjk=is_chinese,
                line_spacing=_real_line_spacing(is_chinese),
                metrics_path=_resolve_font_metrics_path(is_chinese, packing),
            )
            probe = ("财务尽职调查评述示例文字内容测算每行可容纳字数上限" * 12) if is_chinese else (
                "financial due diligence commentary sample text measuring how many words fit " * 6)
            wrapped = measurer.wrap(probe, box.width_pt)
            if not wrapped:
                return static_chi, static_eng
            per_line = max(1.0, len(probe) / len(wrapped))
            # How many lines the box HOLDS, measured -- not the configured
            # guess. The width here has always been measured while the line
            # count was a static 4, which is wrong in both directions on a
            # per-machine template: this Mac's band holds 2.99 lines (so 4
            # over-fills it), while the band that prompted "exe sum 非常短"
            # is evidently taller. target_lines survives only as the fallback
            # for when the shape can't be measured, which is exactly the role
            # the static width fallback already plays.
            from fdd_utils.text_metrics import POWERPOINT_LINE_PITCH_FACTOR
            pitch = font_pt * POWERPOINT_LINE_PITCH_FACTOR * _real_line_spacing(is_chinese)
            # Measure the SPACE, not the drawing: the band's own height plus
            # whatever sits unclaimed between it and the table below. Using
            # box.height_pt alone is circular -- it can only ever ask for
            # enough text to fill the box someone happened to draw.
            usable_pt = box.height_pt + spare_pt
            fitting_lines = (usable_pt / pitch) if pitch > 0 else 0
            lines = fitting_lines if fitting_lines >= 1 else target_lines
            # 0.92: aim just under a full box so a slightly long sentence
            # doesn't push a whole extra line past the table beneath it.
            chars = int(per_line * lines * 0.92)
            logger.info(
                "Executive summary target: %d chars (band %.0fx%.0fpt + %.1fpt spare "
                "below = %.2f lines at %.1fpt, ~%.0f chars/line)",
                chars, box.width_pt, box.height_pt, spare_pt, fitting_lines, font_pt, per_line,
            )
            # No max(static, ...) floor once BOTH dimensions are measured.
            # The floor existed because the line count was a guess; raising a
            # measured 127-char band to a static 144 just overflows it, which
            # is the same mistake as shrinking a box to absorb 0.3 lines. If
            # the resulting summary is too short for the reader's taste, the
            # band itself is too small -- the log line above says by how much,
            # so that is now a decision that can be made on a real number.
            if fitting_lines >= 1:
                return (chars, static_eng) if is_chinese else (static_chi, int(chars / 6))
            if is_chinese:
                return max(static_chi, chars), static_eng
            # English targets are in WORDS; ~6 chars per word including space.
            return static_chi, max(static_eng, int(chars / 6))
        except Exception as exc:
            logger.debug("Could not size summary targets from the real shape: %s", exc)
            return static_chi, static_eng


    @classmethod
    def strip_table_detail_for_summary(cls, text: str, is_chinese: bool) -> str:
        """Truncates one account's commentary to its lead-in only, at the
        same "明细如下："/"the breakdown is set out below" handoff phrase
        _split_table_commentary already splits on for PPTX rendering --
        callers building a page/section-level summary blob (ui.py, both the
        single-file and batch paths) should use ONLY this lead-in, never
        the per-component "-"/"➢" bullets after it.

        Root cause this exists to prevent: the non-AI fallback summary
        (_generate_page_summary) picks each account's "first sentence" via
        a naive splitter that doesn't know "明细如下：" isn't a real sentence
        end -- for a table account's RAW (unsplit) commentary, that made
        its "first sentence" run straight through the handoff phrase and
        into the whole first "➢" bullet, producing a summary that visibly
        spliced in the table's own detail bullets and cut off mid-sentence
        (confirmed against a real reference case export's coSummaryShape). Safe
        to call on every account unconditionally, table-bearing or not --
        _split_table_commentary already falls back to the whole text
        unchanged when the handoff phrase isn't present.
        """
        generator = cls.__new__(cls)
        lead_in, _post = generator._split_table_commentary(text, is_chinese)
        return cls._drop_itemised_breakdown(lead_in, is_chinese)

    @staticmethod
    def _drop_itemised_breakdown(text: str, is_chinese: bool) -> str:
        """Cut an account's numbered composition out of the summary source.

        The per-account bullets now enumerate composition, which is what the
        deliverable wants THERE. It is not what a page summary wants: the band
        is a few lines covering seven accounts, so an enumeration inherited
        wholesale reads as one item per account and nothing else. A real export
        opened with "货币资金余额为1.23亿元，主要包括：1）某银行某支行
        #6028#-美元资本金户8,404.9万元；截至..." -- a bank account number in the
        executive summary, and the other components never mentioned.

        Keeps the sentence up to the enumeration handoff, so the headline
        balance survives and only the itemisation goes.
        """
        body = str(text or "")
        if not body:
            return body
        # "...合计674.5万元，主要包括：1）..." -> "...合计674.5万元。"
        handoff = re.search(
            r"[，,]?\s*(?:主要)?(?:包括|包含|comprising|consisting of)\s*[:：]\s*(?=[1-9１]\s*[）)])",
            body)
        if not handoff or not body[:handoff.start()].strip():
            return body
        head = body[:handoff.start()].rstrip("，,、 ")
        tail = body[handoff.end():]
        # End of the enumeration = the first real sentence end after it. Match
        # 。！？ directly, and an ASCII full stop only when it is NOT a decimal
        # point -- splitting on every "." turned "8,404.9万元" into a sentence
        # boundary and left "9万元；2）..." stranded in the summary.
        end = re.search(r"[。！？]|(?<!\d)\.(?!\d)", tail)
        # Anything after the enumeration is kept: the ageing, impairment and
        # procedure observations live there and are what a summary most wants.
        rest = tail[end.end():].lstrip() if end else ""
        joiner = "。" if is_chinese else ". "
        return (head + joiner + rest).strip()

    def _build_page_summary_source(self, slide_accounts: List[Dict]) -> Tuple[str, str]:
        """Build the exact slide commentary set used for summary generation.

        The summary source is each account's LEAD-IN only, never its table
        detail. strip_table_detail_for_summary was written for exactly this
        and its docstring records the symptom it prevents -- the naive
        first-sentence splitter in _generate_page_summary doesn't know
        "明细如下：" isn't a sentence end, so a table account's raw commentary
        yields a "first sentence" that runs through the handoff phrase and
        swallows the whole first "➢" bullet, splicing table detail into the
        executive summary and cutting off mid-sentence. Both ui.py call
        sites already strip; the export path never did, so the deterministic
        in-export summary reintroduced the same defect the moment it started
        filling the band rather than leaving it blank.

        page_commentary stays UNSTRIPPED -- it is the AI path's own input and
        should still see everything.
        """
        commentary_parts = []
        summary_source_parts = []
        summary_parts = []

        for account_data in slide_accounts or []:
            commentary = str(account_data.get("commentary", "") or "").strip()
            summary = str(account_data.get("summary", "") or "").strip()
            if commentary:
                commentary_parts.append(commentary)
                lead_in = self.strip_table_detail_for_summary(
                    commentary, _account_is_chinese(account_data),
                )
                summary_source_parts.append(lead_in or commentary)
            if summary:
                summary_parts.append(summary)

        page_commentary = "\n\n".join(commentary_parts).strip()
        page_summary_source = (
            "\n\n".join(summary_source_parts).strip() or " ".join(summary_parts).strip()
        )
        return page_commentary, page_summary_source


    def _generate_page_summary(self, commentary: str, is_chinese: bool) -> str:
        """Fallback (non-AI) page summary.

        Sentences are taken in RANK ORDER across account paragraphs -- every
        account's opening sentence first, then every account's second, and so
        on -- until the measured character budget for the box is used up. The
        rank ordering keeps the original intent (the summary spans the whole
        page rather than covering only the first account); filling to the
        budget is what stops it under-running the box.

        The old version took exactly one sentence per account and stopped at
        a static max_sentences (4). Both limits starved the band on a real
        deck: a 7-account page showed 4 sentences / 168 chars against a ~276
        char budget, and a single-account page showed one sentence / 31
        chars -- reported as "exe sum 非常短". max_chars is derived from the
        real coSummaryShape width x target_lines, so it already expresses
        "as much as the box holds" far better than a hand-tuned sentence
        count, and now governs on its own.
        """
        if not commentary or not commentary.strip():
            return ""
        is_chinese_text = is_chinese or detect_chinese_text(commentary)
        _t_chi, _t_eng = self._summary_length_targets(is_chinese_text)
        max_chars = _t_chi if is_chinese_text else _t_eng * 6

        # Each account block is separated by "\n\n".
        blocks = [b.strip() for b in commentary.split("\n\n") if b.strip()]
        by_block = [_split_text_sentences(b, is_chinese_text) for b in blocks]
        by_block = [s for s in by_block if s]
        if not by_block:
            by_block = [_split_text_sentences(commentary, is_chinese_text)]
            by_block = [s for s in by_block if s]
            if not by_block:
                return ""

        sep = "" if is_chinese_text else " "
        picked: List[str] = []
        used = 0
        for rank in range(max(len(s) for s in by_block)):
            for sentences in by_block:
                if rank >= len(sentences):
                    continue
                candidate = sentences[rank].strip()
                if not candidate:
                    continue
                cost = len(candidate) + (len(sep) if picked else 0)
                if picked and used + cost > max_chars:
                    continue
                picked.append(candidate)
                used += cost

        if not picked:
            picked = [by_block[0][0]]

        summary = sep.join(picked).strip()
        if len(summary) > max_chars:
            summary = summary[:max_chars].rstrip(" ,;:/-") + "…"
        return summary.strip()


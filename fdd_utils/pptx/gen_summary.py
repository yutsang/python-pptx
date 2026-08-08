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
            for slide in self.presentation.slides:
                shape = find_shape_by_name(slide.shapes, "coSummaryShape")
                if shape is not None:
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
            fitting_lines = (box.height_pt / pitch) if pitch > 0 else 0
            lines = fitting_lines if fitting_lines >= 1 else target_lines
            # 0.92: aim just under a full box so a slightly long sentence
            # doesn't push a whole extra line past the table beneath it.
            chars = int(per_line * lines * 0.92)
            logger.info(
                "Executive summary target: %d chars (band %.0fx%.0fpt holds %.2f lines "
                "at %.1fpt, ~%.0f chars/line)",
                chars, box.width_pt, box.height_pt, fitting_lines, font_pt, per_line,
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


    def _resolve_summary_model_type(self, is_chinese: bool) -> str:
        cached = getattr(self, "_summary_model_type_cache", None)
        if cached:
            return cached

        if self.model_type:
            resolved_model_type = str(self.model_type).strip()
            self._summary_model_type_cache = resolved_model_type
            return resolved_model_type

        try:
            from fdd_utils.ai import FDDConfig

            config = load_yaml_file(package_file_path("config.yml")) or {}
            requested_model_type = (
                str(self.model_type).strip()
                if self.model_type
                else str(((config.get("default") or {}).get("ai_provider")) or "deepseek")
            )
            config_manager = FDDConfig(
                language="Chi" if is_chinese else "Eng",
                model_type=requested_model_type,
            )
            resolved_model_type = str(config_manager.model_type or requested_model_type)
        except Exception as exc:
            logger.warning("Could not resolve PPTX summary model type, defaulting to deepseek: %s", exc)
            resolved_model_type = "deepseek"

        self._summary_model_type_cache = resolved_model_type
        return resolved_model_type


    def _call_with_timeout(self, func, timeout_seconds: float, timeout_label: str):
        if timeout_seconds <= 0:
            return func()

        result_container = {"value": None, "error": None, "completed": False}

        def _run():
            try:
                result_container["value"] = func()
            except Exception as exc:
                result_container["error"] = exc
            finally:
                result_container["completed"] = True

        worker = threading.Thread(target=_run, daemon=True)
        worker.start()
        worker.join(timeout=timeout_seconds)

        if not result_container["completed"]:
            raise TimeoutError(f"{timeout_label} timed out after {timeout_seconds:.1f} seconds")
        if result_container["error"] is not None:
            raise result_container["error"]
        return result_container["value"]


    def _call_with_timeout_retry(
        self,
        func,
        timeout_seconds: float,
        max_retries: int,
        timeout_label: str,
    ):
        """Call ``func`` with a per-attempt timeout and retry on TimeoutError
        or other transient failures. Raises the last exception if all retries
        fail. Use ``max_retries >= 1`` (1 means "no retry, just run once")."""
        attempts = max(1, int(max_retries or 1))
        last_error: Optional[BaseException] = None
        for attempt in range(1, attempts + 1):
            label = (
                timeout_label
                if attempts == 1
                else f"{timeout_label} (attempt {attempt}/{attempts})"
            )
            try:
                return self._call_with_timeout(func, timeout_seconds, label)
            except TimeoutError as te:
                last_error = te
                logger.warning(
                    "%s timed out after %.1fs; %s",
                    label,
                    timeout_seconds,
                    "retrying" if attempt < attempts else "giving up",
                )
            except Exception as exc:
                last_error = exc
                logger.warning(
                    "%s errored (%s); %s",
                    label,
                    exc,
                    "retrying" if attempt < attempts else "giving up",
                )
        assert last_error is not None
        raise last_error


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


    def _validate_ai_summary(
        self,
        source_text: str,
        draft_summary: str,
        is_chinese: bool,
        ai_helper: Optional[Any] = None,
    ) -> Optional[str]:
        summary_settings = self._summary_settings()
        if not bool(summary_settings.get("enable_validation", True)):
            return draft_summary

        try:
            from fdd_utils.ai import AIClient

            model_type = ai_helper.model_type if ai_helper is not None else self._resolve_summary_model_type(is_chinese)
            if model_type == "local" and not bool(summary_settings.get("local_enable_validation", False)):
                logger.info("Skipping PPTX summary validation for local model; using draft summary directly")
                return draft_summary

            max_input_chars = int(summary_settings.get("max_input_chars", 1400))
            validation_max_tokens = int(summary_settings.get("validation_max_tokens", 90))
            max_numeric_sentences = int(summary_settings.get("max_numeric_sentences", 1))
            validation_timeout_seconds = float(summary_settings.get("validation_timeout_seconds", 25) or 25)
            # Sized to the real summary box, not a static config number --
            # see _summary_length_targets.
            target_chars_chi, target_words_eng = self._summary_length_targets(is_chinese)
            max_sentences_chi = int(summary_settings.get("max_sentences_chi", 4))
            max_sentences_eng = int(summary_settings.get("max_sentences_eng", 4))

            if is_chinese:
                prompt = f"""请校验以下PPT执行摘要草稿，使其适合作为财务PPT摘要框内容。

要求：
1. 只保留与原始评论一致的高层结论、趋势和核心驱动。
2. 控制在{max_sentences_chi}句话以内，长度约{target_chars_chi}字 —— 若草稿明显短于目标长度，请补充其他高层要点以达到目标，不要仅做压缩。
3. 最多保留{max_numeric_sentences}个数字或百分比，除非删除后会影响结论准确性。
4. 删除重复、堆叠金额和逐项罗列，但不得为了简短而牺牲目标长度或丢失核心趋势、驱动和结论。
5. 优先合并重复句、删去铺垫和次要背景，只保留最重要的业务含义。
6. 只输出最终摘要，不要解释。

原始评论：
{source_text[:max_input_chars]}

摘要草稿：
{draft_summary}"""
            else:
                prompt = f"""Validate the draft executive summary for a financial PPT summary box.

Requirements:
1. Keep only source-supported themes, trend, and core driver.
2. Limit the result to no more than {max_sentences_eng} sentences and about {target_words_eng} words —
   if the draft runs noticeably shorter than the target, add other high-level points to reach it
   rather than just compressing further.
3. Keep at most {max_numeric_sentences} number or percentage unless removing it would make the summary inaccurate.
4. Remove repeated phrasing, stacked figures, scene-setting language, and account-by-account detail,
   but do not sacrifice the target length or the key trend, driver, or conclusion just to be terse.
5. Output only the final validated summary paragraph.

Source commentary:
{source_text[:max_input_chars]}

Draft summary:
{draft_summary}"""

            ai_helper = ai_helper or AIClient(
                model_type=model_type,
                language='Chi' if is_chinese else 'Eng',
                model_name=self.model_name,
            )
            validation_max_retries = int(summary_settings.get("validation_max_retries", 2) or 2)
            response = self._call_with_timeout_retry(
                lambda: ai_helper.get_response(
                    user_prompt=prompt,
                    system_prompt=(
                        "You validate executive summaries for financial presentation slides. "
                        "Keep only source-supported, concise, presentation-ready conclusions."
                    ),
                    temperature=float(summary_settings.get("validation_temperature", 0.1) or 0.1),
                    max_tokens=validation_max_tokens,
                ),
                timeout_seconds=validation_timeout_seconds,
                max_retries=validation_max_retries,
                timeout_label="PPTX summary validation",
            )
            validated_summary = str((response or {}).get("content") or "").strip()
            if _looks_like_blocked_ai_content(validated_summary):
                logger.warning("PPTX summary validation returned blocked/network HTML content; using draft summary fallback")
                return draft_summary
            return validated_summary or draft_summary
        except Exception as exc:
            logger.warning("Could not validate AI summary: %s", exc)
            return draft_summary


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
        (confirmed against a real Crescent export's coSummaryShape). Safe
        to call on every account unconditionally, table-bearing or not --
        _split_table_commentary already falls back to the whole text
        unchanged when the handoff phrase isn't present.
        """
        generator = cls.__new__(cls)
        lead_in, _post = generator._split_table_commentary(text, is_chinese)
        return lead_in


    @classmethod
    def generate_section_summary(
        cls,
        commentary: str,
        *,
        is_chinese: bool,
        language: str = "english",
        model_type: Optional[str] = None,
        model_name: Optional[str] = None,
    ) -> Optional[str]:
        """Top-level helper: generate one executive summary from concatenated
        commentary for a BS or IS section. Designed to be called from the UI
        during the AI commentary phase so the PPTX export becomes pure XML
        (no AI calls during export).

        Returns the summary string, or None if AI is disabled / fails.
        """
        try:
            generator = cls.__new__(cls)
            # Use the same config-merged settings the full generator uses,
            # otherwise the timeout/retry defaults from config.yml are
            # ignored (would fire at 10s instead of the configured value).
            generator.pptx_settings = _load_pptx_settings()
            generator.model_type = model_type
            generator.model_name = model_name
            generator.language = language
            result = generator._generate_ai_summary(commentary, commentary, is_chinese)
            if result is None:
                # AI timed out or unavailable — fall back to the rule-based
                # summary so coSummaryShape is never left blank.
                result = generator._generate_page_summary(commentary, is_chinese) or None
            return result
        except Exception as exc:
            logger.warning("generate_section_summary failed: %s", exc)
            return None


    def _generate_ai_summary(self, commentary: str, summary_source: str, is_chinese: bool) -> Optional[str]:
        """Generate and validate AI summary from page commentary."""
        try:
            from fdd_utils.ai import AIClient
            summary_settings = self._summary_settings()
            if not bool(summary_settings.get("enable_ai", True)):
                logger.info("PPTX summary AI disabled by config; using fallback summary")
                return None
            model_type = self._resolve_summary_model_type(is_chinese)
            max_input_chars = int(summary_settings.get("max_input_chars", 1600))
            max_tokens = int(summary_settings.get("max_tokens", 180))
            max_numeric_sentences = int(summary_settings.get("max_numeric_sentences", 1))
            # Use a shorter timeout for local models — they either answer fast
            # or they're not running; long waits just block the export.
            _is_local = str(model_type or "").lower() == "local"
            generation_timeout_seconds = float(
                summary_settings.get("local_generation_timeout_seconds", 10)
                if _is_local else
                summary_settings.get("generation_timeout_seconds", 20)
            )
            # Sized to the real summary box, not a static config number --
            # see _summary_length_targets.
            target_chars_chi, target_words_eng = self._summary_length_targets(is_chinese)
            max_sentences_chi = int(summary_settings.get("max_sentences_chi", 4))
            max_sentences_eng = int(summary_settings.get("max_sentences_eng", 4))
            source_text = str(commentary or summary_source or "").strip()
            if not source_text:
                return None

            if is_chinese:
                prompt = f"""请将以下财务评论改写成适合PPT摘要框的高层执行摘要。

目标长度：约{target_chars_chi}字，控制在{max_sentences_chi}句话以内 —— 请写满这个长度，不要明显短于目标。

要求：
1. 保留高层结论、趋势和核心驱动，可覆盖一个以上要点以达到目标长度。
2. 除非极其必要，最多保留{max_numeric_sentences}个数字或百分比。
3. 不要逐项复述账户，不要堆叠金额细节。
4. 语气要像管理层摘要，写成一个紧凑自然的短段落。
5. 优先删去次要说明、重复铺垫和账户层级细节，只保留最重要的业务结论、驱动和影响，但不要为了简短而牺牲目标长度。

原始内容：
{source_text[:max_input_chars]}"""
            else:
                prompt = f"""Write a short executive summary for a PPT summary box based on the following financial commentary.

Target length: about {target_words_eng} words, with no more than {max_sentences_eng} sentences —
write to fill this length, do not stop noticeably short of it.

Requirements:
1. Cover overall trend, key driver, and business implication — span more than one theme if needed to reach the target length.
2. Keep it high level and presentation-friendly.
3. Include at most {max_numeric_sentences} number or percentage unless absolutely necessary.
4. Do not list account-by-account detail or repeat many figures.
5. Write one compact management-style paragraph. Remove secondary detail, scene-setting language, and repeated wording, but do not sacrifice the target length just to be terse.
6. Remove secondary detail, scene-setting language, and repeated wording.

Original content:
{source_text[:max_input_chars]}"""

            ai_helper = AIClient(
                model_type=model_type,
                language='Chi' if is_chinese else 'Eng',
                model_name=self.model_name,
            )
            generation_max_retries = int(
                summary_settings.get("local_generation_max_retries", 1)
                if _is_local else
                summary_settings.get("generation_max_retries", 1)
            )
            response = self._call_with_timeout_retry(
                lambda: ai_helper.get_response(
                    user_prompt=prompt,
                    system_prompt=(
                        "You write concise executive summaries for financial presentation slides. "
                        "Prefer themes, drivers, and implications over detailed figures."
                    ),
                    temperature=float(summary_settings.get("generation_temperature", 0.2) or 0.2),
                    max_tokens=max_tokens,
                ),
                timeout_seconds=generation_timeout_seconds,
                max_retries=generation_max_retries,
                timeout_label="PPTX summary generation",
            )
            from fdd_utils.ai import strip_thinking
            summary = strip_thinking(str((response or {}).get("content") or "")).strip()
            if _looks_like_blocked_ai_content(summary):
                logger.warning(
                    "PPTX summary generation returned blocked/network HTML content; falling back to compact summary"
                )
                return None

            if summary:
                return self._validate_ai_summary(source_text, summary, is_chinese, ai_helper=ai_helper)
        except Exception as e:
            logger.warning("Could not generate AI summary: %s", e)
            logger.debug(traceback.format_exc())
        
        return None


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


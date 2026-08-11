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
    lead_promises_table,
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




class _PackingMixin:
    """Deciding which account's commentary lands in which slot: the first pass, the
    DP optimisation, and the rebalance passes that run after it.
    
    Several comments in here record passes that were tried and reverted. They are
    load-bearing -- read them before changing a tolerance.

    Mixed into PowerPointGenerator; `self` is the generator.
    """

    def _packing_settings(self, statement_type: Optional[str] = None) -> Dict[str, Any]:
        packing = dict(self.pptx_settings.get("commentary_packing") or {})
        if not statement_type:
            return packing
        overrides = ((packing.get("statement_overrides") or {}).get(statement_type) or {})
        if not overrides:
            return packing
        return _merge_nested_dict(packing, overrides)


    def _tail_overflow_tolerance_units(self, statement_type: Optional[str] = None) -> float:
        raw = self._packing_settings(statement_type).get(
            "tail_overflow_tolerance_lines", self._TAIL_OVERFLOW_TOLERANCE_UNITS
        )
        try:
            return max(0.0, float(raw))
        except (TypeError, ValueError):
            return self._TAIL_OVERFLOW_TOLERANCE_UNITS


    def _slot_names_for_actual_slide(self, actual_slide_idx: int, start_slide: int) -> List[str]:
        """Which slot names this slide's OWN template shapes actually
        support. Mirrors _distribute_content_across_slots's own identically
        -named local closure (slot_names_for_actual_slide) exactly,
        including its "first slide of statement -> single" fallback rule --
        this MUST agree with that closure's notion of what a slide offers,
        or a table account could be "placed" onto a slot name the packer's
        own accounting disagrees with, silently overlapping real content.
        `actual_slide_idx` is an absolute index into self.presentation.
        slides, NOT the slide_idx slot_distribution tuples carry (those are
        0-based from start_slide -- see _append_table_accounts_to_distribution).
        """
        if self.presentation is not None and 0 <= actual_slide_idx < len(self.presentation.slides):
            slide = self.presentation.slides[actual_slide_idx]
            has_left = find_shape_by_name(slide.shapes, "textMainBullets_L") is not None
            has_right = find_shape_by_name(slide.shapes, "textMainBullets_R") is not None
            has_single = find_shape_by_name(slide.shapes, "textMainBullets") is not None
            if has_left and has_right:
                return ["L", "R"]
            if has_single:
                return ["single"]
        return ["single"] if actual_slide_idx == start_slide - 1 else ["L", "R"]


    def _append_table_accounts_to_distribution(
        self, table_items: List[Dict[str, Any]],
        slot_distribution: List[Tuple[int, str, List[Dict[str, Any]]]],
        *, max_slides: int, start_slide: int, is_chinese_databook: bool = False,
        trailing_items: Optional[List[Dict[str, Any]]] = None,
    ) -> List[Tuple[int, str, List[Dict[str, Any]]]]:
        """Assigns each table-bearing account to a (slide_idx, slot_name),
        preferring to join an EARLIER table account's own slot when there's
        room (see _TABLE_SLOT_CAPACITY_PT/_TABLE_SLOT_PACK_THRESHOLD) rather
        than always claiming a fresh one -- a small table (e.g. 税金及附加,
        3 components) claiming a WHOLE column regardless of the normal
        packer's own output left most of that column empty; multiple small
        table accounts can share a column exactly the way multiple ordinary
        accounts already share one, just via a running EMU offset instead
        of shared-textframe paragraphs (see _render_table_accounts_stack).

        Never shares a slot with a NON-table account from the normal
        packer's own pool, and never touches that pool's already-finished
        distribution -- both of those are what caused the original
        two-tables-on-top-of-each-other bug this design replaced.

        trailing_items (optional) are ordinary accounts that come AFTER the
        LAST table account in the statement's own reading order -- e.g. 投资
        收益/营业外支出 following 财务费用 in a real Crescent IS deck. The
        caller (apply_structured_data_to_slides) excludes these from the
        normal packer's input pool entirely before it ever runs, precisely
        so this function can try them here instead: each is placed AFTER all
        table_items have claimed their slots, using ONLY its own lead-in
        height (no table/source/explanation arms -- it's plain commentary),
        against the exact same slot_fill_pt pool table_items just filled --
        which is exactly the empty space below a table that prompted this
        whole feature. A trailing item that fits into no existing slot
        claims a fresh one via the identical fallback table_items already
        use below, so it always renders somewhere; this function still never
        reopens or second-guesses the normal pool's own finished output.

        slide_idx here is 0-based FROM start_slide, matching exactly what
        _distribute_content_across_slots itself returns (its caller,
        apply_structured_data_to_slides, converts to an absolute slide via
        `start_slide - 1 + slide_idx` when it actually renders) -- using an
        absolute index here instead double-applies that offset downstream.
        """
        used = {(slide_idx, slot_name) for slide_idx, slot_name, _ in slot_distribution}
        result = list(slot_distribution)
        slot_fill_pt: Dict[Tuple[int, str], float] = {}
        def _cap_for(key: Tuple[int, str]) -> float:
            # Per-slot real capacity, not one shared constant -- a "single"
            # first-page slot and an L/R continuation slot differ by ~8%.
            return (self._slot_capacity_pt(key[0], key[1], start_slide)
                    * self._TABLE_SLOT_PACK_THRESHOLD)


        def _append_to_slot(key: Tuple[int, str], item: Dict[str, Any], block_pt: float) -> None:
            slide_idx, slot_name = key
            for i, (s, n, accounts) in enumerate(result):
                if s == slide_idx and n == slot_name:
                    result[i] = (s, n, accounts + [item])
                    break
            slot_fill_pt[key] = slot_fill_pt.get(key, 0.0) + block_pt
            if item.get("_presentation_table"):
                slots_with_table.add(key)

        def _open_new_slot(item: Dict[str, Any], block_pt: float) -> Optional[Tuple[int, str]]:
            for slide_idx in range(max_slides):
                actual_slide_idx = start_slide - 1 + slide_idx
                for slot_name in self._slot_names_for_actual_slide(actual_slide_idx, start_slide):
                    if (slide_idx, slot_name) in used:
                        continue
                    used.add((slide_idx, slot_name))
                    result.append((slide_idx, slot_name, [item]))
                    slot_fill_pt[(slide_idx, slot_name)] = block_pt
                    if item.get("_presentation_table"):
                        slots_with_table.add((slide_idx, slot_name))
                    return (slide_idx, slot_name)
            logger.warning(
                "No free slot for account %s within %s slides -- not rendered.",
                item.get("mapping_key"), max_slides,
            )
            return None

        # Content flows in READING ORDER through one column at a time, the
        # way a document does -- fill the current column, then move to the
        # next. This replaces a first-fit search over every open slot,
        # which could place a LATER account into an EARLIER column that
        # still had room: a real export had 投资收益 (last in the income
        # statement) rendered on slide 4 between 税金及附加 and 管理费用,
        # because it happened to fit there. Sequential flow makes that
        # impossible by construction rather than by an extra ordering
        # check, and is also what makes the lead/table split below safe --
        # nothing can ever land between a lead-in and its own table.
        cursor: Optional[Tuple[int, str]] = None

        # Start the flow in the LAST slot the normal packer used, rather than
        # on a brand-new one. That slot is very often barely filled -- a real
        # deck left the IS statement page's commentary column at 33% while
        # the next page's right column sat COMPLETELY EMPTY, roughly two
        # blank columns between them.
        #
        # This slot used to be off-limits: every packer slot went into `used`
        # so a table could never share one, because two renderers writing the
        # same slot once produced overlapping tables. That objection is gone
        # -- since the shared-frame rewrite, _render_table_accounts_stack
        # writes its lead-ins as PARAGRAPHS in the slot's own text frame and
        # reserves table space inside it, and the render dispatch already
        # routes a slot holding any table account through that one path. A
        # mixed slot therefore has a single writer.
        #
        # Only the LAST slot is eligible, which is what keeps reading order:
        # the table accounts come after the plain ones in the statement, so
        # appending to the final packer slot continues the document; appending
        # to an earlier one would interleave them.
        if slot_distribution:
            for slide_idx, slot_name, accounts in reversed(slot_distribution):
                if not accounts:
                    continue
                key = (slide_idx, slot_name)
                # Same currency flow() measures in, so the fit test below is
                # comparing like with like.
                slot_fill_pt[key] = sum(
                    self._estimate_lead_in_pt(a, is_chinese_databook) for a in accounts
                )
                cursor = key
                logger.debug(
                    "Table flow continues in the packer's last slot %s (%.0f of %.0fpt used)",
                    key, slot_fill_pt[key], _cap_for(key),
                )
                break

        # The same tail tolerance the ordinary packer honours. This path did
        # not, so a trailing account that fitted the current column was still
        # pushed into a fresh one: a real deck left 所得税费用 alone in a 13%
        # column while the column before it had room -- the fill diagnostic
        # scored that exact move as a GENUINE GAP ("costs 2.57, remaining
        # 2.61"). The project team accepts 1-2 lines protruding; they do not
        # accept a near-empty column beside a full one.
        # statement_type is not threaded into this function, so this reads the
        # top-level commentary_packing value rather than a per-statement
        # override -- the setting the user actually edits.
        _tail_tol_pt = (
            self._tail_overflow_tolerance_units()
            * _planning_std_lh_pt(is_chinese_databook)
        )

        def _fits(key: Optional[Tuple[int, str]], block_pt: float) -> bool:
            if key is None:
                return False
            return slot_fill_pt.get(key, 0.0) + block_pt <= _cap_for(key) + _tail_tol_pt

        def _remaining(key: Optional[Tuple[int, str]]) -> float:
            """Space left in a slot, for deciding whether a partial block is
            worth leaving behind rather than merely whether it fits."""
            if key is None:
                return 0.0
            return max(0.0, _cap_for(key) + _tail_tol_pt - slot_fill_pt.get(key, 0.0))

        # A column draws ONE source caption no matter how many subtables it
        # holds (see _render_table_accounts_stack, which gives it to the last
        # one only). The per-account estimate charges every table for its own,
        # so the second and later tables in a column have to be refunded it --
        # otherwise the planner believes a column is fuller than it renders,
        # by exactly this much per extra table. Measured on a real export: it
        # over-charged a column by 14.0pt and rejected a block that fitted with
        # 14pt to spare, missing the boundary by 0.1pt.
        _source_cost_pt = self._TABLE_SOURCE_LINE_PT + 2.0
        slots_with_table: set = set()

        def _refunded(key: Optional[Tuple[int, str]], block_pt: float,
                      parts_pt: Optional[Tuple[float, float, float]]):
            """block/parts as they will actually RENDER into `key`."""
            if parts_pt is None or key is None or key not in slots_with_table:
                return block_pt, parts_pt
            lead, table, expl = parts_pt
            return block_pt - _source_cost_pt, (lead, table - _source_cost_pt, expl)

        def flow(item: Dict[str, Any], block_pt: float,
                 parts_pt: Optional[Tuple[float, float, float]] = None) -> None:
            """Places one account at the flow cursor, splitting it at a
            natural boundary when the whole block won't fit in the current
            column but a leading portion of it still will.

            ONE split point: lead-in + table stay here, and only the
            trailing explanation continues in the next column under a
            "（续）" heading. That is the real reference deck's own
            convention for its 营业成本 bullets.

            A lead-in never separates from its own table -- see the removed
            branch below for why. The table itself is never divided (no
            repeated header) per explicit user instruction. parts_pt=None
            marks an account with nothing to split (a plain trailing
            account), which flows whole."""
            nonlocal cursor
            # What this block costs IN THIS COLUMN -- a second table here does
            # not pay for a second source caption (see _refunded).
            block_pt, parts_pt = _refunded(cursor, block_pt, parts_pt)
            if _fits(cursor, block_pt):
                _append_to_slot(cursor, item, block_pt)
                return

            if parts_pt is not None and self._TABLE_ALLOW_LEAD_TABLE_SPLIT:
                lead_pt, table_pt, explain_pt = parts_pt
                heading = self._TABLE_CONTINUATION_HEADING_PT
                # 1. Keep the lead-in AND table here; only the explanation
                #    moves. Worth it only if there IS an explanation to move.
                if explain_pt > 0 and _fits(cursor, lead_pt + table_pt):
                    _append_to_slot(cursor, dict(item, _render_parts=("lead", "table")),
                                    lead_pt + table_pt)
                    cursor = _open_new_slot(dict(item, _render_parts=("expl",)),
                                            explain_pt + heading)
                    return
                # 2. Keep only the lead-in here; table + explanation move.
                #
                #    READ THE HISTORY BEFORE WIDENING THIS. It was built once
                #    to the letter of "表格前如果這個point有文字 而這一邊放不下
                #    表格 那表格可以放下一個", fired on EVERY table account, and
                #    was removed as the worst of the three outcomes:
                #      - the column ended on "...明细如下：" with no 明细 under
                #        it, which simply reads as broken;
                #      - the table then needed a "（续）" heading in the next
                #        column -- the repeated header the same instruction
                #        asked us to avoid;
                #      - and the stranded lead-in left the rest of its own
                #        column empty (one real column sat ~50% blank).
                #
                #    Restored deliberately narrowed, on the user's later
                #    "如果像明細如下這種當然不要 但如果是只是說明關係…" and their
                #    explicit acceptance of the （续）heading. Two gates, and
                #    both matter:
                #
                #    (a) The lead must NOT announce the table
                #        (lead_promises_table). That is the whole of the first
                #        objection above: a lead ending "明细如下：" is a
                #        sentence completed BY the table, and separating them
                #        breaks it. A lead that already stands as a complete
                #        statement does not break.
                #    (b) It must actually buy something -- the lead has to fill
                #        a real part of what is left, not strand two words at
                #        the bottom of the column. That is the third objection:
                #        the old version fired even when the lead was tiny.
                #
                #    The table itself is still never divided (no repeated
                #    header inside a table) per the original instruction.
                lead_text = str(item.get("commentary") or "")
                if (
                    lead_pt > 0
                    and not lead_promises_table(lead_text)
                    and _fits(cursor, lead_pt)
                    and lead_pt >= self._TABLE_SPLIT_MIN_LEAD_FILL * _remaining(cursor)
                ):
                    _append_to_slot(cursor, dict(item, _render_parts=("lead",)), lead_pt)
                    cursor = _open_new_slot(
                        dict(item, _render_parts=("table", "expl")),
                        table_pt + explain_pt + heading,
                    )
                    return

            cursor = _open_new_slot(item, block_pt)

        for item in table_items:
            table = item.get("_presentation_table") or {}
            parts_pt = self._estimate_table_account_parts_pt(item, table, is_chinese_databook)
            flow(item, sum(parts_pt), parts_pt=parts_pt)

        for item in (trailing_items or []):
            flow(item, self._estimate_lead_in_pt(item, is_chinese_databook))

        # Represent every OTHER slot on a slide this pass wrote to, as an
        # explicit empty entry. The render loop fills and clears purely by
        # walking the distribution, so a slot that simply never appears in
        # it is visited by neither branch and keeps whatever raw text the
        # template shipped with -- a real export leaked a literal
        # "Placeholder - placeholder" into slide 5's empty right column
        # this way. _distribute_content_across_slots already does exactly
        # this for the ordinary packing pool; the table pass needed its
        # own copy because it can leave a slot untouched that the ordinary
        # pass never saw either.
        touched_slides = {slide_idx for slide_idx, _slot, accounts in result if accounts}
        present = {(slide_idx, slot) for slide_idx, slot, _accounts in result}
        for slide_idx in sorted(touched_slides):
            actual_slide_idx = start_slide - 1 + slide_idx
            for slot_name in self._slot_names_for_actual_slide(actual_slide_idx, start_slide):
                if (slide_idx, slot_name) not in present:
                    result.append((slide_idx, slot_name, []))

        return result


    def _distribute_content_across_slots(
        self,
        structured_data: List[Dict],
        max_slides: int = 4,
        start_slide: int = 1,
        statement_type: Optional[str] = None,
    ):
        """
        Distribute content across textbox slots based on capacity.
        Slot structure is derived from the actual template slides when they exist.
        Auto-added slides follow the template convention: page 1 keeps a single right-side
        commentary box beside the table, while later slides use left/right commentary slots.
        
        Returns: List of (slide_idx, slot_idx, [account_data], is_partial, continuation_of)
        """
        if not structured_data:
            return []

        # Find a textbox shape to calculate capacity
        sample_shape = None
        for slide in self.presentation.slides:
            for alt_name in ["textMainBullets", "textMainBullets_L", "textMainBullets_R"]:
                shape = find_shape_by_name(slide.shapes, alt_name)
                if shape:
                    sample_shape = shape
                    break
            if sample_shape:
                break
        max_lines_per_textbox = (
            self._calculate_max_lines_for_textbox(sample_shape, statement_type=statement_type)
            if sample_shape
            else 40
        )
        
        logger.info("\n%s", '='*80)
        logger.info("CONTENT DISTRIBUTION STARTING")
        logger.info("%s", '='*80)
        logger.info("Total accounts: %s", len(structured_data))
        logger.info("Max lines per textbox: %s", max_lines_per_textbox)
        if sample_shape:
            logger.info("Sample shape height: %.2f inches", sample_shape.height / 914400)
            logger.info("Estimated capacity: %s lines", max_lines_per_textbox)
        logger.info("%s\n", '='*80)
        
        # _slot_names_for_actual_slide is the single source of truth for
        # this (also used by _append_table_accounts_to_distribution) --
        # see its docstring for why the two must never diverge.
        def slot_names_for_actual_slide(actual_slide_idx: int) -> List[str]:
            return self._slot_names_for_actual_slide(actual_slide_idx, start_slide)

        # Define slot structure: (slide_idx, slot_name)
        slots: List[Tuple[int, str]] = []
        for slide_idx in range(max_slides):
            actual_slide_idx = start_slide - 1 + slide_idx
            for slot_name in slot_names_for_actual_slide(actual_slide_idx):
                slots.append((slide_idx, slot_name))

        slot_shapes: Dict[int, Any] = {}
        for slot_idx, (slide_idx, slot_name) in enumerate(slots):
            actual_slide_idx = start_slide - 1 + slide_idx
            slot_shape = None
            if 0 <= actual_slide_idx < len(self.presentation.slides):
                slide = self.presentation.slides[actual_slide_idx]
                slot_shape = self._resolve_commentary_slot_shape(slide, slot_name)
            slot_shapes[slot_idx] = slot_shape or sample_shape

        logger.info("Total slots available: %s", len(slots))
        
        # Distribution result: [(slide_idx, slot_name, [account_data])]
        distribution = []
        
        current_slot_idx = 0
        current_slot_content = []
        current_slot_lines = 0
        previous_category = None

        def slot_capacity_for(slot_idx: int, *, is_chinese: bool, slot_name_override: Optional[str] = None) -> int:
            _slide_idx, derived_slot_name = slots[slot_idx]
            slot_name_local = slot_name_override or derived_slot_name
            slot_shape_local = slot_shapes.get(slot_idx)
            capacity = self._calculate_max_lines_for_textbox(
                slot_shape_local,
                is_chinese=is_chinese,
                slot_name=slot_name_local,
                statement_type=statement_type,
            )
            if slot_name_local == 'L':
                capacity = int(capacity * 0.98)
            return capacity
        
        for account_idx, account_data in enumerate(structured_data):
            mapping_key_debug = account_data.get('mapping_key', account_data.get('account_name', ''))
            logger.info("\nAccount %s/%s: %s", account_idx + 1, len(structured_data), mapping_key_debug)
            if current_slot_idx >= len(slots):
                dropped_accounts = len(structured_data) - account_idx
                logger.warning(
                    "Ran out of commentary slots; dropping %s remaining account(s) starting from '%s'",
                    dropped_accounts,
                    mapping_key_debug,
                )
                break
            
            category = account_data.get('category', '')
            mapping_key = account_data.get('mapping_key', account_data.get('account_name', ''))
            commentary = account_data.get('commentary', '')

            slide_idx_check, slot_name_check = slots[current_slot_idx]
            is_chinese_content = any('\u4e00' <= c <= '\u9fff' for c in commentary)
            chars_setting = 35 if is_chinese_content else 70
            category_lines = 1 if (category and category != previous_category) else 0
            content_lines = self._calculate_content_lines(
                '',
                mapping_key,
                commentary,
                slot_name=slot_name_check,
                shape=slot_shapes.get(current_slot_idx),
                is_chinese=is_chinese_content,
                statement_type=statement_type,
            )
            total_lines = category_lines + content_lines
            logger.info("  Category: '%s', Lines: cat=%s, content=%s, total=%s", category, category_lines, content_lines, total_lines)
            logger.info("  Commentary length: %s chars, Language: %s, Chars/line: %s", len(commentary), 'Chinese' if is_chinese_content else 'English', chars_setting)

            adjusted_capacity = slot_capacity_for(current_slot_idx, is_chinese=is_chinese_content, slot_name_override=slot_name_check)
            logger.info("  Current slot %s (%s): %s/%s lines used", current_slot_idx, slot_name_check, current_slot_lines, adjusted_capacity)

            if current_slot_lines + total_lines <= adjusted_capacity:
                current_slot_content.append(account_data)
                current_slot_lines += total_lines
                previous_category = category
                logger.info("  Slot %s (%s): Added '%s' (%s lines), now %s/%s lines used", current_slot_idx, slot_name_check, mapping_key, total_lines, current_slot_lines, adjusted_capacity)
            else:
                remaining_lines = adjusted_capacity - current_slot_lines
                logger.info("  Doesn't fit. Remaining: %s lines, Content: %s lines", remaining_lines, content_lines)

                # NOTE: the tail-overflow tolerance is deliberately NOT applied
                # here. Measured on a 14-account workload, keeping an account
                # whole at THIS point made the final deck worse (4 split
                # fragments -> 5): this greedy pass only produces the starting
                # point for _optimize_slot_fill, which re-packs everything
                # anyway, so suppressing a split here just perturbs the DP's
                # input and it re-splits elsewhere. The tolerance belongs at
                # the pass that actually decides to cut an account in half --
                # see _rebalance_overflowing_boundaries.
                next_slot_idx = current_slot_idx + 1

                split_remaining_min = float(self._packing_settings(statement_type).get("split_min_remaining_lines", 3))
                split_content_min = int(self._packing_settings(statement_type).get("split_min_content_lines", 5))
                if remaining_lines > split_remaining_min and content_lines > split_content_min:
                    logger.info("  Attempting to split content...")
                    paragraphs = commentary.split('\n\n')
                    if len(paragraphs) == 1:
                        paragraphs = commentary.split('\n')

                    chars_per_line = self._estimate_chars_per_line(
                        slot_name_check,
                        is_chinese_content,
                        shape=slot_shapes.get(current_slot_idx),
                        statement_type=statement_type,
                    )
                    available_for_commentary = remaining_lines - category_lines - 1

                    # Convert float line-units to visual display lines for the
                    # paragraph-fitting loop below.  available_for_commentary is
                    # in "std_lh units" (one unit = line_h + space_after pt),
                    # but para_lines is computed via chars_per_line and counts
                    # visual display lines.  Multiply by (std_lh / line_h) so
                    # both are in the same unit.
                    #   Both languages: std_lh=9+3=12pt, line_h=9pt → factor ≈ 1.333
                    #   (matches _real_font_size_pt/_real_line_spacing/_PARA_SPACE_AFTER
                    #   -- _fill_text_main_bullets_with_category_and_key hardcodes
                    #   Pt(3) space_after / line_spacing=1.0 for every paragraph,
                    #   any language; there's no separate Chinese 10pt/0.95 value.)
                    _lh_est = 9.0
                    _std_lh_est = _lh_est + self._PARA_SPACE_AFTER
                    available_visual = available_for_commentary * (_std_lh_est / _lh_est)

                    if available_for_commentary > 0:
                        part1_commentary = None
                        part2_commentary = None
                        part1_paragraphs = []
                        part1_lines_used = 0
                        split_index = 0

                        for i, para in enumerate(paragraphs):
                            para_lines = max(1, (len(para) + chars_per_line - 1) // chars_per_line)
                            if part1_lines_used + para_lines <= available_visual:
                                part1_paragraphs.append(para)
                                part1_lines_used += para_lines
                                split_index = i + 1
                            else:
                                break

                        if split_index == len(paragraphs) and part1_paragraphs:
                            # All paragraphs fit in remaining space per heuristic.
                            # Pillow slightly over-counted (< ~1-2 lines) — tolerate
                            # it and force-add to current slot rather than leave a
                            # half-empty slot with nothing below.
                            current_slot_content.append(account_data)
                            current_slot_lines += total_lines
                            previous_category = category
                            logger.info(
                                "  Heuristic fit: forced '%s' into slot %s "
                                "(Pillow over-count tolerated, total_lines=%.1f)",
                                mapping_key, current_slot_idx, total_lines,
                            )
                            continue  # skip move-to-next-slot fallthrough
                        elif part1_paragraphs and split_index < len(paragraphs):
                            # Clean paragraph-boundary split — always safe.
                            part1_commentary = '\n\n'.join(part1_paragraphs).strip()
                            part2_commentary = '\n\n'.join(paragraphs[split_index:]).strip()
                        elif not part1_paragraphs and len(paragraphs) > 0:
                            para = paragraphs[0]
                            chars_available = int(max(1, available_visual * chars_per_line))

                            if len(para) > chars_available:
                                # Only split at SENTENCE boundaries (period,
                                # Chinese full-stop, "!", "?"). No commas,
                                # word-breaks, or hard char cuts — those
                                # produce ugly mid-row fragments.
                                #
                                # Strategy: collect EVERY sentence ending up
                                # to 5 % past chars_available, then pick the
                                # one closest to chars_available without
                                # going over it. This packs the current slot
                                # tight — like a human would — rather than
                                # grabbing the first boundary found and
                                # leaving 5–6 rows of empty space below.
                                hard_cap = min(len(para), int(chars_available * 1.05))
                                _split_min_ratio = float(
                                    self._packing_settings(statement_type).get("split_min_fill_ratio", 0.6)
                                )
                                min_fill = max(1, int(chars_available * _split_min_ratio))
                                end_positions: List[int] = []
                                for end_char in ['. ', '。', '! ', '！', '? ', '？']:
                                    start = 0
                                    while True:
                                        pos = para.find(end_char, start, hard_cap)
                                        if pos < 0:
                                            break
                                        end_positions.append(pos + len(end_char))
                                        start = pos + 1
                                # Keep only splits that still leave the slot
                                # at least 15 % filled — avoids cutting after
                                # the first tiny opening sentence.
                                candidates = [p for p in end_positions if p >= min_fill]
                                best_split = max(candidates) if candidates else None

                                if best_split is None:
                                    # No sentence boundary fits — fall back to
                                    # word boundary to use all available lines
                                    # rather than leave the slot empty.
                                    word_end = para.rfind(' ', 0, hard_cap)
                                    if word_end > 0:
                                        best_split = word_end + 1
                                    else:
                                        # Chinese has no spaces, so the word
                                        # fallback above NEVER fires and this
                                        # dropped straight to a raw character
                                        # cut. That is the real defect behind
                                        # the reported "...某某系统" | "工程第四建设
                                        # 有限公司" -- a cut
                                        # through the middle of a company
                                        # name. _snap_split_before_number
                                        # can't rescue it either: "系统" is a
                                        # legitimate jieba token boundary, so
                                        # nothing downstream sees a problem.
                                        # A clause boundary is always a better
                                        # cut than an arbitrary character, and
                                        # is only used when no sentence
                                        # boundary was available at all.
                                        clause_end = max(
                                            (para.rfind(c, min_fill, hard_cap)
                                             for c in ('，', '；', '、', '：', ',', ';')),
                                            default=-1,
                                        )
                                        if clause_end > 0:
                                            best_split = clause_end + 1
                                        elif chars_available < len(para):
                                            best_split = chars_available  # last-resort hard cut

                                # Never split inside a number -- this is a
                                # SEPARATE, independent split implementation
                                # from _split_commentary_at_boundary (this is
                                # the FIRST-ever split an account gets, during
                                # the initial greedy distribution, before any
                                # rebalance pass runs) and _snap_split_before_
                                # number's fix there never touched this one --
                                # confirmed as the real source of a production
                                # case reading "...室外工程人民币2," at the
                                # bottom of one slide and "818.7万元..."
                                # continuing the next, AFTER that other fix
                                # already shipped.
                                if best_split:
                                    best_split = self._snap_split_before_number(para, best_split)

                                # Slice ONCE on the finalised best_split. Previously the
                                # word-boundary/hard-cut fallback recomputed best_split but
                                # never sliced, leaving part1/part2 unset → UnboundLocalError
                                # or stale text from a prior account bleeding onto the slide.
                                if best_split:
                                    part1_commentary = para[:best_split].strip()
                                    remaining_para = para[best_split:].strip()
                                    if len(paragraphs) > 1:
                                        part2_commentary = remaining_para + '\n\n' + '\n\n'.join(paragraphs[1:])
                                    else:
                                        part2_commentary = remaining_para
                                else:
                                    # No boundary and no cut possible — keep the whole
                                    # paragraph in part1 rather than corrupting the slide.
                                    part1_commentary = para
                                    part2_commentary = '\n\n'.join(paragraphs[1:]) if len(paragraphs) > 1 else None
                            else:
                                part1_commentary = para
                                part2_commentary = '\n\n'.join(paragraphs[1:]) if len(paragraphs) > 1 else ""
                        else:
                            part1_commentary = commentary
                            part2_commentary = None

                        if part1_commentary and part2_commentary:
                            account_part1 = account_data.copy()
                            account_part1['commentary'] = part1_commentary
                            account_part1['is_partial'] = True
                            account_part1['part_num'] = 1
                            current_slot_content.append(account_part1)
                            
                            # Save current slot
                            slide_idx, slot_name = slots[current_slot_idx]
                            distribution.append((slide_idx, slot_name, current_slot_content))
                            logger.info("Split '%s': Part 1 (%s chars) to slot %s, Part 2 (%s chars) to next slot", mapping_key, len(part1_commentary), current_slot_idx, len(part2_commentary))

                            if current_slot_idx + 1 >= len(slots):
                                logger.warning(
                                    "Ran out of commentary slots after splitting '%s'; dropping the remaining continuation",
                                    mapping_key,
                                )
                                break
                            current_slot_idx += 1

                            account_part2 = account_data.copy()
                            account_part2['commentary'] = part2_commentary
                            account_part2['is_continuation'] = True
                            account_part2['part_num'] = 2
                            account_part2['original_key'] = mapping_key

                            next_slot_name = slots[current_slot_idx][1]
                            part2_lines = self._calculate_content_lines(
                                '',
                                mapping_key,
                                part2_commentary,
                                slot_name=next_slot_name,
                                shape=slot_shapes.get(current_slot_idx),
                                is_chinese=is_chinese_content,
                            )
                            current_slot_content = [account_part2]
                            current_slot_lines = part2_lines
                            previous_category = None
                            continue
                else:
                    logger.info("  Not splitting: remaining_lines=%s, content_lines=%s", remaining_lines, content_lines)

                if current_slot_content:
                    slide_idx, slot_name = slots[current_slot_idx]
                    distribution.append((slide_idx, slot_name, current_slot_content))
                    logger.info("  Slot %s (%s): FULL with %s accounts, %s lines used", current_slot_idx, slot_name, len(current_slot_content), current_slot_lines)

                current_slot_idx += 1
                if current_slot_idx >= len(slots):
                    logger.warning(
                        "Ran out of commentary slots while placing '%s'; dropping that account from the remaining slides",
                        mapping_key,
                    )
                    break

                slide_idx_new, slot_name_new = slots[current_slot_idx]
                moved_account = account_data.copy()
                moved_account["commentary"] = commentary
                moved_category_lines = 1 if category else 0
                moved_lines = self._calculate_content_lines(
                    '',
                    mapping_key,
                    moved_account["commentary"],
                    slot_name=slot_name_new,
                    shape=slot_shapes.get(current_slot_idx),
                    is_chinese=is_chinese_content,
                    statement_type=statement_type,
                )
                current_slot_content = [moved_account]
                current_slot_lines = moved_category_lines + moved_lines
                previous_category = category
                logger.info("  Moving '%s' to next slot %s (%s), %s lines", mapping_key, current_slot_idx, slot_name_new, current_slot_lines)
        
        # Save last slot if it has content
        if current_slot_content and current_slot_idx < len(slots):
            slide_idx, slot_name = slots[current_slot_idx]
            distribution.append((slide_idx, slot_name, current_slot_content))
            logger.info("  Slot %s (%s): FINAL with %s accounts, %s lines", current_slot_idx, slot_name, len(current_slot_content), current_slot_lines)
        
        slot_position_map = {slot: idx for idx, slot in enumerate(slots)}

        logger.info("\nDistribution complete: %s slots filled", len(distribution))
        # Diagnostic fill-ratio logging used to recompute Pillow measurements
        # for every (slot, account) pair after the packer was already done.
        # That added 1-3s per export with no functional value. Skip the
        # recompute and just log slot composition.
        if logger.isEnabledFor(logging.DEBUG):
            for distribution_idx, (slide_idx, slot_name, accounts) in enumerate(distribution):
                slot_idx = slot_position_map.get((slide_idx, slot_name), distribution_idx)
                logger.debug(
                    "  Slot %s (Slide %s, %s): %s accounts",
                    slot_idx, slide_idx, slot_name, len(accounts),
                )
        
        # --- Fill optimization pass: pull accounts forward into under-filled slots ---
        distribution = self._optimize_slot_fill(
            distribution,
            slot_shapes=slot_shapes,
            slot_meta=slots,
            statement_type=statement_type,
        )

        # A slide that ends up with real content in one slot (e.g. L) but
        # never had ANY entry at all for one of its OTHER slots (e.g. R) --
        # not drained to empty by a rebalance pass, just never touched by
        # the greedy first pass above, which only appends a distribution
        # entry "if current_slot_content" (line ~2776) and so silently
        # skips a slot that ran out of content before reaching it -- needs
        # that other slot represented too, even with zero accounts.
        # _optimize_slot_fill can't fix this on its own: it only ever
        # builds its internal slot list FROM the `distribution` it's given
        # (see its "for slide_idx, slot_name, _accounts in distribution"
        # loop), so a slot that was never in `distribution` to begin with
        # never reaches it as input, let alone its output -- unlike the
        # "consolidated to empty" case _optimize_slot_fill's own output
        # filter already handles (5da5d38), this one is invisible to it
        # entirely. Confirmed via a real export: the resulting shape is
        # left holding whatever raw template placeholder text it shipped
        # with (a literal "Placeholder – placeholder"), because the
        # render loop's per-slot fill AND its "clear if empty" fallback
        # both key off `slot_contents.items()` -- a slot name that's
        # simply not a key in that dict is never visited by either branch.
        slides_with_content = {slide_idx for slide_idx, _slot_name, accounts in distribution if accounts}
        present_pairs = {(slide_idx, slot_name) for slide_idx, slot_name, _accounts in distribution}
        for slide_idx, slot_name in slots:
            if slide_idx in slides_with_content and (slide_idx, slot_name) not in present_pairs:
                distribution.append((slide_idx, slot_name, []))

        return distribution


    def _compute_slot_used_lines(
        self,
        accounts: List[Dict],
        slot_name: str,
        slot_shape=None,
        statement_type: Optional[str] = None,
    ) -> float:
        """Return used line-units for *accounts* in this slot (float).

        Uses the same accounting as ``slot_cost`` in the DP: each category
        header costs 1 line unit, and each account's commentary costs the
        float value returned by ``_calculate_content_lines`` (actual pt /
        std_lh, no ceil).  Comparing against int capacity from
        ``_calculate_max_lines_for_textbox`` gives accurate fill ratios.
        """
        used: float = 0.0
        prev_cat = None
        for account in accounts:
            cat = str(account.get("category", "") or "")
            # A continuation account never gets its own category-header
            # paragraph rendered (the fill loop explicitly skips it via
            # "not is_continuation" -- a "(cont'd)" fragment belongs to
            # whatever category its first part already introduced, not a
            # second one) -- so it must not be charged the 1.0-line cost
            # for one here either. A continuation is very often the FIRST
            # account placed in a slot (that's how it got split off), so
            # skipping this made every such slot's "used" belief exactly
            # 1.0 line higher than what actually gets rendered -- a
            # believed-vs-actual gap that compounds with the DP's own
            # front-loading and was a real contributor to slots reading as
            # under-filled despite the packing math saying they were full.
            # prev_cat only advances on an actually-rendered header (mirrors
            # the render loop's own current_category, which likewise never
            # moves off a skipped continuation) -- otherwise a later account
            # sharing the continuation's category would wrongly believe its
            # own header was already shown further up.
            if cat and cat != prev_cat and not account.get("is_continuation"):
                if prev_cat is None:
                    # The slot's OWN first category header gets NO space_before
                    # (matches the render's "p_category.space_before = Pt(3)
                    # if current_category else Pt(0)" -- current_category is
                    # still unset the very first time). Charging the full
                    # 1.0-unit (line_h+gap) here overstated every slot's
                    # opening header by one gap's worth -- confirmed by
                    # replicating the render's own paragraph-by-paragraph pt
                    # math for a real 7-account/2-category slot and comparing
                    # totals directly (0.23-unit gap, ~3pt, matched exactly).
                    _approx_line_h = _real_font_size_pt(False) * _real_line_spacing(False)
                    _approx_std_lh = _approx_line_h + _real_para_gap_pt(False)
                    used += (_approx_line_h / _approx_std_lh) if _approx_std_lh > 0 else 1.0
                else:
                    used += 1.0   # category header (same as slot_cost)
                prev_cat = cat
            used += self._calculate_content_lines(
                "",
                _account_cost_key(account),
                account.get("commentary", ""),
                slot_name=slot_name,
                shape=slot_shape,
                is_chinese=_account_is_chinese(account),
                statement_type=statement_type,
            )
            # A presentation-table account's `commentary` is only its lead-in;
            # the table itself and the explanation below it are reserved as
            # blank lines by _render_table_accounts_stack and were not counted
            # here at all. On a real deck that read a slot the renderer fills
            # to 94% as 10% used -- and once inspect_databook's fill
            # diagnostic started modelling the real slot assignment, that
            # under-count immediately produced false "GENUINE GAP" findings
            # (an account "fits in the 16.7 remaining lines" whose table alone
            # needs 24). Safe to charge here because table accounts never
            # reach the packer: they are pulled out before
            # _distribute_content_across_slots and appended afterwards, so
            # the only callers that ever see one are the render-time autofit
            # gate and the diagnostic -- both of which want the real height.
            # ...charged per PART, matching _render_parts, which is how a
            # split table account is actually rendered: the first fragment
            # draws ("lead", "table") and the continuation draws ("expl",)
            # only. Charging the whole block to both put 管理费用's table on
            # two slots and read a 67%-full column as 117%. Absent
            # _render_parts means the account renders whole.
            table = account.get("_presentation_table")
            if table:
                _is_chi = _account_is_chinese(account)
                _parts = account.get("_render_parts")
                _lead_pt, _table_pt, _explain_pt = self._estimate_table_account_parts_pt(
                    account, table, _is_chi,
                )
                _std_lh = _planning_std_lh_pt(_is_chi)
                if _std_lh > 0:
                    _extra = 0.0
                    if _parts is None or "table" in _parts:
                        _extra += _table_pt
                    if _parts is None or "expl" in _parts:
                        _extra += _explain_pt
                    used += _extra / _std_lh
        return max(0.0, used)


    def _rebalance_lopsided_lr_pairs(
        self,
        assignment: List[List[Dict[str, Any]]],
        slots: List[Dict[str, Any]],
        statement_type: Optional[str],
    ) -> None:
        """Mutates `assignment` in place. Fixes same-page L/R pairs where the
        DP left one column completely empty while the other is significantly
        full.

        Root cause: _optimize_slot_fill's lexicographic objective is
        (num_nonempty_slots, underfill_penalty). An EMPTY slot contributes
        NEITHER — it costs the DP nothing — while a non-empty slot under
        target_fill_min_ratio incurs a real penalty. So whenever a page's
        total content doesn't reach ~2x a single slot's capacity, "one
        slot full + one slot empty" scores strictly better than "two slots
        each moderately full" (e.g. one at 90% + one empty beats two at
        45% each: 5% penalty vs 50%+50%). That is a reasonable objective
        for deciding how many PAGES to use, but applied to the two columns
        of a single page it produces a visually broken half-blank layout a
        reader would never expect.

        This is a post-pass, not a DP objective change, to keep the fix
        narrowly scoped: find same-slide (L, R) pairs with exactly one
        side empty, and if the full side has 2+ accounts, look for the
        split point whose two halves are most evenly balanced (by the same
        real line-cost function the DP itself uses) without overflowing
        either box. If no such split exists, leave the pair as the DP
        produced it — this never risks introducing an overflow that wasn't
        there before.
        """
        by_slide: Dict[int, Dict[str, int]] = {}
        for s_i, slot in enumerate(slots):
            if slot.get("slot_name") in ("L", "R"):
                by_slide.setdefault(slot["slide_idx"], {})[slot["slot_name"]] = s_i

        for slide_idx, pair in by_slide.items():
            if "L" not in pair or "R" not in pair:
                continue
            l_i, r_i = pair["L"], pair["R"]
            l_accts, r_accts = assignment[l_i], assignment[r_i]
            if bool(l_accts) == bool(r_accts):
                continue  # both empty, or both already non-empty -- nothing to rebalance
            full_i, empty_i = (l_i, r_i) if l_accts else (r_i, l_i)
            full_accts = assignment[full_i]
            if len(full_accts) < 2:
                continue  # can't split a single account here -- that's the greedy first pass's job

            full_name = slots[full_i]["slot_name"]
            empty_name = slots[empty_i]["slot_name"]
            full_cap = slots[full_i]["capacity"]
            empty_cap = slots[empty_i]["capacity"]
            full_shape = slots[full_i]["shape"]
            empty_shape = slots[empty_i]["shape"]

            best_k, best_diff = None, None
            for k in range(1, len(full_accts)):
                first, rest = full_accts[:k], full_accts[k:]
                first_lines = self._compute_slot_used_lines(
                    first, empty_name, slot_shape=empty_shape, statement_type=statement_type,
                )
                rest_lines = self._compute_slot_used_lines(
                    rest, full_name, slot_shape=full_shape, statement_type=statement_type,
                )
                if first_lines > empty_cap or rest_lines > full_cap:
                    continue
                diff = abs(first_lines - rest_lines)
                if best_diff is None or diff < best_diff:
                    best_diff, best_k = diff, k

            if best_k is None:
                continue  # no split keeps both sides within capacity -- leave as-is

            assignment[empty_i] = full_accts[:best_k]
            assignment[full_i] = full_accts[best_k:]
            logger.info(
                "  Rebalanced lopsided L/R pair on slide %s: moved %s account(s) from slot %s "
                "into previously-empty slot %s",
                slide_idx, best_k, full_i, empty_i,
            )


    def _consolidate_tiny_stub_lr_pairs(
        self,
        assignment: List[List[Dict[str, Any]]],
        slots: List[Dict[str, Any]],
        statement_type: Optional[str],
    ) -> None:
        """Mutates `assignment` in place. Front-loading -- filling L to
        near-capacity before spilling anything into R -- is the intended
        packing philosophy here (minimises total page count), not a bug to
        fix. But confirmed against real production decks: it regularly
        leaves R holding a single orphaned fragment at ~5-20% fill next to
        an L column in the 90s% -- not what a human preparing the same
        page would produce (they'd either use both columns for real or
        fold a leftover sliver into the fuller one, never leave a
        near-blank second column standing next to a full one).

        This only ever fires on a same-slide (L, R) pair where R is
        NON-empty (bool(r_accts) is True, so _rebalance_lopsided_lr_pairs'
        "one side fully empty" case never overlaps with this one) but its
        fill ratio reads as an orphaned stub rather than real content, AND
        R's entire content fits into L's own remaining capacity whole --
        folds it in and leaves R properly, cleanly empty (a genuinely
        blank box now correctly renders that way -- see the
        "if not account_data_list" shape-clearing fix alongside this).
        Never attempts a partial move here: trading one awkward stub for a
        smaller one doesn't serve the "look like a human made this" goal.
        """
        STUB_FILL_THRESHOLD = 0.20  # below this, a slot reads as leftover, not real content

        by_slide: Dict[int, Dict[str, int]] = {}
        for s_i, slot in enumerate(slots):
            if slot.get("slot_name") in ("L", "R"):
                by_slide.setdefault(slot["slide_idx"], {})[slot["slot_name"]] = s_i

        for slide_idx, pair in by_slide.items():
            if "L" not in pair or "R" not in pair:
                continue
            l_i, r_i = pair["L"], pair["R"]
            l_accts, r_accts = assignment[l_i], assignment[r_i]
            if not l_accts or not r_accts:
                continue  # nothing to fold, or _rebalance_lopsided_lr_pairs's job

            l_slot, r_slot = slots[l_i], slots[r_i]
            r_used = self._compute_slot_used_lines(
                r_accts, r_slot["slot_name"], slot_shape=r_slot["shape"], statement_type=statement_type,
            )
            r_cap = r_slot["capacity"]
            if r_cap <= 0 or (r_used / r_cap) >= STUB_FILL_THRESHOLD:
                continue  # R holds real content, not just a leftover stub -- leave it alone

            combined_used = self._compute_slot_used_lines(
                l_accts + r_accts, l_slot["slot_name"], slot_shape=l_slot["shape"], statement_type=statement_type,
            )
            if combined_used > l_slot["capacity"]:
                continue  # doesn't fit whole -- a partial move would just trade one stub for another

            assignment[l_i] = l_accts + r_accts
            assignment[r_i] = []
            logger.info(
                "  Consolidated tiny stub on slide %s: folded R (%.0f%% fill) into L -- "
                "avoids an orphaned near-empty trailing column",
                slide_idx, r_used / r_cap * 100,
            )


    def _try_partial_split_into_gap(
        self,
        cur_accts: List[Dict[str, Any]],
        nxt_accts: List[Dict[str, Any]],
        cur_used: float,
        cur_cap: int,
        cur_name: str,
        cur_shape,
        nxt_cap: int,
        nxt_name: str,
        nxt_shape,
        is_last_cur: bool,
        is_last_nxt: bool,
        target_fill: float,
        statement_type: Optional[str],
    ) -> bool:
        """Mutates `cur_accts`/`nxt_accts` in place if it commits. Attempts
        to move the FRONT part of nxt_accts[0]'s commentary back into
        cur_accts' remaining capacity, leaving the rest as a continuation
        at the front of nxt_accts. Generalizes the initial forward-fill
        pass's one-time split into a repeatable operation any boundary can
        trigger -- so a single account can end up split across more than
        two slots when that's what it takes to close a gap, instead of
        being force-moved whole (leaving the gap unfilled) whenever it
        doesn't fit. Only commits when it strictly lowers the pair's
        summed underfill penalty and both fragments clear a minimum-fill
        safeguard -- never overflows a box, never drops text.
        """
        # This floor is what caps the achievable fill: refusing to act below
        # it leaves exactly that much of every slot empty. At the old
        # hardcoded 1.5 a 23.9-line slot could never exceed
        # (23.9-1.5)/23.9 = 94%, which is precisely the fill reported on a
        # real deck. 1.0 is the meaningful floor -- below one rendered line
        # there is nothing worth moving, and _split_commentary_at_boundary's
        # own min_available_visual already refuses smaller budgets.
        # Budget against capacity PLUS the protrusion allowance, not capacity
        # alone. The acceptance test below already lets a candidate exceed
        # cur_cap by that much, so gating entry on the strict gap made the
        # two disagree: a real deck left a slot at 25.6/26.0 -- a 0.4-line
        # gap, under the 1.0 floor -- so nothing was even attempted, while a
        # ~2-line sentence from the next column would have been accepted had
        # it been offered. "小小溢出其實問題也不大" is exactly this case.
        tail_tol = self._tail_overflow_tolerance_units(statement_type)
        gap = (cur_cap + tail_tol) - cur_used
        if gap < float(self._packing_settings(statement_type).get("min_gap_to_split_lines", 1.0) or 1.0):
            return False

        head_acct = nxt_accts[0]
        cur_last_cat = str(cur_accts[-1].get("category", "") or "") if cur_accts else ""
        head_cat = str(head_acct.get("category", "") or "")
        category_gap_cost = 1.0 if (head_cat and head_cat != cur_last_cat) else 0.0
        text_budget = gap - category_gap_cost
        if text_budget < 1.0:
            return False

        is_chinese = _account_is_chinese(head_acct)
        # The split-point estimate (char-count based) and the real measurement
        # (_compute_slot_used_lines, font-metric based) don't perfectly agree --
        # shrink the requested budget and retry a few times if the chosen head
        # measures slightly over cur_cap, rather than giving up on the first
        # miss. Each retry backs off by the exact measured overage (plus a
        # small margin), so this converges in 1-2 extra tries in practice.
        part1 = None
        trial_cur_used = None
        remaining_budget = text_budget
        for _attempt in range(4):
            split_result = self._split_commentary_at_boundary(
                str(head_acct.get("commentary", "") or ""),
                remaining_budget,
                slot_name=cur_name,
                is_chinese=is_chinese,
                shape=cur_shape,
                statement_type=statement_type,
                key_prefix=f"■ {_account_cost_key(head_acct)} - ",
                # `gap` above is already capacity + allowance.
                overflow_allowance_units=0.0,
            )
            if not split_result:
                return False
            head_text, tail_text = split_result

            candidate = head_acct.copy()
            candidate["commentary"] = head_text
            candidate["is_partial"] = True
            candidate["part_num"] = int(head_acct.get("part_num") or 1)
            candidate["original_key"] = head_acct.get("original_key", head_acct.get("mapping_key"))

            candidate_used = self._compute_slot_used_lines(
                cur_accts + [candidate], cur_name, slot_shape=cur_shape, statement_type=statement_type,
            )
            # The tail tolerance belongs HERE, on the acceptance test, not
            # inside the splitter. This pass PULLS content into an underfilled
            # slot, so letting the result protrude a line or two fills the box
            # -- which is what "能夠再放1-2行才滿" asks for. (An earlier attempt
            # put the tolerance inside _split_commentary_at_boundary as a
            # "refuse to split" branch; that made the caller decline the pull
            # entirely and left the slot EMPTIER. Accepting more content and
            # declining to split are opposite things.)
            if candidate_used <= cur_cap + self._tail_overflow_tolerance_units(statement_type):
                part1, trial_cur_used = candidate, candidate_used
                break
            overage = candidate_used - cur_cap
            remaining_budget -= overage + 0.25
            if remaining_budget < 1.0:
                return False

        if part1 is None:
            return False  # estimate/actual mismatch never converged -- bail out safely

        part2 = head_acct.copy()
        part2["commentary"] = tail_text
        part2["is_continuation"] = True
        part2["part_num"] = int(head_acct.get("part_num") or 1) + 1
        part2["original_key"] = head_acct.get("original_key", head_acct.get("mapping_key"))

        trial_nxt_accts = [part2] + nxt_accts[1:]
        trial_nxt_used = self._compute_slot_used_lines(
            trial_nxt_accts, nxt_name, slot_shape=nxt_shape, statement_type=statement_type,
        )
        if trial_nxt_used > nxt_cap:
            return False

        orig_nxt_used = self._compute_slot_used_lines(
            nxt_accts, nxt_name, slot_shape=nxt_shape, statement_type=statement_type,
        )
        cur_penalty = 0.0 if is_last_cur else max(0.0, target_fill - (cur_used / cur_cap if cur_cap else 0.0))
        nxt_penalty = 0.0 if is_last_nxt else max(0.0, target_fill - (orig_nxt_used / nxt_cap if nxt_cap else 0.0))
        current_total = cur_penalty + nxt_penalty

        alt_cur_penalty = 0.0 if is_last_cur else max(0.0, target_fill - (trial_cur_used / cur_cap if cur_cap else 0.0))
        alt_nxt_penalty = 0.0 if is_last_nxt else max(0.0, target_fill - (trial_nxt_used / nxt_cap if nxt_cap else 0.0))
        alt_total = alt_cur_penalty + alt_nxt_penalty

        if alt_total >= current_total - 1e-9:
            return False

        cur_accts.append(part1)
        nxt_accts[0] = part2
        logger.info(
            "  Split '%s' across boundary: moved a fragment into slot %s, continuation stays in slot %s "
            "(pair penalty %.3f -> %.3f)",
            head_acct.get("mapping_key", "?"), cur_name, nxt_name, current_total, alt_total,
        )
        return True


    def _rebalance_underfilled_boundaries(
        self,
        assignment: List[List[Dict[str, Any]]],
        slots: List[Dict[str, Any]],
        statement_type: Optional[str],
    ) -> None:
        """Mutates `assignment` in place. Generalizes
        _rebalance_lopsided_lr_pairs beyond the "one slot totally empty"
        case: for every adjacent slot boundary (in packing order,
        including same-slide L/R pairs), if the next slot's first account
        would both fit in the current slot's remaining capacity AND
        strictly lower the DP's own lexicographic underfill penalty summed
        across the pair, move it back. If the whole account doesn't fit,
        falls back to _try_partial_split_into_gap to move just enough of
        its commentary to close the gap, leaving the rest as a
        continuation -- so a single account can be split across more than
        two slots when that's what filling every boundary takes.

        This is the exact same boundary check inspect_databook.py's
        analyze_population_fill uses to verify (not just guess) that a
        boundary is a genuine gap rather than already optimal -- turned
        from a passive diagnostic into an active fix. Runs to a fixed
        point (bounded by len(assignment) passes) since a move can only
        ever enable ANOTHER move at the same or a later boundary, never
        undo one, so a short chain of small imbalances in a row all get
        resolved, not just the first.
        """
        packing = self._packing_settings(statement_type)
        target_fill = float(packing.get("target_fill_min_ratio", 0.95) or 0.95)
        n = len(assignment)

        for _pass in range(n):
            changed = False
            for i in range(n - 1):
                nxt_accts = assignment[i + 1]
                if not nxt_accts:
                    continue
                cur_accts = assignment[i]
                cur_slot, nxt_slot = slots[i], slots[i + 1]
                cur_cap, nxt_cap = cur_slot["capacity"], nxt_slot["capacity"]
                cur_shape, nxt_shape = cur_slot["shape"], nxt_slot["shape"]
                cur_name, nxt_name = cur_slot["slot_name"], nxt_slot["slot_name"]

                cur_used = self._compute_slot_used_lines(
                    cur_accts, cur_name, slot_shape=cur_shape, statement_type=statement_type,
                )
                alt_cur_used = self._compute_slot_used_lines(
                    cur_accts + [nxt_accts[0]], cur_name, slot_shape=cur_shape, statement_type=statement_type,
                )
                is_last_cur = (i == n - 1)
                is_last_nxt = (i + 1 == n - 1)

                if alt_cur_used > cur_cap:
                    if self._try_partial_split_into_gap(
                        cur_accts, nxt_accts, cur_used, cur_cap, cur_name, cur_shape,
                        nxt_cap, nxt_name, nxt_shape, is_last_cur, is_last_nxt,
                        target_fill, statement_type,
                    ):
                        changed = True
                    continue  # whole move doesn't fit -- partial split (if any) already handled

                # A same-slide L/R pair with nxt down to its LAST account is exactly
                # what _rebalance_lopsided_lr_pairs exists to prevent (a visibly empty
                # column next to a full one, exposing the template's raw placeholder
                # text) -- but when nxt is also the statement's last slot overall, its
                # penalty is unconditionally exempt (0.0 below regardless of fill), so
                # this whole-move would look like a free win and drain nxt right back
                # to empty, silently undoing that earlier fix. Refuse this specific
                # move; every other boundary (including L/R pairs with 2+ accounts
                # left in nxt) is untouched.
                same_slide_lr_pair = (
                    cur_slot["slide_idx"] == nxt_slot["slide_idx"]
                    and {cur_name, nxt_name} == {"L", "R"}
                )
                if same_slide_lr_pair and len(nxt_accts) == 1:
                    continue

                nxt_used = self._compute_slot_used_lines(
                    nxt_accts, nxt_name, slot_shape=nxt_shape, statement_type=statement_type,
                )
                alt_nxt_used = self._compute_slot_used_lines(
                    nxt_accts[1:], nxt_name, slot_shape=nxt_shape, statement_type=statement_type,
                )

                cur_penalty = 0.0 if is_last_cur else max(0.0, target_fill - (cur_used / cur_cap if cur_cap else 0.0))
                nxt_penalty = 0.0 if is_last_nxt else max(0.0, target_fill - (nxt_used / nxt_cap if nxt_cap else 0.0))
                current_total = cur_penalty + nxt_penalty

                alt_cur_penalty = 0.0 if is_last_cur else max(0.0, target_fill - (alt_cur_used / cur_cap if cur_cap else 0.0))
                alt_nxt_penalty = 0.0 if is_last_nxt else max(0.0, target_fill - (alt_nxt_used / nxt_cap if nxt_cap else 0.0))
                alt_total = alt_cur_penalty + alt_nxt_penalty

                if alt_total < current_total - 1e-9:
                    moved = nxt_accts.pop(0)
                    cur_accts.append(moved)
                    changed = True
                    logger.info(
                        "  Rebalanced underfilled boundary: moved '%s' from slot %s into slot %s "
                        "(pair penalty %.3f -> %.3f)",
                        moved.get("mapping_key", "?"), i + 1, i, current_total, alt_total,
                    )
            if not changed:
                break


    def _try_partial_split_overflow_forward(
        self,
        cur_accts: List[Dict[str, Any]],
        nxt_accts: List[Dict[str, Any]],
        cur_cap: int,
        cur_name: str,
        cur_shape,
        nxt_cap: int,
        nxt_name: str,
        nxt_shape,
        statement_type: Optional[str],
    ) -> bool:
        """Mutates `cur_accts`/`nxt_accts` in place if it commits. Mirrors
        _try_partial_split_into_gap but for the opposite direction: cur is
        already overflowing, so shrink its LAST account down to whatever
        genuinely fits within cur's own capacity and push the trimmed-off
        tail forward as a continuation at the FRONT of nxt. Only commits
        when the split fits both sides — never leaves cur still over
        capacity, never pushes nxt over its own.
        """
        tail_acct = cur_accts[-1]
        other_used = self._compute_slot_used_lines(
            cur_accts[:-1], cur_name, slot_shape=cur_shape, statement_type=statement_type,
        )
        available_for_tail = cur_cap - other_used
        if available_for_tail < 1.0:
            return False

        is_chinese = _account_is_chinese(tail_acct)
        part1 = None
        head_text = tail_text = ""
        remaining_budget = available_for_tail
        for _attempt in range(4):
            split_result = self._split_commentary_at_boundary(
                str(tail_acct.get("commentary", "") or ""),
                remaining_budget,
                slot_name=cur_name,
                is_chinese=is_chinese,
                shape=cur_shape,
                statement_type=statement_type,
                min_fill_ratio=0.3,
                key_prefix=f"■ {_account_cost_key(tail_acct)} - ",
            )
            if not split_result:
                return False
            head_text, tail_text = split_result

            candidate = tail_acct.copy()
            candidate["commentary"] = head_text
            candidate["is_partial"] = True
            candidate["part_num"] = int(tail_acct.get("part_num") or 1)
            candidate["original_key"] = tail_acct.get("original_key", tail_acct.get("mapping_key"))

            candidate_used = self._compute_slot_used_lines(
                cur_accts[:-1] + [candidate], cur_name, slot_shape=cur_shape, statement_type=statement_type,
            )
            if candidate_used <= cur_cap:
                part1 = candidate
                break
            overage = candidate_used - cur_cap
            remaining_budget -= overage + 0.25
            if remaining_budget < 1.0:
                return False

        if part1 is None:
            return False  # estimate/actual mismatch never converged -- bail out safely

        part2 = tail_acct.copy()
        part2["commentary"] = tail_text
        part2["is_continuation"] = True
        part2["part_num"] = int(tail_acct.get("part_num") or 1) + 1
        part2["original_key"] = tail_acct.get("original_key", tail_acct.get("mapping_key"))

        trial_nxt_accts = [part2] + nxt_accts
        trial_nxt_used = self._compute_slot_used_lines(
            trial_nxt_accts, nxt_name, slot_shape=nxt_shape, statement_type=statement_type,
        )
        if trial_nxt_used > nxt_cap:
            return False

        cur_accts[-1] = part1
        nxt_accts.insert(0, part2)
        logger.info(
            "  Split overflowing '%s' forward across boundary: head stays in current slot, "
            "continuation moves into next slot",
            tail_acct.get("mapping_key", "?"),
        )
        return True


    def _rebalance_overflowing_boundaries(
        self,
        assignment: List[List[Dict[str, Any]]],
        slots: List[Dict[str, Any]],
        statement_type: Optional[str],
    ) -> None:
        """Mutates `assignment` in place. Complements
        _rebalance_underfilled_boundaries, which only ever pulls a
        FOLLOWING slot's content BACKWARD into a preceding slot's positive
        gap — it never pushes content FORWARD out of an already-
        overflowing slot into a following slot with spare room, because
        the DP's own lexicographic penalty formula caps an overflowing
        slot's penalty at 0.0 (identical to a perfectly-filled slot), so
        overflow never registers as something worth fixing to either of
        the existing passes (confirmed via a real GPT-5.5/workbench IS
        export: one column measured 103%/OVERFLOW RISK while its same-
        slide sibling sat at 34% underfilled, and the population
        diagnostic's own genuine-gap check — which reuses that same
        penalty formula — reported no fixable gap at all).

        This pass looks at raw used-vs-capacity instead of the penalty
        formula specifically to catch that blind spot: for every adjacent
        boundary where the current slot's real content exceeds its
        capacity and the next slot has spare room, tries moving the
        current slot's LAST account forward whole; if that doesn't fit
        (or cur only has one account to begin with), falls back to
        _try_partial_split_overflow_forward. Never commits a move that
        would newly overflow the next slot.
        """
        n = len(assignment)
        for _pass in range(n):
            changed = False
            for i in range(n - 1):
                cur_accts = assignment[i]
                if not cur_accts:
                    continue
                nxt_accts = assignment[i + 1]
                cur_slot, nxt_slot = slots[i], slots[i + 1]
                cur_cap, nxt_cap = cur_slot["capacity"], nxt_slot["capacity"]
                cur_shape, nxt_shape = cur_slot["shape"], nxt_slot["shape"]
                cur_name, nxt_name = cur_slot["slot_name"], nxt_slot["slot_name"]

                cur_used = self._compute_slot_used_lines(
                    cur_accts, cur_name, slot_shape=cur_shape, statement_type=statement_type,
                )
                if cur_used <= cur_cap:
                    continue  # not overflowing -- nothing for this pass to do at this boundary

                nxt_used = self._compute_slot_used_lines(
                    nxt_accts, nxt_name, slot_shape=nxt_shape, statement_type=statement_type,
                ) if nxt_accts else 0.0
                if nxt_used >= nxt_cap:
                    continue  # next slot has no spare room either -- can't help here

                # Moving just the SINGLE last account used to be the only
                # whole-account move attempted -- if the accounts BEFORE it
                # already exceed cur_cap on their own (several trailing
                # accounts collectively overflow, not just the tail one),
                # `rest_used <= cur_cap` was never true, this whole branch
                # fell through, and _try_partial_split_overflow_forward's own
                # `other_used` (the same all-but-last cost) was equally over
                # cur_cap, so IT bailed immediately too -- nothing ever moved
                # at all. Confirmed via a real production case: a slot at
                # 109% sat next to a COMPLETELY EMPTY next slot (0% -- not
                # merely underfilled) because neither path could ever fire.
                # Grow the batch of trailing whole accounts to move one at a
                # time until what remains in cur actually fits, or until
                # moving the next one would overflow nxt -- whichever comes
                # first. Single-account overflow (the original case) is
                # k=1 here, unchanged in behavior.
                move_count = 0
                for k in range(1, len(cur_accts)):
                    rest_used = self._compute_slot_used_lines(
                        cur_accts[:-k], cur_name, slot_shape=cur_shape, statement_type=statement_type,
                    )
                    trial_nxt_accts = cur_accts[-k:] + nxt_accts
                    trial_nxt_used = self._compute_slot_used_lines(
                        trial_nxt_accts, nxt_name, slot_shape=nxt_shape, statement_type=statement_type,
                    )
                    if trial_nxt_used > nxt_cap:
                        break
                    move_count = k
                    if rest_used <= cur_cap:
                        break
                if move_count:
                    moved = cur_accts[-move_count:]
                    del cur_accts[-move_count:]
                    nxt_accts[0:0] = moved
                    changed = True
                    logger.info(
                        "  Rebalanced overflowing boundary: moved %d whole account(s) forward from slot %s "
                        "(was %.1f/%s) into slot %s",
                        move_count, i, cur_used, cur_cap, i + 1,
                    )
                    continue

                # No whole account could move. The only remaining option is
                # cutting one in half -- and a small overflow is not worth
                # that. The project team accepts 1-2 lines protruding below
                # the box; what they rejected was a split landing mid-name
                # ("某某系统" | "工程第四建设有限公司"). Note this gate is
                # deliberately AFTER the whole-account move above: moving an
                # account forward is free and still happens for any overflow.
                if cur_used <= cur_cap + self._tail_overflow_tolerance_units(statement_type):
                    logger.info(
                        "  Slot %s left overflowing %.2f line(s) over %s on purpose — "
                        "within tail tolerance, so not split",
                        i, cur_used - cur_cap, cur_cap,
                    )
                    continue

                if self._try_partial_split_overflow_forward(
                    cur_accts, nxt_accts, cur_cap, cur_name, cur_shape,
                    nxt_cap, nxt_name, nxt_shape, statement_type,
                ):
                    changed = True
            if not changed:
                break


    def _extend_continuation_in_place(
        self,
        cur_accts: List[Dict[str, Any]],
        nxt_accts: List[Dict[str, Any]],
        cur_name: str,
        cur_shape,
        cur_cap: float,
        statement_type: Optional[str],
    ) -> bool:
        """Mutates cur_accts/nxt_accts in place. Pulls more of nxt_accts[0]
        (a direct continuation of cur_accts[-1], same original account)
        into cur_accts[-1]'s own paragraph, instead of adding it as a new
        one. Returns True if anything changed.

        Why this exists: a NEW paragraph costs at minimum line_h+para_gap
        (~1.0 std_lh unit) even for a single character, since the wrap+
        cost formula charges a full line height for any non-empty text
        plus the fixed per-paragraph gap. _maximize_forward_fill's normal
        "add next account as a new paragraph" path can therefore never
        fill a gap smaller than that floor, however much real room is
        left. Extending an EXISTING paragraph's text only adds wrapped
        LINES (no extra para_gap), so it can use up a fractional-line gap
        the new-paragraph path mathematically cannot.
        """
        head_acct = cur_accts[-1]
        tail_acct = nxt_accts[0]
        head_commentary = str(head_acct.get("commentary", "") or "")
        tail_commentary = str(tail_acct.get("commentary", "") or "")
        if not tail_commentary.strip():
            return False
        combined = (head_commentary + tail_commentary).strip()

        other_used = self._compute_slot_used_lines(
            cur_accts[:-1], cur_name, slot_shape=cur_shape, statement_type=statement_type,
        )
        budget = cur_cap - other_used
        if budget <= 0:
            return False

        # First check whether the WHOLE combined text now fits -- if so,
        # absorb tail_acct entirely rather than leaving an artificially
        # truncated remainder.
        candidate_whole = dict(head_acct)
        candidate_whole["commentary"] = combined
        whole_used = self._compute_slot_used_lines(
            cur_accts[:-1] + [candidate_whole], cur_name, slot_shape=cur_shape, statement_type=statement_type,
        )
        if whole_used <= cur_cap:
            cur_accts[-1] = candidate_whole
            nxt_accts.pop(0)
            return True

        # Deliberately NOT using _split_commentary_at_boundary here: its
        # rough char-count pre-check can decide "the whole paragraph fits"
        # (a coarser, disagreeing estimate from the accurate whole_used
        # check just above) and return None on that basis alone, before
        # ever reaching its own accurate-measurer validation -- a real
        # false negative observed here (whole_used said no, but the
        # function still gave up immediately). A direct binary search
        # against the same accurate _compute_slot_used_lines this whole
        # pass already trusts sidesteps that disagreement entirely.
        lo, hi = len(head_commentary), len(combined) - 1
        best_len = len(head_commentary)
        while lo <= hi:
            mid = (lo + hi) // 2
            candidate = dict(head_acct)
            candidate["commentary"] = combined[:mid].strip()
            used = self._compute_slot_used_lines(
                cur_accts[:-1] + [candidate], cur_name, slot_shape=cur_shape, statement_type=statement_type,
            )
            if used <= cur_cap:
                best_len = mid
                lo = mid + 1
            else:
                hi = mid - 1

        if best_len <= len(head_commentary):
            return False

        # The binary search above optimises for FILL only, so best_len is a
        # raw character offset -- and _snap_split_before_number below only
        # rescues numbers and jieba tokens, not names (系统|工程 is a valid
        # token boundary, which is how a real deck ended up cutting
        # 某某系统 | 工程第四建设有限公司). Snap back to the last natural
        # sentence/clause boundary at or before best_len -- it necessarily
        # still fits, and it must still be a real extension of the head.
        _marks = ('。', '！', '？', '，', '；', '、', '.', ',', ';')
        _back = max((combined.rfind(c, 0, best_len + 1) for c in _marks), default=-1)
        if _back >= 0 and _back + 1 > len(head_commentary):
            best_len = _back + 1
        else:
            # No clause boundary anywhere inside the text this pass wanted to
            # pull forward. Extending anyway would keep the binary search's
            # raw character offset -- a mid-word cut, which is what showed up
            # the moment the gap gate started admitting smaller gaps. There
            # is nothing here worth a bad cut, so decline the extension.
            return False

        best_len = self._snap_split_before_number(combined, best_len)
        if best_len <= len(head_commentary):
            return False
        new_head = combined[:best_len].strip()
        new_tail = combined[best_len:].strip()
        if not new_head or not new_tail:
            return False

        candidate = dict(head_acct)
        candidate["commentary"] = new_head
        candidate_used = self._compute_slot_used_lines(
            cur_accts[:-1] + [candidate], cur_name, slot_shape=cur_shape, statement_type=statement_type,
        )
        if candidate_used > cur_cap:
            return False

        cur_accts[-1] = candidate
        tail_acct["commentary"] = new_tail
        return True


    def _maximize_forward_fill(
        self,
        assignment: List[List[Dict[str, Any]]],
        slots: List[Dict[str, Any]],
        statement_type: Optional[str],
    ) -> None:
        """Mutates `assignment` in place. User-confirmed requirement:
        given a FIXED amount of content, pack it the way a person would --
        fill the first slot as close to its own capacity as possible,
        THEN move on to the next slot with whatever's left, and so on.
        This is a plain greedy maximal-fill-in-reading-order pass,
        deliberately NOT balance-seeking (an empty or lightly-filled
        TRAILING slot is fine and expected once content genuinely runs
        out -- that's a separate, not-yet-requested question of whether
        to use fewer total slides). It directly supersedes _optimize_
        slot_fill's own DP objective ("minimise the MAXIMUM fill ratio
        across slots"), which spreads content for evenness rather than
        maximising each slot in turn -- runs LAST, after every other
        rebalance pass, since it's the final word on "is slot i as full
        as slot i's own capacity allows" and nothing after it should
        undo that.

        For each boundary in order: repeatedly pull the next slot's
        first account into the current slot, whole if it fits; if it
        doesn't fit whole, split it (same paragraph/sentence/word
        boundary search _try_partial_split_into_gap uses) to use up
        whatever room remains, leaving the rest as a continuation at the
        front of the next slot. Keeps going until the current slot has
        no meaningful room left (<0.5 lines) or the next slot is empty,
        then moves on to the next boundary.
        """
        n = len(assignment)
        for i in range(n - 1):
            cur_slot, nxt_slot = slots[i], slots[i + 1]
            cur_name = cur_slot["slot_name"]
            cur_shape = cur_slot["shape"]
            cur_cap = cur_slot["capacity"]

            while assignment[i + 1]:
                cur_accts = assignment[i]
                nxt_accts = assignment[i + 1]
                cur_used = self._compute_slot_used_lines(
                    cur_accts, cur_name, slot_shape=cur_shape, statement_type=statement_type,
                )
                gap = cur_cap - cur_used
                # Lowered from 0.5 -- with the accurate-measurer backoff/trim
                # now in _split_commentary_at_boundary (5fd871e) and the
                # number/currency-safe split points (3149651/f88de44), a
                # small gap can safely be attempted rather than written off;
                # the 1.0-unit floors below this used to compound into a
                # worst-case ~2-line leftover (this gate PLUS a category-
                # transition cost eating into the split attempt), which is
                # exactly the "some pages 2 lines short" the user reported
                # even after those other fixes landed.
                if gap < 0.2:
                    break

                head_acct = nxt_accts[0]

                # A brand-new paragraph costs AT LEAST line_h+para_gap
                # (~1.0 std_lh unit) no matter how short its text is -- the
                # wrap+cost formula charges a full line height for even a
                # single character, plus the fixed per-paragraph gap. So
                # whenever gap < ~1.0 and the next slot's head is a DIRECT
                # continuation of the account already at the end of THIS
                # slot, extending that existing paragraph's own text (which
                # only adds wrapped LINES, no extra para_gap) can use up a
                # gap the "add as a new paragraph" logic below can never
                # fill by construction -- confirmed via a real production
                # case: gap=0.641 units (8.85pt) sat permanently unfilled
                # because even a 1-character new split costs >=13.8pt
                # (1 line + the paragraph gap), a mathematical floor, not a
                # tuning knob. Try this FIRST, before the new-paragraph path.
                if (
                    cur_accts
                    and head_acct.get("is_continuation")
                    and cur_accts[-1].get("mapping_key") == head_acct.get("mapping_key")
                    and cur_accts[-1].get("original_key", cur_accts[-1].get("mapping_key"))
                        == head_acct.get("original_key", head_acct.get("mapping_key"))
                ):
                    if self._extend_continuation_in_place(
                        cur_accts, nxt_accts, cur_name, cur_shape, cur_cap, statement_type,
                    ):
                        continue

                whole_used = self._compute_slot_used_lines(
                    cur_accts + [head_acct], cur_name, slot_shape=cur_shape, statement_type=statement_type,
                )
                # Prefer keeping the account WHOLE over a tidy split, up to
                # the tail tolerance. Strict capacity here was the last place
                # still ignoring it: the team accepts 1-2 lines protruding and
                # has said so repeatedly, but this line split an account the
                # moment it exceeded capacity by any amount at all. On a real
                # deck that cut 其他非流动资产 across the page-1/page-2
                # boundary while page 1 rendered at 94% -- roughly 1.5 unused
                # lines under the split. The tolerance is the same value
                # inspect_pptx.py already declines to flag as overflow, so
                # this cannot produce a warning that was not already accepted.
                if whole_used <= cur_cap + self._tail_overflow_tolerance_units(statement_type):
                    cur_accts.append(nxt_accts.pop(0))
                    continue

                # Doesn't fit whole -- split it to use up exactly what room remains.
                cur_last_cat = str(cur_accts[-1].get("category", "") or "") if cur_accts else ""
                head_cat = str(head_acct.get("category", "") or "")
                category_gap_cost = 1.0 if (head_cat and head_cat != cur_last_cat) else 0.0
                text_budget = gap - category_gap_cost
                # Lowered from 1.0 in step with _split_commentary_at_boundary's
                # own internal floor (also lowered below) -- see the gap<0.2
                # comment above for why 1.0 here was leaving up to ~2 real
                # lines on the table in the worst case.
                if text_budget < 0.5:
                    break

                is_chinese = _account_is_chinese(head_acct)
                part1 = None
                head_text = tail_text = ""
                remaining_budget = text_budget
                # 8 attempts, and each failed candidate shrinks the budget by
                # 1.5x its overage (not 1x) -- _split_commentary_at_boundary's
                # own split-point search is a coarse chars_per_line estimate,
                # not the accurate Pillow/client-metrics measurer this loop's
                # own candidate_used check uses, so the two can disagree by a
                # small, fairly consistent margin. A 1x reduction often lands
                # on the exact same sentence/comma boundary again (nothing
                # between the old and new budget crossed a punctuation mark),
                # burning attempts without changing the answer; overshooting
                # the reduction converges in fewer tries.
                for _attempt in range(8):
                    split_result = self._split_commentary_at_boundary(
                        str(head_acct.get("commentary", "") or ""),
                        remaining_budget,
                        slot_name=cur_name,
                        is_chinese=is_chinese,
                        shape=cur_shape,
                        statement_type=statement_type,
                        # Lower than the other split passes' 0.3 -- this
                        # pass's whole point is "use up as much of the
                        # remaining gap as possible," so a smaller carved-
                        # off fragment is still worth taking here even
                        # where it wouldn't be for a balance-oriented move.
                        min_fill_ratio=0.15,
                        key_prefix=f"■ {_account_cost_key(head_acct)} - ",
                        # Also lower than the shared 1.0 default, in step
                        # with this pass's own gap<0.2/text_budget<0.5 gates
                        # above -- same reasoning: this is the last-word,
                        # maximize-fill pass, not a balance-oriented one.
                        min_available_visual=0.5,
                    )
                    if not split_result:
                        break
                    head_text, tail_text = split_result
                    candidate = head_acct.copy()
                    candidate["commentary"] = head_text
                    candidate["is_partial"] = True
                    candidate["part_num"] = int(head_acct.get("part_num") or 1)
                    candidate["original_key"] = head_acct.get("original_key", head_acct.get("mapping_key"))
                    candidate_used = self._compute_slot_used_lines(
                        cur_accts + [candidate], cur_name, slot_shape=cur_shape, statement_type=statement_type,
                    )
                    if candidate_used <= cur_cap:
                        part1 = candidate
                        break
                    overage = candidate_used - cur_cap
                    remaining_budget -= (overage * 1.5) + 0.25
                    if remaining_budget < 0.5:
                        break

                if part1 is None:
                    break  # can't split to fit either -- nothing more fits in this slot

                part2 = head_acct.copy()
                part2["commentary"] = tail_text
                part2["is_continuation"] = True
                part2["part_num"] = int(head_acct.get("part_num") or 1) + 1
                part2["original_key"] = head_acct.get("original_key", head_acct.get("mapping_key"))

                cur_accts.append(part1)
                nxt_accts[0] = part2
                logger.info(
                    "  Maximized forward fill: split '%s' to top up slot %s (gap was %.1f)",
                    head_acct.get("mapping_key", "?"), i, gap,
                )
                continue


    def _consolidate_trailing_near_empty_slot(
        self,
        assignment: List[List[Dict[str, Any]]],
        slots: List[Dict[str, Any]],
        statement_type: Optional[str],
    ) -> None:
        """User-reported case: a whole trailing slide held only a single
        leftover sentence spilled from the slide before it (e.g. one
        continuation of an already-mostly-placed account), while that
        prior slide was already packed to ~100% -- _maximize_forward_fill
        correctly refuses to grow a slot past its own capacity, so this
        sliver never had anywhere to go and sat alone on an otherwise
        blank page. Eliminating a whole near-empty trailing page is worth
        more than staying strictly under 100% on the page before it, so
        this runs LAST and explicitly accepts a small bounded overflow
        (up to 15% over nominal capacity -- well within what
        _apply_bounded_autofit's shrink, floored at _BOUNDED_AUTOFIT_MIN_
        SCALE=0.70, can visually absorb) specifically to empty this slot
        out entirely. An emptied slot's slide becomes eligible for the
        existing unused-slide removal, collapsing the whole page.

        Only the true tail of the WHOLE statement -- not every small
        slot -- since folding a small-but-legitimate MIDDLE slot forward
        would just recreate the same problem one slot earlier.
        """
        n = len(assignment)
        if n < 2:
            return
        last_i = n - 1
        while last_i > 0 and not assignment[last_i]:
            last_i -= 1
        if last_i <= 0:
            return  # nothing trails, or the whole statement is one slot

        last_slot = slots[last_i]
        last_accts = assignment[last_i]
        last_used = self._compute_slot_used_lines(
            last_accts, last_slot["slot_name"], slot_shape=last_slot["shape"], statement_type=statement_type,
        )
        # Only a genuinely tiny leftover -- not a legitimately-substantial
        # trailing slot that just happens to be under-filled because its
        # own statement's content ran out (that's expected and fine, per
        # _maximize_forward_fill's own docstring).
        if last_used > 3.0 or last_used > 0.25 * last_slot["capacity"]:
            return

        prev_i = last_i - 1
        prev_slot = slots[prev_i]
        prev_accts = assignment[prev_i]
        if not prev_accts:
            return
        combined_used = self._compute_slot_used_lines(
            prev_accts + last_accts, prev_slot["slot_name"], slot_shape=prev_slot["shape"], statement_type=statement_type,
        )
        if combined_used <= prev_slot["capacity"] * 1.15:
            assignment[prev_i] = prev_accts + last_accts
            assignment[last_i] = []
            logger.info(
                "  Consolidated trailing near-empty slot %s (%.1f units) into slot %s (now %.1f/%.1f)",
                last_i, last_used, prev_i, combined_used, prev_slot["capacity"],
            )


    def _optimize_slot_fill(
        self,
        distribution: List[tuple],
        slot_shapes: Optional[Dict[int, Any]] = None,
        slot_meta: Optional[List[Tuple[int, str]]] = None,
        statement_type: Optional[str] = None,
    ) -> List[tuple]:
        """Dynamic-programming balanced partition.

        Flattens all accounts into reading order, then partitions them into
        contiguous groups (one per slot) so that the maximum slot fill ratio
        is minimised. Line counts come from _compute_slot_used_lines measured
        against each slot's actual shape, so when Pillow fitting is enabled
        this uses real font metrics. Preserves reading order; drops trailing
        empty slots.

        DP: dp[s][i] = min achievable "max fill ratio" when placing
        accounts[0..i] into slots[0..s]. Transition: slot s takes a suffix
        accounts[j+1..i]; combine with dp[s-1][j]. O(S * N^2) states, but
        N ≤ ~20 accounts and S ≤ ~8 slots in practice, so this is trivial.
        """
        if not distribution:
            return distribution

        slot_lookup: Dict[Tuple[int, str], Any] = {}
        if slot_meta and slot_shapes:
            for slot_idx, (s_idx, s_name) in enumerate(slot_meta):
                slot_lookup[(s_idx, s_name)] = slot_shapes.get(slot_idx)

        def _resolve_shape(slide_idx: int, slot_name: str):
            shape = slot_lookup.get((slide_idx, slot_name))
            if shape is not None:
                return shape
            try:
                slide = self.presentation.slides[slide_idx]
            except Exception:
                return None
            return self._resolve_commentary_slot_shape(slide, slot_name)

        flat_accounts: List[Dict[str, Any]] = []
        for _slide_idx, _slot_name, accounts in distribution:
            flat_accounts.extend(accounts)

        if not flat_accounts:
            return distribution

        slots: List[Dict[str, Any]] = []
        is_chinese_any = any(_account_is_chinese(a) for a in flat_accounts)
        for slide_idx, slot_name, _accounts in distribution:
            shape = _resolve_shape(slide_idx, slot_name)
            capacity = self._calculate_max_lines_for_textbox(
                shape,
                is_chinese=is_chinese_any,
                slot_name=slot_name,
                statement_type=statement_type,
            )
            slots.append({
                "slide_idx": slide_idx,
                "slot_name": slot_name,
                "shape": shape,
                # Float, not int(...) -- see _calculate_max_lines_for_textbox;
                # re-flooring its already-precise return value here threw the
                # same up-to-one-line margin away a second time.
                "capacity": max(1.0, float(capacity or 1)),
            })

        N = len(flat_accounts)
        S = len(slots)

        # ── Pre-compute per-account content lines for each unique slot type ──
        # Key: (slot_name, shape_width_emu).  Two slots that share the same
        # name and width get the same measurements, so we only call
        # _calculate_content_lines (and Pillow when enabled) once per
        # (account, slot_type) pair — O(N × slot_types) total instead of the
        # O(S × N²) calls that the old range-based approach produced.
        _acct_cost: Dict[Tuple[int, str, int], float] = {}
        seen_slot_types: set = set()
        for slot in slots:
            shape = slot["shape"]
            w_key = int(shape.width) if shape and hasattr(shape, "width") else 0
            type_key = (slot["slot_name"], w_key)
            if type_key in seen_slot_types:
                continue
            seen_slot_types.add(type_key)
            for a_i, account in enumerate(flat_accounts):
                _acct_cost[(a_i, slot["slot_name"], w_key)] = self._calculate_content_lines(
                    "",
                    _account_cost_key(account),
                    account.get("commentary", ""),
                    slot_name=slot["slot_name"],
                    shape=shape,
                    is_chinese=_account_is_chinese(account),
                    statement_type=statement_type,
                )

        # ── Tight packing: use minimum slots, expand only if infeasible ─────────
        # The DP objective (min max-fill-ratio) spreads content across ALL
        # available slots at ~80% fill.  We want ~90-95%.  Fix: compute the
        # minimum number of slots that can hold all content, try that first;
        # if infeasible (split accounts can push content above S_min capacity),
        # expand by one and retry until feasible or S_orig is reached.
        _slots_all = list(slots)
        _S_orig = S

        import math as _math
        _est_sname = slots[0]["slot_name"] if slots else "single"
        _est_wkey = (
            int(slots[0]["shape"].width)
            if slots and slots[0].get("shape") and hasattr(slots[0]["shape"], "width")
            else 0
        )
        _total_est: float = 0.0
        _prev_cat_e: Optional[str] = None
        for _a_i, _acct_e in enumerate(flat_accounts):
            _cat_e = str(_acct_e.get("category", "") or "")
            if _cat_e and _cat_e != _prev_cat_e:
                _total_est += 1.0
            _prev_cat_e = _cat_e
            _total_est += _acct_cost.get((_a_i, _est_sname, _est_wkey), 2.0)
        _min_cap = min(slot["capacity"] for slot in slots) if slots else 1
        S_min = max(1, _math.ceil(_total_est / _min_cap))

        # cost_cache and slot_cost are defined before the retry loop.
        # slot_cost captures `slots` by reference — updating slots = _slots_all[:S_try]
        # inside the loop automatically changes what slot_cost sees.
        cost_cache: Dict[Tuple[int, int, int], float] = {}

        # Progressive relax factors for DP feasibility. Start at 1.0 (strict
        # capacity) and widen until the DP finds a partition. The final
        # factor (very large) guarantees feasibility, so the DP ALWAYS
        # returns a balanced result — we never fall through to greedy, which
        # would force-place oversized accounts and break the slide layout.
        # shape_height_utilization is the "natural" first relaxation because
        # PPT auto-fit can absorb that much overflow at render time.
        _packing_relax = self._packing_settings(statement_type)
        # Default 1.15 -> 1.00 (2026-08-04). This relax lets the packer
        # treat a slot as N% taller than it is, to avoid dropping content
        # or adding a page. It was safe only because the cost model used
        # to OVER-count height (a trailing paragraph gap per block), so a
        # "105% full" slot really rendered at ~100%. Now that the model
        # reproduces PowerPoint's own BoundHeight exactly, that hidden
        # buffer is gone and the same relax produced a real, measured
        # 6.6pt overflow (slide 2 L: 366.6pt of text in a 358.9pt box).
        # 1.00 means "fill the slot completely, never past it", which is
        # exactly the behaviour the user asked for after seeing slide 1
        # land at precisely 0pt spare: "可以用盡空間 比較好".
        _shape_util = float(_packing_relax.get("shape_height_utilization", 1.00) or 1.00)
        # NOTE: the second tier is a hardcoded 1.05 FLOOR, so configuring
        # shape_height_utilization BELOW 1.05 has no effect. Removing the
        # floor was tried (2026-08-05) and made things materially worse:
        # with it gone the next tier is 1.35, and a real 6-account export
        # front-loaded slide 1 to 109% while slide 2 sat at 63%/0%. The
        # 1.05 tier is what keeps a near-fit from falling all the way to
        # 1.35. Residual effect: a slot can still render ~1.05x full (a
        # real export measured -3.6pt, about a third of a line). Closing
        # that properly means making the DP add a slot instead of
        # relaxing -- i.e. the S_min expansion logic, not this ladder.
        _relax_factors: List[float] = [1.0, max(1.05, _shape_util), 1.35, 1.6, 10.0]

        # Front-loading target: slots before the LAST used one should be packed
        # to at least this fill ratio. target_fill_min_ratio existed in config
        # already (default 0.95) but was never read anywhere — this is the
        # first real consumer of it.
        _target_min_fill = float(_packing_relax.get("target_fill_min_ratio", 0.95) or 0.95)

        def slot_cost(s: int, j: int, i: int) -> float:
            """Return float line-units for placing flat_accounts[j..i] in slot s.
            Category headers cost 1.0; account content costs the float from
            _calculate_content_lines (actual_pt / std_lh, no ceil)."""
            if j > i:
                return 0.0
            key = (s, j, i)
            if key in cost_cache:
                return cost_cache[key]
            slot = slots[s]
            shape = slot["shape"]
            w_key = int(shape.width) if shape and hasattr(shape, "width") else 0
            sname = slot["slot_name"]
            used: float = 0.0
            prev_cat = None
            for a_i in range(j, i + 1):
                account = flat_accounts[a_i]
                cat = str(account.get("category", "") or "")
                if cat and cat != prev_cat:
                    used += 1.0  # category header line
                prev_cat = cat
                used += _acct_cost.get((a_i, sname, w_key), 0.0)
            # Each account was costed with its own trailing paragraph gap
            # included (whole_box=False), which is right for every account
            # except the LAST one in the slot -- that final gap renders as
            # invisible padding at the bottom of the frame. Refund exactly
            # one, once, per slot.
            if i >= j:
                used -= _real_para_gap_pt(True) / (
                    _planning_std_lh_pt(True) or 1.0
                )
            cost_cache[key] = used
            return used

        INF = float("inf")
        # Lexicographic DP state: (num_nonempty_slots, underfill_penalty).
        #
        # underfill_penalty is a TEXT-JUSTIFICATION-style cost (like the classic
        # "optimal paragraph layout" problem), NOT a load-balancing cost. Every
        # slot except the LAST one in this attempt's range is penalised for
        # falling short of _target_min_fill; the last slot is exempt (a lighter
        # final page is normal and expected — a lighter FIRST/MIDDLE page is the
        # bug this replaces). Compared as Python tuples: fewer non-empty slots
        # wins first, then lower total penalty. This front-loads content instead
        # of spreading it evenly — the previous "minimise the single worst slot"
        # objective was solved by keeping EVERY slot moderately empty (see
        # commit history: 45%/72% instead of 83%/30%), which is the opposite of
        # what a reader expects from a paginated document.
        INF_ST = (INF, INF)
        dp: List[List[Tuple[float, float]]] = []
        # split[s][i] = j such that slot s holds flat_accounts[j+1..i]; j == i
        # means slot s is empty (carries i through from previous slot).
        split: List[List[int]] = []

        # Progressive relax loop. For each factor, run the DP at full S_orig
        # slots — because the DP's own tight-packing objective will already
        # leave trailing slots empty when the content fits in fewer. The
        # final 10× factor guarantees feasibility, so we never need to fall
        # through to greedy.
        _dp_solved = False
        _solved_factor = 1.0
        S = _S_orig
        slots = _slots_all[:S]
        for _cap_mult in _relax_factors:
            cost_cache.clear()

            dp = [[INF_ST] * N for _ in range(S)]
            split = [[-1] * N for _ in range(S)]

            # Slot 0 is exempt from the underfill penalty only if it's also the
            # LAST slot in this attempt (S == 1) — a single-slot statement is
            # allowed to be light. Otherwise it must justify to the target.
            _cap0 = slots[0]["capacity"] * _cap_mult
            for i in range(N):
                lines = slot_cost(0, 0, i)
                if lines <= _cap0:
                    ratio0 = lines / slots[0]["capacity"]
                    penalty0 = 0.0 if S == 1 else max(0.0, _target_min_fill - ratio0)
                    dp[0][i] = (1.0, penalty0)
                split[0][i] = -1

            for s in range(1, S):
                cap_true = slots[s]["capacity"]
                cap_check = cap_true * _cap_mult
                is_last_slot = (s == S - 1)
                # j == -1 below means "slot s starts fresh, slots 0..s-1
                # contribute nothing" — it always compares as strictly better
                # than any real dp[s-1][j] (fewer non-empty slots wins
                # lexicographically first), so whenever it's available it's
                # ALWAYS chosen, even when an earlier slot could genuinely
                # hold content. That skipped a real coSummaryShape+table
                # slide (slot 0, smaller capacity than a plain L/R breakdown
                # slide) whenever its first content chunk was too big for
                # slot 0 alone but fit a bigger slot 1 — the DP then routed
                # everything through slot 1+, leaving slot 0 with ZERO
                # accounts. That slide was then dropped as "unused" and the
                # table-embedding code fell back to the next used slide (an
                # ordinary breakdown slide with no room for a table),
                # producing the table-overlaps-commentary bug. Only allow the
                # free (0,0) bypass when slot s-1's entire row is infeasible
                # at this relax factor — i.e. skipping ahead is the only way
                # to make progress, not a free lexicographic win over a slot
                # that could actually hold something. Real dp[s-1][j]
                # transitions (j >= 0 below) already cover every case where
                # an earlier slot legitimately holds some/all prefix content.
                prev_row_feasible = any(dp[s - 1][x][0] < INF for x in range(N))
                for i in range(N):
                    # Case A: slot s non-empty, holds accounts[j+1..i]
                    for j in range(-1, i):
                        if j < 0:
                            if prev_row_feasible:
                                continue
                            prev_state: Tuple[float, float] = (0.0, 0.0)
                        else:
                            prev_state = dp[s - 1][j]
                            if prev_state[0] >= INF:
                                continue
                        lines = slot_cost(s, j + 1, i)
                        if lines > cap_check:
                            continue
                        ratio = lines / cap_true
                        penalty = 0.0 if is_last_slot else max(0.0, _target_min_fill - ratio)
                        curr_state = (
                            prev_state[0] + 1.0,
                            prev_state[1] + penalty,
                        )
                        if curr_state <= dp[s][i]:
                            dp[s][i] = curr_state
                            split[s][i] = j
                    # Case B: slot s empty — carry dp[s-1][i] forward unchanged
                    if dp[s - 1][i] < dp[s][i]:
                        dp[s][i] = dp[s - 1][i]
                        split[s][i] = i  # marker: slot s is empty

            if dp[S - 1][N - 1][0] < INF:
                _dp_solved = True
                _solved_factor = _cap_mult
                break

            logger.info(
                "  DP infeasible at relax × %.2f; widening tolerance",
                _cap_mult,
            )

        _used_slots = int(dp[S - 1][N - 1][0]) if _dp_solved else S
        _final_penalty = dp[S - 1][N - 1][1] if _dp_solved else INF
        # logger.warning, not info: inspect_databook.py's export-log
        # analysis only surfaces WARNING/ERROR, so at INFO this line was
        # captured but never shown -- which is exactly the number needed
        # to tell "the DP thinks it packed tight but render disagrees"
        # apart from "the DP had to relax". Cheap to emit (once per
        # statement) and directly actionable, so it earns the level.
        # Also written straight to a file next to the export. Two
        # successive attempts to surface this through logging (INFO, then
        # WARNING) produced nothing on the user's machine -- something in
        # their logging setup filters it -- and this one number is what
        # decides between two opposite fixes for the residual overflow.
        # A file write can't be filtered by a handler or level.
        try:
            _msg = ("tight-pack at 1.0" if _solved_factor <= 1.0
                    else f"RELAXED to x{_solved_factor:.2f}")
            with open("dp_packing_report.txt", "a", encoding="utf-8") as _fh:
                _fh.write(
                    f"{statement_type}: {_msg} | slots used {_used_slots} of {_S_orig} "
                    f"(min {S_min}) | underfill penalty {_final_penalty:.2f} "
                    f"| target_min {_target_min_fill*100:.0f}%\n"
                )
        except Exception:
            pass

        # DEBUG, not WARNING: these were escalated during the capacity-gap
        # investigation purely so they'd survive a filtered logging setup.
        # That investigation is closed, and this console is shared, so a
        # routine packing decision must not read as a warning. The same
        # numbers still go to dp_packing_report.txt above.
        if _solved_factor > 1.0:
            logger.debug(
                "  DP feasible with relax × %.2f; using %s of %s slots, underfill penalty %.2f (target_min=%.0f%%)",
                _solved_factor, _used_slots, _S_orig, _final_penalty, _target_min_fill * 100,
            )
        else:
            logger.debug(
                "  DP tight-pack: using %s of %s slots (min=%s), underfill penalty %.2f (target_min=%.0f%%)",
                _used_slots, _S_orig, S_min, _final_penalty, _target_min_fill * 100,
            )

        # Reconstruct the assignment.
        assignment: List[List[Dict[str, Any]]] = [[] for _ in range(S)]
        i = N - 1
        for s in range(S - 1, -1, -1):
            j = split[s][i]
            if j == i:
                assignment[s] = []
                continue
            assignment[s] = list(flat_accounts[j + 1 : i + 1])
            i = j
            if i < 0:
                break

        # The DP's own result is only the starting point -- six passes
        # below mutate it, so the "tight-pack at 1.0" figure reported above
        # describes the DP, NOT what finally renders. A real export showed
        # the DP at 99% and the rendered slot at 110%; the difference is
        # necessarily introduced here. Record each pass's effect on every
        # slot's cost so the responsible one is identifiable from a single
        # run instead of by elimination.
        def _slot_fill_snapshot() -> List[Tuple[float, float]]:
            out: List[Tuple[float, float]] = []
            for _s_i, _slot in enumerate(slots):
                _used = 0.0
                _prev = None
                for _a in assignment[_s_i]:
                    _c = str(_a.get("category", "") or "")
                    if _c and _c != _prev:
                        _used += 1.0
                    _prev = _c
                    _is_chi = bool(_a.get("is_chinese"))
                    _used += self._calculate_content_lines(
                        "", _rendered_bullet_label(_a, _is_chi),
                        _a.get("commentary", ""), slot_name=_slot["slot_name"],
                        shape=_slot["shape"], is_chinese=_is_chi,
                    )
                if assignment[_s_i]:
                    _used -= _real_para_gap_pt(True) / (_planning_std_lh_pt(True) or 1.0)
                out.append((_used, float(_slot["capacity"])))
            return out

        _pass_trace: List[Tuple[str, List[Tuple[float, float]]]] = [
            ("after DP", _slot_fill_snapshot())
        ]
        for _pass_name, _pass_fn in (
            ("rebalance_lopsided_lr_pairs", self._rebalance_lopsided_lr_pairs),
            ("consolidate_tiny_stub_lr_pairs", self._consolidate_tiny_stub_lr_pairs),
            ("rebalance_underfilled_boundaries", self._rebalance_underfilled_boundaries),
            ("rebalance_overflowing_boundaries", self._rebalance_overflowing_boundaries),
            ("maximize_forward_fill", self._maximize_forward_fill),
            ("consolidate_trailing_near_empty_slot", self._consolidate_trailing_near_empty_slot),
        ):
            _pass_fn(assignment, slots, statement_type)
            _pass_trace.append((_pass_name, _slot_fill_snapshot()))

        try:
            with open("dp_packing_report.txt", "a", encoding="utf-8") as _fh:
                _fh.write(f"\n--- {statement_type}: per-pass slot fill (used/capacity) ---\n")
                _prev_snap = None
                for _name, _snap in _pass_trace:
                    _cells = []
                    for _k, (_u, _c) in enumerate(_snap):
                        _pct = (_u / _c * 100) if _c else 0.0
                        _mark = "!" if _u > _c else " "
                        _chg = ""
                        if _prev_snap and abs(_prev_snap[_k][0] - _u) > 0.01:
                            _chg = f"({_prev_snap[_k][0]:.1f}->{_u:.1f})"
                        _cells.append(f"s{_k}:{_pct:5.1f}%{_mark}{_chg}")
                    _fh.write(f"  {_name:<38} " + "  ".join(_cells) + "\n")
                    _prev_snap = _snap
                _fh.write("  ('!' = over capacity. The first pass where a '!' appears "
                          "is the one that caused the overflow.)\n")
        except Exception as _exc:
            logger.debug("Could not write per-pass fill trace: %s", _exc)

        for s_i, slot in enumerate(slots):
            lines = slot_cost(s_i, 0, -1) if not assignment[s_i] else self._compute_slot_used_lines(
                assignment[s_i],
                slot["slot_name"],
                slot_shape=slot["shape"],
                statement_type=statement_type,
            )
            logger.info(
                "  Balanced DP slot %s (%s): %s/%s lines, accts=%s",
                s_i, slot["slot_name"], lines, slot["capacity"],
                [a.get("mapping_key", "?") for a in assignment[s_i]],
            )

        # A slot with no accounts is normally dropped entirely -- e.g. an
        # unused TRAILING slide the DP decided wasn't needed at all. But an
        # L/R pair where ONE side is intentionally empty (e.g.
        # _consolidate_tiny_stub_lr_pairs folding a tiny stub into its
        # sibling) is NOT that case: the slide itself is genuinely in use
        # (its other slot has real content), so the empty slot must still
        # be included here -- otherwise slides_content never gets an entry
        # for it at all, the per-slot render loop's "if not
        # account_data_list: clear the shape" code never even runs (there's
        # nothing in the dict to iterate over), and the shape is left
        # showing whatever raw template placeholder text it shipped with
        # (confirmed via a real export: a literal "Placeholder –
        # placeholder" on the untouched R box). Dropping is still correct
        # for a slide where EVERY slot is empty (truly unused).
        slide_has_content = {
            slot["slide_idx"] for s_i, slot in enumerate(slots) if assignment[s_i]
        }
        rebuilt = [
            (slot["slide_idx"], slot["slot_name"], _merge_contd_pairs(assignment[s_i]))
            for s_i, slot in enumerate(slots)
            if assignment[s_i] or slot["slide_idx"] in slide_has_content
        ]
        return rebuilt


    def _plan_slot_distribution(
        self,
        structured_data: List[Dict],
        *,
        max_slides: int,
        start_slide: int,
        statement_type: str,
        is_chinese_databook: bool,
    ) -> List[Tuple[int, str, List[Dict]]]:
        """Decide which accounts land in which slot -- the whole plan,
        including presentation-table accounts.

        Extracted so the export and inspect_databook.py's fill diagnostic
        run the SAME planner. They did not: the diagnostic called
        _distribute_content_across_slots on every account, while the export
        pulls table accounts out of that pool first and appends them
        afterwards. On a real deck that made the diagnostic report an IS
        page as 100% full with 4 accounts when the shipped page held one
        account at 33%, and conclude "no genuine packing gaps" about a
        layout it was not describing. A diagnostic that models a different
        algorithm than the one shipping is worse than no diagnostic.

        Callers are expected to have run _prepare_structured_data_for_slides
        already, since the export needs the prepared items for other things
        too.
        """
        tables_enabled = self._presentation_tables_enabled()
        table_style = self._presentation_table_style() if tables_enabled else "table"
        table_items: List[Dict[str, Any]] = []
        normal_items: List[Dict[str, Any]] = []
        last_table_pos: Optional[int] = None
        tagged_normal: List[Tuple[int, Dict[str, Any]]] = []
        for pos, item in enumerate(structured_data):
            table = _presentation_table_for_account(item) if tables_enabled else None
            if table and table_style == "sublist":
                # Fallback style: the account is NEVER pulled out of the
                # normal packing pool -- the table dict becomes plain text
                # appended to its own commentary, so it inherits the whole
                # existing text pipeline (packing, splitting, cross-column
                # continuation, overflow handling) instead of the dedicated
                # native-table rendering path below.
                item = dict(item)
                is_chinese = bool(item.get("is_chinese"))
                lead_in, post_table_text = self._split_table_commentary(
                    item.get("commentary", ""), is_chinese,
                )
                source_multiplier = 1
                financial_data = item.get("financial_data")
                if hasattr(financial_data, "attrs"):
                    source_multiplier = financial_data.attrs.get("source_multiplier") or 1
                sublist_text = _sublist_text_for_table(table, is_chinese, source_multiplier)
                parts = [p for p in (lead_in, sublist_text, post_table_text) if p]
                item["commentary"] = "\n".join(parts)
                tagged_normal.append((pos, item))
            elif table:
                item = dict(item)
                lead_in, post_table_text = self._split_table_commentary(
                    item.get("commentary", ""), bool(item.get("is_chinese")),
                )
                item["commentary"] = lead_in
                item["_presentation_table"] = table
                item["_post_table_text"] = post_table_text
                table_items.append(item)
                last_table_pos = pos
            else:
                tagged_normal.append((pos, item))

        # Accounts positioned AFTER the LAST table account in the
        # statement's own reading order (e.g. 投资收益/营业外支出 following
        # 财务费用 on a real Crescent IS) are withheld from the normal
        # packer's input pool and instead flowed into the table slots'
        # own leftover space below -- see _append_table_accounts_to_
        # distribution's trailing_items. Excluding them BEFORE the packer
        # ever runs (rather than trying to reopen its finished output
        # afterwards) mirrors exactly how table_items themselves have
        # always been excluded from this same pool. Accounts BEFORE or
        # BETWEEN table accounts are untouched -- only the true reading-
        # order tail is eligible, so this can't reorder anything the
        # packer would otherwise have placed earlier.
        trailing_items: List[Dict[str, Any]] = []
        if last_table_pos is not None:
            for pos, item in tagged_normal:
                if pos > last_table_pos:
                    trailing_items.append(item)
                else:
                    normal_items.append(item)
        else:
            normal_items = [item for _pos, item in tagged_normal]

        # Distribute content across textbox slots based on capacity
        slot_distribution = self._distribute_content_across_slots(
            normal_items,
            max_slides=max_slides,
            start_slide=start_slide,
            statement_type=statement_type,
        )

        # One shared column-width set for every subtable in this statement,
        # computed before any of them render (see the method's docstring).
        self._precompute_uniform_table_column_widths(table_items, is_chinese_databook)

        if table_items or trailing_items:
            slot_distribution = self._append_table_accounts_to_distribution(
                table_items, slot_distribution, max_slides=max_slides, start_slide=start_slide,
                is_chinese_databook=is_chinese_databook, trailing_items=trailing_items,
            )

        return slot_distribution


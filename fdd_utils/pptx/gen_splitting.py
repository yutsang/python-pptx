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




class _SplittingMixin:
    """Cutting one account's commentary at a place a reader would accept: sentence
    and clause boundaries, jieba word boundaries, and the guards that stop a cut
    landing inside a company name or before a bare number.

    Mixed into PowerPointGenerator; `self` is the generator.
    """

    def _split_commentary_at_boundary(
        self,
        commentary: str,
        available_std_lh_units: float,
        *,
        slot_name: str,
        is_chinese: bool,
        shape=None,
        statement_type: Optional[str] = None,
        min_fill_ratio: float = 0.5,
        key_prefix: str = "",
        min_available_visual: float = 1.0,
        overflow_allowance_units: Optional[float] = None,
    ) -> Optional[Tuple[str, str]]:
        """Find a clean split point (paragraph boundary, else sentence
        boundary, else word boundary) so the head of `commentary` fits
        within `available_std_lh_units` std_lh-units. Pure/side-effect
        free version of the paragraph-fit + sentence-boundary logic the
        initial forward-fill pass uses for its one-time split, so a
        boundary rebalance can invoke it repeatedly -- the same account's
        remainder may need splitting again at a later pass, which is how
        a single account ends up spread across more than two slots.

        Returns None if nothing usable can be carved off (already fits
        whole, or every candidate split leaves a sliver below
        `min_fill_ratio` of the available space).

        `min_available_visual` is the floor below which this refuses to
        even try -- default 1.0 (roughly one real line) for
        balance-oriented callers. _maximize_forward_fill, whose whole
        purpose is squeezing out the LAST bit of a nearly-full slot,
        passes a lower value; every other caller keeps the safer default.
        """
        commentary = str(commentary or "")
        if not commentary.strip():
            return None

        chars_per_line = self._estimate_chars_per_line(
            slot_name, is_chinese, shape=shape, statement_type=statement_type,
        )
        # Both languages: matches _real_font_size_pt/_real_line_spacing (see
        # the sibling occurrence in _distribute_content_across_slots for the
        # full explanation) -- no separate Chinese 10pt/0.95 value actually
        # gets applied to a rendered paragraph.
        _lh_est = 9.0
        _std_lh_est = _lh_est + self._PARA_SPACE_AFTER
        available_visual = max(0.0, available_std_lh_units) * (_std_lh_est / _lh_est)
        if available_visual < min_available_visual:
            return None

        paragraphs = commentary.split('\n\n')
        if len(paragraphs) == 1:
            paragraphs = commentary.split('\n')

        part1_paragraphs: List[str] = []
        part1_lines_used = 0
        split_index = 0
        for para in paragraphs:
            para_lines = max(1, (len(para) + chars_per_line - 1) // chars_per_line)
            if part1_lines_used + para_lines <= available_visual:
                part1_paragraphs.append(para)
                part1_lines_used += para_lines
                split_index += 1
            else:
                break

        if part1_paragraphs and split_index == len(paragraphs):
            return None  # whole thing already fits -- caller shouldn't be splitting

        if part1_paragraphs:
            part1 = '\n\n'.join(part1_paragraphs).strip()
            part2 = '\n\n'.join(paragraphs[split_index:]).strip()
            if part1 and part2 and part1_lines_used >= available_visual * min_fill_ratio:
                return part1, part2
            return None

        # No whole paragraph fits by the coarse ceil-per-line estimate above --
        # try a sentence boundary within paragraph 0 using a finer character
        # budget instead. The two estimates can disagree (ceil-per-line rounds
        # up to a whole line even for a small fractional overage) -- that's
        # expected, not an error, so don't bail out just because this looser
        # measure says the paragraph would fit; let the boundary search below
        # decide, and downstream capacity re-validation is the real safety net.
        para = paragraphs[0]
        chars_available = int(max(1, available_visual * chars_per_line))
        hard_cap = min(len(para), int(chars_available * 1.05))
        min_fill = max(1, int(chars_available * min_fill_ratio))
        end_positions: List[int] = []
        for end_char in ['. ', '。', '! ', '！', '? ', '？']:
            start = 0
            while True:
                pos = para.find(end_char, start, hard_cap)
                if pos < 0:
                    break
                end_positions.append(pos + len(end_char))
                start = pos + 1
        # Exclude a candidate sitting at (or past) the paragraph's own end --
        # when the finer char-budget is generous enough to reach the last
        # sentence in the paragraph, picking it would leave nothing for the
        # tail, i.e. no real split at all.
        candidates = [p for p in end_positions if min_fill <= p < len(para)]
        best_split = max(candidates) if candidates else None

        if best_split is None:
            word_end = para.rfind(' ', 0, min(hard_cap, len(para) - 1))
            if word_end > 0 and word_end >= min_fill:
                best_split = word_end + 1

        if best_split is None and is_chinese:
            # Chinese has no spaces between words, so the Latin word-boundary
            # fallback above always misses on CJK text -- a long single
            # sentence with no "." within budget (e.g. one continuous
            # amount-heavy clause well past 72 chars before its first "。")
            # had NO fallback at all here and always returned None, silently
            # giving up on pulling any of it forward even with real budget
            # to spare. Comma is the natural next-best CJK clause boundary.
            #
            # The line below used to end "...if even that isn't present in
            # range, a hard character cut is fine for CJK (unlike Latin text,
            # cutting between two Chinese characters doesn't break a word)".
            # A real deck disproved that: it cut
            #   "...主要为应付某某系统" | "工程第四建设有限公司..."
            # straight through a company name. jieba can't rescue it either --
            # 系统|工程 IS a legitimate token boundary, so _snap_split_before_
            # number sees nothing wrong. CJK absolutely does have words; it
            # just doesn't mark them with spaces.
            _cut = min(hard_cap, len(para) - 1)
            comma_end = max(
                para.rfind('，', 0, _cut), para.rfind(',', 0, _cut),
                para.rfind('；', 0, _cut), para.rfind(';', 0, _cut),
                para.rfind('、', 0, _cut),
            )
            if comma_end > 0:
                # Accept a clause boundary even BELOW min_fill: it still
                # beats a cut through the middle of a name.
                #
                # An earlier version returned None here instead, on the
                # theory that "no split, let the box protrude" was the
                # project team's stated preference. Measured, that made
                # things WORSE (per-slot fill 96.9% -> 93.6%): returning
                # None doesn't make the account protrude, it makes the
                # caller decline to pull it in at all, so the slot is left
                # emptier than before -- the reported "放少了1-2行". The
                # tolerance only helps where content STAYS PUT, which is
                # _rebalance_overflowing_boundaries, not here.
                best_split = comma_end + 1
            else:
                hard_end = _cut
                if hard_end >= min_fill:
                    best_split = hard_end

        if not best_split:
            return None

        # Validate against the REAL measurer, not the crude chars_per_line
        # estimate used to find best_split above -- the two can disagree,
        # and because wrapped-line cost is quantized (a candidate a few
        # characters shorter often still wraps to the SAME line count),
        # _maximize_forward_fill's retry loop would see this function
        # propose the exact same "still too big" candidate across several
        # attempts (a small budget cut rarely crosses the crude estimate's
        # own boundary even when the real candidate measurably overshoots),
        # then give up once its overage-scaled reduction blew past the
        # retry floor -- confirmed via a real production case: a 0.44-unit
        # (well under half a line) overage was enough to end the whole
        # retry after a single attempt, discarding a boundary that
        # genuinely had 1.3 real units of room. Back off through the same
        # candidate boundaries already found above (largest to smallest),
        # accepting the first whose ACCURATE wrapped cost actually fits.
        try:
            from fdd_utils.text_metrics import get_measurer, text_box_from_shape
            if shape is not None:
                packing = self._packing_settings(statement_type)
                family = _measurer_family(is_chinese, packing)
                font_size_pt = _real_font_size_pt(is_chinese)
                line_spacing = _real_line_spacing(is_chinese)
                _mpath = _resolve_font_metrics_path(is_chinese, packing)
                measurer = get_measurer(
                    family, font_size_pt, is_cjk=is_chinese, line_spacing=line_spacing,
                    metrics_path=_mpath,
                )
                box = text_box_from_shape(shape)
                line_h = measurer.line_height_pt()
                para_gap = _real_para_gap_pt(is_chinese)
                std_lh = line_h + para_gap
                budget_pt = max(0.0, available_std_lh_units) * std_lh
                min_fill_pt = budget_pt * min_fill_ratio

                # Every candidate boundary found above, largest first --
                # sentence ends, then the word/comma/hard-cut fallback
                # position, each de-duplicated and capped at best_split
                # (never consider a candidate BIGGER than the crude
                # estimate already chose).
                backoff_candidates = sorted(
                    {p for p in candidates if p <= best_split} | {best_split}, reverse=True,
                )
                accepted = None
                for cand_pos in backoff_candidates:
                    cand_head = para[:cand_pos].strip()
                    if not cand_head:
                        continue
                    # key_prefix mirrors _calculate_content_lines' own
                    # "■ {mapping_key} - " prepended to an account's first
                    # paragraph -- omitting it here (the original bug) made
                    # every candidate measure smaller than what the caller's
                    # actual _compute_slot_used_lines check would find, so
                    # a "fixed" head that trimmed real characters off could
                    # still fail the caller's check by the exact same margin.
                    wrapped = measurer.wrap(
                        key_prefix + cand_head,
                        max(10.0, box.width_pt - self._BULLET_HANGING_INDENT_PT),
                        first_line_width_pt=box.width_pt,
                    )
                    cand_pt = len(wrapped) * line_h + para_gap
                    if cand_pt <= budget_pt and cand_pt >= min_fill_pt:
                        accepted = cand_pos
                        break
                if accepted is not None and accepted != best_split:
                    best_split = accepted
                elif accepted is None:
                    # Nothing in the existing candidate list fits even
                    # accurately-measured -- fall back to trimming the
                    # crude choice down word-by-word (Latin) or char-by-
                    # char (CJK) until it does, rather than giving up
                    # outright on a boundary that may still have room. This
                    # is the last-resort branch (no sentence/comma boundary
                    # produced anything accurately-fitting either) -- accept
                    # the first fit regardless of min_fill_ratio here, since
                    # by construction it's the LARGEST candidate that fits
                    # (we're trimming down from best_split), so nothing
                    # smaller would meet min_fill_ratio either; a small-but-
                    # real split still beats the caller getting nothing back
                    # and abandoning a boundary that had real room to spare.
                    trimmed = best_split
                    while trimmed > 0:
                        cand_head = para[:trimmed].strip()
                        if cand_head:
                            wrapped = measurer.wrap(
                                key_prefix + cand_head,
                                max(10.0, box.width_pt - self._BULLET_HANGING_INDENT_PT),
                                first_line_width_pt=box.width_pt,
                            )
                            cand_pt = len(wrapped) * line_h + para_gap
                            if cand_pt <= budget_pt:
                                best_split = trimmed
                                break
                        trimmed -= max(1, len(cand_head) // 10) if cand_head else 5
                    else:
                        return None  # nothing, down to zero, ever fit

                # Extend FORWARD past whatever best_split the crude search
                # above settled on, up to the TRUE available budget. The
                # crude per-character estimate that located candidates/
                # hard_cap above (_AVG_CHAR_WIDTH_CHI/_WORD_WRAP_SLACK)
                # systematically under-counts real Chinese glyph density
                # (~10.9pt/char effective vs the real ~9pt at this font
                # size), and that same 8% "word-wrap slack" margin is
                # justified for Latin text (room for a whole word that
                # might not fit) but doesn't apply to CJK at all (every
                # character is its own valid break point) -- so even the
                # "largest candidate that fits" from the search above
                # routinely stops well short of the box's true capacity.
                # This is exactly the "沒有用盡那一行才cut" a user report
                # confirmed by hand-counting real rendered lines against
                # this file's own prediction: a whole further line of real,
                # usable room was never even considered as a candidate.
                # Binary-search directly against the accurate measurer (not
                # the crude heuristic) for the true maximum that fits, then
                # prefer snapping to the nearest natural sentence/comma
                # boundary AT OR BEFORE that true maximum so the cut still
                # reads naturally wherever that costs nothing.
                lo, hi = best_split, len(para)
                true_max = best_split
                while lo <= hi:
                    mid = (lo + hi) // 2
                    cand_head = para[:mid].strip()
                    if not cand_head:
                        lo = mid + 1
                        continue
                    wrapped = measurer.wrap(
                        key_prefix + cand_head,
                        max(10.0, box.width_pt - self._BULLET_HANGING_INDENT_PT),
                        first_line_width_pt=box.width_pt,
                    )
                    cand_pt = len(wrapped) * line_h + para_gap
                    if cand_pt <= budget_pt:
                        true_max = mid
                        lo = mid + 1
                    else:
                        hi = mid - 1

                # Choose the split point that best FILLS the budget, out of
                # every real sentence/clause boundary within reach.
                #
                # The old rule -- "take any boundary that fits" -- silently
                # wasted space: measured on the reviewed deck's own 应付账款
                # text, a 1.5-line budget stopped at a 1.00-line boundary and
                # left half a line empty, because the next boundary needed
                # 1.78 lines, i.e. 0.28 over. That is nowhere near the ~2
                # lines of protrusion the project team accepts, and it is the
                # reported "能夠再放1-2行才滿". Callers re-validate with the
                # same tolerance (see _try_partial_split_into_gap), so an
                # overshoot chosen here survives rather than being trimmed.
                # Callers that already widened their own budget by the
                # protrusion allowance pass 0 here; adding it a second time
                # let the search run to capacity + 2x the allowance, which
                # overshot every real boundary and fell back to a raw
                # character cut (measured: one appeared the moment
                # _try_partial_split_into_gap started pre-inflating its gap).
                _allow = (self._tail_overflow_tolerance_units(statement_type)
                          if overflow_allowance_units is None
                          else float(overflow_allowance_units))
                tol_pt = _allow * (line_h + para_gap)

                def _head_pt(pos: int) -> Optional[float]:
                    head_txt = para[:pos].strip()
                    if not head_txt:
                        return None
                    return len(measurer.wrap(
                        key_prefix + head_txt,
                        max(10.0, box.width_pt - self._BULLET_HANGING_INDENT_PT),
                        first_line_width_pt=box.width_pt,
                    )) * line_h + para_gap

                # Largest position still within budget+tolerance, so the
                # boundary scan below never has to look further than useful.
                lo2, hi2, tol_max = best_split, len(para) - 1, best_split
                while lo2 <= hi2:
                    mid2 = (lo2 + hi2) // 2
                    cand_pt = _head_pt(mid2)
                    if cand_pt is None:
                        lo2 = mid2 + 1
                        continue
                    if cand_pt <= budget_pt + tol_pt:
                        tol_max = mid2
                        lo2 = mid2 + 1
                    else:
                        hi2 = mid2 - 1

                positions = set()
                for end_char in ('. ', '。', '! ', '！', '? ', '？',
                                 '，', ',', '；', ';', '、'):
                    scan = min_fill
                    while True:
                        found = para.find(end_char, scan, tol_max + 1)
                        if found < 0:
                            break
                        cand = found + len(end_char)
                        if min_fill <= cand < len(para):
                            positions.add(cand)
                        scan = found + 1

                best_under = best_over = None
                under_pt = over_pt = 0.0
                for cand in sorted(positions):
                    cand_pt = _head_pt(cand)
                    if cand_pt is None:
                        continue
                    if cand_pt <= budget_pt:
                        if best_under is None or cand > best_under:
                            best_under, under_pt = cand, cand_pt
                    elif cand_pt <= budget_pt + tol_pt:
                        if best_over is None or cand < best_over:
                            best_over, over_pt = cand, cand_pt
                # Only spend the protrusion when staying inside the budget
                # would actually waste something -- half a line is the point
                # at which the gap is visible in a rendered deck.
                _WASTE_PT = 0.5 * (line_h + para_gap)
                if best_under is not None and (budget_pt - under_pt) < _WASTE_PT:
                    best_split = best_under
                elif best_over is not None:
                    best_split = best_over
                elif best_under is not None:
                    best_split = best_under
                elif true_max > best_split:
                    best_split = true_max
        except Exception:
            pass  # measurer unavailable -- fall through with the crude best_split

        # Never split inside a number -- a real production case did exactly
        # this ("...室外工程人民币2," at the end of one slide, "818.7万元..."
        # continuing the next): the hard-cut/character-trim fallbacks above
        # have no concept of a numeric literal, so a raw character offset
        # can land between the "2" and the "8" of "2,818.7". Backing up
        # only ever REMOVES characters from head (never adds), so whatever
        # capacity check already accepted this position stays satisfied.
        best_split = self._snap_split_before_number(para, best_split)
        best_split = self._snap_before_dangling_connective(para, best_split, min_fill)

        head = para[:best_split].strip()
        tail_rest = para[best_split:].strip()
        tail = (tail_rest + '\n\n' + '\n\n'.join(paragraphs[1:])).strip() if len(paragraphs) > 1 else tail_rest
        if not head or not tail:
            return None
        return head, tail


    # Tokens that cannot end a column: each one is grammatically waiting for
    # something that is now in the next column. The number guard above keeps a
    # figure whole but says nothing about what is left BEHIND it, so a real
    # deck ended a column on "…长期待摊费用余额为" with "355.8万元…" resuming
    # overleaf -- the amount was intact and the sentence was still broken.
    # Longest first: the scan takes the first match, and "主要为" must win
    # over "为".
    _DANGLING_TAIL_TOKENS = (
        "主要包括", "主要系", "主要为", "分别为", "合计为", "约为", "增至", "降至", "升至",
        "包括", "共计", "计为", "达到", "人民币",
        "为", "系", "是", "约", "达", "及", "和", "与", "或", "的", "至", "从", "由", "在", "于",
    )
    _CLAUSE_BOUNDARY_CHARS = "，,；;。、）)"

    @classmethod
    def _snap_before_dangling_connective(cls, text: str, pos: int, min_fill: int) -> Optional[int]:
        """Back `pos` off a split that would leave the head grammatically
        hanging, or refuse the split entirely.

        Retreats only when a full sentence end is available to retreat TO, and
        otherwise leaves the split where it was. It deliberately never refuses:
        asked to choose, the user took a half sentence over a blank line, so
        this may improve a break but must never cost a line of fill.
        """
        if pos <= 0 or pos > len(text):
            return pos
        head = text[:pos].rstrip()
        if not head:
            return pos
        if not any(head.endswith(token) for token in cls._DANGLING_TAIL_TOKENS):
            return pos
        # Once the head is hanging, the only retreat worth taking is to a full
        # SENTENCE end. Backing up to the nearest comma just trades one
        # fragment for another -- on the real case it produced a bare
        # "截至<date>，", an adverbial with no clause attached, which reads no
        # better than the "…余额为" it replaced.
        #
        # No sentence end in range means the split stays exactly where it was.
        # Refusing here was tried and is wrong for this deck: it empties the
        # line instead of filling it, and the user's explicit preference is a
        # half sentence over a blank line.
        cut = max(head.rfind(ch) for ch in "。！？!?")
        if cut <= 0:
            return pos
        candidate = cut + 1
        return candidate if candidate >= min_fill else pos

    @classmethod
    def _snap_before_org_name(cls, text: str, pos: int) -> int:
        """Back `pos` up to before an organisation name it would cut into.

        Fires only when the text between `pos` and an org-name tail ahead of
        it contains no punctuation -- i.e. pos really is inside one name, not
        merely before a sentence that happens to mention a company. Backs up
        to just after the previous punctuation (where the name can only have
        started), and declines when that would throw away most of the head:
        losing the whole slot is a worse outcome than the cut it avoids.
        Only ever moves `pos` earlier, so it can't invalidate a capacity
        check that already passed.
        """
        if pos <= 0 or pos >= len(text):
            return pos
        window = text[pos:pos + 24]
        hits = [window.find(s) for s in cls._ORG_NAME_TAILS]
        hits = [h for h in hits if h >= 0]
        if not hits:
            return pos
        # Punctuation before the tail means pos isn't inside that name.
        if any(ch in cls._SENTENCE_PUNCT for ch in window[:min(hits)]):
            return pos
        prev = max((text.rfind(ch, 0, pos) for ch in cls._SENTENCE_PUNCT), default=-1)
        if prev < 0:
            return pos
        snapped = prev + 1
        return snapped if snapped >= pos * 0.3 else pos


    @classmethod
    def _snap_split_before_number(cls, text: str, pos: int) -> int:
        """If `pos` sits inside a numeric literal (digits, thousands commas,
        decimal point), back up to the start of that literal so the whole
        number stays together on one side of the split. Also backs up past
        an immediately-preceding currency marker (CNY/RMB/人民币/$/¥/...) so
        a bare currency symbol doesn't get stranded with its amount on the
        far side of the split, past any word jieba's segmentation says
        `pos` would split in half (see _jieba_word_boundary_snap), and past
        any _PROTECTED_CJK_COMPOUNDS entry as a fallback/supplement when
        jieba isn't installed or doesn't treat it as one token. Only ever
        moves pos earlier, so it can't undo a capacity check that already
        accepted this position."""
        if pos <= 0 or pos >= len(text):
            return pos
        # Organisation names first: this is the one case jieba actively
        # MASKS rather than merely misses, because every internal boundary of
        # a company name is a real word boundary to it. Applied before the
        # other snaps (and not at the end) because this function has several
        # early returns; each snap only ever moves pos earlier, so running
        # this first is safe and covers every exit path.
        pos = cls._snap_before_org_name(text, pos)
        if pos <= 0:
            return pos
        numeric_chars = set('0123456789,.')
        jieba_snap = _jieba_word_boundary_snap(text, pos)
        if jieba_snap is not None:
            pos = jieba_snap
            if pos <= 0:
                return pos
        # Units glued directly to the number immediately before them --
        # magnitude units (784万元) and date suffixes (2024年/12月/31日)
        # are the same shape of bug: jieba correctly tokenizes the number
        # and the unit as SEPARATE tokens (confirmed: jieba.cut gives
        # ['784', '万元'] and ['31', '日']), so once jieba has already
        # snapped `pos` to that clean token boundary, the straddle-scan
        # below (which only fires when `pos` sits INSIDE the unit itself)
        # never gets a chance to also pull the number back -- jieba
        # correctly stops a mid-unit split but, on its own, still leaves
        # the number stranded from its own unit one token earlier. Two
        # real production cases hit this: "784万" | "元..." before jieba
        # existed, and "...2024年12月31" | "日及2025年12月31日..." AFTER
        # jieba was added (jieba's OWN snap masked the marker loop's
        # number-pullback special case below, since by the time that loop
        # runs, `pos` no longer straddles "万元" -- it sits cleanly right
        # before it). Checking here, unconditionally, right after the
        # jieba snap and before the straddle-scan, closes both cases.
        for _suffix in ('万元', '亿元', '年', '月', '日'):
            if text[pos:pos + len(_suffix)] == _suffix and text[pos - 1] in numeric_chars:
                num_start = pos
                while num_start > 0 and text[num_start - 1] in numeric_chars:
                    num_start -= 1
                if num_start < pos:
                    return num_start
        # `pos` landing strictly INSIDE a multi-character marker/compound
        # itself (e.g. between "人民" and "币", or "分" and "别") is a
        # different failure mode from the number/marker-boundary case
        # below -- neither text[pos-1] nor text[pos] is a digit, so the
        # numeric-literal check never fires and this position was never
        # being checked at all.
        for marker in cls._CURRENCY_MARKERS + cls._PROTECTED_CJK_COMPOUNDS:
            if len(marker) < 2:
                continue  # single-char symbols ($/¥/£/€) can't be split mid-marker
            # Check every position `start` such that the marker, if present
            # there, would straddle pos (start < pos < start+len(marker)).
            # str.find's end argument is EXCLUSIVE, so searching up to `pos`
            # can never match an occurrence whose last character IS at pos
            # -- exactly the case being guarded against -- hence the direct
            # index scan instead.
            for start in range(max(0, pos - len(marker) + 1), pos):
                if text[start:start + len(marker)] == marker:
                    # 万元/亿元 are magnitude units directly attached to the
                    # number before them ("784万元") -- stranding the bare
                    # number from its own unit reads almost as broken as
                    # splitting the unit itself would. Back up further past
                    # any digits/comma/decimal-point run immediately before
                    # the unit, same treatment as a currency marker glued
                    # to its amount below.
                    if marker in ('万元', '亿元'):
                        num_start = start
                        while num_start > 0 and text[num_start - 1] in numeric_chars:
                            num_start -= 1
                        if num_start < start:
                            return num_start
                    return start
        if text[pos - 1] in numeric_chars and text[pos] in numeric_chars:
            start = pos
            while start > 0 and text[start - 1] in numeric_chars:
                start -= 1
            pos = start
        # pos is now either unchanged (wasn't mid-number) or at the start of
        # a numeric literal -- either way, check for a currency marker (and
        # an optional single space) immediately before it, and back up past
        # that too if found.
        if pos < len(text) and text[pos] in set('0123456789'):
            for marker in cls._CURRENCY_MARKERS:
                candidate_start = pos - len(marker)
                if candidate_start >= 0 and text[candidate_start:pos] == marker:
                    return candidate_start
                candidate_start_sp = pos - len(marker) - 1
                if candidate_start_sp >= 0 and text[candidate_start_sp:pos] == marker + ' ':
                    return candidate_start_sp
        return pos


    def _split_table_commentary(self, commentary: str, is_chinese: bool) -> Tuple[str, str]:
        """Splits an account's raw commentary into (lead_in, post_table_text)
        at the handoff phrase ai.py's _detail_table_guidance asks the model
        to end its short lead-in with. Everything up to and including that
        phrase is the lead-in (rendered above the table); anything after it
        is the optional "-"/"➢" explanatory bullets a real deliverable puts
        BELOW the table (provider, charging basis, contract terms) -- a real
        photo comparison against the project team's own Crescent deck showed
        this second part entirely missing before: the whole commentary was
        being treated as lead-in and hard-capped at ~220/340 chars, silently
        dropping any such bullets the model had written past that point with
        nowhere for them to ever render.

        Each side is truncated separately and more generously than the old
        single whole-commentary cap, since each now has its own dedicated
        vertical space (lead-in above the table, explanation below it) --
        see _presentation_table_extra_text_height_pt, which sizes both.

        Falls back to (whole commentary, '') if the model didn't include the
        handoff phrase -- exactly the old behaviour, so a non-compliant
        generation degrades to "no post-table text" rather than misplacing
        content."""
        handoff = self._TABLE_HANDOFF_CHI if is_chinese else self._TABLE_HANDOFF_ENG
        text = (commentary or "").strip()
        idx = text.lower().find(handoff.lower())
        if idx < 0:
            lead_limit = 220 if is_chinese else 340
            return _truncate_text_at_boundary(text, lead_limit, is_chinese), ""
        split_at = idx + len(handoff)
        lead_in = text[:split_at].strip()
        post_table = text[split_at:].strip()
        lead_limit = 140 if is_chinese else 220
        # The real, confirmed-fitting Crescent example (营业成本, the TALLEST
        # of the 4 known tables -- 15 rows) carried ~250 chars of "-"/"➢"
        # explanation. This cap allows meaningfully more room than that
        # observed-good case without being open-ended -- a table already
        # near the tallest slot capacity plus an unbounded explanation is
        # the one real overflow risk this feature has, since it's genuinely
        # new vertical space that wasn't being claimed before.
        post_limit = 450 if is_chinese else 700
        return (
            _truncate_text_at_boundary(lead_in, lead_limit, is_chinese),
            _truncate_text_at_boundary(post_table, post_limit, is_chinese),
        )


    def _drop_orphan_trailing_punctuation(self, key_prefix: str, text: str, is_chinese: bool) -> str:
        """Delete a sentence-ending mark that is the only thing pushing a
        paragraph onto one more line.

        A lone "。" starting a line is wrong in Chinese typography and the
        project team will not accept it. The usual remedies are unavailable
        here: shrinking the font is explicitly ruled out, and the kinsoku
        controls PowerPoint should honour (eaLnBrk, hangingPunct, the run's
        own lang) were all set correctly across three attempts and changed
        nothing in the real render. So take the remedy the team named
        themselves -- "直接刪除句號 他不是那麼重要".

        Measured the way POWERPOINT breaks, not the way this repo does. Our
        own wrapper already implements kinsoku, so it hangs the mark on the
        previous line for free and a line-count comparison here can never
        detect the problem -- the first version of this check was written
        that way and provably never fired. Instead: wrap without the mark,
        then ask whether the mark still fits on the resulting last line. If
        it does not, PowerPoint (which does no hanging) puts it alone on the
        next line, and it goes.
        """
        from fdd_utils.text_metrics import get_measurer, text_box_from_shape

        stripped = (text or "").rstrip()
        if not stripped or stripped[-1] not in self._ORPHANABLE_END_PUNCT:
            return text
        shape = self._measurement_slot_shape()
        box = text_box_from_shape(shape)
        if box is None or box.width_pt <= 0:
            return text
        packing = self._packing_settings()
        measurer = get_measurer(
            _measurer_family(is_chinese, packing),
            _real_font_size_pt(is_chinese), is_cjk=is_chinese,
            line_spacing=_real_line_spacing(is_chinese),
            metrics_path=_resolve_font_metrics_path(is_chinese, packing),
        )
        body_width = max(1.0, box.width_pt - self._BULLET_HANGING_INDENT_PT)
        lines = measurer.wrap(
            key_prefix + stripped[:-1], body_width, first_line_width_pt=box.width_pt,
        )
        if not lines:
            return text
        if measurer.text_width_pt(lines[-1] + stripped[-1]) > body_width:
            return stripped[:-1]
        return text


    def _split_single_into_lr(self, slide, source_shape):
        """Clone a full-width commentary box into two half-width boxes side by
        side (named textMainBullets_L / textMainBullets_R), replacing the
        original. Used when a page was assigned two logical slots (L, R) but
        the underlying template slide only has ONE commentary text box —
        without this, both slots resolve to the SAME shape (see
        _resolve_commentary_slot_shape), so L-content and R-content silently
        collide into one full-width box instead of sitting side by side like
        the BS pages. Idempotent per slide (cached by slide element id)."""
        cache = getattr(self, "_split_lr_cache", None)
        if cache is None:
            cache = self._split_lr_cache = {}
        slide_key = id(slide._element)
        cached = cache.get(slide_key)
        if cached:
            return cached

        from copy import deepcopy
        orig_left = int(source_shape.left)
        orig_width = int(source_shape.width)
        gutter = max(0, int(orig_width * 0.03))
        half_width = max(1, (orig_width - gutter) // 2)

        # Left half = the original shape, resized in place.
        source_shape.left = orig_left
        source_shape.width = half_width
        try:
            source_shape.name = "textMainBullets_L"
        except Exception:
            pass

        # Right half = a deep-copied XML clone, repositioned and renamed. The
        # clone starts with a copy of the original's text — clear it so the
        # packer fills it fresh rather than duplicating whatever was there.
        new_element = deepcopy(source_shape._element)
        source_shape._element.addnext(new_element)
        right_shape = None
        for shape in slide.shapes:
            if shape._element is new_element:
                right_shape = shape
                break
        if right_shape is None:
            # Clone insertion failed for some reason — fall back to treating
            # the (now half-width) original as both slots rather than crashing.
            cache[slide_key] = (source_shape, source_shape)
            return cache[slide_key]

        right_shape.left = orig_left + half_width + gutter
        right_shape.width = half_width
        try:
            right_shape.name = "textMainBullets_R"
        except Exception:
            pass
        if getattr(right_shape, "has_text_frame", False):
            right_shape.text_frame.clear()

        cache[slide_key] = (source_shape, right_shape)
        logger.info(
            "Split single full-width commentary box into L/R halves on a slide "
            "(page was assigned two content slots but the template only had one box)."
        )
        return cache[slide_key]


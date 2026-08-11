from __future__ import annotations

# PowerPointGenerator is assembled from these; see each module's docstring.
from .gen_measurement import _MeasurementMixin
from .gen_packing import _PackingMixin
from .gen_splitting import _SplittingMixin
from .gen_tables import _TablesMixin
from .gen_summary import _SummaryMixin

from . import helpers as _helpers
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


class PowerPointGenerator(
    _MeasurementMixin,
    _PackingMixin,
    _SplittingMixin,
    _TablesMixin,
    _SummaryMixin,
):
    """Main PowerPoint generator class"""

    def __init__(
        self,
        template_path: str,
        language: str = 'english',
        row_limit: int = 20,
        model_type: Optional[str] = None,
        model_name: Optional[str] = None,
    ):
        self.template_path = template_path
        self.language = language.lower()
        self.row_limit = row_limit
        self.model_type = str(model_type or "").strip() or None
        self.model_name = str(model_name or "").strip() or None
        self.presentation = None
        self.pptx_settings = _load_pptx_settings()

    def load_template(self):
        """Load the PowerPoint template"""
        if not os.path.exists(self.template_path):
            raise FileNotFoundError(f"Template not found: {self.template_path}")

        self.presentation = Presentation(self.template_path)
        logger.info("Loaded template: %s", self.template_path)

    # ---- attribute access kept for callers outside this package ----
    # These four moved to helpers.py because they need no instance state, but
    # tooling outside fdd_utils/pptx calls them ON the generator --
    # inspect_databook.py's population diagnostic and
    # ad-hoc/pptx-probes/diagnose_presentation_tables.py. Deleting the
    # attribute broke that (AttributeError at section 10, after a paid-for AI
    # run had already completed). The helper stays the single implementation;
    # these only keep `gen.<name>(...)` resolving.

    def find_shape_by_name(self, shapes, name: str):
        return _helpers.find_shape_by_name(shapes, name)

    def _prepare_structured_data_for_slides(self, structured_data):
        return _helpers._prepare_structured_data_for_slides(structured_data)

    def _presentation_table_for_account(self, account_data):
        return _helpers._presentation_table_for_account(account_data)

    def _expand_commentary_to_cover_summary(self, slide):
        return _helpers._expand_commentary_to_cover_summary(slide)




    def _resolve_commentary_slot_shape(self, slide, slot_name: str, used_shape_ids=None):
        """Resolve the best text box for a commentary slot on a slide."""
        used_shape_ids = used_shape_ids or set()
        preferred_names = {
            "single": [
                "textMainBullets",
                "textMainBullets_L",
                "textMainBullets_R",
                "Text-commentary",
                "Content",
                "MainContent",
                "Body",
            ],
            "L": [
                "textMainBullets_L",
                "Text-commentary",
                "Content",
                "MainContent",
                "Body",
            ],
            "R": [
                "textMainBullets_R",
                "Text-commentary",
                "Content",
                "MainContent",
                "Body",
            ],
        }

        for name in preferred_names.get(slot_name, []):
            shape = find_shape_by_name(slide.shapes, name)
            if shape and getattr(shape, "has_text_frame", False) and id(shape) not in used_shape_ids:
                return shape

        # No dedicated _L/_R box. Only fall back to a single full-width box for
        # an ACTUAL single-column slot; for "L"/"R" that would make both slots
        # resolve to the same physical shape (content collision -> renders as
        # one full-width box instead of two side-by-side columns). Split it
        # into two half-width boxes instead, mirroring the BS page layout.
        if slot_name in ("L", "R"):
            single_shape = find_shape_by_name(slide.shapes, "textMainBullets")
            if single_shape and getattr(single_shape, "has_text_frame", False) and id(single_shape) not in used_shape_ids:
                left_shape, right_shape = self._split_single_into_lr(slide, single_shape)
                return left_shape if slot_name == "L" else right_shape

        generic_candidates = [
            shape for shape in slide.shapes
            if _is_commentary_text_shape(shape) and id(shape) not in used_shape_ids
        ]
        if not generic_candidates:
            return None

        if slot_name == "L":
            return min(generic_candidates, key=lambda shape: (getattr(shape, "left", 0), -getattr(shape, "width", 0)))
        if slot_name == "R":
            return max(generic_candidates, key=lambda shape: (getattr(shape, "left", 0), getattr(shape, "width", 0)))
        return max(generic_candidates, key=lambda shape: (getattr(shape, "width", 0), -getattr(shape, "left", 0)))





    # How far an account may protrude BELOW its box, in std_lh line-units,
    # when the only alternative is splitting it across slots. The project
    # team explicitly accepts 1-2 lines sticking out; they do not accept a
    # split landing mid-name. Applied only at the split decision, never as a
    # general capacity increase -- every slot is still packed to its real
    # capacity, and a slot that simply has more content than fits still
    # splits as before.
    _TAIL_OVERFLOW_TOLERANCE_UNITS = 2.0










    # Average rendered character width (pt) for the fonts we use.
    # English: Arial 9pt mixed text ≈ 5.0 pt/char (incl. spaces & punctuation).
    # A small word-wrap slack (≈8 %) is subtracted because lines always break
    # at a word/character boundary, not at the exact pixel edge.
    #
    # Chinese was 10.0, on the reasoning that a CJK glyph is square so one em
    # at 10pt is 10pt wide. Two things were wrong with that: the commentary
    # renders at 9pt, not 10, and financial prose is not pure CJK -- every
    # amount, date and ASCII marker in it is a HALF-width glyph. Measured with
    # the shipped client metrics over six real commentary sentences: 7.52
    # pt/char weighted, ranging 6.6 (an amount-heavy multi-period sentence) to
    # 8.6. At 10.0 a 4.78in column was estimated to hold 30 characters where it
    # really holds 42, so every split candidate the search proposed was about
    # a dozen characters short of the line -- which is what left a column
    # ending on "…余额为" with a third of its last line empty.
    #
    # 8.0 rather than the measured 7.52: this only sizes the SEARCH, and the
    # candidate it proposes is then validated against the real measurer, so
    # a slightly conservative value costs a little fill while an over-generous
    # one costs retry rounds. Note this constant no longer affects capacity in
    # production at all -- _calculate_content_lines uses it only in its
    # no-font-available fallback, and the real metrics load in every run.
    _AVG_CHAR_WIDTH_ENG = 5.0
    _AVG_CHAR_WIDTH_CHI = 8.0
    _WORD_WRAP_SLACK    = 0.92   # use 92 % of the theoretical line width














    # -- Presentation-detail tables (per-account report-ready breakdowns) --
    # workbook.py's extract_presentation_detail_table finds a small,
    # human-readable breakdown table some account sheets carry alongside
    # their GL-coded main schedule (e.g. 管理费用's 会计服务费/审计费/...). This
    # renders it as a native PowerPoint table next to that account's own
    # commentary bullet. Config-gated (pptx_settings.presentation_tables.
    # enabled), off by default until verified against a real export.
    #
    # Reuses the generic table primitives above (add_table, _fit_table_
    # columns, _format_table_value, _set_paragraph_left_indent,
    # _resolve_table_style_id) rather than _embed_statement_table/
    # _fill_table_placeholder -- those are hard-coupled to the full BS/IS
    # overview grid (category-header rows, row-role tiering by keyword) that
    # a flat ~3-15 row account breakdown doesn't have and doesn't need.
    # Tightened one point per row 2026-08-11 to buy the income statement its
    # third page back. The arithmetic, in the planning unit (13.80pt for
    # Chinese): the IS holds 78.0 units against 75.9 of capacity across two
    # pages, i.e. it misses by 2.1 -- and the tail tolerance is 2.0, so it was
    # missing by a tenth of a line. All of the slack sat on IS page 1, which
    # carried only 营业收入 (7.2 of 23.9) because 营业成本 could not follow it
    # there: even split at its one legal boundary (lead + table stay, the
    # explanation moves on), that block needed ~266pt against ~252pt free.
    # Reading order forbids promoting a later account into the gap, and the
    # lead-only split was tried and deleted (the column ends on "明细如下："
    # with no detail under it), so the row metrics were the only lever left.
    #
    # A 16-row subtable goes 216pt -> 199pt, which puts lead+table at ~249pt
    # against ~252pt free -- it fits, with the page-2 remainder at 52.8 units
    # against 52.0 capacity, inside the same 2.0 tolerance.
    #
    # 11pt is the conservative end of what is safe, not the tightest that
    # fits. The floor is PowerPoint's own line pitch plus the cell margins
    # _set_cell applies: 7pt data font x 1.2 + 0.01in top + 0.01in bottom =
    # 9.84pt. Anything at or under that gets auto-grown by PowerPoint, which
    # would break the height reservation the same way a wrapped cell does.
    #
    # Tightened a second time (2026-08-11) after the trim above still left the
    # income statement a unit and a bit short of two pages. Safe because the
    # cell margins came down with them: _set_cell now writes 0.005in top and
    # bottom rather than 0.01in, so the floor at which PowerPoint auto-grows a
    # row -- font x 1.2 + margins -- drops from 9.84pt to 9.12pt for the 7pt
    # data font, and from 10.44 to 9.72 for the 7.5pt header. Every row below
    # clears its own floor by at least 0.9pt, where the previous 11pt row had
    # only 0.16pt of headroom against the old margins.
    _TABLE_TITLE_ROW_PT = 14.0
    _TABLE_HEADER_ROW_PT = 12.0
    _TABLE_DATA_ROW_PT = 10.0
    _TABLE_CHILD_ROW_PT = 10.0
    _TABLE_TOTAL_ROW_PT = 12.0
    _TABLE_SOURCE_LINE_PT = 12.0
    # Pure whitespace, not text-safety margin -- trimmed 2026-07-31 (6->4,
    # and the source-line box's own "+4" padding at its two call sites ->
    # "+2") to buy back a few points of bottom-of-slide margin on the
    # tallest real stack (营业成本: 16-row table + long lead-in + long
    # explanation measured at only 0.313in/22.6pt clear of the slide's
    # bottom edge on a real Crescent export). Safe to shrink because
    # neither is sized against variable AI text -- the source line is a
    # single fixed short string, and this gap is pure spacing between two
    # already-independently-sized shapes.
    # 4.0 -> 2.0 (2026-08-04). Real PowerPoint BoundHeight measurements
    # decomposed the visible gap under a lead-in exactly: 2.2pt of
    # paragraph spacing we count but BoundHeight doesn't show, 3.5pt from
    # the safety factor, and this 4pt spacer -- 9.7pt total, matching the
    # measured 9.7pt to the decimal. 2pt still reads as a deliberate
    # separation between the text and the table below it.
    _TABLE_GAP_ABOVE_PT = 2.0
    _TABLE_GAP_BELOW_PT = 4.0
    # Shared margin for both text blocks sized against a table (lead-in
    # above it, explanatory bullets below it): inspect_pptx.py re-derives
    # capacity/usage independently from the exported XML rather than
    # sharing this instance's live measurement state, so its verdict can
    # disagree with what was used to size the shape. 1.3x still left a
    # real generated lead-in at 104% (⚠️ OVERFLOW RISK) on a real Crescent
    # export -- raised rather than chasing exact parity between two
    # independently-implemented measurers, since the safe direction here
    # is a touch more headroom, not exact agreement.
    #
    # NOT the lever for a stack running close to the SLIDE's own bottom
    # edge (as opposed to text overflowing its own box) -- raising this
    # makes every box it sizes (lead-in, explanation) TALLER, which pushes
    # whatever comes after it further DOWN the slide. For a single tall
    # stack like 营业成本 that's already tight against the slide bottom,
    # raising this shrinks that margin further, not the other way round.
    _TEXT_HEIGHT_SAFETY_FACTOR = 1.6
    # Render-time counterpart for the two ACTUAL shape heights this feature
    # sets directly (table lead-in, table explanation). The full 1.6x
    # applied to an ACTUAL rendered box instead (as it was until
    # 2026-08-03) creates VISIBLE blank space below the real text and
    # before the table -- confirmed against a real Crescent export where a
    # ~3-line lead-in got sized for ~4.8 "line units", roughly 2 blank
    # lines, matching a real user report ("表格頂部有一些空白...大約2-3行").
    # 1.6x's own history (raised from 1.3x after THAT still read 104%
    # overflow-risk on inspect_pptx.py's independent re-check) means this
    # can't just be zeroed out -- some margin is real. Iteratively tuned
    # (1.35 -> 1.15 -> 1.25) against TWO real texts from an actual Crescent
    # export (营业成本's and 财务费用's own lead-in/explanation) via
    # inspect_pptx.py's independent re-check each time -- 1.15 looked fine
    # on the first (98%/94% fill) but the SECOND, shorter-but-differently-
    # wrapping explanation hit 101% (⚠️ OVERFLOW RISK), proving a single
    # real sample isn't enough evidence to land on a value. 1.25 clears
    # both with real margin (85-91% fill, not razor-thin) -- if a future
    # real case still overflows at 1.25, that's a real, reproducible data
    # point to retune from, not a guess.
    #
    # ALSO used (2026-08-03, same day) by the PLANNING-time estimates
    # below (_estimate_lead_in_pt / _estimate_table_account_block_height_
    # pt) -- these used to stay at the full 1.6x on the theory that a more
    # conservative planning estimate is the safe direction (better to
    # under-pack than overflow). That theory stopped holding once render
    # itself moved to 1.25x: planning at 1.6x now systematically believes
    # LESS room remains after a table block than render will actually
    # need, so a trailing account that would genuinely have fit was never
    # even tried -- confirmed against two real Crescent table pages
    # showing 1.5-2.2in of real, measured, unexplained blank column below
    # a finished table+explanation(+trailing account) stack. Planning and
    # render now use the SAME factor so the estimate that drives "does
    # this fit" decisions actually predicts what render will produce,
    # instead of two different numbers pulling in opposite directions.
    # 1.25 -> 1.10 (2026-08-04). At 1.25 EVERY box rendered at exactly
    # 1/1.25 = 78-80% fill -- the user saw that constant 20-25% tail as
    # "無論當中的字是多少行 下面都會有一行的空間". An earlier attempt at 1.10
    # DID overflow (103-154% on real client metrics) and had to be
    # reverted, but that was before the missing-inset bug was found: each
    # box was silently 7.2pt short of its own content, and this factor was
    # the only thing covering it. With insets now added properly, 1.10 was
    # re-validated by sweeping every real lead-in AND explanation text
    # from a real export, each scored with inspect_pptx.py's own
    # independent formula rather than the renderer's: worst case 87%
    # (lead) / 91% (explanation), i.e. real margin. The same sweep shows
    # 1.05 reaching 95% and 1.00 hitting 100% -- so 1.10 is the last value
    # with genuine headroom, not merely the first that happens to pass.
    _TABLE_RENDER_HEIGHT_SAFETY_FACTOR = 1.10






    # The literal handoff phrase ai.py's _detail_table_guidance asks the
    # model to end its short lead-in with -- the one point in the real
    # convention's two-part structure (lead-in, then optional "-"/"➢"
    # explanatory bullets) that's reliably the SAME string every time,
    # so it's what _split_table_commentary splits on.
    _TABLE_HANDOFF_CHI = "明细如下："
    _TABLE_HANDOFF_ENG = "the breakdown is set out below"




    # Planning-time-only assumption (no real shape exists yet for a slot a
    # table account hasn't been placed into) of how much vertical room a
    # slot has, used purely to decide whether a SECOND table account can
    # join one that already has a first -- render time positions every
    # account's own block at a running offset regardless of this estimate
    # (see _render_table_accounts_stack), so getting this exactly right
    # isn't load-bearing for correctness, only for how good the packing
    # decision is. ~331pt (single) to ~360pt (L/R) measured directly from
    # the real template (scratchpad/measure_slot_capacity.py); table
    # accounts land in L/R far more often than single (single is usually
    # claimed by the statement's own first/executive-summary slide), so
    # this leans toward that number rather than splitting the difference.
    _TABLE_SLOT_CAPACITY_PT = 355.0
    # Real accounts were only ADDED to a slot's already-claimed footprint at
    # 90% of the assumed capacity, not 100% -- deliberate margin for the
    # planning estimate (heuristic, no real font metrics) to disagree with
    # the real-metrics render-time measurement without the combined result
    # actually overflowing the slot. Not tighter than this: each account's
    # OWN estimate already carries the full _TEXT_HEIGHT_SAFETY_FACTOR
    # margin individually, so stacking a second, separate margin on top of
    # that here would mostly just suppress legitimate packing rather than
    # add real safety -- confirmed against a real Crescent export that two
    # of the four real accounts (税金及附加 ~167pt, 财务费用 ~222pt once its
    # own real, fairly long variance-flagging explanation is counted) don't
    # fit together at ANY reasonable threshold (389pt combined vs a
    # 331-360pt slot) -- that's a genuine size fact, not a threshold this
    # constant should be tuned to paper over.
    #
    # 0.90 -> 0.92 (2026-08-04), a deliberately small raise. Two errors
    # this threshold had been quietly absorbing were fixed the same day:
    # the planning estimates were ~14% low (font x 1.0 instead of
    # PowerPoint's real 1.2 line pitch -- see _planning_std_lh_pt) and it
    # multiplied one shared capacity constant instead of each slot's real
    # height (see _slot_capacity_pt). Measured after both fixes, the
    # estimate still runs ~5% under what render produces, so the SAFE
    # ceiling is about 0.95/1.05 ~= 0.90; 0.92 takes the small genuine
    # gain and no more. A first attempt at 0.96 was rejected on this
    # arithmetic alone -- it works out to ~101% of a column's real
    # capacity, i.e. overflow by construction.
    # 0.92 -> 0.98 (2026-08-04). Every reason this sat below 1.0 has now
    # been removed one at a time: the shared-capacity constant (now each
    # slot's real height, _slot_capacity_pt), the ~14% planning line-pitch
    # error (_planning_std_lh_pt), and finally the trailing-paragraph-gap
    # over-count -- the cost model now reproduces PowerPoint's own
    # BoundHeight exactly, so there is no measurement bias left for this
    # threshold to absorb. 0.98 keeps a 2% cushion for the render-time
    # rounding this estimate can't see, while letting a table's own block
    # actually use the column it was measured against; the user's
    # standing preference after seeing a slot land at exactly 0pt spare
    # is "用盡空間 比較好".
    _TABLE_SLOT_PACK_THRESHOLD = 0.98

    # How much of a column's REMAINING space a stranded lead-in has to fill
    # before it is worth separating it from its own table (see flow()'s
    # branch 2). The version of that split which was removed had no such
    # test and fired even when the lead was two words, leaving a column
    # essentially blank. At 0.5 the lead has to take at least half of what
    # is left, so the split only happens when it genuinely uses the column.
    _TABLE_SPLIT_MIN_LEAD_FILL = 0.5








    # Cell padding (0.04in left + 0.04in right, matching _set_cell's
    # margin_left/margin_right) and the child-row left indent (0.12in,
    # matching _set_cell's indent_emu for kind=="child") -- both added on
    # top of the raw measured text width below.
    _TABLE_CELL_PADDING_PT = 5.76
    _TABLE_CHILD_INDENT_PT = 8.64
    _TABLE_MIN_COLUMN_PT = 25.2  # 0.35in floor, guards a pathologically short column

    # A textbox's own top+bottom insets (OOXML default 0.05in each = 7.2pt
    # total). Height formulas that size a box from its TEXT's measured
    # height must add this back, or the box ends up exactly one inset
    # short of holding its own content -- the direct cause of lead-in and
    # explanation boxes sitting at 88-101% fill (and a real 101% OVERFLOW
    # RISK on a real export) instead of the intended ~80%.
    _TEXTBOX_INSET_PT = 7.2

    # Lets a table account's intro sentence finish one column while its
    # table starts the next (see place_table_item). Flip to False to fall
    # straight back to the previous whole-block-only behaviour without
    # touching anything else.
    _TABLE_ALLOW_LEAD_TABLE_SPLIT = True
    # Budget for the "科目名（续）" heading drawn above a continued fragment.
    # Must hold one full 9pt std_lh line (font x 1.2 pitch + paragraph gap
    # ~= 13pt) with the render safety margin -- a first attempt at 16pt
    # measured 148% full on real client metrics, because the textbox's own
    # default 3.6pt top/bottom insets eat into it. The heading renders with
    # those insets zeroed (see _render_continuation_heading) so this budget
    # is available to the text itself.
    _TABLE_CONTINUATION_HEADING_PT = 18.0



















    


    def refresh_project_placeholders(self, project_name: str):
        """Refresh placeholder tokens such as [PROJECT], [Current], and [Total]."""
        if not self.presentation or not project_name:
            return

        display_entity = str(project_name).strip()
        total_slides = len(self.presentation.slides)
        if not display_entity or total_slides <= 0:
            return

        base_replacements = {
            "[PROJECT]": display_entity,
            "[Entity]": display_entity,
            "[Company]": display_entity,
        }

        for slide_index, slide in enumerate(self.presentation.slides):
            proj_title_shape = find_shape_by_name(slide.shapes, "projTitle")
            if not proj_title_shape or not proj_title_shape.has_text_frame:
                continue

            replacements = dict(base_replacements)
            replacements["[Current]"] = str(slide_index + 1)
            replacements["[Total]"] = str(total_slides)
            current_text = proj_title_shape.text

            if any(token in current_text for token in replacements):
                replace_text_preserve_formatting(proj_title_shape, replacements)

            # Chinese reports: the template's own title text is an English
            # scaffold ("Entity overview - Project [PROJECT] (1/4)") with
            # only the [PROJECT] token substituted above -- for a Chinese
            # entity name that reads as English label + Chinese name mixed
            # mid-title. Strip the English lead-in so the title is just the
            # entity name (plus any pagination suffix the template kept),
            # e.g. "昆明经开 (1/4)" instead of "Entity overview - Project
            # 昆明经开 (1/4)".
            if self.language == 'chinese':
                stripped = proj_title_shape.text
                stripped = re.sub(
                    r'(?i)^\s*entity\s+overview\s*[-–:]\s*(project\s*)?',
                    '', stripped,
                )
                if stripped != proj_title_shape.text:
                    replace_text_preserve_formatting(
                        proj_title_shape, {proj_title_shape.text: stripped},
                    )

    def update_project_titles(self, project_name: str, statement_type: str = 'BS'):
        """Update project titles in presentation"""
        if not self.presentation:
            return

        display_entity = str(project_name or "").strip()
        self.refresh_project_placeholders(display_entity)

        # Define title templates based on language and statement type
        if statement_type.upper() == 'BS':
            if self.language == 'chinese':
                title_template = f"资产负债表概览 - {display_entity}"
            else:
                title_template = f"Entity Overview - {display_entity}"
        elif statement_type.upper() == 'IS':
            if self.language == 'chinese':
                title_template = f"利润表概览 - {display_entity}"
            else:
                title_template = f"Income Statement - {display_entity}"
        else:
            if self.language == 'chinese':
                title_template = f"财务报表概览 - {display_entity}"
            else:
                title_template = f"Financial Report - {display_entity}"

        # Update titles in all slides
        for slide_index, slide in enumerate(self.presentation.slides):
            current_slide_number = slide_index + 1
            proj_title_shape = find_shape_by_name(slide.shapes, "projTitle")

            if proj_title_shape:
                current_text = proj_title_shape.text
                if "[PROJECT]" in current_text:
                    replacements = {
                        "[PROJECT]": display_entity,
                        "[Current]": str(current_slide_number),
                        "[Total]": str(len(self.presentation.slides))
                    }
                    replace_text_preserve_formatting(proj_title_shape, replacements)
                else:
                    # Replace the entire title
                    if proj_title_shape.has_text_frame:
                        proj_title_shape.text_frame.text = title_template

    def generate_full_report(self, markdown_content: str, summary_md: Optional[str] = None,
                           output_path: str = None):
        """Generate full PowerPoint report from markdown content"""
        if not self.presentation:
            self.load_template()

        # Process markdown content
        processed_content = _process_markdown_content(markdown_content)

        # Apply content to presentation
        self._apply_content_to_presentation(processed_content)

        # Save if output path provided
        if output_path:
            self.save(output_path)


    def _apply_content_to_presentation(self, sections: Dict):
        """Apply processed content to presentation slides"""
        if not self.presentation:
            logger.warning("No presentation loaded")
            return

        logger.info("Applying %s sections to presentation with %s slides", len(sections), len(self.presentation.slides))

        # Find content placeholders and fill them
        slide_idx = 0
        for slide in self.presentation.slides:
            if slide_idx >= len(sections):
                logger.warning("More slides (%s) than sections (%s)", len(self.presentation.slides), len(sections))
                break

            account_name = list(sections.keys())[slide_idx]
            section_data = sections[account_name]

            logger.info("Processing slide %s for account: %s", slide_idx + 1, account_name)

            # Find content shape using flexible name matching
            content_shape = find_content_shape(slide.shapes)
            if content_shape:
                logger.info("Found content shape '%s' on slide %s", content_shape.name, slide_idx + 1)
                if content_shape.has_text_frame:
                    # Apply content to shape
                    _fill_content_shape(content_shape, section_data)
                    logger.info("Applied content to slide %s", slide_idx + 1)
                else:
                    logger.warning("Content shape found but has no text_frame on slide %s", slide_idx + 1)
            else:
                logger.warning("No content shape found on slide %s, available shapes: %s", slide_idx + 1, [s.name if hasattr(s, 'name') else 'unnamed' for s in slide.shapes])
                # Try to use the first available text frame as fallback
                for shape in slide.shapes:
                    if hasattr(shape, 'has_text_frame') and shape.has_text_frame:
                        shape_name = getattr(shape, 'name', 'unnamed')
                        if 'title' not in shape_name.lower() and 'proj' not in shape_name.lower():
                            logger.info("Using fallback shape '%s' on slide %s", shape_name, slide_idx + 1)
                            _fill_content_shape(shape, section_data)
                            break

            slide_idx += 1







    # Space-after (pt) applied to every paragraph in _fill_text_main_bullets.
    # Matches _fill_text_main_bullets_with_category_and_key's own hardcoded
    # p_key.space_after = Pt(3) (see the "Matches _PARA_SPACE_AFTER (cost
    # estimator)" comment right there in that function) -- the ACTUAL,
    # currently-live commentary-bullet renderer, unlike get_space_after_for_
    # text/get_space_before_for_text below, which belong to a different,
    # legacy code path (_fill_content_shape, reached only from the unused
    # markdown generate() flow) and were never actually applied to a
    # textMainBullets run.
    _PARA_SPACE_AFTER = 3.0







    # Rendered bullet paragraphs hang-indent their WRAPPED lines by 0.15"
    # (p_key.left_indent = Inches(0.15), first_line_indent = Inches(-0.15)):
    # line 1 starts at the box margin and spans the full width, every wrapped
    # line starts 10.8pt further right. Cost estimates wrap at
    # (width - 10.8pt) -- exact for wrapped lines, ~1 char conservative for
    # line 1. Measuring every line at full width under-counted long
    # paragraphs by up to a line each; that error was masked while the old
    # line-height model over-charged the vertical axis, and became real
    # (uncaught-by-autofit) overflow once the 1.2x-size pitch fix landed.
    _BULLET_HANGING_INDENT_PT = 10.8  # 0.15 inch







    # Currency markers that read as visually broken when stranded at the end
    # of a slide with their amount on the next one (real production case:
    # "...bad debt losses of CNY" / "484,000 in FY25..." -- the number itself
    # was intact, only the currency prefix got separated from it).
    _CURRENCY_MARKERS = ('CNY', 'RMB', 'USD', 'HKD', '人民币', '¥', '$', '£', '€')

    # Common Chinese words/compounds this domain's commentary is full of --
    # protected the same way as _CURRENCY_MARKERS (never split strictly
    # inside one). The mid-word/mid-marker fix for "人民币" alone wasn't
    # enough: the SAME hard-cut/character-trim fallback that used to land
    # inside "人民币" lands just as easily inside any other bare CJK
    # compound, since Chinese has no spaces to signal a word boundary in
    # the first place. Confirmed in real production output on the SAME
    # page: "...人民币784万" | "元变为..." (万元 split), "...2026年06月30日
    # 分" | "别为人民币16万元..." (分别 split), "...2025年" | "度为人民币
    # 17万元..." (年度 split) -- three different compounds, not a single
    # fixable special case. Not exhaustive (full Chinese word segmentation
    # is a different-sized problem) -- this is the domain-specific set
    # observed breaking in real financial-commentary text plus the most
    # obviously analogous high-frequency terms; the mechanism only ever
    # backs a split up (never forces a worse one), so a broader list here
    # is safe to extend further if new cases turn up.
    _PROTECTED_CJK_COMPOUNDS = (
        '万元', '亿元', '年度', '分别', '期间', '变为', '转为', '增至', '降至',
        '合计', '其中', '构成', '包括', '管理层', '余额', '截至', '备注',
        '数据', '期末', '期初', '账款', '借款', '贷款', '利息', '费用',
        '资产', '负债', '权益', '收入', '成本', '折旧', '摊销', '核对',
        '差异', '发生', '形成', '增加', '维持', '原值', '净值', '账面',
        '进一步', '未发生', '无余额', '未形成',
    )

    # Organisation-name tails. A split anywhere inside a company name reads
    # as a mistake even though every generic protection above passes it:
    # jieba tokenizes 某某系统|工程|第四|建设|有限公司 into legitimate words,
    # so a cut between any two of them looks like a clean word boundary.
    # Detecting the name's START is unreliable; detecting that pos is
    # heading INTO one of these tails is not.
    _ORG_NAME_TAILS = (
        '有限公司', '有限责任公司', '股份有限公司', '公司', '银行', '集团',
        '事务所', '研究院', '设计院', '工程局', '合伙企业', '分行', '支行',
    )
    _SENTENCE_PUNCT = '。，；、：！？,;.:!?'
















    def apply_structured_data_to_slides(self, structured_data: List[Dict], start_slide: int,
                                       project_name: str, statement_type: str, is_chinese_databook: bool = False,
                                       pre_generated_summary: Optional[str] = None):
        """Apply structured data directly to slides (slides 1-4 for BS, 5-8 for IS).

        If ``pre_generated_summary`` is provided, it's used directly for the
        first slide's coSummaryShape — no AI call from inside PPTX export.
        """
        if not self.presentation:
            self.load_template()

        # Remembered so save() can decide, once for the whole deck, whether
        # to strip the template's static English "Commentary" label bands
        # (see _localize_label_bands) -- both statement passes
        # (BS then IS) report the same databook language, so whichever
        # runs last simply re-confirms the same value.
        self._is_chinese_mode = is_chinese_databook

        stage_started_at = time.perf_counter()
        logger.info("Applying %s accounts to slides starting at %s", len(structured_data), start_slide)

        # Normalize commentary and store originals for fill optimization
        structured_data = _prepare_structured_data_for_slides(structured_data)

        # Continuation slides (every slide of this statement after the first)
        # lose their coSummaryShape and gain that area as extra commentary
        # space. The executive summary stays only on the first slide of
        # each statement, which cuts AI summary calls from up to 8 to 2.
        max_slides = int(self.pptx_settings.get("max_commentary_slides_per_statement", 4) or 4)
        first_slide_idx = start_slide - 1
        for offset in range(1, max_slides):
            cont_idx = first_slide_idx + offset
            if cont_idx >= len(self.presentation.slides):
                break
            _expand_commentary_to_cover_summary(self.presentation.slides[cont_idx])

        # Presentation-table accounts are pulled out of the packing pool
        # entirely rather than fed to the DP/greedy packer with an inflated
        # cost. A first attempt padded commentary cost to make a table
        # account "claim" its whole slot -- verified against a real render
        # that this does NOT work: the packer still placed a second account
        # in the same slot, then SPLIT it across two slides when the
        # combined cost overflowed (the overflow-split logic works purely
        # off the commentary text, with no concept of "this text is padding
        # for a table, do not cut it"). Two tables ended up drawn at the
        # exact same on-slide position. Removing them before packing runs
        # sidesteps that entirely -- the packer never knows they exist, so
        # its own sharing/splitting behaviour can't touch them.
        slot_distribution = self._plan_slot_distribution(
            structured_data, max_slides=max_slides, start_slide=start_slide,
            statement_type=statement_type, is_chinese_databook=is_chinese_databook,
        )

        # Group slot distribution by slide for easier processing
        slides_content = {}  # {slide_idx: {'single': [...], 'L': [...], 'R': [...]}}
        for slot_slide_idx, slot_name, account_list in slot_distribution:
            if slot_slide_idx not in slides_content:
                slides_content[slot_slide_idx] = {}
            slides_content[slot_slide_idx][slot_name] = account_list
        
        # Ensure we have enough slides
        if slides_content:
            max_slide_idx = max(slides_content.keys())
            needed_slides = start_slide + max_slide_idx
            current_slide_count = len(self.presentation.slides)
            
            if needed_slides > current_slide_count:
                # Add slides if needed
                if current_slide_count > 0:
                    slide_layout = self.presentation.slides[0].slide_layout
                    for _ in range(needed_slides - current_slide_count):
                        self.presentation.slides.add_slide(slide_layout)
        
        # Track which slides are used
        used_slide_indices = set()
        summary_jobs: List[Dict[str, Any]] = []
        
        # Apply content to slides
        for slide_idx in sorted(slides_content.keys()):
            actual_slide_idx = start_slide - 1 + slide_idx  # Convert to 0-based
            if actual_slide_idx >= len(self.presentation.slides):
                logger.warning("Slide index %s exceeds available slides", actual_slide_idx + 1)
                continue
            
            slide = self.presentation.slides[actual_slide_idx]
            slot_contents = slides_content[slide_idx]  # {'single': [...], 'L': [...], 'R': [...]}
            # A slide only counts as "used" if at least one of its slots has
            # real content -- a slide whose slots are ALL empty (e.g. after
            # _consolidate_trailing_near_empty_slot folds its one leftover
            # sliver into the previous slide) should be eligible for the
            # unused-slide removal below, not kept around as a blank page.
            if any(slot_contents.values()):
                used_slide_indices.add(actual_slide_idx)

            # Note: Financial tables are filled by embed_financial_tables()

            # Collect all accounts on this slide for summary generation
            all_slide_accounts = []
            for slot_name, accounts in slot_contents.items():
                all_slide_accounts.extend(accounts)

            used_slot_shape_ids = set()
            
            # Fill each slot (single, L, or R) on this slide
            for slot_name, account_data_list in slot_contents.items():
                if not account_data_list:
                    # An INTENTIONALLY empty slot (e.g.
                    # _consolidate_tiny_stub_lr_pairs folded its one tiny
                    # fragment into the earlier slot rather than leave an
                    # orphaned near-empty column) must still have its
                    # shape's text actively cleared -- otherwise the shape
                    # keeps whatever raw template placeholder sample text
                    # it shipped with (e.g. a visible "Placeholder –
                    # placeholder"), which reads as broken output, not as
                    # the clean blank box an intentionally-unused slot
                    # should be.
                    empty_shape = self._resolve_commentary_slot_shape(
                        slide, slot_name, used_shape_ids=used_slot_shape_ids,
                    )
                    if empty_shape is not None and empty_shape.has_text_frame:
                        empty_shape.text_frame.clear()
                        used_slot_shape_ids.add(id(empty_shape))
                    continue

                # Find the appropriate shape based on slot_name
                bullets_shape = self._resolve_commentary_slot_shape(
                    slide,
                    slot_name,
                    used_shape_ids=used_slot_shape_ids,
                )
                if not bullets_shape and slot_name == "single":
                    bullets_shape = self._resolve_commentary_slot_shape(
                        slide,
                        "L",
                        used_shape_ids=used_slot_shape_ids,
                    ) or self._resolve_commentary_slot_shape(
                        slide,
                        "R",
                        used_shape_ids=used_slot_shape_ids,
                    )
                if not bullets_shape:
                    bullets_shape = _add_commentary_slot_shape(slide, slot_name)
                
                if not bullets_shape.has_text_frame:
                    logger.warning("Slide %s: Shape for slot '%s' has no text frame", actual_slide_idx + 1, slot_name)
                    continue
                used_slot_shape_ids.add(id(bullets_shape))

                # A slot containing at least one table account (possibly
                # followed by plain trailing accounts flowed into its
                # leftover space -- see _append_table_accounts_to_
                # distribution's trailing_items) renders via its own
                # dedicated-shape-per-account path instead of the
                # shared-text-frame one below, since that one assumes all of
                # a slot's content lives in one text frame with nothing else
                # interleaved -- not true once a table (a separate shape)
                # sits between one account's lead-in and the next account's
                # own lead-in. A slot with ONLY plain accounts (no table at
                # all) still goes through the normal shared-text-frame path
                # unchanged -- that's the common case and needs none of this.
                if account_data_list and any(a.get("_presentation_table") for a in account_data_list):
                    self._render_table_accounts_stack(
                        slide, bullets_shape, account_data_list, is_chinese_databook,
                        statement_type=statement_type,
                    )
                    continue

                # Fill this slot
                tf = bullets_shape.text_frame
                tf.clear()
                tf.word_wrap = True
                # TextFrame.clear() always leaves exactly one (now-empty)
                # paragraph behind; every real line below is added via
                # add_paragraph() rather than reusing it, so it survives as a
                # permanent blank leading line PowerPoint actually renders
                # (~1 line-height of capacity silently lost on every slot in
                # the deck) even though it contributes zero to this file's
                # own content-line accounting (empty paragraphs are filtered
                # out by both the DP/packing math and inspect_pptx.py's
                # measurement) -- a real vs. believed capacity mismatch that
                # was part of why slots read as "stuck" under their true
                # capacity. Removed once real content exists, right before
                # moving to the next slot.
                _leading_empty_p = tf.paragraphs[0]._p

                # Most slots fit within strict 1.0x capacity and keep the
                # exact 9pt/10pt size (noAutofit) -- no drift, no surprise
                # shrink. The DP occasionally packs a slot beyond that
                # (shape_height_utilization / the looser relax tiers) on
                # the assumption PowerPoint's autofit absorbs the excess;
                # noAutofit alone would silently clip that content instead,
                # so give exactly those slots a bounded normAutofit shrink
                # (never below _BOUNDED_AUTOFIT_MIN_SCALE) so the overflow
                # the DP already decided was "worth it" is actually visible.
                #
                # But ONLY past the tail tolerance. Writing a normAutofit does
                # not shrink text by the fontScale computed here: PowerPoint
                # re-runs its own shrink with its own font metrics and steps
                # down in coarse increments, so it lands wherever it lands.
                # Measured against a real deck (VBA BoundHeight, the ground
                # truth): a slot overflowing by 0.3 lines -- fontScale ~99%
                # written here -- rendered at ~87%, 27 lines at 10.31pt
                # instead of 10.8pt, leaving 52.8pt of visible blank under a
                # box the packer had filled to 101%. Every other box in that
                # deck measured >= 10.8pt/line; this was the only shrunk one
                # and the only one over capacity. Below the tolerance, letting
                # the tail protrude is both what the project team asked for
                # ("1-2 lines sticking out is fine") and strictly better
                # looking than a whole column shrunk to leave a gap.
                slot_is_chinese = any(_account_is_chinese(a) for a in account_data_list)
                slot_capacity = self._calculate_max_lines_for_textbox(
                    bullets_shape, is_chinese=slot_is_chinese, slot_name=slot_name,
                    statement_type=statement_type,
                )
                slot_used = self._compute_slot_used_lines(
                    account_data_list, slot_name, slot_shape=bullets_shape,
                    statement_type=statement_type,
                )
                _tail_tol = self._tail_overflow_tolerance_units(statement_type)
                if slot_capacity > 0 and (slot_used - slot_capacity) > _tail_tol:
                    self._apply_bounded_autofit(tf, slot_capacity / slot_used)
                else:
                    _force_no_autofit(tf)  # keep text at 9pt/10pt, never shrink
                from pptx.enum.text import MSO_VERTICAL_ANCHOR
                tf.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP
                
                # Fixed deck-wide font size — 9pt Arial, no exceptions,
                # no per-slot / per-language variation. This matches the
                # hardcoded return from get_font_size_for_text() and
                # guarantees identical typography on every slide.
                slot_font_size = 9
                logger.info(
                    "Slide %s, slot '%s': Filling with %s accounts at %spt",
                    actual_slide_idx + 1, slot_name, len(account_data_list), slot_font_size,
                )

                # Fill with accounts, grouped by category
                # Show category header only once per category group
                current_category = None
                for account_idx, account_data in enumerate(account_data_list):
                    category = account_data.get('category', '')
                    mapping_key = account_data.get('mapping_key', account_data.get('account_name', ''))
                    # display_name is whatever text sat in the source Financials
                    # sheet's row (often English even for a Chinese-language
                    # report, e.g. Kunshan's own sheets are English-labelled) --
                    # use the pre-resolved Chinese alias for a Chinese report
                    # instead of leaving that English label in an otherwise
                    # fully-translated bullet.
                    display_name = (
                        account_data.get('display_name_zh') or account_data.get('display_name', mapping_key)
                        if is_chinese_databook
                        else account_data.get('display_name', mapping_key)
                    )
                    commentary = account_data.get('commentary', '')
                    clause_reviews = account_data.get('clause_reviews', [])
                    is_chinese = account_data.get('is_chinese', False)
                    is_continuation = account_data.get('is_continuation', False)
                    
                    # Skip category header if this is a continuation of a split account
                    # Show category header only when category changes
                    if category and category != current_category and not is_continuation:
                        # Add category header - use Chinese if databook is Chinese
                        p_category = tf.add_paragraph()
                        p_category.level = 0
                        try:
                            p_category.left_indent = Inches(0.21)
                            p_category.first_line_indent = Inches(-0.19)
                            p_category.space_before = Pt(3) if current_category else Pt(0)
                            p_category.space_after = Pt(0)
                            p_category.line_spacing = 1.0
                            _apply_east_asian_line_breaking(p_category)
                        except:
                            pass
                        
                        run_category = p_category.add_run()
                        # Use Chinese category name if databook is Chinese
                        category_text = translate_category_to_chinese(category) if is_chinese_databook else category
                        
                        run_category.text = category_text
                        run_category.font.size = Pt(slot_font_size)
                        run_category.font.name = 'Arial'
                        self._set_east_asian_typeface(run_category)
                        self._declare_run_language(run_category)
                        run_category.font.bold = False
                        try:
                            from pptx.dml.color import RGBColor
                            run_category.font.color.rgb = RGBColor(0, 51, 102)  # Dark blue
                        except:
                            pass
                        
                        current_category = category
                    
                    # Fill commentary with key formatting
                    # For continuation accounts, show "(cont'd)" or "(续)" after key name
                    if is_continuation:
                        # More prominent continuation marker
                        if is_chinese_databook:
                            display_name_with_cont = f"{display_name} (续)"
                        else:
                            display_name_with_cont = f"{display_name} (cont'd)"
                        
                        # Log continuation for debugging
                        logger.info("Displaying continuation: %s", display_name_with_cont)
                    else:
                        display_name_with_cont = display_name
                    
                    self._fill_text_main_bullets_with_category_and_key(
                        tf, None, display_name_with_cont, commentary, is_chinese,
                        is_chinese_databook=is_chinese_databook, needs_continuation=False,
                        font_size_pt=slot_font_size,
                        clause_reviews=clause_reviews,
                    )
                    # Table-bearing accounts never reach this loop -- they're
                    # intercepted by the table-only-slot dispatch above and
                    # rendered via _render_table_accounts_stack instead.

                if _leading_empty_p.getparent() is not None and not (_leading_empty_p.text or "").strip():
                    _leading_empty_p.getparent().remove(_leading_empty_p)

            page_commentary, page_summary_source = self._build_page_summary_source(all_slide_accounts)

            # Collect coSummaryShape jobs and fill after summaries are generated.
            summary_shape = find_shape_by_name(slide.shapes, "coSummaryShape")
            if summary_shape and summary_shape.has_text_frame:
                summary_shape.text_frame.clear()
                _force_no_autofit(summary_shape.text_frame)
                # Claim the gap the template leaves between this band and the
                # table below. _summary_length_targets sizes the text against
                # the same measurement, so the box and the request stay in step.
                self._grow_summary_band_into_spare(slide, summary_shape)
                if page_summary_source:
                    summary_jobs.append({
                        "slide_idx": actual_slide_idx,
                        "summary_shape": summary_shape,
                        "page_commentary": page_commentary,
                        "page_summary_source": page_summary_source,
                        "is_chinese": is_chinese_databook,
                        "font_is_chinese": all_slide_accounts[0].get('is_chinese', False) if all_slide_accounts else False,
                    })
                else:
                    # Every way this band ends up blank currently looks
                    # identical from the outside: shape missing, no source
                    # text, generator returned nothing, a blank pre-generated
                    # summary. Each one silently does nothing, so "exe sum
                    # 一直未有東西" cost several rounds of guessing. Name the
                    # cause at WARNING level -- inspect_databook.py's export
                    # log analysis surfaces warnings, so the next real run
                    # says which one it was instead of just showing a blank.
                    logger.warning(
                        "Slide %s: coSummaryShape found but NO summary source text "
                        "(%s accounts on this slide, %s with commentary) — band left blank",
                        actual_slide_idx + 1, len(all_slide_accounts),
                        sum(1 for a in all_slide_accounts if str(a.get("commentary", "") or "").strip()),
                    )
            elif actual_slide_idx == start_slide - 1:
                logger.warning(
                    "Slide %s: coSummaryShape NOT FOUND on the first slide of %s — band "
                    "left blank. The lookup is an exact name match; check the template.",
                    actual_slide_idx + 1, statement_type,
                )
            else:
                # Continuation slides have their band removed on purpose by
                # _expand_commentary_to_cover_summary, so its absence here is
                # the designed layout, not a fault. Warning on it put three
                # false lines into the export log analysis of a run whose
                # summaries had in fact all worked.
                logger.info(
                    "Slide %s: continuation slide, no summary band by design",
                    actual_slide_idx + 1,
                )

            logger.info("Filled slide %s with %s accounts across %s slots", actual_slide_idx + 1, len(all_slide_accounts), len(slot_contents))

        # If a pre-generated summary was supplied (computed during the AI
        # commentary phase), use it directly and skip the in-PPTX AI call.
        # The pre-generated summary applies to the FIRST slide of this
        # statement only — continuation slides have coSummaryShape removed.
        pre_summary_text = str(pre_generated_summary or "").strip() if pre_generated_summary else ""
        if pre_summary_text and summary_jobs:
            summary_jobs.sort(key=lambda j: j["slide_idx"])
            first_job = summary_jobs[0]
            summary_results = {first_job["slide_idx"]: pre_summary_text}
            jobs_to_fill = [first_job]
        elif summary_jobs:
            # No pre-generated summary supplied. Calling LLM during PPTX export
            # is slow when the API is flaky (3 retries × 30s × N slides) and
            # the user reported 10+ min hangs. Skip the in-export AI call —
            # leave coSummaryShape blank rather than wait. The user can re-run
            # the AI generation step to refresh summaries when the API is
            # responsive.
            # Skipping the LLM call is right; leaving the box EMPTY was not.
            # The front page of a deliverable shipped with a blank summary
            # band, and because the template's own placeholder in that shape
            # is set to 2.17pt, what showed instead was an illegible sliver
            # of "Entity overview – Project [PROJECT]" -- reported as "exe
            # sum 一直未有東西".
            #
            # _generate_page_summary already exists for exactly this case:
            # it takes the opening sentence of each account paragraph, so it
            # spans the whole page, costs no tokens and no wall time. Use it
            # rather than shipping nothing.
            logger.info(
                "No pre-generated summary supplied; using the deterministic "
                "page summary instead of an in-export LLM call (which would "
                "add 1-3 min per slide)."
            )
            summary_jobs.sort(key=lambda j: j["slide_idx"])
            summary_results = {}
            jobs_to_fill = []
            for job in summary_jobs:
                _src = job.get("page_summary_source") or job.get("page_commentary") or ""
                text = self._generate_page_summary(_src, bool(job.get("is_chinese")))
                if not text:
                    logger.warning(
                        "Slide %s: _generate_page_summary returned nothing from %s chars "
                        "of source — summary band left blank",
                        job["slide_idx"] + 1, len(_src),
                    )
                if text:
                    summary_results[job["slide_idx"]] = text
                    jobs_to_fill.append(job)
        else:
            # No job was ever collected. Either no slide in this statement has
            # a coSummaryShape, or none had source text -- both already warned
            # per-slide above, so this only records that the fill stage ran
            # with nothing to do rather than silently falling through.
            logger.warning(
                "%s: no coSummaryShape jobs collected — every summary band for "
                "this statement stays blank", statement_type,
            )
            summary_results = {}
            jobs_to_fill = []
        for job in jobs_to_fill:
            final_summary = str(summary_results.get(job["slide_idx"]) or "").strip()
            if not final_summary:
                logger.warning("Slide %s: summary resolved to empty text — band left blank",
                               job["slide_idx"] + 1)
                continue
            logger.info("Slide %s: wrote %s-char executive summary into coSummaryShape",
                        job["slide_idx"] + 1, len(final_summary))
            summary_shape = job["summary_shape"]
            p = summary_shape.text_frame.paragraphs[0] if summary_shape.text_frame.paragraphs else summary_shape.text_frame.add_paragraph()
            p.text = final_summary
            _apply_east_asian_line_breaking(p)
            for run in p.runs:
                run.font.size = get_font_size_for_text(final_summary, force_chinese_mode=job["font_is_chinese"])
                run.font.name = get_font_name_for_text(final_summary)
                # This fill site sets font.name from a helper rather than the
                # literal 'Arial' the other 17 sites use, so it was missed
                # when the CJK typeface and language were added everywhere
                # else. The summary band is the first Chinese text on the
                # page; it needs them at least as much.
                self._set_east_asian_typeface(run)
                self._declare_run_language(run)
        
        # Record the FIRST used commentary slide of this statement as a slide
        # OBJECT (not an index). embed_financial_tables targets this so the BS/IS
        # table lands on the correct page even after slides are added/removed.
        if used_slide_indices:
            first_used_idx = min(used_slide_indices)
            if 0 <= first_used_idx < len(self.presentation.slides):
                if not hasattr(self, "_statement_table_slides"):
                    self._statement_table_slides = {}
                self._statement_table_slides[statement_type] = self.presentation.slides[first_used_idx]

        # Note: Unused slides will be removed at the end, after all content and tables are embedded
        # Store unused slides for later removal
        statement_slide_range = list(range(start_slide - 1, min(start_slide + 3, len(self.presentation.slides))))
        unused_slides = [idx for idx in statement_slide_range if idx not in used_slide_indices]
        if unused_slides:
            # Store for later removal - don't remove now
            if not hasattr(self, '_unused_slides_to_remove'):
                self._unused_slides_to_remove = []
            self._unused_slides_to_remove.extend(unused_slides)
            logger.info("Marked %s unused slides for %s for later removal: %s", len(unused_slides), statement_type, [idx + 1 for idx in unused_slides])
        logger.info(
            "PPTX stage apply_structured_data_to_slides[%s] took %.2fs across %s populated slides",
            statement_type,
            time.perf_counter() - stage_started_at,
            len(slides_content),
        )
    
    def _remove_slides(self, slide_indices):
        """Remove slides by indices (from backup method)"""
        # Sort in reverse order to maintain indices while removing
        for slide_idx in sorted(slide_indices, reverse=True):
            if slide_idx < len(self.presentation.slides):
                try:
                    # Use XML-based removal (from backup method)
                    xml_slides = self.presentation.slides._sldIdLst
                    slides = list(xml_slides)
                    
                    if slide_idx < len(slides):
                        # Get the slide element to remove
                        slide_element = slides[slide_idx]
                        # Remove relationship
                        rId = slide_element.rId
                        self.presentation.part.drop_rel(rId)
                        # Remove from XML
                        xml_slides.remove(slide_element)
                        logger.info("Removed slide %s", slide_idx + 1)
                    else:
                        logger.warning("Slide index %s out of range (only %s slides)", slide_idx, len(slides))
                except Exception as e:
                    logger.warning("Could not remove slide %s: %s", slide_idx + 1, e)
                    logger.debug(traceback.format_exc())
    

    
    

    _CJK_RANGE = ("一", "鿿")

    def _declare_run_language(self, run) -> None:
        """Declare the run's LANGUAGE, which is what selects the
        line-breaking algorithm.

        Third attempt at the "。 starts a line" defect, and the first at this
        mechanism. The previous two targeted the RULE (paragraph eaLnBrk /
        hangingPunct) and the FONT (<a:ea> typeface); both are set correctly
        and neither changed the render. The missing piece is that our runs
        declared no language at all -- <a:rPr> carried no lang -- so
        PowerPoint fell back to the document default, and this template is an
        English one. Latin line-breaking has no 禁则处理, which is exactly the
        behaviour observed: a line beginning with 。

        lang is set from the run's own text so this is correct in an English
        deck too; altLang says what to do with East Asian characters inside
        an otherwise-Latin run, which is the common case here ("■ ", " - ",
        and a Chinese sentence are separate runs of one bullet).
        """
        try:
            text = run.text or ""
            has_cjk = any(self._CJK_RANGE[0] <= ch <= self._CJK_RANGE[1] for ch in text)
            rPr = run._r.get_or_add_rPr()
            rPr.set("lang", "zh-CN" if has_cjk else "en-US")
            rPr.set("altLang", "zh-CN")
        except Exception as exc:
            logger.debug("Could not declare run language: %s", exc)

    def _set_east_asian_typeface(self, run) -> None:
        """Declare the East Asian typeface on a run, alongside the Latin one.

        Every run this file creates sets font.name='Arial', which python-pptx
        writes as <a:latin>. Nothing declared <a:ea>, so PowerPoint had to
        pick a CJK fallback itself -- and, more importantly, may classify the
        run as Latin script and apply LATIN line-breaking to it. That is the
        most likely reason a real deck still broke a line before a 、 even
        after eaLnBrk/hangingPunct were set: the paragraph asked for East
        Asian rules while the run looked like Latin text.

        It also removes a real inconsistency: the packer MEASURES Chinese
        text with the configured CJK metrics (Microsoft YaHei by default) but
        was telling PowerPoint the font was Arial. <a:ea> applies only to
        East Asian characters, so declaring it cannot affect Latin text and
        is safe to set unconditionally.
        """
        try:
            from pptx.oxml.ns import qn
            family = self._packing_settings().get("font_family_chi") or "Microsoft YaHei"
            rPr = run._r.get_or_add_rPr()
            for existing in rPr.findall(qn('a:ea')):
                rPr.remove(existing)
            ea = rPr.makeelement(qn('a:ea'), {'typeface': family})
            # <a:ea> must follow <a:latin> in the CT_TextCharacterProperties
            # sequence; append after it when present, else at the end.
            latin = rPr.find(qn('a:latin'))
            if latin is not None:
                latin.addnext(ea)
            else:
                rPr.append(ea)
        except Exception as exc:
            logger.debug("Could not set East Asian typeface: %s", exc)



    # A slot the DP could only pack by relaxing past strict 1.0x capacity
    # (shape_height_utilization / the later 1.35x, 1.6x tiers) has no room
    # to actually hold its content at 9pt with noAutofit -- that combination
    # silently CLIPS the overflow at the shape edge, contradicting the DP's
    # own relax-factor comment ("PPT auto-fit can absorb that much
    # overflow"). 0.70 is a floor, not the typical case: the DP's own
    # relax tiers below its 10.0x always-feasible last resort top out at
    # 1.6x (needing ~63% scale), so this floor mainly guards against that
    # last-resort tier producing illegibly small text rather than clipping.
    _BOUNDED_AUTOFIT_MIN_SCALE = 0.70

    @classmethod
    def _apply_bounded_autofit(cls, text_frame, font_scale: float) -> None:
        """Set ``<a:normAutofit fontScale="..." lnSpcReduction="..."/>`` so
        PowerPoint actually shrinks text to fit instead of clipping it, for
        a slot the DP intentionally packed beyond strict 1.0x capacity.
        ``font_scale`` is capacity/used, clamped to
        ``_BOUNDED_AUTOFIT_MIN_SCALE`` so an extreme overflow (the DP's
        10.0x always-feasible last-resort tier) still clips rather than
        rendering illegibly small text -- a bounded, not unlimited, shrink.
        """
        scale = max(cls._BOUNDED_AUTOFIT_MIN_SCALE, min(1.0, float(font_scale)))
        font_pct = int(round(scale * 100000))
        # PowerPoint reduces line spacing somewhat less aggressively than
        # font size when auto-fitting -- half the font shrink, floored at 0.
        line_pct = int(round(max(0.0, (1.0 - scale) * 0.5) * 100000))
        try:
            from pptx.oxml.ns import qn
            from pptx.oxml import parse_xml
            bodyPr = text_frame._txBody.bodyPr
            for tag in ("a:spAutoFit", "a:normAutofit", "a:noAutofit"):
                for child in bodyPr.findall(qn(tag)):
                    bodyPr.remove(child)
            bodyPr.append(parse_xml(
                f'<a:normAutofit xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
                f'fontScale="{font_pct}" lnSpcReduction="{line_pct}"/>'
            ))
        except Exception as exc:
            logger.debug("Could not apply bounded normAutofit on text frame: %s", exc)




    def _add_runs_for_line(
        self,
        paragraph,
        line: str,
        clause_segments: Optional[List[Tuple[str, str]]],
        font_size_pt: int,
    ) -> None:
        """Add one or more runs to `paragraph` for `line`, applying clause colours when applicable."""
        from pptx.dml.color import RGBColor

        def _apply_run_format(run, color_rgb: Optional[Tuple[int, int, int]]):
            run.font.size = Pt(font_size_pt)
            run.font.name = 'Arial'
            self._set_east_asian_typeface(run)
            self._declare_run_language(run)
            run.font.bold = False
            try:
                if color_rgb is not None:
                    run.font.color.rgb = RGBColor(*color_rgb)
                else:
                    run.font.color.rgb = RGBColor(0, 0, 0)
            except Exception:
                pass

        if not clause_segments:
            run = paragraph.add_run()
            run.text = line
            _apply_run_format(run, None)
            return

        # Find which segments overlap with this line. Since segments span the
        # whole commentary, we need to figure out where this line fits. For
        # simplicity, walk the segments while consuming this line's characters.
        remaining = line
        for segment_text, category in clause_segments:
            if not remaining:
                break
            if not segment_text:
                continue
            # Skip leading characters of the segment that aren't in the line
            # (e.g. the segment may start before this line begins).
            overlap_start = remaining.find(segment_text)
            if overlap_start == 0:
                run = paragraph.add_run()
                run.text = segment_text
                _apply_run_format(run, _category_to_rgb(category))
                remaining = remaining[len(segment_text):]
            elif overlap_start > 0:
                # Plain prefix before this segment
                run = paragraph.add_run()
                run.text = remaining[:overlap_start]
                _apply_run_format(run, None)
                # Then the segment
                run = paragraph.add_run()
                run.text = segment_text
                _apply_run_format(run, _category_to_rgb(category))
                remaining = remaining[overlap_start + len(segment_text):]
            # else: segment doesn't appear on this line, skip it
        if remaining:
            run = paragraph.add_run()
            run.text = remaining
            _apply_run_format(run, None)


    _ORPHANABLE_END_PUNCT = "。．.；;"


    def _fill_text_main_bullets_with_category_and_key(self, text_frame, category: str, display_name: str,
                                                      commentary: str, is_chinese: bool, is_chinese_databook: bool = False,
                                                      needs_continuation: bool = False, font_size_pt: int = 9,
                                                      clause_reviews: Optional[List[Dict[str, Any]]] = None):
        """
        Fill textMainBullets shape with commentary formatted as:
        - Category as first level (dark blue Arial 9) - only if category is provided
        - Key name with filled round bullet + space + key name (black bold Arial 9) + "-" (not bold) + plain text
        - Indentation 0.15" with special hanging 0.15", spacing after 6pt
        - When clause_reviews is provided, non-data-backed clauses are coloured:
          orange (213, 94, 0) for 'reasoning', red (200, 16, 46) for 'hallucination'.
        """
        from pptx.util import Inches
        from pptx.dml.color import RGBColor
        from pptx.enum.text import PP_ALIGN

        # Before anything measures or renders this text, and before the
        # clause segments that carry hallucination highlighting are cut from
        # it, so every downstream consumer sees the same string.
        commentary = self._drop_orphan_trailing_punctuation(
            f"■ {display_name} - " if display_name else "", commentary, is_chinese,
        )

        clause_segments = _build_clause_segments(commentary, clause_reviews) if clause_reviews else None

        # Add category as first level (if category exists and is not None)
        # Note: category is now handled at slide level, so this is only for individual calls
        if category:
            p_category = text_frame.add_paragraph()
            p_category.level = 0
            try:
                p_category.left_indent = Inches(0.21)
                p_category.first_line_indent = Inches(-0.19)
                p_category.space_before = Pt(0)
                p_category.space_after = Pt(0)
                p_category.line_spacing = 1.0
                _apply_east_asian_line_breaking(p_category)
            except:
                pass
            
            run_category = p_category.add_run()
            # The parallel header site in apply_structured_data_to_slides
            # translates; this one never did, despite taking
            # is_chinese_databook as a parameter. Only the presentation-table
            # renderer passes a real category here (the other caller passes
            # None and writes its own header), so the symptom was narrow and
            # easy to miss: a Chinese deck whose table pages alone carried an
            # English section heading -- "Expenses" sitting above 税金及附加
            # and 营业成本, while every non-table page read 流动资产/营业收入.
            run_category.text = (
                translate_category_to_chinese(category) if is_chinese_databook else category
            )
            run_category.font.size = Pt(font_size_pt)
            run_category.font.name = 'Arial'
            self._set_east_asian_typeface(run_category)
            self._declare_run_language(run_category)
            run_category.font.bold = False
            try:
                run_category.font.color.rgb = RGBColor(0, 51, 102)  # Dark blue
            except:
                pass
        
        # Add key name with grey char + space + key name (black bold) + "-" (not bold) + plain text
        p_key = text_frame.add_paragraph()
        p_key.level = 0  # No bullet level, we'll use grey character
        try:
            # Set formatting
            p_key.left_indent = Inches(0.15)  # 0.15" indent
            p_key.first_line_indent = Inches(-0.15)  # 0.15" special hanging
            p_key.space_before = Pt(0)
            p_key.space_after = Pt(3)  # Matches _PARA_SPACE_AFTER (cost estimator)
            p_key.line_spacing = 1.0
            _apply_east_asian_line_breaking(p_key)
        except Exception as e:
            logger.warning("Could not set paragraph formatting: %s", e)
            pass
        
        # Grey char (U+25A0) + space
        run_bullet = p_key.add_run()
        run_bullet.text = '\u25A0 '  # U+25A0 (black square) + space
        run_bullet.font.size = Pt(font_size_pt)
        run_bullet.font.name = 'Arial'
        self._set_east_asian_typeface(run_bullet)
        self._declare_run_language(run_bullet)
        run_bullet.font.bold = False
        try:
            run_bullet.font.color.rgb = RGBColor(128, 128, 128)  # Grey
        except:
            pass

        # Key name (black bold)
        run_key = p_key.add_run()
        run_key.text = display_name
        run_key.font.size = Pt(font_size_pt)
        run_key.font.name = 'Arial'
        self._set_east_asian_typeface(run_key)
        self._declare_run_language(run_key)
        run_key.font.bold = True
        try:
            run_key.font.color.rgb = RGBColor(0, 0, 0)  # Black
        except:
            pass

        # "-" (not bold)
        run_dash = p_key.add_run()
        run_dash.text = " - "
        run_dash.font.size = Pt(font_size_pt)
        run_dash.font.name = 'Arial'
        self._set_east_asian_typeface(run_dash)
        self._declare_run_language(run_dash)
        run_dash.font.bold = False
        try:
            run_dash.font.color.rgb = RGBColor(0, 0, 0)  # Black
        except:
            pass
        
        # Plain text (commentary content)
        commentary_lines = commentary.split('\n')
        first_line_added = False
        for line_idx, line in enumerate(commentary_lines):
            line = line.strip()
            if not line:
                continue

            if not first_line_added:
                target_paragraph = p_key
                first_line_added = True
            else:
                # Subsequent lines as new paragraphs (indented continuation)
                p_text = text_frame.add_paragraph()
                p_text.level = 0  # No bullet for continuation
                try:
                    p_text.left_indent = Inches(0.15)  # 0.15" indent (same as key text)
                    p_text.first_line_indent = Inches(0)  # No hanging for continuation lines
                    p_text.space_before = Pt(0)
                    p_text.space_after = Pt(3)
                    p_text.line_spacing = 1.0
                    _apply_east_asian_line_breaking(p_text)
                except:
                    pass
                target_paragraph = p_text

            self._add_runs_for_line(
                target_paragraph,
                line,
                clause_segments=clause_segments,
                font_size_pt=font_size_pt,
            )

        # Note: "(continued)" is now added to category header, not here
    
    



    


    # Chinese text shown in the Commentary label band in Chinese mode --
    # a single word, easy to change if the team prefers another term.
    _COMMENTARY_LABEL_ZH = "评述"

    def _localize_label_bands(self):
        """Translate the template's static English "Commentary" label band
        (Text-commentary / _L / _R, present on every slide) to Chinese
        when the deliverable is Chinese. English mode is left untouched.

        History -- why this TRANSLATES instead of deleting: ebf2179
        DELETED these shapes outright (the real reference deck has no
        such band at all), but the next two real Chinese exports both
        came back with the two statement-first pages (BS/IS page 1)
        rendering COMPLETELY BLANK in real PowerPoint, with abnormally
        slow file-open -- while python-pptx still saw every shape and
        all text intact in the same files, and the continuation pages
        kept rendering fine. Those two blank pages are exactly the two
        where the production template's band is a native PLACEHOLDER
        shape (the _L/_R variants on continuation pages are plain text
        boxes there); deleting a placeholder from a slide whose layout
        still defines it is the prime suspect for PowerPoint's renderer
        choking. Not reproducible on this Mac (no real PowerPoint, and
        the local template uses plain text boxes throughout), so the
        deletion was replaced by this strictly-safer in-place text swap:
        shape, placeholder wiring, fill and geometry all stay untouched
        -- only the literal run text changes, run-by-run so the run's
        own font formatting is preserved. If the blank-page symptom
        somehow survives even this, the one remaining suspect from the
        same commit on those two pages is the grid-table header restyle
        in _fill_table_placeholder.

        Matched by SHAPE NAME, at save time, so _calculate_table_bounds'
        text-based label-anchor fallback (matches the literal word
        "commentary" during layout) has long since finished by the time
        the text changes.
        """
        if not getattr(self, "_is_chinese_mode", False):
            return
        label_names = {"Text-commentary", "Text-commentary_L", "Text-commentary_R"}
        for slide in self.presentation.slides:
            for shape in slide.shapes:
                if _shape_name(shape) not in label_names:
                    continue
                if not getattr(shape, "has_text_frame", False):
                    continue
                try:
                    for para in shape.text_frame.paragraphs:
                        for run in para.runs:
                            if (run.text or "").strip().lower() == "commentary":
                                run.text = self._COMMENTARY_LABEL_ZH
                except Exception as exc:
                    logger.debug("Could not localize Commentary label band: %s", exc)

    def save(self, output_path: str):
        """Save the presentation"""
        if not self.presentation:
            raise ValueError("No presentation loaded")

        self._localize_label_bands()

        # Ensure output directory exists. dirname() is "" for a bare filename
        # in the current directory, and makedirs("") raises FileNotFoundError
        # -- so any caller passing "out.pptx" rather than a path crashed here
        # after doing all the generation work.
        out_dir = os.path.dirname(output_path)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        self.presentation.save(output_path)
        logger.info("Presentation saved to: %s", output_path)

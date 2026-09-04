"""Functions lifted out of PowerPointGenerator.

Each one read no instance state and called nothing that did, so `self`
was carrying nothing. Lifting them shortens the class without changing
what any of them computes.
"""
from __future__ import annotations

from __future__ import annotations
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

def find_shape_by_name(shapes, name: str):
    """Find shape by name in slide (case-insensitive), recursive"""
    name_lower = name.lower()
    for shape in shapes:
        if hasattr(shape, 'name') and (shape.name == name or shape.name.lower() == name_lower):
            return shape
        
        # Check for group
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            found = find_shape_by_name(shape.shapes, name)
            if found:
                return found
    return None


def _is_commentary_text_shape(shape) -> bool:
    if not getattr(shape, "has_text_frame", False):
        return False
    shape_name = str(getattr(shape, "name", "") or "").lower()
    excluded_tokens = (
        "title",
        "projtitle",
        "summary",
        "cosummaryshape",
        "table",
        "subtitle",
    )
    return not any(token in shape_name for token in excluded_tokens)


def _add_commentary_slot_shape(slide, slot_name: str):
    top = Inches(2.22)
    width = Inches(4.78)
    height = Inches(4.13)
    if slot_name == "L":
        left = Inches(0.13)
    elif slot_name == "R":
        left = Inches(5.09)
    else:
        # Page 1 template uses a single commentary box on the right beside the table.
        left = Inches(5.09)
    return slide.shapes.add_textbox(left, top, width, height)


def _prepare_structured_data_for_slides(structured_data: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    prepared: List[Dict[str, Any]] = []
    for account_data in structured_data or []:
        item = dict(account_data or {})
        commentary = _normalize_slide_commentary_text(item.get("commentary", ""))
        if commentary:
            item["original_commentary"] = commentary
        item["commentary"] = commentary  # Keep full length; fill optimizer handles fit
        prepared.append(item)
    return prepared


def _shape_name(shape) -> str:
    return str(getattr(shape, "name", "") or "")


def _shape_has_table(shape) -> bool:
    try:
        if getattr(shape, "has_table", False):
            return True
    except Exception:
        pass

    try:
        table = getattr(shape, "table", None)
        return table is not None
    except Exception:
        return False


def _resolve_table_target_shape(slide, statement_type: str):
    """Resolve the best existing target for a BS/IS table on a slide."""
    statement_type = (statement_type or "").upper()
    preferred_names = [
        "Table Placeholder",
        "Table Placeholder 2",
        "Content Placeholder 2",
    ]
    if statement_type == "IS":
        preferred_names.extend(["Table 3", "Table 2"])
    preferred_names.extend(["Table", "table", "TABLE"])

    for name in preferred_names:
        shape = find_shape_by_name(slide.shapes, name)
        if shape:
            return shape

    named_table_candidates = []
    table_candidates = []
    text_placeholder_candidates = []
    for shape in slide.shapes:
        shape_name = _shape_name(shape)
        shape_name_lower = shape_name.lower()
        if "table" in shape_name_lower and "placeholder" in shape_name_lower:
            text_placeholder_candidates.append(shape)
            continue
        if _shape_has_table(shape):
            table_candidates.append(shape)
            continue
        if "table" in shape_name_lower:
            named_table_candidates.append(shape)

    if text_placeholder_candidates:
        return text_placeholder_candidates[0]
    if table_candidates:
        return table_candidates[0]
    if named_table_candidates:
        return named_table_candidates[0]
    return None


def _read_table_style_id(tbl_element) -> Optional[str]:
    """Read <a:tableStyleId> (the style GUID) from a table's XML, or None."""
    try:
        from pptx.oxml.ns import qn
        tblPr = tbl_element.find(qn("a:tblPr"))
        if tblPr is None:
            return None
        el = tblPr.find(qn("a:tableStyleId"))
        return el.text.strip() if (el is not None and el.text) else None
    except Exception:
        return None


def _set_table_style_id(tbl_element, style_id: str) -> None:
    """Set the table's style GUID so PowerPoint renders it with that (e.g.
    UpSlide) table style instead of the python-pptx default."""
    from pptx.oxml.ns import qn
    tblPr = tbl_element.find(qn("a:tblPr"))
    if tblPr is None:
        tblPr = tbl_element.makeelement(qn("a:tblPr"), {})
        tbl_element.insert(0, tblPr)  # tblPr must be the first child of <a:tbl>
    for el in tblPr.findall(qn("a:tableStyleId")):
        tblPr.remove(el)
    style_el = tblPr.makeelement(qn("a:tableStyleId"), {})
    style_el.text = style_id
    tblPr.append(style_el)


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


_GRID_CELL_PADDING_PT = 0.08 * 72   # margin_left + margin_right, per _fill_table_placeholder


def _measure_grid_column_widths_pt(
    df, total_width_pt: float, *, data_font_pt: float, header_font_pt: float,
    is_chinese: bool, packing: Optional[Dict[str, Any]] = None,
) -> Optional[List[float]]:
    """Real per-glyph column widths for the BS/IS overview grid.

    The weight-and-character-count allocation below is the last place in the
    deck still sizing by len(): it counts a full-width 益 and a half-width 7 as
    one unit each. That is what left 归属于母公司所有者权益合计 needing 91.0pt
    in a column given 81.9pt -- the one wrap warning that survived every other
    measurement fix.

    Returns None when the client metrics are unavailable, so the caller keeps
    the old estimate rather than measuring with the wrong font.
    """
    try:
        from fdd_utils.text_metrics import get_measurer
    except Exception:
        return None
    packing = packing or {}
    metrics_path = _resolve_font_metrics_path(is_chinese, packing)
    if not metrics_path:
        return None
    try:
        family = _measurer_family(is_chinese, packing)
        m_data = get_measurer(family, data_font_pt, is_cjk=is_chinese, metrics_path=metrics_path)
        m_head = get_measurer(family, header_font_pt, is_cjk=is_chinese, metrics_path=metrics_path)
    except Exception:
        return None

    needs: List[float] = []
    for col_idx, col_name in enumerate(df.columns):
        # The header is measured AS RENDERED: _fill_table_placeholder writes
        # 2023年12月31日, not the ISO column name. Measuring "2023-12-31" makes
        # a date column look 10 narrow ASCII characters wide when it draws six
        # full-width glyphs, and the column then wraps.
        widest = m_head.text_width_pt(_format_header_period(str(col_name), is_chinese))
        for value in df.iloc[:, col_idx].tolist():
            text = "" if value is None else str(value)
            if not text or text == "nan":
                continue
            widest = max(widest, m_data.text_width_pt(text))
        # Leaf rows in column 0 carry a 0.12in indent (see _fill_table_placeholder).
        indent = (0.12 * 72) if col_idx == 0 else 0.0
        needs.append(widest + indent + _GRID_CELL_PADDING_PT)

    if not needs or total_width_pt <= 0:
        return None
    required = sum(needs)
    if required <= 0:
        return None
    if required <= total_width_pt:
        # Spare width is shared in proportion to what each column needs.
        # Giving it all to the row label was tried and is wrong: on real data
        # that column needs 87pt and would have taken 160, squeezing the date
        # columns to 39pt against the 54pt their header actually draws.
        needs = [w * total_width_pt / required for w in needs]
    else:
        # Cannot fit even measured -- scale everything down together, which is
        # still better than the count-based guess it replaces.
        needs = [w * total_width_pt / required for w in needs]
    return needs


def _fit_table_columns(table, df, *, data_font_pt: Optional[float] = None,
                       header_font_pt: Optional[float] = None,
                       is_chinese: Optional[bool] = None,
                       packing: Optional[Dict[str, Any]] = None):
    """Allocate width by role and content length for better readability."""
    if len(table.columns) == 0:
        return

    try:
        total_width = sum(col.width for col in table.columns)
    except Exception:
        total_width = 0
    if total_width <= 0:
        return

    # Measured where the caller knows the rendered font sizes and the client
    # metrics are present; the weight heuristic below is the fallback.
    if data_font_pt and header_font_pt:
        measured = _measure_grid_column_widths_pt(
            df, total_width / 12700.0,
            data_font_pt=data_font_pt, header_font_pt=header_font_pt,
            is_chinese=bool(is_chinese), packing=packing,
        )
        if measured:
            for col_idx, width_pt in enumerate(measured[: len(table.columns)]):
                table.columns[col_idx].width = max(int(Inches(0.35)), int(width_pt * 12700))
            return

    # A CJK character renders roughly 2x as wide as a Latin
    # character/digit at the same point size, but max_len here is a
    # raw character COUNT -- an 11-character Chinese row label like
    # "一年内到期的非流动负债" measures the same "length" as an
    # 11-character Latin one that's actually half as wide on the page.
    # Using the same /10 divisor for both meant max_len/10 almost never
    # exceeded the 2.0 floor for realistic Chinese labels (would need
    # 20+ characters), so column 0's weight was effectively a FIXED
    # 2.0 regardless of actual label length -- combined with the 0.12in
    # left-indent every leaf row gets, longer labels routinely wrapped
    # to 2 lines, each one rendering roughly 2x its neighbors' height
    # (PowerPoint auto-grows a row to fit wrapped text; the nominal
    # row.height set elsewhere is a floor, not a cap).
    # Measured over EVERY row, not a head(25) sample. The sample was silently
    # capping this: a real 33-row balance sheet carries its longest label,
    # 归属于母公司所有者权益合计 (13 chars), at row 30 -- outside the sample --
    # so max_len saw only 11 (一年内到期的非流动负债) and column 0 was sized 2
    # characters too narrow. Confirmed against the shipped client metrics:
    # the label needs 78.0pt at 6pt, the sampled weight gave 71.5pt of usable
    # cell, and PowerPoint auto-grew that row to fit the wrap.
    col0_series = df.iloc[:, 0].astype(str) if len(df.columns) else pd.Series(dtype=str)
    is_cjk_labels = any(
        any('一' <= ch <= '鿿' for ch in str(v)) for v in col0_series.tolist()
    )

    weights = []
    for col_idx, col_name in enumerate(df.columns[: len(table.columns)]):
        col_series = df.iloc[:, col_idx].astype(str) if col_idx < len(df.columns) else pd.Series(dtype=str)
        max_len = max([len(str(col_name))] + [len(val) for val in col_series.tolist()]) if len(col_series) else len(str(col_name))
        col_name_str = str(col_name).lower()
        if col_idx == 0:
            weight = (
                # Ceiling lowered 4.2 -> 3.0 at the same time as the head(25)
                # removal above, and for that reason: the sample used to cap
                # this weight by accident, and without a tighter ceiling a
                # single very long label could now claim width the DATE
                # columns need. These two trade directly against each other
                # (see the date-column comment below). At 3.0 a date column
                # still measures 56.1pt of usable cell against the 53.8pt
                # 2023年12月31日 needs at 7pt, and 3.0 still admits a 15-char
                # label -- longer than any real statement line seen so far.
                max(2.2, min(3.0, max_len / 5)) if is_cjk_labels
                else max(2.0, min(3.2, max_len / 10))
            )
        elif any(token in col_name_str for token in ["20", "19", "date", "年", "月"]):
            # The weight is computed from the ISO column NAME (2023-12-31, 10
            # chars) but a Chinese deck renders 2023年12月31日 -- same character
            # count, three of them full-width, so ~1.4x the drawn width. At the
            # old 1.4 floor that came to 53.0pt of usable cell against 57.7pt of
            # text, and the header wrapped to two lines, growing the row.
            weight = max(1.9, min(2.4, max_len / 8))
        else:
            weight = max(1.2, min(1.9, max_len / 9))
        weights.append(weight)

    total_weight = sum(weights) or 1
    assigned = 0
    for col_idx, weight in enumerate(weights):
        if col_idx == len(weights) - 1:
            width = total_width - assigned
        else:
            width = int(total_width * weight / total_weight)
            assigned += width
        table.columns[col_idx].width = max(int(Inches(0.7)), width)


def _format_table_value(value, is_numeric_column: bool) -> str:
    def _fmt_number(n: float) -> str:
        if n == 0:
            return "-"
        # Accounting convention: negatives in parentheses, not with a minus sign.
        return f"({abs(n):,.0f})" if n < 0 else f"{n:,.0f}"

    if pd.isna(value):
        return ""
    if isinstance(value, (int, float)) and is_numeric_column:
        return _fmt_number(float(value))

    text_val = str(value).strip()
    if is_numeric_column:
        numeric_candidate = text_val.replace(",", "")
        if re.fullmatch(r"-?\d+(\.\d+)?", numeric_candidate):
            try:
                return _fmt_number(float(numeric_candidate))
            except Exception:
                return text_val
    return text_val


def _set_paragraph_left_indent(
    paragraph, left_indent_emu: int, first_line_indent_emu: int = 0,
) -> None:
    """Set a paragraph's left indent (marL) and first-line offset (indent)
    directly on its <a:pPr> XML.

    first_line_indent_emu defaults to 0 (every line at marL). Pass a NEGATIVE
    value equal to -marL for a hanging indent: line 1 starts back at the left
    edge and every wrapped line sits at marL. That is the "■ key - text ..."
    bullet shape the reference deliverable uses, and the shape
    _calculate_content_lines has always measured (it gives paragraph 1 the full
    box width via first_line_width_pt and every later line
    _BULLET_HANGING_INDENT_PT less).

    _Paragraph has NO left_indent property in this python-pptx version
    (only alignment/level/line_spacing/font are exposed) -- `paragraph.
    left_indent = Inches(...)` doesn't raise, but that's because plain
    Python objects accept arbitrary ad-hoc attribute assignment; it
    silently creates a throwaway instance attribute with ZERO effect on
    the underlying XML, discarded the moment the object is garbage
    collected. Confirmed by round-tripping through a real save+reload:
    the "set" value reads back fine within the SAME Python session (the
    fake attribute is still sitting right there), but a freshly loaded
    Presentation() from that same saved file raises AttributeError on
    the same read -- proof nothing was ever written. marL/indent are
    real OOXML attributes on <a:pPr> (ECMA-376 CT_TextParagraphProperties)
    that python-pptx just doesn't wrap with a friendly property; setting
    them via the raw element (same get_or_add_pPr() pattern python-pptx's
    own oxml layer uses internally) is the only way that actually
    persists.
    """
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set('marL', str(int(left_indent_emu)))
    pPr.set('indent', str(int(first_line_indent_emu)))


def _sublist_text_for_table(table: Dict[str, Any], is_chinese_databook: bool, source_multiplier: float = 1,
    max_items: int = 5,
) -> str:
    """Converts a presentation_detail_table dict (extract_presentation_
    detail_table's return shape) into plain text lines for "sublist"
    style. Component lines show ONLY the LATEST period's figure -- a
    full table's worth of per-period, per-component detail written out
    as prose would be worse than the empty space this whole feature
    exists to fill. The total line keeps every period inline, matching
    how OI/OC accounts already state multi-year figures in this
    project's own reference style. Top-level rows only: nested children
    are already rolled into their own parent's total (e.g. 物业管理费's
    第三方/上海熙麦 sub-vendors), and a plain-text bullet list is not the
    place for two levels of indentation -- this already keeps
    same-nature items merged under their shared parent.

    When there are more than max_items top-level rows, only the
    max_items-1 largest (by the latest period's absolute value) are
    shown individually, ranked descending; everything past that is
    rolled into one final "其他"/"Other" line (summed, not dropped --
    the account's own real total still fully accounts for it even
    though this specific line doesn't itemise it). Keeps a long
    component list (e.g. 管理费用's 8 rows) from producing an equally
    long, table-like bullet list -- exactly what "sublist" style trades
    the native table's own per-component precision for.

    Values are in the same raw-yuan internal scale every account's df
    uses (see _render_presentation_table's own docstring) -- divided
    back down by source_multiplier here, at text-building time only,
    same as the native-table path does at render time (cadbce8)."""
    divisor = source_multiplier if source_multiplier and source_multiplier != 0 else 1

    def _scaled(v):
        return v / divisor if isinstance(v, (int, float)) else v

    periods = table.get("periods") or []
    period_labels = table.get("period_labels") or {}
    rows = table.get("rows") or []
    total_row = table.get("total_row") or {}
    if not periods or not rows:
        return ""

    latest_period = periods[-1]
    marker = "- "
    items: List[Tuple[str, float]] = []
    for row in rows:
        label = row.get("label", "")
        value = _scaled((row.get("values") or {}).get(latest_period))
        if value is None or not label:
            continue
        items.append((label, value))

    lines: List[str] = []
    if len(items) > max(1, max_items):
        ranked = sorted(items, key=lambda item: abs(item[1]), reverse=True)
        shown, rest = ranked[: max(1, max_items - 1)], ranked[max(1, max_items - 1):]
        for label, value in shown:
            lines.append(f"{marker}{label}：{_format_table_value(value, is_numeric_column=True)}")
        if rest:
            other_label = "其他" if is_chinese_databook else "Other"
            other_value = sum(v for _l, v in rest)
            lines.append(f"{marker}{other_label}：{_format_table_value(other_value, is_numeric_column=True)}")
    else:
        for label, value in items:
            lines.append(f"{marker}{label}：{_format_table_value(value, is_numeric_column=True)}")

    if total_row:
        total_label = total_row.get("label") or ("合计" if is_chinese_databook else "Total")
        total_values = total_row.get("values") or {}
        parts = []
        for period in periods:
            v = _scaled(total_values.get(period))
            if v is None:
                continue
            label = period_labels.get(period, period)
            text_val = _format_table_value(v, is_numeric_column=True)
            parts.append(f"{label}{text_val}" if is_chinese_databook else f"{text_val} in {label}")
        if parts:
            joiner = "，" if is_chinese_databook else ", "
            sep = "：" if is_chinese_databook else ": "
            lines.append(f"{marker}{total_label}{sep}{joiner.join(parts)}")

    return "\n".join(lines)


def _presentation_table_for_account(account_data: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    financial_data = (account_data or {}).get("financial_data")
    if not isinstance(financial_data, pd.DataFrame):
        return None
    try:
        table = (financial_data.attrs or {}).get("presentation_detail_table")
    except Exception:
        return None
    if not table or not table.get("rows"):
        return None
    return table


#: Accounts a due-diligence report breaks out as a matter of course, best
#: first. Used only to RANK once more accounts qualify than max_per_statement
#: allows -- it is not a whitelist, and an account absent from it still gets a
#: table when there is room (ranked by component count, i.e. by how much the
#: table shows that a sentence could not).
#:
#: 营业收入/营业成本/税金及附加/管理费用/销售费用/财务费用 lead because the
#: reader's question about an income statement is always "made of what", and
#: those breakdowns are the answer. 固定资产/应收账款/应付账款/预收款项 lead the
#: balance sheet for the same reason. Deliberately absent: 货币资金 (one bank
#: account), 长期借款 (one loan), 预付款项 (typically two small items) -- their
#: breakdown is a sentence, not a table.
_TABLE_PRIORITY_DEFAULT = (
    "OI", "OC", "Tax and Surcharges", "GA", "Selling cost", "Fin Exp",
    "NCA", "AR", "AP", "Advances", "OP", "OR", "Tax payable", "IA",
)

#: The income statement outranks the balance sheet for the deck's whole table
#: budget, on the user's call: "我認為是全deck5個而且優先是is 因為bs需要用到的
#: 比較少". A balance-sheet account's question is "how much", which its own
#: line in the embedded BS grid already answers; an income-statement account's
#: is "made of what", which only a breakdown answers.
#:
#: This is a strict order, not a weighting -- if the income statement alone
#: fills the budget the balance sheet gets NO table. That is the intended
#: reading of "全deck5個 + IS優先"; inspect_databook's 3c prints every
#: rejection, so a balance-sheet shutout is visible before the deck is built.
_STATEMENT_PRIORITY = ("IS", "BS")


def _select_presentation_tables(
    candidates: List[Tuple[Dict[str, Any], Dict[str, Any]]],
    settings: Optional[Dict[str, Any]] = None,
) -> Tuple[List[Tuple[Dict[str, Any], Dict[str, Any]]], List[Tuple[Dict[str, Any], str]]]:
    """Which table-bearing accounts actually get a native table.

    Every account that CAN have a table used to get one. A real entity
    export then put 17 tables on 6 slides -- five on one page -- and read as
    a data dump rather than a report. Two independent reasons an account
    should not spend a whole slot on a table:

    1. **The table says nothing the sentence didn't.** 管理费用's read
       "主要包括管理费用-行政及管理层调整。明细如下：" above a table whose
       two rows were 管理费用-行政 and 管理层调整. Same for 长期借款, 预付
       款项, 销售费用, 货币资金, 应付账款 -- six of the seventeen. A table
       earns its space by showing MORE than a sentence can carry, and a
       sentence carries about three components comfortably; hence min_rows.
    2. **There is only so much page.** Past a handful, tables stop being
       exhibits and become the deck. max_per_deck caps them; _STATEMENT_
       PRIORITY then _TABLE_PRIORITY_DEFAULT decide which survive.

    The cap is DECK-wide, so this has to be called once with both statements'
    accounts (see _select_deck_subtables) -- capping each statement separately
    doubles the real budget. Called with one statement's accounts it still
    works, it just cannot know what the other statement wants.

    A dropped account is NOT losing content: it returns to the ordinary
    commentary pool with its full text (the caller trims the now-dangling
    "明细如下：" handoff), which is exactly how it rendered before subtables
    were switched on at all.

    Returns (kept, dropped) where dropped carries a human-readable reason --
    inspect_databook.py's 3c prints it, so the deck's table count is always
    explainable without reading this code.
    """
    tables = (settings or {}).get("presentation_tables") or {}

    def _int_setting(name: str, default: int) -> int:
        try:
            value = int(tables.get(name, default))
        except (TypeError, ValueError):
            return default
        return value if value >= 0 else default

    min_rows = _int_setting("min_rows", 3)
    max_tables = _int_setting("max_per_deck", 5)
    priority = tables.get("priority")
    if not isinstance(priority, (list, tuple)) or not priority:
        priority = _TABLE_PRIORITY_DEFAULT
    rank_of = {str(key).strip().lower(): i for i, key in enumerate(priority)}

    kept: List[Tuple[Dict[str, Any], Dict[str, Any]]] = []
    dropped: List[Tuple[Dict[str, Any], str]] = []

    def _component_count(table: Dict[str, Any]) -> int:
        # Section labels (原值 / 累计折旧 / 净值) are structure, not components.
        # Counting them would let a two-item table clear min_rows on the
        # strength of its own headings.
        return len([r for r in (table.get("rows") or []) if not r.get("is_header")])

    substantial: List[Tuple[Dict[str, Any], Dict[str, Any]]] = []
    for item, table in candidates:
        n_rows = _component_count(table)
        if n_rows < min_rows:
            dropped.append((
                item,
                f"only {n_rows} component(s), below min_rows={min_rows} -- "
                f"the lead-in sentence already names them",
            ))
            continue
        substantial.append((item, table))

    if max_tables == 0:
        return [], dropped + [(item, "max_per_deck=0") for item, _ in substantial]

    def _sort_key(pair: Tuple[Dict[str, Any], Dict[str, Any]]) -> Tuple[int, int, int, str]:
        item, table = pair
        key = str(item.get("mapping_key") or item.get("account_name") or "").strip().lower()
        statement = str(item.get("_statement_type") or "").strip().upper()
        statement_rank = (_STATEMENT_PRIORITY.index(statement)
                          if statement in _STATEMENT_PRIORITY else len(_STATEMENT_PRIORITY))
        # Absent from the priority list is not a demerit against a LOWER-
        # ranked listed account, only against the listed ones -- so unlisted
        # accounts sort after the list and among themselves by how much their
        # table actually shows.
        return (statement_rank, rank_of.get(key, len(rank_of)),
                -_component_count(table), key)

    ordered = sorted(substantial, key=_sort_key)
    kept = ordered[:max_tables]
    for item, table in ordered[max_tables:]:
        statement = str(item.get("_statement_type") or "").strip().upper()
        note = f" -- {statement} ranks after {_STATEMENT_PRIORITY[0]}" if statement == "BS" else ""
        dropped.append((
            item,
            f"over max_per_deck={max_tables} "
            f"({_component_count(table)} component(s), ranked below the ones kept){note}",
        ))

    # Restore the statement's own reading order: the ranking decides WHICH
    # tables survive, never where they land. _append_table_accounts_to_
    # distribution flows them sequentially and relies on that order.
    position = {id(item): i for i, (item, _t) in enumerate(candidates)}
    kept.sort(key=lambda pair: position.get(id(pair[0]), 0))
    return kept, dropped


#: Set on every account the deck-level pass has ruled on. Its ABSENCE is what
#: tells _plan_slot_distribution no such pass ran (a single-statement caller,
#: or a diagnostic), so it can fall back to deciding for itself.
_SUBTABLE_DECISION_KEY = "_subtable_ok"


def _select_deck_subtables(
    statements: List[Tuple[str, List[Dict[str, Any]]]],
    settings: Optional[Dict[str, Any]] = None,
) -> List[Tuple[Dict[str, Any], str]]:
    """Rules on the WHOLE deck's tables at once, tagging each account.

    max_per_deck cannot be enforced from inside _plan_slot_distribution: that
    runs once per statement, so a cap applied there is a cap per statement and
    the deck gets twice what was asked for. Worse, the balance sheet is
    planned first (slide 1 before slide 5), so a naive running budget would be
    spent before the income statement was considered at all -- the exact
    opposite of _STATEMENT_PRIORITY.

    So the decision is made HERE, before any planning, where both statements
    are visible at once. Each account dict is tagged in place; the tag
    survives _prepare_structured_data_for_slides' shallow dict() copy.

    Returns the rejections, for the caller to log.
    """
    candidates: List[Tuple[Dict[str, Any], Dict[str, Any]]] = []
    for statement_type, items in statements:
        for item in items or []:
            table = _presentation_table_for_account(item)
            if table is None:
                continue
            item["_statement_type"] = statement_type
            candidates.append((item, table))

    kept, dropped = _select_presentation_tables(candidates, settings)
    keep_ids = {id(item) for item, _table in kept}
    for item, _table in candidates:
        item[_SUBTABLE_DECISION_KEY] = id(item) in keep_ids
    return dropped


def _strip_table_handoff(text: str, handoff: str) -> str:
    """Removes a "明细如下：" whose table is gone, wherever in the text it sits.

    The phrase is a promise the model was told to make (see
    _detail_table_guidance); left pointing at nothing it reads as a rendering
    failure, which is how the missing tables were reported in the first place.
    It was only stripped when it TRAILED, on the reasoning that the usual
    shape -- phrase, then "➢" explanation bullets -- reads fine. Two real
    bullets say otherwise:

        管理费用主要包括管理费用-行政及管理层调整。明细如下：
        管理费用-行政于2026年1-6月期间发生15.5万元...        <- no marker at all
        销售费用主要包括...和管理层调整，明细如下：销售费用-代理及佣金...  <- same line

    Neither is a list. Both promise 明细 and hand over a sentence.

    A LONE "➢" line under the phrase is the same complaint from the other
    side -- a one-item list, flagged by the user as forced point form -- so it
    is folded back into the paragraph. Two or more are left alone: that is
    real point form and reads as intended.
    """
    body = (text or "").strip()
    if not body or not handoff:
        return body
    idx = body.lower().find(handoff.lower())
    if idx < 0:
        return body

    head = body[:idx].rstrip()
    tail = body[idx + len(handoff):].strip()

    # "...主要包括 A 及 B。" already ends properly; "...和管理层调整，" does not.
    if head and head[-1] in "，,、；;：:":
        head = head[:-1].rstrip()
    is_chinese = any("\u4e00" <= ch <= "\u9fff" for ch in head)
    if head and head[-1] not in "。.！!？?":
        head += "。" if is_chinese else "."
    if not tail:
        return head

    lines = [ln.strip() for ln in tail.split("\n") if ln.strip()]
    marked = [ln for ln in lines if ln.startswith(("\u27a2", "\u2023", "\u2022", "-"))]
    if len(lines) == 1 and len(marked) <= 1:
        only = lines[0].lstrip("\u27a2\u2023\u2022- ").strip()
        return f"{head}{only}" if is_chinese else f"{head} {only}"
    return head + "\n" + "\n".join(lines)


def _truncate_text_at_boundary(text: str, limit: int, is_chinese: bool) -> str:
    """Cuts `text` to at most `limit` chars at a sentence boundary where
    possible. Shared by the lead-in and the post-table explanation --
    same safety-net shape, different caps (see _split_table_commentary)."""
    text = (text or "").strip()
    if len(text) <= limit:
        return text
    boundary_chars = "。；;.!?！？"
    cut = text[:limit]
    # A "." between two digits is a DECIMAL POINT, not a sentence end.
    # Taking it as one truncated a real deck's 营业成本 lead-in at
    # "...较2025年度下降74." -- the last "boundary" in the string was the
    # point inside 74.9%, so the figure was cut in half and the rest of
    # the sentence, including the "明细如下：" handoff, was thrown away.
    # Same defect class as the mid-number split _snap_split_before_number
    # already guards in the packing path; this truncation path never had
    # the guard.
    best = -1
    for pos in range(len(cut) - 1, -1, -1):
        ch = cut[pos]
        if ch not in boundary_chars:
            continue
        if (
            ch == "."
            and pos > 0
            and cut[pos - 1].isdigit()
            and pos + 1 < len(text)
            and text[pos + 1].isdigit()
        ):
            continue
        best = pos
        break
    if best >= int(limit * 0.4):
        return cut[: best + 1]
    return cut.rstrip() + ("…" if is_chinese else "...")


def _planning_std_lh_pt(is_chinese: bool) -> float:
    """One std_lh unit as RENDER actually produces it, for the
    shape-less planning estimates.

    The estimates used to compute font_size x line_spacing + para_gap
    (9 x 1.0 + 2.2 = 11.2pt), but PowerPoint's real line pitch is
    1.2 x the point size (POWERPOINT_LINE_PITCH_FACTOR -- researched
    and confirmed separately, see project memory), so render uses
    10.8 + 2.2 = 13.0pt. Planning was therefore under-estimating every
    block by ~14%, which _TABLE_SLOT_PACK_THRESHOLD was quietly
    absorbing -- two compensating errors that together left real
    columns filled to only ~60% of their true capacity."""
    from fdd_utils.text_metrics import POWERPOINT_LINE_PITCH_FACTOR
    return (_real_font_size_pt(is_chinese) * POWERPOINT_LINE_PITCH_FACTOR
            * _real_line_spacing(is_chinese) + _real_para_gap_pt(is_chinese))


def _rendered_bullet_label(account_data: Dict[str, Any], is_chinese_databook: bool) -> str:
    """The label a bullet ACTUALLY renders with ("■ <label> - ...").

    Cost estimates must measure this, not the raw mapping_key: in a
    Chinese deck the mapping_key is the English short code
    ("Tax and Surcharges"), which is far wider than the Chinese name
    that really renders ("税金及附加") -- 352pt vs 315pt against a
    329.8pt box for one real lead-in, i.e. the estimate believed the
    line wrapped when it doesn't. Every such lead-in box came out one
    whole line too tall, which is exactly the "height 似乎是固定的...
    表格不是緊貼comments" the user reported."""
    mapping_key = account_data.get("mapping_key", account_data.get("account_name", ""))
    if is_chinese_databook:
        return account_data.get("display_name_zh") or account_data.get("display_name", mapping_key)
    return account_data.get("display_name", mapping_key)


def _textbox_usable_and_inset_pt(shape) -> Tuple[float, float]:
    """(usable text height, total top+bottom inset) in points for a
    shape, read from its real bodyPr insets. Falls back to the OOXML
    default when they aren't declared."""
    raw_pt = int(shape.height) / 12700
    try:
        from fdd_utils.text_metrics import text_box_from_shape
        usable_pt = text_box_from_shape(shape).height_pt
    except Exception:
        usable_pt = max(1.0, raw_pt - PowerPointGenerator._TEXTBOX_INSET_PT)
    return usable_pt, max(0.0, raw_pt - usable_pt)


def _table_unit_label(is_chinese_databook: bool) -> str:
    return "人民币千元" if is_chinese_databook else "CNY'000"


def _table_source_multiplier(account_data: Dict[str, Any]) -> float:
    """The account's own raw-yuan -> display-unit divisor. Single
    definition, since the renderer, the width precompute and the
    AI-prompt side all need the identical value (an earlier 1000x
    display bug came from exactly this being derived twice)."""
    financial_data = (account_data or {}).get("financial_data")
    if hasattr(financial_data, "attrs"):
        return financial_data.attrs.get("source_multiplier") or 1
    return 1


#: Splits "固定资产-房屋建筑物" into its category and its item. Fullwidth and
#: em dashes included because the source workbooks use all three.
_LABEL_CATEGORY_SPLIT = re.compile(r"\s*[-－—]\s*")


def _group_plan_rows_by_category(
    entries: List[Dict[str, Any]], title: str = "", depth: int = 0,
) -> List[Dict[str, Any]]:
    """Turns a flat run of "CATEGORY-item" rows into a heading plus its items.

    A 固定资产 table read as thirteen undifferentiated lines -- 固定资产-房屋
    建筑物, 固定资产-机械设备, ..., 累计折旧-房屋建筑物, ... -- which is the
    schedule's own row order and tells the reader nothing about which block
    they are in. The grouping is already IN the labels, so no interpretation
    is needed to recover it:

        固定资产                      累计折旧
          房屋建筑物  225,000    ->     房屋建筑物  -50,000
          机械设备     17,000           机械设备    -8,000

    Shortening each member to the part after the category is the other half
    of the win: the repeated prefix was eating the label column, which is the
    column that wraps.

    Conservative by construction:
      - only a run of 2+ CONSECUTIVE rows sharing a category groups, so
        nothing is reordered and a lone 管理层调整 stays a plain row;
      - if every row lands in ONE group the grouping is dropped -- a single
        heading over the whole table costs a row and says nothing;
      - a table whose rows already carry explicit children is left alone
        rather than structured twice.
    """
    if not entries or any(e.get("kind") != "data" for e in entries):
        return entries

    def _category(label: str) -> Optional[str]:
        parts = _LABEL_CATEGORY_SPLIT.split(str(label or "").strip(), maxsplit=1)
        if len(parts) != 2:
            return None
        head, tail = parts[0].strip(), parts[1].strip()
        return head if head and tail else None

    runs: List[Tuple[Optional[str], List[Dict[str, Any]]]] = []
    for entry in entries:
        category = _category(entry["label"])
        if runs and runs[-1][0] == category and category is not None:
            runs[-1][1].append(entry)
        else:
            runs.append((category, [entry]))

    grouped_runs = [r for r in runs if r[0] is not None and len(r[1]) >= 2]
    if not grouped_runs:
        return entries

    own = str(title or "").strip()
    # One category covering every row earns no heading -- a heading over the
    # whole table says what the row beneath it already says -- but those rows
    # still shed the prefix, which is the part that makes the label column
    # wrap. Returning untouched instead kept 应缴税费- on every row of an
    # 应交税费 table, and the bank-account code prefix on every row of a
    # 货币资金 one.
    covers_everything = (len(grouped_runs) == 1
                         and len(grouped_runs[0][1]) == len(entries))
    out: List[Dict[str, Any]] = []
    for category, members in runs:
        if category is None or len(members) < 2:
            out.extend(members)
            continue
        # A heading that repeats the table's own title band says it twice --
        # the 税金及附加 table would read 税金及附加 / 税金及附加 / 房产税 / ...
        # Its members still lose the repeated prefix, they simply stay at top
        # level, which is what "belongs to the account named above" looks like.
        # A category that is NOT the title (累计折旧 under 固定资产) still gets
        # its heading, because there it is carrying real information.
        heading = category != own and not covers_everything
        if heading:
            out.append({"label": category, "values": {}, "kind": "group"})
        stripped = [
            {**member,
             "label": _LABEL_CATEGORY_SPLIT.split(member["label"], maxsplit=1)[1].strip(),
             "kind": "grouped" if heading else "data"}
            for member in members
        ]
        if heading or depth:
            out.extend(stripped)
        else:
            # Stripping one level only moved the repetition down a level: an
            # 管理费用 table came out as 行政-保险费 / 行政-审计费 / 行政-AMF /
            # 员工-工作餐费 ..., with 行政- on ten of fourteen rows. No heading
            # was added here, so a heading found one level down is the FIRST
            # visible one and cannot read as a peer of an outer level. Only
            # that case recurses, which keeps the table at a single heading
            # level however deep the source account path runs.
            out.extend(_group_plan_rows_by_category(stripped, "", depth + 1))
    return out


def _name_net_block(plan: List[Dict[str, Any]], total_label: Any,
                    is_chinese_databook: bool) -> List[Dict[str, Any]]:
    """Names the third block of a roll forward, when the arithmetic says it is
    the net one.

    A fixed-asset schedule is laid out in three tiers -- cost, accumulated
    depreciation, net -- but only the middle one carries a label of its own,
    because 累计折旧 is the only row the sheet writes with no figures. So the
    deck printed 房屋建筑物 / 机械设备 / 室外工程, then [累计折旧] and the same
    three names again, then the same three a THIRD time with nothing to say
    the third block is the carrying amount:

        房屋建筑物   281,740
        [累计折旧]
            房屋建筑物   (38,032)
        房屋建筑物   243,709      <- net, and the reader is not told

    Nothing here is inferred from wording. The third block is only named if
    every component in it equals its own cost plus its own depreciation, which
    is what "net" means; one row that does not tie and the plan is returned
    untouched. The name comes from the table's own total row (净值小计 ->
    净值) and falls back to 净值 only once the arithmetic has already proved
    what the block is.
    """
    groups = [i for i, entry in enumerate(plan) if entry["kind"] == "group"]
    if len(groups) != 1:
        return plan
    head = groups[0]
    before, after = plan[:head], plan[head + 1:]
    if len(before) < 2 or any(e["kind"] not in ("data", "grouped") for e in before + after):
        return plan
    # The first block's opening label comes round twice after the heading:
    # once opening the depreciation block, once opening the third one.
    starts = [i for i, entry in enumerate(after) if entry["label"] == before[0]["label"]]
    if len(starts) != 2 or starts[0] != 0:
        return plan
    mid, third = after[:starts[1]], after[starts[1]:]
    if len(mid) < 2 or len(third) < 2:
        return plan

    cost = {e["label"]: (e["values"] or {}) for e in before}
    depreciation = {e["label"]: (e["values"] or {}) for e in mid}
    ties = 0
    for entry in third:
        gross, dep = cost.get(entry["label"]), depreciation.get(entry["label"])
        if gross is None or dep is None:
            continue                      # e.g. a trailing 管理层调整
        for period, net in (entry["values"] or {}).items():
            g, d = gross.get(period), dep.get(period)
            if not all(isinstance(v, (int, float)) for v in (net, g, d)):
                continue
            if abs(net - (g + d)) > max(1.0, abs(net) * 0.005):
                return plan               # does not tie -- not a net block
            ties += 1
    if ties < 2:
        return plan

    label = str(total_label or "").strip()
    for suffix in ("小计", "合计", "总计"):
        if label.endswith(suffix) and len(label) > len(suffix):
            label = label[: -len(suffix)].strip()
            break
    else:
        label = ""
    if not label:
        label = "净值" if is_chinese_databook else "Net book value"
    return (plan[:head + 1 + len(mid)]
            + [{"label": label, "values": {}, "kind": "group"}]
            + [{**entry, "kind": "grouped"} for entry in third])


def _build_presentation_table_plan(table: Dict[str, Any], is_chinese_databook: bool, source_multiplier: float,
) -> List[Dict[str, Any]]:
    """Flattens a presentation table's rows -> children (indented) ->
    total into the single ordered render plan both the renderer and
    the uniform-width precompute measure against. Values are divided
    back down to display units here (see _render_presentation_table's
    docstring for why that division belongs at display time)."""
    divisor = source_multiplier if source_multiplier and source_multiplier != 0 else 1

    def _scaled(values: Dict[str, float]) -> Dict[str, float]:
        return {period: (v / divisor if isinstance(v, (int, float)) else v)
                for period, v in (values or {}).items()}

    rows = table.get("rows") or []
    # The schedule's OWN section labels (原值 / 累计折旧 / 净值), carried through
    # as value-less rows. When they are present they are the authority on this
    # table's structure and the label-prefix heuristic must stand down --
    # running both would head the same block twice.
    has_sheet_headings = any(row.get("is_header") for row in rows)

    plan: List[Dict[str, Any]] = []
    for row in rows:
        if row.get("is_header"):
            plan.append({"label": row.get("label", ""), "values": {}, "kind": "group"})
            continue
        kind = "grouped" if (has_sheet_headings and any(
            e["kind"] == "group" for e in plan)) else "data"
        plan.append({"label": row.get("label", ""), "values": _scaled(row.get("values")), "kind": kind})
        for child in (row.get("children") or []):
            plan.append({"label": child.get("label", ""), "values": _scaled(child.get("values")), "kind": "child"})
    if not has_sheet_headings:
        plan = _group_plan_rows_by_category(plan, str(table.get("title") or ""))
    total_row = table.get("total_row")
    plan = _name_net_block(plan, (total_row or {}).get("label"), is_chinese_databook)
    if total_row:
        plan.append({"label": total_row.get("label", "合计" if is_chinese_databook else "Total"),
                     "values": _scaled(total_row.get("values")), "kind": "total"})
    return plan


def _explanation_render_text(post_table_text: str, is_chinese_databook: bool) -> str:
    """The post-table explanation exactly as it RENDERS -- one marker-
    prefixed line per source line. Single definition so the planner
    measures the same string the renderer writes; the two-character
    "➢ " prefix is worth a whole wrapped line on a full-width line."""
    marker = "➢ " if is_chinese_databook else "- "
    raw_lines = [ln.strip() for ln in (post_table_text or "").split("\n") if ln.strip()]
    if not raw_lines:
        raw_lines = [(post_table_text or "").strip()]
    return "\n".join(
        ln if ln.startswith(("➢", "-", "•")) else f"{marker}{ln}" for ln in raw_lines
    )


# A section heading is not the same word as the account it introduces. The
# shared category translation renders Revenue as 营业收入 -- correct as a label
# for the revenue LINE, but as a section heading it repeats the 营业收入 row
# directly beneath it. Suppressing the duplicate left the income statement
# opening with no heading at all while the expense block still had one, which
# read as lopsided. The statement wants the broader word for the section.
_GRID_SECTION_HEADING_ZH = {
    "Revenue": "收入",
    "Operating revenue": "收入",
    "Operating Revenue": "收入",
}


def _grid_section_heading(category: str, is_chinese_mode: bool) -> str:
    if not is_chinese_mode:
        return category
    if category in _GRID_SECTION_HEADING_ZH:
        return _GRID_SECTION_HEADING_ZH[category]
    return translate_category_to_chinese(category)


def _insert_category_header_rows(df, mappings: Optional[Dict[str, Any]], is_chinese_mode: bool):
    """Insert a blank-figures header row ("流动资产" / "Current assets"
    / etc.) into `df` whenever a leaf line item's mapped category
    (mappings.yml -- the SAME per-account "category" field the
    commentary bullets already group by) changes from the previous
    one. Total/subtotal rows (same keyword detection the later styling
    pass uses) never update the running category tracker and never
    trigger an insertion themselves -- a subtotal belongs to whatever
    category the items above it were in, not a category of its own.

    A real Financials-sheet check this session (inspect_financials_
    structure.py against the Kunshan databook) confirmed the RAW
    extracted sheet has no such header rows at all -- straight from a
    leaf item to "Total current assets" -- so this is what actually
    produces the reference format's ("IMG_0035") header rows, since
    nothing upstream of this table provides them on its own.

    Returns `df` unchanged if there's no mappings to categorise
    against (never silently drops rows in that case).
    """
    if not mappings or df is None or df.empty:
        return df

    total_keywords = list(
        {'total', '合计', '总计', '小计', 'subtotal', 'sub-total', 'sub total'}
        | set(SUMMARY_ACCOUNT_SKIP_KEYWORDS)
    )

    new_rows = []
    current_category = None
    seen_categories: set[str] = set()
    for _, row in df.iterrows():
        label = str(row.iloc[0]).strip()
        label_lower = label.lower()
        is_total = any(kw in label_lower for kw in total_keywords)

        if not is_total and label:
            mapping_key = find_mapping_key(label, mappings)
            category = str((mappings.get(mapping_key) or {}).get('category', '') or '') if mapping_key else ''
            if category and category != current_category:
                header_label = _grid_section_heading(category, is_chinese_mode)
                # A category is opened ONCE. A balance sheet's categories are
                # contiguous so this never mattered there, but an income
                # statement runs in statement order and its categories
                # interleave -- Revenue, Expenses, ... then a single Revenue
                # line (投资收益) partway down, then Expenses again. Re-opening
                # on every change reprinted "营业收入" and "费用" mid-statement,
                # which reads as a second, restarted section rather than as
                # the running P&L it is.
                #
                # A header identical to the row it introduces is also dropped:
                # the Revenue category translates to the same words as the
                # 营业收入 account line itself, so the two rendered as a
                # duplicated label one above the other.
                duplicates_next = header_label.strip() == label
                if header_label not in seen_categories and not duplicates_next:
                    header_row = {col: (header_label if i == 0 else pd.NA) for i, col in enumerate(df.columns)}
                    new_rows.append(header_row)
                # Marked as opened even when the header was suppressed for
                # duplicating the row below it: that row IS the heading, in
                # the reader's eyes. Without this the category counts as
                # unopened and re-opens later -- which is how "营业收入"
                # reappeared further down the income statement.
                seen_categories.add(header_label)
                current_category = category

        new_rows.append(row.to_dict())

    return pd.DataFrame(new_rows, columns=df.columns)


_LEAD_PROMISES_TABLE_RE = re.compile(
    r"(?:明细|明細|构成|構成|情况|情況|分析|列示|详情|詳情)?\s*"
    r"(?:如下|见下表|見下表|列示如下|详见下表|詳見下表)\s*[:：]?\s*$"
)


def lead_promises_table(text: str) -> bool:
    """Does this lead-in END by announcing the table that follows it?

    "…明细如下：" only makes sense with the detail directly beneath it. A
    column that ends on that phrase with nothing under it reads as broken,
    which is why splitting a lead from its own table was removed once before
    (see _append_table_accounts_to_distribution's flow()).

    A lead that instead stands on its own -- a complete statement of the
    balance and its drivers, with the table there to support rather than
    complete it -- can be separated, because the sentence is finished either
    way. This is the test that tells the two apart.
    """
    body = str(text or "").strip()
    if not body:
        return False
    # Only the final clause matters: an earlier "如下" inside a long paragraph
    # is not what the reader is left hanging on.
    tail = re.split(r"[。；;\n]", body)[-1].strip() or body
    return bool(_LEAD_PROMISES_TABLE_RE.search(tail))


def find_content_shape(shapes):
    """Find content shape by trying multiple possible names"""
    # Try different possible names for content shapes
    possible_names = [
        'Content',
        'Text-commentary',
        'textMainBullets',
        'Text',
        'Commentary',
        'MainContent',
        'Body'
    ]
    
    for name in possible_names:
        shape = find_shape_by_name(shapes, name)
        if shape and shape.has_text_frame:
            return shape
    
    # If no named shape found, try to find any text frame shape that's not a title
    for shape in shapes:
        if hasattr(shape, 'has_text_frame') and shape.has_text_frame:
            shape_name = getattr(shape, 'name', '')
            # Skip title shapes and other non-content shapes
            if shape_name and 'title' not in shape_name.lower() and 'proj' not in shape_name.lower():
                return shape
    
    return None


def replace_text_preserve_formatting(shape, replacements: Dict[str, str]) -> bool:
    """Replace text while preserving formatting"""
    if not shape.has_text_frame:
        return False

    replaced = False

    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            original_text = run.text
            for old_text, new_text in replacements.items():
                if old_text in run.text:
                    run.text = run.text.replace(old_text, new_text)
            if run.text != original_text:
                replaced = True

    if not replaced:
        current_text = shape.text_frame.text
        updated_text = current_text
        for old_text, new_text in replacements.items():
            updated_text = updated_text.replace(old_text, new_text)
        if updated_text != current_text:
            shape.text_frame.text = updated_text
            replaced = True

    return replaced


def _process_markdown_content(content: str) -> Dict:
    """Process markdown content into structured data"""
    if not content:
        logger.warning("Empty content provided to _process_markdown_content")
        return {}

    logger.info("Processing markdown content, length: %s", len(content))
    logger.debug("Content preview (first 500 chars): %s", content[:500])

    # Split by headers (## Account Name)
    sections = re.split(r'^##\s+(.+)$', content, flags=re.MULTILINE)

    logger.info("Found %s sections after splitting", len(sections))

    processed_sections = {}

    # Process each section
    for i in range(1, len(sections), 2):
        if i + 1 < len(sections):
            account_name = sections[i].strip()
            account_content = sections[i + 1].strip()

            logger.info("Processing section: %s, content length: %s", account_name, len(account_content))

            processed_sections[account_name] = {
                'content': account_content,
                'is_chinese': detect_chinese_text(account_content)
            }

    logger.info("Processed %s sections", len(processed_sections))
    return processed_sections


def _fill_content_shape(shape, section_data: Dict):
    """Fill content shape with processed data"""
    if not shape.has_text_frame:
        logger.warning("Shape does not have text_frame")
        return

    content = section_data.get('content', '')
    is_chinese = section_data.get('is_chinese', False)

    logger.info("Filling shape with content length: %s", len(content))

    # Clear existing content
    shape.text_frame.clear()
    
    if not content or not content.strip():
        logger.warning("No content to fill")
        return
    
    # Split content into paragraphs if it contains newlines
    content_lines = content.split('\n')
    
    # Add content with proper formatting
    for idx, line in enumerate(content_lines):
        line = line.strip()
        if not line and idx > 0:
            # Skip empty lines except add a paragraph break
            continue
        
        if idx == 0:
            # Use first paragraph or create one
            if shape.text_frame.paragraphs:
                p = shape.text_frame.paragraphs[0]
            else:
                p = shape.text_frame.add_paragraph()
        else:
            p = shape.text_frame.add_paragraph()
        
        p.text = line
        
        # Apply formatting to runs
        for run in p.runs:
            run.font.size = get_font_size_for_text(line, force_chinese_mode=is_chinese)
            run.font.name = get_font_name_for_text(line)

        # Set paragraph formatting
        p.space_after = get_space_after_for_text(line, force_chinese_mode=is_chinese)
        p.space_before = get_space_before_for_text(line, force_chinese_mode=is_chinese)
        p.line_spacing = get_line_spacing_for_text(line, force_chinese_mode=is_chinese)
    
    logger.info("Successfully filled shape with %s paragraphs", len([l for l in content_lines if l.strip()]))


def font_metrics_filename(is_chinese: bool) -> str:
    return "msyh_chi.json" if is_chinese else "arial_eng.json"


def font_metrics_candidates(is_chinese: bool, configured: Optional[str] = None) -> List[str]:
    """Every place the client-font metrics may live, best first.

    The files live INSIDE the package (fdd_utils/font_metrics/) so a deployment
    that ships only fdd_utils/ and fdd_app.py carries them. The repo root is
    still searched after it, because that is where they used to sit and an
    older checkout or an absolute config entry should keep working.

    Searching both rather than picking one on purpose: the defaults and the
    real location disagreeing is exactly what broke this before (e732e2d) --
    the default pointed at fdd_utils/font_metrics/ while the files were at the
    root, so every relative path missed and the exporter silently measured with
    a system font while the checker used the client's.
    """
    name = font_metrics_filename(is_chinese)
    # fdd_utils/pptx/helpers.py -> fdd_utils/ -> repo root
    _pkg = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    _root = os.path.dirname(_pkg)
    out: List[str] = []
    for raw in (configured, os.path.join("font_metrics", name)):
        if not raw:
            continue
        p = str(raw).strip()
        if not p:
            continue
        if os.path.isabs(p):
            out.append(p)
        else:
            out.extend([os.path.join(_pkg, p), os.path.join(_root, p)])
    return out


def _resolve_font_metrics_path(is_chinese: bool, packing: Dict[str, Any]) -> Optional[str]:
    """Path to the client-font metrics.json (dumped via dump_font_metrics.py),
    so line-fitting measures with the font the client's PowerPoint renders.
    Language-specific key wins, then a single shared key, then the shipped
    default -- each tried inside the package first and at the repo root
    second (see font_metrics_candidates)."""
    key = "font_metrics_path_chi" if is_chinese else "font_metrics_path_eng"
    configured = packing.get(key) or packing.get("font_metrics_path")
    for candidate in font_metrics_candidates(is_chinese, configured):
        if os.path.exists(candidate):
            return candidate
    return None


def _measurer_family(is_chinese: bool, packing: Dict[str, Any]) -> str:
    """System-font family for the Pillow fallback (overridable in config)."""
    key = "font_family_chi" if is_chinese else "font_family_eng"
    return str(packing.get(key) or ("Microsoft YaHei" if is_chinese else "Arial"))


def _real_font_size_pt(is_chinese: bool) -> float:
    """Font size actually applied to the run (get_font_size_for_text) —
    a single deck-wide 9pt regardless of language, NOT the 10pt some
    capacity/content code used to assume for Chinese."""
    return get_font_size_for_text("", force_chinese_mode=is_chinese).pt


def _real_line_spacing(is_chinese: bool) -> float:
    """Line spacing actually applied to a commentary bullet run.

    _fill_text_main_bullets_with_category_and_key hardcodes
    line_spacing = 1.0 on every paragraph it creates (category header,
    key line, and continuation lines alike) -- unconditionally, not
    gated on is_chinese at all. get_line_spacing_for_text's 0.9-for-
    Chinese value belongs to the separate, legacy _fill_content_shape
    path (markdown generate() flow) and was never actually the value
    applied to a live textMainBullets paragraph. A user-supplied real-
    client-metrics capacity check (inspect_single_slot.py against a
    real Windows export) directly caught this: assuming 0.9 line
    spacing + a 6-9pt inter-paragraph gap that never actually renders
    made the computed capacity roughly 30% smaller than the box's true
    capacity -- "the tool says 94% full" against a box the user could
    still visibly type 5-7 more lines into.
    """
    return 1.0


def _real_para_gap_pt(is_chinese: bool) -> float:
    """Total vertical gap PowerPoint actually renders between two
    consecutive bullet paragraphs.

    _fill_text_main_bullets_with_category_and_key hardcodes
    space_before = Pt(0) and space_after = Pt(3) on every paragraph
    (category header, key line, continuation line) -- 3pt total,
    REGARDLESS of language. It never calls get_space_after_for_text /
    get_space_before_for_text at all; those getters' 4-9pt values
    belong to the separate legacy _fill_content_shape path. See
    _real_line_spacing's docstring for how this was actually caught.

    2.2, not 3.0 (2026-08-04): the requested space_after XML value is
    still literally Pt(3) at render time, unchanged -- this is not a
    claim that PowerPoint renders less than what's asked for. It's a
    correction to how much of that requested space this codebase's
    OWN capacity/content-cost formula should count against a box's
    available room, back-solved from real, empirical spare-capacity
    measurements the user made in real PowerPoint on two independent,
    differently-sized, differently-shaped boxes (a single-column table
    page's textMainBullets and a plain L-column continuation page) --
    both independently implied a real std_lh of ~13.0pt against this
    formula's previous 13.8pt (line_h 10.8, PROVEN correct separately,
    see POWERPOINT_LINE_PITCH_FACTOR's own history -- so the gap
    isolates to para_gap specifically: implied ~2.2pt, not 3.0pt).
    Deliberately not landing exactly on 2.2 without a second real
    cross-check on the render side too -- see the commit message this
    shipped in for the full reasoning and what still needs re-verifying.
    """
    return 3.0


def _account_is_chinese(account: Dict) -> bool:
    """Language flag for MEASUREMENT (which glyph-width table to wrap
    with). Uses the account's own is_chinese when present (set by the
    payload builder via contains_predominantly_chinese_text); otherwise
    detects from the commentary instead of silently defaulting to
    English -- measuring CJK text with Arial's advance table (which has
    no CJK glyphs) under-counted lines badly enough that genuinely
    overflowing slots passed the render-time autofit gate as 'fits'."""
    v = (account or {}).get("is_chinese")
    if v is not None:
        return bool(v)
    return contains_predominantly_chinese_text(str((account or {}).get("commentary", "")))


def _account_cost_key(account: Dict, is_chinese_databook: bool = False) -> str:
    """The key text whose rendered width the cost model should charge: the
    label the renderer will actually draw, plus the continuation marker it
    appends (' (续)' / \" (cont'd)\").

    The marker half was always right. The BASE was not -- this charged
    `mapping_key`, while the renderer writes `display_name_zh or display_name`
    (generation.py's run_key.text). On a Chinese deck those are different
    strings: 'CIP' against '在建工程', 'Long-term deferred expenses' against
    '长期待摊费用'. The gap runs to tens of points, in both directions.

    Worth knowing why this took so long to prove. Measured against real
    PowerPoint the fix scored a DEAD HEAT with the old behaviour twice --
    24/26 bullets each -- and was twice rejected as not worth the risk. It only
    became visible once per-character <a:ea>/<a:latin> pricing landed
    (CompositeMetrics in text_metrics.py) and the ruler stopped being the
    dominant error: on the very next run the old behaviour fell to 23/26 while
    this scored 26/26 at zero net line error. Fixing the biggest error is what
    made the next one measurable; neither was visible through the other.
    """
    key = _rendered_bullet_label(account, is_chinese_databook)
    if account.get("is_continuation"):
        key = f"{key} (续)" if _account_is_chinese(account) else f"{key} (cont'd)"
    return str(key or "")


def _jieba_word_boundary_snap(text: str, pos: int) -> Optional[int]:
    """If jieba is installed, segment `text` and return the start index
    of whichever word strictly contains `pos`, or None if pos already
    sits on a word boundary (or jieba is unavailable/errors).

    This is the GENERAL version of the curated _PROTECTED_CJK_COMPOUNDS
    list below -- it was found to have a real gap: "结清" (settle) split
    as "...或交割前结" / "清安排..." in real production output, the
    SECOND compound found broken after the first round of fixes
    (人民币/万元/分别/年度) -- confirming a fixed list will always be
    one case behind whatever the AI writes next, since Chinese has no
    spaces to mark word boundaries structurally. jieba is a real
    Chinese-word-segmentation library (context-aware -- correctly
    keeps "784"/"万元" as separate tokens but "分别"/"年度"/"结清" as
    single ones); used here ONLY for its segmentation, no other
    behaviour change. Optional dependency, imported lazily so a
    machine without it still runs (falls back to the curated list
    below, unchanged) rather than failing PPTX generation outright.
    """
    try:
        import jieba  # type: ignore
    except ImportError:
        return None
    try:
        offset = 0
        for word in jieba.cut(text):
            word_len = len(word)
            if offset < pos < offset + word_len:
                return offset
            offset += word_len
            if offset >= pos:
                break
        return None
    except Exception:
        return None


def _merge_contd_pairs(accounts: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Merge any consecutive run of (part1, cont'd-part2, cont'd-part3, ...)
    fragments that landed in the same slot.  This happens when the DP
    re-balances: a split was created because an earlier slot was almost
    full, but the resulting pieces all fit together in the slot the DP
    actually chose.  Merging removes the spurious (cont'd) label(s) and
    restores the original single account.

    Only ever merged a single PAIR until a real screenshot (IMG_0076)
    showed an orphaned "(续)" bullet sitting right after its own
    already-rendered head -- a 3-way split (a middle fragment that is
    BOTH is_partial [it got re-split by a later rebalance pass] AND
    is_continuation [it continues the fragment before it]) only had its
    first two pieces merged; the third was never considered because the
    old loop looked at exactly one `nxt`, not the whole chain. Confirmed
    via direct reproduction + tracing: a 3-part split landed as
    [is_partial, is_partial+is_continuation, is_continuation] in one
    slot, and the old pairwise merge produced "merged(1+2)" followed by
    untouched "3" as its own still-(续)-labelled bullet."""
    result: List[Dict[str, Any]] = []
    i = 0
    n = len(accounts)
    while i < n:
        acct = accounts[i]
        if not acct.get("is_partial"):
            result.append(acct)
            i += 1
            continue
        base_key = acct.get("mapping_key")
        run = [acct]
        j = i + 1
        while (
            j < n
            and accounts[j].get("is_continuation")
            and accounts[j].get("original_key", accounts[j].get("mapping_key")) == base_key
        ):
            run.append(accounts[j])
            j += 1
        if len(run) > 1:
            combined = run[0].copy()
            combined["commentary"] = " ".join(
                str(a.get("commentary", "") or "").strip() for a in run
            ).strip()
            combined.pop("is_partial", None)
            combined.pop("part_num", None)
            # A middle fragment re-split by a later rebalance pass can
            # itself be is_continuation=True (it continues a head that
            # sits in an EARLIER slot) as well as is_partial=True (it
            # got split again, with its own tail in THIS run) -- if
            # run[0] is one of those, keep its is_continuation/
            # original_key on the merged result so the "(续)" label
            # renders correctly against the real earlier-slot head;
            # only drop them when run[0] is a genuine, non-continuation
            # first part (the common case).
            if not run[0].get("is_continuation"):
                combined.pop("is_continuation", None)
                combined.pop("original_key", None)
            result.append(combined)
        else:
            result.append(acct)
        i = j
    return result


def _expand_commentary_to_cover_summary(slide) -> bool:
    """Remove coSummaryShape from a continuation slide and expand the
    commentary box(es) upward to fill the freed area.

    Returns True if the operation modified the slide. Called only on
    continuation slides (i.e., not the first slide of a BS/IS statement)
    so the AI executive summary stays on the first slide only.
    """
    summary_shape = find_shape_by_name(slide.shapes, "coSummaryShape")
    if summary_shape is None:
        return False
    try:
        co_top = int(summary_shape.top)
        co_height = int(summary_shape.height)
    except Exception:
        return False
    co_bottom = co_top + co_height

    for slot_name in ("textMainBullets", "textMainBullets_L", "textMainBullets_R"):
        box = find_shape_by_name(slide.shapes, slot_name)
        if box is None:
            continue
        try:
            box_top = int(box.top)
            box_height = int(box.height)
        except Exception:
            continue
        # Only expand boxes located below the summary shape — avoid
        # accidentally covering tables / titles that sit above it.
        if box_top >= co_bottom:
            extension = box_top - co_top
            box.top = co_top
            box.height = box_height + extension

    try:
        sp = summary_shape._element
        sp.getparent().remove(sp)
    except Exception as exc:
        logger.warning("Could not remove coSummaryShape on continuation slide: %s", exc)
        return False
    return True


# ECMA-376 CT_TableCellProperties fixes the ORDER of tcPr's children:
#   lnL, lnR, lnT, lnB, lnTlToBr, lnBlToTr, cell3D, <fill>, headers, extLst
# PowerPoint enforces it. Appending a border to tcPr puts it AFTER the
# <a:solidFill> that cell.fill.solid() already wrote, the sequence is invalid,
# and PowerPoint silently drops the border -- which is why repeated attempts to
# restyle the total-row rules "never took effect" while fills always did.
# Every border therefore has to be INSERTED at its ranked position, not
# appended.
_TCPR_CHILD_ORDER = (
    "lnL", "lnR", "lnT", "lnB", "lnTlToBr", "lnBlToTr", "cell3D",
    "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill",
    "headers", "extLst",
)


def _tcpr_insert_ordered(tcPr, element) -> None:
    """Place `element` among tcPr's children in schema order."""
    from lxml import etree

    def _rank(el) -> int:
        name = etree.QName(el).localname
        return _TCPR_CHILD_ORDER.index(name) if name in _TCPR_CHILD_ORDER else len(_TCPR_CHILD_ORDER)

    my_rank = _rank(element)
    for existing in tcPr:
        if _rank(existing) > my_rank:
            existing.addprevious(element)
            return
    tcPr.append(element)


def _clear_cell_border(cell, border_position='top'):
    """Explicitly draw NO border on one edge of a cell.

    Not the same as never calling _set_cell_border: a cell left alone
    inherits whatever the table style GUID draws, and the UpSlide style this
    deck uses puts a white hairline between cells. That is invisible against
    white data rows but reads as a seam straight across the solid blue header
    band -- which is what "中間不要白色border" was pointing at. An explicit
    <a:noFill/> is the only thing that overrides the style; a border with no
    fill child is treated as "inherit", not "none".
    """
    from pptx.oxml.xmlchemy import OxmlElement

    tag_map = {'top': 'lnT', 'bottom': 'lnB', 'left': 'lnL', 'right': 'lnR'}
    tag_name = tag_map.get(border_position)
    if not tag_name:
        return

    tcPr = cell._tc.get_or_add_tcPr()
    ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    ln = tcPr.find(f"{ns}{tag_name}")
    if ln is None:
        ln = OxmlElement(f"a:{tag_name}")
        _tcpr_insert_ordered(tcPr, ln)
    # Same reason _set_cell_border clears first: append() never replaces, so a
    # leftover <a:solidFill> would still be there next to the <a:noFill/>.
    for child in list(ln):
        ln.remove(child)
    ln.append(OxmlElement('a:noFill'))


def _set_cell_border(cell, border_position='top', color_rgb=None, width=Pt(1)):
    """Set cell border"""
    from pptx.oxml.xmlchemy import OxmlElement
    
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    
    # Map position to tag name
    tag_map = {'top': 'lnT', 'bottom': 'lnB', 'left': 'lnL', 'right': 'lnR'}
    tag_name = tag_map.get(border_position)
    if not tag_name:
        return
        
    # Check if line element exists
    ln = tcPr.find(f"{{http://schemas.openxmlformats.org/drawingml/2006/main}}{tag_name}")
    if ln is None:
        ln = OxmlElement(f"a:{tag_name}")
        # Ordered, not appended -- see _tcpr_insert_ordered. Appending here put
        # the border after the cell's fill and PowerPoint dropped it.
        _tcpr_insert_ordered(tcPr, ln)
        
    # Set properties
    ln.set('w', str(int(width)))
    ln.set('cap', 'flat')
    ln.set('cmpd', 'sng')
    ln.set('algn', 'ctr')

    # Calling this twice on the same side (e.g. a full-grid pass, then a
    # heavier total-row override) previously left BOTH the old and new
    # <a:solidFill>/<a:prstDash>/<a:round>/<a:headEnd>/<a:tailEnd>
    # children on `ln` -- append() never replaces, so the element ended
    # up with duplicates and PowerPoint's rendering of that is
    # undefined (in practice, whichever child renderers pick up first).
    # Clear any existing children before appending the new ones so a
    # second call genuinely overrides the first, not just adds to it.
    for child in list(ln):
        ln.remove(child)

    # Set color
    if color_rgb:
        solidFill = OxmlElement('a:solidFill')
        srgbClr = OxmlElement('a:srgbClr')
        # Convert RGBColor or tuple to hex string
        hex_color = "000000"
        if isinstance(color_rgb, str):
            hex_color = color_rgb.replace('#', '')
        elif isinstance(color_rgb, tuple) and len(color_rgb) == 3:
            hex_color = f"{color_rgb[0]:02x}{color_rgb[1]:02x}{color_rgb[2]:02x}"
        # If it's an RGBColor object, user should pass str or tuple for this low-level func
            
        srgbClr.set('val', hex_color)
        solidFill.append(srgbClr)
        ln.append(solidFill)
        
        prstDash = OxmlElement('a:prstDash')
        prstDash.set('val', 'solid')
        ln.append(prstDash)
        
        round_ = OxmlElement('a:round')
        ln.append(round_)
        
        headEnd = OxmlElement('a:headEnd')
        headEnd.set('type', 'none')
        headEnd.set('w', 'med')
        headEnd.set('len', 'med')
        ln.append(headEnd)
        
        tailEnd = OxmlElement('a:tailEnd')
        tailEnd.set('type', 'none')
        tailEnd.set('w', 'med')
        tailEnd.set('len', 'med')
        ln.append(tailEnd)


def _apply_east_asian_line_breaking(paragraph) -> None:
    """Turn on East Asian line-breaking (禁则处理) and hanging punctuation
    for one paragraph.

    Without this a real deck put a full stop at the START of a line, and
    in the worst case a lone "。" on its own line under a paragraph that
    otherwise ended cleanly. Chinese typography forbids a line beginning
    with closing punctuation (。，）」etc.); the rule that prevents it is
    a PARAGRAPH property, and the template declares no <a:pPr> at all, so
    nothing was asserting it.

    eaLnBrk       -- apply East Asian line-break rules rather than Latin
                     ones. Our runs carry font.name='Arial' (a Latin
                     typeface) even for Chinese text, which is exactly
                     the case where PowerPoint may otherwise fall back to
                     Latin breaking.
    hangingPunct  -- let trailing punctuation hang past the right margin
                     instead of being pushed onto the next line, which is
                     what keeps "米。" together.

    Set explicitly rather than relied on as a schema default -- the
    observed render proves the default was not being applied here.
    """
    try:
        pPr = paragraph._p.get_or_add_pPr()
        pPr.set("eaLnBrk", "1")
        pPr.set("hangingPunct", "1")
    except Exception as exc:
        logger.debug("Could not set East Asian line-breaking: %s", exc)


def _force_no_autofit(text_frame) -> None:
    """Set the text frame's bodyPr autofit to ``<a:noAutofit/>`` so
    PowerPoint never shrinks the text to fit the shape. The template
    ships with ``<a:spAutoFit/>`` (resize shape to fit text), which in
    some viewers falls back to shrinking the text when the shape can't
    grow. Forcing ``noAutofit`` keeps the text at the exact point size
    we set (9pt / 10pt); overflow is simply clipped at the shape edge."""
    try:
        from lxml import etree  # noqa: F401
        from pptx.oxml.ns import qn
        bodyPr = text_frame._txBody.bodyPr
        # Remove any existing autofit child (spAutoFit / normAutofit / noAutofit).
        for tag in ("a:spAutoFit", "a:normAutofit", "a:noAutofit"):
            for child in bodyPr.findall(qn(tag)):
                bodyPr.remove(child)
        from pptx.oxml import parse_xml
        bodyPr.append(parse_xml(
            '<a:noAutofit xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'
        ))
    except Exception as exc:
        logger.debug("Could not force noAutofit on text frame: %s", exc)


def _build_clause_segments(
    commentary: str,
    clause_reviews: Optional[List[Dict[str, Any]]],
) -> Optional[List[Tuple[str, str]]]:
    """Split commentary into (text, category) segments using clause_reviews.

    Returns None if no clauses match. Falls back to a single 'data-backed'
    segment for any text not matched by any clause review (so unmatched
    prose stays black).
    """
    if not commentary or not clause_reviews:
        return None
    # Sort clauses by their position in the commentary
    positions: List[Tuple[int, int, str]] = []
    used_starts: set = set()
    for review in clause_reviews:
        clause_text = str(review.get("clause") or "").strip()
        if not clause_text:
            continue
        category = str(review.get("category") or ("data-backed" if review.get("supported") else "hallucination")).lower()
        search_from = 0
        # Find first non-overlapping occurrence
        while True:
            idx = commentary.find(clause_text, search_from)
            if idx == -1:
                break
            if idx in used_starts:
                search_from = idx + 1
                continue
            used_starts.add(idx)
            positions.append((idx, idx + len(clause_text), category))
            break
    if not positions:
        return None
    positions.sort()
    # Merge overlaps by sorting and skipping fully-contained overlaps
    cleaned: List[Tuple[int, int, str]] = []
    for start, end, cat in positions:
        if cleaned and start < cleaned[-1][1]:
            continue
        cleaned.append((start, end, cat))
    # Build segments from start to end of commentary
    segments: List[Tuple[str, str]] = []
    cursor = 0
    for start, end, cat in cleaned:
        if start > cursor:
            segments.append((commentary[cursor:start], "data-backed"))
        segments.append((commentary[start:end], cat))
        cursor = end
    if cursor < len(commentary):
        segments.append((commentary[cursor:], "data-backed"))
    return segments


def _category_to_rgb(category: str) -> Optional[Tuple[int, int, int]]:
    if category == "hallucination":
        return (200, 16, 46)  # red
    if category == "reasoning":
        return (213, 94, 0)  # orange
    return None


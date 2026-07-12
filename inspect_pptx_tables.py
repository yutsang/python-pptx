#!/usr/bin/env python3
"""Dump the exact cell-level formatting of every table in a .pptx.

Use this against a real, correctly-formatted past report to extract a
concrete spec (row order, subtotal placement, number format, fonts,
fills, borders) that fdd_utils/pptx.py's embed_financial_tables() should
be matching. Also works against our own generated output, so the same
tool can diff "what we produce" against "what the company format is".

Usage:
    python inspect_pptx_tables.py <path.pptx> [--slide N] [--table N]
"""
from __future__ import annotations

import argparse
import sys
from typing import Any, Optional

from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Emu


def _emu_to_in(value: Optional[int]) -> Optional[float]:
    if value is None:
        return None
    return round(Emu(value).inches, 3)


def _color_str(color_format) -> str:
    try:
        if color_format.type is None:
            return "inherited"
        from pptx.enum.dml import MSO_THEME_COLOR
        ctype = color_format.type
        if str(ctype).startswith("MSO_COLOR_TYPE.THEME") or ctype == 2:
            try:
                return f"theme:{MSO_THEME_COLOR(color_format.theme_color).name}"
            except Exception:
                return f"theme:{color_format.theme_color}"
        return f"#{color_format.rgb}"
    except Exception:
        return "?"


def _fill_str(fill) -> str:
    try:
        ftype = fill.type
        if ftype is None:
            return "inherited"
        name = str(ftype).split(".")[-1] if ftype is not None else "NONE"
        if "SOLID" in name.upper():
            return f"solid {_color_str(fill.fore_color)}"
        return name
    except Exception:
        return "?"


def _run_font_desc(run) -> str:
    f = run.font
    try:
        name = f.name or "(inherit)"
        size = f"{f.size.pt:g}pt" if f.size else "(inherit)"
        bold = "B" if f.bold else ""
        italic = "I" if f.italic else ""
        style = "".join(x for x in (bold, italic) if x) or "-"
        color = _color_str(f.color) if f.color and f.color.type is not None else "inherited"
        return f"{name} {size} {style} color={color}"
    except Exception as exc:
        return f"(font read error: {exc})"


_BORDER_TAGS = (("L", "a:lnL"), ("R", "a:lnR"), ("T", "a:lnT"), ("B", "a:lnB"))


def _cell_borders(cell) -> str:
    tc = cell._tc
    tcPr = tc.find(qn("a:tcPr"))
    if tcPr is None:
        return "no tcPr"
    parts = []
    for label, tag in _BORDER_TAGS:
        ln = tcPr.find(qn(tag))
        if ln is None:
            continue
        w = ln.get("w")
        w_pt = round(int(w) / 12700, 2) if w else None
        fill = ln.find(qn("a:noFill"))
        if fill is not None:
            parts.append(f"{label}=none")
            continue
        solid = ln.find(qn("a:solidFill"))
        color = "?"
        if solid is not None:
            srgb = solid.find(qn("a:srgbClr"))
            if srgb is not None:
                color = f"#{srgb.get('val')}"
            else:
                scheme = solid.find(qn("a:schemeClr"))
                if scheme is not None:
                    color = f"theme:{scheme.get('val')}"
        parts.append(f"{label}={w_pt}pt {color}")
    return "; ".join(parts) if parts else "no explicit borders (inherits table style)"


def _cell_fill(cell) -> str:
    try:
        return _fill_str(cell.fill)
    except Exception as exc:
        return f"(fill read error: {exc})"


def dump_table(slide_idx: int, shape, table_idx_on_slide: int) -> None:
    tbl = shape.table
    n_rows = len(tbl.rows)
    n_cols = len(tbl.columns)
    left_in = _emu_to_in(shape.left)
    top_in = _emu_to_in(shape.top)
    width_in = _emu_to_in(shape.width)
    height_in = _emu_to_in(shape.height)

    print("=" * 100)
    print(
        f"SLIDE {slide_idx + 1}  table #{table_idx_on_slide}  shape_name={shape.name!r}  "
        f"pos=({left_in}in, {top_in}in)  size=({width_in}in x {height_in}in)  "
        f"rows={n_rows} cols={n_cols}"
    )
    tbl_elem = tbl._tbl
    tblPr = tbl_elem.find(qn("a:tblPr"))
    if tblPr is not None:
        style_id_elem = tblPr.find(qn("a:tableStyleId"))
        style_id = style_id_elem.text if style_id_elem is not None else None
        print(
            f"  tableStyleId={style_id}  firstRow={tblPr.get('firstRow')}  "
            f"firstCol={tblPr.get('firstCol')}  bandRow={tblPr.get('bandRow')}  "
            f"bandCol={tblPr.get('bandCol')}"
        )
    print("  column widths (in):", [_emu_to_in(col.width) for col in tbl.columns])
    print("  row heights   (in):", [_emu_to_in(row.height) for row in tbl.rows])
    print("-" * 100)

    for r in range(n_rows):
        for c in range(n_cols):
            cell = tbl.cell(r, c)
            if cell.is_spanned:
                continue
            text = cell.text
            span = ""
            if cell.is_merge_origin:
                span = f"  [merged {cell.span_height}x{cell.span_width}]"
            paragraphs = cell.text_frame.paragraphs
            fonts = []
            alignments = []
            for p in paragraphs:
                alignments.append(str(p.alignment) if p.alignment is not None else "inherited")
                for run in p.runs:
                    fonts.append(_run_font_desc(run))
            font_desc = fonts[0] if fonts else "(no runs / empty)"
            align_desc = alignments[0] if alignments else "inherited"
            is_total_like = any(
                kw in text.lower() or kw in text
                for kw in ("total", "合计", "总计", "小计", "subtotal")
            )
            flag = "  *** TOTAL/SUBTOTAL-LIKE ***" if is_total_like and text.strip() else ""
            print(f"  R{r}C{c}{span}: text={text!r}")
            print(f"        font={font_desc}  align={align_desc}  fill={_cell_fill(cell)}{flag}")
            print(f"        borders: {_cell_borders(cell)}")
    print()


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("pptx_path", help="Path to a .pptx file")
    ap.add_argument("--slide", type=int, default=None, help="1-indexed slide number to restrict to")
    ap.add_argument("--table", type=int, default=None, help="1-indexed table-on-slide number to restrict to (requires --slide)")
    args = ap.parse_args()

    prs = Presentation(args.pptx_path)
    found_any = False
    for slide_idx, slide in enumerate(prs.slides):
        if args.slide is not None and (slide_idx + 1) != args.slide:
            continue
        table_idx_on_slide = 0
        for shape in slide.shapes:
            if not getattr(shape, "has_table", False):
                continue
            table_idx_on_slide += 1
            if args.table is not None and table_idx_on_slide != args.table:
                continue
            found_any = True
            dump_table(slide_idx, shape, table_idx_on_slide)

    if not found_any:
        print("No table shapes found matching the given filters.")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())

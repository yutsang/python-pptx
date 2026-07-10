"""PPTX export inspection tool — run this against a FILE ALREADY EXPORTED by
the pipeline (Streamlit UI or CLI test script) to catch the layout classes of
bug that used to require opening PowerPoint by eye:

  1. L/R column collision — a BS/IS content page rendered as one full-width
     commentary box instead of two side-by-side halves (template only had a
     single textMainBullets box for a page that needed two logical slots).
  2. Table/commentary overlap — the embedded financial statement table drawn
     on top of (instead of beside) the commentary text box.
  3. Overflow risk — commentary text that, at production font size/spacing,
     wraps to more lines than the box has vertical room for (same Pillow
     glyph-width measurement the packer itself uses, so a "no overflow" here
     means the packer's own capacity estimate agrees with this independent
     re-check).
  4. Fill ratio (utilisation) — how full each slot ended up, matching the
     packer's own target_fill_min_ratio concept; flags slots that are
     suspiciously empty EXCEPT the last slot of a statement (a lighter tail
     page is normal, not a bug).

Everything here is DETERMINISTIC, geometry + font-metrics only — no AI calls,
no PowerPoint installation required. It does NOT replace opening the real
PPTX in PowerPoint at least once per template/font change (subtle rendering
quirks — kerning, hinting, OS-level font substitution — are still only fully
verifiable there), but it catches the structural classes of bug automatically
so that step becomes a spot-check instead of a full manual read-through.

Usage:
    python inspect_pptx.py path/to/exported.pptx
    python inspect_pptx.py path/to/exported.pptx --config fdd_utils/config.yml
"""
from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass
from typing import List, Optional

from pptx import Presentation
from pptx.util import Emu

from fdd_utils.financial_common import load_yaml_file
from fdd_utils.text_metrics import get_measurer, text_box_from_shape

DEFAULT_CONFIG_CANDIDATES = ["fdd_utils/config.yml", "fdd_utils/config.example.yml"]

# Mirrors fdd_utils/pptx.py's own get_font_size_for_text/get_line_spacing_for_text
# (the functions that actually set the run's real formatting) so this
# independent check agrees with what really gets rendered. Both are a flat
# 9pt/1.0 regardless of language for English; Chinese previously assumed
# 10pt/0.95 here (matching a since-fixed bug in pptx.py's own capacity/content
# calculators) — real values are 9pt/0.9, same as English's font size.
FONT_SIZE_ENG = 9.0
FONT_SIZE_CHI = 9.0
LINE_SPACING_ENG = 1.0
LINE_SPACING_CHI = 0.9
MIN_FILL_RATIO_WARN = 0.40  # below this on a non-last slot -> utilisation flag


def _is_chinese_text(text: str) -> bool:
    return any("一" <= ch <= "鿿" for ch in text)


def _slot_of(shape_name: str) -> str:
    name = (shape_name or "").lower()
    if name.endswith("_l"):
        return "L"
    if name.endswith("_r"):
        return "R"
    return "single"


def _load_config(path: Optional[str]) -> dict:
    candidates = [path] if path else DEFAULT_CONFIG_CANDIDATES
    for candidate in candidates:
        if not candidate:
            continue
        try:
            cfg = load_yaml_file(candidate)
            if cfg:
                return cfg
        except (FileNotFoundError, OSError):
            continue
    return {}


@dataclass
class ShapeInfo:
    name: str
    slot: str
    left_in: float
    top_in: float
    width_in: float
    height_in: float
    text: str
    n_chars: int
    capacity_lines: int
    wrapped_lines: int
    fill_ratio: float
    overflow: bool


def _bbox_overlap(a, b) -> bool:
    """True if two (left, top, width, height) EMU boxes overlap by area."""
    ax1, ay1, ax2, ay2 = a[0], a[1], a[0] + a[2], a[1] + a[3]
    bx1, by1, bx2, by2 = b[0], b[1], b[0] + b[2], b[1] + b[3]
    return ax1 < bx2 and bx1 < ax2 and ay1 < by2 and by1 < ay2


def inspect_pptx(pptx_path: str, config: dict, *, quiet: bool = False) -> dict:
    """Runs the full layout inspection and returns a structured summary
    (used both by this file's own CLI and by inspect_databook.py's combined
    export+inspect flow). Pass quiet=True to suppress the per-slide print
    lines and keep only the final summary print (caller still gets full
    detail back in the returned dict either way)."""
    _print = (lambda *a, **k: None) if quiet else print
    packing_cfg = ((config.get("pptx") or {}).get("commentary_packing") or {})
    metrics_eng = packing_cfg.get("font_metrics_path_eng") or "fdd_utils/font_metrics/arial_eng.json"
    metrics_chi = packing_cfg.get("font_metrics_path_chi") or "fdd_utils/font_metrics/msyh_chi.json"
    family_eng = packing_cfg.get("font_family_eng") or "Arial"
    family_chi = packing_cfg.get("font_family_chi") or "Microsoft YaHei"

    eng_measurer = get_measurer(family_eng, FONT_SIZE_ENG, is_cjk=False,
                                 line_spacing=LINE_SPACING_ENG, metrics_path=metrics_eng)
    chi_measurer = get_measurer(family_chi, FONT_SIZE_CHI, is_cjk=True,
                                 line_spacing=LINE_SPACING_CHI, metrics_path=metrics_chi)
    _print(f"Measurement source: ENG={eng_measurer.source}  CHI={chi_measurer.source}")
    if eng_measurer.source == "system-font" or chi_measurer.source == "system-font":
        _print("⚠️  Falling back to a system-installed font for at least one language — "
               "this machine's font may not match the client's PowerPoint font. Check "
               f"that {metrics_eng!r} / {metrics_chi!r} exist and are readable.")

    prs = Presentation(pptx_path)
    _print(f"\nTotal slides: {len(prs.slides)}\n")

    total_warnings = 0
    warning_details: List[str] = []
    slide_reports: List[dict] = []

    for slide_idx, slide in enumerate(prs.slides):
        commentary_shapes = [
            s for s in slide.shapes
            if "textmainbullets" in (getattr(s, "name", "") or "").lower() and s.has_text_frame
        ]
        table_shapes = [s for s in slide.shapes if getattr(s, "has_table", False)]
        has_summary = any("summary" in (getattr(s, "name", "") or "").lower() for s in slide.shapes)

        if not commentary_shapes:
            continue

        _print(f"=== Slide {slide_idx + 1} ===  (table={bool(table_shapes)}  coSummaryShape={has_summary})")

        infos: List[ShapeInfo] = []
        for shape in commentary_shapes:
            slot = _slot_of(shape.name)
            text = shape.text_frame.text
            box = text_box_from_shape(shape)
            measurer = chi_measurer if _is_chinese_text(text) else eng_measurer
            line_h = measurer.line_height_pt()
            capacity = int(box.height_pt // line_h) if line_h > 0 else 0
            wrapped = measurer.wrap(text, box.width_pt) if text.strip() else []
            n_lines = len(wrapped)
            fill_ratio = (n_lines / capacity) if capacity > 0 else 0.0
            infos.append(ShapeInfo(
                name=shape.name, slot=slot,
                left_in=Emu(shape.left).inches if shape.left is not None else -1,
                top_in=Emu(shape.top).inches if shape.top is not None else -1,
                width_in=Emu(shape.width).inches if shape.width is not None else -1,
                height_in=Emu(shape.height).inches if shape.height is not None else -1,
                text=text, n_chars=len(text), capacity_lines=capacity,
                wrapped_lines=n_lines, fill_ratio=fill_ratio,
                overflow=n_lines > capacity,
            ))

        for i, info in enumerate(infos):
            is_last_slot_on_slide = (i == len(infos) - 1)
            flags = []
            if info.overflow:
                flags.append("⚠️ OVERFLOW RISK")
            if info.n_chars > 0 and info.fill_ratio < MIN_FILL_RATIO_WARN and not is_last_slot_on_slide:
                flags.append(f"📉 under-filled ({info.fill_ratio:.0%})")
            flag_str = ("  " + "  ".join(flags)) if flags else ""
            if flags:
                total_warnings += 1
                warning_details.append(f"Slide {slide_idx + 1} [{info.slot}] {info.name}: {', '.join(flags)}")
            _print(f"  [{info.slot:6s}] {info.name:24s} left={info.left_in:5.2f}in width={info.width_in:5.2f}in "
                   f"chars={info.n_chars:4d} capacity={info.capacity_lines:3d}L wraps_to={info.wrapped_lines:3d}L{flag_str}")

        # 1. L/R collision: a page with no table/summary (i.e. NOT the
        # designed single-column table slide) but only a single unsplit slot.
        slots_seen = {info.slot for info in infos}
        if slots_seen == {"single"} and not table_shapes and not has_summary:
            _print("  ❌ L/R COLLISION SUSPECTED — only a 'single' (full-width) commentary "
                   "slot on a page with no table/coSummaryShape, i.e. this looks like an L/R "
                   "content page that collapsed into one box instead of two side-by-side halves.")
            total_warnings += 1
            warning_details.append(f"Slide {slide_idx + 1}: L/R collision suspected")

        # 2. Table/commentary bounding-box overlap.
        for table_shape in table_shapes:
            t_box = (table_shape.left, table_shape.top, table_shape.width, table_shape.height)
            for info, shape in zip(infos, commentary_shapes):
                if info.n_chars == 0:
                    continue
                c_box = (shape.left, shape.top, shape.width, shape.height)
                if None in t_box or None in c_box:
                    continue
                if _bbox_overlap(t_box, c_box):
                    _print(f"  ❌ TABLE/COMMENTARY OVERLAP — '{table_shape.name}' overlaps '{info.name}'.")
                    total_warnings += 1
                    warning_details.append(f"Slide {slide_idx + 1}: table overlaps '{info.name}'")

        slide_reports.append({
            "slide": slide_idx + 1, "table": bool(table_shapes), "coSummaryShape": has_summary,
            "shapes": [i.__dict__ for i in infos],
        })
        _print()

    _print("=" * 78)
    if total_warnings == 0:
        _print("✅ No layout warnings found across all slides.")
    else:
        _print(f"⚠️  {total_warnings} warning(s) found — see ❌/⚠️/📉 markers above.")
    _print("Reminder: this is a geometry + font-metrics check, not a substitute for opening")
    _print("the file in real PowerPoint at least once per template/font change.")

    return {
        "total_slides": len(prs.slides),
        "content_slides": len(slide_reports),
        "total_warnings": total_warnings,
        "warning_details": warning_details,
        "measurement_source_eng": eng_measurer.source,
        "measurement_source_chi": chi_measurer.source,
        "slide_reports": slide_reports,
    }


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pptx_path", help="Path to an already-exported .pptx file")
    ap.add_argument("--config", default=None, help="Path to config.yml (default: tries fdd_utils/config.yml then config.example.yml)")
    args = ap.parse_args()

    config = _load_config(args.config)
    result = inspect_pptx(args.pptx_path, config)
    return 1 if result["total_warnings"] else 0


if __name__ == "__main__":
    sys.exit(main())

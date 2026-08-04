#!/usr/bin/env python3
"""Read-only diagnostic for the "real spare capacity exceeds computed
capacity" gap reported on both table-lead-in and plain commentary boxes.

Dumps every internal number _calculate_max_lines_for_textbox's "exact by
construction" formula depends on (raw shape height, tIns/bIns margins,
measurer.line_height_pt(), para_gap, std_lh, computed capacity) for named
real boxes in a real exported .pptx -- so a genuine mismatch between the
formula and real PowerPoint rendering can be pinpointed to ONE specific
input, instead of guessing at the combined effect.

Usage:
    python diagnose_capacity_gap.py "Crescent_....pptx"

No PPTX is modified. Safe to run any time.
"""
import argparse
import sys

from pptx import Presentation
from pptx.util import Emu

from fdd_utils.text_metrics import get_measurer, text_box_from_shape

FONT_SIZE_PT = 9.0
LINE_SPACING = 1.0
PARA_GAP_PT = 2.2  # keep in lockstep with fdd_utils/pptx.py's _real_para_gap_pt


def _is_chinese(text: str) -> bool:
    return any('一' <= c <= '鿿' for c in text)


def _load_config():
    try:
        import yaml
        with open("fdd_utils/config.yml", encoding="utf-8") as f:
            return yaml.safe_load(f) or {}
    except Exception:
        return {}


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("pptx_path")
    ap.add_argument("--extra-lines", type=float, default=None,
                     help="If you've empirically measured N real spare lines on the "
                          "FIRST shape checked, pass it here to back-solve the implied "
                          "real line_height_pt this machine's metrics should be using.")
    args = ap.parse_args()

    config = _load_config()
    packing_cfg = ((config.get("pptx") or {}).get("commentary_packing") or {})
    metrics_eng = packing_cfg.get("font_metrics_path_eng") or "fdd_utils/font_metrics/arial_eng.json"
    metrics_chi = packing_cfg.get("font_metrics_path_chi") or "fdd_utils/font_metrics/msyh_chi.json"
    family_eng = packing_cfg.get("font_family_eng") or "Arial"
    family_chi = packing_cfg.get("font_family_chi") or "Microsoft YaHei"

    eng_measurer = get_measurer(family_eng, FONT_SIZE_PT, is_cjk=False,
                                 line_spacing=LINE_SPACING, metrics_path=metrics_eng)
    chi_measurer = get_measurer(family_chi, FONT_SIZE_PT, is_cjk=True,
                                 line_spacing=LINE_SPACING, metrics_path=metrics_chi)
    print(f"Measurement source: ENG={eng_measurer.source}  CHI={chi_measurer.source}")
    print(f"  eng line_height_pt() = {eng_measurer.line_height_pt():.4f}")
    print(f"  chi line_height_pt() = {chi_measurer.line_height_pt():.4f}")
    print()

    prs = Presentation(args.pptx_path)
    checked = 0
    first_shape_result = None

    for slide_idx, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if not getattr(shape, "has_text_frame", False):
                continue
            name = (getattr(shape, "name", "") or "")
            is_target = (
                "textmainbullets" in name.lower()
                or name.startswith("TextBox")
            )
            if not is_target:
                continue
            text = shape.text_frame.text or ""
            if len(text.strip()) < 15:
                continue  # skip stray short labels (e.g. a leftover "评述" band)

            checked += 1
            is_chi = _is_chinese(text)
            measurer = chi_measurer if is_chi else eng_measurer

            raw_height_pt = Emu(shape.height).pt
            box = text_box_from_shape(shape)
            tIns_bIns_pt = raw_height_pt - box.height_pt

            line_h = measurer.line_height_pt()
            std_lh = line_h + PARA_GAP_PT
            capacity_units = max(1.0, box.height_pt / std_lh) if std_lh > 0 else 0.0

            print(f"Slide {slide_idx+1} | {name!r} | {'CHI' if is_chi else 'ENG'} | "
                  f"{len(text)} chars")
            print(f"  raw shape.height       = {raw_height_pt:.3f}pt ({raw_height_pt/72:.4f}in)")
            print(f"  tIns+bIns (margins)    = {tIns_bIns_pt:.3f}pt")
            print(f"  box.height_pt (usable) = {box.height_pt:.3f}pt")
            print(f"  measurer.line_height_pt() = {line_h:.4f}pt   (source={measurer.source})")
            print(f"  para_gap_pt            = {PARA_GAP_PT:.3f}pt")
            print(f"  std_lh (line_h+gap)    = {std_lh:.4f}pt")
            print(f"  capacity_units         = {capacity_units:.4f}L")
            print(f"  IDENTITY CHECK: capacity_units * std_lh = {capacity_units*std_lh:.3f}pt "
                  f"(should equal box.height_pt {box.height_pt:.3f}pt)")

            if first_shape_result is None:
                first_shape_result = (box.height_pt, std_lh, capacity_units, line_h)
                if args.extra_lines is not None:
                    real_capacity_units = capacity_units + args.extra_lines
                    implied_std_lh = box.height_pt / real_capacity_units
                    implied_line_h = implied_std_lh - PARA_GAP_PT
                    print()
                    print(f"  >>> Back-solved from your reported {args.extra_lines} extra real lines:")
                    print(f"      implied REAL std_lh      = {implied_std_lh:.4f}pt "
                          f"(currently computed: {std_lh:.4f}pt, delta={std_lh-implied_std_lh:+.4f}pt)")
                    print(f"      implied REAL line_height_pt = {implied_line_h:.4f}pt "
                          f"(currently computed: {line_h:.4f}pt, delta={line_h-implied_line_h:+.4f}pt)")
            print()

            if checked >= 6:
                print("(stopping after 6 shapes -- enough for comparison)")
                print()
                break
        if checked >= 6:
            break

    if checked == 0:
        print("No textMainBullets*/TextBox shapes with text found.")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())

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
    python inspect_pptx.py path/to/a/folder/          # loops every .pptx in it, prints a summary
"""
from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional

from pptx import Presentation
from pptx.util import Emu

from fdd_utils.financial_common import load_yaml_file
from fdd_utils.text_metrics import get_measurer, text_box_from_shape

DEFAULT_CONFIG_CANDIDATES = ["fdd_utils/config.yml", "fdd_utils/config.example.yml"]

# Mirrors _fill_text_main_bullets_with_category_and_key -- the function that
# ACTUALLY sets a textMainBullets run's formatting -- not get_font_size_for_
# text/get_line_spacing_for_text/get_space_after_for_text/get_space_before_
# for_text, which belong to a separate, legacy code path (_fill_content_shape,
# only reached from the unused markdown generate() flow) and were never the
# values really applied to a live commentary bullet. Caught via a real
# Windows client-metrics export + inspect_single_slot.py: assuming a 9pt
# inter-paragraph gap (this file's old PARA_GAP_CHI, itself copied from a
# since-fixed pptx.py bug) against a real hardcoded 3pt gap inflated "capacity
# used" by roughly 30% -- a box the user could still type 5-7 more lines into
# was already being reported as 94% full.
FONT_SIZE_ENG = 9.0
FONT_SIZE_CHI = 9.0
LINE_SPACING_ENG = 1.0
LINE_SPACING_CHI = 1.0
PARA_GAP_ENG = 3.0
PARA_GAP_CHI = 3.0
MIN_FILL_RATIO_WARN = 0.40  # below this on a non-last slot -> utilisation flag


def _is_chinese_text(text: str, threshold: float = 0.3) -> bool:
    # Predominantly-Chinese (mirrors fdd_utils.financial_common's
    # contains_predominantly_chinese_text / pptx.py's account-level
    # is_chinese flag), not "contains any CJK character" -- an English
    # commentary box that merely names a Chinese counterparty/person
    # still wraps as Latin-script prose in the real render, and measuring
    # it with CJK metrics here would misreport its true fill.
    if not text:
        return False
    chinese_chars = sum(1 for ch in text if "一" <= ch <= "鿿")
    return (chinese_chars / len(text)) > threshold


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
    capacity_lines: float
    wrapped_lines: int
    content_units: float
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
        summary_shapes = [
            s for s in slide.shapes
            if "summary" in (getattr(s, "name", "") or "").lower() and s.has_text_frame
        ]
        has_summary = bool(summary_shapes)

        if not commentary_shapes:
            continue

        _print(f"=== Slide {slide_idx + 1} ===  (table={bool(table_shapes)}  coSummaryShape={has_summary})")
        for s_shape in summary_shapes:
            s_text = s_shape.text_frame.text.strip()
            if s_text:
                preview = s_text[:120] + ("..." if len(s_text) > 120 else "")
                _print(f"  [{s_shape.name}] executive summary ({len(s_text)} chars): {preview!r}")
            else:
                _print(f"  ⚠️  [{s_shape.name}] executive summary shape is EMPTY (0 chars)")
                total_warnings += 1
                warning_details.append(f"Slide {slide_idx + 1}: executive summary ({s_shape.name}) is empty")

        infos: List[ShapeInfo] = []
        for shape in commentary_shapes:
            slot = _slot_of(shape.name)
            text = shape.text_frame.text
            box = text_box_from_shape(shape)
            is_chi = _is_chinese_text(text)
            measurer = chi_measurer if is_chi else eng_measurer
            line_h = measurer.line_height_pt()
            para_gap = PARA_GAP_CHI if is_chi else PARA_GAP_ENG
            std_lh = line_h + para_gap
            # Float, not int()-floored -- matches fdd_utils/pptx.py's
            # _calculate_max_lines_for_textbox (fixed in 5bbec43) and
            # inspect_single_slot.py (fixed in 3ceb2a4). This file's own
            # copy of the capacity formula was missed by both of those
            # fixes, so it kept discarding up to a full std_lh unit of
            # real box height and understating fill_ratio against what the
            # live packer actually computes.
            capacity = (box.height_pt / std_lh) if std_lh > 0 else 0.0

            # Bullet paragraphs (the "■ key - ..." line, rendered as p_key)
            # hang-indent: left_indent=0.15" / first_line_indent=-0.15", so
            # LINE 1 spans the box's FULL width and only WRAPPED continuation
            # lines (2+) are 10.8pt narrower. Continuation paragraphs (a
            # second '\n'-split commentary line, rendered as p_text with
            # first_line_indent=0) are narrow on EVERY line, no exception.
            # Mirror fdd_utils/pptx.py's _BULLET_HANGING_INDENT_PT /
            # first_line_width_pt so this independent re-check uses the same
            # effective width as the packer AND the real render.
            hang_w = max(10.0, box.width_pt - 10.8)

            # Literal wrapped-line count, for display only (chars=/wraps_to=).
            wrapped = measurer.wrap(text, hang_w) if text.strip() else []
            n_lines = len(wrapped)

            # Content cost in the SAME std_lh units as capacity: one para_gap
            # PER PARAGRAPH (not per wrapped physical line), mirroring
            # fdd_utils/pptx.py's _calculate_content_lines. Comparing a
            # literal physical-line count against a std_lh-unit capacity
            # is apples-to-oranges (std_lh bundles a full para_gap into every
            # "line", so it under-counts how many literal lines actually fit)
            # and produced false OVERFLOW RISK flags on ordinary multi-line
            # paragraphs before this was unit-matched.
            paras = [p for p in text.split("\n") if p.strip()] if text.strip() else []
            content_pt = sum(
                len(measurer.wrap(
                    p, hang_w,
                    first_line_width_pt=box.width_pt if p.lstrip().startswith("■") else None,
                )) * line_h + para_gap
                for p in paras
            )
            content_units = (content_pt / std_lh) if std_lh > 0 else 0.0

            fill_ratio = (content_units / capacity) if capacity > 0 else 0.0
            infos.append(ShapeInfo(
                name=shape.name, slot=slot,
                left_in=Emu(shape.left).inches if shape.left is not None else -1,
                top_in=Emu(shape.top).inches if shape.top is not None else -1,
                width_in=Emu(shape.width).inches if shape.width is not None else -1,
                height_in=Emu(shape.height).inches if shape.height is not None else -1,
                text=text, n_chars=len(text), capacity_lines=capacity,
                wrapped_lines=n_lines, content_units=content_units, fill_ratio=fill_ratio,
                overflow=content_units > capacity,
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
                   f"chars={info.n_chars:4d} capacity={info.capacity_lines:5.1f}L used={info.content_units:5.1f}L "
                   f"fill={info.fill_ratio:.0%} "
                   f"(raw wraps_to={info.wrapped_lines:3d}L, NOT comparable to capacity){flag_str}")

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


def _dump_text_and_check_duplicates(result: dict) -> int:
    """Prints every commentary shape's full text (not just the char-count
    summary line inspect_pptx() already prints), and scans every '■ '-led
    bullet across the WHOLE file for duplicates -- i.e. the same account's
    commentary appearing on more than one slide/slot. Returns the number of
    duplicate bullets found (0 = clean)."""
    print("\n" + "=" * 78)
    print("  FULL TEXT DUMP (per slide/shape)")
    print("=" * 78)
    bullet_locations: dict[str, list[str]] = {}
    for slide_report in result["slide_reports"]:
        slide_no = slide_report["slide"]
        for shape in slide_report["shapes"]:
            text = shape.get("text") or ""
            if not text.strip():
                continue
            print(f"\n--- Slide {slide_no} [{shape['name']}] ---")
            print(text)
            for line in text.split("\n"):
                stripped = line.strip()
                if stripped.startswith("■"):
                    # Key on the label + first ~40 chars of body -- enough to
                    # catch a genuine full-bullet repeat without false-flagging
                    # two different bullets that happen to share a label.
                    key = stripped[:60]
                    bullet_locations.setdefault(key, []).append(f"slide {slide_no} [{shape['name']}]")

    print("\n" + "=" * 78)
    print("  DUPLICATE-BULLET SCAN (same bullet text appearing on >1 slide/slot)")
    print("=" * 78)
    duplicates = {k: v for k, v in bullet_locations.items() if len(v) > 1}
    if not duplicates:
        print("✅ No duplicate bullets found -- every account's commentary appears exactly once.")
    else:
        for key, locations in duplicates.items():
            print(f"  ❌ DUPLICATE: {key!r}...")
            for loc in locations:
                print(f"      - {loc}")
    return len(duplicates)


# English "CNY<comma-grouped integer>" -- requires at least one comma group,
# which only ever appears on the exact-integer form ("CNY238,366"), never on
# the "CNY<X> million/thousand" decimal form ("CNY7.9 million" has no comma
# at all, so it can never match here regardless of its own decimal digit).
_ENG_CNY_INT_RE = re.compile(r"CNY\s?(\d{1,3}(?:,\d{3})+)\b")

# Currency-amount tokens, captured as a plain (possibly decimal) number so the
# VALUE can be checked for zero in Python rather than pattern-matching "0" as
# text -- matching text like "0" would also match the trailing ".0" inside an
# ordinary non-zero number such as "457.0万元" or "CNY570.0 million".
_NUMBER_WITH_UNIT_RES = [
    re.compile(r"CNY\s?(\d+(?:\.\d+)?)\s?(?:million|thousand)?"),
    re.compile(r"(?<![\d.])(\d+(?:\.\d+)?)万元"),
    re.compile(r"(?<![\d.])(\d+(?:\.\d+)?)亿元"),
    re.compile(r"人民币\s?(\d+(?:\.\d+)?)元(?!\S)"),
]


def _find_zero_currency_mentions(text: str) -> List[str]:
    hits: List[str] = []
    for pattern in _NUMBER_WITH_UNIT_RES:
        for m in pattern.finditer(text):
            try:
                if float(m.group(1)) == 0:
                    hits.append(m.group(0))
            except ValueError:
                continue
    return hits


def _check_number_formatting_and_zero_wording(result: dict) -> int:
    """Scans every '■ '-led bullet for two classes of issue flagged from real
    reports: (1) English sub-million CNY amounts more precise than the
    intended nearest-thousand rounding (e.g. 'CNY238,366' instead of
    'CNY238,000') -- these read as excessive, inconsistent-with-Chinese
    detail; (2) a literal zero-value currency mention that should have been
    reworded as 'nil'/'未发生' instead (e.g. 'CNY0', '人民币0.0万元').
    Returns the total number of flagged bullets (0 = clean)."""
    print("\n" + "=" * 78)
    print("  NUMBER-FORMATTING / ZERO-WORDING SCAN")
    print("=" * 78)
    print(
        "Flags two things per bullet: (a) English sub-million CNY amounts that\n"
        "aren't rounded to the nearest thousand (over-precise vs the Chinese\n"
        "report's own 万-unit rounding for the same figure), (b) a literal zero\n"
        "currency mention that should read as 'nil'/'未发生' instead. Neither is\n"
        "necessarily wrong on its own -- a genuinely sub-CNY10,000 amount is\n"
        "correctly exact, and a materiality-threshold '0%' isn't a currency\n"
        "mention -- so treat this as a worklist to skim, not an automatic fail.\n"
    )
    flagged = 0
    for slide_report in result["slide_reports"]:
        slide_no = slide_report["slide"]
        for shape in slide_report["shapes"]:
            text = shape.get("text") or ""
            for line in text.split("\n"):
                stripped = line.strip()
                if not stripped.startswith("■"):
                    continue
                label = stripped[:60]
                issues: List[str] = []

                for m in _ENG_CNY_INT_RE.finditer(stripped):
                    value = int(m.group(1).replace(",", ""))
                    if 10_000 <= value < 1_000_000 and value % 1000 != 0:
                        issues.append(f"over-precise amount {m.group(0)!r} (not rounded to nearest thousand)")

                for hit in _find_zero_currency_mentions(stripped):
                    issues.append(f"literal zero mention {hit!r} (should read as 'nil'/'未发生')")

                if issues:
                    flagged += 1
                    print(f"  Slide {slide_no} [{shape['name']}] {label!r}...")
                    for issue in issues:
                        print(f"      - {issue}")

    if not flagged:
        print("✅ No over-precise amounts or literal zero mentions found.")
    return flagged


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pptx_path", help="Path to an already-exported .pptx file, or a directory to scan "
                                       "for every .pptx in it (e.g. a folder of batch-exported decks)")
    ap.add_argument("--config", default=None, help="Path to config.yml (default: tries fdd_utils/config.yml then config.example.yml)")
    ap.add_argument("--dump-text", action="store_true",
                     help="Also print every commentary shape's full text and scan for (1) duplicate "
                          "bullets across the whole file (same account's commentary appearing on "
                          "more than one slide/slot), (2) over-precise English sub-million CNY amounts "
                          "(not rounded to the nearest thousand), (3) literal zero-value currency "
                          "mentions that should read as 'nil'/'未发生' instead.")
    args = ap.parse_args()

    config = _load_config(args.config)

    input_path = Path(args.pptx_path)
    if input_path.is_dir():
        pptx_files = sorted(input_path.glob("*.pptx"))
        if not pptx_files:
            print(f"No .pptx files found in {input_path}")
            return 1
    else:
        pptx_files = [input_path]

    total_warnings = 0
    total_duplicates = 0
    total_wording_flags = 0
    per_file_summary: List[tuple] = []
    for pptx_file in pptx_files:
        if len(pptx_files) > 1:
            print(f"\n{'=' * 90}\n{pptx_file.name}\n{'=' * 90}")
        result = inspect_pptx(str(pptx_file), config)
        duplicate_count = 0
        wording_flag_count = 0
        if args.dump_text:
            duplicate_count = _dump_text_and_check_duplicates(result)
            wording_flag_count = _check_number_formatting_and_zero_wording(result)
        total_warnings += result["total_warnings"]
        total_duplicates += duplicate_count
        total_wording_flags += wording_flag_count
        per_file_summary.append((pptx_file.name, result["total_warnings"], duplicate_count, wording_flag_count))

    if len(pptx_files) > 1:
        print(f"\n{'=' * 90}\nSUMMARY ({len(pptx_files)} file(s))\n{'=' * 90}")
        for name, warnings, duplicates, wording_flags in per_file_summary:
            flags = []
            if warnings:
                flags.append(f"{warnings} layout warning(s)")
            if duplicates:
                flags.append(f"{duplicates} duplicate bullet(s)")
            if wording_flags:
                flags.append(f"{wording_flags} wording flag(s)")
            status = "⚠️ " + ", ".join(flags) if flags else "✅ clean"
            print(f"  {name}: {status}")

    return 1 if (total_warnings or total_duplicates or total_wording_flags) else 0


if __name__ == "__main__":
    sys.exit(main())

"""inspect_render_truth.py — ask real PowerPoint where it actually broke the lines.

Everything else in this repo MODELS PowerPoint's layout (Pillow/metrics.json glyph
widths, a 1.2x line pitch, a 3pt paragraph gap), and every checker we have
measures with the same measurer the packer used -- so the two agree by
construction and a clean `inspect_pptx.py` run is not evidence that PowerPoint
agrees. The only real oracle was measure_boundheight.bas, pasted into the VBE by
hand and read back by eye. That manual loop is why the capacity investigation ran
15 rounds instead of 3.

This is the automated replacement. On Windows it drives PowerPoint over COM,
pulls back what its layout engine actually produced, and scores the model
against it.

The ground truth
----------------
    TextRange.BoundHeight          the height PowerPoint gave the text
    Paragraphs(k).Lines.Count      how many lines it broke paragraph k into

The per-paragraph count is the sharp instrument. A whole-shape height can be
right for the wrong reasons -- two errors cancelling is exactly what a mixed
slot produces -- but a paragraph PowerPoint broke into 4 lines where we
predicted 3 cannot be.

What it scores
--------------
For every "■ key - ..." bullet it computes three predictions and asks which one
PowerPoint agrees with:

    A  what the packer does today  regular-weight widths, and the label is the
                                   mapping_key (_account_cost_key), which is NOT
                                   the text the renderer writes
    B  A + the rendered label      the display_name/display_name_zh actually on
                                   the slide (_rendered_bullet_label already
                                   exists for this; 5 of its 7 call sites in
                                   gen_packing.py still use _account_cost_key)
    C  B + a bold-aware key run    generation.py sets run_key.font.bold = True,
                                   but text_metrics.Measurer holds one
                                   regular-weight table and has no weight
                                   parameter

If C scores no better than A on real content, the fixes are not worth making and
the remaining error is somewhere else. That is the point: decide from measured
truth, before writing anything into the packer.

It also reports
---------------
  * real fill % = BoundHeight / usable box height. The honest answer to "why do
    the pages look under-filled" -- unlike inspect_pptx.py's fill%, this
    numerator came from PowerPoint.
  * pitch = BoundHeight / Lines per shape. Below the nominal 10.8pt means
    PowerPoint re-ran its own autofit shrink and ignored the fontScale we wrote.
  * pitch and gap least-squares-fitted out of the real numbers, both ways round
    on whether the last paragraph's space_after counts -- so 10.8/3.0 gets
    re-measured rather than re-argued. (3.0 was already restored once, 63e4120,
    after 2.2 was back-solved and proved wrong. Re-fit, don't re-litigate.)

Usage
-----
    python inspect_render_truth.py exported_deck.pptx
    python inspect_render_truth.py exported_deck.pptx --lines --csv truth.csv
    python inspect_render_truth.py exported_deck.pptx --model-only   # no PowerPoint

Requires pywin32 (`pip install pywin32`) and PowerPoint. The deck is opened
READ-ONLY and closed without saving; PowerPoint is only quit if this script was
the one that started it.
"""

from __future__ import annotations

import argparse
import csv
import os
import sys
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Tuple

from pptx import Presentation

# The SAME measurer production uses. Imported, never reimplemented -- a fourth
# copy of the formula is exactly the drift this tool exists to detect.
from fdd_utils.financial_common import load_yaml_file
from fdd_utils.text_metrics import get_measurer, text_box_from_shape
from fdd_utils.pptx.helpers import (
    _measurer_family,
    _real_font_size_pt,
    _real_line_spacing,
    _real_para_gap_pt,
    _resolve_font_metrics_path,
)
from fdd_utils.pptx.payloads import _load_pptx_settings
from inspect_pptx import _is_chinese_text, _slot_of

# generation.py's p_key: left_indent 0.15" / first_line_indent -0.15", so line 1
# spans the full box and wrapped continuation lines sit 10.8pt narrower.
BULLET_HANGING_INDENT_PT = 10.8

BULLET_MARKER = "■"
# _explanation_render_text prefixes a post-table explanation with "➢ " in a
# Chinese deck and "- " in an English one.
EXPLAIN_MARKERS = ("➢", "- ", "• ")

# Bold faces for variant C. Windows first -- that is where this runs.
_BOLD_FONTS = {
    False: [r"C:\Windows\Fonts\arialbd.ttf",
            "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
            "/usr/share/fonts/truetype/msttcorefonts/Arial_Bold.ttf"],
    True: [r"C:\Windows\Fonts\msyhbd.ttc",
           r"C:\Windows\Fonts\msyh.ttc",
           "/Library/Fonts/Microsoft/Microsoft YaHei Bold.ttf"],
}


# ---------------------------------------------------------------------------
# Records
# ---------------------------------------------------------------------------

@dataclass
class ParaRow:
    slide: int
    shape: str
    slot: str
    index: int              # 1-based, matches PowerPoint's Paragraphs(k)
    kind: str               # bullet | continuation | category | explain | blank
    has_bold: bool
    chars: int
    lines_a: int            # today's packer
    lines_b: int            # + rendered label
    lines_c: int            # + bold-aware key run
    label: str = ""         # the text actually drawn in the bold run
    mapping_key: str = ""   # what the packer would have charged instead
    real_lines: Optional[int] = None
    space_after_pt: float = 0.0
    space_before_pt: float = 0.0
    selfcheck_ok: bool = True
    text: str = ""

    def delta(self, variant: str) -> Optional[int]:
        if self.real_lines is None:
            return None
        return self.real_lines - getattr(self, f"lines_{variant}")


@dataclass
class ShapeRow:
    slide: int
    shape: str
    slot: str
    is_chinese: bool
    box_w_pt: float
    box_h_pt: float
    model_lines: int
    model_pt: float
    real_lines: Optional[int] = None
    real_bound_pt: Optional[float] = None
    paras: List[ParaRow] = field(default_factory=list)

    @property
    def real_pitch(self) -> Optional[float]:
        if not self.real_lines or self.real_bound_pt is None:
            return None
        return self.real_bound_pt / self.real_lines

    @property
    def real_fill(self) -> Optional[float]:
        if self.real_bound_pt is None or self.box_h_pt <= 0:
            return None
        return self.real_bound_pt / self.box_h_pt


# ---------------------------------------------------------------------------
# Width sources
# ---------------------------------------------------------------------------

def _bold_font(is_cjk: bool, size_pt: float):
    """Pillow face for the BOLD key run. Pillow rather than a MetricsTable
    because there is no bold metrics.json and fontTools is not in
    requirements.txt -- but for Arial the JSON was dumped from the same file,
    so the two agree. Returns None when no bold face is installed, in which
    case variant C is simply not scored rather than silently faked."""
    from PIL import ImageFont
    for path in _BOLD_FONTS[is_cjk]:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size=size_pt)
            except Exception:
                continue
    return None


def _is_cjk_char(ch: str) -> bool:
    cp = ord(ch)
    return (0x3000 <= cp <= 0x303F or 0x3400 <= cp <= 0x4DBF or 0x4E00 <= cp <= 0x9FFF
            or 0xF900 <= cp <= 0xFAFF or 0xFF00 <= cp <= 0xFFEF)


def _atomize(text: str) -> List[str]:
    """Same atoms text_metrics.wrap_paragraph uses: one CJK char, or a maximal
    run of non-CJK non-space."""
    out, i, n = [], 0, len(text)
    while i < n:
        if text[i].isspace():
            i += 1
            continue
        if _is_cjk_char(text[i]):
            out.append(text[i]); i += 1
            continue
        j = i
        while j < n and not text[j].isspace() and not _is_cjk_char(text[j]):
            j += 1
        out.append(text[i:j]); i = j
    return out


def _wrap_runs(runs: Sequence[Tuple[str, bool]], measurer, bold_font,
               first_width_pt: float, width_pt: float) -> int:
    """Line count for a paragraph whose runs have DIFFERENT weights.

    A tool-local greedy wrapper, because the production `measurer.wrap()` takes
    one string at one weight and cannot express "this run is bold". It uses the
    production measurer for every regular run, so the widths are the same
    source; only the loop is local. `_selfcheck` below proves that loop agrees
    with `measurer.wrap()` whenever nothing is bold -- if it ever stops
    agreeing, the tool says so and variant C is not to be trusted.
    """
    def w(text: str, is_bold: bool) -> float:
        if is_bold and bold_font is not None:
            return float(bold_font.getlength(text))
        return measurer.text_width_pt(text)

    lines, cur, limit = 1, 0.0, first_width_pt
    prev_atom = ""
    for text, is_bold in runs:
        for atom in _atomize(text):
            sep = 0.0
            if cur > 0 and prev_atom and not _is_cjk_char(atom[0]) and not _is_cjk_char(prev_atom[-1]):
                sep = w(" ", is_bold)
            aw = w(atom, is_bold)
            if cur > 0 and cur + sep + aw > limit:
                lines += 1
                cur, limit = aw, width_pt
            else:
                cur += sep + aw
            prev_atom = atom
    return lines


# ---------------------------------------------------------------------------
# mapping_key recovery
# ---------------------------------------------------------------------------

def _alias_index() -> Dict[str, List[str]]:
    """rendered label -> the mapping_key(s) it could have come from.

    The deck only carries the label the renderer wrote; the packer charged the
    mapping_key. mappings.yml is what links them, and it is the same file the
    payload builder used, so this recovers variant A's input exactly rather
    than approximating it.
    """
    raw = load_yaml_file("fdd_utils/mappings.yml") or {}
    raw = raw.get("mappings", raw)
    idx: Dict[str, set] = {}
    for key, cfg in raw.items():
        if not isinstance(cfg, dict):
            continue
        for alias in [key] + list(cfg.get("aliases") or []):
            idx.setdefault(str(alias).strip(), set()).add(key)
    return {k: sorted(v) for k, v in idx.items()}


# ---------------------------------------------------------------------------
# The model half (runs anywhere -- this is what --model-only exercises)
# ---------------------------------------------------------------------------

def _para_kind(text: str, space_after_pt: float, starts_bold: bool) -> str:
    """Classify a paragraph the way the renderer built it.

    A category header is the only non-blank paragraph the renderer gives
    space_after = Pt(0) (generation.py's p_category); every bullet and
    continuation gets Pt(3). Read from the shape rather than pattern-matched,
    because prose continuations look exactly like category text.
    """
    stripped = text.strip()
    if not stripped:
        return "blank"
    if stripped.startswith(BULLET_MARKER):
        return "bullet"
    if stripped.startswith(EXPLAIN_MARKERS):
        return "explain"
    if space_after_pt == 0.0 and not starts_bold:
        return "category"
    return "continuation"


def _collect_model(deck_path: str, want_slide: Optional[int],
                   want_shape: Optional[str]) -> Tuple[List[ShapeRow], Dict[str, str], List[str]]:
    prs = Presentation(deck_path)
    packing = (_load_pptx_settings() or {})
    packing = packing.get("packing", packing)
    alias_idx = _alias_index()

    measurers, bold_fonts, env = {}, {}, {}
    for is_chi in (False, True):
        mpath = _resolve_font_metrics_path(is_chi, packing)
        m = get_measurer(
            _measurer_family(is_chi, packing), _real_font_size_pt(is_chi),
            is_cjk=is_chi, line_spacing=_real_line_spacing(is_chi), metrics_path=mpath,
        )
        measurers[is_chi] = m
        bold_fonts[is_chi] = _bold_font(is_chi, _real_font_size_pt(is_chi))
        tag = "CHI" if is_chi else "ENG"
        bold_note = "bold face NOT FOUND — variant C not scored"
        if bold_fonts[is_chi] is not None:
            bold_note = f"bold face {os.path.basename(getattr(bold_fonts[is_chi], 'path', '?'))}"
        env[tag] = (f"{m.source:<14} size={m.size_pt}pt  spacing={m.line_spacing}  "
                    f"line_h={m.line_height_pt():.2f}pt  gap={_real_para_gap_pt(is_chi):.2f}pt"
                    f"\n                 {mpath or '(no metrics.json — system font)'}"
                    f"\n                 {bold_note}")

    rows: List[ShapeRow] = []
    warnings: List[str] = []
    for s_idx, slide in enumerate(prs.slides, start=1):
        if want_slide and s_idx != want_slide:
            continue
        for shape in slide.shapes:
            name = str(getattr(shape, "name", "") or "")
            if not getattr(shape, "has_text_frame", False):
                continue
            tf = shape.text_frame
            text = tf.text or ""
            if not text.strip():
                continue
            # Commentary slots, plus the unnamed textboxes
            # _render_table_accounts_stack drops beside a table.
            if not (name.startswith("textMainBullets") or BULLET_MARKER in text):
                continue
            if want_shape and want_shape.lower() not in name.lower():
                continue

            box = text_box_from_shape(shape)
            is_chi = _is_chinese_text(text)
            measurer = measurers[is_chi]
            bold_font = bold_fonts[is_chi]
            line_h = measurer.line_height_pt()
            hang_w = max(10.0, box.width_pt - BULLET_HANGING_INDENT_PT)

            row = ShapeRow(slide=s_idx, shape=name or "(unnamed)", slot=_slot_of(name),
                           is_chinese=is_chi, box_w_pt=box.width_pt,
                           box_h_pt=box.height_pt, model_lines=0, model_pt=0.0)

            for p_idx, para in enumerate(tf.paragraphs, start=1):
                p_text = para.text or ""
                sa = para.space_after.pt if para.space_after is not None else 0.0
                sb = para.space_before.pt if para.space_before is not None else 0.0
                runs = list(para.runs)
                has_bold = any(bool(r.font.bold) for r in runs)
                starts_bold = bool(runs and runs[0].font.bold)
                kind = _para_kind(p_text, sa, starts_bold)

                label = mapping_key = ""
                selfcheck_ok = True
                if kind == "blank":
                    n_a = n_b = n_c = 1
                else:
                    # Production call, unmodified. This is variant B: the text
                    # that is really on the slide, regular weight throughout.
                    first_w = box.width_pt if kind == "bullet" else None
                    n_b = max(1, len(measurer.wrap(p_text, hang_w, first_line_width_pt=first_w)))
                    n_a = n_c = n_b

                    if kind == "bullet":
                        bold_run = next((r.text for r in runs if r.font.bold), "")
                        label = (bold_run or "").strip()
                        head, sep_found, tail = p_text.partition(bold_run) if bold_run else ("", "", "")
                        first_w = box.width_pt

                        # Self-check: the tool-local mixed wrapper must agree
                        # with the production wrapper when nothing is bold.
                        if sep_found:
                            n_plain = _wrap_runs([(head, False), (bold_run, False), (tail, False)],
                                                 measurer, bold_font, first_w, hang_w)
                            selfcheck_ok = (n_plain == n_b)
                            if not selfcheck_ok:
                                warnings.append(
                                    f"slide {s_idx} {name} para {p_idx}: local wrapper says "
                                    f"{n_plain} lines where measurer.wrap says {n_b}")

                            # Variant C: same text, key run measured bold.
                            if bold_font is not None:
                                n_c = _wrap_runs([(head, False), (bold_run, True), (tail, False)],
                                                 measurer, bold_font, first_w, hang_w)

                            # Variant A: what the packer charged -- mapping_key
                            # in place of the rendered label.
                            cands = alias_idx.get(label, [])
                            if len(cands) == 1:
                                mapping_key = cands[0]
                                if mapping_key != label:
                                    n_a = max(1, len(measurer.wrap(
                                        head + mapping_key + tail, hang_w,
                                        first_line_width_pt=first_w)))

                row.paras.append(ParaRow(
                    slide=s_idx, shape=row.shape, slot=row.slot, index=p_idx,
                    kind=kind, has_bold=has_bold, chars=len(p_text),
                    lines_a=n_a, lines_b=n_b, lines_c=n_c, label=label,
                    mapping_key=mapping_key, space_after_pt=sa, space_before_pt=sb,
                    selfcheck_ok=selfcheck_ok, text=p_text,
                ))
                row.model_lines += n_b
                row.model_pt += n_b * line_h + sa + sb
            # The final paragraph's space_after is invisible padding at the
            # bottom of the frame, not occupied height -- the same correction
            # _calculate_content_lines makes. Whether BoundHeight agrees is one
            # of the things the CALIBRATION fit settles.
            if row.paras:
                row.model_pt -= row.paras[-1].space_after_pt
            rows.append(row)
    return rows, env, warnings


# ---------------------------------------------------------------------------
# The ground-truth half (Windows + PowerPoint only)
# ---------------------------------------------------------------------------

def _attach_powerpoint():
    """Return (app, started_by_us). Reuses a running PowerPoint when there is
    one, so this never quits an instance the user already had open with their
    own work in it."""
    import win32com.client as win32
    try:
        return win32.GetActiveObject("PowerPoint.Application"), False
    except Exception:
        return win32.Dispatch("PowerPoint.Application"), True


def _sub_range(text_range, member: str, index: Optional[int] = None):
    """Get TextRange2.Lines()/.Paragraphs(), which are METHODS.

    VBA lets you write `tr.Lines.Count` because it resolves the default
    arguments for you; pywin32 does not, and `tr.Lines` there is a bound method
    whose `.Count` is the method object's own attribute -- it does not raise, it
    silently returns something meaningless. Every ground-truth number this tool
    prints comes through here for exactly that reason.
    """
    attr = getattr(text_range, member)
    try:
        return attr() if index is None else attr(index)
    except TypeError:
        return attr   # some pywin32/typelib combinations expose it as a property


def _normalize_com_text(text: str) -> str:
    """COM returns \\r for paragraph breaks and \\x0b for soft line breaks."""
    return str(text or "").replace("\r\n", "\n").replace("\r", "\n").replace("\x0b", "\n").strip()


def _fill_ground_truth(deck_path: str, rows: List[ShapeRow], warnings: List[str]) -> str:
    """Open the deck in real PowerPoint and record what its layout engine did.
    Read-only, closed without saving."""
    app, started_by_us = _attach_powerpoint()
    version, pres = "", None
    try:
        try:
            app.Visible = True   # PowerPoint refuses invisible automation
        except Exception:
            pass
        version = f"PowerPoint {getattr(app, 'Version', '?')} build {getattr(app, 'Build', '?')}"
        # Positional, not keyword: late-bound Dispatch resolves named arguments
        # through GetIDsOfNames and that is not reliable across Office builds.
        # Open(FileName, ReadOnly, Untitled, WithWindow); msoTrue = -1, and
        # WithWindow MUST be true or PowerPoint never lays the text out and
        # BoundHeight comes back as 0.
        pres = app.Presentations.Open(os.path.abspath(deck_path), -1, 0, -1)

        by_slide: Dict[int, List[ShapeRow]] = {}
        for r in rows:
            by_slide.setdefault(r.slide, []).append(r)

        for s_idx, wanted in by_slide.items():
            if s_idx > pres.Slides.Count:
                continue
            sld = pres.Slides(s_idx)
            unmatched = list(wanted)
            for j in range(1, sld.Shapes.Count + 1):
                shp = sld.Shapes(j)
                try:
                    if not shp.HasTextFrame or not shp.TextFrame2.HasText:
                        continue
                except Exception:
                    continue
                com_name = str(shp.Name or "")
                tr = shp.TextFrame2.TextRange
                com_text = _normalize_com_text(tr.Text)

                target = next((r for r in unmatched if r.shape == com_name), None)
                if target is None:
                    target = next(
                        (r for r in unmatched
                         if _normalize_com_text("\n".join(p.text for p in r.paras)) == com_text),
                        None)
                if target is None:
                    continue
                unmatched.remove(target)

                target.real_lines = int(_sub_range(tr, "Lines").Count)
                target.real_bound_pt = float(tr.BoundHeight)
                n_com = int(_sub_range(tr, "Paragraphs").Count)
                for p in target.paras:
                    if p.index <= n_com:
                        try:
                            p.real_lines = int(_sub_range(_sub_range(tr, "Paragraphs", p.index),
                                                          "Lines").Count)
                        except Exception:
                            p.real_lines = None
                if n_com != len(target.paras):
                    # Not cosmetic: the two sides disagree about what a
                    # paragraph even is, so every per-paragraph delta below it
                    # compares different things.
                    warnings.append(f"slide {s_idx} {com_name}: PowerPoint sees {n_com} "
                                    f"paragraphs, python-pptx sees {len(target.paras)} — "
                                    f"per-paragraph rows may be misaligned")
            for leftover in unmatched:
                warnings.append(f"slide {s_idx} {leftover.shape}: no matching shape in "
                                f"PowerPoint — not scored")
    finally:
        try:
            if pres is not None:
                pres.Close()
        except Exception:
            pass
        if started_by_us:
            try:
                app.Quit()
            except Exception:
                pass
    return version


# ---------------------------------------------------------------------------
# Calibration
# ---------------------------------------------------------------------------

def _fit_pitch_and_gap(rows: List[ShapeRow]) -> Optional[Dict[str, float]]:
    """Least-squares-fit pitch and gap out of the real measurements:

        BoundHeight = lines * pitch + gap_count * gap

    Fitted twice -- charging a gap for every paragraph, and for all but the last
    -- because which of those BoundHeight includes is precisely the correction
    _calculate_content_lines makes on faith.
    """
    try:
        import numpy as np
    except ImportError:
        return None
    usable = [r for r in rows if r.real_lines and r.real_bound_pt]
    if len(usable) < 3:
        return None

    out: Dict[str, float] = {"n": float(len(usable))}
    for label, drop_last in (("gap_per_para", False), ("gap_between_paras", True)):
        A, b = [], []
        for r in usable:
            n_gaps = sum(1 for p in r.paras if p.space_after_pt > 0)
            if drop_last and r.paras and r.paras[-1].space_after_pt > 0:
                n_gaps -= 1
            A.append([r.real_lines, n_gaps])
            b.append(r.real_bound_pt)
        A, b = np.array(A, float), np.array(b, float)
        sol, *_ = np.linalg.lstsq(A, b, rcond=None)
        out[f"{label}_pitch"] = float(sol[0])
        out[f"{label}_gap"] = float(sol[1])
        out[f"{label}_rmse"] = float(np.sqrt(np.mean((A @ sol - b) ** 2)))
    return out


# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------

VARIANTS = (("a", "today's packer (mapping_key, regular)"),
            ("b", "+ rendered label"),
            ("c", "+ bold-aware key run"))


def _report(rows: List[ShapeRow], env: Dict[str, str], version: str,
            warnings: List[str], show_lines: bool, model_only: bool) -> int:
    print("=" * 96)
    print("RENDER TRUTH — real PowerPoint layout vs this repo's model")
    print("=" * 96)
    print("\nMeasurement source (the ruler production itself uses):")
    for tag in ("ENG", "CHI"):
        print(f"  [{tag}] {env.get(tag, '?')}")
    print(f"  ground truth: {version or 'SKIPPED (--model-only)'}")

    if warnings:
        print("\n  !! " + "\n  !! ".join(warnings[:12]))
        if len(warnings) > 12:
            print(f"  !! ... and {len(warnings) - 12} more")

    # ---- per shape -------------------------------------------------------
    print("\n" + "-" * 96)
    print("PER SHAPE   fill% is BoundHeight/box — the numerator came from PowerPoint, not us.")
    print("            pitch below the nominal line_h means PowerPoint re-ran its own autofit.")
    print("-" * 96)
    print(f"{'sl':>3} {'shape':<24}{'lang':>5}{'box_h':>8}{'mLines':>7}{'rLines':>7}{'d':>4}"
          f"{'model_pt':>10}{'BoundH':>9}{'pitch':>7}{'fill%':>8}")
    n_shape_bad = 0
    for r in sorted(rows, key=lambda x: (x.slide, x.shape)):
        d = "" if r.real_lines is None else f"{r.real_lines - r.model_lines:+d}"
        if d not in ("", "+0"):
            n_shape_bad += 1
        print(f"{r.slide:>3} {r.shape[:24]:<24}{'CHI' if r.is_chinese else 'ENG':>5}"
              f"{r.box_h_pt:8.1f}{r.model_lines:7d}"
              f"{(r.real_lines if r.real_lines is not None else ''):>7}{d:>4}{r.model_pt:10.1f}"
              f"{(f'{r.real_bound_pt:.1f}' if r.real_bound_pt is not None else ''):>9}"
              f"{(f'{r.real_pitch:.2f}' if r.real_pitch else ''):>7}"
              f"{(f'{r.real_fill*100:.1f}' if r.real_fill else ''):>8}")

    all_paras = [p for r in rows for p in r.paras]
    scored = [p for p in all_paras if p.real_lines is not None]

    # The one thing that IS verifiable without PowerPoint: that the tool-local
    # mixed-weight wrapper reproduces the production wrapper exactly when no run
    # is bold. If it does not, variant C is measuring the tool's own bug rather
    # than the bold key run, so this has to be checked before the numbers mean
    # anything -- and it is checked on every run, not just --model-only.
    checked = [p for p in all_paras if p.kind == "bullet"]
    failed = [p for p in checked if not p.selfcheck_ok]
    print("\n" + "-" * 96)
    print(f"SELF-CHECK  local mixed-weight wrapper vs production measurer.wrap(): "
          f"{len(checked) - len(failed)}/{len(checked)} agree")
    print("-" * 96)
    if failed:
        print("  variant C is NOT trustworthy on the rows listed in the warnings above.")

    if model_only:
        print("\n" + "=" * 96)
        print("MODEL-ONLY RUN — no verdict on accuracy (that needs PowerPoint).")
        print("Re-run on Windows without --model-only.")
        print("=" * 96)
        return 0

    # ---- the scoreboard: which variant does PowerPoint agree with? -------
    print("\n" + "-" * 96)
    print("VARIANT SCOREBOARD  — how often each prediction matches PowerPoint, per paragraph")
    print("-" * 96)
    bullets = [p for p in scored if p.kind == "bullet"]
    if not bullets:
        print("  no ■ bullets scored")
    else:
        print(f"{'variant':<40}{'exact':>8}{'rate':>8}{'net lines':>11}{'worst':>7}")
        for v, name in VARIANTS:
            deltas = [p.delta(v) for p in bullets]
            exact = sum(1 for d in deltas if d == 0)
            print(f"{name:<40}{exact:>8}{exact/len(bullets)*100:7.1f}%"
                  f"{sum(deltas):>+11d}{max(abs(d) for d in deltas):>7}")
        print(f"\n  ({len(bullets)} bullets scored; 'net lines' is real minus predicted — "
              f"positive means\n   the model UNDER-counts and the slot will overflow, negative "
              f"means it over-counts\n   and the slot is left under-filled)")
        best = max("abc", key=lambda v: sum(1 for p in bullets if p.delta(v) == 0))
        print(f"\n  -> PowerPoint agrees most with variant '{best.upper()}'.")
        if best == "a":
            print("     The proposed fixes do NOT help on this deck; the error is elsewhere.")

    # non-bullet paragraphs use one prediction for all three variants
    others = [p for p in scored if p.kind != "bullet"]
    if others:
        bad = [p for p in others if p.delta("b") != 0]
        print(f"\n  non-bullet paragraphs: {len(bad)}/{len(others)} mis-wrapped, "
              f"net {sum(p.delta('b') for p in others):+d} lines")

    # ---- mis-wrapped paragraphs -----------------------------------------
    wrong = [p for p in scored if p.delta("c") != 0]
    print("\n" + "-" * 96)
    print("STILL MIS-WRAPPED UNDER THE BEST VARIANT (C)  — this is what remains unexplained")
    print("-" * 96)
    if not wrong:
        print(f"  none — all {len(scored)} paragraphs wrapped exactly as variant C predicts")
    else:
        print(f"{'sl':>3} {'shape':<20}{'#':>4}{'kind':>13}{'A':>3}{'B':>3}{'C':>3}"
              f"{'real':>5}{'d':>4}  text")
        for p in wrong:
            print(f"{p.slide:>3} {p.shape[:20]:<20}{p.index:>4}{p.kind:>13}"
                  f"{p.lines_a:>3}{p.lines_b:>3}{p.lines_c:>3}{p.real_lines:>5}"
                  f"{p.delta('c'):>+4}  {p.text[:30]}")

    if any(not p.selfcheck_ok for p in all_paras):
        print("\n  !! the tool-local mixed wrapper disagreed with measurer.wrap() on some "
              "paragraphs;\n     variant C is NOT trustworthy on those rows (listed in the "
              "warnings above).")

    # ---- calibration -----------------------------------------------------
    fit = _fit_pitch_and_gap(rows)
    print("\n" + "-" * 96)
    print("CALIBRATION  (fitted from the real measurements, not assumed)")
    print("-" * 96)
    if fit is None:
        print("  (needs numpy and at least 3 measured shapes)")
    else:
        print(f"  model constants:  pitch {_real_font_size_pt(False) * 1.2:.2f}pt   "
              f"gap {_real_para_gap_pt(False):.2f}pt")
        for label in ("gap_per_para", "gap_between_paras"):
            print(f"  fit [{label:<18}] pitch {fit[label+'_pitch']:6.2f}pt   "
                  f"gap {fit[label+'_gap']:5.2f}pt   rmse {fit[label+'_rmse']:5.2f}pt")
        best_fit = min(("gap_per_para", "gap_between_paras"), key=lambda k: fit[k + "_rmse"])
        print(f"  -> BoundHeight here is best explained by '{best_fit}' "
              f"(n={int(fit['n'])} shapes)")

    if show_lines and wrong:
        print("\n" + "-" * 96)
        print("FULL TEXT OF THE PARAGRAPHS STILL MIS-WRAPPED")
        print("-" * 96)
        for p in wrong:
            print(f"\n  slide {p.slide} {p.shape} para {p.index} [{p.kind}] "
                  f"A={p.lines_a} B={p.lines_b} C={p.lines_c} real={p.real_lines}"
                  + (f"\n  label drawn: {p.label!r}   packer charged: {p.mapping_key!r}"
                     if p.label else ""))
            print(f"    {p.text}")

    print("\n" + "=" * 96)
    net = sum(p.delta("c") for p in scored)
    print(f"VERDICT: under the best variant, {len(wrong)} of {len(scored)} paragraphs are still "
          f"mis-wrapped\n         (net {net:+d} lines deck-wide); {n_shape_bad} of {len(rows)} "
          f"shapes off on total lines.")
    print("=" * 96)
    return 1 if wrong else 0


def _write_csv(path: str, rows: List[ShapeRow]) -> None:
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(["slide", "shape", "slot", "para", "kind", "label", "mapping_key",
                    "chars", "lines_A", "lines_B", "lines_C", "real_lines",
                    "delta_A", "delta_B", "delta_C", "space_before_pt",
                    "space_after_pt", "selfcheck_ok", "text"])
        for r in rows:
            for p in r.paras:
                w.writerow([p.slide, p.shape, p.slot, p.index, p.kind, p.label,
                            p.mapping_key, p.chars, p.lines_a, p.lines_b, p.lines_c,
                            "" if p.real_lines is None else p.real_lines,
                            *["" if p.delta(v) is None else p.delta(v) for v in "abc"],
                            f"{p.space_before_pt:.2f}", f"{p.space_after_pt:.2f}",
                            int(p.selfcheck_ok), p.text])
    print(f"\nwrote {path}")


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("deck", help="an EXPORTED .pptx (not the template — it has no commentary)")
    ap.add_argument("--lines", action="store_true",
                    help="print the full text of every still-mis-wrapped paragraph")
    ap.add_argument("--slide", type=int, help="restrict to one slide (1-based)")
    ap.add_argument("--shape", help="restrict to shapes whose name contains this")
    ap.add_argument("--csv", help="write the per-paragraph table here")
    ap.add_argument("--model-only", action="store_true",
                    help="skip PowerPoint entirely (exercises the prediction half off Windows)")
    args = ap.parse_args()

    if not os.path.exists(args.deck):
        print(f"No such file: {args.deck}", file=sys.stderr)
        return 2

    rows, env, warnings = _collect_model(args.deck, args.slide, args.shape)
    if not rows:
        print("No commentary shapes found. Is this an exported deck rather than the template?",
              file=sys.stderr)
        return 2

    version = ""
    if not args.model_only:
        if sys.platform != "win32":
            print(f"This needs Windows + PowerPoint (running on {sys.platform}). "
                  f"Use --model-only to exercise the prediction half.", file=sys.stderr)
            return 2
        try:
            import win32com.client  # noqa: F401
        except ImportError:
            print("pywin32 is not installed. Run:  pip install pywin32", file=sys.stderr)
            return 2
        version = _fill_ground_truth(args.deck, rows, warnings)

    rc = _report(rows, env, version, warnings, args.lines, args.model_only)
    if args.csv:
        _write_csv(args.csv, rows)
    return rc


if __name__ == "__main__":
    raise SystemExit(main())

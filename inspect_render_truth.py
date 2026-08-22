"""inspect_render_truth.py — ask real PowerPoint where it actually broke the lines.

Everything else in this repo MODELS PowerPoint's layout (Pillow/metrics.json glyph
widths, a 1.2x line pitch, a 3pt paragraph gap). Those numbers have been wrong
before, and the only oracle so far was measure_boundheight.bas -- a VBA module the
user had to paste into the VBE by hand, run, and copy numbers back from. That
manual loop is why the capacity investigation took 15 rounds instead of 3.

This replaces it. On Windows with PowerPoint installed it drives the real
application over COM, pulls the ground truth PowerPoint's own layout engine
produced, and diffs it against the model THIS repo uses -- importing the very
same measurer production imports, so there is never a second ruler to drift.

What the ground truth actually is
---------------------------------
    TextRange.BoundHeight     the height PowerPoint gave the text
    TextRange.Lines.Count     how many lines it really broke into
    Paragraphs(k).Lines.Count the same, per paragraph -- this is the useful one,
                              because it says WHICH paragraph we mis-wrapped
    Lines(i).Text             the literal line PowerPoint drew

The per-paragraph line count is the sharp instrument. An aggregate height can
be right for the wrong reasons (two errors cancelling); a paragraph that
PowerPoint broke into 4 lines where we predicted 3 cannot.

Known model gaps this is built to measure
-----------------------------------------
1. The account name renders BOLD (generation.py's run_key.font.bold = True) but
   text_metrics.Measurer holds one regular-weight table and has no weight
   parameter. Arial Bold runs 4-10% wider. Every "■ key - ..." paragraph is
   therefore measured with the wrong ruler on its first line, and only those
   paragraphs. The report tags each mismatched paragraph with whether it holds
   a bold run, so the hypothesis either shows up in that column or dies there.
2. normAutofit: PowerPoint re-runs its own shrink and ignores the fontScale we
   write. The per-shape `pitch` column is BoundHeight/Lines -- anything below
   the nominal 10.8pt is PowerPoint having shrunk the box out from under us.
3. The 10.8pt pitch and 3.0pt gap themselves. The CALIBRATION section
   least-squares-fits both out of the real measurements rather than trusting
   the constants. 3.0 was already restored once (63e4120) after 2.2 was
   back-solved and proved wrong -- so re-fit, don't re-argue.

Usage
-----
    python inspect_render_truth.py exported_deck.pptx
    python inspect_render_truth.py exported_deck.pptx --lines      # show the diffs
    python inspect_render_truth.py exported_deck.pptx --csv out.csv
    python inspect_render_truth.py exported_deck.pptx --model-only # no PowerPoint

Requires pywin32 (`pip install pywin32`) and PowerPoint. The presentation is
opened READ-ONLY and closed without saving; PowerPoint is only quit if this
script was the one that started it.
"""

from __future__ import annotations

import argparse
import csv
import os
import sys
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from pptx import Presentation

# The SAME measurer production uses. Imported, never reimplemented -- a fourth
# copy of the formula is exactly the drift this tool exists to detect.
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

BULLET_MARKER = "■"    # ■  account lead-in, holds the BOLD key run
# _explanation_render_text prefixes a post-table explanation with "➢ " in a
# Chinese deck and "- " in an English one.
EXPLAIN_MARKERS = ("➢", "- ", "• ")


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
    model_lines: int
    real_lines: Optional[int] = None
    space_after_pt: float = 0.0
    space_before_pt: float = 0.0
    text: str = ""

    @property
    def delta(self) -> Optional[int]:
        return None if self.real_lines is None else self.real_lines - self.model_lines


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


# ---------------------------------------------------------------------------
# The model half (runs anywhere -- this is what --model-only exercises)
# ---------------------------------------------------------------------------

def _para_kind(text: str, space_after_pt: float, starts_bold: bool) -> str:
    """Classify a paragraph the way the renderer built it.

    A category header is the only non-blank paragraph the renderer gives
    space_after = Pt(0) (generation.py's p_category); every bullet and
    continuation gets Pt(3). That is a more reliable signal than guessing
    from the text, which is why it is read from the shape rather than
    pattern-matched.
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
                   want_shape: Optional[str]) -> Tuple[List[ShapeRow], Dict[str, str]]:
    """Walk the deck with python-pptx and predict every paragraph's line count
    with the production measurer."""
    prs = Presentation(deck_path)
    packing = _load_pptx_settings() or {}
    packing = packing.get("packing", packing)

    measurers = {}
    env: Dict[str, str] = {}
    for is_chi in (False, True):
        mpath = _resolve_font_metrics_path(is_chi, packing)
        m = get_measurer(
            _measurer_family(is_chi, packing), _real_font_size_pt(is_chi),
            is_cjk=is_chi, line_spacing=_real_line_spacing(is_chi), metrics_path=mpath,
        )
        measurers[is_chi] = m
        tag = "CHI" if is_chi else "ENG"
        env[tag] = (f"{m.source:<14} size={m.size_pt}pt  spacing={m.line_spacing}  "
                    f"line_h={m.line_height_pt():.2f}pt  gap={_real_para_gap_pt(is_chi):.2f}pt"
                    + (f"\n                 {mpath}" if mpath else ""))

    rows: List[ShapeRow] = []
    for s_idx, slide in enumerate(prs.slides, start=1):
        if want_slide and s_idx != want_slide:
            continue
        for shape in slide.shapes:
            name = str(getattr(shape, "name", "") or "")
            if not getattr(shape, "has_text_frame", False):
                continue
            # Commentary slots, plus the free textboxes _render_table_accounts_stack
            # drops beside a table (unnamed, so matched by having bullet text).
            tf = shape.text_frame
            text = tf.text or ""
            if not text.strip():
                continue
            if not (name.startswith("textMainBullets") or BULLET_MARKER in text):
                continue
            if want_shape and want_shape.lower() not in name.lower():
                continue

            box = text_box_from_shape(shape)
            is_chi = _is_chinese_text(text)
            measurer = measurers[is_chi]
            line_h = measurer.line_height_pt()
            hang_w = max(10.0, box.width_pt - BULLET_HANGING_INDENT_PT)

            row = ShapeRow(
                slide=s_idx, shape=name or "(unnamed)", slot=_slot_of(name),
                is_chinese=is_chi, box_w_pt=box.width_pt, box_h_pt=box.height_pt,
                model_lines=0, model_pt=0.0,
            )
            for p_idx, para in enumerate(tf.paragraphs, start=1):
                p_text = para.text or ""
                sa = para.space_after.pt if para.space_after is not None else 0.0
                sb = para.space_before.pt if para.space_before is not None else 0.0
                runs = list(para.runs)
                has_bold = any(bool(r.font.bold) for r in runs)
                starts_bold = bool(runs and runs[0].font.bold)
                kind = _para_kind(p_text, sa, starts_bold)

                if kind == "blank":
                    n = 1
                else:
                    n = max(1, len(measurer.wrap(
                        p_text, hang_w,
                        # Only the ■ lead-in hangs; p_text continuations are
                        # narrow on every line (first_line_indent = 0).
                        first_line_width_pt=box.width_pt if kind == "bullet" else None,
                    )))
                row.paras.append(ParaRow(
                    slide=s_idx, shape=row.shape, slot=row.slot, index=p_idx,
                    kind=kind, has_bold=has_bold, chars=len(p_text), model_lines=n,
                    space_after_pt=sa, space_before_pt=sb, text=p_text,
                ))
                row.model_lines += n
                row.model_pt += n * line_h + sa + sb
            # The final paragraph's space_after is invisible padding at the
            # bottom of the frame, not occupied height -- same correction
            # _calculate_content_lines makes. Whether PowerPoint's BoundHeight
            # agrees is one of the things the CALIBRATION fit settles.
            if row.paras:
                row.model_pt -= row.paras[-1].space_after_pt
            rows.append(row)
    return rows, env


# ---------------------------------------------------------------------------
# The ground-truth half (Windows + PowerPoint only)
# ---------------------------------------------------------------------------

def _attach_powerpoint():
    """Return (app, started_by_us). Reuses a running PowerPoint when there is
    one, so this never quits an instance the user already had open with their
    own work in it."""
    import win32com.client as win32
    try:
        app = win32.GetActiveObject("PowerPoint.Application")
        return app, False
    except Exception:
        return win32.Dispatch("PowerPoint.Application"), True


def _fill_ground_truth(deck_path: str, rows: List[ShapeRow]) -> str:
    """Open the deck in real PowerPoint and record what its layout engine did.

    Read-only, closed without saving. Shapes are matched to the python-pptx
    pass by (slide index, shape name); an unnamed table-stack textbox is
    matched by its text instead.
    """
    app, started_by_us = _attach_powerpoint()
    version = ""
    pres = None
    try:
        try:
            app.Visible = True   # PowerPoint refuses invisible automation
        except Exception:
            pass
        version = f"PowerPoint {getattr(app, 'Version', '?')} build {getattr(app, 'Build', '?')}"
        # Positional, not keyword: late-bound Dispatch resolves named arguments
        # through GetIDsOfNames and it is not reliable across Office builds.
        # Signature is Open(FileName, ReadOnly, Untitled, WithWindow); msoTrue
        # is -1, and WithWindow MUST be true or PowerPoint never lays the text
        # out and BoundHeight comes back as 0.
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
                        None,
                    )
                if target is None:
                    continue
                unmatched.remove(target)

                target.real_lines = int(_sub_range(tr, "Lines").Count)
                target.real_bound_pt = float(tr.BoundHeight)
                n_com = int(_sub_range(tr, "Paragraphs").Count)
                for p in target.paras:
                    if p.index <= n_com:
                        try:
                            para_range = _sub_range(tr, "Paragraphs", p.index)
                            p.real_lines = int(_sub_range(para_range, "Lines").Count)
                        except Exception:
                            p.real_lines = None
                if n_com != len(target.paras):
                    # Not cosmetic: it means the two sides disagree about what a
                    # paragraph even is, so every per-paragraph delta below it
                    # is comparing different things.
                    print(f"  !! slide {s_idx} {com_name}: PowerPoint sees {n_com} paragraphs, "
                          f"python-pptx sees {len(target.paras)} -- per-paragraph rows may be misaligned")
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


def _normalize_com_text(text: str) -> str:
    """COM returns \\r for paragraph breaks and \\x0b for soft line breaks."""
    return str(text or "").replace("\r\n", "\n").replace("\r", "\n").replace("\x0b", "\n").strip()


def _sub_range(text_range, member: str, index: Optional[int] = None):
    """Get TextRange2.Lines()/.Paragraphs()/.Runs(), which are METHODS.

    VBA lets you write `tr.Lines.Count` because it resolves the default
    arguments for you; pywin32 does not, and `tr.Lines` there is a bound
    method whose `.Count` is the method's own attribute count -- it does not
    raise, it silently returns something meaningless. Every ground-truth
    number this tool prints comes through here for exactly that reason.
    """
    attr = getattr(text_range, member)
    try:
        return attr() if index is None else attr(index)
    except TypeError:
        # Some pywin32/typelib combinations expose it as a property already.
        return attr


# ---------------------------------------------------------------------------
# Calibration
# ---------------------------------------------------------------------------

def _fit_pitch_and_gap(rows: List[ShapeRow]) -> Optional[Dict[str, float]]:
    """Least-squares-fit pitch and gap out of the real measurements:

        BoundHeight = lines * pitch + gap_count * gap

    Fitted twice -- once charging a gap for every paragraph, once for all but
    the last -- because which of those PowerPoint's BoundHeight includes is
    precisely the correction _calculate_content_lines makes on faith. The fit
    with the smaller residual is the one that is true on this machine.
    """
    try:
        import numpy as np
    except ImportError:
        return None
    usable = [r for r in rows if r.real_lines and r.real_bound_pt]
    if len(usable) < 3:
        return None

    out: Dict[str, float] = {}
    for label, drop_last in (("gap_per_para", False), ("gap_between_paras", True)):
        A, b = [], []
        for r in usable:
            n_gaps = sum(1 for p in r.paras if p.space_after_pt > 0)
            if drop_last and r.paras and r.paras[-1].space_after_pt > 0:
                n_gaps -= 1
            A.append([r.real_lines, n_gaps])
            b.append(r.real_bound_pt)
        A, b = np.array(A, dtype=float), np.array(b, dtype=float)
        sol, *_ = np.linalg.lstsq(A, b, rcond=None)
        resid = float(np.sqrt(np.mean((A @ sol - b) ** 2)))
        out[f"{label}_pitch"] = float(sol[0])
        out[f"{label}_gap"] = float(sol[1])
        out[f"{label}_rmse"] = resid
    out["n"] = float(len(usable))
    return out


# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------

def _report(rows: List[ShapeRow], env: Dict[str, str], version: str,
            show_lines: bool, model_only: bool) -> int:
    print("=" * 92)
    print("RENDER TRUTH — real PowerPoint layout vs this repo's model")
    print("=" * 92)
    print(f"\nMeasurement source (the ruler production itself uses):")
    for tag in ("ENG", "CHI"):
        print(f"  [{tag}] {env.get(tag, '?')}")
    if version:
        print(f"  ground truth: {version}")
    if model_only:
        print("  ground truth: SKIPPED (--model-only) — every real_* column below is blank")

    # ---- per shape -------------------------------------------------------
    print("\n" + "-" * 92)
    print("PER SHAPE  (pitch = BoundHeight/Lines; below the nominal line_h means "
          "PowerPoint shrank it)")
    print("-" * 92)
    hdr = (f"{'sl':>3} {'shape':<24}{'lang':>5}{'box_h':>8}{'mLines':>7}{'rLines':>7}"
           f"{'d':>4}{'model_pt':>10}{'BoundH':>9}{'pitch':>7}")
    print(hdr)
    n_shape_bad = 0
    for r in sorted(rows, key=lambda x: (x.slide, x.shape)):
        d = "" if r.real_lines is None else f"{r.real_lines - r.model_lines:+d}"
        if d not in ("", "+0"):
            n_shape_bad += 1
        print(f"{r.slide:>3} {r.shape[:24]:<24}{'CHI' if r.is_chinese else 'ENG':>5}"
              f"{r.box_h_pt:8.1f}{r.model_lines:7d}"
              f"{(r.real_lines if r.real_lines is not None else ''):>7}{d:>4}"
              f"{r.model_pt:10.1f}"
              f"{(f'{r.real_bound_pt:.1f}' if r.real_bound_pt is not None else ''):>9}"
              f"{(f'{r.real_pitch:.2f}' if r.real_pitch else ''):>7}")

    # ---- per paragraph ---------------------------------------------------
    all_paras = [p for r in rows for p in r.paras]
    scored = [p for p in all_paras if p.delta is not None]
    wrong = [p for p in scored if p.delta != 0]

    print("\n" + "-" * 92)
    print("MIS-WRAPPED PARAGRAPHS  (this is the sharp instrument — an aggregate "
          "height can be right\n                        for the wrong reasons; a "
          "paragraph line count cannot)")
    print("-" * 92)
    if model_only:
        print("  (needs PowerPoint)")
    elif not wrong:
        print(f"  none — all {len(scored)} paragraphs wrapped exactly as predicted")
    else:
        print(f"{'sl':>3} {'shape':<20}{'#':>4}{'kind':>13}{'bold':>6}{'chars':>7}"
              f"{'model':>6}{'real':>5}{'d':>4}  text")
        for p in wrong:
            print(f"{p.slide:>3} {p.shape[:20]:<20}{p.index:>4}{p.kind:>13}"
                  f"{('yes' if p.has_bold else '-'):>6}{p.chars:>7}"
                  f"{p.model_lines:>6}{p.real_lines:>5}{p.delta:>+4}  {p.text[:34]}")

    # ---- the bold hypothesis, stated as a number -------------------------
    if scored and not model_only:
        print("\n" + "-" * 92)
        print("BY PARAGRAPH KIND  (the bold key-name run only exists on 'bullet' rows,")
        print("                    so a bold-width error can ONLY show up there)")
        print("-" * 92)
        print(f"{'kind':>13}{'n':>6}{'wrong':>7}{'rate':>8}{'net lines':>11}")
        for kind in ("bullet", "continuation", "category", "explain", "blank"):
            grp = [p for p in scored if p.kind == kind]
            if not grp:
                continue
            bad = [p for p in grp if p.delta != 0]
            net = sum(p.delta for p in grp)
            print(f"{kind:>13}{len(grp):>6}{len(bad):>7}{len(bad)/len(grp)*100:7.1f}%{net:>+11d}")
        bold_grp = [p for p in scored if p.has_bold]
        plain_grp = [p for p in scored if not p.has_bold]
        if bold_grp and plain_grp:
            br = sum(1 for p in bold_grp if p.delta != 0) / len(bold_grp) * 100
            pr = sum(1 for p in plain_grp if p.delta != 0) / len(plain_grp) * 100
            print(f"\n  paragraphs WITH a bold run: {br:.1f}% mis-wrapped   "
                  f"without: {pr:.1f}% mis-wrapped")
            if br > pr + 10:
                print("  -> consistent with the bold key-name run being measured with the "
                      "regular-weight table.")
            elif bold_grp and br <= pr + 10:
                print("  -> NOT consistent with a bold-width problem; the bold run is being "
                      "absorbed by wrap quantisation.")

    # ---- calibration -----------------------------------------------------
    fit = None if model_only else _fit_pitch_and_gap(rows)
    print("\n" + "-" * 92)
    print("CALIBRATION  (fitted from the real measurements, not assumed)")
    print("-" * 92)
    if fit is None:
        print("  (needs PowerPoint and numpy, and at least 3 measured shapes)")
    else:
        nominal_pitch = _real_font_size_pt(False) * 1.2
        print(f"  model constants:  pitch {nominal_pitch:.2f}pt   gap {_real_para_gap_pt(False):.2f}pt")
        for label in ("gap_per_para", "gap_between_paras"):
            print(f"  fit [{label:<18}] pitch {fit[label+'_pitch']:6.2f}pt   "
                  f"gap {fit[label+'_gap']:5.2f}pt   rmse {fit[label+'_rmse']:5.2f}pt")
        best = min(("gap_per_para", "gap_between_paras"), key=lambda k: fit[k + "_rmse"])
        print(f"  -> BoundHeight on this machine is best explained by '{best}' "
              f"(n={int(fit['n'])} shapes)")

    # ---- line-level diff -------------------------------------------------
    if show_lines and wrong:
        print("\n" + "-" * 92)
        print("LINE-LEVEL DIFF  (model wrap vs the text PowerPoint drew is not "
              "retrievable per-\n                 paragraph over COM, so this shows OUR "
              "breaks for the paragraphs\n                 PowerPoint counted differently)")
        print("-" * 92)
        print("  re-run the model wrap for these by hand if you need the exact glyph "
              "position;\n  the paragraph text is printed in full below.\n")
        for p in wrong:
            print(f"  slide {p.slide} {p.shape} para {p.index} "
                  f"[{p.kind}, bold={'yes' if p.has_bold else 'no'}] "
                  f"model={p.model_lines} real={p.real_lines}")
            print(f"    {p.text}\n")

    # ---- verdict ---------------------------------------------------------
    print("\n" + "=" * 92)
    if model_only:
        print("MODEL-ONLY RUN — no verdict. Re-run on Windows without --model-only.")
        return 0
    net = sum(p.delta for p in scored)
    print(f"VERDICT: {len(wrong)} of {len(scored)} paragraphs mis-wrapped "
          f"({len(wrong)/max(1,len(scored))*100:.1f}%), net {net:+d} lines across the deck; "
          f"{n_shape_bad} of {len(rows)} shapes off.")
    print("=" * 92)
    return 1 if wrong else 0


def _write_csv(path: str, rows: List[ShapeRow]) -> None:
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(["slide", "shape", "slot", "para", "kind", "has_bold", "chars",
                    "model_lines", "real_lines", "delta", "space_before_pt",
                    "space_after_pt", "text"])
        for r in rows:
            for p in r.paras:
                w.writerow([p.slide, p.shape, p.slot, p.index, p.kind,
                            int(p.has_bold), p.chars, p.model_lines,
                            "" if p.real_lines is None else p.real_lines,
                            "" if p.delta is None else p.delta,
                            f"{p.space_before_pt:.2f}", f"{p.space_after_pt:.2f}", p.text])
    print(f"\nwrote {path}")


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("deck", help="an EXPORTED .pptx (not the template — it has no commentary)")
    ap.add_argument("--lines", action="store_true", help="print the full text of every mis-wrapped paragraph")
    ap.add_argument("--slide", type=int, help="restrict to one slide (1-based)")
    ap.add_argument("--shape", help="restrict to shapes whose name contains this")
    ap.add_argument("--csv", help="write the per-paragraph table here")
    ap.add_argument("--model-only", action="store_true",
                    help="skip PowerPoint entirely (lets the model half be exercised off-Windows)")
    args = ap.parse_args()

    if not os.path.exists(args.deck):
        print(f"No such file: {args.deck}", file=sys.stderr)
        return 2

    rows, env = _collect_model(args.deck, args.slide, args.shape)
    if not rows:
        print("No commentary shapes found. Is this an exported deck rather than the template?",
              file=sys.stderr)
        return 2

    version = ""
    if not args.model_only:
        if sys.platform != "win32":
            print(f"This needs Windows + PowerPoint (running on {sys.platform}). "
                  f"Use --model-only to exercise the model half.", file=sys.stderr)
            return 2
        try:
            import win32com.client  # noqa: F401
        except ImportError:
            print("pywin32 is not installed. Run:  pip install pywin32", file=sys.stderr)
            return 2
        version = _fill_ground_truth(args.deck, rows)

    rc = _report(rows, env, version, args.lines, args.model_only)
    if args.csv:
        _write_csv(args.csv, rows)
    return rc


if __name__ == "__main__":
    raise SystemExit(main())

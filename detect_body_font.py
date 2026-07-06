"""Detect which fonts a .pptx actually uses — especially the COMMENTARY BODY font —
so you know which installed .ttf/.otf to dump with dump_font_metrics.py.

Stdlib only; safe to copy standalone to any machine.

Usage:
    python detect_body_font.py template.pptx

What it reports:
  1. Raw diagnostic: how many 'typeface' attributes exist at all (catches the
     quoting/regex issues that make ad-hoc one-liners print []).
  2. All typeface values used, with counts (handles single- OR double-quoted XML).
  3. Theme fonts: major (headings) and minor (body) for latin / east-asian.
  4. Slide-master bodyStyle default fonts.
  5. Fonts explicitly set inside commentary shapes (textMainBullets / Text-commentary /
     coSummaryShape), with +mn-lt / +mn-ea / +mj-lt tokens resolved to real names.
  6. A conclusion: which EN + CN fonts to dump for text measurement.
"""
from __future__ import annotations

import re
import sys
import zipfile
from collections import Counter
from xml.etree import ElementTree as ET

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"

# Matches typeface='...' and typeface="..." alike.
TYPEFACE_RE = re.compile(r"typeface=[\"']([^\"']*)[\"']")

COMMENTARY_NAME_HINTS = ("bullet", "commentary", "summary")


def _decode(b: bytes) -> str:
    return b.decode("utf-8", "ignore")


def parse_theme_fonts(zf: zipfile.ZipFile) -> dict:
    """{'major_latin': 'X', 'major_ea': 'Y', 'minor_latin': ..., 'minor_ea': ...}"""
    out = {}
    for name in zf.namelist():
        if not (name.startswith("ppt/theme/") and name.endswith(".xml")):
            continue
        try:
            root = ET.fromstring(zf.read(name))
        except ET.ParseError:
            continue
        for kind, key in ((f"{A}majorFont", "major"), (f"{A}minorFont", "minor")):
            for grp in root.iter(kind):
                latin = grp.find(f"{A}latin")
                ea = grp.find(f"{A}ea")
                if latin is not None and latin.get("typeface"):
                    out.setdefault(f"{key}_latin", latin.get("typeface"))
                if ea is not None and ea.get("typeface"):
                    out.setdefault(f"{key}_ea", ea.get("typeface"))
        break  # theme1 is the presentation theme
    return out


def resolve(token: str, theme: dict) -> str:
    """Resolve +mn-lt / +mn-ea / +mj-lt / +mj-ea tokens to real font names."""
    mapping = {
        "+mn-lt": theme.get("minor_latin", "?"),
        "+mn-ea": theme.get("minor_ea", "") or theme.get("minor_latin", "?"),
        "+mj-lt": theme.get("major_latin", "?"),
        "+mj-ea": theme.get("major_ea", "") or theme.get("major_latin", "?"),
    }
    resolved = mapping.get(token, token)
    return f"{token} -> {resolved}" if token in mapping else token


def master_body_fonts(zf: zipfile.ZipFile) -> list:
    """Default body fonts from slide-master txStyles/bodyStyle (lvl1)."""
    found = []
    for name in zf.namelist():
        if not (name.startswith("ppt/slideMasters/") and name.endswith(".xml")):
            continue
        try:
            root = ET.fromstring(zf.read(name))
        except ET.ParseError:
            continue
        for body in root.iter(f"{P}bodyStyle"):
            lvl1 = body.find(f"{A}lvl1pPr")
            if lvl1 is None:
                continue
            def_rpr = lvl1.find(f"{A}defRPr")
            if def_rpr is None:
                continue
            for tag, label in ((f"{A}latin", "latin"), (f"{A}ea", "ea")):
                el = def_rpr.find(tag)
                if el is not None and el.get("typeface"):
                    found.append((name.split("/")[-1], label, el.get("typeface")))
    return found


def commentary_shape_fonts(zf: zipfile.ZipFile) -> dict:
    """{shape_name: Counter({'latin:X': n, 'ea:Y': n})} for commentary-ish shapes."""
    per_shape: dict = {}
    slide_names = sorted(n for n in zf.namelist()
                         if n.startswith("ppt/slides/slide") and n.endswith(".xml"))
    for sname in slide_names:
        try:
            root = ET.fromstring(zf.read(sname))
        except ET.ParseError:
            continue
        for sp in root.iter(f"{P}sp"):
            cnv = sp.find(f"{P}nvSpPr/{P}cNvPr")
            shape_name = (cnv.get("name") if cnv is not None else "") or ""
            if not any(h in shape_name.lower() for h in COMMENTARY_NAME_HINTS):
                continue
            counter = per_shape.setdefault(shape_name, Counter())
            for rpr_tag in (f"{A}rPr", f"{A}defRPr", f"{A}endParaRPr"):
                for rpr in sp.iter(rpr_tag):
                    for tag, label in ((f"{A}latin", "latin"), (f"{A}ea", "ea")):
                        el = rpr.find(tag)
                        if el is not None and el.get("typeface"):
                            counter[f"{label}:{el.get('typeface')}"] += 1
    return per_shape


def main() -> int:
    if len(sys.argv) < 2:
        print("usage: python detect_body_font.py <file.pptx>")
        return 1
    path = sys.argv[1]
    zf = zipfile.ZipFile(path)
    xml_names = [n for n in zf.namelist() if n.endswith(".xml")]

    # 1. Raw diagnostic — bytes-level, immune to quoting/regex mishaps.
    raw_count = sum(zf.read(n).count(b"typeface") for n in xml_names)
    print(f"[diag] 'typeface' occurrences across {len(xml_names)} xml parts: {raw_count}")
    if raw_count == 0:
        print("       (0 is abnormal — is this a valid pptx? does it use external theme parts?)")

    # 2. All typeface values (single- or double-quoted).
    counter: Counter = Counter()
    for n in xml_names:
        counter.update(v for v in TYPEFACE_RE.findall(_decode(zf.read(n))) if v)
    print("\n== All typeface values (count) ==")
    for face, cnt in counter.most_common(15) or [("(none)", 0)]:
        print(f"  {cnt:>4}  {face}")

    # 3. Theme fonts.
    theme = parse_theme_fonts(zf)
    print("\n== Theme fonts ==")
    print(f"  heading (major): latin={theme.get('major_latin', '(none)')} | ea={theme.get('major_ea', '(none)')}")
    print(f"  body    (minor): latin={theme.get('minor_latin', '(none)')} | ea={theme.get('minor_ea', '(none)')}")

    # 4. Slide-master body defaults.
    masters = master_body_fonts(zf)
    print("\n== Slide-master bodyStyle defaults ==")
    if masters:
        for fname, label, face in masters:
            print(f"  {fname}: {label} = {resolve(face, theme)}")
    else:
        print("  (none set — body inherits theme minor fonts)")

    # 5. Commentary shapes.
    shapes = commentary_shape_fonts(zf)
    print("\n== Fonts inside commentary shapes (textMainBullets / Text-commentary / coSummaryShape) ==")
    if shapes:
        for shape_name, cnt in shapes.items():
            faces = ", ".join(f"{resolve(k.split(':', 1)[1], theme)} ({k.split(':')[0]}×{v})"
                              for k, v in cnt.most_common())
            print(f"  {shape_name}: {faces or '(no explicit font — inherits master/theme)'}")
    else:
        print("  (no commentary-named shapes found, or none set explicit fonts —")
        print("   they inherit the master/theme body fonts above)")

    # 6. Conclusion.
    def _pick(label: str) -> str:
        # explicit shape font wins; else master default; else theme minor.
        for cnt in shapes.values():
            for k, _ in cnt.most_common():
                if k.startswith(label + ":") and not k.split(":", 1)[1].startswith("+"):
                    return k.split(":", 1)[1]
        for _, lab, face in masters:
            if lab == label and not face.startswith("+"):
                return face
        key = "minor_latin" if label == "latin" else "minor_ea"
        return theme.get(key, "") or "(unknown)"

    en, cn = _pick("latin"), _pick("ea")
    print("\n== CONCLUSION ==")
    print(f"  EN body font  : {en}")
    print(f"  CN body font  : {cn or '(falls back to latin / system substitution)'}")
    print("\nNext: dump the INSTALLED font files for measurement, e.g.:")
    print(f'  python dump_font_metrics.py "C:\\Windows\\Fonts\\<{en}>.ttf" > metrics_eng.json')
    print(f'  python dump_font_metrics.py "C:\\Windows\\Fonts\\<{cn}>.ttf" > metrics_chi.json')
    print("Then set pptx.commentary_packing.font_metrics_path_eng/chi in fdd_utils/config.yml.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

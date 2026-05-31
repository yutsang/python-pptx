"""Extract fonts embedded in a .pptx and report whether they're usable for text measurement.

A .pptx is a ZIP. If the author ticked "Embed fonts in the file", the font binaries
live at ppt/fonts/fontN.fntdata (raw TTF/OTF), and ppt/presentation.xml's
<p:embeddedFontLst> maps each typeface + style (regular/bold/italic) to a relationship
id resolved via ppt/_rels/presentation.xml.rels.

Usage:
    python extract_pptx_fonts.py "Client.template.pptx" [-o out_dir]

What it tells you:
  - which typefaces are embedded, and the style variants present
  - whether each font is FULL or only a SUBSET (subset = only the glyphs used in the
    deck, so it CANNOT measure arbitrary new commentary reliably)
  - the OS/2 fsType embedding permission (licensing note)
  - extracts each font to out_dir as a proper .ttf/.otf and verifies Pillow can load it
"""
from __future__ import annotations

import argparse
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


def _localname(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _detect_ext(head: bytes) -> str:
    if head[:4] in (b"\x00\x01\x00\x00", b"true", b"ttcf"):
        return ".ttf"
    if head[:4] == b"OTTO":
        return ".otf"
    if head[:4] in (b"wOFF", b"wOF2"):
        return ".woff"
    return ".bin"


def _parse_embedded_font_list(zf: zipfile.ZipFile) -> list[dict]:
    """Return [{typeface, styles: {regular: rId, bold: rId, ...}}] from presentation.xml."""
    try:
        xml = zf.read("ppt/presentation.xml")
    except KeyError:
        return []
    root = ET.fromstring(xml)
    fonts: list[dict] = []
    for ef in root.iter():
        if _localname(ef.tag) != "embeddedFont":
            continue
        entry: dict = {"typeface": None, "styles": {}}
        for child in ef:
            ln = _localname(child.tag)
            if ln == "font":
                entry["typeface"] = child.attrib.get("typeface")
            elif ln in ("regular", "bold", "italic", "boldItalic"):
                rid = next((v for k, v in child.attrib.items() if _localname(k) == "id"), None)
                if rid:
                    entry["styles"][ln] = rid
        if entry["typeface"]:
            fonts.append(entry)
    return fonts


def _parse_rels(zf: zipfile.ZipFile) -> dict[str, str]:
    """rId -> target path under ppt/ for font relationships."""
    try:
        xml = zf.read("ppt/_rels/presentation.xml.rels")
    except KeyError:
        return {}
    root = ET.fromstring(xml)
    out: dict[str, str] = {}
    for rel in root:
        rid = rel.attrib.get("Id")
        target = rel.attrib.get("Target", "")
        if rid and "font" in (rel.attrib.get("Type", "").lower() + target.lower()):
            # targets are usually "fonts/font1.fntdata", relative to ppt/
            out[rid] = "ppt/" + target.lstrip("/")
    return out


def _analyse(font_path: Path) -> str:
    """Report glyph coverage, subset likelihood, fsType, em, via fontTools + Pillow."""
    notes: list[str] = []
    try:
        from fontTools.ttLib import TTFont
        tt = TTFont(str(font_path), fontNumber=0, lazy=True)
        upm = tt["head"].unitsPerEm if "head" in tt else "?"
        cmap = tt.getBestCmap() if tt.get("cmap") else {}
        n_cmap = len(cmap)
        n_glyphs = len(tt.getGlyphOrder())
        has_cjk = any(0x4E00 <= cp <= 0x9FFF for cp in cmap)
        # Heuristic: a full Latin font maps ~200+ codepoints; a full CJK font thousands.
        if has_cjk:
            subset = n_cmap < 3000
            kind = "CJK"
        else:
            subset = n_cmap < 150
            kind = "Latin"
        verdict = "SUBSET (only chars used in the deck — NOT safe to measure new text)" if subset \
            else "FULL coverage — usable for measuring arbitrary new commentary"
        fstype = tt["OS/2"].fsType if "OS/2" in tt else None
        fs_map = {0: "Installable (no restriction)", 2: "Restricted license", 4: "Preview & Print",
                  8: "Editable"}
        fs_note = fs_map.get(fstype, f"fsType={fstype}") if fstype is not None else "fsType: n/a"
        notes.append(f"      kind={kind}  glyphs={n_glyphs}  cmap_codepoints={n_cmap}  unitsPerEm={upm}")
        notes.append(f"      coverage: {verdict}")
        notes.append(f"      embedding: {fs_note}")
        tt.close()
    except Exception as exc:
        notes.append(f"      [fontTools analysis failed: {exc}]")
    try:
        from PIL import ImageFont
        ImageFont.truetype(str(font_path), 12)
        notes.append("      Pillow load: OK (measurable)")
    except Exception as exc:
        notes.append(f"      Pillow load: FAILED ({exc})")
    return "\n".join(notes)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx", help="path to the .pptx")
    ap.add_argument("-o", "--out", default="extracted_fonts", help="output directory")
    args = ap.parse_args()

    pptx = Path(args.pptx)
    if not pptx.exists():
        print(f"❌ not found: {pptx}")
        return 1

    out_dir = Path(args.out)
    with zipfile.ZipFile(pptx) as zf:
        names = zf.namelist()
        font_members = [n for n in names if n.startswith("ppt/fonts/")]
        embedded = _parse_embedded_font_list(zf)
        rels = _parse_rels(zf)

        print(f"📂 {pptx.name}")
        print(f"   ppt/fonts/ entries: {len(font_members)}")
        print(f"   <embeddedFont> declarations: {len(embedded)}")

        if not font_members:
            print()
            print("⚠️  NO embedded fonts in this file.")
            print("   The font is NOT inside the .pptx, so it cannot be extracted.")
            print("   Ask the client to re-save with: File > Options > Save >")
            print("     ☑ Embed fonts in the file  →  ◉ Embed all characters (NOT just used)")
            print("   ('Embed all characters' is required to measure NEW text; the default")
            print("    'embed only characters used' produces a subset that can't measure new commentary.)")
            return 0

        out_dir.mkdir(parents=True, exist_ok=True)
        # Map rId -> typeface/style for nice filenames.
        rid_label: dict[str, str] = {}
        for e in embedded:
            for style, rid in e["styles"].items():
                rid_label[rid] = f"{e['typeface']}-{style}"

        extracted: list[Path] = []
        # Extract via rels mapping when possible, else dump every fntdata.
        targets = rels if rels else {f"_{i}": m for i, m in enumerate(font_members)}
        for rid, member in targets.items():
            if member not in names:
                continue
            data = zf.read(member)
            ext = _detect_ext(data[:8])
            label = rid_label.get(rid, Path(member).stem)
            safe = "".join(c if c.isalnum() or c in "-_." else "_" for c in label)
            dest = out_dir / f"{safe}{ext}"
            dest.write_bytes(data)
            extracted.append(dest)
            print(f"\n✅ {label}  →  {dest}  ({len(data):,} bytes, {ext})")
            print(_analyse(dest))

        print(f"\nDone. {len(extracted)} font file(s) in {out_dir}/")
        print("Next: if coverage is FULL, point text_metrics.resolve_font_path at these files")
        print("      (pass the absolute path as the font 'family') to measure with the client's exact font.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

"""Dump a font's measurement metrics as TEXT (JSON) — for text-only egress channels.

Run this IN the environment that HAS the font (or the .pptx with the embedded font).
It emits a compact JSON metrics table to stdout. Copy that text out and feed it to
the generation side via fdd_utils.text_metrics.MetricsTable.from_json(...). No font
binary needs to leave the secure environment — only advance widths + global metrics,
which is all that line-wrapping/fit calculation requires (glyph OUTLINES are not needed).

Usage:
    python dump_font_metrics.py FONT.ttf            > metrics.json
    python dump_font_metrics.py Client.pptx          > metrics.json   # extracts embedded font(s)
    python dump_font_metrics.py FONT.ttf --ascii-only > metrics.json  # tiniest output

Compression: most glyphs in a font share one advance (the em width, esp. for CJK which is
full-width), so we store a single default_advance + only the EXCEPTIONS. A full CJK font
(~20k chars) compresses to a few hundred exception entries → a few KB of pasteable text.
"""
from __future__ import annotations

import argparse
import io
import json
import sys
import zipfile
from collections import Counter


_SFNT_SIGS = (b"\x00\x01\x00\x00", b"OTTO", b"true", b"ttcf")


def _unwrap_eot(data: bytes) -> bytes:
    """Return the raw TTF/OTF inside a PowerPoint .fntdata (EOT) blob.

    EOT layout: [EOTSize u32][FontDataSize u32][Version u32][Flags u32]...[fontdata].
    Font data sits at the end and may be XOR-masked with 0x50. Returns the input
    unchanged if it already looks like a plain sfnt; raises if it can't be unwrapped.
    """
    import struct
    if data[:4] in _SFNT_SIGS:
        return data
    cands = []
    if len(data) >= 16:
        eot_size, fds, _ver, flags = struct.unpack_from("<IIII", data, 0)
        segs = []
        if 0 < fds <= len(data):
            segs.append(data[len(data) - fds:])
        if 0 < eot_size <= len(data) and 0 < fds <= eot_size:
            segs.append(data[eot_size - fds: eot_size])
        for seg in segs:
            cands.append(seg)
            if flags & 0x10000000:  # TTEMBED_XORENCRYPTDATA
                cands.append(bytes(b ^ 0x50 for b in seg))
        if flags & 0x00000004:  # TTEMBED_TTCOMPRESSED
            raise ValueError("EOT MicroType-compressed — cannot decode (ask client for the .ttf)")
    for sig in _SFNT_SIGS:
        i = data.find(sig)
        if i > 0:
            cands.append(data[i:])
    for c in cands:
        if c[:4] in _SFNT_SIGS:
            return c
    raise ValueError("not a TTF/OTF and not unwrappable (first16=%s)" % data[:16].hex())


def _load_embedded_from_pptx(path: str) -> list[tuple[str, bytes]]:
    out: list[tuple[str, bytes]] = []
    with zipfile.ZipFile(path) as zf:
        for name in zf.namelist():
            if name.startswith("ppt/fonts/"):
                out.append((name, zf.read(name)))
    return out


def _metrics_from_bytes(data: bytes, source: str, ascii_only: bool) -> dict:
    from fontTools.ttLib import TTFont
    data = _unwrap_eot(data)   # PowerPoint embeds fonts as EOT — unwrap to raw sfnt first
    tt = TTFont(io.BytesIO(data), fontNumber=0, lazy=True)

    upm = int(tt["head"].unitsPerEm)
    hhea = tt["hhea"] if "hhea" in tt else None
    os2 = tt["OS/2"] if "OS/2" in tt else None
    name_tbl = tt["name"] if "name" in tt else None
    family = None
    if name_tbl is not None:
        rec = name_tbl.getName(1, 3, 1, 0x409) or name_tbl.getName(1, 1, 0, 0)
        if rec:
            family = str(rec)

    cmap = tt.getBestCmap() if tt.get("cmap") else {}
    hmtx = tt["hmtx"]

    def adv(gname: str) -> int:
        try:
            return int(hmtx[gname][0])
        except Exception:
            return upm

    advances: dict[int, int] = {}
    for cp, gname in cmap.items():
        if ascii_only and not (cp < 0x3000 or 0x3000 <= cp <= 0x303F or 0xFF00 <= cp <= 0xFFEF):
            continue
        advances[cp] = adv(gname)

    default_advance = Counter(advances.values()).most_common(1)[0][0] if advances else upm
    exceptions = {str(cp): w for cp, w in advances.items() if w != default_advance}

    has_cjk = any(0x4E00 <= cp <= 0x9FFF for cp in cmap)
    return {
        "schema": "fdd-font-metrics/1",
        "family": family or source,
        "source": source,
        "units_per_em": upm,
        "ascent": int(hhea.ascent) if hhea else (int(os2.sTypoAscender) if os2 else upm),
        "descent": int(hhea.descent) if hhea else (int(os2.sTypoDescender) if os2 else 0),
        "line_gap": int(hhea.lineGap) if hhea else (int(os2.sTypoLineGap) if os2 else 0),
        "cjk_full_width": has_cjk,
        "cmap_codepoints": len(cmap),
        "default_advance": int(default_advance),
        "exceptions": exceptions,   # codepoint(str) -> advance(units); everything else = default
    }


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("font_or_pptx")
    ap.add_argument("--ascii-only", action="store_true",
                    help="dump only ASCII/Latin + CJK punctuation widths (smallest output)")
    args = ap.parse_args()

    p = args.font_or_pptx
    payloads: list[tuple[str, bytes]] = []
    if p.lower().endswith((".pptx", ".potx", ".ppsx", ".zip")):
        payloads = _load_embedded_from_pptx(p)
        if not payloads:
            print("// NO embedded fonts in this file — ask client to embed (all characters).",
                  file=sys.stderr)
            return 2
    else:
        with open(p, "rb") as fh:
            payloads = [(p, fh.read())]

    results = []
    for source, data in payloads:
        try:
            results.append(_metrics_from_bytes(data, source, args.ascii_only))
        except Exception as exc:  # pragma: no cover - defensive
            print(f"// failed on {source}: {exc}", file=sys.stderr)

    out = results[0] if len(results) == 1 else {"schema": "fdd-font-metrics/1", "fonts": results}
    # Compact but valid JSON — this is the TEXT you copy out.
    print(json.dumps(out, ensure_ascii=True, separators=(",", ":"), sort_keys=True))
    sizes = [r["family"] + f" (exceptions={len(r['exceptions'])})" for r in results]
    print(f"// dumped: {', '.join(sizes)}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

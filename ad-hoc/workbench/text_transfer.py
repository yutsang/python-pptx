#!/usr/bin/env python3
"""Carry files through a channel that only carries text, and check they arrived whole. Standard library only: copy this file as it is.

    # on the machine that has the files
    python text_transfer.py encode FILE_OR_FOLDER [FILE_OR_FOLDER ...] --out OUT_FOLDER [--chunk-kb 200]

    # on the machine that receives the text (the parts may be pasted into .md or .txt files, in any order, with other text around them)
    python text_transfer.py decode PART_FILE_OR_FOLDER [...] --out OUT_FOLDER

What travels is the files' exact bytes (tar, then zlib, then base64), so an Excel file arrives with the same sha256 it left with: its
indentation, formats, hidden sheets and everything else are in it because nothing was read out of it. Each part ends with a CRC of its
own text, and the archive carries every file's sha256, so a damaged or missing part is named and nothing is written from a bad copy.

--chunk-kb 0 writes one part. A part is plain text, 76 characters a line; a code fence around it, a quote mark in front of each line or
trailing spaces from a chat window do not matter. The text holds whatever was encoded, so treat it like the files themselves.
"""
from __future__ import annotations

import argparse
import base64
import hashlib
import io
import json
import os
import re
import sys
import tarfile
import zlib
from pathlib import Path
from typing import Dict, List, Optional, Tuple

VERSION = 1
LINE = 76
MANIFEST_NAME = "__manifest__.json"
HEADER = re.compile(r"FDDXFER\s+(\d+)\s+part\s+(\d+)/(\d+)\s+([0-9a-f]{16})")
CRC = re.compile(r"^\W*CRC\s+([0-9a-f]{8})\W*$")


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def collect(paths: List[str]) -> List[Tuple[str, Path]]:
    """[(name inside the archive, file)] -- a folder keeps its own name as the first part of every path in it."""
    out: List[Tuple[str, Path]] = []
    for raw in paths:
        p = Path(raw)
        if p.is_file():
            out.append((p.name, p))
        elif p.is_dir():
            for f in sorted(p.rglob("*")):
                if f.is_file():
                    out.append(((Path(p.name) / f.relative_to(p)).as_posix(), f))
        else:
            raise SystemExit(f"not found: {raw}")
    names = [n for n, _ in out]
    if len(set(names)) != len(names):
        raise SystemExit("two of the files would have the same name inside the archive")
    if not out:
        raise SystemExit("nothing to encode")
    return out


def build_archive(files: List[Tuple[str, Path]]) -> Tuple[bytes, Dict[str, dict]]:
    listing = {name: {"bytes": f.stat().st_size, "sha256": sha256_bytes(f.read_bytes())} for name, f in files}
    manifest = json.dumps({"version": VERSION, "files": listing}, ensure_ascii=False, indent=1).encode("utf-8")
    buf = io.BytesIO()
    with tarfile.open(fileobj=buf, mode="w") as tar:
        info = tarfile.TarInfo(MANIFEST_NAME)
        info.size = len(manifest)
        tar.addfile(info, io.BytesIO(manifest))
        for name, f in files:
            data = f.read_bytes()
            info = tarfile.TarInfo(name)
            info.size = len(data)
            tar.addfile(info, io.BytesIO(data))
    return zlib.compress(buf.getvalue(), 9), listing


def encode(paths: List[str], out: Path, chunk_kb: int = 200) -> dict:
    files = collect(paths)
    packed, listing = build_archive(files)
    archive_id = sha256_bytes(packed)[:16]
    text = base64.b64encode(packed).decode("ascii")
    lines = [text[i:i + LINE] for i in range(0, len(text), LINE)]
    per_part = max(1, (chunk_kb * 1024) // (LINE + 1)) if chunk_kb > 0 else len(lines)
    groups = [lines[i:i + per_part] for i in range(0, len(lines), per_part)] or [[]]
    out.mkdir(parents=True, exist_ok=True)
    if any(out.glob("part-*.txt")):
        raise SystemExit("the output folder already holds parts; use an empty one")
    for n, group in enumerate(groups, 1):
        payload = "".join(group)
        body = [f"FDDXFER {VERSION} part {n}/{len(groups)} {archive_id}", *group, f"CRC {zlib.crc32(payload.encode('ascii')) & 0xffffffff:08x}"]
        (out / f"part-{n:03d}.txt").write_text("\n".join(body) + "\n", encoding="ascii")
    (out / "README.txt").write_text(
        f"{len(groups)} part(s), archive {archive_id}. Paste each part's text into a file on the other machine (.md or .txt, in any order, other text\n"
        "around it is fine), then:  python text_transfer.py decode <those files or their folder> --out <empty folder>\n", encoding="utf-8")
    return {"files": len(files), "bytes": sum(v["bytes"] for v in listing.values()), "parts": len(groups), "text_chars": len(text), "archive": archive_id}


def read_parts(sources: List[str]) -> Tuple[Dict[int, str], Optional[str], int, List[str]]:
    """({part number: payload}, archive id, parts expected, problems) from every file named, or every file in a folder named."""
    texts: List[str] = []
    for raw in sources:
        p = Path(raw)
        for f in (sorted(x for x in p.rglob("*") if x.is_file()) if p.is_dir() else [p]):
            try:
                texts.append(f.read_text(encoding="utf-8", errors="ignore"))
            except OSError:
                continue
    parts: Dict[int, str] = {}
    problems: List[str] = []
    archive_id: Optional[str] = None
    expected = 0
    for text in texts:
        lines = text.splitlines()
        i = 0
        while i < len(lines):
            m = HEADER.search(lines[i])
            if not m:
                i += 1
                continue
            version, number, total, ident = int(m.group(1)), int(m.group(2)), int(m.group(3)), m.group(4)
            payload: List[str] = []
            i += 1
            crc: Optional[str] = None
            while i < len(lines):
                c = CRC.match(lines[i])
                if c:
                    crc = c.group(1)
                    break
                if HEADER.search(lines[i]):
                    break
                cleaned = re.sub(r"[^A-Za-z0-9+/=]", "", lines[i])
                if cleaned and not lines[i].strip().startswith("```"):
                    payload.append(cleaned)
                i += 1
            if version != VERSION:
                problems.append(f"part {number}: written by another version of this tool ({version})")
                continue
            joined = "".join(payload)
            if crc is None:
                problems.append(f"part {number}/{total}: the CRC line is missing (the text was cut off)")
                continue
            if f"{zlib.crc32(joined.encode('ascii')) & 0xffffffff:08x}" != crc:
                problems.append(f"part {number}/{total}: damaged in copying (CRC differs); copy it again")
                continue
            if archive_id is not None and ident != archive_id:
                problems.append(f"part {number}: belongs to another archive ({ident}, not {archive_id})")
                continue
            archive_id, expected = ident, total
            parts[number] = joined
    return parts, archive_id, expected, problems


def safe_target(out: Path, name: str) -> Path:
    target = (out / name).resolve()
    if os.path.isabs(name) or ".." in Path(name).parts or out.resolve() not in target.parents:
        raise SystemExit(f"refused: {name!r} would be written outside the output folder")
    return target


def decode(sources: List[str], out: Path) -> dict:
    parts, archive_id, expected, problems = read_parts(sources)
    if not parts and not problems:
        return {"ok": False, "problems": ["no part of an archive was found in what was given"]}
    missing = [n for n in range(1, expected + 1) if n not in parts]
    if missing:
        problems.append(f"missing part(s): {', '.join(map(str, missing))} of {expected}")
    if problems:
        return {"ok": False, "problems": problems, "have": sorted(parts), "of": expected}
    packed = base64.b64decode("".join(parts[n] for n in range(1, expected + 1)))
    if sha256_bytes(packed)[:16] != archive_id:
        return {"ok": False, "problems": ["the parts are all intact but do not add up to the archive they name"]}
    raw = zlib.decompress(packed)
    with tarfile.open(fileobj=io.BytesIO(raw), mode="r") as tar:
        members = {m.name: m for m in tar.getmembers() if m.isfile()}
        manifest = json.loads(tar.extractfile(members[MANIFEST_NAME]).read().decode("utf-8"))
        wanted = manifest["files"]
        if out.exists() and any(out.iterdir()):
            return {"ok": False, "problems": ["the output folder is not empty; nothing was written"]}
        bad = []
        staged: Dict[str, bytes] = {}
        for name, meta in wanted.items():
            if name not in members:
                bad.append(f"{name}: not in the archive")
                continue
            data = tar.extractfile(members[name]).read()
            if sha256_bytes(data) != meta["sha256"] or len(data) != meta["bytes"]:
                bad.append(f"{name}: arrived changed")
            staged[name] = data
        for name in staged:
            safe_target(out, name)
        if bad:
            return {"ok": False, "problems": bad}
        for name, data in staged.items():
            target = safe_target(out, name)
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(data)
    return {"ok": True, "files": len(staged), "bytes": sum(len(d) for d in staged.values()), "archive": archive_id,
            "sha256": {name: wanted[name]["sha256"] for name in staged}}


def main(argv: Optional[List[str]] = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except (ValueError, OSError):
                pass
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    e = sub.add_parser("encode")
    e.add_argument("paths", nargs="+")
    e.add_argument("--out", required=True)
    e.add_argument("--chunk-kb", type=int, default=200, help="text per part, in KB; 0 = one part")
    d = sub.add_parser("decode")
    d.add_argument("sources", nargs="+")
    d.add_argument("--out", required=True)
    args = ap.parse_args(argv)
    if args.cmd == "encode":
        r = encode(args.paths, Path(args.out), args.chunk_kb)
        print(f"encoded {r['files']} file(s), {r['bytes']:,} bytes into {r['parts']} part(s), {r['text_chars']:,} characters of text; archive {r['archive']}")
        return 0
    r = decode(args.sources, Path(args.out))
    if not r["ok"]:
        print("NOT RESTORED:")
        for p in r["problems"]:
            print("  -", p)
        if r.get("of"):
            print(f"  have part(s) {r['have']} of {r['of']}")
        return 2
    print(f"restored {r['files']} file(s), {r['bytes']:,} bytes; archive {r['archive']}; every file's sha256 matches the one it left with")
    return 0


if __name__ == "__main__":
    sys.exit(main())

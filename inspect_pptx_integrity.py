"""Structural OOXML integrity checker for a .pptx file.

python-pptx's own reader is lenient (it happily opens files that PowerPoint
itself would flag as needing repair), so a clean `Presentation(path)` load
does NOT prove the file is valid OOXML. This script re-checks the raw zip
archive against the rules PowerPoint actually enforces:

  1. No two zip members share the same archive name (a duplicate name is
     invalid in a zip and silently shadows one entry with the other).
  2. Every part's content-type is declared in [Content_Types].xml (via an
     <Override> for that exact partname, or a <Default> for its extension).
  3. Every relationship (package-level _rels/.rels, and each part's own
     <partdir>/_rels/<partname>.rels) whose TargetMode is not "External"
     points at a zip member that actually exists.
  4. No .rels file declares the same r:id twice.
  5. presentation.xml's <p:sldIdLst> has no duplicate sldId values, and
     every r:id it references resolves in ppt/_rels/presentation.xml.rels.

Usage: python inspect_pptx_integrity.py <file.pptx> [more.pptx ...]
Exit code is non-zero if any file has problems.
"""
import posixpath
import sys
import zipfile
from xml.etree import ElementTree as ET

REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
CT_NS = "{http://schemas.openxmlformats.org/package/2006/content-types}"
P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def _rels_path_for(partname: str) -> str:
    d, f = posixpath.split(partname)
    return posixpath.join(d, "_rels", f + ".rels")


def check(path: str) -> list:
    problems = []
    zf = zipfile.ZipFile(path)
    names = zf.namelist()

    # 1. duplicate zip member names
    seen = {}
    for n in names:
        seen[n] = seen.get(n, 0) + 1
    for n, count in seen.items():
        if count > 1:
            problems.append(f"DUPLICATE ZIP ENTRY: {n!r} appears {count} times -- invalid zip/OPC package")

    name_set = set(names)

    # 2. content-types coverage
    try:
        ct_root = ET.fromstring(zf.read("[Content_Types].xml"))
    except KeyError:
        problems.append("MISSING [Content_Types].xml")
        ct_root = None
    if ct_root is not None:
        defaults = {el.get("Extension").lower() for el in ct_root.findall(f"{CT_NS}Default")}
        overrides = {el.get("PartName") for el in ct_root.findall(f"{CT_NS}Override")}
        for n in names:
            if n in ("[Content_Types].xml",) or n.endswith(".rels"):
                continue
            partname = "/" + n
            ext = posixpath.splitext(n)[1].lstrip(".").lower()
            if partname not in overrides and ext not in defaults:
                problems.append(f"NO CONTENT-TYPE declared for {partname!r} (ext={ext!r})")

    # 3 & 4. relationship targets resolve, no duplicate rIds
    rels_files = [n for n in names if n.endswith(".rels")]
    for rels_name in rels_files:
        try:
            root = ET.fromstring(zf.read(rels_name))
        except ET.ParseError as e:
            problems.append(f"UNPARSEABLE RELS: {rels_name} ({e})")
            continue
        base_dir = posixpath.dirname(posixpath.dirname(rels_name))  # strip "_rels/x.rels"
        seen_rids = {}
        for rel in root.findall(f"{REL_NS}Relationship"):
            rid = rel.get("Id")
            seen_rids[rid] = seen_rids.get(rid, 0) + 1
            mode = rel.get("TargetMode", "Internal")
            if mode == "External":
                continue
            target = rel.get("Target")
            resolved = posixpath.normpath(posixpath.join(base_dir, target)).lstrip("/")
            if resolved not in name_set:
                problems.append(f"DANGLING RELATIONSHIP: {rels_name} rId={rid} -> {resolved!r} (not in package)")
        for rid, count in seen_rids.items():
            if count > 1:
                problems.append(f"DUPLICATE RELATIONSHIP ID: {rels_name} rId={rid} declared {count} times")

    # 5. sldIdLst sanity
    if "ppt/presentation.xml" in name_set:
        pres_root = ET.fromstring(zf.read("ppt/presentation.xml"))
        pres_rels_name = "ppt/_rels/presentation.xml.rels"
        pres_rids = set()
        if pres_rels_name in name_set:
            pres_rels_root = ET.fromstring(zf.read(pres_rels_name))
            pres_rids = {rel.get("Id") for rel in pres_rels_root.findall(f"{REL_NS}Relationship")}
        sld_ids = {}
        for sld in pres_root.findall(f"{P_NS}sldIdLst/{P_NS}sldId"):
            sld_id = sld.get("id")
            sld_ids[sld_id] = sld_ids.get(sld_id, 0) + 1
            rid = sld.get(f"{R_NS}id")
            if rid not in pres_rids:
                problems.append(f"SLIDE RELATIONSHIP MISSING: sldId={sld_id} r:id={rid} not in {pres_rels_name}")
        for sld_id, count in sld_ids.items():
            if count > 1:
                problems.append(f"DUPLICATE SLIDE ID: sldId={sld_id} appears {count} times")

    return problems


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    any_problems = False
    for path in sys.argv[1:]:
        print(f"\n=== {path} ===")
        problems = check(path)
        if not problems:
            print("  OK -- no structural problems found")
        else:
            any_problems = True
            for p in problems:
                print(f"  ! {p}")
    sys.exit(1 if any_problems else 0)


if __name__ == "__main__":
    main()

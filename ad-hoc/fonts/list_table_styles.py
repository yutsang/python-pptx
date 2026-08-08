"""List table styles + the style GUID each table uses, in a .pptx.

Use this to find your firm's UpSlide table-style GUID, then either:
  - leave config.yml pptx.table_style_id blank (auto-detect uses a table already
    in the template), OR
  - paste the GUID into config.yml pptx.table_style_id to force it.

Usage:
    python list_table_styles.py "template.pptx"
"""
from __future__ import annotations

import sys
import zipfile
from xml.etree import ElementTree as ET

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def main() -> int:
    if len(sys.argv) < 2:
        print("usage: python list_table_styles.py <file.pptx>")
        return 1
    path = sys.argv[1]

    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()

        # 1. Registered table styles (ppt/tableStyles.xml): name + GUID.
        print("== Registered table styles (ppt/tableStyles.xml) ==")
        if "ppt/tableStyles.xml" in names:
            root = ET.fromstring(zf.read("ppt/tableStyles.xml"))
            default = root.attrib.get("def", "")
            found = False
            for el in root.iter():
                if _local(el.tag) == "tblStyle":
                    found = True
                    sid = el.attrib.get("styleId", "")
                    nm = el.attrib.get("styleName", "")
                    mark = "  <- default" if sid == default else ""
                    print(f"  {sid}  |  {nm}{mark}")
            if not found:
                print("  (none defined in tableStyles.xml)")
        else:
            print("  (no ppt/tableStyles.xml — only PowerPoint built-in styles are used)")

        # 2. Which GUID each actual table in the deck references.
        print("\n== Style GUID used by each table in the slides ==")
        any_table = False
        slide_names = sorted(n for n in names if n.startswith("ppt/slides/slide") and n.endswith(".xml"))
        for sname in slide_names:
            root = ET.fromstring(zf.read(sname))
            for tbl in root.iter():
                if _local(tbl.tag) != "tbl":
                    continue
                any_table = True
                sid = None
                for child in tbl.iter():
                    if _local(child.tag) == "tableStyleId":
                        sid = (child.text or "").strip()
                        break
                print(f"  {sname.split('/')[-1]:<16}  styleId = {sid or '(default / none)'}")
        if not any_table:
            print("  (no tables found in any slide)")

    print("\nTip: paste the GUID your firm's styled table uses into config.yml")
    print("     pptx.table_style_id  (or leave blank to auto-detect it).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

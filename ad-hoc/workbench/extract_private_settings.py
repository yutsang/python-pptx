#!/usr/bin/env python3
"""Write private-config-additions.yml from an OLD checkout whose gateway header names and slide source line were hard-coded in code.

    python ad-hoc/workbench/extract_private_settings.py [--root FOLDER_WITH_THE_OLD_CODE] [--out private-config-additions.yml]

Run it BEFORE the old files are replaced by the public version. It reads fdd_utils/ai/client.py, fdd_utils/pptx/gen_tables.py and (if it still
has them) ad-hoc/workbench/test_dashscope_connectivity.py, and writes only the three settings that moved into config:

    workbench.header_names   charge_code / region_override -> the HTTP header names the gateway wants
    dashscope.header_names   the same, for the connectivity probe
    pptx.source_labels       the slide source line, zh and en

The output file is gitignored and holds company-specific text, so this script prints which settings it found and how long each is, never the
text. Then:  python apply_private_settings.py fdd_utils/config.yml private-config-additions.yml
Standard library plus PyYAML (already a requirement of the tool).
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import yaml

CLIENT = "fdd_utils/ai/client.py"
TABLES = "fdd_utils/pptx/gen_tables.py"
PROBE = "ad-hoc/workbench/test_dashscope_connectivity.py"

# 'name': str(self.config_details.get('charge_code') or '0000')
_CLIENT_HEADER = re.compile(r"""['"]([^'"]+)['"]\s*:\s*str\(\s*self\.config_details\.get\(\s*['"](charge_code|region_override)['"]""")
# "name": charge_code,
_PROBE_HEADER = re.compile(r"""['"]([^'"]+)['"]\s*:\s*(charge_code|region_override)\s*,""")
# source_run.text = "zh text" if is_chinese_databook else "en text"
_SOURCE = re.compile(r"""source_run\.text\s*=\s*(["'])(.+?)\1\s+if\s+is_chinese_databook\s+else\s+(["'])(.+?)\3""")


def headers_from(text: str, pattern: "re.Pattern[str]") -> Dict[str, str]:
    found: Dict[str, str] = {}
    for name, field in pattern.findall(text):
        found.setdefault(field, name)
    return found


def source_labels_from(text: str) -> Dict[str, str]:
    m = _SOURCE.search(text)
    return {"zh": m.group(2), "en": m.group(4)} if m else {}


def read(root: Path, rel: str) -> Optional[str]:
    p = root / rel
    return p.read_text(encoding="utf-8") if p.is_file() else None


def extract(root: Path) -> Tuple[dict, List[str]]:
    """(additions, report lines without any of the extracted text)."""
    additions: dict = {}
    report: List[str] = []
    client = read(root, CLIENT)
    workbench = headers_from(client, _CLIENT_HEADER) if client else {}
    if len(workbench) == 2:
        additions["workbench"] = {"header_names": workbench}
        report.append("workbench.header_names: found both (" + ", ".join(f"{k} {len(v)} chars" for k, v in workbench.items()) + ")")
    else:
        report.append(f"workbench.header_names: NOT found ({'file missing' if client is None else 'the old hard-coded form is not in ' + CLIENT}; "
                      "if this checkout is already the public version they are in config.yml instead)")
    probe = read(root, PROBE)
    dash = headers_from(probe, _PROBE_HEADER) if probe else {}
    if len(dash) == 2:
        additions["dashscope"] = {"header_names": dash}
        report.append("dashscope.header_names: found both (only the connectivity probe uses them)")
    else:
        report.append("dashscope.header_names: not found (only the connectivity probe uses it; skip unless you run that probe with auth: workbench)")
    tables = read(root, TABLES)
    labels = source_labels_from(tables) if tables else {}
    if labels:
        additions["pptx"] = {"source_labels": labels}
        report.append(f"pptx.source_labels: found zh ({len(labels['zh'])} chars) and en ({len(labels['en'])} chars)")
    else:
        report.append("pptx.source_labels: NOT found (" + ("file missing" if tables is None else "the old hard-coded form is not in " + TABLES) + ")")
    return additions, report


def main(argv: Optional[List[str]] = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except (ValueError, OSError):
                pass
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--root", default=".", help="folder that holds the old fdd_utils/ (default: current folder)")
    ap.add_argument("--out", default="private-config-additions.yml")
    args = ap.parse_args(argv)
    out = Path(args.out)
    if out.exists():
        print(f"{out} already exists; nothing written. Review it, or move it away and run again.")
        return 2
    additions, report = extract(Path(args.root))
    for line in report:
        print(line)
    if not additions:
        print("nothing found, nothing written")
        return 1
    out.write_text(yaml.safe_dump(additions, allow_unicode=True, sort_keys=False), encoding="utf-8")
    print(f"wrote {out} with {', '.join(additions)}; the text itself was not printed")
    return 0


if __name__ == "__main__":
    sys.exit(main())

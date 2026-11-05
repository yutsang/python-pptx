#!/usr/bin/env python3
"""Look up the repo's own code by what it does. Free, a few seconds, no model; prints names, places and trust evidence, never client data.

    python ad-hoc/workbench/code_map.py find "characters per line wrap width" [-k 8]
    python ad-hoc/workbench/code_map.py explain fdd_utils.text_metrics.wrap_text_with_metrics [--source 40]
    python ad-hoc/workbench/code_map.py overview [fdd_utils/pptx]

The same three lookups are agent tools (find_capability, explain_capability, code_overview). English words match best.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from fdd_utils.agent import codemap  # noqa: E402


def main(argv=None) -> int:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except (ValueError, OSError):
                pass
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    f = sub.add_parser("find")
    f.add_argument("query")
    f.add_argument("-k", type=int, default=8)
    e = sub.add_parser("explain")
    e.add_argument("name")
    e.add_argument("--source", type=int, default=0)
    o = sub.add_parser("overview")
    o.add_argument("package", nargs="?", default="")
    args = ap.parse_args(argv)
    m = codemap.get_map()
    if args.cmd == "find":
        for h in m.find(args.query, args.k):
            print(f"{h['name']}  [{h['where']}]  {h['lines']} lines, {h['callers']} caller(s), {len(h['tested_by'])} test file(s), {h['markers']} warning/measurement comment(s)")
            print(f"    {h['signature']}")
            if h["summary"]:
                print(f"    {h['summary']}")
            print(f"    touches: {', '.join(h['signals']) or 'nothing flagged'}   |   {h['reuse']}")
        return 0
    if args.cmd == "explain":
        print(json.dumps(m.explain(args.name, args.source), ensure_ascii=False, indent=1))
        return 0
    print(json.dumps(m.overview(args.package), ensure_ascii=False, indent=1))
    return 0


if __name__ == "__main__":
    sys.exit(main())

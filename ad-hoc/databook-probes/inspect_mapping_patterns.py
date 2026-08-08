#!/usr/bin/env python3
"""Audits the per-account `patterns` in mappings.yml -- what they are, where
they go, and which ones risk putting words in the model's mouth.

What they do: PromptEngine.render_prompt injects an account's patterns into
the subagent_1 (Generator) prompt only, rendered as "Example 1: ...",
"Example 2: ...". The Auditor and Validator never see them. So they shape
the FIRST draft's phrasing and structure, and nothing else.

Why that needs auditing: several are not style templates but complete
sentences asserting FACTS -- that bank statements were or were not
obtained, that no differences were noted. An example is the strongest
signal a prompt can carry; a model shown "We have not obtained the bank
statements yet" will tend to write it whether or not the databook says so.
Two patterns for the same account can even assert opposite things.

Flags raised per pattern:
  * canned assurance -- claims work was performed, or that nothing was
    found, as fixed text rather than something drawn from the remarks;
  * canned negative -- asserts something was NOT done/obtained;
  * fully-formed sentence -- a complete sentence with only <PLACEHOLDER>
    slots, which invites filling in the blanks rather than writing;
  * contradiction -- two patterns for one account assert opposite things.

Read-only.

Usage:
    python inspect_mapping_patterns.py
    python inspect_mapping_patterns.py --account Cash --show-full
"""

# moved into ad-hoc/ -- put the repo root back on sys.path so
# `import fdd_utils...` still resolves when run from anywhere.
import sys as _sys
from pathlib import Path as _Path
_sys.path.insert(0, str(_Path(__file__).resolve().parents[2]))
import argparse
import re
import sys

sys.path.insert(0, ".")

import yaml

_ASSURANCE_POSITIVE = (
    "未发现", "未见", "无重大", "无差异", "已核对", "核对了", "核对至", "检查了", "获取并",
    "no material difference", "no differences", "no significant difference",
    "we checked", "we obtained", "we have checked", "we agreed",
)
_ASSURANCE_NEGATIVE = (
    "尚未取得", "尚未获取", "未取得", "未获取",
    "have not obtained", "have not received", "not yet obtained", "have not been provided",
)
_PLACEHOLDER_RE = re.compile(r"<[A-Z_]+>")


def _flags_for(text: str):
    flags = []
    low = str(text or "").lower()
    pos = [w for w in _ASSURANCE_POSITIVE if w.lower() in low]
    neg = [w for w in _ASSURANCE_NEGATIVE if w.lower() in low]
    if pos:
        flags.append(("canned assurance", f"asserts work done / nothing found: {pos[:3]}"))
    if neg:
        flags.append(("canned negative", f"asserts something NOT done: {neg[:3]}"))
    placeholders = _PLACEHOLDER_RE.findall(str(text))
    stripped = _PLACEHOLDER_RE.sub("", str(text)).strip()
    # A complete sentence carrying real words around a few slots reads as a
    # fill-in-the-blanks template; a short structural hint does not.
    if placeholders and len(stripped) >= 25:
        flags.append(("fill-in-the-blanks",
                      f"{len(placeholders)} slot(s) inside {len(stripped)} chars of fixed prose"))
    return flags, bool(pos), bool(neg)


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--mappings", default="fdd_utils/mappings.yml")
    ap.add_argument("--account", default=None, help="only audit this mapping key")
    ap.add_argument("--show-full", action="store_true", help="print each pattern in full")
    args = ap.parse_args()

    data = yaml.safe_load(open(args.mappings, encoding="utf-8"))
    accounts = {k: v for k, v in data.items()
                if isinstance(v, dict) and v.get("patterns")}
    if args.account:
        accounts = {k: v for k, v in accounts.items() if k == args.account}
        if not accounts:
            print(f"❌ {args.account!r} has no patterns (or does not exist).")
            return 1

    print("=" * 78)
    print("MAPPING PATTERN AUDIT")
    print("=" * 78)
    print("Patterns are injected into the subagent_1 (Generator) prompt only,")
    print("as 'Example 1: ...'. The Auditor and Validator never see them.\n")

    total = flagged = 0
    by_flag = {}
    contradictions = []
    worst = []

    for key, cfg in sorted(accounts.items()):
        patterns = cfg["patterns"]
        items = list(patterns.items()) if isinstance(patterns, dict) else [("1", patterns)]
        acct_pos = acct_neg = False
        acct_flags = []
        for pname, text in items:
            if not text or str(text).strip().upper() == "N/A":
                continue
            total += 1
            flags, pos, neg = _flags_for(text)
            acct_pos = acct_pos or pos
            acct_neg = acct_neg or neg
            if flags:
                flagged += 1
                for kind, _d in flags:
                    by_flag[kind] = by_flag.get(kind, 0) + 1
                acct_flags.append((pname, text, flags))
        # A contradiction only matters when the variants are UNLABELLED. Where
        # each pattern's key states its precondition ("only where the remarks
        # show the statements were not obtained"), the conflict is deliberate
        # and useful -- it teaches the model to select on the data instead of
        # copying one arbitrarily. render_prompt passes those conditions
        # through to the example line.
        labelled = all(re.search(r"[（(].+[)）]\s*$", str(pn).strip()) for pn, _t in items)
        if acct_pos and acct_neg and not labelled:
            contradictions.append(key)
        if acct_flags:
            worst.append((key, acct_flags))

    for key, acct_flags in worst:
        print(f"--- {key}")
        for pname, text, flags in acct_flags:
            shown = str(text) if args.show_full else str(text)[:170]
            print(f"  [{pname}] {shown}")
            for kind, detail in flags:
                print(f"      ⚠️ {kind}: {detail}")
        if key in contradictions:
            print(f"      ❌ CONTRADICTION: this account's patterns assert BOTH that work was")
            print(f"         done/nothing found AND that something was not obtained. Whichever")
            print(f"         the model copies may be the opposite of what the databook says.")
        print()

    print("=" * 78)
    print("SUMMARY")
    print("=" * 78)
    print(f"  accounts with patterns : {len(accounts)}")
    print(f"  patterns total         : {total}")
    print(f"  patterns flagged       : {flagged} ({flagged / total * 100:.0f}%)" if total else "")
    for kind, n in sorted(by_flag.items(), key=lambda kv: -kv[1]):
        print(f"    {n:>3d} x  {kind}")
    if contradictions:
        print(f"\n  ❌ {len(contradictions)} account(s) whose patterns contradict each other:")
        print(f"     {', '.join(contradictions)}")
    print("\n  A pattern should show STRUCTURE and REGISTER -- how a sentence is built,")
    print("  which verbs to use -- not assert a fact. Anything the report says about")
    print("  work performed or differences found has to come from that account's own")
    print("  remarks, otherwise the example decides it and the data does not.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

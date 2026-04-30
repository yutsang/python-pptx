"""Inject anonymised patterns from _patterns_input.csv into mappings.yml.

For each (Lang, Items) row in the CSV, append the comment as a pattern to
the matching mapping's `patterns:` block (creating it if absent), keyed by
language. Existing project-team patterns are kept as `Pattern 1`; the
imported ones become `Pattern 2`, `Pattern 3`, … so the LLM sees both the
original abstract guidance and concrete real-world examples.

Items in the CSV that have no matching key in mappings.yml are reported on
stdout — those are the candidate "missing accounts" the user asked about.
"""

from __future__ import annotations

import csv
import sys
from collections import OrderedDict, defaultdict
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
CSV_PATH = REPO / "scripts" / "_patterns_input.csv"
YAML_PATH = REPO / "fdd_utils" / "mappings.yml"

# Normalise variants in the CSV to canonical mapping-keys used in mappings.yml.
ALIAS_TO_CANON = {
    "OCA": "Other CA",
    "Long-term deferred expense": "Long-term deferred expenses",
    "OC & GA": "OC",          # split combined headers into both items
    "OC & GP margin": "OC",
    "OI & OC": "OI",
}
# Combined CSV headers also push to a SECOND target (the GA / OC half).
COMBINED_SECONDARY = {
    "OC & GA": "GA",
    "OC & GP margin": "GA",
    "OI & OC": "OC",
}


def _load_csv():
    rows = []
    with CSV_PATH.open(encoding="utf-8") as f:
        for r in csv.DictReader(f):
            r["Items"] = (r.get("Items") or "").strip()
            r["Lang"] = (r.get("Lang") or "").strip()
            r["Comments"] = (r.get("Comments") or "").strip()
            if not r["Items"] or not r["Comments"]:
                continue
            rows.append(r)
    return rows


def _bucket_patterns(rows):
    """{(canon_item, Lang): [text, ...]} keeping insertion order, deduped."""
    buckets: dict[tuple[str, str], list[str]] = defaultdict(list)
    for r in rows:
        item = r["Items"]
        targets = [ALIAS_TO_CANON.get(item, item)]
        if item in COMBINED_SECONDARY:
            targets.append(COMBINED_SECONDARY[item])
        for tgt in targets:
            key = (tgt, r["Lang"])
            text = r["Comments"]
            if text not in buckets[key]:
                buckets[key].append(text)
    return buckets


def _read_yaml_lines() -> list[str]:
    return YAML_PATH.read_text(encoding="utf-8").splitlines(keepends=True)


def _detect_top_level_keys(lines: list[str]) -> "OrderedDict[str, tuple[int, int]]":
    """Return {key: (start_line, end_line_exclusive)} for every top-level item."""
    sections: OrderedDict[str, tuple[int, int]] = OrderedDict()
    current_key = None
    start = 0
    for i, line in enumerate(lines):
        if line and not line.startswith((" ", "\t", "#", "\n")) and ":" in line:
            head = line.split(":", 1)[0].strip()
            if head:
                if current_key is not None:
                    sections[current_key] = (start, i)
                current_key = head
                start = i
    if current_key is not None:
        sections[current_key] = (start, len(lines))
    return sections


def _patterns_block_lines(buckets_for_item: dict[str, list[str]]) -> list[str]:
    """Build the YAML block for one item's `patterns:` field, mixing Eng/Chi."""
    out = ["  patterns:\n"]
    n = 0
    # Keep order Eng → Chi for readability, max 3 of each language.
    for lang in ("EN", "CN"):
        for text in buckets_for_item.get(lang, [])[:3]:
            n += 1
            text_clean = text.replace("\r\n", "\n").replace("\r", "\n").rstrip()
            out.append(f"    Pattern {n}: |\n")
            for tline in text_clean.split("\n"):
                out.append(f"      {tline}\n")
    return out


def _replace_patterns_block(item_lines: list[str], new_block: list[str]) -> list[str]:
    """Rewrite the existing `  patterns:` block in this item's slice, or append."""
    out: list[str] = []
    skipping = False
    replaced = False
    for line in item_lines:
        if not skipping and line.startswith("  patterns:"):
            out.extend(new_block)
            skipping = True
            replaced = True
            continue
        if skipping:
            stripped_indent = line[: len(line) - len(line.lstrip())]
            # Stay inside the patterns block while it's indented further than 2.
            if line.strip() == "" or len(stripped_indent) >= 4:
                continue
            skipping = False
        out.append(line)
    if not replaced:
        # Append at end of slice with a leading newline if needed.
        if out and not out[-1].endswith("\n"):
            out[-1] = out[-1] + "\n"
        out.extend(new_block)
    return out


def main() -> int:
    rows = _load_csv()
    if not rows:
        print("No rows in CSV; aborting.", file=sys.stderr)
        return 2

    buckets = _bucket_patterns(rows)
    # Group by item: {item: {lang: [texts]}}
    by_item: dict[str, dict[str, list[str]]] = defaultdict(lambda: defaultdict(list))
    for (item, lang), texts in buckets.items():
        by_item[item][lang].extend(texts)

    lines = _read_yaml_lines()
    sections = _detect_top_level_keys(lines)

    matched = []
    missing = []
    new_lines = list(lines)

    # Process items in REVERSE file order so byte offsets stay valid as we
    # splice replacement blocks in-place.
    target_items = sorted(by_item.keys(), key=lambda k: -sections.get(k, (0, 0))[0])
    for item in target_items:
        if item not in sections:
            missing.append(item)
            continue
        start, end = sections[item]
        block = _patterns_block_lines(by_item[item])
        rewritten = _replace_patterns_block(new_lines[start:end], block)
        new_lines = new_lines[:start] + rewritten + new_lines[end:]
        matched.append(item)
        # Recompute sections because lengths may have shifted.
        sections = _detect_top_level_keys(new_lines)

    YAML_PATH.write_text("".join(new_lines), encoding="utf-8")
    print(f"Updated {len(matched)} items in mappings.yml:")
    for k in sorted(matched):
        print(f"  ✓ {k}")
    if missing:
        print(f"\nNot found in mappings.yml ({len(missing)} item(s)):")
        for k in sorted(missing):
            print(f"  ⚠ {k}")
        print("\nAdd these as new entries if you want them processed by the AI.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

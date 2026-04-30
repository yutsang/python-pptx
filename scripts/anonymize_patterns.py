"""Anonymise (脫敏) a patterns workbook for safe sharing.

Input  : an .xlsx with columns Lang, Project, Type, Items, Comments
         (header on row 1 — same shape as the redd_patterns.xlsx sample).
Output : a sibling .xlsx with the same shape but
           - the Project column blanked,
           - all numeric amounts replaced with the literal token "<AMT>",
           - all dates replaced with "<DATE>",
           - all percentages replaced with "<PCT>",
           - all Chinese / English entity names that match a heuristic
             "<Pinyin/Adjective> <Industry-keyword>" pattern replaced
             with "<ENTITY>".
         The sentence STRUCTURE (clauses, ordering, connectives,
         "1)/2)/3)" listing style) is preserved so the result is still
         useful as a writing-style reference for the LLM.

Usage (from repo root):

    python scripts/anonymize_patterns.py /path/to/redd_patterns.xlsx
    # → writes /path/to/redd_patterns.anonymised.xlsx
    #   …also dumps a CSV preview next to it for quick review.

Run with --dry-run to print 5 sample rows to stdout without writing.
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import List

import pandas as pd


# --- regex sanitisers ---------------------------------------------------------
# Order matters: more specific patterns first.

_DATE_PATTERNS: List[re.Pattern] = [
    # 31 December 2021 / 31 Dec 2021 / December 2021 / 2021-12-31
    re.compile(
        r"\b(?:\d{1,2}\s+)?"
        r"(January|February|March|April|May|June|July|August|"
        r"September|October|November|December|"
        r"Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)"
        r"(?:\s+\d{1,2})?(?:[,\s]+\d{4})\b",
        re.IGNORECASE,
    ),
    # 2021-12-31 / 2021/12/31 / 31/12/2021
    re.compile(r"\b\d{4}[-/]\d{1,2}[-/]\d{1,2}\b"),
    re.compile(r"\b\d{1,2}[-/]\d{1,2}[-/]\d{4}\b"),
    # Chinese date: 2021年12月31日 / 2021年12月
    re.compile(r"\d{4}\s*年\s*\d{1,2}\s*月(?:\s*\d{1,2}\s*日)?"),
    # FY21 / FY2021 / 1H21 / 2H21 / Q1 2021
    re.compile(r"\bFY\d{2,4}\b", re.IGNORECASE),
    re.compile(r"\b[12]H\d{2,4}\b", re.IGNORECASE),
    re.compile(r"\bQ[1-4](?:\s*\d{2,4})?\b", re.IGNORECASE),
]

_PERCENT = re.compile(r"-?\d+(?:\.\d+)?\s*%")

# Currency amounts — match BEFORE bare numbers.
_AMOUNT_PATTERNS: List[re.Pattern] = [
    re.compile(
        r"(?:CNY|RMB|USD|HKD|EUR|GBP|JPY|US\$|HK\$|RMB\$|\$)"
        r"\s*\d[\d,\.]*\s*"
        r"(?:thousand|million|billion|K|M|B|K\b|m\b|bn\b)?",
        re.IGNORECASE,
    ),
    # Chinese amount: 人民币59.3百万元 / 人民币0.2万元 / 59.3百万 / 4.4亿元
    re.compile(
        r"(?:人民币|港币|美元)?"
        r"\s*\d[\d,\.]*\s*"
        r"(?:百万元|百万|千万元|千万|万元|万|亿元|亿|千元|千)",
    ),
    # bare number with thousands separator (likely an amount): 1,234.56
    re.compile(r"\b\d{1,3}(?:,\d{3})+(?:\.\d+)?\b"),
    # bare decimal/integer with size suffix: 12.3M, 4.4 million, 59.3 百万
    re.compile(
        r"\b\d+(?:\.\d+)?\s*"
        r"(?:thousand|million|billion|百万|千万|万|亿|千)\b",
        re.IGNORECASE,
    ),
]

# Entity names: Pinyin/English word + industry/business keyword.
# Catches Guangdong JDL, Foshan Wanyuan, Dongguan XYZ Logistics, etc.
_ENTITY_KEYWORDS = (
    "Logistics|Industrial|Investment|Holding|Holdings|Property|Properties|"
    "Real Estate|Group|Enterprise|Limited|Ltd|Co\\.?|Corporation|Company|"
    "Bank|Branch|Construction|Development|Management|Services|"
    "Trading|Trust|Fund|Capital|Partners|Technology|Tech|Energy|Power|"
    "工业|地产|物业|集团|有限公司|股份|银行|分行|开发|管理|服务|物流|信托|基金|"
    "投资|实业|科技|能源|电力"
)
_ENTITY = re.compile(
    r"\b([A-Z][a-zA-Z]+(?:\s+[A-Z][a-zA-Z]+){0,3})\s+(?:" + _ENTITY_KEYWORDS + r")\b"
)
_ENTITY_CHI = re.compile(
    r"[\u4e00-\u9fff]{2,8}(?:" + _ENTITY_KEYWORDS + r")"
)
# "Guangdong JDL" / "Foshan ABC" — Capitalized place(s) followed by a 2-5
# letter all-caps acronym. Catches Pinyin-prefixed entity names that the
# keyword regex above misses because no industry word follows.
_ENTITY_ACRONYM = re.compile(
    r"\b(?:[A-Z][a-z]+\s+){1,3}[A-Z]{2,5}(?=\b)"
)


def _scrub(text: str) -> str:
    if not isinstance(text, str) or not text.strip():
        return text
    out = text
    for rx in _DATE_PATTERNS:
        out = rx.sub("<DATE>", out)
    out = _PERCENT.sub("<PCT>", out)
    for rx in _AMOUNT_PATTERNS:
        out = rx.sub("<AMT>", out)
    out = _ENTITY.sub("<ENTITY>", out)
    out = _ENTITY_CHI.sub("<ENTITY>", out)
    out = _ENTITY_ACRONYM.sub("<ENTITY>", out)
    # Final pass: bare numbers ≥ 4 digits (account ids, share counts, etc.)
    out = re.sub(r"\b\d{4,}\b", "<NUM>", out)
    # Collapse double spaces introduced by replacement.
    out = re.sub(r"[ \t]{2,}", " ", out).strip()
    return out


# --- workbook driver ----------------------------------------------------------

EXPECTED_COLS = ("Lang", "Project", "Type", "Items", "Comments")


def anonymise_workbook(in_path: Path, out_path: Path | None = None, dry_run: bool = False) -> Path:
    df = pd.read_excel(in_path, dtype=str).fillna("")
    missing = [c for c in EXPECTED_COLS if c not in df.columns]
    if missing:
        raise SystemExit(
            f"Workbook is missing expected columns {missing}. "
            f"Got columns: {list(df.columns)}"
        )

    df = df[list(EXPECTED_COLS)].copy()
    df["Project"] = ""  # blank the project column entirely
    df["Comments"] = df["Comments"].map(_scrub)
    # Items column may carry account names — those are generic
    # ("Cash", "AR"), so we leave them untouched.

    if dry_run:
        print(df.head(5).to_string(index=False))
        return in_path

    out_path = out_path or in_path.with_suffix(".anonymised.xlsx")
    df.to_excel(out_path, index=False)
    df.to_csv(out_path.with_suffix(".csv"), index=False)
    print(f"Wrote {out_path}")
    print(f"Wrote {out_path.with_suffix('.csv')} (preview)")
    return out_path


def main(argv: List[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("input", type=Path, help="Path to redd_patterns.xlsx")
    ap.add_argument("-o", "--output", type=Path, default=None, help="Output xlsx path")
    ap.add_argument("--dry-run", action="store_true", help="Print 5 rows, no write")
    args = ap.parse_args(argv)
    if not args.input.exists():
        print(f"Input not found: {args.input}", file=sys.stderr)
        return 2
    anonymise_workbook(args.input, args.output, dry_run=args.dry_run)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

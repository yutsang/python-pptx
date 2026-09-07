#!/usr/bin/env python3
"""Unit tests for typed defects, spans, fact provenance and the direction check.

Free, offline, no databook and no LLM: every frame here is built in this file,
so this runs on any machine and is the regression gate for the pieces of
validator.py that a replay cannot see (span arithmetic, serialisation, and the
direction check's skip rules, which fire on data no archived run contains).

    PYTHONPATH=. python ad-hoc/workbench/test_validator_defects.py
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pandas as pd
import yaml

REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from fdd_utils.ai.validator import (  # noqa: E402
    DEFECT_CODES,
    SourceIndex,
    collect_direction_findings,
    extract_amount_spans,
    segment_clauses,
    verify_commentary,
)
from fdd_utils.workbook import INTERNAL_ROW_KEY  # noqa: E402

FAILURES: list = []


def check(name: str, condition: bool, detail: str = "") -> None:
    if condition:
        print(f"  ok    {name}")
    else:
        FAILURES.append(name)
        print(f"  FAIL  {name}  {detail}")


def frame(rows, *, sheet="Cash", movements=None, row_types=None, notes=None) -> pd.DataFrame:
    """A minimal detail_analysis-shaped frame: label column, INTERNAL_ROW_KEY,
    one or more date columns, plus the attrs the verifier reads."""
    df = pd.DataFrame(rows)
    df.attrs["source_sheet_name"] = sheet
    df.attrs["source_multiplier"] = 1000
    df.attrs["integrity"] = {"sheet_name": sheet, "effective_date": "2025-12-31"}
    df.attrs["row_types_by_description"] = row_types or {}
    df.attrs["significant_movements"] = movements or []
    df.attrs["display_description_map"] = {r: r for r in df[df.columns[0]].tolist()}
    if notes:
        df.attrs["supporting_notes"] = notes
    return df


# ---------------------------------------------------------------------------
# amount spans
# ---------------------------------------------------------------------------

def test_amount_spans() -> None:
    print("\nextract_amount_spans — the span is the whole amount, unit included")
    cases = [
        ("余额为7.8万元，较上期", 78000.0, "7.8万元"),          # _AMT_WAN + 元 suffix
        ("增至7.8亿元。", 780000000.0, "7.8亿元"),              # _AMT_YI + 元 suffix
        ("人民币1,234,567元的应付款", 1234567.0, "人民币1,234,567元"),  # prefix AND suffix
        ("balance of CNY5.8 million at", 5800000.0, "CNY5.8 million"),
        ("CNY 835.3K as at", 835300.0, "CNY 835.3K"),          # the K suffix wave 1 added
        ("198m of bank loans", 198000000.0, "198m"),           # the bare m suffix
        ("1,234,567元", 1234567.0, "1,234,567元"),             # grouped + suffix
        ("US$4,500 paid", 4500.0, "US$4,500"),
        ("余额为7.8万", 78000.0, "7.8万"),                      # no suffix to extend over
        ("5，271.8万元", 52718000.0, "5，271.8万元"),           # fullwidth comma kept in place
    ]
    for text, value, expected_text in cases:
        found = extract_amount_spans(text)
        hit = [(v, s, e) for v, s, e in found if abs(v - value) < 1e-6]
        if not hit:
            check(f"span {text!r}", False, f"value {value} not parsed from {found}")
            continue
        _v, s, e = hit[0]
        check(f"span {text!r}", text[s:e] == expected_text,
              f"got {text[s:e]!r}, want {expected_text!r}")

    # A trailing currency word that belongs to the NEXT figure is not swallowed.
    text = "1,234,000 CNY5,678,000 total"
    spans = extract_amount_spans(text)
    check("a following prefix is not eaten as this amount's suffix",
          any(text[s:e] == "1,234,000" for _v, s, e in spans),
          f"{[(text[s:e]) for _v, s, e in spans]}")


# ---------------------------------------------------------------------------
# clause spans in final_content coordinates
# ---------------------------------------------------------------------------

def test_clause_spans() -> None:
    print("\nverify_commentary — every span slices back to its own clause")
    df = frame({
        "Cash": ["Cash on hand", "Deposits with banks", "Total"],
        INTERNAL_ROW_KEY: [4, 5, 6],
        "2025-12-31": [1_000_000.0, 4_000_000.0, 5_000_000.0],
    }, row_types={"Cash on hand": "detail", "Deposits with banks": "detail", "Total": "total"})
    content = ("现金余额为人民币500.0万元，其中库存现金人民币100.0万元。"
               "该余额较上期增加，主要由于经营活动现金流改善。")
    reviews = verify_commentary(content, df)
    spanned = [r for r in reviews if r.get("span")]
    check("every clause review carries a span", len(spanned) == len(segment_clauses(content)),
          f"{len(spanned)} spans for {len(segment_clauses(content))} clauses")
    for review in spanned:
        s, e = review["span"]
        if content[s:e] != review["clause"]:
            check("final_content[start:end] == clause", False,
                  f"{content[s:e]!r} != {review['clause']!r}")
            return
    check("final_content[start:end] == clause", True)

    # amount spans are rebased onto the content, not the clause
    for review in reviews:
        for amount in review.get("amounts") or []:
            s, e = amount["span"]
            if not content[s:e].strip():
                check("amount spans are content-relative", False, f"empty slice at {amount['span']}")
                return
    check("amount spans are content-relative", True)


# ---------------------------------------------------------------------------
# facts, provenance, classify_miss
# ---------------------------------------------------------------------------

def test_facts_and_classify_miss() -> None:
    print("\nSourceIndex — facts carry provenance; classify_miss keeps its own rule")
    df = frame({
        "Cash": ["Cash on hand", "Deposits with banks", "Total"],
        INTERNAL_ROW_KEY: [4, 5, 6],
        "2025-12-31": [1_000_000.0, 4_000_000.0, 5_000_000.0],
    }, row_types={"Cash on hand": "detail", "Deposits with banks": "detail", "Total": "total"})
    source = SourceIndex.from_df(df)

    fact = source.matches(4_000_000.0)
    check("matches returns the fact, not a bool", isinstance(fact, dict), repr(fact))
    check("the fact names its sheet, row and column",
          fact and fact["sheet"] == "Cash" and fact["row_idx"] == 5
          and fact["col_label"] == "2025-12-31" and fact["row_desc"] == "Deposits with banks",
          repr(fact))

    verdict = source.classify_miss(4_000.0)  # the real figure is 1000x this
    check("a x1000 miss classifies as a scale error",
          verdict["code"] == "AMOUNT_SCALE_ERROR", repr(verdict))
    check("the scale error names its expected value and source",
          verdict.get("expected") == 4_000_000.0 and verdict["source_ref"]["row_idx"] == 5,
          repr(verdict))
    check("a unique factor is not flagged ambiguous", verdict.get("ambiguous") is False)

    stray = source.classify_miss(7_777_777.0)
    check("an unrelated figure is AMOUNT_UNSUPPORTED", stray["code"] == "AMOUNT_UNSUPPORTED",
          repr(stray))

    # The CLI path: a bare float list still works and still classifies.
    plain = SourceIndex([1_000_000.0, 4_000_000.0])
    check("a bare List[float] still builds an index", plain.matches(1_000_000.0) is not None)
    plain_verdict = plain.classify_miss(4_000.0)
    check("bare floats still classify, with a null source_ref",
          plain_verdict["code"] == "AMOUNT_SCALE_ERROR"
          and plain_verdict["source_ref"]["sheet"] is None, repr(plain_verdict))

    # Window sums: bounded to the main frame, no total row, one row type. Row 6
    # is the Total, so no window may reach it. (Asserting on the row_range, not
    # on the value: the two detail rows here sum to the same 5,000,000 the Total
    # holds, which is what a real sheet looks like.)
    windows = [f for f in source.facts if f["kind"] == "window_sum"]
    check("a same-type run of detail rows is still pooled", len(windows) >= 1, repr(windows))
    check("no window sum reaches the Total row",
          all(f["row_range"][1] < 6 for f in windows), repr(windows))


def test_serialisation() -> None:
    print("\nserialisation — the new shape survives yaml.dump and json.dumps")
    df = frame({
        "Cash": ["Cash on hand", "Total"],
        INTERNAL_ROW_KEY: [4, 5],
        "2025-12-31": [1_000_000.0, 1_000_000.0],
    })
    content = "货币资金余额为人民币100.0万元，较上期无重大变动。截至2222年01月01日仍未使用。"
    results = {"Cash": {"agent_4_validation": {
        "final_content": content,
        "clause_reviews": verify_commentary(content, df),
    }}}
    try:
        yaml.dump(results, allow_unicode=True)
        ok_yaml = True
    except Exception as exc:  # noqa: BLE001 — the point is that nothing raises
        ok_yaml = False
        print("       yaml:", exc)
    try:
        json.dumps(results, ensure_ascii=False)
        ok_json = True
    except Exception as exc:  # noqa: BLE001
        ok_json = False
        print("       json:", exc)
    check("yaml.dump does not raise", ok_yaml)
    check("json.dumps does not raise", ok_json)

    codes = {r.get("code") for r in results["Cash"]["agent_4_validation"]["clause_reviews"] if r.get("code")}
    check("every emitted code is in DEFECT_CODES", codes <= DEFECT_CODES, repr(codes))
    check("the invented date is caught", "DATE_UNSUPPORTED" in codes, repr(codes))
    check("CLAIM_MISSING never enters clause_reviews", "CLAIM_MISSING" not in codes)


# ---------------------------------------------------------------------------
# M3 direction check
# ---------------------------------------------------------------------------

def movement(**kw):
    base = {"description": "租金收入", "from_period": "2023-12-31", "to_period": "2024-12-31",
            "from_value": 100_000.0, "to_value": 130_000.0, "percent_change": 30.0}
    base.update(kw)
    return base


def test_direction() -> None:
    print("\n_direction_reviews — fires only on a paired, unambiguous contradiction")
    rows = {"Revenue": ["租金收入", "Total"], INTERNAL_ROW_KEY: [4, 5],
            "2024-12-31": [130_000.0, 130_000.0]}

    df = frame(rows, movements=[movement()])
    check("a correct pairing is silent",
          not collect_direction_findings("租金收入较上期增加30.0%，主要由于出租率提升。", df))

    found = collect_direction_findings("租金收入较上期下降30.0%，主要由于出租率下滑。", df)
    check("a flipped direction word is caught", len(found) == 1, repr(found))
    if found:
        check("the finding is typed and spanned",
              found[0]["code"] == "DIRECTION_MISMATCH" and len(found[0]["span"]) == 2,
              repr(found[0]))

    check("a percentage that matches no movement is ignored",
          not collect_direction_findings("租金收入较上期下降12.5%。", df))
    check("a movement whose description is absent from the clause is ignored",
          not collect_direction_findings("本科目余额较上期下降30.0%。", df))
    check("a percentage with no direction word is ignored",
          not collect_direction_findings("租金收入的变动幅度为30.0%。", df))
    check("a clause with both direction words is ignored",
          not collect_direction_findings("租金收入下降30.0%后又增加。", df))

    nil = frame(rows, movements=[movement(from_value=0.0, percent_change=None)])
    check("a from-nil movement is skipped (no percentage to quote)",
          not collect_direction_findings("租金收入较上期下降30.0%。", nil))

    negative = frame(rows, movements=[movement(from_value=-100_000.0, to_value=-130_000.0,
                                               percent_change=30.0)])
    check("a negative-base movement is skipped",
          not collect_direction_findings("租金收入较上期下降30.0%。", negative))

    flipped = frame(rows, movements=[movement(from_value=-100_000.0, to_value=130_000.0,
                                              percent_change=-230.0)])
    check("a sign-flip movement is skipped",
          not collect_direction_findings("租金收入较上期增加230.0%。", flipped))

    ambiguous = frame(rows, movements=[movement(), movement(description="租金收入",
                                                           percent_change=-30.0,
                                                           from_value=130_000.0,
                                                           to_value=91_000.0)])
    check("two movements matching one magnitude with opposite signs are skipped",
          not collect_direction_findings("租金收入较上期下降30.0%。", ambiguous))

    check("a frame with no movements is silent",
          not collect_direction_findings("租金收入较上期下降30.0%。", frame(rows)))

    # The guard: a detector that raises must not take the pass down.
    class Exploding:
        @property
        def attrs(self):
            raise RuntimeError("boom")

    check("collect_direction_findings swallows a detector crash",
          collect_direction_findings("anything", Exploding()) == [])


def main() -> None:
    test_amount_spans()
    test_clause_spans()
    test_facts_and_classify_miss()
    test_serialisation()
    test_direction()
    print()
    if FAILURES:
        print(f"❌ {len(FAILURES)} failure(s): {FAILURES}")
        sys.exit(1)
    print("✅ all checks passed")


if __name__ == "__main__":
    main()

"""Zero-token regression and real-workbook probe for insight coverage.

    python ad-hoc/databook-probes/probe_insight_coverage.py
    python ad-hoc/databook-probes/probe_insight_coverage.py path/to/databook.xlsx
"""
from __future__ import annotations

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from fdd_utils.ui import build_insight_summary  # noqa: E402
from fdd_utils.workbook import process_workbook_data  # noqa: E402


def _sheet(
    status: str,
    nonzero: int,
    reached: int,
    *,
    reason: str = "",
    kind: str = "",
    hidden: bool = False,
):
    return {
        "status": status,
        "reason": reason,
        "kind": kind,
        "hidden": hidden,
        "title": "",
        "accounts": [],
        "blocks": [],
        "coverage": {
            "numeric_nonzero": nonzero,
            "reached_nonzero": reached,
            "text": 0,
            "text_reached": 0,
            "reached_by": {},
        },
    }


def _inspect_workbook(path: str) -> int:
    state = process_workbook_data(
        temp_path=path,
        entity_name="Coverage probe",
        selected_sheet=None,
    )
    insight = build_insight_summary(
        ai_results={},
        dfs=state.get("dfs"),
        reconciliation=state.get("reconciliation"),
        resolution=state.get("resolution"),
    )
    findings = [
        item for item in insight.get("visible_issues") or []
        if str(item.get("basis") or "").startswith("workbook_digest.coverage")
    ]
    print(f"accounts={len(state.get('dfs') or {})} coverage_findings={len(findings)}")
    for item in findings:
        print(f"{item['basis']}: {item['issue']}")
        print(f"  evidence={item.get('evidence_ids') or []}")
    print("coverage_questions:")
    for question in insight.get("client_questions") or []:
        if (
            question.startswith("What are the populated tabs ")
            or question.startswith("Mapped schedules exposed ")
        ):
            print(f"- {question}")
    return 0


def _run_regression() -> int:
    digest = {
        "reach_marked": True,
        "sheets": {
            "Mapped schedule": _sheet("mapped", 10, 4),
            "Unused support": _sheet("unmapped", 20, 0, reason="no mapping resolved"),
            "Financials": _sheet("financials", 30, 0),
            "Empty notes": _sheet("unmapped", 0, 0),
            "TB": _sheet("unmapped", 1000, 0, reason="no mapping resolved"),
            "2025": _sheet(
                "unmapped", 100, 0,
                reason="no_alias_scored_above_45",
                kind="financial_schedule",
            ),
            "_TM_hidden support": _sheet(
                "unmapped", 100, 0,
                reason="no mapping resolved",
                hidden=True,
            ),
            "Duplicate support": _sheet(
                "unmapped", 100, 0,
                reason="sheet_taken_by_AR",
            ),
        },
    }
    insight = build_insight_summary(
        ai_results={},
        resolution={
            "workbook_digest": digest,
            "unresolved_sheets": ["Unused support", "Empty notes"],
        },
    )
    by_basis = {item.get("basis"): item for item in insight.get("visible_issues") or []}
    tabs = by_basis.get("workbook_digest.coverage.unanalysed_tabs") or {}
    cells = by_basis.get("workbook_digest.coverage.mapped_cells") or {}
    questions = "\n".join(insight.get("client_questions") or [])
    archived = build_insight_summary(
        ai_results={},
        resolution={"unresolved_sheets": ["Legacy support"]},
    )
    archived_bases = {
        item.get("basis") for item in archived.get("visible_issues") or []
    }
    tests = [
        ("populated unmapped tab named", "Unused support" in str(tabs.get("issue") or "")),
        ("empty tab excluded", "Empty notes" not in str(tabs.get("issue") or "")),
        ("Financials excluded", "Financials" not in str(tabs.get("issue") or "")),
        ("technical/source tabs excluded", all(
            name not in str(tabs.get("issue") or "")
            for name in ("TB", "2025", "_TM_hidden support", "Duplicate support")
        )),
        ("mapped cell ratio", "4 of 10" in str(cells.get("issue") or "")),
        ("stable tab evidence", "coverage:Unused support" in (tabs.get("evidence_ids") or [])),
        ("unanalysed tab question", "What are the populated tabs Unused support used for" in questions),
        ("mapped cell question", "Do the unreached cells in Mapped schedule" in questions),
        ("coarse finding suppressed", "resolution.unresolved_sheets" not in by_basis),
        ("archived fallback preserved", "resolution.unresolved_sheets" in archived_bases),
    ]
    for name, passed in tests:
        print(f"{'PASS' if passed else 'FAIL'}  {name}")
    ok = all(passed for _name, passed in tests)
    print("\n" + ("ALL CASES PASS" if ok else "FAILURES ABOVE"))
    return 0 if ok else 1


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("workbook", nargs="?")
    args = parser.parse_args()
    return _inspect_workbook(args.workbook) if args.workbook else _run_regression()


if __name__ == "__main__":
    raise SystemExit(main())

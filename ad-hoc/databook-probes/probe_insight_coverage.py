"""Zero-token regression probe for workbook coverage in the insight summary.

    python ad-hoc/databook-probes/probe_insight_coverage.py
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from fdd_utils.ui import build_insight_summary  # noqa: E402


def _sheet(status: str, nonzero: int, reached: int, *, reason: str = ""):
    return {
        "status": status,
        "reason": reason,
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


def main() -> int:
    digest = {
        "reach_marked": True,
        "sheets": {
            "Mapped schedule": _sheet("mapped", 10, 4),
            "Unused support": _sheet("unmapped", 20, 0, reason="no mapping resolved"),
            "Financials": _sheet("financials", 30, 0),
            "Empty notes": _sheet("unmapped", 0, 0),
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


if __name__ == "__main__":
    raise SystemExit(main())

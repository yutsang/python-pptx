"""Zero-token regression probe for executive-summary amount grounding.

The summary is allowed to quote a figure from any account in its statement,
but a sentence containing a figure in none of those pools must not ship.

    python ad-hoc/databook-probes/probe_summary_grounding.py
"""
from __future__ import annotations

import os
import sys
import tempfile
from unittest.mock import patch

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from fdd_utils.ai.evidence import AccountEvidence, write_evidence  # noqa: E402
from fdd_utils.pptx import PowerPointGenerator  # noqa: E402
from fdd_utils.ui import build_section_summaries  # noqa: E402


def _evidence(key: str, *values: float) -> AccountEvidence:
    return AccountEvidence(
        mapping_key=key,
        facts=[
            {
                "value": value,
                "kind": "cell",
                "sheet": key,
                "row_desc": key,
                "col_label": "latest",
                "multiplier": 1.0,
            }
            for value in values
        ],
    )


def main() -> int:
    pools = [_evidence("Cash", 120_000_000), _evidence("Revenue", 50_000_000)]
    cases = [
        (
            "cross-account union",
            "货币资金为1.2亿元，应收余额为0.5亿元。收入整体保持稳定。",
            "货币资金为1.2亿元，应收余额为0.5亿元。收入整体保持稳定。",
            0,
        ),
        (
            "unsupported numeric sentence",
            "货币资金为9.9亿元。收入整体保持稳定。",
            "收入整体保持稳定。",
            1,
        ),
        (
            "non-numeric summary",
            "流动性保持稳定，经营表现仍受成本压力影响。",
            "流动性保持稳定，经营表现仍受成本压力影响。",
            0,
        ),
        (
            "all numeric sentences rejected",
            "货币资金为9.9亿元。",
            "",
            1,
        ),
    ]

    ok = True
    print(f"{'case':30s} {'result':8s} details")
    print("-" * 80)
    for name, draft, expected, dropped_expected in cases:
        grounded, report = PowerPointGenerator.ground_section_summary(
            draft, pools, is_chinese=True)
        dropped = len(report.get("dropped_sentences") or [])
        passed = grounded == expected and dropped == dropped_expected
        ok = ok and passed
        print(
            f"{name:30s} {'PASS' if passed else 'FAIL':8s} "
            f"checked={report.get('checked_amounts')} dropped={dropped} "
            f"output={grounded!r}"
        )

    with tempfile.TemporaryDirectory() as run_folder:
        write_evidence(run_folder, pools[0])
        run_results = {
            "Cash": {"final": "货币资金为1.2亿元，流动性保持稳定。"},
            "__run__": {"run_folder": run_folder},
        }
        grounding = {}
        with patch.object(
            PowerPointGenerator,
            "generate_section_summary",
            return_value="货币资金为9.9亿元。流动性保持稳定。",
        ):
            summaries = build_section_summaries(
                ai_results=run_results,
                mappings={"Cash": {"type": "BS", "aliases": []}},
                is_chinese_db=True,
                grounding_out=grounding,
            )
        integration_ok = (
            summaries.get("BS") == "流动性保持稳定。"
            and len((grounding.get("BS") or {}).get("dropped_sentences") or []) == 1
        )
        ok = ok and integration_ok
        print(
            f"{'production build seam':30s} {'PASS' if integration_ok else 'FAIL':8s} "
            f"output={summaries.get('BS')!r}"
        )

        with patch.object(
            PowerPointGenerator,
            "generate_section_summary",
            return_value="货币资金为9.9亿元。",
        ):
            rejected = build_section_summaries(
                ai_results=run_results,
                mappings={"Cash": {"type": "BS", "aliases": []}},
                is_chinese_db=True,
            )
        notice_ok = "未通过来源核对" in rejected.get("BS", "")
        ok = ok and notice_ok
        print(
            f"{'all-unsupported fail closed':30s} {'PASS' if notice_ok else 'FAIL':8s} "
            f"output={rejected.get('BS')!r}"
        )

    print("\n" + ("ALL CASES PASS" if ok else "FAILURES ABOVE"))
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())

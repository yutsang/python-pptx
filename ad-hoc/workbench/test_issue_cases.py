"""The place refused sentences are kept (fdd_utils/agent/issues.py): written from a session, replayed with no workbook.

    python ad-hoc/workbench/test_issue_cases.py
"""
from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from test_claims import P1, P2, build, ref  # noqa: E402

from fdd_utils.agent import issues  # noqa: E402
from fdd_utils.ai import claims as cl  # noqa: E402


class Fixture:
    """An account with one right and one wrong sentence, as a session would hold it after a run."""

    def __init__(self):
        rows = [("预提成本", "detail", {P1: 0, P2: 48600}), ("仓库建造成本", "detail", {P1: 600000, P2: 600000}),
                ("Total", "total", {P1: 600000, P2: 648600})]
        self.contract = build(rows=rows, notes=[], account="在建工程")
        cl.standard_derived(self.contract)
        total = ref(self.contract, P2)
        self.claims = [
            {"account": "在建工程", "kind": "amount", "text": "2025年末余额为64.86万元。", "periods": [P2], "cites": [total]},
            {"account": "在建工程", "kind": "amount", "text": "2025年末余额为70.00万元。", "periods": [P2], "cites": [total]}]
        for c in self.claims:
            c["source_version"] = self.contract.source_version
        outcome = cl.verify_candidate(self.claims, {"在建工程": self.contract}, default_account="在建工程")
        candidate = SimpleNamespace(claims=self.claims, outcome=outcome, accepted=False, validated=False)
        self.source = SimpleNamespace(candidates={"在建工程": candidate}, contracts={"在建工程": self.contract})


class KeptAndReplayed(unittest.TestCase):
    def test_only_the_refused_sentence_is_kept_and_it_stays_open_while_it_is_wrong(self):
        fx = Fixture()
        with tempfile.TemporaryDirectory() as tmp:
            written = issues.capture(fx.source, Path(tmp), entity="Entity A", code="abc123")
            self.assertEqual(len(written), 1)
            result = issues.replay(written[0])
        self.assertEqual([r["claim"] for r in result["was"]], [1])
        self.assertEqual(result["status"], "open")
        self.assertEqual([r["claim"] for r in result["now"]], [1])

    def test_a_case_whose_refusal_the_code_no_longer_makes_is_fixed(self):
        # the stored refusal is edited into a right sentence: the verifier as it is now passes it, so the case reads as fixed
        import json
        fx = Fixture()
        with tempfile.TemporaryDirectory() as tmp:
            path = issues.capture(fx.source, Path(tmp), entity="Entity A")[0]
            case = json.loads(path.read_text(encoding="utf-8"))
            case["claims"][1]["text"] = "2025年末余额为64.86万元。"
            path.write_text(json.dumps(case, ensure_ascii=False), encoding="utf-8")
            result = issues.replay(path)
            self.assertEqual(result["status"], "fixed")
            self.assertEqual(issues.summary([result])["by_status"], {"fixed": 1})

    def test_an_account_with_nothing_refused_leaves_no_case(self):
        fx = Fixture()
        fx.source.candidates["在建工程"].outcome = cl.verify_candidate(fx.claims[:1], fx.source.contracts, default_account="在建工程")
        fx.source.candidates["在建工程"].claims = fx.claims[:1]
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(issues.capture(fx.source, Path(tmp), entity="Entity A"), [])

    def test_the_summary_names_the_codes_the_open_cases_are_refused_with(self):
        fx = Fixture()
        with tempfile.TemporaryDirectory() as tmp:
            result = issues.replay(issues.capture(fx.source, Path(tmp), entity="Entity A")[0])
        s = issues.summary([result])
        self.assertEqual(s["cases"], 1)
        self.assertTrue(s["open_by_code"])


if __name__ == "__main__":
    unittest.main()

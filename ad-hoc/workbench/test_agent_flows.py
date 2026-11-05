"""The five flows on a real workbook, end to end, with no provider.

    python ad-hoc/workbench/test_agent_flows.py ["databook.xlsx"]      # default: $FDD_FLOW_DATABOOK, else databook.xlsx

1  databook -> evidence checked period by period -> numerical summary          (no model)
2  a question -> the agent investigates -> an answer with sources and states   (scripted policy)
3  edit one account's paragraph -> candidate -> verification -> deck export     (scripted policy)
4  the sources change -> semantic diff -> affected content                      (a changed copy of the workbook)
5  the exported deck comes back edited -> edits kept, conflicts found, relinks  (the deck from 3)

The "model" is a scripted policy, so this proves the plumbing and the gates on real frames, a real
template and the real packing. It proves nothing about how a real model writes or chooses.
Skipped (exit 0, says so) when the workbook is not on this machine. Leaves a stub run folder under
fdd_utils/logs, as any stub run does.
"""
from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import unittest
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(ROOT))

WORKBOOK = next((a for a in sys.argv[1:] if a.endswith(".xlsx")), os.environ.get("FDD_FLOW_DATABOOK", "databook.xlsx"))
HAVE_WORKBOOK = (ROOT / WORKBOOK).exists()

from fdd_utils.agent import controller as C  # noqa: E402
from fdd_utils.agent import tools as T  # noqa: E402
from fdd_utils.agent import updates as U  # noqa: E402
from fdd_utils.agent.session import SourceSession  # noqa: E402
from fdd_utils.ai import claims as cl  # noqa: E402
from fdd_utils.pptx import lineage as L  # noqa: E402
from test_agent_controller import Policy, evidence_of, returned, tool  # noqa: E402

_STATE = {}


def setUpModule():
    if not HAVE_WORKBOOK:
        return
    os.chdir(ROOT)
    before = set(p.name for p in (ROOT / "fdd_utils" / "logs").glob("run_*"))
    subprocess.run([sys.executable, "-X", "utf8", str(HERE / "stub_inspect_run.py"), WORKBOOK, "--limit", "6"],
                   capture_output=True, text=True, timeout=600, env=dict(os.environ, PYTHONIOENCODING="utf-8"))
    new = sorted(set(p.name for p in (ROOT / "fdd_utils" / "logs").glob("run_*")) - before)
    _STATE["base_run"] = new[-1] if new else None
    _STATE["tmp"] = tempfile.mkdtemp(prefix="agent_flows_")
    _STATE["session"] = SourceSession.load(WORKBOOK, base_run=_STATE["base_run"])


def tearDownModule():
    if _STATE.get("tmp"):
        shutil.rmtree(_STATE["tmp"], ignore_errors=True)


def edit_account(session):
    """An account the base run commented on, with a total and two periods."""
    for key in session.base_results or {}:
        contract = session.contracts.get(key)
        if contract and len(contract.periods) >= 2 and contract.total(contract.periods[-1]):
            return key
    raise unittest.SkipTest("the base run commented on no usable account")


def date_text(period):
    y, m, d = (int(x) for x in period.split("-"))
    return f"{y}年{m}月{d}日"


@unittest.skipUnless(HAVE_WORKBOOK, f"{WORKBOOK} is not on this machine")
class Flow1_Summary(unittest.TestCase):
    def test_every_number_in_the_summary_traces_to_a_fact(self):
        session = _STATE["session"]
        result = T.call_tool(session, "numerical_summary")
        self.assertTrue(result.ok, result.error)
        trace = result.data["traceability"]
        self.assertTrue(trace["complete"], trace)
        self.assertGreater(trace["numbers_printed"], 20)
        self.assertEqual(len(result.data["rows"]), len(session.contracts))

    def test_contracts_exist_for_every_account_and_checks_have_a_state(self):
        session = _STATE["session"]
        info = T.call_tool(session, "inspect_sources").data
        self.assertEqual(info["statement_source"], "financials")
        self.assertTrue(all(a["checks"] for a in info["accounts"] if a["periods"]))


@unittest.skipUnless(HAVE_WORKBOOK, f"{WORKBOOK} is not on this machine")
class Flow2_Ask(unittest.TestCase):
    def test_a_question_is_answered_with_sources_and_verification_state(self):
        session = _STATE["session"]
        key = edit_account(session)
        contract = session.contracts[key]
        last = contract.periods[-1]

        def policy(user, results):
            # a real account's evidence is long and the conversation truncates it, so the scripted
            # policy asks the session for what the model would have read
            if not any(n == "get_account_evidence" for n, _ in results) and "get_account_evidence(" not in user:
                return tool("get_account_evidence", account=key)
            view = cl.evidence_view(contract)
            total = next(t for t in view["totals"] if t["period"] == last and t["basis"] == "authority_adjusted")
            return {"answer": {"claims": [{"account": key, "kind": "amount", "periods": [last], "cites": [total["ref"]],
                                           "text": f"截至{date_text(last)}，{key}余额{abs(total['yuan']) / 1e4:,.1f}万元"}], "unresolved": []}}
        run = C.AgentController(session, Policy(policy)).run(f"{key}的期末余额是多少？")
        self.assertEqual(run.stop_reason, C.STOP_ANSWERED, run.steps)
        self.assertEqual(run.answer["verification"]["failed"], 0)
        source = run.answer["sources"][0]
        self.assertTrue(source["sheet"] and source["row"] is not None, source)
        self.assertEqual(run.answer["source_version"], session.source_version)


@unittest.skipUnless(HAVE_WORKBOOK, f"{WORKBOOK} is not on this machine")
class Flow3_Edit(unittest.TestCase):
    def test_edit_validate_export_and_the_touched_slides_are_found(self):
        session = SourceSession.load(WORKBOOK, base_run=_STATE["base_run"])
        key = edit_account(session)
        contract = session.contracts[key]
        last = contract.periods[-1]
        out_dir = _STATE["tmp"]

        def policy(user, results):
            if "get_account_evidence(" not in user:
                return tool("get_account_evidence", account=key)
            view = cl.evidence_view(contract)
            total = next(t for t in view["totals"] if t["period"] == last and t["basis"] == "authority_adjusted")
            if "validate_candidate(" not in user:
                return tool("validate_candidate", account=key, claims=[{"kind": "amount", "periods": [last], "cites": [total["ref"]],
                            "text": f"截至{date_text(last)}，{key}余额{abs(total['yuan']) / 1e4:,.1f}万元"}])
            if "export_candidate(" not in user:
                return tool("export_candidate", out_path=os.path.join(out_dir, "edit.pptx"))
            return {"done": {"account": key, "export": True, "summary": "ok"}}
        controller = C.AgentController(session, Policy(policy))
        run = controller.run(f"改写{key}的评论", mode="edit", account=key, export=True)
        export = session.exports[-1]
        _STATE["export"], _STATE["key"], _STATE["edit_session"] = export, key, session
        # the stub run's own Prepayments clause is an unsupported figure: the deck inherits it and is NOT deliverable
        self.assertTrue(os.path.exists(export["path"]), export)
        self.assertTrue(export["touched_slides"][key], "the edited text was found on a slide")
        self.assertEqual(export["manifest"]["elements"] >= 1, True)
        self.assertTrue(os.path.exists(export["path"] + L.INTERNAL_SUFFIX))
        self.assertTrue(os.path.exists(export["path"] + ".delivery.json"))
        self.assertEqual(export["delivery"]["status"], "blocked")
        self.assertIn("_DRAFT-BLOCKED", export["path"])
        step = [s for s in run.steps if s.get("tool") == "export_candidate"][0]
        self.assertEqual((step["ok"], step["error"]), (False, "DELIVERY_BLOCKED"))      # a draft was made; the tool did not call it a success

    def test_a_candidate_that_fails_verification_cannot_be_exported(self):
        session = SourceSession.load(WORKBOOK, base_run=_STATE["base_run"])
        key = edit_account(session)
        last = session.contracts[key].periods[-1]
        ref = next(f["ref"] for f in session.contracts[key].facts if f["period"] == last and f["role"] == "total" and f["basis"] == "authority_adjusted")
        bad = T.call_tool(session, "validate_candidate", {"account": key, "claims": [{"kind": "amount", "periods": [last], "cites": [ref],
                          "text": f"截至{date_text(last)}，{key}余额9,999.9万元"}]})
        self.assertFalse(bad.ok)
        self.assertEqual(T.call_tool(session, "export_candidate", {"accounts": [key]}).error["code"], "NOT_VALIDATED")


@unittest.skipUnless(HAVE_WORKBOOK, f"{WORKBOOK} is not on this machine")
class Flow4_5_UpdateAndReturn(unittest.TestCase):
    def exported(self):
        if "export" not in _STATE:
            Flow3_Edit("test_edit_validate_export_and_the_touched_slides_are_found").test_edit_validate_export_and_the_touched_slides_are_found()
        return _STATE["export"], _STATE["key"], _STATE["edit_session"]

    def test_an_untouched_returned_deck_reads_as_unchanged(self):
        export, key, session = self.exported()
        report = U.analyze_returned(export["path"], session)
        self.assertEqual(report["mode"], "full")
        self.assertFalse(report["source_moved"])
        self.assertTrue(all(a["action"] == "keep" for a in report["actions"]), report["summary"])

    def test_a_user_edit_in_the_returned_deck_is_found_and_kept(self):
        export, key, session = self.exported()
        edited = os.path.join(_STATE["tmp"], "returned.pptx")
        shutil.copyfile(export["path"], edited)
        shutil.copyfile(export["path"] + L.INTERNAL_SUFFIX, edited + L.INTERNAL_SUFFIX)
        from pptx import Presentation
        prs = Presentation(edited)
        changed = False
        for slide in prs.slides:
            for shape in slide.shapes:
                if not getattr(shape, "has_text_frame", False):
                    continue
                for para in shape.text_frame.paragraphs:
                    if para.text.startswith("■") and "万元" in para.text and para.runs:
                        para.runs[-1].text = para.runs[-1].text + " 管理层另行说明。"
                        changed = True
                        break
                if changed:
                    break
            if changed:
                break
        prs.save(edited)
        report = U.analyze_returned(edited, session)
        self.assertEqual(report["summary"].get("reverify"), 1, report["summary"])
        self.assertEqual(report["summary"].get("conflict", 0), 0)

    def test_a_changed_workbook_gives_a_semantic_diff_and_the_affected_paragraph(self):
        export, key, session = self.exported()
        contract = session.contracts[key]
        last = contract.periods[-1]
        # the sources move: the cited total changes (a copy of the session with that fact shifted)
        shifted = SourceSession.load(WORKBOOK, base_run=_STATE["base_run"])
        for fact in shifted.contracts[key].facts:
            if fact["role"] == "total" and fact["period"] == last:
                fact["value"] += 1_000_000.0
        shifted.binding.workbook_sha256 = "moved"
        for c in shifted.contracts.values():
            c.binding = shifted.binding.to_dict()
        diff = U.diff_sources(L.read_internal(export["path"])["baseline"], shifted.contracts)
        self.assertEqual(diff["changed_accounts"], [key])
        self.assertTrue(any(c["new"] - c["old"] == 1_000_000.0 for c in diff["accounts"][key]["changed"]))
        report = U.analyze_returned(export["path"], shifted)
        acted = {a["element"]: a["action"] for a in report["actions"]}
        mine = [e["id"] for e in L.read_manifest(export["path"])["elements"] if e["account"] == key]
        self.assertTrue(all(acted[i] == "auto_update" for i in mine), acted)
        self.assertTrue(all(a == "keep" for i, a in acted.items() if i not in mine), "only the paragraph that cites the fact is affected")
        self.assertTrue(report["source_moved"])


@unittest.skipUnless(HAVE_WORKBOOK, f"{WORKBOOK} is not on this machine")
class Cli(unittest.TestCase):
    """fdd_agent.py: free commands always run; model commands need the flag; a scripted file stands in for the model."""

    def run_cli(self, *argv, flag=None):
        import contextlib
        import io
        import fdd_agent
        saved = os.environ.pop("FDD_AGENT", None)
        if flag is not None:
            os.environ["FDD_AGENT"] = flag
        out = io.StringIO()
        try:
            with contextlib.redirect_stdout(out):
                code = fdd_agent.main(list(argv))
        finally:
            os.environ.pop("FDD_AGENT", None)
            if saved is not None:
                os.environ["FDD_AGENT"] = saved
        return code, out.getvalue()

    def test_summary_needs_no_flag_and_no_model(self):
        code, text = self.run_cli("summary", WORKBOOK)
        self.assertEqual(code, 0, text[-400:])
        self.assertIn("no model call", text)
        self.assertRegex(text, r"traceability: (\d+) of \1 printed numbers resolve")

    def test_ask_is_refused_while_the_flag_is_off_and_says_how_to_run_once(self):
        code, text = self.run_cli("ask", WORKBOOK, "余额？", "--scripted", os.devnull)
        self.assertEqual(code, 2)
        self.assertIn("agent path is off", text)
        self.assertIn("FDD_AGENT=1", text)

    def test_a_scripted_conversation_runs_end_to_end_with_the_flag(self):
        session = _STATE["session"]
        key = edit_account(session)
        contract = session.contracts[key]
        last = contract.periods[-1]
        view = cl.evidence_view(contract)
        total = next(t for t in view["totals"] if t["period"] == last and t["basis"] == "authority_adjusted")
        replies = [tool("get_account_evidence", account=key),
                   {"answer": {"claims": [{"account": key, "kind": "amount", "periods": [last], "cites": [total["ref"]],
                                           "text": f"截至{date_text(last)}，{key}余额{abs(total['yuan']) / 1e4:,.1f}万元"}], "unresolved": []}}]
        path = os.path.join(_STATE["tmp"], "script.json")
        Path(path).write_text(json.dumps(replies, ensure_ascii=False), encoding="utf-8")
        code, text = self.run_cli("ask", WORKBOOK, f"{key}的期末余额？", "--scripted", path, flag="1")
        self.assertEqual(code, 0, text[-600:])
        self.assertIn("stop: answered", text)
        self.assertIn("verification: 1 verified", text)

    def test_a_conversation_that_runs_out_ends_with_a_report_and_exit_2(self):
        path = os.path.join(_STATE["tmp"], "short.json")
        Path(path).write_text(json.dumps([tool("inspect_sources")]), encoding="utf-8")
        code, text = self.run_cli("ask", WORKBOOK, "?", "--scripted", path, flag="1")
        self.assertEqual(code, 2)
        self.assertIn("why it stopped", text)


if __name__ == "__main__":
    if not HAVE_WORKBOOK:
        print(f"SKIPPED: {WORKBOOK} is not on this machine (these flows need a real workbook).")
        sys.exit(0)
    unittest.main(verbosity=2, argv=[a for a in sys.argv if not a.endswith(".xlsx")])

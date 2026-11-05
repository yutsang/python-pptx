"""The continuous workflow on a real workbook and its archived run, with no provider.

    FDD_FLOW_WORKBOOK="<databook.xlsx>" FDD_FLOW_BASE_RUN=<run_id> python ad-hoc/workbench/test_workspace_flows.py

investigate (the four reading tools) -> ask -> follow up -> "put that in the report" -> review the draft -> accept ->
export -> the workbook is saved with one figure changed -> what stands and what does not -> a returned deck with a
hand edit -> update. The "model" is a scripted policy, so this proves the plumbing, the gates and the files on real
frames, the real template and the real packing; it says nothing about how a real model writes or chooses.

The workbook is COPIED to a temp folder first and every change is made to the copy: the original is hashed before and
after and the test fails if it moved. Without the two environment variables the tests are skipped (exit 0, says so):
the names of real workbooks are not kept in the repository.
"""
from __future__ import annotations

import os
import re
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(ROOT))

WORKBOOK = os.environ.get("FDD_FLOW_WORKBOOK", "")
BASE_RUN = os.environ.get("FDD_FLOW_BASE_RUN", "")
HAVE = bool(WORKBOOK) and Path(WORKBOOK).exists() and bool(BASE_RUN)

from fdd_utils.agent import controller as C  # noqa: E402
from fdd_utils.agent import tools as T  # noqa: E402
from fdd_utils.agent import updates as U  # noqa: E402
from fdd_utils.agent import workspace as W  # noqa: E402
from fdd_utils.ai import claims as cl  # noqa: E402
from fdd_utils.pptx import lineage as L  # noqa: E402
from fdd_utils.run_identity import file_sha256  # noqa: E402
from test_agent_controller import Policy, evidence_of, tool  # noqa: E402
from flow_helpers import bump_cached_values, date_text, find_figure, sheet_part, wan  # noqa: E402,F401
from flow_helpers import pick_account as _pick_account  # noqa: E402

_STATE: dict = {}


def setUpModule():
    if not HAVE:
        return
    os.chdir(ROOT)
    _STATE["tmp"] = tempfile.mkdtemp(prefix="flows_")
    _STATE["original_sha"] = file_sha256(WORKBOOK)
    _STATE["copy"] = os.path.join(_STATE["tmp"], "copy.xlsx")
    shutil.copyfile(WORKBOOK, _STATE["copy"])


def tearDownModule():
    if HAVE:
        assert file_sha256(WORKBOOK) == _STATE["original_sha"], "the original workbook was modified"
        thread = _STATE.get("thread")
        if thread is not None:       # no names, no figures: which branch the real data took
            check = (thread.work.source.contracts[thread.key].check(thread.last, "total_vs_financials") or {}).get("status")
            print(f"\nflow summary: latest-period check {check}; cells changed in the copy {len(_STATE.get('bumped') or [])}; "
                  f"draft after a tab-only change {(_STATE.get('effect_tab_only') or {}).get('status')}, after a coordinated update {(_STATE.get('effect') or {}).get('status')}; "
                  f"returned deck {_STATE.get('update')}; "
                  f"base run check {thread.work.source.base_run_check['status']}; original workbook unchanged: yes", file=sys.stderr)
        shutil.rmtree(_STATE["tmp"], ignore_errors=True)


def pick_account(source):
    return _pick_account(source, _STATE["copy"])


@unittest.skipUnless(HAVE, "set FDD_FLOW_WORKBOOK and FDD_FLOW_BASE_RUN (a real workbook and the archived run written from it)")
class Thread(unittest.TestCase):
    """One session, from the first question to the update of a returned deck. Order matters: each test builds on the last."""

    @classmethod
    def setUpClass(cls):
        cls.work = W.WorkSession.open(_STATE["copy"], base_run=BASE_RUN, model=None, budget=C.Budget(max_steps=10))
        picked = pick_account(cls.work.source)
        if picked is None:
            raise unittest.SkipTest("no account in this workbook suits the flow (needs a base commentary, two periods, an unreplaced total)")
        cls.key, cls.last, cls.prev, cls.raw, cls.row, cls.unit = picked
        _STATE["thread"] = cls

    def policy_ask(self, evidence_first=True):
        key, last, prev = self.key, self.last, self.prev

        def policy(user, results):
            names = [n for n, _ in results]
            if "get_account_evidence" not in names and "get_account_evidence(" not in user:
                return tool("get_account_evidence", account=key)
            if "compute_analysis(" not in user:
                return tool("compute_analysis", account=key, op="movement", period_from=prev, period=last)
            if "read_source_block(" not in user:
                view = cl.evidence_view(self.work.source.contracts[key])
                raw_ref = next(t["ref"] for t in view["totals"] if t["period"] == last and t["basis"] == "raw_schedule")
                return tool("read_source_block", account=key, ref=raw_ref)
            view = cl.evidence_view(self.work.source.contracts[key])
            total = next(t for t in view["totals"] if t["period"] == last and t["basis"] == "authority_adjusted")
            move = next(d for d in self.work.source.contracts[key].derived if d["kind"] == "movement" and d["periods"] == [prev, last])
            word = "增加" if move["value"] > 0 else "减少"
            return {"answer": {"claims": [
                {"account": key, "kind": "amount", "periods": [last], "cites": [total["ref"]], "text": f"截至{date_text(last)}，{key}余额{wan(total['yuan'])}"},
                {"account": key, "kind": "movement", "periods": [prev, last], "cites": [move["ref"]],
                 "text": f"较{date_text(prev)}{word}{wan(move['value'])}，变动幅度{abs(move['pct']) * 100:.1f}%"}],
                "unresolved": [f"{key}：变动原因未能在工作簿中核实"]}}
        return Policy(policy)

    # -- the four reading tools on real material --------------------------------------------------------------
    def test_1_the_reading_tools_return_location_version_and_limits_on_real_data(self):
        s = self.work.source
        ev = T.call_tool(s, "get_account_evidence", {"account": self.key})
        ref = next(t["ref"] for t in ev.data["totals"] if t["basis"] == "raw_schedule" and t["period"] == self.last)
        block = T.call_tool(s, "read_source_block", {"account": self.key, "ref": ref})
        self.assertTrue(block.ok, block.error)
        self.assertEqual(block.data["highlight_row"], self.row)
        self.assertIn(self.raw["sheet"], block.data["location"])
        self.assertEqual(block.data["source"]["source_version"], s.source_version)
        self.assertIn("formulas are neither shown", block.data["limits"]["values"])
        # the highlighted row really holds the figure the fact says it does (units aside)
        cells = [c for r in block.data["rows"] if r["row"] == self.row for c in r["cells"].values() if isinstance(c, (int, float))]
        self.assertTrue(any(abs(c * self.unit - self.raw["value"]) < 1.0 for c in cells), "the block's row holds the fact's figure")
        self.assertTrue(block.data["figure_found_in"], "the cells that hold the figure are named")
        self.assertTrue(all(abs(m["value"] * {"yuan": 1, "thousands": 1e3, "ten-thousands": 1e4}[m["printed_in"]] - abs(self.raw["value"])) < 1.0
                            for m in block.data["figure_found_in"]))
        search = T.call_tool(s, "search_notes", {"query": self.key, "account": self.key})
        self.assertTrue(search.ok, search.error)
        self.assertIn("only visible sheets are read", search.data["limits"])
        self.assertEqual(search.data["source"]["source_version"], s.source_version)
        cross = T.call_tool(s, "cross_account_evidence", {"account": self.key})
        adj = T.call_tool(s, "adjustment_evidence", {"account": self.key})
        for result in (cross, adj):
            self.assertTrue(result.ok, result.error)
            self.assertTrue(result.data["limits"])
            self.assertEqual(result.data["source"]["source_version"], s.source_version)
        self.assertEqual({p["period"] for p in adj.data["periods"]}, set(s.contracts[self.key].periods))

    # -- ask, follow up, put it in the report ---------------------------------------------------------------------
    def test_2_ask_then_follow_up_reuses_the_evidence_without_fetching_it_again(self):
        self.work.model = self.policy_ask()
        first = self.work.ask(f"{self.key}的变动来自哪里？")
        self.assertTrue(first.ok, first.report)
        self.assertEqual(first.tool_path, ["get_account_evidence", "compute_analysis", "read_source_block"])
        d_ref = next(c["cites"][0] for c in first.claims if c["kind"] == "movement")
        self.assertTrue(d_ref.startswith("D"))

        def follow(user, results):
            self.assertIn("EARLIER IN THIS SESSION", user)
            return {"answer": {"claims": [{"account": self.key, "kind": "movement", "periods": [self.prev, self.last], "cites": [d_ref],
                                           "text": first.claims[1]["text"]}], "unresolved": [f"{self.key}：变动原因未能在工作簿中核实"]}}
        calls = self.work.model_calls
        self.work.model = Policy(follow)
        second = self.work.ask("这个变动的幅度是多少？")
        self.assertTrue(second.ok, second.report)
        self.assertEqual(second.tool_path, [])
        self.assertEqual(self.work.model_calls, calls + 1)

    def test_3_put_that_in_the_report_makes_an_accepted_draft_that_shows_what_it_replaces(self):
        before = self.work.model_calls
        turn = self.work.report("把这段放入报告")
        self.assertTrue(turn.ok, turn.question or turn.report)
        self.assertEqual(self.work.model_calls, before)
        view = next(v for v in self.work.candidates_view() if v["account"] == self.key)
        self.assertEqual(view["status"], "accepted")
        self.assertEqual(view["origin"], "promoted")
        self.assertTrue(view["replaces"], "the base run has commentary for this account")
        self.assertTrue(view["changes"]["replaces_existing"])
        self.assertTrue(any("变动原因未能在工作簿中核实" in u for u in view["unresolved"]))

    # -- export, keeping everything else ---------------------------------------------------------------------------------
    def test_4_export_keeps_the_other_accounts_and_finds_the_touched_slides(self):
        out = os.path.join(_STATE["tmp"], "report.pptx")
        result = self.work.export(out_path=out)
        data = result.get("data") or {}
        self.assertTrue(data.get("path") and os.path.exists(data["path"]), result.get("error"))
        _STATE["export"] = data
        _STATE["exported_text"] = cl.render_claims(self.work.source.candidates[self.key].claims)
        self.assertTrue(data["touched_slides"][self.key], "the edited paragraph is on a slide")
        self.assertEqual(data["manifest"]["elements"] >= 1, True)
        self.assertTrue(os.path.exists(data["path"] + L.INTERNAL_SUFFIX))
        codes = {f["code"]: f["severity"] for f in data["delivery"]["findings"]}
        status = self.work.source.base_run_check["status"]
        self.assertEqual("BASE_RUN_UNVERIFIED" in codes, status == "unverified", codes)
        self.assertIn("SUMMARY_NOT_REGENERATED", codes)
        # the other accounts' commentary is still in the deck
        from fdd_utils.agent.export import slides_with_text, text_fragments
        kept = 0
        for other, result in (self.work.source.base_results or {}).items():
            if other == self.key or str(other).startswith("__"):
                continue
            text = self.work.base_text(other)
            if len(text) > 40 and slides_with_text(data["path"], text_fragments(text)[:3]):
                kept += 1
        self.assertGreater(kept, 3, "the accounts nobody edited keep their commentary")

    # -- the workbook is saved with one figure changed ------------------------------------------------------------------------
    def refresh_after(self, changes):
        """Apply `changes` [(sheet, row, shown value, delta)] to the copy, check the session notices, read it again."""
        s = self.work.source
        old_version = s.source_version
        touched = []
        for sheet, row, shown, delta in changes:
            touched += bump_cached_values(_STATE["copy"], sheet, row, shown, delta)
        self.assertTrue(touched, "the controlled change touched at least one cell")
        self.assertTrue(s.stale_reason(), "the session notices the file changed")
        self.assertEqual(T.call_tool(s, "get_account_evidence", {"account": self.key}).error["code"], "STALE_SOURCE")   # not answered from the old reading
        result = self.work.refresh()
        self.assertTrue(result["ok"], result)
        self.assertNotEqual(s.source_version, old_version)
        return touched, {e["account"]: e for e in result["candidate_effects"]}[self.key]

    def test_5_a_tab_changes_but_the_deck_follows_the_statement_so_the_draft_stands_checked_again(self):
        s = self.work.source
        delta = 100.0 if self.unit >= 1e3 else 100000.0 / max(self.unit, 1.0)
        check = (s.contracts[self.key].check(self.last, "total_vs_financials") or {}).get("status")
        touched, effect = self.refresh_after([(self.raw["sheet"], self.row, self.raw["value"] / self.unit, delta)])
        _STATE["bumped"], _STATE["effect_tab_only"] = touched, effect
        candidate = s.candidates[self.key]
        if check == "pass":
            # the tab now disagrees with the statement; the deck's figure is the statement's, unchanged
            self.assertEqual(effect["status"], "carried", effect)
            self.assertTrue(candidate.validated and candidate.carried and candidate.accepted)
        else:
            self.assertEqual(effect["status"], "invalidated", effect)        # no statement to follow: the tab IS the figure
        self.assertTrue(all(t.stale for t in self.work.turns if t.mode in ("ask", "promote")))

    def test_5b_a_coordinated_update_moves_the_cited_figure_and_the_draft_says_which_figure_moved(self):
        s = self.work.source
        contract = s.contracts[self.key]
        fin = next((f for f in contract.facts if f["period"] == self.last and f["role"] == "total" and f["basis"] == "financials"), None)
        if fin is None:
            raise unittest.SkipTest("this account has no statement figure: the tab-only change already moved it")
        sheet = s.financials_sheet
        spots = [x for x in find_figure(_STATE["copy"], sheet, fin["value"]) if x[3] == self.unit or True]
        rows = {x[0] for x in spots}
        if len(rows) != 1:
            raise unittest.SkipTest(f"the statement figure sits on {len(rows)} rows of {sheet}: not unique, so not changed")
        delta = 100.0 if self.unit >= 1e3 else 100000.0 / max(self.unit, 1.0)
        value_before = s.candidates[self.key].cite_ids and next(iter(s.candidates[self.key].cite_ids.values()))["value"]
        (row, _col, shown, k) = spots[0]
        touched, effect = self.refresh_after([(sheet, row, abs(shown) if shown >= 0 else shown, delta if shown >= 0 else -delta)])
        _STATE["effect"], _STATE["coordinated"] = effect, True
        candidate = s.candidates[self.key]
        self.assertEqual(effect["status"], "invalidated", effect)
        self.assertTrue(effect["reasons"], "it says which figure moved")
        self.assertFalse(candidate.validated)
        self.assertFalse(candidate.accepted)                              # acceptance does not survive a changed premise
        self.assertEqual(next(v for v in self.work.candidates_view() if v["account"] == self.key)["status"], "invalid")
        self.assertEqual(self.work.export().get("error", {}).get("code"), "NO_ACCEPTED_CANDIDATE")
        others = [e for a, e in {e["account"]: e for e in (self.work.events[-1]["effects"])}.items() if a != self.key]
        self.assertTrue(all(e["status"] == "carried" for e in others), "drafts about other accounts are unaffected")

    def test_6_after_a_change_a_new_draft_can_be_made_and_the_old_deck_is_read_against_the_new_sources(self):
        s = self.work.source
        export = _STATE.get("export")
        self.assertTrue(export)
        diff = T.call_tool(s, "diff_sources", {"baseline_deck": export["path"]})
        self.assertTrue(diff.ok, diff.error)
        self.assertIn(self.key, diff.data["changed_accounts"])         # the deck's own record of the sources says which account moved
        report = self.work.read_deck(export["path"])
        self.assertTrue(report["ok"], report.get("error"))
        self.assertTrue(report["data"]["source_moved"])
        _STATE["report"] = report["data"]

    def returned_copy(self, name, hand_edit_text):
        """A copy of the exported deck in which the paragraph holding `hand_edit_text` has a sentence added by hand."""
        export = _STATE["export"]
        path = os.path.join(_STATE["tmp"], name)
        shutil.copyfile(export["path"], path)
        shutil.copyfile(export["path"] + L.INTERNAL_SUFFIX, path + L.INTERNAL_SUFFIX)
        from pptx import Presentation
        from fdd_utils.agent.export import text_fragments
        marker = "【经理手工补充】"
        wanted = [re.sub(r"\s+", "", x) for x in text_fragments(hand_edit_text)[:2] if x]
        prs, edited = Presentation(path), False
        for slide in prs.slides:
            for shape in slide.shapes:
                if not getattr(shape, "has_text_frame", False):
                    continue
                for para in shape.text_frame.paragraphs:
                    flat = re.sub(r"\s+", "", para.text)
                    if para.runs and any(w and w in flat for w in wanted) and not edited:
                        para.runs[-1].text = para.runs[-1].text + marker
                        edited = True
        prs.save(path)
        self.assertTrue(edited, "found the paragraph to edit by hand")
        return path, marker

    @staticmethod
    def deck_text(path):
        from pptx import Presentation
        return "\n".join(p.text for sl in Presentation(path).slides for sh in sl.shapes if getattr(sh, "has_text_frame", False)
                         for p in sh.text_frame.paragraphs)

    def new_accepted_draft(self):
        self.work.model = self.policy_ask()
        turn = self.work.ask(f"{self.key}的变动来自哪里？")
        self.assertTrue(turn.ok, (turn.stop_reason, turn.question, {k: v for k, v in turn.report.items() if k != "known"}))
        promoted = self.work.report("把这段放入报告")
        self.assertTrue(promoted.ok, (promoted.stop_reason, promoted.question, promoted.report.get("promoted")))

    def test_7_a_hand_edit_in_a_returned_deck_is_kept_and_a_hand_edited_paragraph_is_never_overwritten(self):
        moved = bool(_STATE.get("coordinated"))
        other = next(k for k in self.work.source.base_results if k != self.key and not str(k).startswith("__") and len(self.work.base_text(k)) > 40)
        self.new_accepted_draft()
        # A: the user edited ANOTHER account's paragraph; the changed figure is in this one
        returned_a, marker = self.returned_copy("returned_a.pptx", self.work.base_text(other))
        result_a = self.work.update_deck(returned_a, os.path.join(_STATE["tmp"], "updated_a.pptx"))
        self.assertTrue(result_a["data"].get("out_path") and os.path.exists(result_a["data"]["out_path"]), result_a.get("error"))
        text_a = self.deck_text(result_a["data"]["out_path"])
        self.assertIn(marker, text_a, "a hand edit in a paragraph the update does not touch is kept")
        # B: the user edited THIS account's paragraph
        returned_b, marker_b = self.returned_copy("returned_b.pptx", _STATE["exported_text"])
        result_b = self.work.update_deck(returned_b, os.path.join(_STATE["tmp"], "updated_b.pptx"))
        text_b = self.deck_text(result_b["data"]["out_path"])
        self.assertIn(marker_b, text_b, "a hand-edited paragraph is never overwritten")
        if moved:
            self.assertTrue(result_a["data"]["applied"], "the paragraph nobody touched is brought up to the new figure")
            self.assertGreaterEqual(result_b["data"]["waiting_for_a_person"], 1, "the conflict is reported for a person")
            self.assertEqual(result_b["data"]["applied"], [])
        else:
            self.assertEqual(result_a["data"]["applied"], [], "the cited figure did not move, so nothing is rewritten")
            self.assertEqual(result_b["data"]["applied"], [])
        _STATE["update"] = {"moved": moved, "a_applied": len(result_a["data"]["applied"]), "b_waiting": result_b["data"]["waiting_for_a_person"]}


@unittest.skipUnless(HAVE, "set FDD_FLOW_WORKBOOK and FDD_FLOW_BASE_RUN")
class ChatCli(unittest.TestCase):
    """fdd_agent.py chat, driven through a pipe with a scripted file in place of the model: the first thing a person would try."""

    def test_the_whole_thread_from_the_command_line_and_the_session_file_it_leaves(self):
        import json
        import subprocess
        work = os.path.join(_STATE["tmp"], "cli.xlsx")
        shutil.copyfile(WORKBOOK, work)
        _STATE["copy"], keep = work, _STATE["copy"]
        try:
            source = W.WorkSession.open(work, base_run=BASE_RUN, model=None).source
            picked = pick_account(source)
        finally:
            _STATE["copy"] = keep
        if picked is None:
            self.skipTest("no account suits the flow")
        key, last, _prev, _raw, _row, _unit = picked
        view = cl.evidence_view(source.contracts[key])
        total = next(t for t in view["totals"] if t["period"] == last and t["basis"] == "authority_adjusted")
        replies = [tool("get_account_evidence", account=key),
                   {"answer": {"claims": [{"account": key, "kind": "amount", "periods": [last], "cites": [total["ref"]],
                                           "text": f"截至{date_text(last)}，{key}余额{wan(total['yuan'])}"}], "unresolved": [f"{key}：原因未能核实"]}}]
        script, session_file, deck = (os.path.join(_STATE["tmp"], n) for n in ("cli_script.json", "cli.agent-session.json", "cli_out.pptx"))
        Path(script).write_text(json.dumps(replies, ensure_ascii=False), encoding="utf-8")
        typed = f"{key}的期末余额是多少？\n把这段放入报告\n/candidates\n/export {deck}\n/status\n/quit\n"
        run = subprocess.run([sys.executable, "-X", "utf8", str(ROOT / "fdd_agent.py"), "chat", work, "--base-run", BASE_RUN, "--session", session_file,
                              "--scripted", script, "--max-model-calls", "5"], input=typed, capture_output=True, encoding="utf-8", errors="replace", cwd=str(ROOT), timeout=900,
                             env=dict(os.environ, FDD_AGENT="1", PYTHONIOENCODING="utf-8"))
        out = run.stdout
        self.assertEqual(run.returncode, 0, run.stderr[-600:] + out[-600:])
        self.assertIn("stop: answered", out)
        self.assertIn("PROMOTE", out)                                        # "put that in the report" was a promotion, not a second model call
        self.assertIn("[ACCEPTED]", out)
        self.assertIn("UNRESOLVED", out)
        self.assertTrue(os.path.exists(deck) or any(p.startswith("cli_out") for p in os.listdir(_STATE["tmp"])), out[-800:])
        self.assertIn("model calls 2", out)                                  # the two scripted replies, nothing else
        self.assertTrue(os.path.exists(session_file))
        if os.name != "nt":                                                    # Windows has no owner-only mode bits
            self.assertEqual(oct(os.stat(session_file).st_mode & 0o777), "0o600")
        shown = subprocess.run([sys.executable, "-X", "utf8", str(ROOT / "fdd_agent.py"), "session", "show", session_file],
                               capture_output=True, encoding="utf-8", errors="replace", cwd=str(ROOT), env=dict(os.environ, PYTHONIOENCODING="utf-8"))
        self.assertIn("promote", shown.stdout)
        self.assertIn("accepted", shown.stdout)
        self.assertIn("workbook was not loaded", shown.stdout)


if __name__ == "__main__":
    if not HAVE:
        print("SKIPPED: set FDD_FLOW_WORKBOOK and FDD_FLOW_BASE_RUN (a real workbook and the archived run written from it).")
        sys.exit(0)
    unittest.main(verbosity=2)

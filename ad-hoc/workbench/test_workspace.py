"""A working session across turns: follow-ups, "put that in the report", sources that move, and the new reading tools.

Zero tokens. The model is a scripted policy; the workbooks are tiny files written here with openpyxl; the
contracts are the synthetic ones the other tests use. This proves the thread -- what carries, what is invalidated,
what is refused -- and nothing about how a real model chooses.

    python ad-hoc/workbench/test_workspace.py
"""
from __future__ import annotations

import json
import os
import stat
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parents[1]))

from fdd_utils.agent import controller as C  # noqa: E402
from fdd_utils.agent import tools as T  # noqa: E402
from fdd_utils.agent import workspace as W  # noqa: E402
from fdd_utils.agent.session import SourceSession  # noqa: E402
from fdd_utils.ai import claims as cl  # noqa: E402
from fdd_utils.ai import evidence_contract as ec  # noqa: E402
from fdd_utils.run_identity import file_sha256  # noqa: E402
from test_agent_controller import ADJUSTMENT, REPLACED_ROWS, Policy, evidence_of, replaced_session, session_for, tool, total_ref  # noqa: E402
from test_claims import P1, P2, ROWS, build  # noqa: E402

ACCOUNT = "应收账款"
OTHER = "其他应收款"


def write_book(path: str, remark: str = "应收甲公司的租金因装修免租期延迟收取", hidden_remark: str = "隐藏页：应收款项另有一笔不应被看到的调整"):
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "AR"
    ws["A1"], ws["A2"] = "应收账款", "人民币千元"
    for r, (name, a, b) in enumerate([("甲公司", 600, 700), ("乙公司", 400, 300), ("合计", 1000, 1000)], 4):
        ws.cell(r, 1, name), ws.cell(r, 2, a), ws.cell(r, 3, b)
    ws["A9"] = remark
    ws["D9"] = "备注"
    note = wb.create_sheet("Notes")
    note["B2"] = "其他应收款的调整说明：重分类至应收账款，详见审计调整 4 号"
    hidden = wb.create_sheet("Hidden")
    hidden["A1"] = hidden_remark
    hidden.sheet_state = "hidden"
    wb.save(path)


def session_on(path: str, rows=ROWS, financials=None, notes=None, **kw) -> SourceSession:
    contract = build(rows, financials=financials, notes=notes if notes is not None else [], **kw)
    for f in contract.facts:
        f["sheet"], f["row"] = "AR", 3 + ([r[0] for r in rows].index(f["line"]) if f.get("line") in [r[0] for r in rows] else len(rows) - 1)
    contract.account = ACCOUNT
    return session_for(contract, path)


def answer_policy(text="2025年12月31日应收账款余额130.0万元", account=ACCOUNT):
    def policy(user, results):
        view = evidence_of(results)
        if view is None:
            return tool("get_account_evidence", account=account)
        return {"answer": {"claims": [{"account": account, "kind": "amount", "periods": [P2], "cites": [total_ref(view, P2)], "text": text}],
                           "unresolved": ["原因未能从工作簿核实"]}}
    return Policy(policy)


class Thread(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="work_")
        self.book = os.path.join(self.tmp, "wb.xlsx")
        write_book(self.book)
        self.source = session_on(self.book)
        self.source.base_results = {ACCOUNT: {"final": "应收账款余额较低。其他情况正常。"}, "现金": {"final": "现金余额稳定。"}}
        self.model = answer_policy()
        self.work = W.WorkSession(self.source, model=self.model, budget=C.Budget(max_steps=6))

    def test_a_follow_up_can_cite_what_the_first_turn_fetched_without_fetching_it_again(self):
        first = self.work.ask("应收账款期末余额是多少？")
        self.assertTrue(first.ok, first.report)
        self.assertEqual(first.tool_path, ["get_account_evidence"])
        ref = first.claims[0]["cites"][0]
        seen_prompts = []

        def second(user, results):
            seen_prompts.append(user)
            return {"answer": {"claims": [{"account": ACCOUNT, "kind": "amount", "periods": [P2], "cites": [ref],
                                           "text": "截至2025年12月31日，应收账款余额为130.0万元"}], "unresolved": []}}
        self.work.model = Policy(second)
        turn = self.work.ask("再说一次期末余额")
        self.assertTrue(turn.ok, turn.report)
        self.assertEqual(turn.tool_path, [])                                 # it did not fetch the evidence again
        self.assertIn("EARLIER IN THIS SESSION", seen_prompts[0])
        self.assertIn("应收账款余额130.0万元", seen_prompts[0])
        # the same citation from a controller with no memory is refused: that is what the memory buys
        cold = C.AgentController(self.source, Policy(second)).run("再说一次期末余额")
        self.assertNotEqual(cold.stop_reason, C.STOP_ANSWERED)
        self.assertIn("UNSEEN_REF", json.dumps(cold.steps, ensure_ascii=False))

    def test_put_that_in_the_report_makes_a_draft_from_the_answer_with_no_model_call(self):
        self.work.ask("应收账款期末余额是多少？")
        calls = self.model.calls
        turn = self.work.report("把这段放入报告")
        self.assertTrue(turn.ok, turn.question or turn.report)
        self.assertEqual(self.model.calls, calls)                            # promotion is deterministic
        self.assertEqual(turn.mode, "promote")
        candidate = self.source.candidates[ACCOUNT]
        self.assertTrue(candidate.validated and candidate.accepted)
        self.assertEqual(candidate.origin, "promoted")
        view = self.work.candidates_view()[0]
        self.assertEqual(view["status"], "accepted")
        self.assertEqual(view["replaces"], "应收账款余额较低。其他情况正常。")
        self.assertTrue(view["changes"]["added"] and view["changes"]["removed"])
        self.assertIn("原因未能从工作簿核实", view["unresolved"])               # the open item travels with the draft

    def test_a_request_to_change_the_wording_goes_to_the_model_with_the_evidence_it_already_has(self):
        self.work.ask("应收账款期末余额是多少？")
        prompts = []

        def edit(user, results):
            prompts.append(user)
            view_refs = self.work.memory.history[-1]["claims"][0]["cites"]
            if "validate_candidate(" not in user:
                return tool("validate_candidate", account=ACCOUNT, claims=[{"kind": "amount", "periods": [P2], "cites": view_refs,
                                                                          "text": "应收账款2025年12月31日余额为130.0万元"}])
            return {"done": {"account": ACCOUNT, "export": False, "summary": "shortened"}}
        self.work.model = Policy(edit)
        turn = self.work.report("把这段改写得简短一点再放入报告")             # no account named: taken from the last answer
        self.assertTrue(turn.ok, turn.report)
        self.assertEqual(turn.mode, "edit")
        self.assertIn("EARLIER IN THIS SESSION", prompts[0])
        self.assertNotIn("get_account_evidence", turn.tool_path)             # it did not investigate again
        self.assertTrue(self.source.candidates[ACCOUNT].accepted)

    def test_several_accounts_in_the_last_answer_means_asking_which_not_guessing(self):
        other = build(account=OTHER)
        self.source.contracts[OTHER] = other
        other.binding = self.source.binding.to_dict()

        def both(user, results):
            v1 = evidence_of(results)
            if v1 is None:
                return tool("get_account_evidence", account=ACCOUNT)
            if "get_account_evidence(" in user and user.count("get_account_evidence(") < 2:
                return tool("get_account_evidence", account=OTHER)
            claims = []
            for acct, contract in ((ACCOUNT, self.source.contracts[ACCOUNT]), (OTHER, other)):
                view = cl.evidence_view(contract)
                claims.append({"account": acct, "kind": "amount", "periods": [P2], "cites": [total_ref(view, P2)], "text": f"{acct}余额130.0万元"})
            return {"answer": {"claims": claims, "unresolved": []}}
        self.work.model = Policy(both)
        self.assertTrue(self.work.ask("两个科目的余额").ok)
        calls = self.work.model.calls
        turn = self.work.edit("缩短")
        self.assertFalse(turn.ok)
        self.assertEqual(turn.stop_reason, "needs_user")
        self.assertIn(ACCOUNT, turn.question)
        self.assertEqual(self.work.model.calls, calls)
        both_turn = self.work.report("把这段放入报告")                         # a plain placement keeps every account's claims
        self.assertEqual(sorted(both_turn.accounts), sorted([ACCOUNT, OTHER]))

    def test_with_no_earlier_answer_there_is_nothing_to_put_in_the_report(self):
        turn = self.work.report("把这段放入报告")
        self.assertFalse(turn.ok)
        self.assertIn("Ask a question first", turn.question)
        self.assertEqual(self.model.calls, 0)

    def test_the_model_call_cap_is_a_hard_stop(self):
        self.work.max_model_calls = 1
        turn = self.work.ask("应收账款期末余额是多少？")
        self.assertFalse(turn.ok)
        self.assertEqual(turn.stop_reason, C.STOP_SESSION)             # a cap over the whole session, not a provider failure
        self.assertIn("cap (1)", turn.report["budget"]["message"])
        self.assertEqual(turn.report["budget"]["kind"], "calls")
        self.assertEqual(self.work.model_calls, 1)

    def test_requests_are_told_apart_by_their_words(self):
        for text, want in (("把这段放入报告", "report"), ("把这段放入簡報", "report"), ("put that in the report", "report"),
                           ("把这段缩短后放入报告", "report_edit"), ("rewrite it shorter and add it to the deck", "report_edit"),
                           ("应收账款为什么增加？", "ask"), ("这份报告的应收账款是多少", "ask"), ("", "ask")):
            self.assertEqual(W.classify_request(text), want, text)


class Persistence(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="work_")
        self.book = os.path.join(self.tmp, "wb.xlsx")
        write_book(self.book)

    def fresh_load(self, **kw):
        source = session_on(self.book, **kw)
        source.load_params = {"workbook": self.book}
        return source

    def open_with(self, source, path, **kw):
        with mock.patch.object(SourceSession, "load", classmethod(lambda cls, **_k: source)):
            return W.WorkSession.open(self.book, session_path=path, model=answer_policy(), **kw)

    def test_a_saved_session_resumes_with_its_drafts_derived_figures_and_the_open_question(self):
        path = os.path.join(self.tmp, "s.agent-session.json")
        work = W.WorkSession(self.fresh_load(), model=answer_policy(), path=path, budget=C.Budget(max_steps=6))
        work.source.base_results = {ACCOUNT: {"final": "旧"}}
        contract = work.source.contracts[ACCOUNT]
        share = cl.derive_share(contract, "甲公司", P2)                      # D1 first,
        movement = cl.derive_movement(contract, P1, P2)                      # then D2: the order is part of the session
        work.ask("应收账款期末余额是多少？")
        work.report("把这段放入报告")
        self.assertEqual((share["ref"], movement["ref"]), ("D1", "D2"))
        mode = stat.S_IMODE(os.stat(path).st_mode)
        self.assertEqual(mode, 0o600)
        saved = json.load(open(path, encoding="utf-8"))
        self.assertNotIn("api_key", json.dumps(saved))
        resumed = self.open_with(self.fresh_load(), path)
        self.assertFalse(resumed.opening_notice)
        self.assertEqual([t.mode for t in resumed.turns], ["ask", "promote"])
        candidate = resumed.source.candidates[ACCOUNT]
        self.assertTrue(candidate.accepted and candidate.validated)
        contract2 = resumed.source.contracts[ACCOUNT]
        self.assertEqual([d["ref"] for d in contract2.derived], ["D1", "D2"])
        self.assertEqual(resumed.memory.seen_refs, work.memory.seen_refs)
        self.assertTrue(resumed.unresolved())                                # "原因未能从工作簿核实" is still open

    def test_a_session_file_is_refused_where_git_would_track_it(self):
        with self.assertRaisesRegex(ValueError, "not git-ignored"):
            W.WorkSession(self.fresh_load(), path=str(W.REPO_ROOT / "notes_for_me.json"))
        W.WorkSession(self.fresh_load(), path=os.path.join(self.tmp, "ok.json"))              # outside the repo
        W.WorkSession(self.fresh_load(), path=str(W.REPO_ROOT / "fdd_utils" / "output" / "agent_sessions" / "x.json"))   # ignored

    def test_where_git_cannot_answer_only_the_output_folder_is_accepted(self):
        """A copied or downloaded tree is not a git clone: `git check-ignore` exits 128 there. That is no answer, so fall back to the folder rule."""
        import subprocess
        from unittest import mock
        inside = str(W.REPO_ROOT / "fdd_utils" / "output" / "agent_sessions" / "x.json")
        outside = str(W.REPO_ROOT / "notes_for_me.json")

        def git_says(code):
            return mock.patch.object(W.subprocess, "run", return_value=subprocess.CompletedProcess([], code))
        with git_says(128):
            W.WorkSession._check_path(inside)
            with self.assertRaisesRegex(ValueError, "not git-ignored"):
                W.WorkSession._check_path(outside)
        with git_says(0):
            W.WorkSession._check_path(outside)                                   # git itself says it is ignored
        with git_says(1):
            with self.assertRaisesRegex(ValueError, "not git-ignored"):
                W.WorkSession._check_path(inside)                                # git says it would be tracked: its word wins over the folder rule
        with mock.patch.object(W.subprocess, "run", side_effect=FileNotFoundError("git")):
            W.WorkSession._check_path(inside)                                    # no git installed at all
            with self.assertRaisesRegex(ValueError, "not git-ignored"):
                W.WorkSession._check_path(outside)

    def test_reopening_after_the_workbook_changed_decides_each_draft_by_what_it_cites(self):
        path = os.path.join(self.tmp, "s.json")
        work = W.WorkSession(self.fresh_load(), model=answer_policy(), path=path, budget=C.Budget(max_steps=6))
        work.ask("应收账款期末余额是多少？")
        work.report("把这段放入报告")
        moved = [("甲公司", "detail", {P1: 600000, P2: 700000}), ("乙公司", "detail", {P1: 400000, P2: 300000}),
                 ("丙公司", "detail", {P1: 0, P2: 300000}), ("Total", "total", {P1: 1000000, P2: 1500000})]
        write_book(self.book, remark="saved again with other numbers")       # the file the session was saved against is gone
        resumed = self.open_with(self.fresh_load(rows=moved), path)
        self.assertTrue(resumed.opening_notice["source_changed"])
        candidate = resumed.source.candidates[ACCOUNT]
        self.assertFalse(candidate.validated)
        self.assertIn("2025-12-31", candidate.invalid_reason)
        self.assertFalse(candidate.accepted)
        self.assertTrue(all(t.stale for t in resumed.turns))
        self.assertEqual(resumed.export().get("error", {}).get("code"), "NO_BASE_COMMENTARY")


class SourceMoves(unittest.TestCase):
    """The workbook is saved while the session is open."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="work_")
        self.book = os.path.join(self.tmp, "wb.xlsx")
        write_book(self.book)
        self.source = session_on(self.book)
        self.source.load_params = {"workbook": self.book}
        self.source.base_results = {ACCOUNT: {"final": "旧评论。"}}
        self.work = W.WorkSession(self.source, model=answer_policy(), budget=C.Budget(max_steps=6))
        self.work.ask("应收账款期末余额是多少？")
        self.work.report("把这段放入报告")

    def change_file(self, rows=None, remark="另一句备注"):
        write_book(self.book, remark=remark)
        fresh = session_on(self.book, rows=rows or ROWS)
        fresh.load_params = {"workbook": self.book}
        return mock.patch.object(SourceSession, "load", classmethod(lambda cls, **_k: fresh))

    def test_a_change_elsewhere_in_the_file_leaves_the_draft_standing_and_checked_again(self):
        old_version = self.source.source_version
        with self.change_file():
            effects = self.work.refresh()["candidate_effects"]
        self.assertEqual([e["status"] for e in effects], ["carried"])
        candidate = self.source.candidates[ACCOUNT]
        self.assertTrue(candidate.validated and candidate.accepted and candidate.carried)
        self.assertEqual(candidate.source_version, self.source.source_version)
        self.assertNotEqual(self.source.source_version, old_version)

    def test_a_change_to_a_figure_the_draft_cites_invalidates_it_and_says_which(self):
        moved = [("甲公司", "detail", {P1: 600000, P2: 700000}), ("乙公司", "detail", {P1: 400000, P2: 300000}),
                 ("丙公司", "detail", {P1: 0, P2: 300000}), ("Total", "total", {P1: 1000000, P2: 1500000})]
        with self.change_file(rows=moved):
            effects = self.work.refresh()["candidate_effects"]
        self.assertEqual([e["status"] for e in effects], ["invalidated"])
        reason = effects[0]["reasons"][0]
        self.assertIn("2025-12-31", reason)
        self.assertIn("1,300,000", reason)
        self.assertIn("1,500,000", reason)
        candidate = self.source.candidates[ACCOUNT]
        self.assertFalse(candidate.validated)
        self.assertFalse(candidate.accepted)                                 # acceptance does not survive a changed premise
        self.assertEqual(self.work.candidates_view()[0]["status"], "invalid")
        refused = self.work.export()
        self.assertEqual(refused["error"]["code"], "NO_ACCEPTED_CANDIDATE")

    def test_answers_given_before_the_change_cannot_become_drafts_and_the_model_is_told_not_to_reuse_them(self):
        moved = [("甲公司", "detail", {P1: 600000, P2: 700000}), ("乙公司", "detail", {P1: 400000, P2: 300000}),
                 ("丙公司", "detail", {P1: 0, P2: 300000}), ("Total", "total", {P1: 1000000, P2: 1500000})]
        prompts = []

        def probe(user, results):
            prompts.append(user)
            return tool("inspect_sources")
        with self.change_file(rows=moved):
            self.work.refresh()
        turn = self.work.promote()
        self.assertFalse(turn.ok)
        self.assertEqual(turn.stop_reason, "answer_is_stale")
        self.work.model = Policy(probe)
        self.work.budget = C.Budget(max_steps=1)
        self.work.ask("现在呢？")
        self.assertIn("the workbook has changed since: do not reuse", prompts[0])
        self.assertEqual(self.work.memory.seen_refs, set())                  # nothing fetched before the change may be cited now

    def test_a_turn_notices_a_saved_workbook_by_itself_and_says_what_it_did(self):
        with self.change_file():
            turn = self.work.ask("应收账款期末余额是多少？")
        self.assertTrue(any("the workbook had changed" in n for n in turn.notes), turn.notes)
        self.assertTrue(any(e["type"] == "sources_changed" for e in self.work.events))


class Dependencies(unittest.TestCase):
    """A draft rests on more than the figures it quotes: the other lines of its periods, what the checks and permissions
    allowed, the remarks, the figures a computed number came from. Cited figures that did not move are not enough to
    carry an accepted draft over a change."""

    P3 = "2026-12-31"
    OFFSETTING = [("甲公司", "detail", {P1: 600000, P2: 800000}), ("乙公司", "detail", {P1: 400000, P2: 200000}),
                  ("丙公司", "detail", {P1: 0, P2: 300000}), ("Total", "total", {P1: 1000000, P2: 1300000})]

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="work_")
        self.book = os.path.join(self.tmp, "wb.xlsx")
        write_book(self.book)
        self.source = session_on(self.book, notes=cl_notes())
        self.source.load_params = {"workbook": self.book}
        self.source.base_results = {ACCOUNT: {"final": "旧评论。"}}
        self.work = W.WorkSession(self.source, model=answer_policy(), budget=C.Budget(max_steps=6))

    def accepted_total_draft(self, text="2025年12月31日应收账款余额130.0万元", cause=False):
        contract = self.source.contracts[ACCOUNT]
        claim = {"kind": "amount", "periods": [P2], "cites": [next(f["ref"] for f in contract.facts if f["period"] == P2 and f["role"] == "total"
                                                                  and f["basis"] == "authority_adjusted")], "text": text}
        if cause:
            claim["cause_cites"] = [contract.notes[0]["ref"]]
        result = T.call_tool(self.source, "validate_candidate", {"account": ACCOUNT, "claims": [claim]})
        self.assertTrue(result.ok, result.error)
        self.assertTrue(self.work.accept(ACCOUNT))
        return self.source.candidates[ACCOUNT]

    def moved_to(self, rows=None, mutate=None, **kw):
        write_book(self.book, remark="saved again")
        fresh = session_on(self.book, rows=rows or ROWS, notes=cl_notes(), **kw)
        fresh.load_params = {"workbook": self.book}
        if mutate:
            mutate(fresh.contracts[ACCOUNT])
        return mock.patch.object(SourceSession, "load", classmethod(lambda cls, **_k: fresh))

    def refresh_effect(self, **kw):
        with self.moved_to(**kw):
            return self.work.refresh()["candidate_effects"][0]

    def test_a_cited_total_that_did_not_move_still_changes_what_a_person_accepted_when_its_parts_did(self):
        candidate = self.accepted_total_draft()
        effect = self.refresh_effect(rows=self.OFFSETTING)                    # 甲公司 +100,000 and 乙公司 -100,000: the total is the same
        self.assertEqual(effect["status"], "carried_changed", effect)
        self.assertTrue(candidate.validated and candidate.carried)
        self.assertFalse(candidate.accepted)                                  # a decision made against another account is not carried
        text = " ".join(candidate.context_changes)
        self.assertIn("甲公司", text)
        self.assertIn("800,000", text)
        self.assertEqual(self.work.candidates_view()[0]["needs_review"], True)
        self.assertEqual(self.work.export().get("error", {}).get("code"), "NO_ACCEPTED_CANDIDATE")
        self.assertTrue(any(e["type"] == "acceptance_withdrawn" and e["by"] == "source_change" for e in self.work.events))
        # accepting again is the acknowledgement, and the reasons it acknowledged are kept
        self.assertTrue(self.work.accept(ACCOUNT))
        self.assertEqual(candidate.context_changes, [])
        self.assertTrue(any(e["type"] == "accepted" and e["acknowledged"] for e in self.work.events))

    def test_a_change_in_a_period_the_draft_is_not_about_leaves_the_acceptance_alone(self):
        candidate = self.accepted_total_draft()                               # about 2025-12-31 only
        rows = [("甲公司", "detail", {P1: 650000, P2: 700000}), ("乙公司", "detail", {P1: 350000, P2: 300000}),
                ("丙公司", "detail", {P1: 0, P2: 300000}), ("Total", "total", {P1: 1000000, P2: 1300000})]
        effect = self.refresh_effect(rows=rows)
        self.assertEqual(effect["status"], "carried", effect)
        self.assertTrue(candidate.accepted and not candidate.context_changes)

    def test_a_new_period_changes_which_period_is_the_latest_the_draft_was_read_against(self):
        candidate = self.accepted_total_draft()
        rows = [(n, t, dict(v, **{self.P3: 5.0 if t == "detail" else 15.0})) for n, t, v in ROWS]
        effect = self.refresh_effect(rows=rows, periods=(P1, P2, self.P3))
        self.assertEqual(effect["status"], "carried_changed", effect)
        self.assertIn("added " + self.P3, " ".join(effect["reasons"]))
        self.assertIn(f"latest period is now {self.P3}", " ".join(effect["reasons"]))
        self.assertFalse(candidate.accepted)

    def test_nothing_moved_means_nothing_to_ask_again(self):
        candidate = self.accepted_total_draft()
        effect = self.refresh_effect()
        self.assertEqual(effect["status"], "carried")
        self.assertTrue(candidate.accepted and not candidate.context_changes)

    def test_a_computed_figure_with_the_same_value_but_different_inputs_is_not_the_same_figure(self):
        contract = self.source.contracts[ACCOUNT]
        derived = T.call_tool(self.source, "compute_analysis", {"account": ACCOUNT, "op": "movement", "period_from": P1, "period": P2})
        self.assertTrue(derived.ok, derived.error)
        claim = {"kind": "movement", "periods": [P1, P2], "cites": [derived.data["derived"]["ref"]],
                 "text": "应收账款较2024年12月31日增加30.0万元"}
        result = T.call_tool(self.source, "validate_candidate", {"account": ACCOUNT, "claims": [claim]})
        self.assertTrue(result.ok, result.error)
        self.assertTrue(self.work.accept(ACCOUNT))
        shifted = [("甲公司", "detail", {P1: 700000, P2: 800000}), ("乙公司", "detail", {P1: 400000, P2: 300000}),
                   ("丙公司", "detail", {P1: 0, P2: 300000}), ("Total", "total", {P1: 1100000, P2: 1400000})]     # still +300,000
        effect = self.refresh_effect(rows=shifted)
        self.assertEqual(effect["status"], "invalidated", effect)
        self.assertIn("computed from", " ".join(effect["reasons"]))
        self.assertFalse(self.source.candidates[ACCOUNT].accepted)

    def test_a_note_that_is_the_same_text_but_now_belongs_to_another_period_no_longer_supports_the_cause(self):
        self.accepted_total_draft(text="2025年12月31日应收账款余额130.0万元，主要系甲公司租金延迟收取所致", cause=True)

        def move_note(contract):
            contract.notes[0]["period"] = P1
        effect = self.refresh_effect(mutate=move_note)
        self.assertEqual(effect["status"], "invalidated", effect)
        self.assertIn("now belongs to", " ".join(effect["reasons"]))

    def test_a_check_that_came_out_differently_is_named_even_though_every_figure_and_verdict_is_the_same(self):
        candidate = self.accepted_total_draft()

        def flip(contract):
            for check in contract.checks:
                if check["period"] == P2 and check["kind"] == ec.CHECK_TOTAL_VS_FINANCIALS:
                    check["status"] = "not_checked" if check["status"] != "not_checked" else "pass"
        effect = self.refresh_effect(mutate=flip)
        self.assertEqual(effect["status"], "carried_changed", effect)
        self.assertIn(f"{P2}: check {ec.CHECK_TOTAL_VS_FINANCIALS}", " ".join(candidate.context_changes))
        self.assertFalse(candidate.accepted)

    def test_a_permission_that_changed_is_named(self):
        candidate = self.accepted_total_draft()

        def narrow(contract):
            contract.permissions[P2] = dict(contract.permissions[P2], explain_with_detail=not contract.permissions[P2].get("explain_with_detail"))
        effect = self.refresh_effect(mutate=narrow)
        self.assertEqual(effect["status"], "carried_changed", effect)
        self.assertIn("what the period permits changed (explain_with_detail", " ".join(candidate.context_changes))

    def test_a_draft_saved_without_its_context_cannot_prove_the_account_did_not_move(self):
        path = os.path.join(self.tmp, "s.json")
        self.work.path = path
        candidate = self.accepted_total_draft()
        self.work.save()
        saved = json.load(open(path, encoding="utf-8"))
        self.assertIn(P2, saved["candidates"][ACCOUNT]["context"]["by_period"])        # the context is what is kept, with the cites
        del saved["candidates"][ACCOUNT]["context"]
        json.dump(saved, open(path, "w", encoding="utf-8"))
        # the sources are the same bytes: what they hold now is what the draft was written against, so it is recorded then
        same = session_on(self.book, notes=cl_notes())
        same.load_params = {"workbook": self.book}
        with mock.patch.object(SourceSession, "load", classmethod(lambda cls, **_k: same)):
            resumed = W.WorkSession.open(self.book, session_path=path, model=answer_policy())
        self.assertTrue(resumed.source.candidates[ACCOUNT].context)
        # the sources are other bytes: nothing to compare with, and that is said
        saved["source"]["source_version"] = "other"
        json.dump(saved, open(path, "w", encoding="utf-8"))
        with mock.patch.object(SourceSession, "load", classmethod(lambda cls, **_k: same)):
            reopened = W.WorkSession.open(self.book, session_path=path, model=answer_policy())
        effect = reopened.opening_notice["candidate_effects"][0]
        self.assertEqual(effect["status"], "carried_changed", effect)
        self.assertIn("was not recorded", " ".join(effect["reasons"]))
        self.assertFalse(reopened.source.candidates[ACCOUNT].accepted)

    def test_withdrawing_acceptance_by_hand_is_recorded_too(self):
        self.accepted_total_draft()
        self.assertTrue(self.work.unaccept(ACCOUNT))
        self.assertTrue(any(e["type"] == "acceptance_withdrawn" and e["by"] == "user" for e in self.work.events))


def cl_notes():
    from test_claims import NOTES
    return list(NOTES)


class Reading(unittest.TestCase):
    """read_source_block, search_notes, cross_account_evidence, adjustment_evidence on a small real file."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="work_")
        self.book = os.path.join(self.tmp, "wb.xlsx")
        write_book(self.book)
        self.session = session_on(self.book)

    def call(self, name, **args):
        return T.call_tool(self.session, name, args)

    def test_a_block_comes_back_with_cells_location_version_and_limits(self):
        ev = self.call("get_account_evidence", account=ACCOUNT)
        ref = next(t["ref"] for t in ev.data["totals"] if t["basis"] == "raw_schedule")
        result = self.call("read_source_block", account=ACCOUNT, ref=ref)
        self.assertTrue(result.ok, result.error)
        d = result.data
        self.assertTrue(d["location"].startswith("AR!A"), d["location"])
        self.assertEqual(d["source"]["source_version"], self.session.source_version)
        self.assertEqual(d["source"]["file"], "wb.xlsx")
        self.assertIn("formulas are neither shown", d["limits"]["values"])
        self.assertIn("合计", json.dumps(d["rows"], ensure_ascii=False))
        self.assertEqual(d["sheet_top"][0]["text"], "应收账款")

    def test_the_extractor_row_is_zero_based_and_the_block_is_around_the_excel_row(self):
        item = next(f for f in self.session.contracts[ACCOUNT].facts if f["role"] == "component" and f["line"] == "甲公司")
        d = self.call("read_source_block", account=ACCOUNT, ref=item["ref"]).data
        self.assertEqual(d["highlight_row"], item["row"] + 1)

    def test_a_hidden_sheet_and_an_unknown_sheet_are_refused_and_the_choices_are_visible_ones(self):
        hidden = self.call("read_source_block", sheet="Hidden", row=1)
        self.assertEqual(hidden.error["code"], "HIDDEN_SHEET")
        unknown = self.call("read_source_block", sheet="Nope", row=1)
        self.assertEqual(unknown.error["code"], "NO_SUCH_SHEET")
        self.assertNotIn("Hidden", unknown.data["visible_sheets"])
        self.assertEqual(self.call("read_source_block").error["code"], "BAD_ARGS")

    def test_the_block_is_bounded(self):
        d = self.call("read_source_block", sheet="AR", row=1, rows_after=500, max_cols=500).data
        self.assertLessEqual(d["last_row"] - d["first_row"] + 1, 60)
        self.assertEqual(d["limits"]["columns_read"], 40)

    def test_a_read_after_the_file_changed_is_refused_not_answered_from_the_old_reading(self):
        write_book(self.book, remark="changed")
        self.assertEqual(self.call("read_source_block", sheet="AR", row=1).error["code"], "STALE_SOURCE")

    def test_search_finds_remarks_on_visible_sheets_only_and_files_them_on_the_account(self):
        result = self.call("search_notes", query="装修 免租期", account=ACCOUNT)
        self.assertTrue(result.ok, result.error)
        hit = result.data["hits"][0]
        self.assertEqual((hit["sheet"], hit["address"]), ("AR", "A9"))
        self.assertTrue(hit["ref"].startswith("N"))
        self.assertIn(hit["ref"], result.basis)                              # the model may cite it
        self.assertEqual(result.data["searched"]["hidden_sheets_skipped"], 1)
        self.assertEqual(self.call("search_notes", query="不应被看到").data["hits"], [])
        self.assertIn("only visible sheets are read", result.data["limits"])
        again = self.call("search_notes", query="装修 免租期", account=ACCOUNT).data["hits"][0]
        self.assertEqual(again["ref"], hit["ref"])                           # not filed twice

    def test_a_searched_remark_can_be_cited_for_a_cause_and_is_marked_not_verified(self):
        ref = self.call("search_notes", query="装修 免租期", account=ACCOUNT).data["hits"][0]["ref"]
        view = cl.evidence_view(self.session.contracts[ACCOUNT])
        claim = {"account": ACCOUNT, "kind": "cause", "periods": [P2], "cites": [total_ref(view, P2)], "cause_cites": [ref],
                 "text": "2025年12月31日应收账款余额130.0万元，主要系甲公司租金因装修免租期延迟收取所致"}
        verdict = cl.verify_claim(claim, self.session.contracts)
        self.assertEqual(verdict["verdict"], "verified_inference", verdict["findings"])
        self.assertIn("CAUSE_NOTE_RELATED_UNCHECKED", {f["code"] for f in verdict["findings"]})

    def test_search_notes_do_not_count_as_a_change_in_the_workbook(self):
        before = self.session.snapshot()[ACCOUNT]["notes"]
        self.call("search_notes", query="装修", account=ACCOUNT)
        self.assertEqual(self.session.snapshot()[ACCOUNT]["notes"], before)

    def test_an_empty_query_is_an_error(self):
        self.assertEqual(self.call("search_notes", query="  ").error["code"], "BAD_ARGS")

    def test_adjustment_evidence_names_entries_and_the_gap_when_there_is_none(self):
        explained = replaced_session(with_adjustment=True)
        gap = replaced_session(with_adjustment=False)
        e = T.call_tool(explained, "adjustment_evidence", {"account": ACCOUNT}).data
        g = T.call_tool(gap, "adjustment_evidence", {"account": ACCOUNT}).data
        p1 = next(p for p in e["periods"] if p["period"] == P1)
        self.assertTrue(p1["replaced_by_financials"])
        self.assertEqual(p1["difference_tab_minus_financials"], -200000.0)
        self.assertEqual(len(p1["adjustment_entries"]), 1)
        self.assertTrue(p1["adjustment_entries"][0]["ref"].startswith("N"))
        self.assertTrue(p1["entries_add_up_to_difference"])                  # "+20.0万元" is the 200,000 gap
        self.assertEqual(e["periods_with_a_gap_and_no_entry"], [])
        self.assertEqual(g["periods_with_a_gap_and_no_entry"], [P1])
        self.assertIsNone(next(p for p in g["periods"] if p["period"] == P1)["entries_add_up_to_difference"])
        self.assertTrue(e["limits"])

    def test_cross_account_evidence_says_what_it_did_not_cover(self):
        d = self.call("cross_account_evidence", account=ACCOUNT)
        self.assertTrue(d.ok, d.error)
        self.assertEqual(d.data["edges"], [])
        self.assertTrue(any("not a cause" in x for x in d.data["limits"]))
        self.assertEqual(d.data["source"]["source_version"], self.session.source_version)
        self.assertEqual((d.data["failed_checks"], d.data["failed_checks_omitted"]), ([], 0))
        self.assertEqual(d.data["coverage"]["touching_this_account"], 0)

    def with_links(self, links):
        self.session.text_cache[("cross", self.session.source_version)] = {"links": links}

    def three_links(self):
        from fdd_utils.ai.facts import _edge
        return [
            _edge(f"{ACCOUNT}::甲公司", OTHER, "shared_counterparty", [P1, P2], "normalised company-like breakdown label is identical", True,
                  {"normalised_label": "甲公司"}),
            _edge(f"{ACCOUNT}::breakdown", ACCOUNT, "component_to_total", [P1, P2], "|breakdown - total| <= max(1.0, 1% x total)", False,
                  {"component_count": 3, "tied_periods": [P1], "differing_periods": [P2], "used_total_row": True}),
            _edge("WorkbookSheet::Notes#block:0#row:5", ACCOUNT, "supporting_schedule_tieout", [P2], "labelled total row ties to account", False,
                  {"sheet": "Notes", "row": 5, "values": {P2: 1250000.0}, "account_values": {P2: 1300000.0}, "matched_accounts": [ACCOUNT, OTHER], "ambiguous": True})]

    def test_a_check_that_did_not_tie_comes_back_with_what_was_compared_and_why_but_nothing_that_can_be_cited(self):
        self.with_links(self.three_links())
        d = self.call("cross_account_evidence", account=ACCOUNT)
        self.assertTrue(d.ok, d.error)
        self.assertEqual([e["kind"] for e in d.data["edges"]], ["shared_counterparty"])                 # only what passed is support
        self.assertEqual(d.data["tested_and_not_tied"], {"component_to_total": 1, "supporting_schedule_tieout": 1})   # the counts are still there
        kinds = {c["kind"]: c for c in d.data["failed_checks"]}
        self.assertEqual(set(kinds), {"component_to_total", "supporting_schedule_tieout"})
        tie = kinds["component_to_total"]
        self.assertEqual((tie["result"], tie["supports_a_claim"]), ("did not tie", False))
        self.assertIn(f"differ in {P2}", tie["why"])
        self.assertEqual(tie["compared"]["differing_periods"], [P2])
        self.assertIn("|breakdown - total|", tie["test"])
        ambiguous = kinds["supporting_schedule_tieout"]
        self.assertIn("more than one account", ambiguous["why"])
        self.assertEqual(ambiguous["compared"]["account_values"], {P2: 1300000.0})                       # the numbers that did not meet
        self.assertEqual(ambiguous["compared"]["values"], {P2: 1250000.0})
        self.assertIn(OTHER, ambiguous["compared"]["matched_accounts"])
        self.assertEqual(d.data["coverage"], {"accounts_examined": 1, "relationships_tested_in_workbook": 3, "touching_this_account": 3, "passed": 1, "failed": 2})
        self.assertEqual(d.basis, [], "no fact, derived figure or note ref appears: a failed check is not a thing a claim can cite")
        self.assertTrue(any("cannot be cited" in x for x in d.data["limits"]))

    def test_the_list_of_failed_checks_is_bounded_and_says_how_many_it_left_out(self):
        self.with_links(self.three_links())
        d = self.call("cross_account_evidence", account=ACCOUNT, max_failed=1)
        self.assertEqual((len(d.data["failed_checks"]), d.data["failed_checks_omitted"]), (1, 1))
        self.assertEqual(self.call("cross_account_evidence", account=ACCOUNT, max_failed=0).data["failed_checks_omitted"], 2)

    def test_what_a_failed_check_compared_is_trimmed_not_dumped(self):
        from fdd_utils.ai.facts import _edge
        evidence = {f"key{i}": ("x" * 300 if i == 0 else i) for i in range(20)}
        evidence["differing_periods"] = [f"2020-{m:02d}-28" for m in range(1, 13)]
        self.with_links([_edge(f"{ACCOUNT}::breakdown", ACCOUNT, "component_to_total", [P1], "t", False, evidence)])
        compared = self.call("cross_account_evidence", account=ACCOUNT).data["failed_checks"][0]["compared"]
        self.assertLessEqual(len(compared), 9)
        self.assertIn("...", compared)
        self.assertLessEqual(len(compared["key0"]), 103)
        why = self.call("cross_account_evidence", account=ACCOUNT).data["failed_checks"][0]["why"]
        self.assertIn(" ...", why)


class BaseRun(unittest.TestCase):
    """Commentary carried over from an earlier run is only as good as the proof that the run described this workbook."""

    def setUp(self):
        import pandas as pd
        from fdd_utils.agent.export import session_delivery
        from fdd_utils.ai.evidence import frame_fingerprint
        from fdd_utils.agent.session import check_base_run
        self.check, self.delivery, self.fp = check_base_run, session_delivery, frame_fingerprint
        self.dfs = {"A": pd.DataFrame({"x": [1.0, 2.0]}), "B": pd.DataFrame({"x": [3.0, 4.0]})}
        self.now = {"workbook_sha256": "abc", "entity": "", "sheet": "Financials", "financials_sha256": None, "financials_sheet": "", "language": "Chi"}
        self.results = {"A": {"final": "a"}, "B": {"final": "b"}, "__run_health__": {}}

    def loaded(self, identity=None, fingerprints=None):
        return {"results": self.results, "identity": identity, "fingerprints": fingerprints or {}}

    def test_an_identity_that_matches_proves_the_run(self):
        self.assertEqual(self.check(self.loaded(dict(self.now)), self.dfs, self.now)["status"], "verified")

    def test_fingerprints_alone_prove_it_when_every_account_has_one(self):
        fp = {k: self.fp(v) for k, v in self.dfs.items()}
        self.assertEqual(self.check(self.loaded(None, fp), self.dfs, self.now)["status"], "verified")

    def test_nothing_to_prove_it_is_unverified_not_assumed(self):
        out = self.check(self.loaded(), self.dfs, self.now)
        self.assertEqual(out["status"], "unverified")
        self.assertEqual(out["accounts"]["no_fingerprint"], ["A", "B"])
        half = self.check(self.loaded(None, {"A": self.fp(self.dfs["A"])}), self.dfs, self.now)
        self.assertEqual((half["status"], half["accounts"]["matched"]), ("unverified", 1))

    def test_a_workbook_saved_after_the_run_is_judged_account_by_account_when_the_run_has_fingerprints(self):
        fp = {k: self.fp(v) for k, v in self.dfs.items()}
        edited_identity = dict(self.now, workbook_sha256="older-bytes")
        untouched = self.check(self.loaded(edited_identity, fp), self.dfs, self.now)
        self.assertEqual((untouched["status"], untouched["workbook_edited_since"]), ("verified", True))     # edited elsewhere
        fp_changed = dict(fp, A="stale-hash")
        out = self.check(self.loaded(edited_identity, fp_changed), self.dfs, self.now)
        self.assertEqual((out["status"], out["accounts"]["changed"]), ("edited_since", ["A"]))
        # without fingerprints a saved-after workbook cannot be sorted into touched and untouched accounts
        blind = self.check(self.loaded(edited_identity), self.dfs, self.now)
        self.assertEqual(blind["status"], "mismatch")
        self.assertTrue(any("saved after the run" in r for r in blind["reasons"]))
        # another sheet or entity is a mismatch whatever the fingerprints say
        other = self.check(self.loaded(dict(self.now, entity="other"), fp), self.dfs, self.now)
        self.assertEqual(other["status"], "mismatch")

    def test_another_workbook_or_changed_data_is_a_mismatch_that_names_what_differs(self):
        other = dict(self.now, workbook_sha256="zzz")
        self.assertEqual(self.check(self.loaded(other), self.dfs, self.now)["status"], "mismatch")
        fp = {"A": "stale-hash", "B": self.fp(self.dfs["B"])}
        out = self.check(self.loaded(None, fp), self.dfs, self.now)
        self.assertEqual((out["status"], out["accounts"]["changed"]), ("edited_since", ["A"]))
        gone = self.check(dict(self.loaded(dict(self.now)), results={"Z": {"final": "z"}}), self.dfs, self.now)
        self.assertEqual((gone["status"], gone["accounts"]["missing"]), ("mismatch", ["Z"]))

    def decision(self, check, summaries):
        contract = build()
        source = session_for(contract)
        source.base_results, source.base_summaries, source.base_run_check = {"应收账款": {"final": "x"}}, summaries, check
        results = {"应收账款": {"final": "x", "agent_4_validation": {"clause_reviews": [{"clause": "x", "supported": True, "category": "data-backed"}]}},
                   "__run_health__": {"zero_successful_calls": False, "accounts_on_fallback_bullet": [], "accounts_on_passthrough": [],
                                      "accounts_on_error_text": [], "accounts_without_prompt": []}}
        return self.delivery(source, results, edited=["应收账款"])

    def test_the_delivery_decision_says_what_the_base_run_does_and_does_not_prove(self):
        codes_of = lambda d: {f.code: f.severity for f in d.findings}
        unverified = self.decision({"status": "unverified", "reasons": ["x"], "accounts": {"in_run": 26, "matched": 0, "changed": [], "missing": [], "no_fingerprint": ["A"]}}, {})
        self.assertEqual(codes_of(unverified)["BASE_RUN_UNVERIFIED"], "review")
        self.assertEqual(codes_of(unverified)["SUMMARY_NOT_REGENERATED"], "review")          # no saved summary: the band is a splice
        self.assertEqual(unverified.status, "review_required")
        mismatch = self.decision({"status": "mismatch", "reasons": ["identity differs -- workbook_sha256"],
                                  "accounts": {"in_run": 2, "matched": 0, "changed": ["应收账款", "现金"], "missing": [], "no_fingerprint": []}}, {"BS": "s"})
        found = [f for f in mismatch.findings if f.code in ("BASE_RUN_MISMATCH", "BASE_ACCOUNT_CHANGED")]
        self.assertEqual(mismatch.status, "blocked")
        self.assertEqual([f.account for f in found if f.code == "BASE_ACCOUNT_CHANGED"], ["现金"])   # the edited account is replaced, the rest is stale
        since = self.decision({"status": "edited_since", "reasons": ["x"], "accounts": {"in_run": 2, "matched": 0, "changed": ["应收账款", "现金"], "missing": [], "no_fingerprint": []}}, {"BS": "s"})
        self.assertEqual({f.code for f in since.findings if f.severity == "blocked" and f.code.startswith("BASE")}, {"BASE_ACCOUNT_CHANGED"})
        self.assertNotIn("BASE_RUN_MISMATCH", {f.code for f in since.findings})                         # only the changed accounts, not the whole run
        verified = self.decision({"status": "verified", "reasons": [], "accounts": {"in_run": 1, "matched": 1, "changed": [], "missing": [], "no_fingerprint": []}}, {"BS": "s"})
        self.assertEqual(codes_of(verified)["SUMMARY_NOT_REGENERATED"], "info")
        self.assertNotIn("BASE_RUN_UNVERIFIED", codes_of(verified))


if __name__ == "__main__":
    unittest.main(verbosity=2)

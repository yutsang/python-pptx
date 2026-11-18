"""The whole-databook run (fdd_utils/agent/book.py) and what the app files from it. Zero tokens: the writer is scripted, the workbook a tiny file.

    python ad-hoc/workbench/test_book.py

What this proves: every account in scope is written, the ones the verifier accepts are finished and the one it refuses is carried with
its sentence marked, the runaway guard is the per-account one, the session says it wrote the whole book (so it exports its drafts), an
account can be written again with the person's guidance, and the state the app reads is filed. What it cannot show is how a model writes
or how the page looks; the scripted run through a real workbook is baseline_snapshot.py's.
"""
from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parents[1]))

from fdd_utils.agent import book as B  # noqa: E402
from fdd_utils.agent import controller as C  # noqa: E402
from fdd_utils.agent import flag as F  # noqa: E402
from fdd_utils.agent import routing as R  # noqa: E402
from fdd_utils.agent import workspace as W  # noqa: E402
from test_agent_workflow import GOOD, OTHER, Writer, amount_claim, writer_view  # noqa: E402
from test_claims import build as build_contract  # noqa: E402
from test_workspace import ACCOUNT, P1, P2, session_on, write_book  # noqa: E402


def two_account_work(fn, config=None):
    tmp = tempfile.mkdtemp(prefix="book_")
    book = os.path.join(tmp, "wb.xlsx")
    write_book(book)
    source = session_on(book)
    source.load_params = {"workbook": book}
    other = build_contract(account=OTHER)
    other.binding = source.binding.to_dict()
    source.contracts[OTHER] = other
    return W.WorkSession(source, model=Writer(fn), model_name="m", policy=R.Policy(mode="single"), budget=C.Budget(max_steps=8),
                         prepare=True, config=config or {"agent": {}})


def finishes_one_account(system, user, n):
    """The right balance for ACCOUNT. For OTHER the right balance and a wrong movement, which the verifier refuses however often it is
    asked to repair it (so OTHER ends as a draft with one verified sentence and one marked)."""
    view = writer_view(user)
    claim = amount_claim(view, P2, GOOD)
    if view["account"] == ACCOUNT:
        return {"draft": {"claims": [claim], "unresolved": []}} if "REJECTED" not in user else {"fix": [{"index": 0, "claim": claim}]}
    move = next(d["ref"] for d in view["derived"] if d["kind"] == "movement")
    wrong = {"kind": "movement", "periods": [P1, P2], "cites": [move], "text": "较2024年12月31日增加99.0万元"}
    if "REJECTED" in user:
        return {"fix": [{"index": 1, "claim": wrong}], "add": []}
    return {"draft": {"claims": [claim, wrong], "unresolved": []}}


class RunBook(unittest.TestCase):
    def test_every_account_in_scope_is_written_the_refused_one_is_carried_marked_and_the_session_says_it_wrote_the_book(self):
        work = two_account_work(finishes_one_account)
        seen = []
        result = B.run_book(work, on_account=lambda key, turn: seen.append(key), overview=False, deck=False, issues_dir=None)
        self.assertEqual(sorted(seen), sorted([ACCOUNT, OTHER]))
        self.assertEqual(result.finished, [ACCOUNT])
        self.assertEqual(result.marked, [OTHER], "the draft the gate refused is still in the deck, its wrong sentence marked")
        self.assertIn(ACCOUNT, result.ai_results)
        self.assertIn(OTHER, result.ai_results)
        reviews = result.ai_results[OTHER]["agent_4_validation"]["clause_reviews"]
        self.assertTrue(any(r["category"] == "hallucination" and not r["supported"] for r in reviews), reviews)
        self.assertTrue(work.deck.get("whole_book"), "an export of this session is a first generation without being told")
        self.assertEqual(result.stopped, "")
        lines = result.summary_lines("wb.xlsx")
        self.assertIn("accounts in scope 2   finished 1", lines[1])
        self.assertIn("drafts carried with marked sentences 1", lines[1])

    def test_the_guard_on_model_calls_is_per_account_not_the_ask_and_edit_cap(self):
        self.assertEqual(B.book_cap({}, 10), 40 * 12)
        self.assertEqual(B.book_cap({"agent": {"max_calls_per_account": 5}}, 10), 60)
        self.assertIsNone(B.book_cap({"agent": {"max_calls_per_account": 0}}, 10))
        self.assertEqual(F.max_calls_per_account({"agent": {"max_model_calls": 60}}), 40, "max_model_calls is for one ask or edit session")
        work = two_account_work(finishes_one_account, config={"agent": {"max_model_calls": 60, "max_calls_per_account": 7}})
        B.run_book(work, overview=False, deck=False, issues_dir=None)
        self.assertEqual(work.max_model_calls, 7 * (2 + 2), "a book of two accounts: the per-account guard, not the 60 of an ask session")
        explicit = two_account_work(finishes_one_account)
        B.run_book(explicit, overview=False, deck=False, issues_dir=None, max_calls=3)
        self.assertEqual(explicit.max_model_calls, 3)

    def test_a_book_is_held_to_the_per_account_token_guard_not_the_session_cap_of_an_ask(self):
        self.assertEqual(B.book_token_cap({}, 10), 120_000 * 12)
        self.assertEqual(B.book_token_cap({"agent": {"max_tokens_per_account": 5000}}, 10), 60_000)
        self.assertIsNone(B.book_token_cap({"agent": {"max_tokens_per_account": 0}}, 10))
        work = two_account_work(finishes_one_account, config={"agent": {"max_session_tokens": 400_000, "max_tokens_per_account": 7000}})
        work.shared.limits.max_tokens = 400_000                       # what agent_limits reads from max_session_tokens
        B.run_book(work, overview=False, deck=False, issues_dir=None)
        self.assertEqual(work.shared.limits.max_tokens, 7000 * 4, "the session cap of 400,000 is replaced for a book")
        explicit = two_account_work(finishes_one_account)
        B.run_book(explicit, overview=False, deck=False, issues_dir=None, max_tokens=1234)
        self.assertEqual(explicit.shared.limits.max_tokens, 1234)

    def test_the_run_stops_and_says_so_when_the_cap_is_reached_and_nothing_is_raised(self):
        work = two_account_work(finishes_one_account)
        result = B.run_book(work, overview=False, deck=False, issues_dir=None, max_calls=1)
        self.assertTrue(result.stopped, "a cap that stopped the run is in the result")
        self.assertLess(len(result.finished), 2)

    def test_an_account_can_be_written_again_with_the_persons_guidance(self):
        work = two_account_work(finishes_one_account)
        B.run_book(work, overview=False, deck=False, issues_dir=None)
        sent = []

        def again(system, user, n):
            sent.append(user)
            return finishes_one_account(system, user, n)
        work.model = Writer(again)
        result = B.redraft(work, ACCOUNT, "请突出余额与上年末的差异")
        self.assertTrue(any("请突出余额与上年末的差异" in u for u in work.model.users), "the guidance reaches the writer")
        self.assertIn("agent_4_validation", result)
        self.assertEqual(result["agent_4_validation"]["clause_reviews"][0]["supported"], True)

    def test_the_refused_account_is_kept_as_a_case_and_nothing_else_is_written_to_disk_without_a_folder(self):
        work = two_account_work(finishes_one_account)
        with tempfile.TemporaryDirectory() as folder:
            result = B.run_book(work, overview=False, deck=False, issues_dir=folder)
            kept = list(Path(folder).rglob("*.json"))
        self.assertEqual(result.issue_cases, 1)
        self.assertEqual(len(kept), 1)


class CommandLine(unittest.TestCase):
    def test_fdd_agent_book_runs_the_same_book_and_prints_its_summary(self):
        import contextlib
        import io
        import fdd_agent
        work = two_account_work(finishes_one_account)
        out = io.StringIO()
        with mock.patch.object(B, "open_book", return_value=work) as opened, contextlib.redirect_stdout(out):
            code = fdd_agent.main(["book", "wb.xlsx", "--model", "dashscope", "--model-name", "m", "--no-deck", "--max-calls", "50", "--max-tokens", "90000"])
        text = out.getvalue()
        self.assertEqual(opened.call_args.kwargs["model_type"], "dashscope")
        self.assertEqual(opened.call_args.kwargs["model_name"], "m")
        self.assertIn("accounts in scope 2   finished 1", text)
        self.assertIn("drafts carried with marked sentences 1", text)
        self.assertEqual(code, 2, "no deck was asked for, so nothing is ready to deliver")
        self.assertEqual((work.max_model_calls, work.shared.limits.max_tokens), (50, 90000))


class AppState(unittest.TestCase):
    def test_what_the_app_files_from_a_run_and_that_a_blocked_deck_is_still_offered(self):
        from fdd_utils.ui import agent_run
        work = two_account_work(finishes_one_account)
        result = B.run_book(work, overview=False, deck=False, issues_dir=None)
        with tempfile.TemporaryDirectory() as tmp:
            deck = os.path.join(tmp, "x_DRAFT-BLOCKED.pptx")
            Path(deck).write_bytes(b"deck")
            result.export = {"ok": False, "data": {"path": deck, "delivery": {"status": "blocked", "findings": []}},
                             "error": {"code": "DELIVERY_BLOCKED", "message": "draft"}}
            state = agent_run.book_state(result)
        self.assertEqual(state["pptx_download_data"], b"deck")
        self.assertTrue(state["pptx_ready"])
        self.assertEqual(state["delivery_decision"]["status"], "blocked")
        self.assertIs(state[agent_run.WORK_KEY], work)
        self.assertEqual(sorted(state["ai_results"]), sorted(["__run_health__", ACCOUNT, OTHER]))

    def test_no_deck_means_not_ready_and_no_bytes(self):
        from fdd_utils.ui import agent_run
        work = two_account_work(finishes_one_account)
        result = B.run_book(work, overview=False, deck=False, issues_dir=None)
        state = agent_run.book_state(result)
        self.assertFalse(state["pptx_ready"])
        self.assertNotIn("pptx_download_data", state)

    def test_the_batch_phase_returns_the_bundle_the_batch_view_swaps_in(self):
        from fdd_utils.ui import agent_run
        work = two_account_work(finishes_one_account)
        extracted = {"_internal": {"entity_name": "甲实体", "temp_path": work.source.workbook_path, "selected_sheet": "AR", "financials_from": None,
                                   "financials_sheet": None, "mapping_overrides": None, "effective_language": "Chn",
                                   "raw_state": {"dfs": {}, "workbook_list": []}, "reconciliation": None, "resolution": None}}
        with tempfile.TemporaryDirectory() as tmp:
            deck = os.path.join(tmp, "e.pptx")
            Path(deck).write_bytes(b"deck")
            real_run = B.run_book

            def run(w, accounts=None, **kw):
                done = real_run(w, accounts, overview=False, deck=False, issues_dir=None, on_account=kw.get("on_account"))
                done.export = {"ok": True, "data": {"path": deck, "delivery": {"status": "ready"}}}
                return done
            progress = []
            with mock.patch.object(B, "open_book", return_value=work), mock.patch.object(B, "run_book", side_effect=run):
                out = agent_run.run_for_entity(extracted=extracted, model_type="dashscope", model_name="m",
                                               progress_callback=lambda done, total, key: progress.append((done, total, key)))
        self.assertEqual(out["status"], "ok")
        self.assertEqual([p[:2] for p in progress], [(1, 2), (2, 2)])
        bundle = out["state"]
        for key in ("ai_results", "section_summaries", "delivery_decision", "pptx_download_data", "pptx_download_filename", "agent_work",
                    "dfs", "entity_name", "language"):
            self.assertIn(key, bundle)
        self.assertEqual(bundle["entity_name"], "甲实体")


if __name__ == "__main__":
    unittest.main()

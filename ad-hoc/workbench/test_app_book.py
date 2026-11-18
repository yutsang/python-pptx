"""The app's commentary run, driven with Streamlit's own test harness. Zero tokens: the run is the scripted whole-book run of test_book.py.

    python ad-hoc/workbench/test_app_book.py

Proves: the first load starts the run once and files its results where the tabs read them, a click on the page does not start (and pay
for) another, a run that failed says why and waits for the person, an account can be written again from its tab, and the redo button
clears the run. It cannot show how the page looks, and it does not read a workbook (that is baseline_snapshot.py's, on real books).
"""
from __future__ import annotations

import os
import sys
import unittest
from pathlib import Path
from unittest import mock

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parents[1]))

from streamlit.testing.v1 import AppTest  # noqa: E402

from fdd_utils.agent import book as B  # noqa: E402
from fdd_utils.ui import agent_run  # noqa: E402
from test_book import finishes_one_account, two_account_work  # noqa: E402
from test_workspace import ACCOUNT  # noqa: E402

OTHER = "其他应收款"


def page(mappings):
    """The commentary section of the processed view, as render_processed_view calls it, over a loaded workbook with no commentary yet."""
    import streamlit as st
    from fdd_utils.ui import ai_panel
    st.session_state.setdefault("dfs", {"loaded": object()})
    for key, value in (("ai_results", None), ("language", "Chn"), ("model_type", "dashscope"), ("model_name", "m"), ("account_comments", {}),
                       ("pptx_ready", False), ("temp_path", "wb.xlsx")):
        st.session_state.setdefault(key, value)
    ai_panel.render_ai_generation_section(st.session_state, lambda model: model)
    ai_panel.render_generated_content(st.session_state, {}, mappings, lambda model: model)


class Run(unittest.TestCase):
    def setUp(self):
        self.work = two_account_work(finishes_one_account)
        self.runs = []
        self.mappings = {ACCOUNT: {"type": "BS"}, OTHER: {"type": "BS"}}

    def scripted_run(self, session_state, *, progress_callback=None):
        self.runs.append(1)
        accounts = self.work.entity_accounts()
        done = []

        def on_account(key, turn):
            done.append(key)
            if progress_callback:
                progress_callback(len(done), len(accounts), key)
        return B.run_book(self.work, accounts, on_account=on_account, overview=False, deck=False, issues_dir=None)

    def app(self):
        return AppTest.from_function(page, kwargs={"mappings": self.mappings}, default_timeout=30)

    def test_the_first_load_writes_the_book_once_and_a_click_on_the_page_does_not_write_it_again(self):
        with mock.patch.object(agent_run, "run_for_state", side_effect=self.scripted_run):
            at = self.app().run()
            self.assertFalse(at.exception, at.exception)
            self.assertEqual(len(self.runs), 1)
            self.assertEqual(sorted(at.session_state["ai_results"]), sorted(["__run_health__", ACCOUNT, OTHER]))
            self.assertIs(at.session_state["agent_work"], self.work)
            html = " ".join(m.value for m in at.markdown)
            self.assertIn("fdd-final-commentary", html)
            self.assertIn("fdd-hallucination-clause", html, "the sentence the verifier refused is painted red in its tab")
            at.run()                                                              # any later interaction re-runs the script
            self.assertEqual(len(self.runs), 1, "the page does not start a second paid run")

    def test_a_run_that_fails_says_why_and_is_not_started_again_by_the_next_click(self):
        calls = []

        def boom(session_state, *, progress_callback=None):
            calls.append(1)
            raise ValueError("the primary model asks for provider 'dashscope', which is not configured here")
        with mock.patch.object(agent_run, "run_for_state", side_effect=boom):
            at = self.app().run()
            self.assertEqual(len(calls), 1)
            self.assertIn("not configured here", at.session_state["book_error"])
            at.run()
            self.assertEqual(len(calls), 1, "a failed run is not retried by a click on the page")
            self.assertTrue(any("did not produce commentary" in e.value for e in at.error), [e.value for e in at.error])
            self.assertEqual(at.button(key="book_retry").label, "Try again")
        with mock.patch.object(agent_run, "run_for_state", side_effect=self.scripted_run):
            at.button(key="book_retry").click().run()
            self.assertEqual(len(self.runs), 1, "Try again starts it")
            self.assertIsNone(at.session_state["book_error"])

    def test_an_account_is_written_again_from_its_tab_with_the_persons_guidance(self):
        with mock.patch.object(agent_run, "run_for_state", side_effect=self.scripted_run):
            at = self.app().run()
        seen = []

        def redraft(work, account, guidance=""):
            seen.append((account, guidance))
            return {"final": "重写后的文字", "agent_4_validation": {"clause_reviews": [], "final_content": "重写后的文字"}}
        with mock.patch.object(B, "redraft", side_effect=redraft):
            at.text_area(key=f"bs_{ACCOUNT}_reprompt_comment").set_value("请强调变动").run()
            at.button(key=f"bs_{ACCOUNT}_reprompt_button").click().run()
        self.assertFalse(at.exception, at.exception)
        self.assertEqual(seen, [(ACCOUNT, "请强调变动")])
        self.assertEqual(at.session_state["ai_results"][ACCOUNT]["final"], "重写后的文字")
        self.assertFalse(at.session_state["pptx_ready"], "the deck written before is out of date")

    def test_a_model_that_did_not_finish_the_account_leaves_the_earlier_text_and_says_so(self):
        with mock.patch.object(agent_run, "run_for_state", side_effect=self.scripted_run):
            at = self.app().run()
        before = at.session_state["ai_results"][ACCOUNT]["final"]
        with mock.patch.object(B, "redraft", return_value={}):
            at.button(key=f"bs_{ACCOUNT}_reprompt_button").click().run()
        self.assertEqual(at.session_state["ai_results"][ACCOUNT]["final"], before)
        self.assertTrue(any("earlier text stands" in w.value for w in at.warning), [w.value for w in at.warning])


class Demo(unittest.TestCase):
    def test_the_replayed_demo_needs_no_model_and_no_session(self):
        from fdd_utils.ui import ai_panel
        keys = ["Cash"]
        sleeps = []
        with mock.patch.object(ai_panel.time, "sleep", side_effect=lambda s: sleeps.append(s)):
            class Slot:
                def progress(self, *_a, **_k): return None
                def info(self, *_a, **_k): return None
            results = ai_panel._run_demo_ai(keys, 1, Slot(), Slot())
        self.assertIn("Cash", results)
        self.assertTrue(results["Cash"].get("final") or results["Cash"].get("subagent_4") or results["Cash"].get("subagent_1"))


if __name__ == "__main__":
    unittest.main()

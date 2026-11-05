"""The agent workbench panel in the app, driven with Streamlit's own test harness. Zero tokens, no provider.

    python ad-hoc/workbench/test_agent_panel.py

The panel is a face over WorkSession (tested in test_workspace.py); these tests check what it shows and that it is
absent while the agent path is off. The "model" is a scripted policy. What this cannot show is how the page looks.
"""
from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parents[1]))

from streamlit.testing.v1 import AppTest  # noqa: E402

from fdd_utils.agent import controller as C  # noqa: E402
from fdd_utils.agent import workspace as W  # noqa: E402
from test_workspace import ACCOUNT, answer_policy, session_on, write_book  # noqa: E402


def app(work, path):
    import streamlit as st
    from fdd_utils.ui.agent_panel import render_agent_panel
    st.session_state.setdefault("temp_path", path)
    render_agent_panel(st.session_state, factory=lambda _s: work)


def nested_expanders(node, inside=False):
    """Labels of expanders found inside another expander. Older Streamlit versions raise on this at run time; a newer one may not."""
    found = []
    for child in getattr(node, "children", {}).values():
        is_expander = getattr(child, "type", "") == "expander"
        if is_expander and inside:
            found.append(getattr(child, "label", "?"))
        found += nested_expanders(child, inside or is_expander)
    return found


class Panel(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="panel_")
        self.book = os.path.join(self.tmp, "wb.xlsx")
        write_book(self.book)
        source = session_on(self.book)
        source.base_results = {ACCOUNT: {"final": "应收账款余额较低。"}}
        self.work = W.WorkSession(source, model=answer_policy(), budget=C.Budget(max_steps=6))
        self.saved = os.environ.pop("FDD_AGENT", None)

    def tearDown(self):
        os.environ.pop("FDD_AGENT", None)
        if self.saved is not None:
            os.environ["FDD_AGENT"] = self.saved

    def run_app(self):
        return AppTest.from_function(app, kwargs={"work": self.work, "path": self.book}, default_timeout=20).run()

    def test_while_the_agent_path_is_off_the_panel_renders_nothing(self):
        os.environ["FDD_AGENT"] = "0"
        at = self.run_app()
        self.assertFalse(at.exception)
        self.assertEqual(len(at.expander), 0)
        self.assertEqual(len(at.button), 0)

    def test_ask_then_put_it_in_the_report_then_the_draft_is_shown_accepted_with_what_is_open(self):
        os.environ["FDD_AGENT"] = "1"
        at = self.run_app()
        self.assertFalse(at.exception, at.exception)
        at.button(key="agent_start").click().run()
        self.assertEqual(self.work.turns, [])                                # starting a session asks the model nothing
        at.text_area(key="agent_request").set_value("应收账款期末余额是多少？").run()
        at.button(key="agent_ask").click().run()
        self.assertFalse(at.exception, at.exception)
        self.assertEqual(len(self.work.turns), 1)
        self.assertIn("应收账款余额130.0万元", " ".join(m.value for m in at.markdown))
        self.assertTrue(any("原因未能从工作簿核实" in w.value for w in at.warning))      # the open question is on the page
        calls = self.work.model.calls
        at.text_area(key="agent_request").set_value("把这段放入报告").run()
        at.button(key="agent_report").click().run()
        self.assertFalse(at.exception, at.exception)
        self.assertEqual(self.work.model.calls, calls)                       # a plain placement is not a model call
        self.assertEqual(self.work.candidates_view()[0]["status"], "accepted")
        page = " ".join(m.value for m in at.markdown)
        self.assertIn("ACCEPTED", page)
        self.assertIn("What it replaces", page)
        self.assertEqual(nested_expanders(at.main), [], "Streamlit refuses an expander inside an expander; the panel is itself one")
        self.assertIn("应收账款余额较低。", page)                              # what the draft would replace is shown
        self.assertFalse(at.button(key="agent_export").disabled)

    def test_an_export_with_no_base_run_is_refused_in_words_not_silently(self):
        os.environ["FDD_AGENT"] = "1"
        self.work.source.base_results = None
        at = self.run_app()
        at.button(key="agent_start").click().run()
        at.text_area(key="agent_request").set_value("应收账款期末余额是多少？").run()
        at.button(key="agent_ask").click().run()
        at.button(key="agent_report").click().run()
        at.button(key="agent_export").click().run()
        self.assertTrue(any("NO_BASE_COMMENTARY" in e.value for e in at.error), [e.value for e in at.error])
        # the refusal says how to get a whole first deck: the legacy command, then binding the session to that run
        self.assertTrue(any("--run-ai --export-pptx" in c.value for c in at.code), [c.value for c in at.code])
        self.assertTrue(any("/base-run" in c.value for c in at.code))

    def test_overrides_a_partial_update_and_a_doubtful_base_run_are_shown(self):
        os.environ["FDD_AGENT"] = "1"
        self.work.source.load_params = dict(self.work.source.load_params or {}, mapping_overrides={"应收账款": "AR"})
        self.work.source.base_run_check = {"status": "edited_since", "reasons": ["account data differs"]}
        self.work.deck = {"last_update": {"complete": False, "unfinished": [{"action": "conflict"}], "next_steps": ["decide the conflict"]}}
        at = self.run_app()
        at.button(key="agent_start").click().run()
        self.assertFalse(at.exception, at.exception)
        captions = " ".join(c.value for c in at.caption)
        warnings = " ".join(w.value for w in at.warning)
        self.assertIn("应收账款 → AR", captions)
        self.assertIn("edited_since", warnings)
        self.assertIn("PARTIAL", warnings)

    def test_a_draft_that_no_longer_stands_says_why_and_cannot_be_accepted(self):
        os.environ["FDD_AGENT"] = "1"
        at = self.run_app()
        at.button(key="agent_start").click().run()
        at.text_area(key="agent_request").set_value("应收账款期末余额是多少？").run()
        at.button(key="agent_ask").click().run()
        at.button(key="agent_report").click().run()
        candidate = self.work.source.candidates[ACCOUNT]
        candidate.invalid_reason, candidate.accepted = "应收账款 2025-12-31: 1,300,000 is now 1,500,000", False
        at = self.run_app()                                                  # a fresh page over the same session object
        at.button(key="agent_start").click().run()
        self.assertTrue(any("1,500,000" in e.value for e in at.error), [e.value for e in at.error])
        self.assertEqual([b for b in at.button if b.key == f"agent_accept_{ACCOUNT}"], [])

    def test_a_draft_the_account_moved_around_shows_what_moved_and_asks_for_a_second_look(self):
        os.environ["FDD_AGENT"] = "1"
        at = self.run_app()
        at.button(key="agent_start").click().run()
        at.text_area(key="agent_request").set_value("应收账款期末余额是多少？").run()
        at.button(key="agent_ask").click().run()
        at.text_area(key="agent_request").set_value("把这段放入报告").run()
        at.button(key="agent_report").click().run()
        candidate = self.work.source.candidates[ACCOUNT]
        candidate.accepted = False
        candidate.context_changes = ["2025-12-31: the account's other figures moved (1 changed, 0 new, 0 gone): 甲公司 [raw_schedule] 700,000 -> 800,000"]
        at.run()
        self.assertTrue(any("The account moved around this draft" in w.value and "甲公司" in w.value for w in at.warning), [w.value for w in at.warning])
        self.assertIn("agent_accept_" + ACCOUNT, [b.key for b in at.button])
        self.assertTrue(any(b.label == "Accept again, knowing what moved" for b in at.button), [b.label for b in at.button])
        self.assertIn("tokens charged", " ".join(c.value for c in at.caption))



if __name__ == "__main__":
    unittest.main(verbosity=2)

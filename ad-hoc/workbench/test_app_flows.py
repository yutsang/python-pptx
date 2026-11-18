"""The real app (fdd_app.py), driven through its single-file and batch views on a real workbook with a SCRIPTED writer. Zero tokens.

    FDD_FLOW_WORKBOOK=<a databook .xlsx> python ad-hoc/workbench/test_app_flows.py

Skipped without a workbook: name one in FDD_FLOW_WORKBOOK, or leave databook.xlsx in the repo folder. Proves that the page, the extraction,
the commentary run, the deck, the delivery banner and the batch loop (two entities, the ZIP and the combined deck) hold together after the
legacy pipeline was removed. Says nothing about how a model writes, how the page looks, or real PowerPoint.
"""
from __future__ import annotations

import os
import sys
import unittest
from pathlib import Path
from unittest import mock

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(ROOT))

WORKBOOK = os.environ.get("FDD_FLOW_WORKBOOK") or str(ROOT / "databook.xlsx")
TEMPLATE = ROOT / "fdd_utils" / "template.pptx"


class Upload:
    def __init__(self, name: str, data: bytes):
        self.name, self.data, self.size = name, data, len(data)

    def getvalue(self) -> bytes:
        return self.data

    def read(self, *_a) -> bytes:
        return self.data


@unittest.skipUnless(os.path.exists(WORKBOOK) and TEMPLATE.exists(), "needs a databook (FDD_FLOW_WORKBOOK) and fdd_utils/template.pptx")
class Flows(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.data = Path(WORKBOOK).read_bytes()

    def app(self, uploads, run_for_state=None, batch_entity=None):
        import streamlit as st
        from streamlit.testing.v1 import AppTest
        import fdd_utils.ui as ui
        from fdd_utils.ui import agent_run
        patches = [mock.patch.object(st, "file_uploader", side_effect=lambda *a, **k: uploads if k.get("key") == "file_uploader" else None)]
        if run_for_state:
            patches.append(mock.patch.object(agent_run, "run_for_state", side_effect=run_for_state))
        if batch_entity:
            patches.append(mock.patch.object(ui, "batch_run_book_for_entity", side_effect=batch_entity))
        for p in patches:
            p.start()
            self.addCleanup(p.stop)
        return AppTest.from_file(str(ROOT / "fdd_app.py"), default_timeout=600)

    @staticmethod
    def scripted_work(path):
        import entity_trial
        from fdd_utils.agent import book as B
        return B.open_book(path, model=entity_trial.ScriptedWriter("flash"), model_name="scripted")

    def test_single_file_process_then_the_commentary_run_then_the_deck_and_the_banner(self):
        from fdd_utils.agent import book as B

        def scripted_run(session_state, *, progress_callback=None):
            work = self.scripted_work(session_state.get("temp_path"))
            return B.run_book(work, issues_dir=None, on_account=lambda key, turn: progress_callback and progress_callback(1, 1, key))
        at = self.app([Upload("databook.xlsx", self.data)], run_for_state=scripted_run)
        at.run()
        self.assertFalse(at.exception, at.exception)
        at.text_input[0].set_value("HN").run()
        next(b for b in at.button if "Process Data" in b.label).click().run()
        self.assertFalse(at.exception, [e.value[:300] for e in at.exception])
        state = at.session_state
        written = [k for k in (state["ai_results"] or {}) if not str(k).startswith("__")]
        self.assertTrue(written, "the commentary the run wrote is filed where the tabs read it")
        self.assertTrue(state["pptx_ready"] and len(state["pptx_download_data"]) > 10_000, "the deck is ready as bytes")
        self.assertEqual(type(state["agent_work"]).__name__, "WorkSession")
        self.assertIn(state["delivery_decision"]["status"], ("ready", "review", "blocked"))
        self.assertTrue(any("BS" == t.label for t in at.tabs), [t.label for t in at.tabs][:5])
        if state["delivery_decision"]["status"] != "ready":
            self.assertTrue(any("DELIVERY" in e.value or "REVIEW" in e.value for e in list(at.error) + list(at.warning)),
                            "a deck that is not ready says so on the page")
        calls = state["agent_work"].model_calls
        at.run()
        self.assertEqual(state["agent_work"].model_calls, calls, "a click on the page writes nothing again")

    def test_batch_two_entities_go_through_extraction_the_agent_the_zip_and_the_combined_deck(self):
        from fdd_utils.ui import agent_run
        import fdd_utils.ui as ui
        from fdd_utils.agent import book as B
        from fdd_utils.agent.workspace import WorkSession
        import entity_trial
        real = agent_run.run_for_entity
        seen = []

        def scripted_entity(*, extracted, model_type="local", model_name=None, progress_callback=None):
            seen.append(extracted["_internal"]["entity_name"])
            with mock.patch.object(B, "open_book", side_effect=lambda path, **kw: WorkSession.open(path, model=entity_trial.ScriptedWriter("flash"), model_name="scripted")):
                return real(extracted=extracted, model_type=model_type, model_name=model_name, progress_callback=progress_callback)
        at = self.app([Upload("Entity One.xlsx", self.data), Upload("Entity Two.xlsx", self.data)], batch_entity=scripted_entity)
        at.run()
        self.assertFalse(at.exception, at.exception)
        for field in at.text_input:
            if field.label == "Entity name" and not field.value:
                field.set_value("Ent " + field.key[-3:])
        at.run()
        at.button(key="batch_start").click().run()
        self.assertFalse(at.exception, [e.value[:300] for e in at.exception])
        state = at.session_state
        self.assertTrue(state["batch_processing_done"])
        self.assertEqual(len(seen), 2)
        self.assertEqual(state["batch_failed_entities"], [])
        self.assertGreater(len(state["batch_zip_data"]), 10_000)
        self.assertTrue(state["batch_combined_pptx"], "two decks are merged into one")


if __name__ == "__main__":
    unittest.main()

"""One hidden-sheet rule at every place a Financials sheet is chosen or read.

Zero tokens; small workbooks written here with openpyxl:

    python ad-hoc/workbench/test_sheet_visibility.py

Entry points covered: the agent session, the CLI's auto-detection, the CLI's explicit --sheet /
--financials-sheet, the CLI's roll-up matching, the resolver's reference sheet, the UI's sheet options
(preflight) and the one function that reads the Financials sheet. A static guard fails if a new module
matches sheet names against "financials" without going through the visibility rule. In every case a
VISIBLE sheet must behave as it did before.
"""
from __future__ import annotations

import os
import re
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

import pandas as pd

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
sys.path.insert(0, str(ROOT))

from fdd_utils.workbook import visibility as V  # noqa: E402


def make_book(tmp, name="wb.xlsx", **states):
    """A workbook whose sheets are named by the keywords, in order, each with the given state."""
    from openpyxl import Workbook
    wb = Workbook()
    first = True
    for title, state in states.items():
        ws = wb.active if first else wb.create_sheet()
        first = False
        ws.title = title.replace("__", " - ")
        ws["A1"] = "Balance sheet"
        ws["A2"] = "Cash"
        ws["B2"] = 100
        ws.sheet_state = state
    path = os.path.join(tmp, name)
    wb.save(path)
    return path


class Primitives(unittest.TestCase):
    def test_states_filter_and_find(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = make_book(tmp, Data="visible", Financials="hidden", Other="veryHidden", Fin2="visible")
            self.assertEqual(V.sheet_states(path), {"Data": "visible", "Financials": "hidden", "Other": "veryHidden", "Fin2": "visible"})
            self.assertEqual(V.filter_visible(path, ["Financials", "Data", "Fin2", "Nope"]), ["Data", "Fin2"])
            self.assertEqual(V.hidden_among(path, ["Financials", "Data", "Other"]), ["Financials", "Other"])
            self.assertIsNone(V.find_visible_financials_sheet(path))

    def test_unreadable_visibility_chooses_nothing_and_refuses_a_named_sheet(self):
        original = V.sheet_states
        V.sheet_states = lambda path: None
        try:
            self.assertEqual(V.filter_visible("x.xlsx", ["Financials"]), [])
            self.assertIsNone(V.find_visible_financials_sheet("x.xlsx"))
            with self.assertRaises(V.HiddenSheetError):
                V.refuse_hidden("x.xlsx", "Financials")
        finally:
            V.sheet_states = original

    def test_a_missing_file_is_unreadable_not_visible(self):
        self.assertIsNone(V.sheet_states("/nonexistent/none.xlsx"))

    def test_a_buffer_is_outside_what_can_be_checked(self):
        import io
        V.refuse_hidden(io.BytesIO(b""), "Financials")           # no exception


class AgentSession(unittest.TestCase):
    def test_it_is_the_shared_rule(self):
        from fdd_utils.agent import session
        self.assertIs(session.refuse_hidden, V.refuse_hidden)
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(session.find_financials_sheet(make_book(tmp, Data="visible", Financials="visible")), "Financials")
            self.assertIsNone(session.find_financials_sheet(make_book(tmp, "h.xlsx", Data="visible", Financials="hidden")))


class FunnelReader(unittest.TestCase):
    """statements.extract_balance_sheet_and_income_statement is where the sheet is read."""

    def test_a_hidden_sheet_is_refused_before_anything_is_read(self):
        from fdd_utils.workbook import extract_balance_sheet_and_income_statement
        with tempfile.TemporaryDirectory() as tmp:
            path = make_book(tmp, Data="visible", Financials="hidden")
            with self.assertRaisesRegex(V.HiddenSheetError, "hidden"):
                extract_balance_sheet_and_income_statement(workbook_path=path, sheet_name="Financials")

    def test_a_visible_sheet_is_read_as_before(self):
        from fdd_utils.workbook import extract_balance_sheet_and_income_statement
        with tempfile.TemporaryDirectory() as tmp:
            path = make_book(tmp, Data="visible", Financials="visible")
            result = extract_balance_sheet_and_income_statement(workbook_path=path, sheet_name="Financials")
            self.assertEqual(set(result), {"balance_sheet", "income_statement", "project_name"})


class CliAutoDetect(unittest.TestCase):
    def setUp(self):
        import inspect_databook
        self.cli = inspect_databook

    def names(self, path):
        with pd.ExcelFile(path) as book:                                    # closed before the temp folder goes (Windows refuses otherwise)
            return self.cli._resolve_financials_sheets(book, path)

    def test_visible_sheets_keep_the_old_choice_and_order(self):
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(self.names(make_book(tmp, Data="visible", Financials="visible")), ["Financials"])
            self.assertEqual(self.names(make_book(tmp, "b.xlsx", Data="visible", Financials__NB="visible", Financials__HN="visible")),
                             ["Financials - NB", "Financials - HN"])
            self.assertEqual(self.names(make_book(tmp, "c.xlsx", Data="visible", 上海松江Financials="visible")), ["上海松江Financials"])

    def test_hidden_ones_are_dropped_and_a_hidden_exact_match_does_not_beat_a_visible_one(self):
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(self.names(make_book(tmp, Data="visible", Financials="hidden")), [])
            self.assertEqual(self.names(make_book(tmp, "b.xlsx", Data="visible", Financials="hidden", Financials__NB="visible")), ["Financials - NB"])
            self.assertEqual(self.names(make_book(tmp, "c.xlsx", Financials__NB="veryHidden", Financials__HN="visible")), ["Financials - HN"])

    def test_unreadable_visibility_chooses_nothing(self):
        original = V.sheet_states
        V.sheet_states = lambda path: None
        try:
            with tempfile.TemporaryDirectory() as tmp:
                self.assertEqual(self.names(make_book(tmp, Data="visible", Financials="visible")), [])
        finally:
            V.sheet_states = original


class CliExplicitAndRollup(unittest.TestCase):
    @unittest.skipUnless((ROOT / "fdd_utils" / "config.yml").exists(), "the CLI reads fdd_utils/config.yml, which a fresh clone does not have (copy config.example.yml)")
    def test_a_named_hidden_sheet_stops_the_cli_with_a_message_not_a_crash(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = make_book(tmp, Data="visible", Financials="hidden")
            run = subprocess.run([sys.executable, "-X", "utf8", str(ROOT / "inspect_databook.py"), path, "--sheet", "Financials"],
                                 capture_output=True, encoding="utf-8", errors="replace", cwd=str(ROOT), timeout=240, env=dict(os.environ, PYTHONIOENCODING="utf-8"))
            self.assertIn("is hidden", run.stdout + run.stderr)
            self.assertNotIn("Traceback", run.stderr)

    def test_hidden_roll_up_sheets_are_not_matched_to_entities(self):
        import inspect_databook as cli
        with tempfile.TemporaryDirectory() as tmp:
            hidden_roll = make_book(tmp, "Group.Portfolio I.databook. 主表.xlsx", Cover="visible", EntityA="hidden")
            entity = Path(tmp) / "Group.Portfolio I.EntityA.xlsx"
            entity.write_bytes(b"")
            self.assertEqual(cli.resolve_auto_rollup_targets([Path(hidden_roll), entity]), {})
            visible_roll = make_book(tmp, "Group.Portfolio II.databook. 主表.xlsx", Cover="visible", EntityB="visible")
            entity_b = Path(tmp) / "Group.Portfolio II.EntityB.xlsx"
            entity_b.write_bytes(b"")
            targets = cli.resolve_auto_rollup_targets([Path(visible_roll), entity_b])
            self.assertEqual({k.name: v[1] for k, v in targets.items()}, {"Group.Portfolio II.EntityB.xlsx": "EntityB"})


class ResolverAndUi(unittest.TestCase):
    def test_the_resolver_never_takes_a_hidden_sheet_as_its_reference(self):
        from fdd_utils.workbook import _pick_financial_summary_sheet
        profiles = {"Financials": {"sheet_kind": "financial_summary", "is_hidden": True},
                    "Financials - NB": {"sheet_kind": "financial_summary", "is_hidden": False}}
        self.assertEqual(_pick_financial_summary_sheet(profiles), "Financials - NB")
        self.assertIsNone(_pick_financial_summary_sheet({"Financials": {"sheet_kind": "financial_summary", "is_hidden": True}}))

    def test_the_ui_options_exclude_hidden_sheets(self):
        from fdd_utils.workbook.preflight import build_workbook_preflight, get_financial_sheet_options
        with tempfile.TemporaryDirectory() as tmp:
            options = get_financial_sheet_options(build_workbook_preflight(make_book(tmp, Data="visible", Financials="hidden", Fin2="visible")))
            self.assertNotIn("Financials", options)


class StaticGuard(unittest.TestCase):
    def test_every_caller_of_the_cli_resolver_passes_the_path(self):
        """The resolver needs the path to read sheet visibility; a one-argument call is a TypeError that only
        shows when that script is run (it broke the stage replay once)."""
        offenders = []
        for path in list(ROOT.rglob("*.py")):
            if ".venv" in path.parts or path.name == "test_sheet_visibility.py":
                continue
            for number, line in enumerate(path.read_text(encoding="utf-8", errors="ignore").splitlines(), 1):
                if "_resolve_financials_sheets(" in line and "def _resolve_financials_sheets" not in line and not re.search(r"_resolve_financials_sheets\([^,()]*(\([^()]*\))?[^,()]*,", line):
                    offenders.append(f"{path.relative_to(ROOT)}:{number}")
        self.assertEqual(offenders, [])

    def test_no_module_matches_sheet_names_against_financials_without_the_rule(self):
        """A file that both lists sheet names and matches "financials" must reference the visibility rule
        (or the profile's is_hidden / the preflight's visible-only list). Fails when a new one is added."""
        offenders = []
        paths = [p for p in (ROOT / "fdd_utils").rglob("*.py") if "/agent/" not in p.as_posix() or True]
        paths += [ROOT / "inspect_databook.py", ROOT / "fdd_app.py", ROOT / "fdd_agent.py"]
        for path in paths:
            text = path.read_text(encoding="utf-8")
            lines = text.splitlines()
            near = any(re.search(r"sheet_names|sheetnames|\.worksheets", lines[i]) and
                       any(re.search(r"financials", l, re.IGNORECASE) for l in lines[max(0, i - 8): i + 8])
                       for i in range(len(lines)))
            if near and not re.search(r"visibility|is_hidden|_visible_non_blank_sheets|filter_visible", text):
                offenders.append(path.relative_to(ROOT).as_posix())
        self.assertEqual(offenders, [], "a Financials sheet picked by name without the hidden-sheet rule")


if __name__ == "__main__":
    unittest.main(verbosity=2)

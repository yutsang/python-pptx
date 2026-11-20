"""verify_update.py: what it shows of each step, how it reads a step's end, and that it runs a command and keeps all of its output.

    python ad-hoc/workbench/test_verify_update.py

Does not run the gate or the comparison themselves; those are verify_update.py's real run.
"""
from __future__ import annotations

import io
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

import verify_update as vu  # noqa: E402


class Reading(unittest.TestCase):
    def test_the_gate_is_read_from_its_last_verdict_line(self):
        self.assertEqual(vu.gate_verdict(["ok   test_a  3 test(s)", "ALL GREEN: 830 test(s), revision x"]), (True, "ALL GREEN: 830 test(s), revision x"))
        ok, text = vu.gate_verdict(["FAIL test_b  4 test(s)", "      x test_b.T.t: boom", "NOT GREEN: 830 test(s), revision x"])
        self.assertFalse(ok)
        self.assertTrue(text.startswith("NOT GREEN"))
        self.assertFalse(vu.gate_verdict(["nothing"])[0])

    def test_the_comparison_is_read_from_its_result_block(self):
        lines = [" book ...", "==== RESULT ====", "code that took the snapshot: 3986ffef323b   code now: c64cb648b6fa",
                 "books compared: 18   left out (they do not run on their own): 2", "IDENTICAL (contracts, ...)"]
        ok, text = vu.compare_verdict(lines)
        self.assertTrue(ok)
        self.assertIn("books compared: 18", text)
        self.assertFalse(vu.compare_verdict(["==== RESULT ====", "code that took the snapshot: a   code now: b", "3 DIFFERENCE LINE(S)"])[0])
        self.assertFalse(vu.compare_verdict(["the old code's snapshot did not finish (exit 1)"])[0])
        self.assertFalse(vu.compare_verdict(["==== RESULT ====", "NOT A COMPARISON: both fingerprints are the same code."])[0])

    def test_what_is_shown_of_the_comparison_is_the_progress_the_result_and_the_differences(self):
        shown = [l for l in [" Some book.xlsx ...", "  Some book.xlsx   contracts 18 ab12", "Financials authority NOT applied", "==== RESULT ====",
                             "2 DIFFERENCE LINE(S)", "  A.xlsx: CONTRACTS differ in 1 of 24", "    k: was 1  now 2", "tqdm 50%"] if vu.show_compare(l)]
        self.assertEqual(shown, [" Some book.xlsx ...", "==== RESULT ====", "2 DIFFERENCE LINE(S)", "  A.xlsx: CONTRACTS differ in 1 of 24", "    k: was 1  now 2"])


class Folder(unittest.TestCase):
    def test_a_clean_and_complete_folder_passes_and_either_kind_of_difference_fails(self):
        with mock.patch("stale_files.compare", return_value=([], [])):
            self.assertEqual(vu.folder_verdict()[0], True)
        with mock.patch("stale_files.compare", return_value=(["old.py"], ["new.py"])):
            ok, lines = vu.folder_verdict()
        self.assertFalse(ok)
        self.assertIn("    not in main: old.py", lines)
        self.assertIn("    missing here: new.py", lines)


class Running(unittest.TestCase):
    def test_a_command_is_run_all_its_output_is_kept_and_only_what_is_asked_for_is_shown(self):
        log = io.StringIO()
        printed = io.StringIO()
        with mock.patch("sys.stdout", printed):
            code, lines = vu.run_and_show([sys.executable, "-c", "print('keep me'); print('noise'); print('中文 ok')"], HERE,
                                          lambda l: l.startswith("keep"), log)
        self.assertEqual(code, 0)
        self.assertEqual(lines, ["keep me", "noise", "中文 ok"])
        self.assertEqual(log.getvalue().splitlines(), lines)
        self.assertEqual(printed.getvalue().strip(), "keep me")


if __name__ == "__main__":
    unittest.main()

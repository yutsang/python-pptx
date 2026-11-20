"""baseline_snapshot.py goes on past a book it cannot open, and holds the new code to the same answer for it.

    python ad-hoc/workbench/test_baseline_snapshot.py

Found on the company machine 2026-11-20: the first file of databooks/ was a roll-up master that needs an entity, and its error ended the
whole save, so the 18 books after it were never read.
"""
from __future__ import annotations

import sys
import unittest
from pathlib import Path
from unittest import mock

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parents[1]))

import baseline_snapshot as bs  # noqa: E402


class Unopenable(unittest.TestCase):
    def snapshot_with_two_books_that_do_not_open(self):
        def refuse(book):
            raise RuntimeError(f"No eligible accounts after reconciliation filtering. ({Path(book).name})")
        with mock.patch.object(bs, "local_books", return_value={"h1": "A.xlsx", "h2": "B.xlsx"}), mock.patch.object(bs, "open_work", side_effect=refuse):
            return bs.snapshot()

    def test_every_book_is_tried_and_the_reason_is_kept(self):
        snap = self.snapshot_with_two_books_that_do_not_open()
        self.assertEqual(sorted(snap["books"]), ["A.xlsx", "B.xlsx"])
        self.assertIn("No eligible accounts", snap["books"]["A.xlsx"]["error"])

    def test_the_same_answer_twice_is_no_difference_and_a_changed_one_is_named(self):
        snap = self.snapshot_with_two_books_that_do_not_open()
        self.assertEqual(bs.diff(snap, snap), [])
        ran = {"A.xlsx": {"contracts": {}}, "B.xlsx": snap["books"]["B.xlsx"]}
        lines = bs.diff({"books": ran}, snap)
        self.assertEqual(len(lines), 1)
        self.assertIn("A.xlsx: ERROR was none, it ran", lines[0])


if __name__ == "__main__":
    unittest.main()

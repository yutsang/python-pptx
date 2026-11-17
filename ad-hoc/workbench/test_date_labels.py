"""Period labels at the source: which header row a tab's dates are read from, and what a header cell parses to.

Zero tokens, no workbook on disk (the tab is a DataFrame):

    python ad-hoc/workbench/test_date_labels.py

Written for an acceptance book on which 15 of 21 source tabs gave every account the period labels 2023-01-01, 2024-01-01 and
2025-01-01 where the sheet's header says December month ends. Two things were wrong together. A Chinese date broken across a
line or spaced out ("2023年\\n12月31日") did not parse, so the real date row scored nothing; and a helper row of real date cells
formatted as a bare year did parse, so it was chosen. The verifier checks a sentence against the evidence, so a wrong label
inside the evidence passed it.
"""
from __future__ import annotations

import sys
import unittest
from datetime import datetime as dt
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

from fdd_utils.workbook.inspector import date_row_index  # noqa: E402
from fdd_utils.workbook.schedules import normalize_financial_schedule  # noqa: E402
from fdd_utils.workbook.statements import parse_date  # noqa: E402


def iso(value):
    return value.strftime("%Y-%m-%d") if value else None


def tab(date_cells, helper=None):
    """A source tab shaped like the client's: title row (optionally a helper row of dates above the stage row), stage row, date
    row, two lines and a total."""
    helper = helper if helper is not None else [None, None, None, None, None, None]
    return pd.DataFrame([
        ["Back", None, None, None, None, None, None, None, None],
        [None, None, "AR", *helper, None],
        [None, None, None, "Mgt acc", None, None, None, "Audit adjustment", None],
        ["Category", None, "CNY'000", *date_cells, None],
        [1, None, "Customer A", 100.0, 110.0, 120.0, 130.0, None, None],
        [2, None, "Customer B", 50.0, 60.0, 70.0, 80.0, None, None],
        [None, None, "Total", 150.0, 170.0, 190.0, 210.0, 0, 0],
        [None, None, "Check", 0, 0, 0, 0, 0, 0],
    ])


NEWLINE_DATES = ["2023年\n12月31日", "2024年\n12月31日", "2025年\n12月31日", "2026年\n1月31日", "2023年\n12月31日", "2024年\n12月31日"]
YEAR_STARTS = [dt(2023, 1, 1), dt(2024, 1, 1), dt(2025, 1, 1), dt(2026, 1, 1), dt(2023, 1, 1), dt(2024, 1, 1)]


class HeaderCells(unittest.TestCase):
    def test_a_chinese_date_broken_or_spaced_is_one_date(self):
        for text in ("2023年\n12月31日", "2023年\r\n12月31日", "2023 年 12 月 31 日", "2023年 12月31日", "2023年　12月31日", "2023年 12月31日"):
            for bare in (True, False):
                with self.subTest(text=text, allow_bare_number=bare):
                    self.assertEqual(iso(parse_date(text, allow_bare_number=bare)), "2023-12-31")

    def test_a_period_spaced_out_keeps_its_end(self):
        self.assertEqual(iso(parse_date("2026年 1-6 月")), "2026-06-30")
        self.assertEqual(iso(parse_date("2023 年")), "2023-12-31")

    def test_dates_written_with_dots_are_dates_when_choosing_the_row(self):
        self.assertEqual(iso(parse_date("2023.12.31", allow_bare_number=False)), "2023-12-31")

    def test_numbers_are_still_refused_when_choosing_the_row(self):
        for value in ("2023", 2023, 2023.0, "1,234", "1,234.56", "2023.12", "12.", ".5", 45291, "45291"):
            with self.subTest(value=value):
                self.assertIsNone(parse_date(value, allow_bare_number=False))

    def test_text_that_only_contains_spaces_is_left_alone(self):
        self.assertIsNone(parse_date("备注 2023 年 说明"))
        self.assertIsNone(parse_date("2023年12月31日(已审)"))


class RowChoice(unittest.TestCase):
    def test_a_row_of_year_starts_loses_to_a_row_of_period_ends_even_with_more_cells(self):
        df = tab([dt(2023, 12, 31), dt(2024, 12, 31), dt(2025, 12, 31), None, None, None], helper=YEAR_STARTS)
        self.assertEqual(date_row_index(df, 2, lambda v: parse_date(v, allow_bare_number=False), max_distance=2), 3)

    def test_the_order_of_the_rows_does_not_decide(self):
        df = tab(YEAR_STARTS, helper=[dt(2023, 12, 31), dt(2024, 12, 31), dt(2025, 12, 31), None, None, None])
        self.assertEqual(date_row_index(df, 2, lambda v: parse_date(v, allow_bare_number=False), max_distance=2), 1)

    def test_a_sheet_whose_only_dates_are_year_starts_still_gets_its_row(self):
        df = tab(YEAR_STARTS)
        self.assertEqual(date_row_index(df, 2, lambda v: parse_date(v, allow_bare_number=False), max_distance=2), 3)

    def test_a_label_detector_that_returns_iso_strings_is_judged_the_same_way(self):
        from fdd_utils.workbook.inspector import _date_row_index
        df = tab([dt(2023, 12, 31), dt(2024, 12, 31), dt(2025, 12, 31), None, None, None], helper=YEAR_STARTS)
        self.assertEqual(_date_row_index(df, 2), 3)

    def test_no_dates_at_all(self):
        df = pd.DataFrame([["a", "b"], ["c", "d"], ["e", "f"]])
        self.assertIsNone(date_row_index(df, 1, lambda v: parse_date(v, allow_bare_number=False)))


class WholeTab(unittest.TestCase):
    """The extractor, on the shape that went wrong."""

    def columns(self, df):
        out = normalize_financial_schedule("x.xlsx", "AR", sheet_df=df)
        return [(c["stage"], c["date"]) for c in out["columns"]]

    def test_line_broken_dates_under_a_year_helper_row_give_the_period_ends(self):
        got = self.columns(tab(NEWLINE_DATES, helper=YEAR_STARTS))
        self.assertEqual([d for _s, d in got], ["2023-12-31", "2024-12-31", "2025-12-31", "2026-01-31", "2023-12-31", "2024-12-31"])
        self.assertEqual(got[0][0], "Mgt acc")
        self.assertEqual(got[4][0], "Audit adjustment")

    def test_spaced_dates_without_a_helper_row_are_read(self):
        spaced = [d.replace("年\n", " 年 ").replace("月", " 月 ").replace("日", " 日") for d in NEWLINE_DATES]
        self.assertEqual([d for _s, d in self.columns(tab(spaced))][:4], ["2023-12-31", "2024-12-31", "2025-12-31", "2026-01-31"])

    def test_dates_already_clean_are_unchanged(self):
        clean = ["2023-12-31", "2024-12-31", "2025-12-31", "2026-01-31", "2023-12-31", "2024-12-31"]
        self.assertEqual([d for _s, d in self.columns(tab(clean))], clean)


if __name__ == "__main__":
    unittest.main(verbosity=1)

"""The code map an agent looks capabilities up in -- offline, standard library only.

    python ad-hoc/workbench/test_codemap.py

A synthetic tree checks what the index records and what it refuses to guess; one smoke test builds the real repo's map.
"""
from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parents[1]))

from fdd_utils.agent import codemap as C  # noqa: E402
from fdd_utils.agent import tools as T  # noqa: E402

MEASURE = '''"""Measuring text for slides."""


def chars_per_line(width_pt, size_pt):
    """How many characters of this size fit on one line of this width."""
    # measured against real PowerPoint: do not round before dividing, the error compounds over a paragraph
    return int(width_pt / (size_pt * 0.5))


def paragraph_lines(text, width_pt, size_pt):
    """Lines a paragraph takes."""
    per = chars_per_line(width_pt, size_pt)
    return -(-len(text) // per)


def unused_helper(x):
    return x


class Splitter:
    def cut(self, text):
        """Cut text at a sentence boundary."""
        return text.split("。")


def shared_name():
    return 1
'''
UI = '''import streamlit as st


def show_table(df):
    """Show a table in the page."""
    st.write(df)


def shared_name():
    return 2
'''
TEST = '''from pkg.measure import chars_per_line, paragraph_lines


def test_it():
    assert chars_per_line(100, 10) == 20
    assert paragraph_lines("x" * 41, 100, 10) == 3
'''


def tree():
    root = Path(tempfile.mkdtemp(prefix="codemap_"))
    (root / "pkg").mkdir()
    (root / "pkg" / "measure.py").write_text(MEASURE, encoding="utf-8")
    (root / "pkg" / "ui.py").write_text(UI, encoding="utf-8")
    (root / "pkg" / "broken.py").write_text("def oops(:\n", encoding="utf-8")
    (root / "pkg" / "test_hidden.py").write_text("def chars_per_line(): pass\n", encoding="utf-8")
    (root / "logs").mkdir()
    (root / "logs" / "skipped.py").write_text("def should_not_appear(): pass\n", encoding="utf-8")
    (root / "ad-hoc").mkdir()
    (root / "ad-hoc" / "test_measure.py").write_text(TEST, encoding="utf-8")
    return root


class Index(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.root = tree()
        cls.map = C.CodeMap(cls.root)

    def test_functions_methods_and_classes_are_indexed_and_skipped_places_are_not(self):
        names = set(self.map.symbols)
        self.assertIn("pkg.measure.chars_per_line", names)
        self.assertIn("pkg.measure.Splitter.cut", names)
        self.assertIn("pkg.measure.Splitter", names)
        self.assertNotIn("pkg.broken.oops", names)                         # a file that does not parse is skipped, not fatal
        self.assertFalse([n for n in names if "should_not_appear" in n or "test_hidden" in n])

    def test_a_found_capability_carries_its_trust_evidence(self):
        top = self.map.find("characters per line width")[0]
        self.assertEqual(top["name"], "pkg.measure.chars_per_line")
        self.assertEqual(top["markers"], 1)                                # the "measured against real PowerPoint" comment
        self.assertEqual(top["tested_by"], ["ad-hoc/test_measure.py"])
        self.assertEqual(top["callers"], 1)
        self.assertTrue(top["reuse"].startswith("likely usable as it is"), top["reuse"])
        self.assertNotIn("no test names it", top["reuse"])                  # it has one
        self.assertIn("(width_pt, size_pt)", top["signature"].replace(" ", "").replace("(width_pt,size_pt)", "(width_pt, size_pt)"))

    def test_code_that_touches_the_page_is_flagged_to_wrap_or_rewrite(self):
        hit = self.map.find("show a table in the page")[0]
        self.assertEqual(hit["name"], "pkg.ui.show_table")
        self.assertIn("streamlit", hit["signals"])
        self.assertTrue(hit["reuse"].startswith("wrap or rewrite"), hit["reuse"])

    def test_code_nobody_calls_and_no_test_mentions_is_called_unproven(self):
        sym = self.map.symbols["pkg.measure.unused_helper"]
        self.assertTrue(sym.reuse().startswith("unproven"), sym.reuse())

    def test_an_untested_method_is_told_to_be_called_through_its_class_and_not_to_be_trusted_on_tests(self):
        hint = self.map.symbols["pkg.measure.Splitter.cut"].reuse()
        self.assertIn("needs an instance of its class", hint)
        self.assertIn("no test names it", hint)

    def test_many_warning_comments_are_called_out_in_the_hint(self):
        sym = self.map.symbols["pkg.measure.paragraph_lines"]
        sym.markers = ["a", "b", "c"]
        try:
            self.assertIn("carries 3 warning/measurement comments", sym.reuse())
        finally:
            sym.markers = []

    def test_a_short_generic_name_is_not_credited_with_tests_it_may_not_have(self):
        self.assertEqual(self.map.symbols["pkg.measure.Splitter.cut"].tests, [])      # "cut" is too generic to attribute by name

    def test_a_name_defined_twice_is_marked_ambiguous_and_explain_asks_for_the_full_name(self):
        self.assertTrue(self.map.symbols["pkg.ui.shared_name"].ambiguous)
        out = self.map.explain("shared_name")
        self.assertFalse(out["found"])
        self.assertEqual(len(out["ambiguous"]), 2)

    def test_explain_gives_callers_callees_and_bounded_source_only_when_asked(self):
        plain = self.map.explain("pkg.measure.paragraph_lines")
        self.assertTrue(plain["found"])
        self.assertEqual(plain["calls"], ["pkg.measure.chars_per_line"])
        self.assertNotIn("source", plain)
        with_source = self.map.explain("paragraph_lines", source_lines=2)
        self.assertEqual(len(with_source["source"].splitlines()), 2)
        self.assertTrue(with_source["source_truncated"])
        self.assertIn("measured against real PowerPoint", " ".join(self.map.explain("chars_per_line")["markers_text"]))

    def test_an_unknown_name_says_so_and_a_query_of_only_stop_words_finds_nothing(self):
        self.assertFalse(self.map.explain("no_such_thing")["found"])
        self.assertEqual(self.map.find("the of a"), [])

    def test_a_chinese_only_query_finds_nothing_in_english_documented_code_and_the_tool_says_to_use_english(self):
        self.assertEqual(self.map.find("字数 折行"), [])
        self.assertIn("English", T.find_capability(None, "字数")["limits"])

    def test_the_answer_stays_small(self):
        import json
        self.assertLess(len(json.dumps(self.map.find("characters per line width", 8), ensure_ascii=False)), 4000)


class Tools(unittest.TestCase):
    def test_the_three_lookups_are_registered_and_read_no_evidence(self):
        for name in ("find_capability", "explain_capability", "code_overview"):
            self.assertIn(name, T.TOOLS)
            self.assertFalse(T.TOOLS[name].reads_evidence)


class RealRepo(unittest.TestCase):
    def test_the_real_map_finds_the_text_measurement_code_that_was_tuned_by_hand(self):
        m = C.get_map()
        self.assertGreater(len(m.symbols), 500)
        names = " ".join(h["name"] for h in m.find("count characters per line wrap text width", 8))
        self.assertTrue("text_metrics" in names or "gen_measurement" in names, names)
        self.assertNotIn("for_test", " ".join(m.modules))


if __name__ == "__main__":
    unittest.main(verbosity=2)

"""slim_compare.py: the old tree is built from a tag zip without touching anything else, and the end of a check is read right.

    python ad-hoc/workbench/test_slim_compare.py

Does not fetch anything and does not run the old code; the real run is `slim_compare.py prepare` then `run`.
"""
from __future__ import annotations

import argparse
import contextlib
import io
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

import slim_compare as sc  # noqa: E402


def make_zip(path: Path, names) -> None:
    with zipfile.ZipFile(path, "w") as z:
        for name in names:
            z.writestr(name, "x = 1\n")


class Unpack(unittest.TestCase):
    def test_the_top_folder_of_a_github_archive_is_dropped(self):
        with tempfile.TemporaryDirectory() as folder:
            zip_path, dest = Path(folder) / "t.zip", Path(folder) / "old"
            make_zip(zip_path, ["repo-pre-slim/", "repo-pre-slim/fdd_utils/ai/pipeline.py", "repo-pre-slim/ask_run.py", "repo-pre-slim/ad-hoc/中文.py"])
            self.assertEqual(sc.unpack_tag_zip(zip_path, dest), 3)
            self.assertTrue((dest / "fdd_utils" / "ai" / "pipeline.py").exists())
            self.assertTrue((dest / "ad-hoc" / "中文.py").exists())

    def test_a_path_that_climbs_out_of_the_folder_is_refused(self):
        with tempfile.TemporaryDirectory() as folder:
            zip_path, dest = Path(folder) / "t.zip", Path(folder) / "old"
            make_zip(zip_path, ["repo/../../outside.py"])
            with self.assertRaises(ValueError):
                sc.unpack_tag_zip(zip_path, dest)
            self.assertFalse((Path(folder) / "outside.py").exists())


class Material(unittest.TestCase):
    def test_config_template_sessions_and_the_tool_are_copied_and_a_missing_one_is_named(self):
        with tempfile.TemporaryDirectory() as folder:
            root, old = Path(folder) / "repo", Path(folder) / "old"
            for rel, text in (("fdd_utils/config.yml", "a: 1\n"), ("fdd_utils/output/agent_trial/x/a.agent-session.json", "{}"),
                              ("fdd_utils/output/agent_trial/y.pptx", "deck"), ("ad-hoc/workbench/baseline_snapshot.py", "# tool\n")):
                (root / rel).parent.mkdir(parents=True, exist_ok=True)
                (root / rel).write_text(text, encoding="utf-8")
            (old / "ad-hoc" / "workbench").mkdir(parents=True)
            got = sc.copy_local_material(root, old)
            self.assertEqual(got["copied"], ["fdd_utils/config.yml"])
            self.assertEqual(got["missing"], ["fdd_utils/template.pptx"])
            self.assertEqual(got["sessions"], 1)
            self.assertTrue((old / "fdd_utils/output/agent_trial/x/a.agent-session.json").exists())
            self.assertFalse((old / "fdd_utils/output/agent_trial/y.pptx").exists(), "only the session files a snapshot reads")
            self.assertEqual((old / "ad-hoc/workbench/baseline_snapshot.py").read_text(encoding="utf-8"), "# tool\n")


class Prepare(unittest.TestCase):
    def test_a_folder_this_tool_did_not_make_is_left_alone(self):
        with tempfile.TemporaryDirectory() as folder:
            old = Path(folder) / "mine"
            old.mkdir()
            (old / "notes.txt").write_text("keep", encoding="utf-8")
            args = argparse.Namespace(old=str(old), zip="", again=True, tag=sc.TAG)
            with contextlib.redirect_stdout(io.StringIO()):
                self.assertEqual(sc.prepare(args), 2)
            self.assertEqual((old / "notes.txt").read_text(encoding="utf-8"), "keep")


class Trees(unittest.TestCase):
    def test_what_differs_between_two_trees_ignores_line_endings_machine_files_and_outputs(self):
        with tempfile.TemporaryDirectory() as folder:
            old, new = Path(folder) / "old", Path(folder) / "new"
            for root, text in ((old, b"a = 1\n"), (new, b"a = 1\r\n")):
                for rel, body in (("fdd_utils/a.py", text), ("fdd_utils/config.yml", b"key: " + root.name.encode()),
                                  ("fdd_utils/output/x/run.py", root.name.encode()), ("same.py", b"s\n")):
                    (root / rel).parent.mkdir(parents=True, exist_ok=True)
                    (root / rel).write_bytes(body)
            self.assertEqual(sc.files_that_differ(old, new), [])
            (new / "fdd_utils/a.py").write_bytes(b"a = 2\n")
            (new / "only_new.py").write_bytes(b"x\n")
            self.assertEqual(sc.files_that_differ(old, new), ["fdd_utils/a.py", "only_new.py"])

    def test_only_the_pre_slim_tag_is_known_by_two_files_any_other_tag_is_taken_as_it_came(self):
        with tempfile.TemporaryDirectory() as folder:
            old = Path(folder)
            self.assertEqual(len(sc.tree_problems(old, sc.TAG)), 1, "no pipeline.py: not the old code")
            self.assertEqual(sc.tree_problems(old, "ship-20261120"), [])
            (old / "fdd_utils/agent").mkdir(parents=True)
            (old / "fdd_utils/ai").mkdir(parents=True)
            (old / "fdd_utils/agent/book.py").write_text("x\n", encoding="utf-8")
            (old / "fdd_utils/ai/pipeline.py").write_text("x\n", encoding="utf-8")
            self.assertEqual(len(sc.tree_problems(old, sc.TAG)), 1, "book.py is there: not the old code")
            self.assertEqual(sc.tree_problems(old, "ship-20261120"), [], "a later tag has book.py and that is right")


class Verdict(unittest.TestCase):
    def test_identical_names_both_fingerprints_and_counts_the_books(self):
        v = sc.verdict_of([" book A ...", "snapshot code 3986ffef323b (2026-11-20 10:00:00)  now c64cb648b6fa", "IDENTICAL",
                           "  A.xlsx   contracts  24 d79b31aa  sentences  1502 b4aa4210  paid sessions  0/   0 44136fa3  no deck taken",
                           "  B.xlsx   not in the comparison: RuntimeError: x"])
        self.assertEqual((v["old"], v["new"], v["identical"], v["compared"], v["left_out"]), ("3986ffef323b", "c64cb648b6fa", True, 1, 1))

    def test_differences_are_kept_and_never_read_as_identical(self):
        v = sc.verdict_of(["snapshot code aaa (x)  now bbb", "2 DIFFERENCE LINE(S)", "  A.xlsx: CONTRACTS differ in 1 of 24", "    k: was 1  now 2"])
        self.assertFalse(v["identical"])
        self.assertEqual(v["differences"], 2)
        self.assertEqual(len(v["diff_lines"]), 2)


if __name__ == "__main__":
    unittest.main()

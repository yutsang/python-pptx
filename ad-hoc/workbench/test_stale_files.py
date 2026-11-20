"""stale_files.py finds what a folder holds that main does not, and main_files.txt is main's own list.

    python ad-hoc/workbench/test_stale_files.py
"""
from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

import stale_files as sf  # noqa: E402


def touch(root: Path, rel: str) -> None:
    path = root / rel
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("x = 1\n", encoding="utf-8")


class Folder(unittest.TestCase):
    def test_a_file_main_does_not_list_is_found_and_the_folders_that_are_not_code_are_left_alone(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder) / "repo"
            for rel in ("fdd_utils/keep.py", "fdd_utils/ai/pipeline.py", "fdd_utils/config.yml", "ad-hoc/probe_old.py",
                        "output/run/helper.py", "databooks/macro.py", ".venv/lib/x.py", "fdd_utils/__pycache__/a.py"):
                touch(root, rel)
            (root / "venv2").mkdir()
            (root / "venv2" / "pyvenv.cfg").write_text("home = x\n", encoding="utf-8")
            touch(root, "venv2/lib/y.py")
            extra, missing = sf.compare(root, {"fdd_utils/keep.py", "fdd_utils/needed.py"})
            self.assertEqual(extra, ["ad-hoc/probe_old.py", "fdd_utils/ai/pipeline.py"])     # config.yml is local; output, databooks and venvs are not code of main
            self.assertEqual(missing, ["fdd_utils/needed.py"])

    def test_moving_aside_keeps_the_sub_folders_and_deletes_nothing(self):
        with tempfile.TemporaryDirectory() as folder:
            root, dest = Path(folder) / "repo", Path(folder) / "aside"
            touch(root, "fdd_utils/ai/pipeline.py")
            self.assertEqual(sf.move_aside(root, ["fdd_utils/ai/pipeline.py"], dest), ["fdd_utils/ai/pipeline.py"])
            self.assertFalse((root / "fdd_utils/ai/pipeline.py").exists())
            self.assertEqual((dest / "fdd_utils/ai/pipeline.py").read_text(encoding="utf-8"), "x = 1\n")

    def test_restoring_puts_back_what_is_under_a_path_and_never_overwrites(self):
        with tempfile.TemporaryDirectory() as folder:
            root, aside = Path(folder) / "repo", Path(folder) / "repo_not_in_main_20261120"
            for rel in ("fdd_rewrite/fdd/a.py", "fdd_rewrite/tests/t.py", "ad-hoc/workbench/test_study_x.py", "fdd_utils/ai/pipeline.py"):
                touch(aside, rel)
            touch(root, "fdd_rewrite/tests/t.py")                          # a file already stands where one would come back to
            back, kept = sf.restore(root, aside, ["fdd_rewrite", "ad-hoc\\workbench\\test_study_x.py"])
            self.assertEqual(sorted(back), ["ad-hoc/workbench/test_study_x.py", "fdd_rewrite/fdd/a.py"])
            self.assertEqual(kept, ["fdd_rewrite/tests/t.py"])
            self.assertTrue((aside / "fdd_utils/ai/pipeline.py").exists(), "what was not asked for stays aside")
            self.assertTrue((aside / "fdd_rewrite/tests/t.py").exists())
            self.assertEqual(sf.aside_folder(root), aside)


class Manifest(unittest.TestCase):
    def test_main_files_txt_is_what_git_tracks(self):
        tracked = sf.git_files(sf.ROOT)
        if not tracked or not (sf.ROOT / ".git").exists():
            self.skipTest("no git checkout here")
        listed = sf.manifest_paths(sf.MANIFEST.read_text(encoding="utf-8"))
        self.assertEqual(sorted(set(tracked) - listed), [], "tracked but not in main_files.txt: run `python ad-hoc/workbench/stale_files.py --write`")
        self.assertEqual(sorted(listed - set(tracked)), [], "in main_files.txt but not tracked: run `python ad-hoc/workbench/stale_files.py --write`")


if __name__ == "__main__":
    unittest.main()

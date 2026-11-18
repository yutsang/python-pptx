"""Every import in the repo resolves (import_audit.py), and the audit itself can fail.

    python ad-hoc/workbench/test_import_audit.py
"""
from __future__ import annotations

import sys
import unittest
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parents[1]))

import import_audit as ia  # noqa: E402


class Audit(unittest.TestCase):
    def test_a_missing_name_and_a_missing_module_are_reported(self):
        cache = {}
        self.assertIn("no nonexistent_name_xyz", ia.check("fdd_utils.ai.claims", ["verify_claim", "nonexistent_name_xyz"], cache))
        self.assertIn("ModuleNotFoundError", ia.check("fdd_utils.ai.nothing_here", ["x"], cache))
        self.assertIsNone(ia.check("fdd_utils.ai.claims", ["verify_claim"], cache))
        self.assertIsNone(ia.check("fdd_utils.agent", ["controller"], cache), "a module inside a package is a name that resolves")

    def test_every_import_in_the_repo_resolves(self):
        problems = ia.audit(include_scripts=True)
        self.assertEqual(problems, {}, "\n".join(f"{f}: {p}" for f, ps in problems.items() for p in ps))


if __name__ == "__main__":
    unittest.main()

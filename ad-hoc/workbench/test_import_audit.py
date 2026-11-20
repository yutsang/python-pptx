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

    def test_an_import_that_fails_is_named_with_the_file_its_statement_is_in(self):
        import tempfile
        with tempfile.TemporaryDirectory() as folder:
            (Path(folder) / "audit_probe_bad.py").write_text("from fdd_utils.financial_common import a_name_that_is_not_there\n", encoding="utf-8")
            (Path(folder) / "audit_probe_good.py").write_text("import json\n", encoding="utf-8")
            sys.path.insert(0, folder)
            try:
                message = ia.import_failure("audit_probe_bad")
                self.assertIn("a_name_that_is_not_there", message)
                self.assertIn("audit_probe_bad.py line 1", message)
                self.assertIsNone(ia.import_failure("audit_probe_good"))
            finally:
                sys.path.remove(folder)
                sys.modules.pop("audit_probe_bad", None)
                sys.modules.pop("audit_probe_good", None)

    def test_every_import_in_the_repo_resolves(self):
        problems = ia.audit(include_scripts=True)
        self.assertEqual(problems, {}, "\n".join(f"{f}: {p}" for f, ps in problems.items() for p in ps))

    def test_every_method_a_class_calls_on_itself_exists(self):
        problems = ia.method_audit()
        self.assertEqual(problems, {}, "\n".join(f"{f}: {p}" for f, ps in problems.items() for p in ps))

    def test_every_root_script_the_package_imports_is_in_the_deploy_zip(self):
        # the server runs fdd_app.py and fdd_utils/ only (make_deploy_package.py); a root script the package imports at run time has to ship too
        import make_deploy_package as pack
        needed = set()
        for path in ia.py_files():
            if not path.relative_to(ia.ROOT).as_posix().startswith("fdd_utils/"):
                continue
            for base, _names, _line in ia.imports_of(path):
                if "." not in base and (ia.ROOT / f"{base}.py").exists():
                    needed.add(f"{base}.py")
        self.assertEqual(sorted(needed - set(pack.ROOT_FILES)), [], "imported by fdd_utils but not in make_deploy_package.ROOT_FILES")
        for name in pack.ROOT_FILES:
            self.assertTrue((ia.ROOT / name).exists(), name)


if __name__ == "__main__":
    unittest.main()

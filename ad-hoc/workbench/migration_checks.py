"""One command for every offline acceptance suite of the agent migration.

    python ad-hoc/workbench/migration_checks.py [--json out.json] [--only NAME ...] [--list]

Zero tokens. Every suite is a unittest module under ad-hoc/workbench; this runs
them in one process with the provider blocked (run_identity.offline_provider_block),
so a suite that reaches for a model fails the run instead of quietly spending.

Exit status is 1 when any test fails or errors, when a listed suite is missing or
runs no tests, or when any provider call was attempted. The JSON report carries
the code revision and library versions so a result can be tied to the code that
produced it; it never contains workbook content.

    --selfcheck   runs a deliberately failing suite through the same machinery
                  and exits 0 only if the runner reports it as a failure (a gate
                  that cannot fail is not a gate)
"""
from __future__ import annotations

import argparse
import importlib.util
import json
import subprocess
import os
import sys
import time
import unittest
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
sys.path.insert(0, str(ROOT))

#: Registration order is run order. Add a suite here when its module lands.
SUITES = [
    "test_reconciliation_selection",
    "test_delivery_decision",
    "test_evidence_contract",
    "test_sheet_visibility",
    "test_claims",
    "test_agent_tools",
    "test_agent_controller",
    "test_deck_lineage",
    "test_agent_flows",
    "test_workspace",
    "test_routing",
    "test_baseline_pack",
    "test_text_transfer",
    "test_databook_dump",
    "test_extract_private_settings",
    "test_codemap",
    "test_agent_workflow",
    "test_agent_panel",
    "test_workspace_flows",      # real workbook + archived run named by FDD_FLOW_WORKBOOK / FDD_FLOW_BASE_RUN; skipped without them
]


def _revision() -> str:
    try:
        out = subprocess.run(["git", "-C", str(ROOT), "rev-parse", "--short", "HEAD"],
                             capture_output=True, encoding="utf-8", errors="replace", timeout=5)
        dirty = subprocess.run(["git", "-C", str(ROOT), "status", "--porcelain"],
                               capture_output=True, encoding="utf-8", errors="replace", timeout=5).stdout.strip()
        return out.stdout.strip() + ("+dirty" if dirty else "")
    except Exception:
        return "unknown"


def _load(name: str):
    path = HERE / f"{name}.py"
    if not path.exists():
        return None
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def run_suite(name: str, module=None) -> dict:
    module = module or _load(name)
    if module is None:
        return {"suite": name, "status": "MISSING", "tests": 0, "failures": [], "errors": [], "seconds": 0.0}
    suite = unittest.defaultTestLoader.loadTestsFromModule(module)
    # utf-8 on purpose: the default is the Windows code page (cp1252), which cannot hold the Chinese in a failure message and would crash the runner
    stream = open(os.devnull, "w", encoding="utf-8", errors="replace")
    started = time.perf_counter()
    result = unittest.TextTestRunner(stream=stream, verbosity=0).run(suite)
    stream.close()
    failed = [t.id() for t, _ in result.failures] + [t.id() for t, _ in result.errors]
    return {
        "suite": name,
        "status": "PASS" if result.wasSuccessful() and result.testsRun else ("NO_TESTS" if not result.testsRun else "FAIL"),
        "tests": result.testsRun,
        "skipped": len(result.skipped),
        "failures": [{"test": t.id(), "message": m.strip().splitlines()[-1][:300]} for t, m in result.failures],
        "errors": [{"test": t.id(), "message": m.strip().splitlines()[-1][:300]} for t, m in result.errors],
        "failed_ids": failed,
        "seconds": round(time.perf_counter() - started, 2),
    }


def versions() -> dict:
    out = {"python": sys.version.split()[0]}
    for name in ("pandas", "openpyxl", "pptx", "yaml"):
        try:
            module = __import__(name)
            out[name] = getattr(module, "__version__", "?")
        except Exception:
            out[name] = None
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--json", default=None, metavar="PATH", help="write the machine-readable report here")
    ap.add_argument("--only", nargs="*", default=None, help="run just these suites")
    ap.add_argument("--list", action="store_true")
    ap.add_argument("--selfcheck", action="store_true")
    ap.add_argument("--fingerprint", nargs="*", default=[], metavar="FILE",
                    help="record the sha256 of these workbooks in the report (the sources a real-data suite ran against)")
    args = ap.parse_args()

    if args.list:
        for name in SUITES:
            print(name, "(present)" if (HERE / f"{name}.py").exists() else "(not yet written)")
        return 0

    from fdd_utils import run_identity

    if args.selfcheck:
        class _Broken(unittest.TestCase):
            def test_must_fail(self):
                self.assertTrue(False, "deliberate")
        module = type(sys)("broken_suite")
        module._Broken = _Broken
        with run_identity.offline_provider_block():
            report = run_suite("broken_suite", module)
        caught = report["status"] == "FAIL" and report["tests"] == 1
        print("selfcheck:", "the runner reports a failing suite as FAIL" if caught else "RUNNER DID NOT CATCH A FAILURE")
        return 0 if caught else 1

    wanted = args.only or SUITES
    reports = []
    with run_identity.offline_provider_block() as attempts:
        for name in wanted:
            reports.append(run_suite(name))
    ok = all(r["status"] == "PASS" for r in reports) and not attempts
    from fdd_utils.run_identity import file_sha256
    summary = {
        "ok": ok,
        "revision": _revision(),
        "versions": versions(),
        "sources": {Path(f).name: file_sha256(f) for f in args.fingerprint},
        "limits": [
            "no provider is configured on the machine that produced this report: model behaviour is not covered",
            "policies in the controller and flow suites are scripted; they test the harness, not a model",
            "python-pptx only: real PowerPoint (custom XML part, tags, line breaks) was not run",
        ],
        "provider_call_attempts": list(attempts),
        "suites": reports,
        "total_tests": sum(r["tests"] for r in reports),
    }
    for r in reports:
        mark = {"PASS": "ok  ", "MISSING": "miss", "NO_TESTS": "none", "FAIL": "FAIL"}[r["status"]]
        skipped = f"  ({r['skipped']} SKIPPED: needs material this machine run did not name)" if r.get("skipped") else ""
        print(f"{mark} {r['suite']:34s} {r['tests']:3d} test(s)  {r['seconds']:6.2f}s{skipped}")
        for item in r["failures"] + r["errors"]:
            print(f"       x {item['test']}: {item['message']}")
    print(f"provider calls attempted: {len(attempts)}")
    print(f"{'ALL GREEN' if ok else 'NOT GREEN'}: {summary['total_tests']} test(s), revision {summary['revision']}")
    if args.json:
        Path(args.json).write_text(json.dumps(summary, ensure_ascii=False, indent=1), encoding="utf-8")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())

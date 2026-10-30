"""Does inspect_databook.py --from-run fail when it should? Zero tokens, ~2 minutes.

One canned-provider run (stub_inspect_run.py) supplies a real run folder; each
case then re-exports it offline with something broken on purpose and checks
the exit code and the reason in <deck>.reexport.json:

    baseline            the run as written                      -> PASS, exit 0
    wrong workbook      another databook against this run        -> FAIL identity
    overflow            one account's text made 400x longer      -> FAIL layout overflow
    unknown account     commentary for an account not in the book -> FAIL before export
    dropped account     the payload loses an account's bullet     -> FAIL never reached the deck
    no summary          section_summaries.json deleted           -> FAIL; PASS with --summary-fallback
    model call          export path tries AIClient.get_response  -> FAIL, call blocked

    python ad-hoc/workbench/test_offline_reexport.py ["Project Gold Kunshan.databook.xlsx" "databook-Foshan Wanyuan_rebuilt.xlsx"]

The copies of the run folder live in a temp directory; fdd_utils/logs gains
one stub run folder, as any stub run leaves.
"""
from __future__ import annotations

import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import yaml

ROOT = Path(__file__).resolve().parents[2]
PY = sys.executable

#: Runs inspect_databook.main() with one patch applied first.
_DRIVER = """
import sys, types
sys.path.insert(0, {root!r})
import inspect_databook as I
{patch}
sys.argv = {argv!r}
raise SystemExit(I.main())
"""

#: The export path tries a model call before exporting.
_MODEL_CALL = """
_orig = I.export_and_inspect_pptx
def _wrapped(*a, **k):
    from fdd_utils.ai.client import AIClient
    try:
        AIClient.get_response(types.SimpleNamespace(agent_name="injected_probe"), "hello")
    except Exception:
        pass
    return _orig(*a, **k)
I.export_and_inspect_pptx = _wrapped
"""

#: The payload builder loses the first balance-sheet account.
_DROP_PAYLOAD = """
import fdd_utils.pptx as P
_orig = P.build_pptx_structured_payloads
def _dropping(*a, **k):
    out = _orig(*a, **k)
    out["BS"] = list(out.get("BS") or [])[1:]
    return out
P.build_pptx_structured_payloads = _dropping
"""


def _reexport(databook: str, run: str, out_dir: str, *extra: str, patch: str = ""):
    argv = ["inspect_databook.py", databook, "--from-run", run, "--pptx-out", out_dir, *extra]
    cmd = ([PY, "-c", _DRIVER.format(root=str(ROOT), argv=argv, patch=patch)] if patch else [PY, *argv])
    proc = subprocess.run(cmd, cwd=str(ROOT), capture_output=True, text=True, encoding="utf-8", errors="replace")
    reports = sorted(Path(out_dir).glob("*.reexport.json"), key=lambda p: p.stat().st_mtime)
    report = json.loads(reports[-1].read_text(encoding="utf-8")) if reports else {}
    return proc.returncode, report, proc.stdout[-2000:]


def main() -> int:
    books = sys.argv[1:] or ["Project Gold Kunshan.databook.xlsx", "databook-Foshan Wanyuan_rebuilt.xlsx"]
    databook, other = books[0], books[1]
    logs = ROOT / "fdd_utils" / "logs"
    before = {p.name for p in logs.glob("run_*")}
    stub = subprocess.run([PY, "ad-hoc/workbench/stub_inspect_run.py", databook, "--export"], cwd=str(ROOT),
                          capture_output=True, text=True, encoding="utf-8", errors="replace")
    new = sorted((p for p in logs.glob("run_*") if p.name not in before), key=lambda p: p.name)
    if stub.returncode != 0 or not new:
        print(stub.stdout[-3000:], stub.stderr[-3000:])
        print("❌ the stub run produced no run folder")
        return 1
    run_folder = new[-1]
    print(f"stub run: {run_folder}")

    work = Path(tempfile.mkdtemp(prefix="reexport_probe_"))
    out = str(work / "decks")

    def copy(name: str) -> Path:
        target = work / name
        shutil.copytree(run_folder, target)
        return target

    def edit_results(folder: Path, change) -> None:
        path = folder / "results.yml"
        data = yaml.safe_load(path.read_text(encoding="utf-8"))
        change(data)
        path.write_text(yaml.dump(data, allow_unicode=True), encoding="utf-8")

    def first_account(data):
        return next(k for k, v in data.items() if not str(k).startswith("__") and isinstance(v, dict))

    cases = []

    code, rep, _ = _reexport(databook, str(run_folder), out)
    cases.append(("baseline", code == 0 and rep.get("verdict") == "PASS", code, rep))

    code, rep, _ = _reexport(other, str(run_folder), out)
    cases.append(("wrong workbook", code == 1 and any("identity differs" in p or "account" in p
                                                       for p in rep.get("problems", [])), code, rep))

    long_run = copy("overflow")

    def _inflate(data):
        key = first_account(data)
        data[key]["final"] = (str(data[key].get("final") or "") + "。") * 400
    edit_results(long_run, _inflate)
    code, rep, _ = _reexport(databook, str(long_run), out)
    cases.append(("overflow", code == 1 and any(p.startswith("layout overflow") for p in rep.get("problems", [])),
                  code, rep))

    extra_run = copy("unknown_account")
    edit_results(extra_run, lambda d: d.__setitem__("不存在的科目", {"final": "截至2025年12月31日，余额为1.0万元。"}))
    code, rep, _ = _reexport(databook, str(extra_run), out)
    cases.append(("unknown account", code == 1 and any("not in this workbook" in p
                                                        for p in rep.get("problems", [])), code, rep))

    code, rep, _ = _reexport(databook, str(run_folder), out, patch=_DROP_PAYLOAD)
    cases.append(("dropped account", code == 1 and any("never reached the deck" in p
                                                        for p in rep.get("problems", [])), code, rep))

    bare_run = copy("no_summary")
    (bare_run / "section_summaries.json").unlink()
    code, rep, _ = _reexport(databook, str(bare_run), out)
    cases.append(("no summary", code == 1 and any("executive summary" in p for p in rep.get("problems", [])),
                  code, rep))
    code, rep, _ = _reexport(databook, str(bare_run), out, "--summary-fallback")
    cases.append(("no summary + --summary-fallback", code == 0 and rep.get("verdict") == "PASS", code, rep))

    code, rep, _ = _reexport(databook, str(run_folder), out, patch=_MODEL_CALL)
    cases.append(("model call", code == 1 and len(rep.get("provider_call_attempts") or []) == 1, code, rep))

    ok = True
    for name, passed, code, rep in cases:
        print(f"{'ok  ' if passed else 'FAIL'}  {name:34s} exit={code} verdict={rep.get('verdict')}")
        for problem in (rep.get("problems") or [])[:2]:
            print(f"        {problem[:150]}")
        ok = ok and passed
    shutil.rmtree(work, ignore_errors=True)
    print("\n" + ("✅ all cases behave" if ok else "❌ cases above did not"))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())

"""Free; no model call. One command to run after a new version has been copied over this folder, and one block to paste back.

    python ad-hoc/workbench/verify_update.py                       the folder is complete and clean, and the gate is green (about 3 minutes)
    python ad-hoc/workbench/verify_update.py --since ship-20261120  and what the agent path does is what it did at that tag (5 to 15 minutes more)
    python ad-hoc/workbench/verify_update.py --since TAG --deck     and the scripted whole deck of every book, slide by slide (hours)

What each step is:
    FOLDER  stale_files.py: .py/.yml files here that main does not have, and files of main that are not here (a copy laid over an
            older folder keeps what main deleted, and a partial copy lacks what it did not carry)
    GATE    migration_checks.py: the whole offline test gate
    SINCE   slim_compare.py: the code of tag TAG is built next to this folder, both snapshot the databooks, and the two are compared
            (contracts, the verdict on every generated sentence, the saved paid sessions; with --deck also the decks). A change that was
            meant to alter what the verifier says shows up here as differences; say so when you paste it back.

The full output of every step goes to fdd_utils/output/verify_<time>.log; what is printed is the part that answers.
"""
from __future__ import annotations

import argparse
import subprocess
import sys
import time
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
sys.path.insert(0, str(HERE))


def run_and_show(cmd: List[str], cwd: Path, show: Callable[[str], bool], log) -> Tuple[int, List[str]]:
    """Run cmd, write every line to `log`, print the lines `show` accepts as they come, and return (exit code, all lines)."""
    import os
    env = dict(os.environ)
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    proc = subprocess.Popen(cmd, cwd=str(cwd), env=env, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding="utf-8",
                            errors="replace", bufsize=1)
    lines: List[str] = []
    assert proc.stdout is not None
    for raw in proc.stdout:
        line = raw.rstrip("\n").split("\r")[-1]
        if not line.strip():
            continue
        lines.append(line)
        log.write(line + "\n")
        log.flush()
        if show(line):
            print(line, flush=True)
    return proc.wait(), lines


def show_gate(line: str) -> bool:
    return line.startswith(("FAIL", "NOT GREEN", "ALL GREEN")) or line.startswith("      x")


def show_compare(line: str) -> bool:
    """A book being read (` name ...`), the result block, and the lines that say what differs."""
    if line.startswith(" ") and not line.startswith("  ") and line.rstrip().endswith("..."):
        return True
    if "==== RESULT ====" in line or " DIFFERENCE LINE(S)" in line:
        return True
    return line.startswith(("code that took", "books compared", "IDENTICAL", "NOT A COMPARISON", "no verdict", "PROBLEM")) or (
        line.startswith("  ") and ("differ" in line or " was " in line))


def gate_verdict(lines: List[str]) -> Tuple[bool, str]:
    last = next((l for l in reversed(lines) if l.startswith(("ALL GREEN", "NOT GREEN"))), "")
    return last.startswith("ALL GREEN"), last or "the gate printed no verdict"


def compare_verdict(lines: List[str]) -> Tuple[bool, str]:
    after = lines[next((i for i, l in enumerate(lines) if "==== RESULT ====" in l), len(lines)):]
    text = " | ".join(l.strip() for l in after[1:6])
    if any(l.startswith("IDENTICAL") for l in after):
        return True, text
    return False, text or "no result block: the comparison did not finish (see the log)"


def folder_verdict() -> Tuple[bool, List[str]]:
    import stale_files as sf
    manifest = sf.manifest_paths(sf.MANIFEST.read_text(encoding="utf-8"))
    extra, missing = sf.compare(ROOT, manifest)
    lines = [f"NOT IN MAIN: {len(extra)}   IN MAIN, NOT HERE: {len(missing)}"]
    lines += [f"    not in main: {p}" for p in extra[:15]] + [f"    missing here: {p}" for p in missing[:15]]
    if len(extra) > 15 or len(missing) > 15:
        lines.append("    ... (python ad-hoc/workbench/stale_files.py lists all)")
    return not extra and not missing, lines


def fingerprints() -> str:
    try:
        sys.path.insert(0, str(ROOT))
        import entity_trial
        from fdd_utils.agent import book
        return f"code {entity_trial.code_fingerprint()['short']}   product {book.code_fingerprint()['short']}"
    except Exception as exc:                                              # noqa: BLE001
        return f"fingerprint unavailable ({type(exc).__name__})"


def main(argv: Optional[List[str]] = None) -> int:
    try:
        sys.stdout.reconfigure(errors="replace")                         # type: ignore[attr-defined]
    except Exception:                                                    # noqa: BLE001
        pass
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--since", default="", metavar="TAG", help="also compare what the agent path does with the code of this tag")
    ap.add_argument("--deck", action="store_true", help="with --since: the scripted whole deck of every book too (hours)")
    ap.add_argument("--dir", default="", help="with --since: the folder of databooks (default: databooks if it exists)")
    ap.add_argument("--old", default="", help="with --since: where the old code is built (default: next to this folder)")
    args = ap.parse_args(argv)
    sys.path.insert(0, str(ROOT))
    out = ROOT / "fdd_utils" / "output"
    out.mkdir(parents=True, exist_ok=True)
    log_path = out / f"verify_{time.strftime('%Y%m%d_%H%M%S')}.log"
    results: Dict[str, Tuple[bool, str]] = {}
    with log_path.open("w", encoding="utf-8") as log:
        print(f"VERIFY UPDATE  {time.strftime('%Y-%m-%d %H:%M')}   {fingerprints()}", flush=True)
        print("\n=== 1  FOLDER: complete, and nothing in it that main does not have", flush=True)
        ok, lines = folder_verdict()
        print("\n".join(lines))
        results["FOLDER"] = (ok, lines[0])
        print("\n=== 2  GATE: the offline tests (about 3 minutes)", flush=True)
        code, lines = run_and_show([sys.executable, str(HERE / "migration_checks.py")], ROOT, show_gate, log)
        results["GATE"] = gate_verdict(lines)
        if args.since:
            slim = [sys.executable, str(HERE / "slim_compare.py")]
            print(f"\n=== 3  SINCE {args.since}: the old code is built next to this folder", flush=True)
            code, lines = run_and_show(slim + ["prepare", "--tag", args.since, "--again"] + (["--old", args.old] if args.old else []), ROOT,
                                       lambda l: l.startswith(("OLD TREE", "  files of tag", "PROBLEM", "OK.", "the download", "put the zip")), log)
            if code != 0:
                results["SINCE"] = (False, f"the old code of {args.since} could not be built (see the log)")
            else:
                print(f"=== 3  SINCE {args.since}: both sides read the databooks (a book per line)", flush=True)
                flags = (["--deck"] if args.deck else []) + (["--dir", args.dir] if args.dir else []) + (["--old", args.old] if args.old else [])
                code, lines = run_and_show(slim + ["run"] + flags, ROOT, show_compare, log)
                results["SINCE"] = compare_verdict(lines)
        print(f"\n==== VERIFY UPDATE RESULT   {fingerprints()}")
        for name, (ok, text) in results.items():
            print(f"{name:7s} {'PASS' if ok else 'FAIL'}   {text}")
        if not args.since:
            print("SINCE   not run (add --since TAG to compare with an earlier version)")
        else:
            import slim_compare
            print(f"the old code stays in {args.old or slim_compare.default_old_tree()} and holds a copy of your config.yml: delete that folder when you are done")
        print(f"full output: {log_path}")
    return 0 if all(ok for ok, _ in results.values()) else 1


if __name__ == "__main__":
    raise SystemExit(main())

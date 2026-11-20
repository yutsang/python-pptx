"""Free; no model call. Is what the agent path does today what it did before the slim (tag pre-slim-20261119)? Two steps.

    python ad-hoc/workbench/slim_compare.py prepare       build the old code in a folder NEXT TO this one (python-pptx_old_tree); nothing here is touched
    python ad-hoc/workbench/slim_compare.py run           snapshot the books with the old code, then check them with the code in this folder

    prepare  [--zip PATH]   the old code comes from the tag's zip on GitHub, or from a zip you name (a download that fails says so)
             [--again]      build it again (only a folder this tool made is ever replaced)
    run      [--dir databooks]   the folder of databooks (default: databooks if it exists, else this folder)
             [--deck]            also the scripted whole deck of every book, slide by slide (about 3 hours for 17 books; the default
                                 reads contracts, the verdict on every generated sentence and the saved paid sessions: minutes)
             [--books TEXT]      only the books whose file name contains TEXT

The old tree lives outside this folder on purpose: a tree inside it would be read by the import audit, the stale-file listing and the
snapshot's own search for saved sessions. It receives copies of what a tree needs and a checkout does not hold (fdd_utils/config.yml,
fdd_utils/template.pptx, every *.agent-session.json under fdd_utils/output) and this folder's baseline_snapshot.py, which runs on the old
code too. What it prints at the end: the fingerprint of the code that took the snapshot (it must be the old one), the fingerprint now,
and IDENTICAL or the difference lines. This proves what baseline_snapshot.py hashes and nothing else: not how a model writes, not
PowerPoint.
"""
from __future__ import annotations

import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
TAG = "pre-slim-20261119"
REPO_URL = "https://github.com/yutsang/python-pptx"
LABEL = "slim-old"
MARKER = ".slim_old_tree"
#: a file only the old code has, and one only the code of today has
ONLY_OLD, ONLY_NEW = "fdd_utils/ai/pipeline.py", "fdd_utils/agent/book.py"
#: git-ignored, per machine, and needed to run: copied from this folder
LOCAL_FILES = ("fdd_utils/config.yml", "fdd_utils/template.pptx")


def default_old_tree(root: Path = ROOT) -> Path:
    return root.parent / f"{root.name}_old_tree"


def download(tag: str, to: Path) -> None:
    import urllib.request
    url = f"{REPO_URL}/archive/refs/tags/{tag}.zip"
    with urllib.request.urlopen(url, timeout=120) as response, to.open("wb") as out:
        shutil.copyfileobj(response, out)


def unpack_tag_zip(zip_path: Path, dest: Path) -> int:
    """The files of a GitHub tag archive (one top-level folder, which is dropped) under dest. Returns the number of files."""
    count = 0
    dest = dest.resolve()
    with zipfile.ZipFile(zip_path) as archive:
        for info in archive.infolist():
            if info.is_dir():
                continue
            parts = info.filename.split("/", 1)
            if len(parts) < 2 or not parts[1]:
                continue
            target = (dest / parts[1]).resolve()
            if dest not in target.parents:
                raise ValueError(f"the zip holds a path outside its folder: {info.filename}")
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(archive.read(info))
            count += 1
    return count


def copy_local_material(root: Path, old: Path) -> Dict[str, object]:
    """What a tree needs and a checkout does not hold. Returns what was copied (and what was not there)."""
    out: Dict[str, object] = {"copied": [], "missing": [], "sessions": 0, "session_bytes": 0}
    for rel in LOCAL_FILES:
        source = root / rel
        if source.exists():
            (old / rel).parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(source, old / rel)
            out["copied"].append(rel)             # type: ignore[union-attr]
        else:
            out["missing"].append(rel)            # type: ignore[union-attr]
    output = root / "fdd_utils" / "output"
    if output.exists():
        for session in output.rglob("*.agent-session.json"):
            target = old / session.relative_to(root)
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(session, target)
            out["sessions"] = int(out["sessions"]) + 1                      # type: ignore[call-overload]
            out["session_bytes"] = int(out["session_bytes"]) + session.stat().st_size   # type: ignore[call-overload]
    shutil.copyfile(root / "ad-hoc" / "workbench" / "baseline_snapshot.py", old / "ad-hoc" / "workbench" / "baseline_snapshot.py")
    return out


def utf8_env() -> Dict[str, str]:
    env = dict(os.environ)
    env["PYTHONUTF8"] = "1"                 # a piped child on Windows would otherwise write the console's code page and fail on a Chinese file name
    env["PYTHONIOENCODING"] = "utf-8"
    return env


def tee(cmd: List[str], cwd: Path) -> Tuple[int, List[str]]:
    """Run cmd, show its output as it comes, and return (exit code, its lines)."""
    lines: List[str] = []
    proc = subprocess.Popen(cmd, cwd=str(cwd), env=utf8_env(), stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                            text=True, encoding="utf-8", errors="replace", bufsize=1)
    assert proc.stdout is not None
    for raw in proc.stdout:
        line = raw.rstrip("\n").split("\r")[-1]          # a progress bar redraws itself with \r: keep what it ended on
        if line.strip():
            lines.append(line)
            print(line, flush=True)
    return proc.wait(), lines


def verdict_of(lines: Iterable[str]) -> Dict[str, object]:
    """What the end of `baseline_snapshot.py check` says: the two fingerprints, IDENTICAL or the difference lines, and how many books."""
    out: Dict[str, object] = {"old": None, "new": None, "identical": False, "differences": 0, "diff_lines": [], "compared": 0, "left_out": 0}
    after_verdict = False
    for line in lines:
        m = re.match(r"snapshot code (\S+) .* now (\S+)\s*$", line)
        if m:
            out["old"], out["new"] = m.group(1), m.group(2)
            continue
        if line.strip() == "IDENTICAL":
            out["identical"], after_verdict = True, True
            continue
        m = re.match(r"(\d+) DIFFERENCE LINE\(S\)", line.strip())
        if m:
            out["differences"], after_verdict = int(m.group(1)), True
            continue
        if after_verdict:
            if " contracts " in line and " sentences " in line:
                out["compared"] = int(out["compared"]) + 1             # type: ignore[call-overload]
            elif "not in the comparison" in line:
                out["left_out"] = int(out["left_out"]) + 1             # type: ignore[call-overload]
            elif line.startswith("  ") and not out["identical"] and len(out["diff_lines"]) < 40:   # type: ignore[arg-type]
                out["diff_lines"].append(line)                          # type: ignore[union-attr]
    return out


def prepare(args: argparse.Namespace) -> int:
    old = Path(args.old).resolve() if args.old else default_old_tree()
    if old.exists() and any(old.iterdir()):
        if not (old / MARKER).exists():
            print(f"{old} exists and was not made by this tool: nothing touched. Name another folder with --old.")
            return 2
        if not args.again:
            print(f"{old} is already built (tag {(old / MARKER).read_text(encoding='utf-8').strip()}). `run` uses it; --again builds it anew.")
            return check_tree(old)
        shutil.rmtree(old)
    old.mkdir(parents=True, exist_ok=True)
    zip_path: Optional[Path] = Path(args.zip) if args.zip else None
    scratch = None
    try:
        if zip_path is None:
            scratch = Path(tempfile.mkdtemp(prefix="slim_tag_"))
            zip_path = scratch / "tag.zip"
            print(f"fetching tag {args.tag} from {REPO_URL} ...", flush=True)
            try:
                download(args.tag, zip_path)
            except Exception as exc:                                              # noqa: BLE001
                print(f"the download failed: {type(exc).__name__}: {str(exc)[:160]}")
                print(f"put the zip of tag {args.tag} anywhere and run again with --zip <that file>")
                shutil.rmtree(old, ignore_errors=True)
                return 2
        if not zipfile.is_zipfile(zip_path):
            print(f"{zip_path} is not a zip file")
            shutil.rmtree(old, ignore_errors=True)
            return 2
        count = unpack_tag_zip(zip_path, old)
        (old / MARKER).write_text(args.tag + "\n", encoding="utf-8")
    finally:
        if scratch is not None:
            shutil.rmtree(scratch, ignore_errors=True)
    material = copy_local_material(ROOT, old)
    print(f"OLD TREE {old}")
    print(f"  files of tag {args.tag}: {count}")
    print(f"  copied from this folder: {', '.join(material['copied']) or 'nothing'}"                       # type: ignore[arg-type]
          + (f"   NOT FOUND here: {', '.join(material['missing'])}" if material["missing"] else ""))      # type: ignore[arg-type]
    print(f"  saved sessions copied: {material['sessions']} ({int(material['session_bytes']) / 1e6:.1f} MB)")   # type: ignore[call-overload]
    print("  ad-hoc/workbench/baseline_snapshot.py: this folder's, which also runs on the old code")
    return check_tree(old)


def check_tree(old: Path) -> int:
    """Is it the old code, and does it start under this Python?"""
    problems = []
    if not (old / ONLY_OLD).exists():
        problems.append(f"{ONLY_OLD} is missing: this is not the old code")
    if (old / ONLY_NEW).exists():
        problems.append(f"{ONLY_NEW} is there: this is not the old code")
    print(f"  {ONLY_OLD} (only the old code has it): {'yes' if (old / ONLY_OLD).exists() else 'NO'}")
    print(f"  {ONLY_NEW} (only the code of today has it): {'yes, WRONG' if (old / ONLY_NEW).exists() else 'no'}")
    probe = ("import sys; sys.path[:0] = ['.', 'ad-hoc/workbench']; import entity_trial, baseline_snapshot; "
             "from fdd_utils.agent.workspace import WorkSession; print('imports ok')")
    done = subprocess.run([sys.executable, "-c", probe], cwd=str(old), env=utf8_env(), capture_output=True, text=True, encoding="utf-8", errors="replace")
    ok = done.returncode == 0 and "imports ok" in done.stdout
    print(f"  the old code starts under this Python: {'yes' if ok else 'NO'}")
    if not ok:
        problems.append("the old code does not start: " + (done.stderr.strip().splitlines() or ["no message"])[-1][:200])
    for problem in problems:
        print(f"PROBLEM {problem}")
    if not problems:
        print("OK. Next: python ad-hoc/workbench/slim_compare.py run")
    return 1 if problems else 0


def run(args: argparse.Namespace) -> int:
    old = Path(args.old).resolve() if args.old else default_old_tree()
    if not (old / MARKER).exists() or not (old / ONLY_OLD).exists():
        print(f"{old} is not built: run `python ad-hoc/workbench/slim_compare.py prepare` first")
        return 2
    books = Path(args.dir).resolve() if args.dir else ((ROOT / "databooks") if (ROOT / "databooks").exists() else ROOT)
    flags = ["--dir", str(books)] + ([] if args.deck else ["--no-deck"]) + (["--books", args.books] if args.books else [])
    script = Path("ad-hoc") / "workbench" / "baseline_snapshot.py"
    print(f"=== 1 of 2: the snapshot, taken with the OLD code ({old}) over {books}", flush=True)
    code, _ = tee([sys.executable, str(old / script), "save", LABEL] + flags, old)
    saved = old / "fdd_utils" / "output" / "baseline" / f"{LABEL}.json"
    if code != 0 or not saved.exists():
        print(f"the old code's snapshot did not finish (exit {code}); the lines above say where")
        return 1
    target = ROOT / "fdd_utils" / "output" / "baseline" / f"{LABEL}.json"
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(saved, target)
    print(f"\n=== 2 of 2: the same books, checked with the code in {ROOT}", flush=True)
    code, lines = tee([sys.executable, str(ROOT / script), "check", LABEL] + flags, ROOT)
    verdict = verdict_of(lines)
    print("\n==== RESULT ====")
    print(f"code that took the snapshot: {verdict['old']}   code now: {verdict['new']}")
    if verdict["identical"]:
        print(f"books compared: {verdict['compared']}   left out (they do not run on their own): {verdict['left_out']}")
    if verdict["old"] and verdict["old"] == verdict["new"]:
        print("NOT A COMPARISON: both fingerprints are the same code. The old tree was not used.")
        return 1
    if verdict["identical"]:
        print("IDENTICAL" + ("" if args.deck else " (contracts, sentence verdicts, saved sessions; decks not taken: add --deck)"))
        return 0
    print(f"{verdict['differences']} DIFFERENCE LINE(S)" if verdict["differences"] else "no verdict was printed (the check did not finish)")
    for line in verdict["diff_lines"]:                                              # type: ignore[union-attr]
        print(line)
    return 1


def main(argv: Optional[List[str]] = None) -> int:
    try:
        sys.stdout.reconfigure(errors="replace")           # type: ignore[attr-defined]
    except Exception:                                      # noqa: BLE001
        pass
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("action", choices=["prepare", "run"])
    ap.add_argument("--zip", default="", help="prepare: a zip of the tag instead of fetching it")
    ap.add_argument("--again", action="store_true", help="prepare: build the old tree anew")
    ap.add_argument("--tag", default=TAG, help=f"prepare: the tag to compare with (default {TAG})")
    ap.add_argument("--old", default="", help="where the old tree is (default: next to this folder)")
    ap.add_argument("--dir", default="", help="run: the folder of databooks")
    ap.add_argument("--deck", action="store_true", help="run: also the scripted whole deck of every book")
    ap.add_argument("--books", default="", help="run: only the books whose file name contains this")
    args = ap.parse_args(argv)
    return prepare(args) if args.action == "prepare" else run(args)


if __name__ == "__main__":
    raise SystemExit(main())

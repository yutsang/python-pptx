"""Free. Which files in this folder are not files of main, and which files of main are not in this folder.

    python ad-hoc/workbench/stale_files.py            report
    python ad-hoc/workbench/stale_files.py --move     move the .py/.yml files main does not have to a folder next to this one (reversible)
    python ad-hoc/workbench/stale_files.py --restore [PATH ...]   put back what --move put aside: everything, or what starts with each PATH
    python ad-hoc/workbench/stale_files.py --write    rewrite main_files.txt from git (only where git is: the maintainer's copy, before a push)

The company folder is not a git checkout: a new version is copied over the old one, and a file that main has since deleted stays where
it was (found 2026-11-20: 30 leftover files, each importing code the slim removed, failed the import audit and told nobody which they
were). main_files.txt is every file git tracks at the time of the push; this compares the folder with it. A file this does not
list is either one main removed or one that was never in it (work done on that machine); either way it is the person's to keep or
move, so --move puts it aside and deletes nothing.
"""
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import time
from pathlib import Path
from typing import Iterable, List, Optional, Set, Tuple

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
MANIFEST = HERE / "main_files.txt"
#: folders that hold a person's own material, output or scratch, not code of main
SKIP_DIRS = {".venv", "__pycache__", ".git", "node_modules", "output", "for_test", "databooks", "pptx_previews", "logs", "docs", "site-packages"}
#: local, git-ignored files that belong in the folder
LOCAL_FILES = {"fdd_utils/config.yml"}
KINDS = (".py", ".yml")


def manifest_paths(text: str) -> Set[str]:
    return {line.strip() for line in text.splitlines() if line.strip() and not line.startswith("#")}


def folder_files(root: Path) -> List[str]:
    """Every .py and .yml file of the folder, as a posix path relative to it, outside the folders that are not code."""
    import os
    found: List[str] = []
    for folder, dirs, files in os.walk(root):
        here = Path(folder)
        dirs[:] = sorted(d for d in dirs if d not in SKIP_DIRS and not (here / d / "pyvenv.cfg").exists())
        for name in sorted(files):
            if name.endswith(KINDS):
                found.append((here / name).relative_to(root).as_posix())
    return found


def compare(root: Path, manifest: Set[str]) -> Tuple[List[str], List[str]]:
    """(files here that main does not list, files main lists that are not here)."""
    extra = [p for p in folder_files(root) if p not in manifest and p not in LOCAL_FILES]
    missing = sorted(p for p in manifest if not (root / p).exists())
    return extra, missing


def move_aside(root: Path, extra: Iterable[str], dest: Path) -> List[str]:
    moved = []
    for rel in extra:
        target = dest / rel
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(root / rel), str(target))
        moved.append(rel)
    return moved


def aside_folder(root: Path) -> Optional[Path]:
    """The newest folder --move made next to this one."""
    found = sorted(root.parent.glob(f"{root.name}_not_in_main_*"))
    return found[-1] if found else None


def restore(root: Path, aside: Path, prefixes: Iterable[str] = ()) -> Tuple[List[str], List[str]]:
    """Move files back from `aside` to the same place under root: all of them, or those at or under one of the paths. A file whose place
    is taken is left where it is. Returns (put back, left because the place was taken)."""
    wanted = [p.replace("\\", "/").strip("/") for p in prefixes if p.strip()]
    back: List[str] = []
    kept: List[str] = []
    for path in sorted(aside.rglob("*")):
        if not path.is_file():
            continue
        rel = path.relative_to(aside).as_posix()
        if wanted and not any(rel == w or rel.startswith(w + "/") for w in wanted):
            continue
        target = root / rel
        if target.exists():
            kept.append(rel)
            continue
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(path), str(target))
        back.append(rel)
    return back, kept


def git_files(root: Path) -> Optional[List[str]]:
    try:
        out = subprocess.run(["git", "-C", str(root), "ls-files"], capture_output=True, text=True, encoding="utf-8", check=True).stdout
    except (OSError, subprocess.CalledProcessError):
        return None
    return sorted(line for line in out.splitlines() if line.strip())


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--move", action="store_true", help="move the files main does not have to a folder next to this one")
    ap.add_argument("--write", action="store_true", help="rewrite main_files.txt from git")
    ap.add_argument("--restore", nargs="*", default=None, metavar="PATH", help="put back what --move put aside (all of it, or what is at or under each PATH)")
    ap.add_argument("--from", dest="aside", default="", help="the folder --move made (default: the newest one next to this folder)")
    args = ap.parse_args(argv)
    if args.restore is not None:
        aside = Path(args.aside) if args.aside else aside_folder(ROOT)
        if aside is None or not aside.exists():
            print("no folder made by --move was found next to this one")
            return 2
        back, kept = restore(ROOT, aside, args.restore)
        print(f"PUT BACK {len(back)} file(s) from {aside}" + (f", only what is at or under: {', '.join(args.restore)}" if args.restore else ""))
        for rel in back[:200]:
            print(f"    {rel}")
        if kept:
            print(f"LEFT in the aside folder, their place is taken: {len(kept)}")
        return 0
    if args.write:
        files = git_files(ROOT)
        if not files:
            print("git is not available here: main_files.txt is written from a git checkout")
            return 2
        MANIFEST.write_text("# every file git tracks, written by `stale_files.py --write` before a push; compared with a folder by `stale_files.py`\n"
                            + "\n".join(files) + "\n", encoding="utf-8")
        print(f"wrote {MANIFEST.relative_to(ROOT).as_posix()}: {len(files)} files")
        return 0
    manifest = manifest_paths(MANIFEST.read_text(encoding="utf-8"))
    extra, missing = compare(ROOT, manifest)
    print(f"FOLDER {ROOT.name} against main_files.txt ({len(manifest)} files of main)")
    print(f"NOT IN MAIN: {len(extra)} .py/.yml file(s)")
    for rel in extra:
        print(f"    {rel}")
    print(f"IN MAIN, NOT HERE: {len(missing)} file(s)")
    for rel in missing:
        print(f"    {rel}")
    if args.move and extra:
        dest = ROOT.parent / f"{ROOT.name}_not_in_main_{time.strftime('%Y%m%d')}"
        moved = move_aside(ROOT, extra, dest)
        print(f"MOVED {len(moved)} file(s) to {dest} (same sub-folders; move one back to undo it)")
    elif extra:
        print("nothing moved. --move puts the files above aside in a folder next to this one; it deletes nothing.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

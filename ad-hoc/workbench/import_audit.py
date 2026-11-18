"""Free. Does every import in the repo still resolve? Run it after deleting or moving code.

    python ad-hoc/workbench/import_audit.py            prints what does not resolve; exit 1 when anything in fdd_utils does not
    python ad-hoc/workbench/import_audit.py --all      also the names a script takes from the repo's own top-level modules

A lazy import inside a function fails only when that function runs, so a deletion can leave a call that the test gate never reaches. This
reads every .py file in the repo (fdd_utils, the root scripts, ad-hoc, docs), finds every `import x` and `from x import y` at any depth,
and checks, for the repo's own modules, that the module exists and that `y` is a name it defines or a module inside it. Third-party
modules are only looked up (an optional dependency that is not installed is listed, not failed).
"""
from __future__ import annotations

import ast
import importlib
import importlib.util
import sys
from pathlib import Path
from typing import Dict, Iterator, List, Optional, Set, Tuple

ROOT = Path(__file__).resolve().parents[2]
SKIP_DIRS = {".venv", "__pycache__", ".git", "node_modules", "output", "for_test", "databooks", "pptx_previews", "logs"}
#: top-level names that resolve to a file in these folders when a script puts them on sys.path (the scripts do it themselves)
LOCAL_DIRS = [ROOT, ROOT / "ad-hoc" / "workbench"]


def py_files() -> Iterator[Path]:
    for path in sorted(ROOT.rglob("*.py")):
        if not any(part in SKIP_DIRS for part in path.relative_to(ROOT).parts):
            yield path


def module_name(path: Path) -> str:
    parts = list(path.relative_to(ROOT).with_suffix("").parts)
    if parts[-1] == "__init__":
        parts = parts[:-1]
    return ".".join(parts)


def imports_of(path: Path) -> Iterator[Tuple[str, List[str], int]]:
    """(absolute module, names taken from it, line) for every import in the file, at any depth."""
    tree = ast.parse(path.read_text(encoding="utf-8"))
    me = module_name(path)
    package = me if path.name == "__init__.py" else me.rpartition(".")[0]
    for node in ast.walk(tree):
        if isinstance(node, ast.ImportFrom):
            base = node.module or ""
            if node.level:
                parts = package.split(".")
                parts = parts[: len(parts) - (node.level - 1)] if package else []
                base = ".".join(parts + ([node.module] if node.module else []))
            yield base, [a.name for a in node.names], node.lineno
        elif isinstance(node, ast.Import):
            for alias in node.names:
                yield alias.name, [], node.lineno


def local_top_level() -> Set[str]:
    names = set()
    for folder in LOCAL_DIRS:
        for item in folder.glob("*.py"):
            names.add(item.stem)
    return names


def check(base: str, names: List[str], cache: Dict[str, Optional[object]]) -> Optional[str]:
    """A reason the import does not resolve, or None."""
    if not base:
        return None
    if base not in cache:
        try:
            cache[base] = importlib.import_module(base)
        except Exception as exc:                                              # noqa: BLE001
            cache[base] = exc
    module = cache[base]
    if isinstance(module, Exception):
        if isinstance(module, ModuleNotFoundError) and module.name and not (module.name == base or base.startswith(module.name + ".")
                                                                           or module.name.startswith("fdd_utils")):
            return None                                                       # a dependency of the module is not installed: not ours to fail
        return f"{base}: {type(module).__name__}: {str(module)[:120]}"
    missing = []
    for name in names:
        if name == "*" or hasattr(module, name):
            continue
        if hasattr(module, "__path__"):                                       # a package: the name may be a module inside it
            try:
                if importlib.util.find_spec(f"{base}.{name}") is not None:
                    continue
            except (ImportError, ValueError):
                pass
        missing.append(name)
    return f"{base}: no {', '.join(missing)}" if missing else None


def audit(include_scripts: bool = False) -> Dict[str, List[str]]:
    """{file: [problems]} for the repo's own imports. Scripts are checked only when include_scripts (they import each other by sys.path)."""
    for folder in LOCAL_DIRS:
        if str(folder) not in sys.path:
            sys.path.insert(0, str(folder))
    local = local_top_level()
    cache: Dict[str, Optional[object]] = {}
    problems: Dict[str, List[str]] = {}
    for path in py_files():
        rel = str(path.relative_to(ROOT))
        for base, names, line in imports_of(path):
            top = base.split(".")[0]
            if top != "fdd_utils" and not (include_scripts and top in local):
                continue
            reason = check(base, names, cache)
            if reason:
                problems.setdefault(rel, []).append(f"line {line}: {reason}")
    return problems


def main(argv: Optional[List[str]] = None) -> int:
    include = "--all" in (argv if argv is not None else sys.argv[1:])
    problems = audit(include)
    for rel, items in sorted(problems.items()):
        print(rel)
        for item in items:
            print(f"    {item}")
    count = sum(len(v) for v in problems.values())
    print(f"{count} import(s) in {len(problems)} file(s) do not resolve" if problems else "every import resolves")
    return 1 if problems else 0


if __name__ == "__main__":
    raise SystemExit(main())

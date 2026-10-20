"""Zip what the server runs: fdd_app.py and fdd_utils/, nothing else.

fdd_utils/logs/ and fdd_utils/output/ (and __pycache__) are left out: they
hold past prompts and outputs, i.e. client figures, and the server keeps its
own. config.yml and template.pptx ARE included, since the app needs them --
the zip therefore carries the API key.

    python make_deploy_package.py        # -> fdd_deploy.zip, overwritten each time
"""
from __future__ import annotations

import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SKIP_DIRS = {"logs", "output", "__pycache__"}


def main() -> int:
    out = ROOT / "fdd_deploy.zip"
    count = 0
    with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.write(ROOT / "fdd_app.py", arcname="fdd_app.py")
        count += 1
        for path in sorted((ROOT / "fdd_utils").rglob("*")):
            rel = path.relative_to(ROOT)
            if not path.is_file() or SKIP_DIRS & set(rel.parts) or path.suffix in (".pyc", ".pyo"):
                continue
            archive.write(path, arcname=str(rel))
            count += 1
    print(f"{out.name}: {count} files, {out.stat().st_size / 1e6:.1f} MB (includes config.yml)")
    return 0


if __name__ == "__main__":
    sys.exit(main())

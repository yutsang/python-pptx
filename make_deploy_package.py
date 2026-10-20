"""Build the zip that goes on the server: the app and nothing else.

The repo root carries a lot that only matters on a developer machine -- ad-hoc/
probes, docs, PLAN files, the inspect_* CLIs, deck photos, test databooks. The
running app needs only:

    fdd_app.py, requirements.txt, run_app.bat, .streamlit/config.toml
    fdd_utils/  (code, mappings.yml, prompts.yml, font_metrics/,
                 and the two files git does not carry: config.yml, template.pptx)

fdd_utils/logs/ and fdd_utils/output/ are left out on purpose: they hold every
prompt and output of past runs, i.e. client figures, and the server makes its
own. The zip DOES contain config.yml, i.e. the API key -- copy it straight to
the server and do not share it.

    python make_deploy_package.py            # -> fdd_deploy_<date>.zip next to this file
"""
from __future__ import annotations

import sys
import zipfile
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
TOP_FILES = ["fdd_app.py", "requirements.txt", "run_app.bat", ".streamlit/config.toml"]
SKIP_DIRS = {"logs", "output", "__pycache__"}
REQUIRED = ["fdd_utils/config.yml", "fdd_utils/template.pptx",
            "fdd_utils/font_metrics/arial_eng.json", "fdd_utils/font_metrics/msyh_chi.json",
            "fdd_utils/mappings.yml", "fdd_utils/prompts.yml"]


def main() -> int:
    missing = [p for p in REQUIRED if not (ROOT / p).exists()]
    if missing:
        print("❌ missing, the server would not run without them:")
        for p in missing:
            print("   " + p)
        return 1
    out = ROOT / f"fdd_deploy_{datetime.now():%Y%m%d_%H%M}.zip"
    count = 0
    with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name in TOP_FILES:
            path = ROOT / name
            if path.exists():
                archive.write(path, arcname=name)
                count += 1
        for path in sorted((ROOT / "fdd_utils").rglob("*")):
            rel = path.relative_to(ROOT)
            if not path.is_file() or SKIP_DIRS & set(rel.parts) or path.suffix in (".pyc", ".pyo"):
                continue
            archive.write(path, arcname=str(rel))
            count += 1
    size_mb = out.stat().st_size / 1e6
    print(f"✅ {out.name}: {count} files, {size_mb:.1f} MB")
    print("   contains config.yml (API key): copy it to the server directly, do not share it.")
    print("   On the server: unzip into an empty folder, then from THAT folder:")
    print("     pip install -r requirements.txt")
    print("     streamlit run fdd_app.py --server.address 0.0.0.0 --server.port 8501")
    return 0


if __name__ == "__main__":
    sys.exit(main())

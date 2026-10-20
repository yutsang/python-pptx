"""inspect_databook's --run-ai path with a canned provider: zero tokens, seconds.

Every AIClient call returns the same short Chinese bullet (and an empty
Validator verdict), so nothing about commentary QUALITY can be read off this.
What it does exercise is everything around the model: extraction, the compact
output, the per-file TIME line, 5-7, the 6c cross-account pass, the scorecard
row and, with --export, payloads/export/inspect. Use it to check plumbing
before asking for a real run that costs ~15 minutes and tokens.

    python ad-hoc/workbench/stub_inspect_run.py "Project Gold Kunshan.databook.xlsx" [--export] [--limit 6]
"""
from __future__ import annotations

import argparse
import json
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

import fdd_utils.ai.client as _client  # noqa: E402

_client.AIClient.validate_config = lambda self: None
_client.AIClient.initialize_client = lambda self: (None, "stub")


def _canned(self, user_prompt="", system_prompt="", **_kwargs):
    if "4" in str(self.agent_name) or "Validator" in str(self.agent_name):
        content = json.dumps({"clause_reviews": [], "is_valid": True})
    else:
        content = "截至2025年12月31日，余额为100.0万元，主要包括甲公司50.0万元。"
    return {"content": content, "prompt_tokens": 10, "completion_tokens": 10, "total_tokens": 20}


_client.AIClient.get_response = _canned

import inspect_databook as I  # noqa: E402


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("path")
    ap.add_argument("--export", action="store_true")
    ap.add_argument("--limit", type=int, default=6)
    args = ap.parse_args()
    out_dir = tempfile.mkdtemp(prefix="stub_inspect_")
    # Never append canned scores to the real scorecard.
    I.SCORECARD_PATH = Path(out_dir) / "scorecard.jsonl"
    I.inspect_one(args.path, None, "", True, "local", None, limit=args.limit,
                  export_pptx=args.export, pptx_out_dir=out_dir)
    print(f"(stub outputs in {out_dir})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

"""Dump the exact source around a few critical fdd_utils/ai.py fragments.

Used to compare a Windows machine's copy of ai.py (downloaded manually from
GitHub, not git-pulled -- so its exact version is otherwise unknown) against
what the repo currently has, for the specific question of whether the
deepseek branch actually sends "max_tokens" to the API. Paste this script's
output back for a direct line-by-line comparison instead of guessing from
symptoms alone.

Usage: python inspect_ai_py_version.py [path to ai.py, default fdd_utils/ai.py]
"""
import sys

path = sys.argv[1] if len(sys.argv) > 1 else "fdd_utils/ai.py"
with open(path, encoding="utf-8") as f:
    lines = f.read().splitlines()


def dump_around(pattern, context_after=10, context_before=2, label=""):
    found = False
    for i, line in enumerate(lines):
        if pattern in line:
            found = True
            print(f"\n--- {label or pattern} (line {i + 1}) ---")
            start = max(0, i - context_before)
            end = min(len(lines), i + context_after)
            for j in range(start, end):
                print(f"{j + 1:5d}: {lines[j]}")
    if not found:
        print(f"\n--- {label or pattern}: NOT FOUND ---")


print(f"=== {path} ({len(lines)} lines) ===")
dump_around("if max_tokens:", context_after=8, label="max_tokens param gate (should set params['max_tokens'])")
dump_around("elif self.model_type == 'deepseek':", context_after=10, label="deepseek dispatch branch")
dump_around('"expected_max_output_tokens":', context_after=2, label="expected_max_output_tokens field definition")

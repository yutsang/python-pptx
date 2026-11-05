#!/usr/bin/env python3
"""Which config.yml is actually being loaded, and what's in its pptx:
section -- for when _presentation_tables_enabled() (or any other pptx:
setting) doesn't reflect an edit that was just made. Isolates two possible
causes: editing a config.yml that isn't the one this process reads (e.g. a
second/older copy of the repo on the same machine), versus a YAML
indentation mistake that silently drops the key during parsing.

Read-only.

Usage:
    python check_config_path.py
"""
import os
import sys

sys.path.insert(0, ".")

from fdd_utils.financial_common import package_file_path, load_yaml_file

resolved_path = package_file_path("config.yml")
abs_path = os.path.abspath(resolved_path)
print(f"This process's own working directory : {os.getcwd()}")
print(f"config.yml resolves to                : {resolved_path}")
print(f"absolute path                         : {abs_path}")
print(f"file exists                           : {os.path.exists(abs_path)}")
if os.path.exists(abs_path):
    print(f"last modified                         : "
          f"{__import__('datetime').datetime.fromtimestamp(os.path.getmtime(abs_path))}")
    print(f"file size                             : {os.path.getsize(abs_path)} bytes")

print()
try:
    config = load_yaml_file(abs_path)
    pptx_section = (config or {}).get("pptx") or {}
    print(f"Top-level keys in this file           : {sorted((config or {}).keys())}")
    print(f"Keys under pptx:                      : {sorted(pptx_section.keys())}")
    print(f"pptx.presentation_tables (raw)        : {pptx_section.get('presentation_tables')!r}")
    print(f"pptx.table_style_id (raw)             : {pptx_section.get('table_style_id')!r}")
except Exception as exc:
    print(f"Could not load/parse this file: {type(exc).__name__}: {exc}")
    print("If this is a YAML syntax/indentation error, that alone would explain "
          "keys silently missing -- yaml.safe_load either raises (shown above) "
          "or, for a pure indentation mistake, can attach presentation_tables "
          "as a sibling of the wrong key instead of under pptx: at all.")

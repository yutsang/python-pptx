"""Apply private routing/attribution settings without replacing credentials."""
from __future__ import annotations

import argparse
import copy
from pathlib import Path
import shutil
import time

import yaml


def apply_settings(config_path: Path, additions_path: Path) -> Path:
    original = yaml.safe_load(config_path.read_text(encoding="utf-8"))
    additions = yaml.safe_load(additions_path.read_text(encoding="utf-8"))
    if not isinstance(original, dict) or not isinstance(additions, dict):
        raise ValueError("Both settings files must contain YAML mappings")
    allowed = {"workbench": "header_names", "dashscope": "header_names", "pptx": "source_labels"}
    updated = copy.deepcopy(original)
    for section, fields in additions.items():
        if section not in allowed or not isinstance(fields, dict) or set(fields) != {allowed[section]}:
            raise ValueError("Additions may contain only routing header names and source labels")
        name = allowed[section]
        values = fields[name]
        if not isinstance(values, dict):
            raise ValueError("Routing names and source labels must be YAML mappings")
        target = updated.setdefault(section, {}).setdefault(name, {})
        for key, value in values.items():
            if key in target and target[key] != value:
                raise ValueError("An existing private setting differs; review it before replacing")
            target[key] = value
    backup = config_path.with_name(f"{config_path.name}.{time.time_ns()}.bak")
    shutil.copy2(config_path, backup)
    config_path.write_text(yaml.safe_dump(updated, allow_unicode=True, sort_keys=False), encoding="utf-8")
    return backup


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("config", type=Path)
    parser.add_argument("additions", type=Path)
    args = parser.parse_args()
    backup = apply_settings(args.config, args.additions)
    print(f"Private settings updated; all existing credentials retained. Backup: {backup.name}")


if __name__ == "__main__":
    main()

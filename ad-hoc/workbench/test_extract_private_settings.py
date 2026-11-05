"""Recovering the moved gateway header names and slide source line from an old checkout -- offline, standard library plus PyYAML.

    python ad-hoc/workbench/test_extract_private_settings.py

The trees written here are synthetic and use placeholder names; nothing real is read.
"""
from __future__ import annotations

import contextlib
import io
import sys
import tempfile
import unittest
from pathlib import Path

import yaml

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent.parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(ROOT))

import extract_private_settings as X  # noqa: E402
import apply_private_settings as A  # noqa: E402

CHARGE, REGION = "x-example-charge", "x-example-region"
ZH, EN = "资料来源：示例来源", "Source: Example attribution"

OLD_CLIENT = f"""
            headers = {{
                'Ocp-Apim-Subscription-Key': subscription_key,
                '{CHARGE}': str(self.config_details.get('charge_code') or '0000'),
                "{REGION}": str(self.config_details.get("region_override") or 'westeurope'),
            }}
"""
NEW_CLIENT = """
            header_names = self.config_details.get('header_names') or {}
            for field, default in [('charge_code', '0000'), ('region_override', 'westeurope')]:
                name = header_names.get(field)
"""
OLD_TABLES = f"""
        source_run = source_p.add_run()
        source_run.text = "{ZH}" if is_chinese_databook else "{EN}"
        source_run.font.name = 'Arial'
"""
NEW_TABLES = """
        source_labels = self.pptx_settings.get("source_labels") or {}
        source_run.text = (source_labels.get("zh") or "x") if is_chinese_databook else (source_labels.get("en") or "y")
"""
OLD_PROBE = f"""
        headers = {{
            "Authorization": "x",
            "{CHARGE}": charge_code,
            "{REGION}": region_override,
        }}
"""


def tree(client: str, tables: str, probe: str = "") -> Path:
    root = Path(tempfile.mkdtemp(prefix="oldtree_"))
    for rel, text in ((X.CLIENT, client), (X.TABLES, tables), (X.PROBE, probe)):
        if text:
            (root / rel).parent.mkdir(parents=True, exist_ok=True)
            (root / rel).write_text(text, encoding="utf-8")
    return root


def run(root: Path, out: Path):
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        rc = X.main(["--root", str(root), "--out", str(out)])
    return rc, buf.getvalue()


class Extract(unittest.TestCase):
    def test_old_hard_coded_values_are_written_to_the_additions_file_and_never_printed(self):
        root = tree(OLD_CLIENT, OLD_TABLES, OLD_PROBE)
        out = root / "additions.yml"
        rc, printed = run(root, out)
        self.assertEqual(rc, 0, printed)
        got = yaml.safe_load(out.read_text(encoding="utf-8"))
        self.assertEqual(got, {
            "workbench": {"header_names": {"charge_code": CHARGE, "region_override": REGION}},
            "dashscope": {"header_names": {"charge_code": CHARGE, "region_override": REGION}},
            "pptx": {"source_labels": {"zh": ZH, "en": EN}},
        })
        for secret in (CHARGE, REGION, ZH, EN, "Example attribution"):
            self.assertNotIn(secret, printed)
        self.assertIn("found both", printed)

    def test_what_it_writes_is_accepted_by_apply_private_settings_and_credentials_survive(self):
        root = tree(OLD_CLIENT, OLD_TABLES)
        out = root / "additions.yml"
        self.assertEqual(run(root, out)[0], 0)
        cfg = root / "config.yml"
        cfg.write_text(yaml.safe_dump({"workbench": {"api_key": "KEEP-ME", "header_names": {}}, "pptx": {"source_labels": {}}}), encoding="utf-8")
        A.apply_settings(cfg, out)
        merged = yaml.safe_load(cfg.read_text(encoding="utf-8"))
        self.assertEqual(merged["workbench"]["api_key"], "KEEP-ME")
        self.assertEqual(merged["workbench"]["header_names"], {"charge_code": CHARGE, "region_override": REGION})
        self.assertEqual(merged["pptx"]["source_labels"], {"zh": ZH, "en": EN})

    def test_a_checkout_that_is_already_the_public_version_yields_nothing_and_writes_nothing(self):
        root = tree(NEW_CLIENT, NEW_TABLES)
        out = root / "additions.yml"
        rc, printed = run(root, out)
        self.assertEqual(rc, 1)
        self.assertFalse(out.exists())
        self.assertIn("NOT found", printed)

    def test_one_missing_half_of_the_pair_is_not_written_as_if_it_were_whole(self):
        only_charge = OLD_CLIENT.replace(f'"{REGION}": str(self.config_details.get("region_override") or \'westeurope\'),', "")
        root = tree(only_charge, OLD_TABLES)
        out = root / "additions.yml"
        rc, printed = run(root, out)
        self.assertEqual(rc, 0)
        self.assertNotIn("workbench", yaml.safe_load(out.read_text(encoding="utf-8")))
        self.assertIn("workbench.header_names: NOT found", printed)

    def test_an_existing_output_file_is_left_alone(self):
        root = tree(OLD_CLIENT, OLD_TABLES)
        out = root / "additions.yml"
        out.write_text("keep: me\n", encoding="utf-8")
        rc, _ = run(root, out)
        self.assertEqual(rc, 2)
        self.assertEqual(out.read_text(encoding="utf-8"), "keep: me\n")

    def test_missing_files_are_reported_not_raised(self):
        root = Path(tempfile.mkdtemp())
        rc, printed = run(root, root / "additions.yml")
        self.assertEqual(rc, 1)
        self.assertIn("file missing", printed)


if __name__ == "__main__":
    unittest.main(verbosity=2)

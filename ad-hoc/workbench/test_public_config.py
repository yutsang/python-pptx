"""Offline checks for settings kept outside the public source archive."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
import sys
import unittest
import tempfile
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

from fdd_utils.ai.client import AIClient
from fdd_utils.pptx.generation import PowerPointGenerator
from pptx import Presentation
from pptx.util import Inches


class PrivateGatewayConfiguration(unittest.TestCase):
    def test_one_time_settings_update_preserves_credentials_and_other_preferences(self):
        import yaml
        from apply_private_settings import apply_settings
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            config = root / "config.yml"
            additions = root / "private-config-additions.yml"
            original = {"workbench": {"api_key": "unchanged-test-key", "api_base": "https://example.invalid", "charge_code": "unchanged-charge"}, "pptx": {"row_limit": 17}}
            patch_data = {"workbench": {"header_names": {"charge_code": "x-example-charge"}}, "pptx": {"source_labels": {"en": "Private attribution"}}}
            config.write_text(yaml.safe_dump(original))
            additions.write_text(yaml.safe_dump(patch_data))
            backup = apply_settings(config, additions)
            actual = yaml.safe_load(config.read_text())
            self.assertEqual(yaml.safe_load(backup.read_text()), original)
            self.assertEqual(actual["workbench"]["api_key"], original["workbench"]["api_key"])
            self.assertEqual(actual["workbench"]["api_base"], original["workbench"]["api_base"])
            self.assertEqual(actual["workbench"]["charge_code"], original["workbench"]["charge_code"])
            self.assertEqual(actual["pptx"]["row_limit"], 17)
            self.assertEqual(actual["workbench"]["header_names"], patch_data["workbench"]["header_names"])

    def test_settings_update_refuses_to_replace_existing_credentials(self):
        from apply_private_settings import apply_settings
        with tempfile.TemporaryDirectory() as directory:
            config = Path(directory) / "config.yml"
            additions = Path(directory) / "additions.yml"
            original = "workbench:\n  api_key: unchanged-test-key\n"
            config.write_text(original)
            additions.write_text("workbench:\n  api_key: replacement-test-key\n")
            with self.assertRaises(ValueError):
                apply_settings(config, additions)
            self.assertEqual(config.read_text(), original)

    def test_sdk_receives_the_configured_routing_header_names_and_values(self):
        client = AIClient.__new__(AIClient)
        client.model_type = "workbench"
        client.config_details = {
            "api_key": "test-key",
            "api_base": "https://gateway.example.invalid",
            "api_version": "test-version",
            "chat_model": "test-model",
            "charge_code": "example-charge",
            "region_override": "example-region",
            "header_names": {
                "charge_code": "x-example-charge",
                "region_override": "x-example-region",
            },
        }
        with patch("fdd_utils.ai.client.AzureOpenAI") as sdk, patch("fdd_utils.ai.client.httpx.Client"):
            _, model = client.initialize_client()
        expected = {
            "Ocp-Apim-Subscription-Key": "test-key",
            "x-example-charge": "example-charge",
            "x-example-region": "example-region",
        }
        self.assertEqual(sdk.call_args.kwargs["default_headers"], expected)
        self.assertEqual(client._workbench_headers, expected)
        self.assertEqual(model, "test-model")

    def test_probe_uses_the_same_private_header_mapping(self):
        spec = importlib.util.spec_from_file_location("connectivity_probe", ROOT / "ad-hoc/workbench/test_dashscope_connectivity.py")
        probe = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(probe)
        headers = probe._headers("test-key", "workbench", charge_code="example-charge", region_override="example-region", header_names={"charge_code": "x-example-charge", "region_override": "x-example-region"})
        self.assertEqual(headers["x-example-charge"], "example-charge")
        self.assertEqual(headers["x-example-region"], "example-region")
        self.assertEqual(headers["Authorization"], "Bearer test-key")


class PublicDemonstration(unittest.TestCase):
    def test_every_replayed_account_is_explicitly_fictional(self):
        data = json.loads((ROOT / "fdd_utils/demo_results.json").read_text())
        self.assertTrue({"Cash", "OI", "OC", "__BS_summary__", "__IS_summary__"} <= data.keys())
        for key, value in data.items():
            if key.startswith("__"):
                self.assertIn("fictional", value.lower())
                continue
            for stage in ["subagent_1", "subagent_2", "subagent_3", "subagent_4", "final"]:
                self.assertIn("fictional", value[stage].lower(), (key, stage))
            self.assertIsInstance(value["agent_4_validation"]["clause_reviews"], list)

    def test_demo_replays_aliases_and_section_summaries_without_a_provider(self):
        from fdd_utils.ui.ai_panel import _run_demo_ai
        from unittest.mock import Mock
        with patch("fdd_utils.ui.ai_panel.time.sleep"):
            result = _run_demo_ai(["Cash at bank", "营业收入", "Cost"], 0, Mock(), Mock())
        for key in ["Cash at bank", "营业收入", "Cost"]:
            self.assertIn("fictional", result[key]["final"].lower())
        self.assertTrue({"__BS_summary__", "__IS_summary__"} <= result.keys())


class PrivateAttribution(unittest.TestCase):
    def test_rendered_tables_use_private_or_generic_attribution_in_both_languages(self):
        table = {
            "title": "Example table", "periods": ["A", "B"],
            "period_labels": {"A": "A", "B": "B"},
            "rows": [{"label": "Example item", "values": {"A": 100000, "B": 120000}}],
            "total_row": {"label": "Total", "values": {"A": 100000, "B": 120000}},
        }
        for chinese, labels, expected in [
            (False, {"en": "Private attribution"}, "Private attribution"),
            (True, {"zh": "私有來源"}, "私有來源"),
            (False, {}, "Source: Management information; independent analysis"),
            (True, {}, "资料来源：管理层信息；独立分析"),
        ]:
            with self.subTest(chinese=chinese, private=bool(labels)):
                with patch("fdd_utils.pptx.generation._load_pptx_settings", return_value={}):
                    generator = PowerPointGenerator("unused-template")
                generator.pptx_settings["source_labels"] = labels
                presentation = Presentation()
                slide = presentation.slides.add_slide(presentation.slide_layouts[6])
                generator._render_presentation_table(slide, Inches(0.5), Inches(0.5), Inches(5), table, chinese)
                self.assertIn(expected, [shape.text for shape in slide.shapes if shape.has_text_frame])


if __name__ == "__main__":
    unittest.main(verbosity=2)

"""config_report.py: the report a person pastes back must hold no secret, and --set-flash-default must change one line and nothing else.

    python ad-hoc/workbench/test_config_report.py
"""
from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import config_report as cr  # noqa: E402

TEXT = '''# a comment that stays
default:
  ai_provider: "dashscope"
  model: "dashscope_qwen3.7-max"   # starts here
dashscope:
  api_key: "sk-THISISASECRETKEYVALUE0123456789"
  api_base: "https://gateway.example.internal/aiworkbench/openAi/v1"
  charge_code: "9999"
  available_models:
    - id: "dashscope_deepseek-v4-flash-0731"
      kind: chat
      label: "Flash"
    - id: "dashscope_qwen3.7-max"
      kind: chat
      label: "Qwen"
workbench:
  api_key: ""
workbench:
  api_key: "another-secret"
'''
EXAMPLE = TEXT.replace("sk-THISISASECRETKEYVALUE0123456789", "").replace("https://gateway.example.internal/aiworkbench/openAi/v1", "")


def files(tmp: str):
    cfg, ex = Path(tmp) / "config.yml", Path(tmp) / "config.example.yml"
    cfg.write_text(TEXT, encoding="utf-8")
    ex.write_text(EXAMPLE, encoding="utf-8")
    return cfg, ex


class Masking(unittest.TestCase):
    def test_no_secret_host_or_code_reaches_the_report(self):
        with tempfile.TemporaryDirectory() as tmp:
            cfg, ex = files(tmp)
            text = cr.report(cfg, ex, verbose=True, section="dashscope")
        for leak in ("THISISASECRET", "gateway.example", "9999", "another-secret"):
            self.assertNotIn(leak, text)
        self.assertIn("dashscope.api_key: set", text)

    def test_a_model_id_is_not_taken_for_a_key(self):
        self.assertFalse(cr.secret("default.model", "dashscope_deepseek-v4-flash-0731-with-a-long-suffix-xx"))
        self.assertTrue(cr.secret("x.token_value", "sk-abc"))

    def test_a_duplicate_top_level_block_is_named(self):
        with tempfile.TemporaryDirectory() as tmp:
            cfg, ex = files(tmp)
            _data, dups = cr.load(cfg)
        self.assertEqual([k for k, _n in dups], ["workbench"])


class SetFlashDefault(unittest.TestCase):
    def test_one_line_changes_the_comment_and_the_rest_stay_and_a_copy_is_kept(self):
        with tempfile.TemporaryDirectory() as tmp:
            cfg, _ex = files(tmp)
            result = cr.set_flash_default(cfg)
            after = cfg.read_text(encoding="utf-8")
            kept = list(Path(tmp).glob("config.yml.bak-*"))
        self.assertTrue(result.startswith("CHANGED"), result)
        self.assertIn('model: "dashscope_deepseek-v4-flash-0731"   # starts here', after)
        self.assertEqual(after.replace("dashscope_deepseek-v4-flash-0731", "dashscope_qwen3.7-max", 1), TEXT)
        self.assertEqual(len(kept), 1)

    def test_it_is_idempotent_and_refuses_when_there_is_no_single_flash_model(self):
        with tempfile.TemporaryDirectory() as tmp:
            cfg, _ex = files(tmp)
            cr.set_flash_default(cfg)
            self.assertTrue(cr.set_flash_default(cfg).startswith("already set"))
            cfg.write_text(TEXT.replace("dashscope_deepseek-v4-flash-0731", "something-else"), encoding="utf-8")
            self.assertTrue(cr.set_flash_default(cfg).startswith("NOT CHANGED"))


# the company machine's shape: the default provider is not in the dropdown, so the app starts on the first entry (Pro here)
WORKBENCH_FIRST = (TEXT.replace('ai_provider: "dashscope"', 'ai_provider: "workbench"')
                   .replace('model: "dashscope_qwen3.7-max"', 'model: "gpt-6-luna"')
                   .replace('workbench:\n  api_key: ""\nworkbench:\n  api_key: "another-secret"\n', 'workbench:\n  api_key: "another-secret"\n')
                   .replace('    - id: "dashscope_deepseek-v4-flash-0731"\n',
                            '    - id: "dashscope_deepseek-v4-pro-0813"\n      kind: chat\n      label: "Pro"\n'
                            '    - id: "dashscope_deepseek-v4-flash-0731"\n'))


class StartsOnFlash(unittest.TestCase):
    def test_a_default_provider_that_is_not_in_the_dropdown_starts_the_app_on_the_first_entry(self):
        with tempfile.TemporaryDirectory() as tmp:
            cfg = Path(tmp) / "config.yml"
            cfg.write_text(WORKBENCH_FIRST, encoding="utf-8")
            data, _ = cr.load(cfg)
            self.assertEqual(cr.starts_on(data)["model_name"], "dashscope_deepseek-v4-pro-0813")
            cfg.write_text(WORKBENCH_FIRST.replace('"gpt-6-luna"', '"dashscope_deepseek-v4-flash-0731"'), encoding="utf-8")
            data, _ = cr.load(cfg)
            self.assertEqual(cr.starts_on(data)["model_name"], "dashscope_deepseek-v4-pro-0813", "the model alone does not move it: the provider decides")

    def test_both_lines_are_set_and_the_session_starts_on_flash(self):
        with tempfile.TemporaryDirectory() as tmp:
            cfg = Path(tmp) / "config.yml"
            cfg.write_text(WORKBENCH_FIRST, encoding="utf-8")
            result = cr.set_flash_default(cfg)
            after = cfg.read_text(encoding="utf-8")
            data, _ = cr.load(cfg)
            kept = list(Path(tmp).glob("config.yml.bak-*"))
        self.assertTrue(result.startswith("CHANGED"), result)
        self.assertEqual(cr.starts_on(data)["model_name"], "dashscope_deepseek-v4-flash-0731")
        restored = after.replace('ai_provider: "dashscope"', 'ai_provider: "workbench"', 1).replace(
            'model: "dashscope_deepseek-v4-flash-0731"   # starts here', 'model: "gpt-6-luna"   # starts here', 1)
        self.assertEqual(restored, WORKBENCH_FIRST, "nothing but those two lines moved")
        self.assertEqual(len(kept), 1)

    def test_a_file_that_already_says_flash_is_left_alone_whatever_its_quoting(self):
        unquoted = TEXT.replace('ai_provider: "dashscope"', "ai_provider: dashscope").replace(
            'model: "dashscope_qwen3.7-max"   # starts here', "model: dashscope_deepseek-v4-flash-0731")
        with tempfile.TemporaryDirectory() as tmp:
            cfg = Path(tmp) / "config.yml"
            cfg.write_text(unquoted, encoding="utf-8")
            result = cr.set_flash_default(cfg)
            text = cfg.read_text(encoding="utf-8")
            kept = list(Path(tmp).glob("config.yml.bak-*"))
        self.assertTrue(result.startswith("already set"), result)
        self.assertEqual((text, kept), (unquoted, []))

    def test_it_puts_the_file_back_when_flash_is_hidden_and_the_app_would_not_start_on_it(self):
        hidden = WORKBENCH_FIRST.replace('      kind: chat\n      label: "Flash"\n', '      kind: chat\n      label: "Flash"\n      listed: false\n')
        with tempfile.TemporaryDirectory() as tmp:
            cfg = Path(tmp) / "config.yml"
            cfg.write_text(hidden, encoding="utf-8")
            result = cr.set_flash_default(cfg)
            text = cfg.read_text(encoding="utf-8")
        self.assertTrue(result.startswith("RESTORED") or result.startswith("NOT CHANGED"), result)
        self.assertEqual(text, hidden)


class DropIgnoredBlocks(unittest.TestCase):
    def test_the_report_says_what_the_ignored_block_holds_without_showing_it(self):
        with tempfile.TemporaryDirectory() as tmp:
            cfg, ex = files(tmp)
            text = cr.report(cfg, ex)
        self.assertIn("ignored 'workbench' block", text)
        self.assertNotIn("another-secret", text)

    def test_the_earlier_block_goes_and_what_the_app_reads_is_the_same(self):
        with tempfile.TemporaryDirectory() as tmp:
            cfg, _ex = files(tmp)
            before, _ = cr.load(cfg)
            result = cr.drop_ignored_blocks(cfg)
            after, dups = cr.load(cfg)
            text = cfg.read_text(encoding="utf-8")
            kept = list(Path(tmp).glob("config.yml.bak-*"))
        self.assertTrue(result.startswith("REMOVED"), result)
        self.assertEqual((after, dups), (before, []))
        self.assertEqual(text.count("workbench:"), 1)
        self.assertIn("# a comment that stays", text)
        self.assertEqual(len(kept), 1)

    def test_a_block_the_scanner_cannot_cut_cleanly_puts_the_file_back(self):
        odd = "a: {x: 1,\ny: 2}\na: 3\nb: 1\n"
        with tempfile.TemporaryDirectory() as tmp:
            cfg = Path(tmp) / "config.yml"
            cfg.write_text(odd, encoding="utf-8")
            result = cr.drop_ignored_blocks(cfg)
            text = cfg.read_text(encoding="utf-8")
        self.assertTrue(result.startswith("RESTORED"), result)
        self.assertEqual(text, odd)

    def test_nothing_to_do_without_a_duplicate(self):
        with tempfile.TemporaryDirectory() as tmp:
            cfg = Path(tmp) / "config.yml"
            cfg.write_text("a: 1\nb: 2\n", encoding="utf-8")
            self.assertTrue(cr.drop_ignored_blocks(cfg).startswith("NOT CHANGED"))


if __name__ == "__main__":
    unittest.main()

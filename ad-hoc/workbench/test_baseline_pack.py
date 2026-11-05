"""The baseline pack: what pairs a run with a databook, what is refused, what is listed as missing, what never leaves -- offline.

Zero tokens. The material is written here: a stand-in databook (any bytes; the tool pairs by hash), run folders, a ledger and a deck.
This proves the rules of the tool. It says nothing about any real client's material, none of which is on this machine.

    python ad-hoc/workbench/test_baseline_pack.py
"""
from __future__ import annotations

import contextlib
import hashlib
import io
import json
import os
import sys
import tempfile
import time
import unittest
from pathlib import Path
from unittest import mock

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parents[1]))

import baseline_pack as B  # noqa: E402

START = time.mktime(time.strptime("20260101100000", "%Y%m%d%H%M%S"))
# built at run time so no literal in this file looks like a credential to a secret scanner
FAKE_SK = "sk" + "-" + "abcdefghijklmnopqrstuv12345"
FAKE_BEARER_TOKEN = "abcdefghijklmnopqrstuvwxyz" + "012345"
FAKE_CONFIG = {"dashscope": {"api_key": "KEY-VALUE-IN-USE-1234", "api_base": "https://gateway.internal.example/v1", "chat_model": "m"}}


def iso(offset: float) -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime(START + offset))


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


class World:
    """A temp tree: a databook, logs/run_<ts>/, usage.jsonl, a deck. Each piece can be left out or damaged."""

    def __init__(self, *, evidence=True, summaries=True, deck=True, ledger=True, identity=True, texts=None, run_name="run_20260101_100000",
                 databook_bytes=b"original databook bytes"):
        self.tmp = Path(tempfile.mkdtemp(prefix="pack_"))
        self.book = self.tmp / "ClientNameDatabook.xlsx"
        self.book.write_bytes(databook_bytes)
        self.logs = self.tmp / "logs"
        self.run = self.logs / run_name
        self.run.mkdir(parents=True)
        accounts = texts or {"AR": "应收账款余额100.0万元，主要为甲公司。", "AP": "应付账款余额80.0万元。", "Cash": "货币资金50.0万元，较上年增加。"}
        import yaml
        (self.run / "results.yml").write_text(yaml.safe_dump({k: {"final": v} for k, v in accounts.items()}, allow_unicode=True), encoding="utf-8")
        if identity:
            (self.run / "run_identity.json").write_text(json.dumps({
                "workbook": self.book.name, "workbook_sha256": sha(self.book), "entity": "ClientEntity", "sheet": "Financials",
                "financials_from": "", "financials_sha256": None, "financials_sheet": "", "language": "Chi", "model_type": "dashscope",
                "model_name": "m", "accounts": sorted(accounts), "code_revision": "abc1234", "config_sha256": "0" * 64}), encoding="utf-8")
        (self.run / "checkpoint.jsonl").write_text("".join(json.dumps({"mapping_key": k, "fingerprint": "fp-" + k}) + "\n" for k in accounts), encoding="utf-8")
        (self.run / "digest.json").write_text("{}", encoding="utf-8")
        (self.run / "selection.json").write_text("{}", encoding="utf-8")
        (self.run / "summary.txt").write_text("RUN 20260101_100000  accounts 3  calls ok 9 / failed 0\n", encoding="utf-8")
        (self.run / "data.yml").write_text("AR: {prompt: 'the whole stage prompt'}\n", encoding="utf-8")
        (self.run / "audit.jsonl").write_text('{"x": 1}\n', encoding="utf-8")
        (self.run / "processing.log").write_text("AI call attempt 1 failed: https://gateway.internal.example/v1/chat\n", encoding="utf-8")
        (self.run / "config.yml").write_text("api" + "_key: " + "KEY-VALUE-IN-USE-1234\n", encoding="utf-8")
        if evidence:
            (self.run / "evidence").mkdir()
            for k in accounts:
                (self.run / "evidence" / f"{k}.json").write_text(json.dumps({"account": k}), encoding="utf-8")
        if summaries:
            (self.run / "section_summaries.json").write_text(json.dumps({"summaries": {"BS": "summary"}}), encoding="utf-8")
        for p in self.run.rglob("*"):
            if p.is_file():
                os.utime(p, (START + 300, START + 300))
        self.ledger = self.tmp / "usage.jsonl"
        if ledger:
            rows = [
                {"time": iso(10), "stage": "Generator", "client_agent": "subagent_1", "provider": "dashscope", "model": "m", "ok": True, "duration_s": 3.2,
                 "prompt_tokens": 1000, "completion_tokens": 200, "total_tokens": 1200, "cached_prompt_tokens": 100, "reasoning_tokens": 50,
                 "estimated_prompt_tokens": 900, "token_source": "provider", "file": self.book.name, "account": "AR", "entity": "ClientEntity", "session": "s1"},
                {"time": iso(20), "stage": "Auditor", "client_agent": "subagent_2", "provider": "dashscope", "model": "m", "ok": True,
                 "prompt_tokens": 2000, "completion_tokens": 300, "estimated_prompt_tokens": 1800, "token_source": "provider", "file": self.book.name, "account": "AR"},
                {"time": iso(30), "stage": "Generator", "provider": "dashscope", "model": "m", "ok": True, "estimated_prompt_tokens": 700,
                 "token_source": "estimate", "file": self.book.name, "account": "AP"},
                {"time": iso(40), "stage": "Generator", "provider": "dashscope", "model": "m", "ok": False, "estimated_prompt_tokens": 500,
                 "error": "APIConnectionError: could not reach https://gateway.internal.example/v1/chat", "file": self.book.name, "account": "Cash"},
                {"time": iso(50), "stage": "Generator", "provider": "dashscope", "model": "m", "ok": True, "prompt_tokens": 0, "completion_tokens": 0,
                 "token_source": "response_cache", "saved_prompt_tokens": 800, "saved_completion_tokens": 100, "file": self.book.name, "account": "AR"},
                {"time": iso(60), "stage": "Generator", "provider": "dashscope", "model": "m", "ok": True, "prompt_tokens": 9_999_999, "completion_tokens": 1,
                 "token_source": "provider", "file": "SomeOtherClient.xlsx", "account": "X", "entity": "OtherClient"},
                {"time": iso(-86400), "stage": "Generator", "provider": "dashscope", "model": "m", "ok": True, "prompt_tokens": 7_777_777, "completion_tokens": 1,
                 "token_source": "provider", "file": self.book.name},
                {"time": iso(60), "stage": "Generator", "provider": "dashscope", "model": "m", "ok": True, "prompt_tokens": 5_555_555, "completion_tokens": 1,
                 "token_source": "provider"},
            ]
            self.ledger.write_text("".join(json.dumps(r, ensure_ascii=False) + "\n" for r in rows), encoding="utf-8")
        self.decks = self.tmp / "decks"
        self.decks.mkdir()
        self.deck = self.decks / "accepted.pptx"
        if deck:
            self.deck.write_bytes(b"pptx bytes")
            Path(str(self.deck) + ".lineage.json").write_text(json.dumps({"binding": {"workbook_sha256": sha(self.book)}}), encoding="utf-8")
            Path(str(self.deck) + ".delivery.json").write_text(json.dumps({"status": "ready"}), encoding="utf-8")

    def args(self, *extra, deep=False):
        argv = ["--databook", str(self.book), "--logs-dir", str(self.logs), "--ledger", str(self.ledger), "--deck-dir", str(self.decks)]
        argv += list(extra) + ([] if deep else ["--no-deep"])
        return B.make_parser().parse_args(argv)

    def plan(self, *extra, **kw):
        return B.build_plan(self.args(*extra, **kw))


def run_main(world: World, *extra, patch_deep=None, config=None) -> tuple:
    argv = ["--databook", str(world.book), "--logs-dir", str(world.logs), "--ledger", str(world.ledger), "--deck-dir", str(world.decks)] + list(extra)
    out = io.StringIO()
    with contextlib.redirect_stdout(out), mock.patch.object(B, "load_config", lambda: FAKE_CONFIG if config is None else config):
        if patch_deep is not None:
            with mock.patch.object(B, "deep_check", patch_deep):
                code = B.main(argv)
        else:
            code = B.main(argv)
    return code, out.getvalue()


class Pairing(unittest.TestCase):
    def test_a_run_whose_identity_names_the_databook_hash_is_the_baseline_and_a_complete_one(self):
        w = World()
        plan = w.plan()
        self.assertEqual(plan["status"], "complete", plan["missing"] + plan["problems"])
        self.assertEqual(plan["run"]["id"], "run_20260101_100000")
        self.assertEqual((plan["proof"]["status"], plan["proof"]["kind"]), ("verified", "identity"))
        self.assertEqual(plan["run"]["kind"], "real")
        self.assertEqual(Path(plan["deck"]["path"]).name, "accepted.pptx")
        self.assertEqual(plan["missing"], [])

    def test_a_databook_with_other_bytes_is_not_paired_by_name_size_or_anything_but_its_hash(self):
        w = World()
        rebuilt = w.tmp / w.book.name.replace("Databook", "Databook_rebuilt")
        rebuilt.write_bytes(b"original databook bytes, re-saved")             # same look, other bytes: a text-rebuilt or re-saved copy
        argv = B.make_parser().parse_args(["--databook", str(rebuilt), "--logs-dir", str(w.logs), "--ledger", str(w.ledger), "--no-deep"])
        plan = B.build_plan(argv)
        self.assertEqual(plan["status"], "not_a_baseline")
        self.assertTrue(any("no archived run names this databook's hash" in p for p in plan["problems"]), plan["problems"])
        self.assertTrue(any("re-saved, rebuilt from a text export, a demo copy" in p for p in plan["problems"]))
        self.assertEqual(plan["missing"][0]["piece"], "the archived run for this databook")
        self.assertIn("list_archived_runs.py", plan["missing"][0]["how_to_get"])

    def test_naming_the_run_by_hand_does_not_get_past_the_identity_check(self):
        w = World()
        w.book.write_bytes(b"the same file saved again")                       # the run names the earlier bytes
        plan = w.plan("--run", "run_20260101_100000")
        self.assertEqual(plan["status"], "not_a_baseline")
        self.assertEqual(plan["proof"]["status"], "mismatch")
        self.assertTrue(any("other bytes" in r for r in plan["proof"]["reasons"]))

    def test_a_run_with_no_identity_file_is_not_accepted_without_the_per_account_proof(self):
        w = World(identity=False)
        plan = w.plan()
        self.assertEqual(plan["status"], "not_a_baseline")                      # nothing names the hash, so no run is paired
        plan = w.plan("--run", "run_20260101_100000")
        self.assertEqual(plan["status"], "not_a_baseline")
        self.assertEqual(plan["proof"]["status"], "mismatch")
        self.assertIn("the run has no identity file", plan["proof"]["reasons"])
        gap = next(m for m in plan["missing"] if m["piece"] == "proof that this run was written from this databook")
        self.assertIn("run_identity.json", gap["how_to_get"])
        self.assertIn("checkpoint.jsonl", gap["how_to_get"])

    def test_the_deep_proof_decides_and_anything_short_of_verified_is_not_a_baseline(self):
        w = World()
        for status in ("edited_since", "unverified", "mismatch", "error"):
            with mock.patch.object(B, "deep_check", lambda *a, _s=status, **k: {"status": _s, "reasons": ["why " + _s], "accounts": {}}):
                plan = w.plan(deep=True)
            self.assertEqual(plan["status"], "not_a_baseline", status)
            self.assertTrue(any(status in p for p in plan["problems"]), plan["problems"])
        with mock.patch.object(B, "deep_check", lambda *a, **k: {"status": "verified", "reasons": [], "accounts": {"in_run": 3, "matched": 3}}):
            plan = w.plan(deep=True)
        self.assertEqual(plan["status"], "complete")

    def test_a_run_with_no_identity_can_be_shown_to_describe_the_databook_by_fingerprints(self):
        w = World(identity=False)
        with mock.patch.object(B, "deep_check", lambda *a, **k: {"status": "verified", "reasons": [], "accounts": {"in_run": 3, "matched": 3}}):
            plan = w.plan("--run", "run_20260101_100000", deep=True)
        self.assertEqual((plan["status"], plan["proof"]["kind"]), ("complete" if not plan["missing"] else "usable_with_gaps", "fingerprints"))

    def test_several_runs_naming_the_databook_pick_a_real_run_with_most_accounts_and_list_the_rest(self):
        w = World()
        import shutil
        later = w.logs / "run_20260102_100000"
        shutil.copytree(w.run, later)
        import yaml
        (later / "results.yml").write_text(yaml.safe_dump({"AR": {"final": "x" * 30}}), encoding="utf-8")       # one account only
        plan = w.plan()
        self.assertEqual(plan["run"]["id"], "run_20260101_100000")
        self.assertEqual([a["run"] for a in plan["alternatives"]], ["run_20260102_100000"])

    def test_a_stub_run_is_not_a_model_baseline(self):
        same = "as at2025-12-31,balance为100.0万元,主要包括甲公司50.0ten thousand RMB."
        w = World(texts={"AR": same, "AP": same, "Cash": same}, ledger=False)
        plan = w.plan()
        self.assertEqual(plan["status"], "not_a_baseline")
        self.assertTrue(any("looks like a stub run" in p for p in plan["problems"]))


class Missing(unittest.TestCase):
    def test_what_is_absent_is_listed_with_why_it_matters_and_how_to_get_it(self):
        w = World(evidence=False, summaries=False, deck=False, ledger=False)
        plan = w.plan()
        self.assertEqual(plan["status"], "usable_with_gaps")
        pieces = {m["piece"] for m in plan["missing"]}
        self.assertTrue({"run/evidence/", "run/section_summaries.json", "the accepted PPTX with its lineage", "fdd_utils/logs/usage.jsonl"} <= pieces, pieces)
        for m in plan["missing"]:
            self.assertTrue(m["why_it_matters"] and m["how_to_get"], m)
        self.assertEqual(plan["alternatives"], [])

    def test_a_ledger_that_exists_but_holds_nothing_for_this_run_is_a_gap_not_a_silent_zero(self):
        w = World()
        w.ledger.write_text(json.dumps({"time": iso(10), "file": "Other.xlsx", "ok": True, "prompt_tokens": 5}) + "\n", encoding="utf-8")
        plan = w.plan()
        self.assertIn("the ledger lines of the run", {m["piece"] for m in plan["missing"]})
        self.assertEqual(plan["status"], "usable_with_gaps")

    def test_a_deck_with_no_lineage_cannot_be_paired_and_is_counted_not_guessed(self):
        w = World(deck=False)
        (w.decks / "some_deck.pptx").write_bytes(b"x")
        plan = w.plan()
        self.assertIsNone(plan["deck"])
        self.assertEqual(plan["decks_seen_without_lineage"], 1)
        gap = next(m for m in plan["missing"] if m["piece"].startswith("the accepted PPTX"))
        self.assertIn("1 deck(s) were seen with no sidecar", gap["how_to_get"])
        # a deck given by hand is used, its missing sidecars are listed
        plan = w.plan("--deck", str(w.decks / "some_deck.pptx"))
        self.assertEqual(Path(plan["deck"]["path"]).name, "some_deck.pptx")
        self.assertEqual({m["piece"] for m in plan["missing"] if m["piece"].startswith("deck")}, {"deck.lineage.json", "deck.delivery.json"})

    def test_a_databook_whose_name_says_it_is_a_copy_is_usable_but_never_complete_and_says_why(self):
        for name in ("Client_rebuilt.xlsx", "[DEMO]Client.xlsx", "Client copy.xlsx"):
            self.assertTrue(B.derived_name_hint(name), name)
        self.assertFalse(B.derived_name_hint("Client databook.xlsx"))
        w = World()
        copy = w.tmp / "Client_rebuilt.xlsx"
        copy.write_bytes(w.book.read_bytes())                                    # same bytes the run names, under a name that says it is a copy
        plan = B.build_plan(B.make_parser().parse_args(["--databook", str(copy), "--logs-dir", str(w.logs), "--ledger", str(w.ledger), "--deck-dir", str(w.decks), "--no-deep"]))
        self.assertEqual(plan["status"], "usable_with_gaps")
        self.assertIn("the databook as the client sent it", {m["piece"] for m in plan["missing"]})
        out = w.tmp / "pack"
        self.assertTrue(B.write_pack(plan, out, config=FAKE_CONFIG)["ok"])
        manifest = json.loads((out / "MANIFEST.json").read_text(encoding="utf-8"))
        self.assertIn("cannot be told from the file", manifest["databook"]["nature"])
        self.assertIn("suggests a rebuild", manifest["databook"]["nature"])

    def rebuilt(self, w, source_sha=None):
        copy = w.tmp / "Rebuilt.xlsx"
        copy.write_bytes(b"a rebuild: other bytes")
        Path(str(copy) + ".rebuilt.json").write_text(json.dumps({"rebuilt_from_a_text_dump": True, "source_sha256": source_sha or sha(w.book)}), encoding="utf-8")
        return copy

    def plan_for(self, w, book, *extra):
        argv = ["--databook", str(book), "--logs-dir", str(w.logs), "--ledger", str(w.ledger), "--deck-dir", str(w.decks), *extra]
        return B.build_plan(B.make_parser().parse_args(argv))

    def test_a_rebuild_is_found_by_the_file_it_was_rebuilt_from_but_only_the_account_proof_can_make_it_a_baseline(self):
        w = World()
        copy = self.rebuilt(w)
        plan = self.plan_for(w, copy, "--no-deep")                              # the hash alone cannot show a rebuild is the run's workbook
        self.assertEqual(plan["status"], "not_a_baseline")
        self.assertEqual(plan["run"]["id"], "run_20260101_100000", "the run was found through the marker's source hash")
        self.assertTrue(any("a rebuild from a text dump" in r for r in plan["proof"]["reasons"]), plan["proof"])
        with mock.patch.object(B, "deep_check", lambda *a, **k: {"status": "verified", "reasons": [], "accounts": {"in_run": 3, "matched": 3}}):
            plan = self.plan_for(w, copy)
        self.assertEqual(plan["status"], "usable_with_gaps")                    # a rebuild is never "complete"
        self.assertEqual(plan["proof"]["kind"], "fingerprints")
        self.assertIn("the original databook bytes", {m["piece"] for m in plan["missing"]})
        out = w.tmp / "pack"
        self.assertTrue(B.write_pack(plan, out, config=FAKE_CONFIG)["ok"])
        nature = json.loads((out / "MANIFEST.json").read_text(encoding="utf-8"))["databook"]["nature"]
        self.assertIn("REBUILT from a text dump", nature)
        self.assertIn("not the original", nature)

    def test_a_workbook_saved_again_but_matching_every_fingerprint_is_labelled_as_fingerprints_not_identity(self):
        w = World()
        w.book.write_bytes(b"the same data saved again")
        with mock.patch.object(B, "deep_check", lambda *a, **k: {"status": "verified", "reasons": [], "accounts": {"in_run": 3, "matched": 3}}):
            plan = w.plan("--run", "run_20260101_100000", deep=True)
        self.assertEqual(plan["proof"]["kind"], "fingerprints")
        self.assertIn("bytes differ", plan["proof"]["note"])
        self.assertEqual(plan["status"], "usable_with_gaps")
        self.assertIn("the original databook bytes", {m["piece"] for m in plan["missing"]})

    def test_no_databook_is_not_a_baseline(self):
        w = World()
        w.book.unlink()
        self.assertEqual(w.plan()["status"], "not_a_baseline")


class Ledger(unittest.TestCase):
    def test_only_this_files_lines_inside_the_runs_window_are_kept_and_error_text_is_cut_to_its_type(self):
        w = World()
        plan = w.plan()
        rows = plan["ledger"]["rows"]
        self.assertEqual(len(rows), 5)
        blob = json.dumps(rows)
        self.assertNotIn("Other", blob)
        self.assertNotIn("gateway.internal", blob)                              # the error text carried an address
        self.assertEqual([r["error_type"] for r in rows if "error_type" in r], ["APIConnectionError"])
        self.assertNotIn("9999999", blob)
        self.assertNotIn("7777777", blob)
        self.assertNotIn("5555555", blob)

    def test_provider_reported_estimated_and_response_cache_numbers_are_kept_apart(self):
        s = B.summarise_ledger(World().plan()["ledger"]["rows"])
        self.assertEqual(s["provider_reported"], {"calls": 2, "prompt": 3000, "completion": 500, "cached_prompt": 100, "reasoning": 50})
        self.assertEqual(B.total_provider_tokens(s), 3500)
        self.assertEqual(s["calls_without_provider_usage"], 2)                  # the estimate-only call and the failed one
        self.assertEqual(s["calls_failed"], 1)
        self.assertEqual(s["response_cache_hits"], 1)
        self.assertEqual(s["response_cache_saved"], {"prompt": 800, "completion": 100})
        self.assertEqual(s["by_stage"]["Generator"]["calls"], 3)
        self.assertIn("SDK", s["unmeasured"])

    def test_the_ledger_cannot_confirm_a_figure_it_does_not_hold(self):
        total = B.total_provider_tokens(B.summarise_ledger(World().plan()["ledger"]["rows"]))
        self.assertFalse(800_000 <= total <= 1_000_000)                          # the check a person makes against "800k to 1M per entity"


class Pack(unittest.TestCase):
    def write(self, world, *extra, config=None):
        out = world.tmp / "pack"
        plan = world.plan(*extra)
        return out, plan, B.write_pack(plan, out, config=FAKE_CONFIG if config is None else config)

    def test_a_pack_holds_the_whitelist_and_never_the_config_the_log_or_the_prompts(self):
        w = World()
        out, plan, result = self.write(w)
        self.assertTrue(result["ok"], result)
        names = {str(p.relative_to(out)) for p in out.rglob("*") if p.is_file()}
        self.assertIn("databook/ClientNameDatabook.xlsx", names)
        self.assertIn("run/run_20260101_100000/results.yml", names)
        self.assertIn("run/run_20260101_100000/run_identity.json", names)
        self.assertIn("run/run_20260101_100000/evidence/AR.json", names)
        self.assertIn("run/run_20260101_100000/section_summaries.json", names)
        self.assertIn("deck/accepted.pptx", names)
        self.assertIn("deck/accepted.pptx.lineage.json", names)
        self.assertIn("usage/usage_slice.jsonl", names)
        self.assertIn("MANIFEST.json", names)
        for forbidden in ("config.yml", "processing.log", "data.yml", "audit.jsonl"):
            self.assertFalse([n for n in names if n.endswith(forbidden)], forbidden)
        self.assertTrue(B.verify_pack(out, config=FAKE_CONFIG)["ok"])
        manifest = json.loads((out / "MANIFEST.json").read_text(encoding="utf-8"))
        self.assertEqual(manifest["status"], "complete")
        self.assertEqual(manifest["databook"]["sha256"], sha(w.book))
        self.assertEqual(manifest["tokens_provider_reported_total"], 3500)
        self.assertTrue(any("text export" in x for x in manifest["not_included"]))
        self.assertTrue(all(e["sha256"] for e in manifest["files"]))

    def test_the_stage_prompts_come_only_on_request(self):
        w = World()
        out, _, result = self.write(w, "--include-prompts")
        self.assertTrue(result["ok"])
        self.assertTrue((out / "run/run_20260101_100000/data.yml").exists())
        self.assertFalse((out / "run/run_20260101_100000/processing.log").exists())

    def test_a_key_in_a_file_to_be_packed_refuses_the_pack_and_names_the_file_and_kind_not_the_value(self):
        w = World()
        (w.run / "summary.txt").write_text(f"note: api_key = {FAKE_SK}\n", encoding="utf-8")
        out, plan, result = self.write(w)
        self.assertFalse(result["ok"])
        self.assertTrue(result["refused"])
        self.assertTrue(any(str(name).endswith("summary.txt") for name, _ in result["hits"]), result["hits"])
        kinds = {k for _, k in result["hits"]}
        self.assertTrue({"an OpenAI-style key", "an api key assignment"} & kinds, kinds)
        self.assertNotIn(FAKE_SK, json.dumps(result))
        self.assertFalse(out.exists(), "nothing was written")

    def test_a_value_taken_from_the_config_in_use_is_found_even_when_it_looks_like_nothing(self):
        w = World()
        (w.run / "summary.txt").write_text("the model said KEY-VALUE-IN-USE-1234 once\n", encoding="utf-8")
        out, plan, result = self.write(w)
        self.assertFalse(result["ok"])
        self.assertIn("a value taken from the config in use", {k for _, k in result["hits"]})
        self.assertFalse(out.exists())

    def test_a_gateway_address_in_the_evidence_refuses_unless_it_was_looked_at(self):
        w = World()
        (w.run / "evidence" / "AR.json").write_text(json.dumps({"note": "see https://portal.example.org/files/1"}), encoding="utf-8")
        out, plan, result = self.write(w)
        self.assertFalse(result["ok"])
        self.assertEqual({k for _, k in result["hits"]}, {"an https address"})
        try:
            B.ALLOW_ADDRESSES = True
            _, _, again = self.write(World_with_url(), config={})
            self.assertTrue(again["ok"])
        finally:
            B.ALLOW_ADDRESSES = False

    def test_the_binary_databook_is_not_searched_it_is_the_clients_own_file(self):
        w = World(databook_bytes=("sk" + "-abcdefghijklmnopqrstuvwxyz0123 inside a workbook").encode())
        _, plan, result = self.write(w)
        self.assertTrue(result["ok"], result)

    def test_the_databook_can_be_left_out_when_it_travels_another_way_and_the_pack_still_verifies(self):
        w = World()
        out = w.tmp / "pack"
        result = B.write_pack(w.plan(), out, config=FAKE_CONFIG, leave_out_databook=True)
        self.assertTrue(result["ok"])
        self.assertFalse((out / "databook").exists())
        manifest = json.loads((out / "MANIFEST.json").read_text(encoding="utf-8"))
        self.assertFalse(manifest["databook"]["packed"])
        self.assertEqual(manifest["databook"]["sha256"], sha(w.book))
        self.assertTrue(B.verify_pack(out, config=FAKE_CONFIG)["ok"])

    def test_an_output_folder_that_is_not_empty_is_not_written_over(self):
        w = World()
        out = w.tmp / "pack"
        out.mkdir()
        (out / "old.txt").write_text("x", encoding="utf-8")
        result = B.write_pack(w.plan(), out, config=FAKE_CONFIG)
        self.assertFalse(result["ok"])
        self.assertIn("not empty", result["reason"])

    def test_a_location_git_would_track_is_refused(self):
        with self.assertRaises(SystemExit) as caught:
            B.check_out_path(B.ROOT / "notes_for_me")
        self.assertIn("not git-ignored", str(caught.exception))
        B.check_out_path(B.ROOT / "for_test" / "baseline-pack")
        B.check_out_path(Path(tempfile.mkdtemp()))

    def test_where_git_cannot_answer_only_for_test_is_accepted(self):
        """In a copied or downloaded tree `git check-ignore` exits 128: no answer, so fall back to the documented ignored folder."""
        import subprocess
        from unittest import mock

        def git_says(code):
            return mock.patch.object(B.subprocess, "run", return_value=subprocess.CompletedProcess([], code))
        with git_says(128):
            B.check_out_path(B.ROOT / "for_test" / "baseline-pack")
            with self.assertRaises(SystemExit):
                B.check_out_path(B.ROOT / "notes_for_me")
        with git_says(1):
            with self.assertRaises(SystemExit):
                B.check_out_path(B.ROOT / "for_test" / "baseline-pack")              # git's word wins over the folder rule
        with git_says(0):
            B.check_out_path(B.ROOT / "notes_for_me")
        with mock.patch.object(B.subprocess, "run", side_effect=FileNotFoundError("git")):
            B.check_out_path(B.ROOT / "for_test" / "baseline-pack")
            with self.assertRaises(SystemExit):
                B.check_out_path(B.ROOT / "notes_for_me")


def World_with_url():
    w = World()
    (w.run / "evidence" / "AR.json").write_text(json.dumps({"note": "see https://portal.example.org/files/1"}), encoding="utf-8")
    return w


class Verify(unittest.TestCase):
    def pack(self):
        w = World()
        out = w.tmp / "pack"
        self.assertTrue(B.write_pack(w.plan(), out, config=FAKE_CONFIG)["ok"])
        return w, out

    def test_a_pack_that_was_changed_a_file_that_was_added_or_one_that_must_never_be_there_fails(self):
        w, out = self.pack()
        self.assertTrue(B.verify_pack(out, config=FAKE_CONFIG)["ok"])
        (out / "run/run_20260101_100000/results.yml").write_text("AR: {final: changed}\n", encoding="utf-8")
        result = B.verify_pack(out, config=FAKE_CONFIG)
        self.assertFalse(result["ok"])
        self.assertTrue(any("changed since it was packed" in p for p in result["problems"]))
        w, out = self.pack()
        (out / "config.yml").write_text("api" + "_key: " + "nothing-secret-here-but-it-is-a-config\n", encoding="utf-8")
        result = B.verify_pack(out, config=FAKE_CONFIG)
        self.assertTrue(any("must never be packed" in p for p in result["problems"]), result)
        self.assertTrue(any("not in the manifest" in p for p in result["problems"]), result)

    def test_a_secret_that_arrived_after_packing_is_found_by_verify(self):
        w, out = self.pack()
        path = out / "run/run_20260101_100000/summary.txt"
        manifest = json.loads((out / "MANIFEST.json").read_text(encoding="utf-8"))
        path.write_text(f"Authorization: Bearer {FAKE_BEARER_TOKEN}\n", encoding="utf-8")
        for e in manifest["files"]:
            if e["path"].endswith("summary.txt"):
                e["sha256"] = sha(path)
        (out / "MANIFEST.json").write_text(json.dumps(manifest), encoding="utf-8")
        result = B.verify_pack(out, config=FAKE_CONFIG)
        self.assertTrue(any("a bearer token" in p for p in result["problems"]), result)

    def test_swapping_in_a_different_databook_is_caught_even_if_the_manifest_is_rewritten(self):
        w, out = self.pack()
        book = out / "databook" / w.book.name
        book.write_bytes(b"a rebuilt databook")
        manifest = json.loads((out / "MANIFEST.json").read_text(encoding="utf-8"))
        for e in manifest["files"]:
            if e["path"].startswith("databook/"):
                e["sha256"] = sha(book)
        (out / "MANIFEST.json").write_text(json.dumps(manifest), encoding="utf-8")
        result = B.verify_pack(out, config=FAKE_CONFIG)
        self.assertFalse(result["ok"])
        self.assertTrue(any("not the one the manifest names" in p for p in result["problems"]), result)


class Command(unittest.TestCase):
    def test_the_printed_summary_carries_no_file_entity_or_account_names_and_the_dry_run_writes_nothing(self):
        w = World()
        code, text = run_main(w, "--no-deep")
        self.assertEqual(code, 0, text)
        self.assertIn("baseline: COMPLETE", text)
        self.assertIn("provider-reported tokens: 3,000 in + 500 out = 3,500", text)
        self.assertIn("estimates are not added", text)
        for private in ("ClientNameDatabook", "ClientEntity", "OtherClient", "应收账款", "gateway.internal"):
            self.assertNotIn(private, text)
        self.assertIn("Dry run", text)
        self.assertFalse((w.tmp / "pack").exists())
        _, named = run_main(w, "--no-deep", "--names")
        self.assertIn("ClientNameDatabook", named)

    def test_exit_codes_are_0_complete_1_with_gaps_2_not_a_baseline(self):
        self.assertEqual(run_main(World(), "--no-deep")[0], 0)
        self.assertEqual(run_main(World(evidence=False), "--no-deep")[0], 1)
        w = World()
        w.book.write_bytes(b"another file")
        code, text = run_main(w, "--no-deep")
        self.assertEqual(code, 2)
        self.assertIn("Nothing was written", text)

    def test_writing_through_the_command_ends_with_a_verify_and_the_pack_is_usable(self):
        w = World()
        out = B.ROOT / "for_test" / f"pack-test-{os.getpid()}"
        try:
            code, text = run_main(w, "--no-deep", "--out", str(out))
            self.assertEqual(code, 0, text)
            self.assertIn("verify: PASS", text)
            self.assertNotIn(str(out), text)                                      # the path is not printed either unless --names
            self.assertTrue(B.verify_pack(out, config=FAKE_CONFIG)["ok"])
            code, text = run_main(w, "--no-deep", "--out", str(out))              # a second write onto a non-empty folder
            self.assertEqual(code, 2)
            self.assertIn("not empty", text)
        finally:
            import shutil
            shutil.rmtree(out, ignore_errors=True)

    def test_a_secret_stops_the_command_and_it_says_which_kind_in_which_file_without_the_value(self):
        w = World()
        (w.run / "summary.txt").write_text(f"Bearer {FAKE_BEARER_TOKEN}\n", encoding="utf-8")
        out = B.ROOT / "for_test" / f"pack-test-{os.getpid()}"
        code, text = run_main(w, "--no-deep", "--out", str(out))
        self.assertEqual(code, 2)
        self.assertIn("REFUSED", text)
        self.assertIn("a bearer token in", text)
        self.assertNotIn(FAKE_BEARER_TOKEN, text)
        self.assertFalse(out.exists())

    def test_verify_mode_reports_pass_or_fail(self):
        w = World()
        out = w.tmp / "pack"
        B.write_pack(w.plan(), out, config=FAKE_CONFIG)
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf), mock.patch.object(B, "load_config", lambda: FAKE_CONFIG):
            self.assertEqual(B.main(["--verify", str(out)]), 0)
            (out / "README.txt").write_text("edited", encoding="utf-8")            # README is not hashed: it holds no data
            self.assertEqual(B.main(["--verify", str(out)]), 0)
            (out / "extra.txt").write_text("x", encoding="utf-8")
            self.assertEqual(B.main(["--verify", str(out)]), 2)
        self.assertIn("PASS", buf.getvalue())
        self.assertIn("FAIL", buf.getvalue())


class Carrying(unittest.TestCase):
    """The two ways material gets from the machine that has it to the one that reviews it, end to end."""

    def test_the_exact_route_a_pack_goes_through_text_and_still_verifies_with_the_same_databook_hash(self):
        import text_transfer as X
        w = World()
        pack = w.tmp / "pack"
        self.assertTrue(B.write_pack(w.plan(), pack, config=FAKE_CONFIG)["ok"])
        text = w.tmp / "text"
        X.encode([str(pack)], text, 20)
        got = w.tmp / "got"
        self.assertTrue(X.decode([str(text)], got)["ok"])
        self.assertTrue(B.verify_pack(got / "pack", config=FAKE_CONFIG)["ok"])
        self.assertEqual(sha(got / "pack" / "databook" / w.book.name), sha(w.book))
        manifest = json.loads((got / "pack" / "MANIFEST.json").read_text(encoding="utf-8"))
        self.assertEqual(manifest["proof"]["kind"], "identity")

    def test_the_fallback_route_a_dump_for_the_databook_and_a_pack_without_it_ends_as_a_labelled_rebuild(self):
        import databook_dump as D
        import text_transfer as X
        from test_databook_dump import make_book
        real = Path(tempfile.mkdtemp()) / "real.xlsx"
        make_book(real)
        w = World(databook_bytes=real.read_bytes())
        # the machine that has the files
        pack = w.tmp / "pack"
        self.assertTrue(B.write_pack(w.plan(), pack, config=FAKE_CONFIG, leave_out_databook=True)["ok"])
        records, _ = D.dump_workbook(str(w.book))
        dump = w.tmp / "dump.txt"
        D.write_dump(records, dump)
        X.encode([str(pack)], w.tmp / "text", 20)
        # the machine that reviews
        got = w.tmp / "got"
        self.assertTrue(X.decode([str(w.tmp / "text")], got)["ok"])
        rebuilt = w.tmp / "Rebuilt.xlsx"
        D.restore_workbook(D.parse_dump(dump), rebuilt)
        self.assertEqual(D.dump_workbook(str(rebuilt))[1]["digest"], D.dump_workbook(str(w.book))[1]["digest"])
        argv = ["--databook", str(rebuilt), "--logs-dir", str(got / "pack" / "run"), "--ledger", str(got / "pack" / "usage" / "usage_slice.jsonl"),
                "--deck-dir", str(got / "pack" / "deck")]
        with mock.patch.object(B, "deep_check", lambda *a, **k: {"status": "verified", "reasons": [], "accounts": {"in_run": 3, "matched": 3}}):
            plan = B.build_plan(B.make_parser().parse_args(argv))
        self.assertEqual(plan["run"]["id"], "run_20260101_100000")
        self.assertEqual(plan["status"], "usable_with_gaps")
        self.assertEqual(plan["databook"]["rebuilt"]["source_sha256"], sha(w.book))
        self.assertEqual(plan["ledger"]["summary"]["provider_reported"]["calls"], 2)
        self.assertEqual(Path(plan["deck"]["path"]).name, "accepted.pptx")
        self.assertIn("the original databook bytes", {m["piece"] for m in plan["missing"]})


class RunKind(unittest.TestCase):
    def test_the_ledger_and_not_the_runs_own_calls_ok_line_says_whether_a_model_was_called(self):
        w = World()
        start, end = B.run_window(w.run)
        rows = B.ledger_rows(w.ledger, {w.book.name}, start, end)
        self.assertEqual(B.run_kind(w.run, rows), "real")
        self.assertEqual(B.run_kind(w.run, []), "unknown")                       # "calls ok 9" in summary.txt proves nothing
        same = "x" * 40
        stub = World(texts={"A": same, "B": same, "C": same}, ledger=False)
        self.assertEqual(B.run_kind(stub.run, []), "stub-like")


if __name__ == "__main__":
    unittest.main(verbosity=2)

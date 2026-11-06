"""The connected workflow pieces added for the company acceptance -- offline, scripted models, no provider.

    python ad-hoc/workbench/test_agent_workflow.py

Accounting (reasoning inside completion, the reported model id, retries, output allowance), the three configurations on the same
request (Flash-only, Pro-only, Flash->Pro), task kinds that start on the escalation model, the dry-run request preview, the new
reading tools (component pages, mapping explanation, value provenance), manual mapping overrides (user-only, persisted, rebinding
drafts), the deck state and partial updates, the first-generation path, and claims that try to pass by relabelling themselves.
"""
from __future__ import annotations

import json
import os
import stat
import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parents[1]))

from fdd_utils.agent import controller as C  # noqa: E402
from fdd_utils.agent import routing as R  # noqa: E402
from fdd_utils.agent import tools as T  # noqa: E402
from fdd_utils.agent import workspace as W  # noqa: E402
from fdd_utils.agent.session import capture_context, context_changes  # noqa: E402
from fdd_utils.ai import claims as cl  # noqa: E402
from test_agent_controller import Policy, evidence_of, tool, total_ref  # noqa: E402
from test_workspace import ACCOUNT, P1, P2, answer_policy, build, session_on, write_book  # noqa: E402


def reply(content="{}", **usage):
    def call(system, user, max_tokens):
        call.max_tokens = max_tokens
        return {"content": content, **usage}
    return call


def spec(role, call, model="m", **kw):
    return R.RoleSpec(role, call, provider="scripted", model=model, **kw)


# -- the books ----------------------------------------------------------------------------------------------------------------------

class Books(unittest.TestCase):
    def test_reasoning_inside_completion_is_shown_and_not_added(self):
        router = R.RoutedModel(spec("primary", reply("x", prompt_tokens=100, completion_tokens=50, reasoning_tokens=30)))
        router("s", "u", 200)
        record = router.budget.records[0]
        self.assertEqual((record.charged, record.reasoning, record.reasoning_separate), (150, 30, False))
        self.assertIn("shown, not added", router.budget.summary()["reasoning_basis"])

    def test_a_provider_that_counts_reasoning_outside_completion_has_it_added_once(self):
        router = R.RoutedModel(spec("primary", reply("x", prompt_tokens=100, completion_tokens=10, reasoning_tokens=40)))
        router("s", "u", 200)
        record = router.budget.records[0]
        self.assertTrue(record.reasoning_separate)
        self.assertEqual(record.charged, 150)

    def test_the_model_the_provider_says_answered_is_recorded_beside_the_configured_one(self):
        router = R.RoutedModel(spec("primary", reply("x", prompt_tokens=1, completion_tokens=1, model_reported="gw-model-2026-08-13"), model="gw-model"))
        router("s", "u", 10)
        record = router.budget.records[0]
        self.assertEqual((record.model, record.model_reported), ("gw-model", "gw-model-2026-08-13"))
        self.assertEqual(router.budget.summary()["by_role"]["primary"]["models_reported"], ["gw-model-2026-08-13"])

    def test_retries_are_counted_by_why_they_happened(self):
        calls = {"n": 0}

        def flaky(system, user, max_tokens):
            calls["n"] += 1
            if calls["n"] == 1:
                raise RuntimeError("gateway timeout at https://gw.example/v1")
            return {"content": "x"}
        router = R.RoutedModel(spec("primary", flaky))
        router.begin_turn(1)
        with self.assertRaises(RuntimeError):
            router("s", "u", 10)
        router("s", "u", 10)                                           # after a failed call
        router.mark_last_reply_unused()
        router("s", "u", 10)                                           # after a discarded reply
        router.note_rejection()
        router("s", "u", 10)                                           # after the gate sent the answer back
        afters = [r.after for r in router.budget.records]
        self.assertEqual(afters, ["", "failed", "discarded", "rejected"])
        self.assertEqual(router.budget.summary()["retries"], {"after_failed_call": 1, "after_discarded_reply": 1, "after_rejection": 1})
        self.assertNotIn("gw.example", json.dumps(router.budget.to_dict()))

    def test_a_role_output_allowance_caps_the_call_and_is_what_the_cap_reserves(self):
        call = reply("x")
        router = R.RoutedModel(spec("primary", call, max_output=300), budget=R.SharedBudget(R.Limits(max_tokens=400)))
        router("s", "u", 1500)
        self.assertEqual((call.max_tokens, router.budget.records[0].max_output), (300, 300))
        with self.assertRaises(R.BudgetExhausted):                     # 300 reserved again on top of what was charged
            R.RoutedModel(spec("primary", reply("x"), max_output=600), budget=R.SharedBudget(R.Limits(max_tokens=400)))("s", "u", 1500)
        self.assertIn("not a hard limit", router.budget.summary()["reservation"])


# -- routing: configured task kinds, and the three configurations on the same request -----------------------------------------------

class Configurations(unittest.TestCase):
    def test_a_configured_task_kind_starts_on_the_escalation_model_within_the_caps(self):
        router = R.RoutedModel(spec("primary", reply()), spec("escalation", reply()), policy=R.Policy(mode="escalate", escalate_tasks=["adjustment_bridge"]))
        started = router.begin_turn(1, R.task_kinds("请解释应收账款的调整差异"))
        self.assertEqual((router.role, started["reason"]), ("escalation", "configured_task"))
        self.assertIsNone(router.begin_turn(2, R.task_kinds("应收账款余额是多少？")))
        self.assertEqual(router.role, "primary")
        capped = R.RoutedModel(spec("primary", reply()), spec("escalation", reply()), policy=R.Policy(mode="escalate", escalate_tasks=["adjustment_bridge"]),
                               budget=R.SharedBudget(R.Limits(max_escalations=0)))
        self.assertIsNone(capped.begin_turn(1, ["adjustment_bridge"]))
        single = R.RoutedModel(spec("primary", reply()), policy=R.Policy(mode="single", escalate_tasks=["adjustment_bridge"]))
        self.assertIsNone(single.begin_turn(1, ["adjustment_bridge"]))

    def _work(self, model, escalation=None, mode="single"):
        tmp = tempfile.mkdtemp(prefix="cfg_")
        book = os.path.join(tmp, "wb.xlsx")
        write_book(book)
        source = session_on(book)
        source.load_params = {"workbook": book}
        return W.WorkSession(source, model=model, escalation=escalation, model_name="m", policy=R.Policy(mode=mode), budget=C.Budget(max_steps=8))

    def test_flash_only_pro_only_and_flash_then_pro_run_the_same_request_on_one_set_of_books(self):
        from test_routing import handoff_reader_policy, wrong_amount_policy
        request = "应收账款期末余额是多少？"
        flash = self._work(answer_policy())
        pro = self._work(answer_policy())
        routed = self._work(wrong_amount_policy(), escalation=handoff_reader_policy([]), mode="escalate")
        turns = {name: work.ask(request) for name, work in (("flash_only", flash), ("pro_only", pro), ("routing", routed))}
        for name, turn in turns.items():
            self.assertTrue(turn.ok, (name, turn.report))
        self.assertEqual(set(flash.shared.summary()["by_role"]), {"primary"})
        self.assertEqual(set(routed.shared.summary()["by_role"]), {"primary", "escalation"})
        books = routed.shared.summary()
        self.assertEqual(books["escalations"], 1)
        self.assertEqual(books["calls"], sum(v["calls"] for v in books["by_role"].values()))
        self.assertEqual(turns["routing"].report["routing"]["escalations"][0]["reason"], "capability")


    def test_in_edit_mode_failed_validations_are_not_progress_and_repeated_ones_escalate(self):
        attempts = {"n": 0}

        def wrong_each_time(user, results):
            view = evidence_of(results)
            if view is None and "TOOL get_account_evidence(" not in user:
                return tool("get_account_evidence", account=ACCOUNT)
            view = view or cl.evidence_view(self._source.contracts[ACCOUNT])
            attempts["n"] += 1
            amount = 150.0 + attempts["n"]                                   # a different wrong figure each time: a new digest, still wrong
            return tool("validate_candidate", account=ACCOUNT, claims=[{"account": ACCOUNT, "kind": "amount", "periods": [P2],
                        "cites": [total_ref(view, P2)], "text": f"2025年12月31日应收账款余额{amount:.1f}万元"}])

        def fixes_it(user, results):
            names = [n for n, _ in results]
            view = cl.evidence_view(self._source.contracts[ACCOUNT])
            if "validate_candidate" in names and any(r.get("ok") for n, r in results if n == "validate_candidate"):
                return {"done": {"account": ACCOUNT, "export": False, "summary": "fixed"}}
            return tool("validate_candidate", account=ACCOUNT, claims=[{"account": ACCOUNT, "kind": "amount", "periods": [P2],
                        "cites": [total_ref(view, P2)], "text": "2025年12月31日应收账款余额130.0万元"}])
        work = self._work(Policy(wrong_each_time), escalation=Policy(fixes_it), mode="escalate")
        self._source = work.source
        turn = work.edit("重写应收账款的段落", ACCOUNT)
        self.assertTrue(turn.ok, turn.report)
        routing = turn.report["routing"]
        self.assertEqual([e["reason"] for e in routing["escalations"]], ["capability"])
        self.assertEqual(attempts["n"], 2)                                   # two rejected drafts, then the second model


# -- the dry-run request preview and the thinking body ---------------------------------------------------------------------------

class Preview(unittest.TestCase):
    def _client(self, model_type="dashscope", details=None):
        from fdd_utils.ai.client import AIClient
        client = AIClient.__new__(AIClient)
        client.model_type, client.use_heuristic = model_type, False
        client.config_details = dict(details or {"chat_model": "gw-flash"})
        client.model = client.config_details["chat_model"]
        client.temperature, client.max_tokens, client.top_p = 0.1, 1000, None
        client.frequency_penalty = client.presence_penalty = None
        client.logger = __import__("logging").getLogger("test.preview")
        client.client = SimpleNamespace(base_url="https://gateway.internal.example.com/v1")
        return client

    def test_the_declared_thinking_body_is_sent_only_when_declared_and_asked_for(self):
        declared = self._client(details={"chat_model": "m", "thinking_extra_body": {"on": {"enable_thinking": True}, "off": {"enable_thinking": False}}})
        self.assertEqual(declared._thinking_extra_body(True, None), {"enable_thinking": True})
        self.assertEqual(declared._thinking_extra_body(False, None), {"enable_thinking": False})
        self.assertIsNone(declared._thinking_extra_body(None, None))
        self.assertIsNone(self._client()._thinking_extra_body(True, None))

    def test_the_preview_runs_the_clients_own_request_code_and_sends_nothing(self):
        from fdd_utils.agent import probe
        client = self._client(details={"chat_model": "gw-flash", "thinking_extra_body": {"on": {"enable_thinking": True}}})
        role = SimpleNamespace(client=client, provider="dashscope", requested_provider="dashscope", provider_fallback=False, chat_model="gw-flash",
                               thinking="on", reasoning_effort="", thinking_applied=True, thinking_unverified=True, max_output=700,
                               params={"allow_thinking": True})
        out = probe.preview(role)
        self.assertEqual((out["model"], out["max_output"], out["extra_body"]), ("gw-flash", 700, {"enable_thinking": True}))
        self.assertTrue(out["thinking_unverified"])
        self.assertFalse(out["sent"])
        self.assertEqual(out["endpoint_host"], "*.example.com")         # the address itself is not shown
        self.assertNotIn("internal", json.dumps(out))
        self.assertEqual(client.client.base_url, "https://gateway.internal.example.com/v1")     # the real client is put back

    def test_a_role_that_falls_back_to_another_provider_is_refused_in_the_working_session(self):
        tmp = tempfile.mkdtemp(prefix="fb_")
        book = os.path.join(tmp, "wb.xlsx")
        write_book(book)
        work = W.WorkSession(session_on(book), model_type="dashscope", config={"agent": {}})
        fake = SimpleNamespace(provider="deepseek", provider_fallback=True, chat_model="deepseek-chat", thinking_applied=False, thinking_unverified=False)
        with mock.patch.object(C, "AIClientModel", lambda **kw: fake):
            with self.assertRaisesRegex(ValueError, "fall back"):
                work._router()
            work.config = {"agent": {"allow_provider_fallback": True}}
            work._clients.clear()
            self.assertEqual(work._router().roles["primary"].provider_requested, "dashscope")


# -- the reading tools -----------------------------------------------------------------------------------------------------------

class ReadingTools(unittest.TestCase):
    def test_components_come_in_pages_with_refs_that_become_citable(self):
        rows = [(f"公司{i:02d}", "detail", {P1: 1000 * (i + 1), P2: 2000 * (i + 1)}) for i in range(14)]
        rows.append(("Total", "total", {P1: sum(1000 * (i + 1) for i in range(14)), P2: sum(2000 * (i + 1) for i in range(14))}))
        source = session_on(os.path.join(tempfile.mkdtemp(), "wb.xlsx"), rows=rows)
        view = T.call_tool(source, "get_account_evidence", {"account": ACCOUNT}).data
        self.assertGreater(view["components_not_shown"], 0)
        first = T.call_tool(source, "get_account_components", {"account": ACCOUNT, "period": P2, "limit": 10})
        self.assertTrue(first.ok, first.error)
        self.assertEqual((first.data["total_components"], first.data["returned"], first.data["next_offset"]), (14, 10, 10))
        second = T.call_tool(source, "get_account_components", {"account": ACCOUNT, "period": P2, "offset": 10, "limit": 10})
        self.assertEqual((second.data["returned"], second.data["next_offset"]), (4, None))
        refs = {c["ref"] for c in first.data["components"]} | {c["ref"] for c in second.data["components"]}
        self.assertEqual(len(refs), 14)
        self.assertTrue(set(second.basis) <= refs | set(second.basis))
        self.assertTrue(set(c["ref"] for c in second.data["components"]) <= set(second.basis))     # the controller learns these refs
        past = T.call_tool(source, "get_account_components", {"account": ACCOUNT, "offset": 99})
        self.assertEqual(past.error["code"], "NOT_FOUND")
        bad = T.call_tool(source, "get_account_components", {"account": ACCOUNT, "period": "2019-12-31"})
        self.assertEqual(bad.error["code"], "NOT_FOUND")
        self.assertEqual(first.data["source"]["source_version"], source.source_version)
        self.assertTrue(first.data["limits"]["values"])

    def test_the_mapping_explanation_names_candidates_ambiguity_and_overrides_and_changes_nothing(self):
        source = session_on(os.path.join(tempfile.mkdtemp(), "wb.xlsx"))
        source.resolution = {"resolved": {ACCOUNT: {"sheet_name": "AR", "resolution_method": "alias", "score": 96.0, "matched_alias": "应收账款"}},
                             "ambiguities": {ACCOUNT: [{"sheet_name": "AR", "score": 96.0}, {"sheet_name": "AR (2)", "score": 95.5}]},
                             "candidate_map": {}, "override_issues": [{"mapping_key": ACCOUNT, "issue_type": "invalid_override"}],
                             "unresolved_sheets": [{"sheet_name": "Misc"}]}
        source.load_params = {"mapping_overrides": {ACCOUNT: "AR"}}
        before = source.source_version
        out = T.call_tool(source, "explain_mapping", {"account": ACCOUNT})
        self.assertTrue(out.ok, out.error)
        d = out.data
        self.assertEqual((d["sheet"], d["method"], d["ambiguous"], d["runner_up_gap"]), ("AR", "alias", True, 0.5))
        self.assertEqual(d["override"], {"applied": True, "sheet": "AR"})
        self.assertEqual([c["sheet_name"] for c in d["candidates"]], ["AR", "AR (2)"])
        self.assertEqual(len(d["override_issues"]), 1)
        self.assertIn("only the user", d["limits"]["decides"])
        self.assertEqual(source.source_version, before)
        self.assertFalse([n for n in T.TOOLS if "override" in n], "no tool can change a mapping")

    def test_a_figure_found_twice_on_its_row_is_ambiguous_and_none_is_chosen(self):
        from openpyxl import Workbook
        tmp = tempfile.mkdtemp()
        path = os.path.join(tmp, "wb.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "AR"
        ws["A1"], ws["A2"] = "应收账款", "人民币千元"
        for r, (name, a, b) in enumerate([("甲公司", 700, 700), ("乙公司", 400, 300), ("合计", 1100, 1000)], 4):
            ws.cell(r, 1, name), ws.cell(r, 2, a), ws.cell(r, 3, b)
        wb.save(path)
        rows = [("甲公司", "detail", {P1: 700000, P2: 700000}), ("乙公司", "detail", {P1: 400000, P2: 300000}),
                ("Total", "total", {P1: 1100000, P2: 1000000})]
        source = session_on(path, rows=rows)
        ref = next(f["ref"] for f in source.contracts[ACCOUNT].facts if f["line"] == "甲公司" and f["period"] == P2)
        out = T.call_tool(source, "read_source_block", {"account": ACCOUNT, "ref": ref})
        self.assertTrue(out.ok, out.error)
        prov = out.data["figure_provenance"]
        self.assertEqual((prov["exact"], prov["status"], len(prov["candidates"])), (False, "ambiguous", 2))
        ref_b = next(f["ref"] for f in source.contracts[ACCOUNT].facts if f["line"] == "乙公司" and f["period"] == P2)
        one = T.call_tool(source, "read_source_block", {"account": ACCOUNT, "ref": ref_b}).data["figure_provenance"]
        self.assertEqual((one["exact"], one["status"]), (False, "one_candidate"))


# -- manual mapping overrides: the user's, persisted, rebinding drafts ------------------------------------------------------------

class MappingOverrides(unittest.TestCase):
    def test_an_override_is_the_users_changes_the_source_version_rebinds_drafts_and_survives_a_reopen(self):
        tmp = tempfile.mkdtemp(prefix="ovr_")
        book = os.path.join(tmp, "wb.xlsx")
        write_book(book)
        seen = []

        def load(cls_or_none=None, **kw):
            seen.append(dict(kw.get("mapping_overrides") or {}))
            fresh = session_on(book)
            overrides = kw.get("mapping_overrides") or {}
            fresh.binding.mapping = "ovr-" + "-".join(f"{k}={v}" for k, v in sorted(overrides.items())) if overrides else ""
            for contract in fresh.contracts.values():
                contract.binding = fresh.binding.to_dict()
            fresh.load_params = {"workbook": book, "mapping_overrides": dict(overrides)}
            return fresh
        path = os.path.join(W.DEFAULT_DIR, f"test-override-{os.getpid()}.agent-session.json")
        with mock.patch.object(W.SourceSession, "load", side_effect=lambda **kw: load(**kw)):
            work = W.WorkSession.open(book, session_path=path, model=answer_policy(), budget=C.Budget(max_steps=6))
            work.ask("应收账款期末余额是多少？")
            work.report("把这段放入报告")
            before = work.source.source_version
            self.assertTrue(work.source.candidates[ACCOUNT].accepted)
            result = work.set_mapping_override(ACCOUNT, "AR")
            self.assertTrue(result.get("ok"), result)
            self.assertNotEqual(work.source.source_version, before)
            effect = next(e for e in result["candidate_effects"] if e["account"] == ACCOUNT)
            self.assertIn(effect["status"], ("carried_changed", "invalidated"))
            self.assertFalse(work.source.candidates[ACCOUNT].accepted)          # what the person accepted was read another way
            self.assertTrue(any("mapping override" in r for r in effect["reasons"]), effect["reasons"])
            self.assertEqual(work.events[-1]["type"], "mapping_override")
            self.assertEqual(work.events[-1]["by"], "user")
            reopened = W.WorkSession.open(session_path=path, model=answer_policy())
            self.assertEqual(reopened.mapping_overrides, {ACCOUNT: "AR"})
            self.assertEqual(seen[-1], {ACCOUNT: "AR"})
            cleared = reopened.clear_mapping_override(ACCOUNT)
            self.assertTrue(cleared.get("ok"))
            self.assertEqual(reopened.mapping_overrides, {})
        os.remove(path)

    def test_the_source_version_is_unchanged_when_there_is_no_override(self):
        from fdd_utils.ai import evidence_contract as ec
        plain = ec.SourceBinding(workbook_sha256="abc", sheet="Financials")
        self.assertEqual(plain.source_version, ec.SourceBinding(workbook_sha256="abc", sheet="Financials", mapping="").source_version)
        self.assertNotEqual(plain.source_version, ec.SourceBinding(workbook_sha256="abc", sheet="Financials", mapping="x").source_version)


# -- what a draft rests on: the tab, the unit, the mapping ---------------------------------------------------------------------------

class SourceProfile(unittest.TestCase):
    def test_a_change_of_tab_unit_or_override_is_named_even_when_every_figure_is_the_same(self):
        contract = build()
        cl.standard_derived(contract)
        claims = [{"account": ACCOUNT, "kind": "amount", "periods": [P2], "cites": [total_ref(cl.evidence_view(contract), P2)],
                   "text": "2025年12月31日应收账款余额130.0万元"}]
        context = capture_context(contract, claims, {}, {"sheet": "AR", "multiplier": 1000, "mapping_method": "alias", "override": None})
        self.assertEqual(context_changes(contract, context, profile_now={"sheet": "AR", "multiplier": 1000, "mapping_method": "alias", "override": None}), [])
        moved = context_changes(contract, context, profile_now={"sheet": "AR (2)", "multiplier": 1, "mapping_method": "override", "override": "AR (2)"})
        text = " | ".join(moved)
        self.assertIn("now read from tab 'AR (2)'", text)
        self.assertIn("unit its tab is printed in changed", text)
        self.assertIn("manual mapping override changed", text)


# -- the deck, partial updates and the first deck ------------------------------------------------------------------------------------

class Decks(unittest.TestCase):
    def _work(self, **kw):
        tmp = tempfile.mkdtemp(prefix="deck_")
        book = os.path.join(tmp, "wb.xlsx")
        write_book(book)
        source = session_on(book)
        source.load_params = {"workbook": book, "entity": "甲实体"}
        return W.WorkSession(source, model=answer_policy(), budget=C.Budget(max_steps=6), **kw), tmp

    def test_without_a_base_run_the_export_says_how_to_write_the_whole_entity_first(self):
        work, _ = self._work()
        work.ask("应收账款期末余额是多少？")
        work.report("把这段放入报告")
        out = work.export()
        self.assertFalse(out["ok"])
        self.assertEqual(out["error"]["code"], "NO_BASE_COMMENTARY")
        plan = out["error"]["first_generation"]
        self.assertIn("--run-ai --export-pptx", plan["steps"][0]["command"])
        self.assertIn("--entity", plan["steps"][0]["command"])
        self.assertIn("/base-run", plan["steps"][1]["command"])
        self.assertIn("not_offered", plan)

    def test_a_base_run_that_does_not_describe_the_workbook_is_refused_at_export(self):
        work, _ = self._work()
        work.source.base_results = {ACCOUNT: {"final": "旧"}}
        work.source.base_run_check = {"status": "mismatch", "reasons": ["identity differs -- entity"]}
        work.ask("应收账款期末余额是多少？")
        work.report("把这段放入报告")
        self.assertEqual(work.export()["error"]["code"], "BASE_RUN_MISMATCH")

    def test_the_deck_state_is_kept_in_the_session_file_and_comes_back(self):
        work, tmp = self._work()
        work.read_deck(os.path.join(tmp, "no-such.pptx"))
        self.assertEqual(work.deck["read"]["error"], "NO_SUCH_FILE")
        work.deck["last_update"] = {"complete": False, "unfinished": [{"action": "conflict"}]}
        path = work.save(os.path.join(W.DEFAULT_DIR, f"test-deck-{os.getpid()}.agent-session.json"))
        data = json.load(open(path, encoding="utf-8"))
        self.assertEqual(data["deck"]["last_update"]["complete"], False)
        blob = json.dumps(data)
        self.assertNotIn("api_key", blob)
        self.assertNotIn('"config"', blob)
        if os.name != "nt":
            self.assertEqual(stat.S_IMODE(os.stat(path).st_mode), 0o600)
        restored = W.WorkSession(work.source, model=answer_policy())
        restored._restore(data)
        self.assertEqual(restored.deck["last_update"]["unfinished"], [{"action": "conflict"}])
        os.remove(path)

    def test_an_update_that_left_anything_is_partial_and_says_what_and_what_next(self):
        from fdd_utils.agent import updates as U
        from fdd_utils.pptx import lineage as L
        work, tmp = self._work()
        deck = os.path.join(tmp, "returned.pptx")
        Path(deck).write_bytes(b"pptx")
        report = {"summary": {"conflict": 1, "auto_update": 1}, "mode": "full",
                  "actions": [{"element": "e1", "action": "conflict", "reason": "the user edited the paragraph and a figure moved"},
                              {"element": "e2", "action": "auto_update", "reason": "a figure moved"}]}
        manifest = {"elements": [{"id": "e1", "account": ACCOUNT}, {"id": "e2", "account": "现金"}]}
        applied = {"applied": [], "skipped": [{"element": "e1", "action": "conflict", "reason": "the user edited the paragraph and a figure moved"}],
                   "out_path": os.path.join(tmp, "out.pptx"), "touched_slides": []}
        with mock.patch.object(U, "analyze_returned", return_value=report), mock.patch.object(L, "read_manifest", return_value=manifest), \
                mock.patch.object(U, "apply_updates", return_value=applied), mock.patch.object(L, "restamp", return_value=None), \
                mock.patch("fdd_utils.agent.export.layout_problems", return_value={"problems": []}):
            out = work.update_deck(deck, os.path.join(tmp, "out.pptx"))
        self.assertTrue(out["ok"], out)
        data = out["data"]
        self.assertFalse(data["complete"])
        self.assertEqual({u["action"] for u in data["unfinished"]}, {"conflict", "needs_candidate"})
        self.assertTrue(any("conflicts" in s for s in data["next_steps"]))
        self.assertEqual(data["original_kept"], deck)
        self.assertFalse(work.deck["last_update"]["complete"])
        self.assertEqual(Path(deck).read_bytes(), b"pptx")                    # the returned deck itself is never written


# -- the app panel opens the session with what the app holds ----------------------------------------------------------------------------

class PanelOpen(unittest.TestCase):
    def test_the_panel_passes_the_apps_overrides_and_lets_configured_agent_models_decide(self):
        from fdd_utils.ui import agent_panel as P
        state = {"temp_path": "wb.xlsx", "entity_name": "E", "selected_sheet": "Financials", "last_run_folder": "run_1",
                 "mapping_overrides": {"应收账款": "AR"}, "model_type": "local", "model_name": "qwen"}
        seen = {}
        with mock.patch.object(W.WorkSession, "open", side_effect=lambda *a, **k: seen.update(k) or "work"), \
                mock.patch("fdd_utils.agent.flag.role_config", return_value={"provider": "dashscope", "model": "flash"}):
            P._open_work(state)
        self.assertEqual(seen["mapping_overrides"], {"应收账款": "AR"})
        self.assertEqual((seen["model_type"], seen["model_name"]), ("", None))       # agent.models.primary decides
        with mock.patch.object(W.WorkSession, "open", side_effect=lambda *a, **k: seen.update(k) or "work"), \
                mock.patch("fdd_utils.agent.flag.role_config", return_value={"provider": "", "model": ""}):
            P._open_work(state)
        self.assertEqual((seen["model_type"], seen["model_name"]), ("local", "qwen"))  # nothing configured: the sidebar's model


# -- a wrong figure cannot pass by relabelling itself ------------------------------------------------------------------------------------

class Relabelling(unittest.TestCase):
    def test_the_same_wrong_sentence_fails_whatever_kind_inference_or_cross_account_it_declares(self):
        contract = build()
        cl.standard_derived(contract)
        ref = total_ref(cl.evidence_view(contract), P2)
        base = {"account": ACCOUNT, "periods": [P2], "cites": [ref], "text": "2025年12月31日应收账款余额150.0万元"}
        for variant in ({}, {"kind": "descriptive"}, {"kind": "cause"}, {"inference": True}, {"inference": "true"}, {"cross_account": True},
                        {"kind": "annualized", "inference": True}, {"periods": [P1, P2]}, {"cites": [ref, "D1"]}):
            claim = dict(base, **variant)
            outcome = cl.verify_candidate([claim], {ACCOUNT: contract}, default_account=ACCOUNT)
            self.assertFalse(outcome["ok"], variant)
            self.assertEqual(outcome["failed"] + outcome["unverified"], 1, variant)


if __name__ == "__main__":
    unittest.main(verbosity=2)

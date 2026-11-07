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
from test_agent_controller import Policy, evidence_of, returned, tool, total_ref  # noqa: E402
from test_workspace import ACCOUNT, OTHER, P1, P2, answer_policy, build, session_on, write_book  # noqa: E402


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

    def _work(self, model, escalation=None, mode="single", prepare=False):
        tmp = tempfile.mkdtemp(prefix="cfg_")
        book = os.path.join(tmp, "wb.xlsx")
        write_book(book)
        source = session_on(book)
        source.load_params = {"workbook": book}
        # the tool loop's rules are what these tests hold: the prepared write is tested on its own (PreparedWrite)
        return W.WorkSession(source, model=model, escalation=escalation, model_name="m", policy=R.Policy(mode=mode), budget=C.Budget(max_steps=8),
                             prepare=prepare)

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


# -- the prepared write: the controller reads, the model writes once, the verifier checks, a repair touches only what failed -------------

def writer_view(user):
    """The evidence JSON the prepared write puts in the prompt (one compact line after 'EVIDENCE for ...')."""
    return json.loads(user.split("EVIDENCE for", 1)[1].split("\n", 1)[1].split("\n\n", 1)[0])


class Writer:
    """A scripted writer: fn(system, user, n) -> the reply object; records what it was sent."""

    def __init__(self, fn):
        self.fn, self.systems, self.users = fn, [], []

    def __call__(self, system, user, max_tokens):
        self.systems.append(system)
        self.users.append(user)
        return {"content": json.dumps(self.fn(system, user, len(self.users)), ensure_ascii=False), "prompt_tokens": 100, "completion_tokens": 20}


def amount_claim(view, period, text):
    return {"kind": "amount", "periods": [period], "cites": [total_ref(view, period)], "text": text}


GOOD = "2025年12月31日应收账款余额130.0万元"


class PreparedWrite(unittest.TestCase):
    def _work(self, model, escalation=None, mode="single", extra_account=False, config=None):
        tmp = tempfile.mkdtemp(prefix="prep_")
        book = os.path.join(tmp, "wb.xlsx")
        write_book(book)
        source = session_on(book)
        source.load_params = {"workbook": book}
        if extra_account:
            from test_claims import build as build_contract
            other = build_contract(account=OTHER)
            other.binding = source.binding.to_dict()
            source.contracts[OTHER] = other
        return W.WorkSession(source, model=model, escalation=escalation, model_name="m", policy=R.Policy(mode=mode),
                             budget=C.Budget(max_steps=8), prepare=True, config=config or {"agent": {}})

    SEEDED = {"agent": {"seed_facts": True}}
    REVIEWED = {"agent": {"seed_facts": True, "review_meaning": True}}
    BALANCE, MOVE = "截至2025年12月31日，余额为130.0万元。", "较2024年12月31日增加30.0万元。"
    CAUSE = {"kind": "cause", "periods": [P2], "cites": [], "inference": True, "text": "增加主要系回款节奏变化所致"}

    def test_python_writes_the_balance_and_movement_and_the_model_only_adds(self):
        writer = Writer(lambda system, user, n: {"add": [], "unresolved": []})
        work = self._work(writer, config=self.SEEDED)
        turn = work.edit("撰写应收账款的段落", ACCOUNT)
        self.assertTrue(turn.ok, turn.report)
        self.assertIn("KEPT (Python wrote these", writer.users[0])
        self.assertIn(self.BALANCE, writer.users[0])
        self.assertEqual([c["text"] for c in work.source.candidates[ACCOUNT].claims], [self.BALANCE, self.MOVE])
        self.assertEqual((work.shared.calls, turn.report["prepared"]["seeded"]), (1, 2))

    def test_an_addition_that_fails_is_repaired_and_the_python_sentences_stay(self):
        def fn(system, user, n):
            view = writer_view(user)
            if n == 1:
                return {"add": [amount_claim(view, P1, "2024年12月31日应收账款余额150.0万元")]}
            return {"fix": [{"index": 0, "claim": amount_claim(view, P2, "改写")}, {"index": 2, "drop": "数字不符"}]}
        writer = Writer(fn)
        work = self._work(writer, config=self.SEEDED)
        turn = work.edit("撰写应收账款的段落", ACCOUNT)
        self.assertTrue(turn.ok, turn.report)
        self.assertIn("[2]", writer.users[1].split("REJECTED", 1)[1])
        self.assertEqual([c["text"] for c in work.source.candidates[ACCOUNT].claims], [self.BALANCE, self.MOVE])
        self.assertIn("数字不符", turn.unresolved)

    def _reviewed(self, after_review):
        def fn(system, user, n):
            if "REVIEW:" in user:
                return after_review(writer_view(user))
            return {"add": [dict(self.CAUSE)]}
        writer = Writer(fn)
        reviewer = Writer(lambda system, user, n: {"objections": [{"index": 2, "why": "附注未说明回款节奏"}], "missing": []})
        work = self._work(writer, escalation=reviewer, mode="escalate", config=self.REVIEWED)
        return work, work.edit("撰写应收账款的段落", ACCOUNT), writer, reviewer

    def test_the_stronger_model_reads_a_paragraph_with_a_cause_and_its_objection_is_repaired(self):
        work, turn, writer, reviewer = self._reviewed(lambda view: {"fix": [{"index": 2, "drop": "原因未能从附注核实"}]})
        self.assertTrue(turn.ok, turn.report)
        self.assertIn("Judge meaning only", reviewer.systems[0])
        self.assertEqual(turn.report["review"]["outcome"], "repaired")
        self.assertEqual([c["text"] for c in work.source.candidates[ACCOUNT].claims], [self.BALANCE, self.MOVE])
        books = work.shared.summary()
        self.assertEqual((books["by_purpose"]["review"]["calls"], books["escalations"]), (1, 0), "a review is not an escalation")
        self.assertEqual(len(writer.users), 2)

    def test_a_repair_after_review_that_does_not_verify_keeps_the_verified_paragraph(self):
        work, turn, writer, reviewer = self._reviewed(lambda view: {"fix": [{"index": 2, "claim": amount_claim(view, P2, "余额为999.0万元")}]})
        self.assertTrue(turn.ok, turn.report)
        self.assertTrue(turn.report["review"]["outcome"].startswith("reverted"))
        self.assertEqual([c["text"] for c in work.source.candidates[ACCOUNT].claims], [self.BALANCE, self.MOVE, self.CAUSE["text"]])
        self.assertTrue(work.source.candidates[ACCOUNT].validated)
        self.assertTrue(any(u.startswith("REVIEW:") for u in turn.unresolved), "the objection is not lost")

    def test_python_sentences_verify_in_english_and_for_an_income_statement_line(self):
        work = self._work(Writer(lambda *a: {}))
        contract = work.source.contracts[ACCOUNT]
        view = T.get_account_evidence(work.source, ACCOUNT)
        english = C.AgentController(work.source, lambda *a: {}, language="Eng", seed=True)._seed_claims(ACCOUNT, view)
        self.assertEqual([c["text"] for c in english], ["As at 31 December 2025, the balance was CNY1.3 million.",
                                                        "This was an increase of CNY0.3 million from 31 December 2024."])
        contract.statement_type = "IS"
        flow = C.AgentController(work.source, lambda *a: {}, language="Chi", seed=True)._seed_claims(ACCOUNT, view)
        self.assertEqual(flow[0]["text"], "截至2025年12月31日止期间，发生额为130.0万元。")
        self.assertEqual((C._money(-70_677_440, True), C._money(3_091, True), C._money(2.5e8, True)), ("-7,067.7万元", "0.31万元", "2.50亿元"))

    def test_a_plain_paragraph_with_nothing_to_judge_is_not_sent_for_review(self):
        reviewer = Writer(lambda system, user, n: {"objections": [], "missing": []})
        work = self._work(Writer(lambda system, user, n: {"add": []}), escalation=reviewer, mode="escalate", config=self.REVIEWED)
        turn = work.edit("撰写应收账款的段落", ACCOUNT)
        self.assertTrue(turn.ok, turn.report)
        self.assertEqual((len(reviewer.users), turn.report["review"]["reasons"]), (0, []))

    def test_one_account_is_written_in_one_call_with_no_tool_menu_from_evidence_the_controller_read(self):
        writer = Writer(lambda system, user, n: {"draft": {"claims": [amount_claim(writer_view(user), P2, GOOD)], "unresolved": []}})
        work = self._work(writer)
        turn = work.edit("撰写应收账款的段落", ACCOUNT, accept=True)
        self.assertTrue(turn.ok, turn.report)
        self.assertEqual((len(writer.users), work.shared.calls, turn.tool_path), (1, 1, []))
        self.assertNotIn("Tools:", writer.systems[0])
        self.assertNotIn("validate_candidate", writer.systems[0])
        self.assertIn("EVIDENCE for 应收账款", writer.users[0])
        self.assertLess(writer.users[0].index("EVIDENCE for"), writer.users[0].index("REQUEST:"), "the part that repeats comes first")
        self.assertTrue(work.source.candidates[ACCOUNT].accepted)

    def test_the_loop_would_have_sent_the_tool_menu_every_call(self):
        loop_system = C.system_prompt("edit", self._work(Writer(lambda *a: {})).source, "Chi")
        self.assertGreater(len(loop_system), 2 * len(C.writer_system_prompt("Chi")))

    def test_a_rejected_claim_is_repaired_alone_and_the_verified_one_stays_byte_for_byte(self):
        def fn(system, user, n):
            view = writer_view(user)
            if n == 1:
                return {"draft": {"claims": [amount_claim(view, P2, GOOD), amount_claim(view, P1, "2024年12月31日应收账款余额150.0万元")],
                                  "unresolved": []}}
            return {"fix": [{"index": 0, "claim": amount_claim(view, P2, "改写已核实的句子")},
                            {"index": 1, "claim": amount_claim(view, P1, "2024年12月31日应收账款余额100.0万元")}]}
        writer = Writer(fn)
        work = self._work(writer)
        turn = work.edit("撰写应收账款的段落", ACCOUNT)
        self.assertTrue(turn.ok, turn.report)
        repair = writer.users[1]
        self.assertIn("KEPT", repair)
        self.assertIn("[0] " + GOOD, repair)
        self.assertIn("REJECTED", repair)
        self.assertIn("150.0万元", repair.split("REJECTED", 1)[1])
        final = work.source.candidates[ACCOUNT].claims
        self.assertEqual(final[0]["text"], GOOD, "a fix aimed at a kept claim is ignored")
        self.assertEqual(len(writer.users), 2)

    def test_a_draft_that_leaves_out_the_latest_balance_is_sent_back_until_it_states_it(self):
        def fn(system, user, n):
            view = writer_view(user)
            move = next(d["ref"] for d in view["derived"] if d["kind"] == "movement")
            if n == 1:
                return {"draft": {"claims": [{"kind": "movement", "periods": [P1, P2], "cites": [move], "text": "较2024年末增加30.0万元"}],
                                  "unresolved": []}}
            return {"fix": [], "add": [amount_claim(view, P2, GOOD)]}
        writer = Writer(fn)
        work = self._work(writer)
        turn = work.edit("撰写应收账款的段落", ACCOUNT)
        self.assertTrue(turn.ok, turn.report)
        self.assertIn("REQUIRED_NOT_STATED", writer.users[1])
        self.assertEqual(len(work.source.candidates[ACCOUNT].claims), 2)

    def test_the_writer_can_hand_the_turn_to_the_tool_loop_with_the_evidence_already_read(self):
        def fn(system, user, n):
            if n == 1:
                return {"need_evidence": "需要查附注"}
            if "Tools:" not in system:
                raise AssertionError("the loop sends the tool menu")
            view = evidence_of(returned(user))
            if "validate_candidate(" not in user:
                return tool("validate_candidate", account=ACCOUNT, claims=[amount_claim(view, P2, GOOD)])
            return {"done": {"account": ACCOUNT, "export": False, "summary": "ok"}}
        writer = Writer(fn)
        work = self._work(writer)
        turn = work.edit("撰写应收账款的段落", ACCOUNT)
        self.assertTrue(turn.ok, turn.report)
        self.assertIn("TOOL get_account_evidence(", writer.users[1], "the loop starts with the evidence the controller read")
        self.assertNotIn("get_account_evidence", turn.tool_path, "the model did not have to fetch it")
        self.assertEqual(turn.report.get("prepared_fallback"), "需要查附注")

    def test_the_first_model_writes_and_the_second_repairs_after_repeated_failures(self):
        def wrong(system, user, n):
            view = writer_view(user)
            claim = amount_claim(view, P2, f"2025年12月31日应收账款余额{150 + n:.1f}万元")
            return {"draft": {"claims": [claim], "unresolved": []}} if n == 1 else {"fix": [{"index": 0, "claim": claim}]}
        primary = Writer(wrong)
        pro = Writer(lambda system, user, n: {"fix": [{"index": 0, "claim": amount_claim(writer_view(user), P2, GOOD)}]})
        work = self._work(primary, escalation=pro, mode="escalate")
        turn = work.edit("撰写应收账款的段落", ACCOUNT)
        self.assertTrue(turn.ok, turn.report)
        self.assertEqual([e["reason"] for e in turn.report["routing"]["escalations"]], ["capability"])
        self.assertEqual((len(primary.users), len(pro.users)), (2, 1))
        self.assertIn("REJECTED", pro.users[0], "the stronger model gets the rejected claim, not the whole task again")

    def test_an_entity_is_drafted_account_by_account_with_a_fresh_memory_and_a_first_generation_deck_names_what_is_missing(self):
        def fn(system, user, n):
            view = writer_view(user)
            text = GOOD if view["account"] == ACCOUNT else "其他应收款余额999.0万元"
            claim = amount_claim(view, P2, text)
            return {"draft": {"claims": [claim], "unresolved": []}} if "REJECTED" not in user else {"fix": [{"index": 0, "claim": claim}]}
        writer = Writer(fn)
        work = self._work(writer, extra_account=True)
        done = work.draft_entity()
        self.assertEqual(done["ok"], [ACCOUNT])
        self.assertEqual(done["not_finished"], [OTHER])
        later = [u for u in writer.users if "EVIDENCE for 其他应收款" in u]
        self.assertTrue(later)
        self.assertTrue(all("EARLIER IN THIS SESSION" not in u and GOOD not in u for u in later), "no account's text rides along with the next")
        from fdd_utils.agent.export import merged_results, session_delivery
        work.source.base_results = {OTHER: {"final": "旧流程的文字"}}
        results = merged_results(work.source, [ACCOUNT], first_generation=True)
        self.assertNotIn(OTHER, results, "a first generation never borrows a base run's text")
        delivery = session_delivery(work.source, results, edited=[ACCOUNT], first_generation=True)
        missing = [f for f in delivery.findings if f.code == "ACCOUNT_NOT_DRAFTED"]
        self.assertEqual([(f.severity, f.account) for f in missing], [("blocked", OTHER)])
        self.assertEqual(delivery.status, "blocked")


class WholeDatabookTrial(unittest.TestCase):
    """entity_trial.py: the arms are what they say, and the legacy row is what the ledger recorded for that run, nothing else."""

    CFG = {"agent": {"models": {"primary": {"provider": "dashscope", "model": "flash-id"}, "escalation": {"provider": "dashscope", "model": "pro-id"}}},
           "dashscope": {"available_models": [{"id": "pro-id", "price_in": 12, "price_out": 24}, {"id": "flash-id", "price_in": 1, "price_out": 2}]}}

    def args(self, **kw):
        return SimpleNamespace(**{"max_output": 4000, "escalate_after": 1, **kw})

    def test_each_arm_is_the_model_and_the_path_it_names(self):
        import entity_trial as E
        got = {n: E.arm_setup(self.CFG, n, self.args()) for n in E.ARMS}
        self.assertEqual((got["pro"]["primary"]["model"], got["pro"]["escalation"], got["pro"]["prepare"]), ("pro-id", None, True))
        self.assertEqual((got["flash"]["primary"]["model"], got["flash"]["escalation"]), ("flash-id", None))
        self.assertEqual((got["mixed"]["primary"]["model"], got["mixed"]["escalation"]["model"], got["mixed"]["routing"], got["mixed"]["escalate_after"]),
                         ("flash-id", "pro-id", "escalate", 1))
        self.assertEqual((got["loop"]["prepare"], got["loop"]["routing"]), (False, "escalate"))
        self.assertEqual([(got[n].get("seed"), got[n].get("review")) for n in ("mixed", "seeded", "reviewed")],
                         [(False, False), (True, False), (True, True)])
        self.assertEqual(got["reviewed"]["escalation"]["model"], "pro-id", "the reviewer is the stronger model")
        cfg = E.arm_config(self.CFG, got["reviewed"])
        self.assertEqual((cfg["agent"]["seed_facts"], cfg["agent"]["review_meaning"]), (True, True))
        self.assertTrue(all(s["primary"]["max_output"] == 4000 for s in got.values()))
        cfg = E.arm_config(self.CFG, got["pro"])
        self.assertEqual((cfg["agent"]["models"]["primary"]["model"], cfg["agent"]["prepared_write"], cfg["agent"]["routing"]), ("pro-id", True, "single"))
        self.assertEqual(self.CFG["agent"]["models"]["primary"]["model"], "flash-id", "the arm's config is a copy")

    def test_the_legacy_row_is_the_ledger_for_that_runs_file_and_time_window(self):
        import entity_trial as E
        tmp = Path(tempfile.mkdtemp(prefix="trial_"))
        run = tmp / "run_20261030_180626"
        run.mkdir()
        (run / "results.yml").write_text("{}", encoding="utf-8")
        (run / "run_identity.json").write_text(json.dumps({"workbook": "C:/x/book.xlsx", "written_at": "2026-10-30T18:40:00"}), encoding="utf-8")
        rows = [
            {"time": "2026-10-30T18:07:00", "file": "book.xlsx", "provider": "dashscope", "model": "pro-id", "prompt_tokens": 1000, "completion_tokens": 100,
             "cached_prompt_tokens": 400},
            {"time": "2026-10-30T18:30:00", "file": "book.xlsx", "provider": "dashscope", "model": "pro-id", "ok": False},
            {"time": "2026-10-30T18:31:00", "file": "book.xlsx", "provider": "dashscope", "model": "pro-id", "token_source": "response_cache"},
            {"time": "2026-10-30T18:32:00", "file": "other.xlsx", "provider": "dashscope", "model": "pro-id", "prompt_tokens": 9999},
            {"time": "2026-10-30T20:00:00", "file": "book.xlsx", "provider": "dashscope", "model": "pro-id", "prompt_tokens": 9999},
        ]
        ledger = tmp / "usage.jsonl"
        ledger.write_text("\n".join(json.dumps(r) for r in rows), encoding="utf-8")
        got = E.legacy_usage(str(run), self.CFG, ledger=ledger)
        self.assertEqual({k: got["total"][k] for k in ("calls", "failed", "response_cache", "in", "cached", "out")},
                         {"calls": 1, "failed": 1, "response_cache": 1, "in": 1000, "cached": 400, "out": 100})
        self.assertEqual(got["total"]["cost"], round((1000 * 12 + 100 * 24) / 1e6, 4))
        self.assertIn("error", E.legacy_usage(str(run), self.CFG, ledger=tmp / "missing.jsonl"), "no ledger: said, not estimated")

    def test_the_scripted_writer_writes_the_latest_balance_it_is_shown(self):
        import entity_trial as E
        tmp = tempfile.mkdtemp(prefix="trial_w_")
        book = os.path.join(tmp, "wb.xlsx")
        write_book(book)
        source = session_on(book)
        source.load_params = {"workbook": book}
        work = W.WorkSession(source, model=E.ScriptedWriter("x"), model_name="m", budget=C.Budget(max_steps=8), prepare=True)
        turn = work.edit("撰写应收账款的段落", ACCOUNT, accept=True)
        self.assertTrue(turn.ok, turn.report)
        self.assertEqual(work.shared.calls, 1)


class AcceptanceAllowance(unittest.TestCase):
    """A cap declared on the paid acceptance command is ONE ceiling over the whole command. A review (2026-11-06) found the
    entity runs opened with fresh caps of their own: a declared 90 calls allowed 120 with the routing entity, 180 with all."""

    TOTAL = {"max_calls": 90, "max_tokens": 900_000, "max_cost": 9.0}

    @staticmethod
    def drain(total, runs, spend=lambda share: share):
        import agent_acceptance as A
        al, given = A.Allowance(total, runs), []
        for i in range(runs):
            share = al.next(f"run{i}")
            given.append(share)
            if share is not None:
                used = spend(share)
                al.charge({"calls": used["max_calls"] or 0, "charged_tokens": used["max_tokens"] or 0, "cost": used["max_cost"]})
        return al, given

    def test_runs_that_each_spend_their_whole_share_stay_within_the_total(self):
        for runs in (3, 4, 6):                      # the three task runs; + the routing entity; + every entity
            al, given = self.drain(self.TOTAL, runs)
            self.assertLessEqual(al.spent["max_calls"], 90, runs)
            self.assertLessEqual(al.spent["max_tokens"], 900_000, runs)
            self.assertLessEqual(al.spent["max_cost"], 9.0 + 1e-9, runs)
            self.assertTrue(all(g is not None for g in given), runs)

    def test_a_cap_below_the_number_of_runs_leaves_the_last_runs_unstarted(self):
        al, given = self.drain({"max_calls": 2, "max_tokens": None, "max_cost": None}, 3)
        self.assertEqual([g and g["max_calls"] for g in given], [1, 1, None])
        self.assertEqual(al.spent["max_calls"], 2)
        self.assertIn("used up", al.report()["runs"][-1]["not_started"])

    def test_token_and_cost_totals_run_out_like_calls(self):
        al, given = self.drain({"max_calls": None, "max_tokens": 1000, "max_cost": None}, 3, spend=lambda s: {**s, "max_tokens": 1000})
        self.assertEqual(given[1:], [None, None])
        al, given = self.drain({"max_calls": None, "max_tokens": None, "max_cost": 1.0}, 3, spend=lambda s: {**s, "max_cost": 1.0})
        self.assertEqual(given[1:], [None, None])

    def test_what_a_run_leaves_passes_to_the_runs_after_it(self):
        al, given = self.drain({"max_calls": 90, "max_tokens": None, "max_cost": None}, 4, spend=lambda s: {**s, "max_calls": 0})
        self.assertEqual([g["max_calls"] for g in given], [23, 30, 45, 90])

    def test_no_declared_cap_is_no_cap(self):
        al, given = self.drain({"max_calls": None, "max_tokens": None, "max_cost": None}, 6)
        self.assertTrue(all(g == {"max_calls": None, "max_tokens": None, "max_cost": None} for g in given))

    def test_the_paid_command_hands_out_caps_that_add_up_to_no_more_than_the_total(self):
        import agent_acceptance as A
        seen = []

        def spend_all(name, setup, databook, base_run, workdir, limits, **_):
            seen.append((name, limits))
            return {"books": {"calls": limits.max_calls, "charged_tokens": limits.max_tokens, "cost": None}}
        setups = {n: {"routing": "off", "primary": {"provider": "x"}} for n in A.CONFIGS}
        for entity_configs, runs in (("routing", 4), ("all", 6)):
            seen.clear()
            args = SimpleNamespace(databook="book.xlsx", base_run="run_x", i_accept_paid_calls=True, entity=True,
                                   entity_configs=entity_configs, max_calls=90, max_tokens=900_000, max_cost=0.0)
            with mock.patch.object(A, "_config", return_value={}), mock.patch.object(A, "preview_roles", return_value={}), \
                    mock.patch.object(A, "preflight", return_value=[]), mock.patch.object(A, "configurations", return_value=setups), \
                    mock.patch.object(A, "run_configuration", side_effect=spend_all), mock.patch.object(A, "run_entity", side_effect=spend_all), \
                    mock.patch.object(A, "_save", return_value=A.ROOT / "x.json"), mock.patch.object(A, "summary_text", return_value=""), \
                    mock.patch.object(A, "quality_dump", return_value={"path": "", "line": ""}), mock.patch("builtins.print"):
                self.assertEqual(A.cmd_paid(args), 0)
            self.assertEqual(len(seen), runs, entity_configs)
            self.assertLessEqual(sum(l.max_calls for _n, l in seen), 90, entity_configs)
            self.assertLessEqual(sum(l.max_tokens for _n, l in seen), 900_000, entity_configs)


if __name__ == "__main__":
    unittest.main(verbosity=2)

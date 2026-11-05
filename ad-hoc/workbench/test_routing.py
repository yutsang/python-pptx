"""Which model takes the next step, what a failure means for it, and the one set of books -- offline.

Zero tokens. Every model here is a scripted callable. This proves the controller's rules: that a sentence the evidence in hand can
repair is a rewrite and not a question for the user, that only the model failing to manage with what it had can send a turn to
the escalation model, what that model is handed, and that every call of every role is counted once, with what the provider
reported kept apart from what was estimated. It says nothing about whether a real second model does better; that is a paid
measurement and is not made here.

    python ad-hoc/workbench/test_routing.py
"""
from __future__ import annotations

import json
import os
import re
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parents[1]))

from fdd_utils.agent import controller as C  # noqa: E402
from fdd_utils.agent import routing as R  # noqa: E402
from fdd_utils.agent import workspace as W  # noqa: E402
from fdd_utils.agent.session import SourceSession  # noqa: E402
from fdd_utils.ai import claims as cl  # noqa: E402
from test_agent_controller import Policy, WrongFinancials, clean_session, evidence_of, tool, total_ref  # noqa: E402
from test_claims import P1, P2, build, ref  # noqa: E402

ACCOUNT = "应收账款"
ROOT = HERE.parents[1]


# -- the emitted codes ---------------------------------------------------------------------------------------------------------

#: where a model-facing code can be raised, and the shape of the call that raises it
EMITTERS = {
    "fdd_utils/ai/claims.py": [r'(?:fail|review|info)\(\s*"([A-Z][A-Z_]+)"', r'\(\s*"(?:failed|review|info)"\s*,\s*"([A-Z][A-Z_]+)"'],
    "fdd_utils/agent/tools.py": [r'ToolError\(\s*"([A-Z][A-Z_]+)"', r'fail\(\s*"([A-Z][A-Z_]+)"'],
    "fdd_utils/agent/investigate.py": [r'ToolError\(\s*"([A-Z][A-Z_]+)"', r'fail\(\s*"([A-Z][A-Z_]+)"'],
    "fdd_utils/agent/export.py": [r'ToolError\(\s*"([A-Z][A-Z_]+)"'],
    "fdd_utils/agent/updates.py": [r'ToolError\(\s*"([A-Z][A-Z_]+)"'],
    "fdd_utils/agent/controller.py": [r'(?:codes|_gate_codes) = \["([A-Z][A-Z_]+)"'],
}


def emitted_codes() -> dict:
    found: dict = {}
    for name, patterns in EMITTERS.items():
        text = (ROOT / name).read_text(encoding="utf-8")
        for pattern in patterns:
            for m in re.finditer(pattern, text):
                found.setdefault(m.group(1), set()).add(name)
    return found


class Classification(unittest.TestCase):
    def test_every_code_the_verifier_the_tools_and_the_controller_can_raise_has_been_classified_by_a_person(self):
        found = emitted_codes()
        self.assertGreater(len(found), 50, "the scan found the emitters")
        for known in ("DIRECTION_UNBOUND", "SIGN_ERROR", "UNSEEN_REF", "VALIDATION_FAILED", "NOT_JSON"):
            self.assertIn(known, found)
        unclassified = sorted(c for c in found if not R.is_classified(c))
        self.assertEqual(unclassified, [], f"classify these in routing.CLASS (which move do they ask for?): {unclassified}")

    def test_the_gate_that_catches_a_forgotten_code_does_catch_one(self):
        self.assertFalse(R.is_classified("A_CODE_NOBODY_CLASSIFIED"))
        self.assertEqual(R.classify("A_CODE_NOBODY_CLASSIFIED"), R.SYSTEM)       # the floor: never a reason to spend more

    def test_a_sentence_the_evidence_in_hand_can_repair_is_a_rewrite_not_a_question_for_the_user(self):
        # the ones that used to be filed under "needs the user": an unbound direction word, an unstated sign, a statement about
        # another account, a total the permissions do not let it state
        for code in ("DIRECTION_UNBOUND", "DIRECTION_UNSUPPORTED", "PERCENT_SIGN_UNSTATED", "CROSS_ACCOUNT_UNBOUND", "TOTAL_NOT_DESCRIBABLE",
                     "LINE_UNBOUND", "PERIOD_UNBOUND", "PERIODS_UNDECLARED", "MOVEMENT_NOT_PERMITTED", "SIGN_ERROR", "WRONG_PERIOD",
                     "UNIT_ERROR", "AMOUNT_UNCITED", "ANNUALIZED_AS_FORECAST", "CAUSE_NOTE_UNRELATED", "LAYOUT_PROBLEM"):
            self.assertEqual(R.classify(code), R.REWRITE, code)

    def test_only_a_conflict_a_blocked_source_or_a_missing_file_is_for_a_person(self):
        for code in ("conflict:pairing", "conflict:应收账款:2025-12-31", "source:应收账款", "HIDDEN_SHEET", "DELIVERY_BLOCKED", "NO_BASELINE"):
            self.assertEqual(R.classify(code), R.HUMAN, code)

    def test_missing_evidence_is_neither_a_rewrite_nor_a_person(self):
        for code in ("UNSEEN_REF", "UNKNOWN_REF", "CAUSE_UNSUPPORTED", "NOT_COMPUTABLE", "replaced:应收账款:2025-12-31"):
            self.assertEqual(R.classify(code), R.EVIDENCE, code)
        self.assertEqual(R.classify("stale"), R.REFRESH)
        self.assertEqual(R.classify("STALE_SOURCE"), R.REFRESH)

    def test_the_model_failing_to_manage_is_rewrite_or_format_and_nothing_that_waits_on_anything_else(self):
        self.assertTrue(R.is_capability_failure(["SIGN_ERROR", "WRONG_PERIOD"]))
        self.assertTrue(R.is_capability_failure(["NOT_JSON"]))
        self.assertTrue(R.is_capability_failure(["SIGN_ERROR", "DETAIL_NESTING_UNKNOWN"]), "an info code changes nothing")
        self.assertFalse(R.is_capability_failure(["SIGN_ERROR", "UNSEEN_REF"]), "it is also waiting on evidence")
        self.assertFalse(R.is_capability_failure(["conflict:pairing"]))
        self.assertFalse(R.is_capability_failure(["TOOL_EXCEPTION"]))
        self.assertFalse(R.is_capability_failure(["STALE_SOURCE"]))
        self.assertFalse(R.is_capability_failure([]))

    def test_an_unbound_direction_word_is_fixed_by_rewording_from_the_refs_already_cited_with_no_question_and_no_escalation(self):
        session = clean_session()
        contract = session.contracts[ACCOUNT]
        controller = C.AgentController(session, Policy(lambda u, r: tool("inspect_sources")))
        every = [ref(contract, P1, "甲公司"), ref(contract, P2, "甲公司"),
                 cl.derive_share(contract, "甲公司", P1)["ref"], cl.derive_share(contract, "甲公司", P2)["ref"]]
        controller.seen_refs |= set(every)
        controller.consulted.append(ACCOUNT)

        def answer(text, cites, kind):
            return {"claims": [{"account": ACCOUNT, "kind": kind, "periods": [P1, P2], "cites": cites, "text": text}], "unresolved": []}
        accepted, feedback, _ = controller._gate_answer(answer("甲公司较上年下降", every, "share"))
        self.assertFalse(accepted)
        self.assertIn("DIRECTION_UNBOUND", controller._gate_codes)
        self.assertEqual(R.classify_all(controller._gate_codes), {R.REWRITE: ["DIRECTION_UNBOUND"]})
        self.assertIn("[rewrite]", feedback)
        self.assertNotIn("[human]", feedback)
        self.assertIn("Do not ask the user", feedback)
        # the same evidence, reworded so the word attaches to one quantity: accepted
        accepted, feedback, _ = controller._gate_answer(answer("甲公司余额增加", [every[0], every[1]], "movement"))
        self.assertTrue(accepted, feedback)

    def test_the_hint_names_each_class_present_and_a_person_comes_first(self):
        hint = R.next_step_hint(["SIGN_ERROR", "conflict:pairing", "UNSEEN_REF"])
        self.assertLess(hint.index("[human]"), hint.index("[evidence]"))
        self.assertLess(hint.index("[evidence]"), hint.index("[rewrite]"))
        self.assertEqual(R.next_step_hint(["DETAIL_NESTING_UNKNOWN"]), "")


# -- escalation at the controller ----------------------------------------------------------------------------------------------

def spec(role, fn, model, price=None, thinking=""):
    return R.RoleSpec(role, fn, provider="test", model=model, thinking=thinking, price=price)


def wrong_amount_policy(prompts=None, text="2025年12月31日应收账款余额150.0万元"):
    """Fetches the evidence once, then keeps giving an answer whose figure does not match what it cites."""
    def fn(user, results):
        if prompts is not None:
            prompts.append(user)
        view = evidence_of(results)
        if view is None:
            return tool("get_account_evidence", account=ACCOUNT)
        return {"answer": {"claims": [{"account": ACCOUNT, "kind": "amount", "periods": [P2], "cites": [total_ref(view, P2)], "text": text}], "unresolved": []}}
    return Policy(fn)


def handoff_reader_policy(prompts):
    """Knows nothing but the hand-off: reads the ref of the 2025 total from it and answers correctly."""
    def fn(user, results):
        prompts.append(user)
        found = re.search(r"ref=(\w+) period=2025-12-31 basis=authority_adjusted", user)
        if not found:
            return tool("get_account_evidence", account=ACCOUNT)
        return {"answer": {"claims": [{"account": ACCOUNT, "kind": "amount", "periods": [P2], "cites": [found.group(1)],
                                       "text": "2025年12月31日应收账款余额130.0万元"}], "unresolved": []}}
    return Policy(fn)


def routed(primary_fn, escalation_fn=None, *, mode="escalate", limits=None, **policy):
    escalation = spec(R.ROLE_ESCALATION, escalation_fn, "big-model") if escalation_fn is not None else None
    return R.RoutedModel(spec(R.ROLE_PRIMARY, primary_fn, "small-model"), escalation, policy=R.Policy(mode=mode, **policy),
                         budget=R.SharedBudget(limits or R.Limits()))


class Escalation(unittest.TestCase):
    def test_two_wrong_answers_in_a_row_on_evidence_it_had_hands_the_task_to_the_escalation_model_once(self):
        pro_prompts, small_prompts = [], []
        router = routed(wrong_amount_policy(small_prompts), handoff_reader_policy(pro_prompts))
        run = C.AgentController(clean_session(), router).run("应收账款期末余额是多少？")
        self.assertEqual(run.stop_reason, C.STOP_ANSWERED, run.report)
        roles = [r.role for r in router.budget.records]
        self.assertEqual(roles, ["primary", "primary", "primary", "escalation"])         # evidence, wrong, wrong -> hand-off -> right
        routing = run.report["routing"]
        self.assertEqual(routing["stop_class"], "ok")
        self.assertEqual(routing["role_at_end"], "escalation")
        self.assertEqual(len(routing["escalations"]), 1)
        record = routing["escalations"][0]
        self.assertEqual((record["from"], record["to"], record["reason"], record["model"]), ("primary", "escalation", "capability", "big-model"))
        self.assertTrue(set(record["classes"]) <= {R.REWRITE}, record)
        self.assertEqual(run.usage["by_role"]["primary"]["calls"], 3)
        self.assertEqual(run.usage["by_role"]["escalation"]["calls"], 1)
        self.assertEqual(run.usage["by_role"]["escalation"]["model"], "big-model")

    def test_the_hand_off_is_a_task_state_and_not_the_conversation(self):
        pro_prompts = []
        router = routed(wrong_amount_policy(), handoff_reader_policy(pro_prompts))
        run = C.AgentController(clean_session(), router).run("应收账款期末余额是多少？")
        handoff = run.report["routing"]["escalations"][0]["handoff"]
        prompt = pro_prompts[0]
        self.assertIn("TASK STATE", handoff)
        self.assertIn("应收账款期末余额是多少", handoff)                                    # the goal
        self.assertIn("150.0万元", handoff)                                               # the draft that failed
        self.assertRegex(handoff, r"ref=\w+ period=2025-12-31 basis=authority_adjusted .*value=1300000")   # what it cited, to cite again
        self.assertIn("get_account_evidence", handoff)                                   # what was already done
        self.assertIn("AMOUNT", handoff)                                                  # why the verifier objected
        self.assertIn("[rewrite]", handoff)
        self.assertIn("budget left", handoff)
        self.assertIn(handoff, prompt)                                                    # it is what the second model reads
        # none of the first model's transcript is in front of the second one
        self.assertNotIn("TOOL get_account_evidence(", prompt)
        self.assertNotIn("CONTROLLER: answer not accepted", prompt)
        self.assertLess(R.estimate_tokens(handoff), R.Policy().handoff_max_tokens)
        self.assertEqual(record_size(run), len(handoff))

    def test_the_hand_off_stays_within_its_size_however_much_there_is_to_say(self):
        claims = [{"kind": "amount", "periods": [P2], "cites": [f"F{i}"], "text": "金额" * 80} for i in range(8)]
        cited = [{"ref": f"F{i}", "period": P2, "basis": "raw_schedule", "line": f"明细{i}" * 6, "value": i * 1000} for i in range(40)]
        actions = [{"tool": "get_account_evidence", "args": {"account": "x" * 200}, "ok": True} for _ in range(30)]
        text = R.build_handoff(goal="问" * 3000, mode="ask", account="a", source_version="v1", draft={"claims": claims, "objections": ["x" * 900] * 12},
                               cited=cited, actions=actions, failures=[{"code": "SIGN_ERROR", "class": "rewrite", "count": 2}],
                               issues=[{"severity": "blocking", "account": "a", "message": "m" * 500, "resolved": False}] * 10,
                               remaining={"calls": 3, "tokens": None}, max_tokens=900)
        self.assertLessEqual(R.estimate_tokens(text), 900 * 1.15)
        self.assertIn("TASK STATE", text)
        self.assertIn("budget left: calls 3", text)

    def test_evidence_that_exists_but_was_not_read_is_fetched_by_the_controller_not_by_a_stronger_model(self):
        calls, seen = [], []
        session = clean_session()
        real_ref = total_ref(cl.evidence_view(session.contracts[ACCOUNT]), P2)

        def skips_reading(user, results):
            seen.append(user)
            return {"answer": {"claims": [{"account": ACCOUNT, "kind": "amount", "periods": [P2], "cites": [real_ref],
                                           "text": "2025年12月31日应收账款余额130.0万元"}], "unresolved": []}}
        router = routed(Policy(skips_reading), Policy(lambda u, r: calls.append(u) or tool("inspect_sources")))
        run = C.AgentController(session, router).run("应收账款？")
        self.assertTrue(run.ok, run.report)                                  # the second answer cites what the controller fetched
        self.assertEqual(calls, [])
        self.assertEqual(router.budget.escalations, [])
        self.assertEqual(run.steps[0].get("controller_fetched"), [ACCOUNT])
        self.assertIn("fetched for you", seen[1])
        self.assertEqual(router.turn_log[0]["classes"], ["evidence"])

    def test_a_ref_that_exists_nowhere_after_the_evidence_was_read_is_a_capability_failure_and_escalates_once(self):
        calls = []

        def invents(user, results):
            view = evidence_of(results)
            if view is None:
                return tool("get_account_evidence", account=ACCOUNT)
            return {"answer": {"claims": [{"account": ACCOUNT, "kind": "amount", "periods": [P2], "cites": ["F999"], "text": "2025年12月31日应收账款余额130.0万元"}],
                               "unresolved": []}}
        router = routed(Policy(invents), Policy(lambda u, r: calls.append(u) or tool("inspect_sources")))
        run = C.AgentController(clean_session(), router).run("应收账款？")
        self.assertFalse(run.ok)
        self.assertEqual(len(router.budget.escalations), 1)
        self.assertEqual(router.budget.escalations[0]["reason"], "capability")
        self.assertTrue(calls)
        self.assertIn(run.report["routing"]["stop_class"], ("capability_insufficient_after_escalation", "evidence_insufficient"))

    def test_a_cause_the_tools_were_asked_about_and_found_nothing_for_is_absent_never_escalated(self):
        calls = []

        def keeps_the_cause(user, results):
            names = [n for n, _ in results]
            view = evidence_of(results)
            if view is None:
                return tool("get_account_evidence", account=ACCOUNT)
            if "search_notes" not in names:
                return tool("search_notes", query="不存在的原因关键词", account=ACCOUNT)
            return {"answer": {"claims": [{"account": ACCOUNT, "kind": "amount", "periods": [P2], "cites": [total_ref(view, P2)],
                                           "text": "2025年12月31日应收账款余额130.0万元，主要由于新增大客户"}], "unresolved": []}}
        session = clean_session()
        for contract in session.contracts.values():
            contract.notes = []                                             # no remark anywhere: the cause has nothing behind it
        router = routed(Policy(keeps_the_cause), Policy(lambda u, r: calls.append(u) or tool("inspect_sources")))
        from fdd_utils.agent import tools as T
        nothing = lambda session, query, account="", max_hits=12: {"query": query, "hits": [], "total_matches": 0, "source": {}, "limits": {}}
        with mock.patch.object(T.TOOLS["search_notes"], "fn", nothing):    # the synthetic session has no file to scan; the search finds nothing
            run = C.AgentController(session, router).run("应收账款为什么增加？")
        self.assertFalse(run.ok)
        self.assertEqual(calls, [])
        self.assertEqual(router.budget.escalations, [])
        self.assertTrue(any(s.get("absent") == ["CAUSE_UNSUPPORTED"] for s in run.steps), run.steps)
        self.assertEqual(run.report["routing"]["stop_class"], "evidence_insufficient")

    def test_a_conflict_only_a_person_can_settle_never_sends_a_turn_to_the_stronger_model(self):
        session = WrongFinancials().session()
        calls = []

        def forgetful(user, results):
            names = [n for n, _ in results]
            if "inspect_sources" not in names:
                return tool("inspect_sources")
            if "get_account_evidence" not in names:
                return tool("get_account_evidence", account=ACCOUNT)
            return {"answer": {"claims": [{"account": ACCOUNT, "kind": "descriptive", "periods": [], "cites": [], "text": "已查看应收账款"}], "unresolved": []}}
        router = routed(Policy(forgetful), Policy(lambda u, r: calls.append(u) or tool("inspect_sources")))
        run = C.AgentController(session, router).run("应收账款？")
        self.assertEqual(run.stop_reason, C.STOP_REJECTED)
        self.assertEqual(calls, [])
        self.assertEqual(run.report["routing"]["stop_class"], "needs_human")

    def test_single_model_mode_never_escalates_and_says_the_model_did_not_manage(self):
        router = routed(wrong_amount_policy(), None, mode="single")
        run = C.AgentController(clean_session(), router).run("应收账款期末余额是多少？")
        self.assertEqual(run.stop_reason, C.STOP_REJECTED)
        self.assertEqual({r.role for r in router.budget.records}, {"primary"})
        self.assertEqual(run.report["routing"]["stop_class"], "capability_insufficient")
        self.assertEqual(run.report["routing"]["escalations"], [])

    def test_the_stronger_model_failing_too_is_reported_as_that_and_does_not_escalate_again(self):
        router = routed(wrong_amount_policy(), wrong_amount_policy_from_handoff())
        run = C.AgentController(clean_session(), router).run("应收账款期末余额是多少？")
        self.assertNotEqual(run.stop_reason, C.STOP_ANSWERED)
        self.assertEqual(len(router.budget.escalations), 1)
        self.assertEqual(run.report["routing"]["stop_class"], "capability_insufficient_after_escalation")

    def test_a_session_wide_escalation_cap_of_zero_keeps_every_turn_on_the_first_model(self):
        router = routed(wrong_amount_policy(), handoff_reader_policy([]), limits=R.Limits(max_escalations=0))
        run = C.AgentController(clean_session(), router).run("应收账款期末余额是多少？")
        self.assertEqual(router.budget.escalations, [])
        self.assertEqual(run.report["routing"]["stop_class"], "capability_insufficient")

    def test_an_unusable_reply_twice_is_the_same_kind_of_failure_and_the_discarded_replies_are_counted(self):
        def raw(text):
            return lambda system, user, max_tokens: {"content": text, "prompt_tokens": 50, "completion_tokens": 5}
        router = R.RoutedModel(spec(R.ROLE_PRIMARY, raw("no json here"), "small"), spec(R.ROLE_ESCALATION, handoff_reader_policy([]), "big"),
                               policy=R.Policy(mode="escalate"), budget=R.SharedBudget())
        run = C.AgentController(clean_session(), router).run("应收账款期末余额是多少？")
        self.assertEqual(len(router.budget.escalations), 1)
        self.assertEqual(router.budget.escalations[0]["codes"], ["NOT_JSON"])
        self.assertEqual(router.budget.summary()["replies_discarded"], 2)
        self.assertIn(run.stop_reason, (C.STOP_ANSWERED, C.STOP_NO_PROGRESS, C.STOP_REJECTED, C.STOP_STEPS))

    def test_a_step_that_makes_progress_clears_the_count(self):
        router = routed(wrong_amount_policy(), handoff_reader_policy([]), escalate_after=3)
        event = R.StepEvent
        self.assertFalse(router.observe(event(1, ["SIGN_ERROR"])).escalate)
        self.assertFalse(router.observe(event(2, ["SIGN_ERROR"])).escalate)
        self.assertFalse(router.observe(event(3, [], progress=True)).escalate)             # evidence arrived
        self.assertFalse(router.observe(event(4, ["SIGN_ERROR"])).escalate)
        self.assertFalse(router.observe(event(5, ["SIGN_ERROR"])).escalate)
        self.assertTrue(router.observe(event(6, ["SIGN_ERROR"])).escalate)
        router.begin_turn(2)
        self.assertEqual(router.role, R.ROLE_PRIMARY, "each turn starts on the first model")

    def test_a_step_that_also_waits_on_evidence_or_a_person_does_not_count(self):
        router = routed(wrong_amount_policy(), handoff_reader_policy([]))
        for step in range(1, 6):
            self.assertFalse(router.observe(R.StepEvent(step, ["SIGN_ERROR", "UNSEEN_REF"])).escalate)
            self.assertFalse(router.observe(R.StepEvent(step, ["conflict:pairing"])).escalate)

    def test_routing_set_to_escalate_without_an_escalation_model_is_an_error_not_a_silent_single_mode(self):
        with self.assertRaisesRegex(ValueError, "no escalation model"):
            routed(wrong_amount_policy(), None, mode="escalate")


def wrong_amount_policy_from_handoff():
    def fn(user, results):
        found = re.search(r"ref=(\w+) period=2025-12-31 basis=authority_adjusted", user)
        if not found:
            return tool("get_account_evidence", account=ACCOUNT)
        return {"answer": {"claims": [{"account": ACCOUNT, "kind": "amount", "periods": [P2], "cites": [found.group(1)],
                                       "text": "2025年12月31日应收账款余额150.0万元"}], "unresolved": []}}
    return Policy(fn)


def record_size(run) -> int:
    return run.report["routing"]["escalations"][0]["handoff_chars"]


# -- the books ----------------------------------------------------------------------------------------------------------------

def reply(text="ok", **extra):
    return lambda system, user, max_tokens: {"content": text, **extra}


class Books(unittest.TestCase):
    def call(self, router, n=1, system="s" * 70, user="u" * 700, max_tokens=100):
        for _ in range(n):
            router(system, user, max_tokens)

    def test_provider_numbers_and_estimates_are_kept_apart_and_never_added_together(self):
        a = R.RoutedModel(spec("primary", reply("答" * 20, prompt_tokens=1000, completion_tokens=50, cached_prompt_tokens=400, reasoning_tokens=30), "m1"))
        self.call(a)
        a.roles["primary"] = spec("primary", reply("答" * 20), "m1")                  # a provider that reports nothing
        self.call(a)
        books = a.budget.summary()
        self.assertEqual(books["provider_reported"], {"calls": 1, "prompt": 1000, "completion": 50, "cached_prompt": 400, "reasoning": 30})
        self.assertEqual(books["calls_without_provider_usage"], 1)
        self.assertEqual(books["estimated"]["calls"], 2)                                 # an estimate exists for every call sent, so the two can be compared
        self.assertGreater(books["estimated"]["prompt"], 0)
        self.assertEqual(books["charged_tokens"], 1050 + a.budget.records[1].prompt_est + a.budget.records[1].completion_est)
        self.assertIn("provider-reported where the provider reported", books["charged_basis"])
        self.assertTrue(any("SDK" in note for note in books["unmeasured"]))

    def test_the_cap_is_applied_to_what_the_provider_reported_not_to_the_estimate(self):
        router = R.RoutedModel(spec("primary", reply("x", prompt_tokens=5000, completion_tokens=100), "m"), budget=R.SharedBudget(R.Limits(max_tokens=5300)))
        self.call(router)                                                             # estimate ~ 250; the provider said 5,100
        self.assertEqual(router.budget.charged_tokens, 5100)
        with self.assertRaises(R.BudgetExhausted) as caught:
            self.call(router)
        self.assertEqual(caught.exception.kind, "tokens")
        self.assertEqual(router.budget.calls, 1, "a call refused before it was made is not a call")

    def test_a_call_that_failed_is_counted_and_charged_its_prompt_estimate(self):
        def boom(system, user, max_tokens):
            raise TimeoutError("gateway timeout")
        router = R.RoutedModel(spec("primary", boom, "m"))
        with self.assertRaises(TimeoutError):
            self.call(router)
        record = router.budget.records[0]
        self.assertFalse(record.ok)
        self.assertEqual((record.usage_source, record.charged), ("failed", record.prompt_est))
        self.assertGreater(record.prompt_est, 0)
        self.assertEqual(router.budget.summary()["failed_calls"], 1)
        self.assertIn("TimeoutError", record.error)

    def test_a_failure_message_keeps_its_type_and_loses_any_gateway_address(self):
        def boom(system, user, max_tokens):
            raise ConnectionError("could not reach https://gateway.internal.example/v1/chat/completions after 2 tries")
        router = R.RoutedModel(spec("primary", boom, "m"))
        with self.assertRaises(ConnectionError):
            self.call(router)
        error = router.budget.records[0].error
        self.assertTrue(error.startswith("ConnectionError: could not reach <address>"), error)
        self.assertNotIn("gateway.internal", json.dumps(router.budget.to_dict()))

    def test_an_answer_from_the_response_cache_costs_nothing_and_is_counted_apart(self):
        router = R.RoutedModel(spec("primary", reply("x", prompt_tokens=0, completion_tokens=0, token_source="response_cache",
                                                       saved_prompt_tokens=800, saved_completion_tokens=40), "m"))
        self.call(router)
        books = router.budget.summary()
        self.assertEqual(books["charged_tokens"], 0)
        self.assertEqual(books["response_cache"], {"hits": 1, "saved_prompt": 800, "saved_completion": 40})
        self.assertEqual(books["provider_reported"]["calls"], 0)
        self.assertEqual(books["estimated"]["calls"], 0)

    def test_every_role_and_every_attempt_draws_on_the_same_cap(self):
        limits = R.Limits(max_calls=4)
        pro_prompts = []
        router = routed(wrong_amount_policy(), handoff_reader_policy(pro_prompts), limits=limits)
        run = C.AgentController(clean_session(), router).run("应收账款期末余额是多少？")
        self.assertEqual(router.budget.calls, 4)
        self.assertEqual(run.stop_reason, C.STOP_ANSWERED, "3 first-model calls and 1 escalation call fit exactly")
        # a second turn on the same books: the first model's call is refused because the escalation model already used the allowance
        again = C.AgentController(clean_session(), router).run("再问一次")
        self.assertEqual(again.stop_reason, C.STOP_SESSION)
        self.assertEqual(again.report["budget"]["kind"], "calls")
        self.assertEqual(router.budget.summary()["by_role"]["escalation"]["calls"], 1)
        self.assertIn("a cap over the whole working session", C.AgentController(clean_session(), router)._stop_report(again)["why_stopped"])

    def test_a_cost_cap_needs_a_price_for_every_role_in_use_and_never_borrows_one(self):
        with self.assertRaisesRegex(ValueError, "no price is configured for: primary"):
            R.RoutedModel(spec("primary", reply(), "m", price=None), budget=R.SharedBudget(R.Limits(max_cost=1.0)))
        router = R.RoutedModel(spec("primary", reply("x", prompt_tokens=1_000_000, completion_tokens=100_000), "m", price=(12.0, 24.0)),
                               budget=R.SharedBudget(R.Limits(max_cost=14.4005)))
        self.call(router, max_tokens=10)
        self.assertEqual(router.budget.cost, 14.4)                                        # 1.0M x 12 + 0.1M x 24 yuan per million
        with self.assertRaises(R.BudgetExhausted) as caught:
            self.call(router, max_tokens=10)
        self.assertEqual(caught.exception.kind, "cost")

    def test_the_books_survive_a_resumed_session_so_a_cap_is_a_cap_on_the_whole_work(self):
        router = R.RoutedModel(spec("primary", reply("x", prompt_tokens=10, completion_tokens=1), "m"), budget=R.SharedBudget(R.Limits(max_calls=3)))
        self.call(router, n=2)
        saved = json.loads(json.dumps(router.budget.to_dict()))
        fresh = R.RoutedModel(spec("primary", reply("x", prompt_tokens=10, completion_tokens=1), "m"), budget=R.SharedBudget(R.Limits(max_calls=3)))
        fresh.budget.restore(saved)
        self.assertEqual((fresh.budget.calls, fresh.budget.charged_tokens), (2, 22))
        self.call(fresh)
        with self.assertRaises(R.BudgetExhausted):
            self.call(fresh)

    def test_the_records_hold_numbers_and_names_and_no_prompt_text(self):
        router = R.RoutedModel(spec("primary", reply("答案正文", prompt_tokens=10, completion_tokens=1), "m"))
        router("SYSTEM-SECRET-TEXT", "USER-SECRET-TEXT", 50)
        blob = json.dumps(router.budget.to_dict(), ensure_ascii=False)
        self.assertNotIn("SECRET", blob)
        self.assertNotIn("答案正文", blob)


# -- thinking, and what the ledger is told -------------------------------------------------------------------------------------

class Thinking(unittest.TestCase):
    def test_what_each_provider_branch_actually_sends(self):
        self.assertEqual(C.thinking_params("dashscope", ""), ({}, False))                           # nothing requested, nothing sent
        self.assertEqual(C.thinking_params("dashscope", "off"), ({}, False))                        # asked for, but config declares no field for it
        self.assertEqual(C.thinking_params("deepseek", "high"), ({}, False))
        # a provider block that declares its thinking body (thinking_extra_body) gets the request, labelled unverified elsewhere
        self.assertEqual(C.thinking_params("dashscope", "on", declared=True), ({"allow_thinking": True}, True))
        self.assertEqual(C.thinking_params("dashscope", "off", declared=True), ({"allow_thinking": False}, True))
        self.assertEqual(C.thinking_params("deepseek", "high", declared=True), ({"allow_thinking": True, "reasoning_effort": "high"}, True))
        self.assertEqual(C.thinking_params("dashscope", "", declared=True), ({}, False))
        self.assertEqual(C.thinking_params("workbench", "on", reasoning_effort="high"), ({"allow_thinking": True, "reasoning_effort": "high"}, True))
        self.assertEqual(C.thinking_params("local", "off"), ({"allow_thinking": False}, True))
        self.assertEqual(C.thinking_params("local", "on"), ({"allow_thinking": True}, True))
        self.assertEqual(C.thinking_params("workbench", "low"), ({"allow_thinking": True, "reasoning_effort": "low"}, True))

    def test_a_role_that_asked_for_thinking_a_branch_does_not_send_says_so_in_its_records(self):
        router = R.RoutedModel(R.RoleSpec("primary", reply(), provider="dashscope", model="m", thinking="off", thinking_applied=False))
        router("s", "u", 10)
        self.assertEqual(router.budget.records[0].thinking, "off (not sent by this provider branch)")
        sent = R.RoutedModel(R.RoleSpec("primary", reply(), provider="local", model="m", thinking="off", thinking_applied=True))
        sent("s", "u", 10)
        self.assertEqual(sent.budget.records[0].thinking, "off")

    def test_the_agents_calls_are_labelled_as_the_agent_in_the_ledger_and_not_as_the_validator(self):
        from fdd_utils import usage_ledger
        seen = {}

        class FakeClient:
            def __init__(self, **kw):
                seen["init"] = kw
                self.config_details = {"chat_model": "gateway-model-id"}

            def get_response(self, user, system, **kw):
                seen["kwargs"] = kw
                usage_ledger.record_llm_call(client_agent="subagent_4", ok=True, prompt_tokens=1, completion_tokens=1)
                return {"content": "x", "prompt_tokens": 1, "completion_tokens": 1, "cached_prompt_tokens": 7, "reasoning_tokens": 3,
                        "token_usage_source": "provider"}
        tmp = Path(tempfile.mkdtemp()) / "usage.jsonl"
        with mock.patch("fdd_utils.ai.client.AIClient", FakeClient), mock.patch.object(usage_ledger, "USAGE_PATH", tmp):
            model = C.AIClientModel("dashscope", "m", thinking="off", ledger_stage="agent_escalation")
            out = model("s", "u", 10)
        row = json.loads(tmp.read_text(encoding="utf-8").splitlines()[0])
        self.assertEqual(row["stage"], "agent_escalation")
        self.assertNotEqual(row["stage"], "Validator")
        self.assertNotIn("allow_thinking", seen["kwargs"])                  # dashscope with no declared thinking body: nothing is sent
        self.assertNotIn("reasoning_effort", seen["kwargs"])
        self.assertEqual(seen["init"]["agent_name"], C.AGENT_IDENTITY)      # its own identity, not subagent_4's
        self.assertEqual((out["cached_prompt_tokens"], out["reasoning_tokens"]), (7, 3))
        self.assertEqual(model.chat_model, "gateway-model-id")


# -- the working session ---------------------------------------------------------------------------------------------------------

def session_with_book():
    from test_workspace import session_on, write_book
    tmp = tempfile.mkdtemp(prefix="route_")
    book = os.path.join(tmp, "wb.xlsx")
    write_book(book)
    source = session_on(book)
    source.load_params = {"workbook": book}
    return source, tmp


class InAWorkingSession(unittest.TestCase):
    def test_a_turn_records_which_model_answered_why_it_escalated_what_it_was_handed_and_what_each_call_used(self):
        source, tmp = session_with_book()
        pro_prompts = []
        work = W.WorkSession(source, model=wrong_amount_policy(), escalation=handoff_reader_policy(pro_prompts), model_name="small-model",
                             policy=R.Policy(mode="escalate"), budget=C.Budget(max_steps=8))
        turn = work.ask("应收账款期末余额是多少？")
        self.assertTrue(turn.ok, turn.report)
        routing = turn.report["routing"]
        self.assertEqual(routing["stop_class"], "ok")
        self.assertEqual(routing["escalations"][0]["to"], "escalation")
        self.assertIn("TASK STATE", routing["escalations"][0]["handoff"])
        self.assertEqual([r["role"] for r in turn.usage["call_records"]], ["primary", "primary", "primary", "escalation"])
        status = work.status()
        self.assertEqual(status["routing"], "escalate")
        self.assertEqual(status["budget"]["escalations"], 1)
        self.assertEqual(set(status["budget"]["by_role"]), {"primary", "escalation"})
        self.assertEqual(work.model_calls, 4)

    def test_the_saved_session_keeps_the_books_without_any_prompt_and_a_resumed_one_keeps_counting(self):
        source, tmp = session_with_book()
        path = os.path.join(tmp, "s.agent-session.json")
        work = W.WorkSession(source, model=wrong_amount_policy(), escalation=handoff_reader_policy([]), policy=R.Policy(mode="escalate"),
                             path=path, limits=R.Limits(max_calls=5), budget=C.Budget(max_steps=8))
        self.assertTrue(work.ask("应收账款期末余额是多少？").ok)
        saved = json.load(open(path, encoding="utf-8"))
        self.assertEqual(len(saved["budget"]["records"]), 4)
        self.assertEqual(len(saved["budget"]["escalations"]), 1)
        self.assertNotIn("system", json.dumps(saved["budget"]))
        with mock.patch.object(SourceSession, "load", classmethod(lambda cls, **_k: source)):
            resumed = W.WorkSession.open(source.workbook_path, session_path=path, model=wrong_amount_policy(), limits=R.Limits(max_calls=5))
        self.assertEqual(resumed.model_calls, 4)
        self.assertEqual(resumed.shared.remaining()["calls"], 1)

    def test_the_cap_on_calls_set_the_old_way_is_the_same_cap_in_the_books(self):
        source, tmp = session_with_book()
        work = W.WorkSession(source, model=wrong_amount_policy(), max_model_calls=1, budget=C.Budget(max_steps=8))
        turn = work.ask("应收账款期末余额是多少？")
        self.assertEqual(turn.stop_reason, C.STOP_SESSION)
        self.assertEqual(turn.report["budget"]["kind"], "calls")
        self.assertEqual(work.model_calls, 1)

    def test_the_config_helpers_read_the_agent_block_and_leave_the_default_alone(self):
        from fdd_utils.agent import flag as F
        self.assertEqual(F.agent_policy({}).mode, "single")
        cfg = {"agent": {"routing": "escalate", "escalate_after": 3, "max_session_tokens": 400000, "max_session_cost": 0, "max_model_calls": 60,
                         "models": {"escalation": {"provider": "dashscope", "model": "pro-id", "thinking": "off"}},
                         }, "dashscope": {"available_models": [{"id": "pro-id", "price_in": 12, "price_out": 24}, {"id": "flash-id"}]}}
        self.assertEqual((F.agent_policy(cfg).mode, F.agent_policy(cfg).escalate_after), ("escalate", 3))
        limits = F.agent_limits(cfg)
        self.assertEqual((limits.max_tokens, limits.max_cost, limits.max_calls), (400000, None, 60))
        self.assertEqual(F.role_config(cfg, "escalation"), {"provider": "dashscope", "model": "pro-id", "thinking": "off",
                                                             "reasoning_effort": "", "max_output": None})
        cfg["agent"]["models"]["primary"] = {"provider": "dashscope", "model": "flash-id", "reasoning_effort": "low", "max_output": 900}
        self.assertEqual(F.role_config(cfg, "primary"), {"provider": "dashscope", "model": "flash-id", "thinking": "",
                                                          "reasoning_effort": "low", "max_output": 900})
        self.assertEqual(F.agent_policy(cfg).escalate_tasks, [])
        cfg["agent"]["escalate_tasks"] = ["adjustment_bridge"]
        self.assertEqual(F.agent_policy(cfg).escalate_tasks, ["adjustment_bridge"])
        self.assertEqual(F.price_for(cfg, "dashscope", "pro-id"), (12.0, 24.0))
        self.assertIsNone(F.price_for(cfg, "dashscope", "flash-id"), "an entry with no price has none; nothing is borrowed")
        self.assertIsNone(F.price_for(cfg, "deepseek", "pro-id"))
        self.assertEqual(F.parse_role("dashscope::flash-id"), {"provider": "dashscope", "model": "flash-id", "thinking": "",
                                                               "reasoning_effort": "", "max_output": None})
        with self.assertRaises(ValueError):
            F.agent_policy({"agent": {"routing": "both"}})


if __name__ == "__main__":
    unittest.main(verbosity=2)

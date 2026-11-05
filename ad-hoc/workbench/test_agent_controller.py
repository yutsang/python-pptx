"""The single controller: results change the path, and what code can enforce is enforced (Phase 5).

Zero tokens: the "model" is a scripted policy that reads the conversation the controller
shows it and replies with one JSON object. That is enough to test the controller -- the
gates, the open issues, the budgets, the stop reports -- and to show that two situations
send the same policy down different tool paths because of what the tools returned. It is
NOT evidence about how a real model chooses; that needs a provider (see the handoff).

    python ad-hoc/workbench/test_agent_controller.py
"""
from __future__ import annotations

import json
import os
import re
import sys
import tempfile
import unittest
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parents[1]))

from fdd_utils.agent import controller as C  # noqa: E402
from fdd_utils.agent.session import SourceSession  # noqa: E402
from fdd_utils.ai import evidence_contract as ec  # noqa: E402
from fdd_utils.run_identity import file_sha256  # noqa: E402
from test_claims import P1, P2, ROWS, build  # noqa: E402

ADJUSTMENT = "「Audit ADJ#4」其他应收款重分类至应收账款 +20.0万元"
REPLACED = {P1: {"tab": 800000.0, "financials": 1000000.0}}
REPLACED_ROWS = [("甲公司", "detail", {P1: 500000, P2: 700000}), ("乙公司", "detail", {P1: 300000, P2: 300000}),
                 ("丙公司", "detail", {P1: 0, P2: 300000}), ("Total", "total", {P1: 1000000, P2: 1300000})]


def session_for(contract, workbook_path: str = "") -> SourceSession:
    binding = ec.SourceBinding(entity="E", workbook="wb.xlsx", workbook_sha256=(file_sha256(workbook_path) if workbook_path else "aaaa"),
                               financials_sheet="Financials", statement_source="financials")
    return SourceSession(binding=binding, dfs={}, contracts={contract.account: contract}, workbook_path=workbook_path, language="Chi")


def clean_session():
    return session_for(build())


def replaced_session(with_adjustment: bool):
    authority = {P1: dict(REPLACED[P1], **({"adjustments": [ADJUSTMENT]} if with_adjustment else {}))}
    return session_for(build(REPLACED_ROWS, financials={P1: 1000000.0, P2: 1300000.0}, authority=authority, notes=None))


# --- a scripted policy ---------------------------------------------------------------------------------------

def returned(user: str):
    """[(tool, parsed result)] for every tool result in the conversation shown to the model."""
    out = []
    for m in re.finditer(r"TOOL (\w+)\(.*?\) RETURNED:\n(\{.*?\})(?=\n\nTOOL |\n\nCONTROLLER|\n\nSYSTEM|\n\nYour next JSON)", user, re.S):
        try:
            out.append((m.group(1), json.loads(m.group(2))))
        except ValueError:
            pass
    return out


def tool(name, **args):
    return {"tool": name, "args": args}


class Policy:
    """Calls fn(user, results) -> action dict; records the replies it gave."""

    def __init__(self, fn):
        self.fn, self.calls = fn, 0

    def __call__(self, system, user, max_tokens):
        self.calls += 1
        return {"content": json.dumps(self.fn(user, returned(user)), ensure_ascii=False), "prompt_tokens": 100, "completion_tokens": 20}


def evidence_of(results):
    for name, result in reversed(results):
        if name == "get_account_evidence" and result.get("ok"):
            return result["data"]
    return None


def total_ref(view, period):
    return next(t["ref"] for t in view["totals"] if t["period"] == period and t["basis"] == "authority_adjusted")


def investigator(user, results):
    """inspect -> evidence -> (adjustment note? cite it : say it is unresolved) -> answer. Reacts to what came back."""
    seen = [n for n, _ in results]
    if "inspect_sources" not in seen:
        return tool("inspect_sources")
    view = evidence_of(results)
    if view is None:
        return tool("get_account_evidence", account="应收账款")
    claims = [{"account": "应收账款", "kind": "amount", "periods": [P2], "cites": [total_ref(view, P2)],
               "text": "2025年12月31日应收账款余额130.0万元"}]
    adjustment = next((n for n in view["notes"] if n["kind"] == "adjustment"), None)
    replaced = any(row.get("total_vs_financials") == "mismatch" for row in view["checks"].values())
    if not replaced:
        return {"answer": {"claims": claims, "unresolved": []}}
    if adjustment:
        claims.append({"account": "应收账款", "kind": "cause", "periods": [P1], "cites": [total_ref(view, P1)],
                       "cause_cites": [adjustment["ref"]], "text": "2024年12月31日余额按报表列示，差异系其他应收款重分类至应收账款所致"})
        return {"answer": {"claims": claims, "unresolved": []}}
    return {"answer": {"claims": claims, "unresolved": ["应收账款 2024年末报表金额与本表不一致，工作底稿中找不到调整分录，原因未能核实"]}}


class Paths(unittest.TestCase):
    """The same policy, two situations, two tool paths -- and different answers."""

    def run_policy(self, session, policy=investigator, **kw):
        return C.AgentController(session, Policy(policy)).run("应收账款的变动来自哪里？", **kw)

    def test_clean_account_is_answered_after_looking_once(self):
        run = self.run_policy(clean_session())
        self.assertEqual(run.stop_reason, C.STOP_ANSWERED, run.steps)
        self.assertEqual(run.tool_path, ["inspect_sources", "get_account_evidence"])
        self.assertEqual(run.answer["verification"]["failed"], 0)
        self.assertTrue(run.answer["sources"])
        self.assertEqual(run.answer["sources"][0]["row"], 7)      # traced to a sheet row, not just a number

    def test_a_replaced_total_with_an_adjustment_note_is_explained_from_the_note(self):
        run = self.run_policy(replaced_session(with_adjustment=True))
        self.assertEqual(run.stop_reason, C.STOP_ANSWERED, run.steps)
        causes = [c for c in run.answer["claims"] if c["kind"] == "cause"]
        self.assertEqual(len(causes), 1)
        self.assertTrue(causes[0]["cause_cites"])
        self.assertEqual(run.answer["unresolved"], [])
        self.assertEqual(run.answer["verification"]["failed"], 0)

    def test_a_replaced_total_with_no_adjustment_is_reported_as_unresolved_not_invented(self):
        run = self.run_policy(replaced_session(with_adjustment=False))
        self.assertEqual(run.stop_reason, C.STOP_ANSWERED, run.steps)
        self.assertFalse([c for c in run.answer["claims"] if c["kind"] == "cause"])
        self.assertTrue(run.answer["unresolved"])

    def test_the_same_policy_ends_differently_because_of_what_the_tools_returned(self):
        a = self.run_policy(replaced_session(with_adjustment=True))
        b = self.run_policy(replaced_session(with_adjustment=False))
        c = self.run_policy(clean_session())
        self.assertEqual(a.tool_path, b.tool_path)                # same investigation...
        self.assertNotEqual(json.dumps(a.answer["claims"], ensure_ascii=False), json.dumps(b.answer["claims"], ensure_ascii=False))
        self.assertEqual((len(a.answer["unresolved"]), len(b.answer["unresolved"]), len(c.answer["unresolved"])), (0, 1, 0))


def adaptive(user, results):
    """The next TOOL depends on what the last one returned -- three situations, three investigations."""
    names = [n for n, _ in results]
    if "inspect_sources" not in names:
        return tool("inspect_sources")
    view = evidence_of(results)
    if view is None:
        return tool("get_account_evidence", account="应收账款")
    replaced = [p for p, row in view["checks"].items() if row.get("total_vs_financials") == "mismatch"]
    adjustment = next((n for n in view["notes"] if n["kind"] == "adjustment"), None)
    if not replaced:                                         # nothing odd: quantify the change
        if "compute_analysis" not in names:
            return tool("compute_analysis", account="应收账款", op="movement", period_from=P1, period=P2)
        ref = next(r for n, r in results if n == "compute_analysis")["data"]["derived"]["ref"]
        return {"answer": {"claims": [{"account": "应收账款", "kind": "movement", "periods": [P1, P2], "cites": [ref],
                                       "text": "较2024年末增加30.0万元，增幅30.0%"}], "unresolved": []}}
    period = replaced[0]
    if adjustment:                                           # the books say why: read the entry
        if "trace_claim" not in names:
            return tool("trace_claim", account="应收账款", ref=adjustment["ref"])
        return {"answer": {"claims": [{"account": "应收账款", "kind": "cause", "periods": [period], "cites": [total_ref(view, period)],
                                       "cause_cites": [adjustment["ref"]], "text": "2024年12月31日余额按报表列示，差异系其他应收款重分类至应收账款所致"}],
                           "unresolved": []}}
    if "compute_analysis" not in names:                      # nothing explains it: size the disagreement
        return tool("compute_analysis", account="应收账款", op="difference", period=period)
    ref = next(r for n, r in results if n == "compute_analysis")["data"]["derived"]["ref"]
    return {"answer": {"claims": [{"account": "应收账款", "kind": "difference", "periods": [period], "cites": [ref],
                                   "text": "2024年12月31日账面明细与报表相差20.0万元"}],
                       "unresolved": ["应收账款 2024年末差异的原因：工作底稿中没有对应的调整分录"]}}


class Divergence(unittest.TestCase):
    def test_what_a_result_shows_decides_which_tool_comes_next(self):
        runs = {name: C.AgentController(s, Policy(adaptive)).run("应收账款有什么要注意？")
                for name, s in (("clean", clean_session()), ("explained", replaced_session(True)), ("unexplained", replaced_session(False)))}
        for name, run in runs.items():
            self.assertEqual(run.stop_reason, C.STOP_ANSWERED, (name, run.steps))
            self.assertEqual(run.answer["verification"]["failed"], 0, name)
        paths = {name: tuple((s["tool"], (s.get("args") or {}).get("op")) for s in run.steps if s.get("tool")) for name, run in runs.items()}
        self.assertEqual(paths["clean"][-1], ("compute_analysis", "movement"))
        self.assertEqual(paths["explained"][-1], ("trace_claim", None))
        self.assertEqual(paths["unexplained"][-1], ("compute_analysis", "difference"))
        self.assertEqual(len(set(paths.values())), 3, "three situations, three different investigations")
        self.assertEqual(runs["unexplained"].answer["unresolved"][0][:4], "应收账款")


class Gates(unittest.TestCase):
    def test_answering_before_looking_is_sent_back_and_the_path_changes(self):
        def eager(user, results):
            if "CONTROLLER: answer not accepted" in user:
                return investigator(user, results)             # corrected by the gate (the note stays in the conversation)
            return {"answer": {"claims": [{"account": "应收账款", "kind": "amount", "periods": [P2], "cites": ["F1"],
                                           "text": "2025年12月31日应收账款余额130.0万元"}], "unresolved": []}}
        run = C.AgentController(clean_session(), Policy(eager)).run("应收账款余额？")
        self.assertEqual(run.stop_reason, C.STOP_ANSWERED)
        self.assertFalse(run.steps[0]["accepted"])
        self.assertIn("UNSEEN_REF", run.steps[0]["feedback"])
        # the ref it cited exists, its evidence had not been read: the controller fetched it (a free tool call, not a model step)
        self.assertEqual(run.steps[0].get("controller_fetched"), ["应收账款"])
        self.assertEqual(run.tool_path, ["inspect_sources"])

    def test_a_blocking_issue_must_be_investigated_or_named(self):
        def naive(user, results):
            view = evidence_of(results)
            if view is None:
                return tool("get_account_evidence", account="应收账款")
            if "CONTROLLER: answer not accepted" in user:
                return {"answer": {"claims": [{"account": "应收账款", "kind": "amount", "periods": [P2], "cites": [total_ref(view, P2)],
                                               "text": "2025年12月31日应收账款余额130.0万元"}],
                                   "unresolved": ["应收账款 2024年末的报表差异原因未核实"]}}
            return {"answer": {"claims": [{"account": "应收账款", "kind": "amount", "periods": [P2], "cites": [total_ref(view, P2)],
                                           "text": "2025年12月31日应收账款余额130.0万元"}], "unresolved": []}}
        run = C.AgentController(replaced_session(False), Policy(naive)).run("应收账款？")
        self.assertEqual(run.stop_reason, C.STOP_ANSWERED)
        self.assertEqual([s.get("accepted") for s in run.steps if s.get("action") == "answer"], [False, True])
        self.assertIn("open blocking issue", run.steps[1]["feedback"] if "feedback" in run.steps[1] else run.steps[-2]["feedback"])

    def test_a_failed_claim_is_sent_back_never_softened(self):
        def slip(user, results):
            view = evidence_of(results)
            if view is None:
                return tool("get_account_evidence", account="应收账款")
            text = "2025年12月31日应收账款余额130.0亿元" if "CONTROLLER" not in user else "2025年12月31日应收账款余额130.0万元"
            return {"answer": {"claims": [{"account": "应收账款", "kind": "amount", "periods": [P2], "cites": [total_ref(view, P2)], "text": text}],
                               "unresolved": []}}
        run = C.AgentController(clean_session(), Policy(slip)).run("余额？")
        self.assertEqual(run.stop_reason, C.STOP_ANSWERED)
        self.assertIn("UNIT_ERROR", run.steps[1]["feedback"])
        self.assertIn("万元", run.answer["text"])

    def test_a_stated_cause_needs_a_note_or_must_be_marked_as_inference(self):
        def cause(user, results):
            view = evidence_of(results)
            if view is None:
                return tool("get_account_evidence", account="应收账款")
            marked = "CONTROLLER" in user
            return {"answer": {"claims": [{"account": "应收账款", "kind": "amount", "periods": [P2], "cites": [total_ref(view, P2)],
                                           "text": "2025年12月31日应收账款余额130.0万元，主要系租金回款延迟所致", "inference": marked}],
                               "unresolved": []}}
        run = C.AgentController(clean_session(), Policy(cause)).run("为什么增加？")
        self.assertEqual(run.stop_reason, C.STOP_ANSWERED)
        self.assertIn("CAUSE_UNSUPPORTED", run.steps[1]["feedback"])
        self.assertEqual(run.answer["verification"]["inference"], 1)

    def test_answers_sent_back_too_often_end_the_run_with_a_report(self):
        def stubborn(user, results):
            return {"answer": {"claims": [{"account": "应收账款", "kind": "amount", "periods": [P2], "cites": ["F1"], "text": "余额130.0万元"}], "unresolved": []}}
        run = C.AgentController(clean_session(), Policy(stubborn), budget=C.Budget(max_rejections=2)).run("余额？")
        self.assertEqual(run.stop_reason, C.STOP_REJECTED)
        self.assertFalse(run.ok)
        self.assertIsNone(run.answer)
        self.assertIn("why_stopped", run.report)


class Asking(unittest.TestCase):
    def test_an_ordinary_gap_is_not_a_question_for_the_user(self):
        def lazy(user, results):
            if "CONTROLLER: do not ask yet" in user:
                return investigator(user, results)             # told to investigate; the note stays in the conversation
            return {"ask_user": "你想看哪个科目？"}
        run = C.AgentController(clean_session(), Policy(lazy)).run("给我看看")
        self.assertEqual(run.stop_reason, C.STOP_ANSWERED)
        self.assertFalse(run.steps[0]["allowed"])

    def test_a_real_conflict_may_be_put_to_the_user(self):
        conflicted = session_for(build(financials={P1: 1000000.0, P2: 1500000.0}))     # tab 130万 vs Financials 150万, no rule chose
        def asker(user, results):
            if evidence_of(results) is None:
                return tool("get_account_evidence", account="应收账款")
            return {"ask_user": "本表与报表在2025年末不一致，应以哪一个为准？"}
        run = C.AgentController(conflicted, Policy(asker)).run("2025年末余额？")
        self.assertEqual(run.stop_reason, C.STOP_NEEDS_USER)
        self.assertIn("以哪一个为准", run.question)
        self.assertTrue(any(i["id"].startswith("conflict:") for i in run.issues))

    def test_an_ambiguous_account_is_asked_about_before_any_model_call(self):
        s = clean_session()
        s.contracts["应收票据"] = build(account="应收票据", sheet="Notes")
        model = Policy(lambda u, r: tool("inspect_sources"))
        run = C.AgentController(s, model).run("改写评论", mode="edit", account="应收")
        self.assertEqual((run.stop_reason, model.calls), (C.STOP_NEEDS_USER, 0))


class WrongFinancials(unittest.TestCase):
    def session(self):
        contracts = {name: build(account=name, sheet=name, financials={P1: 5_000_000.0, P2: 9_000_000.0}) for name in ("应收账款", "预付款项", "存货", "固定资产")}
        binding = ec.SourceBinding(entity="E", workbook_sha256="aaaa", financials_sheet="Financials", statement_source="financials")
        return SourceSession(binding=binding, dfs={}, contracts=contracts, language="Chi")

    def test_a_financials_sheet_that_ties_to_almost_nothing_is_flagged_and_may_be_asked_about(self):
        s = self.session()
        from fdd_utils.agent import tools as tools_module
        self.assertTrue(tools_module.financials_pairing(s)["suspect"])
        def policy(user, results):
            if not results:
                return tool("inspect_sources")
            return {"ask_user": "Financials 与各科目几乎都对不上，请确认这是该公司该期间的报表。"}
        run = C.AgentController(s, Policy(policy)).run("应收账款余额？")
        self.assertEqual(run.stop_reason, C.STOP_NEEDS_USER)
        self.assertTrue(any(i["id"] == "conflict:pairing" for i in run.issues))

    def test_an_answer_is_refused_while_the_pairing_is_in_doubt_unless_it_says_so(self):
        s = self.session()
        def policy(user, results):
            names = [n for n, _ in results]
            if "inspect_sources" not in names:
                return tool("inspect_sources")
            if "get_account_evidence" not in names:
                return tool("get_account_evidence", account="应收账款")
            return {"answer": {"claims": [{"account": "应收账款", "kind": "descriptive", "periods": [], "cites": [], "text": "已查看应收账款"}], "unresolved": []}}
        run = C.AgentController(s, Policy(policy), budget=C.Budget(max_rejections=1)).run("应收账款？")
        self.assertEqual(run.stop_reason, C.STOP_REJECTED)
        self.assertIn("neither investigated nor named", run.steps[-1]["feedback"])

    def test_a_deck_exported_against_a_suspect_sheet_is_blocked_however_good_its_claims(self):
        from fdd_utils.agent.export import merged_results, session_delivery
        from fdd_utils.agent.session import Candidate, claims_digest
        from fdd_utils.ai import claims as cl
        s = self.session()
        c = s.contracts["应收账款"]
        d = cl.derive_difference(c, P2)                  # against this sheet the balance cannot be stated, the gap can
        claims = [{"account": "应收账款", "kind": "difference", "periods": [P2], "cites": [d["ref"]],
                   "text": "2025年12月31日账面明细与报表相差770.0万元", "source_version": s.source_version}]
        outcome = cl.verify_candidate(claims, s.contracts, default_account="应收账款")
        self.assertTrue(outcome["ok"])
        s.candidates["应收账款"] = Candidate("应收账款", claims, outcome, s.source_version, claims_digest(claims))
        decision = session_delivery(s, merged_results(s, ["应收账款"]))
        self.assertEqual(decision.status, "blocked")
        self.assertIn("FINANCIALS_PAIRING_SUSPECT", [f.code for f in decision.findings])


class Budgets(unittest.TestCase):
    def test_repeating_the_same_call_stops_for_no_progress(self):
        run = C.AgentController(clean_session(), Policy(lambda u, r: tool("inspect_sources"))).run("?")
        self.assertEqual(run.stop_reason, C.STOP_NO_PROGRESS)
        self.assertEqual(len(run.steps), 3)

    def test_step_token_and_time_budgets_end_the_run_with_a_report(self):
        s = clean_session()
        steps = C.AgentController(s, Policy(investigator), budget=C.Budget(max_steps=1)).run("余额？")
        self.assertEqual(steps.stop_reason, C.STOP_STEPS)
        self.assertEqual(steps.report["accounts_consulted"], [])
        self.assertEqual(C.AgentController(s, Policy(investigator), budget=C.Budget(max_tokens=50)).run("?").stop_reason, C.STOP_TOKENS)
        self.assertEqual(C.AgentController(s, Policy(investigator), budget=C.Budget(max_seconds=-1)).run("?").stop_reason, C.STOP_TIME)

    def test_the_stop_report_lists_what_is_known_and_what_is_open(self):
        s = replaced_session(False)
        def looks_then_stalls(user, results):
            return tool("get_account_evidence", account="应收账款") if evidence_of(results) is None else tool("inspect_sources")
        run = C.AgentController(s, Policy(looks_then_stalls), budget=C.Budget(max_no_progress=2)).run("?")
        self.assertEqual(run.stop_reason, C.STOP_NO_PROGRESS)
        known = run.report["known"][0]
        self.assertEqual((known["account"], known["balance_yuan"]), ("应收账款", 1300000.0))
        self.assertTrue(any("2024-12-31" in g for g in run.report["gaps"]))
        self.assertTrue(run.report["next_steps"])

    def test_a_model_error_ends_the_run_cleanly(self):
        def broken(system, user, max_tokens):
            raise ConnectionError("gateway down")
        run = C.AgentController(clean_session(), broken).run("?")
        self.assertEqual(run.stop_reason, C.STOP_MODEL_ERROR)
        self.assertIn("gateway down", run.report["model_error"])

    def test_usage_is_counted_and_provider_figures_kept_apart_from_estimates(self):
        run = C.AgentController(clean_session(), Policy(investigator)).run("余额？")
        self.assertEqual(run.usage["model_calls"], 3)
        self.assertTrue(run.usage["provider_reported"])
        self.assertEqual((run.usage["provider_prompt_tokens"], run.usage["provider_completion_tokens"]), (300, 60))
        self.assertGreater(run.usage["prompt_tokens_est"], 0)


class Failures(unittest.TestCase):
    def test_a_source_that_moves_mid_run_stops_the_answer(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "wb.xlsx")
            Path(path).write_bytes(b"one")
            s = session_for(build(), workbook_path=path)
            def moving(user, results):
                if evidence_of(results) is None and not any(n == "get_account_evidence" for n, _ in results):
                    return tool("get_account_evidence", account="应收账款")
                if not any(n == "inspect_sources" for n, _ in results):
                    Path(path).write_bytes(b"two")                       # the workbook is replaced under the session
                    return tool("get_account_evidence", account="应收账款")
                if "CONTROLLER" in user:
                    return {"ask_user": "工作簿在运行中被替换，请确认以哪个版本为准。"}
                view = evidence_of(results) or {"totals": [{"period": P2, "basis": "authority_adjusted", "ref": "F1"}]}
                return {"answer": {"claims": [{"account": "应收账款", "kind": "amount", "periods": [P2], "cites": [total_ref(view, P2)],
                                               "text": "2025年12月31日应收账款余额130.0万元"}], "unresolved": []}}
            def script(user, results):
                names = [n for n, _ in results]
                if not names:
                    return tool("get_account_evidence", account="应收账款")
                if names == ["get_account_evidence"]:
                    Path(path).write_bytes(b"two")
                    return tool("get_account_evidence", account="应收账款")
                if "CONTROLLER" in user:
                    return {"ask_user": "工作簿在运行中被替换，请确认以哪个版本为准。"}
                view = evidence_of(results)
                return {"answer": {"claims": [{"account": "应收账款", "kind": "amount", "periods": [P2], "cites": [total_ref(view, P2)],
                                               "text": "2025年12月31日应收账款余额130.0万元"}], "unresolved": []}}
            run = C.AgentController(s, Policy(script)).run("余额？")
            self.assertEqual(run.stop_reason, C.STOP_NEEDS_USER, run.steps)
            self.assertEqual(run.steps[1]["error"], "STALE_SOURCE")
            self.assertTrue(any(i["id"] == "stale" for i in run.issues))


class Edit(unittest.TestCase):
    def policy(self, user, results):
        view = evidence_of(results)
        if view is None:
            return tool("get_account_evidence", account="应收账款")
        validated = [r for n, r in results if n == "validate_candidate"]
        ref = total_ref(view, P2)
        if not validated:
            bad = {"account": "应收账款", "kind": "amount", "periods": [P2], "cites": [ref], "text": "2025年12月31日应收账款余额13.0万元"}
            return tool("validate_candidate", account="应收账款", claims=[bad])
        last = validated[-1]
        if not last["ok"]:
            assert last["error"]["code"] == "VALIDATION_FAILED"
            good = {"account": "应收账款", "kind": "amount", "periods": [P2], "cites": [ref], "text": "2025年12月31日应收账款余额130.0万元"}
            return tool("validate_candidate", account="应收账款", claims=[good])
        return {"done": {"account": "应收账款", "export": False, "summary": "已改写"}}

    def test_a_rejected_draft_is_revised_until_it_validates_then_finished(self):
        s = clean_session()
        run = C.AgentController(s, Policy(self.policy)).run("把应收账款评论改为只讲期末余额", mode="edit", account="应收账款")
        self.assertEqual(run.stop_reason, C.STOP_DONE, run.steps)
        self.assertEqual(run.tool_path, ["get_account_evidence", "validate_candidate", "validate_candidate"])
        self.assertEqual([s_["ok"] for s_ in run.steps if s_.get("tool") == "validate_candidate"], [False, True])
        self.assertTrue(s.candidates["应收账款"].validated)
        self.assertIn("130.0万元", run.answer["text"])

    def test_done_without_a_validated_candidate_is_refused(self):
        run = C.AgentController(clean_session(), Policy(lambda u, r: {"done": {"account": "应收账款"}}),
                                budget=C.Budget(max_rejections=2)).run("改", mode="edit", account="应收账款")
        self.assertEqual(run.stop_reason, C.STOP_REJECTED)

    def test_done_with_export_requires_the_export(self):
        s = clean_session()
        def policy(user, results):
            out = self.policy(user, results)
            if "done" in out:
                out["done"]["export"] = True
            return out
        run = C.AgentController(s, Policy(policy), budget=C.Budget(max_rejections=1)).run("改", mode="edit", account="应收账款", export=True)
        self.assertEqual(run.stop_reason, C.STOP_REJECTED)
        self.assertIn("export", run.steps[-1]["feedback"])


class ExportProof(unittest.TestCase):
    """'done' with export=True needs a SUCCESSFUL export of THIS draft against THESE sources."""

    def validated_session(self, text="2025年12月31日应收账款余额130.0万元"):
        from fdd_utils.agent import tools as tm
        s = clean_session()
        c = s.contracts["应收账款"]
        self.assertTrue(tm.call_tool(s, "validate_candidate", {"account": "应收账款", "claims": [
            {"kind": "amount", "periods": [P2], "cites": [total_ref(C_view(c), P2)], "text": text}]}).ok)
        return s

    def gate(self, s):
        return C.AgentController(s, Policy(lambda u, r: {}))._gate_done({"account": "应收账款"}, "应收账款", True)

    def record(self, s, **kw):
        base = {"path": "x.pptx", "source_version": s.source_version, "touched_slides": {"应收账款": [1]}, "ok": True,
                "candidates": {"应收账款": s.candidates["应收账款"].claims_digest}, "delivery": {"status": "ready"}}
        base.update(kw)
        s.exports.append(base)

    def test_a_blocked_or_failed_export_record_does_not_count(self):
        s = self.validated_session()
        self.record(s, ok=False, delivery={"status": "blocked"})
        accepted, feedback, _ = self.gate(s)
        self.assertFalse(accepted)
        self.assertIn("blocked", feedback)

    def test_an_export_of_an_earlier_draft_does_not_prove_a_later_one(self):
        from fdd_utils.agent import tools as tm
        s = self.validated_session()
        self.record(s)                                              # draft A, exported fine
        self.assertTrue(self.gate(s)[0])
        c = s.contracts["应收账款"]
        self.assertTrue(tm.call_tool(s, "validate_candidate", {"account": "应收账款", "claims": [
            {"kind": "amount", "periods": [P1], "cites": [total_ref(C_view(c), P1)], "text": "2024年12月31日应收账款余额100.0万元"}]}).ok)   # draft B
        self.assertFalse(self.gate(s)[0], "A's export must not prove B")
        self.record(s)                                              # B exported
        self.assertTrue(self.gate(s)[0])

    def test_an_export_made_against_older_sources_does_not_count(self):
        s = self.validated_session()
        self.record(s, source_version="0000")
        self.assertFalse(self.gate(s)[0])

    def test_a_record_without_the_new_fields_is_not_accepted(self):
        s = self.validated_session()
        s.exports.append({"source_version": s.source_version, "touched_slides": {"应收账款": [1]}})
        self.assertFalse(self.gate(s)[0])


def C_view(contract):
    from fdd_utils.ai import claims as cl
    return cl.evidence_view(contract)


if __name__ == "__main__":
    unittest.main(verbosity=2)

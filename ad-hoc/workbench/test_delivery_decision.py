"""Delivery decision: when a deck is ready, needs review, or is blocked (Phase 1).

Zero tokens. Synthetic run results only:

    python ad-hoc/workbench/test_delivery_decision.py

The cases that matter most are the ones where the pipeline's own bookkeeping
looks healthy: retries exhausted, an arbiter that picked the best attempt, a run
whose calls all succeeded. None of those makes a deck with an unresolved invented
figure deliverable.
"""
from __future__ import annotations

import json
import os
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

from fdd_utils.delivery import (  # noqa: E402
    STATUS_BLOCKED, STATUS_READY, STATUS_REVIEW, DeliveryDecision, Finding, apply_to_presentation, banner,
    combine_decisions, decide_delivery, labelled_path, read_sidecar, unassessed_decision, write_sidecar,
)
from fdd_utils.workbook.selection import Decision, SelectionResult  # noqa: E402


def decide(**kw):
    """decide_delivery with the inputs a real export always has, so a test names only what it varies."""
    if "scope" not in kw:
        kw.setdefault("bs_is_results", {"statement_source": "financials"})
    return decide_delivery(**kw)



def review(clause, supported=True, category="", reason=""):
    return {"clause": clause, "supported": supported, "category": category, "reason": reason, "code": None}


def account(text="x", reviews=None, **extra):
    out = {"final": text, "subagent_4": text, "agent_4_validation": {"clause_reviews": reviews if reviews is not None else [review(text)]}}
    out.update(extra)
    return out


def results(**accounts):
    out = dict(accounts)
    out["__run_health__"] = {"zero_successful_calls": False, "accounts_on_fallback_bullet": [],
                             "accounts_on_passthrough": [], "accounts_on_error_text": [], "accounts_without_prompt": []}
    out["__run__"] = {"accounts": {k: {"phase": "ACCEPTED"} for k in accounts}}
    return out


class AccountLevel(unittest.TestCase):
    def test_clean_run_is_ready(self):
        d = decide(ai_results=results(a=account()))
        self.assertEqual(d.status, STATUS_READY)
        self.assertTrue(d.deliverable)

    def test_residual_hallucination_blocks_even_after_retries_and_arbiter(self):
        bad = account("收入增长50%", [review("收入增长50%", False, "hallucination", "50% not in the data")],
                      feedback_retries=2,
                      feedback_arbiter={"chosen": "retry_1", "scores": {"original": 3, "retry_1": 1, "retry_2": 2}})
        d = decide(ai_results=results(a=bad))
        self.assertEqual(d.status, STATUS_BLOCKED)
        finding = d.with_severity("blocked")[0]
        self.assertEqual((finding.code, finding.account), ("ACCOUNT_RESIDUAL_HALLUCINATION", "a"))
        self.assertIn("arbiter kept 'retry_1'", finding.message)

    def test_arbiter_that_restored_a_clean_attempt_is_not_blocked(self):
        clean = account("ok", [review("ok")], feedback_retries=2,
                        feedback_arbiter={"chosen": "original", "scores": {"original": 0, "retry_1": 1, "retry_2": 2}})
        self.assertEqual(decide(ai_results=results(a=clean)).status, STATUS_READY)

    def test_reasoning_flag_alone_is_disclosed_not_gating(self):
        flagged = account("预计系结转所致", [review("预计系结转所致", False, "reasoning", "inference")])
        d = decide(ai_results=results(a=flagged))
        self.assertEqual(d.status, STATUS_READY)
        self.assertEqual(d.with_severity("info")[0].code, "ACCOUNT_REASONING_FLAGS")

    def test_ungrounded_or_empty_account_is_blocked(self):
        d = decide(ai_results=results(a=account("text", reviews=[])))
        self.assertEqual(d.status, STATUS_BLOCKED)
        self.assertEqual(d.findings[0].code, "ACCOUNT_UNGROUNDED")

    def test_error_placeholder_blocks_and_fallback_bullet_asks_for_review(self):
        r = results(a=account(), b=account())
        r["__run_health__"]["accounts_on_error_text"] = ["a"]
        r["__run_health__"]["accounts_on_fallback_bullet"] = ["b"]
        d = decide(ai_results=r)
        self.assertEqual(d.status, STATUS_BLOCKED)
        self.assertEqual(d.accounts, {"a": STATUS_BLOCKED, "b": STATUS_REVIEW})

    def test_dead_run_is_blocked(self):
        r = results(a=account())
        r["__run_health__"]["zero_successful_calls"] = True
        self.assertEqual(decide(ai_results=r).status, STATUS_BLOCKED)

    def test_tool_failure_is_a_finding_not_a_pass(self):
        # an extra finding from a failed tool (agent path) is honoured as given
        d = decide(ai_results=results(a=account()),
                            extra_findings=[Finding("TOOL_FAILED", "blocked", "validate_candidate raised", "a")])
        self.assertEqual(d.status, STATUS_BLOCKED)

    def test_no_results_at_all_is_blocked_not_clean(self):
        self.assertEqual(decide().status, STATUS_BLOCKED)


class DeckLevel(unittest.TestCase):
    def selection(self, *decisions, available=True, unmatched=()):
        return SelectionResult(decisions=list(decisions), reconciliation_available=available,
                               unmatched_financials=list(unmatched))

    def test_bs_diff_exclusion_is_surfaced_for_review(self):
        sel = self.selection(Decision("应收账款", False, "excluded", "bs_diff", "does not tie"))
        d = decide(ai_results=results(a=account()), selection=sel)
        self.assertEqual(d.status, STATUS_REVIEW)
        self.assertEqual(d.accounts["应收账款"], STATUS_REVIEW)

    def test_account_with_no_reconciliation_row_is_surfaced(self):
        sel = self.selection(Decision("关联方", False, "excluded", "no_reconciliation_row", "no row"))
        self.assertEqual(decide(ai_results=results(a=account()), selection=sel).status, STATUS_REVIEW)

    def test_dormant_tab_is_silent(self):
        sel = self.selection(Decision("空", False, "excluded", "no_activity", "all zero"))
        self.assertEqual(decide(ai_results=results(a=account()), selection=sel).status, STATUS_READY)

    def test_unchecked_history_is_disclosed_not_gating_but_interim_asks_for_review(self):
        unchecked = self.selection(Decision("a", True, "unchecked_history", "financials_line_nil_recently", "nil"))
        d = decide(ai_results=results(a=account()), selection=unchecked)
        self.assertEqual(d.status, STATUS_READY)
        self.assertEqual(d.with_severity("info")[0].code, "SELECTION_UNCHECKED_HISTORY")
        interim = self.selection(Decision("a", True, "prior_period", "interim_latest_period", "partial"))
        self.assertEqual(decide(ai_results=results(a=account()), selection=interim).status, STATUS_REVIEW)

    def test_financials_line_without_a_tab_asks_for_review(self):
        sel = self.selection(unmatched=[{"statement": "BS", "financials_account": "Retained earnings"}])
        self.assertEqual(decide(ai_results=results(a=account()), selection=sel).status, STATUS_REVIEW)

    def test_synthesized_statements_are_disclosed_as_review(self):
        d = decide(ai_results=results(a=account()), bs_is_results={"statement_source": "synthesized"})
        self.assertEqual(d.status, STATUS_REVIEW)
        self.assertEqual(d.findings[0].code, "STATEMENT_SYNTHESIZED")
        real = decide(ai_results=results(a=account()), bs_is_results={"statement_source": "financials"})
        self.assertEqual(real.status, STATUS_READY)

    def test_wrong_entity_or_financials_blocks(self):
        d = decide(ai_results=results(a=account()), identity_ok=False, identity_note="workbook sha differs")
        self.assertEqual(d.status, STATUS_BLOCKED)
        self.assertIn("workbook sha differs", next(f for f in d.findings if f.code == "IDENTITY_MISMATCH").message)

    def test_worst_part_decides_a_combined_deck(self):
        ready = decide(ai_results=results(a=account()))
        blocked = decide(ai_results=results(a=account("t", [review("t", False, "hallucination", "no")])))
        combined = combine_decisions([("entity-1", ready), ("entity-2", blocked.to_dict()), ("entity-3", None)])
        self.assertEqual(combined.status, STATUS_BLOCKED)
        self.assertIn("entity-2/a", combined.accounts)
        self.assertIn("entity-3", combined.accounts)      # a part with no decision is unassessed, not clean
        self.assertEqual(combine_decisions([("only", ready)]).status, STATUS_READY)


class Silence(unittest.TestCase):
    """Nothing to iterate is not nothing wrong."""

    SENTINEL = {"zero_successful_calls": False, "accounts_on_fallback_bullet": [], "accounts_on_passthrough": [],
                "accounts_on_error_text": [], "accounts_without_prompt": []}

    def test_empty_results_are_blocked(self):
        for empty in ({}, {"__run_health__": self.SENTINEL}):
            d = decide(ai_results=empty)
            self.assertEqual(d.status, STATUS_BLOCKED, empty)
            self.assertFalse(d.deliverable)
            self.assertEqual(d.with_severity("blocked")[0].code, "NO_CONTENT")

    def test_an_admitted_account_with_no_commentary_is_reported(self):
        sel = SelectionResult(decisions=[Decision("a", True, "reconciled", "tied_out", ""), Decision("b", True, "reconciled", "tied_out", "")])
        d = decide(ai_results=results(a=account()), selection=sel)
        self.assertEqual((d.status, d.accounts["b"]), (STATUS_REVIEW, STATUS_REVIEW))
        self.assertEqual(d.findings[0].code, "ACCOUNT_MISSING")

    def test_every_admitted_account_missing_is_no_content(self):
        sel = SelectionResult(decisions=[Decision("z", True, "reconciled", "tied_out", "")])
        self.assertEqual(decide(ai_results=results(a=account()), selection=sel).status, STATUS_BLOCKED)

    def test_text_that_is_empty_is_blocked(self):
        d = decide(ai_results=results(a=account("", [review("x")])))
        self.assertEqual((d.status, d.findings[0].code), (STATUS_BLOCKED, "ACCOUNT_EMPTY_TEXT"))

    def test_a_missing_statement_source_asks_for_review_and_an_account_preview_does_not(self):
        d = decide_delivery(ai_results=results(a=account()))
        self.assertEqual(d.status, STATUS_REVIEW)
        self.assertIn("INPUT_MISSING", [f.code for f in d.findings])
        self.assertEqual(decide_delivery(ai_results=results(a=account()), scope="account").status, STATUS_READY)

    def test_many_missing_accounts_are_one_finding_not_a_wall(self):
        d = decide(ai_results=results(a=account()), expected_accounts=["a"] + [f"m{i}" for i in range(12)])
        self.assertEqual([f.code for f in d.findings if f.code == "ACCOUNT_MISSING"], ["ACCOUNT_MISSING"])
        self.assertIn("12 of the 13", next(f for f in d.findings if f.code == "ACCOUNT_MISSING").message)

    def test_the_expected_accounts_can_be_given_directly(self):
        d = decide(ai_results=results(a=account()), expected_accounts=["a", "b", "c"])
        self.assertEqual(sorted(d.accounts), ["b", "c"])


class Label(unittest.TestCase):
    def test_file_name_suffix(self):
        ready = DeliveryDecision()
        review_ = decide(ai_results=results(a=account()), bs_is_results={"statement_source": "synthesized"})
        blocked = decide()
        self.assertEqual(labelled_path("x/Deck.pptx", ready), "x/Deck.pptx")
        self.assertEqual(labelled_path("x/Deck.pptx", review_), "x/Deck_REVIEW.pptx")
        self.assertEqual(labelled_path("x/Deck.pptx", blocked), "x/Deck_DRAFT-BLOCKED.pptx")
        self.assertEqual(labelled_path("x/Deck_REVIEW.pptx", review_), "x/Deck_REVIEW.pptx")   # idempotent
        self.assertEqual(labelled_path("x/Deck.pptx", None), "x/Deck_REVIEW.pptx")             # no decision: fail closed

    def test_unassessed_is_review_and_says_so(self):
        d = unassessed_decision()
        self.assertEqual((d.status, d.assessed), (STATUS_REVIEW, False))
        self.assertIn("REVIEW REQUIRED", banner(d))

    def test_sidecar_and_document_properties_round_trip(self):
        from pptx import Presentation
        d = decide(ai_results=results(a=account("t", [review("t", False, "hallucination", "no")])))
        with tempfile.TemporaryDirectory() as tmp:
            deck = os.path.join(tmp, "Deck.pptx")
            prs = Presentation()
            apply_to_presentation(prs, d)
            prs.save(deck)
            write_sidecar(deck, d)
            reread = Presentation(deck).core_properties
            self.assertEqual(reread.category, "DRAFT - NOT DELIVERABLE")
            self.assertIn("ACCOUNT_RESIDUAL_HALLUCINATION".lower(), json.dumps(read_sidecar(deck)).lower())
            self.assertEqual(read_sidecar(deck)["status"], STATUS_BLOCKED)
            self.assertEqual(DeliveryDecision.from_dict(read_sidecar(deck)).status, STATUS_BLOCKED)


if __name__ == "__main__":
    unittest.main(verbosity=2)

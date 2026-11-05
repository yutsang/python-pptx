"""Artifact manifest, returned decks and source updates (Phase 6).

Zero tokens. Decks are built here with python-pptx (a commentary box per column, one
paragraph per account: a bold '■ label - ' run and a body run), the manifest is built by the
production code, and each case damages the deck the way a returned one can be damaged:

    python ad-hoc/workbench/test_deck_lineage.py

Nothing here claims PowerPoint preserves the custom XML part or the alt-text tags: python-pptx
does, a real PowerPoint save is NOT tested. The cases for a lost tag and a stripped manifest are
what protect against that being false.
"""
from __future__ import annotations

import copy
import json
import os
import sys
import tempfile
import unittest
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parents[1]))

from pptx import Presentation  # noqa: E402
from pptx.util import Emu  # noqa: E402

from fdd_utils.agent import updates as U  # noqa: E402
from fdd_utils.ai import claims as cl  # noqa: E402
from fdd_utils.pptx import lineage as L  # noqa: E402
from test_agent_tools import make_session  # noqa: E402
from test_claims import P1, P2, ROWS, build, ref  # noqa: E402

PREPAY_ROWS = [("供应商", "detail", {P1: 150000, P2: 200000}), ("Total", "total", {P1: 150000, P2: 200000})]
TEXT = "2025年12月31日应收账款余额130.0万元"
TEXT_CHANGED = "2025年12月31日应收账款余额150.0万元"
ROWS_CHANGED = [("甲公司", "detail", {P1: 600000, P2: 700000}), ("乙公司", "detail", {P1: 400000, P2: 500000}),
                ("丙公司", "detail", {P1: 0, P2: 300000}), ("Total", "total", {P1: 1000000, P2: 1500000})]


def sess(**kwargs):
    """A session with both accounts the test decks talk about."""
    session = make_session(**kwargs)
    session.contracts["预付款项"] = build(account="预付款项", sheet="Prepay", rows=PREPAY_ROWS)
    for contract in session.contracts.values():
        contract.binding = session.binding.to_dict()
    return session


def make_deck(path, slides):
    """slides: [[(shape name, [(label, body), ...]), ...], ...]"""
    prs = Presentation()
    for boxes in slides:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        for i, (name, paragraphs) in enumerate(boxes):
            shape = slide.shapes.add_textbox(Emu(400000 + i * 4600000), Emu(1200000), Emu(4400000), Emu(4500000))
            shape.name = name
            tf = shape.text_frame
            for j, (label, body) in enumerate(paragraphs):
                p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
                r1 = p.add_run()
                r1.text = f"■ {label} - "
                r1.font.bold = True
                r2 = p.add_run()
                r2.text = body
                r2.font.size = Emu(152400)
    prs.save(path)


def stamped(tmp, session, slides, claims=None, contracts=None, name="deck.pptx"):
    path = os.path.join(tmp, name)
    make_deck(path, slides)
    items = [{"mapping_key": a, "display_name": a, "display_name_zh": a, "commentary": c}
             for a, c in {"应收账款": TEXT, "预付款项": "2025年12月31日预付款项余额20.0万元"}.items()]
    contracts = contracts or session.contracts
    public, internal = L.build_manifest(path, items, is_chinese=True, source_version=session.source_version,
                                        contracts=contracts, claims_by_account=claims or {})
    L.stamp_deck(path, public, internal)
    return path, public, internal


def one_slide():
    return [[("textMainBullets", [("应收账款", TEXT), ("预付款项", "2025年12月31日预付款项余额20.0万元")])]]


def claims_for(contract):
    return {"应收账款": [{"id": "c1", "text": TEXT, "cites": [ref(contract, P2)]}]}


def two_sessions():
    old = sess()
    new = sess(rows=ROWS_CHANGED)
    new.binding.workbook_sha256 = "bbbb"            # a changed workbook has a different hash, hence a different source version
    for contract in new.contracts.values():
        contract.binding = new.binding.to_dict()
    return old, new


class Manifest(unittest.TestCase):
    def test_elements_tags_and_embedded_part_round_trip(self):
        old = sess()
        with tempfile.TemporaryDirectory() as tmp:
            path, public, internal = stamped(tmp, old, one_slide(), claims_for(old.contracts["应收账款"]))
            self.assertEqual([e["account"] for e in public["elements"]], ["应收账款", "预付款项"])
            embedded = L.read_manifest(path)
            self.assertEqual(embedded["deck_id"], public["deck_id"])
            tagged = [L.shape_tags(s) for s in Presentation(path).slides[0].shapes]
            self.assertEqual(tagged[0], (public["deck_id"], ["E1", "E2"]))
            self.assertTrue(os.path.exists(path + L.INTERNAL_SUFFIX))

    def test_the_embedded_part_carries_no_figure(self):
        old = sess()
        with tempfile.TemporaryDirectory() as tmp:
            path, public, _ = stamped(tmp, old, one_slide(), claims_for(old.contracts["应收账款"]))
            blob = json.dumps(L.read_manifest(path), ensure_ascii=False)
            for forbidden in ("1300000", "1,300,000", "130.0", "700000", "2025年12月31日应收账款余额"):
                self.assertNotIn(forbidden, blob)
            self.assertTrue(public["elements"][0]["cite_hashes"])         # but the hashes are there

    def test_cites_are_recorded_for_claim_written_paragraphs_only(self):
        old = sess()
        with tempfile.TemporaryDirectory() as tmp:
            _, public, _ = stamped(tmp, old, one_slide(), claims_for(old.contracts["应收账款"]))
            self.assertTrue(public["elements"][0]["cites"])
            self.assertEqual(public["elements"][1]["cites"], [])

    def test_strip_manifest_removes_the_part_and_the_tags(self):
        old = sess()
        with tempfile.TemporaryDirectory() as tmp:
            path, _, _ = stamped(tmp, old, one_slide())
            out = L.strip_manifest(path, os.path.join(tmp, "outside.pptx"))
            self.assertIsNone(L.read_manifest(out))
            self.assertEqual(L.shape_tags(Presentation(out).slides[0].shapes[0]), (None, []))
            self.assertIsNotNone(L.read_manifest(path), "the working copy keeps its manifest")


class ReturnedDeck(unittest.TestCase):
    def analyze(self, tmp, session_old, session_new, edit=None, mutate=None, keep_internal=True, claims=True):
        path, public, internal = stamped(tmp, session_old, one_slide(), claims_for(session_old.contracts["应收账款"]) if claims else None)
        if edit:
            prs = Presentation(path)
            for p in prs.slides[0].shapes[0].text_frame.paragraphs:
                if edit[0] in p.text:
                    p.runs[-1].text = edit[1]
            prs.save(path)
        if mutate:
            mutate(path)
        if not keep_internal:
            os.remove(path + L.INTERNAL_SUFFIX)
        return path, U.analyze_returned(path, session_new)

    def actions(self, report):
        return {a["element"]: a["action"] for a in report["actions"]}

    def test_nothing_changed_means_keep(self):
        old = sess()
        with tempfile.TemporaryDirectory() as tmp:
            _, report = self.analyze(tmp, old, sess())
            self.assertEqual(self.actions(report), {"E1": "keep", "E2": "keep"})
            self.assertEqual(report["mode"], "full")

    def test_source_change_only_updates_the_paragraph_that_cites_the_changed_fact(self):
        old, new = two_sessions()
        with tempfile.TemporaryDirectory() as tmp:
            _, report = self.analyze(tmp, old, new)
            self.assertEqual(self.actions(report)["E1"], "auto_update")
            self.assertEqual(self.actions(report)["E2"], "keep")
            self.assertIn("应收账款", report["source_diff"]["changed_accounts"])
            self.assertTrue(report["source_moved"])

    def test_without_the_internal_baseline_the_change_is_found_but_not_described(self):
        old, new = two_sessions()
        with tempfile.TemporaryDirectory() as tmp:
            _, report = self.analyze(tmp, old, new, keep_internal=False)
            self.assertEqual((report["mode"], report["history"]), ("fingerprint_only", "partial"))
            self.assertEqual(self.actions(report)["E1"], "auto_update")
            self.assertEqual(report["source_diff"]["mode"], "hashes")
            self.assertIn("not the", report["source_diff"]["note"] + "not the")     # states what it cannot say

    def test_user_edit_with_unchanged_sources_is_kept_and_its_numbers_rechecked(self):
        old = sess()
        with tempfile.TemporaryDirectory() as tmp:
            _, report = self.analyze(tmp, old, sess(), edit=("应收账款余额", "2025年12月31日应收账款余额130.0万元，其中甲公司70.0万元，另有一笔99.9万元待确认"))
            action = next(a for a in report["actions"] if a["element"] == "E1")
            self.assertEqual(action["action"], "reverify")
            self.assertIn("130.0万元", action["check"]["amounts_found"])
            self.assertIn("99.9万元", action["check"]["amounts_not_found"])
            self.assertIn("not claim-verified", action["check"]["state"])

    def test_user_edit_and_source_change_is_a_conflict_not_an_overwrite(self):
        old, new = two_sessions()
        with tempfile.TemporaryDirectory() as tmp:
            path, report = self.analyze(tmp, old, new, edit=("应收账款余额", "应收账款期末余额130.0万元，管理层口径"))
            action = next(a for a in report["actions"] if a["element"] == "E1")
            self.assertEqual((action["action"], action["base_known"]), ("conflict", True))
            out = os.path.join(tmp, "updated.pptx")
            result = U.apply_updates(path, new, report, {"E1": TEXT_CHANGED}, out)
            self.assertEqual(result["applied"], [])
            self.assertEqual(result["skipped"][0]["action"], "conflict")
            texts = [p.text for p in Presentation(out).slides[0].shapes[0].text_frame.paragraphs]
            self.assertIn("管理层口径", texts[0])                    # the user's words are still there

    def test_a_concrete_merge_proposal_keeps_what_each_side_changed(self):
        base = "余额为130.0万元。主要为甲公司。账龄较短。"
        user = "余额为130.0万元。主要为甲公司，管理层确认可收回。账龄较短。"
        new = "余额为150.0万元。主要为甲公司。账龄较短。"
        merged = L.three_way(base, user, new)
        self.assertEqual(merged["mode"], "three_way")
        self.assertIn("150.0万元", merged["merged"])               # the source change taken
        self.assertIn("管理层确认可收回", merged["merged"])         # the user's change kept
        self.assertEqual(merged["conflicts"], [])
        both = L.three_way(base, "余额为140.0万元。主要为甲公司。账龄较短。", new)
        self.assertEqual(len(both["conflicts"]), 1)
        self.assertEqual(both["conflicts"][0]["user"], "余额为140.0万元。")
        self.assertEqual(L.three_way(None, user, new)["mode"], "no_base")

    def test_update_in_place_keeps_the_label_run_and_changes_only_what_changed(self):
        old, new = two_sessions()
        with tempfile.TemporaryDirectory() as tmp:
            path, report = self.analyze(tmp, old, new)
            out = os.path.join(tmp, "updated.pptx")
            result = U.apply_updates(path, new, report, {"E1": TEXT_CHANGED}, out)
            self.assertEqual([a["element"] for a in result["applied"]], ["E1"])
            self.assertEqual(result["touched_slides"], [1])
            paragraphs = Presentation(out).slides[0].shapes[0].text_frame.paragraphs
            self.assertEqual(paragraphs[0].text, f"■ 应收账款 - {TEXT_CHANGED}")
            self.assertTrue(paragraphs[0].runs[0].font.bold)              # the label run is untouched
            self.assertEqual(paragraphs[1].text, "■ 预付款项 - 2025年12月31日预付款项余额20.0万元")

    def test_a_copied_shape_is_reported_and_the_one_in_place_is_the_original(self):
        old = sess()
        def copy_onto_slide_two(path):
            prs = Presentation(path)
            slide2 = prs.slides.add_slide(prs.slide_layouts[6])
            slide2.shapes._spTree.append(copy.deepcopy(prs.slides[0].shapes[0]._element))
            prs.save(path)
        with tempfile.TemporaryDirectory() as tmp:
            _, report = self.analyze(tmp, old, sess(), mutate=copy_onto_slide_two)
            self.assertTrue(report["duplicate_tags"])
            self.assertEqual({r["id"]: r["slide_now"] for r in report["elements"]}, {"E1": 1, "E2": 1})
            self.assertEqual(self.actions(report), {"E1": "keep", "E2": "keep"})

    def test_a_copy_beside_the_original_leaves_no_clear_original(self):
        old = sess()
        def copy_in_place(path):
            prs = Presentation(path)
            prs.slides[0].shapes._spTree.append(copy.deepcopy(prs.slides[0].shapes[0]._element))
            prs.save(path)
        with tempfile.TemporaryDirectory() as tmp:
            _, report = self.analyze(tmp, old, sess(), mutate=copy_in_place)
            self.assertEqual({r["status"] for r in report["elements"]}, {"ambiguous"})
            self.assertEqual({a["action"] for a in report["actions"]}, {"inspect_and_relink"})

    def test_reordered_slides_are_moves_not_edits(self):
        old = sess()
        def swap(path):
            prs = Presentation(path)
            ids = prs.slides._sldIdLst
            ids.insert(0, ids[-1])
            prs.save(path)
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "deck.pptx")
            make_deck(path, [[("textMainBullets", [("应收账款", TEXT)])], [("textMainBullets", [("预付款项", "2025年12月31日预付款项余额20.0万元")])]])
            items = [{"mapping_key": a, "display_name": a, "display_name_zh": a, "commentary": c}
                     for a, c in {"应收账款": TEXT, "预付款项": "2025年12月31日预付款项余额20.0万元"}.items()]
            public, internal = L.build_manifest(path, items, is_chinese=True, source_version=old.source_version, contracts=old.contracts)
            L.stamp_deck(path, public, internal)
            swap(path)
            report = U.analyze_returned(path, sess())
            self.assertEqual({r["id"]: (r["status"], r["slide_now"]) for r in report["elements"]}, {"E1": ("moved", 2), "E2": ("moved", 1)})
            self.assertTrue(report["reordered_or_moved"])
            self.assertEqual({a["action"] for a in report["actions"]}, {"keep"})

    def test_a_lost_tag_is_relinked_by_text_and_a_reworded_paragraph_by_similarity(self):
        old = sess()
        def untag(path):
            for shape in Presentation(path).slides[0].shapes:
                pass
            prs = Presentation(path)
            for shape in prs.slides[0].shapes:
                cnv = shape._element.xpath(".//p:cNvPr")[0]
                cnv.attrib.pop("descr", None)
            prs.save(path)
        with tempfile.TemporaryDirectory() as tmp:
            _, report = self.analyze(tmp, old, sess(), mutate=untag)
            self.assertEqual(self.actions(report), {"E1": "keep", "E2": "keep"})
            self.assertEqual({v["method"] for v in report["located"].values()}, {"text"})
        with tempfile.TemporaryDirectory() as tmp:
            _, report = self.analyze(tmp, old, sess(), mutate=untag, edit=("应收账款余额", "2025年12月31日应收账款余额为130.0万元左右"))
            self.assertEqual(report["located"]["E1"]["method"], "similarity")
            self.assertEqual(self.actions(report)["E1"], "reverify")

    def test_a_deleted_paragraph_is_reported_not_recreated(self):
        old = sess()
        def delete(path):
            prs = Presentation(path)
            para = prs.slides[0].shapes[0].text_frame.paragraphs[1]
            para._p.getparent().remove(para._p)
            prs.save(path)
        with tempfile.TemporaryDirectory() as tmp:
            _, report = self.analyze(tmp, old, sess(), mutate=delete)
            self.assertEqual(self.actions(report)["E2"], "inspect_and_relink")
            self.assertEqual(report["located"]["E2"]["status"], "unlocated")

    def test_a_stripped_manifest_claims_no_history(self):
        old = sess()
        with tempfile.TemporaryDirectory() as tmp:
            path, _, _ = stamped(tmp, old, one_slide())
            outside = L.strip_manifest(path, os.path.join(tmp, "outside.pptx"))
            report = U.analyze_returned(outside, sess())
            self.assertEqual((report["mode"], report["history"]), ("none", "unknown"))
            self.assertIn("says nothing about what changed", report["note"])
            self.assertEqual({a["action"] for a in report["actions"]}, {"inspect_and_relink"})
            self.assertTrue(all("amounts_found" in p["check"] for p in report["paragraphs"]))
            self.assertNotIn("source_diff", report)

    def test_several_accounts_in_one_box_and_one_account_across_pages(self):
        old = sess()
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "deck.pptx")
            make_deck(path, [[("textMainBullets_L", [("应收账款", TEXT[:12]), ("预付款项", "2025年12月31日预付款项余额20.0万元")])],
                             [("textMainBullets_L", [("应收账款 (续)", TEXT[12:])])]])
            items = [{"mapping_key": "应收账款", "display_name": "应收账款", "display_name_zh": "应收账款", "commentary": TEXT},
                     {"mapping_key": "预付款项", "display_name": "预付款项", "display_name_zh": "预付款项", "commentary": "2025年12月31日预付款项余额20.0万元"}]
            public, _ = L.build_manifest(path, items, is_chinese=True, source_version=old.source_version, contracts=old.contracts)
            parts = [(e["account"], e["part"], e["parts"], e["continuation"]) for e in public["elements"]]
            self.assertEqual(parts, [("应收账款", 1, 2, False), ("预付款项", 1, 1, False), ("应收账款", 2, 2, True)])

    def test_an_account_spanning_pages_is_not_edited_in_place(self):
        old, new = two_sessions()
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "deck.pptx")
            make_deck(path, [[("textMainBullets_L", [("应收账款", TEXT[:12])])], [("textMainBullets_L", [("应收账款 (续)", TEXT[12:])])]])
            items = [{"mapping_key": "应收账款", "display_name": "应收账款", "display_name_zh": "应收账款", "commentary": TEXT}]
            public, internal = L.build_manifest(path, items, is_chinese=True, source_version=old.source_version, contracts=old.contracts,
                                                claims_by_account=claims_for(old.contracts["应收账款"]))
            for e in public["elements"]:      # the claim spans both pieces; both cite it
                e["cites"], e["cite_hashes"] = public["elements"][0]["cites"] or [old.contracts["应收账款"].fact(ref(old.contracts["应收账款"], P2))["id"]], \
                    {old.contracts["应收账款"].fact(ref(old.contracts["应收账款"], P2))["id"]: L.value_hash(1300000.0)}
            L.stamp_deck(path, public, internal)
            report = U.analyze_returned(path, new)
            result = U.apply_updates(path, new, report, {e["id"]: TEXT_CHANGED for e in public["elements"]}, os.path.join(tmp, "out.pptx"))
            self.assertEqual(result["applied"], [])
            self.assertTrue(all("re-packed" in s["reason"] for s in result["skipped"]))


class Tables(unittest.TestCase):
    def deck_with_table(self, tmp, session):
        path = os.path.join(tmp, "deck.pptx")
        make_deck(path, one_slide())
        prs = Presentation(path)
        shape = prs.slides[0].shapes.add_table(3, 2, Emu(400000), Emu(300000), Emu(3000000), Emu(900000))
        for r, row in enumerate([("科目", "金额"), ("应收账款", "1,300"), ("预付款项", "200")]):
            for c, text in enumerate(row):
                shape.table.cell(r, c).text = text
        prs.save(path)
        items = [{"mapping_key": a, "display_name": a, "display_name_zh": a, "commentary": c}
                 for a, c in {"应收账款": TEXT, "预付款项": "2025年12月31日预付款项余额20.0万元"}.items()]
        public, internal = L.build_manifest(path, items, is_chinese=True, source_version=session.source_version, contracts=session.contracts)
        L.stamp_deck(path, public, internal)
        return path

    def test_a_hand_edited_statement_table_is_seen_and_left_alone(self):
        old = sess()
        with tempfile.TemporaryDirectory() as tmp:
            path = self.deck_with_table(tmp, old)
            self.assertEqual(U.analyze_returned(path, sess())["tables"], {"T1": {"status": "unchanged", "slide_now": 1, "moved": False}})
            prs = Presentation(path)
            table = next(s for s in prs.slides[0].shapes if getattr(s, "has_table", False)).table
            table.cell(1, 1).text = "1,500"
            prs.save(path)
            report = U.analyze_returned(path, sess())
            self.assertEqual(report["tables"]["T1"]["status"], "user_edited")
            self.assertIn({"element": "T1", "action": "table_edited"}, [{k: a[k] for k in ("element", "action")} for a in report["actions"]])

    def test_a_table_with_its_tag_and_title_lost_is_not_guessed_at(self):
        old = sess()
        with tempfile.TemporaryDirectory() as tmp:
            path = self.deck_with_table(tmp, old)
            prs = Presentation(path)
            for shape in prs.slides[0].shapes:
                if getattr(shape, "has_table", False):
                    shape._element.xpath(".//p:cNvPr")[0].attrib.pop("descr", None)
                    shape.table.cell(0, 0).text = "改过的标题"
            prs.save(path)
            self.assertEqual(U.analyze_returned(path, sess())["tables"]["T1"]["status"], "unlocated")


class Dependencies(unittest.TestCase):
    """A figure can stay the same while the reason it may be stated, or the note that explained it, does not."""

    CAUSE_TEXT = "2025年12月31日应收账款余额130.0万元，主要系甲公司因装修免租期延迟收取租金所致"

    def stamped_cause_deck(self, tmp, session):
        path = os.path.join(tmp, "deck.pptx")
        make_deck(path, [[("textMainBullets", [("应收账款", self.CAUSE_TEXT)])]])
        c = session.contracts["应收账款"]
        claims = {"应收账款": [{"id": "c1", "text": self.CAUSE_TEXT, "cites": [ref(c, P2)], "cause_cites": [c.notes[0]["ref"]]}]}
        items = [{"mapping_key": "应收账款", "display_name": "应收账款", "display_name_zh": "应收账款", "commentary": self.CAUSE_TEXT}]
        public, internal = L.build_manifest(path, items, is_chinese=True, source_version=session.source_version,
                                            contracts=session.contracts, claims_by_account=claims)
        L.stamp_deck(path, public, internal)
        return path, public

    def moved(self, session, tag="bbbb"):
        session.binding.workbook_sha256 = tag
        for contract in session.contracts.values():
            contract.binding = session.binding.to_dict()
        return session

    def test_the_manifest_records_what_the_paragraph_depends_on_beyond_its_values(self):
        old = sess()
        with tempfile.TemporaryDirectory() as tmp:
            _, public = self.stamped_cause_deck(tmp, old)
            deps = public["elements"][0]["deps"]
            self.assertTrue(deps["permissions"] and deps["checks"] and len(deps["notes"]) == 1)

    def test_same_figures_but_the_note_behind_the_cause_changed(self):
        old = sess()
        with tempfile.TemporaryDirectory() as tmp:
            path, _ = self.stamped_cause_deck(tmp, old)
            new = self.moved(sess())
            new.contracts["应收账款"].notes[0]["text"] = "应收甲公司的租金已于本期全部收回，装修免租期已结束"
            report = U.analyze_returned(path, new)
            action = next(a for a in report["actions"] if a["element"] == "E1")
            self.assertEqual(action["action"], "auto_update")
            self.assertIn("note(s) the cause rested on changed", action["reason"])
            self.assertEqual(report["source_diff"]["mode"], "values")
            self.assertEqual(report["source_diff"]["changed_accounts"], ["应收账款"])     # found even though no value moved

    def test_same_figures_but_what_the_period_permits_changed(self):
        old = sess()
        with tempfile.TemporaryDirectory() as tmp:
            path, _ = self.stamped_cause_deck(tmp, old)
            new = self.moved(sess(financials={P1: 1000000.0, P2: 1500000.0}))        # Financials now disagrees; the tab total is unchanged
            self.assertEqual(new.contracts["应收账款"].total(P2), 1300000.0)
            report = U.analyze_returned(path, new)
            action = next(a for a in report["actions"] if a["element"] == "E1")
            self.assertEqual(action["action"], "auto_update")
            self.assertIn("permit", action["reason"])
            diff = U.diff_sources(L.read_internal(path)["baseline"], new.contracts)
            fields = {(c["period"], c["field"]) for c in diff["accounts"]["应收账款"]["permission_changes"]}
            self.assertIn((P2, "describe_total"), fields)

    def test_the_public_manifest_alone_still_notices(self):
        old = sess()
        with tempfile.TemporaryDirectory() as tmp:
            path, _ = self.stamped_cause_deck(tmp, old)
            os.remove(path + L.INTERNAL_SUFFIX)
            new = self.moved(sess(financials={P1: 1000000.0, P2: 1500000.0}))
            report = U.analyze_returned(path, new)
            self.assertEqual(report["source_diff"]["mode"], "hashes")
            self.assertEqual(next(a for a in report["actions"] if a["element"] == "E1")["action"], "auto_update")

    def test_nothing_changed_nothing_flagged(self):
        old = sess()
        with tempfile.TemporaryDirectory() as tmp:
            path, _ = self.stamped_cause_deck(tmp, old)
            report = U.analyze_returned(path, sess())
            self.assertEqual(next(a for a in report["actions"] if a["element"] == "E1")["action"], "keep")


class SourceDiff(unittest.TestCase):
    def test_a_line_inserted_in_the_source_is_an_added_fact_and_leaves_an_unrelated_paragraph_alone(self):
        old = sess()
        rows = [("甲公司", "detail", {P1: 600000, P2: 700000}), ("乙公司", "detail", {P1: 400000, P2: 300000}),
                ("丙公司", "detail", {P1: 0, P2: 300000}), ("丁公司", "detail", {P1: 0, P2: 0}), ("Total", "total", {P1: 1000000, P2: 1300000})]
        new = sess(rows=rows)
        with tempfile.TemporaryDirectory() as tmp:
            _, _, internal = stamped(tmp, old, one_slide(), claims_for(old.contracts["应收账款"]))
            diff = U.diff_sources(internal["baseline"], new.contracts)
            added = diff["accounts"]["应收账款"]["added"]
            self.assertTrue(any("丁公司" in i for i in added), added)
            self.assertEqual(diff["accounts"]["应收账款"]["changed"], [])        # no existing figure moved
            self.assertEqual(diff["affected_fact_ids"], [])                      # so no cited fact is affected


    def test_a_baseline_says_from_what_to_what(self):
        old, new = two_sessions()
        with tempfile.TemporaryDirectory() as tmp:
            _, _, internal = stamped(tmp, old, one_slide())
            diff = U.diff_sources(internal["baseline"], new.contracts)
            changes = {c["id"]: (c["old"], c["new"]) for c in diff["accounts"]["应收账款"]["changed"]}
            self.assertEqual(changes["应收账款|2025-12-31|authority_adjusted|#total"], (1300000.0, 1500000.0))
            self.assertEqual(changes["应收账款|2025-12-31|raw_schedule|乙公司"], (300000.0, 500000.0))

    def test_without_any_baseline_the_question_is_not_answered(self):
        from fdd_utils.agent import tools as T
        s = sess()
        r = T.call_tool(s, "diff_sources")
        self.assertEqual((r.ok, r.error["code"]), (False, "NO_BASELINE"))
        self.assertIn("cannot", r.error["message"])

    def test_a_changed_check_and_a_new_period_are_changes_too(self):
        old = sess()
        later = build(rows=[(r[0], r[1], dict(r[2], **{"2026-01-31": r[2][P2]})) for r in ROWS], periods=(P1, P2, "2026-01-31"))
        with tempfile.TemporaryDirectory() as tmp:
            _, _, internal = stamped(tmp, old, one_slide())
            diff = U.diff_sources(internal["baseline"], {"应收账款": later})
            self.assertEqual(diff["accounts"]["应收账款"]["periods_added"], ["2026-01-31"])


if __name__ == "__main__":
    unittest.main(verbosity=2)

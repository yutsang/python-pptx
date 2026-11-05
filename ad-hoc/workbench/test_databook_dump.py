"""A workbook as text, and a rebuild from it, checked against the repo's own readers -- offline, no model.

    python ad-hoc/workbench/test_databook_dump.py

The workbook is written here with openpyxl: a visible sheet with indentation, bold, number formats, outline levels, a hidden row and
column and merged cells; a hidden sheet and a very hidden one. The check that matters is the last class: the style signals the
extractor reads (fdd_utils.workbook.schedules) come out the same from the rebuild as from the original.
"""
from __future__ import annotations

import io
import json
import sys
import tempfile
import unittest
import zipfile
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parents[1]))

import databook_dump as D  # noqa: E402


def make_book(path: Path) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font
    wb = Workbook()
    ws = wb.active
    ws.title = "应收账款"
    ws["A1"], ws["A2"] = "应收账款", "人民币千元"
    ws["C3"], ws["D3"] = datetime(2024, 12, 31), datetime(2025, 12, 31)
    ws["C3"].number_format = ws["D3"].number_format = "yyyy-mm-dd"
    rows = [("按客户", None, None, 0, False, 0), ("甲公司", 600, 700.5, 1, False, 1), ("乙公司", 400, 300, 2, False, 1), ("合计", 1000, 1000.5, 0, True, 0)]
    for r, (name, a, b, indent, bold, level) in enumerate(rows, 4):
        c = ws.cell(r, 1, name)
        c.alignment = Alignment(horizontal="left", indent=indent)
        c.font = Font(bold=bold)
        if a is not None:
            ws.cell(r, 3, a).number_format = "#,##0"
            ws.cell(r, 4, b).number_format = "#,##0.0"
        ws.row_dimensions[r].outlineLevel = level
    ws["A9"] = "=备注：以 = 开头的文字不是公式"
    ws["A9"].data_type = "s"
    ws["B10"], ws["B11"], ws["B12"] = True, "#DIV/0!", -1234.5678901234
    ws["A13"] = "   leading spaces kept   "
    ws["E3"] = 0.123
    ws["E3"].number_format = "0.0%"
    ws["F5"] = "tab\tand\nnewline"
    ws.row_dimensions[8].hidden = True
    ws.column_dimensions["G"].hidden = True
    ws["G4"] = "in a hidden column"
    ws.merge_cells("A1:B1")
    h = wb.create_sheet("Hidden notes")
    h["A1"] = "hidden sheet text"
    h.sheet_state = "hidden"
    v = wb.create_sheet("VeryHidden")
    v["A1"] = 1
    v.sheet_state = "veryHidden"
    empty = wb.create_sheet("Empty")
    empty["A1"] = None
    wb.save(path)


class World:
    def __init__(self):
        self.tmp = Path(tempfile.mkdtemp(prefix="dump_"))
        self.book = self.tmp / "Book.xlsx"
        make_book(self.book)
        self.dump = self.tmp / "dump.txt"
        self.records, self.summary = D.dump_workbook(str(self.book))
        D.write_dump(self.records, self.dump)

    def restore(self, name="rebuilt.xlsx"):
        out = self.tmp / name
        D.restore_workbook(D.parse_dump(self.dump), out)
        return out


class Dump(unittest.TestCase):
    def test_what_is_in_the_dump_is_counted_so_a_person_can_see_the_indentation_arrived(self):
        w = World()
        s = w.summary
        self.assertEqual((s["sheets"], s["hidden_sheets"]), (4, 2))
        self.assertEqual(s["indented"], 2)                                         # 甲公司 and 乙公司; a zero indent is not indentation
        self.assertEqual(s["bold"], 1)
        self.assertEqual(s["outline_rows"], 2)
        self.assertEqual(s["merged_ranges"], 1)
        self.assertGreaterEqual(s["formatted"], 7)
        sheet = next(r for r in w.records if r["k"] == "sheet" and r["name"] == "应收账款")
        self.assertEqual(sheet["hidden_rows"], [8])
        self.assertEqual(sheet["hidden_cols"], ["G"])
        self.assertEqual(sheet["merged"], ["A1:B1"])
        self.assertEqual(sheet["outline"], {"5": 1, "6": 1})
        self.assertEqual({r["name"]: r["state"] for r in w.records if r["k"] == "sheet"},
                         {"应收账款": "visible", "Hidden notes": "hidden", "VeryHidden": "veryHidden", "Empty": "visible"})

    def test_cells_keep_their_exact_position_type_and_value(self):
        w = World()
        rows = {r["r"]: r["c"] for r in w.records if r["k"] == "row" and r["sheet"] == 0}
        a5 = next(c for c in rows[5] if c[0] == "A")
        self.assertEqual(a5[:3], ["A", "s", "甲公司"])
        self.assertIn("i1", a5[3])
        d3 = next(c for c in rows[3] if c[0] == "C")
        self.assertEqual(d3[:3], ["C", "d", "2024-12-31T00:00:00"])
        self.assertEqual(next(c for c in rows[5] if c[0] == "C")[2], 600)
        self.assertEqual(next(c for c in rows[5] if c[0] == "D")[2], 700.5)
        self.assertEqual(next(c for c in rows[9] if c[0] == "A")[:3], ["A", "s", "=备注：以 = 开头的文字不是公式"])
        self.assertEqual(next(c for c in rows[10] if c[0] == "B")[:3], ["B", "b", True])
        self.assertEqual(next(c for c in rows[11] if c[0] == "B")[:3], ["B", "e", "#DIV/0!"])
        self.assertEqual(next(c for c in rows[13] if c[0] == "A")[2], "   leading spaces kept   ")
        self.assertEqual(next(c for c in rows[5] if c[0] == "F")[2], "tab\tand\nnewline")

    def test_an_empty_sheet_is_kept_as_a_sheet_and_a_formula_without_a_cached_value_is_not_a_cell(self):
        w = World()
        self.assertIn("Empty", [r["name"] for r in w.records if r["k"] == "sheet"])
        # write a formula cell with no cached value by hand: the extractor reads values, so there is nothing to carry
        raw = w.tmp / "f.xlsx"
        with zipfile.ZipFile(w.book) as src, zipfile.ZipFile(raw, "w") as dst:
            for item in src.infolist():
                data = src.read(item.filename)
                if item.filename == "xl/worksheets/sheet1.xml":
                    data = data.replace(b"</sheetData>", b'<row r="40"><c r="A40"><f>SUM(C5:C6)</f></c></row></sheetData>')
                dst.writestr(item, data)
        records, _ = D.dump_workbook(str(raw))
        self.assertFalse([r for r in records if r["k"] == "row" and r["sheet"] == 0 and r["r"] == 40])


class Restore(unittest.TestCase):
    def test_the_rebuild_reads_back_with_the_same_digest_and_every_captured_property(self):
        w = World()
        out = w.restore()
        _, summary = D.dump_workbook(str(out))
        self.assertEqual(summary["digest"], w.summary["digest"])
        from openpyxl import load_workbook
        wb = load_workbook(out)
        ws = wb["应收账款"]
        self.assertEqual(ws["A5"].alignment.indent, 1)
        self.assertEqual(ws["A6"].alignment.indent, 2)
        self.assertTrue(ws["A7"].font.b)
        self.assertEqual(ws["C5"].number_format, "#,##0")
        self.assertEqual(ws["E3"].number_format, "0.0%")
        self.assertEqual(ws["C3"].value, datetime(2024, 12, 31))
        self.assertEqual(ws["A9"].value, "=备注：以 = 开头的文字不是公式")
        self.assertEqual(ws["A9"].data_type, "s")
        self.assertTrue(ws.row_dimensions[8].hidden)
        self.assertEqual(ws.row_dimensions[5].outlineLevel, 1)
        self.assertTrue(ws.column_dimensions["G"].hidden)
        self.assertEqual([str(r) for r in ws.merged_cells.ranges], ["A1:B1"])
        self.assertEqual([(s.title, s.sheet_state) for s in wb.worksheets],
                         [("应收账款", "visible"), ("Hidden notes", "hidden"), ("VeryHidden", "veryHidden"), ("Empty", "visible")])

    def test_a_rebuild_says_it_is_one_in_three_places_and_its_bytes_are_not_the_originals(self):
        w = World()
        out = w.restore()
        marker = json.loads(Path(str(out) + ".rebuilt.json").read_text(encoding="utf-8"))
        self.assertTrue(marker["rebuilt_from_a_text_dump"])
        self.assertEqual(marker["source_sha256"], w.records[0]["source_sha256"])
        with zipfile.ZipFile(out) as z:
            self.assertIn("REBUILT FROM A TEXT DUMP", z.read("docProps/core.xml").decode("utf-8"))
        self.assertIsNotNone(D.is_rebuilt(str(out)))
        Path(str(out) + ".rebuilt.json").unlink()
        self.assertIsNotNone(D.is_rebuilt(str(out)), "the mark inside the file survives losing the file beside it")
        self.assertIsNone(D.is_rebuilt(str(w.book)))
        import hashlib
        self.assertNotEqual(hashlib.sha256(out.read_bytes()).hexdigest(), w.records[0]["source_sha256"])

    def test_a_dump_pasted_into_markdown_between_prose_and_fences_restores(self):
        w = World()
        text = w.dump.read_text(encoding="utf-8")
        md = w.tmp / "paste.md"
        md.write_text("# the dump\n\nsome words\n\n```json\n" + text + "```\n\nmore words\n", encoding="utf-8")
        records = D.parse_dump(md)
        self.assertEqual(D.check_records(records), [])
        self.assertEqual(len(records), len(w.records))

    def test_a_line_lost_changed_or_cut_off_in_copying_is_found_before_anything_is_written(self):
        w = World()
        lines = w.dump.read_text(encoding="utf-8").splitlines()
        # a row line dropped
        dropped = [ln for i, ln in enumerate(lines) if i != 8]
        (w.tmp / "d1.txt").write_text("\n".join(dropped), encoding="utf-8")
        self.assertTrue(any("do not match the digest" in p for p in D.check_records(D.parse_dump(w.tmp / "d1.txt"))))
        # one value changed
        changed = [ln.replace("甲公司", "甲公司X") if "甲公司" in ln else ln for ln in lines]
        (w.tmp / "d2.txt").write_text("\n".join(changed), encoding="utf-8")
        self.assertTrue(D.check_records(D.parse_dump(w.tmp / "d2.txt")))
        # the end cut off
        (w.tmp / "d3.txt").write_text("\n".join(lines[:-3]), encoding="utf-8")
        self.assertTrue(any("no end line" in p for p in D.check_records(D.parse_dump(w.tmp / "d3.txt"))))
        # the start cut off
        (w.tmp / "d4.txt").write_text("\n".join(lines[2:]), encoding="utf-8")
        self.assertTrue(any("no meta line" in p for p in D.check_records(D.parse_dump(w.tmp / "d4.txt"))))

    def test_the_command_reports_equal_and_refuses_to_overwrite(self):
        import contextlib
        w = World()
        buf = io.StringIO()
        out = w.tmp / "cmd.xlsx"
        with contextlib.redirect_stdout(buf):
            self.assertEqual(D.main(["restore", str(w.dump), "--out", str(out)]), 0)
            self.assertEqual(D.main(["restore", str(w.dump), "--out", str(out)]), 2)
        text = buf.getvalue()
        self.assertIn("digest EQUAL", text)
        self.assertIn("a rebuild, not the original", text)
        self.assertNotIn("甲公司", text)
        self.assertNotIn("应收账款", text)


class AgainstTheExtractor(unittest.TestCase):
    """The proof that matters: the signals fdd_utils/workbook/schedules reads are the same from the rebuild as from the original."""

    def signals(self, path: Path):
        import pandas as pd
        from openpyxl import load_workbook
        from fdd_utils.workbook import schedules as S
        df = pd.read_excel(path, sheet_name="应收账款", header=None, engine="openpyxl")
        ws = load_workbook(path)["应收账款"]
        sig = S._extract_sheet_style_signals(ws, df, 0)
        return df, {k: sig[k] for k in ("indent_rows", "bold_rows", "outline_levels", "merged_ranges", "number_formats")}

    def test_indent_bold_outline_merges_and_number_formats_come_out_the_same(self):
        w = World()
        out = w.restore()
        df_a, a = self.signals(w.book)
        df_b, b = self.signals(out)
        self.assertEqual(a, b)
        self.assertTrue(a["indent_rows"] and any(level > 0 for _, level, _ in a["indent_rows"]), "the original has indentation to lose")
        self.assertEqual(a["bold_rows"] and len(a["bold_rows"]), 1)
        self.assertTrue(a["number_formats"])

    def test_the_frames_pandas_reads_are_the_same_cell_for_cell(self):
        import pandas as pd
        w = World()
        out = w.restore()
        for name in ("应收账款", "Hidden notes", "VeryHidden"):
            a = pd.read_excel(w.book, sheet_name=name, header=None, engine="openpyxl")
            b = pd.read_excel(out, sheet_name=name, header=None, engine="openpyxl")
            pd.testing.assert_frame_equal(a, b, check_dtype=False)

    def test_the_repos_own_style_reader_on_the_raw_xml_agrees_with_this_one(self):
        from fdd_utils.workbook import schedules as S
        w = World()
        with zipfile.ZipFile(w.book) as z:
            repo = S._xlsx_styles(z)
            mine = D.read_styles(z)
        self.assertEqual(repo["indent"], [m["indent"] for m in mine])
        self.assertEqual(repo["bold"], [m["bold"] for m in mine])
        self.assertEqual(repo["numfmt"], [m["numfmt"] for m in mine])


if __name__ == "__main__":
    unittest.main(verbosity=2)

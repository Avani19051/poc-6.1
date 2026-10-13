#!/usr/bin/env python3
"""Tests for mcmaster_lookup.py  (run:  python -m unittest tests_mcmaster_lookup -v)."""

import tempfile
import unittest
from pathlib import Path

import openpyxl

import case2_mcmaster as m


class TestMpnValidation(unittest.TestCase):
    def test_normalise(self):
        self.assertEqual(m.normalise_mpn(" 91864a028 "), "91864A028")
        self.assertEqual(m.normalise_mpn("91 864 A028"), "91864A028")

    def test_valid_patterns(self):
        for mpn in ("91864A028", "90044A249", "91251A107", "91251A105"):
            self.assertRegex(mpn, m.MPN_RE)

    def test_invalid_patterns(self):
        for mpn in ("", "91864", "ABC", "91864A028X", "918641A028", "hello"):
            self.assertNotRegex(mpn, m.MPN_RE)

    def test_product_url(self):
        self.assertEqual(m.PRODUCT_URL.format(mpn="91864a028"),
                         "https://www.mcmaster.com/products/91864a028/")


class TestParsers(unittest.TestCase):
    def test_rows(self):
        rows = [
            ["Thread Size", "1/4\""],
            ["Material", "Black-Oxide Alloy Steel"],
            ["RoHS Compliance", "RoHS 3 (2015/863/EU) Compliant"],
        ]
        out = m.parse_specs_from_rows(rows)
        self.assertEqual(out["Material"], "Black-Oxide Alloy Steel")
        self.assertEqual(out["RoHS Compliance"], "RoHS 3 (2015/863/EU) Compliant")

    def test_rows_ignores_catalog_category(self):
        # "Material Handling" must NOT be mistaken for the "Material" spec.
        rows = [["Material Handling", "Some category"]]
        self.assertNotIn("Material", m.parse_specs_from_rows(rows))

    def test_text_alternating_lines(self):
        text = "Specifications\nMaterial\nBlack-Oxide Alloy Steel\nMid\nRoHS Compliance\nRoHS 3 (2015/863/EU) Compliant\nEnd"
        out = m.parse_specs_from_text(text)
        self.assertEqual(out["Material"], "Black-Oxide Alloy Steel")
        self.assertEqual(out["RoHS Compliance"], "RoHS 3 (2015/863/EU) Compliant")

    def test_text_missing_values(self):
        self.assertEqual(m.parse_specs_from_text("nothing here"), {})

    def test_full_parse_prefers_rows_and_flags_blocked(self):
        rows = [["Material", "18-8 Stainless Steel"], ["RoHS", "Compliant"]]
        text = "Please log in\nTo continue browsing, please log in"
        specs = m.parse_specs(text, rows)
        self.assertEqual(specs.material, "18-8 Stainless Steel")
        self.assertEqual(specs.rohs, "Compliant")
        self.assertTrue(specs.blocked)
        self.assertTrue(specs.is_complete())

    def test_full_parse_not_found(self):
        specs = m.parse_specs("plain page", [])
        self.assertEqual(specs.material, m.NOT_FOUND)
        self.assertEqual(specs.rohs, m.NOT_FOUND)
        self.assertFalse(specs.is_complete())


class TestExcelRoundtrip(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)
        self.table_path = self.dir / "TableA.xlsx"

    def tearDown(self):
        self.tmp.cleanup()

    def test_create_table_from_brief(self):
        m.ensure_table_a(self.table_path)
        wb = openpyxl.load_workbook(self.table_path)
        ws = wb.active
        self.assertEqual([c.value for c in ws[1]], m.HEADERS)
        self.assertEqual(ws.max_row, 1 + len(m.DEFAULT_ROWS))

    def test_load_update_save(self):
        m.ensure_table_a(self.table_path)
        out = self.dir / "TableA_Updated.xlsx"
        table = m.load_table(self.table_path, out)
        self.assertEqual(set(table.mpns),
                         {"91864A028", "90044A249", "91251A107", "91251A105"})
        table.set_cell("91864A028", "Material", "Black-Oxide Alloy Steel")
        table.set_cell("91864A028", "RoHS", "RoHS 3 (2015/863/EU) Compliant")
        table.save(out)

        wb = openpyxl.load_workbook(out)
        ws = wb.active
        row = {str(r[0]): (r[2], r[3]) for r in ws.iter_rows(values_only=True)}
        self.assertEqual(row["91864A028"],
                         ("Black-Oxide Alloy Steel", "RoHS 3 (2015/863/EU) Compliant"))
        self.assertEqual(row["91251A105"], ("?", "?"))  # untouched rows preserved

    def test_rerun_continues_from_previous_output(self):
        m.ensure_table_a(self.table_path)
        out = self.dir / "TableA_Updated.xlsx"
        t1 = m.load_table(self.table_path, out)
        t1.set_cell("90044A249", "Material", "Alloy Steel")
        t1.save(out)
        # A fresh run should read the output and keep the earlier result.
        t2 = m.load_table(self.table_path, out)
        self.assertEqual(t2.sheet.cell(t2.rows_by_mpn["90044A249"],
                                       t2.header["MATERIAL"]).value, "Alloy Steel")

    def test_missing_column_is_reported(self):
        wb = openpyxl.Workbook()
        wb.active.append(["MPN", "MANUFACTURER NAME", "Material"])  # no RoHS
        wb.save(self.table_path)
        with self.assertRaises(RuntimeError):
            m.load_table(self.table_path, self.dir / "nope.xlsx")


if __name__ == "__main__":
    unittest.main(verbosity=2)

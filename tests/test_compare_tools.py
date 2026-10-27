"""Offline tests for compare_workbook_values.py and compare_vba_source.py.

Synthetic packages and mocked extraction only: no Excel, no network, no real
macro project. Every equality test is paired with a negative test that injects
one deliberate difference, so a check that cannot detect anything fails here.
"""
from __future__ import annotations
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import compare_workbook_values as cwv
import compare_vba_source as cvs

CT = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Target="xl/workbook.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"/>
</Relationships>"""


def esc(text: str) -> str:
    return (text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def sheet_xml(rows) -> str:
    out = ['<?xml version="1.0" encoding="UTF-8"?><worksheet '
           'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>']
    for r, row in enumerate(rows, start=1):
        out.append('<row r="%d">' % r)
        for c, val in enumerate(row):
            ref = "%s%d" % (chr(65 + c), r)
            if val is None:
                continue
            if isinstance(val, bool):
                out.append('<c r="%s" t="b"><v>%d</v></c>' % (ref, 1 if val else 0))
            elif isinstance(val, (int, float)):
                out.append('<c r="%s"><v>%s</v></c>' % (ref, val))
            elif str(val).startswith("#"):
                out.append('<c r="%s" t="e"><v>%s</v></c>' % (ref, esc(str(val))))
            else:
                out.append('<c r="%s" t="inlineStr"><is><t>%s</t></is></c>'
                           % (ref, esc(str(val))))
        out.append('</row>')
    out.append('</sheetData></worksheet>')
    return "".join(out)


def build_xlsx(path: Path, sheets: dict[str, list]) -> Path:
    """Minimal but schema-shaped package: one part per sheet, inline strings."""
    names = list(sheets)
    wb = ['<?xml version="1.0" encoding="UTF-8"?>'
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>']
    rels = ['<?xml version="1.0" encoding="UTF-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", ROOT_RELS)
        for i, name in enumerate(names, start=1):
            wb.append('<sheet name="%s" sheetId="%d" r:id="rId%d"/>'
                      % (esc(name), i, i))
            target = "worksheets/sheet%d.xml" % i
            rels.append('<Relationship Id="rId%d" Target="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/>'
                        % (i, target))
            z.writestr("xl/" + target, sheet_xml(sheets[name]))
        wb.append('</sheets></workbook>')
        rels.append('</Relationships>')
        z.writestr("xl/workbook.xml", "".join(wb))
        z.writestr("xl/_rels/workbook.xml.rels", "".join(rels))
    return path


class ValueComparisonTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.dir = Path(self.temp.name)

    def test_identical_packages(self):
        a = build_xlsx(self.dir / "a.xlsx", {"Dane": [["Klucz", 1], ["Inny", 2]]})
        b = build_xlsx(self.dir / "b.xlsx", {"Dane": [["Klucz", 1], ["Inny", 2]]})
        r = cwv.compare(a, b, None, 1000)
        self.assertEqual(r["status"], "IDENTICAL")
        self.assertGreater(r["cells_a"], 0)

    def test_one_changed_cell_is_detected_with_location(self):
        a = build_xlsx(self.dir / "a.xlsx", {"Dane": [["Klucz", 1], ["Inny", 2]]})
        b = build_xlsx(self.dir / "b.xlsx", {"Dane": [["Klucz", 1], ["Inny", 99]]})
        r = cwv.compare(a, b, None, 1000)
        self.assertEqual(r["status"], "DIFFERS")
        self.assertEqual([s["sheet"] for s in r["changed_sheets"]], ["Dane"])
        self.assertEqual(r["changed_sheets"][0]["first_differing_rows"], [2])

    def test_added_rows_are_detected(self):
        a = build_xlsx(self.dir / "a.xlsx", {"K": [["a"], ["b"]]})
        b = build_xlsx(self.dir / "b.xlsx", {"K": [["a"], ["b"], ["c"]]})
        r = cwv.compare(a, b, None, 1000)
        self.assertEqual(r["status"], "DIFFERS")
        self.assertEqual(r["changed_sheets"][0]["rows_b"], 3)

    def test_added_sheet_is_detected(self):
        a = build_xlsx(self.dir / "a.xlsx", {"K": [["a"]]})
        b = build_xlsx(self.dir / "b.xlsx", {"K": [["a"]], "Nowy": [["x"]]})
        r = cwv.compare(a, b, None, 1000)
        self.assertEqual(r["status"], "DIFFERS")
        self.assertEqual(r["only_in_b"], ["Nowy"])

    def test_numeric_forms_compare_by_value(self):
        a = build_xlsx(self.dir / "a.xlsx", {"K": [[1], [True]]})
        b = build_xlsx(self.dir / "b.xlsx", {"K": [[1.0], [True]]})
        self.assertEqual(cwv.compare(a, b, None, 1000)["status"], "IDENTICAL")

    def test_empty_sheet_refuses_equality(self):
        """Non-vacuity: reading nothing must never be reported as IDENTICAL."""
        a = build_xlsx(self.dir / "a.xlsx", {"Pusty": [[]]})
        b = build_xlsx(self.dir / "b.xlsx", {"Pusty": [[]]})
        with self.assertRaises(ValueError):
            cwv.compare(a, b, None, 1000)

    def test_truncation_raises_instead_of_partial_pass(self):
        rows = [[i] for i in range(50)]
        a = build_xlsx(self.dir / "a.xlsx", {"K": rows})
        b = build_xlsx(self.dir / "b.xlsx", {"K": rows})
        with self.assertRaises(ValueError):
            cwv.read_values(a, None, 10)

    def test_non_ooxml_is_rejected(self):
        p = self.dir / "legacy.xls"
        p.write_bytes(b"not a zip")
        with self.assertRaises(ValueError):
            cwv.read_values(p)

    def test_missing_workbook_part_is_rejected(self):
        p = self.dir / "junk.xlsx"
        with zipfile.ZipFile(p, "w") as z:
            z.writestr("[Content_Types].xml", CT)
        with self.assertRaises(ValueError):
            cwv.read_values(p)

    def test_column_index_parsing(self):
        self.assertEqual(cwv._col_index("A1"), 1)
        self.assertEqual(cwv._col_index("Z9"), 26)
        self.assertEqual(cwv._col_index("AA1"), 27)
        self.assertEqual(cwv._col_index("AB12"), 28)
        with self.assertRaises(ValueError):
            cwv._col_index("12")

    def test_moving_a_value_to_a_different_excel_row_is_detected(self):
        """An OOXML <row> is not its array position: A1 -> A5 is a real edit."""
        a = build_xlsx(self.dir / "a.xlsx", {"K": [["cost", 10]]})
        b = build_xlsx(self.dir / "b.xlsx", {"K": [["cost", 10]]})
        with zipfile.ZipFile(b) as z:
            entries = {name: z.read(name) for name in z.namelist()}
        key = "xl/worksheets/sheet1.xml"
        entries[key] = (entries[key].replace(b'r="1"', b'r="5"')
                        .replace(b'r="A1"', b'r="A5"')
                        .replace(b'r="B1"', b'r="B5"'))
        with zipfile.ZipFile(b, "w") as z:
            for name, data in entries.items():
                z.writestr(name, data)
        r = cwv.compare(a, b, None, 1000)
        self.assertEqual(r["status"], "DIFFERS")
        self.assertEqual(r["changed_sheets"][0]["first_differing_rows"], [1, 5])

    def test_exact_numeric_bytes_do_not_collapse_through_float(self):
        """Two OOXML decimal strings that float rounds to one number differ."""
        a = build_xlsx(self.dir / "a.xlsx", {"K": [[9007199254740992]]})
        b = build_xlsx(self.dir / "b.xlsx", {"K": [[9007199254740993]]})
        self.assertEqual(cwv.compare(a, b, None, 1000)["status"], "DIFFERS")

    def test_boolean_is_not_equal_to_text_that_looks_boolean(self):
        a = build_xlsx(self.dir / "a.xlsx", {"K": [[True]]})
        b = build_xlsx(self.dir / "b.xlsx", {"K": [["TRUE"]]})
        self.assertEqual(cwv.compare(a, b, None, 1000)["status"], "DIFFERS")

    def test_false_cell_row_address_is_rejected(self):
        a = build_xlsx(self.dir / "a.xlsx", {"K": [["x"]]})
        with zipfile.ZipFile(a) as z:
            entries = {name: z.read(name) for name in z.namelist()}
        key = "xl/worksheets/sheet1.xml"
        entries[key] = entries[key].replace(b'<row r="1">', b'<row r="2">')
        with zipfile.ZipFile(a, "w") as z:
            for name, data in entries.items():
                z.writestr(name, data)
        with self.assertRaises(cwv.PackageError):
            cwv.read_values(a)

    def test_scope_statement_present_in_report(self):
        a = build_xlsx(self.dir / "a.xlsx", {"K": [["a"]]})
        b = build_xlsx(self.dir / "b.xlsx", {"K": [["a"]]})
        self.assertIn("NOT compared", cwv.compare(a, b, None, 1000)["scope"])


class VbaSourceComparisonTests(unittest.TestCase):
    """Extraction is mocked; the comparison contract and its guards are not."""

    def test_identical_source(self):
        src = {"Mod1.bas": "Sub A()\nEnd Sub\n", "Mod2.bas": "Sub B()\nEnd Sub\n"}
        with patch.object(cvs, "extract_vba_source", side_effect=[src, dict(src)]):
            r = cvs.compare(Path("a.xlsm"), Path("b.xlsm"))
        self.assertEqual(r["status"], "IDENTICAL")
        self.assertEqual(r["source_lines_a"], 4)

    def test_changed_component_is_located(self):
        a = {"Mod1.bas": "Sub A()\n  x = 1\nEnd Sub\n"}
        b = {"Mod1.bas": "Sub A()\n  x = 2\nEnd Sub\n"}
        with patch.object(cvs, "extract_vba_source", side_effect=[a, b]):
            r = cvs.compare(Path("a.xlsm"), Path("b.xlsm"))
        self.assertEqual(r["status"], "DIFFERS")
        self.assertEqual(r["changed_components"], ["Mod1.bas"])
        self.assertEqual(r["changed_detail"][0]["first_differing_lines"], [2])

    def test_added_component_is_detected(self):
        a = {"Mod1.bas": "Sub A()\nEnd Sub\n"}
        b = {"Mod1.bas": "Sub A()\nEnd Sub\n", "Mod2.bas": "Sub B()\nEnd Sub\n"}
        with patch.object(cvs, "extract_vba_source", side_effect=[a, b]):
            r = cvs.compare(Path("a.xlsm"), Path("b.xlsm"))
        self.assertEqual(r["status"], "DIFFERS")
        self.assertEqual(r["only_in_b"], ["Mod2.bas"])

    def test_empty_extraction_refuses_equality(self):
        """The false-PASS this tool exists to prevent."""
        with patch.object(cvs, "extract_vba_source", return_value={}):
            with self.assertRaises(ValueError):
                cvs.compare(Path("a.xlsm"), Path("b.xlsm"))

    def test_zero_line_extraction_refuses_equality(self):
        with patch.object(cvs, "extract_vba_source", side_effect=[{"M.bas": ""}, {"M.bas": ""}]):
            with self.assertRaises(ValueError):
                cvs.compare(Path("a.xlsm"), Path("b.xlsm"))

    def test_non_ooxml_is_rejected(self):
        with tempfile.TemporaryDirectory() as t:
            p = Path(t) / "legacy.xls"
            p.write_bytes(b"not a zip")
            with self.assertRaises(ValueError):
                cvs.extract_vba_source(p)

    def test_macro_free_package_is_rejected(self):
        with tempfile.TemporaryDirectory() as t:
            p = Path(t) / "plain.xlsx"
            with zipfile.ZipFile(p, "w") as z:
                z.writestr("[Content_Types].xml", CT)
            with self.assertRaises(ValueError):
                cvs.extract_vba_source(p)

    def test_bad_ole_header_is_rejected(self):
        with self.assertRaises(ValueError):
            cvs.CompoundFile(b"x" * 600)

    def test_short_ole_header_is_rejected(self):
        with self.assertRaises(ValueError):
            cvs.CompoundFile(bytes.fromhex("D0CF11E0A1B11AE1") + b"\x00" * 100)

    def test_big_endian_ole_is_rejected(self):
        header = bytearray(bytes.fromhex("D0CF11E0A1B11AE1") + b"\x00" * 500)
        header[28:30] = (0xFF00).to_bytes(2, "little")  # byte order mark
        with self.assertRaises(ValueError):
            cvs.CompoundFile(bytes(header))

    def test_undecodable_source_must_not_be_replaced(self):
        """cp1252 0x81 and 0x8D previously both became U+FFFD: false pass."""
        with self.assertRaises(ValueError):
            cvs._decode(b"\x81", "cp1252")
        with self.assertRaises(ValueError):
            cvs._decode(b"\x8d", "cp1252")
        with self.assertRaises(ValueError):
            cvs._decode(b"Attribute VB_Name", None)

    def test_multichunk_vba_source_probe_uses_exact_chunk_length(self):
        # When probing the first chunk, do not include one byte of the second
        # header: strict decompression would then reject a valid container.
        def chunk(raw):
            payload = b"".join(b"\x00" + raw[i:i + 8] for i in range(0, len(raw), 8))
            header = (0xB000 | (len(payload) - 1)).to_bytes(2, "little")
            return header + payload
        one = b'Attribute VB_Name = "Module1"\r\n'
        two = b'Sub Work()\r\nEnd Sub\r\n'
        container = b"\x01" + chunk(one) + chunk(two)
        self.assertEqual(cvs.find_source_start(container), 0)
        self.assertEqual(cvs.vba_decompress(container), one + two)

    def test_correct_declared_polish_codepage_keeps_diacritics(self):
        self.assertEqual(cvs._decode("Źródła".encode("cp1250"), "cp1250"), "Źródła")

    def test_truncated_vba_chunk_is_not_silently_accepted(self):
        # Chunk header advertises two bytes of payload; there is only one.
        malformed = b"\x01" + (0xB001).to_bytes(2, "little") + b"\x00"
        with self.assertRaises(ValueError):
            cvs.vba_decompress(malformed)

    def test_ole_allocation_cycle_must_fail(self):
        reader = cvs.CompoundFile.__new__(cvs.CompoundFile)
        reader._fat = [0]  # sector 0 -> 0 forever
        with self.assertRaises(ValueError):
            reader._chain_sectors(0, 0)

    def test_scope_statement_present_in_report(self):
        src = {"M.bas": "Sub A()\nEnd Sub\n"}
        with patch.object(cvs, "extract_vba_source", side_effect=[src, dict(src)]):
            r = cvs.compare(Path("a.xlsm"), Path("b.xlsm"))
        self.assertIn("compiled behaviour", r["scope"])


if __name__ == "__main__":
    unittest.main()

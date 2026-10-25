"""Real-snapshot namespace-routing and fabricated-member regression tests."""
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from offline_kb import KnowledgeBase
from api_preflight import check, namespaced, normalize_symbol


class APIPreflightTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.kb = KnowledgeBase()

    @classmethod
    def tearDownClass(cls):
        cls.kb.close()

    def test_five_library_examples_have_exact_source(self):
        cases = {
            'vba': ('ReDim', 'Variant', 'DateAdd', 'Collection'),
            'excel': ('Workbook', 'Range', 'Chart'),
            'msforms': ('UserForm', 'TextBox', 'ListBox', 'ComboBox'),
            'office': ('CommandBars', 'FileDialog'),
            'vbide': ('VBProject', 'VBComponent'),
        }
        for ns, names in cases.items():
            for symbol in names:
                with self.subTest(ns=ns, symbol=symbol):
                    r = check(self.kb, ns, symbol)
                    self.assertEqual(r['status'], 'SOURCE_DOCUMENT_FOUND', r)
                    self.assertTrue(r['evidence'])
                    self.assertTrue(all(e['sha256_verified'] for e in r['evidence']))
                    self.assertTrue(all(namespaced(e['source_path'], ns) for e in r['evidence']))

    def test_no_wrong_library_for_chart_or_forms(self):
        result = check(self.kb, 'excel', 'Chart')
        self.assertIn('api/Excel.Chart(object).md', [e['source_path'] for e in result['evidence']])
        self.assertTrue(all('-graph' not in e['source_path'].casefold() for e in result['evidence']))
        for symbol in ('ListBox', 'ComboBox', 'TextBox'):
            self.assertTrue(all(e['source_path'].startswith('Language/Reference/')
                                for e in check(self.kb, 'msforms', symbol)['evidence']))

    def test_nonexistent_form_members_not_claimed(self):
        for symbol in ('ListBox.ItemHeight', 'UserForm.ScaleMode'):
            result = check(self.kb, 'msforms', symbol)
            self.assertEqual(result['status'], 'MEMBER_NOT_CONFIRMED_OWNER_ONLY')
            self.assertEqual({e['evidence_type'] for e in result['evidence']}, {'OWNER_TYPE_CONTEXT'})

    def test_combined_forms_property_page_is_only_a_candidate(self):
        r = check(self.kb, 'msforms', 'UserForm.InsideWidth')
        self.assertEqual(r['status'], 'CANDIDATES_REQUIRE_MEMBER_VERIFICATION')
        self.assertTrue(any(e['source_path'].endswith('insideheight-insidewidth-properties.md')
                            for e in r['evidence']))
        self.assertNotIn('DIRECT_MEMBER_DOCUMENT', {e['evidence_type'] for e in r['evidence']})

    def test_direct_excel_method_is_not_runtime_validation(self):
        r = check(self.kb, 'excel', 'Range.Find')
        self.assertEqual(r['status'], 'MEMBER_DOCUMENT_FOUND_CHECK_OWNER')
        self.assertTrue(any(e['source_path'] == 'api/Excel.Range.Find.md' for e in r['evidence']))
        self.assertIn('not an Excel compile/runtime test', r['warning'])

    def test_vbide_grouped_heading_evidence(self):
        r = check(self.kb, 'vbide', 'VBComponent')
        self.assertEqual(r['status'], 'SOURCE_DOCUMENT_FOUND')
        self.assertEqual(r['evidence'][0]['source_lines'], [131, 142])
        self.assertIn('Visual-Basic-Add-in-Model/', r['evidence'][0]['source_path'])

    def test_namespace_mismatch_and_unsafe_symbol(self):
        with self.assertRaises(ValueError):
            normalize_symbol('msforms', 'Excel.Range.Find')
        with self.assertRaises(ValueError):
            normalize_symbol('excel', 'Range.Find; DROP TABLE documents')
        with self.assertRaises(ValueError):
            normalize_symbol('vba', 'A.B.C')

    def test_unknown_member_not_false_proven(self):
        r = check(self.kb, 'excel', 'Range.FictionalMemberMadeUp')
        self.assertNotEqual(r['status'], 'MEMBER_DOCUMENT_FOUND_CHECK_OWNER')
        self.assertNotEqual(r['status'], 'SOURCE_DOCUMENT_FOUND')


if __name__ == '__main__':
    unittest.main()

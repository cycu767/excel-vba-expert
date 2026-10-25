"""Regression, real-snapshot and security tests for local full-document retrieval."""
from __future__ import annotations
import hashlib
import json
import os
import shutil
import socket
import sqlite3
import stat
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))
from offline_kb import (KnowledgeBase, parse_doc, safe_relative, markdown_targets,
                        api_key, category_for, contained_file)
from import_knowledge import import_repository

SAMPLES = {
    'Range.Find': ('api/Excel.Range.Find.md', ['What', 'After', 'LookIn', 'Nothing', '## Remarks', '## Examples']),
    'Workbooks.Open': ('api/Excel.Workbooks.Open.md', ['FileName', 'UpdateLinks', 'ReadOnly', 'AutomationSecurity', '## Example']),
    'Worksheet.Change': ('api/Excel.Worksheet.Change.md', ['Target', 'more than one cell', 'recalculation', 'Worksheet_Change']),
    'UserForm.Show': ('Language/Reference/User-Interface-Help/show-method.md', ['modal', 'vbModal', 'vbModeless', '## Remarks', 'UserForm2.Show']),
    'ChartObjects.Add': ('api/Excel.ChartObjects.Add.md', ['Left', 'Top', 'Width', 'Height', 'ChartObject', '## Example']),
}


class SnapshotTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.kb = KnowledgeBase()
    @classmethod
    def tearDownClass(cls):
        cls.kb.close()
    def test_supplied_snapshot_count_contract(self):
        c = self.kb.manifest['counts']
        self.assertGreater(c['source_markdown'], 10000)
        self.assertEqual(c['source_markdown'], c['indexed_markdown'])
        self.assertEqual(c['indexed_markdown'], c['local_markdown'])
    def test_full_text_row_count_matches_documents(self):
        fts_count = self.kb.db.execute('SELECT COUNT(*) FROM document_fts').fetchone()[0]
        self.assertEqual(fts_count, self.kb.manifest['counts']['indexed_markdown'])
    def test_every_document_path_is_a_real_file(self):
        count = 0
        for row in self.kb.db.execute('SELECT path FROM documents'):
            self.assertTrue(contained_file(self.kb.repository, row['path']).is_file())
            count += 1
        self.assertEqual(count, self.kb.manifest['counts']['indexed_markdown'])
    def test_all_source_files_and_indexes_hash_verify(self):
        r = self.kb.verify()
        self.assertEqual(r['status'], 'PASS', r['errors'])
        self.assertEqual(r['hash_verified_files'], self.kb.manifest['counts']['source_files'])
    def test_all_documents_have_exact_path_lookup(self):
        rows = self.kb.db.execute('SELECT path FROM documents').fetchall()
        for row in rows:
            self.assertEqual(self.kb.document(row['path'], 'all')['path'], row['path'])
    def test_all_excel_api_names_are_retrievable(self):
        for row in self.kb.db.execute("SELECT path,api_names FROM documents WHERE category='excel-api'"):
            for name in json.loads(row['api_names']):
                hits = self.kb.search(name, exact=True, limit=100)
                self.assertIn(row['path'], [d['path'] for d in hits], name)
    def test_range_find_full_source(self):
        self.assert_sample('Range.Find')
    def test_workbooks_open_full_source(self):
        self.assert_sample('Workbooks.Open')
    def test_worksheet_change_full_source(self):
        self.assert_sample('Worksheet.Change')
    def test_userform_show_full_source(self):
        self.assert_sample('UserForm.Show')
    def test_chartobjects_add_full_source(self):
        self.assert_sample('ChartObjects.Add')
    def assert_sample(self, api):
        path, needles = SAMPLES[api]
        r = self.kb.read(api)
        self.assertEqual(r['document']['path'], path)
        self.assertTrue(r['complete_document'])
        self.assertTrue(r['source_sha256_verified'])
        raw = (self.kb.repository / path).read_bytes()
        self.assertEqual(r['content'], raw.decode('utf-8-sig'))
        self.assertEqual(hashlib.sha256(raw).hexdigest(), r['document']['sha256'])
        for needle in needles:
            self.assertIn(needle, r['content'])
    def test_userform_parameter_parts_are_in_syntax_not_fabricated_heading(self):
        r = self.kb.read('UserForm.Show', section='Syntax')
        self.assertIn('| _modal_|', r['content'])
        with self.assertRaises(LookupError):
            self.kb.read('UserForm.Show', section='Parameters')
    def test_chart_missing_remarks_remains_missing(self):
        with self.assertRaises(LookupError):
            self.kb.read('ChartObjects.Add', section='Remarks')
    def test_chart_parameter_table_gap_is_not_silently_filled(self):
        r = self.kb.read('ChartObjects.Add', section='Parameters')['content']
        self.assertIn('| _Left_|', r)
        self.assertIn('| _Width_|', r)
        self.assertNotIn('| _Top_|', r)
        self.assertNotIn('| _Height_|', r)
    def test_enum_by_type(self):
        self.assertEqual(self.kb.document('XlLookAt')['kind'], 'enumeration')
    def test_enum_by_constant(self):
        self.assertEqual(self.kb.document('xlWhole')['path'], 'api/Excel.XlLookAt.md')
    def test_object_property_event_method_lookup(self):
        for api, kind in [('Excel.Workbooks','object'),('Range.Value2','property'),
                          ('Worksheet_Change','event'),('Range.Find','method')]:
            hits = self.kb.search(api, exact=True)
            self.assertTrue(any(d['kind'] == kind for d in hits), (api, hits))
    def test_namespace_case_and_excel_prefix(self):
        for s in ('range.find','EXCEL.RANGE.FIND','Range.Find method (Excel)'):
            self.assertEqual(self.kb.document(s)['path'], 'api/Excel.Range.Find.md')
    def test_ambiguous_object_does_not_guess_legacy_graph(self):
        with self.assertRaisesRegex(ValueError, 'Ambiguous'):
            self.kb.document('Range')
    def test_exact_fictional_api_miss_not_invented(self):
        self.assertEqual(self.kb.search('Range.SetFancyGradient', exact=True), [])
        with self.assertRaisesRegex(LookupError, 'does not prove'):
            self.kb.read('Range.SetFancyGradient')
    def test_curated_source_id_crosswalk(self):
        self.assertEqual(self.kb.document('MS-FIND')['path'], 'api/Excel.Range.Find.md')
    def test_full_text_search_finds_body_details(self):
        self.assertIn('api/Excel.Range.Find.md', [d['path'] for d in self.kb.search('saved settings Find')])
    def test_scope_and_kind_filters(self):
        r = self.kb.search('Open', kind='method', category='excel-api')
        self.assertTrue(r)
        self.assertTrue(all(d['kind']=='method' and d['category']=='excel-api' for d in r))
    def test_other_office_document_available(self):
        r = self.kb.read('api/Word.Document.md', scope='all')
        self.assertTrue(r['complete_document'])
    def test_queries_cannot_inject_sql_or_fts(self):
        self.kb.search('" OR 1=1; DROP TABLE documents; --')
        self.assertEqual(self.kb.db.execute('SELECT COUNT(*) FROM documents').fetchone()[0],
                         self.kb.manifest['counts']['indexed_markdown'])
    def test_include_sources_are_available_without_network(self):
        result = self.kb.includes('Range.Find')
        self.assertEqual({r['path'] for r in result}, {'includes/addinsnote.md','includes/feedback-boilerplate.md'})
        self.assertTrue(all(r['content'] for r in result))
    def test_link_case_and_parentheses_resolve_offline(self):
        r = self.kb.resolve_link('api/Excel.Range.Find.md', 'excel.range(object).md')
        self.assertEqual(r['path'], 'api/Excel.Range(object).md')
    def test_relative_glossary_link_and_anchor(self):
        r = self.kb.resolve_link('api/Excel.ChartObjects.Add.md','../language/glossary/vbe-glossary.md#point')
        self.assertEqual(r['status'], 'local'); self.assertEqual(r['anchor'], 'point')
    def test_learn_link_resolves_locally_without_request(self):
        r = self.kb.resolve_link('api/Excel.Range.Find.md','https://learn.microsoft.com/en-us/office/vba/api/excel.range.find')
        self.assertEqual(r['path'],'api/Excel.Range.Find.md')
    def test_external_link_not_fetched(self):
        self.assertEqual(self.kb.resolve_link('x.md','https://example.invalid/nope')['status'],'external_not_fetched')
    def test_link_traversal_rejected(self):
        self.assertEqual(self.kb.resolve_link('api/a.md','../../escape.md')['status'],'unsafe_link_rejected')
    def test_partial_read_reports_original_lines(self):
        r = self.kb.read('Range.Find',start_line=20,end_line=25)
        self.assertFalse(r['complete_document']); self.assertEqual(r['start_line'],20)
        self.assertIn('## Syntax',r['content'])
    def test_invalid_ranges_rejected(self):
        for args in ({'start_line':0},{'end_line':99999},{'start_line':30,'end_line':20}):
            with self.assertRaises(ValueError): self.kb.read('Range.Find',**args)
    def test_sqlite_is_readonly(self):
        with self.assertRaises(sqlite3.OperationalError): self.kb.db.execute('DELETE FROM documents')
    def test_lookup_no_socket_calls(self):
        with patch.object(socket,'socket',side_effect=AssertionError('network call')):
            with patch.object(socket,'getaddrinfo',side_effect=AssertionError('DNS call')):
                self.assert_sample('Range.Find')
                self.kb.search('modal modeless')
    def test_cli_works_from_an_unrelated_working_directory(self):
        with tempfile.TemporaryDirectory() as d:
            result = subprocess.run([sys.executable,str(ROOT/'scripts/offline_docs.py'),'read','UserForm.Show','--json'],
                                    cwd=d,capture_output=True,text=True,encoding='utf-8')
        self.assertEqual(result.returncode,0,result.stderr)
        self.assertIn('vbModeless',json.loads(result.stdout)['content'])
    def test_cli_reports_missing_knowledge_instead_of_using_urls(self):
        with tempfile.TemporaryDirectory() as d:
            result = subprocess.run([sys.executable,str(ROOT/'scripts/offline_docs.py'),'--kb',d,'search','Range.Find'],
                                    capture_output=True,text=True,encoding='utf-8')
        self.assertEqual(result.returncode,2)
        self.assertIn('No online or summary-only fallback',result.stderr)
    def test_licenses_and_third_party_notices_preserved(self):
        for path in ('LICENSE','LICENSE-CODE','ThirdPartyNotices'):
            text, info = self.kb.read_file(path)
            self.assertGreater(len(text),500)
            self.assertEqual(info['sha256'],self.kb.manifest['licenses'][path]['sha256'])


class ImportSafetyTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(); self.root = Path(self.tmp.name)
    def tearDown(self):
        self.tmp.cleanup()
    def fixture(self):
        repo = self.root/'repo'; (repo/'api').mkdir(parents=True)
        for name in ('LICENSE','LICENSE-CODE','ThirdPartyNotices','docfx.json'):
            (repo/name).write_text('SYNTHETIC TEST FIXTURE\n',encoding='utf-8')
        (repo/'api/Excel.Range.Find.md').write_text('---\ntitle: Range.Find method (Excel)\napi_name:\n- Excel.Range.Find\n---\n# Range.Find\n\nSynthetic body.\n## Syntax\nexpression.Find What\n',encoding='utf-8')
        return repo
    def zip_fixture(self, repo):
        path = self.root/'source.zip'
        with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as z:
            for p in repo.rglob('*'):
                if p.is_file(): z.write(p,'VBA-Docs-main/'+p.relative_to(repo).as_posix())
            z.comment=b'0'*40
        return path
    def test_directory_import_full_text_search_read_roundtrip(self):
        repo=self.fixture(); dest=self.root/'knowledge'
        before={str(p):p.read_bytes() for p in repo.rglob('*') if p.is_file()}
        m=import_repository(repo,dest)
        self.assertEqual(m['counts']['indexed_markdown'],1)
        with KnowledgeBase(dest) as kb:
            self.assertIn('Synthetic body',kb.read('Range.Find')['content'])
            self.assertEqual(kb.verify()['status'],'PASS')
        self.assertEqual(before,{str(p):p.read_bytes() for p in repo.rglob('*') if p.is_file()})
    def test_zip_import_roundtrip(self):
        path=self.zip_fixture(self.fixture()); original=path.read_bytes()
        dest=self.root/'knowledge';m=import_repository(path,dest)
        self.assertEqual(path.read_bytes(),original)
        self.assertEqual(m['provenance']['revision'],'0'*40)
        with KnowledgeBase(dest) as kb: self.assertTrue(kb.read('Range.Find')['complete_document'])
    def test_destination_overwrite_refused(self):
        repo=self.fixture();dest=self.root/'knowledge';dest.mkdir()
        with self.assertRaises(FileExistsError):import_repository(repo,dest)
    def test_destination_inside_source_refused(self):
        repo=self.fixture()
        with self.assertRaises(ValueError): import_repository(repo,repo/'new')
    def test_invalid_source_keeps_no_partial_destination(self):
        path=self.root/'bad.zip';path.write_bytes(b'no')
        dest=self.root/'knowledge'
        with self.assertRaises(ValueError):import_repository(path,dest)
        self.assertFalse(dest.exists());self.assertFalse(list(self.root.glob('.vba-kb-import-*')))
    def test_zip_traversal_refused(self):
        path=self.zip_fixture(self.fixture())
        with zipfile.ZipFile(path,'a') as z:z.writestr('VBA-Docs-main/../escape.md','bad')
        with self.assertRaises(ValueError):import_repository(path,self.root/'knowledge')
        self.assertFalse((self.root/'escape.md').exists())
    def test_zip_case_collision_refused(self):
        path=self.zip_fixture(self.fixture())
        with zipfile.ZipFile(path,'a') as z:z.writestr('VBA-Docs-main/API/excel.range.find.md','bad')
        with self.assertRaises(ValueError):import_repository(path,self.root/'knowledge')
    def test_zip_symlink_refused(self):
        path=self.zip_fixture(self.fixture())
        with zipfile.ZipFile(path,'a') as z:
            info=zipfile.ZipInfo('VBA-Docs-main/evil');info.create_system=3
            info.external_attr=(stat.S_IFLNK|0o777)<<16;z.writestr(info,'/tmp')
        with self.assertRaises(ValueError):import_repository(path,self.root/'knowledge')
    def test_directory_symlink_refused(self):
        repo=self.fixture()
        try:
            (repo/'linked.md').symlink_to(repo/'LICENSE')
        except OSError as exc:
            if os.name == 'nt':
                self.skipTest('Windows symlink privilege unavailable: ' + str(exc))
            raise
        with self.assertRaises(ValueError):import_repository(repo,self.root/'knowledge')
    def test_corrupted_document_fails_hash_check(self):
        dest=self.root/'knowledge';import_repository(self.fixture(),dest)
        (dest/'repository/api/Excel.Range.Find.md').write_text('Changed',encoding='utf-8')
        with KnowledgeBase(dest) as kb:
            with self.assertRaisesRegex(ValueError,'integrity'):kb.read('Range.Find')
            self.assertEqual(kb.verify()['status'],'FAIL')
    def test_bad_utf8_is_not_silently_replaced(self):
        repo=self.fixture();(repo/'api/bad.md').write_bytes(b'\xffbad')
        with self.assertRaises(UnicodeError):import_repository(repo,self.root/'knowledge')
        self.assertFalse((self.root/'knowledge').exists())
    def test_missing_license_rejected(self):
        repo=self.fixture();(repo/'LICENSE-CODE').unlink()
        with self.assertRaises(ValueError):import_repository(repo,self.root/'knowledge')
    def test_jsonl_points_to_source_content(self):
        dest=self.root/'knowledge';import_repository(self.fixture(),dest)
        row=json.loads((dest/'documents.jsonl').read_text(encoding='utf-8').splitlines()[0])
        self.assertTrue((dest/row['local_path']).is_file())
        self.assertIn('Synthetic body',(dest/row['local_path']).read_text())
    def test_relocation_preserves_paths(self):
        dest=self.root/'old';import_repository(self.fixture(),dest)
        moved=self.root/'different folder';dest.rename(moved)
        with KnowledgeBase(moved) as kb:self.assertTrue(kb.read('Range.Find')['complete_document'])
    def test_nonportable_paths_rejected(self):
        for path in ('../a','a/../../b','/a','a\\b','C:/test','a/CON.txt','a/name:stream','a/../b'):
            with self.subTest(path=path):
                with self.assertRaises(ValueError):safe_relative(path)
    def test_markdown_target_parentheses(self):
        self.assertEqual(markdown_targets('[Range](excel.range(object).md)'),['excel.range(object).md'])
    def test_comment_headings_do_not_become_sections(self):
        d=parse_doc('api/Excel.Test.md',b'# Test\n<!--\n## Hidden\n-->\n## Syntax\nx\n')
        self.assertEqual([s['heading'] for s in d['sections']],['Test','Syntax'])
    def test_alias_normalization(self):
        self.assertEqual(api_key('Worksheet_Change'),'worksheet.change')
    def test_unknown_legacy_encoding_rejected(self):
        with self.assertRaises(UnicodeError):parse_doc('x.md',b'\x81')
    def test_full_word_api_classified_separately(self):
        self.assertEqual(category_for('api/Word.Document.md'),'other-api-word')


if __name__ == '__main__':
    unittest.main()

"""Offline tests with synthetic fixtures and mock COM. No real Excel or network."""
from __future__ import annotations
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import MagicMock, patch
import zipfile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT/'scripts'))
import common
import search_docs
import index_repository
import inspect_vba
import inspect_workbook
import check_sources
import excel_smoke_test
import validate_skill

class TempCase(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
    def write(self, name, text):
        p=self.root/name; p.parent.mkdir(parents=True, exist_ok=True); p.write_text(text,encoding='utf-8');return p

class CommonTests(TempCase):
    def test_new_report_roundtrip(self):
        p=self.root/'sub/report.json'; common.write_new_json(p,{'a':'value'}); self.assertEqual(common.read_json(p),{'a':'value'})
    def test_existing_report_not_overwritten(self):
        p=self.write('report.json','original')
        with self.assertRaises(FileExistsError): common.write_new_json(p,{'new':True})
        self.assertEqual(p.read_text(),'original')
    def test_unicode_bom_decoding(self):
        self.assertEqual(common.decode_source('Option Explicit'.encode('utf-16'))[0],'Option Explicit')
    def test_legacy_encoding_requires_explicit_choice(self):
        raw='Za\u017c\u00f3\u0142\u0107'.encode('cp1250')
        with self.assertRaises(ValueError): common.decode_source(raw)
        self.assertEqual(common.decode_source(raw,'cp1250')[0],'Za\u017c\u00f3\u0142\u0107')
    def test_output_preflight_rejects_existing_path(self):
        p=self.write('old.json','original')
        with self.assertRaises(FileExistsError):common.validate_new_output(p)
        common.validate_new_output(None)
        common.validate_new_output(self.root/'new.json')
    def test_hash_is_stable(self):
        p=self.write('source.bas','Option Explicit\n'); a=common.sha256_file(p)
        self.assertEqual(a,common.sha256_file(p));self.assertEqual(len(a),64)

class SearchTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls): cls.docs=common.read_json(ROOT/'metadata/source-manifest.json')['sources']
    def test_exact_owner_prefix(self):
        self.assertEqual([d['id'] for d in search_docs.search('Excel.Range.Find',self.docs,exact=True)],['MS-FIND'])
    def test_exact_source_id(self):
        self.assertEqual(search_docs.search('MS-EVENTS',self.docs,exact=True)[0]['id'],'MS-EVENTS')
    def test_broad_userform_search(self):
        self.assertTrue(any(d['id']=='MS-QUERYCLOSE' for d in search_docs.search('UserForm QueryClose',self.docs,limit=10)))
    def test_accent_normalization(self):
        self.assertEqual(search_docs.normalize('B\u0142\u0105d \u017c\u00f3\u0142\u0107'),'blad zolc')
    def test_fictional_api_is_missing_not_nonexistent(self):
        self.assertEqual(search_docs.search('Range.SetFancyGradient',self.docs,exact=True),[])
    def test_limit_and_category(self):
        result=search_docs.search('error',self.docs,limit=2,category='errors')
        self.assertLessEqual(len(result),2);self.assertTrue(result)
        self.assertTrue(all(d['category']=='errors' for d in result))
    def test_cli_exact_lookup(self):
        p=subprocess.run([sys.executable,str(ROOT/'scripts/search_docs.py'),'Range.Find','--exact','--json'],capture_output=True,text=True)
        self.assertEqual(p.returncode,0,p.stderr);self.assertEqual(json.loads(p.stdout)['results'][0]['id'],'MS-FIND')
    def test_cli_missing_member_warning(self):
        p=subprocess.run([sys.executable,str(ROOT/'scripts/search_docs.py'),'Range.SetFancyGradient','--exact','--json'],capture_output=True,text=True)
        self.assertEqual(p.returncode,1);r=json.loads(p.stdout)
        self.assertEqual(r['status'],'not_found_in_loaded_index');self.assertIn('not proof',r['warning'])

class RepositoryTests(TempCase):
    def fixture(self):
        self.write('repo/docfx.json','{}');self.write('repo/LICENSE','SYNTHETIC TEST LICENSE')
        self.write('repo/api/Excel.Range.Find.md','---\ntitle: Range.Find method (Excel)\nms.date: 2019-08-14\n---\n# Range.Find\n\nFinds a cell.\n')
        self.write('repo/Language/Concepts/Forms/test.md','# Test form\n\nForm guidance.')
        self.write('repo/excel/Miscellaneous/test.md','# Historical Excel note\n\nNot discarded because of NOINDEX.')
        self.write('repo/word/other.md','# Word\n\nOther Office.')
        self.write('repo/includes/note.md','Support content')
        return self.root/'repo'
    def test_scope_and_counts_on_fixture(self):
        r=index_repository.inventory(self.fixture())
        self.assertEqual(r['repository_total_markdown_files'],5);self.assertEqual(r['indexed_documents'],3)
        self.assertIsNone(r['repository_revision']);self.assertEqual(r['inventory_status'],'local_checkout_scanned')
    def test_include_other_scope(self):
        r=index_repository.inventory(self.fixture(),True);self.assertEqual(r['indexed_documents'],5)
    def test_document_metadata_and_inferred_url(self):
        r=index_repository.inventory(self.fixture());d=next(d for d in r['documents'] if 'Find.md' in d['repo_path'])
        self.assertEqual(d['ms_date'],'2019-08-14');self.assertEqual(d['url'],'https://learn.microsoft.com/en-us/office/vba/api/excel.range.find')
        self.assertEqual(d['url_status'],'inferred_from_path_not_network_checked');self.assertEqual(len(d['sha256']),64)
    def test_no_source_mutation(self):
        root=self.fixture();before={str(p):p.read_bytes() for p in root.rglob('*') if p.is_file()}
        index_repository.inventory(root);self.assertEqual(before,{str(p):p.read_bytes() for p in root.rglob('*') if p.is_file()})
    def test_missing_checkout_sentinel_rejected(self):
        with self.assertRaises(ValueError): index_repository.inventory(self.root)
    def test_invalid_encoding_recorded_as_problem(self):
        root=self.fixture();(root/'api/Excel.Bad.md').write_bytes(b'\xffbad')
        r=index_repository.inventory(root);self.assertEqual(r['inventory_status'],'local_checkout_partial');self.assertEqual(len(r['problems']),1)
    def test_multiline_frontmatter_falls_back_to_heading(self):
        d=index_repository.parse_document('api/Excel.X.md',b'---\ntitle: >\n  long title\n---\n# Actual heading\n\nBody')
        self.assertEqual(d['title'],'Actual heading')
    def test_cli_refuses_report_inside_checkout(self):
        root=self.fixture();p=subprocess.run([sys.executable,str(ROOT/'scripts/index_repository.py'),'--repo',str(root),'--out',str(root/'out.json')],capture_output=True,text=True)
        self.assertEqual(p.returncode,2);self.assertFalse((root/'out.json').exists())

class VbaInspectionTests(TempCase):
    def test_ignores_keywords_in_strings_comments(self):
        s='Option Explicit\nDebug.Print "On Error Resume Next and End"\n\' Kill "file"\nRem Application.EnableEvents = False\n'
        self.assertEqual(inspect_vba.inspect_text(s),[])
    def test_escaped_quotes_and_apostrophe_in_string(self):
        s='Debug.Print "He said ""End"" and don\'t": DoEvents'
        clean=inspect_vba.strip_strings_comments(s);self.assertIn('DoEvents',clean);self.assertNotIn('End',clean)
    def test_end_statement_not_end_sub(self):
        findings=inspect_vba.inspect_text('Option Explicit\nSub Demo()\nEnd\nEnd Sub')
        self.assertEqual([f['line'] for f in findings if f['rule']=='HARD-END'],[3])
    def test_unqualified_vs_qualified_range(self):
        f=inspect_vba.inspect_text('Option Explicit\nws.Range("A1").Value = 1\nRange("B1").Value = 2')
        self.assertEqual([d['line'] for d in f if d['rule']=='UNQUALIFIED-RANGE'],[3])
    def test_risky_calls_are_candidates_not_confirmed(self):
        f=inspect_vba.inspect_text('Option Explicit\nOn Error Resume Next\napp.Quit\n')
        self.assertEqual({d['rule'] for d in f},{'ERROR-SUPPRESSION','EXCEL-QUIT'})
        self.assertTrue(all(d['confidence']=='heuristic_candidate_not_confirmed' for d in f))
    def test_form_companion_presence_only(self):
        p=self.write('Form1.frm','Option Explicit\nOleObjectBlob = "Form1.frx":0000\n')
        self.assertFalse(inspect_vba.inspect_file(p)['form_resources'][0]['exists'])
        self.write('Form1.frx','synthetic fixture not real designer binary')
        self.assertTrue(inspect_vba.inspect_file(p)['form_resources'][0]['exists'])
    def test_vba_file_unchanged(self):
        p=self.write('test.bas','Option Explicit\n');before=p.read_bytes();inspect_vba.inspect_file(p);self.assertEqual(p.read_bytes(),before)
    def test_non_source_extension_rejected(self):
        p=self.write('macro.xlsm','not actual workbook')
        with self.assertRaises(ValueError): inspect_vba.inspect_file(p)
    def test_missing_option_explicit(self):
        self.assertIn('OPTION-EXPLICIT',{f['rule'] for f in inspect_vba.inspect_text('Sub X()\nEnd Sub')})

class WorkbookInspectionTests(TempCase):
    def package(self,names):
        p=self.root/'fixture.xlsm'
        with zipfile.ZipFile(p,'w') as z:
            for n in names:z.writestr(n,'synthetic metadata test; not a valid Excel workbook')
        return p
    def test_part_classification_and_no_mutation(self):
        p=self.package(['[Content_Types].xml','xl/workbook.xml','xl/vbaProject.bin','xl/macrosheets/sheet1.xml','xl/connections.xml','customUI/customUI.xml'])
        before=p.read_bytes();r=inspect_workbook.inspect_package(p)
        self.assertEqual(r['status'],'METADATA_ONLY');self.assertTrue(r['looks_like_excel_package'])
        self.assertTrue(r['part_groups']['vba_project']);self.assertTrue(r['part_groups']['xlm_macrosheets'])
        self.assertEqual(p.read_bytes(),before)
    def test_no_macro_not_safe_verdict(self):
        r=inspect_workbook.inspect_package(self.package(['[Content_Types].xml','xl/workbook.xml']))
        self.assertFalse(r['part_groups']['vba_project']);self.assertIn('does not establish safety',r['warning'])
    def test_traversal_rejected(self):
        with self.assertRaises(ValueError):inspect_workbook.inspect_package(self.package(['../outside.bin']))
    def test_case_collision_rejected(self):
        with self.assertRaises(ValueError):inspect_workbook.inspect_package(self.package(['xl/A.xml','xl/a.xml']))
    def test_unsupported_container(self):
        p=self.write('legacy.xls','not a zip');self.assertEqual(inspect_workbook.inspect_package(p)['status'],'UNSUPPORTED_OR_ENCRYPTED')
    def test_archive_entry_cap(self):
        p=self.package(['one','two'])
        with patch.object(inspect_workbook,'MAX_ENTRIES',1):
            with self.assertRaises(ValueError):inspect_workbook.inspect_package(p)

class NetworkPermissionTests(unittest.TestCase):
    def test_default_makes_no_network_call(self):
        opener=MagicMock();r=check_sources.run_checks([{'id':'A','url':'https://learn.microsoft.com/example'}],False,opener=opener)
        self.assertEqual(r['status'],'DRY_RUN_NO_NETWORK');opener.open.assert_not_called()
    def test_forbids_untrusted_or_insecure_url(self):
        for url in ['http://learn.microsoft.com/x','https://evil.example/x','https://user:pass@github.com/x','https://github.com/x?secret=a']:
            with self.subTest(url=url):
                with self.assertRaises(ValueError):check_sources.validate_url(url)
    def test_mocked_reachability_not_content_validation(self):
        opener=MagicMock();response=opener.open.return_value.__enter__.return_value
        response.status=200;response.geturl.return_value='https://learn.microsoft.com/example'
        r=check_sources.check_one('https://learn.microsoft.com/example',opener)
        self.assertEqual(r['status'],'reachable');self.assertIn('not semantic',r['meaning'])
        response.read.assert_called_once_with(1024)
    def test_mocked_network_failure_reported(self):
        opener=MagicMock();opener.open.side_effect=OSError('synthetic offline error')
        r=check_sources.check_one('https://github.com/example',opener);self.assertEqual(r['status'],'failed')
    def test_redirect_allowlist_enforced(self):
        with self.assertRaises(ValueError):check_sources.RestrictedRedirect().redirect_request(None,None,302,'',{},'https://evil.example')

class MockComTests(TempCase):
    def mock_excel(self,value=8):
        app=MagicMock();book=app.Workbooks.Add.return_value;sheet=book.Worksheets.return_value
        initial=MagicMock();target=MagicMock();target.Value2=value;source=MagicMock()
        sheet.Range.side_effect=[initial,target,source]
        sheet.ChartObjects.return_value.Add.return_value.Chart.SeriesCollection.return_value.Count=1
        book.SaveAs.side_effect=lambda name,**kw:Path(name).write_text('MOCK FILE, NOT A REAL XLSX')
        return app,book
    def test_mocked_success_closes_owned_resources(self):
        app,book=self.mock_excel();r=excel_smoke_test.execute_smoke(lambda:app,self.root)
        self.assertEqual(r['status'],'PASS');self.assertEqual(len(r['tests']),3)
        book.Close.assert_called_once_with(SaveChanges=False);app.Quit.assert_called_once()
    def test_primary_error_preserved_and_quit_attempted(self):
        app,book=self.mock_excel(value=99);book.Close.side_effect=RuntimeError('cleanup close failed')
        r=excel_smoke_test.execute_smoke(lambda:app,self.root)
        self.assertEqual(r['status'],'FAIL');self.assertEqual(r['primary_error']['type'],'AssertionError')
        self.assertIn('99',r['primary_error']['message']);self.assertEqual(r['cleanup_errors'][0]['operation'],'Close');app.Quit.assert_called_once()
    def test_quit_failure_is_not_success(self):
        app,book=self.mock_excel();app.Quit.side_effect=RuntimeError('quit failure')
        r=excel_smoke_test.execute_smoke(lambda:app,self.root);self.assertEqual(r['status'],'FAIL');self.assertIsNone(r['primary_error'])
        self.assertEqual(r['cleanup_errors'][0]['operation'],'Quit')
    def test_existing_output_refused_before_start(self):
        self.write('smoke-result.xlsx','original');factory=MagicMock()
        with self.assertRaises(FileExistsError):excel_smoke_test.execute_smoke(factory,self.root)
        factory.assert_not_called()
    def test_factory_failure_reported(self):
        factory=MagicMock(side_effect=RuntimeError('COM unavailable'))
        r=excel_smoke_test.execute_smoke(factory,self.root);self.assertEqual(r['status'],'FAIL');self.assertEqual(r['cleanup_errors'],[])
    def test_dependency_probe_no_excel(self):
        with patch.object(excel_smoke_test.platform,'system',return_value='Linux'):
            r=excel_smoke_test.dependencies();self.assertFalse(r['excel_started']);self.assertFalse(r['windows'])
    def test_cli_existing_report_refused_before_dependency_probe(self):
        p=self.write('exists.json','original')
        with patch.object(sys,'argv',['excel_smoke_test','--run-local-excel','--out',str(p)]), patch.object(excel_smoke_test,'dependencies') as dep, patch('sys.stderr'):
            self.assertEqual(excel_smoke_test.main(),2)
            dep.assert_not_called()
    def test_blank_sheet_template_not_user_default(self):
        app,book=self.mock_excel()
        excel_smoke_test.execute_smoke(lambda:app,self.root)
        app.Workbooks.Add.assert_called_once_with(Template=-4167)
    def test_cli_default_is_not_run(self):
        p=subprocess.run([sys.executable,str(ROOT/'scripts/excel_smoke_test.py')],capture_output=True,text=True)
        self.assertEqual(p.returncode,0,p.stderr);self.assertEqual(json.loads(p.stdout)['status'],'NOT_RUN')

class PackageTests(TempCase):
    def test_frontmatter_valid(self):
        r=validate_skill.parse_frontmatter((ROOT/'SKILL.md').read_text());self.assertEqual(r['name'],'excel-vba-expert')
    def test_frontmatter_duplicate_key_rejected(self):
        with self.assertRaises(ValueError):validate_skill.parse_frontmatter('---\nname: x\nname: y\ndescription: "text"\n---\n')
    def test_frontmatter_missing_field_rejected(self):
        with self.assertRaises(ValueError):validate_skill.parse_frontmatter('---\nname: x\n---\n')
    def test_invalid_relative_link_detected(self):
        p=self.write('test.md','[bad](missing.md)')
        self.assertTrue(validate_skill.relative_link_errors(self.root,p,p.read_text()))
    def test_escape_link_rejected(self):
        p=self.write('test.md','[bad](../../outside.md)')
        self.assertIn('escapes',validate_skill.relative_link_errors(self.root,p,p.read_text())[0])
    def test_actual_package_validation(self):
        r=validate_skill.validate(ROOT);self.assertEqual(r['status'],'PASS',r['errors'])
    def test_scenario_resource_contracts(self):
        cases=common.read_json(ROOT/'tests/scenarios.json')['scenarios'];self.assertEqual(len(cases),10)
        for case in cases:
            with self.subTest(case=case['id']):
                self.assertEqual(case['execution_status'],'NOT_RUN_WITH_LIVE_AGENT')
                self.assertTrue(case['must_demonstrate']);self.assertTrue(case['must_not_do'])
                for rel in case['required_resources']:self.assertTrue((ROOT/rel).is_file())
    def test_all_python_sources_compile_as_python_only(self):
        for path in list((ROOT/'scripts').glob('*.py'))+list((ROOT/'tests').glob('test_*.py')):
            with self.subTest(file=path.name):compile(path.read_text(encoding='utf-8'),path.name,'exec')

if __name__=='__main__':unittest.main()

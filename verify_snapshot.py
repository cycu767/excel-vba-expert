"""Verify all local source hashes and five API reads; optionally compare original ZIP.

No network or Excel. --out must be a new path. The report's source-section facts
come from the supplied files, not a reconstructed or corrected API specification.
"""
from __future__ import annotations
import argparse
import hashlib
import json
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'scripts'))
from offline_kb import KnowledgeBase, sha256_file

APIS=('Range.Find','Workbooks.Open','Worksheet.Change','UserForm.Show','ChartObjects.Add')


def verify_snapshot(archive: Path | None = None, kb_root: Path | None = None) -> dict:
    with KnowledgeBase(kb_root) as kb:
        verification=kb.verify()
        apis=[]
        for api in APIS:
            r=kb.read(api)
            d=r['document']
            headings={s['heading'].casefold():s for s in d['sections']}
            areas={}
            for name, variants in [('syntax',('syntax',)),('parameters',('parameters',)),
                                   ('remarks',('remarks',)),('examples',('example','examples'))]:
                match=next((headings[k] for k in variants if k in headings),None)
                areas[name]={'present':bool(match),'location':match}
            if api=='UserForm.Show':
                areas['parameters']={'present':True,'location':headings['syntax'],
                                     'note':'Argument-parts table in Syntax; no Parameters heading in source'}
            note=None
            if api=='ChartObjects.Add':
                note='Syntax names Left, Top, Width, Height; source parameter table contains only Left and Width; no Remarks heading. No missing prose was invented.'
            apis.append({'api':api,'path':d['path'],'sha256':d['sha256'],'lines':d['line_count'],
                         'full_document_read':r['complete_document'],
                         'description_present':bool(d['description_excerpt']),
                         'areas':areas,'local_includes':[x.get('path',x.get('status')) for x in kb.includes(api)],
                         'source_gap':note})
        archive_check={'status':'NOT_RUN','reason':'No original --archive supplied'}
        if archive:
            rows={r['path']:dict(r) for r in kb.db.execute('SELECT * FROM files')}
            mismatches=[];count=0;seen=set()
            with zipfile.ZipFile(archive) as z:
                prefix=kb.manifest['provenance'].get('archive_root','')
                prefix=prefix+'/' if prefix else ''
                for info in z.infolist():
                    if info.is_dir():continue
                    if not info.filename.startswith(prefix):
                        mismatches.append(info.filename);continue
                    rel=info.filename[len(prefix):];seen.add(rel)
                    raw=z.read(info);count+=1
                    if rel not in rows or hashlib.sha256(raw).hexdigest()!=rows[rel]['sha256']:
                        mismatches.append(rel)
            missing=sorted(set(rows)-seen)
            archive_check={'status':'PASS' if not mismatches and not missing and count==len(rows) else 'FAIL',
                           'archive_sha256':sha256_file(archive),'compared_source_files':count,
                           'mismatched_files':mismatches,'missing_from_archive':missing,
                           'evidence':'Every archive entry CRC-read and SHA-256 compared to the verified physical source inventory'}
        excel_names=sum(len(json.loads(r[0])) for r in kb.db.execute("SELECT api_names FROM documents WHERE category='excel-api'"))
        return {'created_at_utc':datetime.now(timezone.utc).isoformat(),
                'status':'PASS' if verification['status']=='PASS' and archive_check['status']!='FAIL' else 'FAIL',
                'source_counts':kb.manifest['counts'],'provenance':kb.manifest['provenance'],
                'local_integrity':verification,'original_archive_comparison':archive_check,
                'excel_api_name_entries_exercised_by_unit_tests':excel_names,
                'api_reads':apis,
                'not_claimed':['Live Qwen/Copilot compliance','VBA compile/runtime','Current upstream equivalence','Semantic review of all source pages']}


def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--archive',type=Path)
    p.add_argument('--kb',type=Path)
    p.add_argument('--out',type=Path)
    a=p.parse_args()
    if a.out and a.out.exists():p.error('Output already exists')
    try:
        r=verify_snapshot(a.archive,a.kb)
        if a.out:
            with a.out.open('x',encoding='utf-8') as f:json.dump(r,f,ensure_ascii=False,indent=2);f.write('\n')
        print(json.dumps(r,ensure_ascii=False,indent=2))
        return 0 if r['status']=='PASS' else 1
    except (OSError,ValueError,zipfile.BadZipFile) as e:
        print('ERROR: '+str(e),file=sys.stderr);return 2

if __name__=='__main__':raise SystemExit(main())

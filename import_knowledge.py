"""Import a complete LOCAL VBA-Docs directory or ZIP into a NEW knowledge folder.

Never downloads, executes repository content, changes the input, or overwrites an
existing destination. Preserve source bytes and every non-Git file. Python 3.10+.
"""
from __future__ import annotations

import argparse
import collections
import hashlib
import json
import os
import re
import shutil
import sqlite3
import stat
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import unquote, urlsplit
from offline_kb import (ROOT, SCHEMA_VERSION, MAX_FILE_BYTES, REQUIRED_SOURCE_FILES,
                        api_key, json_write, parse_doc, safe_relative, sha256_file)

MAX_FILES = 100000
MAX_TOTAL_BYTES = 512 * 1024 * 1024


def _check_limits(entries: list[tuple[str, int]]) -> None:
    if len(entries) > MAX_FILES or sum(size for _, size in entries) > MAX_TOTAL_BYTES:
        raise ValueError('Source exceeds safety limits (100000 files / 512 MiB)')
    seen = set()
    for rel, size in entries:
        safe_relative(rel)
        key = rel.casefold()
        if key in seen:
            raise ValueError('Duplicate or case-colliding path: ' + rel)
        seen.add(key)
        if size > MAX_FILE_BYTES:
            raise ValueError('File exceeds 32 MiB limit: ' + rel)
    required = {p.casefold() for p in REQUIRED_SOURCE_FILES}
    if not required.issubset(seen) or not any(p.startswith('api/') for p in seen):
        raise ValueError('Expected VBA-Docs root with api/, docfx.json, LICENSE, LICENSE-CODE and ThirdPartyNotices')


def copy_zip(source: Path, destination: Path) -> tuple[list[dict], dict]:
    inventory = []
    with zipfile.ZipFile(source) as z:
        infos = [i for i in z.infolist() if not i.is_dir()]
        all_names = {i.filename for i in infos}
        roots = []
        for name in all_names:
            if name.endswith('docfx.json'):
                prefix = name[:-len('docfx.json')]
                if all(prefix + s in all_names for s in REQUIRED_SOURCE_FILES):
                    roots.append(prefix)
        if len(roots) != 1:
            raise ValueError('ZIP must contain exactly one complete VBA-Docs root')
        prefix = roots[0]
        entries = []
        for i in infos:
            if not i.filename.startswith(prefix):
                raise ValueError('File outside the repository root: ' + i.filename)
            mode = (i.external_attr >> 16) & 0xFFFF
            if stat.S_ISLNK(mode) or (stat.S_IFMT(mode) not in (0, stat.S_IFREG)):
                raise ValueError('Symlink/special ZIP entry rejected: ' + i.filename)
            if i.flag_bits & 1:
                raise ValueError('Encrypted source ZIP unsupported')
            rel = safe_relative(i.filename[len(prefix):])
            if '.git' in Path(rel).parts:
                raise ValueError('Use a GitHub source ZIP without .git internals, or import the local checkout')
            entries.append((rel, i.file_size))
        _check_limits(entries)
        for i, (rel, size) in zip(infos, entries):
            path = destination / rel
            path.parent.mkdir(parents=True, exist_ok=True)
            raw = z.read(i)  # zipfile checks the stored CRC; never execute source files.
            if len(raw) != size:
                raise ValueError('ZIP size mismatch: ' + rel)
            with path.open('xb') as f:
                f.write(raw)
            inventory.append({'path': rel, 'bytes': size, 'sha256': hashlib.sha256(raw).hexdigest()})
        comment = z.comment.decode('ascii', errors='replace').strip()
    revision = comment if re.fullmatch(r'[0-9a-fA-F]{40}', comment) else None
    return inventory, {'source_type': 'user_supplied_zip', 'source_name': source.name,
                       'archive_sha256': sha256_file(source), 'archive_bytes': source.stat().st_size,
                       'archive_comment': comment, 'revision': revision,
                       'revision_evidence': 'ZIP comment; preserved, not independently authenticated',
                       'archive_root': prefix.rstrip('/'), 'excluded_git_metadata': False}


def copy_directory(source: Path, destination: Path) -> tuple[list[dict], dict]:
    paths = []
    excluded_git = False
    for directory, dirs, files in os.walk(source, followlinks=False):
        base = Path(directory)
        for name in list(dirs):
            child = base / name
            if child.is_symlink():
                raise ValueError('Source symlink rejected: ' + str(child.relative_to(source)))
            if name == '.git':
                dirs.remove(name); excluded_git = True
        dirs.sort()
        for name in sorted(files):
            p = base / name
            if name == '.git':
                excluded_git = True; continue
            if p.is_symlink() or not stat.S_ISREG(p.stat().st_mode):
                raise ValueError('Non-regular source file: ' + str(p.relative_to(source)))
            paths.append((p, safe_relative(p.relative_to(source).as_posix()), p.stat().st_size))
    _check_limits([(rel, size) for _, rel, size in paths])
    inventory = []
    for p, rel, size in paths:
        raw = p.read_bytes()
        if len(raw) != size:
            raise ValueError('Source changed during import: ' + rel)
        out = destination / rel
        out.parent.mkdir(parents=True, exist_ok=True)
        with out.open('xb') as f:
            f.write(raw)
        inventory.append({'path': rel, 'bytes': size, 'sha256': hashlib.sha256(raw).hexdigest()})
    return inventory, {'source_type': 'user_supplied_directory', 'source_name': source.name,
                       'archive_sha256': None, 'revision': None,
                       'revision_evidence': 'No authenticated revision; file hashes identify the imported tree',
                       'excluded_git_metadata': excluded_git}


def create_index(stage: Path, inventory: list[dict]) -> dict:
    db = sqlite3.connect(stage / 'index.sqlite')
    counts = collections.Counter()
    kinds = collections.Counter()
    alias_extra = json.loads((ROOT / 'metadata/api-aliases.json').read_text(encoding='utf-8'))['aliases']
    curated = json.loads((ROOT / 'metadata/source-manifest.json').read_text(encoding='utf-8'))['sources']
    paths_lower = {d['path'].casefold(): d['path'] for d in inventory}
    curated_mapping, curated_unmapped = {}, []
    aliases_by_path = collections.defaultdict(set)
    alias_missing = []
    for alias, path in alias_extra.items():
        actual = paths_lower.get(path.casefold())
        if actual:
            aliases_by_path[actual].add(api_key(alias))
        else:
            alias_missing.append({'alias': alias, 'path': path})
    for rec in curated:
        candidates = [rec.get('repo_path')]
        url = urlsplit(rec['url'])
        if url.netloc.casefold() == 'learn.microsoft.com' and '/office/vba/' in url.path.casefold():
            tail = re.split(r'/office/vba/', unquote(url.path), flags=re.I)[1]
            candidates.extend([tail + '.md', tail + '/index.md'])
        actual = next((paths_lower[p.casefold()] for p in candidates if p and p.casefold() in paths_lower), None)
        if actual and actual.casefold().endswith('.md'):
            curated_mapping[rec['id']] = actual
            aliases_by_path[actual].add(api_key(rec['id']))
        else:
            curated_unmapped.append(rec['id'])
    try:
        db.executescript('''
            PRAGMA journal_mode=DELETE;
            CREATE TABLE files(path TEXT PRIMARY KEY COLLATE NOCASE, sha256 TEXT NOT NULL, bytes INTEGER NOT NULL);
            CREATE TABLE documents(
                id INTEGER PRIMARY KEY, path TEXT UNIQUE COLLATE NOCASE, title TEXT, category TEXT,
                kind TEXT, api_names TEXT, ms_date TEXT, sha256 TEXT, bytes INTEGER,
                encoding TEXT, line_count INTEGER, sections TEXT, description_excerpt TEXT);
            CREATE TABLE aliases(alias TEXT, doc_id INTEGER REFERENCES documents(id));
            CREATE INDEX aliases_lookup ON aliases(alias);
            CREATE INDEX document_scope ON documents(category,kind);
            CREATE VIRTUAL TABLE document_fts USING fts5(title, aliases, body, content='',
                tokenize='unicode61 remove_diacritics 2');
        ''')
        db.executemany('INSERT INTO files VALUES(?,?,?)', [(d['path'], d['sha256'], d['bytes']) for d in inventory])
        with (stage / 'documents.jsonl').open('x', encoding='utf-8', newline='\n') as docout:
            for item in sorted(inventory, key=lambda d: d['path']):
                path = item['path']
                if not path.casefold().endswith('.md'):
                    continue
                raw = (stage / 'repository' / path).read_bytes()
                doc = parse_doc(path, raw)
                doc['aliases'] = sorted(set(doc['aliases']) | aliases_by_path[path])
                text = doc.pop('body_for_index')
                keys = ('path','title','category','kind','api_names','ms_date','sha256','bytes','encoding','line_count','sections','description_excerpt')
                vals = [json.dumps(doc[k], ensure_ascii=False) if k in ('api_names','sections') else doc[k] for k in keys]
                cur = db.execute('INSERT INTO documents(' + ','.join(keys) + ') VALUES(' + ','.join('?' for _ in keys) + ')', vals)
                ident = cur.lastrowid
                db.executemany('INSERT INTO aliases VALUES(?,?)', [(a, ident) for a in doc['aliases']])
                db.execute('INSERT INTO document_fts(rowid,title,aliases,body) VALUES(?,?,?,?)',
                           (ident, doc['title'], ' '.join(doc['aliases']), text))
                doc['local_path'] = 'repository/' + path
                docout.write(json.dumps(doc, ensure_ascii=False, separators=(',', ':')) + '\n')
                counts[doc['category']] += 1
                kinds[doc['kind']] += 1
        db.commit()
        db.execute("INSERT INTO document_fts(document_fts) VALUES('optimize')")
        db.commit()
        db.execute('VACUUM')
        quick = db.execute('PRAGMA integrity_check').fetchone()[0]
        if quick != 'ok':
            raise ValueError('SQLite integrity check failed: ' + quick)
        indexed = db.execute('SELECT COUNT(*) FROM documents').fetchone()[0]
        alias_count = db.execute('SELECT COUNT(*) FROM aliases').fetchone()[0]
    finally:
        db.close()
    with (stage / 'files.jsonl').open('x', encoding='utf-8', newline='\n') as f:
        for item in sorted(inventory, key=lambda d: d['path']):
            f.write(json.dumps(item, ensure_ascii=False, separators=(',', ':')) + '\n')
    json_write(stage / 'curated-crosswalk.json', {'mapped': curated_mapping, 'not_in_this_repo_or_unresolved': curated_unmapped,
                                                'custom_alias_missing_source': alias_missing})
    return {'indexed_markdown': indexed, 'aliases': alias_count, 'by_category': dict(sorted(counts.items())),
            'by_document_kind_heuristic': dict(sorted(kinds.items())),
            'curated_source_ids_mapped': len(curated_mapping), 'curated_source_ids_unmapped': len(curated_unmapped),
            'custom_alias_missing_source': alias_missing}


def import_repository(source: Path, destination: Path) -> dict:
    source = source.expanduser().resolve(strict=True)
    destination = destination.expanduser().absolute()
    if destination.exists() or destination.is_symlink():
        raise FileExistsError('Destination exists; choose a new folder, then switch explicitly: ' + str(destination))
    if source.is_dir() and destination.resolve().is_relative_to(source):
        raise ValueError('Destination must be outside the source repository')
    destination.parent.mkdir(parents=True, exist_ok=True)
    stage = Path(tempfile.mkdtemp(prefix='.vba-kb-import-', dir=destination.parent))
    try:
        repo = stage / 'repository'
        repo.mkdir()
        if source.is_dir():
            inventory, provenance = copy_directory(source, repo)
        elif zipfile.is_zipfile(source):
            inventory, provenance = copy_zip(source, repo)
        else:
            raise ValueError('Provide a local directory or ZIP archive, not a URL')
        index_counts = create_index(stage, inventory)
        md_count = sum(d['path'].casefold().endswith('.md') for d in inventory)
        if md_count != index_counts['indexed_markdown']:
            raise ValueError('Imported/indexed Markdown counts differ')
        manifest = {'schema_version': SCHEMA_VERSION, 'knowledge_name': 'excel-vba-expert-knowledge',
                    'created_at_utc': datetime.now(timezone.utc).isoformat(),
                    'repository': 'MicrosoftDocs/VBA-Docs', 'source_origin': 'https://github.com/MicrosoftDocs/VBA-Docs',
                    'provenance': provenance, 'status': 'COMPLETE_FOR_SUPPLIED_SNAPSHOT',
                    'counts': {'source_files': len(inventory), 'source_bytes': sum(d['bytes'] for d in inventory),
                               'source_markdown': md_count, 'local_markdown': md_count, **index_counts},
                    'index_sha256': sha256_file(stage / 'index.sqlite'),
                    'documents_jsonl_sha256': sha256_file(stage / 'documents.jsonl'),
                    'files_jsonl_sha256': sha256_file(stage / 'files.jsonl'),
                    'licenses': {p: {'local_path': 'repository/' + p, 'sha256': sha256_file(repo / p)}
                                 for p in ('LICENSE','LICENSE-CODE','ThirdPartyNotices')},
                    'source_changes': 'None; repository file bytes copied unchanged. SQLite, JSONL and aliases are separate authored derivatives.',
                    'semantic_review': 'Not exhaustive; original curated 19 modules remain a separate engineering layer.',
                    'limitations': ['Coverage is measured against supplied files, not an independently authenticated current upstream tree.',
                                    'Includes resolve locally; external linked pages and external assets are not downloaded.',
                                    'This is not a DocFX/HTML renderer. Original Markdown can contain publishing directives and source comments.',
                                    'Full-source availability does not prove correctness of Microsoft examples or execution in Excel.']}
        json_write(stage / 'manifest.json', manifest)
        attribution = '''# Microsoft VBA documentation: offline source snapshot

Copyright Microsoft Corporation and contributors. Original repository:
https://github.com/MicrosoftDocs/VBA-Docs

Documentation/content: Creative Commons Attribution 4.0 International, as recorded
in repository/ThirdPartyNotices and repository/LICENSE. Code: MIT, as recorded in
repository/LICENSE-CODE. Preserve all three files, source-level attributions and
notices when copying or redistributing. No Microsoft endorsement is claimed.
License link: https://creativecommons.org/licenses/by/4.0/legalcode

Source files in repository/ are byte-for-byte unchanged. Generated index.sqlite,
documents.jsonl, files.jsonl, lookup aliases and manifests are separate indexing
artifacts, not corrections to or replacement of the original documentation.
Do not treat HTML-commented samples as current recommended code. Read original
remarks and apply the skill's reviewed safety rules before adapting examples.

Use the sibling excel-vba-expert/scripts/offline_docs.py to search/read locally.
No internet, API key, embeddings server, Excel or third-party Python packages are
required. Python 3.10+ with sqlite3 FTS5 is required. manifest.json contains exact
counts and snapshot provenance. documents.jsonl contains local document paths;
files.jsonl contains hashes of every source file.
'''
        (stage / 'ATTRIBUTION.md').write_text(attribution, encoding='utf-8')
        if destination.exists() or destination.is_symlink():
            raise FileExistsError('Destination appeared while importing; refusing replacement')
        stage.rename(destination)
        return manifest
    except BaseException:
        shutil.rmtree(stage, ignore_errors=True)
        raise


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('--source', required=True, type=Path)
    p.add_argument('--dest', required=True, type=Path)
    args = p.parse_args()
    try:
        manifest = import_repository(args.source, args.dest)
        print(json.dumps({'status': manifest['status'], 'destination': str(args.dest),
                          'counts': manifest['counts'], 'provenance': manifest['provenance']}, indent=2))
        return 0
    except (OSError, ValueError, sqlite3.Error, zipfile.BadZipFile, UnicodeError) as exc:
        print(f'ERROR: {exc}', file=sys.stderr)
        return 2


if __name__ == '__main__':
    raise SystemExit(main())

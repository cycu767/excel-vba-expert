"""Local Microsoft VBA documentation storage and retrieval. No network code.

The original files are authoritative; derived SQLite/JSON indexes are disposable.
Python 3.10+ with sqlite3 FTS5 is required. Only the standard library is imported.
"""
from __future__ import annotations

import collections
import hashlib
import json
import os
import posixpath
import re
import sqlite3
import unicodedata
from pathlib import Path, PurePosixPath
from urllib.parse import unquote, urlsplit

ROOT = Path(__file__).resolve().parents[1]
SCHEMA_VERSION = 1
MAX_FILE_BYTES = 32 * 1024 * 1024
EXCEL_CATEGORIES = (
    'excel-api', 'excel-concepts', 'excel-graph', 'excel-overview', 'office-api',
    'office-concepts', 'forms-concepts', 'vba-reference', 'vba-concepts',
    'vba-how-to', 'vbe', 'overview', 'office-mac', 'support',
)
REQUIRED_SOURCE_FILES = ('LICENSE', 'LICENSE-CODE', 'ThirdPartyNotices', 'docfx.json')


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open('rb') as f:
        for block in iter(lambda: f.read(1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()


def json_write(path: Path, value: object) -> None:
    with path.open('x', encoding='utf-8', newline='\n') as f:
        json.dump(value, f, ensure_ascii=False, indent=2)
        f.write('\n')


def normalize(value: str) -> str:
    value = value.casefold().replace('\u0142', 'l')
    return ''.join(c for c in unicodedata.normalize('NFKD', value)
                   if not unicodedata.combining(c))


def api_key(value: str) -> str:
    value = normalize(value).strip().strip('`')
    value = re.sub(r'\s+(?:method|property|event|object|enumeration|statement|function)(?:\s|\().*$', '', value)
    value = re.sub(r'\s+(?:method|property|event|object|enumeration|statement|function)$', '', value)
    return re.sub(r'\s+', '', value).replace('_', '.')


def safe_relative(value: str) -> str:
    """Reject traversal, links, ADS and Windows-special path components."""
    if not value or '\\' in value or value.startswith('/'):
        raise ValueError(f'Unsafe relative path: {value!r}')
    p = PurePosixPath(value)
    parts = value.split('/')
    if any(x in ('', '.', '..') for x in parts):
        raise ValueError(f'Unsafe path component: {value!r}')
    reserved = {'CON', 'PRN', 'AUX', 'NUL'} | {f'{s}{i}' for s in ('COM', 'LPT') for i in range(1, 10)}
    for part in parts:
        if (any(ord(c) < 32 or c in '<>:"|?*' for c in part)
                or part.endswith((' ', '.')) or part.split('.')[0].upper() in reserved):
            raise ValueError(f'Nonportable/unsafe file name: {value!r}')
    if p.is_absolute():
        raise ValueError(f'Absolute path not allowed: {value!r}')
    return p.as_posix()


def contained_file(root: Path, relative: str) -> Path:
    relative = safe_relative(relative)
    root = root.resolve(strict=True)
    current = root
    for part in PurePosixPath(relative).parts:
        current = current / part
        if current.is_symlink():
            raise ValueError(f'Symlink forbidden in knowledge tree: {relative}')
    resolved = current.resolve(strict=True)
    if not resolved.is_relative_to(root) or not resolved.is_file():
        raise ValueError(f'Not a contained regular file: {relative}')
    return resolved


def decode_markdown(raw: bytes) -> tuple[str, str]:
    if raw.startswith((b'\xff\xfe', b'\xfe\xff')):
        return raw.decode('utf-16', errors='strict'), 'utf-16-bom'
    return raw.decode('utf-8-sig', errors='strict'), 'utf-8-sig'


def category_for(path: str) -> str:
    p = path.casefold()
    if p.startswith(('api/excel.', 'api/excel(')):
        return 'excel-api'
    if p.startswith('api/office.'):
        return 'office-api'
    if p.startswith('excel/') and 'graph' in p:
        return 'excel-graph'
    if p.startswith('excel/'):
        return 'excel-concepts'
    if p.startswith('api/overview/excel'):
        return 'excel-overview'
    if p.startswith('api/overview/'):
        return 'overview'
    if p.startswith('language/concepts/forms/'):
        return 'forms-concepts'
    if 'visual-basic-add-in-model' in p:
        return 'vbe'
    if p.startswith('language/how-to/'):
        return 'vba-how-to'
    if p.startswith('language/reference/'):
        return 'vba-reference'
    if p.startswith('language/'):
        return 'vba-concepts'
    if p.startswith('library-reference/'):
        return 'office-concepts'
    if p.startswith('office-mac/'):
        return 'office-mac'
    if p.startswith('includes/'):
        return 'support'
    if p.startswith('f1/'):
        return 'f1-index'
    if p.startswith('api/'):
        return 'other-api-' + PurePosixPath(path).name.split('.')[0].casefold()
    if '/' not in p:
        return 'repository-meta'
    return 'other-' + p.split('/')[0]


def scalar(front: str, key: str) -> str | None:
    m = re.search(r'^' + re.escape(key) + r':[ \t]*(.*)$', front, re.M)
    if not m or m.group(1).strip() in ('', '>', '|', '>-', '|-'):
        return None
    return m.group(1).strip().strip('"\'')


def front_list(front: str, key: str) -> list[str]:
    """Read only a scalar/list of simple strings; never execute YAML tags."""
    m = re.search(r'^' + re.escape(key) + r':[ \t]*(.*)(?:\n|$)', front, re.M)
    if not m:
        return []
    result = []
    initial = m.group(1).strip()
    if initial and initial not in ('>', '|', '>-', '|-'):
        if initial.startswith('['):
            result += [s.strip().strip('"\'') for s in initial.strip('[]').split(',') if s.strip()]
        else:
            result.append(initial.strip('"\''))
    for line in front[m.end():].splitlines():
        item = re.match(r'^\s*-\s+(.+?)\s*$', line)
        if not item:
            break
        result.append(item.group(1).strip('"\''))
    return result


def visible_text(text: str) -> str:
    """Index visible prose/code, not deliberately commented-out old examples."""
    return re.sub(r'<!--[\s\S]*?-->', '', text)


def section_ranges(text: str) -> list[dict]:
    lines = text.splitlines()
    headings = []
    fence = None
    in_comment = False
    for i, line in enumerate(lines, 1):
        if in_comment:
            if '-->' in line:
                in_comment = False
            continue
        if '<!--' in line:
            if '-->' not in line.split('<!--', 1)[1]:
                in_comment = True
            continue
        f = re.match(r'^\s*(`{3,}|~{3,})', line)
        if f:
            c = f.group(1)[0]
            fence = None if fence == c else c if fence is None else fence
            continue
        if fence:
            continue
        h = re.match(r'^(#{1,6})\s+(.+?)\s*#*\s*$', line)
        if h:
            headings.append({'heading': h.group(2), 'level': len(h.group(1)), 'start_line': i})
    for pos, h in enumerate(headings):
        h['end_line'] = next((n['start_line'] - 1 for n in headings[pos+1:]
                              if n['level'] <= h['level']), len(lines))
    return headings


def parse_doc(path: str, raw: bytes) -> dict:
    text, encoding = decode_markdown(raw)
    m = re.match(r'\A---\s*\r?\n(.*?)\r?\n---\s*\r?\n', text, re.S)
    front, body = (m.group(1), text[m.end():]) if m else ('', text)
    h = re.search(r'^#\s+(.+)$', body, re.M)
    title = scalar(front, 'title') or (h.group(1).strip() if h else PurePosixPath(path).stem)
    kind_match = re.search(r'\b(method|property|event|object|enumeration|statement|function|collection)\b', title, re.I)
    kind = kind_match.group(1).lower() if kind_match else 'article'
    api_names = front_list(front, 'api_name')
    candidates = [title, PurePosixPath(path).stem, path, *api_names]
    for candidate in list(candidates):
        if candidate.casefold().startswith('excel.'):
            candidates.append(candidate[6:])
    # Enumeration values are aliases to the complete enumeration page, not invented pages.
    if kind == 'enumeration':
        for row in visible_text(body).splitlines():
            match = re.match(r'^\s*\|\s*([^|]+)\|', row)
            if match:
                constant = match.group(1).strip().strip('*_` ')
                if re.fullmatch(r'(?:xl|mso)[A-Z]\w*', constant):
                    candidates.append(constant)
    aliases = sorted({api_key(a) for a in candidates if a.strip()})
    paragraphs = re.split(r'\n\s*\n', visible_text(body))
    intro = next((x.strip() for x in paragraphs if x.strip() and not x.lstrip().startswith(('#', '[!', '<!--', '|', '```', '---'))), '')
    sections = section_ranges(text)
    return {'path': path, 'title': title, 'category': category_for(path), 'kind': kind,
            'api_names': api_names, 'aliases': aliases, 'ms_date': scalar(front, 'ms.date'),
            'sha256': hashlib.sha256(raw).hexdigest(), 'bytes': len(raw), 'encoding': encoding,
            'line_count': len(text.splitlines()), 'sections': sections,
            'description_excerpt': intro[:600], 'body_for_index': visible_text(body)}


def resolve_kb(explicit: Path | str | None = None) -> Path:
    if explicit:
        return Path(explicit).expanduser().resolve(strict=True)
    env = os.environ.get('EXCEL_VBA_KB')
    if env:
        return Path(env).expanduser().resolve(strict=True)
    config = json.loads((ROOT / 'metadata/knowledge-config.json').read_text(encoding='utf-8'))
    return (ROOT / config['relative_path']).resolve(strict=True)


class KnowledgeBase:
    def __init__(self, root: Path | str | None = None):
        self.root = resolve_kb(root)
        self.manifest = json.loads(contained_file(self.root, 'manifest.json').read_text(encoding='utf-8'))
        if self.manifest.get('schema_version') != SCHEMA_VERSION:
            raise ValueError('Unsupported knowledge-base schema; import/reindex explicitly')
        self.repository = self.root / 'repository'
        if self.repository.is_symlink():
            raise ValueError('The repository directory must not be a symlink')
        db = contained_file(self.root, 'index.sqlite')
        self.db = sqlite3.connect(db.as_uri() + '?mode=ro', uri=True)
        self.db.row_factory = sqlite3.Row
        self.db.execute('PRAGMA query_only=ON')
        self.db.execute('PRAGMA trusted_schema=OFF')

    def __enter__(self):
        return self

    def __exit__(self, *_):
        self.close()

    def close(self):
        self.db.close()

    @staticmethod
    def unpack(row: sqlite3.Row, score: float = 0) -> dict:
        d = dict(row)
        for key in ('api_names', 'sections'):
            d[key] = json.loads(d[key])
        d['score'] = score
        return d

    def search(self, query: str, *, exact: bool = False, scope: str = 'excel',
               category: str | None = None, kind: str | None = None, limit: int = 8) -> list[dict]:
        if not query.strip() or not 1 <= limit <= 100:
            raise ValueError('Non-empty query and limit between 1 and 100 required')
        if scope not in ('excel', 'all'):
            raise ValueError('scope must be excel or all')
        conditions, params = [], []
        if scope == 'excel':
            conditions.append('d.category IN (' + ','.join('?' for _ in EXCEL_CATEGORIES) + ')')
            params.extend(EXCEL_CATEGORIES)
        if category:
            conditions.append('d.category=?'); params.append(category)
        if kind:
            conditions.append('d.kind=?'); params.append(kind)
        where = ' AND ' + ' AND '.join(conditions) if conditions else ''
        rows = self.db.execute('SELECT DISTINCT d.* FROM documents d JOIN aliases a ON a.doc_id=d.id '
                               'WHERE a.alias=?' + where +
                               " ORDER BY CASE WHEN d.category='excel-api' THEN 0 ELSE 1 END, d.path LIMIT ?",
                               [api_key(query), *params, limit]).fetchall()
        results = [self.unpack(r, 1000) for r in rows]
        if exact or len(results) >= limit:
            return results
        words = list(dict.fromkeys(re.findall(r'[a-z0-9]+', normalize(query))))[:24]
        if not words:
            return results
        # Query grammar is generated from tokens, never accepted as raw FTS SQL.
        match = ' AND '.join('"' + word + '"' for word in words)
        ranked = self.db.execute(
            'SELECT d.*, bm25(document_fts, 12.0, 8.0, 1.0) AS rank FROM document_fts '
            'JOIN documents d ON d.id=document_fts.rowid WHERE document_fts MATCH ?' + where +
            ' ORDER BY rank, d.path LIMIT ?', [match, *params, min(100, limit + len(results))]).fetchall()
        seen = {d['path'] for d in results}
        for row in ranked:
            if row['path'] not in seen:
                d = self.unpack(row, -float(row['rank']))
                del d['rank']
                results.append(d); seen.add(d['path'])
        return results[:limit]

    def document(self, selector: str, scope: str = 'excel') -> dict:
        row = self.db.execute('SELECT * FROM documents WHERE path=? COLLATE NOCASE', (selector,)).fetchone()
        if row:
            return self.unpack(row)
        matches = self.search(selector, exact=True, scope=scope, limit=100)
        if not matches:
            raise LookupError('Not found in the supplied snapshot. This does not prove API nonexistence: ' + selector)
        if len(matches) > 1:
            raise ValueError('Ambiguous name. Read an explicit local path: ' + ', '.join(d['path'] for d in matches))
        return matches[0]

    def read_file(self, path: str) -> tuple[str, dict]:
        row = self.db.execute('SELECT * FROM files WHERE path=? COLLATE NOCASE', (path,)).fetchone()
        if not row:
            raise LookupError('Not in local file inventory: ' + path)
        info = dict(row)
        p = contained_file(self.repository, info['path'])
        if p.stat().st_size > MAX_FILE_BYTES:
            raise ValueError('File exceeds read limit')
        raw = p.read_bytes()
        if hashlib.sha256(raw).hexdigest() != info['sha256']:
            raise ValueError('Source integrity mismatch; reimport explicitly: ' + info['path'])
        text, _ = decode_markdown(raw)
        return text, info

    def read(self, selector: str, *, section: str | None = None, start_line: int | None = None,
             end_line: int | None = None, scope: str = 'excel') -> dict:
        doc = self.document(selector, scope)
        text, _ = self.read_file(doc['path'])
        lines = text.splitlines(keepends=True)
        lo, hi = 1, len(lines)
        if section:
            candidates = [s for s in doc['sections'] if normalize(s['heading']) == normalize(section)]
            if not candidates and normalize(section) in ('example', 'examples'):
                candidates = [s for s in doc['sections'] if normalize(s['heading']) in ('example', 'examples')]
            if len(candidates) != 1:
                raise LookupError('Section absent or ambiguous in source; headings: ' +
                                  ', '.join(s['heading'] for s in doc['sections']))
            lo, hi = candidates[0]['start_line'], candidates[0]['end_line']
        if start_line is not None:
            lo = start_line
        if end_line is not None:
            hi = end_line
        if lo < 1 or hi < lo or hi > len(lines):
            raise ValueError(f'Invalid line range; source has {len(lines)} lines')
        return {'document': doc, 'local_path': str(contained_file(self.repository, doc['path'])),
                'start_line': lo, 'end_line': hi, 'total_lines': len(lines),
                'complete_document': lo == 1 and hi == len(lines), 'content': ''.join(lines[lo-1:hi]),
                'source_sha256_verified': True}

    def resolve_link(self, from_path: str, target: str) -> dict:
        target = target.strip().strip('<>')
        split = urlsplit(target)
        if split.scheme and split.scheme not in ('http', 'https'):
            return {'target': target, 'status': 'external_or_unsupported'}
        if split.scheme:
            if split.netloc.casefold() != 'learn.microsoft.com' or '/office/vba/' not in split.path.casefold():
                return {'target': target, 'status': 'external_not_fetched'}
            p = re.split(r'/office/vba/', unquote(split.path), flags=re.I)[1]
        elif target.startswith('~/'):
            p = unquote(split.path[2:])
        elif split.path.startswith('/office/vba/'):
            p = unquote(split.path[len('/office/vba/'):])
        elif split.path.startswith('/'):
            return {'target': target, 'status': 'external_site_path_not_fetched'}
        elif not split.path:
            p = from_path
        else:
            p = posixpath.normpath(posixpath.join(posixpath.dirname(from_path), unquote(split.path)))
        try:
            safe_relative(p)
        except ValueError:
            return {'target': target, 'status': 'unsafe_link_rejected'}
        choices = [p]
        if not p.casefold().endswith(('.md', '.gif', '.png', '.jpg', '.jpeg', '.json', '.yml')):
            choices.extend([p + '.md', p + '/index.md'])
        for choice in choices:
            row = self.db.execute('SELECT path FROM files WHERE path=? COLLATE NOCASE', (choice,)).fetchone()
            if row:
                return {'target': target, 'status': 'local', 'path': row['path'], 'anchor': split.fragment}
        return {'target': target, 'status': 'missing_in_snapshot', 'candidate': p, 'anchor': split.fragment}

    def links(self, selector: str, scope: str = 'excel') -> list[dict]:
        doc = self.document(selector, scope)
        text, _ = self.read_file(doc['path'])
        targets = markdown_targets(visible_text(text))
        return [self.resolve_link(doc['path'], t) for t in dict.fromkeys(targets)]

    def includes(self, selector: str, scope: str = 'excel') -> list[dict]:
        doc = self.document(selector, scope)
        result, seen = [], {doc['path']}
        def walk(path: str, depth: int):
            if depth > 8:
                result.append({'status': 'include_depth_limit', 'path': path}); return
            text, _ = self.read_file(path)
            for target in re.findall(r'\[!include\[[^\]]*\]\(([^)]+)\)\]', visible_text(text), flags=re.I):
                resolved = self.resolve_link(path, target)
                if resolved['status'] != 'local':
                    result.append(resolved); continue
                rel = resolved['path']
                if rel in seen:
                    continue
                seen.add(rel)
                included, info = self.read_file(rel)
                result.append({**resolved, 'content': included, 'sha256': info['sha256']})
                walk(rel, depth + 1)
        walk(doc['path'], 0)
        return result

    def verify(self) -> dict:
        errors = []
        checked = 0
        file_rows = self.db.execute('SELECT path,sha256,bytes FROM files ORDER BY path').fetchall()
        for row in file_rows:
            try:
                path = contained_file(self.repository, row['path'])
                if path.stat().st_size != row['bytes'] or sha256_file(path) != row['sha256']:
                    errors.append({'path': row['path'], 'error': 'hash_or_size_mismatch'})
                else:
                    checked += 1
            except (OSError, ValueError) as exc:
                errors.append({'path': row['path'], 'error': str(exc)})
        count = self.db.execute('SELECT COUNT(*) FROM documents').fetchone()[0]
        local_md = sum(1 for p in self.repository.rglob('*') if p.is_file() and p.suffix.casefold() == '.md')
        m = self.manifest
        if count != m['counts']['indexed_markdown'] or local_md != count:
            errors.append({'error': 'markdown_count_mismatch'})
        if len(file_rows) != m['counts']['source_files']:
            errors.append({'error': 'source_file_count_mismatch'})
        if sha256_file(self.root / 'index.sqlite') != m['index_sha256']:
            errors.append({'error': 'index_hash_mismatch'})
        if sha256_file(self.root / 'documents.jsonl') != m['documents_jsonl_sha256']:
            errors.append({'error': 'portable_index_hash_mismatch'})
        if sha256_file(self.root / 'files.jsonl') != m['files_jsonl_sha256']:
            errors.append({'error': 'file_manifest_hash_mismatch'})
        quick = self.db.execute('PRAGMA quick_check').fetchone()[0]
        if quick != 'ok':
            errors.append({'error': 'sqlite_quick_check', 'detail': quick})
        return {'status': 'PASS' if not errors else 'FAIL', 'checked_files': len(file_rows),
                'hash_verified_files': checked, 'indexed_markdown': count, 'local_markdown': local_md,
                'sqlite_quick_check': quick, 'errors': errors,
                'scope': 'Supplied snapshot, byte integrity and local paths; not semantic or Excel runtime validation'}


def markdown_targets(text: str) -> list[str]:
    """Extract inline targets with balanced parentheses, e.g. Range(object).md."""
    targets = []
    for start in re.finditer(r'\]\(', text):
        pos, depth = start.end(), 1
        end = pos
        while end < len(text) and depth:
            char = text[end]
            if char == '\\':
                end += 2; continue
            if char == '(':
                depth += 1
            elif char == ')':
                depth -= 1
            if depth:
                end += 1
        if depth == 0:
            target = text[pos:end].strip()
            if target:
                targets.append(target.split(' "', 1)[0])
    return targets

"""Offline structural validator for this skill, not a YAML engine or VBA compiler."""
from __future__ import annotations
import argparse
import collections
import csv
import hashlib
import json
import re
import sys
from pathlib import Path
from urllib.parse import unquote, urlsplit
from common import ROOT, read_json, sha256_file, write_new_json

REQUIRED = ['SKILL.md', 'README.md', 'SOURCES.md', 'LICENSES.md', 'LIMITATIONS.md',
            'QUALITY-REPORT.md', 'CHANGELOG.md', 'agents/openai.yaml',
            'references/INDEX.md', 'references/documentation-index.md',
            'metadata/source-manifest.json', 'metadata/source-index.csv',
            'metadata/coverage-report.json', 'metadata/modules.json', 'metadata/routing.json',
            'tests/scenarios.json', 'tests/WINDOWS-ACCEPTANCE.md']

def parse_frontmatter(text: str) -> dict:
    """Validate the package's deliberately restricted two-scalar YAML subset."""
    match = re.match(r'\A---\r?\n(.*?)\r?\n---(?:\r?\n|$)', text, re.S)
    if not match:
        raise ValueError('Missing or malformed YAML frontmatter')
    result = {}
    for line in match.group(1).splitlines():
        parts = re.fullmatch(r'(name|description):\s*(.+)', line)
        if not parts or parts.group(1) in result:
            raise ValueError('Expected unique name/description scalar keys only')
        key, value = parts.groups()
        if value.startswith('"'):
            value = json.loads(value)
        elif not re.fullmatch(r'[a-z0-9]+(?:-[a-z0-9]+)*', value):
            raise ValueError('Use JSON-compatible double quoting for text scalars')
        if not isinstance(value, str) or not value.strip():
            raise ValueError('Empty/non-string frontmatter value')
        result[key] = value
    if set(result) != {'name', 'description'}:
        raise ValueError('Both name and description are required')
    if not re.fullmatch(r'[a-z0-9]+(?:-[a-z0-9]+)*', result['name']) or len(result['name']) > 64:
        raise ValueError('Invalid skill name')
    if len(result['description']) > 1024:
        raise ValueError('Description too long')
    return result

def relative_link_errors(root: Path, path: Path, text: str) -> list[str]:
    errors = []
    # Links in code samples do not declare resources.
    without_fences = re.sub(r'```.*?```', '', text, flags=re.S)
    for target in re.findall(r'(?<!!)\[[^\]]*\]\(([^\s)]+)\)', without_fences):
        if target.startswith(('#', 'https://', 'http://', 'mailto:')):
            continue
        clean = unquote(target.split('#', 1)[0])
        resolved = (path.parent / clean).resolve()
        if resolved != root and root not in resolved.parents:
            errors.append(f'{path.relative_to(root)}: link escapes skill: {target}')
        elif not resolved.exists():
            errors.append(f'{path.relative_to(root)}: missing target: {target}')
    return errors

def validate(root: Path) -> dict:
    root = root.resolve()
    errors, warnings = [], []
    counts = {}
    for required in REQUIRED:
        if not (root / required).is_file():
            errors.append('Missing required file: ' + required)
    if errors:
        return {'status':'FAIL', 'errors':errors, 'warnings':warnings, 'counts':counts}
    try:
        text = (root / 'SKILL.md').read_text(encoding='utf-8')
        header = parse_frontmatter(text)
        if header['name'] != root.name:
            errors.append('Skill folder does not match frontmatter name')
        if len(text.splitlines()) > 500:
            errors.append('SKILL.md exceeds 500 lines')
        counts['skill_lines'] = len(text.splitlines())
        content_hashes = collections.defaultdict(list)
        for path in root.rglob('*'):
            if path.is_symlink():
                errors.append('Symlink in package: ' + str(path.relative_to(root)))
            if not path.is_file() or '__pycache__' in path.parts:
                continue
            if path.suffix == '.md':
                body = path.read_text(encoding='utf-8')
                errors.extend(relative_link_errors(root, path, body))
                if re.search(r'\[TODO\]|TODO:|PLACEHOLDER', body):
                    errors.append('Unfinished authoring placeholder: ' + str(path.relative_to(root)))
                if path.name != 'README.md' and ('/mnt/data/' in body or re.search(r'[Cc]:\\Users\\', body)):
                    errors.append('Nonportable user path: ' + str(path.relative_to(root)))
                content_hashes[hashlib.sha256(body.encode()).hexdigest()].append(str(path.relative_to(root)))
        duplicates = [v for v in content_hashes.values() if len(v) > 1]
        if duplicates:
            errors.append('Duplicate Markdown content: ' + repr(duplicates))
        manifest = read_json(root / 'metadata/source-manifest.json')
        docs = manifest['sources']
        ids = [d['id'] for d in docs]
        urls = [d['url'] for d in docs]
        if len(ids) != len(set(ids)) or len(urls) != len(set(urls)):
            errors.append('Duplicate source IDs or URLs')
        for doc in docs:
            if urlsplit(doc['url']).scheme != 'https' or not doc.get('summary'):
                errors.append('Invalid source record: ' + doc['id'])
        coverage = read_json(root / 'metadata/coverage-report.json')
        if coverage['verified_source_documents'] != len(docs) or coverage['curated_index_records'] != len(docs):
            errors.append('Coverage/source count mismatch')
        if sum(c['verified_and_indexed'] for c in coverage['categories']) != len(docs):
            errors.append('Category totals mismatch')
        with (root / 'metadata/source-index.csv').open(encoding='utf-8-sig', newline='') as stream:
            rows = list(csv.DictReader(stream))
        if {d['id']:d['url'] for d in rows} != {d['id']:d['url'] for d in docs} or len(rows) != len(docs):
            errors.append('CSV/JSON index mismatch')
        modules = read_json(root / 'metadata/modules.json')
        for name, item in modules.items():
            file = root / 'references' / name
            if not file.is_file():
                errors.append('Missing domain module: ' + name)
                continue
            if not set(item['sources']) <= set(ids):
                errors.append('Unknown module source ID: ' + name)
            body = file.read_text(encoding='utf-8')
            for source_id in set(re.findall(r'\b(?:MS|VSC|PYWIN32)-[A-Z0-9-]+\b', body)):
                if source_id not in ids:
                    errors.append('Unknown referenced source ID: ' + source_id)
            if f'(references/{name})' not in text:
                errors.append('Domain module not directly linked by SKILL.md: ' + name)
        routes = read_json(root / 'metadata/routing.json')
        for route, data in routes.items():
            for rel in ['workflows/'+data['workflow']+'.md'] + ['references/'+s+'.md' for s in data['references']]:
                if not (root/rel).is_file():
                    errors.append(f'Route {route}: missing {rel}')
        cases = read_json(root / 'tests/scenarios.json')['scenarios']
        if len({s['id'] for s in cases}) != len(cases):
            errors.append('Duplicate scenario IDs')
        for case in cases:
            if case['route'] not in routes:
                errors.append('Unknown scenario route: ' + case['id'])
            for rel in case['required_resources']:
                if not (root/rel).is_file():
                    errors.append('Missing scenario resource: ' + rel)
            if not case['must_demonstrate'] or not case['must_not_do']:
                errors.append('Incomplete scenario contract: ' + case['id'])
        integrity_path = root / 'metadata/file-integrity.json'
        if integrity_path.exists():
            for rel, digest in read_json(integrity_path)['files'].items():
                resolved = (root/rel).resolve()
                if root not in resolved.parents or not resolved.is_file() or sha256_file(resolved) != digest:
                    errors.append('Integrity mismatch: ' + rel)
        else:
            warnings.append('No optional package integrity manifest')
        counts.update(source_records=len(docs), domain_modules=len(modules), workflows=len(list((root/'workflows').glob('*.md'))),
                      scenarios=len(cases), scripts=len(list((root/'scripts').glob('*.py'))), relative_link_files=len(list(root.rglob('*.md'))))
    except (OSError, ValueError, KeyError, TypeError) as exc:
        errors.append(type(exc).__name__ + ': ' + str(exc))
    return {'status':'FAIL' if errors else 'PASS', 'errors':errors, 'warnings':warnings, 'counts':counts,
            'scope':'Offline structure, contracts and hashes only; not live agent, VBA or COM validation'}

def main() -> int:
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--root',type=Path,default=ROOT)
    p.add_argument('--out',type=Path)
    a=p.parse_args()
    report=validate(a.root)
    try:
        if a.out: write_new_json(a.out,report)
        print(json.dumps(report,indent=2))
        return 0 if report['status']=='PASS' else 1
    except OSError as exc:
        print(f'ERROR: {exc}',file=sys.stderr);return 2

if __name__=='__main__':
    raise SystemExit(main())

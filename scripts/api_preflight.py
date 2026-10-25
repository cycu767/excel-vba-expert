"""Read-only namespace-aware evidence check for Excel VBA API names.

Return source evidence and a confidence boundary, never a claim that an API cannot exist.
Uses the installed unmodified Microsoft snapshot via offline_kb.KnowledgeBase.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import PurePosixPath
from offline_kb import KnowledgeBase

NAMESPACES = ('vba', 'excel', 'msforms', 'office', 'vbide')
VBIDE_GROUP = 'Language/Reference/Visual-Basic-Add-in-Model/'
LANGUAGE = 'Language/Reference/User-Interface-Help/'


def namespaced(path: str, namespace: str) -> bool:
    """Classify by original path, not by a keyword somewhere in an example."""
    p = path.casefold()
    if namespace == 'excel':
        return p.startswith('api/excel.') and '-graph' not in PurePosixPath(p).name
    if namespace == 'office':
        return p.startswith('api/office.')
    if namespace == 'vbide':
        return p.startswith(VBIDE_GROUP.casefold()) or (
            p.startswith(LANGUAGE.casefold()) and 'vba-add-in-object-model' in p)
    if namespace in ('vba', 'msforms'):
        return p.startswith(LANGUAGE.casefold())
    return False


def source_path_for_object(namespace: str, obj: str) -> list[str]:
    if not re.fullmatch(r'[A-Za-z][A-Za-z0-9_]*', obj):
        return []
    if namespace == 'excel':
        return [f'api/Excel.{obj}.md', f'api/Excel.{obj}(object).md']
    if namespace == 'office':
        return [f'api/Office.{obj}.md']
    if namespace == 'msforms':
        return [f'{LANGUAGE}{obj.lower()}-control.md', f'{LANGUAGE}{obj.lower()}-object.md']
    if namespace == 'vba':
        return [f'{LANGUAGE}{obj.lower()}-{kind}.md' for kind in (
            'statement', 'function', 'data-type', 'object', 'keyword', 'operator', 'property', 'method')]
    return []


def normalize_symbol(namespace: str, symbol: str) -> list[str]:
    chunks = symbol.strip().split('.')
    if not all(re.fullmatch(r'[A-Za-z_][A-Za-z0-9_]*', p) for p in chunks):
        raise ValueError('Use a simple API symbol like ListBox.ItemHeight, not code or arbitrary text')
    if chunks[0].casefold() in NAMESPACES:
        if chunks[0].casefold() != namespace:
            raise ValueError('Symbol namespace disagrees with the requested --namespace')
        chunks = chunks[1:]
    if not 1 <= len(chunks) <= 2:
        raise ValueError('Specify Object or Object.Member (optionally prefixed with namespace)')
    return chunks


def heading_section(text: str, heading: str) -> tuple[int, int] | None:
    lines = text.splitlines()
    for n, line in enumerate(lines):
        m = re.match(r'^(#{2,4})\s+(.+?)\s*$', line)
        if m and m.group(2).strip(' `').casefold() == heading.casefold():
            level = len(m.group(1))
            end = len(lines)
            for x in range(n + 1, len(lines)):
                h = re.match(r'^(#{1,6})\s+', lines[x])
                if h and len(h.group(1)) <= level:
                    end = x
                    break
            return (n + 1, end)
    return None


def check(kb: KnowledgeBase, namespace: str, symbol: str) -> dict:
    if namespace not in NAMESPACES:
        raise ValueError('Unknown namespace')
    parts = normalize_symbol(namespace, symbol)
    obj, member = parts[0], parts[1] if len(parts) == 2 else None
    evidence = []
    seen = set()

    def add(path: str, evidence_type: str, heading: str | None = None):
        if path.casefold() in seen or not namespaced(path, namespace):
            return
        try:
            raw, doc = kb.read_file(path)
        except (FileNotFoundError, LookupError, ValueError, OSError):
            return
        line_range = heading_section(raw, heading) if heading else None
        if heading and line_range is None:
            return
        seen.add(path.casefold())
        row = kb.db.execute('SELECT title FROM documents WHERE path=? COLLATE NOCASE', (path,)).fetchone()
        evidence.append({
            'source_path': path, 'title': row['title'] if row else PurePosixPath(path).stem, 'sha256_verified': True,
            'source_lines': list(line_range or (1, len(raw.splitlines()))),
            'evidence_type': evidence_type,
        })

    if namespace == 'vbide':
        # VBIDE reference describes many members in one chapter each, rather than
        # independent api/VBIDE.VBComponent.md files. Check original headings.
        files = ('objects', 'collections', 'properties', 'methods', 'events')
        for type_name in files:
            path = f'{VBIDE_GROUP}{type_name}-visual-basic-add-in-model.md'
            add(path, 'OBJECT_OR_COLLECTION_SECTION' if member is None else
                'GROUPED_MEMBER_SECTION_REQUIRES_OWNER_CHECK',
                heading=obj if member is None else member)
        if member:
            for type_name in ('objects', 'collections'):
                path = f'{VBIDE_GROUP}{type_name}-visual-basic-add-in-model.md'
                add(path, 'OWNER_TYPE_CONTEXT', heading=obj)
    else:
        if member is None:
            for path in source_path_for_object(namespace, obj):
                add(path, 'DIRECT_DOCUMENT')
        else:
            if namespace in ('excel', 'office'):
                add(f'api/{namespace.capitalize()}.{obj}.{member}.md', 'DIRECT_MEMBER_DOCUMENT')
            elif namespace in ('msforms', 'vba'):
                for kind in ('property', 'method', 'event', 'function', 'statement'):
                    add(f'{LANGUAGE}{member.lower()}-{kind}-microsoft-forms.md',
                        'GENERIC_FORMS_MEMBER_REQUIRES_OWNER_CHECK')
                    add(f'{LANGUAGE}{member.lower()}-{kind}.md',
                        'GENERIC_MEMBER_REQUIRES_OWNER_CHECK')
            if namespace in ('msforms', 'vba'):
                # A single MSForms property page can be shared by several controls
                # and can combine names (e.g. InsideHeight/InsideWidth). Return as
                # a candidate, never as evidence of this object's ownership.
                for hit in kb.search(member, exact=False, scope='all', limit=100,
                                     category='vba-reference'):
                    stem = PurePosixPath(hit['path']).stem.casefold().replace('-', '')
                    if member.casefold() in stem:
                        add(hit['path'], 'GENERIC_MEMBER_CANDIDATE_REQUIRES_OWNER_CHECK')
            for path in source_path_for_object(namespace, obj):
                add(path, 'OWNER_TYPE_CONTEXT')

        # Use up to 100 exact alias hits; do not infer absence from first 8.
        queries = [symbol, '.'.join(parts), obj + (' ' + member if member else ''),
                   f'{obj} control' if namespace == 'msforms' else obj]
        for query in dict.fromkeys(queries):
            for hit in kb.search(query, exact=True, scope='all', limit=100):
                if namespaced(hit['path'], namespace):
                    add(hit['path'], 'EXACT_INDEX_CANDIDATE_REQUIRES_REVIEW')

    direct = any(e['evidence_type'] in ('DIRECT_DOCUMENT', 'DIRECT_MEMBER_DOCUMENT', 'OBJECT_OR_COLLECTION_SECTION')
                 for e in evidence)
    if not member and direct:
        status = 'SOURCE_DOCUMENT_FOUND'
    elif member and any(e['evidence_type'] == 'DIRECT_MEMBER_DOCUMENT' for e in evidence):
        status = 'MEMBER_DOCUMENT_FOUND_CHECK_OWNER'
    elif member and evidence and all(e['evidence_type'] == 'OWNER_TYPE_CONTEXT' for e in evidence):
        status = 'MEMBER_NOT_CONFIRMED_OWNER_ONLY'
    elif evidence:
        status = 'CANDIDATES_REQUIRE_MEMBER_VERIFICATION'
    else:
        status = 'NOT_CONFIRMED_IN_SNAPSHOT'
    return {
        'namespace': namespace, 'symbol': symbol, 'status': status,
        'warning': ('A source hit is not an Excel compile/runtime test. A grouped or generic member '
                    'does not establish that the requested class exposes it. A miss is not proof '
                    'that the API does not exist. Use the source and Excel Object Browser.'),
        'evidence': evidence,
    }


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('--namespace', choices=NAMESPACES, required=True)
    p.add_argument('--symbol', required=True)
    p.add_argument('--kb', help='Override knowledge folder')
    p.add_argument('--json', action='store_true')
    args = p.parse_args(argv)
    try:
        with KnowledgeBase(args.kb) as kb:
            result = check(kb, args.namespace, args.symbol)
        if args.json:
            print(json.dumps(result, indent=2, ensure_ascii=False))
        else:
            print(f"{result['namespace']} :: {result['symbol']} -> {result['status']}")
            for e in result['evidence']:
                print(f"  {e['evidence_type']}: {e['source_path']} lines {e['source_lines']}")
            print('  ' + result['warning'])
        return 0
    except (OSError, ValueError, LookupError) as exc:
        print('ERROR: ' + str(exc), file=sys.stderr)
        return 2


if __name__ == '__main__':
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    raise SystemExit(main())

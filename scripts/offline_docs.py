"""Search and read FULL local Microsoft VBA source documents; never use the network."""
from __future__ import annotations
import argparse
import json
import sqlite3
import sys
from pathlib import Path
from offline_kb import KnowledgeBase


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('--kb', type=Path, help='Knowledge folder override; default is the skill sibling')
    sub = p.add_subparsers(dest='command', required=True)
    search = sub.add_parser('search', help='Exact API aliases first; then full-text relevance')
    search.add_argument('query')
    search.add_argument('--exact', action='store_true')
    search.add_argument('--limit', type=int, default=8)
    search.add_argument('--category')
    search.add_argument('--kind', choices=['object','method','property','event','enumeration','statement','function','collection','article'])
    search.add_argument('--scope', choices=['excel','all'], default='excel')
    search.add_argument('--json', action='store_true')
    read = sub.add_parser('read', help='Read a full source file or an explicit original-line range')
    read.add_argument('selector', help='API/curated ID, or exact path returned by search')
    read.add_argument('--scope', choices=['excel','all'], default='excel')
    read.add_argument('--section')
    read.add_argument('--start-line', type=int)
    read.add_argument('--end-line', type=int)
    read.add_argument('--expand-includes', action='store_true', help='Append local include text, without changing source')
    read.add_argument('--json', action='store_true')
    links = sub.add_parser('links', help='Resolve source cross-references to local paths')
    links.add_argument('selector')
    links.add_argument('--scope', choices=['excel','all'], default='excel')
    sub.add_parser('status', help='Snapshot identity, counts and local paths')
    sub.add_parser('verify', help='Hash-check all source files and both derived indexes')
    args = p.parse_args(argv)
    try:
        with KnowledgeBase(args.kb) as kb:
            if args.command == 'status':
                print(json.dumps({'knowledge_root': str(kb.root), **kb.manifest}, ensure_ascii=False, indent=2))
                return 0
            if args.command == 'verify':
                result = kb.verify()
                print(json.dumps(result, indent=2))
                return 0 if result['status'] == 'PASS' else 1
            if args.command == 'links':
                print(json.dumps(kb.links(args.selector, args.scope), ensure_ascii=False, indent=2))
                return 0
            if args.command == 'search':
                results = kb.search(args.query, exact=args.exact, limit=args.limit, scope=args.scope,
                                    category=args.category, kind=args.kind)
                for d in results:
                    d['local_path'] = str(kb.repository / d['path'])
                    d['read_command_selector'] = d['path']
                report = {'status': 'matches' if results else 'not_found_in_snapshot',
                          'scope': args.scope, 'warning': 'A miss is not proof of API nonexistence. Use local path + read to inspect full source.',
                          'results': results}
                if args.json:
                    print(json.dumps(report, ensure_ascii=False, indent=2))
                else:
                    for d in results:
                        print(f"{d['title']} [{d['category']} / {d['kind']}]\n  {d['local_path']}\n  {d['description_excerpt']}\n")
                    if not results:
                        print(report['warning'])
                return 0 if results else 1
            result = kb.read(args.selector, section=args.section, start_line=args.start_line,
                             end_line=args.end_line, scope=args.scope)
            if args.expand_includes:
                result['includes'] = kb.includes(args.selector, args.scope)
            if args.json:
                print(json.dumps(result, ensure_ascii=False, indent=2))
            else:
                print(f"SOURCE: {result['local_path']}\nLINES: {result['start_line']}-{result['end_line']} of {result['total_lines']}\n"
                      f"SHA256 VERIFIED: {result['document']['sha256']}\n\n", end='')
                print(result['content'], end='' if result['content'].endswith('\n') else '\n')
                for included in result.get('includes', []):
                    print('\n--- LOCAL INCLUDE: ' + included.get('path', included.get('target', '')) + ' ---')
                    print(included.get('content', json.dumps(included, ensure_ascii=False)))
            return 0
    except (FileNotFoundError, OSError, ValueError, LookupError, sqlite3.Error) as exc:
        print('ERROR: ' + str(exc) + '\nInstall the separate local knowledge package or use --kb PATH. '
              'No online or summary-only fallback was attempted.', file=sys.stderr)
        return 2


if __name__ == '__main__':
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    raise SystemExit(main())

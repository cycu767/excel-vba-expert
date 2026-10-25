"""Search local authored summaries or a locally built repository index. No network."""
from __future__ import annotations
import argparse
import json
import re
import sys
import unicodedata
from pathlib import Path
from common import ROOT, read_json

def normalize(text: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", text.casefold().replace("\u0142", "l")) if not unicodedata.combining(c))

def tokens(text: str) -> set[str]:
    return set(re.findall(r"[a-z0-9]+", normalize(text)))

def member_key(text: str) -> str:
    t = normalize(text).strip().removeprefix("excel.")
    return re.sub(r"\s+(method|property|event|object).*$", "", t)

def search(query: str, documents: list[dict], limit: int = 8, category: str | None = None, exact: bool = False) -> list[dict]:
    q = normalize(query).strip()
    if not q or limit < 1:
        return []
    qt = tokens(q)
    scored = []
    for d in documents:
        if category and d.get("category") != category:
            continue
        title = normalize(d.get("title", ""))
        ident = normalize(d.get("id", ""))
        tail = d.get("url", "").rstrip("/").split("/")[-1]
        names = {member_key(title), member_key(tail), ident}
        if exact and member_key(q) not in names:
            continue
        text = normalize(" ".join([d.get("title", ""), d.get("summary", ""), d.get("url", ""), " ".join(d.get("keywords", [])), d.get("repo_path") or ""]))
        overlap = qt & tokens(text)
        if not exact and not overlap and q not in text and q != ident:
            continue
        score = (100 if member_key(q) in names else 0) + (35 if q in title else 0) + (20 if q in text else 0)
        score += 8 * len(qt & tokens(title)) + 3 * len(overlap)
        if qt and not qt <= tokens(text):
            score -= 5 * len(qt - tokens(text))
        scored.append((score, d))
    scored.sort(key=lambda item: (-item[0], item[1].get("id", "")))
    return [dict(d, search_score=score) for score, d in scored[:limit]]

def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("query", nargs="?", default="")
    p.add_argument("--index", type=Path, default=ROOT / "metadata/source-manifest.json")
    p.add_argument("--limit", type=int, default=8)
    p.add_argument("--category")
    p.add_argument("--exact", action="store_true")
    p.add_argument("--json", action="store_true")
    p.add_argument("--show", help="Show the complete local record for a source ID")
    args = p.parse_args()
    try:
        data = read_json(args.index)
        docs = data.get("sources", data.get("documents", []))
        if not isinstance(docs, list):
            raise ValueError("Index must contain a sources/documents list")
        if args.show:
            results = [d for d in docs if d.get("id", "").casefold() == args.show.casefold()]
        else:
            if not args.query.strip():
                p.error("Provide a query or --show ID")
            if not 1 <= args.limit <= 100:
                p.error("--limit must be between 1 and 100")
            results = search(args.query, docs, args.limit, args.category, args.exact)
        report = {"status": "matches" if results else "not_found_in_loaded_index", "scope": data.get("inventory_status", "local_checkout"), "warning": "A miss in this index is not proof that an API does not exist. Summaries do not replace full signatures.", "results": results}
        if args.json or args.show:
            print(json.dumps(report, ensure_ascii=False, indent=2))
        else:
            print(report["warning"])
            for d in results:
                print(f"\n{d['id']} | {d['title']} | {d.get('category', '')}\n{d.get('summary', '')}\n{d.get('url', '')}")
            if not results:
                print(report["status"])
        return 0 if results else 1
    except (OSError, ValueError, KeyError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

if __name__ == "__main__":
    raise SystemExit(main())

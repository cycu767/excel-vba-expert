"""Inventory a user-supplied local VBA-Docs checkout. Never download, fetch, or modify it."""
from __future__ import annotations
import argparse
import collections
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import quote
from common import write_new_json, sha256_file

MAX_DOC_BYTES = 4 * 1024 * 1024

def classify(path: str) -> tuple[str, bool]:
    p = path.casefold()
    if p.startswith(("includes/", "f1/")):
        return "support-or-f1", False
    if p.startswith("api/excel."):
        return "excel-api", True
    if p.startswith("api/office."):
        return "office-api", True
    if p.startswith("api/graph.") or "graph-visual-basic" in p:
        return "graph", True
    if p.startswith("api/overview/"):
        return "overview", True
    if p.startswith("excel/"):
        return "excel-concepts", True
    if p.startswith("language/concepts/forms/"):
        return "forms-concepts", True
    if "visual-basic-add-in-model" in p or "add-in-model" in p:
        return "vbe", True
    if p.startswith("language/how-to/"):
        return "vba-how-to", True
    if p.startswith("language/"):
        return "vba-language", True
    if p.startswith("library-reference/"):
        return "office-concepts", True
    return "other-office-or-root", False

def scalar(front: str, key: str) -> str | None:
    match = re.search(r"^" + re.escape(key) + r":\s*(.*?)\s*$", front, re.M)
    if not match:
        return None
    val = match.group(1).strip()
    if val in {"", ">", "|", ">-", "|-"}:
        return None  # No pretend full YAML parser; use a heading fallback.
    return val.strip("\"'")

def parse_document(path: str, raw: bytes) -> dict:
    text = raw.decode("utf-8-sig", errors="strict")
    match = re.match(r"\A---\s*\r?\n(.*?)\r?\n---\s*\r?\n", text, re.S)
    front, body = (match.group(1), text[match.end():]) if match else ("", text)
    heading = re.search(r"^#\s+(.+)$", body, re.M)
    title = scalar(front, "title") or (heading.group(1).strip() if heading else Path(path).stem)
    description = scalar(front, "description")
    if description is None:
        paragraphs = re.split(r"\n\s*\n", body)
        description = next((s.strip() for s in paragraphs if s.strip() and not s.lstrip().startswith(("#", "[!", "```", "|", "<!--"))), "")
    description = re.sub(r"\[([^\]]+)\]\([^)]*\)", r"\1", description)
    description = re.sub(r"\s+", " ", description).strip()[:600]
    category, included = classify(path)
    stem = path[:-3] if path.lower().endswith(".md") else path
    return {"id": "REPO-" + hashlib.sha256(path.encode()).hexdigest()[:16], "title": title,
            "repo_path": path, "category": category, "included": included,
            "url": "https://learn.microsoft.com/en-us/office/vba/" + quote(stem.lower(), safe="/().-"),
            "url_status": "inferred_from_path_not_network_checked", "summary": description,
            "summary_status": "mechanically_extracted_not_curated", "keywords": [Path(path).stem],
            "ms_date": scalar(front, "ms.date"), "sha256": hashlib.sha256(raw).hexdigest(),
            "bytes": len(raw), "deprecation_status": "not_determined",
            "deprecation_keyword_hits": re.findall(r"\b(?:deprecated|obsolete)\b", body, flags=re.I)[:5]}

def inventory(repo: Path, include_other: bool = False) -> dict:
    repo = repo.resolve(strict=True)
    if not (repo / "docfx.json").is_file() or not (repo / "LICENSE").is_file():
        raise ValueError("Expected local VBA-Docs root with docfx.json and LICENSE")
    documents, excluded, problems = [], [], []
    counts = collections.Counter()
    total = 0
    for directory, dirs, files in os.walk(repo, followlinks=False):
        dirs[:] = sorted(d for d in dirs if d != ".git" and not (Path(directory) / d).is_symlink())
        for name in sorted(files):
            p = Path(directory) / name
            if p.suffix.casefold() != ".md":
                continue
            rel = p.relative_to(repo).as_posix()
            total += 1
            category, selected = classify(rel)
            counts[category] += 1
            if p.is_symlink():
                problems.append({"path": rel, "error": "symlink_not_followed"})
                continue
            if not selected and not include_other:
                excluded.append(rel)
                continue
            try:
                if p.stat().st_size > MAX_DOC_BYTES:
                    raise ValueError("document_exceeds_size_limit")
                documents.append(parse_document(rel, p.read_bytes()))
            except (OSError, UnicodeError, ValueError) as exc:
                problems.append({"path": rel, "error": str(exc)})
    revision = None
    if (repo / ".git").exists() and shutil.which("git"):
        try:
            r = subprocess.run(["git", "-C", str(repo), "rev-parse", "HEAD"], capture_output=True, text=True, timeout=10, check=True)
            if re.fullmatch(r"[0-9a-fA-F]{40,64}", r.stdout.strip()):
                revision = r.stdout.strip()
        except (OSError, subprocess.SubprocessError):
            pass
    return {"schema_version": 1, "created_at_utc": datetime.now(timezone.utc).isoformat(),
            "repository": "MicrosoftDocs/VBA-Docs (local checkout; identity not authenticated)",
            "repository_revision": revision, "inventory_status": "local_checkout_partial" if problems else "local_checkout_scanned",
            "repository_total_markdown_files": total, "indexed_documents": len(documents),
            "counts_by_category": dict(sorted(counts.items())), "excluded_paths": excluded,
            "problems": problems, "license_sha256": sha256_file(repo / "LICENSE"),
            "notes": ["Complete scan refers to supplied checkout, not proof that checkout equals upstream.",
                      "No online link verification and no semantic review of generated entries.",
                      "ms.date is author metadata, not last Git modification time.",
                      "Publishing redirects, include resolution and non-Markdown API sources need separate review.",
                      "Canonical Learn URLs are inferred and may need a publishing redirect or special mapping."],
            "documents": documents}

def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--repo", required=True, type=Path)
    p.add_argument("--out", required=True, type=Path)
    p.add_argument("--include-other-office", action="store_true")
    a = p.parse_args()
    try:
        repo = a.repo.resolve(strict=True)
        dest = a.out.resolve()
        if dest == repo or repo in dest.parents:
            raise ValueError("Write the generated index outside the source checkout")
        report = inventory(repo, a.include_other_office)
        write_new_json(dest, report)
        print(json.dumps({k: report[k] for k in ("inventory_status", "repository_revision", "repository_total_markdown_files", "indexed_documents", "counts_by_category")}, indent=2))
        return 1 if report["problems"] else 0
    except (OSError, ValueError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

if __name__ == "__main__":
    raise SystemExit(main())

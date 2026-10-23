# Limitations - version 2.0.0

## What is no longer a limitation
The full supplied repository is physically included in the separate knowledge
package. All 38,187 Markdown files are indexed and readable. The old v1 limitation
"no source mirror / unknown total" has been resolved. Both ZIPs must be installed.

## Remaining limits
1. Snapshot coverage is complete **relative to the supplied archive**, not an
   independently authenticated current Microsoft repository. Its ZIP comment gives
   a revision identifier; the archive SHA-256 identifies exact received bytes.
   No upstream branch comparison or per-file Git history was performed.
2. Documentation source availability is not source correctness or semantic review.
   The 19 expert modules remain selected guidance. Original source omissions,
   old examples, hidden HTML comments and publishing directives remain intact.
   ChartObjects.Add has a partially described parameter table and no Remarks.
3. External pages/assets are not downloaded. Win32 APIs, full VS Code and pywin32
   documentation are not contained in VBA-Docs; the preserved curated layer has
   selected summaries only for those external sources. Local links/includes are
   resolved where a file exists; unresolved/external targets are explicitly marked.
   This is not a complete DocFX site renderer and does not follow publishing redirects.
4. Search is lexical/alias-based, not an embedding or translation system. Prefer
   source API names and English terms. Default search scope is Excel and supporting
   Language/Forms/Office/VBE; --scope all exposes all indexed files. Ambiguous names
   require an explicit source path. A miss is not proof of API nonexistence.
5. Python 3.10+ and SQLite FTS5 are required for the indexed tools. Source Markdown
   and JSONL remain directly readable without the SQLite engine. Ordinary use and
   import require no network or additional Python packages. The strong offline test
   harness separately requires Linux/libseccomp; it was actually run in this build.
6. No Windows/Excel/VBE/MSForms UI or live VS Code/Qwen session was available.
   Runtime, compilation, COM-process exit, UserForm rendering and agent compliance
   remain NOT_RUN. Python tests and mocked COM are not substitutes for these gates.
7. A skill is instructions plus tools, not a network firewall or a local language
   model. The tested process tree had kernel-enforced networking disabled; that
   does not isolate the whole host, the AI client or its cloud model.
8. The preserved VBA scanner is a heuristic, not a VBA parser/compiler or malware
   detector. The workbook inspector checks ZIP metadata, not formulas, code or safe
   execution. Existing examples have not been compiled or tested in real Excel.
9. New imports use a new destination and source limits (100,000 files, 512 MiB,
   32 MiB per file). Symlinks, unsafe paths, collisions and undecodable Markdown
   are rejected. A rejected import is not a partial successful package.
10. Hashes detect content changes but are not an author signature. SQLite and JSONL
    indexes are derived artifacts; reimport explicitly after source changes. No
    background refresh, silent overwrite or automatic fallback to online content.
11. v1 expert module wording about a partial index refers to the retained 70-record
    curated index. The v2 mandatory workflow uses the full-source knowledge layer.
    The original 19 modules, seven workflows and eight Python files are preserved
    byte-for-byte and are not replaced by raw documentation.

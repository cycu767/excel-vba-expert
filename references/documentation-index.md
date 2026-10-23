# Local full-source documentation lookup

## Two cooperating layers
The sibling knowledge package contains every supplied repository file unchanged.
SQLite FTS5 indexes the full visible Markdown text; documents.jsonl is a portable
metadata index with **local source paths**, not substitute content. The repository
files remain independently readable and searchable with a text editor or ripgrep.

The existing 19 expert modules, seven workflows, 70 reviewed source summaries,
examples and original tools remain the separate engineering layer. Use full source
for the API contract and reviewed rules for safe implementation. The old
search_docs.py and index_repository.py remain available, but they are not the new
full-source import/read system. Do not mistake their summaries for Microsoft text.

## Runtime requirements and location
Use Python 3.10+ with sqlite3 FTS5 (standard Windows Python builds normally include
it; import gives an explicit failure if unavailable). No additional Python package,
network service, API key, Excel installation or Git executable is required.
Resolve the default knowledge path from metadata/knowledge-config.json relative
to the installed skill. The default sibling is `../excel-vba-expert-knowledge`.
Explicit `--kb` overrides EXCEL_VBA_KB, which overrides the portable config.

## Mandatory sequence
1. Identify owning namespace and exact member; search locally before uncertain use.
2. Inspect exact hits. Select by actual source path when names are ambiguous.
3. Read the full source, including relevant local includes and linked types/enums.
4. Record description, syntax, parameter/parts table, return type, remarks and
   examples that actually exist. Report omissions; do not manufacture sections.
5. Apply the relevant curated workflow/module, preserving business behavior.
6. Cite local source path, original line numbers and snapshot evidence. Separate
   documentation facts from engineering choices and unexecuted VBA/runtime tests.

## Commands, from any working directory
```text
python "<skill-root>/scripts/offline_docs.py" status
python "<skill-root>/scripts/offline_docs.py" search "Range.Find" --exact --json
python "<skill-root>/scripts/offline_docs.py" read "Range.Find" --expand-includes
python "<skill-root>/scripts/offline_docs.py" read "Range.Find" --section Remarks
python "<skill-root>/scripts/offline_docs.py" read "Range.Find" --start-line 20 --end-line 39
python "<skill-root>/scripts/offline_docs.py" search "Range" --kind object
python "<skill-root>/scripts/offline_docs.py" read "api/Excel.Range(object).md"
python "<skill-root>/scripts/offline_docs.py" read "xlWhole"
python "<skill-root>/scripts/offline_docs.py" search "modal modeless" --limit 5
python "<skill-root>/scripts/offline_docs.py" search "Document" --scope all
python "<skill-root>/scripts/offline_docs.py" links "Range.Find"
python "<skill-root>/scripts/offline_docs.py" --kb "<knowledge-folder>" verify
```
`--json` is supported by search/read. read returns the full file by default, verifies
its hash and never silently truncates. Section names follow the actual source;
UserForm.Show has its argument-parts table within Syntax, not a Parameters heading.
`--expand-includes` appends locally resolved include files without modifying source.
links returns local paths, external-not-fetched or missing-in-snapshot explicitly.
Original case-mismatched relative links are resolved case-insensitively. This is not
an HTML/DocFX renderer; raw directives, source comments and source flaws remain.

## Matching and absence
Search uses exact aliases first, then AND-token full-text ranking. Use source API
names or English keywords; this is not semantic multilingual translation. Exact
aliases include api_name, titles, filename stems, Excel-prefixed names, event-handler
notation, enum constants, explicit MSForms aliases and mapped curated IDs.
Broad results are candidates, not evidence for a claim. A failed search does not
prove API nonexistence. The default scope prioritizes Excel plus Language/Forms/
Office/VBE; `--scope all` makes every indexed Markdown document searchable.
A read selector matching more than one document fails with candidate paths instead
of choosing a legacy Graph or wrong-library page. Use the returned explicit path.

## Raw fallback, without SQLite
Read the Markdown files under `<knowledge-folder>/repository/`. Search filenames,
api_name and titles, or scan documents.jsonl for `path`/`local_path`. For example,
Range.Find is `repository/api/Excel.Range.Find.md`. Do not substitute a URL for text.

## Updating from a local ZIP or directory
```text
python "<skill-root>/scripts/import_knowledge.py" --source "<VBA-Docs-main.zip>" --dest "<new-knowledge-folder>"
python "<skill-root>/scripts/import_knowledge.py" --source "<local-VBA-Docs-directory>" --dest "<another-new-folder>"
python "<skill-root>/scripts/offline_docs.py" --kb "<new-knowledge-folder>" verify
```
Import copies all non-Git source files and licenses byte-for-byte, builds new SQLite
and JSONL indexes and writes provenance/counts/hashes. It rejects unsafe ZIP paths,
case collisions, symlinks, absent license files and undecodable Markdown. It uses a
staging directory and publishes only on success, without replacing a destination.
No fetch, subprocess execution of source content, or background update occurs.

## Evidence and limits
The archive SHA-256 and ZIP-comment revision identify the supplied snapshot, not
an independently authenticated current GitHub checkout. The count includes F1 and
navigation files; it is not a count of unique API members. Classification by title
is heuristic. Every Markdown file is indexed, but not every page is semantically
reviewed. External Win32/VS Code/pywin32 pages are not supplied by VBA-Docs.
See [coverage](../metadata/coverage-report.json), [sources](../SOURCES.md),
[API checks](../OFFLINE-VERIFICATION.md) and [limitations](../LIMITATIONS.md).

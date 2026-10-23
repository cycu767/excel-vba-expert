# Offline verification - full local Microsoft source layer

## What was verified against the supplied archive
All 38,911 source files were read from the actual ZIP and compared by SHA-256 with
the imported inventory. Every physical source file was independently hash-checked.
There were **zero missing files and zero byte mismatches**. All 38,187 Markdown
paths are indexed and physically present; the complete original text is readable.

The automatic suite also looks up every Excel api_name entry (5,462 entries) and
checks that the matching source path is returned. This measures the names present
in the snapshot, not the universe of possible/current Excel members. Every one of
38,187 Markdown documents is separately tested for exact-path retrieval.

## Required API checks
| Requested API | Complete source file | Syntax | Parameters | Remarks | Examples |
|---|---|---|---|---|---|
| Range.Find | api/Excel.Range.Find.md | Present | Present | Present | Present |
| Workbooks.Open | api/Excel.Workbooks.Open.md | Present | Present | Present | Present |
| Worksheet.Change | api/Excel.Worksheet.Change.md | Present | Present | Present | Present |
| UserForm.Show | Language/Reference/User-Interface-Help/show-method.md | Present | Parts table under Syntax | Present | Present |
| ChartObjects.Add | api/Excel.ChartObjects.Add.md | Present | Source table has only Left and Width | Absent in source | Present |

Descriptions and complete source text were read for all five APIs. The reader does
not rely on an authored synopsis. It verifies the source hash, returns original
line ranges and can append full locally resolved include text. The standalone
JSON report stores exact source paths, hashes, line counts and section locations.

### Source gaps are not hidden
ChartObjects.Add lists Left, Top, Width and Height in its syntax but the original
parameter table contains only Left and Width. There is no Remarks heading. The
source has been preserved; this package does not fabricate missing descriptions.
UserForm.Show is an authored lookup alias for Microsoft's Show method page. Its
api_name frontmatter remains Office.Show; the source has not been renamed/changed.

Range.Find includes historical material inside HTML comments. Full raw reads retain
it; the full-text index excludes commented-out material. Do not recommend commented
samples as current API guidance. Official examples may need the skill's error,
ownership and event-cleanup safeguards before use in production.

## Offline proof
The full executed suite passed 117/117 tests, with zero failures, errors or skips.

The test runner can install a Linux seccomp filter that rejects socket creation,
connect, send/receive and related network operations, including in subprocesses.
IPv4 and IPv6 probes fail with EPERM; a child Python process is also blocked.
Source search/read/import fixtures are tested under that filter, not merely with
an assertion that code probably does not use HTTP. Exact run results, counts and
scope are recorded in [offline-test-results.json](metadata/offline-test-results.json).
The filter applies to the test process tree, not the entire host or AI service.

The portable default test guard is a separate Python-only check. It is not called
OS isolation. Windows users can disconnect the machine or deny the agent/runtime
network through their own environment and execute the same lookup/read tests.
The skill does not modify firewall settings or turn a cloud model into a local one.

## Reproduction
From the installed skill root:
```text
python scripts/offline_docs.py verify
python tests/verify_snapshot.py --archive "<original-VBA-Docs-main.zip>" --out "<new-report.json>"
python -m unittest discover -s tests -v
```
Linux with libseccomp.so.2:
```text
python tests/run_offline_tests.py --enforce-os --out "<new-offline-report.json>"
```
The original archive is needed only to repeat byte comparison, not for ordinary
use or the unit suite. Do not claim offline proof if --enforce-os fails; report
the actual available guard. No test compiles or executes VBA in real Excel.

## Evidence files
[Snapshot and API results](metadata/offline-verification.json),
[offline test results](metadata/offline-test-results.json),
[coverage](metadata/coverage-report.json),
[preservation check](metadata/preservation-report.json).

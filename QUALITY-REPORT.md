# Quality report - version 2.0.0

This update separates four kinds of evidence: original-file integrity, local
retrieval, agent behavior, and real Excel execution. The first two are exercised;
the latter two are not claimed. Structured test reports record actual run counts.

## Source integrity and retrieval
All 38,911 input source files match their physical copies by SHA-256. All 38,187
Markdown files exist, have metadata and full-text index entries, and support local
reads. Original archive entry CRCs and SHA-256 were checked, not inferred from names.
The five required API pages were fully read; present/missing source sections were
recorded without inventing missing descriptions. Exact local paths and licenses
were preserved; the curated layer remains separate and intact.

## Executed test evidence
**117/117 tests passed: 59 original tests + 58 new offline-knowledge tests.**
Failures: 0. Errors: 0. Skipped: 0. Linux, Python 3.13.5.

See [offline-test-results.json](metadata/offline-test-results.json) for the full run
log, exact test counts, failures/skips, environment and network guard. The enforced
run blocks network syscalls in the test process and children using Linux seccomp.
Both IPv4/IPv6 probes and a child Python probe were actually denied by the kernel.
This is stronger than mocked HTTP or a flag saying "offline". It is not isolation
of the whole host or AI service.

The suite retains all 59 original tests and adds real-snapshot lookup/read/hash,
full-text, every-file exact-path and every Excel api_name lookup checks, ZIP and
local-directory import roundtrips, encoding and unsafe-path tests, relocation,
ambiguity handling, source-gap preservation and local link/include resolution.

Original COM tests use mocks only. Agent scenario contracts remain test resources,
not evaluated live model responses. No compiler, Excel, MSForms or COM runtime was
used to assert success.

## Preserved artifacts
The 19 expert modules, seven workflows, eight original Python files (seven tools
and one helper) and original VBA/code examples remain byte-for-byte identical to
v1. Navigation, entrypoint and reports were updated; three new source-retrieval
code files and verification resources were added. See the preservation manifest.

## Reproduce
```text
python scripts/validate_skill.py
python scripts/offline_docs.py verify
python -m unittest discover -s tests -v
python tests/run_offline_tests.py --enforce-os --out "<new-report.json>"
```
The last command requires Linux/libseccomp; a failure to install its filter is an
error, not a silent fallback to mocks. All normal source operations need only
Python 3.10+/SQLite FTS5. Windows symlink tests can be explicitly skipped when that
OS privilege is unavailable; the delivered Linux run reports its actual skip count.

## Not executed
Live Windows installation, VS Code/Qwen/Copilot auto-invocation, agent adherence,
VBA compilation, Excel runtime/COM/UI and end-to-end business application tests.
No claim of semantic correctness for all Microsoft examples or current upstream
snapshot equality. The source corpus is measured against the user-supplied ZIP.
